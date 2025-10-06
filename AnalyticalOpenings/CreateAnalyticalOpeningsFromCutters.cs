using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace AnalyticalOpenings
{
    [Transaction(TransactionMode.Manual)]
    public class CreateAnalyticalOpeningsFromCutters : IExternalCommand
    {
        private const double PlaneThickness = 0.01; // feet, thin prism thickness
        private const double NormalParallelTol = 1e-3;
        private const double GeometryTol = 1e-6;
        private static readonly string LogPath = Path.Combine(Path.GetTempPath(), "AnalyticalOpeningsLog.txt");

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uiapp = commandData.Application;
            var uidoc = uiapp.ActiveUIDocument;
            var doc = uidoc.Document;

            SafeLog("=== CreateAnalyticalOpeningsFromCutters START ===");

            IList<Reference> panelRefs;
            IList<Reference> cutterRefs;

            try
            {
                panelRefs = uidoc.Selection.PickObjects(ObjectType.Element, new AnalyticalPanelFilter(),
                    "Выберите аналитические панели, в которых нужно сделать проемы");
                if (panelRefs == null || panelRefs.Count == 0)
                {
                    TaskDialog.Show("Панели", "Не выбрано ни одной панели.");
                    return Result.Cancelled;
                }

                cutterRefs = uidoc.Selection.PickObjects(ObjectType.Element, new CutterFilter(),
                    "Выберите элементы (Generic Model/Shaft), определяющие контуры проемов");
                if (cutterRefs == null || cutterRefs.Count == 0)
                {
                    TaskDialog.Show("Режущие элементы", "Не выбрано ни одного режущего элемента.");
                    return Result.Cancelled;
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Ошибка", ex.Message);
                return Result.Failed;
            }

            var panels = panelRefs
                .Select(r => doc.GetElement(r))
                .OfType<AnalyticalPanel>()
                .ToList();

            var cutters = cutterRefs
                .Select(r => doc.GetElement(r))
                .Where(e => e != null)
                .ToList();

            if (panels.Count == 0)
            {
                TaskDialog.Show("Панели", "Не удалось получить ни одной аналитической панели.");
                return Result.Cancelled;
            }
            if (cutters.Count == 0)
            {
                TaskDialog.Show("Режущие элементы", "Не удалось получить ни одного cutter-элемента.");
                return Result.Cancelled;
            }

            using (var tx = new Transaction(doc, "Создать аналитические проемы"))
            {
                tx.Start();

                foreach (var panel in panels)
                {
                    try
                    {
                        ProcessPanel(doc, panel, cutters);
                    }
                    catch (Exception exPanel)
                    {
                        SafeLog($"Панель {panel.Id}: ошибка обработки — {exPanel.Message}");
                    }
                }

                tx.Commit();
            }

            SafeLog("=== CreateAnalyticalOpeningsFromCutters END ===");
            return Result.Succeeded;
        }

        //————————————————————————————————————————————————————————————————
        // Process one AnalyticalPanel: build prism, intersect cutters,
        // extract unique opening loops, and create AnalyticalOpenings
        //————————————————————————————————————————————————————————————————
        private void ProcessPanel(Document doc, AnalyticalPanel panel, List<Element> cutters)
        {
            SafeLog($"=== Обработка панели {panel.Id} ===");

            // 1. Get panel contour and plane
            var panelLoop = panel.GetOuterContour();
            if (panelLoop == null || !panelLoop.Any())
            {
                SafeLog($"Панель {panel.Id}: пустой контур — пропуск");
                return;
            }

            var panelPlane = GetPanelPlane(panelLoop);
            if (panelPlane == null)
            {
                SafeLog($"Панель {panel.Id}: не удалось определить плоскость — пропуск");
                return;
            }

            double tol = Math.Max(doc.Application.ShortCurveTolerance, GeometryTol);
            var cleanPanelLoop = BuildCleanLoopOrderedSafe(panelLoop, panelPlane, tol);
            if (cleanPanelLoop == null)
            {
                SafeLog($"Панель {panel.Id}: не удалось очистить контур — пропуск");
                return;
            }

            // 2. Build thin prism solid for clipping
            var panelPrism = CreatePanelPrismSolid(cleanPanelLoop, panelPlane, PlaneThickness);
            if (panelPrism == null)
            {
                SafeLog($"Панель {panel.Id}: не удалось создать призму — пропуск");
                return;
            }

            // 3. Collect candidate opening loops from all cutters
            var candidateLoops = new List<CurveLoop>();
            foreach (var cutter in cutters)
            {
                foreach (var cutterSolid in GetElementSolids(cutter))
                {
                    try
                    {
                        var clipped = BooleanOperationsUtils.ExecuteBooleanOperation(
                            cutterSolid, panelPrism, BooleanOperationsType.Intersect);

                        if (clipped == null) continue;

                        var loops = ExtractPlanarOpeningLoops(clipped, panelPlane, tol);
                        candidateLoops.AddRange(loops);
                    }
                    catch (Exception ex)
                    {
                        SafeLog($"Панель {panel.Id} × Cutter {cutter.Id}: ошибка пересечения — {ex.Message}");
                    }
                }
            }

            if (candidateLoops.Count == 0)
            {
                SafeLog($"Панель {panel.Id}: не найдено контуров проемов");
                return;
            }

            // 4. Deduplicate loops
            var uniqueLoops = new HashSet<CurveLoop>(new CurveLoopComparer(tol));
            foreach (var loop in candidateLoops)
            {
                var adjusted = ProjectCurveLoopToPlane(loop, cleanPanelLoop);
                if (adjusted == null) continue;

                var finalLoop = CloseAndClean(adjusted.Select(c => c.GetEndPoint(0)).ToList(), panelPlane, tol);
                if (finalLoop == null) continue;

                if (!uniqueLoops.Add(finalLoop))
                {
                    SafeLog($"Панель {panel.Id}: дубликат петли — пропуск");
                }
            }

            // 5. Create openings for unique loops
            foreach (var loop in uniqueLoops)
            {
                try
                {
                    var centroid = GetCentroid(loop);
                    if (!IsPointInsideCurveLoop(cleanPanelLoop, centroid))
                    {
                        SafeLog($"Панель {panel.Id}: петля вне границ панели — пропуск");
                        continue;
                    }

                    var opening = AnalyticalOpening.Create(doc, loop, panel.Id);
                    if (opening != null)
                        SafeLog($"Панель {panel.Id}: создан проем {opening.Id}");
                    else
                        SafeLog($"Панель {panel.Id}: не удалось создать проем");
                }
                catch (Exception ex)
                {
                    SafeLog($"Панель {panel.Id}: ошибка создания проема — {ex.Message}");
                }
            }
        }

        // ---------- Geometry helpers (mostly adapted from your code) ----------

        private Plane GetPanelPlane(CurveLoop contour)
        {
            var pts = new List<XYZ>();
            foreach (Curve c in contour) pts.Add(c.GetEndPoint(0));
            if (pts.Count < 3) return null;

            for (int i = 0; i < pts.Count - 2; i++)
            {
                var v1 = pts[i + 1] - pts[i];
                for (int j = i + 2; j < pts.Count; j++)
                {
                    var v2 = pts[j] - pts[i];
                    var n = v1.CrossProduct(v2);
                    if (n.GetLength() > GeometryTol)
                        return Plane.CreateByNormalAndOrigin(n.Normalize(), pts[i]);
                }
            }
            return null;
        }

        private CurveLoop BuildCleanLoopOrderedSafe(CurveLoop rawLoop, Plane plane, double tol)
        {
            var verts = new List<XYZ>();
            foreach (Curve c in rawLoop)
            {
                var p0 = ProjectPointToPlane(c.GetEndPoint(0), plane);
                var p1 = ProjectPointToPlane(c.GetEndPoint(1), plane);
                if (verts.Count == 0) verts.Add(p0);
                else if (!verts.Last().IsAlmostEqualTo(p0, tol)) verts.Add(p0);
                if (!p0.IsAlmostEqualTo(p1, tol)) verts.Add(p1);
            }
            if (verts.Count < 3) return null;
            if (!verts.First().IsAlmostEqualTo(verts.Last(), tol)) verts.Add(verts.First());

            verts = CollapseTinyAndColinear(verts, tol);
            if (verts.Count < 4) return null;

            var loop = new CurveLoop();
            for (int i = 0; i < verts.Count - 1; i++)
            {
                var a = verts[i]; var b = verts[i + 1];
                if (a.DistanceTo(b) > tol) loop.Append(Line.CreateBound(a, b));
            }

            var start = loop.First().GetEndPoint(0);
            var end = loop.Last().GetEndPoint(1);
            if (!start.IsAlmostEqualTo(end, tol)) loop.Append(Line.CreateBound(end, start));

            return loop;
        }

        private List<XYZ> CollapseTinyAndColinear(List<XYZ> v, double tol)
        {
            if (v.Count < 3) return v;
            var outV = new List<XYZ>();
            for (int i = 0; i < v.Count; i++)
            {
                if (outV.Count == 0 || outV.Last().DistanceTo(v[i]) > tol) outV.Add(v[i]);
            }
            if (outV.Count < 3) return outV;
            if (!outV.First().IsAlmostEqualTo(outV.Last(), tol)) outV.Add(outV.First());

            int iIdx = 0;
            while (outV.Count > 3 && iIdx < outV.Count - 2)
            {
                var a = outV[iIdx]; var b = outV[iIdx + 1]; var c = outV[iIdx + 2];
                var abVec = b - a; var bcVec = c - b;
                double abLen = abVec.GetLength(), bcLen = bcVec.GetLength();
                if (abLen <= tol || bcLen <= tol) { iIdx++; continue; }
                var ab = abVec / abLen; var bc = bcVec / bcLen;
                double dot = ab.DotProduct(bc);
                double sinMag = ab.CrossProduct(bc).GetLength();
                if (sinMag <= 1e-6 && dot > 0.9999) outV.RemoveAt(iIdx + 1);
                else iIdx++;
            }
            if (!outV.First().IsAlmostEqualTo(outV.Last(), tol)) outV.Add(outV.First());
            return outV;
        }

        private XYZ ProjectPointToPlane(XYZ p, Plane plane)
        {
            double d = (p - plane.Origin).DotProduct(plane.Normal);
            return p - d * plane.Normal;
        }

        private CurveLoop CloseAndClean(List<XYZ> verts, Plane plane, double tol)
        {
            var v = new List<XYZ>();
            foreach (var p in verts)
            {
                var q = ProjectPointToPlane(p, plane);
                if (v.Count == 0 || !v.Last().IsAlmostEqualTo(q, tol)) v.Add(q);
            }
            if (v.Count < 3) return null;
            if (!v.First().IsAlmostEqualTo(v.Last(), tol)) v.Add(v.First());

            v = CollapseTinyAndColinear(v, tol);
            if (v.Count < 4) return null;

            // Ensure CCW orientation
            var n = plane.Normal.Normalize();
            XYZ u = null;
            for (int i = 0; i < v.Count - 1 && u == null; i++)
            {
                var d = v[i + 1] - v[i]; var cand = d - n.DotProduct(d) * n;
                if (cand.GetLength() > tol) u = cand.Normalize();
            }
            if (u == null) u = XYZ.BasisX;
            var w = n.CrossProduct(u).Normalize();

            double area2 = 0;
            for (int i = 0; i < v.Count - 1; i++)
            {
                var a = v[i] - plane.Origin; var b = v[i + 1] - plane.Origin;
                var ax = a.DotProduct(u); var ay = a.DotProduct(w);
                var bx = b.DotProduct(u); var by = b.DotProduct(w);
                area2 += (ax * by - bx * ay);
            }
            if (area2 < 0) v.Reverse();

            var loop = new CurveLoop();
            for (int i = 0; i < v.Count - 1; i++)
            {
                var a = v[i]; var b = v[i + 1];
                if (a.DistanceTo(b) > tol) loop.Append(Line.CreateBound(a, b));
            }
            var start = loop.First().GetEndPoint(0);
            var end = loop.Last().GetEndPoint(1);
            if (!start.IsAlmostEqualTo(end, tol)) loop.Append(Line.CreateBound(end, start));

            return loop;
        }

        private XYZ GetCentroid(CurveLoop loop)
        {
            var pts = new List<XYZ>();
            foreach (Curve c in loop) pts.Add(c.GetEndPoint(0));
            double sx = 0, sy = 0, sz = 0; int n = pts.Count;
            foreach (var p in pts) { sx += p.X; sy += p.Y; sz += p.Z; }
            return new XYZ(sx / n, sy / n, sz / n);
        }

        private bool IsPointInsideCurveLoop(CurveLoop loop, XYZ testPoint)
        {
            var pts = new List<XYZ>();
            foreach (Curve c in loop) pts.Add(c.GetEndPoint(0));
            if (pts.Count < 3) return false;

            var n = (pts[1] - pts[0]).CrossProduct(pts[2] - pts[0]).Normalize();
            var plane = Plane.CreateByNormalAndOrigin(n, pts[0]);

            var u = (pts[1] - pts[0]).Normalize();
            var v = n.CrossProduct(u);

            Func<XYZ, (double x, double y)> proj = p =>
            {
                var d = p - plane.Origin;
                return (d.DotProduct(u), d.DotProduct(v));
            };

            var poly = pts.Select(q => proj(q)).ToList();
            var pt = proj(testPoint);

            bool inside = false;
            for (int i = 0, j = poly.Count - 1; i < poly.Count; j = i++)
            {
                bool cond = ((poly[i].y > pt.y) != (poly[j].y > pt.y)) &&
                            (pt.x < (poly[j].x - poly[i].x) * (pt.y - poly[i].y) / (poly[j].y - poly[i].y) + poly[i].x);
                if (cond) inside = !inside;
            }
            return inside;
        }

        private CurveLoop ProjectCurveLoopToPlane(CurveLoop loop, CurveLoop targetContour)
        {
            var pts = new List<XYZ>();
            foreach (Curve c in targetContour) pts.Add(c.GetEndPoint(0));
            if (pts.Count < 3) return null;
            var n = (pts[1] - pts[0]).CrossProduct(pts[2] - pts[0]).Normalize();
            var plane = Plane.CreateByNormalAndOrigin(n, pts[0]);

            var outLoop = new CurveLoop();
            foreach (Curve c in loop)
            {
                var a = c.GetEndPoint(0); var b = c.GetEndPoint(1);
                var pa = ProjectPointToPlane(a, plane);
                var pb = ProjectPointToPlane(b, plane);
                outLoop.Append(Line.CreateBound(pa, pb));
            }
            return outLoop;
        }

        // ---------- Panel prism and solid utilities ----------

        private Solid CreatePanelPrismSolid(CurveLoop panelLoop, Plane plane, double thickness)
        {
            try
            {
                // Ensure loop is valid and planar
                var loops = new List<CurveLoop> { panelLoop };
                var dir = plane.Normal;
                double dist = Math.Max(thickness, PlaneThickness);
                return GeometryCreationUtilities.CreateExtrusionGeometry(loops, dir, dist);
            }
            catch
            {
                // Try extruding both directions if needed (ensure it crosses cutter)
                try
                {
                    var loops = new List<CurveLoop> { panelLoop };
                    var dir = -plane.Normal;
                    double dist = Math.Max(thickness, PlaneThickness);
                    return GeometryCreationUtilities.CreateExtrusionGeometry(loops, dir, dist);
                }
                catch
                {
                    return null;
                }
            }
        }

        private IEnumerable<Solid> GetElementSolids(Element e)
        {
            var opts = new Options
            {
                ComputeReferences = true,
                DetailLevel = ViewDetailLevel.Fine,
                IncludeNonVisibleObjects = true
            };

            var ge = e.get_Geometry(opts);
            if (ge == null) yield break;

            foreach (var obj in ge)
            {
                if (obj is Solid s && s.Volume > 1e-9)
                {
                    yield return s;
                }
                else if (obj is GeometryInstance gi)
                {
                    var insp = gi.GetInstanceGeometry();
                    foreach (var sub in insp)
                    {
                        if (sub is Solid ss && ss.Volume > 1e-9)
                            yield return ss;
                    }
                }
            }
        }

        private List<CurveLoop> ExtractPlanarOpeningLoops(Solid solid, Plane panelPlane, double tol)
        {
            var loops = new List<CurveLoop>();
            foreach (Face f in solid.Faces)
            {
                if (!(f is PlanarFace pf)) continue;

                var faceNormal = pf.FaceNormal.Normalize();
                var dot = Math.Abs(faceNormal.DotProduct(panelPlane.Normal));
                if (Math.Abs(dot - 1.0) > NormalParallelTol) continue; // not parallel to panel plane

                // Extract edge loops from face; convert to CurveLoops
                var edgeLoops = pf.GetEdgesAsCurveLoops();
                foreach (CurveLoop cl in edgeLoops)
                {
                    // Clean and return
                    var cleaned = BuildCleanLoopOrderedSafe(cl, panelPlane, tol);
                    if (cleaned != null && cleaned.Any())
                        loops.Add(cleaned);
                }
            }
            // Prefer larger loops (outer perimeter) if multiple; keep all as each may represent a hole
            return loops;
        }

        // ---------- Selection filters ----------

        private class AnalyticalPanelFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem) => elem is AnalyticalPanel;
            public bool AllowReference(Reference reference, XYZ position) => false;
        }

        private class CutterFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            {
                var cat = elem.Category;
                if (cat == null) return false;
                var bic = (BuiltInCategory)cat.Id.IntegerValue;

                // Allow Generic Models and Shaft Openings (common for “cutter” inputs)
                if (bic == BuiltInCategory.OST_GenericModel) return true;
                if (bic == BuiltInCategory.OST_ShaftOpening) return true;

                // Also allow regular Opening elements (e.g., wall/floor openings) if present
                if (elem is Opening) return true;

                return false;
            }
            public bool AllowReference(Reference reference, XYZ position) => false;
        }

        // ---------- Logging ----------

        private void SafeLog(string msg)
        {
            try
            {
                File.AppendAllText(LogPath, $"{DateTime.Now:HH:mm:ss} - {msg}{Environment.NewLine}");
            }
            catch { /* no UI noise on logging issues */ }
        }
    }

    class CurveLoopComparer : IEqualityComparer<CurveLoop>
    {
        private readonly double _tol;
        public CurveLoopComparer(double tol) { _tol = tol; }

        public bool Equals(CurveLoop x, CurveLoop y)
        {
            if (x == null || y == null) return false;
            var vx = x.Select(c => c.GetEndPoint(0)).ToList();
            var vy = y.Select(c => c.GetEndPoint(0)).ToList();
            if (vx.Count != vy.Count) return false;

            // Try to align starting point
            for (int shift = 0; shift < vy.Count; shift++)
            {
                bool allEqual = true;
                for (int i = 0; i < vx.Count; i++)
                {
                    var pi = vx[i];
                    var qj = vy[(i + shift) % vy.Count];
                    if (!pi.IsAlmostEqualTo(qj, _tol)) { allEqual = false; break; }
                }
                if (allEqual) return true;
            }
            return false;
        }

        public int GetHashCode(CurveLoop obj)
        {
            if (obj == null) return 0;
            // crude hash: sum of rounded coords
            var pts = obj.Select(c => c.GetEndPoint(0));
            double sx = 0, sy = 0, sz = 0;
            foreach (var p in pts) { sx += Math.Round(p.X / _tol); sy += Math.Round(p.Y / _tol); sz += Math.Round(p.Z / _tol); }
            return sx.GetHashCode() ^ sy.GetHashCode() ^ sz.GetHashCode();
        }
    }



    // ---------- Extension helpers matching your style (XYZ equality, etc.) ----------

    public static class XyzExtensions
    {
        public static bool IsAlmostEqualTo(this XYZ a, XYZ b, double tol)
        {
            return a.DistanceTo(b) <= tol;
        }
    }
}
