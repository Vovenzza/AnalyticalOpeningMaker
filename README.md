# Analytical Openings for Revit

A Revit 2023 add‑in that automatically creates **analytical openings** in selected analytical panels based on intersections with selected cutter elements (Generic Models, Shaft Openings).

---

## ✨ Features

- Select multiple **analytical panels** (vertical, horizontal, any orientation).
- Select **cutter elements** (Generic Models, Shaft Openings, Openings).
- Automatically computes intersection contours between cutters and panels.
- Creates **AnalyticalOpening** objects clipped to the panel’s boundaries.
- Deduplication logic prevents duplicate openings in the same location.
- Detailed logging to a text file for troubleshooting.

---

## 📦 Installation

### Option 1: Using the EXE (recommended for end‑users)

1. Download the latest release `.exe` from the [Releases](../../releases) page.
2. Run the installer/executable.  
   - It will copy the compiled DLL and a ready `.addin` manifest into your Revit add‑ins folder automatically.
   - Default location:  
     ```
     %AppData%\Autodesk\Revit\Addins\2023\
     ```
3. Restart Revit. The command will appear under **Add‑Ins → External Tools**.

### Option 2: Manual deployment (for developers)

1. **Build the project** in Visual Studio (Release configuration).
2. Copy the compiled DLL (e.g. `AnalyticalOpenings.dll`) to a folder of your choice, e.g.:

