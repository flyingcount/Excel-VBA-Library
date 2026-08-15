# Test: add-in menu loads on Excel start

## Setup

1. Load `build/ExcelVbaLib.xlam` via **File → Options → Add-ins → Excel Add-ins** (not COM Add-ins, not XLSTART).
2. If ThisWorkbook was just updated, run `scripts/Inject-ThisWorkbook.ps1` once, then **fully quit Excel**.

## Steps

1. Start Excel with a blank workbook. Do not run `Auto_Open`.
2. Look for **Excel VBA Lib** on the **Add-ins** ribbon tab (classic Worksheet Menu Bar commands land there).

## Expected

- The menu is present within about a second of startup (OnTime waits until Excel is idle).
- Submenus: Benford, Data (with Probability distributions), Sampling, **Matrices** (Create / Operations / Properties / Validation / Decompositions), Worksheet templates.
- Unchecking the add-in removes the menu. Checking it again restores the menu without a restart (`Workbook_AddinInstall`).
- `Application.Run "Auto_Open"` still rebuilds the menu if you need it in the current session.
