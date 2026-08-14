# Export / import notes

## Prerequisites

1. Excel → **File → Options → Trust Center → Trust Center Settings → Macro Settings**
2. Enable **Trust access to the VBA project object model**
3. Save and restart Excel if prompted

Without this, automation cannot list or export modules from `Personal123.xlsb` (you will see an empty VBProject).

## Naming conventions

| Kind | Prefix | Example |
|------|--------|---------|
| Public API standard module | `modApi` | `modApiArrays` (in the add-in) |
| Internal helper module | `modInternal` | `modInternalSheetIO` (in the add-in) |
| Worksheet UDF module | `Fn_` | `Fn_Ageing` (keep Public Function names) |
| Class module | `cls` | `clsTimer` |
| UserForm | form name | `BoxCoxForm` |
| Public procedure | verb + noun | `WriteArrayToSheet` |
| Private helper | same module, `Private` | `Private Function LastUsedRow(...)` |

`Attribute VB_Name` inside each file **must** match the file stem (without extension).

## Add code to the library

Open `ExcelVbaLib.xlam` → `Alt+F11` → Insert → Module (or import a `.bas` **into the add-in**). Save the `.xlam`. Caller workbooks only load the add-in.

To copy from Personal123: export a module, then **File → Import File…** in the add-in project — not into a caller workbook.

## Build the add-in (one project, full call graph)

Caller workbooks should **load `ExcelVbaLib.xlam`**, not import Benford (or other) modules one by one.

```powershell
# From repo root; close the add-in in Excel first if it is already loaded
powershell -ExecutionPolicy Bypass -File scripts/Build-ExcelVbaLib.ps1
```

The script imports the `source/` snapshot (`Internal`, then `Api`, then `Menus`; skips `_export_raw/`), compiles, and writes `build/ExcelVbaLib.xlam`. After that, grow the library in the add-in itself.

### Load and run

1. Excel → Options → Add-ins → Manage **Excel Add-ins** → Go → Browse → `build/ExcelVbaLib.xlam`
2. Use the **Excel VBA Lib** menu, or from the Immediate window:

```vb
Application.Run "BenfordAnalysisFirstDigit", Range("A2:A50")
Application.Run "CreateDateTable", #1/1/2024#, #1/31/2024#
```

Add-in `Public Sub`s do not appear in Alt+F8. Optional: Tools → References → **ExcelVbaLib** for early-bound `Call`.

### Manual import (into the add-in)

Use this only if you are not running the build script:

1. Create a new workbook → save as **Excel Add-in** (`*.xlam`) under `build/ExcelVbaLib.xlam`
2. `Alt+F11` → File → **Import File…** for each `source/**/*.bas` (**Internal** first)
3. Set project name (VBAProject properties) to `ExcelVbaLib`
4. Save the add-in
5. Load it as an Excel add-in (steps above)

## Keep links (calls) intact

- Import **Internal** modules before rewriting callers, or import everything into the **same** `.xlam` project in one pass
- Inside the add-in, keep using `Call modInternalSheetIO` patterns / same-project `Call Helper`
- Only the **boundary** from other workbooks should use:
  - Tools → **References** → `ExcelVbaLib`, or
  - `Application.Run "WriteArrayToSheet", ...`

## Migration wrappers in Personal.xlsb

While moving code out of Personal:

```vb
' Personal.xlsb — temporary shim
Public Sub DumpArray()
    Application.Run "WriteArrayToSheet"  ' public name in the add-in
End Sub
```

Delete shims once callers use the add-in directly.

## Optional: batch export with Excel (after Trust is on)

From PowerShell (illustrative):

```powershell
$excel = New-Object -ComObject Excel.Application
$wb = $excel.Workbooks.Open("...\Data\Personal123.xlsb")
$out = "...\source\_export_raw"
New-Item -ItemType Directory -Force -Path $out | Out-Null
foreach ($comp in $wb.VBProject.VBComponents) {
  if ($comp.Type -in 1, 2, 3) {  # std / class / form
    $comp.Export((Join-Path $out ($comp.Name + ".bas")))
  }
}
$wb.Close($false)
$excel.Quit()
```

Then import chosen modules into **ExcelVbaLib.xlam**, not into a folder tree.

## Do not commit

- Binary add-ins unless you explicitly want them (`build/*.xlam` is gitignored by default)
- Exported junk under `source/_export_raw/` (also ignored)
