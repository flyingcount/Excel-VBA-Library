# Export / import notes

## Prerequisites

1. Excel → **File → Options → Trust Center → Trust Center Settings → Macro Settings**
2. Enable **Trust access to the VBA project object model**
3. Save and restart Excel if prompted

Without this, automation cannot list or export modules from `Personal123.xlsb` (you will see an empty VBProject).

## Naming conventions

| Kind | Prefix | Example |
|------|--------|---------|
| Public API standard module | `modApi` | `modApiArrays.bas` |
| Internal helper module | `modInternal` | `modInternalSheetIO.bas` |
| Class module (optional) | `cls` | `clsTimer.cls` |
| Public procedure | verb + noun | `WriteArrayToSheet` |
| Private helper | same module, `Private` | `Private Function LastUsedRow(...)` |

`Attribute VB_Name` inside each file **must** match the file stem (without extension).

## Manual export (from Personal123.xlsb or the .xlam)

1. Open the workbook/add-in → `Alt+F11`
2. For each module to migrate:
   - Right-click module → **Export File…**
   - Save under `source/Api/` or `source/Internal/` per [ModuleMap.md](ModuleMap.md)
3. Prefer **re-home** code into the target module names in the map (rename on export), rather than keeping legacy names forever
4. Commit the `.bas` / `.cls` text files

## Manual import (into the add-in)

1. Create a new workbook → save as **Excel Add-in** (`*.xlam`) under `build/ExcelVbaLib.xlam`
2. `Alt+F11` → File → **Import File…** for each `source/**/*.bas`
3. Set project name (VBAProject properties) to e.g. `ExcelVbaLib`
4. Save the add-in
5. Excel → Options → Add-ins → manage Excel Add-ins → browse and load `ExcelVbaLib.xlam`

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

Then sort exported files into `Api/` vs `Internal/` using the module map.

## Do not commit

- Binary add-ins unless you explicitly want them (`build/*.xlam` is gitignored by default)
- Exported junk under `source/_export_raw/` (also ignored)
