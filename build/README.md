# Build output

`build/ExcelVbaLib.xlam` is the add-in Excel loads. Source-only changes do **not** update it.

## Rebuild the add-in (no Excel)

From the repo root:

```bash
pip install pyOpenVBA
python scripts/Build-ExcelVbaLib.py
```

That writes `build/ExcelVbaLib.xlam` from `source/Internal`, `Api`, and `Menus` (ThisWorkbook is injected, not imported as a second class).

## Rebuild with Excel (optional compile)

With Excel closed (or at least without `ExcelVbaLib.xlam` loaded):

```powershell
powershell -ExecutionPolicy Bypass -File scripts/Build-ExcelVbaLib.ps1
```

Requires **Trust access to the VBA project object model** (see [docs/ExportImport.md](../docs/ExportImport.md)).

**Replace a loaded add-in:** quit Excel fully, copy the new `build/ExcelVbaLib.xlam` over the loaded file, then start Excel and load the add-in again.

**The add-in is the library.** After a rebuild, grow it in the VBE of `ExcelVbaLib.xlam` or edit `source/` and run `Build-ExcelVbaLib.py` again.

## Load it in Excel

1. File → Options → Add-ins → Manage **Excel Add-ins** → Go
2. Browse to `build/ExcelVbaLib.xlam` → OK
3. Leave **ExcelVbaLib** checked

The **Excel VBA Lib** menu is scheduled from `Workbook_Open` / `Workbook_AddinInstall` (`InstallExcelVbaLibMenu`). Add-ins do not run `Auto_Open` at startup, and Excel also skips an explicit `Call Auto_Open` from the add-in.

If Data macros (e.g. Poisson) are stale after a source change, with Excel open and the add-in loaded:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/Import-AddinModules.ps1
```

That removes and re-imports `modInternalData` then `modApiData` into that `.xlam` and saves. Pass `-XlamPath` if the add-in is not `build\ExcelVbaLib.xlam` under the repo. Pass `-Modules` to replace other components (Internal first).

If the menu is missing after a source change, with Excel open and the add-in loaded:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/Inject-ThisWorkbook.ps1
```

Then restart Excel once. Do **not** put the `.xlam` in XLSTART; load it via **Excel Add-ins**.

Do **not** import `modApiBenford` into the caller workbook. The add-in already contains the full call graph.

## Run Benford

Put numbers in a range, then either:

- Menu **Excel VBA Lib** → **Benford** → **First digit** (prompts for a range), or
- Immediate window / another macro:

```vb
Application.Run "BenfordAnalysisFirstDigit", Range("A2:A50")
```

Add-in macros do not appear in Alt+F8. `Application.Run` or the menu is the supported entry point.

Optional: Tools → References → **ExcelVbaLib**, then you can write `BenfordAnalysisFirstDigit Range("A2:A50")` without `Application.Run`.
