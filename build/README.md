# Build output

## Rebuild the add-in

From the repo root, with Excel closed (or at least without `ExcelVbaLib.xlam` loaded):

```powershell
powershell -ExecutionPolicy Bypass -File scripts/Build-ExcelVbaLib.ps1
```

That imports the snapshot under `source/Internal`, `Api`, and `Menus` into **one** project and saves:

`build/ExcelVbaLib.xlam`

Requires **Trust access to the VBA project object model** (see [docs/ExportImport.md](../docs/ExportImport.md)).

**The add-in is the library.** Add new modules in the VBE of `ExcelVbaLib.xlam`, then save the `.xlam`. Do not drop new `.bas` files into `source/` as a substitute for editing the add-in.

Binary `.xlam` files are gitignored. **`git pull` does not change `build\ExcelVbaLib.xlam`.** After pulling source, refresh that file from the **repo folder** (not `C:\Windows\System32`):

```powershell
cd "C:\Users\chanp\OneDrive\Notebooks\Cursor\Excel-VBA-Library"
powershell -ExecutionPolicy Bypass -File .\scripts\Import-AddinModules.ps1 -All
```

That writes Internal + Api + Menus (including Matrices) into `build\ExcelVbaLib.xlam` and saves. Excel can be open with the add-in loaded. The script lists public names in `modApiMatrices` and fails if those matrix modules are missing. Running **without** `-All` only replaces Data modules, which leaves matrix functions missing or stale.

A full rebuild (add-in **not** loaded) is still:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\Build-ExcelVbaLib.ps1
```

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
