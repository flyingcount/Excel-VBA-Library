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

Binary `.xlam` files are gitignored.

## Load it in Excel

1. File → Options → Add-ins → Manage **Excel Add-ins** → Go
2. Browse to `build/ExcelVbaLib.xlam` → OK
3. Leave **ExcelVbaLib** checked

The **Excel VBA Lib** menu is built in `Workbook_Open` (add-ins do not run `Auto_Open` at startup).

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
