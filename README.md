# Excel VBA Library

**Excel VBA** add-in (`ExcelVbaLib.xlam`) plus migration notes from a long-lived personal macro workbook.

**The library is the add-in.** Add new modules in the VBE of `ExcelVbaLib.xlam`. One Excel add-in project holds the full call graph — do not import library modules into caller workbooks.

Related: [PowerQuery-Library](https://github.com/flyingcount/PowerQuery-Library) (Power Query M functions).

## Layout

```text
.
├── README.md
├── CHANGELOG.md
├── LICENSE
├── docs/
│   ├── ExportImport.md
│   ├── ModuleMap.md
│   ├── PersonalInventory.md
│   └── InfrastructureCatalog.md
├── scripts/
│   ├── Build-ExcelVbaLib.ps1      ← rebuild build/ExcelVbaLib.xlam from the source snapshot
│   ├── Import-AddinModules.ps1    ← replace modules in a specific .xlam (default: build\)
│   ├── Import-Menu13FromPersonal.ps1 ← copy Personal Matrices1/2 into the .xlam
│   ├── Inject-ThisWorkbook.ps1    ← write Workbook_Open into a loaded add-in (no full rebuild)
│   └── Export-PersonalModules.ps1 ← dump Menu13 (and other) modules from Personal.xlsb
├── source/                  ← snapshot of modules already in the add-in (see source/README.md)
│   ├── Api/
│   ├── Internal/
│   ├── Menus/
│   └── _export_raw/         ← Local dumps from Personal123 (gitignored)
├── build/                   ← ExcelVbaLib.xlam
└── Data/                    ← Personal123.xlsb (local; gitignored)
```

## Quick start

1. Enable **Trust access to the VBA project object model** — see [docs/ExportImport.md](docs/ExportImport.md).
2. Build the add-in if you do not already have it (Excel must not have `ExcelVbaLib.xlam` loaded):

   ```powershell
   powershell -ExecutionPolicy Bypass -File scripts/Build-ExcelVbaLib.ps1
   ```

3. Excel → Options → Add-ins → Excel Add-ins → Browse → `build/ExcelVbaLib.xlam`.
   After `git pull`, that file is already in the repo. To rebuild from `source/` **in Windows PowerShell on the PC that has Excel** (not the cloud Linux terminal):

   ```powershell
   powershell -ExecutionPolicy Bypass -File .\scripts\Import-AddinModules.ps1 -All
   powershell -ExecutionPolicy Bypass -File .\scripts\Import-Menu13FromPersonal.ps1
   ```
4. Restart Excel. **Excel VBA Lib** should appear on the **Add-ins** ribbon tab without running `Auto_Open`. If it does not, with Excel open:

   ```powershell
   powershell -ExecutionPolicy Bypass -File scripts/Inject-ThisWorkbook.ps1
   ```

   Then quit Excel fully and start it again.
5. Run tools from that menu, or:

   ```vb
   Application.Run "BenfordAnalysisFirstDigit", Range("A2:A50")
   Application.Run "CreateLinksTemplate"
   Application.Run "ExtractSample", Range("A1:C20"), 50, False
   Application.Run "RandomIntegers"
   Application.Run "DataCombinations"
   Application.Run "ShowCustomLists"
   Application.Run "ListNamedRangeProperties"
   Application.Run "RangeAnalysis"
   Application.Run "FlagDuplicates"
   Application.Run "ListTableProperties"
   Application.Run "BatchListAllFiles_FolderSubfolders"
   Application.Run "MatrixMultiply"
   Application.Run "HistogramTableAndPlot"
   Application.Run "GenerateNormalPlot"
   ```

Grow the library in the add-in (`Alt+F11` on `ExcelVbaLib.xlam`), then save the `.xlam`. Do not import individual `modApi*` / `modInternal*` files into caller workbooks.

## License

MIT — see [LICENSE](LICENSE).
