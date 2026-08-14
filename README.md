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
│   └── Build-ExcelVbaLib.ps1  ← rebuild build/ExcelVbaLib.xlam from the source snapshot
├── source/                  ← snapshot of modules already in the add-in (see source/README.md)
│   ├── Api/
│   ├── Internal/
│   ├── Menus/
│   └── _export_raw/         ← Local dumps from Personal123 (gitignored)
├── build/                   ← ExcelVbaLib.xlam (local; gitignored)
└── Data/                    ← Personal123.xlsb (local; gitignored)
```

## Quick start

1. Enable **Trust access to the VBA project object model** — see [docs/ExportImport.md](docs/ExportImport.md).
2. Build the add-in if you do not already have it (Excel must not have `ExcelVbaLib.xlam` loaded):

   ```powershell
   powershell -ExecutionPolicy Bypass -File scripts/Build-ExcelVbaLib.ps1
   ```

3. Excel → Options → Add-ins → Excel Add-ins → Browse → `build/ExcelVbaLib.xlam`.
4. Run Benford or worksheet templates from the **Excel VBA Lib** menu, or:

   ```vb
   Application.Run "BenfordAnalysisFirstDigit", Range("A2:A50")
   Application.Run "CreateLinksTemplate"
   ```

Grow the library in the add-in (`Alt+F11` on `ExcelVbaLib.xlam`), then save the `.xlam`. Do not import individual `modApi*` / `modInternal*` files into caller workbooks.

## License

MIT — see [LICENSE](LICENSE).
