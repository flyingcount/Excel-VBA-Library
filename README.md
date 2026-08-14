# Excel VBA Library

Git-friendly **Excel VBA** library: curated Api/Internal modules for an `.xlam` add-in, plus migration notes from a long-lived personal macro workbook.

**Design rule:** one Excel add-in project holds the full call graph. Public API modules call Internal helpers with normal `Call` statements — do not split helpers across workbooks.

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
│   └── Build-ExcelVbaLib.ps1  ← rebuild build/ExcelVbaLib.xlam from source/
├── source/                  ← see [source/README.md](source/README.md) for where each Personal123 family goes
│   ├── Api/                 ← Public entry points (CreateDateTable, Benford, …)
│   ├── Internal/            ← Shared helpers
│   ├── Features/            ← Domain packs (Stats, Matrices, Editing, …)
│   ├── Udf/                 ← Worksheet functions (`Fn_*`)
│   ├── Forms/               ← UserForms
│   ├── Classes/             ← Class modules
│   ├── Menus/               ← Ribbon / ThisWorkbook handlers
│   ├── Sandbox/             ← `z_*` / WIP (not imported until reviewed)
│   └── _export_raw/         ← Local dumps from Personal123 (gitignored)
├── build/                   ← ExcelVbaLib.xlam (local; gitignored)
└── Data/                    ← Personal123.xlsb (local; gitignored)
```

## Quick start

1. Enable **Trust access to the VBA project object model** — see [docs/ExportImport.md](docs/ExportImport.md).
2. Build the add-in (Excel must not have `ExcelVbaLib.xlam` loaded):

   ```powershell
   powershell -ExecutionPolicy Bypass -File scripts/Build-ExcelVbaLib.ps1
   ```

3. Excel → Options → Add-ins → Excel Add-ins → Browse → `build/ExcelVbaLib.xlam`.
4. Run Benford from the **Excel VBA Lib** menu, or:

   ```vb
   Application.Run "BenfordAnalysisFirstDigit", Range("A2:A50")
   ```

Do not import individual `modApi*` / `modInternal*` files into caller workbooks. Re-run the build script when you add modules under `source/`.

To migrate more code from Personal123, export into `source/_export_raw/` and re-home using [source/README.md](source/README.md).

## License

MIT — see [LICENSE](LICENSE).
