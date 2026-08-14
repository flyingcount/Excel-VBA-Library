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
├── source/
│   ├── Api/                 ← Public entry points (incl. CreateDateTable)
│   ├── Internal/            ← Shared helpers (array→sheet, date table build, …)
│   └── _export_raw/         ← Local dumps from Personal123 (gitignored)
├── build/                   ← ExcelVbaLib.xlam (local; gitignored)
└── Data/                    ← Personal123.xlsb (local; gitignored)
```

## Quick start

1. Enable **Trust access to the VBA project object model** — see [docs/ExportImport.md](docs/ExportImport.md).
2. Export or refresh from `Data/Personal123.xlsb` into `source/_export_raw/` when needed.
3. Curate shared code into `source/Api` and `source/Internal` using [docs/ModuleMap.md](docs/ModuleMap.md) and [docs/InfrastructureCatalog.md](docs/InfrastructureCatalog.md).
4. Create `build/ExcelVbaLib.xlam`, import curated modules, compile, and load as an Excel add-in.

## License

MIT — see [LICENSE](LICENSE).
