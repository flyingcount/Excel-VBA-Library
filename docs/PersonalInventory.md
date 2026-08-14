# Personal123.xlsb inventory

Exported **2026-08-09** after Trust access was enabled.  
Source workbook: `Data/Personal123.xlsb`  
Raw export: `source/_export_raw/` (gitignored dumps + `inventory.csv`)

## Scale

| Item | Count |
|------|------:|
| VBComponents | 217 |
| Standard modules | ~204 |
| UserForms | 10 |
| Document modules with code | `ThisWorkbook` |
| Procedures catalogued | see `inventory.csv` |

## Natural groupings (keep Call links inside each family first)

| Prefix / area | Role | Home |
|---------------|------|------|
| `Custom_Menu*` domain tools | Stats, XmR, matrices, editing, output, … | `source/Features/<family>/` — [source/README.md](../source/README.md) |
| `Fn_*` | Worksheet-callable UDFs | `source/Udf/` (keep Public Function names) |
| `aPublicProcedures` / `aPublic*` | Shared library hubs | `source/Api/` + `source/Internal/` |
| `Filehandling`, `JPEGS`, `Sampling`, `Maths` | Shared utilities | `Api` / `Internal` per [ModuleMap.md](ModuleMap.md) |
| UserForms | `.frm` / `.frx` | `source/Forms/` |
| `Custom_Menu_Menus`, `ThisWorkbook` | Ribbon / open handlers | `source/Menus/` |
| `z_*` / WIP | Incomplete or snippet templates | `source/Sandbox/` — do not import to `.xlam` until reviewed |
| `Cipher_*`, niche tools | Domain features | `source/Features/Other/` until a family folder exists |

## Critical finding: duplicated sheet/array helpers

See the full classification in **[InfrastructureCatalog.md](InfrastructureCatalog.md)**.

Highlights:

- `WriteArrayToWorksheet`: ~127 call sites, **21 module copies, 12 body variants** (aPublicProcedures is newest, not universal).
- `TurnOff/OnScreeupdates…`: huge fan-in; superseded conceptually by `SpeedOn`/`SpeedOff`.
- `FormatOutputSheet`: 24 copies but **24 different bodies** → feature formatting, not shared infra.
- `zStockCode*` / `z_WIP`: snippet templates; harvest ideas, don’t treat as the runtime library.
- Broader intentional hubs: entire `aPublic*` family (Procedures, RangeValid, RangeFn, Tables, …) — under-adopted by menus.

## Recommended next steps

1. **Done:** Trust access + export to `_export_raw`.
2. Diff several `WriteArrayToWorksheet` copies; pick the best → paste into `source/Internal/modInternalSheetIO.bas`.
3. In **one pilot feature module** (e.g. `Custom_Menu11_Histogram`), delete the private copy and `Call modInternalSheetIO.DumpArray` (or keep the name `WriteArrayToWorksheet` as a thin Public wrapper in SheetIO).
4. Compile; repeat for other features.
5. Build `ExcelVbaLib.xlam` only after Internal helpers stabilize — do not try to import all 200 modules on day one.

## Regenerating the export

```powershell
# From repo root, with Excel closed or allowing reopen:
# Re-run the export script / ask the agent to re-export Data\Personal123.xlsb
```

Refresh `inventory.csv` whenever Personal changes materially.
