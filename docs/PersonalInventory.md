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

Import the next family into **`ExcelVbaLib.xlam`** (same project). Do not re-home into git folders.

| Prefix / area | Role | Home |
|---------------|------|------|
| `Custom_Menu*` domain tools | Stats, XmR, matrices, editing, output, … | Add-in VBE |
| `Fn_*` | Worksheet-callable UDFs | Add-in VBE (keep Public Function names) |
| `aPublicProcedures` / `aPublic*` | Shared library hubs | Already started in the add-in (`modApi*` / `modInternal*`) |
| `Filehandling`, `JPEGS`, `Sampling`, `Maths` | Shared utilities | Add-in, per [ModuleMap.md](ModuleMap.md) |
| UserForms | `.frm` / `.frx` | Add-in VBE |
| `Custom_Menu_Menus`, `ThisWorkbook` | Ribbon / open handlers | Add-in (`modAddinMenu` already ships) |
| `z_*` / WIP | Incomplete or snippet templates | Leave out of the add-in until reviewed |
| `Cipher_*`, niche tools | Domain features | Add-in VBE when wanted |

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
2. Diff several `WriteArrayToWorksheet` copies; pick the best → put it in the add-in (`modInternalSheetIO`).
3. In **one pilot feature module** (e.g. `Custom_Menu11_Histogram`), delete the private copy and `Call` the add-in helper.
4. Compile; repeat for other features.
5. Copy further modules into `ExcelVbaLib.xlam` — do not try to import all 200 modules on day one.

## Regenerating the export

```powershell
# From repo root, with Excel closed or allowing reopen:
# Re-run the export script / ask the agent to re-export Data\Personal123.xlsb
```

Refresh `inventory.csv` whenever Personal changes materially.
