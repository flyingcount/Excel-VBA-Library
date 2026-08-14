# source/

Curated VBA for `ExcelVbaLib.xlam`. Raw dumps from Personal123 stay in `_export_raw/` (gitignored) until they are re-homed here.

**One add-in project holds the full call graph.** Api calls Internal in the same `.xlam`. Do not split a call chain across workbooks.

## Folders

| Folder | What belongs here | Personal123 examples |
|--------|-------------------|----------------------|
| [Api/](Api/) | Public `Sub`/`Function` names other workbooks and the ribbon call | `aPublicProcedures` entry points, `CreateDateTable`, Benford analyses |
| [Internal/](Internal/) | Shared plumbing used by Api/Features; not the external API | `SpeedOn`/`WriteArrayToWorksheet`, named ranges, date-table build |
| [Features/](Features/) | Domain packs from `Custom_Menu*` (may stay one module or later split to Api+Internal) | Histogram, XmR, matrices, editing tools |
| [Udf/](Udf/) | Worksheet-callable functions; **keep the `Fn_*` Public names** | `Fn_Ageing`, `Fn_TaxUK`, `Fn_Zscore`, `Fn_Matrices*` |
| [Forms/](Forms/) | UserForms (`.frm` + `.frx`) | `BoxCoxForm`, `DataCleansingOptionsForm` |
| [Classes/](Classes/) | Class modules (`cls*.cls`) | (none exported yet) |
| [Menus/](Menus/) | Ribbon / command-bar wiring and add-in `ThisWorkbook` handlers | `Custom_Menu_Menus`, `ThisWorkbook` |
| [Sandbox/](Sandbox/) | Experiments and snippet templates — **do not import to the `.xlam` until reviewed** | `z_*`, `z_WIP`, `zStockCode*` |
| [_export_raw/](_export_raw/) | Unsorted dump from `Data/Personal123.xlsb` | all 200+ exported components |

## Where to put the next Personal module

1. Is it a **worksheet UDF** (`Fn_*`)? → `Udf/`
2. Is it a **UserForm**? → `Forms/`
3. Is it a **class**? → `Classes/`
4. Is it **menu/ribbon** wiring? → `Menus/`
5. Is it `z_*` / WIP / stock snippets? → `Sandbox/`
6. Is it **generic Excel plumbing** (speed, sheet I/O, named ranges, folders)? → `Internal/` (+ thin wrapper in `Api/` if other workbooks call it)
7. Is it a **domain tool** (plot, matrix, edit, forecast, cipher, …)? → `Features/<family>/` (see [Features/README.md](Features/README.md))
8. When a Feature is stable, optionally split like Benford: public names → `Api/modApi…`, engine → `Internal/modInternal…`

## Naming

| Kind | File / `Attribute VB_Name` |
|------|----------------------------|
| Public API | `modApi{Area}.bas` |
| Internal helper | `modInternal{Area}.bas` |
| Feature (pre-split) | `mod{Area}.bas` (Personal `Custom_Menu*` names OK while landing) |
| UDF module | keep `Fn_{Area}.bas` so formula names stay stable |
| Class | `cls{Name}.cls` |
| Form | `{FormName}.frm` |

`Attribute VB_Name` **must** match the file stem.

## Import order into the `.xlam`

Run `scripts/Build-ExcelVbaLib.ps1` (it uses this order). Manual import:

1. `Internal/`
2. `Api/`
3. `Udf/` / `Features/` / `Forms/` / `Classes/` / `Menus/` as needed  
Never import `Sandbox/` or `_export_raw/` wholesale.
