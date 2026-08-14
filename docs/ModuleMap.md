# Suggested module map

Based on the export of `Data/Personal123.xlsb` (see [PersonalInventory.md](PersonalInventory.md)).

**Golden source for shared helpers already in Personal:** module `aPublicProcedures`  
(`WriteArrayToWorksheet`, `SpeedOn`/`SpeedOff`, output-sheet helpers, named ranges, etc.)

Many `Custom_Menu*` modules contain **private copies** of the same helpers (~21× `WriteArrayToWorksheet`, ~18× screen-update toggles). Migration = centralize once, then delete duplicates.

## Target layout

Full routing table (Personal prefix → folder): **[source/README.md](../source/README.md)**.

```text
source/
├── Api/                 ← Public names (modApi*.bas)
├── Internal/            ← Shared plumbing (modInternal*.bas)
├── Features/            ← Custom_Menu* domain packs
│   ├── Stats/           ← Menu5 remainder, Menu11 plots/XmR, Menu18
│   ├── Matrices/        ← Menu13
│   ├── Editing/         ← Menu1, 14, 15, 16
│   ├── Charts/          ← Menu3
│   ├── Output/          ← Menu8
│   ├── Workbook/        ← Menu2
│   ├── TimeSeries/      ← Menu27, 28
│   └── Other/           ← cipher, primes, sampling, …
├── Udf/                 ← Fn_* worksheet functions
├── Forms/               ← UserForms
├── Classes/             ← cls*.cls
├── Menus/               ← Custom_Menu_Menus, ThisWorkbook
├── Sandbox/             ← z_* / WIP
└── _export_raw/         ← unsorted Personal dump
```

### Curated so far (Api + Internal)

```text
source/Api/
├── modApiArrays.bas
├── modApiSheets.bas
├── modApiFiles.bas
├── modApiTables.bas
├── modApiDates.bas       ← CreateDateTable
├── modApiBenford.bas     ← Benford / last-two-digit analyses
└── modApiUi.bas

source/Internal/
├── modInternalSheetIO.bas
├── modInternalExcelApp.bas
├── modInternalDateTable.bas
├── modInternalBenford.bas
├── modInternalNamedRanges.bas
├── modInternalText.bas
└── modInternalError.bas
```

### Phase-2 feature packs (do not import all at once)

Keep as separate modules under `source/Features/<family>/` when landing from Personal. Split to Api + Internal only when the public surface is stable (Benford already is).

| Pack | Status | Destination |
|------|--------|-------------|
| Benford | Extracted | `Api/modApiBenford` + `Internal/modInternalBenford` |
| Stats / quality | Later | `Features/Stats/` — `Custom_Menu11_*`, remaining `Custom_Menu5_*`, XmR |
| Matrices | Later | `Features/Matrices/` — `Custom_Menu13_*`; UDFs → `Udf/` |
| Editing / ranges | Later | `Features/Editing/` |
| Charts | Later | `Features/Charts/` |
| Output | Later | `Features/Output/` |
| Workbook | Later | `Features/Workbook/` |
| Time series | Later | `Features/TimeSeries/` |
| UDFs | Later | `Udf/` — keep `Fn_*` names |
| Forms | Later | `Forms/` |
| Menus | Later | `Menus/` |
| `z_*` / WIP | Exclude | `Sandbox/` |

### Benford public surface

| Public procedure | Module | Output sheet (Personal spelling kept) |
|------------------|--------|----------------------------------------|
| `BenfordAnalysisFirstDigit` | `modApiBenford` | Bedford Analysis First digit |
| `BenfordAnalysisSecondDigit` | `modApiBenford` | Bedford Analysis Second digit |
| `BenfordAnalysisThirdDigit` | `modApiBenford` | Bedford Analysis Third digit |
| `BenfordAnalysisTwoDigit` | `modApiBenford` | Bedford Analysis 2 digit |
| `BenfordAnalysisThreeDigit` | `modApiBenford` | Bedford Analysis 3 digit |
| `BenfordAnalysisLastTwoDigit` | `modApiBenford` | Bedford Analysis last 2 digits |
| `Benford2ndDigitProbability` | `modApiBenford` | (UDF / helper) |
| `Benford3rdDigitProbability` | `modApiBenford` | (UDF / helper) |

Each analysis accepts an optional `Range`; if omitted, an InputBox prompts (same as the Personal menu macros).

## Dependency direction (keep this)

```text
Other workbooks / Personal shims
        │
        ▼
   Api  /  Udf  /  Features  /  Menus
        │
        ▼
   Internal
        │
        ▼
   Excel object model
```

- **Api / Features / Udf / Menus → Internal** ✅  
- **Internal → Api** ❌ (causes cycles)  
- **Api → Api** sparingly (prefer one facade calling Internal)  
- **Sandbox** is never imported until promoted

## What goes where

| If the routine… | Put it in |
|-----------------|-----------|
| Is called from other workbooks or from the ribbon | `Api/*` as `Public` |
| Is a worksheet formula (`Fn_*`) | `Udf/` (keep the Public Function name) |
| Is a `Custom_Menu*` domain tool not yet split | `Features/<family>/` — see [source/README.md](../source/README.md) |
| Is a UserForm | `Forms/` |
| Is a class module | `Classes/` |
| Builds menus / `Workbook_Open` | `Menus/` |
| Is `z_*` / WIP / stock snippets | `Sandbox/` |
| Writes/reads arrays to ranges and is shared by many macros | `Internal/modInternalSheetIO` |
| Turns off screen updating / calculation around a block | `Internal/modInternalExcelApp` |
| Is only used by one Api procedure | `Private` in that same Api module |
| Picks a folder or builds a file list | `Api/modApiFiles` (thin) + Internal if parsing is heavy |
| Creates sheets / clears used ranges for output | `Api/modApiSheets` calling SheetIO |

## Suggested public surface (starter names)

Stabilize these names in the add-in; map old Personal names → new names in a short alias table as you migrate.

### `modApiArrays`
| Public procedure | Responsibility |
|------------------|----------------|
| `WriteArrayToSheet` | Write 1D/2D array starting at a target range/sheet |
| `ReadRangeToArray` | Load a range into a variant array |
| `WriteArrayToNewSheet` | Create sheet + dump array (calls SheetIO + Sheets) |

### `modApiSheets`
| Public procedure | Responsibility |
|------------------|----------------|
| `EnsureSheet` | Get or create worksheet by name |
| `ClearSheetData` | Clear data region without killing the sheet |

### `modApiFiles`
| Public procedure | Responsibility |
|------------------|----------------|
| `PickFolder` | Folder picker → path string |
| `ListFiles` | Extensions filter → path array/collection |

### `modApiTables`
| Public procedure | Responsibility |
|------------------|----------------|
| `RangeToListObject` | Promote range to table |
| `ListObjectToArray` | Table body → array |

### `modApiUi`
| Public procedure | Responsibility |
|------------------|----------------|
| `NotifyInfo` / `NotifyError` | User messaging wrappers |

### `modApiBenford`
| Public procedure | Responsibility |
|------------------|----------------|
| `BenfordAnalysisFirstDigit` | Leading-digit 1–9 analysis sheet + charts |
| `BenfordAnalysisSecondDigit` | Second digit 0–9 |
| `BenfordAnalysisThirdDigit` | Third digit 0–9 |
| `BenfordAnalysisTwoDigit` | First two digits 10–99 |
| `BenfordAnalysisThreeDigit` | First three digits 100–999 |
| `BenfordAnalysisLastTwoDigit` | Last two digits 00–99 (uniform expected) |
| `Benford2ndDigitProbability` / `Benford3rdDigitProbability` | Worksheet UDFs |

## Migration order (preserves Call links)

1. **Split `aPublicProcedures.bas`** from `_export_raw/` into Api + Internal (see mapping below) — this is step one, not importing 200 menu modules.
2. Create empty `.xlam`; import Internal then Api; **compile**.
3. In Personal123, change one pilot menu module to `Call` the shared names (or reference the add-in) and **delete its private** `WriteArrayToWorksheet` / `TurnOffScreeupdates…` copies.
4. Repeat pilots (Histogram → LinearRegression → Benford).
5. Only then batch-import feature modules; UDFs (`Fn_*`) can stay in Personal longer if worksheet formulas depend on them.
6. Leave `z_*` / `z_WIP` out of the add-in until reviewed.

### `aPublicProcedures` → scaffold mapping

| Procedure in Personal | Destination |
|----------------------|-------------|
| `SpeedOn` / `SpeedOff` / `TurnOffScreeupdatesAndCalculation` / `TurnOn…` | `modInternalExcelApp` (+ thin Public wrappers in Api if menus need them) |
| `WriteArrayToWorksheet` / `WriteArrayToWorksheetA1` / `WriteRangeToWorksheet` | `modInternalSheetIO` + `modApiArrays` |
| `CreateOutputSheet` / `DeleteOutputSheet` / `CheckExistenceAndDeleteOutputSheet` / `WorksheetExists` | `modApiSheets` |
| `CreateNamedRangeForSheet` / `ValidRangeName` / … | `modInternalNamedRanges` |
| `CreateFolder` / `DeleteFolder` | `modApiFiles` |
| `StatusUpdate` / `StatusClear` | `modApiUi` |

## Example call chain (same project — links stay local)

```vb
' Api/modApiArrays.bas
Public Sub WriteArrayToSheet(ByVal TargetSheet As String, ByVal StartCell As String, ByRef Data As Variant)
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    Call modInternalSheetIO.DumpArray(TargetSheet, StartCell, Data)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("WriteArrayToSheet")
End Sub
```

```vb
' Internal/modInternalSheetIO.bas
Public Sub DumpArray(ByVal TargetSheet As String, ByVal StartCell As String, ByRef Data As Variant)
    ' shared implementation used by several Api routines
End Sub
```

> Note: Internal helpers used from other modules in the same project must be `Public` (or in a standard module without `Private`). Prefer `Public` only inside the add-in; do not document them as part of the external API. Optionally prefix Internal procedure names with `z` / `Int` if you want them sorted away from Api.

## Optional link to Power Query library

| Related PQ library | VBA touchpoint |
|------------------|----------------|
| [PowerQuery-Library](https://github.com/flyingcount/PowerQuery-Library) `src/Files/` | `modApiFiles` + sheet dump of results |
| PQ Cleanse\* / profiling | Rarely needed in VBA; if macros orchestrate PQ, keep orchestration in `modApiUi` / a future `modApiPowerQuery` |
| Output of any macro | Always via `modInternalSheetIO` |

## After Trust access is enabled

Re-run an export of Personal into `source/_export_raw/`, then fill this checklist:

- [ ] List every `Public Sub/Function` and assign an Api module  
- [ ] List shared helpers → Internal  
- [ ] Delete dead code / duplicates  
- [ ] Update this document with the **actual** procedure inventory  
