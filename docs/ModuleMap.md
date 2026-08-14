# Suggested module map

Based on the export of `Data/Personal123.xlsb` (see [PersonalInventory.md](PersonalInventory.md)).

**Golden source for shared helpers already in Personal:** module `aPublicProcedures`  
(`WriteArrayToWorksheet`, `SpeedOn`/`SpeedOff`, output-sheet helpers, named ranges, etc.)

Many `Custom_Menu*` modules contain **private copies** of the same helpers (~21× `WriteArrayToWorksheet`, ~18× screen-update toggles). Migration = centralize once, then delete duplicates.

## Library

**The library is `ExcelVbaLib.xlam`.** Add further Personal123 modules in the add-in VBE. `source/` is only a snapshot of what is already in the add-in (see [source/README.md](../source/README.md)).

### Currently in the add-in

```text
source/Api/
├── modApiArrays.bas
├── modApiSheets.bas
├── modApiFiles.bas
├── modApiTables.bas
├── modApiDates.bas       ← CreateDateTable
├── modApiBenford.bas     ← Benford / last-two-digit analyses
├── modApiWorksheetTemplates.bas ← Links, Actions, Force Field, Assumptions, Questions, Notes workbook
├── modApiSampling.bas    ← ExtractSample (Menu4)
└── modApiUi.bas

source/Internal/
├── modInternalSheetIO.bas
├── modInternalExcelApp.bas
├── modInternalDateTable.bas
├── modInternalBenford.bas
├── modInternalNamedRanges.bas
├── modInternalText.bas
├── modInternalError.bas
├── modInternalWorksheetTemplates.bas
└── modInternalSampling.bas

source/Menus/
├── modAddinMenu.bas
└── ThisWorkbook.cls      ← Workbook_Open / BeforeClose (injected, not imported)
```

### Later Personal packs

Copy the next family into **ExcelVbaLib.xlam** (same project so `Call` stays in-process). Do not land it in git folders first. Leave `z_*` / WIP out until reviewed.

| Pack | Status |
|------|--------|
| Benford | In the add-in (`modApiBenford` + `modInternalBenford`) |
| Worksheet templates | In the add-in (`modApiWorksheetTemplates`) — Personal `Custom_Menu26_wkshtTmplt` |
| Sampling | In the add-in (`modApiSampling`) — Personal `Custom_Menu4_Sample` / `ExtractSample` |
| Stats / quality | Later — `Custom_Menu11_*`, remaining `Custom_Menu5_*`, XmR |
| Matrices | Later — `Custom_Menu13_*`, `Fn_Matrices*` |
| Editing / ranges | Later |
| Charts | Later |
| Output | Later |
| Workbook | Later |
| Time series | Later |
| UDFs | Later — keep `Fn_*` names |
| Forms | Later |
| Menus | Add-in menu is `modAddinMenu`; Personal `Custom_Menu_Menus` later |

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

### Worksheet templates public surface

| Public procedure | Module | Sheet / table |
|------------------|--------|----------------|
| `CreateLinksTemplate` | `modApiWorksheetTemplates` | Links / `Tbl_Links` |
| `CreateActionsTemplate` | `modApiWorksheetTemplates` | Actions / `Tbl_Actions` |
| `CreateForceFieldTemplate` | `modApiWorksheetTemplates` | Force Field / `Tbl_ForceField` |
| `CreateAssumptionsTemplate` | `modApiWorksheetTemplates` | Assumptions / `Tbl_Assumptions` |
| `CreateQuestionsTemplate` | `modApiWorksheetTemplates` | Questions / `Tbl_Questions` |
| `CreateNotesWorkbook` | `modApiWorksheetTemplates` | New workbook + Index |
| `ImportPythonPackages` | `modApiWorksheetTemplates` | Python packages |
| `HyperLinkText` | `modApiWorksheetTemplates` | UDF used on the Links sheet |

### Sampling public surface

| Public procedure | Module | Sheet |
|------------------|--------|--------|
| `ExtractSample` | `modApiSampling` | Sample |

Random sample of input rows (percent of row count), with or without replacement.

## Dependency direction (keep this)

```text
Other workbooks / Personal shims
        │
        ▼
   ExcelVbaLib.xlam (public Subs / UDFs / menu)
        │
        ▼
   Internal helpers in the same add-in
        │
        ▼
   Excel object model
```

- Public entry points and helpers live in **the same `.xlam`** ✅  
- Internal helpers must not call public API modules (avoids cycles)  
- Do not import add-in modules into caller workbooks

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

1. **Split `aPublicProcedures.bas`** from `_export_raw/` into helpers already in the add-in (see mapping below) — this is step one, not importing 200 menu modules.
2. Load `ExcelVbaLib.xlam`; **compile**.
3. In Personal123, change one pilot menu module to `Call` / `Application.Run` the add-in names and **delete its private** `WriteArrayToWorksheet` / `TurnOffScreeupdates…` copies.
4. Repeat pilots (Histogram → LinearRegression). Benford is already in the add-in.
5. Copy further feature modules into the **add-in** project; UDFs (`Fn_*`) can stay in Personal longer if worksheet formulas depend on them.
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
