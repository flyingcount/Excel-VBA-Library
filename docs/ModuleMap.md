# Suggested module map

Based on the export of `Data/Personal123.xlsb` (see [PersonalInventory.md](PersonalInventory.md)).

**Golden source for shared helpers already in Personal:** module `aPublicProcedures`  
(`WriteArrayToWorksheet`, `SpeedOn`/`SpeedOff`, output-sheet helpers, named ranges, etc.)

Many `Custom_Menu*` modules contain **private copies** of the same helpers (~21× `WriteArrayToWorksheet`, ~18× screen-update toggles). Migration = centralize once, then delete duplicates.

## Target layout

```text
source/
├── Api/
│   ├── modApiArrays.bas      ← from aPublicProcedures: WriteArray*, WriteRange*
│   ├── modApiSheets.bas      ← CreateOutputSheet, WorksheetExists, UserSelectWorksheet…
│   ├── modApiFiles.bas       ← CreateFolder/DeleteFolder + Filehandling
│   ├── modApiTables.bas      ← table/list migrations (Menu1/Menu19…)
│   └── modApiUi.bas          ← StatusUpdate/StatusClear, Notify*
└── Internal/
    ├── modInternalSheetIO.bas    ← WriteArrayToWorksheet implementation (single copy)
    ├── modInternalExcelApp.bas   ← SpeedOn/Off, TurnOff/OnScreeupdatesAndCalculation
    ├── modInternalNamedRanges.bas← CheckNamedRangeExists, CreateNamedRange* (dedupe later)
    ├── modInternalText.bas       ← string/null helpers
    └── modInternalError.bas      ← error helpers
```

### Phase-2 feature packs (do not import all at once)

Keep as separate modules under `source/Features/` when ready (names can stay `Custom_Menu…` initially):

| Pack | Examples from export |
|------|----------------------|
| Stats / quality | `Custom_Menu5_*` Benford, `Custom_Menu11_*` plots, XmR |
| Matrices | `Custom_Menu13_*`, `Fn_Matrices*` |
| Editing / ranges | `Custom_Menu1_Editing`, `Custom_Menu14_*` |
| Output | `Custom_Menu8_Output*` |
| UDFs | `Fn_*` (keep Public Function names for worksheet formulas) |
| Menus | `Custom_Menu_Menus`, `ThisWorkbook` open handlers |

## Dependency direction (keep this)

```text
Other workbooks / Personal shims
        │
        ▼
   Api (Public only)
        │
        ▼
   Internal (helpers)
        │
        ▼
   Excel object model
```

- **Api → Internal** ✅  
- **Internal → Api** ❌ (causes cycles)  
- **Api → Api** sparingly (prefer one facade calling Internal)

## What goes where

| If the routine… | Put it in |
|-----------------|-----------|
| Is called from other workbooks or from the ribbon | `Api/*` as `Public` |
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
| `CreateNamedRangeForSheet` / `ValidRangeName` / … | `modInternalNamedRanges` (add when you migrate) |
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
