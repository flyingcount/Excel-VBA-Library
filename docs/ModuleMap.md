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
├── modApiData.bas        ← random fills, combinations, distributions (Menu6)
├── modApiHyperlinks.bas  ← hyperlink inventory / index / follow (Menu21)
├── Custom_Menu13_CreateMatrices.bas
├── Custom_Menu13_Matrices1.bas
├── Custom_Menu13_Matrices2.bas
├── Custom_Menu13_Cholesky.bas
├── Custom_Menu13_EigenDecomp.bas
├── custom_Menu13_Unitary.bas
├── Custom_Menu13_MatrixUtilities.bas
├── Fn_MatricesArray.bas / Fn_MatricesRng.bas / Fn_Matrices2.bas
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
├── modInternalSampling.bas
├── modInternalData.bas
├── modInternalHyperlinks.bas
└── modInternalMatrices.bas

source/Menus/
├── modAddinMenu.bas
└── ThisWorkbook.cls      ← Workbook_Open / AddinInstall / BeforeClose (injected, not imported)
```

### Later Personal packs

Copy the next family into **ExcelVbaLib.xlam** (same project so `Call` stays in-process). Do not land it in git folders first. Leave `z_*` / WIP out until reviewed.

| Pack | Status |
|------|--------|
| Benford | In the add-in (`modApiBenford` + `modInternalBenford`) |
| Worksheet templates | In the add-in (`modApiWorksheetTemplates`) — Personal `Custom_Menu26_wkshtTmplt` |
| Sampling | In the add-in (`modApiSampling`) — Personal `Custom_Menu4_Sample` / `ExtractSample` |
| Data | In the add-in (`modApiData`) — Personal `Custom_Menu6_Data` / `RndFrmRng` / `RndProbDist` |
| Hyperlinks | In the add-in (`modApiHyperlinks`) — Personal Menu21 |
| Matrices | In the add-in (`Custom_Menu13_Matrices1` / `Custom_Menu13_Matrices2` / rest of Menu13) — Personal names kept. Overlay originals with `Import-Menu13FromPersonal.ps1` |
| Stats / quality | Later — `Custom_Menu11_*`, remaining `Custom_Menu5_*`, XmR |
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

### Data public surface

| Public procedure | Module | Notes |
|------------------|--------|--------|
| `DataCombinations` | `modApiData` | Cartesian product of selected columns; writes to the right |
| `RandomIntegers` / `RandomNumbers` / `RandomDates` / `RandomStrings` | `modApiData` | Fill the current selection |
| `RandomFromList` | `modApiData` | Destination then source list (InputBox ranges) |
| `RandomTrueFalse` / `RandomYesNo` / `Random1or0` | `modApiData` | Fill the current selection |
| `CreateYesNoDataset` | `modApiData` | Sheet **Yes No Dataset** (Predicted / Actual) |
| `RandomTestDataTypes` | `modApiData` | Two-column type/value block |
| `RandomBinomialNumbers` / `RandomBernoulliNumbers` / `RandomNormalNumbers` / `RandomPoissonNumbers` / `RandomExponentialNumbers` / `RandomGammaNumbers` / `RandomHypergeometricNumbers` | `modApiData` | **Excel VBA Lib → Data → Probability distributions** |

Personal Menu6 **Prime numbers** (`PrimeGenerator` in `Custom_Menu10_Prime`) is not in this pack.

### Hyperlinks public surface

Personal Menu21 (`Custom_Menu21_Hyperlinks`, `Custom_Menu21_Index`). Menu **Excel VBA Lib → Hyperlinks**. Personal OnAction names are kept.

| Public procedure | Module | Notes |
|------------------|--------|--------|
| `HyperlinkInventory` | `modApiHyperlinks` | Sheet **Hyperlinks**: location, text, ScreenTip, address, SubAddress, type |
| `CreateIndex` | `modApiHyperlinks` | Sheet **Index**; inserts a Back to Index row on unprotected sheets |
| `UpdateIndex` | `modApiHyperlinks` | Rebuilds Index in place; does not insert rows |
| `ShowAllWorksheetsInWorkbook` | `modApiHyperlinks` | Message box with hidden / very hidden flags |
| `RemovingHyperLink` | `modApiHyperlinks` | Deletes workbook links whose Text to display matches the selected cells |
| `OpenHyperlink` | `modApiHyperlinks` | Follows hyperlinks in the selection (Personal showed RangeForm instead) |
| `AddHyperlinksToCurrentSheetA1` | `modApiHyperlinks` | Back-link in the active cell on every other unprotected sheet |

`HyperLinkText` (UDF) already lives on `modApiWorksheetTemplates`. Notes workbook Index (`BuildIndexSheet`) does not insert rows; **Create index** does.

### Matrices public surface

Personal modules (same names in the add-in VBE): `Custom_Menu13_CreateMatrices`, `Custom_Menu13_Matrices1`, `Custom_Menu13_Matrices2`, `Custom_Menu13_Cholesky`, `Custom_Menu13_EigenDecomp`, `custom_Menu13_Unitary`, `Custom_Menu13_MatrixUtilities`, `Fn_MatricesArray`, `Fn_MatricesRng`, `Fn_Matrices2`.

Git snapshot math is array-based in `modInternalMatrices` (Personal dumps are gitignored). After `Import-AddinModules.ps1 -All`, overlay the original Personal modules into the `.xlam` with `scripts/Import-Menu13FromPersonal.ps1` (Windows, Personal.xlsb present). `-AllMenu13` copies the whole family. Do that **after** `-All`; running `-All` again replaces them with the git snapshot.

`git pull` does not update `build\ExcelVbaLib.xlam`. Add-in macros never appear in Alt+F8; worksheet `Mat*` names are registered into Insert Function (category **Excel VBA Lib**) by `RegisterMatrixUdfs` (`Fn_MatricesRng`) when the menu installs.

| Public procedure | Module | Notes |
|------------------|--------|--------|
| `MatrixCreateIdentity` / `Zeros` / `Ones` / `Diagonal` / `Random` / `Hilbert` / `Exchange` / `Toeplitz` / `Vandermonde` / `Companion` | `Custom_Menu13_CreateMatrices` | Create writes at the active cell; vector-based create writes to the right of the selection |
| `MatrixTranspose` / `Add` / `Subtract` / `Scale` / `Multiply` / `MultiplicationHadamard` / `MultiplicationKronecker` / `Outer` / `Dot` / `Inverse` / `Power` / `Determinant` / `Trace` / `DiagExtract` / `Vec` / `Unvec` | `Custom_Menu13_Matrices1` | Result one column to the right of the selection (Hadamard prompts; default is to the right of B) |
| `MatrixHadamardProof` | `Custom_Menu13_MatrixUtilities` | Sheet **Hadamard Proof**: H, Hᵀ, H.HT, then n and I or a not-Hadamard message |
| `MatrixSolve` / `Rank` / `Norm` / `Norm1` / `NormInf` / `IsSymmetric` / `Adjugate` / `PseudoInverse` / `LU` / `Cofactor` / `Minor` | `Custom_Menu13_Matrices2` | Cofactor limited to order 20 |
| `MatrixCholesky` | `Custom_Menu13_Cholesky` | SPD only |
| `MatrixEigen` | `Custom_Menu13_EigenDecomp` | Jacobi; symmetric only. Writes vectors then a λ column |
| `MatrixQR` / `MatrixIsOrthogonal` | `custom_Menu13_Unitary` | QR stacks R below Q. Complex unitary from Personal is not included |
| `Mat*` / `MatrixMultDefined` | `Fn_MatricesArray` / `Fn_MatricesRng` / `Fn_Matrices2` | Worksheet UDFs (`RegisterMatrixUdfs`) |

Numeric arrays only; max side 250. Inverse uses Gauss-Jordan; determinant and rank use LU; eigen is Jacobi (symmetric); QR is modified Gram-Schmidt; pseudoinverse uses normal equations (needs full column or row rank). `vec`/`unvec` are column-major.

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
