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
├── modApiFiles.bas       ← folder file listings (Menu24)
├── modApiTables.bas      ← table inventory (Menu19)
├── modApiDates.bas       ← CreateDateTable
├── modApiBenford.bas     ← Benford / last-two-digit analyses
├── modApiWorksheetTemplates.bas ← Links, Actions, Force Field, Assumptions, Questions, Notes workbook
├── modApiSampling.bas    ← ExtractSample (Menu4)
├── modApiData.bas        ← random fills, combinations, distributions (Menu6)
├── modApiHyperlinks.bas  ← hyperlink inventory / index / follow (Menu21)
├── modApiCustomLists.bas ← custom lists / AutoCorrect (Menu16)
├── modApiRanges.bas      ← named ranges, analysis, cleanse (Menu14)
├── modApiDuplicates.bas  ← flag / colour / count duplicates (Menu15)
├── modApiProtection.bas  ← scroll area, sheet protect, unhide (Menu7)
├── modApiReconciliations.bas ← two-column recon / range compare (Menu20)
├── modApiPowerQuery.bas  ← import/export form, background refresh, Fast Combine, connect all tables (Menu29)
├── frmPQLibrary.frm      ← Import or export queries and functions (Menu29)
├── frmFileListingProgress.frm ← modeless Cancel during Files listing
├── modApiMatrixCreate.bas
├── modApiMatrices1.bas
├── modApiMatrices2.bas
├── modApiCholesky.bas
├── modApiEigenDecomp.bas
├── modApiUnitary.bas
├── modApiMatrixUtilities.bas
├── modApiCovariance.bas
├── modApiAnalysis.bas
├── modApiConfusion.bas
├── modApiResiduals.bas
├── modApiSvd.bas
├── modApiLinearSystem.bas
├── modApiHistogram.bas
├── modApiDistPlots.bas
├── modApiQQPlots.bas
├── modApiLinearRegression.bas
├── modApiLorenz.bas
├── modApiAcf.bas
├── modApiTimeSeries.bas ← analysis / lag differencing (Menu27)
├── modApiDiebold.bas
├── modApiXmR.bas
├── modApiProcessCapability.bas
├── modApiChartSheet.bas
├── Fn_MatricesArray.bas / Fn_MatricesRng.bas / Fn_Matrices2.bas
├── Fn_TimeSeries.bas     ← ACF / ACVF / PACF / Bartlett / Box-Pierce / Ljung-Box UDFs (Menu27)
└── modApiUi.bas

source/Internal/
├── modInternalSheetIO.bas
├── modInternalExcelApp.bas
├── modInternalDateTable.bas
├── modInternalBenford.bas
├── modInternalNamedRanges.bas
├── modInternalRanges.bas
├── modInternalDuplicates.bas
├── modInternalProtection.bas
├── modInternalReconciliations.bas
├── modInternalText.bas
├── modInternalError.bas
├── modInternalWorksheetTemplates.bas
├── modInternalSampling.bas
├── modInternalData.bas
├── modInternalHyperlinks.bas
├── modInternalCustomLists.bas
├── modInternalTables.bas
├── modInternalFiles.bas
├── modInternalPowerQuery.bas
├── modInternalMatrices.bas
├── modInternalAnalysis.bas
├── modInternalConfusion.bas
├── modInternalSvd.bas
├── modInternalTimeSeries.bas
└── modInternalPlots.bas

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
| Custom lists | In the add-in (`modApiCustomLists`) — Personal `Custom_Menu16_CustomLists` / `Custom_Menu16_Autocorrect` |
| Ranges | In the add-in (`modApiRanges` + `modInternalRanges`) — Personal `Custom_Menu14_*` |
| Duplicates | In the add-in (`modApiDuplicates` + `modInternalDuplicates`) — Personal `Custom_Menu15_Duplicates` |
| Protection | In the add-in (`modApiProtection` + `modInternalProtection`) — Personal `Custom_Menu7_Protection` |
| Reconciliations | In the add-in (`modApiReconciliations` + `modInternalReconciliations`) — Personal `Custom_Menu20_*` |
| Tables | In the add-in (`modApiTables`) — Personal `Custom_Menu19_Tables` |
| Files | In the add-in (`modApiFiles`) — Personal `Custom_Menu24_ListFilesInFolder` |
| Power Query | In the add-in (`modApiPowerQuery` + `modInternalPowerQuery` + `frmPQLibrary`) — Personal Menu29 |
| Analysis | In the add-in (`modApiAnalysis` / `modApiCovariance` / `modApiConfusion` / `modApiResiduals` / `modApiSvd` / `modApiLinearSystem`) — Personal Menu18 |
| Plots Charts | In the add-in (`modApiHistogram` / `modApiDistPlots` / `modApiQQPlots` / `modApiLinearRegression` / `modApiLorenz` / `modApiAcf` / `modApiDiebold` / `modApiXmR` / `modApiProcessCapability` / `modApiChartSheet` + `modInternalPlots`) — Personal `Custom_Menu11_*` |
| Time series | In the add-in (`modApiTimeSeries` + `modInternalTimeSeries` + `Fn_TimeSeries`) — Personal `Custom_Menu27_*` |
| Matrices | In the add-in (`modApiMatrices1` / `modApiMatrices2` / rest of Menu13 as `modApi*`). Overlay Personal originals with `Import-Menu13FromPersonal.ps1` (rewrites names to `modApi*`) |
| Stats / quality | Later — remaining `Custom_Menu5_*` |
| Charts | Later — Personal Menu3 arrange/list charts (Menu11 plots are on **Plots Charts**) |
| Output | Later |
| Workbook | Later |
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

### Custom lists public surface

Personal Menu16 (`Custom_Menu16_CustomLists`, `Custom_Menu16_Autocorrect`). Menu **Excel VBA Lib → Custom lists**. Personal OnAction names are kept.

Custom lists and AutoCorrect replacements are **application-wide** (Excel Options), not stored in the workbook.

| Public procedure | Module | Notes |
|------------------|--------|--------|
| `CountCustomLists` | `modApiCustomLists` | Message box: total, built-in (1–4), user-defined |
| `ShowCustomLists` | `modApiCustomLists` | Sheet **Custom List properties**: number, Type (Built-in/User), item count, elements |
| `CreateCustomListByColumn` | `modApiCustomLists` | One list per column; blanks skipped; existing lists skipped; ≥2 values required |
| `CreateCustomListByRow` | `modApiCustomLists` | One list per row (same rules) |
| `DeleteCustomList` | `modApiCustomLists` | List numbers from a range (column A of the inventory). Built-in 1–4 cannot be deleted. Highest numbers first so remaining numbers stay valid. Confirms because the change is Excel-wide |
| `NumCustomLists` | `modApiCustomLists` | Volatile UDF: `Application.CustomListCount` |
| `AutoCorrectEntries_Display` | `modApiCustomLists` | Sheet **Auto correct List** (Personal spelling kept) |
| `AutoCorrectEntries_Add` | `modApiCustomLists` | Two-column range (Replace / With). Adds or overwrites; does not remove other entries. Uses the selected range (Personal scanned column A of the sheet) |

### Ranges public surface

Personal Menu14 (`Custom_Menu14_ColRange`, `Custom_Menu14_FreqAnalysis`, `Custom_Menu14_ManipulateRanges`, `Custom_Menu14_RangeAnalysis`, `Custom_Menu14_SplitRange`). Menu **Excel VBA Lib → Ranges**. Personal OnAction names are kept, including `ConvertNamedRangeGlodalToLocalScope`.

UserForms from that menu (`DataCleansingOptionsForm`, `DataValidationForm`, `MultiSearchAndReplace`) are not included.

| Public procedure | Module | Notes |
|------------------|--------|--------|
| `ListNamedRangeProperties` | `modApiRanges` | Sheet **Range properties**: index, name, RefersTo, address, visible, comment, workbook parameter, scope. Sequential index (Personal used `Name.Index`, which is not a Name property). No `Range.ListNames` dump. |
| `PasteRangeAsPicture` | `modApiRanges` | Pictures.Paste of the selection |
| `HighlightRanges` / `DeHighlightRanges` | `modApiRanges` | Yes = workbook, No = active sheet. Fill ColorIndex 36; clear uses `xlColorIndexNone` (Personal used 0). Constants without `RefersToRange` are skipped |
| `DeleteNamedRanges` | `modApiRanges` | Names listed in the selection. Confirms. Then rewrites **Range properties** |
| `HiddenMasterDataWorkbookScope` / `HiddenMasterDataWorksheetScope` | `modApiRanges` | Two columns: name, value. Hidden constants. Worksheet version prompts for a cell on the owner sheet. `ValidRangeName` sanitises names |
| `NamedRangeIntoNamedRangeColumns` | `modApiRanges` | Hidden sheet-scoped name per column from the header row (header excluded). Header-only / empty header / one cell rejected |
| `CreateHiddenNamedRange` / `CreateHiddenNamedString` / `CreateHiddenNamedNumber` | `modApiRanges` | Hidden workbook names. Names need not be letters-only |
| `NameCreateConstantsAsNamedRanges` | `modApiRanges` | `Divisor_Thousand` = 1000, `Divisor_Millions` = 1000000 |
| `CreateNamedRangeInAllWorksheets` | `modApiRanges` | Same local name on every sheet, same address **on that sheet** (Personal pointed every sheet at the original selection) |
| `HideAllNamedRanges` / `UnhideAllNamedRanges` | `modApiRanges` | Workbook `Names` collection |
| `HideNamedRangesWorksheet` / `UnhideNamedRangesWorksheet` | `modApiRanges` | Sheet-scoped names on the active sheet |
| `HideSpecifiedNamedRanges` / `UnhideSpecifiedNamedRanges` | `modApiRanges` | Names listed in the selection; then **Range properties** |
| `ListUniqueValues` | `modApiRanges` | Sheet **Unique Values**. `RemoveDuplicates` uses every selected column (Personal passed only the last column index) |
| `ConvertNamedRangeLocalToGlobalScope` / `ConvertNamedRangeGlodalToLocalScope` | `modApiRanges` | Keeps RefersTo / Visible / Comment. Local→global strips the sheet prefix. Global→local prompts for the owner sheet |
| `RangeAnalysis` / `RangeAnalysisMessage` | `modApiRanges` | Sheet **Range_Analysis** or a message box. Blanks are not numeric. Even/Odd count integer numbers (Personal’s message box wrote `n/a`) |
| `FrequencyAnalysis` | `modApiRanges` | Sheet **Frequency Analysis**. One row per used character, O(n), empty range rejected (Personal wrote ASCII 1–255 with an n×255 loop) |
| `Remove_AlphaCharactersFromString` / `Remove_NumbersFromString` / `Remove_SpecialCharactersFromString` | `modApiRanges` | Keep digits / letters / letters+digits+space+period |
| `ClearCellsThatOnlyContainWhitespaces` / `RemoveWhiteSpaces` | `modApiRanges` | Trim NBSP/tab/breaks. Regex version is the menu target (Personal OnAction was the Function `RemoveWhiteSpacesFn`) |
| `ListCharactersAndCodesInRange` | `modApiRanges` | Sheet **Characters and Codes**, sorted by code |
| `TransposeARange` | `modApiRanges` | Copy transposed one column to the right of the selection; sheet-qualified `Cells` |
| `SplitRangeAndNameEachColumn` | `modApiRanges` | Workbook names `BaseAll` and `Base_1..n`. Header optional. Single-column works. Cancel on the header question creates nothing (Personal called `TurnOff` on cancel) |

`ValidRangeName` lives on `modInternalNamedRanges` (spaces to `.`, illegal characters stripped, leading digit or A1-style address prefixed). Entire-column selections are intersected with UsedRange. InputBox cancel returns without `End`.

### Duplicates public surface

Personal Menu15 (`Custom_Menu15_Duplicates`). Menu **Excel VBA Lib → Duplicates**. Personal OnAction names are kept (`ColorDuplicates` vs `ColourDuplicateValues*`).

Blank and error cells are skipped. Entire-column selections use UsedRange. Flag/reference tools insert cells to the right of the selected rows only (Personal inserted the whole worksheet column).

| Public procedure | Module | Notes |
|------------------|--------|--------|
| `FlagDuplicates` | `modApiDuplicates` | Single column. Write `Duplicate` beside later occurrences; first occurrence stays blank |
| `ReferenceDuplicates` | `modApiDuplicates` | Single column. Letter `A`, `B`, … on every member of a duplicate group (including the first). Unique values stay blank |
| `ColorDuplicates` | `modApiDuplicates` | Distinct ColorIndex per duplicate group, cycling a palette. Unique cells unfilled. Personal incremented ColorIndex per duplicate cell and could raise “Too many duplicate!” |
| `ColourDuplicateValuesByRow` | `modApiDuplicates` | Red (3) when a value appears more than once in that row. Need ≥2 columns. Clears fill with `xlColorIndexNone` (Personal used 0) |
| `ColourDuplicateValuesByColumn` | `modApiDuplicates` | Green (4) when a value appears more than once in that column. Need ≥2 rows |
| `ColourDuplicateValuesInSelection` | `modApiDuplicates` | Yellow (6) when a value appears more than once in the selection |
| `DuplicateCountFromSelection` | `modApiDuplicates` | Message box: count of cells whose value appears more than once |

### Protection public surface

Personal Menu7 (`Custom_Menu7_Protection`). Menu **Excel VBA Lib → Protection**. Personal OnAction names are kept.

The default password is Personal's `WYSIWYG`: a known convenience value (`DisplayPassword` shows it), not a secret.

| Public procedure | Module | Notes |
|------------------|--------|--------|
| `LimitScrollArea` | `modApiProtection` | `ScrollArea` = first area of the selection. Requires a worksheet and a range on that sheet |
| `ResetScrollArea` | `modApiProtection` | Clears `ScrollArea` |
| `ProtectWorksheet` | `modApiProtection` | Default password, `UserInterfaceOnly:=True` (VBA can edit this session only). Unprotects with the default password first when that works |
| `UnProtectWorksheet` | `modApiProtection` | Default password. Different password → message, not a raw VBA error |
| `DisplayPassword` | `modApiProtection` | Message box with the default password |
| `UnHideAllSheets` | `modApiProtection` | Hidden and very-hidden worksheets become visible; reports the count. Chart sheets unchanged (Worksheets collection) |
| `UnhideAllRowsAndColumns` | `modApiProtection` | Unhides every row and column without `Cells.Select`. Protected sheet → message |

Chart sheets are rejected. Personal used unqualified `Cells.Select` for unhide-rows/columns.

### Reconciliations public surface

Personal Menu20 (`Custom_Menu20_Reconciliation` / `Custom_Menu20_RecnStrings` / `Custom_Menu20_CompareRanges`). Menu **Excel VBA Lib → Reconciliations**. Personal OnAction names are kept.

`CompareTwoRangesForm` is not included; **Compare two ranges** uses InputBoxes.

| Public procedure | Module | Notes |
|------------------|--------|--------|
| `ReconcileTwoColumns` | `modApiReconciliations` | Two numeric columns → sheet **Reconciliation**. 1-for-1 match (Dictionary queues). Statement in A; first dataset in H; second in O. Non-numeric cells skipped for sums. Cancel before `.Columns.Count` |
| `ReconcileTwoColumnsStrings` | `modApiReconciliations` | Two text columns → sheet **Reconciliation Strings**. Binary match, no trim. Unmatched listed as not in the other set |
| `CompareRanges` | `modApiReconciliations` | Same-size ranges, every cell (Personal skipped the first). Yellow fill on diffs; sheet **Range Comparison**; table `Comparison` |

### Tables public surface

Personal Menu19 (`Custom_Menu19_Tables`). Menu **Excel VBA Lib → Tables**. Personal OnAction names are kept.

Personal `CreateCalendar` is not on this menu; use `CreateDateTable` for a calendar dimension.

| Public procedure | Module | Notes |
|------------------|--------|--------|
| `ListTableProperties` | `modApiTables` | Sheet **Table properties**: index, name, sheet, range, hyperlink, rows/columns, filter/header/totals, style, source type, comment, alt text, header range. Sized to the table count (Personal wrote a 1000-row empty block). `ListObject.Active` / Creator / Application are dropped (not useful; `.Active` is not a ListObject property). |
| `ShowAllTablesInWorkbook` | `modApiTables` | Message box: index, name, sheet (hidden flag), address. Personal menu called this but the Sub was commented out; title was `vbOKOnly`. |
| `RangeToListObject` | `modApiTables` | Promote a range to a ListObject |
| `ListObjectToArray` | `modApiTables` | Table body → array |

### Files public surface

Personal Menu24 (`Custom_Menu24_ListFilesInFolder`). Menu **Excel VBA Lib → Files**. Personal OnAction names are kept.

SharePoint download (`Custom_Menu24_ShrpntDownload`) is not included (hard-coded destination path). Workbook-name helpers from that module are not included.

| Public procedure | Module | Notes |
|------------------|--------|--------|
| `ReturnFilesInSelectedFolder` | `modApiFiles` | Folder picker; sheet **File listing** for that folder only. Count-and-estimate prompt only when there are 1,000 or more files. Modeless **Cancel** during count/list writes any rows already collected (partial sheet). Personal wrote two columns at a selected cell and called missing `SelectFolder` / `GetRange`. Same columns as the recursive listing. |
| `BatchListAllFiles_FolderSubfolders` | `modApiFiles` | Same sheet, including subfolders. Counts files first (`Files.Count` per folder). Under 1,000 files, lists immediately; otherwise a prompt shows the count and an estimated listing time so you can cancel. Modeless **Cancel** during the run writes a partial **File listing**. Then collects all files and writes once (Personal reset the row counter in each subfolder and overwrote earlier rows). Hidden / system / junction folders are skipped. Attribute flags are decoded as a combination, not a single Case value. |
| `PickFolder` | `modApiFiles` | Folder picker → path string (starts at Excel's default file path, not `E:\`) |
| `ListFiles` | `modApiFiles` | Paths in one folder matching an extension filter such as `*.csv` |

### Power Query public surface

Personal Menu29 (`Custom_Menu29_PowerQuery`). Menu **Excel VBA Lib → Power Query**. Personal OnAction names are kept.

| Public procedure | Module | Notes |
|------------------|--------|--------|
| `ShowPQLibraryForm` | `modApiPowerQuery` | Form `frmPQLibrary`: import selected `tblPQ_Library` functions (with `fn*` deps) into the active workbook, or export workbook queries to sheet **PQ_Functions**. Looks up `tblPQ_Library` in any open workbook (not only Personal.xlsb). |
| `BackgroundRefreshToggle` | `modApiPowerQuery` | Flips `BackgroundQuery` on every OLEDB connection; reports how many are on/off |
| `IgnorePrivacyToggle` | `modApiPowerQuery` | Flips `Queries.FastCombine` (Ignore Privacy Levels); reports the new state |
| `Add_Connection_All_Tables` | `modApiPowerQuery` | Connection-only query per Excel Table; optional Data Model load. Skips tables that already have a query. One connection per table (Personal added two with the same name when loading to the model) |

### Analysis public surface

Personal Menu18 (`Custom_Menu18_Analysis`, `Covariance`, `ConfusionMatrix`, `CnfsnMtrxTmplt`, `Residuals`, `SVD`, `LinearSystem` / `LinearSysAXB`). Menu **Excel VBA Lib → Analysis**. Personal OnAction names are kept. SVD is also on **Matrices → Decompositions**.

Observations are rows; variables are columns. Vector/matrix tools that take a header use the first row as labels and the rest as numeric data.

| Public procedure | Module | Notes |
|------------------|--------|--------|
| `SVD` | `modApiSvd` | Writes singular values, U, and V. Golub-Reinsch; max side 500. Does not grey-fill or require a numeric output cell |
| `LinearSystem_AXB_v1` / `LinearSystem_AXB_v2` | `modApiLinearSystem` | Solve AX=B via `modInternalMatrices.Solve`. v2 is the Personal menu name; the RefEdit form is not used |
| `ConfusionMatrix` | `modApiConfusion` | Sheet **Confusion Matrix** from Yes/No (case-insensitive; Y/N). Predicted then Actual |
| `ConfusionMatrixOnesAndZeros` | `modApiConfusion` | Same sheet; 1 = positive class |
| `ConfusionMatrixTemplate` | `modApiConfusion` | Sheet **Confusion Matrix Template** with live R1C1 formulae. Personal wrote formula text that never calculated |
| `CalculateVarianceCovarianceMatrix` | `modApiAnalysis` | Population covariance (divide by n), with headers |
| `MatrixCovariance` | `modApiCovariance` | Sample covariance (n-1). Range need not be square |
| `MatrixCovarianceStandardise` | `modApiCovariance` | Sample covariance of standardised columns, with a **Standardised Covariance** label above the matrix |
| `CorrelationMatrix` | `modApiAnalysis` | Population covariance / products of population SDs |
| `ProveVarCovarAndCorrel` | `modApiAnalysis` | Sheet **Cov and Correl**: D R D = covariance, plus a Check block |
| `CalculateMeanVector` | `modApiAnalysis` | 1 × p means under the header |
| `CalculateStandardDeviationPopulationVector` / `Sample` | `modApiAnalysis` | STDEV.P / STDEV.S row vectors |
| `CalculateStdDevProductMatrixPopulation` | `modApiAnalysis` | Outer product of population SDs |
| `ResidualsAnalysis` | `modApiResiduals` | Sheet **Residuals Analysis**; sheet-scoped names `Order` / `Residuals` |
| `ComparePredictedToActual` / `…OneAndZerosOnly` | `modApiConfusion` | Optional worksheet UDFs |

Personal `GetRange` used `End` on cancel (kills Excel). Prompts now return Nothing. `ExtractBody` / `RangeArea` used unqualified `Cells` and could read the wrong sheet.

### Plots Charts public surface

Personal Menu11 (`Custom_Menu11_*`). Menu **Excel VBA Lib → Plots Charts**. Personal OnAction names are kept.

| Public procedure | Module | Notes |
|------------------|--------|--------|
| `HistogramTableAndPlot` / `HistogramFormulaeAndPlot` | `modApiHistogram` | Sheets **Histogram** / **Histogram formulae**. Formulae version uses COUNTIF so bins update. Cancel on range or bin count does nothing. |
| `LinearRegression` / `LinearRegressionV2` | `modApiLinearRegression` | Two numeric columns (x, y). V2 adds residuals, slope CI, and an integer-x fitted grid. V2 original named range is n×2 (Personal used n×n). |
| `ProcessCapabilityChart` | `modApiProcessCapability` | Sheet **Process Capability Plot**. Cp/Cpk are worksheet formulae (Personal called `personal.xlsb` UDFs). LSL/USL swap is fixed. |
| `GenerateBinomialPlot` / `GenerateNormalPlot` / `GenerateLogNormalPlot` / `GeneratePoissonPlot` / `GenerateWeibullPlot` / `GenerateGammaPlot` / `GenerateBetaPlot` / `GenerateExponentialPlot` / `GenerateHypergeometricPlot` / `GenerateLogisticCurve` | `modApiDistPlots` | Formula tables with yellow inputs; PDF+CDF except logistic (one chart). R1C1 via `FormulaR1C1`. Exponential lambda is labelled Lambda (Personal said Alpha). |
| `QQPlotGaussianNormal` / `QQPlotUniform` | `modApiQQPlots` | Sheets **QQ Normal chart** / **QQ Uniform chart**. Blanks rejected. Intercept/R² are formulae (Personal wrote some as `.Value` text). |
| `GiniPlot` | `modApiLorenz` | Sheet **Gini**. Single non-negative column. FastExcel Gini algorithm. |
| `GenerateCorrelogram` / `ACFLower` / `ACFUpper` | `modApiAcf` | Sheet **Correlogram**. PACF Yule-Walker, max 40 lags. The old “stationary?” flag is **Non-decreasing from first obs**. |
| `DieboldMarianoTest` | `modApiDiebold` | Sheet **DieboldMariano**. 1 column = RAND demo forecasts; 3 columns = actual + two forecasts. Two-sided p-value. |
| `XmR` / `GenerateXMRDiagnostics` | `modApiXmR` | Sheets **XmR** / **XMR Diagnostics**. X̄ ± 2.66 MR̄, MR UCL 3.267 MR̄. Diagnostics is a 31-column Western Electric table. |
| `PlotLineChartSheet` | `modApiChartSheet` | Chart sheet named from the header; SE bars + trendline. |

### Time series public surface

Personal Menu27 (`Custom_Menu27_TimeSeries` / `Custom_Menu27_TS_DateDiff`). Menu **Excel VBA Lib → Time series**. Personal OnAction names are kept.

`GeneratePACFList` / `GenerateAVCFList` / `GenerateAutoCovarianceMatrix` are not on this menu (they were private in Personal).

| Public procedure | Module | Notes |
|------------------|--------|--------|
| `TimeSeriesAnalysis` | `modApiTimeSeries` | Sheet **Time Series**. Values for lag 0..n/3. ACF/ACVF/PACF computed from arrays (Personal re-read the range per lag). Screen updating is not turned off before the InputBox |
| `TimeSeriesAnalysisFormula` | `modApiTimeSeries` | Same sheet; live UDFs in `Fn_TimeSeries` (Insert Function category **Excel VBA Lib**). Lags 0..n/4. A1 formulas qualified to this add-in (`='ExcelVbaLib.xlam'!ACVF(...)`). Unqualified `=ACVF(...)` is `#NAME?` until the menu has registered the UDFs |
| `DateDifferencing` | `modApiTimeSeries` | Sheet **Date Diff**. Lag differences (not calendar dates). First difference has n-1 points (Personal wrote n-2). Charts without `.Select` |
| `DateDifferencingFormulae` | `modApiTimeSeries` | Sheet **Date Diff Formula**. Difference cells as R1C1 formulas |
| `ACF` / `ACVF` / `PACF` / `Bartlett` / `BoxPierce` / `BoxPiercePVal` / `BoxPierceTest` / `LjungBox` / `LjungBoxPVal` / `LjungBoxTest` | `Fn_TimeSeries` | Worksheet UDFs (same registration as `Mat*`). Bartlett is two-sided (`|ACF|`). Box-Pierce / Ljung-Box reject white noise when p < alpha (Personal messages were inverted) |

### Matrices public surface

Api modules (VBE names): `modApiMatrixCreate`, `modApiMatrices1`, `modApiMatrices2`, `modApiCholesky`, `modApiEigenDecomp`, `modApiUnitary`, `modApiMatrixUtilities`, `modApiCovariance`, plus `Fn_MatricesArray`, `Fn_MatricesRng`, `Fn_Matrices2`.

Git snapshot math is array-based in `modInternalMatrices` (Personal dumps are gitignored). After `Import-AddinModules.ps1 -All`, overlay the original Personal modules into the `.xlam` with `scripts/Import-Menu13FromPersonal.ps1` (Windows, Personal.xlsb present). `-AllMenu13` copies the whole family. Do that **after** `-All`; running `-All` again replaces them with the git snapshot.

`git pull` updates `build\ExcelVbaLib.xlam`. Add-in macros never appear in Alt+F8; worksheet `Mat*` names are registered into Insert Function (category **Excel VBA Lib**) by `RegisterMatrixUdfs` (`Fn_MatricesRng`) when the menu installs.

| Public procedure | Module | Notes |
|------------------|--------|--------|
| `MatrixCreateIdentity` / `Zeros` / `Ones` / `Diagonal` / `Random` / `Hilbert` / `Exchange` / `Toeplitz` / `Vandermonde` / `Companion` | `modApiMatrixCreate` | Create writes at the active cell; vector-based create writes to the right of the selection |
| `MatrixTranspose` / `Add` / `Subtract` / `Scale` / `Multiply` / `MultiplicationHadamard` / `MultiplicationKronecker` / `Outer` / `Dot` / `Inverse` / `Power` / `Determinant` / `Trace` / `DiagExtract` / `Vec` / `Unvec` | `modApiMatrixUtilities` | Result one column to the right of the selection (Hadamard prompts; default is to the right of B) |
| `MatrixHadamardProof` | `modApiMatrixUtilities` | Sheet **Hadamard Proof**: H, Hᵀ, H.HT, then n and I or a not-Hadamard message |
| `MatrixSolve` / `Rank` / `Norm` / `Norm1` / `NormInf` / `IsSymmetric` / `Adjugate` / `PseudoInverse` / `LU` / `Cofactor` / `Minor` | `modApiMatrices2` | Cofactor limited to order 20 |
| `MatrixCholesky` | `modApiCholesky` | SPD only |
| `MatrixEigen` | `modApiEigenDecomp` | Jacobi; symmetric only. Writes vectors then a λ column |
| `MatrixQR` / `MatrixIsOrthogonal` | `modApiUnitary` | QR stacks R below Q. Complex unitary from Personal is not included |
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
| `ReturnFilesInSelectedFolder` | List files in one folder on **File listing** |
| `BatchListAllFiles_FolderSubfolders` | List files in a folder tree on **File listing** |
| `PickFolder` | Folder picker → path string |
| `ListFiles` | Extensions filter → path array |

### `modApiTables`
| Public procedure | Responsibility |
|------------------|----------------|
| `ListTableProperties` | Inventory sheet of all ListObjects with hyperlinks |
| `ShowAllTablesInWorkbook` | Message box listing table name / sheet / address |
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
| PQ Cleanse\* / profiling | Rarely needed in VBA; workbook PQ connections are `modApiPowerQuery` |
| Output of any macro | Always via `modInternalSheetIO` |

## After Trust access is enabled

Re-run an export of Personal into `source/_export_raw/`, then fill this checklist:

- [ ] List every `Public Sub/Function` and assign an Api module  
- [ ] List shared helpers → Internal  
- [ ] Delete dead code / duplicates  
- [ ] Update this document with the **actual** procedure inventory  
