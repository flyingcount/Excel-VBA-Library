# Changelog

All notable changes to this library are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/).

## [Unreleased]

### Added
- `CreateDateTable` (`source/Api/modApiDates.bas` + `source/Internal/modInternalDateTable.bas`) — write a calendar dimension sheet with fiscal year/period columns.
- Benford analyses (`source/Api/modApiBenford.bas` + `source/Internal/modInternalBenford.bas` + `source/Internal/modInternalNamedRanges.bas`) — first / second / third / two-digit / three-digit / last-two-digit.
- Add-in build: `scripts/Build-ExcelVbaLib.ps1` writes `build/ExcelVbaLib.xlam` (Internal + Api + Menus in one project). Load that file instead of importing modules into caller workbooks.
- Worksheet templates (`source/Api/modApiWorksheetTemplates.bas` + `source/Internal/modInternalWorksheetTemplates.bas`) — Links, Actions, Force Field, Assumptions, Questions, Notes workbook, Python packages.
- Add-in menu loads on Excel start via `ThisWorkbook` `Workbook_Open` / `Workbook_AddinInstall` (`.xlam` does not run `Auto_Open`).
- `ExtractSample` (`source/Api/modApiSampling.bas` + `source/Internal/modInternalSampling.bas`) — random row sample from a range (Personal Menu4).
- Data fills (`source/Api/modApiData.bas` + `source/Internal/modInternalData.bas`) — Personal Menu6 random / combinations / probability distributions. Menu **Excel VBA Lib → Data** (with **Probability distributions** submenu). Prime-number generator from Personal Menu10 is not included.
- `scripts/Import-AddinModules.ps1` — replace modules in `ExcelVbaLib.xlam` (default: `modInternalData`, `modApiData`). `-All` writes Internal + Api + Menus into `build\ExcelVbaLib.xlam`. Pass `-XlamPath` to target a specific add-in file.
- Matrices — `modApiMatrixCreate` / `modApiMatrices1` / `modApiMatrices2` / `modApiCholesky` / `modApiEigenDecomp` / `modApiUnitary` / `modApiMatrixUtilities`, plus `Fn_Matrices*`. Math is in `modInternalMatrices`. Overlay Personal VBA with `scripts/Import-Menu13FromPersonal.ps1` after `-All` (rewrites Personal `Custom_Menu13_*` names to `modApi*`). Extra ops: vec/unvec, companion, cofactor, minor. Personal-only Menu13 tools (diagnostic, formulae, rotations, sheet eigen/Cholesky) sit on **Excel VBA Lib → Matrices** Create / Operations / Decompositions; duplicate identity/zeros/ones/exchange/transpose/dot items are not listed twice.
- Hyperlinks (`source/Api/modApiHyperlinks.bas` + `source/Internal/modInternalHyperlinks.bas`) — Personal Menu21 inventory, index, remove-by-display-text, follow selection, and back-links. Menu **Excel VBA Lib → Hyperlinks**.
- Power Query (`source/Api/modApiPowerQuery.bas` + `source/Internal/modInternalPowerQuery.bas` + `frmPQLibrary`) — Personal Menu29 import/export form, toggle background refresh, Fast Combine, and connect-all-tables. Menu **Excel VBA Lib → Power Query**.
- Custom lists (`source/Api/modApiCustomLists.bas` + `source/Internal/modInternalCustomLists.bas`) — Personal Menu16 list / create / delete custom lists, plus AutoCorrect inventory and add. Menu **Excel VBA Lib → Custom lists**. Lists are application-wide; delete confirms and removes highest list numbers first.
- Tables (`source/Api/modApiTables.bas` + `source/Internal/modInternalTables.bas`) — Personal Menu19 `ListTableProperties` inventory sheet (sized to the table count, with working sheet-qualified hyperlinks) and `ShowAllTablesInWorkbook`. Menu **Excel VBA Lib → Tables**.
- Files (`source/Api/modApiFiles.bas` + `source/Internal/modInternalFiles.bas` + `frmFileListingProgress`) — Personal Menu24 folder listings on sheet **File listing** (this folder, or folder and subfolders). Menu **Excel VBA Lib → Files**. Recursive listing no longer overwrites earlier rows; attributes are decoded as flags. A count-only pass runs first; if there are 1,000 or more files, a prompt shows the count plus an estimated listing time so you can cancel. A modeless Cancel dialog stops the run and writes any files already collected.

- Matrices **Validation → Hadamard Proof** prompts for a range, writes sheet **Hadamard Proof** with H, Hᵀ, and H.HT, then n and I when H Hᵀ = n I, otherwise `Not Hadamard: H.HT not equal to nI`.
- Analysis (Personal Menu18) on **Excel VBA Lib → Analysis**: SVD, AX=B, confusion matrices (Yes/No, 1/0, template), population covariance/correlation/mean/stdev vectors, residuals plot. Sample covariance stays on **Matrices → Operations** (`MatrixCovariance` / `MatrixCovarianceStandardise`) and no longer requires a square range.
- Plots Charts (Personal Menu11) on **Excel VBA Lib → Plots Charts**: histograms, linear regression, process capability, parametric PDF/CDF plots (including logistic curve), QQ plots, Lorenz/Gini, ACF/PACF, Diebold-Mariano, XmR, and a line chart sheet (`modInternalPlots` + `modApiHistogram` / `modApiDistPlots` / `modApiQQPlots` / `modApiLinearRegression` / `modApiLorenz` / `modApiAcf` / `modApiDiebold` / `modApiXmR` / `modApiProcessCapability` / `modApiChartSheet`).
- Ranges (Personal Menu14) on **Excel VBA Lib → Ranges**: named-range inventory / hide / create / scope, unique values, range analysis, character frequency, cleanse, transpose, and split-to-named-columns (`modApiRanges` + `modInternalRanges`, with helpers on `modInternalNamedRanges`). Personal OnAction names are kept. Data-cleanse / validation / multi-replace UserForms are not included.
- Duplicates (Personal Menu15) on **Excel VBA Lib → Duplicates**: flag later occurrences, letter-label duplicate groups, colour groups / by row / by column / in selection, and count duplicate cells (`modApiDuplicates` + `modInternalDuplicates`). Personal OnAction names are kept. Blanks and errors are skipped; flag/reference tools insert cells in the selected rows only.
- Protection (Personal Menu7) on **Excel VBA Lib → Protection**: scroll-area limit, default-password protect/unprotect (`UserInterfaceOnly`), show password, unhide worksheets, unhide rows/columns (`modApiProtection` + `modInternalProtection`). Personal OnAction names are kept. The default password is the Personal convenience value, not a secret.
- Reconciliations (Personal Menu20) on **Excel VBA Lib → Reconciliations**: two-column number/string recon and cell-by-cell range compare (`modApiReconciliations` + `modInternalReconciliations`). Personal OnAction names are kept. Compare uses InputBoxes (`CompareTwoRangesForm` is not included). String recon writes sheet **Reconciliation Strings**; numeric recon still uses **Reconciliation**.
- Time series (Personal Menu27) on **Excel VBA Lib → Time series**: ACF/PACF analysis (values or formulae) and lag differencing with charts (`modApiTimeSeries` + `modInternalTimeSeries` + `Fn_TimeSeries`). Personal OnAction names are kept. `ACF` / `ACVF` / `PACF` and the portmanteau UDFs are registered into Insert Function (category **Excel VBA Lib**) like `Mat*`. Unqualified `=ACVF(...)` is `#NAME?` until that registration runs. **Date differencing with formulae** adds an autocorrelation table for the original series and every difference order; lag-1 ACF < −0.5 is flagged as over-differenced.

### Changed
- Track `build/ExcelVbaLib.xlam` in git (`build/*.xlam` is no longer gitignored).
- Files listings count files first (`Files.Count` per folder, no per-file property reads). Fewer than 1,000 files are listed immediately. At 1,000 or more, a prompt shows the count and a rough listing-time estimate; Cancel skips the detailed walk and does not create **File listing**. Default button is No when there are more than 10,000 files or the high estimate is 30 seconds or more. A modeless **Files** dialog with Cancel stops counting or listing; if any file details were already collected they are written as a partial **File listing**.
- Power Query library form uses a two-pane layout (function list + description / dependencies / M code), Segoe UI, and Import / Export workbook / Close on a bottom button row.
- Power Query import no longer prefixes query names with `Function Library/`. Names are `{category}/{function}`, or just the function name when the library row has no category.
- Personal Menu13/18 modules in the add-in are now `modApi*` (`modApiMatrixCreate`, `modApiMatrices1`, `modApiMatrices2`, `modApiCholesky`, `modApiEigenDecomp`, `modApiUnitary`, `modApiMatrixUtilities`, `modApiCovariance`) instead of `Custom_Menu13_*` / `Custom_Menu18_*`.
- Matrices **Operations** Hadamard product is now **Multiplication-Hadamard** (`MatrixMultiplicationHadamard`). It prompts for an output cell defaulting to the right of matrix B.
- Matrices **Operations** Kronecker product is now **Multiplication-Kronecker** (`MatrixMultiplicationKronecker`; `MatrixKronecker` still runs the same op).
- The library is `ExcelVbaLib.xlam`. Placeholder `source/` folders (`Features/`, `Udf/`, `Forms/`, `Classes/`, `Sandbox/`) were removed; add new modules in the add-in VBE.
- Add-in menu install no longer calls `Auto_Open` from `Workbook_Open` (Excel skips that name for add-ins). `Workbook_Open` / `Workbook_AddinInstall` schedule `InstallExcelVbaLibMenu` via `Application.OnTime`. Re-inject ThisWorkbook with `scripts/Inject-ThisWorkbook.ps1` or a full rebuild.
- Matrices **Properties** submenu holds All properties, determinant, trace, rank, the three norms, condition number, spectral radius, eigenvalues, and eigenvectors. Each writes a label with the value in the cell to its right.

### Fixed
- Power Query import/export failed to compile: `PqLibraryTable` and the other library constants sat after procedures, so VBA treated them as undefined. They now live at the top of `modInternalPowerQuery`.
- `RandomPoissonNumbers` (Data → Probability distributions → Poisson) raised error 438 because Excel has no `WorksheetFunction.Poisson_Inv`. Draws now use Knuth's method, with a rounded normal approximation for large lambda.
- Matrices **Is symmetric** / **Is orthogonal** raised "You cannot change part of an array" because they wrote a Boolean into the cell to the right of the matrix (often a leftover array formula). They now prompt for output like **Size** and write `Symmetric`/`Orthogonal` with TRUE/FALSE in the cell to the right.
- Confusion matrix template formulae now calculate (`FormulaR1C1`). Personal dumped formula strings with `WriteArrayToWorksheet`.
- `MatrixCovariance` no longer requires a square range (observations × variables).
- Residuals analysis writes to the output sheet only (Personal used unqualified `Range` / workbook-level names) and rejects blanks (`IsNumeric` treats empty as numeric).

## [0.1.0] - 2026-08-09

### Added
- Initial Excel VBA library repository (split from PowerQuery-Library).
- Api/Internal module stubs for arrays, sheets, files, tables, UI, and shared helpers.
- Docs: export/import, module map, Personal123 inventory, infrastructure catalog.
