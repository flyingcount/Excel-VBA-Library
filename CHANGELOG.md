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
- `scripts/Import-AddinModules.ps1` — replace modules in `ExcelVbaLib.xlam` (default: `modInternalData`, `modApiData`). `-All` writes Internal + Api + Menus into `build\ExcelVbaLib.xlam` (needed because the `.xlam` is gitignored). Pass `-XlamPath` to target a specific add-in file.
- Matrices (`source/Api/modApiMatrices.bas` + `source/Internal/modInternalMatrices.bas`) — Personal Menu13 create / operate / Cholesky / symmetric eigen / QR, plus `Mat*` worksheet UDFs and `MatrixMultDefined`. Menu **Excel VBA Lib → Matrices**. Personal `.bas` dumps were gitignored; this is an array-based reimplementation of that family (not a line-for-line paste). `scripts/Export-PersonalModules.ps1` exports the original modules when Personal.xlsb is present.

### Changed
- The library is `ExcelVbaLib.xlam`. Placeholder `source/` folders (`Features/`, `Udf/`, `Forms/`, `Classes/`, `Sandbox/`) were removed; add new modules in the add-in VBE.
- Add-in menu install no longer calls `Auto_Open` from `Workbook_Open` (Excel skips that name for add-ins). `Workbook_Open` / `Workbook_AddinInstall` schedule `InstallExcelVbaLibMenu` via `Application.OnTime`. Re-inject ThisWorkbook with `scripts/Inject-ThisWorkbook.ps1` or a full rebuild.

### Fixed
- `RandomPoissonNumbers` (Data → Probability distributions → Poisson) raised error 438 because Excel has no `WorksheetFunction.Poisson_Inv`. Draws now use Knuth's method, with a rounded normal approximation for large lambda.

## [0.1.0] - 2026-08-09

### Added
- Initial Excel VBA library repository (split from PowerQuery-Library).
- Api/Internal module stubs for arrays, sheets, files, tables, UI, and shared helpers.
- Docs: export/import, module map, Personal123 inventory, infrastructure catalog.
