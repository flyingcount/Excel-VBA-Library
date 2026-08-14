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
- Add-in menu loads on Excel start via `ThisWorkbook` `Workbook_Open` (`.xlam` does not run `Auto_Open`).
- `ExtractSample` (`source/Api/modApiSampling.bas` + `source/Internal/modInternalSampling.bas`) — random row sample from a range (Personal Menu4).

### Changed
- The library is `ExcelVbaLib.xlam`. Placeholder `source/` folders (`Features/`, `Udf/`, `Forms/`, `Classes/`, `Sandbox/`) were removed; add new modules in the add-in VBE.

## [0.1.0] - 2026-08-09

### Added
- Initial Excel VBA library repository (split from PowerQuery-Library).
- Api/Internal module stubs for arrays, sheets, files, tables, UI, and shared helpers.
- Docs: export/import, module map, Personal123 inventory, infrastructure catalog.
