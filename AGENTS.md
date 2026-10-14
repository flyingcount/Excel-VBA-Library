# Agent notes — Excel VBA Library

This repository is an **Excel VBA add-in** (`ExcelVbaLib.xlam`). There is no Node/Python package manager, Docker Compose stack, or long-running server. The product runtime is **Microsoft Excel for Windows** with VBA and COM automation.

Standard human workflow: [README.md](README.md), build/inject scripts under `scripts/`, manual tests under `tests/*.md`.

## Cursor Cloud specific instructions

Cloud Linux agents **cannot** run Excel, rebuild `build/ExcelVbaLib.xlam`, or execute the manual test procedures in `tests/`. Those steps require a Windows machine with desktop Excel.

### What works in Cloud

- Edit/read VBA source under `source/Api`, `source/Internal`, `source/Menus`.
- Edit docs (`docs/`), changelogs, and test procedure markdown.
- Review PowerShell under `scripts/` (do not expect COM/`Excel.Application` to succeed here).

### What must run on Windows + Excel

| Action | Command / location |
|--------|--------------------|
| Trust VBA project model | [docs/ExportImport.md](docs/ExportImport.md) |
| Rebuild add-in | `powershell -ExecutionPolicy Bypass -File scripts/Build-ExcelVbaLib.ps1` (Excel must not have the add-in loaded) |
| Inject ThisWorkbook events into a loaded add-in | `powershell -ExecutionPolicy Bypass -File scripts/Inject-ThisWorkbook.ps1` then fully quit and restart Excel |
| Load add-in | Excel → Options → Add-ins → Excel Add-ins → `build/ExcelVbaLib.xlam` |
| Manual regression | Follow `tests/test_*.md` (Immediate window / ribbon) |

### Gotchas

- `.xlam` add-ins **do not run `Auto_Open`**. Menu install is scheduled from `ThisWorkbook` via `Application.OnTime` (`Workbook_Open` / `Workbook_AddinInstall`). After pulling ThisWorkbook/menu changes, re-inject or full-rebuild, then restart Excel.
- `build/*.xlam` and `Data/*.xlsb` are gitignored; each Windows clone builds its own add-in.
- There is no lint/test/build toolchain on Linux for this repo; “green” CI here means source/docs review only.
