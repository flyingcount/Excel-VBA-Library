# Test: worksheet templates

## Setup

1. Load `ExcelVbaLib.xlam` — see [build/README.md](../build/README.md).
2. Open a blank workbook (do not import Personal `Custom_Menu26_wkshtTmplt`).

## Steps

Immediate window:

```vb
Application.Run "CreateLinksTemplate"
Application.Run "CreateActionsTemplate"
Application.Run "CreateForceFieldTemplate"
Application.Run "CreateAssumptionsTemplate"
Application.Run "CreateQuestionsTemplate"
```

Or menu **Excel VBA Lib → Worksheet templates**.

## Expected

| Macro | Sheet | Table | Notes |
|-------|--------|--------|-------|
| `CreateLinksTemplate` | Links | `Tbl_Links` | Sample Google link; `=HyperLinkText(C4)` in F4; Category slicer |
| `CreateActionsTemplate` | Actions | `Tbl_Actions` | Live/Dead, Owner, Action Category slicers; overdue days in red |
| `CreateForceFieldTemplate` | Force Field | `Tbl_ForceField` | Drivers / weights / barriers |
| `CreateAssumptionsTemplate` | Assumptions | `Tbl_Assumptions` | Sample ref A1; Status slicer |
| `CreateQuestionsTemplate` | Questions | `Tbl_Questions` | Sample ref Q1 |

- Re-running a template when the sheet exists prompts to rename it `… OLD`. Cancel leaves the original.
- `CreateNotesWorkbook` builds all five plus an Index, then a Save As dialog (no hardcoded path).
- `ImportPythonPackages` writes `=PY(...)` stubs on **Python packages** and moves that sheet first.
