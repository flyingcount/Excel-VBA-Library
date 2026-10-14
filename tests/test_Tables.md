# Test: Tables (Personal Menu19)

Personal `Custom_Menu19_Tables`, rewritten as `modApiTables` + `modInternalTables`.

## Setup

1. Load `build/ExcelVbaLib.xlam`. Fully quit Excel and restart if the menu was just rebuilt.
2. Use **Excel VBA Lib → Tables**. Do not import Personal `Custom_Menu19_Tables` into the test workbook.

## List table properties

On a workbook with at least one Excel Table (`Insert → Table`), named e.g. `Sales` on Sheet1 covering `A1:C10` (header plus 9 data rows):

**Tables → List table properties**

Expected: sheet **Table properties** with a title in A1 and a header row at A3. Column A (Index) is width 7. One data row per ListObject:

- Index, Name (`Sales`), Display name, Worksheet (`Sheet1`), Range (`$A$1:$C$10`)
- **Link to table** is a hyperlink to that range; clicking it selects the table
- No. rows = 9, No. columns = 3
- Show autofilter / headers / totals, Table style, Source type `Range`
- Header range `$A$1:$C$1`

If the workbook has no tables, a message box says so and no sheet is written.

Personal wrote a pre-sized 1000-row array (mostly blank), used `ListObject.Active` (not a ListObject property), dumped Creator / Application / Parent type, and could leave empty rows on the sheet. This version sizes the array to the table count and uses `ListRows.Count` so an empty table reports 0 rows instead of `"None"`.

A table on a sheet named `O'Brien` still gets a working hyperlink (the sheet name is quoted and apostrophes doubled).

## List tables in a message box

**Tables → List tables in a message box**

Expected: a message box titled **Tables** listing index, table name, sheet (with `(hidden)` / `(very hidden)` when relevant), and address.

Personal Menu19 called `ShowAllTablesInWorkbook` but the Sub was commented out, so the menu item did nothing. The older body passed `vbOKOnly` as the title.

If there are no tables, the box says so. A very long list is truncated and points to **List table properties**.
