# Test: Power Query (Personal Menu29)

Personal `Custom_Menu29_PowerQuery`, rewritten as `modApiPowerQuery` + `modInternalPowerQuery` + `frmPQLibrary`.

## Setup

1. After `git pull`, load `build/ExcelVbaLib.xlam`. To rebuild from source:

   ```powershell
   powershell -ExecutionPolicy Bypass -File .\scripts\Import-AddinModules.ps1 -All
   ```

2. Load `ExcelVbaLib.xlam`. Fully quit Excel and restart if the menu was just rebuilt.
3. Use **Excel VBA Lib → Power Query**. Do not import `Custom_Menu29_PowerQuery` into the test workbook.

## Import or export queries and functions

Needs a table named `tblPQ_Library` (Personal's `PQ_Library` sheet). Open Personal.xlsb first, or pick that workbook when prompted.

**Power Query → Import or export queries and functions**

The form is titled **Power Query library**. Search and the function list are on the left; description, dependencies, and M code (read-only) are on the right. Buttons are **Import**, **Export workbook**, and **Close**.

- Search/select functions, **Import**: creates or replaces queries named `{category}/{function}` (or just the function name when there is no category) in the active workbook, including `fn*` dependencies. Preview uses the function name (Personal used the filtered list index, which showed the wrong M code).
- **Export**: writes every query in the active workbook to sheet **PQ_Functions** (table `PQ_Functions`). Running export again replaces the table body.

The destination/source cannot be the add-in itself.

## Toggle background refresh

In a workbook with at least one Power Query query loaded as a connection (Data → Queries & Connections):

**Power Query → Toggle background refresh**

Expected: a message that background refresh is now ON or OFF for the OLEDB connection(s). Run it again: the state flips. A workbook with no OLEDB connections reports that and writes nothing.

## Toggle privacy (Fast Combine)

**Power Query → Toggle privacy (Fast Combine)**

Expected: a message that Ignore Privacy Levels is now ON, then OFF on the next run. Excel 2016+ is required (`Workbook.Queries`). Cancel is not offered; each run flips the setting.

## Connect all tables

Create two Excel Tables (`Insert → Table`) named `Sales` and `Items` on a sheet. Leave a third range as a plain range (not a table).

**Power Query → Connect all tables**

1. Confirm Yes to create connections.
2. Choose No for Data Model.

Expected: two queries (`Sales`, `Items`) and two connections (`Query - Sales`, `Query - Items`). A summary reports 2 created and 0 already present. Run again: 0 created, 2 already had a query.

With Data Model = Yes on a fresh workbook with one table: one connection is created with `CreateModelConnection:=True` (Personal added a second connection with the same name, which can fail).

Cancel on the first Yes/No box creates nothing.
