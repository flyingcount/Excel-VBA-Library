# Test: Reconciliations (Personal Menu20)

Personal `Custom_Menu20_Reconciliation` / `Custom_Menu20_RecnStrings` / `Custom_Menu20_CompareRanges`, rewritten as `modApiReconciliations` + `modInternalReconciliations`.

`CompareTwoRangesForm` is not included; **Compare two ranges** uses InputBoxes (same as the Personal `CompareRanges` Sub).

## Setup

1. Load `build/ExcelVbaLib.xlam`. Fully quit Excel and restart if the menu was just rebuilt.
2. Use **Excel VBA Lib → Reconciliations**. Do not import Personal `Custom_Menu20_*` modules into the test workbook.
3. Work in a throwaway workbook. Number recon writes sheet **Reconciliation**; string recon writes **Reconciliation Strings**. Each run replaces that sheet.

Cancel on any InputBox does nothing. Entire-column selections use UsedRange. A two-column selection is rejected. Personal checked `.Columns` before `Nothing`, so Cancel could error. Second-range errors said “First range”. Name prompts for the second dataset said “Enter first dataset name”.

## Reconcile two columns of numbers

```
A1: 10
A2: 20
A3: 30
C1: 10
C2: 20
C3: 99
```

Select `A1:A3`. **Reconciliations → Reconcile two columns of numbers**. First range `A1:A3`, name `Bank`. Second range `C1:C3`, name `Books`.

Expected: sheet **Reconciliation** with a timestamp in A1.

- Columns H–M: Bank rows. 10 and 20 are `Matched` with flags `A` and `B`; 30 is unmatched.
- Columns O–T: Books rows. 10 and 20 matched; 99 unmatched.
- Column A statement: Bank total 60, minus unmatched 30, plus Books unmatched 99, Books total 99, unreconciled balance 0 when the unmatched items are applied (Personal’s running-total layout).
- Non-numeric cells are skipped for matching and totals (Personal added them with `+` and could type-mismatch).
- Matching is 1-for-1 in row order (Dictionary queues), not an O(n×m) nested loop.
- Personal also dumped both datasets stacked in column H and a third copy in V; this version writes Bank in H and Books in O only.
- Duplicate unmatched amounts (if any) are listed once below the statement. Personal’s warning block could repeat the same rows.

A single-cell column is allowed (Personal `Range.Value` was a scalar, so `ReDim Preserve` failed).

## Reconcile two columns of strings

```
A1: apple
A2: banana
C1: apple
C2: cherry
```

**Reconciliations → Reconcile two columns of strings**. Names `Left` / `Right`.

Expected: sheet **Reconciliation Strings**. apple matched (`A`). banana listed as `Record not in Right`. cherry listed as `Record not in Left`. No numeric totals. The numeric **Reconciliation** sheet is left unchanged.

## Compare two ranges

```
A1: 1  B1: 2
A2: 3  B2: 4
D1: 1  E1: 9
D2: 3  E2: 4
```

Select `A1:B2` as range 1 and `D1:E2` as range 2. **Reconciliations → Compare two ranges**.

Expected: B1 and E1 filled yellow. Sheet **Range Comparison** with a header at A3 (Index, worksheets, rows, columns, values, Difference flag) covering **all four cells**. Message **1 differences found.** (Personal spelling was `diffrences`.) Table name `Comparison`.

Personal’s loop started at `Cells(2)`, so the first cell was never compared.

Different sizes: message and no sheet. A VBA error in a cell is treated as a difference instead of raising.
