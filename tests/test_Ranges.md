# Test: Ranges (Personal Menu14)

Personal `Custom_Menu14_ColRange` / `Custom_Menu14_FreqAnalysis` / `Custom_Menu14_ManipulateRanges` / `Custom_Menu14_RangeAnalysis` / `Custom_Menu14_SplitRange`, rewritten as `modApiRanges` + `modInternalRanges` + named-range helpers on `modInternalNamedRanges`.

Forms used from Personal Menu14 (`DataCleansingOptionsForm`, `DataValidationForm`, `MultiSearchAndReplace`) are not in this pack.

## Setup

1. Load `build/ExcelVbaLib.xlam`. Fully quit Excel and restart if the menu was just rebuilt.
2. Use **Excel VBA Lib → Ranges**. Do not import Personal `Custom_Menu14_*` modules into the test workbook.
3. Work in a throwaway workbook. Several tools create or hide named ranges.

## List named range properties

Create a workbook-scoped name `ZzzTest` referring to `Sheet1!$A$1`. **Ranges → List named range properties**.

Expected: sheet **Range properties** with a title in A1, timestamp in A2, and a header row at A3: Index, Name, Refers to, Address, Visible, Comment, Workbook parameter, Scope.

- One row for `ZzzTest` with Scope `Workbook` and Address `$A$1`.
- Frozen panes below the header.
- Personal used `Name.Index` (not a Name property, so the column was often blank), wrote through `On Error Resume Next`, and dumped `Range.ListNames` in columns J–K. This version uses a sequential index and skips ListNames.

Cancel on any InputBox does nothing.

## Paste as picture

Select `A1:B2`. **Ranges → Paste as picture**.

Expected: a picture of that range on the active sheet. If the selection is not a range, a message box says to select a range first.

## Colour / remove named-range colour

**Ranges → Colour named ranges** → Yes (workbook).

Expected: cells referred to by named ranges are filled ColorIndex 36 (light yellow). Names that are constants (no `RefersToRange`) are skipped. Cancel stops.

**Ranges → Remove named-range colour** → Yes.

Expected: that fill is cleared (`xlColorIndexNone`). Personal set `ColorIndex = 0`.

## Delete specified names

On **Range properties**, select the Name cell for `ZzzTest`. **Ranges → Delete specified names**. Confirm Yes.

Expected: the name is gone and **Range properties** is rewritten. Cancel on the confirm prompt deletes nothing. Unknown names in the selection are listed in a warning; other names are still deleted.

## Create named ranges

**Hidden workbook-scope constants.** Two columns:

```
A1: ZzzConst
B1: 42
```

Select `A1:B1`. **Ranges → Create named ranges → Hidden workbook-scope constants**.

Expected: hidden name `ZzzConst` with RefersTo `=42`. One-column selection is rejected. Names are sanitised (`ValidRangeName`: spaces to `.`, illegal characters stripped, leading digit prefixed).

**Hidden named range.** Select `A1:A3`. **Create named ranges → Hidden named range**. Name `ZzzBlock`.

Expected: hidden workbook name referring to that range. Cancel on the name box creates nothing. Personal allowed letters-only names; this version accepts a valid Excel name.

**Common constants.** **Create named ranges → Common constants (thousand / million)**.

Expected: visible names `Divisor_Thousand` (`=1000`) and `Divisor_Millions` (`=1000000`).

**Name each column from header.**

```
A1: Alpha
B1: Beta
A2: 1
B2: 2
```

Select `A1:B2`. **Create named ranges → Name each column from header**.

Expected: hidden sheet-scoped names `Alpha` and `Beta` referring to `A2` and `B2` (header excluded). A single-cell range is rejected. A header-only row (one row) is rejected. Empty header cells are rejected. Personal required “letters only” with a buggy mix of numeric/alpha headers.

**Local name on every sheet.** Select `B2`. **Create named ranges → Local name on every sheet**. Name `ZzzLocal`.

Expected: a visible sheet-scoped `ZzzLocal` on every worksheet, each referring to `$B$2` **on that sheet**. Personal pointed every sheet’s name at the original selection’s sheet.

## Hide / unhide

**Hide all names (workbook)** / **Unhide all names (workbook)** flip `.Visible` on every name in the workbook.

**Hide names on active sheet** / **Unhide names on active sheet** affect only sheet-scoped names on the active sheet.

**Hide specified names** / **Unhide specified names** use names listed in the selection (column B of **Range properties** works; local name without a sheet prefix is enough). Then **Range properties** is rewritten.

## List unique values

```
A1: a
A2: a
A3: b
B1: 1
B2: 1
B3: 1
```

Select `A1:B3`. **Ranges → List unique values**.

Expected: sheet **Unique Values** with two rows (`a`/`1` and `b`/`1`). All selected columns are the uniqueness key. Personal passed only the last column index into `RemoveDuplicates`, so two columns collapsed as if only the last column mattered.

## Scope conversion

Create a sheet-scoped name `ZzzScope` on Sheet1. Put `ZzzScope` in a cell. **Ranges → Scope: worksheet to workbook**.

Expected: workbook-scoped `ZzzScope` with the same RefersTo; **Range properties** refreshes. Personal parsed addresses with `InStr`/`Right` and often failed.

**Scope: workbook to worksheet** asks for a cell on the destination sheet. The public name stays `ConvertNamedRangeGlodalToLocalScope` (Personal spelling).

## Analyse range

Select a mixed block (formula, number, text, blank). **Ranges → Analyse range → Analyse to a sheet**.

Expected: sheet **Range_Analysis** with Range, Formula, Arrays, Numeric, Text, Blanks, Error, Even, Odd, Non text, Logical, N/A, totals, rows, columns.

- Blanks are not counted as Numeric (`IsNumeric` on an empty cell is True in VBA). Personal counted blanks as numeric.
- Even/Odd are integer numeric cells. Personal’s message box computed them then wrote `n/a`.
- Entire-column selections are intersected with UsedRange.

**Analyse in a message box** shows the same figures.

Cancel on the range prompt writes nothing / shows no box.

## Character frequency

Select cells containing `Aab`. **Ranges → Analyse range → Character frequency**.

Expected: sheet **Frequency Analysis** with the source address, character count, concatenated text in D1, and one row per **used** character (code, label, count, %). Sorted by code. Space/tab/CR/LF/NBSP are labelled. Personal wrote all ASCII 1–255 (mostly zeros) with an O(n×255) nested loop and divided by length without guarding an empty range.

## Cleanse range

On copies of `A1: A-12b `:

- **Keep digits** → `12`
- **Keep letters** → `Ab`
- **Keep letters, digits, spaces** → `A12b ` (period is also kept, matching Personal)
- **Trim whitespace** on a cell that is `a  b` with NBSP/tab → `a b`
- **Trim whitespace (regex)** does the same via `VBScript.RegExp`. Personal’s menu called `RemoveWhiteSpacesFn` (a Function); this pack’s menu calls `RemoveWhiteSpaces`.

**List unique characters** on `A1: A` `A2: A` writes sheet **Characters and Codes** with one row for `A` (code 65) plus a unique count. Codes are sorted.

## Transpose to the right

Select `A1:B2` with values 1,2 / 3,4. **Ranges → Transpose to the right**.

Expected: a 2×2 block starting one column to the right of the selection (`D1:E2` if the selection starts at A) containing 1,3 / 2,4. Uses the selection’s sheet (`Cells` is qualified). A single cell copies to the same relative destination. Personal used unqualified `Cells` (could hit the wrong sheet) and `Application.Volatile` in a Sub.

## Split into named columns

Select `A1:B3` (optional header in row 1). **Ranges → Split into named columns**. Output name `Output`. Yes = has header.

Expected: workbook names `OutputAll` (whole range), `Output_1` (column A without header), `Output_2` (column B without header). A single-column range still gets `OutputAll` and `Output_1`. Cancel on the header question creates nothing. Personal’s single-column branch used an unset `col` / a zero column counter, and Cancel called `TurnOff` instead of restoring screen updating.
