# Test: Duplicates (Personal Menu15)

Personal `Custom_Menu15_Duplicates`, rewritten as `modApiDuplicates` + `modInternalDuplicates`.

## Setup

1. Load `build/ExcelVbaLib.xlam`. Fully quit Excel and restart if the menu was just rebuilt.
2. Use **Excel VBA Lib → Duplicates**. Do not import Personal `Custom_Menu15_Duplicates` into the test workbook.
3. Work on a throwaway sheet. Flag / unique-reference tools insert cells to the right of the selection (the selected rows only, not the whole worksheet column).

Cancel on any InputBox does nothing. Entire-column selections are intersected with UsedRange. Blank and error cells are skipped (Personal `CountIf` treated blanks as duplicates of each other).

## Flag duplicates (single column)

```
A1: apple
A2: banana
A3: apple
A4: (blank)
A5: apple
```

Select `A1:A5`. **Duplicates → Flag duplicates (single column)**. Accept the default range.

Expected: cells to the right of A3 and A5 contain `Duplicate`. A1 (first apple) and A2 stay blank in the flag column. A4 is not flagged. A two-column selection is rejected. Personal inserted an **entire** worksheet column.

## Unique references for duplicates

On a copy of the same column. **Duplicates → Unique references for duplicates**.

Expected: the apple group is labelled `A` on **every** apple (A1, A3, A5). Banana stays blank. A second duplicate group would be `B`. Personal used `Collection` + `On Error 457` and `Cells(1, n).Address` for letters (unqualified `Cells`).

## Colour duplicate groups

```
A1: x
B1: y
A2: x
B2: z
A3: y
```

Select `A1:B3`. **Duplicates → Colour duplicate groups**.

Expected: both `x` cells share one fill, both `y` cells share a different fill, `z` is unfilled. Groups cycle a ColorIndex palette (Personal incremented ColorIndex on every duplicate cell and could raise "Too many duplicate!" at ColorIndex 57).

## Colour duplicates by row

```
A1: a  B1: a  C1: b
A2: a  B2: c  C2: d
```

Select `A1:C2`. **Duplicates → Colour duplicates by row**.

Expected: A1 and B1 are red (ColorIndex 3). Row 2 is unfilled. Existing fill in the selection is cleared first (`xlColorIndexNone`; Personal used 0). A single-column range is rejected.

## Colour duplicates by column

Same block. **Duplicates → Colour duplicates by column**.

Expected: A1 and A2 are green (ColorIndex 4). A single-row range is rejected.

## Colour duplicates in selection

Same block. **Duplicates → Colour duplicates in selection**.

Expected: A1, B1, and A2 are yellow (ColorIndex 6) because `a` appears three times in the block. A single cell is rejected.

## Count duplicates in selection

Same block. **Duplicates → Count duplicates in selection**.

Expected: message **Number of duplicates in the selection is 3** (cells whose value appears more than once). Blanks are not counted. A single cell is rejected.
