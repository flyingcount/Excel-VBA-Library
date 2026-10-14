# Test: Custom lists (Personal Menu16)

Personal `Custom_Menu16_CustomLists` / `Custom_Menu16_Autocorrect`, rewritten as `modApiCustomLists` + `modInternalCustomLists`.

Custom lists and AutoCorrect replacements are stored in **Excel**, not in the workbook. Tests that create or delete lists change this machine's Excel settings. Use a throwaway list (for example `ZzzListA`, `ZzzListB`, `ZzzListC`) and delete it when finished.

## Setup

1. Load `build/ExcelVbaLib.xlam`. Fully quit Excel and restart if the menu was just rebuilt.
2. Use **Excel VBA Lib → Custom lists**. Do not import Personal Menu16 modules into the test workbook.

## Count

**Custom lists → Count**

Expected: a message box with the total, 4 built-in (days / months), and the user-defined remainder. Notes that lists are stored in Excel.

## List on worksheet

**Custom lists → List on worksheet**

Expected: sheet **Custom List properties** with a title in A1 and a header row at A3: Custom list number, Type, Item count, then elements to the right.

- Lists 1–4 are Type `Built-in` (Sun–Sat, Sunday–Saturday, Jan–Dec, January–December).
- User lists are Type `User`.
- Column A is the list number used by **Delete by list number**.

Personal wrote only the list number and elements (no Type / Item count), used `GetCustomListNum` after `GetCustomListContents`, and `WorksheetFunction.Transpose` on each scalar item.

## Create from columns

On a blank sheet:

```
A1: ZzzListA
A2: ZzzListB
A3: ZzzListC
```

Select `A1:A3`. **Custom lists → Create from columns**. Accept the default range.

Expected: `1 custom list created (column-wise).` Cancel on the InputBox creates nothing.

Run **Create from columns** again on the same range.

Expected: `0 custom lists created (column-wise).` and `1 already existed and was skipped.` Personal `CreateCustomListByColumn` raised (or could raise) if the list already existed; the row-wise Personal routine swallowed that error.

A single-cell selection reports that a list cannot be created from one cell.

Two columns of three values each create two lists (blank cells inside a column are skipped; a column with fewer than two values is skipped).

## Create from rows

```
A5: ZzzRow1
B5: ZzzRow2
C5: ZzzRow3
```

Select `A5:C5`. **Custom lists → Create from rows**.

Expected: one new list of those three items.

## Delete by list number

1. **List on worksheet** so column A has current numbers.
2. Note the number of the `ZzzListA` list (must be > 4).
3. Select that number. **Custom lists → Delete by list number**. Confirm Yes.

Expected: that user list is gone; the inventory sheet is rewritten. Built-in numbers 1–4 in the selection are skipped with a message. Cancel on the confirm prompt deletes nothing.

If several user list numbers are selected, they are deleted **highest first**. Personal deleted in sheet order, so deleting 5 then 6 could remove the list that had just become 5.

## AutoCorrect list

**Custom lists → AutoCorrect list**

Expected: sheet **Auto correct List** (Personal spelling) with title in A1 and Replace: / With: from row 3.

## Add AutoCorrect entries

```
A10: Replace:
B10: With:
A11: zzzteh
B11: zzzthe
```

Select `A10:B11`. **Custom lists → Add AutoCorrect entries**. Accept the default range.

Expected: `1 AutoCorrect entry added or updated.` The header row is skipped. Personal `AutoCorrectEntries_Add` ignored the selection and scanned **column A of the worksheet** (`Columns(1).SpecialCells`).

Type `zzzteh` in a cell; Excel should offer `zzzthe`. Remove the test entry afterwards (File → Options → Proofing → AutoCorrect Options) if you do not want it kept.
