# Test: Hyperlinks (Personal Menu21)

Personal `Custom_Menu21_Hyperlinks` / `Custom_Menu21_Index`, rewritten as `modApiHyperlinks` + `modInternalHyperlinks`.

## Setup

1. Load `build/ExcelVbaLib.xlam`. Fully quit Excel and restart if the menu was just rebuilt.
2. Use **Excel VBA Lib → Hyperlinks**. Do not import the Personal Menu21 modules into the test workbook.

## Inventory

In a workbook, put a hyperlink in `A1` (Insert → Link) to `https://example.com` with display text `Example`.

**Hyperlinks → Inventory**

Expected: sheet **Hyperlinks** with a title in A3 and a header row at A5. One data row: worksheet name, `A1`, `Example`, ScreenTip (may be empty), `https://example.com`, blank SubAddress, Type `Range`. Parent/Creator columns from Personal are dropped; ScreenTip and a readable Type are used instead.

If the workbook has no hyperlinks, a message box says so and no sheet is written.

## Index

On a workbook with at least two worksheets and no sheet named Index:

**Hyperlinks → Create index** — confirm Yes.

Expected:

- Sheet **Index** at the front, with hyperlinks to every worksheet (20 names per column).
- Each other unprotected sheet has a new row 1 with **Back to Index**.
- Protected sheets appear on Index but are not edited.
- Running Create index again reports that Index already exists.

**Hyperlinks → Update index** after adding a new sheet:

Expected: Index is rebuilt in place. Back-links on A1 are refreshed. No extra rows are inserted.

## Remove by text to display

With several hyperlinks that display `Example`:

1. Type `Example` in a cell (or select that text on the Hyperlinks inventory sheet).
2. **Hyperlinks → Remove by text to display**, accept the default range.

Expected: every workbook hyperlink whose Text to display is `Example` is deleted (case-insensitive). A count message appears. Cancel on the InputBox writes nothing. Deleting while looping the Hyperlinks collection (Personal) is avoided; links are removed from last to first.

## Open selected

Select the cell with a remaining web hyperlink. **Hyperlinks → Open selected**.

Expected: Excel follows the link (default browser). Personal Menu21 `OpenHyperlink` opened a RangeForm diagnostic instead; this library version follows the links. More than one selected link asks for confirmation first.

## Back-links to this sheet A1

Select cell `B2` on Sheet1. **Hyperlinks → Back-links to this sheet A1** — confirm Yes.

Expected: every other unprotected worksheet has a hyperlink in `B2` whose text is `Back to 'Sheet1'` and whose target is Sheet1!A1. Protected sheets are skipped. A count message appears.

## List worksheets

**Hyperlinks → List worksheets**

Expected: a message box with index, name, and `(hidden)` / `(very hidden)` when relevant. Personal passed `vbOKOnly` as the title; the title is now **Worksheets**.
