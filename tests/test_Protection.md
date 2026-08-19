# Test: Protection (Personal Menu7)

Personal `Custom_Menu7_Protection`, rewritten as `modApiProtection` + `modInternalProtection`.

The default password is the Personal constant `WYSIWYG`. It is a **known convenience value**, not a secret. **Protection → Show default password** displays it.

## Setup

1. Load `build/ExcelVbaLib.xlam`. Fully quit Excel and restart if the menu was just rebuilt.
2. Use **Excel VBA Lib → Protection**. Do not import Personal `Custom_Menu7_Protection` into the test workbook.
3. Work in a throwaway workbook. Chart sheets are rejected (Personal used `ActiveSheet` and could fail).

## Limit / reset scroll area

On a worksheet, select `A1:D20`. **Protection → Limit scroll area to selection**.

Expected: you cannot move the active cell outside `$A$1:$D$20`. Only the first area of a multi-area selection is used.

**Protection → Reset scroll area**.

Expected: the whole sheet is scrollable again.

If the selection is not a range (for example a shape), a message box says to select a range.

## Protect / unprotect (default password)

**Protection → Protect worksheet (default password)**.

Expected: the sheet is protected. Typing in cells is blocked. VBA can still write this session (`UserInterfaceOnly:=True`; that flag is lost when the workbook is reopened). Personal called `Protect` even if the sheet was already protected (error). This version unprotects with the default password first when that works.

**Protection → Unprotect worksheet (default password)**.

Expected: the sheet is editable. A sheet protected with a different password shows a message instead of a VBA error.

## Show default password

**Protection → Show default password**.

Expected: a message box `The password is : WYSIWYG`. Personal assigned the `MsgBox` result to an unused `Long`.

## Unhide all worksheets

Hide one sheet (right-click → Hide) and very-hide another in the VBE (`xlSheetVeryHidden`). **Protection → Unhide all worksheets**.

Expected: both become visible; a message reports how many. If none were hidden: `There are no hidden worksheets in …`. Chart sheets are not in the Worksheets collection (same as Personal).

## Unhide all rows and columns

Hide row 5 and column C. **Protection → Unhide all rows and columns**.

Expected: row 5 and column C are visible. The selection is unchanged (Personal did `Cells.Select`). If the sheet is protected, a message says to unprotect first.
