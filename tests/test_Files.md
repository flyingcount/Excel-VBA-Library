# Test: Files (Personal Menu24)

Personal `Custom_Menu24_ListFilesInFolder`, rewritten as `modApiFiles` + `modInternalFiles`.

## Setup

1. Load `build/ExcelVbaLib.xlam`. Fully quit Excel and restart if the menu was just rebuilt.
2. Use **Excel VBA Lib → Files**. Do not import Personal Menu24 modules into the test workbook.

## List files in folder

Create a folder with two files (for example `a.txt` and `b.csv`) and one subfolder that also contains a file.

**Files → List files in folder**, pick that folder.

Expected: sheet **File listing** with a header row at A1: Name, Path, Size(Bytes), Type, Created, Last Accessed, Last Modified, Parent Folder, Short Name, Short Path, File Attribute Value, File Attribute Description.

- Two data rows (the subfolder file is omitted). No count-and-estimate prompt (fewer than 1,000 files).
- Size is a number; dates use `yyyy-mm-dd hh:mm:ss`.
- Autofilter is on; row 1 is frozen.
- Cancel on the folder picker writes nothing and does not create the sheet.
- A modeless **Files** dialog stays open while counting and listing. **Cancel** (or Esc) stops the run. If any file details were already collected, they are written to **File listing** with a **Partial listing (cancelled)** title. Cancel during the count (before details) writes nothing.

Personal `ReturnFilesInSelectedFolder` wrote only folder + name at a selected cell and called `SelectFolder` / `GetRange` (those helpers are not in Personal). This version uses the same columns as the recursive listing.

## List files in folder and subfolders

**Files → List files in folder and subfolders**, pick the same folder.

Expected: three data rows (parent files plus the file in the subfolder). Hidden, system, and junction folders are skipped. No confirm prompt under 1,000 files.

If the count is 1,000 or more, a prompt reports the file count, how long counting took, and an estimated listing time. **Yes** continues; **No** writes nothing and does not create the sheet. For more than 10,000 files, **No** is the default button.

Personal `BatchListAllFiles_FolderSubfolders` reset its row counter in each recursive call, so subfolder files overwrote earlier rows starting at row 2. It also wrote the header and AutoFit on every folder, created the output sheet before the picker result was checked, and decoded attributes only when the value matched a single Case (so Archive+Hidden = 34 was `Unknown`).

A folder with no files shows a message and does not keep a blank **File listing** sheet.
