#!/usr/bin/env bash
# Overlay Personal Menu13 Matrices1/2 into ExcelVbaLib.xlam. Excel is Windows-only.
set -euo pipefail
cat <<'EOF'
This Linux workspace cannot read Personal.xlsb or write ExcelVbaLib.xlam.

On your Windows PC, in Windows PowerShell (not this terminal):

   cd "C:\Users\chanp\OneDrive\Notebooks\Cursor\Excel-VBA-Library"
   git pull origin add-create-date-table
   powershell -ExecutionPolicy Bypass -File .\scripts\Import-AddinModules.ps1 -All
   powershell -ExecutionPolicy Bypass -File .\scripts\Import-Menu13FromPersonal.ps1

That last script copies Custom_Menu13_Matrices1 and Custom_Menu13_Matrices2
from Data\Personal.xlsb (or Personal123.xlsb) into build\ExcelVbaLib.xlam.

Whole Menu13 family:

   powershell -ExecutionPolicy Bypass -File .\scripts\Import-Menu13FromPersonal.ps1 -AllMenu13

Then quit Excel fully and start it again. Compile the add-in in the VBE.
EOF
exit 1
