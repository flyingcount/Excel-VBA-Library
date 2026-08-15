#!/usr/bin/env bash
# This add-in can only be compiled by Microsoft Excel on Windows.
# If you ran git-pull / Import-AddinModules in the cloud Linux terminal, that
# cannot update build/ExcelVbaLib.xlam (the file is gitignored and there is no Excel here).
set -euo pipefail
cat <<'EOF'
Import-AddinModules.ps1 must run in Windows PowerShell on the PC that has Excel,
not in this Linux workspace.

1. On your Windows PC, open PowerShell (Start → Windows PowerShell).
2. Run:

   cd "C:\Users\chanp\OneDrive\Notebooks\Cursor\Excel-VBA-Library"
   git pull origin add-create-date-table
   powershell -ExecutionPolicy Bypass -File .\scripts\Import-AddinModules.ps1 -All

3. Quit Excel fully and start it again.

If that folder is a different clone, cd to the clone that contains
build\ExcelVbaLib.xlam, then run the same Import-AddinModules.ps1 -All line.
EOF
exit 1
