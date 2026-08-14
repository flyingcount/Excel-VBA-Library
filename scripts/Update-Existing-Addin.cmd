@echo off
setlocal
REM Update the existing ExcelVbaLib.xlam in this repo (does not create a new add-in).
REM Excel must not have the file open.

set "REPO=%~dp0.."
set "XLAM=%REPO%\build\ExcelVbaLib.xlam"
if exist "C:\Users\chanp\OneDrive\Notebooks\Cursor\Excel-VBA-Library\build\ExcelVbaLib.xlam" (
  set "XLAM=C:\Users\chanp\OneDrive\Notebooks\Cursor\Excel-VBA-Library\build\ExcelVbaLib.xlam"
)

echo Closing Excel so %XLAM% can be written...
taskkill /F /IM EXCEL.EXE >nul 2>&1
timeout /t 2 /nobreak >nul

where python >nul 2>&1
if errorlevel 1 (
  echo Python not found. Install Python, then: pip install pyOpenVBA
  echo Then run: python "%REPO%\scripts\Build-ExcelVbaLib.py" "%XLAM%"
  exit /b 1
)

python -c "import pyopenvba" >nul 2>&1
if errorlevel 1 (
  echo Installing pyOpenVBA...
  python -m pip install pyOpenVBA
)

echo Updating existing add-in:
echo   %XLAM%
python "%REPO%\scripts\Build-ExcelVbaLib.py" "%XLAM%"
if errorlevel 1 (
  echo Update failed. If the file is locked, quit Excel from Task Manager and retry.
  exit /b 1
)

echo Saved. Start Excel, load that add-in, retry Data - Probability distributions - Poisson.
endlocal
