@echo off
REM CMD wrapper. Prefer PowerShell:
REM   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
REM   .\scripts\Update-Existing-Addin.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Update-Existing-Addin.ps1"
if errorlevel 1 exit /b 1
