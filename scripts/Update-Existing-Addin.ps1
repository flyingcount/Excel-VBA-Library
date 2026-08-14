# Update the existing ExcelVbaLib.xlam using Excel (no Python).
# Run from any folder in PowerShell (not cmd, not C:\windows\System32 relative paths):
#
#   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
#   & "C:\Users\chanp\OneDrive\Notebooks\Cursor\Excel-VBA-Library\scripts\Update-Existing-Addin.ps1"
#
# Requires: Excel, and Trust access to the VBA project object model.

[CmdletBinding()]
param(
    [string]$XlamPath
)

$ErrorActionPreference = "Stop"

$here = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($here)) {
    $here = Split-Path -Parent $MyInvocation.MyCommand.Path
}
$repo = (Resolve-Path (Join-Path $here "..")).Path
$importer = Join-Path $here "Import-AddinModules.ps1"
if (-not (Test-Path -LiteralPath $importer)) {
    throw "Missing $importer"
}

if ([string]::IsNullOrWhiteSpace($XlamPath)) {
    $candidates = @(
        "C:\Users\chanp\OneDrive\Notebooks\Cursor\Excel-VBA-Library\build\ExcelVbaLib.xlam",
        (Join-Path $repo "build\ExcelVbaLib.xlam")
    )
    foreach ($c in $candidates) {
        if (Test-Path -LiteralPath $c) {
            $XlamPath = $c
            break
        }
    }
}

if ([string]::IsNullOrWhiteSpace($XlamPath) -or -not (Test-Path -LiteralPath $XlamPath)) {
    throw "Existing add-in not found. Expected build\ExcelVbaLib.xlam under the repo."
}

$src = Join-Path $repo "source\Internal\modInternalData.bas"
if (-not (Test-Path -LiteralPath $src)) {
    throw "Missing $src — pull the repo so the Poisson fix is in source."
}
$srcText = [System.IO.File]::ReadAllText($src)
if ($srcText -match "WorksheetFunction\.Poisson_Inv\(") {
    throw "source\Internal\modInternalData.bas still calls Poisson_Inv. Pull the latest source first."
}
if ($srcText -notmatch "Function RandomPoisson\(") {
    throw "source\Internal\modInternalData.bas is missing RandomPoisson. Pull the latest source first."
}

Write-Host "Will update existing add-in:"
Write-Host "  $XlamPath"
Write-Host "from source (Excel COM, not Python)."

& $importer -RepoRoot $repo -XlamPath $XlamPath -Modules @("modInternalData", "modApiData")
