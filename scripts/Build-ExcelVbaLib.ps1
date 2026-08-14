# Build ExcelVbaLib.xlam
# Packs curated source into one add-in so caller workbooks load a single library.
# Requires: Excel, and Trust access to the VBA project object model.
#
# From repo root:
#   powershell -ExecutionPolicy Bypass -File scripts/Build-ExcelVbaLib.ps1

[CmdletBinding()]
param(
    [string]$RepoRoot
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $here = $PSScriptRoot
    if ([string]::IsNullOrWhiteSpace($here)) {
        $here = Split-Path -Parent $MyInvocation.MyCommand.Path
    }
    $RepoRoot = (Resolve-Path (Join-Path $here "..")).Path
}

$buildDir = Join-Path $RepoRoot "build"
$xlamPath = Join-Path $buildDir "ExcelVbaLib.xlam"

# Import order: Internal first (call graph), then Api, then Menus. Never _export_raw.
# New library code belongs in ExcelVbaLib.xlam, not extra source folders.
$importRoots = @(
    (Join-Path $RepoRoot "source\Internal"),
    (Join-Path $RepoRoot "source\Api"),
    (Join-Path $RepoRoot "source\Menus")
)

$extensions = @("*.bas", "*.cls", "*.frm")

function Get-ImportFiles {
    $files = New-Object System.Collections.Generic.List[string]
    foreach ($root in $importRoots) {
        if (-not (Test-Path $root)) { continue }
        foreach ($ext in $extensions) {
            Get-ChildItem -Path $root -Recurse -File -Filter $ext -ErrorAction SilentlyContinue |
                Sort-Object FullName |
                ForEach-Object {
                    # ThisWorkbook is the add-in document module; inject later, do not Import.
                    if ($_.BaseName -eq "ThisWorkbook") { return }
                    $files.Add($_.FullName)
                }
        }
    }
    return $files
}

function Set-ThisWorkbookEvents {
    param($Workbook, [string]$EventsPath)
    if (-not (Test-Path $EventsPath)) { return }
    $code = Get-Content -LiteralPath $EventsPath -Raw
    $cm = $Workbook.VBProject.VBComponents.Item("ThisWorkbook").CodeModule
    if ($cm.CountOfLines -gt 0) {
        $cm.DeleteLines(1, $cm.CountOfLines)
    }
    $cm.AddFromString($code)
    Write-Host "Wrote ThisWorkbook Open/BeforeClose events."
}

$toImport = Get-ImportFiles
if ($toImport.Count -eq 0) {
    throw "No .bas/.cls/.frm files found under source/Internal, Api, or Menus."
}

Write-Host "Will import $($toImport.Count) component(s):"
$toImport | ForEach-Object { Write-Host "  $_" }

New-Item -ItemType Directory -Force -Path $buildDir | Out-Null

if (Test-Path $xlamPath) {
    try {
        Remove-Item -LiteralPath $xlamPath -Force
    } catch {
        throw "Cannot overwrite $xlamPath. Close Excel if ExcelVbaLib is loaded, then retry."
    }
}

$excel = $null
$wb = $null
$tempXlsx = Join-Path $env:TEMP ("ExcelVbaLib-build-{0}.xlsx" -f [guid]::NewGuid().ToString("N"))

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false

    $wb = $excel.Workbooks.Add()

    try {
        $null = $wb.VBProject.Name
    } catch {
        throw @"
Programmatic access to the VBA project failed.
Enable: File → Options → Trust Center → Trust Center Settings → Macro Settings
→ Trust access to the VBA project object model
Then close Excel and re-run this script.
"@
    }

    $wb.VBProject.Name = "ExcelVbaLib"

    foreach ($file in $toImport) {
        Write-Host "Importing $file"
        [void]$wb.VBProject.VBComponents.Import($file)
    }

    Set-ThisWorkbookEvents -Workbook $wb -EventsPath (Join-Path $RepoRoot "source\Menus\ThisWorkbook.cls")

    # Compile VBA project (command bar control 578). VBE must be reachable.
    try {
        $excel.Visible = $true
        $excel.VBE.MainWindow.Visible = $true
        $compileCtrl = $excel.VBE.CommandBars.FindControl(1, 578)
        if ($null -eq $compileCtrl) {
            $compileCtrl = $excel.VBE.CommandBars.FindControl($null, 578)
        }
        if ($null -ne $compileCtrl) {
            $compileCtrl.Execute()
            Write-Host "Compiled VBA project."
        } else {
            Write-Warning "Could not find Compile command; open the .xlam and compile manually (Debug → Compile VBAProject)."
        }
        $excel.VBE.MainWindow.Visible = $false
        $excel.Visible = $false
    } catch {
        Write-Warning "Compile step failed: $($_.Exception.Message). Open the .xlam and compile manually."
    }

    # 55 = xlOpenXMLAddIn (.xlam)
    $wb.SaveAs($xlamPath, 55)
    Write-Host "Saved $xlamPath"
} finally {
    if ($null -ne $wb) {
        try { $wb.Close($false) } catch { }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch { }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    $wb = $null
    $excel = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if (Test-Path $tempXlsx) { Remove-Item $tempXlsx -Force -ErrorAction SilentlyContinue }
}

if (-not (Test-Path $xlamPath)) {
    throw "Build finished but $xlamPath was not created."
}

Write-Host ""
Write-Host "Load the add-in: Excel -> Options -> Add-ins -> Excel Add-ins -> Browse -> build\ExcelVbaLib.xlam"
Write-Host "Then run Benford from the Excel VBA Lib menu, or:"
Write-Host '  Application.Run "BenfordAnalysisFirstDigit", Range("A2:A50")'
