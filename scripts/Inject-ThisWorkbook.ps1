# Inject ThisWorkbook events into the loaded ExcelVbaLib add-in.
# Excel add-ins do not run Auto_Open. Menu install lives in ThisWorkbook
# (Workbook_Open / AddinInstall) and must be written into the document module,
# not imported as a second class.
#
# With Excel open and ExcelVbaLib.xlam loaded:
#   powershell -ExecutionPolicy Bypass -File scripts/Inject-ThisWorkbook.ps1
#
# Then restart Excel (or wait a second) so Workbook_Open can schedule the menu.

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

$eventsPath = Join-Path $RepoRoot "source\Menus\ThisWorkbook.cls"
$menuPath = Join-Path $RepoRoot "source\Menus\modAddinMenu.bas"
if (-not (Test-Path -LiteralPath $eventsPath)) { throw "Missing $eventsPath" }
if (-not (Test-Path -LiteralPath $menuPath)) { throw "Missing $menuPath" }

function Get-ThisWorkbookCode([string]$Path) {
    $code = [System.IO.File]::ReadAllText($Path)
    if ($code.Length -gt 0 -and [int][char]$code[0] -eq 0xFEFF) {
        $code = $code.Substring(1)
    }
    $out = New-Object System.Collections.Generic.List[string]
    $inHeader = $false
    foreach ($line in ($code -split "`r?`n")) {
        if ($line -match '^VERSION\s+\d') { $inHeader = $true; continue }
        if ($inHeader -and $line -match '^(BEGIN|END)\b') {
            if ($line -match '^END\b') { $inHeader = $false }
            continue
        }
        if ($line -match '^Attribute\s+VB_') { continue }
        $out.Add($line)
    }
    return ($out -join "`r`n").Trim() + "`r`n"
}

$excel = $null
try {
    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
} catch {
    throw "Excel is not running. Open Excel with ExcelVbaLib.xlam loaded, then re-run."
}

$addinWb = $null
try { $addinWb = $excel.Workbooks.Item("ExcelVbaLib.xlam") } catch { $addinWb = $null }
if ($null -eq $addinWb) {
    throw "ExcelVbaLib.xlam is not an open workbook. Load it via File → Options → Add-ins → Excel Add-ins."
}

Write-Host "Updating $($addinWb.FullName)"

$comp = $null
try { $comp = $addinWb.VBProject.VBComponents.Item("modAddinMenu") } catch { $comp = $null }
if ($null -ne $comp) {
    $addinWb.VBProject.VBComponents.Remove($comp)
    Write-Host "  removed modAddinMenu"
}
[void]$addinWb.VBProject.VBComponents.Import($menuPath)
Write-Host "  imported modAddinMenu"

$code = Get-ThisWorkbookCode $eventsPath
$cm = $addinWb.VBProject.VBComponents.Item("ThisWorkbook").CodeModule
if ($cm.CountOfLines -gt 0) {
    $cm.DeleteLines(1, $cm.CountOfLines)
}
$cm.AddFromString($code)
Write-Host "  wrote ThisWorkbook events ($($cm.CountOfLines) lines)"

$foundOpen = $false
for ($i = 1; $i -le $cm.CountOfLines; $i++) {
    if ($cm.Lines($i, 1) -match "Sub Workbook_Open") { $foundOpen = $true; break }
}
if (-not $foundOpen) { throw "ThisWorkbook is missing Workbook_Open after inject." }

$excel.DisplayAlerts = $false
$addinWb.Save()
$excel.DisplayAlerts = $true
Write-Host "Saved."

try {
    $excel.Run("'" + $addinWb.Name + "'!InstallExcelVbaLibMenu")
    Write-Host "Installed menu for this session."
} catch {
    Write-Warning "Could not run InstallExcelVbaLibMenu: $($_.Exception.Message). Restart Excel."
}

Write-Host "Restart Excel once so Workbook_Open runs on a cold start."
