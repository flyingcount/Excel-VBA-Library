# Replace modules in ExcelVbaLib.xlam from the source snapshot.
# Default: Data modules (modInternalData, then modApiData).
#
# Targets the add-in file (not "whatever ExcelVbaLib.xlam is loaded").
# Pass -XlamPath when the file is not build\ExcelVbaLib.xlam under the repo.
#
#   powershell -ExecutionPolicy Bypass -File scripts/Import-AddinModules.ps1
#   powershell -ExecutionPolicy Bypass -File scripts/Import-AddinModules.ps1 -XlamPath "C:\path\to\build\ExcelVbaLib.xlam"
#
# Other modules (Internal first when they depend on each other):
#   powershell -ExecutionPolicy Bypass -File scripts/Import-AddinModules.ps1 -Modules modInternalBenford,modApiBenford
#
# ThisWorkbook is the document module — use Inject-ThisWorkbook.ps1 for that.

[CmdletBinding()]
param(
    [string]$RepoRoot,
    [string]$XlamPath,
    [string[]]$Modules = @("modInternalData", "modApiData")
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $here = $PSScriptRoot
    if ([string]::IsNullOrWhiteSpace($here)) {
        $here = Split-Path -Parent $MyInvocation.MyCommand.Path
    }
    $RepoRoot = (Resolve-Path (Join-Path $here "..")).Path
}

if ($null -eq $Modules -or $Modules.Count -eq 0) {
    throw "Pass at least one module name, e.g. -Modules modInternalData,modApiData"
}

foreach ($modName in $Modules) {
    if ($modName -eq "ThisWorkbook") {
        throw "ThisWorkbook is the add-in document module. Use scripts/Inject-ThisWorkbook.ps1 instead of Import."
    }
}

$searchRoots = @(
    (Join-Path $RepoRoot "source\Internal"),
    (Join-Path $RepoRoot "source\Api"),
    (Join-Path $RepoRoot "source\Menus")
)

function Get-ModuleSourcePath {
    param([string]$ModName)
    foreach ($root in $searchRoots) {
        foreach ($ext in @(".bas", ".cls", ".frm")) {
            $p = Join-Path $root ($ModName + $ext)
            if (Test-Path -LiteralPath $p) { return $p }
        }
    }
    throw "No source file for '$ModName' under source/Internal, Api, or Menus."
}

function Test-SamePath([string]$A, [string]$B) {
    if ([string]::IsNullOrWhiteSpace($A) -or [string]::IsNullOrWhiteSpace($B)) { return $false }
    return [string]::Equals(
        $A.TrimEnd('\', '/'),
        $B.TrimEnd('\', '/'),
        [StringComparison]::OrdinalIgnoreCase)
}

function Get-OpenWorkbookByPath {
    param($Excel, [string]$TargetPath)
    foreach ($wb in @($Excel.Workbooks)) {
        $full = $null
        try { $full = [string]$wb.FullName } catch { continue }
        if (Test-SamePath $full $TargetPath) { return $wb }
    }
    return $null
}

$toImport = New-Object System.Collections.Generic.List[object]
foreach ($modName in $Modules) {
    $path = Get-ModuleSourcePath $modName
    $toImport.Add([pscustomobject]@{ Name = $modName; Path = $path })
}

if ([string]::IsNullOrWhiteSpace($XlamPath)) {
    $XlamPath = Join-Path $RepoRoot "build\ExcelVbaLib.xlam"
}
$XlamPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($XlamPath)
if (Test-Path -LiteralPath $XlamPath) {
    $XlamPath = (Resolve-Path -LiteralPath $XlamPath).Path
}

$excel = $null
$createdExcel = $false
try {
    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
} catch {
    $excel = $null
}

if ($null -eq $excel) {
    if (-not (Test-Path -LiteralPath $XlamPath)) {
        throw @"
Excel is not running and the add-in was not found:
  $XlamPath
Open Excel with that add-in loaded, or pass -XlamPath, or build it first:
  powershell -ExecutionPolicy Bypass -File scripts/Build-ExcelVbaLib.ps1
"@
    }
    $excel = New-Object -ComObject Excel.Application
    $createdExcel = $true
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
}

try {
    $addinWb = Get-OpenWorkbookByPath -Excel $excel -TargetPath $XlamPath
    if ($null -eq $addinWb) {
        if (-not (Test-Path -LiteralPath $XlamPath)) {
            throw "Add-in not open and file not found: $XlamPath"
        }
        Write-Host "Opening $XlamPath"
        $addinWb = $excel.Workbooks.Open($XlamPath)
    }

    try {
        $null = $addinWb.VBProject.Name
    } catch {
        throw @"
Programmatic access to the VBA project failed.
Enable: File → Options → Trust Center → Trust Center Settings → Macro Settings
→ Trust access to the VBA project object model
Then re-run this script.
"@
    }

    Write-Host "Updating $($addinWb.FullName)"

    foreach ($item in $toImport) {
        $comp = $null
        try { $comp = $addinWb.VBProject.VBComponents.Item($item.Name) } catch { $comp = $null }
        if ($null -ne $comp) {
            $addinWb.VBProject.VBComponents.Remove($comp)
            Write-Host "  removed $($item.Name)"
        }
        [void]$addinWb.VBProject.VBComponents.Import($item.Path)
        Write-Host "  imported $($item.Name) from $($item.Path)"
    }

    $excel.DisplayAlerts = $false
    $addinWb.Save()
    if (-not $createdExcel) { $excel.DisplayAlerts = $true }
    Write-Host "Saved $($addinWb.FullName)"

    if ($createdExcel) {
        $addinWb.Close($true)
        Write-Host "Closed Excel. Load that add-in and retry Poisson."
    } else {
        Write-Host "Retry Excel VBA Lib → Data → Probability distributions → Poisson."
    }
} finally {
    if ($createdExcel -and $null -ne $excel) {
        try { $excel.Quit() } catch { }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        $excel = $null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}
