# Replace modules in ExcelVbaLib.xlam from the source snapshot.
# Targets the add-in *file* (default: <repo>\build\ExcelVbaLib.xlam).
# git pull does not update that .xlam — it is gitignored.
#
# Refresh the whole add-in in place (Internal, Api, Menus + ThisWorkbook events):
#   powershell -ExecutionPolicy Bypass -File scripts/Import-AddinModules.ps1 -All
#
# Data modules only (legacy default):
#   powershell -ExecutionPolicy Bypass -File scripts/Import-AddinModules.ps1
#
# Named modules (Internal first when they depend on each other):
#   powershell -ExecutionPolicy Bypass -File scripts/Import-AddinModules.ps1 -Modules modInternalMatrices,modApiMatrices,modAddinMenu
#
# Other copy of the add-in:
#   powershell -ExecutionPolicy Bypass -File scripts/Import-AddinModules.ps1 -All -XlamPath "C:\path\to\ExcelVbaLib.xlam"

[CmdletBinding()]
param(
    [string]$RepoRoot,
    [string]$XlamPath,
    [string[]]$Modules = @("modInternalData", "modApiData"),
    [switch]$All
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $here = $PSScriptRoot
    if ([string]::IsNullOrWhiteSpace($here)) {
        $here = Split-Path -Parent $MyInvocation.MyCommand.Path
    }
    $RepoRoot = (Resolve-Path (Join-Path $here "..")).Path
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

function Get-AllModuleItems {
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($root in $searchRoots) {
        if (-not (Test-Path $root)) { continue }
        foreach ($ext in @("*.bas", "*.cls", "*.frm")) {
            Get-ChildItem -Path $root -File -Filter $ext -ErrorAction SilentlyContinue |
                Where-Object { $_.BaseName -ne "ThisWorkbook" } |
                Sort-Object FullName |
                ForEach-Object {
                    $list.Add([pscustomobject]@{ Name = $_.BaseName; Path = $_.FullName })
                }
        }
    }
    if ($list.Count -eq 0) {
        throw "No .bas/.cls/.frm files found under source/Internal, Api, or Menus."
    }
    return $list
}

function Get-ThisWorkbookCode {
    param([string]$EventsPath)
    $code = [System.IO.File]::ReadAllText($EventsPath)
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

function Set-ThisWorkbookEvents {
    param($Workbook, [string]$EventsPath)
    if (-not (Test-Path $EventsPath)) { return }
    $code = Get-ThisWorkbookCode $EventsPath
    $cm = $Workbook.VBProject.VBComponents.Item("ThisWorkbook").CodeModule
    if ($cm.CountOfLines -gt 0) {
        $cm.DeleteLines(1, $cm.CountOfLines)
    }
    $cm.AddFromString($code)
    $foundOpen = $false
    for ($i = 1; $i -le $cm.CountOfLines; $i++) {
        if ($cm.Lines($i, 1) -match "Sub Workbook_Open") { $foundOpen = $true; break }
    }
    if (-not $foundOpen) {
        throw "ThisWorkbook is missing Workbook_Open after inject."
    }
    Write-Host "  wrote ThisWorkbook events ($($cm.CountOfLines) lines)"
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

function Write-PublicProcs {
    param($Component)
    if ($null -eq $Component) { return }
    $cm = $Component.CodeModule
    Write-Host "Public entry points in $($Component.Name):"
    $count = 0
    for ($i = 1; $i -le $cm.CountOfLines; $i++) {
        $line = $cm.Lines($i, 1)
        if ($line -match '^\s*Public\s+(Sub|Function)\s+(\w+)') {
            Write-Host ("  {0,-8} {1}" -f $Matches[1], $Matches[2])
            $count++
        }
    }
    if ($count -eq 0) {
        Write-Warning "$($Component.Name) has no Public Sub/Function lines after import."
    } else {
        Write-Host "  ($count public names)"
    }
}

function Assert-RequiredComponents {
    param($Workbook, [string[]]$Names, [switch]$Required)
    $missing = New-Object System.Collections.Generic.List[string]
    foreach ($name in $Names) {
        $comp = $null
        try { $comp = $Workbook.VBProject.VBComponents.Item($name) } catch { $comp = $null }
        if ($null -eq $comp) { $missing.Add($name) }
    }
    if ($missing.Count -eq 0) { return }
    $msg = "Add-in is missing: $($missing -join ', '). Run with -All to import Internal + Api + Menus (including Matrices)."
    if ($Required) {
        throw $msg
    }
    Write-Warning $msg
}

if ($All) {
    $toImport = Get-AllModuleItems
} else {
    Write-Warning "Without -All this updates only: $($Modules -join ', '). Matrix UDFs live in modApiMatrices; use -All to refresh the whole add-in."
    if ($null -eq $Modules -or $Modules.Count -eq 0) {
        throw "Pass -All, or at least one module name, e.g. -Modules modInternalMatrices,modApiMatrices,modAddinMenu"
    }
    foreach ($modName in $Modules) {
        if ($modName -eq "ThisWorkbook") {
            throw "ThisWorkbook is the add-in document module. Use -All or scripts/Inject-ThisWorkbook.ps1."
        }
    }
    $toImport = New-Object System.Collections.Generic.List[object]
    foreach ($modName in $Modules) {
        $path = Get-ModuleSourcePath $modName
        $toImport.Add([pscustomobject]@{ Name = $modName; Path = $path })
    }
}

if ([string]::IsNullOrWhiteSpace($XlamPath)) {
    $XlamPath = Join-Path $RepoRoot "build\ExcelVbaLib.xlam"
}
$XlamPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($XlamPath)
if (Test-Path -LiteralPath $XlamPath) {
    $XlamPath = (Resolve-Path -LiteralPath $XlamPath).Path
}

Write-Host "Will update $XlamPath"
$toImport | ForEach-Object { Write-Host "  $($_.Name)" }

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
        Write-Host "  imported $($item.Name)"
    }

    if ($All) {
        Set-ThisWorkbookEvents -Workbook $addinWb -EventsPath (Join-Path $RepoRoot "source\Menus\ThisWorkbook.cls")
    }

    try {
        $excel.Visible = $true
        $excel.VBE.MainWindow.Visible = $true
        $compileCtrl = $excel.VBE.CommandBars.FindControl(1, 578)
        if ($null -eq $compileCtrl) { $compileCtrl = $excel.VBE.CommandBars.FindControl($null, 578) }
        if ($null -ne $compileCtrl) {
            $compileCtrl.Execute()
            Write-Host "Compiled VBA project."
        } else {
            Write-Warning "Could not find Compile VBAProject (id 578). Compile manually in the VBE."
        }
        if ($createdExcel) {
            $excel.VBE.MainWindow.Visible = $false
            $excel.Visible = $false
        }
    } catch {
        Write-Warning "Compile step failed: $($_.Exception.Message)"
    }

    $excel.DisplayAlerts = $false
    $addinWb.Save()
    if (-not $createdExcel) { $excel.DisplayAlerts = $true }
    Write-Host "Saved $($addinWb.FullName)"

    Assert-RequiredComponents -Workbook $addinWb -Names @("modInternalMatrices", "modApiMatrices", "modAddinMenu") -Required:$All
    try {
        Write-PublicProcs $addinWb.VBProject.VBComponents.Item("modApiMatrices")
    } catch {
        Write-Warning "Could not list modApiMatrices public names: $($_.Exception.Message)"
    }

    if (-not $createdExcel) {
        try {
            $excel.Run("'" + $addinWb.Name + "'!InstallExcelVbaLibMenu")
            Write-Host "Rebuilt the Excel VBA Lib menu for this session."
        } catch {
            Write-Warning "Could not run InstallExcelVbaLibMenu: $($_.Exception.Message). Restart Excel."
        }
        try {
            $excel.Run("'" + $addinWb.Name + "'!RegisterMatrixUdfs")
            Write-Host "Registered matrix worksheet functions (Insert Function category Excel VBA Lib)."
        } catch {
            Write-Warning "Could not register matrix UDFs: $($_.Exception.Message). Restart Excel with the add-in loaded."
        }
    } else {
        $addinWb.Close($true)
        Write-Host "Closed Excel. Restart Excel (add-in loaded via Excel Add-ins) so the menu appears."
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
