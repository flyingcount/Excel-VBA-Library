# Copy Personal Menu13 Matrices1 / Matrices2 (original VBA) into ExcelVbaLib.xlam.
# This cloud/Linux workspace cannot see Personal.xlsb. Run on the Windows PC.
#
# After git pull + Import-AddinModules -All (git snapshot), overlay the Personal
# modules, rewriting their VBA names to modApiMatrices1 / modApiMatrices2:
#
#   powershell -ExecutionPolicy Bypass -File scripts/Import-Menu13FromPersonal.ps1
#
# Whole Menu13 family (create, Cholesky, eigen, unitary, utilities, Fn_Matrices*):
#   powershell -ExecutionPolicy Bypass -File scripts/Import-Menu13FromPersonal.ps1 -AllMenu13
#
# Original Public procedure names may differ from the Excel VBA Lib menu OnAction
# names. If VBA compile reports a duplicate Public name (e.g. WriteArrayToWorksheet),
# make that helper Private in the imported Personal module.

[CmdletBinding()]
param(
    [string]$RepoRoot,
    [string]$WorkbookPath,
    [string]$XlamPath,
    [switch]$AllMenu13
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $here = $PSScriptRoot
    if ([string]::IsNullOrWhiteSpace($here)) {
        $here = Split-Path -Parent $MyInvocation.MyCommand.Path
    }
    $RepoRoot = (Resolve-Path (Join-Path $here "..")).Path
}

$personalNames = @(
    "Custom_Menu13_Matrices1",
    "Custom_Menu13_Matrices2"
)
if ($AllMenu13) {
    $personalNames = @(
        "Custom_Menu13_CreateMatrices",
        "Custom_Menu13_Matrices1",
        "Custom_Menu13_Matrices2",
        "Custom_Menu13_Cholesky",
        "Custom_Menu13_EigenDecomp",
        "custom_Menu13_Unitary",
        "Custom_Menu13_MatrixUtilities",
        "Fn_MatricesArray",
        "Fn_MatricesRng",
        "Fn_Matrices2"
    )
}

$moduleRename = [ordered]@{
    "Custom_Menu13_CreateMatrices" = "modApiMatrixCreate"
    "Custom_Menu13_MatrixUtilities" = "modApiMatrixUtilities"
    "Custom_Menu13_EigenDecomp" = "modApiEigenDecomp"
    "Custom_Menu13_Matrices1" = "modApiMatrices1"
    "Custom_Menu13_Matrices2" = "modApiMatrices2"
    "Custom_Menu13_Cholesky" = "modApiCholesky"
    "custom_Menu13_Unitary" = "modApiUnitary"
    "Custom_Menu18_Covariance" = "modApiCovariance"
}

function Convert-PersonalModuleSource {
    param([string]$Path, [string]$NewName)
    $text = [System.IO.File]::ReadAllText($Path)
    foreach ($old in $moduleRename.Keys) {
        $text = $text.Replace($old, [string]$moduleRename[$old])
    }
    $tmp = Join-Path $env:TEMP ($NewName + ".bas")
    [System.IO.File]::WriteAllText($tmp, $text)
    return $tmp
}

$outDir = Join-Path $RepoRoot "source\_export_raw"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

if ([string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $candidates = @(
        (Join-Path $RepoRoot "Data\Personal.xlsb"),
        (Join-Path $RepoRoot "Data\Personal123.xlsb")
    )
    $WorkbookPath = $candidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
}
if ([string]::IsNullOrWhiteSpace($WorkbookPath) -or -not (Test-Path -LiteralPath $WorkbookPath)) {
    throw "Personal workbook not found. Pass -WorkbookPath or put Personal.xlsb in Data\."
}

if ([string]::IsNullOrWhiteSpace($XlamPath)) {
    $XlamPath = Join-Path $RepoRoot "build\ExcelVbaLib.xlam"
}
$XlamPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($XlamPath)
if (-not (Test-Path -LiteralPath $XlamPath)) {
    throw "Add-in not found: $XlamPath. Run Import-AddinModules.ps1 -All first."
}
$XlamPath = (Resolve-Path -LiteralPath $XlamPath).Path

Write-Host "Exporting Menu13 from $WorkbookPath"
Write-Host "Into add-in $XlamPath"

$excel = $null
$createdExcel = $false
$personalWb = $null
$addinWb = $null
try {
    try {
        $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    } catch {
        $excel = New-Object -ComObject Excel.Application
        $createdExcel = $true
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
    }

    $personalWb = $excel.Workbooks.Open($WorkbookPath)
    try { $null = $personalWb.VBProject.Name } catch {
        throw "Enable Trust access to the VBA project object model, then retry."
    }

    $exported = New-Object System.Collections.Generic.List[object]
    foreach ($name in $personalNames) {
        $comp = $null
        try { $comp = $personalWb.VBProject.VBComponents.Item($name) } catch { $comp = $null }
        if ($null -eq $comp) {
            Write-Warning "Personal has no module '$name'."
            continue
        }
        $dest = Join-Path $outDir ($name + ".bas")
        if (Test-Path -LiteralPath $dest) { Remove-Item -LiteralPath $dest -Force }
        $comp.Export($dest)
        Write-Host "  exported $name ($((Get-Item -LiteralPath $dest).Length) bytes)"
        $exported.Add([pscustomobject]@{ Name = $name; Path = $dest })
    }
    if ($exported.Count -eq 0) {
        throw "No Menu13 modules exported from Personal."
    }

    $addinWb = $null
    foreach ($wb in @($excel.Workbooks)) {
        $full = $null
        try { $full = [string]$wb.FullName } catch { continue }
        if ([string]::Equals($full.TrimEnd('\', '/'), $XlamPath.TrimEnd('\', '/'), [StringComparison]::OrdinalIgnoreCase)) {
            $addinWb = $wb
            break
        }
    }
    if ($null -eq $addinWb) {
        $addinWb = $excel.Workbooks.Open($XlamPath)
    }

    try { $null = $addinWb.VBProject.Name } catch {
        throw "Trust access to the VBA project object model is required for the add-in too."
    }

    foreach ($item in $exported) {
        $libName = $item.Name
        if ($moduleRename.Keys -contains $item.Name) {
            $libName = [string]$moduleRename[$item.Name]
        }
        $importPath = $item.Path
        if ($libName -ne $item.Name) {
            $importPath = Convert-PersonalModuleSource -Path $item.Path -NewName $libName
        }
        foreach ($removeName in @($item.Name, $libName)) {
            $comp = $null
            try { $comp = $addinWb.VBProject.VBComponents.Item($removeName) } catch { $comp = $null }
            if ($null -ne $comp) {
                $addinWb.VBProject.VBComponents.Remove($comp)
                Write-Host "  removed $removeName from add-in"
            }
        }
        [void]$addinWb.VBProject.VBComponents.Import($importPath)
        Write-Host "  imported Personal $($item.Name) as $libName"
    }

    $excel.DisplayAlerts = $false
    $addinWb.Save()
    Write-Host "Saved $($addinWb.FullName)"
    Write-Host "Compile the add-in in the VBE (Debug → Compile VBAProject)."
    Write-Host "If Compile fails on a duplicated Public name, make that helper Private in the Personal module."
} finally {
    if ($null -ne $personalWb) {
        try { $personalWb.Close($false) } catch { }
    }
    if ($createdExcel -and $null -ne $excel) {
        if ($null -ne $addinWb) { try { $addinWb.Close($true) } catch { } }
        try { $excel.Quit() } catch { }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}
