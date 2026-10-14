# Export selected VBA modules from Personal.xlsb into source/_export_raw.
#
# Menu13 and other Personal families are gitignored. This workspace often
# has no copy of the workbook; run this on the machine that holds Personal.
#
# From the repo root, with Excel closed or allowing reopen:
#   powershell -ExecutionPolicy Bypass -File scripts/Export-PersonalModules.ps1
#
# Optional:
#   -WorkbookPath "D:\Data\Personal.xlsb"
#   -NameMatch "Custom_Menu13*","Fn_Matrices*"

[CmdletBinding()]
param(
    [string]$RepoRoot,
    [string]$WorkbookPath,
    [string[]]$NameMatch = @("Custom_Menu13*", "Fn_Matrices*", "custom_Menu13*")
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $here = $PSScriptRoot
    if ([string]::IsNullOrWhiteSpace($here)) {
        $here = Split-Path -Parent $MyInvocation.MyCommand.Path
    }
    $RepoRoot = (Resolve-Path (Join-Path $here "..")).Path
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

Write-Host "Exporting from $WorkbookPath"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $null
try {
    $wb = $excel.Workbooks.Open($WorkbookPath)
    try { $null = $wb.VBProject.Name } catch {
        throw "Enable Trust access to the VBA project object model, then retry."
    }
    $exported = 0
    foreach ($comp in @($wb.VBProject.VBComponents)) {
        $name = $comp.Name
        $hit = $false
        foreach ($pat in $NameMatch) {
            if ($name -like $pat) { $hit = $true; break }
        }
        if (-not $hit) { continue }
        $dest = Join-Path $outDir ($name + ".bas")
        if (Test-Path -LiteralPath $dest) { Remove-Item -LiteralPath $dest -Force }
        $comp.Export($dest)
        Write-Host "  $name -> $dest"
        $exported++
    }
    $menus = $null
    try { $menus = $wb.VBProject.VBComponents.Item("Custom_Menu_Menus") } catch { $menus = $null }
    if ($null -ne $menus) {
        $dest = Join-Path $outDir "Custom_Menu_Menus.bas"
        if (Test-Path -LiteralPath $dest) { Remove-Item -LiteralPath $dest -Force }
        $menus.Export($dest)
        Write-Host "  Custom_Menu_Menus -> $dest"
    }
    if ($exported -eq 0) {
        throw "No components matched $($NameMatch -join ', ')."
    }
    Write-Host "Exported $exported module(s) plus menu wiring (if present)."
} finally {
    if ($null -ne $wb) { try { $wb.Close($false) } catch { } }
    try { $excel.Quit() } catch { }
}
