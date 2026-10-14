# Self-contained: patch Poisson 438 in the existing ExcelVbaLib.xlam.
# Does not need Import-AddinModules.ps1 or Python. Paste into PowerShell from any folder.
#
#   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
#   then paste this file, or:
#   & "C:\Users\chanp\OneDrive\Notebooks\Cursor\Excel-VBA-Library\scripts\Patch-Poisson438.ps1"
#
# Requires Excel + Trust access to the VBA project object model.

$ErrorActionPreference = "Stop"
$XlamPath = "C:\Users\chanp\OneDrive\Notebooks\Cursor\Excel-VBA-Library\build\ExcelVbaLib.xlam"

if (-not (Test-Path -LiteralPath $XlamPath)) {
    throw "Add-in not found: $XlamPath"
}

$knuth = @"
Private Function RandomPoisson(ByVal lambda As Double) As Long
    Dim L As Double
    Dim p As Double
    Dim k As Long
    Dim u As Double
    Dim x As Double
    If lambda <= 0# Then
        RandomPoisson = 0
        Exit Function
    End If
    L = Exp(-lambda)
    If L > 0# And lambda < 100# Then
        k = 0
        p = 1#
        Do
            k = k + 1
            p = p * Rnd()
        Loop While p > L
        RandomPoisson = k - 1
        Exit Function
    End If
    u = Rnd()
    If u <= 0 Then u = 0.000000001
    If u >= 1 Then u = 0.999999999
    x = Application.WorksheetFunction.NormInv(u, lambda, Sqr(lambda))
    If x <= 0# Then
        RandomPoisson = 0
    ElseIf x >= 2147483647# Then
        RandomPoisson = 2147483647
    Else
        RandomPoisson = CLng(Int(x + 0.5))
    End If
End Function
"@

$excel = $null
$created = $false
try {
    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
} catch {
    $excel = $null
}
if ($null -eq $excel) {
    $excel = New-Object -ComObject Excel.Application
    $created = $true
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
}

$wb = $null
try {
    foreach ($w in @($excel.Workbooks)) {
        try {
            if ([string]$w.Name -eq "ExcelVbaLib.xlam") { $wb = $w; break }
        } catch { }
    }
    if ($null -eq $wb) {
        Write-Host "Opening $XlamPath"
        $wb = $excel.Workbooks.Open($XlamPath)
    }

    try {
        $null = $wb.VBProject.Name
    } catch {
        throw @"
Excel blocked VBA access.
File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings
-> enable Trust access to the VBA project object model
Then run this again.
"@
    }

    $cm = $wb.VBProject.VBComponents.Item("modInternalData").CodeModule
    $old = $cm.Lines(1, $cm.CountOfLines)
    if ($old -notmatch "Poisson_Inv\(") {
        if ($old -match "Function RandomPoisson\(") {
            Write-Host "Already patched. No Poisson_Inv call in modInternalData."
        } else {
            throw "modInternalData has no Poisson_Inv call and no RandomPoisson. Unexpected module."
        }
    } else {
        $new = $old.Replace(
            "Application.WorksheetFunction.Poisson_Inv(UnitRnd(), lambda)",
            "RandomPoisson(lambda)")
        if ($new -notmatch "Function RandomPoisson\(") {
            $new = $new.TrimEnd() + "`r`n`r`n" + $knuth.Trim() + "`r`n"
        }
        $cm.DeleteLines(1, $cm.CountOfLines)
        $cm.AddFromString($new)
        Write-Host "Replaced Poisson_Inv with RandomPoisson in modInternalData."
    }

    $check = $cm.Lines(1, $cm.CountOfLines)
    if ($check -match "Poisson_Inv\(") {
        throw "Patch failed: Poisson_Inv is still in modInternalData."
    }
    if ($check -notmatch "Function RandomPoisson\(") {
        throw "Patch failed: RandomPoisson was not added."
    }

    $excel.DisplayAlerts = $false
    $wb.Save()
    Write-Host "Saved $($wb.FullName)"
    Write-Host "Retry Excel VBA Lib -> Data -> Probability distributions -> Poisson."
} finally {
    if ($created -and $null -ne $excel) {
        try { if ($null -ne $wb) { $wb.Close($true) } } catch { }
        try { $excel.Quit() } catch { }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}
