Option Explicit

' Document module for ExcelVbaLib.xlam (injected, not imported as a second component).
' Excel does not run Auto_Open when a workbook is loaded as an add-in.
' Calling a procedure named Auto_Open from here is also skipped. Schedule a
' differently named installer after Excel is idle so CommandBars exist.

Private Sub Workbook_Open()
    Call modAddinMenu.ScheduleMenuInstall
End Sub

Private Sub Workbook_AddinInstall()
    Call modAddinMenu.ScheduleMenuInstall
End Sub

Private Sub Workbook_AddinUninstall()
    Call modAddinMenu.RemoveExcelVbaLibMenu
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call modAddinMenu.RemoveExcelVbaLibMenu
End Sub
