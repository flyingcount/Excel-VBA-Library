Option Explicit

' Document module for ExcelVbaLib.xlam (not imported as a separate component).
' Auto_Open does not run when a workbook is loaded as an add-in; Workbook_Open does.

Private Sub Workbook_Open()
    Call Auto_Open
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call Auto_Close
End Sub
