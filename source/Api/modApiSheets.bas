Attribute VB_Name = "modApiSheets"
Option Explicit

' Public API: worksheet lifecycle helpers.

Public Function EnsureSheet(ByVal SheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(SheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
        ws.Name = SheetName
    End If
    Set EnsureSheet = ws
End Function

Public Sub ClearSheetData(ByVal SheetName As String)
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    ActiveWorkbook.Worksheets(SheetName).Cells.Clear
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("ClearSheetData")
End Sub
