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

' Personal Menu13 names (Custom_Menu13_Matrices1 diagnostic / Cholesky / eigen output sheets).

Public Function WorksheetExists(ByVal str_Name As String) As Boolean
    Dim ws As Object
    On Error Resume Next
    Set ws = ActiveWorkbook.Sheets(str_Name)
    On Error GoTo 0
    WorksheetExists = Not ws Is Nothing
End Function

Public Sub CheckExistenceAndDeleteOutputSheet(ByVal str_Name As String)
    If WorksheetExists(str_Name) Then
        Application.DisplayAlerts = False
        ActiveWorkbook.Sheets(str_Name).Delete
        Application.DisplayAlerts = True
    End If
End Sub

Public Sub CreateOutputSheet(ByVal str_Name As String)
    Call CheckExistenceAndDeleteOutputSheet(str_Name)
    ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.ActiveSheet).Name = str_Name
End Sub
