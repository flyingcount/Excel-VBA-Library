Attribute VB_Name = "modInternalNamedRanges"
Option Explicit

' Internal: workbook-scoped named ranges used by Api/feature modules.

Public Function CheckNamedRangeExists(ByVal RangeName As String) As Boolean
    Dim nm As Name
    CheckNamedRangeExists = False
    For Each nm In ActiveWorkbook.Names
        If nm.Name = RangeName Then
            CheckNamedRangeExists = True
            Exit Function
        End If
    Next nm
End Function

Public Sub CreateNamedRange(ByVal RangeName As String, ByVal RangeToName As Range)
    If CheckNamedRangeExists(RangeName) Then ActiveWorkbook.Names(RangeName).Delete
    ActiveWorkbook.Names.Add Name:=RangeName, RefersTo:=RangeToName, Visible:=True
    ActiveWorkbook.Names(RangeName).Comment = "Created " & Now()
End Sub

' Sheet-scoped name so Residuals / Order do not collide with other sheets.
Public Sub CreateSheetNamedRange(ByVal ws As Worksheet, ByVal RangeName As String, ByVal Target As Range)
    On Error Resume Next
    ws.Names(RangeName).Delete
    On Error GoTo 0
    ws.Names.Add Name:=RangeName, RefersTo:=Target
End Sub
