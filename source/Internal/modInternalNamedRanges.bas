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
