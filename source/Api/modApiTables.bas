Attribute VB_Name = "modApiTables"
Option Explicit

' Public API: ListObject / table helpers.
' TODO: migrate Personal.xlsb table utilities here.

Public Function RangeToListObject(ByVal TargetSheet As String, ByVal AddressOrName As String, ByVal TableName As String) As ListObject
    Dim ws As Worksheet
    Dim rng As Range
    Set ws = ActiveWorkbook.Worksheets(TargetSheet)
    Set rng = ws.Range(AddressOrName)
    Set RangeToListObject = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    RangeToListObject.Name = TableName
End Function

Public Function ListObjectToArray(ByVal TargetSheet As String, ByVal TableName As String) As Variant
    ListObjectToArray = ActiveWorkbook.Worksheets(TargetSheet).ListObjects(TableName).DataBodyRange.Value
End Function
