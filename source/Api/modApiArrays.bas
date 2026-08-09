Attribute VB_Name = "modApiArrays"
Option Explicit

' Public API: array dump/load entry points.
' Other workbooks / Personal shims should call these names only.

Public Sub WriteArrayToSheet(ByVal TargetSheet As String, ByVal StartCell As String, ByRef Data As Variant)
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    Call modInternalSheetIO.DumpArray(TargetSheet, StartCell, Data)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("WriteArrayToSheet")
End Sub

Public Function ReadRangeToArray(ByVal TargetSheet As String, ByVal AddressOrName As String) As Variant
    On Error GoTo EH
    ReadRangeToArray = modInternalSheetIO.ReadArray(TargetSheet, AddressOrName)
    Exit Function
EH:
    Call modInternalError.RaiseCurrent("ReadRangeToArray")
End Function

Public Sub WriteArrayToNewSheet(ByVal SheetName As String, ByVal StartCell As String, ByRef Data As Variant)
    On Error GoTo EH
    Call modApiSheets.EnsureSheet(SheetName)
    Call WriteArrayToSheet(SheetName, StartCell, Data)
    Exit Sub
EH:
    Call modInternalError.RaiseCurrent("WriteArrayToNewSheet")
End Sub
