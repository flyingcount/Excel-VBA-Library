Attribute VB_Name = "modApiTables"
Option Explicit

' Public API: Excel Table (ListObject) helpers (Personal Custom_Menu19_Tables).
' Other workbooks / the add-in menu should call these names only.

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

''' @Description: Write every Excel Table in the workbook to a sheet named Table properties, with a hyperlink to each table.
''' @Example: ListTableProperties
Public Sub ListTableProperties()
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    Call modInternalTables.WriteTableInventory(ActiveWorkbook)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("ListTableProperties")
End Sub

''' @Description: List every Excel Table name, sheet, and address in a message box.
''' @Example: ShowAllTablesInWorkbook
Public Sub ShowAllTablesInWorkbook()
    MsgBox modInternalTables.TableListMessage(ActiveWorkbook), vbInformation, "Tables"
End Sub
