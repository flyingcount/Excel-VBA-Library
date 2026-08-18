Attribute VB_Name = "modApiDates"
Option Explicit

' Public API: calendar / date-dimension helpers.

''' @Description: Creates (or refreshes) a date dimension sheet covering StartDate..EndDate inclusive.
''' @Param StartDate: First calendar day.
''' @Param EndDate: Last calendar day.
''' @Param SheetName: Target worksheet name (default "DateTable"). Created if missing; contents cleared.
''' @Param FiscalYearStartMonth: Month number when fiscal year begins (1 = calendar year). Default 1.
''' @Param AsExcelTable: If True, formats the block as a ListObject named "DateTable".
'''
''' @Example:
'''   CreateDateTable #1/1/2024#, #12/31/2026#
'''   CreateDateTable #4/1/2024#, #3/31/2027#, "DimDate", 4
Public Sub CreateDateTable( _
    ByVal StartDate As Date, _
    ByVal EndDate As Date, _
    Optional ByVal SheetName As String = "DateTable", _
    Optional ByVal FiscalYearStartMonth As Long = 1, _
    Optional ByVal AsExcelTable As Boolean = True _
)
    Dim data As Variant
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim lo As ListObject

    On Error GoTo EH

    Call modInternalExcelApp.PushAppState

    data = modInternalDateTable.BuildDateTableArray(StartDate, EndDate, FiscalYearStartMonth)

    Set ws = modApiSheets.EnsureSheet(SheetName)
    ws.Cells.Clear

    ' Drop any prior table on this sheet
    Do While ws.ListObjects.Count > 0
        ws.ListObjects(1).Delete
    Loop

    Call modInternalSheetIO.DumpArray(SheetName, "A1", data)

    lastRow = UBound(data, 1)
    lastCol = UBound(data, 2)

    ws.Columns(1).NumberFormat = "yyyy-mm-dd"
    ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Font.Bold = True
    ws.Columns.AutoFit

    If AsExcelTable And lastRow >= 2 Then
        Set lo = ws.ListObjects.Add( _
            SourceType:=xlSrcRange, _
            Source:=ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)), _
            XlListObjectHasHeaders:=xlYes _
        )
        lo.Name = ValidTableName(SheetName)
    End If

    Call modInternalExcelApp.PopAppState
    Exit Sub

EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("CreateDateTable")
End Sub

Private Function ValidTableName(ByVal SheetName As String) As String
    Dim n As String
    n = Replace(SheetName, " ", "_")
    If Left$(n, 1) Like "[0-9]" Then n = "T_" & n
    ValidTableName = n
End Function
