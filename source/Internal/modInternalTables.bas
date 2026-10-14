Attribute VB_Name = "modInternalTables"
Option Explicit

' Internal: Excel Table (ListObject) inventory
' (Personal Custom_Menu19_Tables ListTableProperties / ShowAllTablesInWorkbook).
' Called from modApiTables. Do not document these as the external API.

Public Const TablePropertiesSheetName As String = "Table properties"
Private Const PropColCount As Long = 16
Private Const LinkCol As Long = 6
Private Const SheetCol As Long = 4
Private Const RangeCol As Long = 5
Private Const NameCol As Long = 2

Public Function CountExcelTables(ByVal wb As Workbook) As Long
    Dim ws As Worksheet
    Dim n As Long
    For Each ws In wb.Worksheets
        n = n + ws.ListObjects.Count
    Next ws
    CountExcelTables = n
End Function

Public Function TableInventoryArray(ByVal wb As Workbook) As Variant
    Dim n As Long
    Dim arr As Variant
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim r As Long
    Dim headerAddr As String

    n = CountExcelTables(wb)
    If n = 0 Then
        TableInventoryArray = Empty
        Exit Function
    End If

    ReDim arr(1 To n + 1, 1 To PropColCount)
    arr(1, 1) = "Index"
    arr(1, 2) = "Name"
    arr(1, 3) = "Display name"
    arr(1, 4) = "Worksheet"
    arr(1, 5) = "Range"
    arr(1, 6) = "Link to table"
    arr(1, 7) = "No. rows"
    arr(1, 8) = "No. columns"
    arr(1, 9) = "Show autofilter"
    arr(1, 10) = "Show headers"
    arr(1, 11) = "Show totals"
    arr(1, 12) = "Table style"
    arr(1, 13) = "Source type"
    arr(1, 14) = "Comment"
    arr(1, 15) = "Alternative text"
    arr(1, 16) = "Header range"

    r = 2
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If lo.ShowHeaders Then
                headerAddr = lo.HeaderRowRange.Address
            Else
                headerAddr = "N/A"
            End If
            arr(r, 1) = r - 1
            arr(r, 2) = lo.Name
            arr(r, 3) = lo.DisplayName
            arr(r, 4) = ws.Name
            arr(r, 5) = lo.Range.Address
            arr(r, 6) = "Link to " & lo.Name
            arr(r, 7) = lo.ListRows.Count
            arr(r, 8) = lo.ListColumns.Count
            arr(r, 9) = lo.ShowAutoFilter
            arr(r, 10) = lo.ShowHeaders
            arr(r, 11) = lo.ShowTotals
            arr(r, 12) = SafeTableStyleName(lo)
            arr(r, 13) = SourceTypeName(lo.SourceType)
            arr(r, 14) = SafeListObjectText(lo, "Comment")
            arr(r, 15) = SafeListObjectText(lo, "AlternativeText")
            arr(r, 16) = headerAddr
            r = r + 1
        Next lo
    Next ws
    TableInventoryArray = arr
End Function

Public Sub WriteTableInventory(ByVal wb As Workbook)
    Dim n As Long
    Dim arr As Variant
    Dim ws As Worksheet
    Dim k As Long
    Dim subAddr As String

    n = CountExcelTables(wb)
    If n = 0 Then
        MsgBox "There are no Excel Tables (ListObjects) in " & wb.Name & ".", vbInformation, "Tables"
        Exit Sub
    End If

    arr = TableInventoryArray(wb)
    Call modApiSheets.CreateOutputSheet(TablePropertiesSheetName)
    Set ws = wb.Worksheets(TablePropertiesSheetName)
    ws.Range("A1").Value = "Excel Tables in " & wb.Name & " (" & CStr(n) & ")"
    ws.Range("A1").Font.Bold = True
    ws.Range("A3").Resize(n + 1, PropColCount).Value = arr
    ws.Range("A3").Resize(1, PropColCount).Font.Bold = True

    For k = 1 To n
        subAddr = "'" & Replace(CStr(arr(k + 1, SheetCol)), "'", "''") & "'!" & CStr(arr(k + 1, RangeCol))
        ws.Hyperlinks.Add Anchor:=ws.Cells(k + 3, LinkCol), Address:="", _
            SubAddress:=subAddr, TextToDisplay:="Link to " & CStr(arr(k + 1, NameCol))
    Next k

    ws.Columns("A:P").AutoFit
    ws.Columns("A").ColumnWidth = 7
    If ws.Columns("N").ColumnWidth > 40 Then ws.Columns("N").ColumnWidth = 40
    If ws.Columns("O").ColumnWidth > 40 Then ws.Columns("O").ColumnWidth = 40
    ws.Activate
    On Error Resume Next
    ActiveWindow.FreezePanes = False
    ActiveWindow.SplitColumn = 2
    ActiveWindow.SplitRow = 3
    ActiveWindow.FreezePanes = True
    On Error GoTo 0
End Sub

Public Function TableListMessage(ByVal wb As Workbook) As String
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim n As Long
    Dim i As Long
    Dim lines As String
    Dim vis As String
    Dim msg As String
    Const MaxLen As Long = 900

    n = CountExcelTables(wb)
    If n = 0 Then
        TableListMessage = "There are no Excel Tables (ListObjects) in " & wb.Name & "."
        Exit Function
    End If

    For Each ws In wb.Worksheets
        Select Case ws.Visible
            Case xlSheetHidden
                vis = " (hidden)"
            Case xlSheetVeryHidden
                vis = " (very hidden)"
            Case Else
                vis = ""
        End Select
        For Each lo In ws.ListObjects
            i = i + 1
            lines = lines & CStr(i) & "  " & lo.Name & " - " & ws.Name & vis & _
                    "  " & lo.Range.Address & vbCrLf
        Next lo
    Next ws

    msg = "Tables in " & wb.Name & " (" & CStr(n) & "):" & vbCrLf & vbCrLf & lines
    If Len(msg) > MaxLen Then
        msg = Left$(msg, MaxLen) & vbCrLf & vbCrLf & _
              "List truncated. Use Tables → List table properties for the full inventory."
    End If
    TableListMessage = msg
End Function

Private Function SourceTypeName(ByVal src As Long) As String
    Select Case src
        Case xlSrcRange
            SourceTypeName = "Range"
        Case xlSrcExternal
            SourceTypeName = "External"
        Case xlSrcQuery
            SourceTypeName = "Query"
        Case xlSrcXml
            SourceTypeName = "XML"
        Case xlSrcModel
            SourceTypeName = "Data Model"
        Case Else
            SourceTypeName = CStr(src)
    End Select
End Function

Private Function SafeTableStyleName(ByVal lo As ListObject) As String
    On Error Resume Next
    SafeTableStyleName = CStr(lo.TableStyle)
    On Error GoTo 0
End Function

Private Function SafeListObjectText(ByVal lo As ListObject, ByVal propName As String) As String
    On Error Resume Next
    Select Case propName
        Case "Comment"
            SafeListObjectText = CStr(lo.Comment)
        Case "AlternativeText"
            SafeListObjectText = CStr(lo.AlternativeText)
    End Select
    On Error GoTo 0
End Function
