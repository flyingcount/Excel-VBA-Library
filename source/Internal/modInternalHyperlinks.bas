Attribute VB_Name = "modInternalHyperlinks"
Option Explicit

' Internal: workbook hyperlink inventory, index, and follow/remove helpers
' (Personal Custom_Menu21_Hyperlinks / Custom_Menu21_Index).
' Called from modApiHyperlinks. Do not document these as the external API.

Public Const HyperlinksSheetName As String = "Hyperlinks"
Public Const IndexSheetName As String = "Index"
Private Const IndexBackRange As String = "A1"
Private Const IndexRowsPerCol As Long = 20

Public Function CountWorkbookHyperlinks(ByVal wb As Workbook) As Long
    Dim ws As Worksheet
    Dim n As Long
    For Each ws In wb.Worksheets
        n = n + ws.Hyperlinks.Count
    Next ws
    CountWorkbookHyperlinks = n
End Function

' Header plus one row per worksheet hyperlink (range or shape).
Public Function HyperlinkInventoryArray(ByVal wb As Workbook) As Variant
    Dim n As Long
    Dim arr As Variant
    Dim ws As Worksheet
    Dim hl As Hyperlink
    Dim r As Long

    n = CountWorkbookHyperlinks(wb)
    ReDim arr(1 To n + 1, 1 To 7)
    arr(1, 1) = "Worksheet"
    arr(1, 2) = "Location"
    arr(1, 3) = "Text to display"
    arr(1, 4) = "ScreenTip"
    arr(1, 5) = "Address"
    arr(1, 6) = "SubAddress"
    arr(1, 7) = "Type"

    r = 2
    For Each ws In wb.Worksheets
        For Each hl In ws.Hyperlinks
            arr(r, 1) = ws.Name
            arr(r, 2) = HyperlinkLocation(hl)
            arr(r, 3) = SafeHyperlinkText(hl)
            arr(r, 4) = SafeScreenTip(hl)
            arr(r, 5) = hl.Address
            arr(r, 6) = hl.SubAddress
            arr(r, 7) = HyperlinkTypeName(hl)
            r = r + 1
        Next hl
    Next ws
    HyperlinkInventoryArray = arr
End Function

Public Sub WriteHyperlinkInventory(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim arr As Variant
    Dim n As Long

    n = CountWorkbookHyperlinks(wb)
    If n = 0 Then
        MsgBox "There are no hyperlinks in " & wb.Name & ".", vbInformation, "Hyperlinks"
        Exit Sub
    End If

    arr = HyperlinkInventoryArray(wb)
    Call modApiSheets.CreateOutputSheet(HyperlinksSheetName)
    Set ws = wb.Worksheets(HyperlinksSheetName)
    Call FormatInventorySheet(ws, arr, n)
End Sub

Public Function RemoveHyperlinksByDisplayText(ByVal wb As Workbook, ByVal texts As Range) As Long
    Dim ws As Worksheet
    Dim i As Long
    Dim removed As Long
    Dim display As String
    Dim needle As String
    Dim cell As Range
    Dim lookup As Object

    Set lookup = CreateObject("Scripting.Dictionary")
    lookup.CompareMode = 1
    Set texts = VisibleTextCells(texts)
    If texts Is Nothing Then
        RemoveHyperlinksByDisplayText = 0
        Exit Function
    End If
    For Each cell In texts.Cells
        If Not IsError(cell.Value) Then
            needle = Trim$(CStr(cell.Value))
            If Len(needle) > 0 Then
                If Not lookup.Exists(needle) Then lookup.Add needle, True
            End If
        End If
    Next cell
    If lookup.Count = 0 Then
        RemoveHyperlinksByDisplayText = 0
        Exit Function
    End If

    For Each ws In wb.Worksheets
        For i = ws.Hyperlinks.Count To 1 Step -1
            display = SafeHyperlinkText(ws.Hyperlinks(i))
            If lookup.Exists(display) Then
                ws.Hyperlinks(i).Delete
                removed = removed + 1
            End If
        Next i
    Next ws
    RemoveHyperlinksByDisplayText = removed
End Function

Public Function FollowHyperlinksInRange(ByVal rng As Range) As Long
    Dim cell As Range
    Dim n As Long
    Dim opened As Long
    Dim area As Range
    Dim target As Range

    If rng Is Nothing Then Exit Function
    For Each area In rng.Areas
        Set target = Application.Intersect(area, area.Worksheet.UsedRange)
        If Not target Is Nothing Then
            For Each cell In target.Cells
                If cell.Hyperlinks.Count > 0 Then n = n + 1
            Next cell
        End If
    Next area

    If n = 0 Then
        MsgBox "No hyperlinks in the selected range.", vbInformation, "Hyperlinks"
        Exit Function
    End If
    If n > 1 Then
        If MsgBox("Open " & CStr(n) & " hyperlinks in the default app/browser?", _
                  vbYesNo + vbQuestion, "Hyperlinks") = vbNo Then
            Exit Function
        End If
    End If

    For Each area In rng.Areas
        Set target = Application.Intersect(area, area.Worksheet.UsedRange)
        If Not target Is Nothing Then
            For Each cell In target.Cells
                If cell.Hyperlinks.Count > 0 Then
                    On Error Resume Next
                    cell.Hyperlinks(1).Follow
                    If Err.Number = 0 Then opened = opened + 1
                    Err.Clear
                    On Error GoTo 0
                End If
            Next cell
        End If
    Next area
    FollowHyperlinksInRange = opened
End Function

' Inserts a Back-to-Index row at the top of each unprotected sheet (Personal CreateIndex).
Public Sub CreateWorkbookIndex(ByVal wb As Workbook, ByVal insertBackRow As Boolean)
    Dim wIndex As Worksheet
    Dim wSheet As Worksheet
    Dim r As Long
    Dim c As Long

    If modInternalWorksheetTemplates.SheetExists(IndexSheetName) Then
        MsgBox "A worksheet named '" & IndexSheetName & "' already exists. Rename it or use Update index.", _
               vbExclamation, "Hyperlinks"
        Exit Sub
    End If

    Set wIndex = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    wIndex.Name = IndexSheetName

    r = 1
    c = 1
    For Each wSheet In wb.Worksheets
        If r > IndexRowsPerCol Then
            c = c + 1
            r = 1
        End If
        If StrComp(wSheet.Name, IndexSheetName, vbTextCompare) = 0 Then
            wIndex.Cells(r, c).Value = IndexSheetName
        Else
            wIndex.Hyperlinks.Add Anchor:=wIndex.Cells(r, c), Address:="", _
                SubAddress:="'" & wSheet.Name & "'!" & IndexBackRange, _
                TextToDisplay:=wSheet.Name
            If Not wSheet.ProtectContents Then
                If insertBackRow Then wSheet.Rows(1).Insert
                Call EnsureBackLink(wSheet)
            End If
        End If
        r = r + 1
    Next wSheet

    Call FormatIndexSheet(wIndex)
End Sub

' Rebuilds Index in place. Does not insert rows; refreshes A1 back-links on unprotected sheets.
Public Sub UpdateWorkbookIndex(ByVal wb As Workbook)
    Dim wIndex As Worksheet
    Dim wSheet As Worksheet
    Dim r As Long
    Dim c As Long

    If Not modInternalWorksheetTemplates.SheetExists(IndexSheetName) Then
        MsgBox "A worksheet named '" & IndexSheetName & "' does not exist. Create an index first.", _
               vbExclamation, "Hyperlinks"
        Exit Sub
    End If

    Set wIndex = wb.Worksheets(IndexSheetName)
    wIndex.Cells.Clear
    wIndex.Hyperlinks.Delete

    r = 1
    c = 1
    For Each wSheet In wb.Worksheets
        If r > IndexRowsPerCol Then
            c = c + 1
            r = 1
        End If
        If StrComp(wSheet.Name, IndexSheetName, vbTextCompare) = 0 Then
            wIndex.Cells(r, c).Value = IndexSheetName
        Else
            wIndex.Hyperlinks.Add Anchor:=wIndex.Cells(r, c), Address:="", _
                SubAddress:="'" & wSheet.Name & "'!" & IndexBackRange, _
                TextToDisplay:=wSheet.Name
            If Not wSheet.ProtectContents Then Call EnsureBackLink(wSheet)
        End If
        r = r + 1
    Next wSheet

    Call FormatIndexSheet(wIndex)
End Sub

Public Function AddBackLinksToActiveSheetA1(ByVal wb As Workbook, ByVal wsActive As Worksheet, ByVal destAddr As String) As Long
    Dim wsTarget As Worksheet
    Dim targetCell As Range
    Dim subAddr As String
    Dim linkText As String
    Dim added As Long

    subAddr = "'" & wsActive.Name & "'!A1"
    linkText = "Back to '" & wsActive.Name & "'"

    For Each wsTarget In wb.Worksheets
        If StrComp(wsTarget.Name, wsActive.Name, vbTextCompare) <> 0 Then
            If Not wsTarget.ProtectContents Then
                On Error Resume Next
                Set targetCell = Nothing
                Set targetCell = wsTarget.Range(destAddr)
                On Error GoTo 0
                If Not targetCell Is Nothing Then
                    Call ClearCellHyperlinks(targetCell)
                    targetCell.ClearContents
                    wsTarget.Hyperlinks.Add Anchor:=targetCell, Address:="", _
                        SubAddress:=subAddr, TextToDisplay:=linkText
                    added = added + 1
                End If
            End If
        End If
    Next wsTarget
    AddBackLinksToActiveSheetA1 = added
End Function

Public Function WorksheetListMessage(ByVal wb As Workbook) As String
    Dim ws As Worksheet
    Dim lines As String
    Dim vis As String
    For Each ws In wb.Worksheets
        Select Case ws.Visible
            Case xlSheetHidden
                vis = " (hidden)"
            Case xlSheetVeryHidden
                vis = " (very hidden)"
            Case Else
                vis = ""
        End Select
        lines = lines & CStr(ws.Index) & "  " & ws.Name & vis & vbCrLf
    Next ws
    WorksheetListMessage = "Worksheets in " & wb.Name & " (" & CStr(wb.Worksheets.Count) & "):" & vbCrLf & vbCrLf & lines
End Function

Private Sub EnsureBackLink(ByVal wSheet As Worksheet)
    Dim cell As Range
    Set cell = wSheet.Range(IndexBackRange)
    Call ClearCellHyperlinks(cell)
    wSheet.Hyperlinks.Add Anchor:=cell, Address:="", _
        SubAddress:="'" & IndexSheetName & "'!A1", _
        TextToDisplay:="Back to Index"
End Sub

Private Sub ClearCellHyperlinks(ByVal cell As Range)
    Dim i As Long
    For i = cell.Hyperlinks.Count To 1 Step -1
        cell.Hyperlinks(i).Delete
    Next i
End Sub

Private Sub FormatInventorySheet(ByVal ws As Worksheet, ByRef arr As Variant, ByVal n As Long)
    Dim lastRow As Long
    lastRow = n + 1
    ws.Range("A5").Resize(lastRow, 7).Value = arr
    ws.Range("A3").Value = "Inventory of hyperlinks in the workbook"
    ws.Range("A3").Font.Bold = True
    ws.Range("A5:G5").Font.Bold = True
    ws.Columns("A:G").AutoFit
    If ws.Columns("C").ColumnWidth > 40 Then ws.Columns("C").ColumnWidth = 40
    If ws.Columns("E").ColumnWidth > 40 Then ws.Columns("E").ColumnWidth = 40
    ws.Activate
End Sub

Private Sub FormatIndexSheet(ByVal wIndex As Worksheet)
    On Error Resume Next
    wIndex.Activate
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0
    wIndex.Range("A1").Font.Name = "Arial Black"
    wIndex.Columns("A:I").AutoFit
End Sub

Private Function HyperlinkLocation(ByVal hl As Hyperlink) As String
    On Error Resume Next
    HyperlinkLocation = hl.Range.Address(False, False)
    If Len(HyperlinkLocation) = 0 Then HyperlinkLocation = CStr(hl.Parent.Name)
    On Error GoTo 0
End Function

Private Function SafeHyperlinkText(ByVal hl As Hyperlink) As String
    On Error Resume Next
    SafeHyperlinkText = CStr(hl.TextToDisplay)
    On Error GoTo 0
End Function

Private Function SafeScreenTip(ByVal hl As Hyperlink) As String
    On Error Resume Next
    SafeScreenTip = CStr(hl.ScreenTip)
    On Error GoTo 0
End Function

Private Function HyperlinkTypeName(ByVal hl As Hyperlink) As String
    On Error Resume Next
    Select Case hl.Type
        Case msoHyperlinkRange
            HyperlinkTypeName = "Range"
        Case msoHyperlinkShape
            HyperlinkTypeName = "Shape"
        Case msoHyperlinkInlineShape
            HyperlinkTypeName = "Inline shape"
        Case Else
            HyperlinkTypeName = CStr(hl.Type)
    End Select
    On Error GoTo 0
End Function

Private Function VisibleTextCells(ByVal rng As Range) As Range
    Dim area As Range
    Dim acc As Range
    Dim target As Range
    If rng Is Nothing Then Exit Function
    For Each area In rng.Areas
        Set target = Application.Intersect(area, area.Worksheet.UsedRange)
        If Not target Is Nothing Then
            If acc Is Nothing Then
                Set acc = target
            Else
                Set acc = Union(acc, target)
            End If
        End If
    Next area
    Set VisibleTextCells = acc
End Function
