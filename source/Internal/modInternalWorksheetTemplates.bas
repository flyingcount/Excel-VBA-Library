Attribute VB_Name = "modInternalWorksheetTemplates"
Option Explicit

' Internal helpers for worksheet templates (Personal Custom_Menu26_wkshtTmplt).
' Called from modApiWorksheetTemplates. Do not document these as the external API.

''' @Description: True if ActiveWorkbook already has a sheet with this name.
Public Function SheetExists(ByVal SheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(SheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

''' @Description: True if any sheet in ActiveWorkbook has a ListObject with this name.
Public Function TableExists(ByVal TableName As String) As Boolean
    Dim ws As Worksheet
    Dim lo As ListObject
    For Each ws In ActiveWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, TableName, vbTextCompare) = 0 Then
                TableExists = True
                Exit Function
            End If
        Next lo
    Next ws
End Function

''' @Description: Create a fresh template sheet. If SheetName exists, offer to rename it to "… OLD".
''' @Return: The new worksheet, or Nothing if the user cancelled or "… OLD" already exists.
Public Function PrepareTemplateSheet(ByVal SheetName As String, ByVal TableName As String) As Worksheet
    Dim wb As Workbook
    Dim oldName As String
    Dim wsOld As Worksheet

    Set wb = ActiveWorkbook
    oldName = SheetName & " OLD"

    If SheetExists(SheetName) Then
        Select Case MsgBox("Worksheet '" & SheetName & "' already exists. Overwrite the old one?", _
                           vbYesNo + vbQuestion, "Excel VBA Lib")
            Case vbNo
                Exit Function
            Case vbYes
                If SheetExists(oldName) Then
                    MsgBox "Old worksheet '" & oldName & "' already exists. Ending procedure.", _
                           vbExclamation, "Excel VBA Lib"
                    Exit Function
                End If
                Set wsOld = wb.Worksheets(SheetName)
                wsOld.Name = oldName
                ' Free the table name so the new sheet can reuse it.
                If TableExists(TableName) Then
                    On Error Resume Next
                    wsOld.ListObjects(TableName).Unlist
                    On Error GoTo 0
                End If
        End Select
    End If

    Set PrepareTemplateSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    PrepareTemplateSheet.Name = SheetName
End Function

''' @Description: Turn the header block on ws into a ListObject named TableName.
Public Function CreateHeaderTable(ByVal ws As Worksheet, ByVal HeaderCell As Range, ByVal TableName As String) As ListObject
    Dim rng As Range
    Set rng = HeaderCell.CurrentRegion
    If TableExists(TableName) Then
        MsgBox "A table named '" & TableName & "' already exists. Ending procedure.", vbExclamation, "Excel VBA Lib"
        Exit Function
    End If
    Set CreateHeaderTable = ws.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=rng, _
        XlListObjectHasHeaders:=xlYes)
    CreateHeaderTable.Name = TableName
End Function

''' @Description: Add a slicer for FieldName. Name is prefixed with TableName so it stays unique in the workbook.
Public Sub AddFieldSlicer(ByVal ws As Worksheet, ByVal lo As ListObject, ByVal FieldName As String, _
                          ByVal TopPt As Double, ByVal LeftPt As Double, _
                          ByVal WidthPt As Double, ByVal HeightPt As Double)
    Dim cache As SlicerCache
    Dim slicerName As String

    slicerName = Replace(Replace(lo.Name & "_" & FieldName, " / ", "_"), " ", "_")
    On Error Resume Next
    Set cache = ActiveWorkbook.SlicerCaches.Add2(lo, FieldName)
    If cache Is Nothing Then Exit Sub
    cache.Slicers.Add ws, , slicerName, FieldName, TopPt, LeftPt, WidthPt, HeightPt
    On Error GoTo 0
End Sub

''' @Description: Index sheet at the front of ActiveWorkbook, with a Back-to-Index link in A1 of each other sheet.
''' Does not insert a row (templates already leave A1 empty).
Public Sub BuildIndexSheet()
    Const indexName As String = "Index"
    Dim wb As Workbook
    Dim wIndex As Worksheet
    Dim wSheet As Worksheet
    Dim r As Long
    Dim c As Long

    Set wb = ActiveWorkbook
    If SheetExists(indexName) Then
        MsgBox "INDEX worksheet already exists. Rename it and re-run.", vbExclamation, "Excel VBA Lib"
        Exit Sub
    End If

    Set wIndex = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    wIndex.Name = indexName
    wIndex.Activate
    ActiveWindow.DisplayGridlines = False

    r = 1
    c = 1
    For Each wSheet In wb.Worksheets
        If r > 20 Then
            c = c + 1
            r = 1
        End If

        If StrComp(wSheet.Name, indexName, vbTextCompare) = 0 Then
            wIndex.Cells(r, c).Value = indexName
        Else
            wIndex.Hyperlinks.Add Anchor:=wIndex.Cells(r, c), Address:="", _
                SubAddress:="'" & wSheet.Name & "'!A1", TextToDisplay:=wSheet.Name
            If Not wSheet.ProtectContents Then
                wSheet.Hyperlinks.Add Anchor:=wSheet.Range("A1"), Address:="", _
                    SubAddress:="'Index'!A1", TextToDisplay:="Back to Index"
            End If
        End If
        r = r + 1
    Next wSheet

    wIndex.Range("A1").Font.Name = "Arial Black"
    wIndex.Columns("A:I").AutoFit
End Sub
