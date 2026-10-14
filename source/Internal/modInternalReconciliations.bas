Attribute VB_Name = "modInternalReconciliations"
Option Explicit

' Internal: two-column recon and range compare (Personal Custom_Menu20_*).
' Called from modApiReconciliations. Do not document these as the external API.

Public Const ReconciliationSheetName As String = "Reconciliation"
Public Const ReconciliationStringsSheetName As String = "Reconciliation Strings"
Public Const RangeComparisonSheetName As String = "Range Comparison"

Private Const ColAmount As Long = 1
Private Const ColItem As Long = 2
Private Const ColFlag As Long = 3
Private Const ColStatus As Long = 4
Private Const ColXref As Long = 5
Private Const ColSource As Long = 6
Private Const DatasetCols As Long = 6

Public Function PromptReconRange(ByVal PromptText As String) As Range
    Dim rng As Range
    Set rng = modInternalRanges.PromptRange(PromptText, "Reconciliations")
    If rng Is Nothing Then Exit Function
    Set rng = modInternalRanges.UsedPart(rng)
    If rng Is Nothing Then
        MsgBox "The selected range has no used cells.", vbExclamation, "Reconciliations"
        Exit Function
    End If
    Set PromptReconRange = rng
End Function

Public Function RequireSingleColumn(ByVal rng As Range, ByVal which As String) As Boolean
    If rng.Columns.Count <> 1 Then
        MsgBox which & " must be a single column.", vbExclamation, "Reconciliations"
        Exit Function
    End If
    RequireSingleColumn = True
End Function

Public Sub ReconcileNumericColumns(ByVal rngA As Range, ByVal nameA As String, ByVal rngB As Range, ByVal nameB As String)
    Dim a As Variant
    Dim b As Variant
    Dim recon As Variant
    Dim warn As Variant
    Dim ws As Worksheet
    a = ColumnBlock(rngA, nameA)
    b = ColumnBlock(rngB, nameB)
    Call MatchPairs(a, b, nameA, nameB, True)
    recon = NumericReconStatement(a, b, nameA, nameB)
    warn = MultiMatchWarning(a, b, True)
    Call WriteReconSheet(a, b, recon, warn, True, ReconciliationSheetName)
    Set ws = ActiveWorkbook.Worksheets(ReconciliationSheetName)
    ws.Activate
End Sub

Public Sub ReconcileStringColumns(ByVal rngA As Range, ByVal nameA As String, ByVal rngB As Range, ByVal nameB As String)
    Dim a As Variant
    Dim b As Variant
    Dim recon As Variant
    Dim warn As Variant
    a = ColumnBlock(rngA, nameA)
    b = ColumnBlock(rngB, nameB)
    Call MatchPairs(a, b, nameA, nameB, False)
    recon = StringReconStatement(a, b, nameA, nameB)
    warn = MultiMatchWarning(a, b, False)
    Call WriteReconSheet(a, b, recon, warn, False, ReconciliationStringsSheetName)
    ActiveWorkbook.Worksheets(ReconciliationStringsSheetName).Activate
End Sub

Public Function CompareTwoRanges(ByVal rngA As Range, ByVal rngB As Range) As Long
    Dim nRows As Long
    Dim nCols As Long
    Dim n As Long
    Dim r As Long
    Dim c As Long
    Dim i As Long
    Dim diffs As Long
    Dim arr As Variant
    Dim ws As Worksheet
    Dim lo As ListObject

    nRows = rngA.Rows.Count
    nCols = rngA.Columns.Count
    If nRows <> rngB.Rows.Count Or nCols <> rngB.Columns.Count Then
        MsgBox "The two ranges are different sizes, so they cannot be compared cell by cell.", vbExclamation, "Reconciliations"
        CompareTwoRanges = -1
        Exit Function
    End If
    n = nRows * nCols
    If n > 1048550 Then
        MsgBox "Output will not fit on a worksheet (more than 1,048,550 cells).", vbExclamation, "Reconciliations"
        CompareTwoRanges = -1
        Exit Function
    End If

    rngA.Interior.ColorIndex = xlColorIndexNone
    rngB.Interior.ColorIndex = xlColorIndexNone

    ReDim arr(1 To n + 1, 1 To 10)
    arr(1, 1) = "Index"
    arr(1, 2) = "Range 1 worksheet"
    arr(1, 3) = "Range 1 row"
    arr(1, 4) = "Range 1 column"
    arr(1, 5) = "Range 1 value"
    arr(1, 6) = "Range 2 worksheet"
    arr(1, 7) = "Range 2 row"
    arr(1, 8) = "Range 2 column"
    arr(1, 9) = "Range 2 value"
    arr(1, 10) = "Difference flag"

    i = 1
    For r = 1 To nRows
        For c = 1 To nCols
            i = i + 1
            arr(i, 1) = i - 1
            arr(i, 2) = rngA.Worksheet.Name
            arr(i, 3) = rngA.Cells(r, c).Row
            arr(i, 4) = rngA.Cells(r, c).Column
            arr(i, 5) = rngA.Cells(r, c).Value
            arr(i, 6) = rngB.Worksheet.Name
            arr(i, 7) = rngB.Cells(r, c).Row
            arr(i, 8) = rngB.Cells(r, c).Column
            arr(i, 9) = rngB.Cells(r, c).Value
            If Not ValuesEqual(rngA.Cells(r, c).Value, rngB.Cells(r, c).Value) Then
                arr(i, 10) = True
                rngA.Cells(r, c).Interior.Color = RGB(255, 255, 0)
                rngB.Cells(r, c).Interior.Color = RGB(255, 255, 0)
                diffs = diffs + 1
            End If
        Next c
    Next r

    Call modApiSheets.CreateOutputSheet(RangeComparisonSheetName)
    Set ws = ActiveWorkbook.Worksheets(RangeComparisonSheetName)
    ws.Range("A1").Value = "Range comparison run on " & Now()
    ws.Range("A3").Resize(n + 1, 10).Value = arr
    ws.Range("A3").Resize(1, 10).Font.Bold = True
    ws.Columns("A:J").AutoFit
    On Error Resume Next
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A3").CurrentRegion, , xlYes)
    If Not lo Is Nothing Then lo.Name = "Comparison"
    On Error GoTo 0
    ws.Activate
    CompareTwoRanges = diffs
End Function

Private Function ColumnBlock(ByVal rng As Range, ByVal srcName As String) As Variant
    Dim n As Long
    Dim i As Long
    Dim raw As Variant
    Dim arr As Variant
    n = rng.Rows.Count
    ReDim arr(1 To n, 1 To DatasetCols)
    If n = 1 Then
        arr(1, ColAmount) = rng.Cells(1, 1).Value
    Else
        raw = rng.Value
        For i = 1 To n
            arr(i, ColAmount) = raw(i, 1)
        Next i
    End If
    For i = 1 To n
        arr(i, ColItem) = i
        arr(i, ColSource) = "Source: " & srcName
    Next i
    ColumnBlock = arr
End Function

Private Sub MatchPairs(ByRef a As Variant, ByRef b As Variant, ByVal nameA As String, ByVal nameB As String, ByVal numeric As Boolean)
    Dim queues As Object
    Dim key As String
    Dim i As Long
    Dim j As Long
    Dim nMatch As Long
    Dim flag As String
    Dim q As Collection

    Set queues = CreateObject("Scripting.Dictionary")
    queues.CompareMode = vbBinaryCompare
    For j = 1 To UBound(b, 1)
        key = MatchKey(b(j, ColAmount), numeric)
        If Len(key) > 0 Then
            If Not queues.Exists(key) Then
                Set q = New Collection
                queues.Add key, q
            End If
            queues(key).Add j
        End If
    Next j

    nMatch = 0
    For i = 1 To UBound(a, 1)
        key = MatchKey(a(i, ColAmount), numeric)
        If Len(key) = 0 Then GoTo NextA
        If Not queues.Exists(key) Then GoTo NextA
        Set q = queues(key)
        If q.Count = 0 Then GoTo NextA
        j = q(1)
        q.Remove 1
        nMatch = nMatch + 1
        flag = IndexToLetters(nMatch)
        a(i, ColFlag) = flag
        a(i, ColStatus) = "Matched"
        a(i, ColXref) = nameB & " item: " & CStr(j)
        b(j, ColFlag) = flag
        b(j, ColStatus) = "Matched"
        b(j, ColXref) = nameA & " item: " & CStr(i)
NextA:
    Next i
End Sub

Private Function NumericReconStatement(ByRef a As Variant, ByRef b As Variant, ByVal nameA As String, ByVal nameB As String) As Variant
    Dim totalA As Double
    Dim totalB As Double
    Dim unmatchedA As Double
    Dim unmatchedB As Double
    Dim i As Long
    Dim r As Long
    Dim n As Long
    Dim arr As Variant
    Dim v As Double

    n = UBound(a, 1) + UBound(b, 1) + 8
    ReDim arr(1 To n, 1 To 4)
    totalA = SumNumeric(a)
    totalB = SumNumeric(b)

    r = 1
    arr(r, 1) = "Item number"
    arr(r, 2) = "First dataset total"
    arr(r, 4) = totalA
    r = r + 1

    For i = 1 To UBound(a, 1)
        If Not IsMatched(a, i) Then
            If TryCDbl(a(i, ColAmount), v) Then
                arr(r, 1) = a(i, ColItem)
                arr(r, 2) = nameA & ": record " & CStr(i)
                arr(r, 3) = -v
                unmatchedA = unmatchedA - v
                r = r + 1
            End If
        End If
    Next i
    arr(r, 2) = "First dataset unreconciled items total"
    arr(r, 4) = unmatchedA
    r = r + 1
    arr(r, 2) = "Sub-total"
    arr(r, 4) = totalA + unmatchedA
    r = r + 1

    For i = 1 To UBound(b, 1)
        If Not IsMatched(b, i) Then
            If TryCDbl(b(i, ColAmount), v) Then
                arr(r, 1) = b(i, ColItem)
                arr(r, 2) = nameB & ": record " & CStr(i)
                arr(r, 3) = v
                unmatchedB = unmatchedB + v
                r = r + 1
            End If
        End If
    Next i
    arr(r, 2) = "Second dataset unreconciled items total"
    arr(r, 4) = unmatchedB
    r = r + 1
    arr(r, 2) = "Grand total"
    arr(r, 3) = unmatchedA + unmatchedB
    arr(r, 4) = totalA + unmatchedA + unmatchedB
    r = r + 1
    arr(r, 2) = "Second dataset total"
    arr(r, 4) = totalB
    r = r + 1
    arr(r, 2) = "Unreconciled balance"
    arr(r, 4) = (totalA + unmatchedA + unmatchedB) - totalB

    NumericReconStatement = TrimRows(arr, r, 4)
End Function

Private Function StringReconStatement(ByRef a As Variant, ByRef b As Variant, ByVal nameA As String, ByVal nameB As String) As Variant
    Dim i As Long
    Dim r As Long
    Dim n As Long
    Dim arr As Variant
    n = UBound(a, 1) + UBound(b, 1) + 2
    ReDim arr(1 To n, 1 To 4)
    r = 1
    arr(r, 1) = "Item number"
    arr(r, 2) = "First dataset"
    r = r + 1
    For i = 1 To UBound(a, 1)
        If Not IsMatched(a, i) Then
            arr(r, 1) = a(i, ColItem)
            arr(r, 2) = nameA & ": record " & CStr(i)
            arr(r, 3) = a(i, ColAmount)
            arr(r, 4) = "Record not in " & nameB
            r = r + 1
        End If
    Next i
    For i = 1 To UBound(b, 1)
        If Not IsMatched(b, i) Then
            arr(r, 1) = b(i, ColItem)
            arr(r, 2) = nameB & ": record " & CStr(i)
            arr(r, 3) = b(i, ColAmount)
            arr(r, 4) = "Record not in " & nameA
            r = r + 1
        End If
    Next i
    StringReconStatement = TrimRows(arr, r - 1, 4)
End Function

Private Function MultiMatchWarning(ByRef a As Variant, ByRef b As Variant, ByVal numeric As Boolean) As Variant
    Dim arr As Variant
    Dim r As Long
    Dim n As Long
    n = UBound(a, 1) + UBound(b, 1) + 4
    ReDim arr(1 To n, 1 To DatasetCols)
    arr(1, 1) = "Warning: listed below, if any, are amounts that appear more than once. Items may have been mismatched."
    arr(3, 1) = "Amount"
    arr(3, 2) = "Item number"
    arr(3, 3) = "Flag"
    arr(3, 4) = "Status"
    arr(3, 5) = "Recon cross reference"
    arr(3, 6) = "Source dataset"
    r = 3
    Call AppendDupUnmatched(arr, r, a, numeric)
    Call AppendDupUnmatched(arr, r, b, numeric)
    If r = 3 Then
        MultiMatchWarning = Empty
    Else
        MultiMatchWarning = TrimRows(arr, r, DatasetCols)
    End If
End Function

Private Sub AppendDupUnmatched(ByRef dest As Variant, ByRef r As Long, ByRef src As Variant, ByVal numeric As Boolean)
    Dim counts As Object
    Dim listed As Object
    Dim i As Long
    Dim key As String
    Set counts = CreateObject("Scripting.Dictionary")
    Set listed = CreateObject("Scripting.Dictionary")
    counts.CompareMode = vbBinaryCompare
    listed.CompareMode = vbBinaryCompare
    For i = 1 To UBound(src, 1)
        key = MatchKey(src(i, ColAmount), numeric)
        If Len(key) = 0 Then GoTo NextCount
        If counts.Exists(key) Then
            counts(key) = counts(key) + 1
        Else
            counts.Add key, 1
        End If
NextCount:
    Next i
    For i = 1 To UBound(src, 1)
        If IsMatched(src, i) Then GoTo NextRow
        key = MatchKey(src(i, ColAmount), numeric)
        If Len(key) = 0 Then GoTo NextRow
        If counts(key) <= 1 Then GoTo NextRow
        If listed.Exists(CStr(i)) Then GoTo NextRow
        r = r + 1
        dest(r, 1) = src(i, ColAmount)
        dest(r, 2) = src(i, ColItem)
        dest(r, 3) = src(i, ColFlag)
        dest(r, 4) = src(i, ColStatus)
        dest(r, 5) = src(i, ColXref)
        dest(r, 6) = src(i, ColSource)
        listed.Add CStr(i), True
NextRow:
    Next i
End Sub

Private Sub WriteReconSheet(ByRef a As Variant, ByRef b As Variant, ByRef recon As Variant, ByRef warn As Variant, ByVal numeric As Boolean, ByVal sheetName As String)
    Dim ws As Worksheet
    Dim nA As Long
    Dim nB As Long
    Dim nRecon As Long
    nA = UBound(a, 1)
    nB = UBound(b, 1)
    Call modApiSheets.CreateOutputSheet(sheetName)
    Set ws = ActiveWorkbook.Worksheets(sheetName)
    ws.Range("A1").Value = "Reconciliation run on " & Now()
    If numeric Then
        ws.Range("E2").Value = "Notes"
        ws.Range("E2").Font.Bold = True
    End If
    Call DumpBlock(ws, "A3", recon)
    Call DumpBlock(ws, "H3", a)
    Call DumpBlock(ws, "O3", b)
    ws.Range("H2:M2").Value = Array("Amount", "Item number", "Flag", "Status", "Recon cross reference", "Source dataset")
    ws.Range("O2:T2").Value = Array("Amount", "Item number", "Flag", "Status", "Recon cross reference", "Source dataset")
    ws.Range("H2:M2").Font.Bold = True
    ws.Range("O2:T2").Font.Bold = True
    ws.Range("H2:M2").Interior.Color = RGB(255, 204, 153)
    ws.Range("O2:T2").Interior.Color = RGB(0, 255, 0)
    If numeric Then
        ws.Columns("C:D").ColumnWidth = 12
        ws.Columns("C:D").NumberFormat = "#,##0.00"
        ws.Columns("H").NumberFormat = "#,##0.00"
        ws.Columns("O").NumberFormat = "#,##0.00"
    Else
        ws.Columns("C:D").ColumnWidth = 30
    End If
    ws.Columns("A:B").AutoFit
    ws.Columns("E").ColumnWidth = 24
    ws.Columns("G").ColumnWidth = 1
    ws.Columns("N").ColumnWidth = 1
    ws.Columns("H:M").AutoFit
    ws.Columns("O:T").AutoFit
    If Not IsEmpty(warn) Then
        nRecon = UBound(recon, 1)
        Call DumpBlock(ws, ws.Cells(nRecon + 5, 1).Address, warn)
        ws.Columns("A:F").AutoFit
    End If
End Sub

Private Sub DumpBlock(ByVal ws As Worksheet, ByVal startCell As String, ByRef data As Variant)
    Dim nRows As Long
    Dim nCols As Long
    If IsEmpty(data) Then Exit Sub
    nRows = UBound(data, 1) - LBound(data, 1) + 1
    nCols = UBound(data, 2) - LBound(data, 2) + 1
    ws.Range(startCell).Resize(nRows, nCols).Value = data
End Sub

Private Function MatchKey(ByVal v As Variant, ByVal numeric As Boolean) As String
    Dim n As Double
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    If numeric Then
        If TryCDbl(v, n) Then MatchKey = "N:" & CStr(n)
    Else
        MatchKey = "S:" & CStr(v)
        If Len(CStr(v)) = 0 Then MatchKey = ""
    End If
End Function

Private Function IsMatched(ByRef arr As Variant, ByVal i As Long) As Boolean
    IsMatched = (CStr(arr(i, ColStatus)) = "Matched")
End Function

Private Function SumNumeric(ByRef arr As Variant) As Double
    Dim i As Long
    Dim v As Double
    Dim s As Double
    For i = 1 To UBound(arr, 1)
        If TryCDbl(arr(i, ColAmount), v) Then s = s + v
    Next i
    SumNumeric = s
End Function

Private Function TryCDbl(ByVal v As Variant, ByRef outN As Double) As Boolean
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    If VarType(v) = vbBoolean Then Exit Function
    If VarType(v) = vbString Then
        If Len(Trim$(CStr(v))) = 0 Then Exit Function
    End If
    If Not IsNumeric(v) Then Exit Function
    outN = CDbl(v)
    TryCDbl = True
End Function

Private Function ValuesEqual(ByVal a As Variant, ByVal b As Variant) As Boolean
    Dim diff As Boolean
    On Error Resume Next
    diff = (a <> b)
    If Err.Number <> 0 Then
        ValuesEqual = False
        Exit Function
    End If
    On Error GoTo 0
    ValuesEqual = Not diff
End Function

Private Function TrimRows(ByRef src As Variant, ByVal lastRow As Long, ByVal nCols As Long) As Variant
    Dim arr As Variant
    Dim r As Long
    Dim c As Long
    If lastRow < 1 Then lastRow = 1
    ReDim arr(1 To lastRow, 1 To nCols)
    For r = 1 To lastRow
        For c = 1 To nCols
            arr(r, c) = src(r, c)
        Next c
    Next r
    TrimRows = arr
End Function

Private Function IndexToLetters(ByVal n As Long) As String
    Dim s As String
    If n < 1 Then n = 1
    Do
        n = n - 1
        s = Chr$(65 + (n Mod 26)) & s
        n = n \ 26
    Loop While n > 0
    IndexToLetters = s
End Function
