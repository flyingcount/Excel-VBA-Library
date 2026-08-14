Attribute VB_Name = "modInternalSampling"
Option Explicit

' Internal: row sampling (Personal Custom_Menu4_Sample / ExtractSample).
' Called from modApiSampling. Do not document these as the external API.

''' @Description: Draw SampleSize rows from Data (1-based 2D) onto the existing Sample sheet.
''' Caller must create/clear that sheet first.
Public Sub RunExtractSample(ByVal Data As Variant, ByVal SampleSize As Long, ByVal WithReplacement As Boolean)
    Dim n As Long
    Dim cols As Long
    Dim picks() As Long
    Dim sample() As Variant
    Dim freq() As Long
    Dim audit() As Variant
    Dim i As Long
    Dim r As Long
    Dim c As Long
    Dim ws As Worksheet
    Dim outRow As Long
    Dim pct As Double

    n = UBound(Data, 1) - LBound(Data, 1) + 1
    cols = UBound(Data, 2) - LBound(Data, 2) + 1
    If SampleSize < 1 Then SampleSize = 1
    If Not WithReplacement And SampleSize > n Then SampleSize = n

    picks = DrawRowIndexes(n, SampleSize, WithReplacement)
    ReDim sample(1 To SampleSize, 1 To cols)
    ReDim freq(1 To n)

    For i = 1 To SampleSize
        r = picks(i)
        freq(r) = freq(r) + 1
        For c = 1 To cols
            sample(i, c) = Data(LBound(Data, 1) + r - 1, LBound(Data, 2) + c - 1)
        Next c
    Next i

    Set ws = ActiveWorkbook.Worksheets("Sample")

    pct = SampleSize / n * 100
    ws.Range("A1").Value = "Sample of " & SampleSize & " from " & n & " (" & Format$(pct, "0.0") & "%) " & _
        IIf(WithReplacement, "with replacement", "without replacement")
    ws.Range("A1").Font.Bold = True

    Call modInternalSheetIO.DumpArray("Sample", "A3", sample)

    outRow = 3 + SampleSize + 2
    ws.Cells(outRow, 1).Value = "Source row (draw order)"
    ws.Cells(outRow, 1).Font.Bold = True
    For i = 1 To SampleSize
        ws.Cells(outRow, 1 + i).Value = picks(i)
    Next i

    outRow = outRow + 3
    ws.Cells(outRow, 1).Value = "Row number"
    ws.Cells(outRow, 2).Value = "Times sampled"
    ws.Range(ws.Cells(outRow, 1), ws.Cells(outRow, 2)).Font.Bold = True

    ReDim audit(1 To n, 1 To 2)
    For i = 1 To n
        audit(i, 1) = i
        audit(i, 2) = freq(i)
    Next i
    Call modInternalSheetIO.DumpArray("Sample", ws.Cells(outRow + 1, 1).Address, audit)

    ws.Columns.AutoFit
End Sub

''' @Description: Turn a Range.Value into a 2-D array (1 row x 1 col if the range is a single cell).
Public Function RangeTo2D(ByVal rng As Range) As Variant
    Dim v As Variant
    Dim a(1 To 1, 1 To 1) As Variant

    v = rng.Value
    If IsArray(v) Then
        RangeTo2D = v
    Else
        a(1, 1) = v
        RangeTo2D = a
    End If
End Function

' 1-based indexes into the source rows.
Private Function DrawRowIndexes(ByVal n As Long, ByVal k As Long, ByVal WithReplacement As Boolean) As Long()
    Dim picks() As Long
    Dim idx() As Long
    Dim i As Long
    Dim j As Long
    Dim tmp As Long

    ReDim picks(1 To k)

    If WithReplacement Then
        For i = 1 To k
            picks(i) = CLng(Application.WorksheetFunction.RandBetween(1, n))
        Next i
    Else
        ' Fisher–Yates shuffle, then take the first k indexes (no replacement, no infinite retry loop).
        ReDim idx(1 To n)
        For i = 1 To n
            idx(i) = i
        Next i
        For i = n To 2 Step -1
            j = CLng(Application.WorksheetFunction.RandBetween(1, i))
            tmp = idx(i)
            idx(i) = idx(j)
            idx(j) = tmp
        Next i
        For i = 1 To k
            picks(i) = idx(i)
        Next i
    End If

    DrawRowIndexes = picks
End Function
