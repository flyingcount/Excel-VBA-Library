Attribute VB_Name = "modApiLorenz"
Option Explicit

' Public API: Personal Custom_Menu11_lorenzGini.
' Lorenz curve and Gini coefficient (FastExcel GINICOEFA algorithm).

Private Const SheetName As String = "Gini"

''' @Description: Lorenz curve and Gini coefficient from a single non-negative numeric column. Writes sheet Gini.
''' @Example: GiniPlot
Public Sub GiniPlot()
    Dim src As Range
    Dim ws As Worksheet
    Dim n As Long
    Dim i As Long
    Dim gini As Variant
    Dim orig As Range
    Dim ch As ChartObject
    On Error GoTo EH
    Set src = modInternalPlots.PromptRange("Select a single column of non-negative values.")
    If src Is Nothing Then Exit Sub
    If src.Columns.Count <> 1 Then
        MsgBox "Data must be in a single column.", vbExclamation, "Plots Charts"
        Exit Sub
    End If
    n = src.Rows.Count
    If n < 1 Then Exit Sub
    If Not modInternalPlots.RangeIsAllNumeric(src) Then
        MsgBox "Not all the data is numeric.", vbExclamation, "Plots Charts"
        Exit Sub
    End If
    If modInternalPlots.RangeHasNegative(src) Then
        MsgBox "There are negative values in the data set.", vbExclamation, "Plots Charts"
        Exit Sub
    End If
    Call modInternalExcelApp.PushAppState
    Set ws = modInternalPlots.OutputSheet(SheetName)
    ws.Range("A1").Value = "Lorenz curve and Gini coefficient"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 12
    Call modInternalPlots.FlattenToColumn(src, ws.Range("M6"))
    Set orig = ws.Range("M6").Resize(n, 1)
    gini = GiniCoefficient(orig)
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "OriginalData", orig)
    ws.Range("I6").Value = "Gini"
    ws.Range("J6").Value = gini
    ws.Range("I7").Value = "Count, N"
    ws.Range("J7").Formula = "=COUNTA(OriginalData)"
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "Count_N", ws.Range("J7"))
    ws.Range("L4").Value = "#"
    ws.Range("M4").Value = "Original data"
    ws.Range("N4").Value = "Original data sorted"
    ws.Range("O4").Value = "Original data sorted %"
    ws.Range("P4").Value = "Original data sorted Cumm %"
    ws.Range("Q4").Value = "Equality line"
    ws.Range("P3").Value = "y-values"
    ws.Range("Q3").Value = "x-values"
    ws.Range("L5").Value = 0
    ws.Range("N5").Value = 0
    ws.Range("O5").Value = 0
    ws.Range("P5").Value = 0
    ws.Range("Q5").Value = 0
    For i = 1 To n
        ws.Cells(5 + i, 12).FormulaR1C1 = "=R[-1]C+1"
        ws.Cells(5 + i, 14).FormulaR1C1 = "=IF(RC[-1]="""","""",SMALL(OriginalData,RC[-2]))"
        ws.Cells(5 + i, 15).FormulaR1C1 = "=RC[-1]/SUM(OriginalDataSorted)"
        ws.Cells(5 + i, 16).FormulaR1C1 = "=RC[-1]+R[-1]C"
        ws.Cells(5 + i, 17).FormulaR1C1 = "=R[-1]C+1/Count_N"
    Next i
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "OriginalDataSorted", orig.Offset(0, 1))
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "GiniXValues", orig.Offset(-1, 4).Resize(n + 1, 1))
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "GiniYValues", orig.Offset(-1, 3).Resize(n + 1, 1))
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "EqualityLine", orig.Offset(-1, 4).Resize(n + 1, 1))
    ws.Range("L3:Q4").HorizontalAlignment = xlCenter
    ws.Range("L3:Q4").WrapText = True
    ws.Columns(11).ColumnWidth = 1
    ws.Columns(12).ColumnWidth = 6
    ws.Range(ws.Columns(13), ws.Columns(17)).ColumnWidth = 6
    ws.Range(ws.Columns(13), ws.Columns(17)).NumberFormat = "0.00"
    ws.Range("I6:J6").BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
    Set ch = modInternalPlots.AddChartAt(ws, ws.Range("A3"), 350, 250)
    With ch.Chart
        .HasTitle = True
        .ChartTitle.Text = "Gini plot"
        .HasLegend = True
        .Legend.Position = xlLegendPositionTop
        Call modInternalPlots.AddXySeries(ch.Chart, "Lorenz", ws.Range("GiniXValues"), ws.Range("GiniYValues"), xlXYScatterSmoothNoMarkers, RGB(0, 0, 0), False)
        Call modInternalPlots.AddXySeries(ch.Chart, "Equality", ws.Range("GiniXValues"), ws.Range("EqualityLine"), xlXYScatterSmoothNoMarkers, RGB(150, 150, 150), False)
        Call modInternalPlots.StyleValueAxes(ch.Chart, "", "", "0.0", "0.00", True)
        .Axes(xlValue).MaximumScale = 1
        .HasLegend = True
    End With
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("GiniPlot")
End Sub

' FastExcel GINICOEFA: https://fastexcel.wordpress.com/2011/05/21/fast-gini/
Private Function GiniCoefficient(ByVal inputRng As Range) As Variant
    Dim vals As Variant
    Dim j As Long
    Dim jRank As Long
    Dim dObs As Double
    Dim dGini As Double
    Dim dSum As Double
    On Error GoTo Fail
    If Application.WorksheetFunction.CountA(inputRng) = 0 Then Exit Function
    vals = inputRng.Value2
    If Not IsArray(vals) Then
        GiniCoefficient = 0
        Exit Function
    End If
    dObs = UBound(vals, 1)
    If dObs < 2 Then
        GiniCoefficient = 0
        Exit Function
    End If
    Call QSortVar(vals, 1, CLng(dObs))
    For j = 1 To dObs
        jRank = dObs - j + 1
        dGini = dGini + vals(j, 1) * jRank
        dSum = dSum + vals(j, 1)
    Next j
    If dSum = 0 Then
        GiniCoefficient = 0
        Exit Function
    End If
    GiniCoefficient = (dObs + 1) / (dObs - 1) - dGini * 2# / (dObs * (dObs - 1) * (dSum / dObs))
    Exit Function
Fail:
    GiniCoefficient = CVErr(xlErrValue)
End Function

Private Sub QSortVar(ByRef inputValues As Variant, ByVal jStart As Long, ByVal jEnd As Long)
    Dim jStart2 As Long
    Dim jEnd2 As Long
    Dim v1 As Variant
    Dim v2 As Variant
    jStart2 = jStart
    jEnd2 = jEnd
    v1 = inputValues((jStart + (jEnd - jStart) * Rnd()), 1)
    While jStart2 < jEnd2
        While inputValues(jStart2, 1) < v1 And jStart2 < jEnd
            jStart2 = jStart2 + 1
        Wend
        While inputValues(jEnd2, 1) > v1 And jEnd2 > jStart
            jEnd2 = jEnd2 - 1
        Wend
        If jStart2 < jEnd2 Then
            v2 = inputValues(jStart2, 1)
            inputValues(jStart2, 1) = inputValues(jEnd2, 1)
            inputValues(jEnd2, 1) = v2
        End If
        If jStart2 <= jEnd2 Then
            jStart2 = jStart2 + 1
            jEnd2 = jEnd2 - 1
        End If
    Wend
    If jStart2 > jStart Then Call QSortVar(inputValues, jStart, jEnd2)
    If jStart2 < jEnd Then Call QSortVar(inputValues, jStart2, jEnd)
End Sub
