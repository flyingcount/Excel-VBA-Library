Attribute VB_Name = "modApiHistogram"
Option Explicit

' Public API: Personal Custom_Menu11_Histogram / HistogramFormulae.
' Histogram table + column chart on a new sheet. Named ranges Histogram_Data (input)
' and HistBins / HistFrequency (output) are workbook-scoped so the chart can find them.

Private Const SheetValues As String = "Histogram"
Private Const SheetFormulae As String = "Histogram formulae"
Private Const TableRow As Long = 12

''' @Description: Frequency histogram (VBA counts) from a numeric range. Prompts for bin count.
''' @Example: HistogramTableAndPlot
Public Sub HistogramTableAndPlot()
    Dim src As Range
    Dim nBins As Variant
    On Error GoTo EH
    Set src = modInternalPlots.PromptRange("Select the numeric data to bin.")
    If src Is Nothing Then Exit Sub
    If Application.WorksheetFunction.Count(src) = 0 Then
        MsgBox "Select numeric data.", vbExclamation, "Plots Charts"
        Exit Sub
    End If
    nBins = modInternalPlots.PromptNumber("How many bins?", 10)
    If IsEmpty(nBins) Then Exit Sub
    If CLng(nBins) < 1 Then nBins = 10
    Call modInternalExcelApp.PushAppState
    Call BuildCountedHistogram(src, CLng(nBins))
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("HistogramTableAndPlot")
End Sub

''' @Description: Frequency histogram with COUNTIF formulae so bins update when the source data changes.
''' @Example: HistogramFormulaeAndPlot
Public Sub HistogramFormulaeAndPlot()
    Dim src As Range
    Dim nBins As Variant
    On Error GoTo EH
    Set src = modInternalPlots.PromptRange("Select the numeric data to bin.")
    If src Is Nothing Then Exit Sub
    If Application.WorksheetFunction.Count(src) = 0 Then
        MsgBox "Select numeric data.", vbExclamation, "Plots Charts"
        Exit Sub
    End If
    nBins = modInternalPlots.PromptNumber("How many bins?", 10)
    If IsEmpty(nBins) Then Exit Sub
    If CLng(nBins) < 1 Then nBins = 10
    Call modInternalExcelApp.PushAppState
    Call BuildFormulaHistogram(src, CLng(nBins))
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("HistogramFormulaeAndPlot")
End Sub

Private Sub BuildCountedHistogram(ByVal src As Range, ByVal nBins As Long)
    Dim ws As Worksheet
    Dim n As Long
    Dim nRows As Long
    Dim i As Long
    Dim j As Long
    Dim v As Double
    Dim binStart As Double
    Dim minV As Double
    Dim maxV As Double
    Dim dataRange As Double
    Dim binWidth As Double
    Dim tbl As Variant
    Dim lo As Double
    Dim hi As Double
    Dim rngOut As Range
    n = src.Cells.Count
    minV = Application.WorksheetFunction.Min(src)
    maxV = Application.WorksheetFunction.Max(src)
    binStart = Application.WorksheetFunction.RoundDown(minV, 0)
    maxV = Application.WorksheetFunction.RoundUp(maxV, 0)
    dataRange = maxV - minV
    If dataRange / nBins < 1 Then
        binWidth = dataRange / nBins
    Else
        binWidth = Application.WorksheetFunction.Round(dataRange / nBins, 0)
        If binWidth = 0 Then binWidth = dataRange / nBins
    End If
    nRows = nBins + 2
    ReDim tbl(1 To nRows, 1 To 4)
    For i = 1 To nRows
        tbl(i, 1) = i
        tbl(i, 4) = 0
        Select Case i
            Case 1
                tbl(i, 2) = "<"
                tbl(i, 3) = binStart
            Case 2
                tbl(i, 2) = binStart
                tbl(i, 3) = binStart + binWidth
            Case nRows
                tbl(i, 2) = ">"
                tbl(i, 3) = binStart + nBins * binWidth
            Case Else
                tbl(i, 2) = tbl(i - 1, 3) + 0.000001
                tbl(i, 3) = tbl(i - 1, 3) + binWidth
        End Select
    Next i
    For i = 1 To n
        If IsNumeric(src.Cells(i).Value) Then
            If Len(CStr(src.Cells(i).Value)) > 0 Then
                v = CDbl(src.Cells(i).Value)
                If v <= tbl(1, 3) Then
                    tbl(1, 4) = tbl(1, 4) + 1
                ElseIf v > tbl(nRows - 1, 3) Then
                    tbl(nRows, 4) = tbl(nRows, 4) + 1
                Else
                    For j = 2 To nRows - 1
                        lo = tbl(j, 2)
                        hi = tbl(j, 3)
                        If v > tbl(j - 1, 3) And v <= hi Then
                            tbl(j, 4) = tbl(j, 4) + 1
                            Exit For
                        End If
                    Next j
                End If
            End If
        End If
        Application.StatusBar = "Binning " & i & " of " & n
    Next i
    Set ws = modInternalPlots.OutputSheet(SheetValues)
    Call WriteHistogramChrome(ws, src, n, nBins, minV, maxV, binWidth)
    Call modInternalPlots.WriteBlock(ws, TableRow, 1, tbl)
    Set rngOut = ws.Range(ws.Cells(TableRow - 1, 1), ws.Cells(TableRow + nRows - 1, 4))
    Call NameHistogramRanges(src, rngOut)
    Call PlotHistogram(ws)
    Application.StatusBar = False
End Sub

Private Sub BuildFormulaHistogram(ByVal src As Range, ByVal nBins As Long)
    Dim ws As Worksheet
    Dim nRows As Long
    Dim i As Long
    Dim rngOut As Range
    nRows = nBins + 2
    Set ws = modInternalPlots.OutputSheet(SheetFormulae)
    Call modInternalNamedRanges.CreateNamedRange("Histogram_Data", src)
    ws.Range("A1").Value = "Histogram Data and Plot"
    ws.Range("A1").Font.Bold = True
    ws.Range("A3").Value = "Bin start"
    ws.Range("B3").Formula = "=ROUNDDOWN(MIN(Histogram_Data),0)"
    ws.Range("A5").Value = "Number of bins"
    ws.Range("B5").Value = nBins
    ws.Range("A6").Value = "Minimum data value"
    ws.Range("B6").Formula = "=MIN(Histogram_Data)"
    ws.Range("A7").Value = "Maximum data value"
    ws.Range("B7").Formula = "=MAX(Histogram_Data)"
    ws.Range("A8").Value = "Data range"
    ws.Range("B8").FormulaR1C1 = "=R7C2-R6C2"
    ws.Range("A4").Value = "Bin range"
    ws.Range("B4").Formula = "=IF(B8/B5<1,B8/B5,ROUND(B8/B5,0))"
    ws.Range("C3").Value = "<- change"
    ws.Range("C4").Value = "<- change"
    ws.Range("A9").Value = "Number of records"
    ws.Range("B9").Formula = "=COUNT(Histogram_Data)"
    ws.Range("C9").Value = "Check total"
    ws.Range("D9").Formula = "=SUM(HistFrequency)"
    ws.Range("E9").FormulaR1C1 = "=IF(RC[-3]=SUM(HistFrequency),""OK"",RC[-3]-SUM(HistFrequency))"
    ws.Cells(TableRow - 1, 1).Value = "Index"
    ws.Cells(TableRow - 1, 2).Value = "From"
    ws.Cells(TableRow - 1, 3).Value = "To"
    ws.Cells(TableRow - 1, 4).Value = "Frequency"
    For i = 1 To nRows
        ws.Cells(TableRow + i - 1, 1).Value = i
        Select Case i
            Case 1
                ws.Cells(TableRow, 2).Value = "<"
                ws.Cells(TableRow, 3).FormulaR1C1 = "=R3C2"
                ws.Cells(TableRow, 4).FormulaR1C1 = "=COUNTIF(Histogram_Data,""<=""&RC[-1])"
            Case 2
                ws.Cells(TableRow + 1, 2).FormulaR1C1 = "=R[-1]C[1]+0.000001"
                ws.Cells(TableRow + 1, 3).FormulaR1C1 = "=R3C2+R4C2"
                ws.Cells(TableRow + 1, 4).FormulaR1C1 = "=COUNTIFS(Histogram_Data,""<=""&RC[-1],Histogram_Data,"">""&R[-1]C[-1])"
            Case nRows
                ws.Cells(TableRow + nRows - 1, 2).Value = ">"
                ws.Cells(TableRow + nRows - 1, 3).FormulaR1C1 = "=IF(R3C2+R5C2*R4C2>R[-1]C,R[-1]C,R3C2+R5C2*R4C2)"
                ws.Cells(TableRow + nRows - 1, 4).FormulaR1C1 = "=COUNTIF(Histogram_Data,"">""&R[-1]C[-1])"
            Case Else
                ws.Cells(TableRow + i - 1, 2).FormulaR1C1 = "=R[-1]C[1]+0.000001"
                ws.Cells(TableRow + i - 1, 3).FormulaR1C1 = "=R[-1]C+R4C2"
                ws.Cells(TableRow + i - 1, 4).FormulaR1C1 = "=COUNTIFS(Histogram_Data,""<=""&RC[-1],Histogram_Data,"">""&R[-1]C[-1])"
        End Select
    Next i
    Set rngOut = ws.Range(ws.Cells(TableRow - 1, 1), ws.Cells(TableRow + nRows - 1, 4))
    Call NameHistogramRanges(src, rngOut)
    ws.Columns("A").AutoFit
    ws.Columns("C").AutoFit
    Call PlotHistogram(ws)
End Sub

Private Sub WriteHistogramChrome(ByVal ws As Worksheet, ByVal src As Range, ByVal n As Long, _
                                 ByVal nBins As Long, ByVal minV As Double, ByVal maxV As Double, _
                                 ByVal binWidth As Double)
    ws.Range("A1").Value = "Histogram Data and Plot"
    ws.Range("A1").Font.Bold = True
    ws.Range("A3").Value = "Bin start"
    ws.Range("B3").Value = minV
    ws.Range("A4").Value = "Bin range"
    ws.Range("B4").Value = binWidth
    ws.Range("A5").Value = "Number of bins"
    ws.Range("B5").Value = nBins
    ws.Range("A6").Value = "Minimum data value"
    ws.Range("B6").Value = minV
    ws.Range("A7").Value = "Maximum data value"
    ws.Range("B7").Value = maxV
    ws.Range("A8").Value = "Data range"
    ws.Range("B8").FormulaR1C1 = "=R[-1]C-R[-2]C"
    ws.Range("A9").Value = "Number of records"
    ws.Range("B9").Value = n
    ws.Range("C9").Value = "Check total"
    ws.Range("D9").Formula = "=SUM(HistFrequency)"
    ws.Range("E9").FormulaR1C1 = "=IF(RC[-3]=SUM(HistFrequency),""OK"",RC[-3]-SUM(HistFrequency))"
    ws.Cells(TableRow - 1, 1).Value = "Index"
    ws.Cells(TableRow - 1, 2).Value = "From"
    ws.Cells(TableRow - 1, 3).Value = "To"
    ws.Cells(TableRow - 1, 4).Value = "Frequency"
    ws.Columns("A").AutoFit
    ws.Columns("C").AutoFit
    Call modInternalNamedRanges.CreateNamedRange("Histogram_Data", src)
End Sub

Private Sub NameHistogramRanges(ByVal src As Range, ByVal rngOut As Range)
    Dim n As Long
    n = rngOut.Rows.Count - 1
    Call modInternalNamedRanges.CreateNamedRange("Histogram_Data", src)
    Call modInternalNamedRanges.CreateNamedRange("HistogramOutput", rngOut)
    Call modInternalNamedRanges.CreateNamedRange("HistBins", rngOut.Cells(2, 3).Resize(n, 1))
    Call modInternalNamedRanges.CreateNamedRange("HistFrequency", rngOut.Cells(2, 4).Resize(n, 1))
    Call modInternalNamedRanges.CreateNamedRange("HistPlotData", Union(rngOut.Cells(2, 3).Resize(n, 1), rngOut.Cells(2, 4).Resize(n, 1)))
End Sub

Private Sub PlotHistogram(ByVal ws As Worksheet)
    Dim ch As ChartObject
    Set ch = modInternalPlots.AddChartAt(ws, ws.Range("F1"), 325, 325)
    With ch.Chart
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Frequency distribution"
        .HasLegend = False
        .SeriesCollection.NewSeries
        With .SeriesCollection(1)
            .Name = "Frequency distribution"
            .XValues = ws.Parent.Names("HistBins").RefersToRange
            .Values = ws.Parent.Names("HistFrequency").RefersToRange
            .ChartType = xlColumnClustered
        End With
        With .Axes(xlCategory)
            .TickLabels.NumberFormat = "0.0"
            .TickLabels.Orientation = 90
            .HasMajorGridlines = False
            .HasMinorGridlines = False
        End With
        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Text = "Frequency"
            .TickLabels.NumberFormat = "0"
            .MinimumScale = 0
            .HasMajorGridlines = False
            .HasMinorGridlines = False
        End With
    End With
End Sub
