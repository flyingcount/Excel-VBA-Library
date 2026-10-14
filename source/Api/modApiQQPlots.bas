Attribute VB_Name = "modApiQQPlots"
Option Explicit

' Public API: Personal Custom_Menu11_QQPlotNormal / QQPlotUniform.

Private Const SheetNormal As String = "QQ Normal chart"
Private Const SheetUniform As String = "QQ Uniform chart"

''' @Description: Normal QQ plot of a numeric range (flattened if multi-column). Writes sheet QQ Normal chart.
''' @Example: QQPlotGaussianNormal
Public Sub QQPlotGaussianNormal()
    Dim src As Range
    Dim ws As Worksheet
    Dim n As Long
    Dim i As Long
    Dim dataRng As Range
    Dim ch As ChartObject
    On Error GoTo EH
    Set src = modInternalPlots.PromptRange("Select the numeric data for the QQ plot.")
    If src Is Nothing Then Exit Sub
    n = src.Cells.Count
    If n < 3 Then
        MsgBox "Need at least 3 data points.", vbExclamation, "Plots Charts"
        Exit Sub
    End If
    If Not modInternalPlots.RangeIsAllNumeric(src) Then
        MsgBox "Every cell must be numeric (blanks are not allowed).", vbExclamation, "Plots Charts"
        Exit Sub
    End If
    Call modInternalExcelApp.PushAppState
    Set ws = modInternalPlots.OutputSheet(SheetNormal)
    ws.Range("A1").Value = "Normal QQ Plot data"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 12
    Call modInternalPlots.WriteHeaderRow(ws.Range("A6"), Array("Original Data", "Rank", "Rank proportion", "Rank-based Z Score", "Fitted y", "Residuals"))
    Call modInternalPlots.FlattenToColumn(src, ws.Range("A7"))
    Set dataRng = ws.Range("A7").Resize(n, 1)
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=dataRng, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange dataRng
        .Header = xlNo
        .Apply
    End With
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "OriginalData", dataRng)
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "Rank", dataRng.Offset(0, 1))
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "ZScores", dataRng.Offset(0, 3))
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "Residuals", dataRng.Offset(0, 5))
    For i = 1 To n
        ws.Cells(6 + i, 2).Value = i
        ws.Cells(6 + i, 3).FormulaR1C1 = "=RC[-1]/(COUNT(OriginalData)+1)"
        ws.Cells(6 + i, 4).FormulaR1C1 = "=NORMSINV(RC[-1])"
        ws.Cells(6 + i, 5).FormulaR1C1 = "=RC[-1]*R3C17+R4C17"
        ws.Cells(6 + i, 6).FormulaR1C1 = "=RC[-1]-RC[-5]"
    Next i
    ws.Range("A2").Value = "Number of data points"
    ws.Range("B2").Value = n
    ws.Range("A3").Value = "Mean"
    ws.Range("B3").Formula = "=AVERAGE(OriginalData)"
    ws.Range("A4").Value = "Standard deviation"
    ws.Range("B4").Formula = "=STDEV.S(OriginalData)"
    ws.Range("P2").Value = "QQ line"
    ws.Range("P2").Font.Bold = True
    ws.Range("P3").Value = "Slope"
    ws.Range("Q3").Formula = "=SLOPE(OriginalData,ZScores)"
    ws.Range("P4").Value = "Intercept"
    ws.Range("Q4").Formula = "=INTERCEPT(OriginalData,ZScores)"
    ws.Range("P5").Value = "R-squared"
    ws.Range("Q5").Formula = "=RSQ(OriginalData,ZScores)"
    ws.Range("P20").Value = "Residuals line"
    ws.Range("P20").Font.Bold = True
    ws.Range("P21").Value = "Slope"
    ws.Range("Q21").Formula = "=SLOPE(Residuals,Rank)"
    ws.Range("P22").Value = "Intercept"
    ws.Range("Q22").Formula = "=INTERCEPT(Residuals,Rank)"
    ws.Range("P23").Value = "R-squared"
    ws.Range("Q23").Formula = "=RSQ(Residuals,Rank)"
    ws.Columns("A").ColumnWidth = 12
    ws.Range("B:F").ColumnWidth = 7
    ws.Range("C:F").NumberFormat = "0.0000"
    Set ch = modInternalPlots.AddChartAt(ws, ws.Range("H1"), 375, 300)
    With ch.Chart
        .HasTitle = True
        .ChartTitle.Text = "QQ Plot - Normal Distribution"
        .HasLegend = False
        Call modInternalPlots.AddXySeries(ch.Chart, "Original Data", ws.Range("ZScores"), dataRng, xlXYScatter, RGB(0, 0, 0), True)
        .SeriesCollection(1).Trendlines.Add
        .SeriesCollection(1).Trendlines(1).DisplayEquation = True
        .SeriesCollection(1).Trendlines(1).DisplayRSquared = True
        Call modInternalPlots.StyleValueAxes(ch.Chart, "", "", "0.00", "0.00", False)
    End With
    Set ch = modInternalPlots.AddChartAt(ws, ws.Range("H22"), 375, 300)
    With ch.Chart
        .HasTitle = True
        .ChartTitle.Text = "Normal QQ plot residuals"
        .HasLegend = False
        Call modInternalPlots.AddXySeries(ch.Chart, "Residuals", ws.Range("Rank"), ws.Range("Residuals"), xlXYScatter, RGB(0, 0, 0), True)
        Call modInternalPlots.StyleValueAxes(ch.Chart, "", "", "0", "0.00", False)
    End With
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("QQPlotGaussianNormal")
End Sub

''' @Description: Uniform QQ plot: sorted values vs a linear grid from min to max. Writes sheet QQ Uniform chart.
''' @Example: QQPlotUniform
Public Sub QQPlotUniform()
    Dim src As Range
    Dim ws As Worksheet
    Dim n As Long
    Dim i As Long
    Dim data As Variant
    Dim ch As ChartObject
    Dim rngData As Range
    On Error GoTo EH
    Set src = modInternalPlots.PromptRange("Select the numeric data for the QQ plot.")
    If src Is Nothing Then Exit Sub
    n = src.Cells.Count
    If n < 3 Then
        MsgBox "Need at least 3 data points.", vbExclamation, "Plots Charts"
        Exit Sub
    End If
    If Not modInternalPlots.RangeIsAllNumeric(src) Then
        MsgBox "Every cell must be numeric (blanks are not allowed).", vbExclamation, "Plots Charts"
        Exit Sub
    End If
    Call modInternalExcelApp.PushAppState
    Set ws = modInternalPlots.OutputSheet(SheetUniform)
    ws.Range("A1").Value = "Uniform distribution QQ Plot data"
    ws.Range("A1").Font.Bold = True
    Call modInternalPlots.WriteHeaderRow(ws.Range("Q9"), Array("Original data", "Original data sorted, y", "Rank", "Theoretical values, x", "Fitted y", "Residuals"))
    ReDim data(1 To n, 1 To 6)
    For i = 1 To n
        data(i, 1) = src.Cells(i).Value
        data(i, 2) = "=IF(RC[-1]="""","""",SMALL(OriginalData,ROW()-ROW(R9C17)))"
        data(i, 3) = i
        If i = 1 Then
            data(i, 4) = "=RC[-2]"
        Else
            data(i, 4) = "=R[-1]C+Increment"
        End If
        data(i, 5) = "=RC[-1]*SlopeFitted+InterceptFitted"
        data(i, 6) = "=RC[-4]-RC[-1]"
    Next i
    Call modInternalPlots.WriteBlock(ws, 10, 17, data)
    Set rngData = ws.Range("Q10").Resize(n, 1)
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "OriginalData", rngData)
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "SortedOriginalData", rngData.Offset(0, 1))
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "Rank", rngData.Offset(0, 2))
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "TheoreticalQuantile", rngData.Offset(0, 3))
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "Residuals", rngData.Offset(0, 5))
    ws.Range("Q2").Value = "Data statistics"
    ws.Range("Q2").Font.Bold = True
    ws.Range("Q3").Value = "Count"
    ws.Range("R3").Formula = "=COUNT(OriginalData)"
    ws.Range("Q4").Value = "Maximum"
    ws.Range("R4").Formula = "=MAX(OriginalData)"
    ws.Range("Q5").Value = "Minimum"
    ws.Range("R5").Formula = "=MIN(OriginalData)"
    ws.Range("Q6").Value = "Data range"
    ws.Range("R6").FormulaR1C1 = "=R[-2]C-R[-1]C"
    ws.Range("Q7").Value = "Increment"
    ws.Range("R7").FormulaR1C1 = "=R[-1]C/(R[-4]C-1)"
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "Increment", ws.Range("R7"))
    ws.Range("A17").Value = "Slope"
    ws.Range("B17").Formula = "=SLOPE(SortedOriginalData,TheoreticalQuantile)"
    ws.Range("A18").Value = "Intercept"
    ws.Range("B18").Formula = "=INTERCEPT(SortedOriginalData,TheoreticalQuantile)"
    ws.Range("A19").Value = "R-squared"
    ws.Range("B19").Formula = "=RSQ(SortedOriginalData,TheoreticalQuantile)"
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "SlopeFitted", ws.Range("B17"))
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "InterceptFitted", ws.Range("B18"))
    ws.Range("I17").Value = "Slope"
    ws.Range("J17").Formula = "=SLOPE(Residuals,Rank)"
    ws.Range("I18").Value = "Intercept"
    ws.Range("J18").Formula = "=INTERCEPT(Residuals,Rank)"
    ws.Range("I19").Value = "R-squared"
    ws.Range("J19").Formula = "=RSQ(Residuals,Rank)"
    ws.Range("Q:V").Font.Size = 6
    ws.Range("Q:V").ColumnWidth = 6
    Set ch = modInternalPlots.AddChartAt(ws, ws.Range("A2"), 375, 250)
    ch.Chart.HasTitle = True
    ch.Chart.ChartTitle.Text = "QQ Plot - Uniform Distribution"
    Call modInternalPlots.AddXySeries(ch.Chart, "QQ", ws.Range("TheoreticalQuantile"), ws.Range("SortedOriginalData"), xlXYScatter, RGB(0, 0, 0), True)
    ch.Chart.SeriesCollection(1).Trendlines.Add
    ch.Chart.SeriesCollection(1).Trendlines(1).DisplayEquation = True
    ch.Chart.SeriesCollection(1).Trendlines(1).DisplayRSquared = True
    Call modInternalPlots.StyleValueAxes(ch.Chart, "Theoretical Data", "Actual", "0.00", "0.00", False)
    Set ch = modInternalPlots.AddChartAt(ws, ws.Range("I2"), 375, 250)
    ch.Chart.HasTitle = True
    ch.Chart.ChartTitle.Text = "Uniform QQ plot residuals"
    Call modInternalPlots.AddXySeries(ch.Chart, "Residuals", ws.Range("Rank"), ws.Range("Residuals"), xlXYScatter, RGB(0, 0, 0), True)
    Call modInternalPlots.StyleValueAxes(ch.Chart, "", "", "0", "0.00", False)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("QQPlotUniform")
End Sub
