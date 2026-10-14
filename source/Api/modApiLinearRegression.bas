Attribute VB_Name = "modApiLinearRegression"
Option Explicit

' Public API: Personal Custom_Menu11_LinearRegression / LinearRegressV2.

Private Const SheetV1 As String = "Linear regression"
Private Const SheetV2 As String = "Linear regression V2"

''' @Description: Simple linear regression of two columns (x, y) with prediction intervals and a 50-point forecast.
''' @Example: LinearRegression
Public Sub LinearRegression()
    Dim src As Range
    Dim ws As Worksheet
    Dim n As Long
    Dim i As Long
    Dim actual As Range
    Dim fcst As Range
    Dim maxX As Double
    Dim ch As ChartObject
    On Error GoTo EH
    Set src = modInternalPlots.PromptRange("Select two columns: x then y.")
    If src Is Nothing Then Exit Sub
    If src.Columns.Count <> 2 Then
        MsgBox "Range must be two columns: x and y.", vbExclamation, "Plots Charts"
        Exit Sub
    End If
    n = src.Rows.Count
    If n < 3 Then
        MsgBox "Need at least 3 rows.", vbExclamation, "Plots Charts"
        Exit Sub
    End If
    If Not modInternalPlots.RangeIsAllNumeric(src) Then
        MsgBox "x and y must be numeric.", vbExclamation, "Plots Charts"
        Exit Sub
    End If
    Call modInternalExcelApp.PushAppState
    Set ws = modInternalPlots.OutputSheet(SheetV1)
    maxX = Application.WorksheetFunction.Max(src.Columns(1))
    For i = 1 To n
        ws.Cells(2 + i, 13).Value = src.Cells(i, 1).Value
        ws.Cells(2 + i, 14).Value = src.Cells(i, 2).Value
        ws.Cells(2 + i, 15).FormulaR1C1 = "=R10C2*SQRT((1/R7C2)+(((RC[-2]-R9C2)^2)/R11C2))"
        ws.Cells(2 + i, 16).FormulaR1C1 = "=RC[-2]-RC[-1]*R12C2"
        ws.Cells(2 + i, 17).FormulaR1C1 = "=RC[-3]+RC[-2]*R12C2"
        ws.Cells(2 + i, 18).Value = "Actual"
    Next i
    Set actual = ws.Range("M3").Resize(n, 6)
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "XDataActual", actual.Columns(1))
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "YDataActual", actual.Columns(2))
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "LowerActual", actual.Columns(4))
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "UpperActual", actual.Columns(5))
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=actual.Columns(1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange actual.Resize(n, 2)
        .Header = xlNo
        .Apply
    End With
    For i = 1 To 50
        ws.Cells(2 + n + i, 13).FormulaR1C1 = "=R[-1]C+R22C2"
        ws.Cells(2 + n + i, 14).FormulaR1C1 = "=FORECAST.LINEAR(RC[-1],YDataActual,XDataActual)"
        ws.Cells(2 + n + i, 15).FormulaR1C1 = "=R10C2*SQRT(1+(1/R7C2)+(((RC[-2]-R9C2)^2)/R11C2))"
        ws.Cells(2 + n + i, 16).FormulaR1C1 = "=RC[-2]-RC[-1]*R12C2"
        ws.Cells(2 + n + i, 17).FormulaR1C1 = "=RC[-3]+RC[-2]*R12C2"
        ws.Cells(2 + n + i, 18).Value = "Forecast"
    Next i
    Set fcst = ws.Range("M3").Offset(n, 0).Resize(50, 6)
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "XDataForecast", fcst.Columns(1))
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "YDataForecast", fcst.Columns(2))
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "LowerForecast", fcst.Columns(4))
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "UpperForecast", fcst.Columns(5))
    Call WriteRegStats(ws, "XDataActual", "YDataActual")
    Call FormatRegSheet(ws, maxX, 10, "M2")
    Set ch = modInternalPlots.AddChartAt(ws, ws.Range("D2"), 400, 300)
    Call PlotRegChart(ch, actual, fcst, False)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("LinearRegression")
End Sub

''' @Description: Linear regression v2: original points, integer-x fitted grid, 100-point forecast, residuals plot, slope CI.
''' @Example: LinearRegressionV2
Public Sub LinearRegressionV2()
    Dim src As Range
    Dim ws As Worksheet
    Dim n As Long
    Dim i As Long
    Dim minX As Double
    Dim maxX As Double
    Dim m As Long
    Dim x As Long
    Dim orig As Range
    Dim actual As Range
    Dim fcst As Range
    Dim ch As ChartObject
    On Error GoTo EH
    Set src = modInternalPlots.PromptRange("Select two columns: x then y.")
    If src Is Nothing Then Exit Sub
    If src.Columns.Count <> 2 Then
        MsgBox "Range must be two columns: x and y.", vbExclamation, "Plots Charts"
        Exit Sub
    End If
    n = src.Rows.Count
    If n < 3 Then
        MsgBox "Need at least 3 rows.", vbExclamation, "Plots Charts"
        Exit Sub
    End If
    If Not modInternalPlots.RangeIsAllNumeric(src) Then
        MsgBox "x and y must be numeric.", vbExclamation, "Plots Charts"
        Exit Sub
    End If
    Call modInternalExcelApp.PushAppState
    Set ws = modInternalPlots.OutputSheet(SheetV2)
    minX = Application.WorksheetFunction.RoundDown(Application.WorksheetFunction.Min(src.Columns(1)), 0)
    maxX = Application.WorksheetFunction.RoundUp(Application.WorksheetFunction.Max(src.Columns(1)), 0)
    For i = 1 To n
        ws.Cells(2 + i, 25).Value = src.Cells(i, 1).Value
        ws.Cells(2 + i, 26).Value = src.Cells(i, 2).Value
        ws.Cells(2 + i, 23).FormulaR1C1 = IIf(i = 1, "1", "=R[-1]C+1")
        ws.Cells(2 + i, 27).FormulaR1C1 = "=RC[-2]*R2C15+R7C15"
        ws.Cells(2 + i, 28).FormulaR1C1 = "=RC[-2]-RC[-1]"
    Next i
    Set orig = ws.Range("Y3").Resize(n, 2)
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "XDataOriginal", orig.Columns(1))
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "YDataOriginal", orig.Columns(2))
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "XDataResidual", ws.Range("W3").Resize(n, 1))
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "YDataResidual", ws.Range("AB3").Resize(n, 1))
    m = CLng(maxX - minX) + 1
    If m < 2 Then m = 2
    For i = 1 To m
        x = CLng(minX) + i - 1
        ws.Cells(2 + i, 17).Value = x
        ws.Cells(2 + i, 18).FormulaR1C1 = "=FORECAST.LINEAR(RC[-1],YDataOriginal,XDataOriginal)"
        ws.Cells(2 + i, 19).FormulaR1C1 = "=R10C2*SQRT((1/R7C2)+(((RC[-2]-R9C2)^2)/R11C2))"
        ws.Cells(2 + i, 20).FormulaR1C1 = "=RC[-2]-RC[-1]*R12C2"
        ws.Cells(2 + i, 21).FormulaR1C1 = "=RC[-3]+RC[-2]*R12C2"
        ws.Cells(2 + i, 22).Value = "Calculated"
    Next i
    Set actual = ws.Range("Q3").Resize(m, 6)
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "XDataActual", actual.Columns(1))
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "YDataActual", actual.Columns(2))
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "LowerActual", actual.Columns(4))
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "UpperActual", actual.Columns(5))
    For i = 1 To 100
        ws.Cells(2 + m + i, 17).FormulaR1C1 = "=R[-1]C+R22C2"
        ws.Cells(2 + m + i, 18).FormulaR1C1 = "=FORECAST.LINEAR(RC[-1],YDataActual,XDataActual)"
        ws.Cells(2 + m + i, 19).FormulaR1C1 = "=R10C2*SQRT(1+(1/R7C2)+(((RC[-2]-R9C2)^2)/R11C2))"
        ws.Cells(2 + m + i, 20).FormulaR1C1 = "=RC[-2]-RC[-1]*R12C2"
        ws.Cells(2 + m + i, 21).FormulaR1C1 = "=RC[-3]+RC[-2]*R12C2"
        ws.Cells(2 + m + i, 22).Value = "Forecast"
    Next i
    Set fcst = ws.Range("Q3").Offset(m, 0).Resize(100, 6)
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "XDataForecast", fcst.Columns(1))
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "YDataForecast", fcst.Columns(2))
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "LowerForecast", fcst.Columns(4))
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "UpperForecast", fcst.Columns(5))
    Call WriteRegStats(ws, "XDataOriginal", "YDataOriginal")
    Call WriteSlopeBlock(ws, "N1", "YDataOriginal", "XDataOriginal")
    Call WriteSlopeBlock(ws, "N21", "YDataResidual", "XDataResidual")
    Call FormatRegSheet(ws, maxX, 1, "Q2")
    ws.Range("Y1").Value = "Original data"
    ws.Range("Y1").Font.Bold = True
    Call modInternalPlots.WriteHeaderRow(ws.Range("X2"), Array("#", "x", "y", "y predicted", "y residual"))
    Set ch = modInternalPlots.AddChartAt(ws, ws.Range("D1"), 400, 300)
    Call PlotRegChart(ch, actual, fcst, True)
    Call modInternalPlots.AddXySeries(ch.Chart, "Original data", orig.Columns(1), orig.Columns(2), xlXYScatter, RGB(0, 0, 0), True)
    Set ch = modInternalPlots.AddChartAt(ws, ws.Range("D21"), 400, 300)
    ch.Chart.HasTitle = True
    ch.Chart.ChartTitle.Text = "Y residuals"
    Call modInternalPlots.AddXySeries(ch.Chart, "Residuals", ws.Range("XDataResidual"), ws.Range("YDataResidual"), xlXYScatter, RGB(0, 0, 0), True)
    ch.Chart.SeriesCollection(1).Trendlines.Add
    ch.Chart.SeriesCollection(1).Trendlines(1).DisplayEquation = True
    ch.Chart.SeriesCollection(1).Trendlines(1).DisplayRSquared = True
    Call modInternalPlots.StyleValueAxes(ch.Chart, "", "y residual", "0.0", "0.0", False)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("LinearRegressionV2")
End Sub

Private Sub WriteRegStats(ByVal ws As Worksheet, ByVal xName As String, ByVal yName As String)
    ws.Range("A5").Value = "Confidence level"
    ws.Range("B5").Value = 0.95
    Call modInternalPlots.HighlightYellow(ws.Range("B5"))
    ws.Range("A6").Value = "Alpha"
    ws.Range("B6").FormulaR1C1 = "=1-R[-1]C"
    ws.Range("A7").Value = "Number of data points, n"
    ws.Range("B7").Formula = "=COUNTA(" & xName & ")"
    ws.Range("A8").Value = "Degrees of freedom, df"
    ws.Range("B8").FormulaR1C1 = "=R[-1]C-2"
    ws.Range("A9").Value = "Mean of x"
    ws.Range("B9").Formula = "=AVERAGE(" & xName & ")"
    ws.Range("A10").Value = "Standard error of Y given X, STEyx"
    ws.Range("B10").Formula = "=STEYX(" & yName & "," & xName & ")"
    ws.Range("A11").Value = "Variance, SSx"
    ws.Range("B11").Formula = "=DEVSQ(" & xName & ")"
    ws.Range("A12").Value = "t-critical value"
    ws.Range("B12").FormulaR1C1 = "=T.INV.2T(R[-6]C,R[-4]C)"
    ws.Range("A13").Value = "Pearson R2"
    ws.Range("B13").Formula = "=RSQ(" & yName & "," & xName & ")"
End Sub

Private Sub WriteSlopeBlock(ByVal ws As Worksheet, ByVal topLeft As String, ByVal yName As String, ByVal xName As String)
    Dim r As Range
    Set r = ws.Range(topLeft)
    r.Value = "Slope stats"
    r.Font.Bold = True
    r.Offset(1, 0).Value = "Slope"
    r.Offset(1, 1).Formula = "=SLOPE(" & yName & "," & xName & ")"
    r.Offset(2, 0).Value = "Margin of error, ME"
    r.Offset(2, 1).FormulaR1C1 = "=R[11]C*R[12]C"
    r.Offset(3, 0).Value = "Confidence interval"
    r.Offset(3, 1).FormulaR1C1 = "=R[-1]C*2"
    r.Offset(4, 0).Value = "Lower"
    r.Offset(4, 1).FormulaR1C1 = "=R[-3]C-R[-2]C"
    r.Offset(5, 0).Value = "Upper"
    r.Offset(5, 1).FormulaR1C1 = "=R[-4]C+R[-3]C"
    r.Offset(6, 0).Value = "Intercept"
    r.Offset(6, 1).Formula = "=INTERCEPT(" & yName & "," & xName & ")"
    r.Offset(8, 0).Value = "Confidence level"
    r.Offset(8, 1).FormulaR1C1 = "=R5C2"
    r.Offset(9, 0).Value = "Alpha"
    r.Offset(9, 1).FormulaR1C1 = "=1-R[-1]C"
    r.Offset(10, 0).Value = "Critical probability"
    r.Offset(10, 1).FormulaR1C1 = "=1-R[-1]C/2"
    r.Offset(11, 0).Value = "N"
    r.Offset(11, 1).Formula = "=COUNTA(" & xName & ")"
    r.Offset(12, 0).Value = "df"
    r.Offset(12, 1).FormulaR1C1 = "=R[-1]C-2"
    r.Offset(13, 0).Value = "T stat critical"
    r.Offset(13, 1).FormulaR1C1 = "=T.INV(R[-4]C,R[-1]C)"
    r.Offset(14, 0).Value = "Standard error, SE"
    r.Offset(14, 1).FormulaR1C1 = "=SQRT(DEVSQ(" & yName & ")/R[-2]C)/SQRT(DEVSQ(" & xName & "))"
End Sub

Private Sub FormatRegSheet(ByVal ws As Worksheet, ByVal maxX As Double, ByVal inc As Double, ByVal hdrCell As String)
    ws.Range("A1").Value = ws.Name
    ws.Range("A1").Font.Bold = True
    ws.Range("A15").Value = "Enter x to get y and lower & upper error margins"
    ws.Range("A16").Value = "x"
    ws.Range("B16").Value = maxX + 1
    Call modInternalPlots.HighlightYellow(ws.Range("B16"))
    ws.Range("A17").Value = "y"
    ws.Range("B17").FormulaR1C1 = "=FORECAST.LINEAR(R[-1]C,YDataActual,XDataActual)"
    ws.Range("A18").Value = "Standard error prediction, SEpred"
    ws.Range("B18").FormulaR1C1 = "=R10C2*SQRT((1+1/R7C2)+(((R[-2]C-R9C2)^2)/R11C2))"
    ws.Range("A19").Value = "Lower"
    ws.Range("B19").FormulaR1C1 = "=R[-2]C-R[-1]C*R12C2"
    ws.Range("A20").Value = "Upper"
    ws.Range("B20").FormulaR1C1 = "=R[-3]C+R[-2]C*R12C2"
    ws.Range("A22").Value = "Forecast increment"
    ws.Range("B22").Value = inc
    Call modInternalPlots.HighlightYellow(ws.Range("B22"))
    Call modInternalPlots.WriteHeaderRow(ws.Range(hdrCell), Array("x", "y", "Standard error, SE", "Lower", "Upper", "Data type"))
End Sub

Private Sub PlotRegChart(ByVal chObj As ChartObject, ByVal actual As Range, ByVal fcst As Range, ByVal smoothActual As Boolean)
    Dim chtType As XlChartType
    If smoothActual Then chtType = xlXYScatterSmoothNoMarkers Else chtType = xlXYScatter
    With chObj.Chart
        .HasTitle = True
        .ChartTitle.Text = "Linear regression and forecast"
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        Call modInternalPlots.AddXySeries(chObj.Chart, "Actuals", actual.Columns(1), actual.Columns(2), chtType, RGB(0, 0, 255), Not smoothActual)
        If Not smoothActual Then
            .SeriesCollection(1).Trendlines.Add
            .SeriesCollection(1).Trendlines(1).DisplayEquation = True
            .SeriesCollection(1).Trendlines(1).DisplayRSquared = True
        End If
        Call modInternalPlots.AddXySeries(chObj.Chart, "Actuals Lower", actual.Columns(1), actual.Columns(4), xlXYScatterSmoothNoMarkers, RGB(180, 180, 180), False)
        Call modInternalPlots.AddXySeries(chObj.Chart, "Actuals Upper", actual.Columns(1), actual.Columns(5), xlXYScatterSmoothNoMarkers, RGB(180, 180, 180), False)
        Call modInternalPlots.AddXySeries(chObj.Chart, "Forecast", fcst.Columns(1), fcst.Columns(2), xlXYScatterSmoothNoMarkers, RGB(0, 150, 0), False)
        Call modInternalPlots.AddXySeries(chObj.Chart, "Forecast Lower", fcst.Columns(1), fcst.Columns(4), xlXYScatterSmoothNoMarkers, RGB(180, 180, 180), False)
        Call modInternalPlots.AddXySeries(chObj.Chart, "Forecast Upper", fcst.Columns(1), fcst.Columns(5), xlXYScatterSmoothNoMarkers, RGB(180, 180, 180), False)
        Call modInternalPlots.StyleValueAxes(chObj.Chart, "x", "y", "0.0", "0.0", False)
        .HasLegend = True
    End With
End Sub
