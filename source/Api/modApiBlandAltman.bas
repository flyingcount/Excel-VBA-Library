Attribute VB_Name = "modApiBlandAltman"
Option Explicit

' Public API: Personal Custom_Menu5_BlandAltman.
' Two numeric columns (method 1, method 2) → sheet BlandAltman with mean/difference and LoA plot.

Private Const SheetName As String = "BlandAltman"

''' @Description: Bland-Altman plot of two paired measurement columns (no header). Writes mean, difference, and ±1.96 SD limits of agreement.
''' @Example: BlandAltmanPlot
Public Sub BlandAltmanPlot()
    Dim src As Range
    Dim ws As Worksheet
    Dim n As Long
    Dim i As Long
    Dim raw As Variant
    Dim avg() As Double
    Dim diff() As Double
    Dim rngAvg As Range
    Dim rngDiff As Range
    Dim meanDiff As Double
    Dim sdDiff As Double
    Dim minX As Double
    Dim maxX As Double
    Dim ch As ChartObject
    On Error GoTo EH
    Set src = modInternalAnalysis.PromptNumericRange( _
        "Select two columns of paired measurements, without a header row.", _
        "Bland-Altman", 2, 2, 2)
    If src Is Nothing Then Exit Sub
    n = src.Rows.Count
    raw = src.Value
    ReDim avg(1 To n, 1 To 1)
    ReDim diff(1 To n, 1 To 1)
    For i = 1 To n
        avg(i, 1) = (CDbl(raw(i, 1)) + CDbl(raw(i, 2))) / 2
        diff(i, 1) = CDbl(raw(i, 1)) - CDbl(raw(i, 2))
    Next i
    meanDiff = Application.WorksheetFunction.Average(diff)
    If n < 2 Then
        MsgBox "Need at least two pairs to compute the standard deviation.", vbExclamation, "Bland-Altman"
        Exit Sub
    End If
    sdDiff = Application.WorksheetFunction.StDev_S(diff)
    minX = Application.WorksheetFunction.Min(avg)
    maxX = Application.WorksheetFunction.Max(avg)
    If minX = maxX Then
        minX = minX - 1
        maxX = maxX + 1
    End If
    Call modInternalExcelApp.PushAppState
    Call modApiSheets.CreateOutputSheet(SheetName)
    Set ws = ActiveWorkbook.Worksheets(SheetName)
    ws.Range("A1").Value = "Bland-Altman plot"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 12
    ws.Range("A3").Value = "Measurement 1"
    ws.Range("B3").Value = "Measurement 2"
    ws.Range("C3").Value = "Average"
    ws.Range("D3").Value = "Difference"
    ws.Range("A3:D3").Font.Bold = True
    ws.Range("A3:D3").HorizontalAlignment = xlCenter
    ws.Range("A3:D3").WrapText = True
    If n = 1 Then
        ws.Range("A4").Resize(1, 2).Value = raw
    Else
        ws.Range("A4").Resize(n, 2).Value = raw
    End If
    ws.Range("C4").Resize(n, 1).Value = avg
    ws.Range("D4").Resize(n, 1).Value = diff
    Set rngAvg = ws.Range("C4").Resize(n, 1)
    Set rngDiff = ws.Range("D4").Resize(n, 1)
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "AverageMeasurement", rngAvg)
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "MeasurementDifference", rngDiff)
    ws.Range("F3").Value = "n"
    ws.Range("G3").Value = n
    ws.Range("F4").Value = "Mean difference"
    ws.Range("G4").Value = meanDiff
    ws.Range("F5").Value = "SD (sample)"
    ws.Range("G5").Value = sdDiff
    ws.Range("F6").Value = "+1.96 SD"
    ws.Range("G6").Value = meanDiff + 1.96 * sdDiff
    ws.Range("F7").Value = "-1.96 SD"
    ws.Range("G7").Value = meanDiff - 1.96 * sdDiff
    ws.Range("F3:F7").HorizontalAlignment = xlRight
    ws.Range("G3:G7").NumberFormat = "#,##0.0000"
    ws.Range("A4:D" & (3 + n)).NumberFormat = "#,##0.0000"
    ws.Columns("A:D").ColumnWidth = 14
    ws.Columns("E").ColumnWidth = 2
    ws.Columns("F").ColumnWidth = 18
    ws.Columns("G").ColumnWidth = 12
    Set ch = ws.ChartObjects.Add(Left:=ws.Range("I2").Left, Top:=ws.Range("I2").Top, Width:=480, Height:=275)
    With ch.Chart
        .ChartType = xlXYScatter
        .HasTitle = True
        .ChartTitle.Text = "Bland-Altman plot - difference vs mean"
        .HasLegend = True
        .SeriesCollection.NewSeries
        With .SeriesCollection(1)
            .Name = "Difference vs mean"
            .XValues = rngAvg
            .Values = rngDiff
            .MarkerStyle = xlMarkerStyleCircle
            .MarkerSize = 5
            .MarkerForegroundColor = RGB(0, 0, 255)
            .MarkerBackgroundColor = RGB(0, 0, 255)
        End With
        Call AddHorizontalLine(ch.Chart, 2, "Mean difference", minX, maxX, meanDiff, RGB(255, 0, 0), False)
        Call AddHorizontalLine(ch.Chart, 3, "+1.96 SD", minX, maxX, meanDiff + 1.96 * sdDiff, RGB(0, 0, 255), True)
        Call AddHorizontalLine(ch.Chart, 4, "-1.96 SD", minX, maxX, meanDiff - 1.96 * sdDiff, RGB(0, 0, 255), True)
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Mean of two measurements"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Difference (method 1 - method 2)"
        .Axes(xlCategory).HasMajorGridlines = False
        .Axes(xlValue).HasMajorGridlines = False
    End With
    ch.Name = "BlandAltmanPlot"
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("BlandAltmanPlot")
End Sub

Private Sub AddHorizontalLine(ByVal ch As Chart, ByVal seriesIndex As Long, ByVal seriesName As String, _
                              ByVal x0 As Double, ByVal x1 As Double, ByVal y As Double, _
                              ByVal lineRgb As Long, ByVal dashed As Boolean)
    ch.SeriesCollection.NewSeries
    With ch.SeriesCollection(seriesIndex)
        .Name = seriesName
        .XValues = Array(x0, x1)
        .Values = Array(y, y)
        .ChartType = xlXYScatterLinesNoMarkers
        .Format.Line.Weight = 0.75
        .Format.Line.ForeColor.RGB = lineRgb
        If dashed Then
            .Format.Line.DashStyle = msoLineDash
        Else
            .Format.Line.DashStyle = msoLineSolid
        End If
        .HasDataLabels = True
        .DataLabels.ShowValue = False
        .DataLabels.ShowSeriesName = False
        .Points(.Points.Count).HasDataLabel = True
        .Points(.Points.Count).DataLabel.Text = seriesName
        If dashed Then
            .Points(.Points.Count).DataLabel.Position = xlLabelPositionAbove
        Else
            .Points(.Points.Count).DataLabel.Position = xlLabelPositionRight
        End If
    End With
End Sub
