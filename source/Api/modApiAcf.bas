Attribute VB_Name = "modApiAcf"
Option Explicit

' Public API: Personal Custom_Menu11_ACF.
' ACF is a worksheet formula; PACF is computed in VBA (Yule-Walker) up to min(n-1, 40) lags.

Private Const SheetName As String = "Correlogram"
Private Const MaxLag As Long = 40

''' @Description: Autocorrelation and partial autocorrelation (correlogram) of a single numeric column.
''' @Example: GenerateCorrelogram
Public Sub GenerateCorrelogram()
    Dim src As Range
    Dim ws As Worksheet
    Dim n As Long
    Dim nLag As Long
    Dim k As Long
    Dim dataRng As Range
    Dim lagRng As Range
    Dim ch As ChartObject
    On Error GoTo EH
    Set src = modInternalPlots.PromptRange("Select a single column of time-series values.")
    If src Is Nothing Then Exit Sub
    If src.Columns.Count <> 1 Then
        MsgBox "Data must be in a single column.", vbExclamation, "Plots Charts"
        Exit Sub
    End If
    n = src.Rows.Count
    If n < 4 Then
        MsgBox "Need at least 4 observations.", vbExclamation, "Plots Charts"
        Exit Sub
    End If
    If Not modInternalPlots.RangeIsAllNumeric(src) Then
        MsgBox "Every cell must be numeric.", vbExclamation, "Plots Charts"
        Exit Sub
    End If
    nLag = n - 1
    If nLag > MaxLag Then nLag = MaxLag
    Call modInternalExcelApp.PushAppState
    Set ws = modInternalPlots.OutputSheet(SheetName)
    ws.Range("A1").Value = "Correlogram"
    ws.Range("A1").Font.Bold = True
    src.Copy
    ws.Range("AF3").Resize(n, 1).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    Set dataRng = ws.Range("AF3").Resize(n, 1)
    For k = 1 To n
        dataRng.Cells(k, 1).Offset(0, 1).Value = k
    Next k
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "Data", dataRng)
    ws.Range("AC2").Value = "Statistics"
    ws.Range("AC3").Value = "Mean"
    ws.Range("AD3").Formula = "=AVERAGE(Data)"
    ws.Range("AC4").Value = "Variance"
    ws.Range("AD4").Formula = "=VAR.P(Data)"
    ws.Range("AC5").Value = "Dev Sq"
    ws.Range("AD5").Formula = "=DEVSQ(Data)"
    ws.Range("AC6").Value = "Count"
    ws.Range("AD6").Formula = "=COUNT(Data)"
    ws.Range("AC7").Value = "Alpha"
    ws.Range("AD7").Value = 0.05
    Call modInternalPlots.HighlightYellow(ws.Range("AD7"))
    ws.Range("AC8").Value = "Standard error"
    ws.Range("AD8").Formula = "=1/SQRT(COUNT(Data))"
    ws.Range("AC9").Value = "PACF Critical value"
    ws.Range("AD9").FormulaR1C1 = "=NORM.INV(1-R[-2]C/2,0,R[-1]C)"
    ws.Range("AC10").Value = "Non-decreasing from first obs"
    ws.Range("AD10").Value = SeriesNonDecreasing(src)
    ws.Range("AC11").Value = "Count ACF > Critical value"
    ws.Range("AD11").Formula = "=COUNTIF(ACFGreaterCritVal,TRUE)"
    ws.Range("AC12").Value = "Count ACF > Critical value %"
    ws.Range("AD12").FormulaR1C1 = "=R[-1]C/R[-6]C"
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "Alpha", ws.Range("AD7"))
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "Standard_error", ws.Range("AD8"))
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "PACF_Critical_value", ws.Range("AD9"))
    Call modInternalPlots.WriteHeaderRow(ws.Range("Q2"), Array("Lag", "ACF", "p-value", "ACF-Critical value", "Comment", "Lower", "Upper", "ACF > critical", "PACF", "PACF lower", "PACF upper"))
    For k = 1 To nLag
        Application.StatusBar = "Lag " & k & " of " & nLag
        ws.Cells(2 + k, 17).Value = k
        ws.Cells(2 + k, 18).FormulaR1C1 = "=SUMPRODUCT(OFFSET(Data,0,0,COUNT(Data)-RC[-1])-AVERAGE(Data),OFFSET(Data,RC[-1],0,COUNT(Data)-RC[-1])-AVERAGE(Data))/DEVSQ(Data)"
        ws.Cells(2 + k, 19).FormulaR1C1 = "=1-NORM.DIST(RC[-1],0,Standard_error,TRUE)"
        ws.Cells(2 + k, 20).FormulaR1C1 = "=ABS(RC[-2])-RC[3]"
        ws.Cells(2 + k, 21).FormulaR1C1 = "=IF(RC[-1]<0,""Is NOT significantly different from zero"",""Is significantly different from zero"")"
        ws.Cells(2 + k, 22).FormulaR1C1 = "=-RC[1]"
        If k = 1 Then
            ws.Cells(2 + k, 23).FormulaR1C1 = "=NORMSINV(1-Alpha/2)/SQRT(COUNT(Data))"
        Else
            ws.Cells(2 + k, 23).FormulaR1C1 = "=NORMSINV(1-Alpha/2)*SQRT((1+2*SUMSQ(R3C18:R[-1]C18))/COUNT(Data))"
        End If
        ws.Cells(2 + k, 24).FormulaR1C1 = "=IF(RC[-4]>0,TRUE,FALSE)"
        ws.Cells(2 + k, 25).Value = PacfAtLag(dataRng, k)
        ws.Cells(2 + k, 26).Formula = "=-PACF_Critical_value"
        ws.Cells(2 + k, 27).Formula = "=PACF_Critical_value"
    Next k
    Set lagRng = ws.Range("Q3").Resize(nLag, 1)
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "Lag", lagRng)
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "ACF", lagRng.Offset(0, 1))
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "ACFUpperCritVal", lagRng.Offset(0, 5))
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "ACFLowerCritVal", lagRng.Offset(0, 6))
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "ACFGreaterCritVal", lagRng.Offset(0, 7))
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "PACF", lagRng.Offset(0, 8))
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "PACFLowerCritVal", lagRng.Offset(0, 9))
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "PACFUpperCritVal", lagRng.Offset(0, 10))
    ws.Columns("AB").ColumnWidth = 1
    ws.Columns("AE").ColumnWidth = 1
    Set ch = modInternalPlots.AddChartAt(ws, ws.Range("A2"), 375, 300)
    ch.Chart.HasTitle = True
    ch.Chart.ChartTitle.Text = "Time-series"
    Call modInternalPlots.AddXySeries(ch.Chart, "Data", dataRng.Offset(0, 1).Resize(n, 1), dataRng, xlXYScatterLines, RGB(0, 0, 0), False)
    Call modInternalPlots.StyleValueAxes(ch.Chart, "", "", "0", "0.00", False)
    Set ch = modInternalPlots.AddChartAt(ws, ws.Range("I1"), 375, 200)
    Call PlotCorrelogram(ch, "Correlogram - ACF", ws.Range("ACF"), ws.Range("ACFUpperCritVal"), ws.Range("ACFLowerCritVal"))
    Set ch = modInternalPlots.AddChartAt(ws, ws.Range("I12"), 375, 200)
    Call PlotCorrelogram(ch, "Correlogram - PACF", ws.Range("PACF"), ws.Range("PACFUpperCritVal"), ws.Range("PACFLowerCritVal"))
    ws.Calculate
    Call ColourTrue(ws.Range("ACFGreaterCritVal"))
    Application.StatusBar = False
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Application.StatusBar = False
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("GenerateCorrelogram")
End Sub

''' @Description: Lower ACF critical value -NormSInv(1-alpha/2)/sqrt(n).
Public Function ACFLower(ByVal rng_Data As Range, ByVal dbl_Aplha As Double) As Double
    Application.Volatile
    ACFLower = -Application.WorksheetFunction.NormSInv(1 - dbl_Aplha / 2) / Sqr(Application.WorksheetFunction.Count(rng_Data))
End Function

''' @Description: Upper ACF critical value NormSInv(1-alpha/2)/sqrt(n).
Public Function ACFUpper(ByVal rng_Data As Range, ByVal dbl_Aplha As Double) As Double
    Application.Volatile
    ACFUpper = Application.WorksheetFunction.NormSInv(1 - dbl_Aplha / 2) / Sqr(Application.WorksheetFunction.Count(rng_Data))
End Function

Private Sub PlotCorrelogram(ByVal chObj As ChartObject, ByVal titleText As String, ByVal acfRng As Range, ByVal upRng As Range, ByVal loRng As Range)
    With chObj.Chart
        .HasTitle = True
        .ChartTitle.Text = titleText
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .SeriesCollection.NewSeries
        With .SeriesCollection(1)
            .Name = titleText
            .Values = acfRng
            .ChartType = xlColumnClustered
        End With
        Call modInternalPlots.AddXySeries(chObj.Chart, "Upper critical value", acfRng, upRng, xlXYScatterLines, RGB(100, 100, 100), False)
        Call modInternalPlots.AddXySeries(chObj.Chart, "Lower critical value", acfRng, loRng, xlXYScatterLines, RGB(100, 100, 100), False)
        Call modInternalPlots.StyleValueAxes(chObj.Chart, "Lag", "", "0", "0.00", False)
        .HasLegend = True
    End With
End Sub

Private Function SeriesNonDecreasing(ByVal rng As Range) As Boolean
    Dim i As Long
    Dim first As Double
    first = CDbl(rng.Cells(1, 1).Value)
    For i = 1 To rng.Cells.Count
        If CDbl(rng.Cells(i, 1).Value) < first Then
            SeriesNonDecreasing = False
            Exit Function
        End If
    Next i
    SeriesNonDecreasing = True
End Function

Private Function PacfAtLag(ByVal rngData As Range, ByVal lag As Long) As Variant
    Dim work As Variant
    On Error GoTo Fail
    work = Application.WorksheetFunction.MMult(Application.WorksheetFunction.MInverse(AutoCovMatrix(rngData, lag)), AvcfVector(rngData, lag))
    If lag = 1 Then
        PacfAtLag = work(1)
    Else
        PacfAtLag = work(lag, 1)
    End If
    Exit Function
Fail:
    PacfAtLag = CVErr(xlErrValue)
End Function

Private Function AutoCovMatrix(ByVal rngData As Range, ByVal lag As Long) As Variant
    Dim a As Variant
    Dim r As Long
    Dim c As Long
    ReDim a(1 To lag, 1 To lag)
    For c = 1 To lag
        For r = 1 To lag
            a(r, c) = Acvf(rngData, Abs(c - r))
        Next r
    Next c
    AutoCovMatrix = a
End Function

Private Function AvcfVector(ByVal rngData As Range, ByVal lag As Long) As Variant
    Dim a As Variant
    Dim i As Long
    ReDim a(1 To lag, 1 To 1)
    For i = 1 To lag
        a(i, 1) = Acvf(rngData, i)
    Next i
    AvcfVector = a
End Function

Private Function Acvf(ByVal rngData As Range, ByVal lag As Long) As Double
    Dim mean As Double
    Dim tot As Double
    Dim i As Long
    Dim n As Long
    n = rngData.Cells.Count
    mean = Application.WorksheetFunction.Average(rngData)
    tot = 0
    For i = 1 To n - lag
        tot = tot + (CDbl(rngData.Cells(i + lag).Value) - mean) * (CDbl(rngData.Cells(i).Value) - mean)
    Next i
    Acvf = tot / n
End Function

Private Sub ColourTrue(ByVal rng As Range)
    Dim c As Range
    For Each c In rng.Cells
        If c.Value = True Then
            c.Interior.ColorIndex = 3
            c.Font.ColorIndex = 2
        Else
            c.Interior.ColorIndex = xlColorIndexNone
            c.Font.ColorIndex = xlColorIndexAutomatic
        End If
    Next c
End Sub
