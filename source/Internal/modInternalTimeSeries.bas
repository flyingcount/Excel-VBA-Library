Attribute VB_Name = "modInternalTimeSeries"
Option Explicit

' Internal: time-series analysis and lag differencing (Personal Custom_Menu27_*).
' Called from modApiTimeSeries. Do not document these as the external API.

Public Const TimeSeriesSheetName As String = "Time Series"
Public Const DateDiffSheetName As String = "Date Diff"
Public Const DateDiffFormulaSheetName As String = "Date Diff Formula"
Public Const TimeSeriesDataName As String = "Time_Series_Data"

Private Const TitleMenu As String = "Time series"
Private Const DateDiffStartRow As Long = 21
Private Const MaxDiffOrders As Long = 24
Private Const AlphaDefault As Double = 0.05

Public Function PromptNumericColumn(ByVal PromptText As String, ByVal minRows As Long) As Range
    Dim rng As Range
    Set rng = modInternalRanges.PromptRange(PromptText, TitleMenu)
    If rng Is Nothing Then Exit Function
    Set rng = modInternalRanges.UsedPart(rng)
    If rng Is Nothing Then
        MsgBox "The selected range has no used cells.", vbExclamation, TitleMenu
        Exit Function
    End If
    If rng.Columns.Count <> 1 Then
        MsgBox "Data must be a single column.", vbExclamation, TitleMenu
        Exit Function
    End If
    If rng.Rows.Count < minRows Then
        MsgBox "Need at least " & CStr(minRows) & " observations.", vbExclamation, TitleMenu
        Exit Function
    End If
    If Application.WorksheetFunction.Count(rng) <> rng.Cells.Count Then
        MsgBox "Every cell must be numeric and non-empty.", vbExclamation, TitleMenu
        Exit Function
    End If
    Set PromptNumericColumn = rng
End Function

Public Sub WriteTimeSeriesAnalysis(ByVal src As Range, ByVal useFormulas As Boolean)
    Dim ws As Worksheet
    Dim n As Long
    Dim nLag As Long
    Dim dataRng As Range
    Dim vals As Variant
    Dim stats As Variant
    Dim i As Long

    n = src.Rows.Count
    If useFormulas Then
        nLag = n \ 4
    Else
        nLag = n \ 3
    End If
    If nLag < 1 Then nLag = 1

    Set ws = modInternalPlots.OutputSheet(TimeSeriesSheetName)
    Call FormatTimeSeriesSheet(ws)

    vals = ColumnDoubles(src)
    For i = 1 To n
        ws.Cells(5 + i, 1).Value = vals(i)
    Next i
    Set dataRng = ws.Range("A6").Resize(n, 1)
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, TimeSeriesDataName, dataRng, True, "Created " & Now())

    If useFormulas Then
        Call WriteFormulaStats(ws, nLag)
    Else
        stats = ValueStatsBlock(vals, nLag)
        ws.Range("C6").Resize(UBound(stats, 1), UBound(stats, 2)).Value = stats
    End If
    ws.Activate
End Sub

Public Sub WriteDateDifferencing(ByVal src As Range, ByVal useFormulas As Boolean)
    Dim ws As Worksheet
    Dim n As Long
    Dim nDiff As Long
    Dim nCols As Long
    Dim arr As Variant
    Dim vals As Variant
    Dim r As Long
    Dim c As Long
    Dim d As Long
    Dim sheetName As String
    Dim headerRow As Long
    Dim lastDataRow As Long
    Dim col As Long

    If useFormulas Then
        sheetName = DateDiffFormulaSheetName
    Else
        sheetName = DateDiffSheetName
    End If

    n = src.Rows.Count
    nDiff = CLng(Round(n / 3, 0))
    If nDiff < 1 Then nDiff = 1
    If nDiff > MaxDiffOrders Then nDiff = MaxDiffOrders
    nCols = nDiff + 2
    vals = ColumnDoubles(src)

    ReDim arr(1 To n + 1, 1 To nCols)
    arr(1, 1) = "Index"
    arr(1, 2) = "Time series data"
    For c = 3 To nCols
        arr(1, c) = c - 2
    Next c
    For r = 1 To n
        arr(r + 1, 1) = r
        arr(r + 1, 2) = vals(r)
    Next r
    For c = 3 To nCols
        d = c - 2
        For r = 1 To n - d
            If useFormulas Then
                arr(r + 1, c) = "=R[1]C[-1]-RC[-1]"
            Else
                arr(r + 1, c) = CDbl(arr(r + 2, c - 1)) - CDbl(arr(r + 1, c - 1))
            End If
        Next r
    Next c

    Set ws = modInternalPlots.OutputSheet(sheetName)
    ws.Hyperlinks.Add Anchor:=ws.Range("D2"), Address:="https://real-statistics.com/time-series-analysis/arima-processes/arima-differencing/", TextToDisplay:="ARIMA differencing"
    headerRow = DateDiffStartRow + 4
    lastDataRow = headerRow + n
    ws.Cells(DateDiffStartRow, 1).Value = "Slope"
    ws.Cells(DateDiffStartRow + 1, 1).Value = "Intercept"
    ws.Cells(DateDiffStartRow + 2, 1).Value = "R squared"
    For col = 2 To nCols
        ws.Cells(DateDiffStartRow, col).FormulaR1C1 = "=SLOPE(R" & (headerRow + 1) & "C" & col & ":R" & lastDataRow & "C" & col & ",R" & (headerRow + 1) & "C1:R" & lastDataRow & "C1)"
        ws.Cells(DateDiffStartRow + 1, col).FormulaR1C1 = "=INTERCEPT(R" & (headerRow + 1) & "C" & col & ":R" & lastDataRow & "C" & col & ",R" & (headerRow + 1) & "C1:R" & lastDataRow & "C1)"
        ws.Cells(DateDiffStartRow + 2, col).FormulaR1C1 = "=RSQ(R" & (headerRow + 1) & "C" & col & ":R" & lastDataRow & "C" & col & ",R" & (headerRow + 1) & "C1:R" & lastDataRow & "C1)"
    Next col

    If useFormulas Then
        Call modInternalPlots.WriteBlock(ws, headerRow, 1, arr)
        Call WriteDiffAcfTable(ws, headerRow, n, nCols)
    Else
        ws.Cells(headerRow, 1).Resize(n + 1, nCols).Value = arr
    End If
    ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, nCols)).Font.Bold = True
    Call PlotDateDiffCharts(ws, ws.Cells(headerRow, 1).Resize(n + 1, nCols), nCols)
    ws.Activate
End Sub

Public Function TsAcf(ByVal rngData As Range, ByVal lag As Long) As Variant
    Dim n As Long
    Dim vals As Variant
    Dim mean As Double
    Dim s0 As Double
    Dim sLag As Double
    If rngData Is Nothing Then
        TsAcf = "Error! Range must be a single column"
        Exit Function
    End If
    If rngData.Columns.Count <> 1 Then
        TsAcf = "Error! Range must be a single column"
        Exit Function
    End If
    n = rngData.Rows.Count
    If lag < 0 Or lag > n - 1 Then
        TsAcf = "Error! Lag cannot be greater than the number of data points"
        Exit Function
    End If
    On Error GoTo Fail
    vals = ColumnDoubles(rngData)
    mean = MeanOf(vals, n)
    s0 = AcvfK(vals, n, mean, 0)
    If s0 = 0 Then
        TsAcf = "Error: divide by zero"
        Exit Function
    End If
    sLag = AcvfK(vals, n, mean, lag)
    TsAcf = sLag / s0
    Exit Function
Fail:
    TsAcf = CVErr(xlErrValue)
End Function

Public Function TsAcvf(ByVal rngData As Range, ByVal lag As Long) As Variant
    Dim n As Long
    Dim vals As Variant
    Dim mean As Double
    If rngData Is Nothing Then
        TsAcvf = CVErr(xlErrRef)
        Exit Function
    End If
    If lag < 0 Then
        TsAcvf = CVErr(xlErrNum)
        Exit Function
    End If
    On Error GoTo Fail
    n = rngData.Rows.Count
    If lag > n - 1 Then
        TsAcvf = CVErr(xlErrNum)
        Exit Function
    End If
    vals = ColumnDoubles(rngData)
    mean = MeanOf(vals, n)
    TsAcvf = AcvfK(vals, n, mean, lag)
    Exit Function
Fail:
    TsAcvf = CVErr(xlErrValue)
End Function

Public Function TsPacf(ByVal rngData As Range, ByVal lag As Long) As Variant
    Dim n As Long
    Dim vals As Variant
    Dim mean As Double
    Dim acvf As Variant
    Dim k As Long
    If lag < 1 Then
        TsPacf = CVErr(xlErrNum)
        Exit Function
    End If
    On Error GoTo Fail
    n = rngData.Rows.Count
    vals = ColumnDoubles(rngData)
    mean = MeanOf(vals, n)
    ReDim acvf(0 To lag)
    For k = 0 To lag
        acvf(k) = AcvfK(vals, n, mean, k)
    Next k
    TsPacf = PacfFromAcvf(acvf, lag)
    Exit Function
Fail:
    TsPacf = CVErr(xlErrValue)
End Function

Public Function TsBartlett(ByVal rngData As Range, ByVal lag As Long, Optional ByVal alpha As Double = AlphaDefault) As Variant
    Dim acfVal As Variant
    Dim n As Long
    Dim se As Double
    Dim crit As Double
    On Error GoTo Fail
    acfVal = TsAcf(rngData, lag)
    If Not IsNumeric(acfVal) Then
        TsBartlett = acfVal
        Exit Function
    End If
    n = Application.WorksheetFunction.Count(rngData)
    se = 1 / Sqr(n)
    crit = Application.WorksheetFunction.Norm_Inv(1 - alpha / 2, 0, se)
    TsBartlett = BartlettText(CDbl(acfVal), crit)
    Exit Function
Fail:
    TsBartlett = CVErr(xlErrValue)
End Function

Public Function TsBoxPierce(ByVal rngData As Range, ByVal lag As Long) As Variant
    TsBoxPierce = Portmanteau(rngData, lag, False)
End Function

Public Function TsLjungBox(ByVal rngData As Range, ByVal lag As Long) As Variant
    TsLjungBox = Portmanteau(rngData, lag, True)
End Function

Public Function TsChiSqPVal(ByVal stat As Variant, ByVal lag As Long) As Variant
    If Not IsNumeric(stat) Then
        TsChiSqPVal = stat
        Exit Function
    End If
    If lag < 1 Then
        TsChiSqPVal = CVErr(xlErrNum)
        Exit Function
    End If
    On Error GoTo Fail
    TsChiSqPVal = Application.WorksheetFunction.ChiSq_Dist_RT(CDbl(stat), lag)
    Exit Function
Fail:
    TsChiSqPVal = CVErr(xlErrValue)
End Function

Public Function TsWhiteNoiseTest(ByVal pVal As Variant, ByVal lag As Long, ByVal alpha As Double, ByVal testName As String) As Variant
    If Not IsNumeric(pVal) Then
        TsWhiteNoiseTest = pVal
        Exit Function
    End If
    TsWhiteNoiseTest = WhiteNoiseText(testName, CDbl(pVal), lag, alpha)
End Function

Private Sub FormatTimeSeriesSheet(ByVal ws As Worksheet)
    With ws.Range("A1")
        .Value = "Time series analysis"
        .Font.Bold = True
    End With
    ws.Hyperlinks.Add Anchor:=ws.Range("G1"), Address:="https://people.duke.edu/~rnau/411arim.htm#spreadsheet", TextToDisplay:="https://people.duke.edu/~rnau/411arim.htm#spreadsheet"
    ws.Hyperlinks.Add Anchor:=ws.Range("G2"), Address:="https://people.duke.edu/~rnau/411arim2.htm", TextToDisplay:="https://people.duke.edu/~rnau/411arim2.htm"
    ws.Hyperlinks.Add Anchor:=ws.Range("G3"), Address:="https://people.duke.edu/~rnau/411arim3.htm", TextToDisplay:="https://people.duke.edu/~rnau/411arim3.htm"
    ws.Hyperlinks.Add Anchor:=ws.Range("G4"), Address:="http://www.real-statistics.com/time-series-analysis/basic-time-series-forecasting/", TextToDisplay:="Basic time-series forecasting"
    ws.Hyperlinks.Add Anchor:=ws.Range("J1"), Address:="http://www.real-statistics.com/time-series-analysis/stochastic-processes/partial-autocorrelation-function/", TextToDisplay:="Partial autocorrelation function"
    ws.Range("A5:M5").Value = Array("Data", "", "Lag", "ACF", "ACVF", "PACF", "Bartlett", _
        "Box-Pierce", "Box-Pierce p-value", "Box-Pierce test", "Ljung-Box", _
        "Ljung-Box p-value", "Ljung-Box test")
    With ws.Range("A5:M5")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With
    ws.Columns(2).ColumnWidth = 1
    ws.Columns(7).ColumnWidth = 65
    ws.Columns(10).AutoFit
    ws.Columns(13).AutoFit
End Sub

Private Function ValueStatsBlock(ByRef vals As Variant, ByVal nLag As Long) As Variant
    Dim n As Long
    Dim mean As Double
    Dim acvf As Variant
    Dim acf As Variant
    Dim arr As Variant
    Dim k As Long
    Dim lag As Long
    Dim bp As Double
    Dim lb As Double
    Dim pBp As Double
    Dim pLb As Double
    Dim se As Double
    Dim crit As Double
    n = UBound(vals)
    mean = MeanOf(vals, n)
    ReDim acvf(0 To nLag)
    ReDim acf(0 To nLag)
    For k = 0 To nLag
        acvf(k) = AcvfK(vals, n, mean, k)
        If acvf(0) = 0 Then
            acf(k) = Empty
        Else
            acf(k) = acvf(k) / acvf(0)
        End If
    Next k
    se = 1 / Sqr(n)
    crit = Application.WorksheetFunction.Norm_Inv(1 - AlphaDefault / 2, 0, se)

    ReDim arr(1 To nLag, 1 To 11)
    For k = 1 To nLag
        lag = k - 1
        arr(k, 1) = lag
        arr(k, 2) = acf(lag)
        arr(k, 3) = acvf(lag)
        If lag >= 1 Then
            arr(k, 4) = PacfFromAcvf(acvf, lag)
            arr(k, 5) = BartlettText(CDbl(acf(lag)), crit)
            bp = BoxPierceFromAcf(acf, n, lag)
            lb = LjungBoxFromAcf(acf, n, lag)
            pBp = Application.WorksheetFunction.ChiSq_Dist_RT(bp, lag)
            pLb = Application.WorksheetFunction.ChiSq_Dist_RT(lb, lag)
            arr(k, 6) = bp
            arr(k, 7) = pBp
            arr(k, 8) = WhiteNoiseText("Box-Pierce", pBp, lag, AlphaDefault)
            arr(k, 9) = lb
            arr(k, 10) = pLb
            arr(k, 11) = WhiteNoiseText("Ljung-Box", pLb, lag, AlphaDefault)
        End If
    Next k
    ValueStatsBlock = arr
End Function

Private Sub WriteDiffAcfTable(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal n As Long, ByVal nCols As Long)
    Dim firstData As Long
    Dim lastData As Long
    Dim acfTitle As Long
    Dim acfHeader As Long
    Dim flagRow As Long
    Dim lag1Row As Long
    Dim nLag As Long
    Dim shortest As Long
    Dim maxLag As Long
    Dim lag As Long
    Dim c As Long
    Dim d As Long
    Dim lastUsed As Long
    Dim p As String
    Dim srcAddr As String
    Dim lag1Rng As Range

    firstData = headerRow + 1
    lastData = headerRow + n
    shortest = n - (nCols - 2)
    If shortest < 2 Then shortest = 2
    nLag = n \ 4
    If nLag < 1 Then nLag = 1
    maxLag = shortest - 1
    If maxLag < 1 Then maxLag = 1
    If nLag > maxLag Then nLag = maxLag

    acfTitle = lastData + 3
    acfHeader = acfTitle + 2
    flagRow = acfHeader + nLag + 2
    lag1Row = acfHeader + 2
    p = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!"

    ws.Cells(acfTitle, 1).Value = "Autocorrelation of the original series and each differenced series"
    ws.Cells(acfTitle, 1).Font.Bold = True
    ws.Cells(acfTitle + 1, 1).Value = "Lag-1 ACF < -0.5 suggests over-differencing. Prefer the smallest d whose lag-1 ACF is not below -0.5."
    ws.Cells(acfHeader, 1).Value = "Lag"
    For c = 2 To nCols
        ws.Cells(acfHeader, c).Value = ws.Cells(headerRow, c).Value
    Next c
    ws.Range(ws.Cells(acfHeader, 1), ws.Cells(acfHeader, nCols)).Font.Bold = True

    For lag = 0 To nLag
        ws.Cells(acfHeader + 1 + lag, 1).Value = lag
        For c = 2 To nCols
            d = c - 2
            lastUsed = lastData - d
            If lastUsed < firstData Then lastUsed = firstData
            If lag > lastUsed - firstData Then GoTo NextCol
            srcAddr = ws.Range(ws.Cells(firstData, c), ws.Cells(lastUsed, c)).Address(True, True, xlA1, False)
            ws.Cells(acfHeader + 1 + lag, c).Formula = "=" & p & "ACF(" & srcAddr & "," & CStr(lag) & ")"
NextCol:
        Next c
    Next lag

    ws.Cells(flagRow, 1).Value = "Lag-1 ACF < -0.5"
    ws.Cells(flagRow, 1).Font.Bold = True
    For c = 2 To nCols
        ws.Cells(flagRow, c).Formula = "=IF(" & ws.Cells(lag1Row, c).Address(True, True, xlA1, False) & "<-0.5,""Over-differenced"",""OK"")"
    Next c
    ws.Range(ws.Cells(acfHeader + 1, 2), ws.Cells(acfHeader + 1 + nLag, nCols)).NumberFormat = "0.000"

    Set lag1Rng = ws.Range(ws.Cells(lag1Row, 2), ws.Cells(lag1Row, nCols))
    lag1Rng.FormatConditions.Delete
    With lag1Rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="-0.5")
        .Interior.Color = RGB(255, 199, 206)
        .Font.Color = RGB(156, 0, 6)
    End With
    With lag1Rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="-0.5")
        .Interior.Color = RGB(198, 239, 206)
        .Font.Color = RGB(0, 97, 0)
    End With
    ws.Columns(1).AutoFit
End Sub

Private Sub WriteFormulaStats(ByVal ws As Worksheet, ByVal nLag As Long)
    Dim k As Long
    Dim r As Long
    Dim lag As Long
    Dim p As String
    p = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!"
    For k = 1 To nLag
        lag = k - 1
        r = 5 + k
        ws.Cells(r, 3).Value = lag
        ws.Cells(r, 4).Formula = "=" & p & "ACF(" & TimeSeriesDataName & ",C" & CStr(r) & ")"
        ws.Cells(r, 5).Formula = "=" & p & "ACVF(" & TimeSeriesDataName & ",C" & CStr(r) & ")"
        If lag >= 1 Then
            ws.Cells(r, 6).Formula = "=" & p & "PACF(" & TimeSeriesDataName & ",C" & CStr(r) & ")"
            ws.Cells(r, 7).Formula = "=" & p & "Bartlett(" & TimeSeriesDataName & ",C" & CStr(r) & ")"
            ws.Cells(r, 8).Formula = "=" & p & "BoxPierce(" & TimeSeriesDataName & ",C" & CStr(r) & ")"
            ws.Cells(r, 9).Formula = "=" & p & "BoxPiercePVal(" & TimeSeriesDataName & ",C" & CStr(r) & ")"
            ws.Cells(r, 10).Formula = "=" & p & "BoxPierceTest(" & TimeSeriesDataName & ",C" & CStr(r) & ")"
            ws.Cells(r, 11).Formula = "=" & p & "LjungBox(" & TimeSeriesDataName & ",C" & CStr(r) & ")"
            ws.Cells(r, 12).Formula = "=" & p & "LjungBoxPVal(" & TimeSeriesDataName & ",C" & CStr(r) & ")"
            ws.Cells(r, 13).Formula = "=" & p & "LjungBoxTest(" & TimeSeriesDataName & ",C" & CStr(r) & ")"
        End If
    Next k
End Sub

Private Function ColumnDoubles(ByVal rng As Range) As Variant
    Dim n As Long
    Dim i As Long
    Dim raw As Variant
    Dim arr As Variant
    n = rng.Rows.Count
    ReDim arr(1 To n)
    If n = 1 Then
        arr(1) = CDbl(rng.Cells(1, 1).Value)
    Else
        raw = rng.Value
        For i = 1 To n
            arr(i) = CDbl(raw(i, 1))
        Next i
    End If
    ColumnDoubles = arr
End Function

Private Function MeanOf(ByRef vals As Variant, ByVal n As Long) As Double
    Dim i As Long
    Dim s As Double
    For i = 1 To n
        s = s + vals(i)
    Next i
    MeanOf = s / n
End Function

Private Function AcvfK(ByRef vals As Variant, ByVal n As Long, ByVal mean As Double, ByVal lag As Long) As Double
    Dim i As Long
    Dim tot As Double
    For i = 1 To n - lag
        tot = tot + (vals(i + lag) - mean) * (vals(i) - mean)
    Next i
    AcvfK = tot / n
End Function

Private Function PacfFromAcvf(ByRef acvf As Variant, ByVal lag As Long) As Variant
    Dim g As Variant
    Dim rhs As Variant
    Dim work As Variant
    Dim r As Long
    Dim c As Long
    ReDim g(1 To lag, 1 To lag)
    ReDim rhs(1 To lag, 1 To 1)
    For c = 1 To lag
        rhs(c, 1) = acvf(c)
        For r = 1 To lag
            g(r, c) = acvf(Abs(c - r))
        Next r
    Next c
    On Error GoTo Fail
    work = Application.WorksheetFunction.MMult(Application.WorksheetFunction.MInverse(g), rhs)
    On Error Resume Next
    PacfFromAcvf = work(lag, 1)
    If Err.Number <> 0 Then
        Err.Clear
        PacfFromAcvf = work(lag)
    End If
    On Error GoTo 0
    Exit Function
Fail:
    PacfFromAcvf = CVErr(xlErrValue)
End Function

Private Function BoxPierceFromAcf(ByRef acf As Variant, ByVal n As Long, ByVal lag As Long) As Double
    Dim i As Long
    Dim tot As Double
    For i = 1 To lag
        tot = tot + acf(i) * acf(i)
    Next i
    BoxPierceFromAcf = n * tot
End Function

Private Function LjungBoxFromAcf(ByRef acf As Variant, ByVal n As Long, ByVal lag As Long) As Double
    Dim i As Long
    Dim tot As Double
    For i = 1 To lag
        tot = tot + (acf(i) * acf(i)) / (n - i)
    Next i
    LjungBoxFromAcf = n * (n + 2) * tot
End Function

Private Function Portmanteau(ByVal rngData As Range, ByVal lag As Long, ByVal ljung As Boolean) As Variant
    Dim n As Long
    Dim k As Long
    Dim vals As Variant
    Dim mean As Double
    Dim s0 As Double
    Dim acf As Variant
    If lag < 1 Then
        Portmanteau = CVErr(xlErrNum)
        Exit Function
    End If
    On Error GoTo Fail
    n = rngData.Rows.Count
    vals = ColumnDoubles(rngData)
    mean = MeanOf(vals, n)
    s0 = AcvfK(vals, n, mean, 0)
    If s0 = 0 Then
        Portmanteau = CVErr(xlErrDiv0)
        Exit Function
    End If
    ReDim acf(1 To lag)
    For k = 1 To lag
        acf(k) = AcvfK(vals, n, mean, k) / s0
    Next k
    If ljung Then
        Portmanteau = LjungBoxFromAcf(acf, n, lag)
    Else
        Portmanteau = BoxPierceFromAcf(acf, n, lag)
    End If
    Exit Function
Fail:
    Portmanteau = CVErr(xlErrValue)
End Function

Private Function BartlettText(ByVal acfVal As Double, ByVal crit As Double) As String
    If Abs(acfVal) > crit Then
        BartlettText = "Bartlett test is statistically different from zero: |ACF| " & Round(acfVal, 2) & " > Crit " & Round(crit, 2)
    Else
        BartlettText = "Bartlett test is Not statistically different from zero: |ACF| " & Round(acfVal, 2) & " < Crit " & Round(crit, 2)
    End If
End Function

Private Function WhiteNoiseText(ByVal testName As String, ByVal pVal As Double, ByVal lag As Long, ByVal alpha As Double) As String
    If pVal < alpha Then
        WhiteNoiseText = testName & " rejects white noise for lags =< " & CStr(lag) & ": p-value " & Round(pVal, 6) & " < Alpha " & Round(alpha, 2)
    Else
        WhiteNoiseText = testName & " does not reject white noise for lags =< " & CStr(lag) & ": p-value " & Round(pVal, 6) & " >= Alpha " & Round(alpha, 2)
    End If
End Function

Private Sub PlotDateDiffCharts(ByVal ws As Worksheet, ByVal tbl As Range, ByVal nCols As Long)
    Call AddLineTrendChart(ws, ws.Range("A4"), tbl.Columns(2), "Time series", "Original time series")
    If nCols >= 3 Then
        Call AddLineTrendChart(ws, ws.Range("B4"), tbl.Columns(3), "1-Differenced time series", "Differenced time series")
    End If
    If nCols >= 4 Then
        Call AddLineTrendChart(ws, ws.Range("L4"), tbl.Columns(4), "2-Differenced time series", "Differenced time series")
    End If
End Sub

Private Sub AddLineTrendChart(ByVal ws As Worksheet, ByVal anchor As Range, ByVal yCol As Range, ByVal titleText As String, ByVal seriesName As String)
    Dim ch As ChartObject
    Set ch = modInternalPlots.AddChartAt(ws, anchor, 400, 235)
    With ch.Chart
        .ChartType = xlLine
        .SetSourceData Source:=yCol
        .HasTitle = True
        With .ChartTitle
            .Text = titleText
            .Font.Size = 12
        End With
        .HasLegend = False
        With .Axes(xlCategory, xlPrimary)
            .HasTitle = True
            .AxisTitle.Text = "Time"
            With .TickLabels
                .Font.Size = 5
                .Orientation = xlTickLabelOrientationHorizontal
            End With
        End With
        With .Axes(xlValue, xlPrimary)
            .HasTitle = True
            .AxisTitle.Text = "Values"
            .HasMajorGridlines = False
            .HasMinorGridlines = False
            .TickLabels.Font.Size = 5
        End With
        With .SeriesCollection(1)
            .Name = seriesName
            .Trendlines.Add
            With .Trendlines(1)
                .DisplayEquation = True
                .DisplayRSquared = True
                .Format.Line.ForeColor.RGB = RGB(255, 0, 0)
                .Format.Line.Weight = 0.1
            End With
            .Format.Line.Weight = 0.75
            .Format.Line.ForeColor.RGB = RGB(0, 0, 0)
            .Format.Line.DashStyle = msoLineSolid
        End With
    End With
End Sub
