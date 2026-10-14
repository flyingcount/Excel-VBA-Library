Attribute VB_Name = "modApiDistPlots"
Option Explicit

' Public API: Personal Custom_Menu11_* distribution / logistic plots.
' Writes a formula table at the selection (or on a dedicated sheet) with yellow input
' cells, then PDF and CDF charts. Formulas are R1C1 so they calculate.

''' @Description: Binomial PMF and CDF for 100 trials, p=0.8 (yellow cell). Writes at the selection.
''' @Example: GenerateBinomialPlot
Public Sub GenerateBinomialPlot()
    Call RunDistAtSelection("Binomial", Array("Prob.", "Trial", "Number of trials", "Binomial distribution", "CDF"), _
        100, 5, 1, 2, 4, 5, "Number of trials", "Frequency", True)
End Sub

''' @Description: Normal PDF and CDF (mean 0, sd 0.5) over ±5 sd. Writes at the selection.
''' @Example: GenerateNormalPlot
Public Sub GenerateNormalPlot()
    Call RunDistAtSelection("Normal", Array("Mean", "Standard deviation", "Z score", "x", "Normal distribution", "CDF"), _
        201, 6, 2, 4, 5, 6, "x", "Probability", True)
End Sub

''' @Description: Log-normal PDF and CDF on a new sheet "Log Normal". Yellow cells: increment, mean, sd, x-start.
''' @Example: GenerateLogNormalPlot
Public Sub GenerateLogNormalPlot()
    Call RunNamedSheetDist("Log Normal", 3, 1)
End Sub

''' @Description: Poisson PMF and CDF (mean 10). Writes at the selection.
''' @Example: GeneratePoissonPlot
Public Sub GeneratePoissonPlot()
    Call RunPoisson()
End Sub

''' @Description: Weibull PDF and CDF (alpha 1, beta 10). Writes at the selection.
''' @Example: GenerateWeibullPlot
Public Sub GenerateWeibullPlot()
    Call RunWeibullOrGamma("Weibull")
End Sub

''' @Description: Gamma PDF and CDF (alpha 1, beta 10). Writes at the selection.
''' @Example: GenerateGammaPlot
Public Sub GenerateGammaPlot()
    Call RunWeibullOrGamma("Gamma")
End Sub

''' @Description: Beta PDF and CDF on [0, 1] (alpha=beta=0.5). Writes at the selection.
''' @Example: GenerateBetaPlot
Public Sub GenerateBetaPlot()
    Call RunBeta()
End Sub

''' @Description: Exponential PDF and CDF (lambda 1.5) on a new sheet "Exponential".
''' @Example: GenerateExponentialPlot
Public Sub GenerateExponentialPlot()
    Call RunExponential()
End Sub

''' @Description: Hypergeometric PMF and CDF (n=100, K=50, N=500) on a new sheet "Hypergeometric".
''' @Example: GenerateHypergeometricPlot
Public Sub GenerateHypergeometricPlot()
    Call RunHypergeometric()
End Sub

''' @Description: Logistic sigmoid a/(1+EXP(-k*(x-x0))) on a new sheet "Logistics curve".
''' @Example: GenerateLogisticCurve
Public Sub GenerateLogisticCurve()
    Call RunLogistic()
End Sub

Private Sub RunDistAtSelection(ByVal kind As String, ByVal headers As Variant, _
                               ByVal nRows As Long, ByVal nCols As Long, _
                               ByVal yellowCols As Long, ByVal xCol As Long, _
                               ByVal pdfCol As Long, ByVal cdfCol As Long, _
                               ByVal xTitle As String, ByVal yTitle As String, _
                               ByVal markers As Boolean)
    Dim ws As Worksheet
    Dim startRow As Long
    Dim startCol As Long
    Dim data As Variant
    On Error GoTo EH
    If TypeName(Selection) <> "Range" Then
        MsgBox "Select a cell where the table should start.", vbExclamation, "Plots Charts"
        Exit Sub
    End If
    Call modInternalExcelApp.PushAppState
    Call modInternalPlots.AnchorAtSelection(startRow, startCol)
    Set ws = ActiveSheet
    data = DistArray(kind, nRows, nCols)
    Call modInternalPlots.WriteBlock(ws, startRow, startCol, data)
    Call FinishDist(ws, startRow, startCol, nRows, nCols, headers, yellowCols, xCol, pdfCol, cdfCol, _
                    kind & " distribution plot", kind & " CDF plot", xTitle, yTitle, markers)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("Generate" & kind & "Plot")
End Sub

Private Function DistArray(ByVal kind As String, ByVal nRows As Long, ByVal nCols As Long) As Variant
    Dim a As Variant
    Dim i As Long
    ReDim a(1 To nRows, 1 To nCols)
    Select Case kind
        Case "Binomial"
            For i = 1 To nRows
                If i = 1 Then a(i, 1) = 0.8 Else a(i, 1) = "=R[-1]C"
                a(i, 2) = i
                a(i, 3) = 100
                a(i, 4) = "=BINOM.DIST(RC[-2],RC[-1],RC[-3],FALSE)"
                a(i, 5) = "=BINOM.DIST(RC[-3],RC[-2],RC[-4],TRUE)"
            Next i
        Case "Normal"
            For i = 1 To nRows
                If i = 1 Then
                    a(i, 1) = 0
                    a(i, 2) = 0.5
                    a(i, 4) = "=RC[-3]-(5*RC[-2])"
                Else
                    a(i, 1) = "=R[-1]C"
                    a(i, 2) = "=R[-1]C"
                    a(i, 4) = "=R[-1]C+((10*RC[-2])/200)"
                End If
                a(i, 3) = "=(RC[1]-RC[-2])/RC[-1]"
                a(i, 5) = "=NORM.DIST(RC[-1],RC[-4],ABS(RC[-3]),FALSE)"
                a(i, 6) = "=NORM.DIST(RC[-2],RC[-5],ABS(RC[-4]),TRUE)"
            Next i
    End Select
    DistArray = a
End Function

Private Sub RunPoisson()
    Dim a As Variant
    Dim i As Long
    Dim ws As Worksheet
    Dim startRow As Long
    Dim startCol As Long
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    Call modInternalPlots.AnchorAtSelection(startRow, startCol)
    Set ws = ActiveSheet
    ReDim a(1 To 100, 1 To 5)
    For i = 1 To 100
        If i = 1 Then
            a(i, 1) = 1
            a(i, 2) = 10
            a(i, 3) = 0
        Else
            a(i, 1) = "=R[-1]C"
            a(i, 2) = "=R[-1]C"
            a(i, 3) = "=R[-1]C+RC[-2]"
        End If
        a(i, 4) = "=POISSON.DIST(RC[-1],RC[-2],FALSE)"
        a(i, 5) = "=POISSON.DIST(RC[-2],RC[-3],TRUE)"
    Next i
    Call modInternalPlots.WriteBlock(ws, startRow, startCol, a)
    Call FinishDist(ws, startRow, startCol, 100, 5, _
        Array("x increment", "Mean", "x", "Poisson distribution", "CDF"), _
        3, 3, 4, 5, "Poisson distribution plot", "Poisson CDF plot", "x", "Probability", True)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("GeneratePoissonPlot")
End Sub

Private Sub RunWeibullOrGamma(ByVal kind As String)
    Dim a As Variant
    Dim i As Long
    Dim ws As Worksheet
    Dim startRow As Long
    Dim startCol As Long
    Dim pdf As String
    Dim cdf As String
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    Call modInternalPlots.AnchorAtSelection(startRow, startCol)
    Set ws = ActiveSheet
    If kind = "Weibull" Then
        pdf = "=WEIBULL.DIST(RC[-1],RC[-3],RC[-2],FALSE)"
        cdf = "=WEIBULL.DIST(RC[-2],RC[-4],RC[-3],TRUE)"
    Else
        pdf = "=GAMMA.DIST(RC[-1],RC[-3],RC[-2],FALSE)"
        cdf = "=GAMMA.DIST(RC[-2],RC[-4],RC[-3],TRUE)"
    End If
    ReDim a(1 To 50, 1 To 6)
    For i = 1 To 50
        If i = 1 Then
            a(i, 1) = 1
            a(i, 2) = 1
            a(i, 3) = 10
            a(i, 4) = 0
        Else
            a(i, 1) = "=R[-1]C"
            a(i, 2) = "=R[-1]C"
            a(i, 3) = "=R[-1]C"
            a(i, 4) = "=R[-1]C+RC[-3]"
        End If
        a(i, 5) = pdf
        a(i, 6) = cdf
    Next i
    Call modInternalPlots.WriteBlock(ws, startRow, startCol, a)
    Call FinishDist(ws, startRow, startCol, 50, 6, _
        Array("x increment", "Alpha", "Beta", "x", kind & " distribution", "CDF"), _
        4, 4, 5, 6, kind & " distribution plot", kind & " CDF plot", "x", "Probability", False)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("Generate" & kind & "Plot")
End Sub

Private Sub RunBeta()
    Dim a As Variant
    Dim i As Long
    Dim ws As Worksheet
    Dim startRow As Long
    Dim startCol As Long
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    Call modInternalPlots.AnchorAtSelection(startRow, startCol)
    Set ws = ActiveSheet
    ReDim a(1 To 100, 1 To 6)
    For i = 1 To 100
        If i = 1 Then
            a(i, 1) = 0.01
            a(i, 2) = 0.5
            a(i, 3) = 0.5
            a(i, 4) = 0
        Else
            a(i, 1) = "=R[-1]C"
            a(i, 2) = "=R[-1]C"
            a(i, 3) = "=R[-1]C"
            a(i, 4) = "=R[-1]C+RC[-3]"
        End If
        a(i, 5) = "=BETA.DIST(RC[-1],RC[-3],RC[-2],FALSE,0,1)"
        a(i, 6) = "=BETA.DIST(RC[-2],RC[-4],RC[-3],TRUE,0,1)"
    Next i
    Call modInternalPlots.WriteBlock(ws, startRow, startCol, a)
    Call FinishDist(ws, startRow, startCol, 100, 6, _
        Array("x increment", "Alpha", "Beta", "x", "Beta distribution", "CDF"), _
        4, 4, 5, 6, "Beta distribution plot", "Beta CDF plot", "x", "Probability", False)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("GenerateBetaPlot")
End Sub

Private Sub RunExponential()
    Dim a As Variant
    Dim i As Long
    Dim ws As Worksheet
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    Set ws = modInternalPlots.OutputSheet("Exponential")
    ReDim a(1 To 50, 1 To 5)
    For i = 1 To 50
        If i = 1 Then
            a(i, 1) = 1
            a(i, 2) = 1.5
            a(i, 3) = 0
        Else
            a(i, 1) = "=R[-1]C"
            a(i, 2) = "=R[-1]C"
            a(i, 3) = "=R[-1]C+RC[-2]"
        End If
        a(i, 4) = "=EXPON.DIST(RC[-1],RC[-2],FALSE)"
        a(i, 5) = "=EXPON.DIST(RC[-2],RC[-3],TRUE)"
    Next i
    Call modInternalPlots.WriteBlock(ws, 2, 1, a)
    Call FinishDist(ws, 2, 1, 50, 5, _
        Array("x increment", "Lambda", "x", "Exponential distribution", "CDF"), _
        2, 3, 4, 5, "Exponential distribution plot", "Exponential CDF plot", "x", "Probability", False)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("GenerateExponentialPlot")
End Sub

Private Sub RunNamedSheetDist(ByVal sheetName As String, ByVal startRow As Long, ByVal startCol As Long)
    Dim a As Variant
    Dim i As Long
    Dim ws As Worksheet
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    Set ws = modInternalPlots.OutputSheet(sheetName)
    ws.Cells(1, 3).Value = "X start"
    ws.Cells(1, 4).Value = 0.1
    Call modInternalPlots.HighlightYellow(ws.Cells(1, 4))
    ReDim a(1 To 100, 1 To 6)
    For i = 1 To 100
        If i = 1 Then
            a(i, 1) = 0.1
            a(i, 2) = 0.5
            a(i, 3) = 0.5
            a(i, 4) = "=ABS(R[-2]C)"
        Else
            a(i, 1) = "=R[-1]C"
            a(i, 2) = "=R[-1]C"
            a(i, 3) = "=R[-1]C"
            a(i, 4) = "=R[-1]C+RC[-3]"
        End If
        a(i, 5) = "=LOGNORM.DIST(RC[-1],RC[-3],ABS(RC[-2]),FALSE)"
        a(i, 6) = "=LOGNORM.DIST(RC[-2],RC[-4],ABS(RC[-3]),TRUE)"
    Next i
    Call modInternalPlots.WriteBlock(ws, startRow, startCol, a)
    Call FinishDist(ws, startRow, startCol, 100, 6, _
        Array("x increment", "Mean", "Standard deviation", "x", "Log-Normal distribution", "CDF"), _
        3, 4, 5, 6, "Log-Normal distribution plot", "Log-Normal CDF plot", "x", "Probability", True)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("GenerateLogNormalPlot")
End Sub

Private Sub RunHypergeometric()
    Dim a As Variant
    Dim i As Long
    Dim ws As Worksheet
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    Set ws = modInternalPlots.OutputSheet("Hypergeometric")
    ReDim a(1 To 101, 1 To 6)
    For i = 1 To 101
        If i = 1 Then
            a(i, 1) = 0
            a(i, 2) = 100
            a(i, 3) = 50
            a(i, 4) = 500
        Else
            a(i, 1) = "=R[-1]C+1"
            a(i, 2) = "=R[-1]C"
            a(i, 3) = "=R[-1]C"
            a(i, 4) = "=R[-1]C"
        End If
        a(i, 5) = "=HYPGEOM.DIST(RC[-4],RC[-3],RC[-2],RC[-1],FALSE)"
        a(i, 6) = "=HYPGEOM.DIST(RC[-5],RC[-4],RC[-3],RC[-2],TRUE)"
    Next i
    Call modInternalPlots.WriteBlock(ws, 2, 1, a)
    Call FinishDist(ws, 2, 1, 101, 6, _
        Array("Number of successes in the sample,k", "Sample size,n", "Number of success in the population,K", "Population size,N", "Hypergeometric distribution", "CDF"), _
        4, 1, 5, 6, "Hypergeometric distribution plot", "Hypergeometric CDF plot", _
        "Number of successes in the sample, k (without replacement)", "Probability", False)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("GenerateHypergeometricPlot")
End Sub

Private Sub RunLogistic()
    Dim a As Variant
    Dim i As Long
    Dim ws As Worksheet
    Dim rng As Range
    Dim ch As ChartObject
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    Set ws = modInternalPlots.OutputSheet("Logistics curve")
    ReDim a(1 To 101, 1 To 6)
    For i = 1 To 101
        If i = 1 Then
            a(i, 1) = 1
            a(i, 2) = 1
            a(i, 3) = 0
            a(i, 4) = 0.25
            a(i, 5) = -50
        Else
            a(i, 1) = "=R[-1]C"
            a(i, 2) = "=R[-1]C"
            a(i, 3) = "=R[-1]C"
            a(i, 4) = "=R[-1]C"
            a(i, 5) = "=R[-1]C+RC[-4]"
        End If
        a(i, 6) = "=RC[-4]/(1+EXP(-RC[-2]*(RC[-1]-RC[-3])))"
    Next i
    Call modInternalPlots.WriteBlock(ws, 2, 1, a)
    Call modInternalPlots.WriteHeaderRow(ws.Cells(1, 1), Array("x increment", "a, Max", "x midpoint", "k, rate", "x", "Logistic curve"))
    Call modInternalPlots.HighlightYellow(ws.Range(ws.Cells(2, 1), ws.Cells(2, 5)))
    Set rng = ws.Range(ws.Cells(2, 1), ws.Cells(102, 6))
    rng.Columns(1).NumberFormat = "0.0"
    Set ch = modInternalPlots.AddChartAt(ws, ws.Cells(2, 8), 260, 260)
    ch.Chart.HasTitle = True
    ch.Chart.ChartTitle.Text = "Logistic function / curve"
    Call modInternalPlots.AddXySeries(ch.Chart, "Logistic", rng.Columns(5), rng.Columns(6), xlXYScatterSmooth, RGB(255, 0, 0), False)
    Call modInternalPlots.StyleValueAxes(ch.Chart, "x", "y", "0", "0.00", False)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("GenerateLogisticCurve")
End Sub

Private Sub FinishDist(ByVal ws As Worksheet, ByVal startRow As Long, ByVal startCol As Long, _
                       ByVal nRows As Long, ByVal nCols As Long, ByVal headers As Variant, _
                       ByVal yellowCols As Long, ByVal xCol As Long, ByVal pdfCol As Long, ByVal cdfCol As Long, _
                       ByVal pdfTitle As String, ByVal cdfTitle As String, _
                       ByVal xTitle As String, ByVal yTitle As String, ByVal markers As Boolean)
    Dim rng As Range
    Dim ch As ChartObject
    Call modInternalPlots.WriteHeaderRow(ws.Cells(startRow - 1, startCol), headers)
    Set rng = ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow + nRows - 1, startCol + nCols - 1))
    Call modInternalPlots.HighlightYellow(ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow, startCol + yellowCols - 1)))
    rng.Font.Size = 8
    Set ch = modInternalPlots.AddChartAt(ws, ws.Cells(startRow, startCol + 6), 260, 260)
    ch.Chart.HasTitle = True
    ch.Chart.ChartTitle.Text = pdfTitle
    Call modInternalPlots.AddXySeries(ch.Chart, "", rng.Columns(xCol), rng.Columns(pdfCol), xlXYScatterSmooth, RGB(255, 0, 0), markers)
    Call modInternalPlots.StyleValueAxes(ch.Chart, xTitle, yTitle, "0", "0.00", True)
    If cdfCol > 0 Then
        Set ch = modInternalPlots.AddChartAt(ws, ws.Cells(startRow, startCol + 12), 260, 260)
        ch.Chart.HasTitle = True
        ch.Chart.ChartTitle.Text = cdfTitle
        Call modInternalPlots.AddXySeries(ch.Chart, "", rng.Columns(xCol), rng.Columns(cdfCol), xlXYScatterSmooth, RGB(255, 0, 0), markers)
        Call modInternalPlots.StyleValueAxes(ch.Chart, xTitle, yTitle, "0", "0.00", True)
    End If
End Sub
