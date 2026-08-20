Attribute VB_Name = "modApiDeming"
Option Explicit

' Public API: Personal Custom_Menu5_DemingRegression.
' Two numeric columns (X, Y) → sheet Deming Regression with live formulae for λ, intercept, slope, and residuals.

Private Const SheetName As String = "Deming Regression"

''' @Description: Deming regression (error in both X and Y). λ = VAR.S(X)/VAR.S(Y). Writes fitted values, residuals, and an XY chart with the fit line.
''' @Example: CalculateDemingRegression
Public Sub CalculateDemingRegression()
    Dim src As Range
    Dim ws As Worksheet
    Dim n As Long
    Dim raw As Variant
    Dim rngX As Range
    Dim rngY As Range
    Dim lastData As Long
    Dim minX As Double
    Dim maxX As Double
    Dim ch As ChartObject
    On Error GoTo EH
    Set src = modInternalAnalysis.PromptNumericRange( _
        "Select two columns of paired X, Y values, without a header row.", _
        "Deming Regression", 2, 2, 3)
    If src Is Nothing Then Exit Sub
    n = src.Rows.Count
    raw = src.Value
    minX = Application.WorksheetFunction.Min(src.Columns(1))
    maxX = Application.WorksheetFunction.Max(src.Columns(1))
    If minX = maxX Then
        minX = minX - 1
        maxX = maxX + 1
    End If
    lastData = 3 + n
    Call modInternalExcelApp.PushAppState
    Call modApiSheets.CreateOutputSheet(SheetName)
    Set ws = ActiveWorkbook.Worksheets(SheetName)
    ws.Range("A1").Value = "Deming regression"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 12
    ws.Range("A2").Value = "Produced " & Format$(Now, "yyyy-mm-dd hh:nn")
    ws.Range("A4").Value = "X bar"
    ws.Range("A5").Value = "Y bar"
    ws.Range("A6").Value = "u (DEVSQ X)"
    ws.Range("A7").Value = "v (DEVSQ Y)"
    ws.Range("A8").Value = "Var.S X"
    ws.Range("A9").Value = "Var.S Y"
    ws.Range("A10").Value = "r"
    ws.Range("A11").Value = "n"
    ws.Range("A12").Value = "Lambda"
    ws.Range("A13").Value = "Alpha"
    ws.Range("A14").Value = "Beta"
    On Error Resume Next
    ws.Range("A12").AddComment "VAR.S(X) / VAR.S(Y). Personal Menu5 default; many texts assume λ = 1."
    Err.Clear
    On Error GoTo EH
    ws.Range("D3").Value = "X"
    ws.Range("E3").Value = "Y"
    ws.Range("F3").Value = "Predicted Y"
    ws.Range("G3").Value = "x-hat"
    ws.Range("H3").Value = "y-hat"
    ws.Range("I3").Value = "Raw residual"
    ws.Range("J3").Value = "X-residual"
    ws.Range("K3").Value = "Y-residual"
    ws.Range("L3").Value = "Optimised residual"
    ws.Range("D3:L3").Font.Bold = True
    ws.Range("D3:L3").HorizontalAlignment = xlCenter
    ws.Range("D3:L3").WrapText = True
    If n = 1 Then
        ws.Range("D4").Resize(1, 2).Value = raw
    Else
        ws.Range("D4").Resize(n, 2).Value = raw
    End If
    Set rngX = ws.Range("D4").Resize(n, 1)
    Set rngY = ws.Range("E4").Resize(n, 1)
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "rng_X", rngX)
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "rng_Y", rngY)
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "X_bar", ws.Range("B4"))
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "Y_bar", ws.Range("B5"))
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "DemingLambda", ws.Range("B12"))
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "Alpha", ws.Range("B13"))
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "Beta", ws.Range("B14"))
    ws.Range("B4").Formula = "=AVERAGE(rng_X)"
    ws.Range("B5").Formula = "=AVERAGE(rng_Y)"
    ws.Range("B6").Formula = "=DEVSQ(rng_X)"
    ws.Range("B7").Formula = "=DEVSQ(rng_Y)"
    ws.Range("B8").Formula = "=VAR.S(rng_X)"
    ws.Range("B9").Formula = "=VAR.S(rng_Y)"
    ws.Range("B10").Formula = "=SUMPRODUCT(rng_X-X_bar,rng_Y-Y_bar)"
    ws.Range("B11").Formula = "=COUNT(rng_X)"
    ws.Range("B12").Formula = "=B8/B9"
    ws.Range("B13").Formula = "=Y_bar-X_bar*Beta"
    ws.Range("B14").Formula = "=(DemingLambda*B7-B6+SQRT((DemingLambda*B7-B6)^2+4*DemingLambda*B10^2))/(2*DemingLambda*B10)"
    ws.Range("F4").Resize(n, 1).FormulaR1C1 = "=Alpha+RC[-2]*Beta"
    ws.Range("G4").Resize(n, 1).FormulaR1C1 = "=RC[-3]+DemingLambda*Beta*RC[2]/(DemingLambda*Beta^2+1)"
    ws.Range("H4").Resize(n, 1).FormulaR1C1 = "=RC[-3]-RC[1]/(DemingLambda*Beta^2+1)"
    ws.Range("I4").Resize(n, 1).FormulaR1C1 = "=RC[-4]-RC[-3]"
    ws.Range("J4").Resize(n, 1).FormulaR1C1 = "=RC[-6]-RC[-3]"
    ws.Range("K4").Resize(n, 1).FormulaR1C1 = "=RC[-6]-RC[-3]"
    ws.Range("L4").Resize(n, 1).FormulaR1C1 = "=SIGN(RC[-3])*SQRT(RC[-2]^2+DemingLambda*RC[-1]^2)"
    ws.Range("I" & (lastData + 1)).FormulaR1C1 = "=SUM(R[-" & n & "]C:R[-1]C)"
    ws.Range("J" & (lastData + 1)).FormulaR1C1 = "=SUM(R[-" & n & "]C:R[-1]C)"
    ws.Range("K" & (lastData + 1)).FormulaR1C1 = "=SUM(R[-" & n & "]C:R[-1]C)"
    ws.Range("L" & (lastData + 1)).FormulaR1C1 = "=SUM(R[-" & n & "]C:R[-1]C)"
    ws.Range("H" & (lastData + 1)).Value = "Sums"
    ws.Range("H" & (lastData + 1)).Font.Bold = True
    ws.Range("B4:B14").NumberFormat = "#,##0.0000"
    ws.Range("D4:L" & (lastData + 1)).NumberFormat = "#,##0.00"
    ws.Columns("A").ColumnWidth = 16
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 2
    ws.Range("D:L").ColumnWidth = 11
    ws.Range("N3").Value = "Fit X"
    ws.Range("O3").Value = "Fit Y"
    ws.Range("N3:O3").Font.Bold = True
    ws.Range("N4").Value = minX
    ws.Range("N5").Value = maxX
    ws.Range("O4").Formula = "=Alpha+N4*Beta"
    ws.Range("O5").Formula = "=Alpha+N5*Beta"
    Set ch = ws.ChartObjects.Add(Left:=ws.Range("N8").Left, Top:=ws.Range("N8").Top, Width:=420, Height:=260)
    With ch.Chart
        .ChartType = xlXYScatter
        .HasTitle = True
        .ChartTitle.Text = "Deming regression"
        .HasLegend = True
        .SeriesCollection.NewSeries
        With .SeriesCollection(1)
            .Name = "Observed"
            .XValues = rngX
            .Values = rngY
            .MarkerStyle = xlMarkerStyleCircle
            .MarkerSize = 5
        End With
        .SeriesCollection.NewSeries
        With .SeriesCollection(2)
            .Name = "Deming fit"
            .XValues = ws.Range("N4:N5")
            .Values = ws.Range("O4:O5")
            .ChartType = xlXYScatterLinesNoMarkers
            .Format.Line.ForeColor.RGB = RGB(255, 0, 0)
            .Format.Line.Weight = 1.25
        End With
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "X"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Y"
        .Axes(xlCategory).HasMajorGridlines = False
        .Axes(xlValue).HasMajorGridlines = False
    End With
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("CalculateDemingRegression")
End Sub
