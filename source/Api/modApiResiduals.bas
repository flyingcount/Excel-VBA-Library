Attribute VB_Name = "modApiResiduals"
Option Explicit

' Public API: Personal Custom_Menu18_Residuals.
' One numeric column → sheet Residuals Analysis with order, sum, slope/intercept/R², and scatter + trendline.

Private Const SheetName As String = "Residuals Analysis"

''' @Description: Plot residuals (one numeric column) against observation order on a new sheet.
''' @Example: ResidualsAnalysis
Public Sub ResidualsAnalysis()
    Dim src As Range
    Dim ws As Worksheet
    Dim n As Long
    Dim i As Long
    Dim order() As Long
    Dim body As Variant
    Dim rngOrder As Range
    Dim rngResid As Range
    Dim ch As ChartObject
    On Error GoTo EH
    If TypeName(Selection) <> "Range" Then
        MsgBox "Select a single column of residual values.", vbExclamation, "Residuals"
        Exit Sub
    End If
    On Error Resume Next
    Set src = Application.InputBox(Prompt:="Select data in a single column range", _
                                   Title:="Residuals", Default:=Selection.Address, Type:=8)
    On Error GoTo 0
    If src Is Nothing Then Exit Sub
    If src.Columns.Count <> 1 Then
        MsgBox "Select a single column.", vbExclamation, "Residuals"
        Exit Sub
    End If
    n = src.Rows.Count
    If n < 3 Then
        MsgBox "Need at least 3 data points.", vbExclamation, "Residuals"
        Exit Sub
    End If
    If Application.WorksheetFunction.Count(src) <> n Then
        MsgBox "Every cell must be numeric (blanks and text are not allowed).", vbExclamation, "Residuals"
        Exit Sub
    End If
    body = src.Value
    ReDim order(1 To n, 1 To 1)
    For i = 1 To n
        order(i, 1) = i
    Next i
    Call modInternalExcelApp.PushAppState
    Call modApiSheets.CreateOutputSheet(SheetName)
    Set ws = ActiveWorkbook.Worksheets(SheetName)
    ws.Range("A1").Value = "Residuals Analysis"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 12
    ws.Range("A3").Value = "Order"
    ws.Range("B3").Value = "Residuals"
    ws.Range("A3:B3").Font.Bold = True
    ws.Range("A3:B3").HorizontalAlignment = xlCenter
    ws.Range("A4").Resize(n, 1).Value = order
    If n = 1 Then
        ws.Range("B4").Value = body
    Else
        ws.Range("B4").Resize(n, 1).Value = body
    End If
    Set rngOrder = ws.Range("A4").Resize(n, 1)
    Set rngResid = ws.Range("B4").Resize(n, 1)
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "Order", rngOrder)
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "Residuals", rngResid)
    ws.Range("D8").Value = "Sum of residuals"
    ws.Range("D8").HorizontalAlignment = xlRight
    ws.Range("E8").Formula = "=SUM(Residuals)"
    ws.Range("E8").NumberFormat = "#,##0.0000"
    ws.Range("D3").Value = "Linear regression statistics:"
    ws.Range("D3").Font.Bold = True
    ws.Range("D4").Value = "Slope"
    ws.Range("E4").Formula = "=SLOPE(Residuals,Order)"
    ws.Range("E4").NumberFormat = "0.0000"
    ws.Range("D5").Value = "Intercept"
    ws.Range("E5").Formula = "=INTERCEPT(Residuals,Order)"
    ws.Range("E5").NumberFormat = "0.0000"
    ws.Range("D6").Value = "R-squared"
    ws.Range("E6").Formula = "=RSQ(Residuals,Order)"
    ws.Range("E6").NumberFormat = "0.0000"
    ws.Columns("A:B").ColumnWidth = 10
    ws.Columns("C").ColumnWidth = 2
    ws.Columns("D").ColumnWidth = 22
    ws.Columns("E").ColumnWidth = 12
    Set ch = ws.ChartObjects.Add(Left:=ws.Range("G2").Left, Top:=ws.Range("G2").Top, Width:=500, Height:=250)
    With ch.Chart
        .ChartType = xlXYScatter
        .HasTitle = True
        .ChartTitle.Text = "Residuals Plot"
        .ChartTitle.Font.Size = 12
        .HasLegend = False
        .SeriesCollection.NewSeries
        With .SeriesCollection(1)
            .Name = "Residuals"
            .XValues = rngOrder
            .Values = rngResid
            .MarkerStyle = xlMarkerStyleX
            .MarkerSize = 3
            .Trendlines.Add
            .Trendlines(1).DisplayEquation = True
            .Trendlines(1).DisplayRSquared = True
        End With
        .Axes(xlCategory).HasTitle = False
        .Axes(xlValue).HasMajorGridlines = False
    End With
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("ResidualsAnalysis")
End Sub
