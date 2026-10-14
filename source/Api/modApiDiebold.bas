Attribute VB_Name = "modApiDiebold"
Option Explicit

' Public API: Personal Custom_Menu11_Diebold.
' One column: actuals plus two demo RAND forecasts (yellow). Three columns: actual, forecast 1, forecast 2.

Private Const SheetName As String = "DieboldMariano"

''' @Description: Diebold-Mariano test of two forecasts vs actuals, with HLN small-sample correction.
''' @Example: DieboldMarianoTest
Public Sub DieboldMarianoTest()
    Dim src As Range
    Dim ws As Worksheet
    Dim n As Long
    Dim i As Long
    Dim useDemo As Boolean
    Dim dataRng As Range
    Dim ch As ChartObject
    On Error GoTo EH
    Set src = modInternalPlots.PromptRange("Select actuals (1 column) or actual + two forecast columns (3 columns).")
    If src Is Nothing Then Exit Sub
    n = src.Rows.Count
    If n < 3 Then
        MsgBox "Need at least 3 rows.", vbExclamation, "Plots Charts"
        Exit Sub
    End If
    If src.Columns.Count <> 1 And src.Columns.Count <> 3 Then
        MsgBox "Select 1 column (demo forecasts) or 3 columns (actual, forecast 1, forecast 2).", vbExclamation, "Plots Charts"
        Exit Sub
    End If
    useDemo = (src.Columns.Count = 1)
    Call modInternalExcelApp.PushAppState
    Set ws = modInternalPlots.OutputSheet(SheetName)
    ws.Range("A1").Value = "Diebold-Mariano Test"
    ws.Range("A1").Font.Bold = True
    If useDemo Then
        ws.Range("A2").Value = "Forecast columns are RAND() placeholders. Replace them, or re-run with three columns: actual, forecast 1, forecast 2."
    End If
    Call modInternalPlots.WriteHeaderRow(ws.Range("H5"), Array("Data", "Forecast 1", "Forecast 2", "", "Index", "Error 1", "Error 2", _
        "Loss-differential, d", "gamma", "Std Err", "Diebold-Mariano statistic", "p-value", "p-value > Alpha", "Comment", "Order", "HLN statistic", "HLN p-value", "HLN comment"))
    For i = 1 To n
        ws.Cells(5 + i, 8).Value = src.Cells(i, 1).Value
        If useDemo Then
            ws.Cells(5 + i, 9).FormulaR1C1 = "=RC[-1]+RAND()*1.0005"
            ws.Cells(5 + i, 10).FormulaR1C1 = "=RC[-2]+RAND()*0.9999"
        Else
            ws.Cells(5 + i, 9).Value = src.Cells(i, 2).Value
            ws.Cells(5 + i, 10).Value = src.Cells(i, 3).Value
        End If
        ws.Cells(5 + i, 12).Value = i
        ws.Cells(5 + i, 13).FormulaR1C1 = "=RC[-5]-RC[-4]"
        ws.Cells(5 + i, 14).FormulaR1C1 = "=RC[-6]-RC[-4]"
        ws.Cells(5 + i, 15).FormulaR1C1 = "=RC[-2]^2-RC[-1]^2"
        ws.Cells(5 + i, 16).FormulaR1C1 = "=IF(RC[6]>=Count,NA(),SUMPRODUCT(OFFSET(Differences,RC[6],0,Count-RC[6],1)-Mean_d,OFFSET(Differences,0,0,Count-RC[6])-Mean_d)/Count)"
        If i = 1 Then
            ws.Cells(5 + i, 17).FormulaR1C1 = "=SQRT(Var_d/Count)"
            ws.Cells(5 + i, 22).Formula = "=Lag"
        Else
            ws.Cells(5 + i, 17).FormulaR1C1 = "=SQRT((Var_d+2*SUMPRODUCT(R6C16:R[-1]C16))/Count)"
            ws.Cells(5 + i, 22).FormulaR1C1 = "=R[-1]C+1"
        End If
        ws.Cells(5 + i, 18).FormulaR1C1 = "=Mean_d/RC[-1]"
        ws.Cells(5 + i, 19).FormulaR1C1 = "=2*(1-NORMSDIST(ABS(RC[-1])))"
        ws.Cells(5 + i, 20).FormulaR1C1 = "=IF(RC[-1]>Alpha,TRUE,FALSE)"
        ws.Cells(5 + i, 21).FormulaR1C1 = "=IF(RC[-1]=TRUE,""NO significant difference between the forecasts"",""Significant difference between the forecasts"")"
        ws.Cells(5 + i, 23).FormulaR1C1 = "=RC[-5]*SQRT((Count+1-2*RC[-1]+RC[-1]*(RC[-1]-1))/Count)"
        ws.Cells(5 + i, 24).FormulaR1C1 = "=T.DIST.2T(ABS(RC[-1]),Count-1)"
        ws.Cells(5 + i, 25).FormulaR1C1 = "=IF(RC[-1]>Alpha,""NO significant difference between the forecasts"",""Significant difference between the forecasts"")"
    Next i
    Set dataRng = ws.Range("H6").Resize(n, 1)
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "Data", dataRng)
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "Differences", dataRng.Offset(0, 7))
    ws.Range("E1").Value = "Count, n"
    ws.Range("F1").Formula = "=COUNT(Data)"
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "Count", ws.Range("F1"))
    ws.Range("E2").Value = "Mean of d"
    ws.Range("F2").Formula = "=AVERAGE(Differences)"
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "Mean_d", ws.Range("F2"))
    ws.Range("E3").Value = "Variance of d"
    ws.Range("F3").Formula = "=VAR.P(Differences)"
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "Var_d", ws.Range("F3"))
    ws.Range("H1").Value = "Lag, k"
    ws.Range("I1").Value = 1
    Call modInternalPlots.HighlightYellow(ws.Range("I1"))
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "Lag", ws.Range("I1"))
    ws.Range("H2").Value = "Alpha"
    ws.Range("I2").Value = 0.05
    Call modInternalPlots.HighlightYellow(ws.Range("I2"))
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "Alpha", ws.Range("I2"))
    ws.Range("H3").Value = "Suggested order, h"
    ws.Range("I3").Formula = "=ROUND(Count^(1/3)+1,0)"
    ws.Range("L1").Value = "Harvey, Leybourne and Newbold (HLN) test - small samples"
    ws.Range("L1").Font.Bold = True
    Call modInternalPlots.HighlightYellow(ws.Range("I6").Resize(n, 1))
    Call modInternalPlots.HighlightYellow(ws.Range("J6").Resize(n, 1))
    ws.Columns("E").ColumnWidth = 16
    ws.Columns("H").ColumnWidth = 17
    ws.Columns("U").ColumnWidth = 44
    Set ch = modInternalPlots.AddChartAt(ws, ws.Range("A5"), 375, 300)
    ch.Chart.HasTitle = True
    ch.Chart.ChartTitle.Text = "Time-series forecasts"
    ch.Chart.HasLegend = True
    ch.Chart.Legend.Position = xlLegendPositionBottom
    Call modInternalPlots.AddXySeries(ch.Chart, "Data", dataRng.Offset(0, 4), dataRng, xlXYScatterLines, RGB(0, 0, 0), False)
    Call modInternalPlots.AddXySeries(ch.Chart, "Forecast 1", dataRng.Offset(0, 4), dataRng.Offset(0, 1), xlXYScatterLines, RGB(120, 120, 120), False)
    Call modInternalPlots.AddXySeries(ch.Chart, "Forecast 2", dataRng.Offset(0, 4), dataRng.Offset(0, 2), xlXYScatterLines, RGB(120, 120, 120), False)
    Call modInternalPlots.StyleValueAxes(ch.Chart, "Lag", "", "0", "0.00", False)
    ch.Chart.HasLegend = True
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("DieboldMarianoTest")
End Sub
