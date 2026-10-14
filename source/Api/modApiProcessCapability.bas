Attribute VB_Name = "modApiProcessCapability"
Option Explicit

' Public API: Personal Custom_Menu11_ProcessCapability.
' Cp / Cpk are worksheet formulae (Personal called personal.xlsb UDFs).

Private Const SheetName As String = "Process Capability Plot"
Private Const NInc As Long = 100

''' @Description: Process capability plot with Cp, Cpk, USL/LSL overlays. Prompts for data, LSL, and USL.
''' @Example: ProcessCapabilityChart
Public Sub ProcessCapabilityChart()
    Dim src As Range
    Dim lsl As Variant
    Dim usl As Variant
    Dim tmp As Double
    Dim ws As Worksheet
    Dim n As Long
    Dim i As Long
    Dim z As Long
    Dim ch As ChartObject
    Dim plot As Range
    On Error GoTo EH
    Set src = modInternalPlots.PromptRange("Select the process output data range.")
    If src Is Nothing Then Exit Sub
    If Not modInternalPlots.RangeIsAllNumeric(src) Then
        MsgBox "Process data must be numeric.", vbExclamation, "Plots Charts"
        Exit Sub
    End If
    lsl = modInternalPlots.PromptNumber("Lower specification limit", -10)
    If IsEmpty(lsl) Then Exit Sub
    usl = modInternalPlots.PromptNumber("Upper specification limit", 10)
    If IsEmpty(usl) Then Exit Sub
    If CDbl(lsl) > CDbl(usl) Then
        tmp = CDbl(lsl)
        lsl = CDbl(usl)
        usl = tmp
    End If
    Call modInternalExcelApp.PushAppState
    Set ws = modInternalPlots.OutputSheet(SheetName)
    ws.Range("A1").Value = "Process capability"
    ws.Range("A1").Font.Bold = True
    ws.Range("A3").Value = "Process data"
    Call modInternalPlots.WriteHeaderRow(ws.Range("Q3"), Array("Z-score", "X", "Normal", "USL", "LSL", "Sigma", "Actual frequency", "Normalised frequency"))
    n = src.Cells.Count
    Call modInternalPlots.FlattenToColumn(src, ws.Range("A4"))
    Call modInternalNamedRanges.CreateNamedRange("ProcessOutputData", ws.Range("A4").Resize(n, 1))
    ws.Range("C4").Value = "Mean"
    ws.Range("D4").Formula = "=AVERAGE(ProcessOutputData)"
    ws.Range("C5").Value = "Std Dev"
    ws.Range("D5").Formula = "=STDEV.S(ProcessOutputData)"
    ws.Range("C6").Value = "Upper specification limit, USL"
    ws.Range("D6").Value = CDbl(usl)
    ws.Range("C7").Value = "Lower specification limit, LSL"
    ws.Range("D7").Value = CDbl(lsl)
    ws.Range("C8").Value = "Number of data points"
    ws.Range("D8").Formula = "=COUNTA(ProcessOutputData)"
    ws.Range("C9").Value = "Plot range"
    ws.Range("D9").FormulaR1C1 = "=R5C4*6*2"
    ws.Range("C10").Value = "Plot increment"
    ws.Range("D10").FormulaR1C1 = "=R[-1]C/100"
    ws.Range("C12").Value = "Process capability, Cp"
    ws.Range("D12").FormulaR1C1 = "=(R6C4-R7C4)/(6*R5C4)"
    ws.Range("E12").Formula = "=IF(D12<1,""Not capable"",IF(D12<1.33,""Barely capable"",IF(D12<1.67,""Capable"",""Excellent"")))"
    ws.Range("C13").Value = "Capability ratio, Cr"
    ws.Range("D13").FormulaR1C1 = "=1/R[-1]C"
    ws.Range("E13").Formula = "=IF(D13>1,""Not capable"",IF(D13>0.75,""Barely capable"",""Capable""))"
    ws.Range("C15").Value = "Process capability Upper, Cpu"
    ws.Range("D15").FormulaR1C1 = "=(R6C4-R4C4)/(3*R5C4)"
    ws.Range("C16").Value = "Process capability Lower, Cpl"
    ws.Range("D16").FormulaR1C1 = "=(R4C4-R7C4)/(3*R5C4)"
    ws.Range("C17").Value = "Process capability centralised, Cpk"
    ws.Range("D17").Formula = "=MIN(D15,D16)"
    ws.Range("E17").Formula = "=IF(D17<1,""Not capable"",IF(D17<1.33,""Barely capable"",IF(D17<1.67,""Capable"",""Excellent"")))"
    For i = 1 To NInc + 1
        ws.Cells(i + 3, 17).FormulaR1C1 = "=-(R4C4-RC[1])/R5C4"
        If i = 1 Then
            ws.Cells(i + 3, 18).FormulaR1C1 = "=(R4C4-6*R5C4)"
        Else
            ws.Cells(i + 3, 18).FormulaR1C1 = "=R[-1]C+R10C4"
        End If
        ws.Cells(i + 3, 19).FormulaR1C1 = "=NORM.DIST(RC[-1],R4C4,R5C4,FALSE)"
        Select Case i
            Case 1
                ws.Cells(i + 3, 23).FormulaR1C1 = "=COUNTIF(ProcessOutputData,""<=""&RC[-5])"
            Case NInc + 1
                ws.Cells(i + 3, 23).FormulaR1C1 = "=COUNTIF(ProcessOutputData,"">""&RC[-5])"
            Case Else
                ws.Cells(i + 3, 23).FormulaR1C1 = "=COUNTIFS(ProcessOutputData,""<=""&RC[-5],ProcessOutputData,"">""&R[-1]C[-5])"
        End Select
        ws.Cells(i + 3, 24).FormulaR1C1 = "=RC[-1]/R8C4"
    Next i
    z = -6
    For i = NInc + 4 To NInc + 16
        ws.Cells(i + 3, 17).Value = z
        If i = NInc + 4 Then
            ws.Cells(i + 3, 18).FormulaR1C1 = "=(R4C4-6*R5C4)"
        Else
            ws.Cells(i + 3, 18).FormulaR1C1 = "=R[-1]C+R5C4"
        End If
        ws.Cells(i + 3, 22).FormulaR1C1 = "=MAX(R4C19:R104C19)*0.75"
        z = z + 1
    Next i
    ws.Range("R105").FormulaR1C1 = "=R6C4"
    ws.Range("T105").FormulaR1C1 = "=MAX(R[-101]C[-1]:R[-1]C[-1])"
    ws.Range("R106").FormulaR1C1 = "=R7C4"
    ws.Range("U106").FormulaR1C1 = "=MAX(R[-102]C[-2]:R[-2]C[-2])"
    ws.Columns(1).ColumnWidth = 8
    ws.Columns(2).ColumnWidth = 1
    ws.Range("C:E").EntireColumn.AutoFit
    ws.Range("Q:X").ColumnWidth = 6
    ws.Range("Q:R").NumberFormat = "0.00"
    ws.Range("S:V").NumberFormat = "0.0000"
    Set plot = ws.Range("Q3").CurrentRegion
    Set ch = modInternalPlots.AddChartAt(ws, ws.Range("F1"), 500, 270)
    With ch.Chart
        .HasTitle = True
        .ChartTitle.Text = "Process capability with upper and lower specification limits"
        .HasLegend = True
        Call modInternalPlots.AddXySeries(ch.Chart, "Normal", plot.Columns(2).Offset(1).Resize(plot.Rows.Count - 1), plot.Columns(3).Offset(1).Resize(plot.Rows.Count - 1), xlXYScatter, RGB(0, 0, 255), True)
        Call modInternalPlots.AddXySeries(ch.Chart, "USL", plot.Columns(2).Offset(1).Resize(plot.Rows.Count - 1), plot.Columns(4).Offset(1).Resize(plot.Rows.Count - 1), xlXYScatter, RGB(255, 0, 0), True)
        Call modInternalPlots.AddXySeries(ch.Chart, "LSL", plot.Columns(2).Offset(1).Resize(plot.Rows.Count - 1), plot.Columns(5).Offset(1).Resize(plot.Rows.Count - 1), xlXYScatter, RGB(255, 0, 0), True)
        Call modInternalPlots.AddXySeries(ch.Chart, "Sigma", plot.Columns(2).Offset(1).Resize(plot.Rows.Count - 1), plot.Columns(6).Offset(1).Resize(plot.Rows.Count - 1), xlXYScatter, RGB(100, 100, 100), True)
        Call modInternalPlots.AddXySeries(ch.Chart, "Frequency", plot.Columns(2).Offset(1).Resize(plot.Rows.Count - 1), plot.Columns(8).Offset(1).Resize(plot.Rows.Count - 1), xlXYScatterLinesNoMarkers, RGB(255, 0, 0), False)
        Call modInternalPlots.StyleValueAxes(ch.Chart, "Outputs", "Probability", "0", "0.00", False)
        .HasLegend = True
    End With
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("ProcessCapabilityChart")
End Sub
