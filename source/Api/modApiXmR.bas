Attribute VB_Name = "modApiXmR"
Option Explicit

' Public API: Personal Custom_Menu11_XmR / XmRFull.
' Constants: E2 = 2.66 (X limits), D4 = 3.267 (MR UCL). First moving range is blank.

Private Const E2 As Double = 2.66
Private Const D4 As Double = 3.267

''' @Description: Individuals and moving-range (XmR) chart from two columns: labels plus numeric values (header required).
''' @Example: XmR
Public Sub XmR()
    Dim src As Range
    Dim ws As Worksheet
    Dim n As Long
    Dim i As Long
    Dim hdrX As String
    Dim hdrY As String
    Dim rngX As Range
    Dim rngY As Range
    Dim rngMR As Range
    Dim rngAvg As Range
    Dim rngUcl As Range
    Dim rngLcl As Range
    Dim rngMrUcl As Range
    Dim rngMrAvg As Range
    Dim ch As ChartObject
    On Error GoTo EH
    Set src = modInternalPlots.PromptRange("Select two columns: labels (with header) and numeric values.")
    If src Is Nothing Then Exit Sub
    If src.Columns.Count <> 2 Then
        MsgBox "The input range must have 2 columns.", vbExclamation, "Plots Charts"
        Exit Sub
    End If
    If src.Rows.Count < 3 Then
        MsgBox "Need a header row and at least 2 data points.", vbExclamation, "Plots Charts"
        Exit Sub
    End If
    If Not modInternalPlots.HeaderRowLooksLikeHeader(src) Then
        MsgBox "The first row must be a non-numeric header.", vbExclamation, "Plots Charts"
        Exit Sub
    End If
    If Not modInternalPlots.RangeIsAllNumeric(src.Columns(2).Offset(1, 0).Resize(src.Rows.Count - 1, 1)) Then
        MsgBox "Column 2 data cells must be numeric.", vbExclamation, "Plots Charts"
        Exit Sub
    End If
    n = src.Rows.Count - 1
    hdrX = CStr(src.Cells(1, 1).Value)
    hdrY = CStr(src.Cells(1, 2).Value)
    Call modInternalExcelApp.PushAppState
    Set ws = modInternalPlots.OutputSheet("XmR")
    Call modInternalPlots.WriteHeaderRow(ws.Range("A5"), Array(hdrX, hdrY, "Moving Range", "Data Average", "Moving Range Average", "XmR UCL", "XmR LCL", "Data comment", "Moving Average UCL", "Moving Average LCL", "Moving range comment"))
    For i = 1 To n
        ws.Cells(5 + i, 1).Value = src.Cells(i + 1, 1).Value
        ws.Cells(5 + i, 2).Value = src.Cells(i + 1, 2).Value
        If i = 1 Then
            ws.Cells(6, 3).Value = ""
        Else
            ws.Cells(5 + i, 3).FormulaR1C1 = "=ABS(RC[-1]-R[-1]C[-1])"
        End If
        ws.Cells(5 + i, 4).Formula = "=AVERAGE(XmR_Data)"
        ws.Cells(5 + i, 5).Formula = "=AVERAGE(XmR_MR)"
        ws.Cells(5 + i, 6).FormulaR1C1 = "=RC[-2]+2.66*RC[-1]"
        ws.Cells(5 + i, 7).FormulaR1C1 = "=RC[-3]-2.66*RC[-2]"
        ws.Cells(5 + i, 8).FormulaR1C1 = "=IF(RC[-6]=RC[-4],""Average"",IF(RC[-6]<RC[-4],""Below"",""Above""))"
        ws.Cells(5 + i, 9).FormulaR1C1 = "=RC[-4]*3.267"
        ws.Cells(5 + i, 10).Value = 0
        ws.Cells(5 + i, 11).FormulaR1C1 = "=IF(RC[-8]=RC[-6],""Average"",IF(RC[-8]<RC[-6],""Below"",""Above""))"
    Next i
    Set rngX = ws.Range("A6").Resize(n, 1)
    Set rngY = ws.Range("B6").Resize(n, 1)
    Set rngMR = ws.Range("C6").Resize(n, 1)
    Set rngAvg = ws.Range("D6").Resize(n, 1)
    Set rngMrAvg = ws.Range("E6").Resize(n, 1)
    Set rngUcl = ws.Range("F6").Resize(n, 1)
    Set rngLcl = ws.Range("G6").Resize(n, 1)
    Set rngMrUcl = ws.Range("I6").Resize(n, 1)
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "XmR_Data", rngY)
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "XmR_MR", rngMR)
    ws.Range("A6:K" & (5 + n)).Font.Size = 8
    ws.Range("C:K").NumberFormat = "#,##0.0000"
    ws.Columns("A:K").ColumnWidth = 10
    Set ch = ws.ChartObjects.Add(Left:=530, Top:=60, Width:=500, Height:=150)
    ch.Chart.HasTitle = True
    ch.Chart.ChartTitle.Text = "XmR Plot - Individual value"
    Call AddLimitChart(ch.Chart, rngX, rngY, rngAvg, rngUcl, rngLcl, hdrX, hdrY)
    Set ch = ws.ChartObjects.Add(Left:=530, Top:=220, Width:=500, Height:=150)
    ch.Chart.HasTitle = True
    ch.Chart.ChartTitle.Text = "XmR Plot - Moving range"
    Call AddLimitChart(ch.Chart, rngX, rngMR, rngMrAvg, rngMrUcl, ws.Range("J6").Resize(n, 1), hdrX, "Moving Range")
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("XmR")
End Sub

''' @Description: XmR diagnostics with Western Electric rules, zone labels, and X / mR charts. One or two columns, no header.
''' @Example: GenerateXMRDiagnostics
Public Sub GenerateXMRDiagnostics()
    Dim src As Range
    Dim ws As Worksheet
    Dim n As Long
    Dim i As Long
    Dim twoCol As Boolean
    Dim lo As ListObject
    On Error GoTo EH
    Set src = modInternalPlots.PromptRange("Select input range column(s) excluding headers (1 or 2 columns).")
    If src Is Nothing Then Exit Sub
    If src.Columns.Count > 2 Then
        MsgBox "Select at most 2 columns: labels/date and values.", vbExclamation, "Plots Charts"
        Exit Sub
    End If
    n = src.Rows.Count
    If n < 2 Then
        MsgBox "Need at least 2 rows.", vbExclamation, "Plots Charts"
        Exit Sub
    End If
    twoCol = (src.Columns.Count = 2)
    Call modInternalExcelApp.PushAppState
    Set ws = modInternalPlots.OutputSheet("XMR Diagnostics")
    Call modInternalPlots.WriteHeaderRow(ws.Range("A3"), Array( _
        "Sequence", "Amount", "X-bar", "X_UCL", "X_LCL", "X_+1SD", "X_-1SD", "X_+2SD", "X_-2SD", _
        "mR", "mR-bar", "mR_UCL", "Zone Label", "Rule1: Beyond 3 Sigma", "Rule2: 8 Same Side", _
        "Rule3: 4 of 5 in Zone B", "Rule4: Trend of 6", "Rule5: 2 of 3 in Zone A", _
        "Rule6: 14 Consecutive Alternations", "Rule Breach Summary", "Dashboard Comment", _
        "Signal Tier", "Sign Deviation", "Zone Band", "Delta", "Trend Sign", "Alternation Flag", _
        "Standard deviation", "Above/Below flag", "Greater/lower flag", "Rolling count of alternating points"))
    For i = 1 To n
        If twoCol Then
            ws.Cells(3 + i, 1).Value = src.Cells(i, 1).Value
            ws.Cells(3 + i, 2).Value = src.Cells(i, 2).Value
        Else
            ws.Cells(3 + i, 1).Value = i
            ws.Cells(3 + i, 2).Value = src.Cells(i, 1).Value
        End If
        ws.Cells(3 + i, 3).Formula = "=AVERAGE($B$4:$B$" & (3 + n) & ")"
        ws.Cells(3 + i, 4).FormulaR1C1 = "=RC[-1]+2.66*RC[7]"
        ws.Cells(3 + i, 5).FormulaR1C1 = "=RC[-2]-2.66*RC[6]"
        ws.Cells(3 + i, 6).FormulaR1C1 = "=RC[-3]+RC[22]"
        ws.Cells(3 + i, 7).FormulaR1C1 = "=RC[-4]-RC[21]"
        ws.Cells(3 + i, 8).FormulaR1C1 = "=RC[-5]+2*RC[20]"
        ws.Cells(3 + i, 9).FormulaR1C1 = "=RC[-6]-2*RC[19]"
        If i = 1 Then
            ws.Cells(4, 10).Value = ""
        Else
            ws.Cells(3 + i, 10).FormulaR1C1 = "=ABS(RC[-8]-R[-1]C[-8])"
        End If
        ws.Cells(3 + i, 11).Formula = "=AVERAGE($J$4:$J$" & (3 + n) & ")"
        ws.Cells(3 + i, 12).FormulaR1C1 = "=3.267*RC[-1]"
        ws.Cells(3 + i, 13).FormulaR1C1 = "=IF(ABS(RC[-11]-RC[-10])<=RC[15],""Zone C"",IF(ABS(RC[-11]-RC[-10])<=2*RC[15],""Zone B"",IF(ABS(RC[-11]-RC[-10])<=3*RC[15],""Zone A"",""Beyond Zone A"")))"
        ws.Cells(3 + i, 14).FormulaR1C1 = "=IF(OR(RC[-12]>RC[-10],RC[-12]<RC[-9]),""Rule 1 breach"","""")"
        If i >= 8 Then
            ws.Cells(3 + i, 15).FormulaR1C1 = "=IF(OR(COUNTIF(R[-7]C[8]:RC[8],-1)=8,COUNTIF(R[-7]C[8]:RC[8],1)=8),""Rule 2 breach"","""")"
        End If
        If i >= 5 Then
            ws.Cells(3 + i, 16).FormulaR1C1 = "=IF(COUNTIF(R[-4]C[-3]:RC[-3],""Zone B"")+COUNTIF(R[-4]C[-3]:RC[-3],""Zone A"")+COUNTIF(R[-4]C[-3]:RC[-3],""Beyond Zone A"")>=4,""Rule 3 breach"","""")"
        End If
        If i >= 6 Then
            ws.Cells(3 + i, 17).FormulaR1C1 = "=IF(OR(COUNTIF(R[-5]C[13]:RC[13],""Greater"")=6,COUNTIF(R[-5]C[13]:RC[13],""Lower"")=6),""Rule 4 breach"","""")"
        End If
        If i >= 3 Then
            ws.Cells(3 + i, 18).FormulaR1C1 = "=IF(AND(COUNTIF(R[-2]C[-5]:RC[-5],""Zone A"")+COUNTIF(R[-2]C[-5]:RC[-5],""Beyond Zone A"")>=2,OR(COUNTIF(R[-2]C[11]:RC[11],""Above"")=3,COUNTIF(R[-2]C[11]:RC[11],""Below"")=3)),""Rule 5 breach"","""")"
        End If
        If i >= 14 Then
            ws.Cells(3 + i, 19).FormulaR1C1 = "=IF(SUM(R[-13]C[8]:RC[8])=14,""Rule 6 breach"","""")"
        End If
        ws.Cells(3 + i, 20).FormulaR1C1 = "=TEXTJOIN(""; "",TRUE,RC[-6]:RC[-1])"
        ws.Cells(3 + i, 21).FormulaR1C1 = "=IF(RC[-1]<>"""",""Breach"",""No breach"")"
        If i = 1 Then
            ws.Cells(4, 22).Value = "Tier 0"
        Else
            ws.Cells(3 + i, 22).FormulaR1C1 = "=IF(RC[-2]<>"""",IF(COUNTIF(R[-1]C[-8]:R[1]C[-8],""Rule 1 breach"")>0,""Tier 1"",IF(COUNTIF(R[-1]C[-5]:R[1]C[-5],""Rule 4 breach"")+COUNTIF(R[-1]C[-3]:R[1]C[-3],""Rule 6 breach"")>0,""Tier 2"",""Tier 3"")),""Tier 0"")"
        End If
        ws.Cells(3 + i, 23).FormulaR1C1 = "=IF(RC[-21]>RC[-20],1,IF(RC[-21]<RC[-20],-1,0))"
        ws.Cells(3 + i, 24).FormulaR1C1 = "=IF(ABS(RC[-22]-RC[-21])<=RC[4],""Zone C"",IF(ABS(RC[-22]-RC[-21])<=2*RC[4],""Zone B"",IF(ABS(RC[-22]-RC[-21])<=3*RC[4],""Zone A"",""Beyond Zone A"")))"
        If i > 1 Then
            ws.Cells(3 + i, 25).FormulaR1C1 = "=RC[-23]-R[-1]C[-23]"
            ws.Cells(3 + i, 27).FormulaR1C1 = "=IF(RC[-1]*R[-1]C[-1]=-1,1,0)"
            ws.Cells(3 + i, 31).FormulaR1C1 = "=IF(SIGN(RC[-6])<>SIGN(R[-1]C[-6]),R[-1]C+1,0)"
        End If
        ws.Cells(3 + i, 26).FormulaR1C1 = "=SIGN(RC[-1])"
        ws.Cells(3 + i, 28).Formula = "=STDEV($B$4:$B$" & (3 + n) & ")"
        ws.Cells(3 + i, 29).FormulaR1C1 = "=IF(RC[-27]=RC[-26],""On"",IF(RC[-27]>RC[-26],""Above"",""Below""))"
        If i > 1 Then
            ws.Cells(3 + i, 30).FormulaR1C1 = "=IF(RC[-28]=R[-1]C[-28],""Equal"",IF(RC[-28]>R[-1]C[-28],""Greater"",""Lower""))"
        End If
    Next i
    ws.Rows(3).Font.Bold = True
    ws.Rows(3).WrapText = True
    ws.Rows(3).RowHeight = 30
    ws.Columns("C:L").NumberFormat = "#,##0.0000"
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A3").CurrentRegion, , xlYes)
    lo.Name = "tbl_XmR_Analysis"
    lo.TableStyle = "TableStyleLight8"
    Call PlotXmrDiagnostics(ws, n)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("GenerateXMRDiagnostics")
End Sub

Private Sub AddLimitChart(ByVal ch As Chart, ByVal xRng As Range, ByVal yRng As Range, _
                          ByVal avgRng As Range, ByVal uclRng As Range, ByVal lclRng As Range, _
                          ByVal xTitle As String, ByVal yTitle As String)
    ch.HasLegend = False
    Call modInternalPlots.AddXySeries(ch, "Data", xRng, yRng, xlXYScatterLines, RGB(0, 0, 0), True)
    Call modInternalPlots.AddXySeries(ch, "Average", xRng, avgRng, xlXYScatterLines, RGB(0, 0, 255), False)
    Call modInternalPlots.AddXySeries(ch, "UCL", xRng, uclRng, xlXYScatterLines, RGB(255, 0, 0), False)
    Call modInternalPlots.AddXySeries(ch, "LCL", xRng, lclRng, xlXYScatterLines, RGB(0, 150, 0), False)
    Call modInternalPlots.StyleValueAxes(ch, xTitle, yTitle, "0", "0.00", False)
End Sub

Private Sub PlotXmrDiagnostics(ByVal ws As Worksheet, ByVal n As Long)
    Dim plot As Range
    Dim ch As ChartObject
    Set plot = ws.Range("A4").Resize(n, 12)
    Set ch = modInternalPlots.AddChartAt(ws, ws.Range("B4"), 475, 300)
    ch.Chart.HasTitle = True
    ch.Chart.ChartTitle.Text = "X bar Plot"
    ch.Chart.HasLegend = True
    Call modInternalPlots.AddXySeries(ch.Chart, "x", plot.Columns(1), plot.Columns(2), xlLine, RGB(0, 180, 0), False)
    Call modInternalPlots.AddXySeries(ch.Chart, "X bar", plot.Columns(1), plot.Columns(3), xlLine, RGB(0, 0, 255), False)
    Call modInternalPlots.AddXySeries(ch.Chart, "UCL", plot.Columns(1), plot.Columns(4), xlLine, RGB(255, 0, 0), False)
    Call modInternalPlots.AddXySeries(ch.Chart, "LCL", plot.Columns(1), plot.Columns(5), xlLine, RGB(255, 0, 0), False)
    Call modInternalPlots.StyleValueAxes(ch.Chart, "Data", "", "0", "0.0", False)
    ch.Chart.HasLegend = True
    Set ch = modInternalPlots.AddChartAt(ws, ws.Range("I4"), 475, 300)
    ch.Chart.HasTitle = True
    ch.Chart.ChartTitle.Text = "Moving Range Plot"
    ch.Chart.HasLegend = True
    Call modInternalPlots.AddXySeries(ch.Chart, "mR", plot.Columns(1), plot.Columns(10), xlLine, RGB(0, 180, 0), False)
    Call modInternalPlots.AddXySeries(ch.Chart, "Upper Control Limit", plot.Columns(1), plot.Columns(12), xlLine, RGB(255, 0, 0), False)
    Call modInternalPlots.StyleValueAxes(ch.Chart, "Data", "", "0", "0.0", False)
    ch.Chart.HasLegend = True
End Sub
