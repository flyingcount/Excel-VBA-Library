Attribute VB_Name = "modApiChartSheet"
Option Explicit

' Public API: Personal Custom_Menu11_ChartSheetPlot.

''' @Description: Line chart on a chart sheet from a single column with a header. Adds SE bars and a trendline.
''' @Example: PlotLineChartSheet
Public Sub PlotLineChartSheet()
    Dim src As Range
    Dim hdr As Range
    Dim body As Range
    Dim chartName As String
    Dim cht As Chart
    On Error GoTo EH
    Set src = modInternalPlots.PromptRange("Select a single column of data including a header cell at the top." & vbCrLf & "Output will be a chart sheet.")
    If src Is Nothing Then Exit Sub
    If src.Columns.Count <> 1 Then
        MsgBox "Select a single column.", vbExclamation, "Plots Charts"
        Exit Sub
    End If
    If src.Rows.Count < 2 Then
        MsgBox "Include a header and at least one data cell.", vbExclamation, "Plots Charts"
        Exit Sub
    End If
    Set hdr = src.Cells(1, 1)
    Set body = src.Resize(src.Rows.Count - 1, 1).Offset(1, 0)
    chartName = CStr(hdr.Value)
    If Len(chartName) = 0 Then chartName = "Unknown"
    Call modInternalExcelApp.PushAppState
    Call modInternalPlots.DeleteChartSheetIfExists(chartName)
    Set cht = ActiveWorkbook.Charts.Add
    cht.Name = chartName
    With cht
        .ChartArea.Format.Line.Visible = False
        .HasTitle = True
        With .ChartTitle
            .Text = chartName & " with standard error bars"
            .Font.Name = "Arial"
            .Font.Size = 16
            .Font.Bold = True
        End With
        .HasLegend = True
        .Legend.Position = xlLegendPositionTop
        .HasDataTable = True
        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Text = "y-axis"
            .HasMajorGridlines = False
            .HasMinorGridlines = False
        End With
        With .Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.Text = "x-axis"
            .HasMajorGridlines = False
            .HasMinorGridlines = False
        End With
        With .SeriesCollection(1)
            .ChartType = xlLine
            .Name = hdr.Value
            .Values = body
            .Format.Line.ForeColor.RGB = RGB(255, 0, 0)
            .Format.Line.Weight = 0.25
            .ErrorBar Direction:=xlY, Include:=xlErrorBarIncludeBoth, Type:=xlErrorBarTypeStError
            .Trendlines.Add
            .Trendlines(1).DisplayEquation = True
            .Trendlines(1).DisplayRSquared = True
        End With
    End With
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("PlotLineChartSheet")
End Sub
