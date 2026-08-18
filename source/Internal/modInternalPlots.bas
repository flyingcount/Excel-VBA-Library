Attribute VB_Name = "modInternalPlots"
Option Explicit

' Internal: shared prompts, sheet I/O, and chart helpers for Plots Charts (Personal Menu11).
' Called from modApi* plot modules. Do not document as the external API.

Public Const PlotYellow As Long = 65535

Public Function PromptRange(ByVal PromptText As String, Optional ByVal Title As String = "Plots Charts") As Range
    Dim rng As Range
    Dim def As Range
    If TypeName(Selection) = "Range" Then Set def = Selection
    On Error Resume Next
    If def Is Nothing Then
        Set rng = Application.InputBox(Prompt:=PromptText, Title:=Title, Type:=8)
    Else
        Set rng = Application.InputBox(Prompt:=PromptText, Title:=Title, Default:=def.Address, Type:=8)
    End If
    On Error GoTo 0
    Set PromptRange = rng
End Function

' Empty if the user cancelled.
Public Function PromptNumber(ByVal PromptText As String, ByVal DefaultValue As Variant, Optional ByVal Title As String = "Plots Charts") As Variant
    Dim resp As Variant
    resp = Application.InputBox(Prompt:=PromptText, Title:=Title, Default:=DefaultValue, Type:=1)
    If VarType(resp) = vbBoolean Then
        PromptNumber = Empty
    Else
        PromptNumber = resp
    End If
End Function

Public Function RangeIsAllNumeric(ByVal rng As Range) As Boolean
    If rng Is Nothing Then Exit Function
    RangeIsAllNumeric = (Application.WorksheetFunction.Count(rng) = rng.Cells.Count)
End Function

Public Function RangeHasNegative(ByVal rng As Range) As Boolean
    Dim c As Range
    If rng Is Nothing Then Exit Function
    For Each c In rng.Cells
        If IsNumeric(c.Value) Then
            If Len(CStr(c.Value)) > 0 Then
                If CDbl(c.Value) < 0 Then
                    RangeHasNegative = True
                    Exit Function
                End If
            End If
        End If
    Next c
End Function

Public Function HeaderRowLooksLikeHeader(ByVal rng As Range) As Boolean
    Dim c As Range
    If rng Is Nothing Then Exit Function
    If rng.Rows.Count < 1 Then Exit Function
    For Each c In rng.Rows(1).Cells
        If Not IsNumeric(c.Value) Or Len(Trim$(CStr(c.Value))) = 0 Then
            HeaderRowLooksLikeHeader = True
            Exit Function
        End If
        If VarType(c.Value) = vbString Then
            HeaderRowLooksLikeHeader = True
            Exit Function
        End If
    Next c
End Function

Public Sub FlattenToColumn(ByVal src As Range, ByVal destTop As Range)
    Dim c As Range
    Dim i As Long
    i = 0
    For Each c In src.Cells
        destTop.Offset(i, 0).Value = c.Value
        i = i + 1
    Next c
End Sub

' Start row/col of the selection; bump off row 1 so a header can sit above the block.
Public Sub AnchorAtSelection(ByRef startRow As Long, ByRef startCol As Long)
    If TypeName(Selection) <> "Range" Then
        startRow = 2
        startCol = 1
        Exit Sub
    End If
    startRow = Selection.Rows(1).Row
    If startRow < 2 Then startRow = 2
    startCol = Selection.Columns(1).Column
End Sub

Public Sub HighlightYellow(ByVal rng As Range)
    rng.Interior.Color = PlotYellow
End Sub

Public Sub WriteHeaderRow(ByVal dest As Range, ByVal headers As Variant)
    Dim n As Long
    n = UBound(headers) - LBound(headers) + 1
    dest.Resize(1, n).Value = headers
    With dest.Resize(1, n)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
    End With
End Sub

' Write a 2D block. Strings starting with "=" are R1C1 formulas.
Public Sub WriteBlock(ByVal ws As Worksheet, ByVal startRow As Long, ByVal startCol As Long, ByRef data As Variant)
    Dim r As Long
    Dim c As Long
    Dim r0 As Long
    Dim c0 As Long
    Dim v As Variant
    Dim dest As Range
    If Not IsArray(data) Then
        ws.Cells(startRow, startCol).Value = data
        Exit Sub
    End If
    r0 = LBound(data, 1)
    On Error Resume Next
    c0 = LBound(data, 2)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        For r = LBound(data) To UBound(data)
            Set dest = ws.Cells(startRow + r - LBound(data), startCol)
            v = data(r)
            If IsFormulaText(v) Then dest.FormulaR1C1 = CStr(v) Else dest.Value = v
        Next r
        Exit Sub
    End If
    On Error GoTo 0
    For r = r0 To UBound(data, 1)
        For c = c0 To UBound(data, 2)
            Set dest = ws.Cells(startRow + r - r0, startCol + c - c0)
            v = data(r, c)
            If IsFormulaText(v) Then
                dest.FormulaR1C1 = CStr(v)
            ElseIf Not IsEmpty(v) Then
                dest.Value = v
            End If
        Next c
    Next r
End Sub

Public Function IsFormulaText(ByVal v As Variant) As Boolean
    If VarType(v) = vbString Then
        If Left$(v, 1) = "=" Then IsFormulaText = True
    End If
End Function

Public Function AddChartAt(ByVal ws As Worksheet, ByVal anchor As Range, ByVal w As Double, ByVal h As Double) As ChartObject
    Set AddChartAt = ws.ChartObjects.Add(Left:=anchor.Left, Top:=anchor.Top, Width:=w, Height:=h)
End Function

Public Sub StyleValueAxes(ByVal ch As Chart, ByVal xTitle As String, ByVal yTitle As String, _
                          Optional ByVal xFormat As String = "0", Optional ByVal yFormat As String = "0.00", _
                          Optional ByVal yMinZero As Boolean = False)
    With ch
        .HasLegend = False
        .Axes(xlCategory).HasMajorGridlines = False
        .Axes(xlCategory).HasMinorGridlines = False
        If Len(xTitle) > 0 Then
            .Axes(xlCategory).HasTitle = True
            .Axes(xlCategory).AxisTitle.Text = xTitle
        End If
        .Axes(xlCategory).TickLabels.NumberFormat = xFormat
        .Axes(xlValue).HasMajorGridlines = False
        .Axes(xlValue).HasMinorGridlines = False
        If Len(yTitle) > 0 Then
            .Axes(xlValue).HasTitle = True
            .Axes(xlValue).AxisTitle.Text = yTitle
        End If
        .Axes(xlValue).TickLabels.NumberFormat = yFormat
        If yMinZero Then .Axes(xlValue).MinimumScale = 0
    End With
End Sub

Public Sub AddXySeries(ByVal ch As Chart, ByVal seriesName As String, ByVal xRng As Range, ByVal yRng As Range, _
                       Optional ByVal chtType As XlChartType = xlXYScatterSmooth, _
                       Optional ByVal lineRgb As Long = 255, Optional ByVal markers As Boolean = False)
    Dim s As Series
    Set s = ch.SeriesCollection.NewSeries
    With s
        .Name = seriesName
        .XValues = xRng
        .Values = yRng
        .ChartType = chtType
        If markers Then
            .MarkerStyle = xlMarkerStyleX
            .MarkerSize = 2
        Else
            .MarkerStyle = xlMarkerStyleNone
        End If
        On Error Resume Next
        .Format.Line.Weight = 0.5
        .Format.Line.ForeColor.RGB = lineRgb
        .Format.Line.DashStyle = msoLineSolid
        On Error GoTo 0
    End With
End Sub

Public Sub DeleteChartSheetIfExists(ByVal chartName As String)
    Dim sh As Object
    If Len(chartName) = 0 Then chartName = "Unknown"
    For Each sh In ActiveWorkbook.Charts
        If StrComp(sh.Name, chartName, vbTextCompare) = 0 Then
            Application.DisplayAlerts = False
            sh.Delete
            Application.DisplayAlerts = True
            Exit Sub
        End If
    Next sh
End Sub

Public Function OutputSheet(ByVal sheetName As String) As Worksheet
    Call modApiSheets.CreateOutputSheet(sheetName)
    Set OutputSheet = ActiveWorkbook.Worksheets(sheetName)
End Function
