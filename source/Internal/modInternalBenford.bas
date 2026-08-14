Attribute VB_Name = "modInternalBenford"
Option Explicit

' Internal: Benford table build, formatting, and charts.
' Called from modApiBenford. Do not document these as the external API.

Public Enum BenfordKind
    bkFirst = 1
    bkSecond = 2
    bkThird = 3
    bkTwo = 4
    bkThree = 5
    bkLastTwo = 6
End Enum

Private Const cnst_OutputRowOffset As Long = 4
Private Const cnst_OutputColumnOffset As Long = 1

Public Function SecondDigitProbability(ByVal SecondDigit As Long) As Double
    Dim counter As Long
    Dim twoDigits As Long
    Dim acc As Double

    For counter = 10 To 90 Step 10
        twoDigits = counter + SecondDigit
        acc = acc + WorksheetFunction.Log10(1 + (1 / twoDigits))
    Next
    SecondDigitProbability = acc
End Function

Public Function ThirdDigitProbability(ByVal ThirdDigit As Long) As Double
    Dim counter As Long
    Dim threeDigits As Long
    Dim acc As Double

    For counter = 100 To 990 Step 10
        threeDigits = counter + ThirdDigit
        acc = acc + WorksheetFunction.Log10(1 + (1 / threeDigits))
    Next
    ThirdDigitProbability = acc
End Function

Public Sub RunAnalysis(ByVal kind As BenfordKind, ByVal InputRange As Range)
    Dim suffix As String

    Application.StatusBar = "Processing input data"
    Call CreateBenfordTable(kind, InputRange)

    Application.StatusBar = "Plotting charts"
    suffix = KindSuffix(kind)
    Call BenfordFrequencyPlot(suffix, KindPlotTitle(kind))
    Call BenfordScatterPlot(suffix)
    If kind = bkFirst Then
        Call BenfordResidualPlot(Range("BenfordBins1stD"), Range("BenfordResidual"))
    End If
End Sub

' =============================================================================
' Table build
' =============================================================================

Private Sub CreateBenfordTable(ByVal kind As BenfordKind, ByVal rng_InputRange As Range)
    Dim ws As Worksheet
    Dim rngBody As Range
    Dim rngOutput As Range
    Dim sheetName As String
    Dim suffix As String
    Dim digitName As String
    Dim totalRow As Long
    Dim n As Long
    Dim digitStart As Long
    Dim digitEnd As Long
    Dim digit As Long
    Dim destRow As Long

    sheetName = KindSheetName(kind)
    suffix = KindSuffix(kind)
    digitName = "BenfordDigit" & suffix
    totalRow = KindTotalRow(kind)
    n = rng_InputRange.Cells.Count

    Call modInternalNamedRanges.CreateNamedRange("BenfordData" & suffix, rng_InputRange)
    Call PrepareOutputSheet(sheetName)
    Set ws = Sheets(sheetName)

    ws.Range("A4:B4") = Array("Data", KindDigitHeader(kind))
    Call WriteInputAndDigits(kind, ws, rng_InputRange, n)

    Set rngBody = ws.Cells(cnst_OutputRowOffset, cnst_OutputColumnOffset).CurrentRegion
    Call modInternalNamedRanges.CreateNamedRange("BenfordData" & suffix, BodyColumn(rngBody, 1))
    Call modInternalNamedRanges.CreateNamedRange(digitName, BodyColumn(rngBody, 2))

    Call WriteHeadings(kind, ws)
    Call WriteFooter(kind, ws, digitName)

    digitStart = KindDigitStart(kind)
    digitEnd = KindDigitEnd(kind)
    For digit = digitStart To digitEnd
        destRow = KindBinRow(kind, digit)
        Call WriteBinRow(kind, ws, destRow, digit, digitName, totalRow)
    Next

    Set rngOutput = ws.Cells(cnst_OutputRowOffset, cnst_OutputColumnOffset + 4).CurrentRegion
    Call modInternalNamedRanges.CreateNamedRange(KindOutputRangeName(kind), rngOutput)
    Call modInternalNamedRanges.CreateNamedRange("BenfordBins" & suffix, BodyColumn(rngOutput, 1))
    Call modInternalNamedRanges.CreateNamedRange("BenfordCount" & suffix, BodyColumn(rngOutput, 2))
    Call modInternalNamedRanges.CreateNamedRange("BenfordActualFrequency" & suffix, BodyColumn(rngOutput, 3))
    Call modInternalNamedRanges.CreateNamedRange("BenfordFrequency" & suffix, BodyColumn(rngOutput, 4))
    Call modInternalNamedRanges.CreateNamedRange("BenfordUpper" & suffix, BodyColumn(rngOutput, 7))
    Call modInternalNamedRanges.CreateNamedRange("BenfordLower" & suffix, BodyColumn(rngOutput, 8))
    Call modInternalNamedRanges.CreateNamedRange("BenfordZStat" & suffix, BodyColumn(rngOutput, 11))
    If kind = bkFirst Then
        Call modInternalNamedRanges.CreateNamedRange("BenfordResidual", BodyColumn(rngOutput, 12))
    End If

    Call FormatOutputSheet(kind, ws)
End Sub

Private Sub WriteInputAndDigits(ByVal kind As BenfordKind, ByVal ws As Worksheet, ByVal rng_InputRange As Range, ByVal n As Long)
    Dim i As Long
    Dim digitFormula As String

    digitFormula = DigitFormulaR1C1(kind)
    For i = 1 To n
        ws.Cells(i + cnst_OutputRowOffset, 1).Value = rng_InputRange.Cells(i).Value
        If kind = bkThree Then
            ws.Cells(i + cnst_OutputRowOffset, 2).Value = StripAndValidateData(rng_InputRange.Cells(i).Value)
        Else
            ws.Cells(i + cnst_OutputRowOffset, 2).FormulaR1C1 = digitFormula
        End If
        Application.StatusBar = "Processing input data record " & i & " of " & n
    Next
End Sub

Private Sub WriteHeadings(ByVal kind As BenfordKind, ByVal ws As Worksheet)
    Dim freqHeader As String
    freqHeader = "Benford frequency"
    If kind = bkLastTwo Then freqHeader = "Expected frequency"

    If kind = bkFirst Then
        ws.Range("D4:O4") = Array("Digit", "Count", "Actual frequency", freqHeader, _
            "Expected count", "Count difference", "Z Upper", "Z Lower", "1/2N", "Act-Exp", "Z-Statistic", "Residual")
    Else
        ws.Range("D4:N4") = Array("Digit", "Count", "Actual frequency", freqHeader, _
            "Expected count", "Count difference", "Z Upper", "Z Lower", "1/2N", "Act-Exp", "Z-Statistic")
    End If
End Sub

Private Sub WriteBinRow(ByVal kind As BenfordKind, ByVal ws As Worksheet, ByVal destRow As Long, ByVal digit As Long, ByVal digitName As String, ByVal totalRow As Long)
    Dim col As Long
    Dim tot As String

    col = cnst_OutputColumnOffset + 3
    tot = "R" & totalRow & "C5"

    ws.Cells(destRow, col).Value = digit
    ws.Cells(destRow, col + 1).FormulaR1C1 = "=COUNTIF(" & digitName & ",RC[-1])"
    ws.Cells(destRow, col + 2).FormulaR1C1 = "=RC[-1]/" & tot & " "

    Select Case kind
        Case bkFirst
            ws.Cells(destRow, col + 3).FormulaR1C1 = "=Log10(1 + (1 / RC[-3]))"
        Case bkSecond
            ws.Cells(destRow, col + 3).Value = SecondDigitProbability(digit)
        Case bkThird
            ws.Cells(destRow, col + 3).Value = ThirdDigitProbability(digit)
        Case bkTwo, bkThree
            ws.Cells(destRow, col + 3).FormulaR1C1 = "=Log10(RC[-3] + 1) - Log10(RC[-3])"
        Case bkLastTwo
            ws.Cells(destRow, col + 3).Value = 0.01
    End Select

    If kind = bkLastTwo Then
        ws.Cells(destRow, col + 4).FormulaR1C1 = "=RC[-1]*" & tot
    Else
        ws.Cells(destRow, col + 4).FormulaR1C1 = "=ROUND(RC[-1]*" & tot & ",0)"
    End If

    If kind = bkFirst Or kind = bkSecond Or kind = bkThird Then
        ws.Cells(destRow, col + 5).FormulaR1C1 = "=RC[-1]-RC[-4]"
    Else
        ws.Cells(destRow, col + 5).FormulaR1C1 = "=RC[-4]-RC[-1]"
    End If

    ws.Cells(destRow, col + 6).FormulaR1C1 = "=RC[-3]+R2C10*SQRT((RC[-3]*(1-RC[-3])/" & tot & "))+RC[2]"
    ws.Cells(destRow, col + 7).FormulaR1C1 = "=RC[-4]-R2C10*SQRT((RC[-4]*(1-RC[-4])/" & tot & "))-RC[1]"
    ws.Cells(destRow, col + 8).FormulaR1C1 = "=1/(2*" & tot & ")"
    ws.Cells(destRow, col + 9).FormulaR1C1 = "=ABS(RC[-7]-RC[-6])"
    ws.Cells(destRow, col + 10).FormulaR1C1 = _
        "=IF(RC[-2]<RC[-1],(RC[-1]-RC[-2])/(SQRT(RC[-7]*(1-RC[-7])/" & tot & ")),RC[-1]/(SQRT(RC[-7]*(1-RC[-7])/" & tot & ")))"

    If kind = bkFirst Then
        ws.Cells(destRow, col + 11).FormulaR1C1 = "=RC[-9]-RC[-8]"
    End If
End Sub

Private Sub WriteFooter(ByVal kind As BenfordKind, ByVal ws As Worksheet, ByVal digitName As String)
    Dim d As Long
    Dim totalRow As Long
    totalRow = KindTotalRow(kind)

    Select Case kind
        Case bkFirst
            With ws
                .Range("D15").FormulaR1C1 = "Total"
                .Range("E15").FormulaR1C1 = "=SUM(R[-10]C:R[-2]C)"
                .Range("H15").FormulaR1C1 = "=SUM(R[-10]C:R[-2]C)"
                .Range("D16").FormulaR1C1 = "0"
                .Range("E16").FormulaR1C1 = "=COUNTIF(" & digitName & ",RC[-1])"
                .Range("D17").FormulaR1C1 = "Not a number"
                .Range("E17").FormulaR1C1 = "=COUNTIF(" & digitName & ",""Not a number"")"
                .Range("D18").FormulaR1C1 = "Grand total"
                .Range("E18").FormulaR1C1 = "=SUM(R[-3]C:R[-1]C)"
                .Range("D19").FormulaR1C1 = "Data count"
                .Range("E19").FormulaR1C1 = "=COUNTA(" & digitName & ")"
                .Range("D20").FormulaR1C1 = "Missed records"
                .Range("E20").FormulaR1C1 = "=R[-2]C-R[-1]C"
            End With
        Case bkSecond, bkThird
            With ws
                .Range("D16").FormulaR1C1 = "Total"
                .Range("E16").FormulaR1C1 = "=SUM(R[-11]C:R[-2]C)"
                .Range("H16").FormulaR1C1 = "=SUM(R[-11]C:R[-2]C)"
                .Range("D17").FormulaR1C1 = "Not a number"
                .Range("E17").FormulaR1C1 = "=COUNTIF(" & digitName & ",""Not a number"")"
                .Range("D18").FormulaR1C1 = "Grand total"
                .Range("E18").FormulaR1C1 = "=SUM(R[-3]C:R[-1]C)"
                .Range("D19").FormulaR1C1 = "Data count"
                .Range("E19").FormulaR1C1 = "=COUNTA(" & digitName & ")"
                .Range("D20").FormulaR1C1 = "Missed records"
                .Range("E20").FormulaR1C1 = "=R[-2]C-R[-1]C"
            End With
        Case bkTwo
            With ws
                .Range("D96").FormulaR1C1 = "Sub-total"
                .Range("E96").FormulaR1C1 = "=SUM(R[-91]C:R[-2]C)"
                .Range("H96").FormulaR1C1 = "=SUM(R[-91]C:R[-2]C)"
                For d = 0 To 9
                    .Cells(97 + d, 4).Value = d
                    .Cells(97 + d, 5).FormulaR1C1 = "=COUNTIF(" & digitName & ",RC[-1])"
                Next
                .Range("D107").FormulaR1C1 = "Not a number"
                .Range("E107").FormulaR1C1 = "=COUNTIF(" & digitName & ",""Not a number"")"
                .Range("D108").FormulaR1C1 = "Grand total"
                .Range("E108").FormulaR1C1 = "=SUM(R[-12]C:R[-1]C)"
                .Range("D109").FormulaR1C1 = "Data count"
                .Range("E109").FormulaR1C1 = "=COUNTA(" & digitName & ")"
                .Range("D110").FormulaR1C1 = "Missed records"
                .Range("E110").FormulaR1C1 = "=R[-2]C-R[-1]C"
            End With
        Case bkThree
            With ws
                .Range("D906").FormulaR1C1 = "Sub-total"
                .Range("E906").FormulaR1C1 = "=SUM(R[-901]C:R[-2]C)"
                .Range("H906").FormulaR1C1 = "=SUM(R[-901]C:R[-2]C)"
                .Range("D907").FormulaR1C1 = "Excluded"
                .Range("E907").FormulaR1C1 = "=COUNTIF(" & digitName & ",""Excluded"")"
                .Range("D908").FormulaR1C1 = "Too short"
                .Range("E908").FormulaR1C1 = "=COUNTIF(" & digitName & ",""Too short"")"
                .Range("D909").FormulaR1C1 = "Grand total"
                .Range("E909").FormulaR1C1 = "=SUM(R[-3]C:R[-1]C)"
                .Range("D910").FormulaR1C1 = "Data count"
                .Range("E910").FormulaR1C1 = "=COUNTA(" & digitName & ")"
                .Range("D911").FormulaR1C1 = "Missed records"
                .Range("E911").FormulaR1C1 = "=R[-2]C-R[-1]C"
            End With
        Case bkLastTwo
            With ws
                .Range("D106").FormulaR1C1 = "Sub-total"
                .Range("E106").FormulaR1C1 = "=SUM(R[-101]C:R[-2]C)"
                .Range("D107").FormulaR1C1 = "Not a number"
                .Range("E107").FormulaR1C1 = "=COUNTIF(" & digitName & ",""Not a number"")"
                .Range("D108").FormulaR1C1 = "Grand total"
                .Range("E108").FormulaR1C1 = "=SUM(R[-2]C:R[-1]C)"
                .Range("D109").FormulaR1C1 = "Data count"
                .Range("E109").FormulaR1C1 = "=COUNTA(" & digitName & ")"
                .Range("D110").FormulaR1C1 = "Missed records"
                .Range("E110").FormulaR1C1 = "=R[-2]C-R[-1]C"
            End With
    End Select
End Sub

Private Sub FormatOutputSheet(ByVal kind As BenfordKind, ByVal ws As Worksheet)
    Dim totalRow As Long
    Dim grandTotalRow As Long
    Dim zName As String

    totalRow = KindTotalRow(kind)
    grandTotalRow = KindGrandTotalRow(kind)
    zName = "BenfordZStat" & KindSuffix(kind)

    With ws
        .Range("D1").Value = "Benford Data and Plot"
        .Range("D2").Value = "Last updated " & Now
        .Range("I1").Value = "Confidence interval"
        .Range("J1").Value = 0.95
        .Range("K1").Value = "<- change value"
        .Range("I2").Value = "Z score"
        .Range("J2").Value = "=NORM.S.INV((1+R[-1]C)/2)"

        .Columns("A").AutoFit
        .Columns("B").ColumnWidth = 13
        .Columns("C").ColumnWidth = 1
        .Columns("D").ColumnWidth = 13
        .Columns("E:H").ColumnWidth = 9
        .Columns("I").ColumnWidth = 10
        .Columns("J:N").ColumnWidth = 9
        If kind = bkFirst Then .Columns("O").ColumnWidth = 9

        If kind = bkLastTwo Then
            .Columns("A").NumberFormat = "General"
            .Columns("D").NumberFormat = "00"
        Else
            .Columns("A").NumberFormat = "#,##0.00"
        End If
        .Columns("E").NumberFormat = "#,##0"
        .Columns("F:H").NumberFormat = "0.0000"
        .Columns("H:I").NumberFormat = "0"
        .Columns("J:N").NumberFormat = "0.0000"
        If kind = bkFirst Then .Columns("O").NumberFormat = "0.0000"

        With .Rows(4)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = True
        End With

        With .Range("E" & totalRow).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Range("E" & grandTotalRow)
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlDouble
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThick
            End With
        End With

        With .Range(zName)
            If .FormatConditions.Count > 0 Then .FormatConditions.Delete
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=$J$2"
            With .FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .Color = 13551615
                .TintAndShade = 0
            End With
        End With

        .Range("A4").CurrentRegion.AutoFilter
        .Range("A1").Select
    End With
End Sub

' =============================================================================
' Kind metadata
' =============================================================================

Private Function KindSheetName(ByVal kind As BenfordKind) As String
    Select Case kind
        Case bkFirst: KindSheetName = "Bedford Analysis First digit"
        Case bkSecond: KindSheetName = "Bedford Analysis Second digit"
        Case bkThird: KindSheetName = "Bedford Analysis Third digit"
        Case bkTwo: KindSheetName = "Bedford Analysis 2 digit"
        Case bkThree: KindSheetName = "Bedford Analysis 3 digit"
        Case bkLastTwo: KindSheetName = "Bedford Analysis last 2 digits"
    End Select
End Function

Private Function KindSuffix(ByVal kind As BenfordKind) As String
    Select Case kind
        Case bkFirst: KindSuffix = "1stD"
        Case bkSecond: KindSuffix = "2ndD"
        Case bkThird: KindSuffix = "3rdD"
        Case bkTwo: KindSuffix = "2D"
        Case bkThree: KindSuffix = "3D"
        Case bkLastTwo: KindSuffix = "Last2D"
    End Select
End Function

Private Function KindDigitHeader(ByVal kind As BenfordKind) As String
    Select Case kind
        Case bkFirst: KindDigitHeader = "First digit"
        Case bkSecond: KindDigitHeader = "Second digit"
        Case bkThird: KindDigitHeader = "Third digit"
        Case bkTwo: KindDigitHeader = "First two digits"
        Case bkThree: KindDigitHeader = "First three digits"
        Case bkLastTwo: KindDigitHeader = "Last two digits"
    End Select
End Function

Private Function KindPlotTitle(ByVal kind As BenfordKind) As String
    Select Case kind
        Case bkFirst: KindPlotTitle = "Benford first digit frequency plot"
        Case bkSecond: KindPlotTitle = "Benford 2nd digit frequency plot"
        Case bkThird: KindPlotTitle = "Benford 3rd digit frequency plot"
        Case bkTwo: KindPlotTitle = "Benford first 2-digit frequency plot"
        Case bkThree: KindPlotTitle = "Benford first 3 digit frequency plot"
        Case bkLastTwo: KindPlotTitle = "Benford last 2-digit frequency plot"
    End Select
End Function

Private Function KindTotalRow(ByVal kind As BenfordKind) As Long
    Select Case kind
        Case bkFirst: KindTotalRow = 15
        Case bkSecond, bkThird: KindTotalRow = 16
        Case bkTwo: KindTotalRow = 96
        Case bkThree: KindTotalRow = 906
        Case bkLastTwo: KindTotalRow = 106
    End Select
End Function

Private Function KindGrandTotalRow(ByVal kind As BenfordKind) As Long
    Select Case kind
        Case bkFirst, bkSecond, bkThird: KindGrandTotalRow = 18
        Case bkTwo, bkLastTwo: KindGrandTotalRow = 108
        Case bkThree: KindGrandTotalRow = 909
    End Select
End Function

Private Function KindDigitStart(ByVal kind As BenfordKind) As Long
    Select Case kind
        Case bkFirst: KindDigitStart = 1
        Case bkSecond, bkThird, bkLastTwo: KindDigitStart = 0
        Case bkTwo: KindDigitStart = 10
        Case bkThree: KindDigitStart = 100
    End Select
End Function

Private Function KindDigitEnd(ByVal kind As BenfordKind) As Long
    Select Case kind
        Case bkFirst: KindDigitEnd = 9
        Case bkSecond, bkThird: KindDigitEnd = 9
        Case bkTwo: KindDigitEnd = 99
        Case bkThree: KindDigitEnd = 999
        Case bkLastTwo: KindDigitEnd = 99
    End Select
End Function

Private Function KindBinRow(ByVal kind As BenfordKind, ByVal digit As Long) As Long
    Select Case kind
        Case bkFirst, bkSecond, bkThird, bkLastTwo
            KindBinRow = digit + 1 + cnst_OutputRowOffset - KindDigitStart(kind)
        Case bkTwo
            KindBinRow = digit - 9 + cnst_OutputRowOffset
        Case bkThree
            KindBinRow = digit - 99 + cnst_OutputRowOffset
    End Select
End Function

Private Function KindOutputRangeName(ByVal kind As BenfordKind) As String
    Select Case kind
        Case bkThird: KindOutputRangeName = "BenfordOuptutData3rd"
        Case bkThree: KindOutputRangeName = "BenfordOutputData3D"
        Case Else: KindOutputRangeName = "BenfordOuptutData"
    End Select
End Function

Private Function DigitFormulaR1C1(ByVal kind As BenfordKind) As String
    Select Case kind
        Case bkFirst
            DigitFormulaR1C1 = "=IF(ISNUMBER(RC[-1]),RIGHT(LEFT(SUBSTITUTE(ABS(RC[-1]),""."","""",1),1),1),""Not a number"")"
        Case bkSecond
            DigitFormulaR1C1 = "=IF(ISNUMBER(RC[-1]),RIGHT(LEFT(SUBSTITUTE(ABS(RC[-1]),""."","""",1),2),1),""Not a number"")"
        Case bkThird
            DigitFormulaR1C1 = "=IF(ISNUMBER(RC[-1]),RIGHT(LEFT(SUBSTITUTE(ABS(RC[-1]),""."","""",1),3),1),""Not a number"")"
        Case bkTwo
            DigitFormulaR1C1 = "=IF(ISNUMBER(RC[-1]),LEFT(SUBSTITUTE(ABS(RC[-1]),""."","""",1),2),""Not a number"")"
        Case bkLastTwo
            DigitFormulaR1C1 = "=IFERROR(RIGHT(SUBSTITUTE(ABS(IF(ISBLANK(RC[-1]),""Not a number"",RC[-1])),""."","""",1),2),""Not a number"")"
        Case Else
            DigitFormulaR1C1 = vbNullString
    End Select
End Function

' =============================================================================
' Shared helpers
' =============================================================================

Private Function BodyColumn(ByVal rngRegion As Range, ByVal colIndex As Long) As Range
    Dim c As Range
    Set c = rngRegion.Columns(colIndex)
    Set BodyColumn = c.Offset(1, 0).Resize(c.Rows.Count - 1, c.Columns.Count)
End Function

Private Sub PrepareOutputSheet(ByVal str_OutputSheetName As String)
    Dim wSheet As Object
    Dim boo_WorksheetExists As Boolean

    boo_WorksheetExists = False
    For Each wSheet In ActiveWorkbook.Sheets
        If wSheet.Name = str_OutputSheetName Then
            boo_WorksheetExists = True
            Exit For
        End If
    Next wSheet

    If boo_WorksheetExists Then
        With Sheets(str_OutputSheetName)
            .Cells.Clear
            If .ChartObjects.Count > 0 Then .ChartObjects.Delete
        End With
    Else
        Sheets.Add After:=ActiveSheet
        ActiveSheet.Name = str_OutputSheetName
    End If
    Sheets(str_OutputSheetName).Select
End Sub

Private Function StripAndValidateData(ByVal var_Data As Variant) As Variant
    Dim var_Temp As Variant

    If IsNumeric(var_Data) Then
        var_Temp = Abs(var_Data)
        var_Temp = WorksheetFunction.Substitute(var_Temp, ".", "", 1)
        var_Temp = RemoveLeadingZeros(var_Temp)
        var_Temp = Left(var_Temp, 3)
        If Len(var_Temp) < 3 Then var_Temp = "Too short"
    Else
        var_Temp = "Excluded"
    End If
    StripAndValidateData = var_Temp
End Function

Private Function RemoveLeadingZeros(ByVal var_Data As Variant) As Variant
    Do While Left(var_Data, 1) = "0"
        var_Data = Mid(var_Data, 2)
    Loop
    RemoveLeadingZeros = var_Data
End Function

' =============================================================================
' Charts (one implementation, parameterized by named-range suffix)
' =============================================================================

Private Sub BenfordFrequencyPlot(ByVal suffix As String, ByVal chartTitle As String)
    Dim chrt As ChartObject
    Dim rngBins As Range
    Dim rngActual As Range
    Dim rngExpected As Range
    Dim rngUpper As Range
    Dim rngLower As Range

    Set rngBins = Range("BenfordBins" & suffix)
    Set rngActual = Range("BenfordActualFrequency" & suffix)
    Set rngExpected = Range("BenfordFrequency" & suffix)
    Set rngUpper = Range("BenfordUpper" & suffix)
    Set rngLower = Range("BenfordLower" & suffix)

    Set chrt = ActiveSheet.ChartObjects.Add( _
        Left:=Range("O1").Left, Width:=250, Top:=Range("O1").Top, Height:=230)

    With chrt.Chart
        .ChartType = xlLine
        .SetSourceData Source:=rngBins
        .HasTitle = True
        .ChartTitle.Text = chartTitle
        .ChartTitle.Font.Size = 12
        .HasLegend = True
        With .Legend
            .Position = xlLegendPositionBottom
            .Font.Size = 5
        End With

        With .Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.Text = "Digits"
            .TickLabels.NumberFormat = "0"
            .HasMajorGridlines = False
            .HasMinorGridlines = False
        End With
        With .Axes(xlValue, xlPrimary)
            .HasTitle = True
            .AxisTitle.Text = "Frequency"
            .TickLabels.NumberFormat = "0.00"
            .HasMajorGridlines = False
            .HasMinorGridlines = False
            .MinimumScale = 0
        End With

        With .FullSeriesCollection(1)
            .Name = "Actual frequency"
            .XValues = rngBins
            .Values = rngActual
            .ChartType = xlLineMarkers
            .Format.Line.Weight = 0.25
        End With
        .SeriesCollection.NewSeries
        With .FullSeriesCollection(2)
            .Name = "Benford frequency"
            .XValues = rngBins
            .Values = rngExpected
            .ChartType = xlLineMarkers
            .Format.Line.Weight = 0.25
        End With
        .SeriesCollection.NewSeries
        With .FullSeriesCollection(3)
            .Name = "Benford Z upper"
            .XValues = rngBins
            .Values = rngUpper
            .ChartType = xlLine
            With .Format.Line
                .Weight = 0.25
                .DashStyle = msoLineSysDash
            End With
        End With
        .SeriesCollection.NewSeries
        With .FullSeriesCollection(4)
            .Name = "Benford Z lower"
            .XValues = rngBins
            .Values = rngLower
            .ChartType = xlLine
            With .Format.Line
                .Weight = 0.25
                .DashStyle = msoLineSysDash
            End With
        End With
    End With
End Sub

Private Sub BenfordScatterPlot(ByVal suffix As String)
    Dim chrt As ChartObject
    Dim rngExpected As Range
    Dim rngActual As Range

    Set rngExpected = Range("BenfordFrequency" & suffix)
    Set rngActual = Range("BenfordActualFrequency" & suffix)

    Set chrt = ActiveSheet.ChartObjects.Add( _
        Left:=Range("O15").Left, Width:=250, Top:=Range("O15").Top, Height:=230)

    With chrt.Chart
        .HasTitle = True
        .ChartTitle.Text = "Actual vs Benford frequency plot"
        .ChartTitle.Font.Size = 12
        .HasLegend = False

        .SeriesCollection.NewSeries
        With .FullSeriesCollection(1)
            .Name = "Actual frequency"
            .XValues = rngExpected
            .Values = rngActual
            .ChartType = xlXYScatter
            .Trendlines.Add
            With .Trendlines(1)
                .DisplayEquation = True
                .DisplayRSquared = True
                .Format.Line.Weight = 0.25
            End With
        End With

        With .Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.Text = "Benford frequency"
            .TickLabels.NumberFormat = "0.00"
            .HasMajorGridlines = False
            .HasMinorGridlines = False
        End With
        With .Axes(xlValue, xlPrimary)
            .HasTitle = True
            .AxisTitle.Text = "Actual frequency"
            .TickLabels.NumberFormat = "0.00"
            .HasMajorGridlines = False
            .HasMinorGridlines = False
            .MinimumScale = 0
        End With
    End With
End Sub

Private Sub BenfordResidualPlot(ByVal XData As Range, ByVal Y1Data As Range)
    Dim chrt As ChartObject

    Set chrt = ActiveSheet.ChartObjects.Add( _
        Left:=Range("J15").Left, Width:=250, Top:=Range("J15").Top, Height:=230)

    With chrt.Chart
        .HasTitle = True
        .ChartTitle.Text = "Benford frequency residuals"
        .ChartTitle.Font.Size = 12
        .HasLegend = False

        With .Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.Text = "Benford digit"
            .TickLabels.NumberFormat = "0"
            .HasMajorGridlines = False
            .HasMinorGridlines = False
        End With
        With .Axes(xlValue, xlPrimary)
            .HasTitle = True
            .AxisTitle.Text = "Residual"
            .TickLabels.NumberFormat = "0.00"
            .HasMajorGridlines = False
            .HasMinorGridlines = False
        End With

        .SeriesCollection.NewSeries
        With .FullSeriesCollection(1)
            .Name = "Benford residuals"
            .XValues = XData
            .Values = Y1Data
            .ChartType = xlColumnClustered
            .Trendlines.Add
            With .Trendlines(1)
                .DisplayEquation = True
                .DisplayRSquared = True
                .Format.Line.Weight = 0.25
            End With
        End With
    End With
End Sub
