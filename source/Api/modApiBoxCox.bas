Attribute VB_Name = "modApiBoxCox"
Option Explicit

' Public API: Personal Custom_Menu5_BoxCox (menu called Boxcox; BoxCoxForm is not used).
' Numeric cells → sheet BoxCox with shifted data, λ grid, and log-likelihood. Prompts replace the unused form.

Private Const SheetName As String = "BoxCox"
Private Const DefaultLambdaStart As Double = -5
Private Const DefaultLambdaEnd As Double = 5
Private Const DefaultLambdaStep As Double = 0.1
Private Const WorkHeaderRow As Long = 13
Private Const WorkFirstCol As Long = 2
Private Const MaxFormulaCells As Long = 50000

''' @Description: Box-Cox λ grid with log-likelihood. Shift (Alpha) maps the smallest value to 1 when any input is ≤ 0. Transform uses shifted data once (Personal added Alpha twice).
''' @Example: BoxCox
Public Sub BoxCox()
    Dim src As Range
    Dim ws As Worksheet
    Dim n As Long
    Dim nTrials As Long
    Dim i As Long
    Dim c As Long
    Dim lambdaStart As Variant
    Dim lambdaEnd As Variant
    Dim lambdaStep As Variant
    Dim lambdas() As Double
    Dim vals() As Double
    Dim cell As Range
    Dim alpha As Double
    Dim minVal As Double
    Dim rngShifted As Range
    Dim rngLL As Range
    Dim maxLL As Double
    Dim bestCol As Long
    On Error GoTo EH
    Set src = modInternalAnalysis.PromptNumericRange( _
        "Select the numeric data to transform (any shape; cells are flattened).", _
        "Box-Cox", 1, 0, 1)
    If src Is Nothing Then Exit Sub
    n = src.Cells.Count
    If n < 2 Then
        MsgBox "Need at least two numeric values.", vbExclamation, "Box-Cox"
        Exit Sub
    End If
    lambdaStart = modInternalPlots.PromptNumber("Lambda start", DefaultLambdaStart, "Box-Cox")
    If IsEmpty(lambdaStart) Then Exit Sub
    lambdaEnd = modInternalPlots.PromptNumber("Lambda end", DefaultLambdaEnd, "Box-Cox")
    If IsEmpty(lambdaEnd) Then Exit Sub
    lambdaStep = modInternalPlots.PromptNumber("Lambda increment", DefaultLambdaStep, "Box-Cox")
    If IsEmpty(lambdaStep) Then Exit Sub
    If CDbl(lambdaStep) <= 0 Then
        MsgBox "Lambda increment must be greater than zero.", vbExclamation, "Box-Cox"
        Exit Sub
    End If
    If CDbl(lambdaEnd) < CDbl(lambdaStart) Then
        MsgBox "Lambda end must be greater than or equal to lambda start.", vbExclamation, "Box-Cox"
        Exit Sub
    End If
    nTrials = CLng(Round((CDbl(lambdaEnd) - CDbl(lambdaStart)) / CDbl(lambdaStep), 0)) + 1
    If nTrials < 1 Then
        MsgBox "No lambda trials to run.", vbExclamation, "Box-Cox"
        Exit Sub
    End If
    If CLng(n) * CLng(nTrials) > MaxFormulaCells Then
        If MsgBox("This will write " & Format$(CLng(n) * CLng(nTrials), "#,##0") & _
                  " transform formulae. Continue?", vbYesNo + vbQuestion, "Box-Cox") <> vbYes Then
            Exit Sub
        End If
    End If
    ReDim lambdas(1 To nTrials)
    For i = 1 To nTrials
        lambdas(i) = CDbl(lambdaStart) + (i - 1) * CDbl(lambdaStep)
        lambdas(i) = Round(lambdas(i), 10)
    Next i
    ReDim vals(1 To n, 1 To 2)
    i = 0
    minVal = CDbl(src.Cells(1, 1).Value)
    For Each cell In src.Cells
        i = i + 1
        vals(i, 1) = i
        vals(i, 2) = CDbl(cell.Value)
        If vals(i, 2) < minVal Then minVal = vals(i, 2)
    Next cell
    ' Strictly positive shifted data: 0 if already > 0, otherwise 1 − min so the smallest value becomes 1.
    If minVal > 0 Then
        alpha = 0
    Else
        alpha = 1 - minVal
    End If
    Call modInternalExcelApp.PushAppState
    Call modApiSheets.CreateOutputSheet(SheetName)
    Set ws = ActiveWorkbook.Worksheets(SheetName)
    ws.Range("B2").Value = "Box-Cox transformation summary"
    ws.Range("B2").Font.Bold = True
    ws.Range("B3").Value = "No. data points"
    ws.Range("C3").Value = n
    ws.Range("B4").Value = "Alpha"
    ws.Range("C4").Value = alpha
    ws.Range("D4").Value = "Shift added to every value so shifted data are > 0 (1 − min when any value is ≤ 0)."
    ws.Range("B5").Value = "Lambda start"
    ws.Range("C5").Value = CDbl(lambdaStart)
    ws.Range("B6").Value = "Lambda end"
    ws.Range("C6").Value = CDbl(lambdaEnd)
    ws.Range("B7").Value = "Lambda increment"
    ws.Range("C7").Value = CDbl(lambdaStep)
    ws.Range("B8").Value = "Number of lambda trials"
    ws.Range("C8").Value = nTrials
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "Alpha", ws.Range("C4"))
    ws.Range("B10").Value = "Log-likelihood"
    ws.Range("B11").Value = "Maximum LL"
    ws.Cells(WorkHeaderRow, WorkFirstCol).Value = "Ref"
    ws.Cells(WorkHeaderRow, WorkFirstCol + 1).Value = "Data"
    ws.Cells(WorkHeaderRow, WorkFirstCol + 2).Value = "Shifted data"
    For i = 1 To nTrials
        ws.Cells(WorkHeaderRow, WorkFirstCol + 2 + i).Value = lambdas(i)
    Next i
    ws.Range(ws.Cells(WorkHeaderRow, WorkFirstCol), ws.Cells(WorkHeaderRow, WorkFirstCol + 2 + nTrials)).Font.Bold = True
    ws.Cells(WorkHeaderRow + 1, WorkFirstCol).Resize(n, 2).Value = vals
    Set rngShifted = ws.Cells(WorkHeaderRow + 1, WorkFirstCol + 2).Resize(n, 1)
    rngShifted.FormulaR1C1 = "=RC[-1]+Alpha"
    Call modInternalNamedRanges.CreateSheetNamedRange(ws, "Shifted_Data", rngShifted)
    For i = 1 To nTrials
        c = WorkFirstCol + 2 + i
        ws.Cells(WorkHeaderRow + 1, c).Resize(n, 1).FormulaR1C1 = _
            "=IF(ABS(R" & WorkHeaderRow & "C)<1E-9,LN(RC" & (WorkFirstCol + 2) & ")," & _
            "((RC" & (WorkFirstCol + 2) & "^R" & WorkHeaderRow & "C)-1)/R" & WorkHeaderRow & "C)"
        ws.Cells(10, c).FormulaR1C1 = _
            "=-COUNT(Shifted_Data)/2*LN(VARP(R" & (WorkHeaderRow + 1) & "C:R" & (WorkHeaderRow + n) & "C))" & _
            "+(R" & WorkHeaderRow & "C-1)*SUMPRODUCT(LN(Shifted_Data))"
    Next i
    Set rngLL = ws.Cells(10, WorkFirstCol + 3).Resize(1, nTrials)
    ws.Range("D11").Formula = "=MAX(" & rngLL.Address(RowAbsolute:=True, ColumnAbsolute:=True) & ")"
    Application.Calculate
    If Not IsError(ws.Range("D11").Value) Then
        maxLL = ws.Range("D11").Value
        bestCol = WorkFirstCol + 3
        For i = 1 To nTrials
            c = WorkFirstCol + 2 + i
            If Not IsError(ws.Cells(10, c).Value) Then
                If ws.Cells(10, c).Value = maxLL Then
                    bestCol = c
                    Exit For
                End If
            End If
        Next i
        With ws.Cells(WorkHeaderRow, bestCol)
            .Interior.Color = 65535
            .Font.Bold = True
        End With
        ws.Cells(10, bestCol).Interior.Color = 65535
        ws.Range("B12").Value = "Best lambda"
        ws.Range("C12").Value = ws.Cells(WorkHeaderRow, bestCol).Value
        ws.Range("C12").Interior.Color = 65535
    End If
    ws.Columns("B").ColumnWidth = 24
    ws.Columns("C").ColumnWidth = 14
    ws.Columns("D").ColumnWidth = 14
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("BoxCox")
End Sub
