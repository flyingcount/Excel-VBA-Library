Attribute VB_Name = "modApiConfusion"
Option Explicit

' Public API: Personal Menu18 confusion matrix (Yes/No, 1/0, and a formula template).

Private Const SheetName As String = "Confusion Matrix"
Private Const TemplateSheetName As String = "Confusion Matrix Template"
Private Const StartRow As Long = 3
Private Const StartCol As Long = 6

''' @Description: Confusion matrix from two columns of Yes/No (Predicted, Actual). Case-insensitive; Y/N allowed.
''' @Example: ConfusionMatrix
Public Sub ConfusionMatrix()
    Call RunConfusion("yesno")
End Sub

''' @Description: Confusion matrix from two columns of 1/0 (Predicted, Actual). 1 is the positive class.
''' @Example: ConfusionMatrixOnesAndZeros
Public Sub ConfusionMatrixOnesAndZeros()
    Call RunConfusion("onezero")
End Sub

''' @Description: Worksheet UDF: TP / FN / FP / TN from Yes/No predicted and actual.
Public Function ComparePredictedToActual(ByVal Predicted As String, ByVal Actual As String) As String
    Dim s As String
    s = modInternalConfusion.ClassifyYesNo(Predicted, Actual)
    If Len(s) = 0 Then
        ComparePredictedToActual = "Error: inputs must be YES or NO"
    Else
        ComparePredictedToActual = s
    End If
End Function

''' @Description: Worksheet UDF: TP / FN / FP / TN from 1/0 predicted and actual.
Public Function ComparePredictedToActualOneAndZerosOnly(ByVal Predicted As Integer, ByVal Actual As Integer) As String
    Dim s As String
    s = modInternalConfusion.ClassifyOneZero(Predicted, Actual)
    If Len(s) = 0 Then
        ComparePredictedToActualOneAndZerosOnly = "Error: inputs must be 1 or 0"
    Else
        ComparePredictedToActualOneAndZerosOnly = s
    End If
End Function

Public Function CheckRangeHasYesNoOnly(ByVal rng As Range) As Boolean
    CheckRangeHasYesNoOnly = modInternalConfusion.RangeIsYesNo(rng)
End Function

Public Function CheckRangeHasOnesAndZerosOnly(ByVal rng As Range) As Boolean
    CheckRangeHasOnesAndZerosOnly = modInternalConfusion.RangeIsOneZero(rng)
End Function

''' @Description: Protected template sheet with TP/FP/FN/TN input cells and live metric formulae.
''' @Example: ConfusionMatrixTemplate
Public Sub ConfusionMatrixTemplate()
    Dim ws As Worksheet
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    Call modApiSheets.CreateOutputSheet(TemplateSheetName)
    Set ws = ActiveWorkbook.Worksheets(TemplateSheetName)
    ws.Range("A1").Value = "Confusion Matrix"
    ws.Range("A1").Font.Bold = True
    ws.Range("C3").Value = "Actual: Yes"
    ws.Range("D3").Value = "Actual: No"
    ws.Range("E3").Value = "Total"
    ws.Range("A4").Value = "Predicted: Yes"
    ws.Range("A5").Value = "Predicted: No"
    ws.Range("A6").Value = "Total"
    ws.Range("C6").FormulaR1C1 = "=SUM(R[-2]C:R[-1]C)"
    ws.Range("D6").FormulaR1C1 = "=SUM(R[-2]C:R[-1]C)"
    ws.Range("E4").FormulaR1C1 = "=SUM(RC[-2]:RC[-1])"
    ws.Range("E5").FormulaR1C1 = "=SUM(RC[-2]:RC[-1])"
    ws.Range("E6").FormulaR1C1 = "=SUM(R[-2]C[-2]:R[-1]C[-1])"
    ws.Range("A8").Value = "Accuracy"
    ws.Range("C8").FormulaR1C1 = "=(R[-4]C+R[-3]C[1])/R[-2]C[2]"
    ws.Range("E8").Value = "Overall, how often is the classifier correct"
    ws.Range("A9").Value = "Mis-classification rate"
    ws.Range("C9").FormulaR1C1 = "=(1-R[-1]C)"
    ws.Range("E9").Value = "Overall, how often is it wrong. AKA Error rate. Same as (1-Accuracy)"
    ws.Range("A10").Value = "True positive rate"
    ws.Range("C10").FormulaR1C1 = "=R[-6]C/(R[-6]C+R[-5]C)"
    ws.Range("E10").Value = "How often does it predict YES correctly. AKA Sensitivity or Recall"
    ws.Range("A11").Value = "False positive rate"
    ws.Range("C11").FormulaR1C1 = "=R[-7]C[1]/(R[-7]C[1]+R[-6]C[1])"
    ws.Range("E11").Value = "How often does it predict YES when it is NO"
    ws.Range("A12").Value = "True negative rate"
    ws.Range("C12").FormulaR1C1 = "=R[-7]C[1]/(R[-7]C[1]+R[-8]C[1])"
    ws.Range("E12").Value = "How often does it predict NO correctly. AKA Specificity"
    ws.Range("A13").Value = "Precision"
    ws.Range("C13").FormulaR1C1 = "=R[-9]C/(R[-9]C[1]+R[-9]C)"
    ws.Range("E13").Value = "When it predicts YES, how often is it correct"
    ws.Range("A14").Value = "Prevalence"
    ws.Range("C14").FormulaR1C1 = "=R[-8]C/R[-8]C[2]"
    ws.Range("E14").Value = "How often does the YES condition actually occur"
    ws.Range("A15").Value = "F-measure"
    ws.Range("C15").FormulaR1C1 = "=(2*R[-5]C*R[-2]C)/(R[-5]C+R[-2]C)"
    ws.Range("E15").Value = "Weighted average of the true positive rate (recall) and precision"
    With ws.Range("C3:E3")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With
    With ws.Range("A4:A6")
        .Font.Bold = True
        .HorizontalAlignment = xlRight
        .WrapText = True
    End With
    Call UnlockCountCell(ws.Range("C4"), 4, """TP = "" 0")
    Call UnlockCountCell(ws.Range("D4"), 3, """FP = "" 0")
    Call UnlockCountCell(ws.Range("C5"), 3, """FN = "" 0")
    Call UnlockCountCell(ws.Range("D5"), 4, """TN = "" 0")
    ws.Columns("A").ColumnWidth = 25
    ws.Columns("B").ColumnWidth = 3
    ws.Columns("C:E").ColumnWidth = 12
    ws.Protect UserInterfaceOnly:=True
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("ConfusionMatrixTemplate")
End Sub

Private Sub RunConfusion(ByVal kind As String)
    Dim rng As Range
    Dim detail As Variant
    Dim labels As Variant
    Dim i As Long
    Dim n As Long
    Dim tag As String
    Dim c As ConfusionCounts
    Dim m As ConfusionMetrics
    Dim block As Variant
    Dim header As Variant
    Dim posLabel As String
    Dim negLabel As String
    Dim ws As Worksheet
    On Error GoTo EH
    Set rng = modInternalConfusion.PromptTwoColumns("Select two columns: Predicted, then Actual")
    If rng Is Nothing Then Exit Sub
    If rng.Columns.Count <> 2 Then
        MsgBox "Range must have exactly two columns (Predicted, Actual).", vbExclamation, "Confusion Matrix"
        Exit Sub
    End If
    If kind = "yesno" Then
        If Not modInternalConfusion.RangeIsYesNo(rng) Then
            MsgBox "Every cell must be Yes or No (Y/N allowed).", vbExclamation, "Confusion Matrix"
            Exit Sub
        End If
        posLabel = "Yes"
        negLabel = "No"
    Else
        If Not modInternalConfusion.RangeIsOneZero(rng) Then
            MsgBox "Every cell must be 1 or 0 (1 = positive class).", vbExclamation, "Confusion Matrix"
            Exit Sub
        End If
        posLabel = "1"
        negLabel = "0"
    End If
    n = rng.Rows.Count
    ReDim detail(1 To n, 1 To 4)
    ReDim labels(1 To n, 1 To 1)
    For i = 1 To n
        detail(i, 1) = rng.Cells(i, 1).Value
        detail(i, 2) = rng.Cells(i, 2).Value
        If kind = "yesno" Then
            tag = modInternalConfusion.ClassifyYesNo(detail(i, 1), detail(i, 2))
        Else
            tag = modInternalConfusion.ClassifyOneZero(detail(i, 1), detail(i, 2))
        End If
        detail(i, 3) = tag
        detail(i, 4) = (tag = "TP" Or tag = "TN")
        labels(i, 1) = tag
    Next i
    c = modInternalConfusion.CountFromLabels(labels)
    If c.TP + c.FP + c.FN + c.TN = 0 Then
        MsgBox "No classified rows.", vbExclamation, "Confusion Matrix"
        Exit Sub
    End If
    m = modInternalConfusion.MetricsFromCounts(c)
    block = modInternalConfusion.MetricsBlock(c, m, posLabel, negLabel)
    ReDim header(1 To 1, 1 To 4)
    header(1, 1) = "Predicted"
    header(1, 2) = "Actual"
    header(1, 3) = "Classification"
    header(1, 4) = "Prediction correct"
    Call modInternalExcelApp.PushAppState
    Call modApiSheets.CreateOutputSheet(SheetName)
    Set ws = ActiveWorkbook.Worksheets(SheetName)
    Call modApiArrays.WriteArrayToWorksheet(header, SheetName, StartRow, 1)
    Call modApiArrays.WriteArrayToWorksheet(detail, SheetName, StartRow + 1, 1)
    Call modApiArrays.WriteArrayToWorksheet(block, SheetName, StartRow, StartCol)
    Call modInternalConfusion.FormatConfusionSheet(ws, StartRow, StartCol)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("ConfusionMatrix")
End Sub

Private Sub UnlockCountCell(ByVal rng As Range, ByVal ColorIndex As Long, ByVal fmt As String)
    rng.Interior.ColorIndex = ColorIndex
    rng.HorizontalAlignment = xlCenter
    rng.VerticalAlignment = xlCenter
    rng.NumberFormat = fmt
    rng.Locked = False
End Sub
