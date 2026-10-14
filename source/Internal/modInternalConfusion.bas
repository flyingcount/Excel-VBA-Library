Attribute VB_Name = "modInternalConfusion"
Option Explicit

' Internal: binary classifier counts and metrics (Personal Menu18 ConfusionMatrix).
' Positive class is Yes / 1. Called from modApiConfusion.

Public Type ConfusionCounts
    TP As Long
    FP As Long
    FN As Long
    TN As Long
End Type

Public Type ConfusionMetrics
    Total As Long
    Accuracy As Double
    MisclassificationRate As Double
    TruePositiveRate As Double
    FalsePositiveRate As Double
    TrueNegativeRate As Double
    FalseNegativeRate As Double
    Precision As Double
    Prevalence As Double
    FalseDiscoveryRate As Double
    FalseOmissionRate As Double
    NegativePredictiveValue As Double
    PositiveLikelihoodRatio As Double
    NegativeLikelihoodRatio As Double
    DiagnosticOddsRatio As Double
    Fmeasure As Double
    FBeta As Double
    Beta As Double
    MCC As Double
    Informedness As Double
    Markedness As Double
    ThreatScore As Double
End Type

Public Function PromptTwoColumns(ByVal PromptText As String) As Range
    Dim rng As Range
    If TypeName(Selection) <> "Range" Then
        MsgBox "Select a two-column range first (Predicted, Actual).", vbExclamation, "Confusion Matrix"
        Exit Function
    End If
    On Error Resume Next
    Set rng = Application.InputBox(Prompt:=PromptText, Title:="Confusion Matrix", Default:=Selection.Address, Type:=8)
    On Error GoTo 0
    Set PromptTwoColumns = rng
End Function

Public Function SafeDiv(ByVal numer As Double, ByVal denom As Double) As Double
    If denom = 0 Then
        SafeDiv = 0
    Else
        SafeDiv = numer / denom
    End If
End Function

Public Function ClassifyYesNo(ByVal predicted As Variant, ByVal actual As Variant) As String
    Dim p As String
    Dim a As String
    p = NormalizeYesNo(predicted)
    a = NormalizeYesNo(actual)
    If Len(p) = 0 Or Len(a) = 0 Then
        ClassifyYesNo = ""
        Exit Function
    End If
    ClassifyYesNo = ClassifyPair(p = "YES", a = "YES")
End Function

Public Function ClassifyOneZero(ByVal predicted As Variant, ByVal actual As Variant) As String
    Dim p As Long
    Dim a As Long
    If Not IsBinary01(predicted) Or Not IsBinary01(actual) Then
        ClassifyOneZero = ""
        Exit Function
    End If
    p = CLng(Val(CStr(predicted)))
    a = CLng(Val(CStr(actual)))
    ClassifyOneZero = ClassifyPair(p = 1, a = 1)
End Function

Public Function RangeIsYesNo(ByVal rng As Range) As Boolean
    Dim cell As Range
    For Each cell In rng.Cells
        If Len(NormalizeYesNo(cell.Value)) = 0 Then
            RangeIsYesNo = False
            Exit Function
        End If
    Next cell
    RangeIsYesNo = True
End Function

Public Function RangeIsOneZero(ByVal rng As Range) As Boolean
    Dim cell As Range
    For Each cell In rng.Cells
        If Not IsBinary01(cell.Value) Then
            RangeIsOneZero = False
            Exit Function
        End If
    Next cell
    RangeIsOneZero = True
End Function

Public Function CountFromLabels(ByRef labels As Variant) As ConfusionCounts
    Dim i As Long
    Dim c As ConfusionCounts
    For i = LBound(labels, 1) To UBound(labels, 1)
        Select Case CStr(labels(i, 1))
            Case "TP": c.TP = c.TP + 1
            Case "FP": c.FP = c.FP + 1
            Case "FN": c.FN = c.FN + 1
            Case "TN": c.TN = c.TN + 1
        End Select
    Next i
    CountFromLabels = c
End Function

Public Function MetricsFromCounts(ByRef c As ConfusionCounts) As ConfusionMetrics
    Dim m As ConfusionMetrics
    Dim tpr As Double
    Dim tnr As Double
    Dim prec As Double
    Dim npv As Double
    Dim fpr As Double
    Dim fnr As Double
    Dim lrPlus As Double
    Dim lrMinus As Double
    m.Total = c.TP + c.FP + c.FN + c.TN
    m.Beta = 1
    m.Accuracy = SafeDiv(CDbl(c.TP + c.TN), CDbl(m.Total))
    m.MisclassificationRate = SafeDiv(CDbl(c.FP + c.FN), CDbl(m.Total))
    tpr = SafeDiv(CDbl(c.TP), CDbl(c.TP + c.FN))
    fpr = SafeDiv(CDbl(c.FP), CDbl(c.FP + c.TN))
    tnr = SafeDiv(CDbl(c.TN), CDbl(c.TN + c.FP))
    fnr = SafeDiv(CDbl(c.FN), CDbl(c.FN + c.TP))
    prec = SafeDiv(CDbl(c.TP), CDbl(c.TP + c.FP))
    npv = SafeDiv(CDbl(c.TN), CDbl(c.TN + c.FN))
    m.TruePositiveRate = tpr
    m.FalsePositiveRate = fpr
    m.TrueNegativeRate = tnr
    m.FalseNegativeRate = fnr
    m.Precision = prec
    m.Prevalence = SafeDiv(CDbl(c.FN + c.TP), CDbl(m.Total))
    m.FalseDiscoveryRate = SafeDiv(CDbl(c.FP), CDbl(c.TP + c.FP))
    m.FalseOmissionRate = SafeDiv(CDbl(c.FN), CDbl(c.FN + c.TN))
    m.NegativePredictiveValue = npv
    lrPlus = SafeDiv(tpr, fpr)
    lrMinus = SafeDiv(fnr, tnr)
    m.PositiveLikelihoodRatio = lrPlus
    m.NegativeLikelihoodRatio = lrMinus
    m.DiagnosticOddsRatio = SafeDiv(lrPlus, lrMinus)
    m.Fmeasure = SafeDiv(2# * tpr * prec, tpr + prec)
    m.FBeta = SafeDiv((1 + m.Beta * m.Beta) * prec * tpr, (m.Beta * m.Beta) * prec + tpr)
    If (c.TP + c.FP) = 0 Or (c.TP + c.FN) = 0 Or (c.TN + c.FP) = 0 Or (c.TN + c.FN) = 0 Then
        m.MCC = 0
    Else
        m.MCC = (CDbl(c.TP) * CDbl(c.TN) - CDbl(c.FP) * CDbl(c.FN)) / _
            Sqr(CDbl(c.TP + c.FP) * CDbl(c.TP + c.FN) * CDbl(c.TN + c.FP) * CDbl(c.TN + c.FN))
    End If
    m.Informedness = tpr + tnr - 1
    m.Markedness = prec + npv - 1
    m.ThreatScore = SafeDiv(CDbl(c.TP), CDbl(c.TP + c.FN + c.FP))
    MetricsFromCounts = m
End Function

Public Function MetricsBlock(ByRef c As ConfusionCounts, ByRef m As ConfusionMetrics, _
                             ByVal posLabel As String, ByVal negLabel As String) As Variant
    Dim a As Variant
    ReDim a(1 To 25, 1 To 6)
    a(1, 2) = "Actual: " & posLabel
    a(1, 3) = "Actual: " & negLabel
    a(1, 4) = "Total"
    a(1, 5) = m.Prevalence
    a(1, 6) = m.Accuracy
    a(2, 1) = "Predicted: " & posLabel
    a(2, 2) = c.TP
    a(2, 3) = c.FP
    a(2, 4) = c.TP + c.FP
    a(2, 5) = m.Precision
    a(2, 6) = m.FalseDiscoveryRate
    a(3, 1) = "Predicted: " & negLabel
    a(3, 2) = c.FN
    a(3, 3) = c.TN
    a(3, 4) = c.FN + c.TN
    a(3, 5) = m.FalseOmissionRate
    a(3, 6) = m.NegativePredictiveValue
    a(4, 1) = "Total"
    a(4, 2) = c.FN + c.TP
    a(4, 3) = c.TN + c.FP
    a(4, 4) = m.Total
    a(5, 2) = m.TruePositiveRate
    a(5, 3) = m.FalsePositiveRate
    a(5, 5) = m.PositiveLikelihoodRatio
    a(5, 6) = m.DiagnosticOddsRatio
    a(6, 2) = m.FalseNegativeRate
    a(6, 3) = m.TrueNegativeRate
    a(6, 5) = m.NegativeLikelihoodRatio
    a(6, 6) = m.Fmeasure
    a(8, 1) = "Accuracy": a(8, 2) = m.Accuracy
    a(8, 3) = "Overall, how often is the classifier correct"
    a(9, 1) = "Mis-classification rate": a(9, 2) = m.MisclassificationRate
    a(9, 3) = "Overall, how often is it wrong. AKA Error rate. Same as (1-Accuracy)"
    a(10, 1) = "True positive rate": a(10, 2) = m.TruePositiveRate
    a(10, 3) = "How often does it predict YES correctly. AKA Sensitivity or Recall"
    a(11, 1) = "False positive rate": a(11, 2) = m.FalsePositiveRate
    a(11, 3) = "How often does it predict YES when it is NO. Type I error"
    a(12, 1) = "True negative rate": a(12, 2) = m.TrueNegativeRate
    a(12, 3) = "How often does it predict NO correctly. AKA Specificity"
    a(13, 1) = "False negative rate": a(13, 2) = m.FalseNegativeRate
    a(13, 3) = "Type II error: predicts NO when it is actually YES"
    a(14, 1) = "Precision": a(14, 2) = m.Precision
    a(14, 3) = "When it predicts YES, how often is it correct"
    a(15, 1) = "Prevalence": a(15, 2) = m.Prevalence
    a(15, 3) = "How often does the YES condition actually occur"
    a(16, 1) = "Positive likelihood ratio, LR+": a(16, 2) = m.PositiveLikelihoodRatio
    a(16, 3) = ">1 strong positive predictor; <1 strong negative; =1 uninformative"
    a(17, 1) = "Negative likelihood ratio, LR-": a(17, 2) = m.NegativeLikelihoodRatio
    a(17, 3) = ">1 strong positive predictor; <1 strong negative; =1 uninformative"
    a(18, 1) = "Diagnostic Odds ratio": a(18, 2) = m.DiagnosticOddsRatio
    a(18, 3) = ">1 useful model; <1 invert the model; =1 no information"
    a(19, 1) = "F1-measure": a(19, 2) = m.Fmeasure
    a(19, 3) = "Harmonic mean of recall and precision"
    a(20, 1) = "Beta": a(20, 2) = m.Beta
    a(20, 3) = "0.5 weights precision; 2 weights recall. Default = 1 (F1)"
    a(21, 1) = "F Beta measure": a(21, 2) = m.FBeta
    a(21, 3) = "(1+b^2)*precision*recall / (b^2*precision + recall) with Beta above"
    a(22, 1) = "Matthews correlation coefficient (MCC)": a(22, 2) = m.MCC
    a(22, 3) = "Worst = -1, best = 1. High only if both classes are predicted well"
    a(23, 1) = "Informedness or Bookmaker Informedness (BM)": a(23, 2) = m.Informedness
    a(23, 3) = "+1 all correct; -1 all inverted. Counterpart to recall"
    a(24, 1) = "Markedness (MK)": a(24, 2) = m.Markedness
    a(24, 3) = "+1 every prediction correct; -1 every prediction incorrect"
    a(25, 1) = "Threat score (TS) or Critical Success Index (CSI)": a(25, 2) = m.ThreatScore
    MetricsBlock = a
End Function

Public Sub FormatConfusionSheet(ByVal ws As Worksheet, ByVal startRow As Long, ByVal startCol As Long)
    Dim anchor As Range
    Set anchor = ws.Cells(startRow, startCol)
    With anchor
        .Value = "Confusion Matrix"
        .Font.Bold = True
        .WrapText = True
        With .Offset(0, 1).Resize(1, 3)
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
            .WrapText = True
        End With
        With .Offset(1, 0).Resize(3, 1)
            .HorizontalAlignment = xlRight
            .Font.Bold = True
            .WrapText = True
        End With
        With .Offset(7, 0).Resize(18, 1)
            .WrapText = True
            .RowHeight = 30
        End With
        .Offset(1, 1).Interior.ColorIndex = 4
        .Offset(1, 1).NumberFormat = """TP = "" 0"
        .Offset(2, 2).Interior.ColorIndex = 4
        .Offset(2, 2).NumberFormat = """TN = "" 0"
        .Offset(2, 1).Interior.ColorIndex = 3
        .Offset(2, 1).NumberFormat = """FN = "" 0"
        .Offset(1, 2).Interior.ColorIndex = 3
        .Offset(1, 2).NumberFormat = """FP = "" 0"
        .Offset(1, 1).Resize(3, 3).HorizontalAlignment = xlCenter
        Call ShadeMetric(.Offset(0, 4), """Prevalence = """ & Chr(10) & " 0.000")
        Call ShadeMetric(.Offset(0, 5), """Accuracy = """ & Chr(10) & " 0.000")
        Call ShadeMetric(.Offset(1, 4), """Precision = """ & Chr(10) & " 0.000")
        Call ShadeMetric(.Offset(1, 5), """False discovery rate = """ & Chr(10) & " 0.000")
        Call ShadeMetric(.Offset(2, 4), """False omission rate = """ & Chr(10) & " 0.000")
        Call ShadeMetric(.Offset(2, 5), """Negative predictive value = """ & Chr(10) & " 0.000")
        Call ShadeMetric(.Offset(4, 1), """Recall = """ & Chr(10) & " 0.000")
        Call ShadeMetric(.Offset(4, 2), """False positive rate = """ & Chr(10) & " 0.000")
        Call ShadeMetric(.Offset(4, 4), """LR+ = """ & Chr(10) & " 0.000")
        Call ShadeMetric(.Offset(4, 5), """Diagnostic odds ratio = """ & Chr(10) & " 0.000")
        Call ShadeMetric(.Offset(5, 1), """False negative rate = """ & Chr(10) & " 0.000")
        Call ShadeMetric(.Offset(5, 2), """Specificity = """ & Chr(10) & " 0.000")
        Call ShadeMetric(.Offset(5, 4), """LR- = """ & Chr(10) & " 0.000")
        Call ShadeMetric(.Offset(5, 5), """F1 score = """ & Chr(10) & " 0.000")
        .Offset(1, 0).Resize(5, 1).RowHeight = 30
    End With
    ws.Range("A3:D3").WrapText = True
    ws.Columns("A:D").ColumnWidth = 6
    ws.Columns("A:D").Font.Size = 8
    ws.Columns("E").ColumnWidth = 1
    ws.Columns("F").ColumnWidth = 24
    ws.Columns("G:H").ColumnWidth = 25
    ws.Columns("I").ColumnWidth = 6
    ws.Columns("J:K").ColumnWidth = 30
End Sub

Private Sub ShadeMetric(ByVal rng As Range, ByVal fmt As String)
    rng.HorizontalAlignment = xlCenter
    rng.VerticalAlignment = xlCenter
    rng.NumberFormat = fmt
    rng.WrapText = True
    rng.Interior.ColorIndex = 15
End Sub

Private Function ClassifyPair(ByVal predPos As Boolean, ByVal actPos As Boolean) As String
    If predPos And actPos Then
        ClassifyPair = "TP"
    ElseIf predPos And Not actPos Then
        ClassifyPair = "FP"
    ElseIf (Not predPos) And actPos Then
        ClassifyPair = "FN"
    Else
        ClassifyPair = "TN"
    End If
End Function

Private Function NormalizeYesNo(ByVal v As Variant) As String
    Dim s As String
    If IsError(v) Then Exit Function
    s = UCase$(Trim$(CStr(v)))
    If s = "YES" Or s = "Y" Then
        NormalizeYesNo = "YES"
    ElseIf s = "NO" Or s = "N" Then
        NormalizeYesNo = "NO"
    End If
End Function

Private Function IsBinary01(ByVal v As Variant) As Boolean
    Dim s As String
    If IsError(v) Then Exit Function
    If IsNumeric(v) Then
        If CDbl(v) = 0 Or CDbl(v) = 1 Then
            IsBinary01 = True
            Exit Function
        End If
    End If
    s = Trim$(CStr(v))
    IsBinary01 = (s = "0" Or s = "1")
End Function
