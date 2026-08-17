Attribute VB_Name = "modApiAnalysis"
Option Explicit

' Public API: Personal Menu18 Analysis (mean / stdev / correlation / variance-covariance).
' Other workbooks / the add-in menu should call these names only.

''' @Description: Write a 1 x p mean vector under the header, for a range that includes a header row.
''' @Example: CalculateMeanVector
Public Sub CalculateMeanVector()
    Call WriteVectorStat("Mean Vector", False, "mean")
End Sub

''' @Description: Write a 1 x p population standard-deviation vector (divide by n).
''' @Example: CalculateStandardDeviationPopulationVector
Public Sub CalculateStandardDeviationPopulationVector()
    Call WriteVectorStat("Standard Deviation Population Vector", True, "sdPop")
End Sub

''' @Description: Write a 1 x p sample standard-deviation vector (divide by n-1).
''' @Example: CalculateStandardDeviationSampleVector
Public Sub CalculateStandardDeviationSampleVector()
    Call WriteVectorStat("Standard Deviation Sample Vector", True, "sdSample")
End Sub

''' @Description: Write the p x p correlation matrix (population covariance divided by products of population SDs).
''' @Example: CorrelationMatrix
Public Sub CorrelationMatrix()
    Call WriteMatrixStat("Correlation Matrix", "correl")
End Sub

''' @Description: Write the p x p population variance-covariance matrix (divide by n).
''' @Example: CalculateVarianceCovarianceMatrix
Public Sub CalculateVarianceCovarianceMatrix()
    Call WriteMatrixStat("Variance-Covariance Matrix", "covPop")
End Sub

''' @Description: Write the p x p matrix of products of population standard deviations (sd_i * sd_j).
''' @Example: CalculateStdDevProductMatrixPopulation
Public Sub CalculateStdDevProductMatrixPopulation()
    Call WriteMatrixStat("Standard Deviation Product Row Matrix - Population", "sdProd")
End Sub

''' @Description: Sheet Cov and Correl showing D Sigma D = covariance, plus a direct covariance check.
''' @Example: ProveVarCovarAndCorrel
Public Sub ProveVarCovarAndCorrel()
    Const SheetName As String = "Cov and Correl"
    Dim src As Range
    Dim dest As Range
    Dim header As Variant
    Dim body As Variant
    Dim sdDiag As Variant
    Dim correl As Variant
    Dim covFromCorrel As Variant
    Dim covDirect As Variant
    Dim variance As Variant
    Dim n As Long
    Dim ws As Worksheet
    On Error GoTo EH
    Set src = modInternalAnalysis.PromptRange("Select the input range including the header.")
    If src Is Nothing Then Exit Sub
    If Not modInternalAnalysis.SplitHeaderBody(src, header, body) Then Exit Sub
    sdDiag = modInternalAnalysis.DiagonalFromVector(modInternalAnalysis.StdDevVector(body, False))
    correl = modInternalAnalysis.CorrelMatrix(body, True)
    covFromCorrel = Application.WorksheetFunction.MMult( _
        Application.WorksheetFunction.MMult(sdDiag, correl), sdDiag)
    covDirect = modInternalAnalysis.CovariancePopulation(body)
    variance = modInternalAnalysis.VarianceVector(body, False)
    n = UBound(sdDiag, 2)
    Call modInternalExcelApp.PushAppState
    Call modApiSheets.CreateOutputSheet(SheetName)
    Set ws = ActiveWorkbook.Worksheets(SheetName)
    ws.Range("A1").Value = "Variance-Covariance matrix and Correlation Matrix relationship"
    ws.Range("A1").Font.Bold = True
    ws.Range("A4").Value = "Diagonal population standard deviation matrix"
    Call modInternalAnalysis.PutArray(ws.Range("A5"), sdDiag)
    ws.Cells(4, n + 1).Value = "X"
    ws.Cells(4, n + 1).HorizontalAlignment = xlCenter
    ws.Cells(4, n + 2).Value = "Correlation Matrix"
    Call modInternalAnalysis.PutArray(ws.Cells(5, n + 2), correl)
    ws.Cells(4, 2 * n + 2).Value = "X"
    ws.Cells(4, 2 * n + 2).HorizontalAlignment = xlCenter
    ws.Cells(4, 2 * n + 3).Value = "Diagonal population standard deviation matrix"
    Call modInternalAnalysis.PutArray(ws.Cells(5, 2 * n + 3), sdDiag)
    ws.Cells(4, 3 * n + 3).Value = "="
    ws.Cells(4, 3 * n + 3).HorizontalAlignment = xlCenter
    ws.Cells(4, 3 * n + 4).Value = "Variance-Covariance Matrix"
    Call modInternalAnalysis.PutArray(ws.Cells(5, 3 * n + 4), covFromCorrel)
    ws.Cells(n + 6, 1).Value = "Variance"
    Call modInternalAnalysis.PutArray(ws.Cells(n + 7, 1), header)
    Call modInternalAnalysis.PutArray(ws.Cells(n + 8, 1), variance)
    ws.Cells(n + 6, 3 * n + 3).Value = "Check"
    Call modInternalAnalysis.PutArray(ws.Cells(n + 6, 3 * n + 4), covDirect)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("ProveVarCovarAndCorrel")
End Sub

Private Sub WriteVectorStat(ByVal Title As String, ByVal NeedTwoRows As Boolean, ByVal kind As String)
    Dim src As Range
    Dim dest As Range
    Dim header As Variant
    Dim body As Variant
    Dim vec As Variant
    On Error GoTo EH
    Set src = modInternalAnalysis.PromptRange("Select data range including the header")
    If src Is Nothing Then Exit Sub
    If Not modInternalAnalysis.SplitHeaderBody(src, header, body) Then Exit Sub
    If NeedTwoRows And UBound(body, 1) < 2 Then
        MsgBox "Need at least two data rows.", vbExclamation, "Analysis"
        Exit Sub
    End If
    Select Case kind
        Case "mean": vec = modInternalAnalysis.MeanVector(body)
        Case "sdPop": vec = modInternalAnalysis.StdDevVector(body, False)
        Case "sdSample": vec = modInternalAnalysis.StdDevVector(body, True)
        Case Else: Err.Raise 5, "WriteVectorStat", "Unknown kind"
    End Select
    Set dest = modInternalAnalysis.PromptRange("Select output location", src.Cells(1).Offset(src.Rows.Count + 1, 0))
    If dest Is Nothing Then Exit Sub
    Call modInternalAnalysis.WriteLabeledVector(dest.Cells(1, 1), Title, header, vec)
    Exit Sub
EH:
    Call modInternalError.RaiseCurrent("WriteVectorStat")
End Sub

Private Sub WriteMatrixStat(ByVal Title As String, ByVal kind As String)
    Dim src As Range
    Dim dest As Range
    Dim header As Variant
    Dim body As Variant
    Dim mat As Variant
    On Error GoTo EH
    Set src = modInternalAnalysis.PromptRange("Select range including header")
    If src Is Nothing Then Exit Sub
    If Not modInternalAnalysis.SplitHeaderBody(src, header, body) Then Exit Sub
    If UBound(body, 1) < 2 Then
        MsgBox "Need at least two data rows.", vbExclamation, "Analysis"
        Exit Sub
    End If
    Select Case kind
        Case "correl": mat = modInternalAnalysis.CorrelMatrix(body)
        Case "covPop": mat = modInternalAnalysis.CovariancePopulation(body)
        Case "sdProd": mat = modInternalAnalysis.StdDevProduct(modInternalAnalysis.StdDevVector(body, False))
        Case Else: Err.Raise 5, "WriteMatrixStat", "Unknown kind"
    End Select
    Set dest = modInternalAnalysis.PromptRange("Select output location", src.Cells(1).Offset(0, src.Columns.Count + 1))
    If dest Is Nothing Then Exit Sub
    Call modInternalAnalysis.WriteLabeledMatrix(dest.Cells(1, 1), Title, header, mat)
    Exit Sub
EH:
    Call modInternalError.RaiseCurrent("WriteMatrixStat")
End Sub
