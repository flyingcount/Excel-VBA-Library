Attribute VB_Name = "Custom_Menu13_Matrices2"
Option Explicit

' Personal Custom_Menu13_Matrices2: solve, rank, norms, LU, adjugate, pseudoinverse,
' cofactor, minor, symmetry tests.

''' @Description: Solve A X = B for X. Selection is A; prompts for B.
Public Sub MatrixSolve()
    Call Custom_Menu13_MatrixUtilities.BinaryOp("solve")
End Sub

''' @Description: Numerical rank of the selection (one cell to the right).
Public Sub MatrixRank()
    Call Custom_Menu13_MatrixUtilities.UnaryOp("rank")
End Sub

''' @Description: Frobenius norm of the selection (one cell to the right).
Public Sub MatrixNorm()
    Call Custom_Menu13_MatrixUtilities.UnaryOp("norm")
End Sub

''' @Description: 1-norm (maximum absolute column sum) of the selection.
Public Sub MatrixNorm1()
    Call Custom_Menu13_MatrixUtilities.UnaryOp("norm1")
End Sub

''' @Description: Infinity-norm (maximum absolute row sum) of the selection.
Public Sub MatrixNormInf()
    Call Custom_Menu13_MatrixUtilities.UnaryOp("normInf")
End Sub

''' @Description: TRUE if the selection is square and symmetric.
Public Sub MatrixIsSymmetric()
    Call Custom_Menu13_MatrixUtilities.UnaryOp("sym")
End Sub

''' @Description: Classical adjugate (det(A) * inverse) of the selected square matrix.
Public Sub MatrixAdjugate()
    Call Custom_Menu13_MatrixUtilities.UnaryOp("adjugate")
End Sub

''' @Description: Moore-Penrose pseudoinverse via normal equations.
Public Sub MatrixPseudoInverse()
    Call Custom_Menu13_MatrixUtilities.UnaryOp("pinv")
End Sub

''' @Description: LU factorization. Writes L, then U immediately below L.
Public Sub MatrixLU()
    Call Custom_Menu13_MatrixUtilities.UnaryOp("lu")
End Sub

''' @Description: Cofactor matrix of the selection (order at most 20).
Public Sub MatrixCofactor()
    Call Custom_Menu13_MatrixUtilities.UnaryOp("cofactor")
End Sub

''' @Description: Minor obtained by deleting a prompted row and column.
Public Sub MatrixMinor()
    Dim rng As Range
    Dim r As Variant
    Dim c As Variant
    Dim a As Variant
    Dim out As Variant
    On Error GoTo EH
    Set rng = Custom_Menu13_MatrixUtilities.NeedSelection()
    If rng Is Nothing Then Exit Sub
    r = Custom_Menu13_MatrixUtilities.PromptNumber("Row to delete (1-based)", 1)
    If IsEmpty(r) Then Exit Sub
    c = Custom_Menu13_MatrixUtilities.PromptNumber("Column to delete (1-based)", 1)
    If IsEmpty(c) Then Exit Sub
    Call modInternalExcelApp.PushAppState
    a = modInternalMatrices.RangeToMatrix(rng)
    out = modInternalMatrices.MinorMatrix(a, CLng(r), CLng(c))
    Call modInternalMatrices.PutMatrix(modInternalMatrices.OutputOrigin(rng), out)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("MatrixMinor")
End Sub
