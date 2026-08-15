Attribute VB_Name = "Custom_Menu13_Matrices1"
Option Explicit

' Personal Custom_Menu13_Matrices1: transpose, arithmetic, inverse, det, trace,
' Kronecker/Hadamard, scale/power, diagonal extract, outer/dot, vec/unvec.
' Results write one column to the right of the selection.

''' @Description: Transpose the selected matrix to the right of the selection.
Public Sub MatrixTranspose()
    Call Custom_Menu13_MatrixUtilities.UnaryOp("transpose")
End Sub

''' @Description: A + B. Prompts for the second matrix.
Public Sub MatrixAdd()
    Call Custom_Menu13_MatrixUtilities.BinaryOp("add")
End Sub

''' @Description: A - B. Prompts for the second matrix.
Public Sub MatrixSubtract()
    Call Custom_Menu13_MatrixUtilities.BinaryOp("sub")
End Sub

''' @Description: Multiply every entry of the selection by a scalar.
Public Sub MatrixScale()
    Dim k As Variant
    k = Custom_Menu13_MatrixUtilities.PromptNumber("Scale factor", 1)
    If IsEmpty(k) Then Exit Sub
    Call Custom_Menu13_MatrixUtilities.UnaryOp("scale", k)
End Sub

''' @Description: A * B (matrix product). Prompts for the second matrix.
Public Sub MatrixMultiply()
    Call Custom_Menu13_MatrixUtilities.BinaryOp("mul")
End Sub

''' @Description: Element-wise A * B. Prompts for the second matrix.
Public Sub MatrixHadamard()
    Call Custom_Menu13_MatrixUtilities.BinaryOp("hadamard")
End Sub

''' @Description: Kronecker product A ⊗ B. Prompts for the second matrix.
Public Sub MatrixKronecker()
    Call Custom_Menu13_MatrixUtilities.BinaryOp("kron")
End Sub

''' @Description: Outer product of two vectors. Selection is the first; prompts for the second.
Public Sub MatrixOuter()
    Call Custom_Menu13_MatrixUtilities.BinaryOp("outer")
End Sub

''' @Description: Dot product of two vectors. Selection is the first; prompts for the second.
Public Sub MatrixDot()
    Call Custom_Menu13_MatrixUtilities.BinaryOp("dot")
End Sub

''' @Description: Inverse of the selected square matrix, written to the right.
Public Sub MatrixInverse()
    Call Custom_Menu13_MatrixUtilities.UnaryOp("inverse")
End Sub

''' @Description: Integer matrix power A^p of the selected square matrix.
Public Sub MatrixPower()
    Dim p As Variant
    p = Custom_Menu13_MatrixUtilities.PromptNumber("Integer power (negative uses the inverse)", 2)
    If IsEmpty(p) Then Exit Sub
    If p <> Int(p) Then
        MsgBox "Power must be an integer.", vbExclamation, "Matrices"
        Exit Sub
    End If
    Call Custom_Menu13_MatrixUtilities.UnaryOp("power", p)
End Sub

''' @Description: Determinant of the selected square matrix (one cell to the right).
Public Sub MatrixDeterminant()
    Call Custom_Menu13_MatrixUtilities.UnaryOp("det")
End Sub

''' @Description: Trace of the selected square matrix (one cell to the right).
Public Sub MatrixTrace()
    Call Custom_Menu13_MatrixUtilities.UnaryOp("trace")
End Sub

''' @Description: Main diagonal of the selection as a column.
Public Sub MatrixDiagExtract()
    Call Custom_Menu13_MatrixUtilities.UnaryOp("diag")
End Sub

''' @Description: Column-major vec(A): stack columns into one column.
Public Sub MatrixVec()
    Call Custom_Menu13_MatrixUtilities.UnaryOp("vec")
End Sub

''' @Description: Unvec a column (column-major) back to a matrix. Prompts for row count.
Public Sub MatrixUnvec()
    Dim nRows As Variant
    nRows = Custom_Menu13_MatrixUtilities.PromptNumber("Number of rows in the result", 1)
    If IsEmpty(nRows) Then Exit Sub
    Call Custom_Menu13_MatrixUtilities.UnaryOp("unvec", nRows)
End Sub
