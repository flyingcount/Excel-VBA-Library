Attribute VB_Name = "Custom_Menu13_Cholesky"
Option Explicit

' Personal Custom_Menu13_Cholesky.

''' @Description: Lower-triangular Cholesky factor L of a symmetric positive-definite selection (A = L L^T).
Public Sub MatrixCholesky()
    Call Custom_Menu13_MatrixUtilities.UnaryOp("chol")
End Sub
