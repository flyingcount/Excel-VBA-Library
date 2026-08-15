Attribute VB_Name = "Custom_Menu13_EigenDecomp"
Option Explicit

' Personal Custom_Menu13_EigenDecomp. Jacobi method; selection must be symmetric.

''' @Description: Symmetric eigen-decomposition. Writes eigenvectors then an eigenvalues column.
Public Sub MatrixEigen()
    Call Custom_Menu13_MatrixUtilities.UnaryOp("eigen")
End Sub
