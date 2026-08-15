Attribute VB_Name = "Custom_Menu13_CreateMatrices"
Option Explicit

' Personal Custom_Menu13_CreateMatrices. Writes a new matrix at the active cell
' (vector-based creates write to the right of the selection).

''' @Description: Write an n x n identity matrix at the active cell.
Public Sub MatrixCreateIdentity()
    Call Custom_Menu13_MatrixUtilities.WriteCreated("identity")
End Sub

''' @Description: Write an n x m zeros matrix at the active cell.
Public Sub MatrixCreateZeros()
    Call Custom_Menu13_MatrixUtilities.WriteCreated("zeros")
End Sub

''' @Description: Write an n x m ones matrix at the active cell.
Public Sub MatrixCreateOnes()
    Call Custom_Menu13_MatrixUtilities.WriteCreated("ones")
End Sub

''' @Description: Write a Hilbert matrix of order n at the active cell.
Public Sub MatrixCreateHilbert()
    Call Custom_Menu13_MatrixUtilities.WriteCreated("hilbert")
End Sub

''' @Description: Write a uniform random matrix at the active cell.
Public Sub MatrixCreateRandom()
    Call Custom_Menu13_MatrixUtilities.WriteCreated("random")
End Sub

''' @Description: Write the exchange (counter-identity) matrix of order n at the active cell.
Public Sub MatrixCreateExchange()
    Call Custom_Menu13_MatrixUtilities.WriteCreated("exchange")
End Sub

''' @Description: Build a diagonal matrix from the selected row or column vector.
Public Sub MatrixCreateDiagonal()
    Call Custom_Menu13_MatrixUtilities.FromSelectionCreate("diag")
End Sub

''' @Description: Symmetric Toeplitz matrix from the selected first row or column.
Public Sub MatrixCreateToeplitz()
    Call Custom_Menu13_MatrixUtilities.FromSelectionCreate("toeplitz")
End Sub

''' @Description: Vandermonde matrix from the selected vector (prompts for column count).
Public Sub MatrixCreateVandermonde()
    Call Custom_Menu13_MatrixUtilities.FromSelectionCreate("vander")
End Sub

''' @Description: Companion matrix of the selected coefficient vector (constant term first).
Public Sub MatrixCreateCompanion()
    Call Custom_Menu13_MatrixUtilities.FromSelectionCreate("companion")
End Sub
