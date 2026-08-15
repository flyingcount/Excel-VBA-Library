Attribute VB_Name = "custom_Menu13_Unitary"
Option Explicit

' Personal custom_Menu13_Unitary: real QR / orthogonality (complex unitary not included).

''' @Description: Thin QR (modified Gram-Schmidt). Writes Q, then R immediately below Q.
Public Sub MatrixQR()
    Call Custom_Menu13_MatrixUtilities.UnaryOp("qr")
End Sub

''' @Description: TRUE if A^T A is the identity (orthogonal / real unitary).
Public Sub MatrixIsOrthogonal()
    Call Custom_Menu13_MatrixUtilities.UnaryOp("orth")
End Sub
