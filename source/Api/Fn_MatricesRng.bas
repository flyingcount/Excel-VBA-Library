Attribute VB_Name = "Fn_MatricesRng"
Option Explicit

' Personal Fn_MatricesRng: range UDFs for Matrices1 operations.
' RegisterMatrixUdfs lists every Mat* name under Insert Function.

Private Const UdfCategory As String = "Excel VBA Lib"

Public Function MatTranspose(ByVal rng As Range) As Variant
    On Error GoTo Fail
    MatTranspose = modInternalMatrices.TransposeMatrix(modInternalMatrices.RangeToMatrix(rng))
    Exit Function
Fail:
    MatTranspose = CVErr(xlErrValue)
End Function

Public Function MatInv(ByVal rng As Range) As Variant
    On Error GoTo Fail
    MatInv = modInternalMatrices.Inverse(modInternalMatrices.RangeToMatrix(rng))
    Exit Function
Fail:
    MatInv = CVErr(xlErrNum)
End Function

Public Function MatDet(ByVal rng As Range) As Variant
    On Error GoTo Fail
    MatDet = modInternalMatrices.Determinant(modInternalMatrices.RangeToMatrix(rng))
    Exit Function
Fail:
    MatDet = CVErr(xlErrNum)
End Function

Public Function MatTrace(ByVal rng As Range) As Variant
    On Error GoTo Fail
    MatTrace = modInternalMatrices.TraceOf(modInternalMatrices.RangeToMatrix(rng))
    Exit Function
Fail:
    MatTrace = CVErr(xlErrValue)
End Function

Public Function MatMult(ByVal a As Range, ByVal b As Range) As Variant
    On Error GoTo Fail
    MatMult = modInternalMatrices.MatrixMultDefined( _
        modInternalMatrices.RangeToMatrix(a), modInternalMatrices.RangeToMatrix(b))
    Exit Function
Fail:
    MatMult = CVErr(xlErrValue)
End Function

Public Function MatrixMultDefined(ByVal a As Range, ByVal b As Range) As Variant
    MatrixMultDefined = MatMult(a, b)
End Function

Public Function MatAdd(ByVal a As Range, ByVal b As Range) As Variant
    On Error GoTo Fail
    MatAdd = modInternalMatrices.AddScaled( _
        modInternalMatrices.RangeToMatrix(a), modInternalMatrices.RangeToMatrix(b), 1)
    Exit Function
Fail:
    MatAdd = CVErr(xlErrValue)
End Function

Public Function MatSub(ByVal a As Range, ByVal b As Range) As Variant
    On Error GoTo Fail
    MatSub = modInternalMatrices.AddScaled( _
        modInternalMatrices.RangeToMatrix(a), modInternalMatrices.RangeToMatrix(b), -1)
    Exit Function
Fail:
    MatSub = CVErr(xlErrValue)
End Function

Public Function MatScale(ByVal rng As Range, ByVal k As Double) As Variant
    On Error GoTo Fail
    MatScale = modInternalMatrices.ScaleMat(modInternalMatrices.RangeToMatrix(rng), k)
    Exit Function
Fail:
    MatScale = CVErr(xlErrValue)
End Function

Public Function MatPow(ByVal rng As Range, ByVal p As Long) As Variant
    On Error GoTo Fail
    MatPow = modInternalMatrices.MatPower(modInternalMatrices.RangeToMatrix(rng), p)
    Exit Function
Fail:
    MatPow = CVErr(xlErrNum)
End Function

Public Function MatHadamard(ByVal a As Range, ByVal b As Range) As Variant
    On Error GoTo Fail
    MatHadamard = modInternalMatrices.Hadamard( _
        modInternalMatrices.RangeToMatrix(a), modInternalMatrices.RangeToMatrix(b))
    Exit Function
Fail:
    MatHadamard = CVErr(xlErrValue)
End Function

Public Function MatKronecker(ByVal a As Range, ByVal b As Range) As Variant
    On Error GoTo Fail
    MatKronecker = modInternalMatrices.Kronecker( _
        modInternalMatrices.RangeToMatrix(a), modInternalMatrices.RangeToMatrix(b))
    Exit Function
Fail:
    MatKronecker = CVErr(xlErrValue)
End Function

Public Function MatDiag(ByVal rng As Range) As Variant
    On Error GoTo Fail
    MatDiag = modInternalMatrices.ExtractDiag(modInternalMatrices.RangeToMatrix(rng))
    Exit Function
Fail:
    MatDiag = CVErr(xlErrValue)
End Function

Public Function MatOuter(ByVal a As Range, ByVal b As Range) As Variant
    On Error GoTo Fail
    MatOuter = modInternalMatrices.OuterProduct( _
        modInternalMatrices.RangeToMatrix(a), modInternalMatrices.RangeToMatrix(b))
    Exit Function
Fail:
    MatOuter = CVErr(xlErrValue)
End Function

Public Function MatDot(ByVal a As Range, ByVal b As Range) As Variant
    On Error GoTo Fail
    MatDot = modInternalMatrices.DotProduct( _
        modInternalMatrices.RangeToMatrix(a), modInternalMatrices.RangeToMatrix(b))
    Exit Function
Fail:
    MatDot = CVErr(xlErrValue)
End Function

Public Function MatVec(ByVal rng As Range) As Variant
    On Error GoTo Fail
    MatVec = modInternalMatrices.Vectorize(modInternalMatrices.RangeToMatrix(rng))
    Exit Function
Fail:
    MatVec = CVErr(xlErrValue)
End Function

Public Function MatUnvec(ByVal rng As Range, ByVal nRows As Long) As Variant
    On Error GoTo Fail
    MatUnvec = modInternalMatrices.Unvectorize(modInternalMatrices.RangeToMatrix(rng), nRows)
    Exit Function
Fail:
    MatUnvec = CVErr(xlErrValue)
End Function

Public Function MatToeplitz(ByVal rng As Range) As Variant
    On Error GoTo Fail
    MatToeplitz = modInternalMatrices.ToeplitzFromVector(modInternalMatrices.RangeToMatrix(rng))
    Exit Function
Fail:
    MatToeplitz = CVErr(xlErrValue)
End Function

Public Function MatVander(ByVal rng As Range, Optional ByVal cols As Variant) As Variant
    On Error GoTo Fail
    If IsMissing(cols) Then
        MatVander = modInternalMatrices.VandermondeFromVector(modInternalMatrices.RangeToMatrix(rng))
    Else
        MatVander = modInternalMatrices.VandermondeFromVector(modInternalMatrices.RangeToMatrix(rng), CLng(cols))
    End If
    Exit Function
Fail:
    MatVander = CVErr(xlErrValue)
End Function

Public Function MatCompanion(ByVal rng As Range) As Variant
    On Error GoTo Fail
    MatCompanion = modInternalMatrices.CompanionFromVector(modInternalMatrices.RangeToMatrix(rng))
    Exit Function
Fail:
    MatCompanion = CVErr(xlErrValue)
End Function

Public Sub RegisterMatrixUdfs()
    Call RegisterUdf("MatTranspose", "Transpose a numeric matrix.", "rng")
    Call RegisterUdf("MatInv", "Inverse of a square numeric matrix.", "rng")
    Call RegisterUdf("MatDet", "Determinant of a square numeric matrix.", "rng")
    Call RegisterUdf("MatTrace", "Trace of a square numeric matrix.", "rng")
    Call RegisterUdf("MatIdentity", "n-by-n identity matrix.", "n")
    Call RegisterUdf("MatZeros", "n-by-m zeros. Omit m for square.", "n,m")
    Call RegisterUdf("MatOnes", "n-by-m ones. Omit m for square.", "n,m")
    Call RegisterUdf("MatHilbert", "n-by-n Hilbert matrix.", "n")
    Call RegisterUdf("MatExchange", "n-by-n exchange (counter-identity) matrix.", "n")
    Call RegisterUdf("MatMult", "Matrix product A*B.", "a,b")
    Call RegisterUdf("MatrixMultDefined", "Matrix product of two ranges (Menu13 helper name).", "a,b")
    Call RegisterUdf("MatAdd", "Element-wise A+B.", "a,b")
    Call RegisterUdf("MatSub", "Element-wise A-B.", "a,b")
    Call RegisterUdf("MatScale", "Multiply every entry by k.", "rng,k")
    Call RegisterUdf("MatPow", "Integer matrix power A^p (negative uses the inverse).", "rng,p")
    Call RegisterUdf("MatHadamard", "Element-wise product.", "a,b")
    Call RegisterUdf("MatKronecker", "Kronecker product.", "a,b")
    Call RegisterUdf("MatChol", "Cholesky factor L of a symmetric positive-definite matrix.", "rng")
    Call RegisterUdf("MatQR", "Thin QR. Q on top, R immediately below.", "rng")
    Call RegisterUdf("MatEigen", "Symmetric eigen: eigenvectors then an eigenvalues column.", "rng")
    Call RegisterUdf("MatLU", "LU factorization. L on top, U immediately below.", "rng")
    Call RegisterUdf("MatNormF", "Frobenius norm.", "rng")
    Call RegisterUdf("MatNorm1", "Maximum absolute column sum.", "rng")
    Call RegisterUdf("MatNormInf", "Maximum absolute row sum.", "rng")
    Call RegisterUdf("MatRank", "Numerical rank.", "rng")
    Call RegisterUdf("MatDiag", "Main diagonal as a column.", "rng")
    Call RegisterUdf("MatOuter", "Outer product of two vectors.", "a,b")
    Call RegisterUdf("MatDot", "Dot product of two vectors.", "a,b")
    Call RegisterUdf("MatVec", "Column-major vec(A).", "rng")
    Call RegisterUdf("MatUnvec", "Unvec a column into nRows rows (column-major).", "rng,nRows")
    Call RegisterUdf("MatAdj", "Classical adjugate (det times inverse).", "rng")
    Call RegisterUdf("MatPInv", "Moore-Penrose pseudoinverse via normal equations.", "rng")
    Call RegisterUdf("MatCofactor", "Cofactor matrix (order at most 20).", "rng")
    Call RegisterUdf("MatMinor", "Minor after deleting row r and column c.", "rng,r,c")
    Call RegisterUdf("MatIsSymmetric", "TRUE if A equals A-transpose within 1e-9.", "rng")
    Call RegisterUdf("MatIsOrthogonal", "TRUE if A^T A is identity within 1e-9.", "rng")
    Call RegisterUdf("MatToeplitz", "Symmetric Toeplitz matrix from the first row or column.", "rng")
    Call RegisterUdf("MatVander", "Vandermonde matrix. Omit cols for a square matrix.", "rng,cols")
    Call RegisterUdf("MatCompanion", "Companion matrix; constant term first.", "rng")
End Sub

Private Sub RegisterUdf(ByVal procName As String, ByVal descr As String, ByVal argHelp As String)
    Dim qualified As String
    qualified = "'" & ThisWorkbook.Name & "'!" & procName
    On Error Resume Next
    Application.MacroOptions Macro:=qualified, Description:=descr, Category:=UdfCategory, ArgumentDescriptions:=Split(argHelp, ",")
    If Err.Number <> 0 Then
        Err.Clear
        Application.MacroOptions Macro:=qualified, Description:=descr, Category:=UdfCategory
    End If
    On Error GoTo 0
End Sub
