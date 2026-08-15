Attribute VB_Name = "modApiMatrices"
Option Explicit

' Public API: matrix create / operate / decompose (Personal Custom_Menu13_* / Fn_Matrices*).
' Menu macros fill to the right of the selection (create writes at the active cell).
' Add-in Subs do not appear in Alt+F8. Worksheet UDFs are registered by RegisterMatrixUdfs.

Private Const UdfCategory As String = "Excel VBA Lib"

''' @Description: Write an n x n identity matrix at the active cell.
Public Sub MatrixCreateIdentity()
    Call WriteCreated("identity")
End Sub

''' @Description: Write an n x m zeros matrix at the active cell.
Public Sub MatrixCreateZeros()
    Call WriteCreated("zeros")
End Sub

''' @Description: Write an n x m ones matrix at the active cell.
Public Sub MatrixCreateOnes()
    Call WriteCreated("ones")
End Sub

''' @Description: Write a Hilbert matrix of order n at the active cell.
Public Sub MatrixCreateHilbert()
    Call WriteCreated("hilbert")
End Sub

''' @Description: Write a uniform random matrix at the active cell.
Public Sub MatrixCreateRandom()
    Call WriteCreated("random")
End Sub

''' @Description: Write the exchange (counter-identity) matrix of order n at the active cell.
Public Sub MatrixCreateExchange()
    Call WriteCreated("exchange")
End Sub

''' @Description: Build a diagonal matrix from the selected row or column vector.
Public Sub MatrixCreateDiagonal()
    Call FromSelectionCreate("diag")
End Sub

''' @Description: Symmetric Toeplitz matrix from the selected first row or column.
Public Sub MatrixCreateToeplitz()
    Call FromSelectionCreate("toeplitz")
End Sub

''' @Description: Vandermonde matrix from the selected vector (prompts for column count).
Public Sub MatrixCreateVandermonde()
    Call FromSelectionCreate("vander")
End Sub

''' @Description: Transpose the selected matrix to the right of the selection.
Public Sub MatrixTranspose()
    Call UnaryOp("transpose")
End Sub

''' @Description: Inverse of the selected square matrix, written to the right.
Public Sub MatrixInverse()
    Call UnaryOp("inverse")
End Sub

''' @Description: Trace of the selected square matrix (one cell to the right).
Public Sub MatrixTrace()
    Call UnaryOp("trace")
End Sub

''' @Description: Determinant of the selected square matrix (one cell to the right).
Public Sub MatrixDeterminant()
    Call UnaryOp("det")
End Sub

''' @Description: Frobenius norm of the selection (one cell to the right).
Public Sub MatrixNorm()
    Call UnaryOp("norm")
End Sub

''' @Description: 1-norm (maximum absolute column sum) of the selection.
Public Sub MatrixNorm1()
    Call UnaryOp("norm1")
End Sub

''' @Description: Infinity-norm (maximum absolute row sum) of the selection.
Public Sub MatrixNormInf()
    Call UnaryOp("normInf")
End Sub

''' @Description: Numerical rank of the selection (one cell to the right).
Public Sub MatrixRank()
    Call UnaryOp("rank")
End Sub

''' @Description: TRUE if the selection is square and symmetric.
Public Sub MatrixIsSymmetric()
    Call UnaryOp("sym")
End Sub

''' @Description: TRUE if A^T A is the identity (orthogonal / unitary real).
Public Sub MatrixIsOrthogonal()
    Call UnaryOp("orth")
End Sub

''' @Description: Multiply every entry of the selection by a scalar.
Public Sub MatrixScale()
    Dim k As Variant
    k = PromptNumber("Scale factor", 1)
    If IsEmpty(k) Then Exit Sub
    Call UnaryOp("scale", k)
End Sub

''' @Description: Integer matrix power A^p of the selected square matrix.
Public Sub MatrixPower()
    Dim p As Variant
    p = PromptNumber("Integer power (negative uses the inverse)", 2)
    If IsEmpty(p) Then Exit Sub
    If p <> Int(p) Then
        MsgBox "Power must be an integer.", vbExclamation, "Matrices"
        Exit Sub
    End If
    Call UnaryOp("power", p)
End Sub

''' @Description: Main diagonal of the selection as a column.
Public Sub MatrixDiagExtract()
    Call UnaryOp("diag")
End Sub

''' @Description: Classical adjugate (det(A) * inverse) of the selected square matrix.
Public Sub MatrixAdjugate()
    Call UnaryOp("adjugate")
End Sub

''' @Description: Moore-Penrose pseudoinverse via normal equations.
Public Sub MatrixPseudoInverse()
    Call UnaryOp("pinv")
End Sub

''' @Description: A + B. Prompts for the second matrix.
Public Sub MatrixAdd()
    Call BinaryOp("add")
End Sub

''' @Description: A - B. Prompts for the second matrix.
Public Sub MatrixSubtract()
    Call BinaryOp("sub")
End Sub

''' @Description: A * B (matrix product). Prompts for the second matrix.
Public Sub MatrixMultiply()
    Call BinaryOp("mul")
End Sub

''' @Description: Element-wise A * B. Prompts for the second matrix.
Public Sub MatrixHadamard()
    Call BinaryOp("hadamard")
End Sub

''' @Description: Kronecker product A ⊗ B. Prompts for the second matrix.
Public Sub MatrixKronecker()
    Call BinaryOp("kron")
End Sub

''' @Description: Solve A X = B for X. Selection is A; prompts for B.
Public Sub MatrixSolve()
    Call BinaryOp("solve")
End Sub

''' @Description: Outer product of two vectors. Selection is the first; prompts for the second.
Public Sub MatrixOuter()
    Call BinaryOp("outer")
End Sub

''' @Description: Dot product of two vectors. Selection is the first; prompts for the second.
Public Sub MatrixDot()
    Call BinaryOp("dot")
End Sub

''' @Description: Lower-triangular Cholesky factor L of a symmetric positive-definite selection (A = L L^T).
Public Sub MatrixCholesky()
    Call UnaryOp("chol")
End Sub

''' @Description: Symmetric eigen-decomposition. Writes eigenvectors then an eigenvalues column.
Public Sub MatrixEigen()
    Call UnaryOp("eigen")
End Sub

''' @Description: Thin QR (modified Gram-Schmidt). Writes Q, then R immediately below Q.
Public Sub MatrixQR()
    Call UnaryOp("qr")
End Sub

''' @Description: LU factorization. Writes L, then U immediately below L.
Public Sub MatrixLU()
    Call UnaryOp("lu")
End Sub

' --- worksheet UDFs (Personal Fn_Matrices* role; keep short Mat* names) ---

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

Public Function MatIdentity(ByVal n As Long) As Variant
    On Error GoTo Fail
    MatIdentity = modInternalMatrices.Identity(n)
    Exit Function
Fail:
    MatIdentity = CVErr(xlErrValue)
End Function

Public Function MatZeros(ByVal n As Long, Optional ByVal m As Variant) As Variant
    On Error GoTo Fail
    If IsMissing(m) Then
        MatZeros = modInternalMatrices.Filled(n, n, 0)
    Else
        MatZeros = modInternalMatrices.Filled(n, CLng(m), 0)
    End If
    Exit Function
Fail:
    MatZeros = CVErr(xlErrValue)
End Function

Public Function MatOnes(ByVal n As Long, Optional ByVal m As Variant) As Variant
    On Error GoTo Fail
    If IsMissing(m) Then
        MatOnes = modInternalMatrices.Filled(n, n, 1)
    Else
        MatOnes = modInternalMatrices.Filled(n, CLng(m), 1)
    End If
    Exit Function
Fail:
    MatOnes = CVErr(xlErrValue)
End Function

Public Function MatHilbert(ByVal n As Long) As Variant
    On Error GoTo Fail
    MatHilbert = modInternalMatrices.Hilbert(n)
    Exit Function
Fail:
    MatHilbert = CVErr(xlErrValue)
End Function

Public Function MatExchange(ByVal n As Long) As Variant
    On Error GoTo Fail
    MatExchange = modInternalMatrices.ExchangeMat(n)
    Exit Function
Fail:
    MatExchange = CVErr(xlErrValue)
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

Public Function MatChol(ByVal rng As Range) As Variant
    On Error GoTo Fail
    MatChol = modInternalMatrices.Cholesky(modInternalMatrices.RangeToMatrix(rng))
    Exit Function
Fail:
    MatChol = CVErr(xlErrNum)
End Function

Public Function MatQR(ByVal rng As Range) As Variant
    On Error GoTo Fail
    MatQR = modInternalMatrices.QRFactors(modInternalMatrices.RangeToMatrix(rng))
    Exit Function
Fail:
    MatQR = CVErr(xlErrNum)
End Function

Public Function MatEigen(ByVal rng As Range) As Variant
    On Error GoTo Fail
    MatEigen = modInternalMatrices.EigenSymmetric(modInternalMatrices.RangeToMatrix(rng))
    Exit Function
Fail:
    MatEigen = CVErr(xlErrNum)
End Function

Public Function MatLU(ByVal rng As Range) As Variant
    On Error GoTo Fail
    MatLU = modInternalMatrices.LUFactors(modInternalMatrices.RangeToMatrix(rng))
    Exit Function
Fail:
    MatLU = CVErr(xlErrNum)
End Function

Public Function MatNormF(ByVal rng As Range) As Variant
    On Error GoTo Fail
    MatNormF = modInternalMatrices.FrobeniusNorm(modInternalMatrices.RangeToMatrix(rng))
    Exit Function
Fail:
    MatNormF = CVErr(xlErrValue)
End Function

Public Function MatNorm1(ByVal rng As Range) As Variant
    On Error GoTo Fail
    MatNorm1 = modInternalMatrices.Norm1(modInternalMatrices.RangeToMatrix(rng))
    Exit Function
Fail:
    MatNorm1 = CVErr(xlErrValue)
End Function

Public Function MatNormInf(ByVal rng As Range) As Variant
    On Error GoTo Fail
    MatNormInf = modInternalMatrices.NormInf(modInternalMatrices.RangeToMatrix(rng))
    Exit Function
Fail:
    MatNormInf = CVErr(xlErrValue)
End Function

Public Function MatRank(ByVal rng As Range) As Variant
    On Error GoTo Fail
    MatRank = modInternalMatrices.RankOf(modInternalMatrices.RangeToMatrix(rng))
    Exit Function
Fail:
    MatRank = CVErr(xlErrValue)
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

Public Function MatAdj(ByVal rng As Range) As Variant
    On Error GoTo Fail
    MatAdj = modInternalMatrices.Adjugate(modInternalMatrices.RangeToMatrix(rng))
    Exit Function
Fail:
    MatAdj = CVErr(xlErrNum)
End Function

Public Function MatPInv(ByVal rng As Range) As Variant
    On Error GoTo Fail
    MatPInv = modInternalMatrices.PseudoInverse(modInternalMatrices.RangeToMatrix(rng))
    Exit Function
Fail:
    MatPInv = CVErr(xlErrNum)
End Function

Public Function MatIsSymmetric(ByVal rng As Range) As Variant
    On Error GoTo Fail
    MatIsSymmetric = modInternalMatrices.IsSymmetric(modInternalMatrices.RangeToMatrix(rng))
    Exit Function
Fail:
    MatIsSymmetric = CVErr(xlErrValue)
End Function

Public Function MatIsOrthogonal(ByVal rng As Range) As Variant
    On Error GoTo Fail
    MatIsOrthogonal = modInternalMatrices.IsOrthogonal(modInternalMatrices.RangeToMatrix(rng))
    Exit Function
Fail:
    MatIsOrthogonal = CVErr(xlErrValue)
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

' Register Mat* names so they appear under Insert Function, category "Excel VBA Lib".
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
    Call RegisterUdf("MatAdj", "Classical adjugate (det times inverse).", "rng")
    Call RegisterUdf("MatPInv", "Moore-Penrose pseudoinverse via normal equations.", "rng")
    Call RegisterUdf("MatIsSymmetric", "TRUE if A equals A-transpose within 1e-9.", "rng")
    Call RegisterUdf("MatIsOrthogonal", "TRUE if A^T A is identity within 1e-9.", "rng")
    Call RegisterUdf("MatToeplitz", "Symmetric Toeplitz matrix from the first row or column.", "rng")
    Call RegisterUdf("MatVander", "Vandermonde matrix. Omit cols for a square matrix.", "rng,cols")
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

Private Function NeedSelection() As Range
    If TypeName(Selection) <> "Range" Then
        MsgBox "Select a range first.", vbExclamation, "Matrices"
        Exit Function
    End If
    Set NeedSelection = Selection
End Function

Private Function PromptNumber(ByVal PromptText As String, ByVal DefaultValue As Variant) As Variant
    Dim resp As Variant
    resp = Application.InputBox(Prompt:=PromptText, Title:="Matrices", Default:=DefaultValue, Type:=1)
    If VarType(resp) = vbBoolean Then
        PromptNumber = Empty
    Else
        PromptNumber = resp
    End If
End Function

Private Function PromptRange(ByVal PromptText As String) As Range
    Dim rng As Range
    On Error Resume Next
    Set rng = Application.InputBox(Prompt:=PromptText, Title:="Matrices", Type:=8)
    On Error GoTo 0
    Set PromptRange = rng
End Function

Private Sub WriteCreated(ByVal kind As String)
    Dim dest As Range
    Dim n As Variant
    Dim m As Variant
    Dim lo As Variant
    Dim hi As Variant
    Dim mat As Variant
    On Error GoTo EH
    If TypeName(Selection) <> "Range" Then
        MsgBox "Select the top-left cell for the new matrix.", vbExclamation, "Matrices"
        Exit Sub
    End If
    Set dest = Selection.Cells(1, 1)
    n = PromptNumber("Number of rows", 3)
    If IsEmpty(n) Then Exit Sub
    If kind = "identity" Or kind = "hilbert" Or kind = "exchange" Then
        m = n
    Else
        m = PromptNumber("Number of columns", CLng(n))
        If IsEmpty(m) Then Exit Sub
    End If
    If kind = "random" Then
        lo = PromptNumber("Minimum value", 0)
        If IsEmpty(lo) Then Exit Sub
        hi = PromptNumber("Maximum value", 1)
        If IsEmpty(hi) Then Exit Sub
        If CDbl(lo) > CDbl(hi) Then
            MsgBox "Minimum cannot be greater than maximum.", vbExclamation, "Matrices"
            Exit Sub
        End If
    End If
    Call modInternalExcelApp.PushAppState
    Select Case kind
        Case "identity"
            mat = modInternalMatrices.Identity(CLng(n))
        Case "zeros"
            mat = modInternalMatrices.Filled(CLng(n), CLng(m), 0)
        Case "ones"
            mat = modInternalMatrices.Filled(CLng(n), CLng(m), 1)
        Case "hilbert"
            mat = modInternalMatrices.Hilbert(CLng(n))
        Case "exchange"
            mat = modInternalMatrices.ExchangeMat(CLng(n))
        Case "random"
            mat = modInternalMatrices.RandomMatrix(CLng(n), CLng(m), CDbl(lo), CDbl(hi))
    End Select
    Call modInternalMatrices.PutMatrix(dest, mat)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("MatrixCreate")
End Sub

Private Sub FromSelectionCreate(ByVal kind As String)
    Dim rng As Range
    Dim mat As Variant
    Dim cols As Variant
    On Error GoTo EH
    Set rng = NeedSelection()
    If rng Is Nothing Then Exit Sub
    If kind = "vander" Then
        cols = PromptNumber("Number of columns (degree + 1)", rng.Cells.Count)
        If IsEmpty(cols) Then Exit Sub
    End If
    Call modInternalExcelApp.PushAppState
    Select Case kind
        Case "diag"
            mat = modInternalMatrices.DiagonalFromVector(modInternalMatrices.RangeToMatrix(rng))
        Case "toeplitz"
            mat = modInternalMatrices.ToeplitzFromVector(modInternalMatrices.RangeToMatrix(rng))
        Case "vander"
            mat = modInternalMatrices.VandermondeFromVector(modInternalMatrices.RangeToMatrix(rng), CLng(cols))
    End Select
    Call modInternalMatrices.PutMatrix(modInternalMatrices.OutputOrigin(rng), mat)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("MatrixCreate")
End Sub

Private Sub UnaryOp(ByVal kind As String, Optional ByVal param As Variant)
    Dim rng As Range
    Dim dest As Range
    Dim a As Variant
    Dim out As Variant
    On Error GoTo EH
    Set rng = NeedSelection()
    If rng Is Nothing Then Exit Sub
    Call modInternalExcelApp.PushAppState
    a = modInternalMatrices.RangeToMatrix(rng)
    Set dest = modInternalMatrices.OutputOrigin(rng)
    Select Case kind
        Case "sym"
            dest.Value = modInternalMatrices.IsSymmetric(a)
        Case "orth"
            dest.Value = modInternalMatrices.IsOrthogonal(a)
        Case "transpose"
            out = modInternalMatrices.TransposeMatrix(a)
            Call modInternalMatrices.PutMatrix(dest, out)
        Case "inverse"
            out = modInternalMatrices.Inverse(a)
            Call modInternalMatrices.PutMatrix(dest, out)
        Case "trace"
            out = modInternalMatrices.ScalarMatrix(modInternalMatrices.TraceOf(a))
            Call modInternalMatrices.PutMatrix(dest, out)
        Case "det"
            out = modInternalMatrices.ScalarMatrix(modInternalMatrices.Determinant(a))
            Call modInternalMatrices.PutMatrix(dest, out)
        Case "norm"
            out = modInternalMatrices.ScalarMatrix(modInternalMatrices.FrobeniusNorm(a))
            Call modInternalMatrices.PutMatrix(dest, out)
        Case "norm1"
            out = modInternalMatrices.ScalarMatrix(modInternalMatrices.Norm1(a))
            Call modInternalMatrices.PutMatrix(dest, out)
        Case "normInf"
            out = modInternalMatrices.ScalarMatrix(modInternalMatrices.NormInf(a))
            Call modInternalMatrices.PutMatrix(dest, out)
        Case "rank"
            out = modInternalMatrices.ScalarMatrix(CDbl(modInternalMatrices.RankOf(a)))
            Call modInternalMatrices.PutMatrix(dest, out)
        Case "chol"
            out = modInternalMatrices.Cholesky(a)
            Call modInternalMatrices.PutMatrix(dest, out)
        Case "eigen"
            out = modInternalMatrices.EigenSymmetric(a)
            Call modInternalMatrices.PutMatrix(dest, out)
        Case "qr"
            out = modInternalMatrices.QRFactors(a)
            Call modInternalMatrices.PutMatrix(dest, out)
        Case "lu"
            out = modInternalMatrices.LUFactors(a)
            Call modInternalMatrices.PutMatrix(dest, out)
        Case "scale"
            out = modInternalMatrices.ScaleMat(a, CDbl(param))
            Call modInternalMatrices.PutMatrix(dest, out)
        Case "power"
            out = modInternalMatrices.MatPower(a, CLng(param))
            Call modInternalMatrices.PutMatrix(dest, out)
        Case "diag"
            out = modInternalMatrices.ExtractDiag(a)
            Call modInternalMatrices.PutMatrix(dest, out)
        Case "adjugate"
            out = modInternalMatrices.Adjugate(a)
            Call modInternalMatrices.PutMatrix(dest, out)
        Case "pinv"
            out = modInternalMatrices.PseudoInverse(a)
            Call modInternalMatrices.PutMatrix(dest, out)
        Case Else
            Err.Raise 5, "MatrixOp", "Unknown unary op."
    End Select
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("MatrixOp")
End Sub

Private Sub BinaryOp(ByVal kind As String)
    Dim rng As Range
    Dim other As Range
    Dim a As Variant
    Dim b As Variant
    Dim out As Variant
    On Error GoTo EH
    Set rng = NeedSelection()
    If rng Is Nothing Then Exit Sub
    Set other = PromptRange("Select the second matrix")
    If other Is Nothing Then Exit Sub
    Call modInternalExcelApp.PushAppState
    a = modInternalMatrices.RangeToMatrix(rng)
    b = modInternalMatrices.RangeToMatrix(other)
    Select Case kind
        Case "add"
            out = modInternalMatrices.AddScaled(a, b, 1)
        Case "sub"
            out = modInternalMatrices.AddScaled(a, b, -1)
        Case "mul"
            out = modInternalMatrices.MatrixMultDefined(a, b)
        Case "hadamard"
            out = modInternalMatrices.Hadamard(a, b)
        Case "kron"
            out = modInternalMatrices.Kronecker(a, b)
        Case "solve"
            out = modInternalMatrices.Solve(a, b)
        Case "outer"
            out = modInternalMatrices.OuterProduct(a, b)
        Case "dot"
            out = modInternalMatrices.ScalarMatrix(modInternalMatrices.DotProduct(a, b))
        Case Else
            Err.Raise 5, "MatrixOp", "Unknown binary op."
    End Select
    Call modInternalMatrices.PutMatrix(modInternalMatrices.OutputOrigin(rng), out)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("MatrixOp")
End Sub
