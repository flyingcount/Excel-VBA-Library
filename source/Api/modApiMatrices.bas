Attribute VB_Name = "modApiMatrices"
Option Explicit

' Public API: matrix create / operate / decompose (Personal Custom_Menu13_* / Fn_Matrices*).
' Menu macros fill to the right of the selection. UDFs are worksheet-callable.

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

''' @Description: Build a diagonal matrix from the selected row or column vector.
Public Sub MatrixCreateDiagonal()
    Dim rng As Range
    Dim mat As Variant
    On Error GoTo EH
    Set rng = NeedSelection()
    If rng Is Nothing Then Exit Sub
    Call modInternalExcelApp.PushAppState
    mat = modInternalMatrices.DiagonalFromVector(modInternalMatrices.RangeToMatrix(rng))
    Call modInternalMatrices.PutMatrix(modInternalMatrices.OutputOrigin(rng), mat)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("MatrixCreateDiagonal")
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

''' @Description: TRUE if the selection is square and symmetric.
Public Sub MatrixIsSymmetric()
    Call UnaryOp("sym")
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

' --- worksheet UDFs (Personal Fn_Matrices* role; keep short Mat* names) ---

Public Function MatTranspose(ByVal rng As Range) As Variant
    On Error GoTo Fail
    MatTranspose = modInternalMatrices.Transpose(modInternalMatrices.RangeToMatrix(rng))
    Exit Function
Fail:
    MatTranspose = CVErr(xlErrValue)
End Function

Public Function MatInv(ByVal rng As Range) As Variant
    On Error GoTo Fail
    MatInv = modInternalMatrices.Inverse(modInternalMatrices.RangeToMatrix(rng))
    Exit Function
Fail:
    MatInv = CVErr(xlErrValue)
End Function

Public Function MatDet(ByVal rng As Range) As Variant
    On Error GoTo Fail
    MatDet = modInternalMatrices.Determinant(modInternalMatrices.RangeToMatrix(rng))
    Exit Function
Fail:
    MatDet = CVErr(xlErrValue)
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
    MatAdd = modInternalMatrices.Add( _
        modInternalMatrices.RangeToMatrix(a), modInternalMatrices.RangeToMatrix(b), 1)
    Exit Function
Fail:
    MatAdd = CVErr(xlErrValue)
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
    MatChol = CVErr(xlErrValue)
End Function

Public Function MatNormF(ByVal rng As Range) As Variant
    On Error GoTo Fail
    MatNormF = modInternalMatrices.FrobeniusNorm(modInternalMatrices.RangeToMatrix(rng))
    Exit Function
Fail:
    MatNormF = CVErr(xlErrValue)
End Function

Public Function MatIsSymmetric(ByVal rng As Range) As Variant
    On Error GoTo Fail
    MatIsSymmetric = modInternalMatrices.IsSymmetric(modInternalMatrices.RangeToMatrix(rng))
    Exit Function
Fail:
    MatIsSymmetric = CVErr(xlErrValue)
End Function

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
    If kind = "identity" Or kind = "hilbert" Then
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

Private Sub UnaryOp(ByVal kind As String)
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
    If kind = "sym" Then
        dest.Value = modInternalMatrices.IsSymmetric(a)
    Else
        Select Case kind
            Case "transpose"
                out = modInternalMatrices.Transpose(a)
            Case "inverse"
                out = modInternalMatrices.Inverse(a)
            Case "trace"
                out = modInternalMatrices.ScalarMatrix(modInternalMatrices.TraceOf(a))
            Case "det"
                out = modInternalMatrices.ScalarMatrix(modInternalMatrices.Determinant(a))
            Case "norm"
                out = modInternalMatrices.ScalarMatrix(modInternalMatrices.FrobeniusNorm(a))
            Case "chol"
                out = modInternalMatrices.Cholesky(a)
            Case "eigen"
                out = modInternalMatrices.EigenSymmetric(a)
            Case "qr"
                out = modInternalMatrices.QRFactors(a)
        End Select
        Call modInternalMatrices.PutMatrix(dest, out)
    End If
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
            out = modInternalMatrices.Add(a, b, 1)
        Case "sub"
            out = modInternalMatrices.Add(a, b, -1)
        Case "mul"
            out = modInternalMatrices.MatrixMultDefined(a, b)
        Case "hadamard"
            out = modInternalMatrices.Hadamard(a, b)
        Case "kron"
            out = modInternalMatrices.Kronecker(a, b)
        Case "solve"
            out = modInternalMatrices.Solve(a, b)
    End Select
    Call modInternalMatrices.PutMatrix(modInternalMatrices.OutputOrigin(rng), out)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("MatrixOp")
End Sub
