Attribute VB_Name = "Custom_Menu13_MatrixUtilities"
Option Explicit

' Personal Custom_Menu13_MatrixUtilities: shared prompts and sheet writes for Menu13.
' Math lives in modInternalMatrices. Internal must not call this module.

Public Function NeedSelection() As Range
    If TypeName(Selection) <> "Range" Then
        MsgBox "Select a range first.", vbExclamation, "Matrices"
        Exit Function
    End If
    Set NeedSelection = Selection
End Function

Public Function PromptNumber(ByVal PromptText As String, ByVal DefaultValue As Variant) As Variant
    Dim resp As Variant
    resp = Application.InputBox(Prompt:=PromptText, Title:="Matrices", Default:=DefaultValue, Type:=1)
    If VarType(resp) = vbBoolean Then
        PromptNumber = Empty
    Else
        PromptNumber = resp
    End If
End Function

Public Function PromptRange(ByVal PromptText As String) As Range
    Dim rng As Range
    On Error Resume Next
    Set rng = Application.InputBox(Prompt:=PromptText, Title:="Matrices", Type:=8)
    On Error GoTo 0
    Set PromptRange = rng
End Function

Public Sub WriteCreated(ByVal kind As String)
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
        Case Else
            Err.Raise 5, "MatrixCreate", "Unknown create kind."
    End Select
    Call modInternalMatrices.PutMatrix(dest, mat)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("MatrixCreate")
End Sub

Public Sub FromSelectionCreate(ByVal kind As String)
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
        Case "companion"
            mat = modInternalMatrices.CompanionFromVector(modInternalMatrices.RangeToMatrix(rng))
        Case Else
            Err.Raise 5, "MatrixCreate", "Unknown create kind."
    End Select
    Call modInternalMatrices.PutMatrix(modInternalMatrices.OutputOrigin(rng), mat)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("MatrixCreate")
End Sub

Public Sub UnaryOp(ByVal kind As String, Optional ByVal param As Variant)
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
        Case "vec"
            out = modInternalMatrices.Vectorize(a)
            Call modInternalMatrices.PutMatrix(dest, out)
        Case "unvec"
            out = modInternalMatrices.Unvectorize(a, CLng(param))
            Call modInternalMatrices.PutMatrix(dest, out)
        Case "cofactor"
            out = modInternalMatrices.CofactorMatrix(a)
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

Public Sub BinaryOp(ByVal kind As String)
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
