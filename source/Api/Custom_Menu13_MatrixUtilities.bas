Attribute VB_Name = "Custom_Menu13_MatrixUtilities"
Option Explicit

' Personal Custom_Menu13_MatrixUtilities: shared prompts and sheet writes for Menu13.
' Math lives in modInternalMatrices. Internal must not call this module.

Private Sub ShowMatrixError()
    Dim desc As String
    desc = Err.Description
    Call modInternalExcelApp.PopAppState
    If Len(desc) = 0 Then desc = "The matrix operation could not be completed."
    MsgBox desc, vbExclamation, "Matrices"
End Sub

Public Function NeedSelection() As Range
    If TypeName(Selection) <> "Range" Then
        MsgBox "Select a range first.", vbExclamation, "Matrices"
        Exit Function
    End If
    Set NeedSelection = Selection
End Function

''' @Description: Size of a matrix as "rows x columns".
Public Function MatrixSize(ByVal var_Input As Variant) As String
    Dim rng As Range
    Dim a As Variant
    If TypeName(var_Input) = "Range" Then
        Set rng = var_Input
        MatrixSize = CStr(rng.Rows.Count) & " x " & CStr(rng.Columns.Count)
        Exit Function
    End If
    a = Custom_Menu13_Matrices1.ToMatrix2D(var_Input)
    MatrixSize = CStr(UBound(a, 1) - LBound(a, 1) + 1) & " x " & CStr(UBound(a, 2) - LBound(a, 2) + 1)
End Function

''' @Description: Write the matrix size two rows below the matrix (output cell can be changed).
Public Sub MatrixWriteSize()
    Dim rng As Range
    Dim dest As Range
    On Error GoTo EH
    Set rng = PromptRange("Select the matrix")
    If rng Is Nothing Then Exit Sub
    Set dest = PromptValidationOutput(rng)
    If dest Is Nothing Then Exit Sub
    Call PutScalarCell(dest, MatrixSize(rng))
    Exit Sub
EH:
    Call ShowMatrixError
End Sub

Private Function PromptValidationOutput(ByVal rng As Range) As Range
    Dim def As Range
    Set def = rng.Cells(rng.Rows.Count, 1).Offset(2, 0)
    Set PromptValidationOutput = PromptRange("Select the output location", def)
End Function

' Write a scalar or a small 2D block without assigning into a leftover array formula / spill.
Private Sub PutScalarCell(ByVal dest As Range, ByVal value As Variant)
    Dim cell As Range
    Dim box As Variant
    Dim nr As Long
    Dim nc As Long
    Set cell = dest.Cells(1, 1)
    On Error Resume Next
    If cell.HasArray Then
        cell.CurrentArray.ClearContents
    End If
    On Error GoTo 0
    If IsArray(value) Then
        nr = UBound(value, 1) - LBound(value, 1) + 1
        nc = UBound(value, 2) - LBound(value, 2) + 1
        On Error Resume Next
        If nc > 1 Then
            If cell.Offset(0, nc - 1).HasArray Then
                cell.Offset(0, nc - 1).CurrentArray.ClearContents
            End If
        End If
        On Error GoTo 0
        cell.Resize(nr, nc).Value = value
    Else
        ReDim box(1 To 1, 1 To 1)
        box(1, 1) = value
        cell.Resize(1, 1).Value = box
    End If
End Sub

Public Sub WriteValidationBool(ByVal kind As String)
    Dim rng As Range
    Dim dest As Range
    Dim result As Boolean
    Dim heading As String
    Dim pair As Variant
    On Error GoTo EH
    Set rng = PromptRange("Select the matrix")
    If rng Is Nothing Then Exit Sub
    Select Case kind
        Case "sym"
            heading = "Symmetric"
            result = modInternalMatrices.IsSymmetric(modInternalMatrices.RangeToMatrix(rng))
        Case "orth"
            heading = "Orthogonal"
            result = modInternalMatrices.IsOrthogonal(modInternalMatrices.RangeToMatrix(rng))
        Case Else
            Err.Raise 5, "MatrixValidation", "Unknown validation kind."
    End Select
    Set dest = PromptValidationOutput(rng)
    If dest Is Nothing Then Exit Sub
    ReDim pair(1 To 1, 1 To 2)
    pair(1, 1) = heading
    pair(1, 2) = result
    Call PutScalarCell(dest, pair)
    Exit Sub
EH:
    Call ShowMatrixError
End Sub

Private Function PropertyHeading(ByVal kind As String) As String
    Select Case kind
        Case "det"
            PropertyHeading = "Determinant"
        Case "trace"
            PropertyHeading = "Trace"
        Case "rank"
            PropertyHeading = "Rank"
        Case "norm"
            PropertyHeading = "Frobenius norm"
        Case "norm1"
            PropertyHeading = "1-norm"
        Case "normInf"
            PropertyHeading = "Infinity-norm"
        Case "cond"
            PropertyHeading = "Condition number"
        Case "spectral"
            PropertyHeading = "Spectral radius"
        Case "eigenvals"
            PropertyHeading = "Eigenvalues"
        Case "eigenvecs"
            PropertyHeading = "Eigenvectors"
        Case Else
            Err.Raise 5, "MatrixProperty", "Unknown property."
    End Select
End Function

Private Function PropertyValueOf(ByRef a As Variant, ByVal kind As String) As Variant
    Dim square As Boolean
    square = (modInternalMatrices.RowsOf(a) = modInternalMatrices.ColsOf(a))
    Select Case kind
        Case "det"
            If square Then
                PropertyValueOf = modInternalMatrices.Determinant(a)
            Else
                PropertyValueOf = "Not defined (not square)"
            End If
        Case "trace"
            If square Then
                PropertyValueOf = modInternalMatrices.TraceOf(a)
            Else
                PropertyValueOf = "Not defined (not square)"
            End If
        Case "rank"
            PropertyValueOf = modInternalMatrices.RankOf(a)
        Case "norm"
            PropertyValueOf = modInternalMatrices.FrobeniusNorm(a)
        Case "norm1"
            PropertyValueOf = modInternalMatrices.Norm1(a)
        Case "normInf"
            PropertyValueOf = modInternalMatrices.NormInf(a)
        Case "cond"
            PropertyValueOf = ConditionNumberOrMessage(a, square)
        Case "spectral"
            PropertyValueOf = SpectralRadiusOrMessage(a, square)
        Case "eigenvals"
            PropertyValueOf = EigenblockOrMessage(a, square, "vals")
        Case "eigenvecs"
            PropertyValueOf = EigenblockOrMessage(a, square, "vecs")
        Case Else
            Err.Raise 5, "MatrixProperty", "Unknown property."
    End Select
End Function

Private Function ConditionNumberOrMessage(ByRef a As Variant, ByVal square As Boolean) As Variant
    On Error GoTo fail
    If Not square Then
        ConditionNumberOrMessage = "Not defined (not square)"
        Exit Function
    End If
    ConditionNumberOrMessage = modInternalMatrices.ConditionNumberInf(a)
    Exit Function
fail:
    ConditionNumberOrMessage = "Undefined (singular)"
End Function

Private Function SpectralRadiusOrMessage(ByRef a As Variant, ByVal square As Boolean) As Variant
    On Error GoTo fail
    If Not square Then
        SpectralRadiusOrMessage = "Not defined (not square)"
        Exit Function
    End If
    SpectralRadiusOrMessage = modInternalMatrices.SpectralRadiusOf(a)
    Exit Function
fail:
    SpectralRadiusOrMessage = Err.Description
End Function

Private Function EigenblockOrMessage(ByRef a As Variant, ByVal square As Boolean, ByVal which As String) As Variant
    On Error GoTo fail
    If Not square Then
        EigenblockOrMessage = "Not defined (not square)"
        Exit Function
    End If
    If Not modInternalMatrices.IsSymmetric(a, 0.000000001) Then
        EigenblockOrMessage = "Needs a symmetric matrix"
        Exit Function
    End If
    If which = "vecs" Then
        EigenblockOrMessage = modInternalMatrices.EigenvectorsOf(a)
    Else
        EigenblockOrMessage = modInternalMatrices.EigenvaluesColumn(a)
    End If
    Exit Function
fail:
    EigenblockOrMessage = Err.Description
End Function

Private Function ScalarPropertyKinds() As Variant
    ScalarPropertyKinds = Array("det", "trace", "rank", "norm", "norm1", "normInf", "cond", "spectral")
End Function

Private Sub PutLabeledProperty(ByVal dest As Range, ByVal heading As String, ByVal value As Variant)
    Dim cell As Range
    Set cell = dest.Cells(1, 1)
    Call PutScalarCell(cell, heading)
    Call PutScalarCell(cell.Offset(0, 1), value)
End Sub

Public Sub WriteProperty(ByVal kind As String)
    Dim rng As Range
    Dim dest As Range
    Dim a As Variant
    On Error GoTo EH
    Set rng = PromptRange("Select the matrix")
    If rng Is Nothing Then Exit Sub
    a = modInternalMatrices.RangeToMatrix(rng)
    Set dest = PromptValidationOutput(rng)
    If dest Is Nothing Then Exit Sub
    Call PutLabeledProperty(dest, PropertyHeading(kind), PropertyValueOf(a, kind))
    Exit Sub
EH:
    Call ShowMatrixError
End Sub

''' @Description: Write every Matrices/Properties value as a label/result list.
Public Sub MatrixPropertiesAll()
    Dim rng As Range
    Dim dest As Range
    Dim a As Variant
    Dim kinds As Variant
    Dim tbl As Variant
    Dim i As Long
    Dim n As Long
    Dim propKind As String
    Dim cell As Range
    Dim val As Variant
    On Error GoTo EH
    Set rng = PromptRange("Select the matrix")
    If rng Is Nothing Then Exit Sub
    a = modInternalMatrices.RangeToMatrix(rng)
    Set dest = PromptValidationOutput(rng)
    If dest Is Nothing Then Exit Sub
    kinds = ScalarPropertyKinds()
    n = UBound(kinds) - LBound(kinds) + 1
    ReDim tbl(1 To n, 1 To 2)
    For i = 1 To n
        propKind = CStr(kinds(LBound(kinds) + i - 1))
        tbl(i, 1) = PropertyHeading(propKind)
        tbl(i, 2) = PropertyValueOf(a, propKind)
    Next i
    Call PutScalarCell(dest, tbl)
    Set cell = dest.Cells(1, 1).Offset(n, 0)
    val = PropertyValueOf(a, "eigenvals")
    Call PutLabeledProperty(cell, PropertyHeading("eigenvals"), val)
    If IsArray(val) Then
        Set cell = cell.Offset(modInternalMatrices.RowsOf(val), 0)
    Else
        Set cell = cell.Offset(1, 0)
    End If
    val = PropertyValueOf(a, "eigenvecs")
    Call PutLabeledProperty(cell, PropertyHeading("eigenvecs"), val)
    Exit Sub
EH:
    Call ShowMatrixError
End Sub

Public Function PromptNumber(ByVal PromptText As String, ByVal DefaultValue As Variant) As Variant
    Dim resp As Variant
    resp = Application.InputBox(Prompt:=PromptText, Title:="Matrices", Default:=DefaultValue, Type:=1)
    If VarType(resp) = vbBoolean Then
        PromptNumber = Empty
    Else
        PromptNumber = resp
    End If
End Function

Public Function PromptRange(ByVal PromptText As String, Optional ByVal DefaultRng As Range) As Range
    Dim rng As Range
    Dim defAddr As String
    If DefaultRng Is Nothing Then
        If TypeName(Selection) = "Range" Then Set DefaultRng = Selection
    End If
    If Not DefaultRng Is Nothing Then defAddr = DefaultRng.Address
    On Error Resume Next
    If Len(defAddr) > 0 Then
        Set rng = Application.InputBox(Prompt:=PromptText, Title:="Matrices", Default:=defAddr, Type:=8)
    Else
        Set rng = Application.InputBox(Prompt:=PromptText, Title:="Matrices", Type:=8)
    End If
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
    If kind = "randomSquare" Then
        n = PromptNumber("Matrix size (rows = columns)", 4)
        If IsEmpty(n) Then Exit Sub
        If n < 1 Or n <> Int(n) Then
            MsgBox "Size must be a positive integer.", vbExclamation, "Matrices"
            Exit Sub
        End If
        m = n
        lo = PromptNumber("From", 0)
        If IsEmpty(lo) Then Exit Sub
        hi = PromptNumber("To", 1)
        If IsEmpty(hi) Then Exit Sub
        If CDbl(lo) > CDbl(hi) Then
            MsgBox "From cannot be greater than To.", vbExclamation, "Matrices"
            Exit Sub
        End If
    Else
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
        Case "randomSquare"
            mat = modInternalMatrices.RandomMatrix(CLng(n), CLng(m), CDbl(lo), CDbl(hi))
        Case Else
            Err.Raise 5, "MatrixCreate", "Unknown create kind."
    End Select
    Call modInternalMatrices.PutMatrix(dest, mat)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call ShowMatrixError
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
    Call ShowMatrixError
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
            Call PutScalarCell(dest, modInternalMatrices.IsSymmetric(a))
        Case "orth"
            Call PutScalarCell(dest, modInternalMatrices.IsOrthogonal(a))
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
    Call ShowMatrixError
End Sub

Public Sub BinaryOp(ByVal kind As String)
    Dim rng As Range
    Dim other As Range
    Dim a As Variant
    Dim b As Variant
    Dim out As Variant
    On Error GoTo EH
    Set rng = PromptRange("Select the first matrix")
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
    Call ShowMatrixError
End Sub

' Excel VBA Lib menu OnAction names (kept here so Personal Custom_Menu13_Matrices1 /
' Custom_Menu13_Matrices2 can live in the add-in under their original names).

Public Sub MatrixTranspose()
    Call UnaryOp("transpose")
End Sub

Public Sub MatrixAdd()
    Call BinaryOp("add")
End Sub

Public Sub MatrixSubtract()
    Call BinaryOp("sub")
End Sub

Public Sub MatrixScale()
    Dim k As Variant
    k = PromptNumber("Scale factor", 1)
    If IsEmpty(k) Then Exit Sub
    Call UnaryOp("scale", k)
End Sub

Public Sub MatrixMultiply()
    Call BinaryOp("mul")
End Sub

Public Sub MatrixHadamard()
    Call BinaryOp("hadamard")
End Sub

Public Sub MatrixKronecker()
    Call BinaryOp("kron")
End Sub

Public Sub MatrixOuter()
    Call BinaryOp("outer")
End Sub

Public Sub MatrixDot()
    Call BinaryOp("dot")
End Sub

Public Sub MatrixInverse()
    Call UnaryOp("inverse")
End Sub

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

Public Sub MatrixDeterminant()
    Call WriteProperty("det")
End Sub

Public Sub MatrixTrace()
    Call WriteProperty("trace")
End Sub

Public Sub MatrixDiagExtract()
    Call UnaryOp("diag")
End Sub

Public Sub MatrixVec()
    Call UnaryOp("vec")
End Sub

Public Sub MatrixUnvec()
    Dim nRows As Variant
    nRows = PromptNumber("Number of rows in the result", 1)
    If IsEmpty(nRows) Then Exit Sub
    Call UnaryOp("unvec", nRows)
End Sub

Public Sub MatrixSolve()
    Call BinaryOp("solve")
End Sub

Public Sub MatrixRank()
    Call WriteProperty("rank")
End Sub

Public Sub MatrixNorm()
    Call WriteProperty("norm")
End Sub

Public Sub MatrixNorm1()
    Call WriteProperty("norm1")
End Sub

Public Sub MatrixNormInf()
    Call WriteProperty("normInf")
End Sub

Public Sub MatrixConditionNumber()
    Call WriteProperty("cond")
End Sub

Public Sub MatrixSpectralRadius()
    Call WriteProperty("spectral")
End Sub

Public Sub MatrixEigenvalues()
    Call WriteProperty("eigenvals")
End Sub

Public Sub MatrixEigenvectors()
    Call WriteProperty("eigenvecs")
End Sub

Public Sub MatrixIsSymmetric()
    Call WriteValidationBool("sym")
End Sub

Public Sub MatrixAdjugate()
    Call UnaryOp("adjugate")
End Sub

Public Sub MatrixPseudoInverse()
    Call UnaryOp("pinv")
End Sub

Public Sub MatrixLU()
    Call UnaryOp("lu")
End Sub

Public Sub MatrixCofactor()
    Call UnaryOp("cofactor")
End Sub

Public Sub MatrixMinor()
    Dim rng As Range
    Dim r As Variant
    Dim c As Variant
    Dim a As Variant
    Dim out As Variant
    On Error GoTo MinorEH
    Set rng = NeedSelection()
    If rng Is Nothing Then Exit Sub
    r = PromptNumber("Row to delete (1-based)", 1)
    If IsEmpty(r) Then Exit Sub
    c = PromptNumber("Column to delete (1-based)", 1)
    If IsEmpty(c) Then Exit Sub
    Call modInternalExcelApp.PushAppState
    a = modInternalMatrices.RangeToMatrix(rng)
    out = modInternalMatrices.MinorMatrix(a, CLng(r), CLng(c))
    Call modInternalMatrices.PutMatrix(modInternalMatrices.OutputOrigin(rng), out)
    Call modInternalExcelApp.PopAppState
    Exit Sub
MinorEH:
    Call ShowMatrixError
End Sub

Public Function IsMatrixEqual(var_Input1 As Variant, var_Input2 As Variant) As Boolean
    Dim lng_Row As Long
    Dim lng_Col As Long
    Dim arr_Input1 As Variant
    Dim arr_Input2 As Variant
    Dim int_Flag As Integer
    int_Flag = 1
    arr_Input1 = Custom_Menu13_Matrices1.ToMatrix2D(var_Input1)
    arr_Input2 = Custom_Menu13_Matrices1.ToMatrix2D(var_Input2)
    For lng_Row = LBound(arr_Input1, 1) To UBound(arr_Input1, 1)
        For lng_Col = LBound(arr_Input1, 2) To UBound(arr_Input1, 2)
            If arr_Input1(lng_Row, lng_Col) = arr_Input2(lng_Row, lng_Col) Then
                int_Flag = int_Flag * 1
            Else
                int_Flag = int_Flag * 0
            End If
        Next lng_Col
    Next lng_Row
    IsMatrixEqual = (int_Flag = 1)
End Function
