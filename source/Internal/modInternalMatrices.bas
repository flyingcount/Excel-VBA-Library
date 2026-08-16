Attribute VB_Name = "modInternalMatrices"
Option Explicit

' Internal: real matrix algebra for Matrices menu ops and Fn_Matrices* UDFs.
' Called from modApiMatrix* / Fn_Matrices*. Do not document these as the external API.
' Arrays are 1-based, 2-dimensional Doubles stored in Variant.

Public Const MatrixMaxN As Long = 250
Public Const CofactorMaxN As Long = 20

Public Function RowsOf(ByRef mat As Variant) As Long
    RowsOf = UBound(mat, 1) - LBound(mat, 1) + 1
End Function

Public Function ColsOf(ByRef mat As Variant) As Long
    ColsOf = UBound(mat, 2) - LBound(mat, 2) + 1
End Function

Public Function RangeToMatrix(ByVal rng As Range) As Variant
    Dim raw As Variant
    Dim mat As Variant
    Dim r As Long
    Dim c As Long
    Dim nr As Long
    Dim nc As Long

    If rng Is Nothing Then
        Err.Raise 5, "RangeToMatrix", "Select a range."
    End If
    If rng.Areas.Count > 1 Then
        Err.Raise 5, "RangeToMatrix", "Select a single contiguous range."
    End If

    nr = rng.Rows.Count
    nc = rng.Columns.Count
    ReDim mat(1 To nr, 1 To nc)
    If rng.Cells.Count = 1 Then
        raw = rng.Value
        mat(1, 1) = ToFiniteDouble(raw, 1, 1)
    Else
        raw = rng.Value
        For r = 1 To nr
            For c = 1 To nc
                mat(r, c) = ToFiniteDouble(raw(r, c), r, c)
            Next c
        Next r
    End If
    RangeToMatrix = mat
End Function

Public Sub PutMatrix(ByVal destTopLeft As Range, ByRef mat As Variant)
    Dim nr As Long
    Dim nc As Long
    nr = RowsOf(mat)
    nc = ColsOf(mat)
    destTopLeft.Resize(nr, nc).Value = mat
End Sub

Public Function OutputOrigin(ByVal src As Range) As Range
    Set OutputOrigin = src.Cells(1, 1).Offset(0, src.Columns.Count + 1)
End Function

Public Sub CheckSize(ByVal n As Long, ByVal m As Long)
    If n < 1 Or m < 1 Then
        Err.Raise 5, "CheckSize", "Matrix dimensions must be at least 1."
    End If
    If n > MatrixMaxN Or m > MatrixMaxN Then
        Err.Raise 5, "CheckSize", "Matrices larger than " & MatrixMaxN & " x " & MatrixMaxN & " are not supported."
    End If
End Sub

Public Function Identity(ByVal n As Long) As Variant
    Dim a As Variant
    Dim i As Long
    Dim j As Long
    Call CheckSize(n, n)
    ReDim a(1 To n, 1 To n)
    For i = 1 To n
        For j = 1 To n
            If i = j Then a(i, j) = 1 Else a(i, j) = 0
        Next j
    Next i
    Identity = a
End Function

Public Function Filled(ByVal n As Long, ByVal m As Long, ByVal value As Double) As Variant
    Dim a As Variant
    Dim i As Long
    Dim j As Long
    Call CheckSize(n, m)
    ReDim a(1 To n, 1 To m)
    For i = 1 To n
        For j = 1 To m
            a(i, j) = value
        Next j
    Next i
    Filled = a
End Function

Public Function RandomMatrix(ByVal n As Long, ByVal m As Long, ByVal lo As Double, ByVal hi As Double) As Variant
    Dim a As Variant
    Dim i As Long
    Dim j As Long
    Call CheckSize(n, m)
    ReDim a(1 To n, 1 To m)
    Randomize
    For i = 1 To n
        For j = 1 To m
            a(i, j) = lo + Rnd() * (hi - lo)
        Next j
    Next i
    RandomMatrix = a
End Function

Public Function Hilbert(ByVal n As Long) As Variant
    Dim a As Variant
    Dim i As Long
    Dim j As Long
    Call CheckSize(n, n)
    ReDim a(1 To n, 1 To n)
    For i = 1 To n
        For j = 1 To n
            a(i, j) = 1 / CDbl(i + j - 1)
        Next j
    Next i
    Hilbert = a
End Function

Public Function DiagonalFromVector(ByRef vec As Variant) As Variant
    Dim n As Long
    Dim a As Variant
    Dim i As Long
    Dim j As Long
    n = RowsOf(vec)
    If ColsOf(vec) <> 1 And RowsOf(vec) <> 1 Then
        Err.Raise 5, "DiagonalFromVector", "Select a single row or column."
    End If
    If ColsOf(vec) > 1 Then n = ColsOf(vec)
    Call CheckSize(n, n)
    ReDim a(1 To n, 1 To n)
    For i = 1 To n
        For j = 1 To n
            a(i, j) = 0
        Next j
        If ColsOf(vec) = 1 Then
            a(i, i) = CDbl(vec(i, 1))
        Else
            a(i, i) = CDbl(vec(1, i))
        End If
    Next i
    DiagonalFromVector = a
End Function

Public Function TransposeMatrix(ByRef a As Variant) As Variant
    Dim nr As Long
    Dim nc As Long
    Dim b As Variant
    Dim i As Long
    Dim j As Long
    nr = RowsOf(a)
    nc = ColsOf(a)
    ReDim b(1 To nc, 1 To nr)
    For i = 1 To nr
        For j = 1 To nc
            b(j, i) = a(i, j)
        Next j
    Next i
    TransposeMatrix = b
End Function

Public Function AddScaled(ByRef a As Variant, ByRef b As Variant, ByVal signB As Double) As Variant
    Dim nr As Long
    Dim nc As Long
    Dim c As Variant
    Dim i As Long
    Dim j As Long
    nr = RowsOf(a)
    nc = ColsOf(a)
    If RowsOf(b) <> nr Or ColsOf(b) <> nc Then
        Err.Raise 5, "Add", "Matrices must have the same shape."
    End If
    ReDim c(1 To nr, 1 To nc)
    For i = 1 To nr
        For j = 1 To nc
            c(i, j) = CDbl(a(i, j)) + signB * CDbl(b(i, j))
        Next j
    Next i
    AddScaled = c
End Function

Public Function Hadamard(ByRef a As Variant, ByRef b As Variant) As Variant
    Dim nr As Long
    Dim nc As Long
    Dim c As Variant
    Dim i As Long
    Dim j As Long
    nr = RowsOf(a)
    nc = ColsOf(a)
    If RowsOf(b) <> nr Or ColsOf(b) <> nc Then
        Err.Raise 5, "Hadamard", "Matrices must have the same shape."
    End If
    ReDim c(1 To nr, 1 To nc)
    For i = 1 To nr
        For j = 1 To nc
            c(i, j) = CDbl(a(i, j)) * CDbl(b(i, j))
        Next j
    Next i
    Hadamard = c
End Function

' Confirmed Personal helper name (InfrastructureCatalog).
Public Function MatrixMultDefined(ByRef a As Variant, ByRef b As Variant) As Variant
    Dim nr As Long
    Dim nk As Long
    Dim nc As Long
    Dim c As Variant
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim acc As Double
    nr = RowsOf(a)
    nk = ColsOf(a)
    If RowsOf(b) <> nk Then
        Err.Raise 5, "MatrixMultDefined", "Inner dimensions must match (A columns = B rows)."
    End If
    nc = ColsOf(b)
    Call CheckSize(nr, nc)
    ReDim c(1 To nr, 1 To nc)
    For i = 1 To nr
        For j = 1 To nc
            acc = 0
            For k = 1 To nk
                acc = acc + CDbl(a(i, k)) * CDbl(b(k, j))
            Next k
            c(i, j) = acc
        Next j
    Next i
    MatrixMultDefined = c
End Function

Public Function Kronecker(ByRef a As Variant, ByRef b As Variant) As Variant
    Dim ra As Long
    Dim ca As Long
    Dim rb As Long
    Dim cb As Long
    Dim c As Variant
    Dim i As Long
    Dim j As Long
    Dim p As Long
    Dim q As Long
    ra = RowsOf(a)
    ca = ColsOf(a)
    rb = RowsOf(b)
    cb = ColsOf(b)
    Call CheckSize(ra * rb, ca * cb)
    ReDim c(1 To ra * rb, 1 To ca * cb)
    For i = 1 To ra
        For j = 1 To ca
            For p = 1 To rb
                For q = 1 To cb
                    c((i - 1) * rb + p, (j - 1) * cb + q) = CDbl(a(i, j)) * CDbl(b(p, q))
                Next q
            Next p
        Next j
    Next i
    Kronecker = c
End Function

Public Function TraceOf(ByRef a As Variant) As Double
    Dim n As Long
    Dim i As Long
    Dim acc As Double
    Call RequireSquare(a, "Trace")
    n = RowsOf(a)
    acc = 0
    For i = 1 To n
        acc = acc + CDbl(a(i, i))
    Next i
    TraceOf = acc
End Function

Public Function FrobeniusNorm(ByRef a As Variant) As Double
    Dim i As Long
    Dim j As Long
    Dim acc As Double
    acc = 0
    For i = LBound(a, 1) To UBound(a, 1)
        For j = LBound(a, 2) To UBound(a, 2)
            acc = acc + CDbl(a(i, j)) * CDbl(a(i, j))
        Next j
    Next i
    FrobeniusNorm = Sqr(acc)
End Function

Public Function IsSymmetric(ByRef a As Variant, Optional ByVal tol As Double = 0.000000001) As Boolean
    Dim n As Long
    Dim i As Long
    Dim j As Long
    If RowsOf(a) <> ColsOf(a) Then
        IsSymmetric = False
        Exit Function
    End If
    n = RowsOf(a)
    For i = 1 To n
        For j = i + 1 To n
            If Abs(CDbl(a(i, j)) - CDbl(a(j, i))) > tol Then
                IsSymmetric = False
                Exit Function
            End If
        Next j
    Next i
    IsSymmetric = True
End Function

Public Function Determinant(ByRef a As Variant) As Double
    Dim lu As Variant
    Dim piv() As Long
    Dim sign As Long
    Call RequireSquare(a, "Determinant")
    lu = CopyMat(a)
    sign = LUFactor(lu, piv)
    If sign = 0 Then
        Determinant = 0
    Else
        Determinant = sign * ProductDiag(lu)
    End If
End Function

Public Function Inverse(ByRef a As Variant) As Variant
    Dim n As Long
    Dim aug As Variant
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim piv As Long
    Dim maxAbs As Double
    Dim tmp As Double
    Dim fac As Double
    Call RequireSquare(a, "Inverse")
    n = RowsOf(a)
    Call CheckSize(n, n)
    ReDim aug(1 To n, 1 To 2 * n)
    For i = 1 To n
        For j = 1 To n
            aug(i, j) = CDbl(a(i, j))
            If i = j Then aug(i, n + j) = 1 Else aug(i, n + j) = 0
        Next j
    Next i
    For k = 1 To n
        piv = k
        maxAbs = Abs(CDbl(aug(k, k)))
        For i = k + 1 To n
            If Abs(CDbl(aug(i, k))) > maxAbs Then
                maxAbs = Abs(CDbl(aug(i, k)))
                piv = i
            End If
        Next i
        If maxAbs < 1E-14 Then
            Err.Raise 5, "Inverse", "Matrix is singular."
        End If
        If piv <> k Then
            For j = 1 To 2 * n
                tmp = aug(k, j)
                aug(k, j) = aug(piv, j)
                aug(piv, j) = tmp
            Next j
        End If
        fac = CDbl(aug(k, k))
        For j = 1 To 2 * n
            aug(k, j) = CDbl(aug(k, j)) / fac
        Next j
        For i = 1 To n
            If i <> k Then
                fac = CDbl(aug(i, k))
                If fac <> 0 Then
                    For j = 1 To 2 * n
                        aug(i, j) = CDbl(aug(i, j)) - fac * CDbl(aug(k, j))
                    Next j
                End If
            End If
        Next i
    Next k
    Inverse = SliceCols(aug, n + 1, 2 * n)
End Function

Public Function Solve(ByRef a As Variant, ByRef b As Variant) As Variant
    Dim invA As Variant
    invA = Inverse(a)
    If RowsOf(b) <> RowsOf(a) Then
        Err.Raise 5, "Solve", "b must have the same number of rows as A."
    End If
    Solve = MatrixMultDefined(invA, b)
End Function

Public Function Cholesky(ByRef a As Variant) As Variant
    Dim n As Long
    Dim L As Variant
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim acc As Double
    Call RequireSquare(a, "Cholesky")
    If Not IsSymmetric(a, 0.000000001) Then
        Err.Raise 5, "Cholesky", "Matrix must be symmetric."
    End If
    n = RowsOf(a)
    ReDim L(1 To n, 1 To n)
    For i = 1 To n
        For j = 1 To i
            acc = CDbl(a(i, j))
            For k = 1 To j - 1
                acc = acc - CDbl(L(i, k)) * CDbl(L(j, k))
            Next k
            If i = j Then
                If acc <= 1E-14 Then
                    Err.Raise 5, "Cholesky", "Matrix is not positive definite."
                End If
                L(i, j) = Sqr(acc)
            Else
                L(i, j) = acc / CDbl(L(j, j))
            End If
        Next j
    Next i
    Cholesky = L
End Function

' Symmetric eigen: Jacobi rotations. Returns n x (n+1): eigenvectors | eigenvalues column.
Public Function EigenSymmetric(ByRef a As Variant) As Variant
    Dim n As Long
    Dim v As Variant
    Dim d As Variant
    Dim i As Long
    Dim j As Long
    Dim p As Long
    Dim q As Long
    Dim iter As Long
    Dim maxOff As Double
    Dim app As Double
    Dim aqq As Double
    Dim apq As Double
    Dim theta As Double
    Dim t As Double
    Dim c As Double
    Dim s As Double
    Dim tau As Double
    Dim k As Long
    Dim out As Variant
    Const maxIter As Long = 100
    Call RequireSquare(a, "Eigen")
    If Not IsSymmetric(a, 0.000000001) Then
        Err.Raise 5, "Eigen", "Eigen uses the Jacobi method and needs a symmetric matrix."
    End If
    n = RowsOf(a)
    d = CopyMat(a)
    v = Identity(n)
    For iter = 1 To maxIter * n * n
        maxOff = 0
        p = 1
        q = 2
        If n = 1 Then Exit For
        For i = 1 To n
            For j = i + 1 To n
                If Abs(CDbl(d(i, j))) > maxOff Then
                    maxOff = Abs(CDbl(d(i, j)))
                    p = i
                    q = j
                End If
            Next j
        Next i
        If maxOff < 0.000000001 Then Exit For
        app = CDbl(d(p, p))
        aqq = CDbl(d(q, q))
        apq = CDbl(d(p, q))
        If Abs(apq) < 1E-18 Then
            t = 0
        Else
            theta = (aqq - app) / (2 * apq)
            If theta >= 0 Then
                t = 1 / (theta + Sqr(1 + theta * theta))
            Else
                t = -1 / (-theta + Sqr(1 + theta * theta))
            End If
        End If
        c = 1 / Sqr(1 + t * t)
        s = t * c
        tau = s / (1 + c)
        For k = 1 To n
            If k <> p And k <> q Then
                app = CDbl(d(k, p))
                aqq = CDbl(d(k, q))
                d(k, p) = app - s * (aqq + tau * app)
                d(p, k) = d(k, p)
                d(k, q) = aqq + s * (app - tau * aqq)
                d(q, k) = d(k, q)
            End If
            app = CDbl(v(k, p))
            aqq = CDbl(v(k, q))
            v(k, p) = app - s * (aqq + tau * app)
            v(k, q) = aqq + s * (app - tau * aqq)
        Next k
        app = CDbl(d(p, p))
        aqq = CDbl(d(q, q))
        d(p, p) = app - t * apq
        d(q, q) = aqq + t * apq
        d(p, q) = 0
        d(q, p) = 0
    Next iter
    ReDim out(1 To n, 1 To n + 1)
    For i = 1 To n
        For j = 1 To n
            out(i, j) = v(i, j)
        Next j
        out(i, n + 1) = d(i, i)
    Next i
    EigenSymmetric = out
End Function

' n x 1 column of eigenvalues from Jacobi (symmetric matrices).
Public Function EigenvaluesColumn(ByRef a As Variant) As Variant
    Dim packed As Variant
    Dim n As Long
    Dim i As Long
    Dim out As Variant
    packed = EigenSymmetric(a)
    n = RowsOf(packed)
    ReDim out(1 To n, 1 To 1)
    For i = 1 To n
        out(i, 1) = packed(i, n + 1)
    Next i
    EigenvaluesColumn = out
End Function

' n x n matrix whose columns are eigenvectors from Jacobi (symmetric matrices).
Public Function EigenvectorsOf(ByRef a As Variant) As Variant
    Dim packed As Variant
    Dim n As Long
    Dim i As Long
    Dim j As Long
    Dim out As Variant
    packed = EigenSymmetric(a)
    n = RowsOf(packed)
    ReDim out(1 To n, 1 To n)
    For i = 1 To n
        For j = 1 To n
            out(i, j) = packed(i, j)
        Next j
    Next i
    EigenvectorsOf = out
End Function

' Infinity-norm condition number ||A||_inf * ||A^{-1}||_inf (Excel MCOND).
Public Function ConditionNumberInf(ByRef a As Variant) As Double
    Call RequireSquare(a, "Condition")
    ConditionNumberInf = NormInf(a) * NormInf(Inverse(a))
End Function

' Spectral radius max |lambda|. Exact for symmetric matrices; power iteration otherwise.
Public Function SpectralRadiusOf(ByRef a As Variant) As Double
    Dim evals As Variant
    Dim i As Long
    Dim best As Double
    Call RequireSquare(a, "SpectralRadius")
    If IsSymmetric(a, 0.000000001) Then
        evals = EigenvaluesColumn(a)
        best = 0
        For i = 1 To RowsOf(evals)
            If Abs(CDbl(evals(i, 1))) > best Then best = Abs(CDbl(evals(i, 1)))
        Next i
        SpectralRadiusOf = best
        Exit Function
    End If
    SpectralRadiusOf = SpectralRadiusPower(a)
End Function

Private Function SpectralRadiusPower(ByRef a As Variant) As Double
    Dim n As Long
    Dim i As Long
    Dim iter As Long
    Dim x As Variant
    Dim y As Variant
    Dim nrm As Double
    Dim ray As Double
    Dim prev As Double
    Const maxIter As Long = 200
    Const tol As Double = 0.000000001
    n = RowsOf(a)
    ReDim x(1 To n, 1 To 1)
    For i = 1 To n
        x(i, 1) = 1
    Next i
    prev = 0
    For iter = 1 To maxIter
        y = MatrixMultDefined(a, x)
        nrm = 0
        For i = 1 To n
            If Abs(CDbl(y(i, 1))) > nrm Then nrm = Abs(CDbl(y(i, 1)))
        Next i
        If nrm < 1E-18 Then
            SpectralRadiusPower = 0
            Exit Function
        End If
        For i = 1 To n
            x(i, 1) = CDbl(y(i, 1)) / nrm
        Next i
        y = MatrixMultDefined(a, x)
        ray = 0
        nrm = 0
        For i = 1 To n
            ray = ray + CDbl(x(i, 1)) * CDbl(y(i, 1))
            nrm = nrm + CDbl(x(i, 1)) * CDbl(x(i, 1))
        Next i
        If nrm > 0 Then ray = ray / nrm
        If iter > 5 And Abs(Abs(ray) - prev) < tol Then Exit For
        prev = Abs(ray)
    Next iter
    SpectralRadiusPower = Abs(ray)
End Function

' Thin QR via modified Gram-Schmidt. Returns [Q | R] stacked as Q then R below? Better two outputs.
' Returns n x (n+m) = Q then R on the right if square. For n x m, Q is n x m, R is m x m.
Public Function QRFactors(ByRef a As Variant) As Variant
    Dim n As Long
    Dim m As Long
    Dim Q As Variant
    Dim R As Variant
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim nrm As Double
    Dim acc As Double
    Dim out As Variant
    n = RowsOf(a)
    m = ColsOf(a)
    If m > n Then
        Err.Raise 5, "QR", "QR needs at least as many rows as columns."
    End If
    Q = CopyMat(a)
    ReDim R(1 To m, 1 To m)
    For j = 1 To m
        nrm = 0
        For i = 1 To n
            nrm = nrm + CDbl(Q(i, j)) * CDbl(Q(i, j))
        Next i
        nrm = Sqr(nrm)
        If nrm < 1E-14 Then
            Err.Raise 5, "QR", "Columns are linearly dependent."
        End If
        R(j, j) = nrm
        For i = 1 To n
            Q(i, j) = CDbl(Q(i, j)) / nrm
        Next i
        For k = j + 1 To m
            acc = 0
            For i = 1 To n
                acc = acc + CDbl(Q(i, j)) * CDbl(Q(i, k))
            Next i
            R(j, k) = acc
            For i = 1 To n
                Q(i, k) = CDbl(Q(i, k)) - acc * CDbl(Q(i, j))
            Next i
        Next k
    Next j
    ReDim out(1 To n + m, 1 To m)
    For i = 1 To n
        For j = 1 To m
            out(i, j) = Q(i, j)
        Next j
    Next i
    For i = 1 To m
        For j = 1 To m
            out(n + i, j) = R(i, j)
        Next j
    Next i
    QRFactors = out
End Function

Public Function ScalarMatrix(ByVal value As Double) As Variant
    Dim a(1 To 1, 1 To 1) As Double
    a(1, 1) = value
    ScalarMatrix = a
End Function

Public Function ScaleMat(ByRef a As Variant, ByVal k As Double) As Variant
    Dim nr As Long
    Dim nc As Long
    Dim b As Variant
    Dim i As Long
    Dim j As Long
    nr = RowsOf(a)
    nc = ColsOf(a)
    ReDim b(1 To nr, 1 To nc)
    For i = 1 To nr
        For j = 1 To nc
            b(i, j) = k * CDbl(a(i, j))
        Next j
    Next i
    ScaleMat = b
End Function

Public Function MatPower(ByRef a As Variant, ByVal p As Long) As Variant
    Dim i As Long
    Dim acc As Variant
    Call RequireSquare(a, "MatPower")
    If p = 0 Then
        MatPower = Identity(RowsOf(a))
        Exit Function
    End If
    If p < 0 Then
        acc = Inverse(a)
        p = -p
    Else
        acc = CopyMat(a)
    End If
    MatPower = Identity(RowsOf(a))
    For i = 1 To p
        MatPower = MatrixMultDefined(MatPower, acc)
    Next i
End Function

Public Function RankOf(ByRef a As Variant) As Long
    Dim lu As Variant
    Dim n As Long
    Dim m As Long
    Dim k As Long
    Dim i As Long
    Dim j As Long
    Dim p As Long
    Dim r As Long
    Dim maxAbs As Double
    Dim tmp As Double
    Dim tol As Double

    lu = CopyMat(a)
    n = RowsOf(lu)
    m = ColsOf(lu)
    If n < m Then
        lu = TransposeMatrix(lu)
        n = RowsOf(lu)
        m = ColsOf(lu)
    End If

    tol = 0.000000001
    r = 0
    For k = 1 To m
        p = r + 1
        maxAbs = 0
        For i = r + 1 To n
            If Abs(CDbl(lu(i, k))) > maxAbs Then
                maxAbs = Abs(CDbl(lu(i, k)))
                p = i
            End If
        Next i
        If maxAbs <= tol Then GoTo NextCol
        r = r + 1
        If p <> r Then
            For j = 1 To m
                tmp = lu(r, j)
                lu(r, j) = lu(p, j)
                lu(p, j) = tmp
            Next j
        End If
        For i = r + 1 To n
            lu(i, k) = CDbl(lu(i, k)) / CDbl(lu(r, k))
            For j = k + 1 To m
                lu(i, j) = CDbl(lu(i, j)) - CDbl(lu(i, k)) * CDbl(lu(r, j))
            Next j
        Next i
NextCol:
    Next k
    RankOf = r
End Function

Public Function ExtractDiag(ByRef a As Variant) As Variant
    Dim n As Long
    Dim i As Long
    Dim b As Variant
    n = RowsOf(a)
    If ColsOf(a) < n Then n = ColsOf(a)
    ReDim b(1 To n, 1 To 1)
    For i = 1 To n
        b(i, 1) = CDbl(a(i, i))
    Next i
    ExtractDiag = b
End Function

Public Function OuterProduct(ByRef u As Variant, ByRef v As Variant) As Variant
    Dim colU As Variant
    Dim colV As Variant
    colU = AsColumn(u)
    colV = AsColumn(v)
    OuterProduct = MatrixMultDefined(colU, TransposeMatrix(colV))
End Function

Public Function DotProduct(ByRef u As Variant, ByRef v As Variant) As Double
    Dim colU As Variant
    Dim colV As Variant
    Dim i As Long
    Dim acc As Double
    colU = AsColumn(u)
    colV = AsColumn(v)
    If RowsOf(colU) <> RowsOf(colV) Then
        Err.Raise 5, "DotProduct", "Vectors must have the same length."
    End If
    acc = 0
    For i = 1 To RowsOf(colU)
        acc = acc + CDbl(colU(i, 1)) * CDbl(colV(i, 1))
    Next i
    DotProduct = acc
End Function

Public Function Norm1(ByRef a As Variant) As Double
    Dim j As Long
    Dim i As Long
    Dim colSum As Double
    Dim best As Double
    best = 0
    For j = 1 To ColsOf(a)
        colSum = 0
        For i = 1 To RowsOf(a)
            colSum = colSum + Abs(CDbl(a(i, j)))
        Next i
        If colSum > best Then best = colSum
    Next j
    Norm1 = best
End Function

Public Function NormInf(ByRef a As Variant) As Double
    Dim i As Long
    Dim j As Long
    Dim rowSum As Double
    Dim best As Double
    best = 0
    For i = 1 To RowsOf(a)
        rowSum = 0
        For j = 1 To ColsOf(a)
            rowSum = rowSum + Abs(CDbl(a(i, j)))
        Next j
        If rowSum > best Then best = rowSum
    Next i
    NormInf = best
End Function

Public Function ExchangeMat(ByVal n As Long) As Variant
    Dim a As Variant
    Dim i As Long
    Dim j As Long
    Call CheckSize(n, n)
    ReDim a(1 To n, 1 To n)
    For i = 1 To n
        For j = 1 To n
            a(i, j) = 0
        Next j
        a(i, n - i + 1) = 1
    Next i
    ExchangeMat = a
End Function

Public Function ToeplitzFromVector(ByRef vec As Variant) As Variant
    Dim col As Variant
    Dim n As Long
    Dim a As Variant
    Dim i As Long
    Dim j As Long
    col = AsColumn(vec)
    n = RowsOf(col)
    Call CheckSize(n, n)
    ReDim a(1 To n, 1 To n)
    For i = 1 To n
        For j = 1 To n
            a(i, j) = CDbl(col(Abs(i - j) + 1, 1))
        Next j
    Next i
    ToeplitzFromVector = a
End Function

Public Function VandermondeFromVector(ByRef vec As Variant, Optional ByVal nCols As Long = 0) As Variant
    Dim col As Variant
    Dim n As Long
    Dim m As Long
    Dim a As Variant
    Dim i As Long
    Dim j As Long
    Dim x As Double
    col = AsColumn(vec)
    n = RowsOf(col)
    If nCols <= 0 Then
        m = n
    Else
        m = nCols
    End If
    Call CheckSize(n, m)
    ReDim a(1 To n, 1 To m)
    For i = 1 To n
        x = 1
        For j = 1 To m
            a(i, j) = x
            x = x * CDbl(col(i, 1))
        Next j
    Next i
    VandermondeFromVector = a
End Function

Public Function LUFactors(ByRef a As Variant) As Variant
    Dim lu As Variant
    Dim piv() As Long
    Dim n As Long
    Dim L As Variant
    Dim U As Variant
    Dim out As Variant
    Dim i As Long
    Dim j As Long
    Call RequireSquare(a, "LU")
    n = RowsOf(a)
    lu = CopyMat(a)
    If LUFactor(lu, piv) = 0 Then
        Err.Raise 5, "LU", "Matrix is singular."
    End If
    ReDim L(1 To n, 1 To n)
    ReDim U(1 To n, 1 To n)
    For i = 1 To n
        For j = 1 To n
            If i > j Then
                L(i, j) = CDbl(lu(i, j))
                U(i, j) = 0
            ElseIf i = j Then
                L(i, j) = 1
                U(i, j) = CDbl(lu(i, j))
            Else
                L(i, j) = 0
                U(i, j) = CDbl(lu(i, j))
            End If
        Next j
    Next i
    ReDim out(1 To 2 * n, 1 To n)
    For i = 1 To n
        For j = 1 To n
            out(i, j) = L(i, j)
            out(n + i, j) = U(i, j)
        Next j
    Next i
    LUFactors = out
End Function

Public Function Adjugate(ByRef a As Variant) As Variant
    Adjugate = ScaleMat(Inverse(a), Determinant(a))
End Function

Public Function PseudoInverse(ByRef a As Variant) As Variant
    Dim at As Variant
    Dim ata As Variant
    Dim aat As Variant
    at = TransposeMatrix(a)
    If RowsOf(a) >= ColsOf(a) Then
        ata = MatrixMultDefined(at, a)
        PseudoInverse = MatrixMultDefined(Inverse(ata), at)
    Else
        aat = MatrixMultDefined(a, at)
        PseudoInverse = MatrixMultDefined(at, Inverse(aat))
    End If
End Function

' Column-major vec(A): stack columns into one column (Personal Matrices1).
Public Function Vectorize(ByRef a As Variant) As Variant
    Dim nr As Long
    Dim nc As Long
    Dim b As Variant
    Dim i As Long
    Dim j As Long
    Dim k As Long
    nr = RowsOf(a)
    nc = ColsOf(a)
    Call CheckSize(nr * nc, 1)
    ReDim b(1 To nr * nc, 1 To 1)
    k = 1
    For j = 1 To nc
        For i = 1 To nr
            b(k, 1) = CDbl(a(i, j))
            k = k + 1
        Next i
    Next j
    Vectorize = b
End Function

Public Function Unvectorize(ByRef vec As Variant, ByVal nRows As Long) As Variant
    Dim col As Variant
    Dim n As Long
    Dim nCols As Long
    Dim a As Variant
    Dim i As Long
    Dim j As Long
    Dim k As Long
    col = AsColumn(vec)
    n = RowsOf(col)
    If nRows < 1 Then
        Err.Raise 5, "Unvectorize", "Number of rows must be at least 1."
    End If
    If n Mod nRows <> 0 Then
        Err.Raise 5, "Unvectorize", "Vector length must be a multiple of the row count."
    End If
    nCols = n / nRows
    Call CheckSize(nRows, nCols)
    ReDim a(1 To nRows, 1 To nCols)
    k = 1
    For j = 1 To nCols
        For i = 1 To nRows
            a(i, j) = CDbl(col(k, 1))
            k = k + 1
        Next i
    Next j
    Unvectorize = a
End Function

' Companion of x^n + c_n x^{n-1} + ... + c_1, with the selected vector = (c_1 .. c_n)
' (constant term first). Result is n x n.
Public Function CompanionFromVector(ByRef vec As Variant) As Variant
    Dim col As Variant
    Dim n As Long
    Dim a As Variant
    Dim i As Long
    Dim j As Long
    col = AsColumn(vec)
    n = RowsOf(col)
    Call CheckSize(n, n)
    ReDim a(1 To n, 1 To n)
    For i = 1 To n
        For j = 1 To n
            a(i, j) = 0
        Next j
    Next i
    For i = 2 To n
        a(i, i - 1) = 1
    Next i
    For i = 1 To n
        a(i, n) = -CDbl(col(i, 1))
    Next i
    CompanionFromVector = a
End Function

Public Function MinorMatrix(ByRef a As Variant, ByVal skipRow As Long, ByVal skipCol As Long) As Variant
    Dim n As Long
    Dim b As Variant
    Dim i As Long
    Dim j As Long
    Dim ri As Long
    Dim cj As Long
    Call RequireSquare(a, "Minor")
    n = RowsOf(a)
    If n < 2 Then
        Err.Raise 5, "Minor", "Minor needs a matrix of order 2 or more."
    End If
    If skipRow < 1 Or skipRow > n Or skipCol < 1 Or skipCol > n Then
        Err.Raise 5, "Minor", "Row and column indices must be between 1 and n."
    End If
    ReDim b(1 To n - 1, 1 To n - 1)
    ri = 1
    For i = 1 To n
        If i <> skipRow Then
            cj = 1
            For j = 1 To n
                If j <> skipCol Then
                    b(ri, cj) = CDbl(a(i, j))
                    cj = cj + 1
                End If
            Next j
            ri = ri + 1
        End If
    Next i
    MinorMatrix = b
End Function

Public Function CofactorMatrix(ByRef a As Variant) As Variant
    Dim n As Long
    Dim c As Variant
    Dim i As Long
    Dim j As Long
    Dim sgn As Long
    Call RequireSquare(a, "Cofactor")
    n = RowsOf(a)
    If n > CofactorMaxN Then
        Err.Raise 5, "Cofactor", "Cofactor matrix is limited to " & CofactorMaxN & " x " & CofactorMaxN & "."
    End If
    If n = 1 Then
        CofactorMatrix = ScalarMatrix(1)
        Exit Function
    End If
    ReDim c(1 To n, 1 To n)
    For i = 1 To n
        For j = 1 To n
            If ((i + j) Mod 2) = 0 Then sgn = 1 Else sgn = -1
            c(i, j) = sgn * Determinant(MinorMatrix(a, i, j))
        Next j
    Next i
    CofactorMatrix = c
End Function

Public Function IsOrthogonal(ByRef a As Variant, Optional ByVal tol As Double = 0.000000001) As Boolean
    Dim at As Variant
    Dim p As Variant
    Dim i As Long
    Dim j As Long
    Dim expect As Double
    Call RequireSquare(a, "IsOrthogonal")
    at = TransposeMatrix(a)
    p = MatrixMultDefined(at, a)
    For i = 1 To RowsOf(p)
        For j = 1 To ColsOf(p)
            If i = j Then expect = 1 Else expect = 0
            If Abs(CDbl(p(i, j)) - expect) > tol Then
                IsOrthogonal = False
                Exit Function
            End If
        Next j
    Next i
    IsOrthogonal = True
End Function

' Square H is Hadamard when H H^T = n I (n = order).
Public Function IsHadamard(ByRef a As Variant, Optional ByVal tol As Double = 0.000000001) As Boolean
    Dim at As Variant
    Dim p As Variant
    Dim n As Long
    Dim i As Long
    Dim j As Long
    Dim expect As Double
    If RowsOf(a) <> ColsOf(a) Then
        IsHadamard = False
        Exit Function
    End If
    n = RowsOf(a)
    at = TransposeMatrix(a)
    p = MatrixMultDefined(a, at)
    For i = 1 To n
        For j = 1 To n
            If i = j Then expect = CDbl(n) Else expect = 0
            If Abs(CDbl(p(i, j)) - expect) > tol Then
                IsHadamard = False
                Exit Function
            End If
        Next j
    Next i
    IsHadamard = True
End Function

' TRUE when every entry is +1 or -1 (within tol).
Public Function HasPlusMinusOneEntries(ByRef a As Variant, Optional ByVal tol As Double = 0.000000001) As Boolean
    Dim i As Long
    Dim j As Long
    Dim v As Double
    For i = 1 To RowsOf(a)
        For j = 1 To ColsOf(a)
            v = CDbl(a(i, j))
            If Abs(v - 1) > tol And Abs(v + 1) > tol Then
                HasPlusMinusOneEntries = False
                Exit Function
            End If
        Next j
    Next i
    HasPlusMinusOneEntries = True
End Function

Private Function AsColumn(ByRef a As Variant) As Variant
    Dim n As Long
    Dim b As Variant
    Dim i As Long
    If ColsOf(a) = 1 Then
        AsColumn = CopyMat(a)
        Exit Function
    End If
    If RowsOf(a) = 1 Then
        n = ColsOf(a)
        ReDim b(1 To n, 1 To 1)
        For i = 1 To n
            b(i, 1) = CDbl(a(1, i))
        Next i
        AsColumn = b
        Exit Function
    End If
    Err.Raise 5, "AsColumn", "Select a single row or column vector."
End Function

Private Function ToFiniteDouble(ByVal v As Variant, ByVal r As Long, ByVal c As Long) As Double
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then
        Err.Raise 5, "RangeToMatrix", "Cell (" & r & "," & c & ") is empty or an error."
    End If
    If VarType(v) = vbString Then
        If Len(Trim$(CStr(v))) = 0 Then
            Err.Raise 5, "RangeToMatrix", "Cell (" & r & "," & c & ") is empty."
        End If
    End If
    If Not IsNumeric(v) Then
        Err.Raise 5, "RangeToMatrix", "Cell (" & r & "," & c & ") is not numeric."
    End If
    ToFiniteDouble = CDbl(v)
End Function

Private Sub RequireSquare(ByRef a As Variant, ByVal opName As String)
    If RowsOf(a) <> ColsOf(a) Then
        Err.Raise 5, opName, opName & " needs a square matrix."
    End If
    Call CheckSize(RowsOf(a), ColsOf(a))
End Sub

Private Function CopyMat(ByRef a As Variant) As Variant
    Dim nr As Long
    Dim nc As Long
    Dim b As Variant
    Dim i As Long
    Dim j As Long
    nr = RowsOf(a)
    nc = ColsOf(a)
    ReDim b(1 To nr, 1 To nc)
    For i = 1 To nr
        For j = 1 To nc
            b(i, j) = CDbl(a(i, j))
        Next j
    Next i
    CopyMat = b
End Function

Private Function SliceCols(ByRef a As Variant, ByVal c1 As Long, ByVal c2 As Long) As Variant
    Dim nr As Long
    Dim b As Variant
    Dim i As Long
    Dim j As Long
    nr = RowsOf(a)
    ReDim b(1 To nr, 1 To c2 - c1 + 1)
    For i = 1 To nr
        For j = c1 To c2
            b(i, j - c1 + 1) = a(i, j)
        Next j
    Next i
    SliceCols = b
End Function

Private Function ProductDiag(ByRef a As Variant) As Double
    Dim i As Long
    Dim acc As Double
    acc = 1
    For i = LBound(a, 1) To UBound(a, 1)
        acc = acc * CDbl(a(i, i))
    Next i
    ProductDiag = acc
End Function

' Returns sign of permutation, or 0 if singular. lu is replaced in place.
Private Function LUFactor(ByRef lu As Variant, ByRef piv() As Long) As Long
    Dim n As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim p As Long
    Dim maxAbs As Double
    Dim tmp As Double
    Dim sign As Long
    n = RowsOf(lu)
    ReDim piv(1 To n)
    For i = 1 To n
        piv(i) = i
    Next i
    sign = 1
    For k = 1 To n
        p = k
        maxAbs = Abs(CDbl(lu(k, k)))
        For i = k + 1 To n
            If Abs(CDbl(lu(i, k))) > maxAbs Then
                maxAbs = Abs(CDbl(lu(i, k)))
                p = i
            End If
        Next i
        If maxAbs < 1E-14 Then
            LUFactor = 0
            Exit Function
        End If
        If p <> k Then
            For j = 1 To n
                tmp = lu(k, j)
                lu(k, j) = lu(p, j)
                lu(p, j) = tmp
            Next j
            i = piv(k)
            piv(k) = piv(p)
            piv(p) = i
            sign = -sign
        End If
        For i = k + 1 To n
            lu(i, k) = CDbl(lu(i, k)) / CDbl(lu(k, k))
            For j = k + 1 To n
                lu(i, j) = CDbl(lu(i, j)) - CDbl(lu(i, k)) * CDbl(lu(k, j))
            Next j
        Next i
    Next k
    LUFactor = sign
End Function
