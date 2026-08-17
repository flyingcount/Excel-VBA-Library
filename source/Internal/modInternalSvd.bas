Attribute VB_Name = "modInternalSvd"
Option Explicit

' Internal: Golub-Reinsch SVD (Personal Custom_Menu18_SVD / Numerical Recipes svdcmp).
' A = U * diag(w) * V^T. U overwrites A (m x n). w is 1-based length n. V is n x n.

Public Const SvdMaxN As Long = 500

Public Sub Decompose(ByRef a As Variant, ByVal m As Long, ByVal n As Long, ByRef w As Variant, ByRef v As Variant)
    Dim rv1() As Double
    Dim i As Long
    Dim its As Long
    Dim j As Long
    Dim jj As Long
    Dim k As Long
    Dim L As Long
    Dim nm As Long
    Dim Anorm As Double
    Dim c As Double
    Dim f As Double
    Dim g As Double
    Dim h As Double
    Dim s As Double
    Dim scale_ As Double
    Dim x As Double
    Dim y As Double
    Dim z As Double
    Dim converged As Boolean

    If m < 1 Or n < 1 Then Err.Raise 5, "SVD", "Matrix must have at least one row and one column."
    If m > SvdMaxN Or n > SvdMaxN Then Err.Raise 5, "SVD", "Matrices larger than " & SvdMaxN & " on a side are not supported."

    ReDim w(1 To n)
    ReDim v(1 To n, 1 To n)
    ReDim rv1(1 To n)
    g = 0#
    scale_ = 0#
    Anorm = 0#

    For i = 1 To n
        L = i + 1
        rv1(i) = scale_ * g
        g = 0#
        s = 0#
        scale_ = 0#
        If i <= m Then
            For k = i To m
                scale_ = scale_ + Abs(CDbl(a(k, i)))
            Next k
            If scale_ <> 0# Then
                For k = i To m
                    a(k, i) = CDbl(a(k, i)) / scale_
                    s = s + CDbl(a(k, i)) * CDbl(a(k, i))
                Next k
                f = CDbl(a(i, i))
                g = -SvdSign(Sqr(s), f)
                h = f * g - s
                a(i, i) = f - g
                For j = L To n
                    s = 0#
                    For k = i To m
                        s = s + CDbl(a(k, i)) * CDbl(a(k, j))
                    Next k
                    f = s / h
                    For k = i To m
                        a(k, j) = CDbl(a(k, j)) + f * CDbl(a(k, i))
                    Next k
                Next j
                For k = i To m
                    a(k, i) = scale_ * CDbl(a(k, i))
                Next k
            End If
        End If
        w(i) = scale_ * g
        g = 0#
        s = 0#
        scale_ = 0#
        If i <= m And i <> n Then
            For k = L To n
                scale_ = scale_ + Abs(CDbl(a(i, k)))
            Next k
            If scale_ <> 0# Then
                For k = L To n
                    a(i, k) = CDbl(a(i, k)) / scale_
                    s = s + CDbl(a(i, k)) * CDbl(a(i, k))
                Next k
                f = CDbl(a(i, L))
                g = -SvdSign(Sqr(s), f)
                h = f * g - s
                a(i, L) = f - g
                For k = L To n
                    rv1(k) = CDbl(a(i, k)) / h
                Next k
                For j = L To m
                    s = 0#
                    For k = L To n
                        s = s + CDbl(a(j, k)) * CDbl(a(i, k))
                    Next k
                    For k = L To n
                        a(j, k) = CDbl(a(j, k)) + s * rv1(k)
                    Next k
                Next j
                For k = L To n
                    a(i, k) = scale_ * CDbl(a(i, k))
                Next k
            End If
        End If
        Anorm = SvdMax(Anorm, Abs(CDbl(w(i))) + Abs(rv1(i)))
    Next i

    For i = n To 1 Step -1
        If i < n Then
            If g <> 0# Then
                For j = L To n
                    v(j, i) = (CDbl(a(i, j)) / CDbl(a(i, L))) / g
                Next j
                For j = L To n
                    s = 0#
                    For k = L To n
                        s = s + CDbl(a(i, k)) * CDbl(v(k, j))
                    Next k
                    For k = L To n
                        v(k, j) = CDbl(v(k, j)) + s * CDbl(v(k, i))
                    Next k
                Next j
            End If
            For j = L To n
                v(i, j) = 0#
                v(j, i) = 0#
            Next j
        End If
        v(i, i) = 1#
        g = rv1(i)
        L = i
    Next i

    For i = SvdMin(m, n) To 1 Step -1
        L = i + 1
        g = CDbl(w(i))
        For j = L To n
            a(i, j) = 0#
        Next j
        If g <> 0# Then
            g = 1# / g
            For j = L To n
                s = 0#
                For k = L To m
                    s = s + CDbl(a(k, i)) * CDbl(a(k, j))
                Next k
                f = (s / CDbl(a(i, i))) * g
                For k = i To m
                    a(k, j) = CDbl(a(k, j)) + f * CDbl(a(k, i))
                Next k
            Next j
            For j = i To m
                a(j, i) = CDbl(a(j, i)) * g
            Next j
        Else
            For j = i To m
                a(j, i) = 0#
            Next j
        End If
        a(i, i) = CDbl(a(i, i)) + 1#
    Next i

    For k = n To 1 Step -1
        converged = False
        For its = 1 To 30
            For L = k To 1 Step -1
                nm = L - 1
                If (Abs(rv1(L)) + Anorm) = Anorm Then GoTo SplitOk
                If nm >= 1 Then
                    If (Abs(CDbl(w(nm))) + Anorm) = Anorm Then GoTo CancelLeft
                End If
            Next L
CancelLeft:
            c = 0#
            s = 1#
            For i = L To k
                f = s * rv1(i)
                rv1(i) = c * rv1(i)
                If (Abs(f) + Anorm) = Anorm Then GoTo SplitOk
                g = CDbl(w(i))
                h = Pythag(f, g)
                w(i) = h
                h = 1# / h
                c = g * h
                s = -f * h
                For j = 1 To m
                    y = CDbl(a(j, nm))
                    z = CDbl(a(j, i))
                    a(j, nm) = (y * c) + (z * s)
                    a(j, i) = -(y * s) + (z * c)
                Next j
            Next i
SplitOk:
            z = CDbl(w(k))
            If L = k Then
                If z < 0# Then
                    w(k) = -z
                    For j = 1 To n
                        v(j, k) = -CDbl(v(j, k))
                    Next j
                End If
                converged = True
                Exit For
            End If
            If its = 30 Then Err.Raise 5, "SVD", "SVD did not converge in 30 iterations."
            x = CDbl(w(L))
            nm = k - 1
            y = CDbl(w(nm))
            g = rv1(nm)
            h = rv1(k)
            f = ((y - z) * (y + z) + (g - h) * (g + h)) / (2# * h * y)
            g = Pythag(f, 1#)
            f = ((x - z) * (x + z) + h * ((y / (f + SvdSign(g, f))) - h)) / x
            c = 1#
            s = 1#
            For j = L To nm
                i = j + 1
                g = rv1(i)
                y = CDbl(w(i))
                h = s * g
                g = c * g
                z = Pythag(f, h)
                rv1(j) = z
                c = f / z
                s = h / z
                f = (x * c) + (g * s)
                g = -(x * s) + (g * c)
                h = y * s
                y = y * c
                For jj = 1 To n
                    x = CDbl(v(jj, j))
                    z = CDbl(v(jj, i))
                    v(jj, j) = (x * c) + (z * s)
                    v(jj, i) = -(x * s) + (z * c)
                Next jj
                z = Pythag(f, h)
                w(j) = z
                If z <> 0# Then
                    z = 1# / z
                    c = f * z
                    s = h * z
                End If
                f = (c * g) + (s * y)
                x = -(s * g) + (c * y)
                For jj = 1 To m
                    y = CDbl(a(jj, j))
                    z = CDbl(a(jj, i))
                    a(jj, j) = (y * c) + (z * s)
                    a(jj, i) = -(y * s) + (z * c)
                Next jj
            Next j
            rv1(L) = 0#
            rv1(k) = f
            w(k) = x
        Next its
        If Not converged Then Err.Raise 5, "SVD", "SVD did not converge in 30 iterations."
    Next k
    Call SortDescending(a, v, w)
End Sub

Public Sub WriteBlocks(ByVal dest As Range, ByRef u As Variant, ByRef w As Variant, ByRef v As Variant, ByVal nRows As Long, ByVal nCols As Long)
    Dim p As Long
    Dim r As Long
    Dim wRow As Variant
    Dim i As Long
    If nRows < nCols Then p = nRows Else p = nCols
    ReDim wRow(1 To 1, 1 To p)
    For i = 1 To p
        wRow(1, i) = w(i)
    Next i
    dest.Value = "Singular Values"
    dest.Font.Bold = True
    dest.Offset(1, 0).Resize(1, p).Value = wRow
    r = 4
    dest.Offset(r - 1, 0).Value = "Left matrix U"
    dest.Offset(r - 1, 0).Font.Bold = True
    dest.Offset(r, 0).Resize(nRows, p).Value = SliceLeft(u, nRows, p)
    r = r + nRows + 2
    dest.Offset(r - 1, 0).Value = "Right matrix V"
    dest.Offset(r - 1, 0).Font.Bold = True
    dest.Offset(r, 0).Resize(nCols, p).Value = SliceLeft(v, nCols, p)
End Sub

Private Function SliceLeft(ByRef mat As Variant, ByVal nRows As Long, ByVal nCols As Long) As Variant
    Dim out As Variant
    Dim r As Long
    Dim c As Long
    ReDim out(1 To nRows, 1 To nCols)
    For r = 1 To nRows
        For c = 1 To nCols
            out(r, c) = mat(r, c)
        Next c
    Next r
    SliceLeft = out
End Function

Private Sub SortDescending(ByRef a As Variant, ByRef v As Variant, ByRef w As Variant)
    Dim swapped As Boolean
    Dim iMin As Long
    Dim iMax As Long
    Dim i As Long
    Dim k As Long
    Dim tmp As Double
    iMin = LBound(w)
    iMax = UBound(w)
    Do
        swapped = False
        For i = iMin To iMax Step 2
            k = i + 1
            If k > iMax Then Exit For
            If Abs(CDbl(w(i))) < Abs(CDbl(w(k))) Then
                tmp = CDbl(w(k))
                w(k) = w(i)
                w(i) = tmp
                Call SwapCol(a, k, i)
                Call SwapCol(v, k, i)
                swapped = True
            End If
        Next i
        If iMin = LBound(w) Then
            iMin = LBound(w) + 1
        Else
            iMin = LBound(w)
        End If
    Loop Until (Not swapped) And iMin = LBound(w)
End Sub

Private Sub SwapCol(ByRef a As Variant, ByVal k As Long, ByVal j As Long)
    Dim i As Long
    Dim tmp As Variant
    For i = 1 To UBound(a, 1)
        tmp = a(i, j)
        a(i, j) = a(i, k)
        a(i, k) = tmp
    Next i
End Sub

Private Function SvdSign(ByVal x As Double, ByVal y As Double) As Double
    If y >= 0 Then SvdSign = Abs(x) Else SvdSign = -Abs(x)
End Function

Private Function SvdMax(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then SvdMax = a Else SvdMax = b
End Function

Private Function SvdMin(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then SvdMin = a Else SvdMin = b
End Function

Private Function Pythag(ByVal x As Double, ByVal y As Double) As Double
    Dim absa As Double
    Dim absb As Double
    absa = Abs(x)
    absb = Abs(y)
    If absa > absb Then
        Pythag = absa * Sqr(1# + (absb / absa) ^ 2)
    ElseIf absb = 0# Then
        Pythag = 0#
    Else
        Pythag = absb * Sqr(1# + (absa / absb) ^ 2)
    End If
End Function
