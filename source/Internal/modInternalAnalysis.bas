Attribute VB_Name = "modInternalAnalysis"
Option Explicit

' Internal: column statistics, covariance, and correlation (Personal Menu18 Analysis / Covariance).
' Called from modApiAnalysis / modApiCovariance. Do not document as the external API.
' Body arrays are 1-based 2D. Covariance uses observations in rows, variables in columns.

Public Function PromptRange(ByVal PromptText As String, Optional ByVal DefaultRng As Range) As Range
    Dim rng As Range
    If DefaultRng Is Nothing Then
        If TypeName(Selection) = "Range" Then Set DefaultRng = Selection
    End If
    On Error Resume Next
    If DefaultRng Is Nothing Then
        Set rng = Application.InputBox(Prompt:=PromptText, Title:="Analysis", Type:=8)
    Else
        Set rng = Application.InputBox(Prompt:=PromptText, Title:="Analysis", Default:=DefaultRng.Address, Type:=8)
    End If
    On Error GoTo 0
    Set PromptRange = rng
End Function

' Header = first row; body = remaining rows as a 2D array. False = cancel or invalid.
Public Function SplitHeaderBody(ByVal rng As Range, ByRef header As Variant, ByRef body As Variant) As Boolean
    Dim bodyRng As Range
    If rng Is Nothing Then Exit Function
    If rng.Areas.Count > 1 Then
        MsgBox "Select a single contiguous range that includes a header row.", vbExclamation, "Analysis"
        Exit Function
    End If
    If rng.Rows.Count < 2 Then
        MsgBox "Select the header row plus at least one data row.", vbExclamation, "Analysis"
        Exit Function
    End If
    header = AsRowVector(rng.Rows(1).Value)
    Set bodyRng = rng.Offset(1).Resize(rng.Rows.Count - 1, rng.Columns.Count)
    If Not RangeIsAllNumeric(bodyRng) Then
        MsgBox "The data rows (below the header) must be numeric.", vbExclamation, "Analysis"
        Exit Function
    End If
    body = AsMatrix(bodyRng.Value, bodyRng.Rows.Count, bodyRng.Columns.Count)
    SplitHeaderBody = True
End Function

Public Function RangeIsAllNumeric(ByVal rng As Range) As Boolean
    If rng Is Nothing Then Exit Function
    RangeIsAllNumeric = (Application.WorksheetFunction.Count(rng) = rng.Cells.Count)
End Function

Public Function AsMatrix(ByVal raw As Variant, ByVal nRows As Long, ByVal nCols As Long) As Variant
    Dim a As Variant
    Dim r As Long
    Dim c As Long
    ReDim a(1 To nRows, 1 To nCols)
    If nRows = 1 And nCols = 1 Then
        a(1, 1) = CDbl(raw)
    ElseIf nRows = 1 Then
        For c = 1 To nCols
            a(1, c) = CDbl(raw(1, c))
        Next c
    ElseIf nCols = 1 Then
        If IsArray(raw) Then
            If ArrayRank(raw) = 1 Then
                For r = 1 To nRows
                    a(r, 1) = CDbl(raw(r))
                Next r
            Else
                For r = 1 To nRows
                    a(r, 1) = CDbl(raw(r, 1))
                Next r
            End If
        Else
            a(1, 1) = CDbl(raw)
        End If
    Else
        For r = 1 To nRows
            For c = 1 To nCols
                a(r, c) = CDbl(raw(r, c))
            Next c
        Next r
    End If
    AsMatrix = a
End Function

Public Function AsRowVector(ByVal raw As Variant) As Variant
    Dim a As Variant
    Dim c As Long
    Dim n As Long
    If Not IsArray(raw) Then
        ReDim a(1 To 1, 1 To 1)
        a(1, 1) = raw
        AsRowVector = a
        Exit Function
    End If
    If ArrayRank(raw) = 1 Then
        n = UBound(raw) - LBound(raw) + 1
        ReDim a(1 To 1, 1 To n)
        For c = 1 To n
            a(1, c) = raw(LBound(raw) + c - 1)
        Next c
    Else
        n = UBound(raw, 2) - LBound(raw, 2) + 1
        ReDim a(1 To 1, 1 To n)
        For c = 1 To n
            a(1, c) = raw(LBound(raw, 1), LBound(raw, 2) + c - 1)
        Next c
    End If
    AsRowVector = a
End Function

Public Function TransposeRow(ByVal rowVec As Variant) As Variant
    Dim n As Long
    Dim c As Long
    Dim col As Variant
    n = UBound(rowVec, 2)
    ReDim col(1 To n, 1 To 1)
    For c = 1 To n
        col(c, 1) = rowVec(1, c)
    Next c
    TransposeRow = col
End Function

Public Function MeanVector(ByRef body As Variant) As Variant
    Dim nRows As Long
    Dim nCols As Long
    Dim r As Long
    Dim c As Long
    Dim tot As Double
    Dim out As Variant
    nRows = UBound(body, 1)
    nCols = UBound(body, 2)
    ReDim out(1 To 1, 1 To nCols)
    For c = 1 To nCols
        tot = 0
        For r = 1 To nRows
            tot = tot + CDbl(body(r, c))
        Next r
        out(1, c) = tot / nRows
    Next c
    MeanVector = out
End Function

' Sample = True uses n-1 (Excel STDEV.S). False uses n (STDEV.P).
Public Function StdDevVector(ByRef body As Variant, Optional ByVal Sample As Boolean = True) As Variant
    Dim means As Variant
    Dim nRows As Long
    Dim nCols As Long
    Dim r As Long
    Dim c As Long
    Dim tot As Double
    Dim denom As Double
    Dim out As Variant
    nRows = UBound(body, 1)
    nCols = UBound(body, 2)
    If Sample Then
        If nRows < 2 Then Err.Raise 5, "StdDevVector", "Sample standard deviation needs at least 2 rows."
        denom = nRows - 1
    Else
        denom = nRows
    End If
    means = MeanVector(body)
    ReDim out(1 To 1, 1 To nCols)
    For c = 1 To nCols
        tot = 0
        For r = 1 To nRows
            tot = tot + (CDbl(body(r, c)) - CDbl(means(1, c))) ^ 2
        Next r
        out(1, c) = Sqr(tot / denom)
    Next c
    StdDevVector = out
End Function

Public Function VarianceVector(ByRef body As Variant, Optional ByVal Sample As Boolean = False) As Variant
    Dim means As Variant
    Dim nRows As Long
    Dim nCols As Long
    Dim r As Long
    Dim c As Long
    Dim tot As Double
    Dim denom As Double
    Dim out As Variant
    nRows = UBound(body, 1)
    nCols = UBound(body, 2)
    If Sample Then
        If nRows < 2 Then Err.Raise 5, "VarianceVector", "Sample variance needs at least 2 rows."
        denom = nRows - 1
    Else
        denom = nRows
    End If
    means = MeanVector(body)
    ReDim out(1 To 1, 1 To nCols)
    For c = 1 To nCols
        tot = 0
        For r = 1 To nRows
            tot = tot + (CDbl(body(r, c)) - CDbl(means(1, c))) ^ 2
        Next r
        out(1, c) = tot / denom
    Next c
    VarianceVector = out
End Function

Public Function Centered(ByRef body As Variant) As Variant
    Dim means As Variant
    Dim nRows As Long
    Dim nCols As Long
    Dim r As Long
    Dim c As Long
    Dim out As Variant
    nRows = UBound(body, 1)
    nCols = UBound(body, 2)
    means = MeanVector(body)
    ReDim out(1 To nRows, 1 To nCols)
    For c = 1 To nCols
        For r = 1 To nRows
            out(r, c) = CDbl(body(r, c)) - CDbl(means(1, c))
        Next r
    Next c
    Centered = out
End Function

Public Function Standardised(ByRef body As Variant, Optional ByVal Sample As Boolean = True) As Variant
    Dim xc As Variant
    Dim sd As Variant
    Dim nRows As Long
    Dim nCols As Long
    Dim r As Long
    Dim c As Long
    Dim s As Double
    Dim out As Variant
    xc = Centered(body)
    sd = StdDevVector(body, Sample)
    nRows = UBound(xc, 1)
    nCols = UBound(xc, 2)
    ReDim out(1 To nRows, 1 To nCols)
    For c = 1 To nCols
        s = CDbl(sd(1, c))
        If s = 0 Then
            For r = 1 To nRows
                out(r, c) = 0
            Next r
        Else
            For r = 1 To nRows
                out(r, c) = CDbl(xc(r, c)) / s
            Next r
        End If
    Next c
    Standardised = out
End Function

Public Function Gramian(ByRef a As Variant) As Variant
    Gramian = Application.WorksheetFunction.MMult(Application.WorksheetFunction.Transpose(a), a)
End Function

Public Function ScaleMatrix(ByRef a As Variant, ByVal divisor As Double) As Variant
    Dim r As Long
    Dim c As Long
    Dim out As Variant
    If divisor = 0 Then Err.Raise 5, "ScaleMatrix", "Cannot divide by zero."
    ReDim out(1 To UBound(a, 1), 1 To UBound(a, 2))
    For r = 1 To UBound(a, 1)
        For c = 1 To UBound(a, 2)
            out(r, c) = CDbl(a(r, c)) / divisor
        Next c
    Next r
    ScaleMatrix = out
End Function

' Population covariance: (X-mean)'(X-mean) / n. Personal VarCovarMatrix.
Public Function CovariancePopulation(ByRef body As Variant) As Variant
    CovariancePopulation = ScaleMatrix(Gramian(Centered(body)), CDbl(UBound(body, 1)))
End Function

' Sample covariance: (X-mean)'(X-mean) / (n-1). Personal MatrixCovariance, without the square-range requirement.
Public Function CovarianceSample(ByRef body As Variant) As Variant
    If UBound(body, 1) < 2 Then Err.Raise 5, "CovarianceSample", "Sample covariance needs at least 2 rows."
    CovarianceSample = ScaleMatrix(Gramian(Centered(body)), CDbl(UBound(body, 1) - 1))
End Function

' Sample covariance of standardised columns = sample correlation.
Public Function CovarianceStandardised(ByRef body As Variant) As Variant
    If UBound(body, 1) < 2 Then Err.Raise 5, "CovarianceStandardised", "Needs at least 2 rows."
    CovarianceStandardised = ScaleMatrix(Gramian(Standardised(body, True)), CDbl(UBound(body, 1) - 1))
End Function

Public Function StdDevProduct(ByRef sdRow As Variant) As Variant
    StdDevProduct = Application.WorksheetFunction.MMult(Application.WorksheetFunction.Transpose(sdRow), sdRow)
End Function

Public Function CorrelMatrix(ByRef body As Variant, Optional ByVal ZeroIfUndefined As Boolean = False) As Variant
    Dim sdProd As Variant
    Dim cov As Variant
    Dim r As Long
    Dim c As Long
    Dim out As Variant
    sdProd = StdDevProduct(StdDevVector(body, False))
    cov = CovariancePopulation(body)
    ReDim out(1 To UBound(cov, 1), 1 To UBound(cov, 2))
    For r = 1 To UBound(cov, 1)
        For c = 1 To UBound(cov, 2)
            If CDbl(sdProd(r, c)) = 0 Then
                If ZeroIfUndefined Then out(r, c) = 0 Else out(r, c) = "n/a"
            Else
                out(r, c) = CDbl(cov(r, c)) / CDbl(sdProd(r, c))
            End If
        Next c
    Next r
    CorrelMatrix = out
End Function

Public Function DiagonalFromVector(ByRef vec As Variant) As Variant
    Dim n As Long
    Dim i As Long
    Dim j As Long
    Dim row As Variant
    Dim out As Variant
    row = AsRowVector(vec)
    n = UBound(row, 2)
    ReDim out(1 To n, 1 To n)
    For i = 1 To n
        For j = 1 To n
            If i = j Then out(i, j) = row(1, j) Else out(i, j) = 0
        Next j
    Next i
    DiagonalFromVector = out
End Function

Public Sub WriteTitle(ByVal dest As Range, ByVal Title As String)
    dest.Cells(1, 1).Value = Title
    dest.Cells(1, 1).Font.Bold = True
End Sub

Public Sub PutArray(ByVal dest As Range, ByRef data As Variant)
    Dim nRows As Long
    Dim nCols As Long
    If Not IsArray(data) Then
        dest.Value = data
        Exit Sub
    End If
    nRows = UBound(data, 1) - LBound(data, 1) + 1
    nCols = UBound(data, 2) - LBound(data, 2) + 1
    dest.Resize(nRows, nCols).Value = data
End Sub

Public Sub WriteLabeledMatrix(ByVal dest As Range, ByVal Title As String, ByVal header As Variant, ByVal mat As Variant)
    WriteTitle dest, Title
    Call PutArray(dest.Offset(1, 1), header)
    Call PutArray(dest.Offset(2, 0), TransposeRow(header))
    Call PutArray(dest.Offset(2, 1), mat)
End Sub

Public Sub WriteLabeledVector(ByVal dest As Range, ByVal Title As String, ByVal header As Variant, ByVal vec As Variant)
    WriteTitle dest, Title
    Call PutArray(dest.Offset(1, 0), header)
    Call PutArray(dest.Offset(2, 0), vec)
End Sub

Private Function ArrayRank(ByRef data As Variant) As Long
    Dim n As Long
    On Error GoTo Fail
    n = UBound(data, 2)
    ArrayRank = 2
    Exit Function
Fail:
    ArrayRank = 1
End Function
