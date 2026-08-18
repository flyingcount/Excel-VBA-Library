Attribute VB_Name = "modInternalData"
Option Explicit

' Internal: random / combinatorial data fills (Personal Custom_Menu6_*).
' Called from modApiData. Do not document these as the external API.

Public Function SelectionRange() As Range
    If TypeName(Selection) <> "Range" Then Exit Function
    Set SelectionRange = Selection
End Function

Public Function PromptRange(ByVal PromptText As String, Optional ByVal DefaultRng As Range) As Range
    Dim rng As Range
    If DefaultRng Is Nothing Then
        If TypeName(Selection) = "Range" Then Set DefaultRng = Selection
    End If
    On Error Resume Next
    If DefaultRng Is Nothing Then
        Set rng = Application.InputBox(Prompt:=PromptText, Title:="Data", Type:=8)
    Else
        Set rng = Application.InputBox(Prompt:=PromptText, Title:="Data", Default:=DefaultRng.Address, Type:=8)
    End If
    On Error GoTo 0
    Set PromptRange = rng
End Function

' Empty if the user cancelled.
Public Function PromptNumber(ByVal PromptText As String, ByVal DefaultValue As Variant) As Variant
    Dim resp As Variant
    resp = Application.InputBox(Prompt:=PromptText, Title:="Data", Default:=DefaultValue, Type:=1)
    If VarType(resp) = vbBoolean Then
        PromptNumber = Empty
    Else
        PromptNumber = resp
    End If
End Function

' Empty if the user cancelled. Default is shown as yyyy-mm-dd.
Public Function PromptDate(ByVal PromptText As String, ByVal DefaultValue As Date) As Variant
    Dim resp As Variant
    On Error GoTo Cancelled
    resp = Application.InputBox(Prompt:=PromptText, Title:="Data", Default:=Format$(DefaultValue, "yyyy-mm-dd"), Type:=2)
    If VarType(resp) = vbBoolean Then
        PromptDate = Empty
        Exit Function
    End If
    PromptDate = CDate(resp)
    Exit Function
Cancelled:
    PromptDate = Empty
End Function

' True = continue.
Public Function ConfirmCellCount(ByVal n As Long, Optional ByVal Limit As Long = 20000) As Boolean
    If n <= Limit Then
        ConfirmCellCount = True
        Exit Function
    End If
    ConfirmCellCount = (MsgBox("There are " & n & " cells to fill, which may take a while. Continue?", _
        vbYesNo + vbQuestion, "Data") = vbYes)
End Function

Public Function BlankValueArray(ByVal rng As Range) As Variant
    Dim a(1 To 1, 1 To 1) As Variant
    If rng.Cells.Count = 1 Then
        BlankValueArray = a
    Else
        BlankValueArray = rng.Value
    End If
End Function

Public Sub PutValueArray(ByVal rng As Range, ByRef data As Variant)
    If rng.Cells.Count = 1 Then
        rng.Value = data(1, 1)
    Else
        rng.Value = data
    End If
End Sub

Public Sub FillUniformIntegers(ByVal rng As Range, ByVal lo As Long, ByVal hi As Long)
    Dim data As Variant
    Dim r As Long
    Dim c As Long
    Dim span As Double
    data = BlankValueArray(rng)
    span = CDbl(hi) - CDbl(lo) + 1
    Randomize
    For r = LBound(data, 1) To UBound(data, 1)
        For c = LBound(data, 2) To UBound(data, 2)
            data(r, c) = lo + Int(Rnd() * span)
        Next c
    Next r
    Call PutValueArray(rng, data)
End Sub

Public Sub FillUniformDoubles(ByVal rng As Range, ByVal lo As Double, ByVal hi As Double)
    Dim data As Variant
    Dim r As Long
    Dim c As Long
    data = BlankValueArray(rng)
    Randomize
    For r = LBound(data, 1) To UBound(data, 1)
        For c = LBound(data, 2) To UBound(data, 2)
            data(r, c) = Rnd() * (hi - lo) + lo
        Next c
    Next r
    Call PutValueArray(rng, data)
End Sub

Public Sub FillRandomDates(ByVal rng As Range, ByVal lo As Date, ByVal hi As Date)
    Dim data As Variant
    Dim r As Long
    Dim c As Long
    Dim span As Double
    data = BlankValueArray(rng)
    span = CDbl(hi) - CDbl(lo)
    If span < 0 Then span = 0
    Randomize
    For r = LBound(data, 1) To UBound(data, 1)
        For c = LBound(data, 2) To UBound(data, 2)
            data(r, c) = CDate(Int(CDbl(lo) + Rnd() * (span + 1)))
        Next c
    Next r
    Call PutValueArray(rng, data)
    rng.NumberFormat = "yyyy-mm-dd"
End Sub

Public Sub FillRandomStrings(ByVal rng As Range, ByVal minLen As Long, ByVal maxLen As Long, ByVal kind As Long)
    Dim data As Variant
    Dim r As Long
    Dim c As Long
    data = BlankValueArray(rng)
    Randomize
    For r = LBound(data, 1) To UBound(data, 1)
        For c = LBound(data, 2) To UBound(data, 2)
            data(r, c) = RandomString(minLen, maxLen, kind)
        Next c
    Next r
    Call PutValueArray(rng, data)
End Sub

' kind: 1 upper, 2 lower, 3 initial capital, 4 digit text, 5 numeric value
Public Function RandomString(ByVal minLen As Long, ByVal maxLen As Long, ByVal kind As Long) As Variant
    Dim n As Long
    Dim i As Long
    Dim baseAscii As Long
    Dim span As Long
    Dim s As String

    If minLen < 1 Or maxLen < 1 Or minLen > maxLen Then
        RandomString = CVErr(xlErrValue)
        Exit Function
    End If
    If kind < 1 Or kind > 5 Then
        RandomString = CVErr(xlErrValue)
        Exit Function
    End If
    If kind = 3 And maxLen < 2 Then
        RandomString = CVErr(xlErrValue)
        Exit Function
    End If

    Select Case kind
        Case 1
            baseAscii = 65: span = 26
        Case 2, 3
            baseAscii = 97: span = 26
        Case 4, 5
            baseAscii = 48: span = 10
    End Select

    n = minLen + Int(Rnd() * (maxLen - minLen + 1))
    If kind = 3 Then
        s = Chr$(65 + Int(Rnd() * 26))
        For i = 2 To n
            s = s & Chr$(baseAscii + Int(Rnd() * span))
        Next i
    Else
        For i = 1 To n
            s = s & Chr$(baseAscii + Int(Rnd() * span))
        Next i
    End If

    If kind = 5 Then
        RandomString = CDbl(s)
    Else
        RandomString = s
    End If
End Function

Public Sub FillCoinFlip(ByVal rng As Range, ByVal trueValue As Variant, ByVal falseValue As Variant)
    Dim data As Variant
    Dim r As Long
    Dim c As Long
    data = BlankValueArray(rng)
    Randomize
    For r = LBound(data, 1) To UBound(data, 1)
        For c = LBound(data, 2) To UBound(data, 2)
            If Rnd() < 0.5 Then
                data(r, c) = trueValue
            Else
                data(r, c) = falseValue
            End If
        Next c
    Next r
    Call PutValueArray(rng, data)
End Sub

Public Sub FillFromList(ByVal dest As Range, ByVal src As Range)
    Dim list() As Variant
    Dim data As Variant
    Dim n As Long
    Dim i As Long
    Dim r As Long
    Dim c As Long
    Dim cell As Range

    n = src.Cells.Count
    ReDim list(1 To n)
    i = 0
    For Each cell In src.Cells
        i = i + 1
        list(i) = cell.Value
    Next cell

    data = BlankValueArray(dest)
    Randomize
    For r = LBound(data, 1) To UBound(data, 1)
        For c = LBound(data, 2) To UBound(data, 2)
            data(r, c) = list(1 + Int(Rnd() * n))
        Next c
    Next r
    Call PutValueArray(dest, data)
End Sub

Public Sub WriteCombinations(ByVal src As Range)
    Dim counts() As Long
    Dim values As Variant
    Dim nCol As Long
    Dim i As Long
    Dim total As Double
    Dim dest As Range
    Dim out() As Variant
    Dim col As Long
    Dim rowCount As Long
    Dim repeats As Long
    Dim cycles As Long
    Dim prior As Long
    Dim outRow As Long
    Dim cycle As Long
    Dim v As Long
    Dim rep As Long

    nCol = src.Columns.Count
    ReDim counts(1 To nCol)
    ReDim values(1 To nCol)
    total = 1
    For i = 1 To nCol
        counts(i) = CountUntilBlank(src.Columns(i))
        If counts(i) < 1 Then
            MsgBox "Each source column needs at least one value.", vbExclamation, "Data"
            Exit Sub
        End If
        total = total * counts(i)
        values(i) = ColumnValues(src.Columns(i), counts(i))
    Next i

    If total > 1000000 Then
        MsgBox "There are over 1,000,000 combinations, which will not fit on a worksheet.", vbExclamation, "Data"
        Exit Sub
    End If
    If Not ConfirmCellCount(CLng(total), 20000) Then Exit Sub

    Set dest = src.Cells(1, 1).Offset(0, nCol + 1)
    dest.Resize(1, 2).ClearContents
    dest.Value = "Number of combinations ="
    dest.Offset(0, 1).Value = total

    ReDim out(1 To CLng(total), 1 To nCol)
    prior = 1
    For col = 1 To nCol
        prior = prior * counts(col)
        repeats = CLng(total / prior)
        cycles = CLng((total / repeats) / counts(col))
        outRow = 1
        For cycle = 1 To cycles
            For v = 1 To counts(col)
                For rep = 1 To repeats
                    out(outRow, col) = values(col)(v)
                    outRow = outRow + 1
                Next rep
            Next v
        Next cycle
    Next col

    dest.Offset(1, 0).Resize(CLng(total), nCol).Value = out
End Sub

Public Sub WriteYesNoDataset(ByVal truePos As Long, ByVal trueNeg As Long, ByVal n As Long)
    Dim ws As Worksheet
    Dim arr() As Variant
    Dim i As Long

    If n < truePos + trueNeg Then
        MsgBox "Dataset size must be at least true positives + true negatives.", vbExclamation, "Data"
        Exit Sub
    End If

    Set ws = ActiveWorkbook.Worksheets("Yes No Dataset")
    ReDim arr(1 To n + 1, 1 To 2)
    arr(1, 1) = "Predicted"
    arr(1, 2) = "Actual"
    For i = 1 To truePos
        arr(i + 1, 1) = "Yes"
        arr(i + 1, 2) = "Yes"
    Next i
    For i = truePos + 1 To truePos + trueNeg
        arr(i + 1, 1) = "No"
        arr(i + 1, 2) = "No"
    Next i
    Randomize
    For i = truePos + trueNeg + 1 To n
        If Rnd() < 0.5 Then
            arr(i + 1, 1) = "Yes"
            arr(i + 1, 2) = "No"
        Else
            arr(i + 1, 1) = "No"
            arr(i + 1, 2) = "Yes"
        End If
    Next i
    Call ShuffleRows(arr, 2, n + 1)
    Call modInternalSheetIO.DumpArray("Yes No Dataset", "A1", arr)
    ws.Columns.AutoFit
End Sub

Public Sub WriteTestDataTypes(ByVal rng As Range)
    Dim dest As Range
    Dim arr() As Variant
    Dim n As Long
    Dim i As Long
    Dim kind As Long
    Dim loDate As Date
    Dim hiDate As Date

    If rng.Columns.Count < 2 Then
        MsgBox "Select at least two columns.", vbExclamation, "Data"
        Exit Sub
    End If

    Set dest = rng.Resize(rng.Rows.Count, 2)
    n = dest.Rows.Count
    If n < 2 Then
        MsgBox "Select at least two rows (header + data).", vbExclamation, "Data"
        Exit Sub
    End If

    ReDim arr(1 To n, 1 To 2)
    arr(1, 1) = "Data type"
    arr(1, 2) = "Data"
    loDate = DateSerial(1990, 1, 1)
    hiDate = Date
    Randomize
    For i = 2 To n
        kind = 1 + Int(Rnd() * 5)
        Select Case kind
            Case 1
                arr(i, 1) = "Number"
                arr(i, 2) = Rnd() * 200 - 100
            Case 2
                arr(i, 1) = "Integer"
                arr(i, 2) = -100 + Int(Rnd() * 201)
            Case 3
                arr(i, 1) = "Boolean"
                arr(i, 2) = (Rnd() < 0.5)
            Case 4
                arr(i, 1) = "Date"
                arr(i, 2) = CDate(Int(CDbl(loDate) + Rnd() * (CDbl(hiDate) - CDbl(loDate) + 1)))
            Case Else
                arr(i, 1) = "String"
                arr(i, 2) = RandomString(1, 10, 2)
        End Select
    Next i
    dest.Value = arr
End Sub

Public Sub FillBinomial(ByVal rng As Range, ByVal trials As Long, ByVal p As Double)
    Dim data As Variant
    Dim r As Long
    Dim c As Long
    If trials < 0 Or p < 0 Or p > 1 Then
        MsgBox "Trials must be >= 0 and probability must be between 0 and 1.", vbExclamation, "Data"
        Exit Sub
    End If
    data = BlankValueArray(rng)
    Randomize
    For r = LBound(data, 1) To UBound(data, 1)
        For c = LBound(data, 2) To UBound(data, 2)
            data(r, c) = Application.WorksheetFunction.Binom_Inv(trials, p, UnitRnd())
        Next c
    Next r
    Call PutValueArray(rng, data)
End Sub

Public Sub FillBernoulli(ByVal rng As Range, ByVal p As Double)
    Dim data As Variant
    Dim r As Long
    Dim c As Long
    data = BlankValueArray(rng)
    Randomize
    For r = LBound(data, 1) To UBound(data, 1)
        For c = LBound(data, 2) To UBound(data, 2)
            If Rnd() < p Then
                data(r, c) = 1
            Else
                data(r, c) = 0
            End If
        Next c
    Next r
    Call PutValueArray(rng, data)
End Sub

Public Sub FillNormal(ByVal rng As Range, ByVal meanVal As Double, ByVal stdev As Double)
    Dim data As Variant
    Dim r As Long
    Dim c As Long
    If stdev <= 0 Then
        MsgBox "Standard deviation must be greater than 0.", vbExclamation, "Data"
        Exit Sub
    End If
    data = BlankValueArray(rng)
    Randomize
    For r = LBound(data, 1) To UBound(data, 1)
        For c = LBound(data, 2) To UBound(data, 2)
            data(r, c) = Application.WorksheetFunction.NormInv(UnitRnd(), meanVal, stdev)
        Next c
    Next r
    Call PutValueArray(rng, data)
End Sub

Public Sub FillPoisson(ByVal rng As Range, ByVal lambda As Double)
    Dim data As Variant
    Dim r As Long
    Dim c As Long
    data = BlankValueArray(rng)
    Randomize
    For r = LBound(data, 1) To UBound(data, 1)
        For c = LBound(data, 2) To UBound(data, 2)
            data(r, c) = RandomPoisson(lambda)
        Next c
    Next r
    Call PutValueArray(rng, data)
End Sub

Public Sub FillExponential(ByVal rng As Range, ByVal lambda As Double)
    Dim data As Variant
    Dim r As Long
    Dim c As Long
    data = BlankValueArray(rng)
    Randomize
    For r = LBound(data, 1) To UBound(data, 1)
        For c = LBound(data, 2) To UBound(data, 2)
            data(r, c) = -Log(1 - UnitRnd()) / lambda
        Next c
    Next r
    Call PutValueArray(rng, data)
End Sub

Public Sub FillGamma(ByVal rng As Range, ByVal alpha As Double, ByVal beta As Double)
    Dim data As Variant
    Dim r As Long
    Dim c As Long
    data = BlankValueArray(rng)
    Randomize
    For r = LBound(data, 1) To UBound(data, 1)
        For c = LBound(data, 2) To UBound(data, 2)
            data(r, c) = Application.WorksheetFunction.GammaInv(UnitRnd(), alpha, beta)
        Next c
    Next r
    Call PutValueArray(rng, data)
End Sub

' n draws without replacement from a population of popSize with kSuccess successes.
Public Sub FillHypergeometric(ByVal rng As Range, ByVal nDraw As Long, ByVal kSuccess As Long, ByVal popSize As Long)
    Dim data As Variant
    Dim r As Long
    Dim c As Long
    data = BlankValueArray(rng)
    Randomize
    For r = LBound(data, 1) To UBound(data, 1)
        For c = LBound(data, 2) To UBound(data, 2)
            data(r, c) = RandomHypergeometric(nDraw, kSuccess, popSize)
        Next c
    Next r
    Call PutValueArray(rng, data)
End Sub

Private Function CountUntilBlank(ByVal col As Range) As Long
    Dim area As Range
    Dim cell As Range
    Dim n As Long
    Set area = Application.Intersect(col, col.Worksheet.UsedRange)
    If area Is Nothing Then Exit Function
    n = 0
    For Each cell In area.Cells
        If Len(CStr(cell.Value)) = 0 Then Exit For
        n = n + 1
    Next cell
    CountUntilBlank = n
End Function

Private Function ColumnValues(ByVal col As Range, ByVal n As Long) As Variant
    Dim area As Range
    Dim a() As Variant
    Dim i As Long
    Set area = Application.Intersect(col, col.Worksheet.UsedRange)
    ReDim a(1 To n)
    For i = 1 To n
        a(i) = area.Cells(i, 1).Value
    Next i
    ColumnValues = a
End Function

Private Function UnitRnd() As Double
    Dim u As Double
    u = Rnd()
    If u <= 0 Then u = 0.000000001
    If u >= 1 Then u = 0.999999999
    UnitRnd = u
End Function

Private Sub ShuffleRows(ByRef arr As Variant, ByVal firstRow As Long, ByVal lastRow As Long)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim tmp As Variant
    If lastRow <= firstRow Then Exit Sub
    For i = lastRow To firstRow + 1 Step -1
        j = firstRow + Int(Rnd() * (i - firstRow + 1))
        For k = LBound(arr, 2) To UBound(arr, 2)
            tmp = arr(i, k)
            arr(i, k) = arr(j, k)
            arr(j, k) = tmp
        Next k
    Next i
End Sub

' Excel has POISSON / POISSON.DIST but no POISSON.INV, so WorksheetFunction.Poisson_Inv
' raises 438. Knuth for ordinary lambda; rounded Normal(mean, Sqrt(mean)) when
' Exp(-lambda) underflows or the Knuth loop would be too long.
Private Function RandomPoisson(ByVal lambda As Double) As Long
    Dim L As Double
    Dim p As Double
    Dim k As Long
    Dim u As Double
    Dim x As Double

    If lambda <= 0# Then
        RandomPoisson = 0
        Exit Function
    End If

    L = Exp(-lambda)
    If L > 0# And lambda < 100# Then
        k = 0
        p = 1#
        Do
            k = k + 1
            p = p * Rnd()
        Loop While p > L
        RandomPoisson = k - 1
        Exit Function
    End If

    u = UnitRnd()
    x = Application.WorksheetFunction.NormInv(u, lambda, Sqr(lambda))
    If x <= 0# Then
        RandomPoisson = 0
    ElseIf x >= 2147483647# Then
        RandomPoisson = 2147483647
    Else
        RandomPoisson = CLng(Int(x + 0.5))
    End If
End Function

Private Function RandomHypergeometric(ByVal nDraw As Long, ByVal kSuccess As Long, ByVal popSize As Long) As Long
    Dim i As Long
    Dim hits As Long
    Dim remainK As Long
    Dim remainN As Long
    remainK = kSuccess
    remainN = popSize
    hits = 0
    For i = 1 To nDraw
        If remainN <= 0 Then Exit For
        If Rnd() < remainK / remainN Then
            hits = hits + 1
            remainK = remainK - 1
        End If
        remainN = remainN - 1
    Next i
    RandomHypergeometric = hits
End Function
