Attribute VB_Name = "modApiData"
Option Explicit

' Public API: random and combinatorial data fills (Personal Custom_Menu6_Data / RndFrmRng / RndProbDist).
' Fill the current selection, or prompt for a range where Personal did.

''' @Description: Cartesian product of values in the selected columns (blank cell ends a column). Writes to the right of the selection.
Public Sub DataCombinations()
    Dim rng As Range
    On Error GoTo EH
    Set rng = modInternalData.SelectionRange
    If rng Is Nothing Then Exit Sub
    Call modInternalExcelApp.PushAppState
    Call modInternalData.WriteCombinations(rng)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("DataCombinations")
End Sub

''' @Description: Uniform random integers in the selected range (inclusive min/max).
Public Sub RandomIntegers()
    Dim rng As Range
    Dim lo As Variant
    Dim hi As Variant
    On Error GoTo EH
    Set rng = RequireSelection()
    If rng Is Nothing Then Exit Sub
    lo = modInternalData.PromptNumber("Minimum integer", 0)
    If IsEmpty(lo) Then Exit Sub
    hi = modInternalData.PromptNumber("Maximum integer", 100)
    If IsEmpty(hi) Then Exit Sub
    If CDbl(lo) > CDbl(hi) Then
        MsgBox "Minimum cannot be greater than maximum.", vbExclamation, "Data"
        Exit Sub
    End If
    Call FillWithPush(rng, "int", CLng(lo), CLng(hi), 0)
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("RandomIntegers")
End Sub

''' @Description: Uniform random real numbers in the selected range.
Public Sub RandomNumbers()
    Dim rng As Range
    Dim lo As Variant
    Dim hi As Variant
    On Error GoTo EH
    Set rng = RequireSelection()
    If rng Is Nothing Then Exit Sub
    lo = modInternalData.PromptNumber("Minimum", -10)
    If IsEmpty(lo) Then Exit Sub
    hi = modInternalData.PromptNumber("Maximum", 10)
    If IsEmpty(hi) Then Exit Sub
    If CDbl(lo) > CDbl(hi) Then
        MsgBox "Minimum cannot be greater than maximum.", vbExclamation, "Data"
        Exit Sub
    End If
    Call FillWithPush(rng, "dbl", CDbl(lo), CDbl(hi), 0)
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("RandomNumbers")
End Sub

''' @Description: Random dates in the selected range. Prompts for earliest and latest date.
Public Sub RandomDates()
    Dim rng As Range
    Dim lo As Variant
    Dim hi As Variant
    On Error GoTo EH
    Set rng = RequireSelection()
    If rng Is Nothing Then Exit Sub
    lo = modInternalData.PromptDate("Earliest date", DateSerial(1990, 1, 1))
    If IsEmpty(lo) Then Exit Sub
    hi = modInternalData.PromptDate("Latest date", Date)
    If IsEmpty(hi) Then Exit Sub
    If CDate(lo) > CDate(hi) Then
        MsgBox "Earliest date cannot be after latest date.", vbExclamation, "Data"
        Exit Sub
    End If
    Call FillWithPush(rng, "date", CDate(lo), CDate(hi), 0)
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("RandomDates")
End Sub

''' @Description: Random strings (or digit text) in the selected range.
''' Type 1=upper, 2=lower, 3=initial capital, 4=digit text, 5=numeric digits.
Public Sub RandomStrings()
    Dim rng As Range
    Dim mn As Variant
    Dim mx As Variant
    Dim kind As Variant
    On Error GoTo EH
    Set rng = RequireSelection()
    If rng Is Nothing Then Exit Sub
    mn = modInternalData.PromptNumber("Minimum length (at least 1)", 8)
    If IsEmpty(mn) Then Exit Sub
    mx = modInternalData.PromptNumber("Maximum length (at least 1)", 8)
    If IsEmpty(mx) Then Exit Sub
    kind = modInternalData.PromptNumber("Type: 1=UPPER 2=lower 3=Initial capital 4=digit text 5=numbers", 3)
    If IsEmpty(kind) Then Exit Sub
    If CLng(mn) < 1 Or CLng(mx) < 1 Or CLng(mn) > CLng(mx) Then
        MsgBox "Lengths must be at least 1, and minimum cannot exceed maximum.", vbExclamation, "Data"
        Exit Sub
    End If
    If CLng(kind) < 1 Or CLng(kind) > 5 Then
        MsgBox "Type must be 1, 2, 3, 4 or 5.", vbExclamation, "Data"
        Exit Sub
    End If
    If CLng(kind) = 3 And CLng(mx) < 2 Then
        MsgBox "Initial-capital strings need a maximum length of at least 2.", vbExclamation, "Data"
        Exit Sub
    End If
    Call FillWithPush(rng, "str", CLng(mn), CLng(mx), CLng(kind))
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("RandomStrings")
End Sub

''' @Description: Fill a destination range with random picks from a source list range.
Public Sub RandomFromList()
    Dim dest As Range
    Dim src As Range
    On Error GoTo EH
    Set dest = modInternalData.PromptRange("Select the range to fill")
    If dest Is Nothing Then Exit Sub
    Set src = modInternalData.PromptRange("Select the list of values to draw from")
    If src Is Nothing Then Exit Sub
    If Not modInternalData.ConfirmCellCount(dest.Cells.Count) Then Exit Sub
    Call modInternalExcelApp.PushAppState
    Call modInternalData.FillFromList(dest, src)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("RandomFromList")
End Sub

''' @Description: Random True/False in the selected range.
Public Sub RandomTrueFalse()
    Call FillCoin("bool")
End Sub

''' @Description: Random Yes/No in the selected range.
Public Sub RandomYesNo()
    Call FillCoin("yesno")
End Sub

''' @Description: Random 1 or 0 in the selected range.
Public Sub Random1or0()
    Call FillCoin("10")
End Sub

''' @Description: Build a Yes/No predicted-vs-actual sheet with specified true positives/negatives.
Public Sub CreateYesNoDataset()
    Dim tp As Variant
    Dim tn As Variant
    Dim n As Variant
    Dim ws As Worksheet
    On Error GoTo EH
    tp = modInternalData.PromptNumber("How many true positives?", 1000)
    If IsEmpty(tp) Then Exit Sub
    tn = modInternalData.PromptNumber("How many true negatives?", 990)
    If IsEmpty(tn) Then Exit Sub
    n = modInternalData.PromptNumber("How many rows? Must be at least true positives + true negatives.", CLng(tp) + CLng(tn))
    If IsEmpty(n) Then Exit Sub
    Call modInternalExcelApp.PushAppState
    Set ws = modApiSheets.EnsureSheet("Yes No Dataset")
    ws.Cells.Clear
    Call modInternalData.WriteYesNoDataset(CLng(tp), CLng(tn), CLng(n))
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("CreateYesNoDataset")
End Sub

''' @Description: Two-column test data (type name + value) starting at a prompted two-column range.
Public Sub RandomTestDataTypes()
    Dim rng As Range
    On Error GoTo EH
    Set rng = modInternalData.PromptRange("Select range for the test data (at least two columns).")
    If rng Is Nothing Then Exit Sub
    Call modInternalExcelApp.PushAppState
    Call modInternalData.WriteTestDataTypes(rng)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("RandomTestDataTypes")
End Sub

''' @Description: Binomial draws in a prompted range (trials, success probability).
Public Sub RandomBinomialNumbers()
    Dim rng As Range
    Dim trials As Variant
    Dim p As Variant
    On Error GoTo EH
    Set rng = PromptFillRange()
    If rng Is Nothing Then Exit Sub
    trials = modInternalData.PromptNumber("Number of trials", 2)
    If IsEmpty(trials) Then Exit Sub
    p = modInternalData.PromptNumber("Probability of success (0 to 1)", 0.5)
    If IsEmpty(p) Then Exit Sub
    If CDbl(p) < 0 Or CDbl(p) > 1 Then
        MsgBox "Probability must be between 0 and 1.", vbExclamation, "Data"
        Exit Sub
    End If
    If CLng(trials) < 0 Then
        MsgBox "Number of trials cannot be negative.", vbExclamation, "Data"
        Exit Sub
    End If
    Call DistFill(rng, "binom", CDbl(trials), CDbl(p), 0)
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("RandomBinomialNumbers")
End Sub

''' @Description: Bernoulli 0/1 in a prompted range. Probability is P(1).
Public Sub RandomBernoulliNumbers()
    Dim rng As Range
    Dim p As Variant
    On Error GoTo EH
    Set rng = PromptFillRange()
    If rng Is Nothing Then Exit Sub
    p = modInternalData.PromptNumber("Event probability P(1), between 0 and 1", 0.5)
    If IsEmpty(p) Then Exit Sub
    If CDbl(p) < 0 Or CDbl(p) > 1 Then
        MsgBox "Probability must be between 0 and 1.", vbExclamation, "Data"
        Exit Sub
    End If
    Call DistFill(rng, "bern", CDbl(p), 0, 0)
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("RandomBernoulliNumbers")
End Sub

''' @Description: Normal draws in a prompted range.
Public Sub RandomNormalNumbers()
    Dim rng As Range
    Dim meanVal As Variant
    Dim sd As Variant
    On Error GoTo EH
    Set rng = PromptFillRange()
    If rng Is Nothing Then Exit Sub
    meanVal = modInternalData.PromptNumber("Mean", 0)
    If IsEmpty(meanVal) Then Exit Sub
    sd = modInternalData.PromptNumber("Standard deviation", 0.5)
    If IsEmpty(sd) Then Exit Sub
    If CDbl(sd) <= 0 Then
        MsgBox "Standard deviation must be greater than 0.", vbExclamation, "Data"
        Exit Sub
    End If
    Call DistFill(rng, "norm", CDbl(meanVal), CDbl(sd), 0)
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("RandomNormalNumbers")
End Sub

''' @Description: Poisson draws in a prompted range (inverse CDF).
Public Sub RandomPoissonNumbers()
    Dim rng As Range
    Dim lam As Variant
    On Error GoTo EH
    Set rng = PromptFillRange()
    If rng Is Nothing Then Exit Sub
    lam = modInternalData.PromptNumber("Lambda (mean)", 1)
    If IsEmpty(lam) Then Exit Sub
    Call DistFill(rng, "pois", CDbl(lam), 0, 0)
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("RandomPoissonNumbers")
End Sub

''' @Description: Exponential draws in a prompted range (rate lambda).
Public Sub RandomExponentialNumbers()
    Dim rng As Range
    Dim lam As Variant
    On Error GoTo EH
    Set rng = PromptFillRange()
    If rng Is Nothing Then Exit Sub
    lam = modInternalData.PromptNumber("Lambda (rate, must be > 0)", 1)
    If IsEmpty(lam) Then Exit Sub
    If CDbl(lam) <= 0 Then
        MsgBox "Lambda must be greater than 0.", vbExclamation, "Data"
        Exit Sub
    End If
    Call DistFill(rng, "exp", CDbl(lam), 0, 0)
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("RandomExponentialNumbers")
End Sub

''' @Description: Gamma draws. Alpha=1 is exponential; beta=2 relates to chi-square.
Public Sub RandomGammaNumbers()
    Dim rng As Range
    Dim a As Variant
    Dim b As Variant
    On Error GoTo EH
    Set rng = PromptFillRange()
    If rng Is Nothing Then Exit Sub
    a = modInternalData.PromptNumber("Alpha (positive). Alpha=1 is Exponential.", 1)
    If IsEmpty(a) Then Exit Sub
    b = modInternalData.PromptNumber("Beta (positive). Beta=2 relates to chi-square.", 1)
    If IsEmpty(b) Then Exit Sub
    If CDbl(a) <= 0 Or CDbl(b) <= 0 Then
        MsgBox "Alpha and beta must be positive.", vbExclamation, "Data"
        Exit Sub
    End If
    Call DistFill(rng, "gamma", CDbl(a), CDbl(b), 0)
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("RandomGammaNumbers")
End Sub

''' @Description: Hypergeometric draws: n draws from a population of N with K successes.
Public Sub RandomHypergeometricNumbers()
    Dim rng As Range
    Dim nDraw As Variant
    Dim kSuc As Variant
    Dim pop As Variant
    On Error GoTo EH
    Set rng = PromptFillRange()
    If rng Is Nothing Then Exit Sub
    nDraw = modInternalData.PromptNumber("Sample size n (draws)", 100)
    If IsEmpty(nDraw) Then Exit Sub
    kSuc = modInternalData.PromptNumber("Successes in the population, K", 50)
    If IsEmpty(kSuc) Then Exit Sub
    pop = modInternalData.PromptNumber("Population size N (must be >= n and >= K)", 500)
    If IsEmpty(pop) Then Exit Sub
    If CLng(nDraw) < 0 Or CLng(kSuc) < 0 Or CLng(pop) < 0 Then
        MsgBox "n, K and N cannot be negative.", vbExclamation, "Data"
        Exit Sub
    End If
    If CLng(pop) < CLng(nDraw) Or CLng(pop) < CLng(kSuc) Then
        MsgBox "Population N must be at least n and at least K.", vbExclamation, "Data"
        Exit Sub
    End If
    Call DistFill(rng, "hyper", CDbl(nDraw), CDbl(kSuc), CDbl(pop))
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("RandomHypergeometricNumbers")
End Sub

Private Function RequireSelection() As Range
    Set RequireSelection = modInternalData.SelectionRange
    If RequireSelection Is Nothing Then
        MsgBox "Select a range first.", vbExclamation, "Data"
        Exit Function
    End If
    If Not modInternalData.ConfirmCellCount(RequireSelection.Cells.Count) Then
        Set RequireSelection = Nothing
    End If
End Function

Private Function PromptFillRange() As Range
    Set PromptFillRange = modInternalData.PromptRange("Select the range to fill")
    If PromptFillRange Is Nothing Then Exit Function
    If Not modInternalData.ConfirmCellCount(PromptFillRange.Cells.Count, 100000) Then
        Set PromptFillRange = Nothing
    End If
End Function

Private Sub FillCoin(ByVal kind As String)
    Dim rng As Range
    On Error GoTo EH
    Set rng = RequireSelection()
    If rng Is Nothing Then Exit Sub
    Call modInternalExcelApp.PushAppState
    Select Case kind
        Case "bool"
            Call modInternalData.FillCoinFlip(rng, True, False)
        Case "yesno"
            Call modInternalData.FillCoinFlip(rng, "Yes", "No")
        Case Else
            Call modInternalData.FillCoinFlip(rng, 1, 0)
    End Select
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("RandomFill")
End Sub

Private Sub FillWithPush(ByVal rng As Range, ByVal kind As String, ByVal a As Variant, ByVal b As Variant, ByVal c As Variant)
    Call modInternalExcelApp.PushAppState
    Select Case kind
        Case "int"
            Call modInternalData.FillUniformIntegers(rng, CLng(a), CLng(b))
        Case "dbl"
            Call modInternalData.FillUniformDoubles(rng, CDbl(a), CDbl(b))
        Case "date"
            Call modInternalData.FillRandomDates(rng, CDate(a), CDate(b))
        Case "str"
            Call modInternalData.FillRandomStrings(rng, CLng(a), CLng(b), CLng(c))
    End Select
    Call modInternalExcelApp.PopAppState
End Sub

Private Sub DistFill(ByVal rng As Range, ByVal kind As String, ByVal a As Double, ByVal b As Double, ByVal c As Double)
    Call modInternalExcelApp.PushAppState
    Select Case kind
        Case "binom"
            Call modInternalData.FillBinomial(rng, CLng(a), b)
        Case "bern"
            Call modInternalData.FillBernoulli(rng, a)
        Case "norm"
            Call modInternalData.FillNormal(rng, a, b)
        Case "pois"
            Call modInternalData.FillPoisson(rng, a)
        Case "exp"
            Call modInternalData.FillExponential(rng, a)
        Case "gamma"
            Call modInternalData.FillGamma(rng, a, b)
        Case "hyper"
            Call modInternalData.FillHypergeometric(rng, CLng(a), CLng(b), CLng(c))
    End Select
    Call modInternalExcelApp.PopAppState
End Sub
