Attribute VB_Name = "modApiBenford"
Option Explicit

' Public API: Benford (and last-two-digit) analyses.
' Other workbooks / Personal shims should call these names only.

''' @Description: First-digit Benford analysis. Prompts if InputRange is omitted.
''' @Example: BenfordAnalysisFirstDigit Range("A2:A100")
Public Sub BenfordAnalysisFirstDigit(Optional ByVal InputRange As Range)
    Call RunBenford(bkFirst, InputRange)
End Sub

''' @Description: Second-digit Benford analysis. Prompts if InputRange is omitted.
''' @Example: BenfordAnalysisSecondDigit Range("A2:A100")
Public Sub BenfordAnalysisSecondDigit(Optional ByVal InputRange As Range)
    Call RunBenford(bkSecond, InputRange)
End Sub

''' @Description: Third-digit Benford analysis. Prompts if InputRange is omitted.
''' @Example: BenfordAnalysisThirdDigit Range("A2:A100")
Public Sub BenfordAnalysisThirdDigit(Optional ByVal InputRange As Range)
    Call RunBenford(bkThird, InputRange)
End Sub

''' @Description: First-two-digits Benford analysis. Prompts if InputRange is omitted.
''' @Example: BenfordAnalysisTwoDigit Range("A2:A100")
Public Sub BenfordAnalysisTwoDigit(Optional ByVal InputRange As Range)
    Call RunBenford(bkTwo, InputRange)
End Sub

''' @Description: First-three-digits Benford analysis. Prompts if InputRange is omitted.
''' @Example: BenfordAnalysisThreeDigit Range("A2:A100")
Public Sub BenfordAnalysisThreeDigit(Optional ByVal InputRange As Range)
    Call RunBenford(bkThree, InputRange)
End Sub

''' @Description: Last-two-digits analysis (uniform 1% expected). Prompts if InputRange is omitted.
''' @Example: BenfordAnalysisLastTwoDigit Range("A2:A100")
Public Sub BenfordAnalysisLastTwoDigit(Optional ByVal InputRange As Range)
    Call RunBenford(bkLastTwo, InputRange)
End Sub

''' @Description: Theoretical Benford probability for a second digit 0–9.
Public Function Benford2ndDigitProbability(ByVal lng_SecondDigit As Long) As Double
    Application.Volatile
    Benford2ndDigitProbability = modInternalBenford.SecondDigitProbability(lng_SecondDigit)
End Function

''' @Description: Theoretical Benford probability for a third digit 0–9.
Public Function Benford3rdDigitProbability(ByVal lng_ThirdDigit As Long) As Double
    Application.Volatile
    Benford3rdDigitProbability = modInternalBenford.ThirdDigitProbability(lng_ThirdDigit)
End Function

Private Sub RunBenford(ByVal kind As BenfordKind, ByVal InputRange As Range)
    Dim rng As Range

    On Error GoTo EH
    Call modInternalExcelApp.PushAppState

    Set rng = ResolveInputRange(InputRange)
    If rng Is Nothing Then GoTo CleanUp

    Call modInternalBenford.RunAnalysis(kind, rng)

CleanUp:
    Application.StatusBar = False
    Call modInternalExcelApp.PopAppState
    Exit Sub

EH:
    Application.StatusBar = False
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("BenfordAnalysis")
End Sub

Private Function ResolveInputRange(ByVal InputRange As Range) As Range
    Dim rng_DefaultRange As Range
    Dim rng As Range

    If Not InputRange Is Nothing Then
        Set ResolveInputRange = InputRange
        Exit Function
    End If

    Set rng_DefaultRange = Selection
    On Error Resume Next
    Set rng = Application.InputBox( _
            Prompt:="Please Select Range", _
            Title:="Range Select", _
            Default:=rng_DefaultRange.Address, _
            Type:=8)
    On Error GoTo 0
    Set ResolveInputRange = rng
End Function
