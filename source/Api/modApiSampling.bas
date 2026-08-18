Attribute VB_Name = "modApiSampling"
Option Explicit

' Public API: random row sample from a range (Personal Custom_Menu4_Sample).
' Other workbooks / the add-in menu should call ExtractSample only.

''' @Description: Writes a random sample of rows to a sheet named Sample.
''' @Param InputRange: Rows to sample. If omitted, an InputBox prompts.
''' @Param SamplePercent: 0–100. If 0 or omitted, an InputBox prompts (default 70).
''' @Param WithReplacement: True/False. If omitted, a Yes/No/Cancel prompt is shown.
'''
''' @Example:
'''   ExtractSample
'''   ExtractSample Range("A2:D100"), 20, False
Public Sub ExtractSample( _
    Optional ByVal InputRange As Range, _
    Optional ByVal SamplePercent As Double = 0, _
    Optional ByVal WithReplacement As Variant _
)
    Dim rng As Range
    Dim data As Variant
    Dim n As Long
    Dim sampleSize As Long
    Dim replaceFlag As Boolean
    Dim pct As Double

    On Error GoTo EH

    Set rng = ResolveInputRange(InputRange)
    If rng Is Nothing Then Exit Sub

    If IsMissing(WithReplacement) Then
        Select Case MsgBox("Sample with replacement?", vbYesNoCancel + vbDefaultButton2 + vbQuestion, "Extract sample")
            Case vbYes
                replaceFlag = True
            Case vbNo
                replaceFlag = False
            Case Else
                Exit Sub
        End Select
    Else
        replaceFlag = CBool(WithReplacement)
    End If

    pct = SamplePercent
    If pct <= 0 Then
        pct = PromptSamplePercent()
        If pct <= 0 Then Exit Sub
    End If
    If pct > 100 Then
        MsgBox "Sample percent must be between 0 and 100.", vbExclamation, "Extract sample"
        Exit Sub
    End If

    data = modInternalSampling.RangeTo2D(rng)
    n = UBound(data, 1) - LBound(data, 1) + 1
    sampleSize = CLng(Fix(pct / 100# * n))
    If sampleSize < 1 Then sampleSize = 1

    Call modInternalExcelApp.PushAppState
    Call modApiSheets.EnsureSheet("Sample").Cells.Clear
    Call modInternalSampling.RunExtractSample(data, sampleSize, replaceFlag)
    Call modInternalExcelApp.PopAppState
    Exit Sub

EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("ExtractSample")
End Sub

Private Function ResolveInputRange(ByVal InputRange As Range) As Range
    Dim rng As Range

    If Not InputRange Is Nothing Then
        Set ResolveInputRange = InputRange
        Exit Function
    End If

    On Error Resume Next
    Set rng = Application.InputBox( _
        Prompt:="Select the range to sample (each row is one item).", _
        Title:="Extract sample", _
        Default:=Selection.Address, _
        Type:=8)
    On Error GoTo 0
    Set ResolveInputRange = rng
End Function

' Returns 0 if the user cancels.
Private Function PromptSamplePercent() As Double
    Dim resp As Variant
    resp = Application.InputBox( _
        Prompt:="Sample size as a percent of the input rows (1–100).", _
        Title:="Extract sample", _
        Default:=70, _
        Type:=1)
    If VarType(resp) = vbBoolean Then
        PromptSamplePercent = 0
    Else
        PromptSamplePercent = CDbl(resp)
    End If
End Function
