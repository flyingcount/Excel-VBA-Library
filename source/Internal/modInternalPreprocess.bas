Attribute VB_Name = "modInternalPreprocess"
Option Explicit

' Internal: column scaling and dummy (one-hot) encoding (Personal Custom_Menu5_Scaling / Custom_menu5_ML).
' Called from modApiPreprocess. Do not document as the external API.
' Body arrays are 1-based 2D (rows x columns).

Public Function PromptRange(ByVal PromptText As String, Optional ByVal Title As String = "Data Preprocessing") As Range
    Dim rng As Range
    Dim def As Range
    If TypeName(Selection) = "Range" Then Set def = Selection
    On Error Resume Next
    If def Is Nothing Then
        Set rng = Application.InputBox(Prompt:=PromptText, Title:=Title, Type:=8)
    Else
        Set rng = Application.InputBox(Prompt:=PromptText, Title:=Title, Default:=def.Address, Type:=8)
    End If
    On Error GoTo 0
    Set PromptRange = rng
End Function

Public Function PromptNumericColumns(ByVal PromptText As String, Optional ByVal MinRows As Long = 2) As Range
    Dim rng As Range
    Set rng = PromptRange(PromptText)
    If rng Is Nothing Then Exit Function
    If rng.Areas.Count > 1 Then
        MsgBox "Select a single contiguous range.", vbExclamation, "Data Preprocessing"
        Exit Function
    End If
    If rng.Rows.Count < MinRows Then
        MsgBox "Need at least " & MinRows & " rows.", vbExclamation, "Data Preprocessing"
        Exit Function
    End If
    If Not modInternalAnalysis.RangeIsAllNumeric(rng) Then
        MsgBox "Every cell must be numeric (blanks and text are not allowed).", vbExclamation, "Data Preprocessing"
        Exit Function
    End If
    Set PromptNumericColumns = rng
End Function

' (x - mean) / population SD per column. Zero SD → 0. Personal ScalingStandardisationFn / STDEV.P.
Public Function ScaleStandardise(ByRef body As Variant) As Variant
    ScaleStandardise = modInternalAnalysis.Standardised(body, False)
End Function

' (x - min) / (max - min) per column. Zero range → 0. Personal ScalingNormalisationFn.
Public Function ScaleNormalise(ByRef body As Variant) As Variant
    Dim nRows As Long
    Dim nCols As Long
    Dim r As Long
    Dim c As Long
    Dim lo As Double
    Dim hi As Double
    Dim span As Double
    Dim out As Variant
    nRows = UBound(body, 1)
    nCols = UBound(body, 2)
    ReDim out(1 To nRows, 1 To nCols)
    For c = 1 To nCols
        lo = CDbl(body(1, c))
        hi = lo
        For r = 2 To nRows
            If CDbl(body(r, c)) < lo Then lo = CDbl(body(r, c))
            If CDbl(body(r, c)) > hi Then hi = CDbl(body(r, c))
        Next r
        span = hi - lo
        If span = 0 Then
            For r = 1 To nRows
                out(r, c) = 0
            Next r
        Else
            For r = 1 To nRows
                out(r, c) = (CDbl(body(r, c)) - lo) / span
            Next r
        End If
    Next c
    ScaleNormalise = out
End Function

' (x - median) / IQR per column, Excel QUARTILE.INC. Zero IQR → 0. Personal ScalingRobustFn.
Public Function ScaleRobust(ByRef body As Variant) As Variant
    Dim nRows As Long
    Dim nCols As Long
    Dim r As Long
    Dim c As Long
    Dim med As Double
    Dim iqr As Double
    Dim col() As Double
    Dim out As Variant
    nRows = UBound(body, 1)
    nCols = UBound(body, 2)
    ReDim col(1 To nRows)
    ReDim out(1 To nRows, 1 To nCols)
    For c = 1 To nCols
        For r = 1 To nRows
            col(r) = CDbl(body(r, c))
        Next r
        med = Application.WorksheetFunction.Median(col)
        iqr = Application.WorksheetFunction.Quartile_Inc(col, 3) - Application.WorksheetFunction.Quartile_Inc(col, 1)
        If iqr = 0 Then
            For r = 1 To nRows
                out(r, c) = 0
            Next r
        Else
            For r = 1 To nRows
                out(r, c) = (CDbl(body(r, c)) - med) / iqr
            Next r
        End If
    Next c
    ScaleRobust = out
End Function

' Distinct values in first-seen order; blanks skipped. 1-based column vector, or Empty if none.
Public Function UniqueColumnValues(ByVal rng As Range) As Variant
    Dim dict As Object
    Dim cell As Range
    Dim v As Variant
    Dim out As Variant
    Dim i As Long
    Set dict = CreateObject("Scripting.Dictionary")
    For Each cell In rng.Cells
        If Not IsError(cell.Value) Then
            If Len(Trim$(CStr(cell.Value))) > 0 Then
                v = cell.Value
                If Not dict.Exists(v) Then dict.Add v, v
            End If
        End If
    Next cell
    If dict.Count = 0 Then Exit Function
    ReDim out(1 To dict.Count, 1 To 1)
    i = 1
    For Each v In dict.Keys
        out(i, 1) = v
        i = i + 1
    Next v
    UniqueColumnValues = out
End Function

' One-hot matrix, nRows x nCategories. Blank / unmatched rows are all zeros.
Public Function DummyMatrix(ByVal rng As Range, ByRef categories As Variant) As Variant
    Dim nRows As Long
    Dim nCat As Long
    Dim r As Long
    Dim c As Long
    Dim out As Variant
    nRows = rng.Rows.Count
    nCat = UBound(categories, 1)
    ReDim out(1 To nRows, 1 To nCat)
    For r = 1 To nRows
        For c = 1 To nCat
            If rng.Cells(r, 1).Value = categories(c, 1) Then
                out(r, c) = 1
            Else
                out(r, c) = 0
            End If
        Next c
    Next r
    DummyMatrix = out
End Function

' Write a 2D block at dest. Title in the row above when dest is not on row 1.
Public Sub WriteBlockToRight(ByVal dest As Range, ByRef data As Variant, Optional ByVal Title As String = "")
    Dim nRows As Long
    Dim nCols As Long
    nRows = UBound(data, 1) - LBound(data, 1) + 1
    nCols = UBound(data, 2) - LBound(data, 2) + 1
    dest.Resize(nRows, nCols).Value = data
    If Len(Title) > 0 And dest.Row > 1 Then
        With dest.Offset(-1, 0)
            .Value = Title
            .Font.Bold = True
        End With
    End If
End Sub
