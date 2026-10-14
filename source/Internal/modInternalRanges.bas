Attribute VB_Name = "modInternalRanges"
Option Explicit

' Internal: range prompts, analysis, cleanse, unique values (Personal Custom_Menu14_*).
' Called from modApiRanges. Do not document these as the external API.

Public Const FrequencySheetName As String = "Frequency Analysis"
Public Const UniqueValuesSheetName As String = "Unique Values"
Public Const CharactersSheetName As String = "Characters and Codes"
Public Const RangeAnalysisSheetName As String = "Range_Analysis"

Public Function PromptRange(ByVal PromptText As String, Optional ByVal Title As String = "Ranges") As Range
    Dim rng As Range
    Dim defAddr As String
    If TypeName(Selection) = "Range" Then defAddr = Selection.Address
    On Error Resume Next
    If Len(defAddr) > 0 Then
        Set rng = Application.InputBox(Prompt:=PromptText, Title:=Title, Default:=defAddr, Type:=8)
    Else
        Set rng = Application.InputBox(Prompt:=PromptText, Title:=Title, Type:=8)
    End If
    On Error GoTo 0
    Set PromptRange = rng
End Function

' Empty if cancelled (InputBox Type 2 returns False).
Public Function PromptString(ByVal PromptText As String, ByVal DefaultValue As String, Optional ByVal Title As String = "Ranges") As Variant
    Dim resp As Variant
    resp = Application.InputBox(Prompt:=PromptText, Title:=Title, Default:=DefaultValue, Type:=2)
    If VarType(resp) = vbBoolean Then
        PromptString = Empty
    Else
        PromptString = CStr(resp)
    End If
End Function

' Empty if cancelled (InputBox Type 1 returns False).
Public Function PromptNumber(ByVal PromptText As String, ByVal DefaultValue As Variant, Optional ByVal Title As String = "Ranges") As Variant
    Dim resp As Variant
    resp = Application.InputBox(Prompt:=PromptText, Title:=Title, Default:=DefaultValue, Type:=1)
    If VarType(resp) = vbBoolean Then
        PromptNumber = Empty
    Else
        PromptNumber = resp
    End If
End Function

' Pick a worksheet by selecting a cell on it. Nothing if cancelled.
Public Function PromptWorksheet(ByVal PromptText As String, Optional ByVal Title As String = "Ranges") As Worksheet
    Dim rng As Range
    On Error Resume Next
    Set rng = Application.InputBox(Prompt:=PromptText, Title:=Title, Type:=8)
    On Error GoTo 0
    If rng Is Nothing Then Exit Function
    Set PromptWorksheet = rng.Worksheet
End Function

' Intersect with UsedRange so entire-column selections do not scan a million empty cells.
Public Function UsedPart(ByVal rng As Range) As Range
    Dim u As Range
    If rng Is Nothing Then Exit Function
    If rng.Rows.Count = rng.Worksheet.Rows.Count Or rng.Columns.Count = rng.Worksheet.Columns.Count Then
        On Error Resume Next
        Set u = rng.Worksheet.UsedRange
        On Error GoTo 0
        If u Is Nothing Then
            Set UsedPart = rng
        Else
            Set UsedPart = Intersect(rng, u)
        End If
    Else
        Set UsedPart = rng
    End If
End Function

Public Function RequireSelectionRange() As Range
    If TypeName(Selection) <> "Range" Then
        MsgBox "Select a range first.", vbExclamation, "Ranges"
        Exit Function
    End If
    Set RequireSelectionRange = Selection
End Function

Public Sub WriteRangeAnalysisSheet(ByVal src As Range)
    Dim stats As Variant
    Dim ws As Worksheet
    Dim i As Long
    stats = RangeStatsArray(src)
    Call modApiSheets.CreateOutputSheet(RangeAnalysisSheetName)
    Set ws = ActiveWorkbook.Worksheets(RangeAnalysisSheetName)
    For i = 1 To UBound(stats, 1)
        ws.Cells(i, 1).Value = stats(i, 1) & " = " & CStr(stats(i, 2))
    Next i
    ws.Columns(1).AutoFit
    ws.Activate
End Sub

Public Function RangeStatsMessage(ByVal src As Range) As String
    Dim stats As Variant
    Dim i As Long
    Dim msg As String
    stats = RangeStatsArray(src)
    msg = "Range attributes:"
    For i = 1 To UBound(stats, 1)
        msg = msg & vbCrLf & stats(i, 1) & " = " & CStr(stats(i, 2))
    Next i
    RangeStatsMessage = msg
End Function

Public Sub WriteFrequencyAnalysis(ByVal src As Range)
    Dim text As String
    Dim dict As Object
    Dim ws As Worksheet
    Dim arr As Variant
    Dim n As Long
    Dim i As Long
    Dim keys As Variant
    Dim code As Long
    Dim total As Long

    text = ConcatenateRangeText(src)
    total = Len(text)
    If total = 0 Then
        MsgBox "The selected range has no text to analyse.", vbExclamation, "Ranges"
        Exit Sub
    End If

    Set dict = CharacterCounts(text)
    n = dict.Count
    ReDim arr(1 To n + 1, 1 To 4)
    arr(1, 1) = "Character code"
    arr(1, 2) = "Character"
    arr(1, 3) = "Frequency"
    arr(1, 4) = "Frequency %"
    keys = SortedDictionaryKeys(dict)
    For i = 1 To n
        code = CLng(keys(i))
        arr(i + 1, 1) = code
        arr(i + 1, 2) = CharacterLabel(code, Empty)
        arr(i + 1, 3) = dict(code)
        arr(i + 1, 4) = dict(code) / total
    Next i

    Call modApiSheets.CreateOutputSheet(FrequencySheetName)
    Set ws = ActiveWorkbook.Worksheets(FrequencySheetName)
    ws.Range("A1").Value = "Text analysed"
    ws.Range("B1").Value = src.Address(External:=True)
    ws.Range("A2").Value = "Characters"
    ws.Range("B2").Value = total
    ws.Range("D1").Value = Left$(text, 32000)
    ws.Range("A4").Resize(n + 1, 4).Value = arr
    ws.Range("A4").Resize(1, 4).Font.Bold = True
    ws.Range("D5").Resize(n, 1).NumberFormat = "0.00%"
    ws.Columns("A:C").AutoFit
    ws.Columns("D").ColumnWidth = 14
    ws.Range("D1").WrapText = True
    ws.Columns("D").ColumnWidth = 40
    ws.Activate
End Sub

Public Sub WriteUniqueCharacters(ByVal src As Range)
    Dim dict As Object
    Dim cell As Range
    Dim i As Long
    Dim code As Long
    Dim ch As String
    Dim ws As Worksheet
    Dim keys As Variant
    Dim n As Long

    Set dict = CreateObject("Scripting.Dictionary")
    For Each cell In src.Cells
        If IsError(cell.Value) Then GoTo NextCell
        ch = CStr(cell.Value)
        For i = 1 To Len(ch)
            code = AscW(Mid$(ch, i, 1))
            If code < 0 Then code = code + 65536
            If Not dict.Exists(code) Then dict.Add code, Mid$(ch, i, 1)
        Next i
NextCell:
    Next cell

    If dict.Count = 0 Then
        MsgBox "No characters found in the selected range.", vbInformation, "Ranges"
        Exit Sub
    End If

    n = dict.Count
    keys = SortedDictionaryKeys(dict)
    Call modApiSheets.CreateOutputSheet(CharactersSheetName)
    Set ws = ActiveWorkbook.Worksheets(CharactersSheetName)
    ws.Range("A1").Value = "Unique character count for the range:"
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = src.Address(External:=True)
    ws.Range("B2").Value = n
    ws.Range("A4").Value = "Character code"
    ws.Range("B4").Value = "Character"
    ws.Range("A4:B4").Font.Bold = True
    For i = 1 To n
        code = CLng(keys(i))
        ws.Cells(i + 4, 1).Value = code
        ws.Cells(i + 4, 2).Value = CharacterLabel(code, dict(code))
    Next i
    ws.Columns("A:B").AutoFit
    ws.Activate
End Sub

Public Sub WriteUniqueValues(ByVal src As Range)
    Dim ws As Worksheet
    Dim cols() As Long
    Dim i As Long
    Dim nCols As Long

    Call modApiSheets.CreateOutputSheet(UniqueValuesSheetName)
    Set ws = ActiveWorkbook.Worksheets(UniqueValuesSheetName)
    src.Copy
    ws.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
    Application.CutCopyMode = False

    nCols = src.Columns.Count
    ReDim cols(1 To nCols)
    For i = 1 To nCols
        cols(i) = i
    Next i
    On Error Resume Next
    ws.UsedRange.RemoveDuplicates Columns:=cols, Header:=xlGuess
    ws.UsedRange.SpecialCells(xlCellTypeBlanks).Delete Shift:=xlShiftUp
    On Error GoTo 0
    ws.Activate
End Sub

Public Sub TransposeToTheRight(ByVal src As Range)
    Dim nRows As Long
    Dim nCols As Long
    Dim dest As Range
    Dim arr As Variant
    Dim out() As Variant
    Dim r As Long
    Dim c As Long

    nRows = src.Rows.Count
    nCols = src.Columns.Count
    Set dest = src.Worksheet.Cells(src.Row, src.Column + nCols + 1).Resize(nCols, nRows)
    dest.ClearFormats
    If src.Cells.Count = 1 Then
        dest.Value = src.Value
        Exit Sub
    End If
    arr = src.Value
    ReDim out(1 To nCols, 1 To nRows)
    For r = 1 To nRows
        For c = 1 To nCols
            out(c, r) = arr(r, c)
        Next c
    Next r
    dest.Value = out
End Sub

Public Sub PasteSelectionAsPicture(ByVal src As Range)
    Application.CutCopyMode = False
    src.Copy
    src.Worksheet.Pictures.Paste.Select
    Application.CutCopyMode = False
End Sub

Public Sub KeepDigitsInRange(ByVal rng As Range)
    Call MapCells(rng, "digits")
End Sub

Public Sub KeepLettersInRange(ByVal rng As Range)
    Call MapCells(rng, "letters")
End Sub

Public Sub KeepLettersAndDigitsInRange(ByVal rng As Range)
    Call MapCells(rng, "alnum")
End Sub

Public Sub TrimWhitespaceInRange(ByVal rng As Range)
    Dim cell As Range
    Dim s As String
    For Each cell In rng.Cells
        If VarType(cell.Value) = vbString Then
            s = CStr(cell.Value)
            s = Replace(s, Chr$(160), " ")
            s = Replace(s, vbTab, " ")
            s = Replace(s, vbCrLf, " ")
            s = Replace(s, vbCr, " ")
            s = Replace(s, vbLf, " ")
            cell.Value = Application.WorksheetFunction.Trim(s)
        End If
    Next cell
End Sub

Public Sub TrimWhitespaceRegexInRange(ByVal rng As Range)
    Dim cell As Range
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.Pattern = "[\s\t\n\r]+"
    For Each cell In rng.Cells
        If Not IsError(cell.Value) Then
            If VarType(cell.Value) = vbString Or Len(CStr(cell.Value)) > 0 Then
                cell.Value = regex.Replace(Trim$(CStr(cell.Value)), " ")
            End If
        End If
    Next cell
End Sub

Private Function RangeStatsArray(ByVal src As Range) As Variant
    Dim cell As Range
    Dim v As Variant
    Dim nFormula As Long
    Dim nArray As Long
    Dim nNumeric As Long
    Dim nText As Long
    Dim nBlank As Long
    Dim nError As Long
    Dim nEven As Long
    Dim nOdd As Long
    Dim nNonText As Long
    Dim nLogical As Long
    Dim nNA As Long
    Dim arr As Variant
    Dim n As Double

    For Each cell In src.Cells
        v = cell.Value
        If cell.HasFormula Then nFormula = nFormula + 1
        If cell.HasArray Then nArray = nArray + 1
        If IsError(v) Then
            nError = nError + 1
            nNonText = nNonText + 1
            If CStr(v) = CStr(CVErr(xlErrNA)) Then nNA = nNA + 1
        ElseIf IsEmpty(v) Then
            nBlank = nBlank + 1
            nNonText = nNonText + 1
        ElseIf VarType(v) = vbBoolean Then
            nLogical = nLogical + 1
            nNonText = nNonText + 1
        ElseIf VarType(v) = vbString Then
            nText = nText + 1
        ElseIf IsNumeric(v) Then
            nNumeric = nNumeric + 1
            nNonText = nNonText + 1
            n = CDbl(v)
            If n = Fix(n) Then
                If (n Mod 2) = 0 Then
                    nEven = nEven + 1
                Else
                    nOdd = nOdd + 1
                End If
            End If
        Else
            nNonText = nNonText + 1
        End If
    Next cell

    ReDim arr(1 To 15, 1 To 2)
    arr(1, 1) = "Range": arr(1, 2) = src.Address(External:=True)
    arr(2, 1) = "Formula": arr(2, 2) = nFormula
    arr(3, 1) = "Arrays": arr(3, 2) = nArray
    arr(4, 1) = "Numeric": arr(4, 2) = nNumeric
    arr(5, 1) = "Text": arr(5, 2) = nText
    arr(6, 1) = "Blanks": arr(6, 2) = nBlank
    arr(7, 1) = "Error": arr(7, 2) = nError
    arr(8, 1) = "Even": arr(8, 2) = nEven
    arr(9, 1) = "Odd": arr(9, 2) = nOdd
    arr(10, 1) = "Non text": arr(10, 2) = nNonText
    arr(11, 1) = "Logical": arr(11, 2) = nLogical
    arr(12, 1) = "N/A": arr(12, 2) = nNA
    arr(13, 1) = "Total cells in range": arr(13, 2) = src.Cells.Count
    arr(14, 1) = "Number of rows": arr(14, 2) = src.Rows.Count
    arr(15, 1) = "Number of columns": arr(15, 2) = src.Columns.Count
    RangeStatsArray = arr
End Function

Private Function ConcatenateRangeText(ByVal src As Range) As String
    Dim cell As Range
    Dim parts As String
    For Each cell In src.Cells
        If Not IsError(cell.Value) Then
            If Not IsEmpty(cell.Value) Then parts = parts & CStr(cell.Value)
        End If
    Next cell
    ConcatenateRangeText = parts
End Function

Private Function CharacterCounts(ByVal text As String) As Object
    Dim dict As Object
    Dim i As Long
    Dim code As Long
    Set dict = CreateObject("Scripting.Dictionary")
    For i = 1 To Len(text)
        code = AscW(Mid$(text, i, 1))
        If code < 0 Then code = code + 65536
        If dict.Exists(code) Then
            dict(code) = dict(code) + 1
        Else
            dict.Add code, 1
        End If
    Next i
    Set CharacterCounts = dict
End Function

Private Function SortedDictionaryKeys(ByVal dict As Object) As Variant
    Dim keys As Variant
    Dim i As Long
    Dim j As Long
    Dim tmp As Variant
    keys = dict.Keys
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If CLng(keys(j)) < CLng(keys(i)) Then
                tmp = keys(i)
                keys(i) = keys(j)
                keys(j) = tmp
            End If
        Next j
    Next i
    ReDim tmp(1 To dict.Count)
    For i = 0 To dict.Count - 1
        tmp(i + 1) = keys(i)
    Next i
    SortedDictionaryKeys = tmp
End Function

Private Function CharacterLabel(ByVal code As Long, ByVal fallback As Variant) As String
    Select Case code
        Case 9: CharacterLabel = "Tab"
        Case 10: CharacterLabel = "Line Feed"
        Case 13: CharacterLabel = "Carriage Return"
        Case 32: CharacterLabel = "Space"
        Case 160: CharacterLabel = "Non-breaking Space"
        Case Else
            If IsEmpty(fallback) Then
                CharacterLabel = ChrW$(code)
            ElseIf VarType(fallback) = vbString Then
                CharacterLabel = CStr(fallback)
            Else
                CharacterLabel = ChrW$(code)
            End If
    End Select
End Function

Private Sub MapCells(ByVal rng As Range, ByVal mode As String)
    Dim cell As Range
    Dim s As String
    Dim i As Long
    Dim ch As String
    Dim keep As String
    For Each cell In rng.Cells
        If IsError(cell.Value) Then GoTo NextCell
        s = CStr(cell.Value)
        keep = ""
        For i = 1 To Len(s)
            ch = Mid$(s, i, 1)
            Select Case mode
                Case "digits"
                    If ch Like "[0-9]" Then keep = keep & ch
                Case "letters"
                    If ch Like "[A-Za-z]" Then keep = keep & ch
                Case "alnum"
                    If ch Like "[A-Za-z0-9. ]" Then keep = keep & ch
            End Select
        Next i
        cell.Value = keep
NextCell:
    Next cell
End Sub
