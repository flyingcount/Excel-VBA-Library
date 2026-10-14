Attribute VB_Name = "modInternalDuplicates"
Option Explicit

' Internal: duplicate flags, labels, fills, and counts (Personal Custom_Menu15_Duplicates).
' Called from modApiDuplicates. Do not document these as the external API.

Private Const FlagLabel As String = "Duplicate"
Private Const ColorByRowIndex As Long = 3
Private Const ColorByColIndex As Long = 4
Private Const ColorInRangeIndex As Long = 6

Public Function PromptDuplicateRange(ByVal PromptText As String) As Range
    Dim rng As Range
    Set rng = modInternalRanges.PromptRange(PromptText, "Duplicates")
    If rng Is Nothing Then Exit Function
    Set rng = modInternalRanges.UsedPart(rng)
    If rng Is Nothing Then
        MsgBox "The selected range has no used cells.", vbExclamation, "Duplicates"
        Exit Function
    End If
    Set PromptDuplicateRange = rng
End Function

Public Function RequireSingleColumn(ByVal rng As Range) As Boolean
    If rng.Columns.Count <> 1 Then
        MsgBox "Select a single column. Flags are written in the cells to the right.", vbExclamation, "Duplicates"
        Exit Function
    End If
    RequireSingleColumn = True
End Function

Public Sub FlagDuplicatesInColumn(ByVal src As Range)
    Dim dict As Object
    Dim i As Long
    Dim n As Long
    Dim key As String
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    n = src.Rows.Count
    Call InsertCellsToTheRight(src)
    For i = 1 To n
        key = DupKey(src.Cells(i, 1).Value)
        If Len(key) = 0 Then GoTo NextI
        If dict.Exists(key) Then
            src.Cells(i, 1).Offset(0, 1).Value = FlagLabel
        Else
            dict.Add key, i
        End If
NextI:
    Next i
End Sub

' Letter labels (A, B, …) on every cell that belongs to a duplicate group, including the first.
Public Sub ReferenceDuplicatesInColumn(ByVal src As Range)
    Dim firstCell As Object
    Dim letters As Object
    Dim cell As Range
    Dim key As String
    Dim nGroups As Long
    Dim letter As String
    Dim prevBar As Variant
    Dim i As Long
    Dim n As Long

    Set firstCell = CreateObject("Scripting.Dictionary")
    Set letters = CreateObject("Scripting.Dictionary")
    firstCell.CompareMode = vbTextCompare
    letters.CompareMode = vbTextCompare
    Call InsertCellsToTheRight(src)

    prevBar = Application.DisplayStatusBar
    Application.DisplayStatusBar = True
    n = src.Cells.Count
    i = 0
    On Error GoTo EH
    For Each cell In src.Cells
        i = i + 1
        Application.StatusBar = "Processing " & CStr(i) & " of " & CStr(n)
        key = DupKey(cell.Value)
        If Len(key) = 0 Then GoTo NextCell
        If firstCell.Exists(key) Then
            If Not letters.Exists(key) Then
                nGroups = nGroups + 1
                letters.Add key, IndexToLetters(nGroups)
                firstCell(key).Offset(0, 1).Value = letters(key)
            End If
            cell.Offset(0, 1).Value = letters(key)
        Else
            firstCell.Add key, cell
        End If
NextCell:
    Next cell
    Application.StatusBar = False
    Application.DisplayStatusBar = prevBar
    Exit Sub
EH:
    Dim nErr As Long
    Dim sSrc As String
    Dim sDesc As String
    nErr = Err.Number
    sSrc = Err.Source
    sDesc = Err.Description
    Application.StatusBar = False
    Application.DisplayStatusBar = prevBar
    Err.Raise nErr, sSrc, sDesc
End Sub

' Distinct fill per duplicate group. Unique and blank cells are left unfilled.
Public Sub ColorDuplicateGroups(ByVal src As Range)
    Dim firstCell As Object
    Dim colors As Object
    Dim cell As Range
    Dim key As String
    Dim nGroups As Long
    Set firstCell = CreateObject("Scripting.Dictionary")
    Set colors = CreateObject("Scripting.Dictionary")
    firstCell.CompareMode = vbTextCompare
    colors.CompareMode = vbTextCompare
    For Each cell In src.Cells
        key = DupKey(cell.Value)
        If Len(key) = 0 Then GoTo NextCell
        If firstCell.Exists(key) Then
            If Not colors.Exists(key) Then
                nGroups = nGroups + 1
                colors.Add key, GroupColorIndex(nGroups)
                firstCell(key).Interior.ColorIndex = colors(key)
            End If
            cell.Interior.ColorIndex = colors(key)
        Else
            firstCell.Add key, cell
        End If
NextCell:
    Next cell
End Sub

Public Sub ColourDuplicatesByRow(ByVal src As Range)
    Dim rw As Range
    If src.Columns.Count = 1 Then
        MsgBox "Need at least two columns to find duplicates within a row.", vbExclamation, "Duplicates"
        Exit Sub
    End If
    src.Interior.ColorIndex = xlColorIndexNone
    For Each rw In src.Rows
        Call ColourWhereCountExceedsOne(rw, ColorByRowIndex)
    Next rw
End Sub

Public Sub ColourDuplicatesByColumn(ByVal src As Range)
    Dim col As Range
    If src.Rows.Count = 1 Then
        MsgBox "Need at least two rows to find duplicates within a column.", vbExclamation, "Duplicates"
        Exit Sub
    End If
    src.Interior.ColorIndex = xlColorIndexNone
    For Each col In src.Columns
        Call ColourWhereCountExceedsOne(col, ColorByColIndex)
    Next col
End Sub

Public Sub ColourDuplicatesInRange(ByVal src As Range)
    If src.Cells.Count = 1 Then
        MsgBox "Need more than one cell to look for duplicates.", vbExclamation, "Duplicates"
        Exit Sub
    End If
    src.Interior.ColorIndex = xlColorIndexNone
    Call ColourWhereCountExceedsOne(src, ColorInRangeIndex)
End Sub

Public Function CountDuplicateCells(ByVal src As Range) As Long
    Dim counts As Object
    Dim cell As Range
    Dim key As String
    Dim n As Long
    If src.Cells.Count = 1 Then
        MsgBox "Need more than one cell to look for duplicates.", vbExclamation, "Duplicates"
        Exit Function
    End If
    Set counts = CountKeys(src)
    For Each cell In src.Cells
        key = DupKey(cell.Value)
        If Len(key) > 0 Then
            If counts(key) > 1 Then n = n + 1
        End If
    Next cell
    CountDuplicateCells = n
End Function

Private Sub ColourWhereCountExceedsOne(ByVal block As Range, ByVal colorIdx As Long)
    Dim counts As Object
    Dim cell As Range
    Dim key As String
    Set counts = CountKeys(block)
    For Each cell In block.Cells
        key = DupKey(cell.Value)
        If Len(key) > 0 Then
            If counts(key) > 1 Then cell.Interior.ColorIndex = colorIdx
        End If
    Next cell
End Sub

Private Function CountKeys(ByVal block As Range) As Object
    Dim dict As Object
    Dim cell As Range
    Dim key As String
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    For Each cell In block.Cells
        key = DupKey(cell.Value)
        If Len(key) > 0 Then
            If dict.Exists(key) Then
                dict(key) = dict(key) + 1
            Else
                dict.Add key, 1
            End If
        End If
    Next cell
    Set CountKeys = dict
End Function

' Empty / error / blank text → skip. Other values compared case-insensitively as text.
Private Function DupKey(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    DupKey = Trim$(CStr(v))
    If Len(DupKey) = 0 Then DupKey = ""
End Function

Private Sub InsertCellsToTheRight(ByVal src As Range)
    src.Offset(0, 1).Resize(src.Rows.Count, 1).Insert Shift:=xlToRight
End Sub

Private Function IndexToLetters(ByVal n As Long) As String
    Dim s As String
    If n < 1 Then n = 1
    Do
        n = n - 1
        s = Chr$(65 + (n Mod 26)) & s
        n = n \ 26
    Loop While n > 0
    IndexToLetters = s
End Function

Private Function GroupColorIndex(ByVal groupNum As Long) As Long
    Dim palette As Variant
    Dim n As Long
    palette = Array(3, 4, 5, 6, 7, 8, 17, 18, 19, 20, 22, 23, 24, 26, 27, 28, 33, 34, 35, 36, 37, 38, 39, 40, 43, 44, 45, 46, 47, 50, 51, 52, 53, 54)
    n = UBound(palette) - LBound(palette) + 1
    GroupColorIndex = palette(LBound(palette) + ((groupNum - 1) Mod n))
End Function
