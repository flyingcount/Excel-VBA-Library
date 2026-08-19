Attribute VB_Name = "modApiDuplicates"
Option Explicit

' Public API: flag, label, colour, and count duplicates (Personal Custom_Menu15_Duplicates).
' Other workbooks / the add-in menu should call these names only.
' Personal OnAction names are kept (ColorDuplicates vs ColourDuplicateValues*).

''' @Description: Insert cells to the right of a single-column range and write Duplicate beside later occurrences of each value (first occurrence is left blank). Blanks and errors are skipped.
''' @Example: FlagDuplicates
Public Sub FlagDuplicates()
    Dim rng As Range
    Set rng = modInternalDuplicates.PromptDuplicateRange("Select a single column to flag later duplicate values.")
    If rng Is Nothing Then Exit Sub
    If Not modInternalDuplicates.RequireSingleColumn(rng) Then Exit Sub
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    Call modInternalDuplicates.FlagDuplicatesInColumn(rng)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("FlagDuplicates")
End Sub

''' @Description: Insert cells to the right of a single-column range and give each duplicate group a letter (A, B, …) on every member, including the first. Unique values stay blank.
''' @Example: ReferenceDuplicates
Public Sub ReferenceDuplicates()
    Dim rng As Range
    Set rng = modInternalDuplicates.PromptDuplicateRange("Select a single column to label duplicate groups.")
    If rng Is Nothing Then Exit Sub
    If Not modInternalDuplicates.RequireSingleColumn(rng) Then Exit Sub
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    Call modInternalDuplicates.ReferenceDuplicatesInColumn(rng)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("ReferenceDuplicates")
End Sub

''' @Description: Fill each duplicate group in the range with a distinct colour. Unique, blank, and error cells are not filled.
''' @Example: ColorDuplicates
Public Sub ColorDuplicates()
    Dim rng As Range
    Set rng = modInternalDuplicates.PromptDuplicateRange("Select the range whose duplicate groups to colour.")
    If rng Is Nothing Then Exit Sub
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    Call modInternalDuplicates.ColorDuplicateGroups(rng)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("ColorDuplicates")
End Sub

''' @Description: Fill cells red when their value appears more than once in the same row of the selection. Need at least two columns.
''' @Example: ColourDuplicateValuesByRow
Public Sub ColourDuplicateValuesByRow()
    Dim rng As Range
    Set rng = modInternalDuplicates.PromptDuplicateRange("Select a range with at least two columns.")
    If rng Is Nothing Then Exit Sub
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    Call modInternalDuplicates.ColourDuplicatesByRow(rng)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("ColourDuplicateValuesByRow")
End Sub

''' @Description: Fill cells green when their value appears more than once in the same column of the selection. Need at least two rows.
''' @Example: ColourDuplicateValuesByColumn
Public Sub ColourDuplicateValuesByColumn()
    Dim rng As Range
    Set rng = modInternalDuplicates.PromptDuplicateRange("Select a range with at least two rows.")
    If rng Is Nothing Then Exit Sub
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    Call modInternalDuplicates.ColourDuplicatesByColumn(rng)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("ColourDuplicateValuesByColumn")
End Sub

''' @Description: Fill cells yellow when their value appears more than once anywhere in the selection.
''' @Example: ColourDuplicateValuesInSelection
Public Sub ColourDuplicateValuesInSelection()
    Dim rng As Range
    Set rng = modInternalDuplicates.PromptDuplicateRange("Select the range to colour duplicate values in.")
    If rng Is Nothing Then Exit Sub
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    Call modInternalDuplicates.ColourDuplicatesInRange(rng)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("ColourDuplicateValuesInSelection")
End Sub

''' @Description: Message box with how many cells in the selection have a value that appears more than once (blank and error cells are not counted).
''' @Example: DuplicateCountFromSelection
Public Sub DuplicateCountFromSelection()
    Dim rng As Range
    Dim n As Long
    Set rng = modInternalDuplicates.PromptDuplicateRange("Select the range to count duplicate cells in.")
    If rng Is Nothing Then Exit Sub
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    n = modInternalDuplicates.CountDuplicateCells(rng)
    Call modInternalExcelApp.PopAppState
    If rng.Cells.Count = 1 Then Exit Sub
    MsgBox "Number of duplicates in the selection is " & CStr(n), vbInformation, "Duplicates"
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("DuplicateCountFromSelection")
End Sub
