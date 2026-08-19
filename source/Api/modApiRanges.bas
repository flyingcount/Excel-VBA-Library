Attribute VB_Name = "modApiRanges"
Option Explicit

' Public API: named ranges, range analysis, and cleanse (Personal Custom_Menu14_*).
' Other workbooks / the add-in menu should call these names only.
' Personal OnAction names are kept, including ConvertNamedRangeGlodalToLocalScope.

''' @Description: Write every named range in the workbook to a sheet named Range properties.
''' @Example: ListNamedRangeProperties
Public Sub ListNamedRangeProperties()
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    Call modInternalNamedRanges.WriteNamedRangeInventory(ActiveWorkbook)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("ListNamedRangeProperties")
End Sub

''' @Description: Paste the selected range as a picture on the active sheet.
''' @Example: PasteRangeAsPicture
Public Sub PasteRangeAsPicture()
    Dim src As Range
    Set src = modInternalRanges.RequireSelectionRange()
    If src Is Nothing Then Exit Sub
    Call modInternalRanges.PasteSelectionAsPicture(src)
End Sub

''' @Description: Fill named-range cells light yellow. Yes = whole workbook, No = active sheet.
''' @Example: HighlightRanges
Public Sub HighlightRanges()
    Dim ans As VbMsgBoxResult
    ans = MsgBox("Colour named ranges on:" & vbCrLf & _
                 "Yes = whole workbook" & vbCrLf & _
                 "No = active sheet" & vbCrLf & _
                 "Cancel = stop", _
                 vbYesNoCancel + vbQuestion, "Ranges")
    Select Case ans
        Case vbYes
            Call modInternalNamedRanges.HighlightNamedRanges(ActiveWorkbook.Names, True)
        Case vbNo
            Call modInternalNamedRanges.HighlightNamedRanges(ActiveSheet.Names, True)
    End Select
End Sub

''' @Description: Clear fill from named-range cells. Yes = whole workbook, No = active sheet.
''' @Example: DeHighlightRanges
Public Sub DeHighlightRanges()
    Dim ans As VbMsgBoxResult
    ans = MsgBox("Remove colour from named ranges on:" & vbCrLf & _
                 "Yes = whole workbook" & vbCrLf & _
                 "No = active sheet" & vbCrLf & _
                 "Cancel = stop", _
                 vbYesNoCancel + vbQuestion, "Ranges")
    Select Case ans
        Case vbYes
            Call modInternalNamedRanges.HighlightNamedRanges(ActiveWorkbook.Names, False)
        Case vbNo
            Call modInternalNamedRanges.HighlightNamedRanges(ActiveSheet.Names, False)
    End Select
End Sub

''' @Description: Delete named ranges whose names are in the selected (or prompted) cells. Confirms first.
''' @Example: DeleteNamedRanges
Public Sub DeleteNamedRanges()
    Dim rng As Range
    Dim missing As String
    Dim ans As VbMsgBoxResult
    Set rng = PromptNameCells("Select cells that contain named range names to delete.")
    If rng Is Nothing Then Exit Sub
    ans = MsgBox("Delete the named ranges listed in " & rng.Address(False, False) & "?" & vbCrLf & vbCrLf & _
                 "This cannot be undone.", _
                 vbYesNo + vbExclamation + vbDefaultButton2, "Ranges")
    If ans <> vbYes Then Exit Sub
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    missing = modInternalNamedRanges.DeleteSpecifiedNames(ActiveWorkbook, rng)
    Call modInternalNamedRanges.WriteNamedRangeInventory(ActiveWorkbook)
    Call modInternalExcelApp.PopAppState
    If Len(missing) > 0 Then MsgBox missing, vbExclamation, "Ranges"
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("DeleteNamedRanges")
End Sub

''' @Description: Create hidden workbook-scoped constants from a two-column range (name, value).
''' @Example: HiddenMasterDataWorkbookScope
Public Sub HiddenMasterDataWorkbookScope()
    Dim rng As Range
    Set rng = PromptWorkingRange("Select two columns: names in the left column, values in the right.")
    If rng Is Nothing Then Exit Sub
    If rng.Columns.Count < 2 Then
        MsgBox "Select two columns: names then values.", vbExclamation, "Ranges"
        Exit Sub
    End If
    On Error GoTo EH
    Call modInternalNamedRanges.CreateHiddenConstantsFromRows(ActiveWorkbook, rng)
    Exit Sub
EH:
    Call modInternalError.RaiseCurrent("HiddenMasterDataWorkbookScope")
End Sub

''' @Description: Create hidden worksheet-scoped constants from a two-column range (name, value). Prompts for the sheet.
''' @Example: HiddenMasterDataWorksheetScope
Public Sub HiddenMasterDataWorksheetScope()
    Dim rng As Range
    Dim ws As Worksheet
    Set rng = PromptWorkingRange("Select two columns: names in the left column, values in the right.")
    If rng Is Nothing Then Exit Sub
    If rng.Columns.Count < 2 Then
        MsgBox "Select two columns: names then values.", vbExclamation, "Ranges"
        Exit Sub
    End If
    Set ws = modInternalRanges.PromptWorksheet("Select a cell on the worksheet that should own the names.")
    If ws Is Nothing Then Exit Sub
    On Error GoTo EH
    Call modInternalNamedRanges.CreateHiddenConstantsFromRows(ActiveWorkbook, rng, ws)
    Exit Sub
EH:
    Call modInternalError.RaiseCurrent("HiddenMasterDataWorksheetScope")
End Sub

''' @Description: Create a hidden sheet-scoped named range for each column, named from the header row (header excluded from the range).
''' @Example: NamedRangeIntoNamedRangeColumns
Public Sub NamedRangeIntoNamedRangeColumns()
    Dim rng As Range
    Set rng = PromptWorkingRange("Select a range with a header row. Each column becomes a hidden sheet-scoped name.")
    If rng Is Nothing Then Exit Sub
    If rng.Cells.Count = 1 Then
        MsgBox "The range must be more than one cell.", vbExclamation, "Ranges"
        Exit Sub
    End If
    On Error GoTo EH
    Call modInternalNamedRanges.NameEachColumnFromHeader(rng)
    Exit Sub
EH:
    Call modInternalError.RaiseCurrent("NamedRangeIntoNamedRangeColumns")
End Sub

''' @Description: Create a hidden workbook-scoped named range for the selected cells.
''' @Example: CreateHiddenNamedRange
Public Sub CreateHiddenNamedRange()
    Dim rng As Range
    Dim nm As Variant
    Set rng = PromptWorkingRange("Select the range to name.")
    If rng Is Nothing Then Exit Sub
    nm = modInternalRanges.PromptString("Enter the name for the range.", "myRange")
    If IsEmpty(nm) Then Exit Sub
    Call modInternalNamedRanges.CreateHiddenRangeName(ActiveWorkbook, CStr(nm), rng)
End Sub

''' @Description: Create a hidden workbook-scoped name that refers to a text constant.
''' @Example: CreateHiddenNamedString
Public Sub CreateHiddenNamedString()
    Dim nm As Variant
    Dim txt As Variant
    nm = modInternalRanges.PromptString("Enter the name (letters, digits, underscore).", "myRange")
    If IsEmpty(nm) Then Exit Sub
    txt = modInternalRanges.PromptString("Enter the text to assign to the name.", "Default text", "Enter text")
    If IsEmpty(txt) Then Exit Sub
    Call modInternalNamedRanges.CreateHiddenConstant(ActiveWorkbook, CStr(nm), CStr(txt))
End Sub

''' @Description: Create a hidden workbook-scoped name that refers to a numeric constant.
''' @Example: CreateHiddenNamedNumber
Public Sub CreateHiddenNamedNumber()
    Dim nm As Variant
    Dim num As Variant
    nm = modInternalRanges.PromptString("Enter the name (letters, digits, underscore).", "myNumber")
    If IsEmpty(nm) Then Exit Sub
    num = modInternalRanges.PromptNumber("Enter the number to assign to the name.", 10, "Enter a number")
    If IsEmpty(num) Then Exit Sub
    Call modInternalNamedRanges.CreateHiddenConstant(ActiveWorkbook, CStr(nm), num)
End Sub

''' @Description: Create workbook names Divisor_Thousand (=1000) and Divisor_Millions (=1000000).
''' @Example: NameCreateConstantsAsNamedRanges
Public Sub NameCreateConstantsAsNamedRanges()
    Call modInternalNamedRanges.CreateCommonConstantNames(ActiveWorkbook)
End Sub

''' @Description: Create the same local-scope name on every worksheet, referring to the same address on that sheet.
''' @Example: CreateNamedRangeInAllWorksheets
Public Sub CreateNamedRangeInAllWorksheets()
    Dim rng As Range
    Dim nm As Variant
    Set rng = PromptWorkingRange("Select the address to name on every worksheet.")
    If rng Is Nothing Then Exit Sub
    nm = modInternalRanges.PromptString("Enter the name for the range.", "myRange")
    If IsEmpty(nm) Then Exit Sub
    Call modInternalNamedRanges.CreateLocalNameOnEverySheet(ActiveWorkbook, CStr(nm), rng)
End Sub

''' @Description: Hide every named range in the workbook (Name Manager still lists hidden names if Show hidden is on).
''' @Example: HideAllNamedRanges
Public Sub HideAllNamedRanges()
    Call modInternalNamedRanges.SetAllNamesVisible(ActiveWorkbook, False)
End Sub

''' @Description: Unhide every named range in the workbook.
''' @Example: UnhideAllNamedRanges
Public Sub UnhideAllNamedRanges()
    Call modInternalNamedRanges.SetAllNamesVisible(ActiveWorkbook, True)
End Sub

''' @Description: Unhide sheet-scoped named ranges on the active worksheet.
''' @Example: UnhideNamedRangesWorksheet
Public Sub UnhideNamedRangesWorksheet()
    Call modInternalNamedRanges.SetSheetNamesVisible(ActiveSheet, True)
End Sub

''' @Description: Hide sheet-scoped named ranges on the active worksheet.
''' @Example: HideNamedRangesWorksheet
Public Sub HideNamedRangesWorksheet()
    Call modInternalNamedRanges.SetSheetNamesVisible(ActiveSheet, False)
End Sub

''' @Description: Unhide named ranges whose names are in the selected (or prompted) cells, then refresh Range properties.
''' @Example: UnhideSpecifiedNamedRanges
Public Sub UnhideSpecifiedNamedRanges()
    Call ApplySpecifiedVisibility(True)
End Sub

''' @Description: Hide named ranges whose names are in the selected (or prompted) cells, then refresh Range properties.
''' @Example: HideSpecifiedNamedRanges
Public Sub HideSpecifiedNamedRanges()
    Call ApplySpecifiedVisibility(False)
End Sub

''' @Description: Copy unique rows from the selected range to a sheet named Unique Values (all selected columns are the key).
''' @Example: ListUniqueValues
Public Sub ListUniqueValues()
    Dim rng As Range
    Set rng = PromptWorkingRange("Select the range to list unique rows from.")
    If rng Is Nothing Then Exit Sub
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    Call modInternalRanges.WriteUniqueValues(rng)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("ListUniqueValues")
End Sub

''' @Description: Convert worksheet-scoped names listed in the selection to workbook scope.
''' @Example: ConvertNamedRangeLocalToGlobalScope
Public Sub ConvertNamedRangeLocalToGlobalScope()
    Dim rng As Range
    Dim missing As String
    Set rng = PromptNameCells("Select cells that contain worksheet-scoped named range names.")
    If rng Is Nothing Then Exit Sub
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    missing = modInternalNamedRanges.ConvertLocalNamesToGlobal(ActiveWorkbook, rng)
    Call modInternalNamedRanges.WriteNamedRangeInventory(ActiveWorkbook)
    Call modInternalExcelApp.PopAppState
    If Len(missing) > 0 Then MsgBox missing, vbExclamation, "Ranges"
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("ConvertNamedRangeLocalToGlobalScope")
End Sub

''' @Description: Convert workbook-scoped names listed in the selection to worksheet scope on a sheet you pick.
''' @Example: ConvertNamedRangeGlodalToLocalScope
Public Sub ConvertNamedRangeGlodalToLocalScope()
    Dim rng As Range
    Dim ws As Worksheet
    Dim missing As String
    Set rng = PromptNameCells("Select cells that contain workbook-scoped named range names.")
    If rng Is Nothing Then Exit Sub
    Set ws = modInternalRanges.PromptWorksheet("Select a cell on the worksheet that should own the names.")
    If ws Is Nothing Then Exit Sub
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    missing = modInternalNamedRanges.ConvertGlobalNamesToLocal(ActiveWorkbook, rng, ws)
    Call modInternalNamedRanges.WriteNamedRangeInventory(ActiveWorkbook)
    Call modInternalExcelApp.PopAppState
    If Len(missing) > 0 Then MsgBox missing, vbExclamation, "Ranges"
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("ConvertNamedRangeGlodalToLocalScope")
End Sub

''' @Description: Count formulae, arrays, numbers, text, blanks, errors, even/odd integers, and size of the selection; write sheet Range_Analysis.
''' @Example: RangeAnalysis
Public Sub RangeAnalysis()
    Dim rng As Range
    Set rng = PromptWorkingRange("Select the range to analyse.")
    If rng Is Nothing Then Exit Sub
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    Call modInternalRanges.WriteRangeAnalysisSheet(rng)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("RangeAnalysis")
End Sub

''' @Description: Same counts as RangeAnalysis, shown in a message box.
''' @Example: RangeAnalysisMessage
Public Sub RangeAnalysisMessage()
    Dim rng As Range
    Set rng = PromptWorkingRange("Select the range to analyse.")
    If rng Is Nothing Then Exit Sub
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    MsgBox modInternalRanges.RangeStatsMessage(rng), vbInformation, "Ranges"
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("RangeAnalysisMessage")
End Sub

''' @Description: Count each distinct character in the selection; write sheet Frequency Analysis (used characters only, with %).
''' @Example: FrequencyAnalysis
Public Sub FrequencyAnalysis()
    Dim rng As Range
    Set rng = PromptWorkingRange("Select the range whose characters to count.")
    If rng Is Nothing Then Exit Sub
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    Call modInternalRanges.WriteFrequencyAnalysis(rng)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("FrequencyAnalysis")
End Sub

''' @Description: Strip letters and symbols, leaving digits, in the selected range.
''' @Example: Remove_AlphaCharactersFromString
Public Sub Remove_AlphaCharactersFromString()
    Dim rng As Range
    Set rng = PromptWorkingRange("Select the range to keep digits only.")
    If rng Is Nothing Then Exit Sub
    Call modInternalRanges.KeepDigitsInRange(rng)
End Sub

''' @Description: Strip digits and symbols, leaving letters, in the selected range.
''' @Example: Remove_NumbersFromString
Public Sub Remove_NumbersFromString()
    Dim rng As Range
    Set rng = PromptWorkingRange("Select the range to keep letters only.")
    If rng Is Nothing Then Exit Sub
    Call modInternalRanges.KeepLettersInRange(rng)
End Sub

''' @Description: Strip symbols, leaving letters, digits, and spaces, in the selected range.
''' @Example: Remove_SpecialCharactersFromString
Public Sub Remove_SpecialCharactersFromString()
    Dim rng As Range
    Set rng = PromptWorkingRange("Select the range to keep letters, digits, and spaces.")
    If rng Is Nothing Then Exit Sub
    Call modInternalRanges.KeepLettersAndDigitsInRange(rng)
End Sub

''' @Description: Collapse NBSP, tabs, and line breaks to spaces and trim, in text cells of the selection.
''' @Example: ClearCellsThatOnlyContainWhitespaces
Public Sub ClearCellsThatOnlyContainWhitespaces()
    Dim rng As Range
    Set rng = PromptWorkingRange("Select the range to trim whitespace.")
    If rng Is Nothing Then Exit Sub
    Call modInternalRanges.TrimWhitespaceInRange(rng)
End Sub

''' @Description: Collapse runs of whitespace to a single space using a regular expression.
''' @Example: RemoveWhiteSpaces
Public Sub RemoveWhiteSpaces()
    Dim rng As Range
    Set rng = PromptWorkingRange("Select the range to trim whitespace (regex).")
    If rng Is Nothing Then Exit Sub
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    Call modInternalRanges.TrimWhitespaceRegexInRange(rng)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("RemoveWhiteSpaces")
End Sub

''' @Description: List each distinct character and its code on a sheet named Characters and Codes.
''' @Example: ListCharactersAndCodesInRange
Public Sub ListCharactersAndCodesInRange()
    Dim rng As Range
    Set rng = PromptWorkingRange("Select the range to list unique characters from.")
    If rng Is Nothing Then Exit Sub
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    Call modInternalRanges.WriteUniqueCharacters(rng)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("ListCharactersAndCodesInRange")
End Sub

''' @Description: Write a transposed copy of the selection one column to the right of it.
''' @Example: TransposeARange
Public Sub TransposeARange()
    Dim rng As Range
    Set rng = modInternalRanges.RequireSelectionRange()
    If rng Is Nothing Then Exit Sub
    Call modInternalRanges.TransposeToTheRight(rng)
End Sub

''' @Description: Create workbook names BaseAll and Base_1..n for each column. Asks whether the first row is a header.
''' @Example: SplitRangeAndNameEachColumn
Public Sub SplitRangeAndNameEachColumn()
    Dim rng As Range
    Dim baseName As Variant
    Dim ans As VbMsgBoxResult
    Dim hasHeader As Boolean
    Set rng = PromptWorkingRange("Select the range to split into named columns.")
    If rng Is Nothing Then Exit Sub
    baseName = modInternalRanges.PromptString("What do you want the output called?", "Output")
    If IsEmpty(baseName) Then Exit Sub
    ans = MsgBox("Does the range have a header row?", vbYesNoCancel + vbDefaultButton1, "Ranges")
    If ans = vbCancel Then Exit Sub
    hasHeader = (ans = vbYes)
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    Call modInternalNamedRanges.SplitColumnsToWorkbookNames(rng, CStr(baseName), hasHeader)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("SplitRangeAndNameEachColumn")
End Sub

Private Function PromptWorkingRange(ByVal PromptText As String) As Range
    Dim rng As Range
    Set rng = modInternalRanges.PromptRange(PromptText)
    If rng Is Nothing Then Exit Function
    Set rng = modInternalRanges.UsedPart(rng)
    If rng Is Nothing Then
        MsgBox "The selected range has no used cells.", vbExclamation, "Ranges"
        Exit Function
    End If
    Set PromptWorkingRange = rng
End Function

Private Function PromptNameCells(ByVal PromptText As String) As Range
    Dim rng As Range
    Set rng = PromptWorkingRange(PromptText)
    If rng Is Nothing Then Exit Function
    Set PromptNameCells = rng
End Function

Private Sub ApplySpecifiedVisibility(ByVal makeVisible As Boolean)
    Dim rng As Range
    Dim missing As String
    Dim label As String
    If makeVisible Then
        label = "unhide"
    Else
        label = "hide"
    End If
    Set rng = PromptNameCells("Select cells that contain named range names to " & label & ".")
    If rng Is Nothing Then Exit Sub
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    missing = modInternalNamedRanges.SetSpecifiedNamesVisible(ActiveWorkbook, rng, makeVisible)
    Call modInternalNamedRanges.WriteNamedRangeInventory(ActiveWorkbook)
    Call modInternalExcelApp.PopAppState
    If Len(missing) > 0 Then MsgBox missing, vbExclamation, "Ranges"
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("ApplySpecifiedVisibility")
End Sub
