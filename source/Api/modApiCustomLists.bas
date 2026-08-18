Attribute VB_Name = "modApiCustomLists"
Option Explicit

' Public API: Excel custom lists and AutoCorrect entries (Personal Custom_Menu16_*).
' Other workbooks / the add-in menu should call these names only.
' Lists and replacements are stored in Excel (application-wide), not in the workbook.

''' @Description: Show how many custom lists Excel currently has (includes four built-in day/month lists).
''' @Example: CountCustomLists
Public Sub CountCustomLists()
    Dim n As Long
    Dim userN As Long
    n = modInternalCustomLists.CustomListCountValue()
    userN = n - modInternalCustomLists.BuiltInCustomListCount
    If userN < 0 Then userN = 0
    MsgBox "There are " & CStr(n) & " custom lists" & vbCrLf & vbCrLf & _
           CStr(modInternalCustomLists.BuiltInCustomListCount) & " built-in (days / months)" & vbCrLf & _
           CStr(userN) & " user-defined" & vbCrLf & vbCrLf & _
           "Custom lists are stored in Excel, not in this workbook.", _
           vbInformation, "Custom lists"
End Sub

''' @Description: Write every custom list to a sheet named Custom List properties (number, type, items).
''' @Example: ShowCustomLists
Public Sub ShowCustomLists()
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    Call modInternalCustomLists.WriteCustomListInventory
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("ShowCustomLists")
End Sub

''' @Description: Create custom lists from a range, one list per column (blank cells skipped). Application-wide.
''' @Example: CreateCustomListByColumn
Public Sub CreateCustomListByColumn()
    Call CreateCustomLists(False)
End Sub

''' @Description: Create custom lists from a range, one list per row (blank cells skipped). Application-wide.
''' @Example: CreateCustomListByRow
Public Sub CreateCustomListByRow()
    Call CreateCustomLists(True)
End Sub

''' @Description: Delete user custom lists whose numbers are in the selected (or prompted) range. Built-in lists 1-4 cannot be deleted. Highest numbers are deleted first so remaining numbers stay valid.
''' @Example: DeleteCustomList
Public Sub DeleteCustomList()
    Dim rng As Range
    Dim deleted As Long
    Dim skippedBuiltIn As Long
    Dim skippedInvalid As Long
    Dim ans As VbMsgBoxResult
    On Error GoTo EH
    Set rng = modInternalCustomLists.PromptListRange( _
        "Select cells that contain custom list numbers to delete." & vbCrLf & vbCrLf & _
        "Use column A of the Custom List properties sheet. Built-in lists 1-" & _
        CStr(modInternalCustomLists.BuiltInCustomListCount) & " cannot be deleted." & vbCrLf & _
        "Lists are stored in Excel, not this workbook.")
    If rng Is Nothing Then Exit Sub
    Set rng = modInternalCustomLists.UsedPart(rng)
    If rng Is Nothing Then
        MsgBox "The selected range has no used cells.", vbExclamation, "Custom lists"
        Exit Sub
    End If
    ans = MsgBox("Delete the user-defined custom lists whose numbers are in " & rng.Address(False, False) & "?" & vbCrLf & vbCrLf & _
                 "This changes Excel for all workbooks and cannot be undone.", _
                 vbYesNo + vbExclamation + vbDefaultButton2, "Custom lists")
    If ans <> vbYes Then Exit Sub
    Call modInternalExcelApp.PushAppState
    deleted = modInternalCustomLists.DeleteCustomListsByNumbers(rng, skippedBuiltIn, skippedInvalid)
    If deleted > 0 Then Call modInternalCustomLists.WriteCustomListInventory
    Call modInternalExcelApp.PopAppState
    MsgBox DeleteResultMessage(deleted, skippedBuiltIn, skippedInvalid), vbInformation, "Custom lists"
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("DeleteCustomList")
End Sub

''' @Description: Worksheet function: number of Excel custom lists (volatile).
''' @Example: =NumCustomLists()
Public Function NumCustomLists() As Long
    Application.Volatile
    NumCustomLists = modInternalCustomLists.CustomListCountValue()
End Function

''' @Description: Write Excel AutoCorrect replacements to a sheet named Auto correct List.
''' @Example: AutoCorrectEntries_Display
Public Sub AutoCorrectEntries_Display()
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    Call modInternalCustomLists.WriteAutoCorrectInventory
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("AutoCorrectEntries_Display")
End Sub

''' @Description: Add or overwrite AutoCorrect entries from a two-column range (Replace | With). Does not remove existing entries that are absent from the range. Application-wide.
''' @Example: AutoCorrectEntries_Add
Public Sub AutoCorrectEntries_Add()
    Dim rng As Range
    Dim n As Long
    On Error GoTo EH
    Set rng = modInternalCustomLists.PromptListRange( _
        "Select a two-column range: Replace in the first column, With in the second." & vbCrLf & vbCrLf & _
        "A header row of Replace / With is skipped. Existing AutoCorrect entries not in the range are left unchanged.", _
        "Custom lists")
    If rng Is Nothing Then Exit Sub
    Set rng = modInternalCustomLists.UsedPart(rng)
    If rng Is Nothing Then
        MsgBox "The selected range has no used cells.", vbExclamation, "Custom lists"
        Exit Sub
    End If
    If rng.Columns.Count < 2 Then
        MsgBox "Select two columns (Replace and With).", vbExclamation, "Custom lists"
        Exit Sub
    End If
    n = modInternalCustomLists.AddAutoCorrectFromRange(rng)
    MsgBox CStr(n) & " AutoCorrect " & Plural(n, "entry", "entries") & " added or updated.", _
           vbInformation, "Custom lists"
    Exit Sub
EH:
    Call modInternalError.RaiseCurrent("AutoCorrectEntries_Add")
End Sub

Private Sub CreateCustomLists(ByVal byRow As Boolean)
    Dim rng As Range
    Dim added As Long
    Dim skippedExist As Long
    Dim tooShort As Long
    Dim orient As String
    On Error GoTo EH
    If byRow Then
        orient = "row"
    Else
        orient = "column"
    End If
    Set rng = modInternalCustomLists.PromptListRange( _
        "Select the range to turn into custom lists (one list per " & orient & ")." & vbCrLf & vbCrLf & _
        "Blank cells are skipped. A " & orient & " needs at least two values. Lists already in Excel are skipped." & vbCrLf & _
        "Custom lists are stored in Excel, not this workbook.")
    If rng Is Nothing Then Exit Sub
    Set rng = modInternalCustomLists.UsedPart(rng)
    If rng Is Nothing Then
        MsgBox "The selected range has no used cells.", vbExclamation, "Custom lists"
        Exit Sub
    End If
    If rng.Cells.Count = 1 Then
        MsgBox "Cannot create a custom list from just one cell.", vbExclamation, "Custom lists"
        Exit Sub
    End If
    Call modInternalCustomLists.AddCustomListsFromRange(rng, byRow, added, skippedExist, tooShort)
    MsgBox CreateResultMessage(added, skippedExist, tooShort, orient), vbInformation, "Custom lists"
    Exit Sub
EH:
    Call modInternalError.RaiseCurrent("CreateCustomList")
End Sub

Private Function CreateResultMessage(ByVal added As Long, ByVal skippedExist As Long, _
    ByVal tooShort As Long, ByVal orient As String) As String
    Dim msg As String
    msg = CStr(added) & " custom " & Plural(added, "list", "lists") & " created (" & orient & "-wise)."
    If skippedExist > 0 Then
        msg = msg & vbCrLf & CStr(skippedExist) & " already existed and " & Plural(skippedExist, "was", "were") & " skipped."
    End If
    If tooShort > 0 Then
        msg = msg & vbCrLf & CStr(tooShort) & " " & orient & Plural(tooShort, "", "s") & _
              " had fewer than two values and " & Plural(tooShort, "was", "were") & " skipped."
    End If
    CreateResultMessage = msg
End Function

Private Function DeleteResultMessage(ByVal deleted As Long, ByVal skippedBuiltIn As Long, _
    ByVal skippedInvalid As Long) As String
    Dim msg As String
    msg = CStr(deleted) & " custom " & Plural(deleted, "list", "lists") & " deleted."
    If skippedBuiltIn > 0 Then
        msg = msg & vbCrLf & CStr(skippedBuiltIn) & " built-in " & Plural(skippedBuiltIn, "list", "lists") & " cannot be deleted."
    End If
    If skippedInvalid > 0 Then
        msg = msg & vbCrLf & CStr(skippedInvalid) & " " & Plural(skippedInvalid, "number was", "numbers were") & " not a current list number."
    End If
    DeleteResultMessage = msg
End Function

Private Function Plural(ByVal n As Long, ByVal one As String, ByVal many As String) As String
    If n = 1 Then
        Plural = one
    Else
        Plural = many
    End If
End Function
