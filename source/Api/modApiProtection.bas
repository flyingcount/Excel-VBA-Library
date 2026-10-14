Attribute VB_Name = "modApiProtection"
Option Explicit

' Public API: scroll area, default-password sheet protect, unhide (Personal Custom_Menu7_Protection).
' Other workbooks / the add-in menu should call these names only.
' Personal OnAction names are kept.
' The default password is the Personal constant WYSIWYG: a known convenience value, not a secret.

''' @Description: Limit scrolling on the active worksheet to the first area of the current selection.
''' @Example: LimitScrollArea
Public Sub LimitScrollArea()
    Dim ws As Worksheet
    Set ws = modInternalProtection.RequireWorksheet()
    If ws Is Nothing Then Exit Sub
    On Error GoTo EH
    Call modInternalProtection.SetScrollAreaToSelection(ws)
    Exit Sub
EH:
    Call modInternalError.RaiseCurrent("LimitScrollArea")
End Sub

''' @Description: Remove any scroll-area limit on the active worksheet.
''' @Example: ResetScrollArea
Public Sub ResetScrollArea()
    Dim ws As Worksheet
    Set ws = modInternalProtection.RequireWorksheet()
    If ws Is Nothing Then Exit Sub
    On Error GoTo EH
    Call modInternalProtection.ClearScrollArea(ws)
    Exit Sub
EH:
    Call modInternalError.RaiseCurrent("ResetScrollArea")
End Sub

''' @Description: Protect the active worksheet with the default password, UserInterfaceOnly so VBA can still edit this session.
''' @Example: ProtectWorksheet
Public Sub ProtectWorksheet()
    Dim ws As Worksheet
    Set ws = modInternalProtection.RequireWorksheet()
    If ws Is Nothing Then Exit Sub
    On Error GoTo EH
    Call modInternalProtection.ProtectSheetDefault(ws)
    Exit Sub
EH:
    MsgBox "Could not protect the worksheet. It may already be protected with a different password.", vbExclamation, "Protection"
End Sub

''' @Description: Unprotect the active worksheet using the default password.
''' @Example: UnProtectWorksheet
Public Sub UnProtectWorksheet()
    Dim ws As Worksheet
    Set ws = modInternalProtection.RequireWorksheet()
    If ws Is Nothing Then Exit Sub
    On Error GoTo EH
    Call modInternalProtection.UnprotectSheetDefault(ws)
    Exit Sub
EH:
    MsgBox "Could not unprotect the worksheet. It may be protected with a different password.", vbExclamation, "Protection"
End Sub

''' @Description: Show the default worksheet-protection password (Personal convenience password, not a secret).
''' @Example: DisplayPassword
Public Sub DisplayPassword()
    MsgBox "The password is : " & modInternalProtection.DefaultProtectPassword, vbInformation, "Protection"
End Sub

''' @Description: Make every hidden and very-hidden worksheet in the workbook visible.
''' @Example: UnHideAllSheets
Public Sub UnHideAllSheets()
    Dim n As Long
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    n = modInternalProtection.UnhideHiddenWorksheets(ActiveWorkbook)
    Call modInternalExcelApp.PopAppState
    If n = 0 Then
        MsgBox "There are no hidden worksheets in " & ActiveWorkbook.Name & ".", vbInformation, "Protection"
    Else
        MsgBox "Made " & CStr(n) & " worksheet(s) visible.", vbInformation, "Protection"
    End If
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("UnHideAllSheets")
End Sub

''' @Description: Unhide every row and column on the active worksheet without changing the selection.
''' @Example: UnhideAllRowsAndColumns
Public Sub UnhideAllRowsAndColumns()
    Dim ws As Worksheet
    Set ws = modInternalProtection.RequireWorksheet()
    If ws Is Nothing Then Exit Sub
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    Call modInternalProtection.UnhideAllRowsAndColumnsOnSheet(ws)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    MsgBox "Could not unhide rows and columns. Unprotect the worksheet first if it is protected.", vbExclamation, "Protection"
End Sub
