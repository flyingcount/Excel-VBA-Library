Attribute VB_Name = "modApiHyperlinks"
Option Explicit

' Public API: Personal Menu21 hyperlinks / workbook index.
' Other workbooks / the add-in menu should call these names only.

''' @Description: Write every workbook hyperlink (sheet, location, text, address) to a sheet named Hyperlinks.
''' @Example: HyperlinkInventory
Public Sub HyperlinkInventory()
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    Call modInternalHyperlinks.WriteHyperlinkInventory(ActiveWorkbook)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("HyperlinkInventory")
End Sub

''' @Description: Create an Index sheet with links to every worksheet, and insert a Back to Index row at the top of each unprotected sheet.
''' @Example: CreateIndex
Public Sub CreateIndex()
    Dim ans As VbMsgBoxResult
    On Error GoTo EH
    ans = MsgBox("Create a sheet named Index with links to every worksheet?" & vbCrLf & vbCrLf & _
                 "A 'Back to Index' row will be inserted at the top of each unprotected sheet.", _
                 vbYesNo + vbQuestion, "Hyperlinks")
    If ans <> vbYes Then Exit Sub
    Call modInternalExcelApp.PushAppState
    Call modInternalHyperlinks.CreateWorkbookIndex(ActiveWorkbook, True)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("CreateIndex")
End Sub

''' @Description: Rebuild an existing Index sheet and refresh A1 back-links. Does not insert rows.
''' @Example: UpdateIndex
Public Sub UpdateIndex()
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    Call modInternalHyperlinks.UpdateWorkbookIndex(ActiveWorkbook)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("UpdateIndex")
End Sub

''' @Description: Delete workbook hyperlinks whose Text to display matches any cell in the selected (or prompted) range.
''' @Example: RemovingHyperLink
Public Sub RemovingHyperLink()
    Dim rng As Range
    Dim n As Long
    On Error GoTo EH
    Set rng = PromptHyperlinkRange("Select cells that contain the hyperlink Text to display")
    If rng Is Nothing Then Exit Sub
    Call modInternalExcelApp.PushAppState
    n = modInternalHyperlinks.RemoveHyperlinksByDisplayText(ActiveWorkbook, rng)
    Call modInternalExcelApp.PopAppState
    MsgBox CStr(n) & " hyperlink(s) removed.", vbInformation, "Hyperlinks"
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("RemovingHyperLink")
End Sub

''' @Description: Follow each hyperlink in the selection (default browser / Excel for internal links).
''' @Example: OpenHyperlink
Public Sub OpenHyperlink()
    Dim rng As Range
    On Error GoTo EH
    If TypeName(Selection) <> "Range" Then
        MsgBox "Select a range that contains hyperlinks.", vbExclamation, "Hyperlinks"
        Exit Sub
    End If
    Set rng = Selection
    Call modInternalHyperlinks.FollowHyperlinksInRange(rng)
    Exit Sub
EH:
    Call modInternalError.RaiseCurrent("OpenHyperlink")
End Sub

''' @Description: On every other worksheet, put a hyperlink in the active cell that returns to A1 of the current sheet.
''' @Example: AddHyperlinksToCurrentSheetA1
Public Sub AddHyperlinksToCurrentSheetA1()
    Dim addr As String
    Dim n As Long
    Dim ans As VbMsgBoxResult
    On Error GoTo EH
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Select a worksheet cell first.", vbExclamation, "Hyperlinks"
        Exit Sub
    End If
    addr = ActiveCell.Address(False, False)
    ans = MsgBox("Place a hyperlink in cell " & addr & " of every other worksheet," & vbCrLf & _
                 "linking back to A1 of '" & ActiveSheet.Name & "'?" & vbCrLf & vbCrLf & _
                 "Protected sheets are skipped. Existing contents of those cells are cleared.", _
                 vbYesNo + vbQuestion, "Hyperlinks")
    If ans <> vbYes Then Exit Sub
    Call modInternalExcelApp.PushAppState
    n = modInternalHyperlinks.AddBackLinksToActiveSheetA1(ActiveWorkbook, ActiveSheet, addr)
    Call modInternalExcelApp.PopAppState
    MsgBox CStr(n) & " hyperlink(s) added.", vbInformation, "Hyperlinks"
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("AddHyperlinksToCurrentSheetA1")
End Sub

''' @Description: List every worksheet name (and hidden state) in a message box.
''' @Example: ShowAllWorksheetsInWorkbook
Public Sub ShowAllWorksheetsInWorkbook()
    MsgBox modInternalHyperlinks.WorksheetListMessage(ActiveWorkbook), vbInformation, "Worksheets"
End Sub

Private Function PromptHyperlinkRange(ByVal PromptText As String) As Range
    Dim rng As Range
    Dim defAddr As String
    If TypeName(Selection) = "Range" Then defAddr = Selection.Address
    On Error Resume Next
    If Len(defAddr) > 0 Then
        Set rng = Application.InputBox(Prompt:=PromptText, Title:="Hyperlinks", Default:=defAddr, Type:=8)
    Else
        Set rng = Application.InputBox(Prompt:=PromptText, Title:="Hyperlinks", Type:=8)
    End If
    On Error GoTo 0
    Set PromptHyperlinkRange = rng
End Function
