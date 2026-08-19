Attribute VB_Name = "modInternalProtection"
Option Explicit

' Internal: scroll limits, default-password sheet protect, unhide (Personal Custom_Menu7_Protection).
' Called from modApiProtection. Do not document these as the external API.
' DefaultProtectPassword is a known convenience password (same as Personal), not a secret.

Public Const DefaultProtectPassword As String = "WYSIWYG"

Public Function RequireWorksheet(Optional ByVal Title As String = "Protection") As Worksheet
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "The active sheet must be a worksheet.", vbExclamation, Title
        Exit Function
    End If
    Set RequireWorksheet = ActiveSheet
End Function

Public Sub SetScrollAreaToSelection(ByVal ws As Worksheet)
    Dim rng As Range
    If TypeName(Selection) <> "Range" Then
        MsgBox "Select a range on this worksheet first.", vbExclamation, "Protection"
        Exit Sub
    End If
    Set rng = Selection
    If Not rng.Worksheet Is ws Then
        MsgBox "Select a range on the active worksheet.", vbExclamation, "Protection"
        Exit Sub
    End If
    ws.ScrollArea = ""
    ws.ScrollArea = rng.Areas(1).Address
End Sub

Public Sub ClearScrollArea(ByVal ws As Worksheet)
    ws.ScrollArea = ""
End Sub

Public Sub ProtectSheetDefault(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Unprotect Password:=DefaultProtectPassword
    On Error GoTo 0
    ws.Protect Password:=DefaultProtectPassword, UserInterfaceOnly:=True
End Sub

Public Sub UnprotectSheetDefault(ByVal ws As Worksheet)
    ws.Unprotect Password:=DefaultProtectPassword
End Sub

Public Function UnhideHiddenWorksheets(ByVal wb As Workbook) As Long
    Dim ws As Worksheet
    Dim n As Long
    For Each ws In wb.Worksheets
        If ws.Visible <> xlSheetVisible Then
            ws.Visible = xlSheetVisible
            n = n + 1
        End If
    Next ws
    UnhideHiddenWorksheets = n
End Function

Public Sub UnhideAllRowsAndColumnsOnSheet(ByVal ws As Worksheet)
    ws.Cells.EntireRow.Hidden = False
    ws.Cells.EntireColumn.Hidden = False
End Sub
