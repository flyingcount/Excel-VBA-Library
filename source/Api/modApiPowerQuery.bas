Attribute VB_Name = "modApiPowerQuery"
Option Explicit

' Public API: Personal Menu29 Power Query helpers.
' Other workbooks / the add-in menu should call these names only.
' Personal OnAction names are kept: ShowPQLibraryForm, BackgroundRefreshToggle,
' IgnorePrivacyToggle, Add_Connection_All_Tables.

Private Const TitlePq As String = "Power Query"

''' @Description: Show the Import or export Power Query queries and functions form (Personal Menu29).
''' @Example: ShowPQLibraryForm
Public Sub ShowPQLibraryForm()
    On Error GoTo EH
    If ActiveWorkbook Is Nothing Then
        MsgBox "Open a workbook first.", vbExclamation, TitlePq
        Exit Sub
    End If
    If ActiveWorkbook.Name = ThisWorkbook.Name Then
        MsgBox "Activate a workbook other than the add-in, then try again.", vbExclamation, TitlePq
        Exit Sub
    End If
    frmPQLibrary.Show
    Exit Sub
EH:
    Call ShowPqError
End Sub

Private Sub ShowPqError()
    Dim desc As String
    desc = Err.Description
    Call modInternalExcelApp.PopAppState
    If Len(desc) = 0 Then desc = "The Power Query operation could not be completed."
    MsgBox desc, vbExclamation, TitlePq
End Sub

''' @Description: Toggle BackgroundQuery on every OLEDB (Power Query) connection in the active workbook, then report the new state.
''' @Example: BackgroundRefreshToggle
Public Sub BackgroundRefreshToggle()
    Dim wb As Workbook
    Dim n As Long
    Dim nOn As Long
    Dim nOff As Long
    On Error GoTo EH
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "Open a workbook first.", vbExclamation, TitlePq
        Exit Sub
    End If
    n = modInternalPowerQuery.ToggleOledbBackgroundQuery(wb, nOn, nOff)
    If n = 0 Then
        MsgBox "No Power Query (OLEDB) connections in this workbook.", vbInformation, TitlePq
    ElseIf nOff = 0 Then
        MsgBox "Background refresh is now ON for " & CStr(nOn) & " connection(s).", vbInformation, TitlePq
    ElseIf nOn = 0 Then
        MsgBox "Background refresh is now OFF for " & CStr(nOff) & " connection(s).", vbInformation, TitlePq
    Else
        MsgBox "Toggled " & CStr(n) & " connection(s). Background refresh is on for " & _
               CStr(nOn) & " and off for " & CStr(nOff) & ".", vbInformation, TitlePq
    End If
    Exit Sub
EH:
    Call ShowPqError
End Sub

''' @Description: Toggle Workbook.Queries.FastCombine (Ignore Privacy Levels). Reports whether Fast Combine is now on or off.
''' @Example: IgnorePrivacyToggle
Public Sub IgnorePrivacyToggle()
    Dim wb As Workbook
    Dim nowOn As Boolean
    On Error GoTo EH
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "Open a workbook first.", vbExclamation, TitlePq
        Exit Sub
    End If
    If Not modInternalPowerQuery.HasQueries(wb) Then
        MsgBox "This workbook has no Power Query Queries collection (Excel 2016 or later is required).", vbExclamation, TitlePq
        Exit Sub
    End If
    nowOn = modInternalPowerQuery.ToggleFastCombine(wb)
    If nowOn Then
        MsgBox "Ignore Privacy Levels (Fast Combine) is now ON." & vbCrLf & vbCrLf & _
               "Power Query will combine data sources without privacy prompts.", vbInformation, TitlePq
    Else
        MsgBox "Ignore Privacy Levels (Fast Combine) is now OFF." & vbCrLf & vbCrLf & _
               "Privacy checks are enforced.", vbInformation, TitlePq
    End If
    Exit Sub
EH:
    Call ShowPqError
End Sub

''' @Description: Create a Power Query connection-only query for every Excel Table that does not already have one. Optionally load to the Data Model.
''' @Example: Add_Connection_All_Tables
Public Sub Add_Connection_All_Tables()
    Dim wb As Workbook
    Dim ans As VbMsgBoxResult
    Dim addToModel As Boolean
    Dim nTables As Long
    Dim created As Long
    Dim t0 As Double
    Dim modelNote As String
    On Error GoTo EH
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "Open a workbook first.", vbExclamation, TitlePq
        Exit Sub
    End If
    If Not modInternalPowerQuery.HasQueries(wb) Then
        MsgBox "This workbook has no Power Query Queries collection (Excel 2016 or later is required).", vbExclamation, TitlePq
        Exit Sub
    End If
    nTables = modInternalPowerQuery.CountWorkbookTables(wb)
    If nTables = 0 Then
        MsgBox "There are no Excel Tables (ListObjects) in this workbook.", vbInformation, TitlePq
        Exit Sub
    End If
    ans = MsgBox("Create Power Query connections for all " & CStr(nTables) & " table(s) in this workbook?" & vbCrLf & vbCrLf & _
                 "Existing queries for those tables are skipped.", _
                 vbYesNo + vbQuestion, TitlePq)
    If ans <> vbYes Then Exit Sub
    ans = MsgBox("Also add the data to the Data Model (Power Pivot)?", _
                 vbYesNo + vbQuestion + vbDefaultButton2, TitlePq)
    If ans = vbYes Then addToModel = True
    t0 = Timer
    Call modInternalExcelApp.PushAppState
    created = modInternalPowerQuery.AddConnectionsForAllTables(wb, addToModel)
    Call modInternalExcelApp.PopAppState
    modelNote = ""
    If addToModel Then modelNote = " (Data Model)"
    MsgBox CStr(created) & " connection(s) created" & modelNote & " in " & Format$(Timer - t0, "0.0") & " seconds." & vbCrLf & vbCrLf & _
           CStr(nTables - created) & " table(s) already had a query or connection.", _
           vbInformation, TitlePq
    Exit Sub
EH:
    Call ShowPqError
End Sub
