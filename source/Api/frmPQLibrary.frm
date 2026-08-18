VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPQLibrary 
   Caption         =   "Power Query library"
   ClientHeight    =   10005
   ClientLeft      =   150
   ClientTop       =   585
   ClientWidth     =   19620
   OleObjectBlob   =   "frmPQLibrary.frx":0000
   StartUpPosition   =   1  'CenterOwner
End
Attribute VB_Name = "frmPQLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Personal frmPQLibrary: browse tblPQ_Library, import selected queries (with fn* deps),
' or export the active workbook's queries to a PQ_Functions table.
' Preview looks up by function name so a search filter does not show the wrong M code.

Private loLib As ListObject
Private mReady As Boolean

Private Sub UserForm_Initialize()
    Call ApplyProfessionalLayout
    Set loLib = modInternalPowerQuery.GetMasterLibraryTable()
    mReady = Not loLib Is Nothing
    If mReady Then Call LoadFunctionList("")
End Sub

Private Sub UserForm_Activate()
    If mReady Then Exit Sub
    MsgBox "No table named tblPQ_Library was found." & vbCrLf & vbCrLf & _
           "Open Personal.xlsb (or another workbook that contains that table) and try again.", _
           vbExclamation, "Power Query"
    Unload Me
End Sub

Private Sub chkSelectAll_Click()
    Dim i As Long
    For i = 0 To lstFunctions.ListCount - 1
        lstFunctions.Selected(i) = chkSelectAll.Value
    Next i
End Sub

Private Sub txtSearch_Change()
    Call LoadFunctionList(txtSearch.Text)
End Sub

Private Sub lstFunctions_Change()
    Dim fnName As String
    Dim r As ListRow
    Dim idx As Long
    idx = lstFunctions.ListIndex
    If idx < 0 Then Exit Sub
    fnName = CStr(lstFunctions.List(idx, 0))
    Set r = FindLibraryRow(fnName)
    If r Is Nothing Then Exit Sub
    txtDescription.Text = modInternalPowerQuery.PqNz(r.Range(1, 5).Value)
    txtMCode.Text = modInternalPowerQuery.PqNz(r.Range(1, 6).Value)
    txtDependencies.Text = modInternalPowerQuery.GetDependenciesFromMCode(txtMCode.Text, fnName)
End Sub

Private Sub cmdImport_Click()
    Dim i As Long
    Dim fnName As String
    Dim total As Long
    Dim done As Long
    Dim dest As Workbook

    Set dest = ActiveWorkbook
    If dest Is Nothing Then
        MsgBox "Open a destination workbook first.", vbExclamation, "Power Query"
        Exit Sub
    End If
    If dest.Name = ThisWorkbook.Name Then
        MsgBox "The destination cannot be the add-in. Activate a workbook, then import.", vbExclamation, "Power Query"
        Exit Sub
    End If

    For i = 0 To lstFunctions.ListCount - 1
        If lstFunctions.Selected(i) Then total = total + 1
    Next i
    If total = 0 Then
        MsgBox "Select one or more functions to import.", vbInformation, "Power Query"
        Exit Sub
    End If

    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    For i = 0 To lstFunctions.ListCount - 1
        If lstFunctions.Selected(i) Then
            done = done + 1
            fnName = CStr(lstFunctions.List(i, 0))
            Application.StatusBar = "Importing " & CStr(done) & " of " & CStr(total) & ": " & fnName
            Call modInternalPowerQuery.ImportFunctionByName(fnName, loLib, dest)
        End If
    Next i
    On Error Resume Next
    dest.RefreshAll
    On Error GoTo 0
    Call modInternalExcelApp.PopAppState
    Application.StatusBar = False
    MsgBox CStr(done) & " function(s) imported (including dependencies).", vbInformation, "Power Query"
    Unload Me
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Application.StatusBar = False
    MsgBox Err.Description, vbExclamation, "Power Query"
End Sub

Private Sub cmdExportWorkbook_Click()
    Dim dest As Workbook
    Dim n As Long
    Set dest = ActiveWorkbook
    If dest Is Nothing Then
        MsgBox "Open a workbook first.", vbExclamation, "Power Query"
        Exit Sub
    End If
    If dest.Name = ThisWorkbook.Name Then
        MsgBox "The source cannot be the add-in. Activate a workbook, then export.", vbExclamation, "Power Query"
        Exit Sub
    End If
    On Error GoTo ExportEH
    Call modInternalExcelApp.PushAppState
    n = modInternalPowerQuery.ExportWorkbookQueriesToPQFunctions(dest)
    Call modInternalExcelApp.PopAppState
    If n = 1 Then
        MsgBox "1 query written to the PQ_Functions table.", vbInformation, "Power Query"
    Else
        MsgBox CStr(n) & " queries written to the PQ_Functions table.", vbInformation, "Power Query"
    End If
    Unload Me
    Exit Sub
ExportEH:
    Call modInternalExcelApp.PopAppState
    MsgBox Err.Description, vbExclamation, "Power Query"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub LoadFunctionList(ByVal filterText As String)
    Dim r As ListRow
    Dim match As Boolean
    Dim ft As String
    Dim fnName As String
    Dim category As String
    Dim description As String

    ft = LCase$(filterText)
    lstFunctions.Clear
    lstFunctions.ColumnCount = 3
    If loLib Is Nothing Then Exit Sub
    If loLib.ListRows.Count = 0 Then Exit Sub

    For Each r In loLib.ListRows
        fnName = modInternalPowerQuery.PqNz(r.Range(1, 1).Value)
        category = modInternalPowerQuery.PqNz(r.Range(1, 2).Value)
        description = modInternalPowerQuery.PqNz(r.Range(1, 5).Value)
        match = True
        If Len(ft) > 0 Then
            match = (InStr(1, LCase$(fnName), ft, vbBinaryCompare) > 0) _
                 Or (InStr(1, LCase$(category), ft, vbBinaryCompare) > 0) _
                 Or (InStr(1, LCase$(description), ft, vbBinaryCompare) > 0)
        End If
        If match Then
            lstFunctions.AddItem fnName
            lstFunctions.List(lstFunctions.ListCount - 1, 1) = category
            lstFunctions.List(lstFunctions.ListCount - 1, 2) = modInternalPowerQuery.PqNz(r.Range(1, 3).Value)
        End If
    Next r
End Sub

Private Function FindLibraryRow(ByVal fnName As String) As ListRow
    Dim r As ListRow
    If loLib Is Nothing Then Exit Function
    For Each r In loLib.ListRows
        If StrComp(modInternalPowerQuery.PqNz(r.Range(1, 1).Value), fnName, vbTextCompare) = 0 Then
            Set FindLibraryRow = r
            Exit Function
        End If
    Next r
End Function

Private Sub ApplyProfessionalLayout()
    Dim m As Single
    Dim g As Single
    Dim iw As Single
    Dim ih As Single
    Dim leftW As Single
    Dim rightL As Single
    Dim rightW As Single
    Dim btnW As Single
    Dim btnWide As Single
    Dim btnH As Single
    Dim btnT As Single
    Dim listT As Single
    Dim prevT As Single
    Dim col1 As Single
    Dim col2 As Single
    Dim col3 As Single
    Dim descH As Single
    Dim depH As Single

    m = 14
    g = 10
    btnW = 84
    btnWide = 118
    btnH = 24
    col1 = 168
    col2 = 118
    col3 = 64
    leftW = col1 + col2 + col3 + 18

    Me.Caption = "Power Query library"
    Me.Width = 780
    Me.Height = 530
    Me.BackColor = RGB(252, 252, 252)
    On Error Resume Next
    Me.Font.Name = "Segoe UI"
    Me.Font.Size = 9
    On Error GoTo 0

    iw = Me.InsideWidth
    ih = Me.InsideHeight
    If leftW > iw * 0.52 Then leftW = iw * 0.52
    rightL = m + leftW + g
    rightW = iw - rightL - m
    btnT = ih - m - btnH
    listT = 108
    descH = 56
    depH = 40

    On Error Resume Next
    Me.Controls("TextBox1").Visible = False
    Me.Controls("TextBox1").TabStop = False
    On Error GoTo 0

    Call PlaceLabel("lblTitle", m, 10, iw - m * 2, 18, "Power Query library", True, 12, RGB(32, 32, 32))
    Call PlaceLabel("lblSubtitle", m, 30, iw - m * 2, 16, _
        "Import selected functions into the active workbook, or export its queries to a PQ_Functions table.", _
        False, 8, RGB(96, 96, 96))
    Call PlaceLabel("lblRule", m, 50, iw - m * 2, 1, "", False, 8, RGB(210, 210, 210))
    Me.Controls("lblRule").BackStyle = 1
    Me.Controls("lblRule").BackColor = RGB(220, 222, 226)

    Call PlaceLabel("lblSearch", m, 60, 42, 16, "Search", False, 9, RGB(64, 64, 64))
    With txtSearch
        .Left = m + 44
        .Top = 58
        .Width = leftW - 44 - 92
        .Height = 20
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .TabIndex = 0
        .SpecialEffect = 0
        .BorderStyle = 1
    End With
    With chkSelectAll
        .Caption = "Select all"
        .Left = m + leftW - 88
        .Top = 60
        .Width = 88
        .Height = 16
        .BackStyle = 0
        .TabIndex = 1
        .Font.Name = "Segoe UI"
        .Font.Size = 9
    End With

    Call PlaceLabel("Label1", m, listT - 16, col1, 14, "Function", True, 8, RGB(80, 80, 80))
    Call PlaceLabel("Label2", m + col1, listT - 16, col2, 14, "Category", True, 8, RGB(80, 80, 80))
    Call PlaceLabel("Label3", m + col1 + col2, listT - 16, col3, 14, "Version", True, 8, RGB(80, 80, 80))

    With lstFunctions
        .Left = m
        .Top = listT
        .Width = leftW
        .Height = btnT - g - listT
        .ColumnCount = 3
        .ColumnWidths = CStr(col1) & ";" & CStr(col2) & ";" & CStr(col3)
        .MultiSelect = 1
        .ListStyle = 1
        .IntegralHeight = True
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .SpecialEffect = 0
        .TabIndex = 2
    End With

    prevT = 58
    Call PlaceLabel("lblDescription", rightL, prevT, rightW, 14, "Description", True, 8, RGB(80, 80, 80))
    Call StylePreviewBox(txtDescription, rightL, prevT + 16, rightW, descH, "Segoe UI", 9, False, 3)
    prevT = prevT + 16 + descH + 8
    Call PlaceLabel("lblDependencies", rightL, prevT, rightW, 14, "Dependencies", True, 8, RGB(80, 80, 80))
    Call StylePreviewBox(txtDependencies, rightL, prevT + 16, rightW, depH, "Segoe UI", 9, True, 4)
    prevT = prevT + 16 + depH + 8
    Call PlaceLabel("lblMCode", rightL, prevT, rightW, 14, "M code", True, 8, RGB(80, 80, 80))
    Call StylePreviewBox(txtMCode, rightL, prevT + 16, rightW, btnT - g - (prevT + 16), "Consolas", 8, True, 5)

    With cmdImport
        .Caption = "Import"
        .Accelerator = "I"
        .Default = True
        .TakeFocusOnClick = False
        .Left = iw - m - btnW - g - btnWide - g - btnW
        .Top = btnT
        .Width = btnW
        .Height = btnH
        .TabIndex = 6
        .Font.Name = "Segoe UI"
        .Font.Size = 9
    End With
    With cmdExportWorkbook
        .Caption = "Export workbook"
        .Accelerator = "E"
        .TakeFocusOnClick = False
        .Left = iw - m - btnW - g - btnWide
        .Top = btnT
        .Width = btnWide
        .Height = btnH
        .TabIndex = 7
        .Font.Name = "Segoe UI"
        .Font.Size = 9
    End With
    With cmdCancel
        .Caption = "Close"
        .Cancel = True
        .TakeFocusOnClick = False
        .Left = iw - m - btnW
        .Top = btnT
        .Width = btnW
        .Height = btnH
        .TabIndex = 8
        .Font.Name = "Segoe UI"
        .Font.Size = 9
    End With
End Sub

Private Sub PlaceLabel(ByVal ctlName As String, ByVal l As Single, ByVal t As Single, _
        ByVal w As Single, ByVal h As Single, ByVal cap As String, _
        ByVal isBold As Boolean, ByVal fontSize As Single, ByVal fore As Long)
    Dim lbl As MSForms.Label
    On Error Resume Next
    Set lbl = Me.Controls(ctlName)
    On Error GoTo 0
    If lbl Is Nothing Then
        Set lbl = Me.Controls.Add("Forms.Label.1", ctlName, True)
    End If
    With lbl
        .Left = l
        .Top = t
        .Width = w
        .Height = h
        .Caption = cap
        .ForeColor = fore
        .BackStyle = 0
        .WordWrap = True
        .Font.Name = "Segoe UI"
        .Font.Size = fontSize
        .Font.Bold = isBold
    End With
End Sub

Private Sub StylePreviewBox(ByVal tb As MSForms.TextBox, ByVal l As Single, ByVal t As Single, _
        ByVal w As Single, ByVal h As Single, ByVal fontName As String, ByVal fontSize As Single, _
        ByVal codeLike As Boolean, ByVal tabIdx As Integer)
    With tb
        .Left = l
        .Top = t
        .Width = w
        .Height = h
        .MultiLine = True
        .WordWrap = Not codeLike
        If codeLike Then
            .ScrollBars = 3
        Else
            .ScrollBars = 2
        End If
        .EnterKeyBehavior = True
        .IntegralHeight = False
        .Locked = True
        .BackColor = RGB(248, 249, 250)
        .SpecialEffect = 0
        .BorderStyle = 1
        .TabIndex = tabIdx
        On Error Resume Next
        .Font.Name = fontName
        If Err.Number <> 0 Then .Font.Name = "Tahoma"
        On Error GoTo 0
        .Font.Size = fontSize
    End With
End Sub
