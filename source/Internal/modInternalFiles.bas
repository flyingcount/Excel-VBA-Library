Attribute VB_Name = "modInternalFiles"
Option Explicit

' Internal: folder file listings (Personal Custom_Menu24_ListFilesInFolder).
' Called from modApiFiles. Do not document these as the external API.

Public Const FileListingSheetName As String = "File listing"
Private Const FileColCount As Long = 12
' Detail listing (property reads) is slower than a Files.Count walk.
Private Const ListingTimeFactorLow As Double = 3#
Private Const ListingTimeFactorHigh As Double = 8#
Private Const WriteFilesPerSecond As Double = 4000#
Private Const PromptIfAtLeast As Long = 1000
Private mListingCancelled As Boolean

Public Function PickFolderPath(Optional ByVal Title As String = "Select folder") As String
    Dim dlg As FileDialog
    Set dlg = Application.FileDialog(msoFileDialogFolderPicker)
    With dlg
        .Title = Title
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath & "\"
        If .Show <> -1 Then
            PickFolderPath = vbNullString
            Exit Function
        End If
        If .SelectedItems.Count = 0 Then
            PickFolderPath = vbNullString
        Else
            PickFolderPath = .SelectedItems(1)
        End If
    End With
End Function

Public Sub RequestListingCancel()
    mListingCancelled = True
End Sub

Public Sub WriteFileListing(ByVal FolderPath As String, ByVal IncludeSubfolders As Boolean)
    Dim fso As Object
    Dim folder As Object
    Dim rows As Collection
    Dim n As Long
    Dim nFolders As Long
    Dim t0 As Double
    Dim countSecs As Double

    On Error GoTo EH

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(FolderPath) Then
        MsgBox "Folder not found:" & vbCrLf & FolderPath, vbExclamation, "Files"
        Exit Sub
    End If

    Set folder = fso.GetFolder(FolderPath)
    mListingCancelled = False
    Call ShowProgress("Counting files...")
    t0 = Timer
    Call CountFiles(folder, IncludeSubfolders, n, nFolders)
    countSecs = ElapsedSeconds(t0)

    If mListingCancelled Then
        Call CloseProgress
        MsgBox "Cancelled.", vbInformation, "Files"
        Exit Sub
    End If

    Call HideProgress

    If n = 0 Then
        Application.StatusBar = False
        MsgBox "No files found in " & FolderPath & ".", vbInformation, "Files"
        Exit Sub
    End If

    If n >= PromptIfAtLeast Then
        If Not ConfirmFileListing(n, nFolders, countSecs) Then Exit Sub
    End If

    mListingCancelled = False
    Call ShowProgress("Listing file details...")
    Set rows = New Collection
    Call CollectFiles(folder, IncludeSubfolders, rows)

    If mListingCancelled Then
        If rows.Count = 0 Then
            Call CloseProgress
            MsgBox "Cancelled before any files were listed.", vbInformation, "Files"
            Exit Sub
        End If
        Call DumpListing(rows, True)
        Call CloseProgress
        MsgBox "Cancelled. Wrote " & Format$(rows.Count, "#,##0") & " file" & PluralS(rows.Count) & " to File listing.", _
               vbInformation, "Files"
        Exit Sub
    End If

    n = rows.Count
    If n = 0 Then
        Call CloseProgress
        MsgBox "No files found in " & FolderPath & ".", vbInformation, "Files"
        Exit Sub
    End If

    Call DumpListing(rows, False)
    Call CloseProgress
    Exit Sub
EH:
    Call CloseProgress
    Err.Raise Err.Number, "WriteFileListing", Err.Description
End Sub

Public Function ListFilePaths(ByVal FolderPath As String, ByVal ExtensionFilter As String) As Variant
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim names() As String
    Dim n As Long
    Dim pat As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(FolderPath) Then
        ListFilePaths = Array()
        Exit Function
    End If

    pat = ExtensionFilter
    If Len(pat) = 0 Then pat = "*.*"
    Set folder = fso.GetFolder(FolderPath)
    If folder.Files.Count = 0 Then
        ListFilePaths = Array()
        Exit Function
    End If

    ReDim names(0 To folder.Files.Count - 1)
    n = 0
    For Each file In folder.Files
        If file.Name Like pat Then
            names(n) = file.Path
            n = n + 1
        End If
    Next file

    If n = 0 Then
        ListFilePaths = Array()
    Else
        ReDim Preserve names(0 To n - 1)
        ListFilePaths = names
    End If
End Function

Private Sub CountFiles(ByVal folder As Object, ByVal IncludeSubfolders As Boolean, _
    ByRef fileCount As Long, ByRef folderCount As Long)

    Dim nHere As Long
    Dim subFolder As Object
    Dim subs As Object

    folderCount = folderCount + 1
    If mListingCancelled Then Exit Sub
    If folderCount Mod 10 = 0 Then
        Call UpdateProgress("Counting files... " & Format$(fileCount, "#,##0") & " files, " & _
            Format$(folderCount, "#,##0") & " folders")
        DoEvents
        If mListingCancelled Then Exit Sub
    End If

    On Error Resume Next
    nHere = folder.Files.Count
    If Err.Number <> 0 Then
        Err.Clear
        nHere = 0
    End If
    On Error GoTo 0
    fileCount = fileCount + nHere

    If Not IncludeSubfolders Then Exit Sub

    On Error Resume Next
    Set subs = folder.SubFolders
    If Err.Number <> 0 Then
        Err.Clear
        Set subs = Nothing
    End If
    On Error GoTo 0
    If subs Is Nothing Then Exit Sub

    For Each subFolder In subs
        If mListingCancelled Then Exit For
        If Not SkipFolder(subFolder) Then
            Call CountFiles(subFolder, True, fileCount, folderCount)
        End If
    Next subFolder
End Sub

Private Function ConfirmFileListing(ByVal nFiles As Long, ByVal nFolders As Long, ByVal countSecs As Double) As Boolean
    Dim loSecs As Double
    Dim hiSecs As Double
    Dim defaultBtn As Long
    Dim msg As String

    loSecs = countSecs * ListingTimeFactorLow + (CDbl(nFiles) / WriteFilesPerSecond)
    hiSecs = countSecs * ListingTimeFactorHigh + (CDbl(nFiles) / (WriteFilesPerSecond / 2#))
    If loSecs < 1# And nFiles < 500 Then loSecs = 0#
    If hiSecs < loSecs Then hiSecs = loSecs
    If nFiles > 2000 And hiSecs < 2# Then hiSecs = 2#

    msg = "About " & Format$(nFiles, "#,##0") & " file" & PluralS(nFiles) & " in " & _
          Format$(nFolders, "#,##0") & " folder" & PluralS(nFolders) & "." & vbCrLf & vbCrLf & _
          "Counting took " & FormatDuration(countSecs) & "." & vbCrLf & _
          "Listing details may take " & FormatDurationRange(loSecs, hiSecs) & _
          " (estimate; network folders can be slower)." & vbCrLf & vbCrLf & _
          "Continue?"

    If nFiles > 10000 Or hiSecs >= 30# Then
        defaultBtn = vbDefaultButton2
    Else
        defaultBtn = vbDefaultButton1
    End If

    ConfirmFileListing = (MsgBox(msg, vbYesNo + vbQuestion + defaultBtn, "Files") = vbYes)
End Function

Private Function ElapsedSeconds(ByVal startTimer As Double) As Double
    Dim secs As Double
    secs = Timer - startTimer
    If secs < 0 Then secs = secs + 86400#
    ElapsedSeconds = secs
End Function

Private Function FormatDuration(ByVal secs As Double) As String
    Dim s As Long
    If secs < 0.5 Then
        FormatDuration = "less than 1 second"
        Exit Function
    End If
    s = CLng(secs)
    If s < 1 Then s = 1
    If s < 60 Then
        FormatDuration = CStr(s) & " second" & PluralS(s)
    ElseIf s < 3600 Then
        FormatDuration = CStr(s \ 60) & " minute" & PluralS(s \ 60)
        If (s Mod 60) >= 15 Then
            FormatDuration = FormatDuration & " " & CStr(s Mod 60) & " second" & PluralS(s Mod 60)
        End If
    Else
        FormatDuration = CStr(s \ 3600) & " hour" & PluralS(s \ 3600)
    End If
End Function

Private Function FormatDurationRange(ByVal loSecs As Double, ByVal hiSecs As Double) As String
    If hiSecs < 0.5 Then
        FormatDurationRange = "less than 1 second"
    ElseIf loSecs < 0.5 Or (hiSecs - loSecs) < 1# Then
        FormatDurationRange = "about " & FormatDuration(hiSecs)
    Else
        FormatDurationRange = "about " & FormatDuration(loSecs) & " to " & FormatDuration(hiSecs)
    End If
End Function

Private Function PluralS(ByVal n As Long) As String
    If n = 1 Then
        PluralS = ""
    Else
        PluralS = "s"
    End If
End Function

Private Sub CollectFiles(ByVal folder As Object, ByVal IncludeSubfolders As Boolean, ByVal rows As Collection)
    Dim file As Object
    Dim subFolder As Object
    Dim files As Object
    Dim subs As Object

    On Error Resume Next
    Set files = folder.Files
    If Err.Number <> 0 Then
        Err.Clear
        Set files = Nothing
    End If
    On Error GoTo 0

    If Not files Is Nothing Then
        For Each file In files
            If mListingCancelled Then Exit For
            rows.Add FileRow(file)
            If rows.Count Mod 50 = 0 Then
                Call UpdateProgress("Listing file details... " & Format$(rows.Count, "#,##0") & " files")
                DoEvents
                If mListingCancelled Then Exit For
            End If
        Next file
    End If

    If Not IncludeSubfolders Then Exit Sub

    On Error Resume Next
    Set subs = folder.SubFolders
    If Err.Number <> 0 Then
        Err.Clear
        Set subs = Nothing
    End If
    On Error GoTo 0

    If subs Is Nothing Then Exit Sub
    For Each subFolder In subs
        If mListingCancelled Then Exit For
        If Not SkipFolder(subFolder) Then
            Call CollectFiles(subFolder, True, rows)
        End If
    Next subFolder
End Sub

Private Function SkipFolder(ByVal folder As Object) As Boolean
    Dim a As Long
    On Error Resume Next
    a = folder.Attributes
    If Err.Number <> 0 Then
        SkipFolder = True
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0
    ' Hidden, system, or reparse point / junction (avoids cycles).
    If (a And 2) <> 0 Then SkipFolder = True
    If (a And 4) <> 0 Then SkipFolder = True
    If (a And 1024) <> 0 Then SkipFolder = True
End Function

Private Function FileRow(ByVal file As Object) As Variant
    Dim a(1 To FileColCount) As Variant
    a(1) = SafeFileProp(file, "Name")
    a(2) = SafeFileProp(file, "Path")
    a(3) = SafeFileProp(file, "Size")
    a(4) = SafeFileProp(file, "Type")
    a(5) = SafeFileProp(file, "DateCreated")
    a(6) = SafeFileProp(file, "DateLastAccessed")
    a(7) = SafeFileProp(file, "DateLastModified")
    a(8) = SafeParentPath(file)
    a(9) = SafeFileProp(file, "ShortName")
    a(10) = SafeFileProp(file, "ShortPath")
    a(11) = SafeFileProp(file, "Attributes")
    a(12) = AttributeDescription(CLng(Val(CStr(a(11)))))
    FileRow = a
End Function

Private Function FileRowsToArray(ByVal rows As Collection) As Variant
    Dim arr As Variant
    Dim i As Long
    Dim j As Long
    Dim oneRow As Variant
    Dim n As Long

    n = rows.Count
    ReDim arr(1 To n + 1, 1 To FileColCount)
    arr(1, 1) = "Name"
    arr(1, 2) = "Path"
    arr(1, 3) = "Size(Bytes)"
    arr(1, 4) = "Type"
    arr(1, 5) = "Created"
    arr(1, 6) = "Last Accessed"
    arr(1, 7) = "Last Modified"
    arr(1, 8) = "Parent Folder"
    arr(1, 9) = "Short Name"
    arr(1, 10) = "Short Path"
    arr(1, 11) = "File Attribute Value"
    arr(1, 12) = "File Attribute Description"

    For i = 1 To n
        oneRow = rows(i)
        For j = 1 To FileColCount
            arr(i + 1, j) = oneRow(j)
        Next j
    Next i
    FileRowsToArray = arr
End Function

Private Sub DumpListing(ByVal rows As Collection, ByVal partial As Boolean)
    Dim arr As Variant
    Dim ws As Worksheet
    Dim n As Long
    Dim headerRow As Long

    n = rows.Count
    arr = FileRowsToArray(rows)
    Call modApiSheets.CreateOutputSheet(FileListingSheetName)
    Set ws = ActiveWorkbook.Worksheets(FileListingSheetName)
    If partial Then
        ws.Range("A1").Value = "Partial listing (cancelled) — " & Format$(n, "#,##0") & " files"
        ws.Range("A1").Font.Bold = True
        headerRow = 3
    Else
        headerRow = 1
    End If
    ws.Cells(headerRow, 1).Resize(n + 1, FileColCount).Value = arr
    Call FormatFileListingSheet(ws, n, headerRow)
End Sub

Private Sub FormatFileListingSheet(ByVal ws As Worksheet, ByVal n As Long, ByVal headerRow As Long)
    Dim lastRow As Long
    lastRow = headerRow + n
    ws.Cells(headerRow, 1).Resize(1, FileColCount).Font.Bold = True
    ws.Cells(headerRow + 1, 3).Resize(n, 1).NumberFormat = "#,##0"
    ws.Cells(headerRow + 1, 5).Resize(n, 3).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    ws.Cells(headerRow, 1).Resize(n + 1, FileColCount).AutoFilter
    ws.Columns("A:L").AutoFit
    If ws.Columns("B").ColumnWidth > 60 Then ws.Columns("B").ColumnWidth = 60
    If ws.Columns("H").ColumnWidth > 50 Then ws.Columns("H").ColumnWidth = 50
    If ws.Columns("J").ColumnWidth > 50 Then ws.Columns("J").ColumnWidth = 50
    If ws.Columns("L").ColumnWidth > 40 Then ws.Columns("L").ColumnWidth = 40
    ws.Activate
    On Error Resume Next
    ActiveWindow.FreezePanes = False
    ActiveWindow.SplitColumn = 0
    ActiveWindow.SplitRow = headerRow
    ActiveWindow.FreezePanes = True
    On Error GoTo 0
    Application.StatusBar = False
End Sub

Private Sub ShowProgress(ByVal msg As String)
    On Error Resume Next
    frmFileListingProgress.SetStatus msg
    If Not frmFileListingProgress.Visible Then
        frmFileListingProgress.Show vbModeless
    End If
    On Error GoTo 0
    DoEvents
End Sub

Private Sub UpdateProgress(ByVal msg As String)
    On Error Resume Next
    frmFileListingProgress.SetStatus msg
    Application.StatusBar = msg
    On Error GoTo 0
End Sub

Private Sub HideProgress()
    On Error Resume Next
    frmFileListingProgress.Hide
    On Error GoTo 0
End Sub

Private Sub CloseProgress()
    On Error Resume Next
    Unload frmFileListingProgress
    Application.StatusBar = False
    On Error GoTo 0
End Sub

Private Function SafeFileProp(ByVal file As Object, ByVal propName As String) As Variant
    On Error Resume Next
    Select Case propName
        Case "Name"
            SafeFileProp = file.Name
        Case "Path"
            SafeFileProp = file.Path
        Case "Size"
            SafeFileProp = file.Size
        Case "Type"
            SafeFileProp = file.Type
        Case "DateCreated"
            SafeFileProp = file.DateCreated
        Case "DateLastAccessed"
            SafeFileProp = file.DateLastAccessed
        Case "DateLastModified"
            SafeFileProp = file.DateLastModified
        Case "ShortName"
            SafeFileProp = file.ShortName
        Case "ShortPath"
            SafeFileProp = file.ShortPath
        Case "Attributes"
            SafeFileProp = file.Attributes
    End Select
    On Error GoTo 0
End Function

Private Function SafeParentPath(ByVal file As Object) As String
    On Error Resume Next
    SafeParentPath = file.ParentFolder.Path
    On Error GoTo 0
End Function

Private Function AttributeDescription(ByVal a As Long) As String
    Dim parts As String
    If a = 0 Then
        AttributeDescription = "Normal"
        Exit Function
    End If
    If (a And 1) <> 0 Then parts = parts & "Read-only, "
    If (a And 2) <> 0 Then parts = parts & "Hidden, "
    If (a And 4) <> 0 Then parts = parts & "System, "
    If (a And 8) <> 0 Then parts = parts & "Volume, "
    If (a And 16) <> 0 Then parts = parts & "Directory, "
    If (a And 32) <> 0 Then parts = parts & "Archive, "
    If (a And 128) <> 0 Then parts = parts & "Normal, "
    If (a And 256) <> 0 Then parts = parts & "Temporary, "
    If (a And 1024) <> 0 Then parts = parts & "Reparse point, "
    If (a And 2048) <> 0 Then parts = parts & "Compressed, "
    If (a And 4096) <> 0 Then parts = parts & "Offline, "
    If (a And 16384) <> 0 Then parts = parts & "Encrypted, "
    If Len(parts) = 0 Then
        AttributeDescription = "Unknown (" & CStr(a) & ")"
    Else
        AttributeDescription = Left$(parts, Len(parts) - 2)
    End If
End Function
