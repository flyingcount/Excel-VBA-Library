Attribute VB_Name = "modApiFiles"
Option Explicit

' Public API: folder/file pickers and simple listings.
' TODO: migrate Personal.xlsb file utilities here.

Public Function PickFolder(Optional ByVal Title As String = "Select folder") As String
    Dim dlg As FileDialog
    Set dlg = Application.FileDialog(msoFileDialogFolderPicker)
    With dlg
        .Title = Title
        .AllowMultiSelect = False
        If .Show <> -1 Then
            PickFolder = vbNullString
            Exit Function
        End If
        PickFolder = .SelectedItems(1)
    End With
End Function

Public Function ListFiles(ByVal FolderPath As String, ByVal ExtensionFilter As String) As Variant
    ' ExtensionFilter example: "*.csv" or "*.xlsx"
    ' TODO: replace with Personal implementation if richer logic exists
    Dim fso As Object, folder As Object, file As Object
    Dim names() As String
    Dim n As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(FolderPath) Then
        ListFiles = Array()
        Exit Function
    End If

    Set folder = fso.GetFolder(FolderPath)
    ReDim names(0 To folder.Files.Count - 1)
    n = 0
    For Each file In folder.Files
        If file.Name Like ExtensionFilter Then
            names(n) = file.Path
            n = n + 1
        End If
    Next file

    If n = 0 Then
        ListFiles = Array()
    Else
        ReDim Preserve names(0 To n - 1)
        ListFiles = names
    End If
End Function
