Attribute VB_Name = "modApiFiles"
Option Explicit

' Public API: folder pickers and file listings (Personal Custom_Menu24_ListFilesInFolder).
' Other workbooks / the add-in menu should call these names only.

Public Function PickFolder(Optional ByVal Title As String = "Select folder") As String
    PickFolder = modInternalFiles.PickFolderPath(Title)
End Function

Public Function ListFiles(ByVal FolderPath As String, ByVal ExtensionFilter As String) As Variant
    ' ExtensionFilter example: "*.csv" or "*.xlsx"
    ListFiles = modInternalFiles.ListFilePaths(FolderPath, ExtensionFilter)
End Function

''' @Description: List files in a chosen folder (not subfolders) on a sheet named File listing. If there are 1,000 or more files, a prompt shows the count and an estimated run time so you can cancel. A modeless dialog with Cancel stops the run and writes any files already collected.
''' @Example: ReturnFilesInSelectedFolder
Public Sub ReturnFilesInSelectedFolder()
    Dim folderPath As String
    On Error GoTo EH
    folderPath = modInternalFiles.PickFolderPath("Please select a folder to list files from")
    If Len(folderPath) = 0 Then Exit Sub
    Call modInternalExcelApp.PushAppState
    Call modInternalFiles.WriteFileListing(folderPath, False)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("ReturnFilesInSelectedFolder")
End Sub

''' @Description: List files in a chosen folder and its subfolders on a sheet named File listing. Hidden, system, and junction folders are skipped. If there are 1,000 or more files, a prompt shows the count and an estimated run time so you can cancel. A modeless dialog with Cancel stops the run and writes any files already collected.
''' @Example: BatchListAllFiles_FolderSubfolders
Public Sub BatchListAllFiles_FolderSubfolders()
    Dim folderPath As String
    On Error GoTo EH
    folderPath = modInternalFiles.PickFolderPath("Please select a folder to list files from (includes subfolders)")
    If Len(folderPath) = 0 Then Exit Sub
    Call modInternalExcelApp.PushAppState
    Call modInternalFiles.WriteFileListing(folderPath, True)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("BatchListAllFiles_FolderSubfolders")
End Sub
