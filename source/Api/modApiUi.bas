Attribute VB_Name = "modApiUi"
Option Explicit

' Public API: lightweight UI / messaging.

Public Sub NotifyInfo(ByVal Message As String, Optional ByVal Title As String = "PowerQuery VBA Lib")
    MsgBox Message, vbInformation, Title
End Sub

Public Sub NotifyError(ByVal Message As String, Optional ByVal Title As String = "PowerQuery VBA Lib")
    MsgBox Message, vbExclamation, Title
End Sub
