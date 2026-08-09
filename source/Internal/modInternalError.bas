Attribute VB_Name = "modInternalError"
Option Explicit

' Internal: centralized error reporting for Api facades.

Public Sub RaiseCurrent(ByVal Context As String)
    Dim msg As String
    msg = Context & vbCrLf & "Err " & Err.Number & ": " & Err.Description
    Err.Raise Err.Number, Context, msg
End Sub

Public Sub LogMessage(ByVal Context As String, ByVal Message As String)
    ' TODO: route to a log sheet or Immediate window as preferred
    Debug.Print Now; Context; Message
End Sub
