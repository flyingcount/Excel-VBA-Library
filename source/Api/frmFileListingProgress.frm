VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFileListingProgress 
   Caption         =   "Files"
   ClientHeight    =   1695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5520
   OleObjectBlob   =   "frmFileListingProgress.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFileListingProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub SetStatus(ByVal msg As String)
    lblStatus.Caption = msg
End Sub

Private Sub cmdCancel_Click()
    cmdCancel.Enabled = False
    lblStatus.Caption = "Stopping..."
    Call modInternalFiles.RequestListingCancel
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode <> 1 Then
        Call modInternalFiles.RequestListingCancel
        lblStatus.Caption = "Stopping..."
        On Error Resume Next
        cmdCancel.Enabled = False
        On Error GoTo 0
        Cancel = True
    End If
End Sub
