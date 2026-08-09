Attribute VB_Name = "modInternalExcelApp"
Option Explicit

' Internal: push/pop Application state (ScreenUpdating, Calculation, EnableEvents).
' Used by Api routines that write heavily to sheets.

Private Type AppState
    ScreenUpdating As Boolean
    EnableEvents As Boolean
    Calculation As XlCalculation
    HasValue As Boolean
End Type

Private mState As AppState

Public Sub PushAppState()
    If Not mState.HasValue Then
        mState.ScreenUpdating = Application.ScreenUpdating
        mState.EnableEvents = Application.EnableEvents
        mState.Calculation = Application.Calculation
        mState.HasValue = True
    End If
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End Sub

Public Sub PopAppState()
    If Not mState.HasValue Then Exit Sub
    Application.ScreenUpdating = mState.ScreenUpdating
    Application.EnableEvents = mState.EnableEvents
    Application.Calculation = mState.Calculation
    mState.HasValue = False
End Sub
