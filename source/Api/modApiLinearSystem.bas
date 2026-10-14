Attribute VB_Name = "modApiLinearSystem"
Option Explicit

' Public API: Personal Menu18 AX = B solver.
' LinearSystem_AXB_v2 was a RefEdit UserForm that duplicated Gauss-Jordan (and a broken integer path).
' Both names now prompt for A, B, and an output cell and use the add-in matrix solver.

''' @Description: Solve AX = B. A must be square and larger than 1 x 1; B must have the same number of rows.
''' @Example: LinearSystem_AXB_v1
Public Sub LinearSystem_AXB_v1()
    Call SolveAXB
End Sub

''' @Description: Same as LinearSystem_AXB_v1 (Personal menu OnAction). The UserForm/RefEdit path is not used.
''' @Example: LinearSystem_AXB_v2
Public Sub LinearSystem_AXB_v2()
    Call SolveAXB
End Sub

Private Sub SolveAXB()
    Dim rngA As Range
    Dim rngB As Range
    Dim dest As Range
    Dim a As Variant
    Dim b As Variant
    Dim x As Variant
    On Error GoTo EH
    If TypeName(Selection) <> "Range" Then
        MsgBox "Select a range first.", vbExclamation, "Linear system"
        Exit Sub
    End If
    On Error Resume Next
    Set rngA = Application.InputBox(Prompt:="Select square matrix A (more than one cell).", _
                                    Title:="Linear system AX=B", Default:=Selection.Address, Type:=8)
    On Error GoTo 0
    If rngA Is Nothing Then Exit Sub
    If rngA.Rows.Count <> rngA.Columns.Count Then
        MsgBox "A must be square.", vbExclamation, "Linear system"
        Exit Sub
    End If
    If rngA.Rows.Count = 1 Then
        MsgBox "A cannot be 1 x 1.", vbExclamation, "Linear system"
        Exit Sub
    End If
    If Application.WorksheetFunction.Count(rngA) <> rngA.Cells.Count Then
        MsgBox "A must be numeric.", vbExclamation, "Linear system"
        Exit Sub
    End If
    On Error Resume Next
    Set rngB = Application.InputBox(Prompt:="Select B (same number of rows as A).", _
                                    Title:="Linear system AX=B", Default:=Selection.Address, Type:=8)
    On Error GoTo 0
    If rngB Is Nothing Then Exit Sub
    If rngB.Rows.Count <> rngA.Rows.Count Then
        MsgBox "A and B must have the same number of rows.", vbExclamation, "Linear system"
        Exit Sub
    End If
    If Application.WorksheetFunction.Count(rngB) <> rngB.Cells.Count Then
        MsgBox "B must be numeric.", vbExclamation, "Linear system"
        Exit Sub
    End If
    On Error Resume Next
    Set dest = Application.InputBox(Prompt:="Select output start cell", Title:="Linear system AX=B", _
                                    Default:=rngB.Cells(1).Offset(0, rngB.Columns.Count + 1).Address, Type:=8)
    On Error GoTo 0
    If dest Is Nothing Then Exit Sub
    a = modInternalMatrices.RangeToMatrix(rngA)
    b = modInternalMatrices.RangeToMatrix(rngB)
    x = modInternalMatrices.Solve(a, b)
    Call modInternalMatrices.PutMatrix(dest.Cells(1, 1), x)
    dest.Cells(1, 1).Resize(modInternalMatrices.RowsOf(x), modInternalMatrices.ColsOf(x)).Select
    Exit Sub
EH:
    Call modInternalError.RaiseCurrent("LinearSystem_AXB")
End Sub
