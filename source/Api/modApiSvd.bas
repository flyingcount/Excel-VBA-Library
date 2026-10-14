Attribute VB_Name = "modApiSvd"
Option Explicit

' Public API: Personal Custom_Menu18_SVD. A = U * diag(s) * V^T written beside the input.

''' @Description: Singular value decomposition of a numeric matrix. Writes singular values, U, and V.
''' @Example: SVD
Public Sub SVD()
    Dim src As Range
    Dim dest As Range
    Dim a As Variant
    Dim w As Variant
    Dim v As Variant
    Dim nRows As Long
    Dim nCols As Long
    On Error GoTo EH
    If TypeName(Selection) <> "Range" Then
        MsgBox "Select a numeric matrix.", vbExclamation, "SVD"
        Exit Sub
    End If
    On Error Resume Next
    Set src = Application.InputBox(Prompt:="Select input matrix", Title:="SVD", Default:=Selection.Address, Type:=8)
    On Error GoTo 0
    If src Is Nothing Then Exit Sub
    If Application.WorksheetFunction.Count(src) <> src.Cells.Count Then
        MsgBox "The matrix must be numeric.", vbExclamation, "SVD"
        Exit Sub
    End If
    nRows = src.Rows.Count
    nCols = src.Columns.Count
    a = modInternalAnalysis.AsMatrix(src.Value, nRows, nCols)
    On Error Resume Next
    Set dest = Application.InputBox(Prompt:="Select output start cell", Title:="SVD", _
                                    Default:=src.Cells(1).Offset(0, nCols + 1).Address, Type:=8)
    On Error GoTo 0
    If dest Is Nothing Then Exit Sub
    Call modInternalSvd.Decompose(a, nRows, nCols, w, v)
    Call modInternalSvd.WriteBlocks(dest.Cells(1, 1), a, w, v, nRows, nCols)
    Exit Sub
EH:
    Call modInternalError.RaiseCurrent("SVD")
End Sub
