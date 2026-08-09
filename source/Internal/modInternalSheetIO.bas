Attribute VB_Name = "modInternalSheetIO"
Option Explicit

' Internal: array <-> range I/O shared by Api modules.
' Paste implementations from Personal.xlsb helpers that "output arrays to worksheets".

Public Sub DumpArray(ByVal TargetSheet As String, ByVal StartCell As String, ByRef Data As Variant)
    Dim ws As Worksheet
    Dim dest As Range
    Dim rows As Long, cols As Long

    ' ActiveWorkbook = caller's book; ThisWorkbook would be the add-in itself.
    Set ws = ActiveWorkbook.Worksheets(TargetSheet)
    Set dest = ws.Range(StartCell)

    If IsEmpty(Data) Then Exit Sub

    If IsArray(Data) Then
        If RankOf(Data) = 1 Then
            rows = UBound(Data) - LBound(Data) + 1
            cols = 1
            dest.Resize(rows, cols).Value = Application.Transpose(Data)
        Else
            rows = UBound(Data, 1) - LBound(Data, 1) + 1
            cols = UBound(Data, 2) - LBound(Data, 2) + 1
            dest.Resize(rows, cols).Value = Data
        End If
    Else
        dest.Value = Data
    End If
End Sub

Public Function ReadArray(ByVal TargetSheet As String, ByVal AddressOrName As String) As Variant
    ReadArray = ActiveWorkbook.Worksheets(TargetSheet).Range(AddressOrName).Value
End Function

Private Function RankOf(ByRef Data As Variant) As Long
    On Error GoTo fail
    Dim n As Long
    n = UBound(Data, 2)
    RankOf = 2
    Exit Function
fail:
    RankOf = 1
End Function
