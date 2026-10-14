Attribute VB_Name = "modApiArrays"
Option Explicit

' Public API: array dump/load entry points.
' Other workbooks / Personal shims should call these names only.

Public Sub WriteArrayToSheet(ByVal TargetSheet As String, ByVal StartCell As String, ByRef Data As Variant)
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    Call modInternalSheetIO.DumpArray(TargetSheet, StartCell, Data)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("WriteArrayToSheet")
End Sub

Public Function ReadRangeToArray(ByVal TargetSheet As String, ByVal AddressOrName As String) As Variant
    On Error GoTo EH
    ReadRangeToArray = modInternalSheetIO.ReadArray(TargetSheet, AddressOrName)
    Exit Function
EH:
    Call modInternalError.RaiseCurrent("ReadRangeToArray")
End Function

Public Sub WriteArrayToNewSheet(ByVal SheetName As String, ByVal StartCell As String, ByRef Data As Variant)
    On Error GoTo EH
    Call modApiSheets.EnsureSheet(SheetName)
    Call WriteArrayToSheet(SheetName, StartCell, Data)
    Exit Sub
EH:
    Call modInternalError.RaiseCurrent("WriteArrayToNewSheet")
End Sub

' Personal Menu13 names used by modApiMatrices1 / Cholesky / eigen / create.

Public Sub WriteArrayToWorksheet(arr_WorkingArray As Variant, str_OutputSheetName As String, _
                                 ByVal lng_StartRow As Long, ByVal lng_StartColumn As Long)
    Dim lng_Rows As Long
    Dim lng_Cols As Long
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Sheets(str_OutputSheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Target sheet '" & str_OutputSheetName & "' not found.", vbCritical, "Output Error"
        Exit Sub
    End If
    If lng_StartRow < 1 Then lng_StartRow = 1
    If lng_StartColumn < 1 Then lng_StartColumn = 1
    Select Case ArrayDimensionCount(arr_WorkingArray)
        Case 1
            lng_Rows = UBound(arr_WorkingArray) - LBound(arr_WorkingArray) + 1
            ws.Cells(lng_StartRow, lng_StartColumn).Resize(lng_Rows, 1).Value = Application.Transpose(arr_WorkingArray)
        Case 2
            lng_Rows = UBound(arr_WorkingArray, 1) - LBound(arr_WorkingArray, 1) + 1
            lng_Cols = UBound(arr_WorkingArray, 2) - LBound(arr_WorkingArray, 2) + 1
            ws.Cells(lng_StartRow, lng_StartColumn).Resize(lng_Rows, lng_Cols).Value = arr_WorkingArray
        Case Else
            MsgBox "The array has more than 2 dimensions. This utility only handles 1D/2D.", vbExclamation
    End Select
End Sub

Public Sub WriteArrayToWorksheetA1(arr_WorkingArray As Variant, _
                                   str_OutputSheetName As String, _
                                   str_TopLeftOutputRange As String)
    Dim lng_StartRow As Long
    Dim lng_StartColumn As Long
    Dim rng_Test As Range
    On Error Resume Next
    Set rng_Test = ActiveWorkbook.Sheets(str_OutputSheetName).Range(str_TopLeftOutputRange)
    If rng_Test Is Nothing Then Set rng_Test = Range(str_TopLeftOutputRange)
    On Error GoTo 0
    If rng_Test Is Nothing Then
        lng_StartRow = 1
        lng_StartColumn = 1
    Else
        lng_StartRow = rng_Test.Row
        lng_StartColumn = rng_Test.Column
    End If
    Call WriteArrayToWorksheet(arr_WorkingArray, str_OutputSheetName, lng_StartRow, lng_StartColumn)
End Sub

Public Function IsRangeValidated(rng_input As Range) As Variant
    If rng_input Is Nothing Then
        IsRangeValidated = "Range is empty"
        Exit Function
    End If
    If Application.CountA(rng_input) = 0 Then
        IsRangeValidated = "Range is completely empty"
        Exit Function
    End If
    IsRangeValidated = True
End Function

Public Function IsRangeNumeric(rng_input As Range) As Boolean
    Dim rng_C As Range
    Dim lng_ErrorCounter As Long
    If rng_input Is Nothing Then
        IsRangeNumeric = False
        Exit Function
    End If
    For Each rng_C In rng_input
        If Not IsNumeric(rng_C.Value) Then lng_ErrorCounter = lng_ErrorCounter + 1
    Next rng_C
    IsRangeNumeric = (lng_ErrorCounter = 0)
End Function

Private Function ArrayDimensionCount(arr_WorkingArray As Variant) As Long
    Dim i As Long
    Dim tmp As Long
    If Not IsArray(arr_WorkingArray) Then
        ArrayDimensionCount = 0
        Exit Function
    End If
    On Error GoTo Done
    i = 0
    Do While True
        i = i + 1
        tmp = UBound(arr_WorkingArray, i)
    Loop
Done:
    ArrayDimensionCount = i - 1
    Err.Clear
End Function
