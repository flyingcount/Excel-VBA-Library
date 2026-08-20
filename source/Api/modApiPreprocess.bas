Attribute VB_Name = "modApiPreprocess"
Option Explicit

' Public API: Personal Custom_Menu5_Scaling / Custom_menu5_ML.
' Column scaling and dummy (one-hot) encoding. Writes to the right of the source range.

''' @Description: Standardise each column: (x − mean) / population SD. Writes two columns to the right of the source.
''' @Example: ScalingStandard
Public Sub ScalingStandard()
    Call ScaleColumns("standard")
End Sub

''' @Description: Min-max normalise each column to [0, 1]. Writes two columns to the right of the source.
''' @Example: ScalingNormalise
Public Sub ScalingNormalise()
    Call ScaleColumns("normalise")
End Sub

''' @Description: Robust-scale each column: (x − median) / IQR (QUARTILE.INC). Writes two columns to the right of the source.
''' @Example: ScalingRobust
Public Sub ScalingRobust()
    Call ScaleColumns("robust")
End Sub

''' @Description: One-hot dummy columns from a single categorical column. Optional header row. Writes two columns to the right.
''' @Example: DummyVariablesForMachineLearning
Public Sub DummyVariablesForMachineLearning()
    Dim src As Range
    Dim cats As Variant
    Dim dummy As Variant
    Dim dest As Range
    Dim nCat As Long
    Dim nRows As Long
    Dim withHeaders As VbMsgBoxResult
    Dim r As Long
    Dim c As Long
    Dim headerRow As Long
    On Error GoTo EH
    Set src = modInternalPreprocess.PromptRange( _
        "Select a single column of categorical data." & vbCrLf & vbCrLf & _
        "Output is written two columns to the right of the first cell.")
    If src Is Nothing Then Exit Sub
    If src.Areas.Count > 1 Then
        MsgBox "Select a single contiguous column.", vbExclamation, "Data Preprocessing"
        Exit Sub
    End If
    If src.Columns.Count <> 1 Then
        MsgBox "Input range must be a single column.", vbExclamation, "Data Preprocessing"
        Exit Sub
    End If
    cats = modInternalPreprocess.UniqueColumnValues(src)
    If IsEmpty(cats) Then
        MsgBox "No non-blank values in the column.", vbExclamation, "Data Preprocessing"
        Exit Sub
    End If
    dummy = modInternalPreprocess.DummyMatrix(src, cats)
    nCat = UBound(cats, 1)
    nRows = src.Rows.Count
    withHeaders = MsgBox("Include a header row of category names (and the original values in the first output column)?", _
                         vbYesNo + vbQuestion, "Data Preprocessing")
    Call modInternalExcelApp.PushAppState
    Set dest = src.Cells(1, 1).Offset(0, 2)
    If withHeaders = vbYes Then
        If dest.Row > 1 Then
            headerRow = dest.Row - 1
            dest.Worksheet.Cells(headerRow, dest.Column).Value = "Value"
            For c = 1 To nCat
                dest.Worksheet.Cells(headerRow, dest.Column + c).Value = cats(c, 1)
            Next c
            dest.Worksheet.Cells(headerRow, dest.Column).Resize(1, nCat + 1).Font.Bold = True
            For r = 1 To nRows
                dest.Cells(r, 1).Value = src.Cells(r, 1).Value
            Next r
            dest.Offset(0, 1).Resize(nRows, nCat).Value = dummy
        Else
            dest.Value = "Value"
            For c = 1 To nCat
                dest.Offset(0, c).Value = cats(c, 1)
            Next c
            dest.Resize(1, nCat + 1).Font.Bold = True
            For r = 1 To nRows
                dest.Offset(r, 0).Value = src.Cells(r, 1).Value
            Next r
            dest.Offset(1, 1).Resize(nRows, nCat).Value = dummy
        End If
    Else
        dest.Resize(nRows, nCat).Value = dummy
    End If
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("DummyVariablesForMachineLearning")
End Sub

Private Sub ScaleColumns(ByVal kind As String)
    Dim src As Range
    Dim body As Variant
    Dim scaled As Variant
    Dim dest As Range
    Dim title As String
    On Error GoTo EH
    Select Case kind
        Case "standard": title = "Scaled data (standardised)"
        Case "normalise": title = "Scaled data (normalised)"
        Case "robust": title = "Scaled data (robust)"
        Case Else: Err.Raise 5, "ScaleColumns", "Unknown kind"
    End Select
    Set src = modInternalPreprocess.PromptNumericColumns("Select numeric columns to scale (no header row).")
    If src Is Nothing Then Exit Sub
    body = modInternalAnalysis.AsMatrix(src.Value, src.Rows.Count, src.Columns.Count)
    Select Case kind
        Case "standard": scaled = modInternalPreprocess.ScaleStandardise(body)
        Case "normalise": scaled = modInternalPreprocess.ScaleNormalise(body)
        Case "robust": scaled = modInternalPreprocess.ScaleRobust(body)
    End Select
    Call modInternalExcelApp.PushAppState
    Set dest = src.Worksheet.Cells(src.Row, src.Column + src.Columns.Count + 2)
    Call modInternalPreprocess.WriteBlockToRight(dest, scaled, title)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("ScaleColumns")
End Sub
