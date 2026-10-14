Attribute VB_Name = "modApiCovariance"
Option Explicit

' Public API: Personal Custom_Menu18_Covariance.
' Sample covariance (n-1). Observations in rows, variables in columns — the range does not need to be square.

''' @Description: Write the sample covariance matrix (n-1) to the right of a numeric block of observations x variables.
''' @Example: MatrixCovariance
Public Sub MatrixCovariance()
    Call WriteCov("sample")
End Sub

''' @Description: Write the sample covariance of standardised columns (sample correlation) to the right of the data.
''' @Example: MatrixCovarianceStandardise
Public Sub MatrixCovarianceStandardise()
    Call WriteCov("standard")
End Sub

Private Sub WriteCov(ByVal kind As String)
    Dim src As Range
    Dim dest As Range
    Dim body As Variant
    Dim out As Variant
    On Error GoTo EH
    Set src = modInternalAnalysis.PromptRange("Select the data range (rows = observations, columns = variables)")
    If src Is Nothing Then Exit Sub
    If src.Rows.Count < 2 Then
        MsgBox "Need at least two observation rows.", vbExclamation, "Covariance"
        Exit Sub
    End If
    If Not modInternalAnalysis.RangeIsAllNumeric(src) Then
        MsgBox "The range must be numeric.", vbExclamation, "Covariance"
        Exit Sub
    End If
    body = modInternalAnalysis.AsMatrix(src.Value, src.Rows.Count, src.Columns.Count)
    If kind = "standard" Then
        out = modInternalAnalysis.CovarianceStandardised(body)
    Else
        out = modInternalAnalysis.CovarianceSample(body)
    End If
    Set dest = modInternalAnalysis.PromptRange("Select output start cell", src.Cells(1).Offset(0, src.Columns.Count + 1))
    If dest Is Nothing Then Exit Sub
    If kind = "standard" Then
        Call modInternalAnalysis.WriteTitle(dest.Cells(1, 1), "Standardised Covariance")
        Call modInternalAnalysis.PutArray(dest.Cells(2, 1), out)
    Else
        Call modInternalAnalysis.PutArray(dest.Cells(1, 1), out)
    End If
    Exit Sub
EH:
    Call modInternalError.RaiseCurrent("MatrixCovariance")
End Sub
