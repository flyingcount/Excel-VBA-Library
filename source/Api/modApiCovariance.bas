Attribute VB_Name = "Custom_Menu18_Covariance"
Option Explicit

Sub MatrixCovariance()
Dim rng_input, rng_Output, rng_InputDefault, rng_OutputDefault As Range
Dim arr_Mtrx, arr_Output As Variant
'---Capture input range-----------------------------------------------------------
Set rng_InputDefault = Selection
'Enable error handling
On Error Resume Next

Set rng_input = Application.InputBox( _
        Prompt:="Please Select Range", _
        Title:="Range Select", _
        Default:=rng_InputDefault.Address, _
        Type:=8)
'---validate input range--------------------------------------------------------
If rng_input Is Nothing Then
    MsgBox "You must select a range. Ending program", vbOKOnly, "Error:"
    Exit Sub
End If
If rng_input.Rows.count <> rng_input.Columns.count Then
    MsgBox "The range must be square. Ending program", vbOKOnly, "Error:"
    Exit Sub
End If
If rng_input.Cells.count = 1 Then
    MsgBox "The range must be more than 1 cell. Ending program", vbOKOnly, "Error:"
    Exit Sub
End If
'Disable error-handling
On Error GoTo 0
'---write range to matrix---------------------------------------------------------
ReDim arr_Mtrx(1 To rng_input.Rows.count, 1 To rng_input.Columns.count)
arr_Mtrx = rng_input

arr_Output = CalculateMatrixCovariance(arr_Mtrx)
'---Capture output range-----------------------------------------------------------
Set rng_OutputDefault = Cells(rng_input.row, rng_input.Column + rng_input.Columns.count + 1)
'Enable error handling
On Error Resume Next

Set rng_Output = Application.InputBox( _
        Prompt:="Please Select Range", _
        Title:="Range Select", _
        Default:=rng_OutputDefault.Address, _
        Type:=8)
'---validate output range--------------------------------------------------------
If rng_Output Is Nothing Then
    Exit Sub
End If
'---output the matrix-------------------------------------------------------------
ActiveSheet.Range(rng_Output.Address).Resize(UBound(arr_Output, 1), UBound(arr_Output, 2)) = arr_Output
End Sub

Sub MatrixCovarianceStandardise()
Dim rng_input, rng_Output, rng_InputDefault, rng_OutputDefault As Range
Dim arr_Mtrx, arr_Output As Variant
'---Capture input range-----------------------------------------------------------
Set rng_InputDefault = Selection
'Enable error handling
On Error Resume Next

Set rng_input = Application.InputBox( _
        Prompt:="Please Select Range", _
        Title:="Range Select", _
        Default:=rng_InputDefault.Address, _
        Type:=8)
'---validate input range--------------------------------------------------------
If rng_input Is Nothing Then
    MsgBox "You must select a range. Ending program", vbOKOnly, "Error:"
    Exit Sub
End If
If rng_input.Rows.count <> rng_input.Columns.count Then
    MsgBox "The range must be square. Ending program", vbOKOnly, "Error:"
    Exit Sub
End If
If rng_input.Cells.count = 1 Then
    MsgBox "The range must be more than 1 cell. Ending program", vbOKOnly, "Error:"
    Exit Sub
End If
'Disable error-handling
On Error GoTo 0
'---write range to matrix---------------------------------------------------------
ReDim arr_Mtrx(1 To rng_input.Rows.count, 1 To rng_input.Columns.count)
arr_Mtrx = rng_input

arr_Output = CalculateMatrixCovarianceStandardise(arr_Mtrx)
'---Capture output range-----------------------------------------------------------
Set rng_OutputDefault = Cells(rng_input.row, rng_input.Column + rng_input.Columns.count + 1)
'Enable error handling
On Error Resume Next

Set rng_Output = Application.InputBox( _
        Prompt:="Please Select Range", _
        Title:="Range Select", _
        Default:=rng_OutputDefault.Address, _
        Type:=8)
'---validate output range--------------------------------------------------------
If rng_Output Is Nothing Then
    Exit Sub
End If
'---output the matrix-------------------------------------------------------------
ActiveSheet.Range(rng_Output.Address).Resize(UBound(arr_Output, 1), UBound(arr_Output, 2)) = arr_Output
End Sub

Private Function CalculateMatrixCovariance(arr_Mtrx As Variant) As Variant
Dim arr_Statistics, arr_SubtractMean, arr_Output As Variant
'calculate the means and Std Dev of each column in the array and write to a 2-row matrix
arr_Statistics = GetMatrixStatistics(arr_Mtrx)
'deduct column mean from each column
arr_SubtractMean = MatrixSubtractColumnMean(arr_Mtrx, arr_Statistics)
'calculate A(transpose).A
arr_Output = GramianMatrix(arr_SubtractMean)
'divide each element of the matrix by n-1
arr_Output = MatrixDivideByScalar(arr_Output, UBound(arr_Mtrx, 1) - 1)
CalculateMatrixCovariance = arr_Output
End Function

Private Function CalculateMatrixCovarianceStandardise(arr_Mtrx As Variant) As Variant
Dim rng_input As Range
Dim arr_Statistics, arr_SubtractMeanDivideSD As Variant
Dim arr_Output As Variant
Dim dbl_mean As Double
'calculate the means and Std Dev of each column in the array and write to a 2-row matrix
arr_Statistics = GetMatrixStatistics(arr_Mtrx)
'deduct column mean from each column
arr_SubtractMeanDivideSD = MatrixSubtractColumnMean(arr_Mtrx, arr_Statistics)
'divide by column std dev from each column
arr_SubtractMeanDivideSD = MatrixDivideColumnStdDev(arr_SubtractMeanDivideSD, arr_Statistics)
'calculate A(transpose).A
arr_Output = GramianMatrix(arr_SubtractMeanDivideSD)
'divide each element of the matrix by n-1
arr_Output = MatrixDivideByScalar(arr_Output, UBound(arr_Mtrx, 1) - 1)
'output the matrix
CalculateMatrixCovarianceStandardise = arr_Output
End Function

Private Function MatrixDivide(arr_Data As Variant, arr_Divisor As Variant) As Variant
Dim lng_RowCounter, lng_ColCounter, lng_CountLessOne As Long
Dim arr_Temp As Variant

ReDim arr_Temp(1 To UBound(arr_Data, 1), 1 To UBound(arr_Data, 2))

If UBound(arr_Data, 2) <> UBound(arr_Divisor, 2) Then
    MsgBox "Both matrices must have the same number of columns", vbOKOnly, "Error"
    Exit Function
End If

For lng_ColCounter = 1 To UBound(arr_Data, 2) Step 1
    For lng_RowCounter = 1 To UBound(arr_Data, 1) Step 1
        arr_Temp(lng_RowCounter, lng_ColCounter) = arr_Data(lng_RowCounter, lng_ColCounter) / arr_Divisor(1, lng_ColCounter)
    Next
Next
MatrixDivide = arr_Temp
End Function

Private Function MatrixDivideByScalar(arr_Data As Variant, lng_Divisor) As Variant
Dim lng_RowCounter, lng_ColCounter, lng_CountLessOne As Long
Dim arr_Temp As Variant

ReDim arr_Temp(1 To UBound(arr_Data, 1), 1 To UBound(arr_Data, 2))

For lng_ColCounter = 1 To UBound(arr_Data, 2) Step 1
    For lng_RowCounter = 1 To UBound(arr_Data, 1) Step 1
        arr_Temp(lng_RowCounter, lng_ColCounter) = arr_Data(lng_RowCounter, lng_ColCounter) / lng_Divisor
    Next
Next
MatrixDivideByScalar = arr_Temp
End Function

Private Function MatrixSubtractColumnMean(arr_Data As Variant, arr_Statistics As Variant) As Variant
Dim lng_RowCounter, lng_ColCounter As Long
Dim arr_Temp As Variant
Dim dbl_Totaliser As Double

ReDim arr_Temp(1 To UBound(arr_Data, 1), 1 To UBound(arr_Data, 2))

For lng_ColCounter = 1 To UBound(arr_Data, 2) Step 1
    For lng_RowCounter = 1 To UBound(arr_Data, 1) Step 1
        arr_Temp(lng_RowCounter, lng_ColCounter) = arr_Data(lng_RowCounter, lng_ColCounter) _
                                                    - arr_Statistics(1, lng_ColCounter)
    Next
Next
MatrixSubtractColumnMean = arr_Temp
End Function

Private Function GetMatrixStatistics(arr_Mtrx As Variant) As Variant
'calculate the means of each column in a matrix and return it as a row matrix
Dim lng_RowCounter, lng_ColCounter As Long
Dim dbl_Totaliser As Double
Dim arr_Statistics As Variant

ReDim arr_Statistics(1 To 5, 1 To UBound(arr_Mtrx, 2))
'populate counts to row 5  in the array
For lng_ColCounter = 1 To UBound(arr_Mtrx, 2) Step 1
    arr_Statistics(5, lng_ColCounter) = UBound(arr_Mtrx, 1)
Next
'populate the mean in row = 1
For lng_ColCounter = 1 To UBound(arr_Mtrx, 2) Step 1
    dbl_Totaliser = 0
    For lng_RowCounter = 1 To UBound(arr_Mtrx, 1) Step 1
        dbl_Totaliser = dbl_Totaliser + arr_Mtrx(lng_RowCounter, lng_ColCounter)
    Next
    'write the mean to row 1 in the array
    arr_Statistics(1, lng_ColCounter) = dbl_Totaliser / arr_Statistics(5, lng_ColCounter)
Next
'populate the std dev in row = 2
For lng_ColCounter = 1 To UBound(arr_Mtrx, 2) Step 1
    dbl_Totaliser = 0
    For lng_RowCounter = 1 To UBound(arr_Mtrx, 1) Step 1
        dbl_Totaliser = dbl_Totaliser + ((arr_Mtrx(lng_RowCounter, lng_ColCounter) - arr_Statistics(1, lng_ColCounter)) ^ 2)
    Next
    'write the std dev to row 2 in the array
    arr_Statistics(2, lng_ColCounter) = Sqr(dbl_Totaliser / (arr_Statistics(5, lng_ColCounter) - 1))
Next
GetMatrixStatistics = arr_Statistics
End Function

Private Function MatrixDivideColumnStdDev(arr_Mtrx As Variant, arr_Statistics As Variant) As Variant
'calculate the means of each column in a matrix and return it as a row matrix
Dim lng_RowCounter, lng_ColCounter As Long
Dim arr_Temp As Variant

ReDim arr_Temp(1 To UBound(arr_Mtrx, 1), 1 To UBound(arr_Mtrx, 2))

'populate the std dev array
For lng_ColCounter = 1 To UBound(arr_Mtrx, 2) Step 1
    For lng_RowCounter = 1 To UBound(arr_Mtrx, 1) Step 1
        arr_Temp(lng_RowCounter, lng_ColCounter) = arr_Mtrx(lng_RowCounter, lng_ColCounter) / arr_Statistics(2, lng_ColCounter)
    Next
Next
MatrixDivideColumnStdDev = arr_Temp
End Function

Private Function GramianMatrix(arr_Matrix As Variant) As Variant
'A(Transpose).A = Gramian matrix
Dim arr_MatrixTranspose As Variant
'define the size of the array
ReDim arr_MatrixTranspose(1 To UBound(arr_Matrix, 2), 1 To UBound(arr_Matrix, 1))

'transpose the matrix
arr_MatrixTranspose = WorksheetFunction.Transpose(arr_Matrix)
'multiply the transposed matrix by itself
GramianMatrix = WorksheetFunction.MMult(arr_MatrixTranspose, arr_Matrix)
End Function
