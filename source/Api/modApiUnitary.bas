Attribute VB_Name = "modApiUnitary"
Option Explicit

' Real QR / orthogonality (complex unitary not included).

''' @Description: Thin QR (modified Gram-Schmidt). Writes Q, then R immediately below Q.
Public Sub MatrixQR()
    Call modApiMatrixUtilities.UnaryOp("qr")
End Sub

''' @Description: TRUE if A^T A is the identity (orthogonal / real unitary).
Public Sub MatrixIsOrthogonal()
    Call modApiMatrixUtilities.WriteValidationBool("orth")
End Sub

' --- Personal123 Unitary ---

Sub CheckArrayIsUnitary()
'matrices are unitary if multiplying by their respective conjugate transposes yields identity matrix
'this subroutine returns TRUE if the matrix is unitary and FALSE if not
Dim rng_Default, rng_OriginalMtrx As Range
Dim rng_Output As Range
Dim rng_DefaultOutput As Range
Dim lng_OutputRefRow, lng_OutputRefCol As Long
Dim arr_Original(), arr_Transpose(), arr_Multiplication() As Variant
Dim boo_MultiplicationDefined As Boolean: boo_MultiplicationDefined = False
Dim boo_Identity As Boolean: boo_Identity = False

Set rng_Default = Selection
'-----------------------------------------------------------------
'Enable error handling
On Error Resume Next
Set rng_OriginalMtrx = Application.InputBox( _
        Prompt:="Select matrix", _
        Title:="Select range", _
        Default:=rng_Default.Address, _
        Type:=8)
'---Validate range----------------------------------------------------
If IsRangeValidated(rng_OriginalMtrx) <> True Then
    MsgBox IsRangeValidated(rng_OriginalMtrx), vbOKOnly, "Error!"
    Exit Sub
End If
'Disable error-handling
On Error GoTo 0
'---size the original array and write the range to an array---------
ReDim arr_Original(1 To rng_OriginalMtrx.Rows.count, 1 To rng_OriginalMtrx.Columns.count)
arr_Original = rng_OriginalMtrx
'---get the transpose matrix----------------------------------------
arr_Transpose = TransposeMatrix(arr_Original)
'---check the matrix multiplication is defined----------------------
boo_MultiplicationDefined = MatrixMultDefined(arr_Original, arr_Transpose)

If boo_MultiplicationDefined = False Then
    MsgBox "Matrix multiplication is undefined. The columns and rows don't equal. Programme will end", vbOKOnly, "Error"
    Exit Sub
End If
'---multiply the matrix by its transpose----------------------------
arr_Multiplication = MultiplyArrays(arr_Original, arr_Transpose)
'---check the multiplication returns an identity matrix----------------
boo_Identity = IsMatrixIdentity(arr_Multiplication)
'---Report output to the user---------------------------------------
If boo_Identity = True Then
    MsgBox "The matrix is UNITARY", vbOKOnly, "Result"
    Else
        MsgBox "The matrix is NOT unitary", vbOKOnly, "Result"
End If
'---set default output range-------------------------------------
lng_OutputRefRow = rng_OriginalMtrx.row
lng_OutputRefCol = rng_OriginalMtrx.Column + rng_OriginalMtrx.Columns.count + 1
Set rng_DefaultOutput = Cells(lng_OutputRefRow, lng_OutputRefCol)
'---capture output range-------------------------------------------
'Enable error handling
On Error Resume Next
Set rng_Output = Application.InputBox( _
        Prompt:="Select output location", _
        Title:="Select range", _
        Default:=rng_DefaultOutput.Address, _
        Type:=8)

'Disable error-handling
On Error GoTo 0
'---output results------------------------------------------------------
rng_Output.Resize(UBound(arr_Multiplication, 1), UBound(arr_Multiplication, 2)) = arr_Multiplication

End Sub

Private Function TransposeMatrix(arr_Input As Variant) As Variant
'returns the transpose of an array
Dim arr_Transpose() As Variant
Dim lng_OutputRefRow, lng_OutputRefCol As Long
'---set the transpose array dimensions--------------------------------------------
'transpose rows = input array columns
'transpose columns = input array rows
'ReDim arr_Transpose(1 To UBound(arr_Input, 2), 1 To UBound(arr_Input, 1))
'---transpose the array and write it to a new array-----------------------------
arr_Transpose = Application.Transpose(arr_Input)
'---write the transposed array to the fuction output-------------------------------
TransposeMatrix = arr_Transpose
End Function

Private Function MatrixMultDefined(arr_One As Variant, arr_Two As Variant) As Boolean
'Checks whether the multiplication of two matrices is defined by checking their sizes. More specifically
'whether the columns of Matrix1= rowsof Matrix2

If UBound(arr_Two, 1) = UBound(arr_One, 2) Then
    MatrixMultDefined = True
        Else
            MatrixMultDefined = False
End If

End Function

Private Function MultiplyArrays(arr_One As Variant, arr_Two As Variant) As Variant
'multiplies two arrays the order is important

MultiplyArrays = WorksheetFunction.MMult(arr_One, arr_Two)

End Function
