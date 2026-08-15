Attribute VB_Name = "Custom_Menu13_Cholesky"
Option Explicit

' Personal Custom_Menu13_Cholesky.

''' @Description: Lower-triangular Cholesky factor L of a symmetric positive-definite selection (A = L L^T).
Public Sub MatrixCholesky()
    Call Custom_Menu13_MatrixUtilities.UnaryOp("chol")
End Sub

' --- Personal123 Custom_Menu13_Cholesky ---

Sub CholeskyDecompositionOutput()
'performs Cholesky decomposition on a symetric matrix
Dim arr_Input, arr_Cholesky, arr_CholeskyTranspose, arr_CheckMatrix As Variant

'---Capture input range-----------------------------------------------------------
Dim rng_WorkingRange, rng_DefaultRange As Range

Set rng_DefaultRange = Selection

Set rng_WorkingRange = Application.InputBox( _
        Prompt:="Please Select Range", _
        Title:="Range Select", _
        Default:=rng_DefaultRange.Address, _
        Type:=8)
'---Gets the number of rows of the input range---------------------------------------
Dim lng_MatrixCols As Long

lng_MatrixCols = rng_WorkingRange.Columns.count
'---Reads a range into an array------------------------------------------------------
arr_Input = rng_WorkingRange.value
'---checks matrix is symmetric------------------------------------------------------
If IsMatrixSymmetric(arr_Input) = False Then
    MsgBox "Matris isn't symmetric. This won't work." & vbCrLf & vbCrLf & "Ending procedure.", vbOKOnly, "Error!"
    Exit Sub
End If
'---creates output sheet------------------------------------------------------------
Dim str_OutputSheetName As String: str_OutputSheetName = "Cholesky Decomposition"

CheckExistenceAndDeleteOutputSheet (str_OutputSheetName)
CreateOutputSheet (str_OutputSheetName)
'---Calculates the Cholesky lower triangular matrix---------------------------------
arr_Cholesky = CholeskyDecomposition(arr_Input)
'---Calculates the Cholesky lower triangular matrix tranpose-------------------------
arr_CholeskyTranspose = WorksheetFunction.Transpose(CholeskyDecomposition(arr_Input))
'---Calculates the check matrix------------------------------------------------------
arr_CheckMatrix = WorksheetFunction.MMult(CholeskyDecomposition(arr_Input), _
                        WorksheetFunction.Transpose(CholeskyDecomposition(arr_Input)))
'---Output Cholesky decomposition------------------------------------------------------
Call WriteArrayToWorksheet(arr_Input, str_OutputSheetName, 5, 1)
Call WriteArrayToWorksheet(arr_Cholesky, str_OutputSheetName, 5, 1 * lng_MatrixCols + 2)
Call WriteArrayToWorksheet(arr_CholeskyTranspose, str_OutputSheetName, 5, 2 * lng_MatrixCols + 3)
Call WriteArrayToWorksheet(arr_CheckMatrix, str_OutputSheetName, 5, 3 * lng_MatrixCols + 4)

With Sheets(str_OutputSheetName)
    With Range("A1")
        .value = "Cholesky decomposition of a symmetric matrix"
        .Font.Bold = True
    End With
    With Range("A4")
        .value = "Original matrix"
        .Font.Bold = True
    End With
    With Cells(4, 1 * lng_MatrixCols + 2)
        .value = "Lower Triangular matrix"
        .Font.Bold = True
    End With
    With Cells(4, 2 * lng_MatrixCols + 3)
        .value = "Transpose"
        .Font.Bold = True
    End With
    With Cells(4, 3 * lng_MatrixCols + 4)
        .value = "Check matrix"
        .Font.Bold = True
    End With
End With
End Sub

Function CholeskyDecomposition(var_Matrix As Variant) As Variant
Dim arr_LowerTriangular() As Double, S As Double
Dim int_NumRows, int_NumCols, j, k, i As Integer
'Defines the size whether the input is a range or an array
If TypeName(var_Matrix) = "Range" Then
    int_NumRows = var_Matrix.Rows.count
    int_NumCols = var_Matrix.Columns.count
    Else
        int_NumRows = UBound(var_Matrix, 1)
        int_NumCols = UBound(var_Matrix, 2)
End If
'checks whether the input is square or not. Ends the procedure if it isn't square
If int_NumRows <> int_NumCols Then
    CholeskyDecomposition = "Matrix must be square"
    Exit Function
End If

ReDim arr_LowerTriangular(1 To int_NumRows, 1 To int_NumRows)

For j = 1 To int_NumRows
    S = 0
    For k = 1 To j - 1
        S = S + arr_LowerTriangular(j, k) ^ 2
    Next k
    arr_LowerTriangular(j, j) = var_Matrix(j, j) - S
    
    If arr_LowerTriangular(j, j) <= 0 Then
        Exit For
    End If
    
    arr_LowerTriangular(j, j) = Sqr(arr_LowerTriangular(j, j))
    
    For i = j + 1 To int_NumRows
        S = 0
        For k = 1 To j - 1
            S = S + arr_LowerTriangular(i, k) * arr_LowerTriangular(j, k)
        Next k
        arr_LowerTriangular(i, j) = (var_Matrix(i, j) - S) / arr_LowerTriangular(j, j)
    Next i
Next j

CholeskyDecomposition = arr_LowerTriangular

End Function
