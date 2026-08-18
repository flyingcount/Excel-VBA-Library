Attribute VB_Name = "modApiEigenDecomp"
Option Explicit

' Symmetric eigen-decomposition (Jacobi). Selection must be symmetric.

''' @Description: Symmetric eigen-decomposition. Writes eigenvectors then an eigenvalues column.
Public Sub MatrixEigen()
    Call modApiMatrixUtilities.UnaryOp("eigen")
End Sub

' --- Personal123 EigenDecomp ---

Sub EigenDecompositionSymmetricMatrix()
'Performs the eigen decomposition of a symmetric matrix into all its component parts
'---Input range----------------------------------------------------------------------
Dim rng_Default, rng_input As Range

Set rng_Default = Selection
'Enable error handling
On Error Resume Next
Set rng_input = Application.InputBox( _
        Prompt:="Select matrix", _
        Title:="Select range", _
        Default:=rng_Default.Address, _
        Type:=8)
'Disable error-handling
On Error GoTo 0
'check that range isn't nothing
If rng_input Is Nothing Then
    MsgBox "Range is empty. Ending procedure.", vbOKOnly, "Error!"
    Exit Sub
End If
'check that range has rows or columns > 1
If rng_input.Rows.count < 2 Or rng_input.Columns.count < 2 Then
    MsgBox "Range must have more than 1 rows and columns. Ending procedure.", vbOKOnly, "Error!"
    Exit Sub
End If
'---assign range to an array----------------------------------------------------------
Dim arr_Input() As Variant
arr_Input = rng_input.value

Dim lng_InputMatrixSize As Long: lng_InputMatrixSize = UBound(arr_Input, 1) - LBound(arr_Input, 1) + 1
'---confirm matrix is symmetric-------------------------------------------------------
If IsMatrixSymmetric(arr_Input) = False Then
    'if not symmetric notify user and end procedure
    MsgBox "Matrix is not symmetric. Ending procedure.", vbOKOnly, "Error!"
    Exit Sub
End If
'---Prepare output sheet--------------------------------------------------------------
Dim str_OutputSheetName As String: str_OutputSheetName = "Eigen Decomposition"

Call CheckExistenceAndDeleteOutputSheet(str_OutputSheetName)
Call CreateOutputSheet(str_OutputSheetName)
'---Generate eigen values and eigen vectors------------------------------------------
Dim arr_Eigen As Variant
ReDim arr_Eigen(1 To lng_InputMatrixSize, 1 To lng_InputMatrixSize + 1)

arr_Eigen = EIGEN_JK(arr_Input)
'---Output eigen values and eigen vectors---------------------------------------------
Dim lng_StartRow As Long: lng_StartRow = 5
Dim lng_StartColumn As Long: lng_StartColumn = 1

With Sheets(str_OutputSheetName)
    With .Range("A1")
    .value = "Eigen Decomposition of Symmetric Matrix into Q A Q(transpose)"
    .Font.Bold = True
    End With
    With .Cells(lng_StartRow, lng_StartColumn)
    .value = "Original Symmetric Matrix"
    .Font.Bold = True
    End With
    With .Cells(lng_StartRow, lng_StartColumn + 1 * lng_InputMatrixSize + 1)
    .value = "Eigen values and eigen vectors"
    .Font.Bold = True
    End With
End With

Call WriteArrayToWorksheet(arr_Input, str_OutputSheetName, lng_StartRow + 1, lng_StartColumn)
Call WriteArrayToWorksheet(arr_Eigen, str_OutputSheetName, lng_StartRow + 1, lng_StartColumn + 1 * lng_InputMatrixSize + 1)
'---Generate QAQTranspose-------------------------------------------------------------
Dim arr_Q, arr_QTRanspose, arr_EigenValueDiagonalMatrix As Variant
Dim lng_CountRow, lng_CountCol As Long

ReDim arr_Q(LBound(arr_Eigen, 1) To UBound(arr_Eigen, 1), LBound(arr_Eigen, 2) To UBound(arr_Eigen, 2) - 1)
ReDim arr_EigenValueDiagonalMatrix(LBound(arr_Eigen, 1) To UBound(arr_Eigen, 1), LBound(arr_Eigen, 2) To UBound(arr_Eigen, 2) - 1)
ReDim arr_QTRanspose(LBound(arr_Eigen, 1) To UBound(arr_Eigen, 1), LBound(arr_Eigen, 2) To UBound(arr_Eigen, 2) - 1)

'Generate Q
For lng_CountRow = LBound(arr_Eigen, 1) To UBound(arr_Eigen, 1) Step 1
    For lng_CountCol = LBound(arr_Eigen, 2) + 1 To UBound(arr_Eigen, 2) Step 1
        arr_Q(lng_CountRow, lng_CountCol - 1) = arr_Eigen(lng_CountRow, lng_CountCol)
    Next lng_CountCol
Next lng_CountRow

'Generate eigen value diagonal matrix
For lng_CountRow = LBound(arr_Eigen, 1) To UBound(arr_Eigen, 1) Step 1
    For lng_CountCol = LBound(arr_Eigen, 2) To UBound(arr_Eigen, 2) - 1 Step 1
        If lng_CountRow = lng_CountCol Then
        arr_EigenValueDiagonalMatrix(lng_CountRow, lng_CountCol) = arr_Eigen(lng_CountRow, 1)
            Else: arr_EigenValueDiagonalMatrix(lng_CountRow, lng_CountCol) = 0
        End If
    Next lng_CountCol
Next lng_CountRow

'Generate Q Transpose
arr_QTRanspose = WorksheetFunction.Transpose(arr_Q)
'---Output Q, A and Q Transpose-------------------------------------------------------
Call WriteArrayToWorksheet(arr_Q, str_OutputSheetName, lng_StartRow + 1 * lng_InputMatrixSize + 3, lng_StartColumn)
Call WriteArrayToWorksheet(arr_EigenValueDiagonalMatrix, str_OutputSheetName, lng_StartRow + 1 * lng_InputMatrixSize + 3, lng_StartColumn + 1 * lng_InputMatrixSize + 1)
Call WriteArrayToWorksheet(arr_QTRanspose, str_OutputSheetName, lng_StartRow + 1 * lng_InputMatrixSize + 3, lng_StartColumn + 2 * lng_InputMatrixSize + 2)
'---Calculate QAQTranspose------------------------------------------------------------
Dim arr_QAQTranspose As Variant

arr_QAQTranspose = WorksheetFunction.MMult(arr_Q, arr_EigenValueDiagonalMatrix)
arr_QAQTranspose = WorksheetFunction.MMult(arr_QAQTranspose, arr_QTRanspose)
'---Output QAQTranspose------------------------------------------------------------
Call WriteArrayToWorksheet(arr_QAQTranspose, str_OutputSheetName, lng_StartRow + 1 * lng_InputMatrixSize + 3, lng_StartColumn + 3 * lng_InputMatrixSize + 3)
'---Calculate Diagonal Matrix Lambda (Eigenvalues) = QTranspose A Q-------------------
Dim arr_DiagonalMatrixLambda As Variant

arr_DiagonalMatrixLambda = WorksheetFunction.MMult(arr_QTRanspose, arr_Input)
arr_DiagonalMatrixLambda = WorksheetFunction.MMult(arr_DiagonalMatrixLambda, arr_Q)
'---Output QAQTranspose------------------------------------------------------------
Call WriteArrayToWorksheet(arr_DiagonalMatrixLambda, str_OutputSheetName, lng_StartRow + 1 * lng_InputMatrixSize + 3, lng_StartColumn + 4 * lng_InputMatrixSize + 4)

With Sheets(str_OutputSheetName)
    With .Cells(lng_StartRow + 1 * lng_InputMatrixSize + 2, lng_StartColumn)
    .value = "Matrix Q"
    .Font.Bold = True
    End With
    With .Cells(lng_StartRow + 1 * lng_InputMatrixSize + 2, lng_StartColumn + 1 * lng_InputMatrixSize + 1)
    .value = "Matrix Lambda, A"
    .Font.Bold = True
    End With
    With .Cells(lng_StartRow + 1 * lng_InputMatrixSize + 2, lng_StartColumn + 2 * lng_InputMatrixSize + 2)
    .value = "Matrix Q Transpose"
    .Font.Bold = True
    End With
    With .Cells(lng_StartRow + 1 * lng_InputMatrixSize + 2, lng_StartColumn + 3 * lng_InputMatrixSize + 3)
    .value = "Q A Q Transpose"
    .Font.Bold = True
    End With
    With .Cells(lng_StartRow + 1 * lng_InputMatrixSize + 2, lng_StartColumn + 4 * lng_InputMatrixSize + 4)
    .value = "Matrix Lambda, A = Q-1 Input matrix Q"
    .Font.Bold = True
    End With
End With
'---Confirm QAQTranspose is identical to input matrix---------------------------------
End Sub

Sub DiagonalEigenvalueSymmetricMatrix()
'Returns the eigenvalue diagonal matrix of a symmetric matrix
'---Input range----------------------------------------------------------------------
Dim rng_Default, rng_input As Range

Set rng_Default = Selection
'Enable error handling
On Error Resume Next
Set rng_input = Application.InputBox( _
        Prompt:="Select matrix", _
        Title:="Select range", _
        Default:=rng_Default.Address, _
        Type:=8)
'Disable error-handling
On Error GoTo 0
'check that range isn't nothing
If rng_input Is Nothing Then
    MsgBox "Range is empty. Ending procedure.", vbOKOnly, "Error!"
    Exit Sub
End If
'check that range has rows or columns > 1
If rng_input.Rows.count < 2 Or rng_input.Columns.count < 2 Then
    MsgBox "Range must have more than 1 rows and columns. Ending procedure.", vbOKOnly, "Error!"
    Exit Sub
End If
'---assign range to an array----------------------------------------------------------
Dim arr_Input() As Variant
arr_Input = rng_input.value

Dim lng_InputMatrixSize As Long: lng_InputMatrixSize = UBound(arr_Input, 1) - LBound(arr_Input, 1) + 1
'---confirm matrix is symmetric-------------------------------------------------------
If IsMatrixSymmetric(arr_Input) = False Then
    'if not symmetric notify user and end procedure
    MsgBox "Matrix is not symmetric. Ending procedure.", vbOKOnly, "Error!"
    Exit Sub
End If
'---Generate eigen values and eigen vectors------------------------------------------
Dim arr_Eigen As Variant
ReDim arr_Eigen(1 To lng_InputMatrixSize, 1 To lng_InputMatrixSize + 1)

arr_Eigen = EIGEN_JK(arr_Input)
'---Generate QAQTranspose-------------------------------------------------------------
Dim arr_Q, arr_QTRanspose, arr_EigenValueDiagonalMatrix As Variant
Dim lng_CountRow, lng_CountCol As Long

ReDim arr_Q(LBound(arr_Eigen, 1) To UBound(arr_Eigen, 1), LBound(arr_Eigen, 2) To UBound(arr_Eigen, 2) - 1)
ReDim arr_EigenValueDiagonalMatrix(LBound(arr_Eigen, 1) To UBound(arr_Eigen, 1), LBound(arr_Eigen, 2) To UBound(arr_Eigen, 2) - 1)
ReDim arr_QTRanspose(LBound(arr_Eigen, 1) To UBound(arr_Eigen, 1), LBound(arr_Eigen, 2) To UBound(arr_Eigen, 2) - 1)

'Generate Q
For lng_CountRow = LBound(arr_Eigen, 1) To UBound(arr_Eigen, 1) Step 1
    For lng_CountCol = LBound(arr_Eigen, 2) + 1 To UBound(arr_Eigen, 2) Step 1
        arr_Q(lng_CountRow, lng_CountCol - 1) = arr_Eigen(lng_CountRow, lng_CountCol)
    Next lng_CountCol
Next lng_CountRow

'Generate eigen value diagonal matrix
'For lng_CountRow = LBound(arr_Eigen, 1) To UBound(arr_Eigen, 1) Step 1
'    For lng_CountCol = LBound(arr_Eigen, 2) To UBound(arr_Eigen, 2) - 1 Step 1
'        If lng_CountRow = lng_CountCol Then
'        arr_EigenValueDiagonalMatrix(lng_CountRow, lng_CountCol) = arr_Eigen(lng_CountRow, 1)
'            Else: arr_EigenValueDiagonalMatrix(lng_CountRow, lng_CountCol) = 0
'        End If
'    Next lng_CountCol
'Next lng_CountRow

'Generate Q Transpose
arr_QTRanspose = WorksheetFunction.Transpose(arr_Q)
'---Calculate Diagonal Matrix Lambda (Eigenvalues) = QTranspose A Q-------------------
Dim arr_DiagonalMatrixLambda As Variant

arr_DiagonalMatrixLambda = WorksheetFunction.MMult(arr_QTRanspose, arr_Input)
arr_DiagonalMatrixLambda = WorksheetFunction.MMult(arr_DiagonalMatrixLambda, arr_Q)
'---capture output range-------------------------------------------
Dim rng_Output, rng_DefaultOutput As Range

Set rng_DefaultOutput = Selection.Offset(0, Selection.Columns.count + 1)
'Enable error handling
On Error Resume Next
Set rng_Output = Application.InputBox( _
        Prompt:="Select output location", _
        Title:="Select range", _
        Default:=rng_DefaultOutput.Address, _
        Type:=8)

'User cancel scenario
If rng_Output Is Nothing Then
    MsgBox "Output location is empty", vbOKOnly, "Error!"
    Exit Sub
End If
'Disable error-handling
On Error GoTo 0
'---Output QAQTranspose------------------------------------------------------------
Call WriteArrayToWorksheetA1(arr_DiagonalMatrixLambda, ActiveSheet.Name, rng_Output.Address)

End Sub

Private Function EIGEN_JK(ByVal m As Variant) As Variant
'***************************************************************************
'**  Function computes the eigenvalues and eigenvectors for a real        **
'**  symmetric positive definite matrix using the "JK Method".  The       **
'**  first column of the return matrix contains the eigenvalues and       **
'**  the rest of the p+1 columns contain the eigenvectors.                **
'**  See:                                                                 **
'**  KAISER,H.F. (1972) "THE JK METHOD: A PROCEDURE FOR FINDING THE       **
'**  EIGENVALUES OF A REAL SYMMETRIC MATRIX", The Computer Journal,       **
'**  VOL.15, 271-273.                                                     **
'***************************************************************************

Dim a() As Variant, Ematrix() As Double
Dim i As Long, j As Long, k As Long, iter As Long, p As Long
Dim den As Double, hold As Double, Sin_ As Double, num As Double
Dim Sin2 As Double, Cos2 As Double, Cos_ As Double, Test As Double
Dim Tan2 As Double, Cot2 As Double, tmp As Double
Const eps As Double = 1E-16
    
    On Error GoTo EndProc
    
    a = m
    p = UBound(a, 1)
    ReDim Ematrix(1 To p, 1 To p + 1)
    
    For iter = 1 To 15
        
        'Orthogonalize pairs of columns in upper off diag
        For j = 1 To p - 1
            For k = j + 1 To p
                
                den = 0#
                num = 0#
                'Perform single plane rotation
                For i = 1 To p
                    num = num + 2 * a(i, j) * a(i, k)   ': numerator eq. 11
                    den = den + (a(i, j) + a(i, k)) * _
                        (a(i, j) - a(i, k))             ': denominator eq. 11
                Next i
                
                'Skip rotation if aij is zero and correct ordering
                If Abs(num) < eps And den >= 0 Then Exit For
                
                'Perform Rotation
                If Abs(num) <= Abs(den) Then
                    Tan2 = Abs(num) / Abs(den)          ': eq. 11
                    Cos2 = 1 / Sqr(1 + Tan2 * Tan2)     ': eq. 12
                    Sin2 = Tan2 * Cos2                  ': eq. 13
                Else
                    Cot2 = Abs(den) / Abs(num)          ': eq. 16
                    Sin2 = 1 / Sqr(1 + Cot2 * Cot2)     ': eq. 17
                    Cos2 = Cot2 * Sin2                  ': eq. 18
                End If
                
                Cos_ = Sqr((1 + Cos2) / 2)              ': eq. 14/19
                Sin_ = Sin2 / (2 * Cos_)                ': eq. 15/20
                
                If den < 0 Then
                    tmp = Cos_
                    Cos_ = Sin_                         ': table 21
                    Sin_ = tmp
                End If
                
                Sin_ = Sgn(num) * Sin_                  ': sign table 21
                
                'Rotate
                For i = 1 To p
                    tmp = a(i, j)
                    a(i, j) = tmp * Cos_ + a(i, k) * Sin_
                    a(i, k) = -tmp * Sin_ + a(i, k) * Cos_
                Next i
                
            Next k
        Next j
        
        'Test for convergence
        Test = Application.SumSq(a)
        If Abs(Test - hold) < eps And iter > 5 Then Exit For
        hold = Test
    Next iter
    
    If iter = 16 Then MsgBox "JK Iteration has not converged."
    
    'Compute eigenvalues/eigenvectors
    For j = 1 To p
        'Compute eigenvalues
        For k = 1 To p
            Ematrix(j, 1) = Ematrix(j, 1) + a(k, j) ^ 2
        Next k
        Ematrix(j, 1) = Sqr(Ematrix(j, 1))
        
        'Normalize eigenvectors
        For i = 1 To p
            If Ematrix(j, 1) <= 0 Then
                Ematrix(i, j + 1) = 0
            Else
                Ematrix(i, j + 1) = a(i, j) / Ematrix(j, 1)
            End If
        Next i
    Next j
        
    EIGEN_JK = Ematrix
    
    Exit Function
    
EndProc:
    MsgBox Prompt:="Error in function EIGEN_JK!" & vbCr & vbCr & _
        "Error: " & Err.Description & ".", Buttons:=48, _
        Title:="Run time error!"
End Function
