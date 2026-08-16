Attribute VB_Name = "modApiMatrixCreate"
Option Explicit

' Create matrices at the active cell (vector-based creates write to the right of the selection).

''' @Description: Write an n x n identity matrix at the active cell.
Public Sub MatrixCreateIdentity()
    Call modApiMatrixUtilities.WriteCreated("identity")
End Sub

''' @Description: Write an n x m zeros matrix at the active cell.
Public Sub MatrixCreateZeros()
    Call modApiMatrixUtilities.WriteCreated("zeros")
End Sub

''' @Description: Write an n x m ones matrix at the active cell.
Public Sub MatrixCreateOnes()
    Call modApiMatrixUtilities.WriteCreated("ones")
End Sub

''' @Description: Write a Hilbert matrix of order n at the active cell.
Public Sub MatrixCreateHilbert()
    Call modApiMatrixUtilities.WriteCreated("hilbert")
End Sub

''' @Description: Write a uniform random matrix at the active cell.
Public Sub MatrixCreateRandom()
    Call modApiMatrixUtilities.WriteCreated("random")
End Sub

''' @Description: Write an n x n uniform random matrix at the active cell.
Public Sub MatrixCreateRandomSquare()
    Call modApiMatrixUtilities.WriteCreated("randomSquare")
End Sub

''' @Description: Write the exchange (counter-identity) matrix of order n at the active cell.
Public Sub MatrixCreateExchange()
    Call modApiMatrixUtilities.WriteCreated("exchange")
End Sub

''' @Description: Build a diagonal matrix from the selected row or column vector.
Public Sub MatrixCreateDiagonal()
    Call modApiMatrixUtilities.FromSelectionCreate("diag")
End Sub

''' @Description: Symmetric Toeplitz matrix from the selected first row or column.
Public Sub MatrixCreateToeplitz()
    Call modApiMatrixUtilities.FromSelectionCreate("toeplitz")
End Sub

''' @Description: Vandermonde matrix from the selected vector (prompts for column count).
Public Sub MatrixCreateVandermonde()
    Call modApiMatrixUtilities.FromSelectionCreate("vander")
End Sub

''' @Description: Companion matrix of the selected coefficient vector (constant term first).
Public Sub MatrixCreateCompanion()
    Call modApiMatrixUtilities.FromSelectionCreate("companion")
End Sub

' --- Personal123 CreateMatrices ---

Sub CreateExchangeMatrix()
Dim rng_Default, rng_input, rng_Output As Range
Dim lng_MatrixSize, lng_Row, lng_Col As Long
Dim arr_Matrix As Variant

lng_MatrixSize = Application.InputBox( _
        Prompt:="How many rows / columns do you want", _
        Title:="Enter the number of rows / columns ", _
        Default:=4, _
        Type:=1)
'Error checking to capture divide by zero
If lng_MatrixSize = 0 Or lng_MatrixSize = "" Then
    MsgBox "Error! Number of rows/columns cannot be zero or empty. I'll use 4 instead."
    lng_MatrixSize = 4
End If

Set rng_Default = Selection
'Enable error handling
On Error Resume Next
Set rng_input = Application.InputBox( _
        Prompt:="Please output location", _
        Title:="Select range", _
        Default:=rng_Default.Address, _
        Type:=8)
'User cancel scenario
If rng_input Is Nothing Then
    Exit Sub
End If
'Disable error-handling
On Error GoTo 0
'--------------------------------------------------------------
Set rng_Output = rng_input.Resize(lng_MatrixSize, lng_MatrixSize)
'--populate matrix array--------------------------------------
ReDim arr_Matrix(1 To lng_MatrixSize, 1 To lng_MatrixSize)

For lng_Col = 1 To lng_MatrixSize Step 1
    For lng_Row = 1 To lng_MatrixSize Step 1
        If lng_Row + lng_Col = lng_MatrixSize + 1 Then
            arr_Matrix(lng_Row, lng_Col) = 1
            Else
                arr_Matrix(lng_Row, lng_Col) = 0
        End If
    Next
Next
rng_Output = arr_Matrix
End Sub

Sub CreateRandomDiagonalMatrix()
Dim rng_Default, rng_input, rng_Output As Range
Dim lng_MatrixSize, lng_Row, lng_Col As Long
Dim arr_Matrix As Variant

lng_MatrixSize = Application.InputBox( _
        Prompt:="How many rows / columns do you want", _
        Title:="Enter the number of rows / columns ", _
        Default:=4, _
        Type:=1)
'Error checking to capture divide by zero
If lng_MatrixSize = 0 Or lng_MatrixSize = "" Then
    MsgBox "Error! Number of rows/columns cannot be zero or empty. I'll use 4 instead."
    lng_MatrixSize = 4
End If

Set rng_Default = Selection
'Enable error handling
On Error Resume Next
Set rng_input = Application.InputBox( _
        Prompt:="Please output location", _
        Title:="Select range", _
        Default:=rng_Default.Address, _
        Type:=8)
'User cancel scenario
If rng_input Is Nothing Then
    Exit Sub
End If
'Disable error-handling
On Error GoTo 0
'--------------------------------------------------------------
Set rng_Output = rng_input.Resize(lng_MatrixSize, lng_MatrixSize)
'--populate matrix array--------------------------------------
ReDim arr_Matrix(1 To lng_MatrixSize, 1 To lng_MatrixSize)

For lng_Col = 1 To lng_MatrixSize Step 1
    For lng_Row = 1 To lng_MatrixSize Step 1
        If lng_Row = lng_Col Then
            arr_Matrix(lng_Row, lng_Col) = WorksheetFunction.RandBetween(0, 100)
            Else
                arr_Matrix(lng_Row, lng_Col) = 0
        End If
    Next
Next
rng_Output = arr_Matrix
End Sub

Sub CreateZeroMatrix()
Dim rng_Default, rng_input, rng_Output As Range
Dim lng_MatrixSize, lng_Row, lng_Col As Long
Dim arr_Matrix As Variant

lng_MatrixSize = Application.InputBox( _
        Prompt:="How many rows / columns do you want", _
        Title:="Enter the number of rows / columns ", _
        Default:=4, _
        Type:=1)
'Error checking to capture divide by zero
If lng_MatrixSize = 0 Or lng_MatrixSize = "" Then
    MsgBox "Error! Number of rows/columns cannot be zero or empty. I'll use 4 instead."
    lng_MatrixSize = 4
End If

Set rng_Default = Selection
'Enable error handling
On Error Resume Next
Set rng_input = Application.InputBox( _
        Prompt:="Please output location", _
        Title:="Select range", _
        Default:=rng_Default.Address, _
        Type:=8)
'User cancel scenario
If rng_input Is Nothing Then
    Exit Sub
End If
'Disable error-handling
On Error GoTo 0
'--------------------------------------------------------------
Set rng_Output = rng_input.Resize(lng_MatrixSize, lng_MatrixSize)
'--populate matrix array--------------------------------------
ReDim arr_Matrix(1 To lng_MatrixSize, 1 To lng_MatrixSize)

For lng_Col = 1 To lng_MatrixSize Step 1
    For lng_Row = 1 To lng_MatrixSize Step 1
        arr_Matrix(lng_Row, lng_Col) = 0
    Next
Next
rng_Output = arr_Matrix
End Sub

Sub CreateMatrixOfOnes()
Dim rng_Default, rng_input, rng_Output As Range
Dim lng_MatrixSize, lng_Row, lng_Col As Long
Dim arr_Matrix As Variant

lng_MatrixSize = Application.InputBox( _
        Prompt:="How many rows / columns do you want", _
        Title:="Enter the number of rows / columns ", _
        Default:=4, _
        Type:=1)
'Error checking to capture divide by zero
If lng_MatrixSize = 0 Or lng_MatrixSize = "" Then
    MsgBox "Error! Number of rows/columns cannot be zero or empty. I'll use 4 instead."
    lng_MatrixSize = 4
End If

Set rng_Default = Selection
'Enable error handling
On Error Resume Next
Set rng_input = Application.InputBox( _
        Prompt:="Please output location", _
        Title:="Select range", _
        Default:=rng_Default.Address, _
        Type:=8)
'User cancel scenario
If rng_input Is Nothing Then
    Exit Sub
End If
'Disable error-handling
On Error GoTo 0
'--------------------------------------------------------------
Set rng_Output = rng_input.Resize(lng_MatrixSize, lng_MatrixSize)
'--populate matrix array--------------------------------------
ReDim arr_Matrix(1 To lng_MatrixSize, 1 To lng_MatrixSize)

For lng_Col = 1 To lng_MatrixSize Step 1
    For lng_Row = 1 To lng_MatrixSize Step 1
        arr_Matrix(lng_Row, lng_Col) = 1
    Next
Next
rng_Output = arr_Matrix
End Sub

Sub CreateRandomSymmetricMatrix()
'Generates a symmetric matrix using the fact that A x ATranspose = Symmetric matrix

'---Get the size of the matrix-----------------------------------------------------
Dim lng_Value, dRand As Long
'Enable error handling
On Error Resume Next
lng_Value = Application.InputBox( _
        Prompt:="What size square matrix do you want", _
        Title:="Input", _
        Default:=3, _
        Type:=1)
On Error GoTo 0

'check matrix size is positive
If lng_Value < 0 Then
    MsgBox "Matrix size must be a positive number greater than 1. Ending procedure.", vbOKOnly, "Error!"
    Exit Sub
End If
'check matrix size is greater than 1
If lng_Value <= 1 Then
    MsgBox "Matrix size must be greater than 1. Ending procedure.", vbOKOnly, "Error!"
    Exit Sub
End If
'Generate a n x n square mtrix
Dim lng_CounterRow, lng_CounterCol As Long
Dim lMin As Long: lMin = 0
Dim lMax  As Long
Dim arr_Matrix As Variant
ReDim arr_Matrix(1 To lng_Value, 1 To lng_Value)
'Get the maximum value for elements in the matrix
lMax = Application.InputBox( _
        Prompt:="Enter the maximum value matrix elements can take. The minimum value is 0.", _
        Title:="Input", _
        Default:=10, _
        Type:=1)
'check matrix size is greater than 0
If lMax < 0 Then
    MsgBox "Matrix elements must be greater than 0. Ending procedure.", vbOKOnly, "Error!"
    Exit Sub
End If
'Take square root of lMax as in generating the symmetric matrix it is multiplied by itself
lMax = CInt(WorksheetFunction.RoundDown(CDbl(lMax) ^ (1 / lng_Value), 0))

Randomize

For lng_CounterRow = LBound(arr_Matrix, 1) To UBound(arr_Matrix, 1) Step 1
    For lng_CounterCol = LBound(arr_Matrix, 2) To UBound(arr_Matrix, 2) Step 1
        ' Calculate random value, where Value >= Min And Value <= Max
        dRand = lMin + Int(Rnd * (lMax - lMin + 1))
        'assign the random integer to the matrix
        arr_Matrix(lng_CounterRow, lng_CounterCol) = dRand
    Next lng_CounterCol
Next lng_CounterRow
'Multiply the n x n square matrix by its transpose
Dim arr_Symmetric() As Variant
ReDim arr_Symmetric(1 To lng_Value, 1 To lng_Value)
arr_Symmetric = WorksheetFunction.MMult(arr_Matrix, WorksheetFunction.Transpose(arr_Matrix))
'---capture output range-------------------------------------------
Dim rng_DefaultOutput, rng_Output As Range
'Enable error handling
On Error Resume Next
'Get the current selection as the default output range
Set rng_DefaultOutput = Selection.Cells(1)

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
'---Write the matrix to the spreadsheet--------------------------------------------------
Call WriteArrayToWorksheetA1(arr_Symmetric, ActiveSheet.Name, rng_Output.Address)

End Sub

Sub CreateAntiDiagonalMatrix()
Dim rng_Default, rng_input, rng_Output As Range
Dim lng_MatrixSize, lng_Row, lng_Col As Long
Dim arr_Matrix As Variant

lng_MatrixSize = Application.InputBox( _
        Prompt:="How many rows / columns do you want", _
        Title:="Enter the number of rows / columns ", _
        Default:=4, _
        Type:=1)
'Error checking to capture divide by zero
If lng_MatrixSize = 0 Or lng_MatrixSize = "" Then
    MsgBox "Error! Number of rows/columns cannot be zero or empty. I'll use 4 instead."
    lng_MatrixSize = 4
End If

Set rng_Default = Selection
'Enable error handling
On Error Resume Next
Set rng_input = Application.InputBox( _
        Prompt:="Please output location", _
        Title:="Select range", _
        Default:=rng_Default.Address, _
        Type:=8)
'User cancel scenario
If rng_input Is Nothing Then
    Exit Sub
End If
'Disable error-handling
On Error GoTo 0
'--------------------------------------------------------------
Set rng_Output = rng_input.Resize(lng_MatrixSize, lng_MatrixSize)
'--populate matrix array--------------------------------------
ReDim arr_Matrix(1 To lng_MatrixSize, 1 To lng_MatrixSize)

For lng_Col = 1 To lng_MatrixSize Step 1
    For lng_Row = 1 To lng_MatrixSize Step 1
        If lng_Row + lng_Col = lng_MatrixSize + 1 Then
            arr_Matrix(lng_Row, lng_Col) = WorksheetFunction.RandBetween(0, 100)
            Else
                arr_Matrix(lng_Row, lng_Col) = 0
        End If
    Next
Next
rng_Output = arr_Matrix
End Sub

Sub CreateIdentityMatrix()
Dim rng_Default, rng_input, rng_Output As Range
Dim lng_MatrixSize, lng_Row, lng_Col As Long
Dim arr_Matrix As Variant

lng_MatrixSize = Application.InputBox( _
        Prompt:="How many rows / columns do you want", _
        Title:="Enter the number of rows / columns ", _
        Default:=4, _
        Type:=1)
'Error checking to capture divide by zero
If lng_MatrixSize = 0 Or lng_MatrixSize = "" Then
    MsgBox "Error! Number of rows/columns cannot be zero or empty. I'll use 4 instead."
    lng_MatrixSize = 4
End If

Set rng_Default = Selection
'Enable error handling
On Error Resume Next
Set rng_input = Application.InputBox( _
        Prompt:="Please Select a cell", _
        Title:="Select range", _
        Default:=rng_Default.Address, _
        Type:=8)
'User cancel scenario
If rng_input Is Nothing Then
    Exit Sub
End If
'Disable error-handling
On Error GoTo 0
'--------------------------------------------------------------
Set rng_Output = rng_input.Resize(lng_MatrixSize, lng_MatrixSize)
'--populate matrix array--------------------------------------
ReDim arr_Matrix(1 To lng_MatrixSize, 1 To lng_MatrixSize)

For lng_Col = 1 To lng_MatrixSize Step 1
    For lng_Row = 1 To lng_MatrixSize Step 1
        If lng_Row = lng_Col Then
            rng_Output.Cells(lng_Row, lng_Col).value = 1
            Else
                rng_Output.Cells(lng_Row, lng_Col).value = 0
        End If
    Next
Next
End Sub
