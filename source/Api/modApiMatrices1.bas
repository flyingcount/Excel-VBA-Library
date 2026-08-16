Attribute VB_Name = "Custom_Menu13_Matrices1"
'Option Private Module
Option Base 1
Option Explicit
Sub MatrixDiagnosticMessage()
'Run all matrix property functions and output report as an array
Application.Volatile

Dim arrTemp(1 To 22, 1 To 2) As Variant
Dim InputRange As Range
Dim str_SheetName As String
Dim arr_Input As Variant

Set InputRange = Selection
arr_Input = InputRange.value

'Set matrix descriptors
arrTemp(1, 1) = "Number of rows"
arrTemp(2, 1) = "Number of columns"
arrTemp(3, 1) = "Number of elements"
arrTemp(4, 1) = "Matrix size is"
arrTemp(5, 1) = "Matrix is numeric"
arrTemp(6, 1) = "Matrix is zero"
arrTemp(7, 1) = "Matrix of ones"
arrTemp(8, 1) = "Matrix is diagonal"
arrTemp(9, 1) = "Matrix is rectangular"
arrTemp(10, 1) = "Matrix is square"
arrTemp(11, 1) = "Matrix is identity, I"
arrTemp(12, 1) = "Matrix is square and scalar"
arrTemp(13, 1) = "Matrix is symmetric"
arrTemp(14, 1) = "Matrix is square and skew-symmetric"
arrTemp(15, 1) = "Matrix trace is"
arrTemp(16, 1) = "Matrix is orthogonal"
arrTemp(17, 1) = "Matrix is triangular lower"
arrTemp(18, 1) = "Matrix is triangular upper"
arrTemp(19, 1) = "Matrix is unitary Inverse"
arrTemp(20, 1) = "Matrix is unitary transpose equal to inverse"
arrTemp(21, 1) = "Matrix determinant = 1?"
arrTemp(22, 1) = "Matrix determinant"

'Determine matrix properties
arrTemp(1, 2) = InputRange.Rows.count
arrTemp(2, 2) = InputRange.Columns.count
arrTemp(3, 2) = arrTemp(1, 2) * arrTemp(2, 2)
arrTemp(4, 2) = arrTemp(1, 2) & " x " & arrTemp(2, 2)
arrTemp(5, 2) = MatrixNumeric(InputRange)
arrTemp(6, 2) = MatrixZero(InputRange)
arrTemp(7, 2) = MatrixOfOnes(InputRange)
arrTemp(8, 2) = IsMatrixDiagonal(InputRange)
arrTemp(9, 2) = IsMatrixRectangular(InputRange)
arrTemp(10, 2) = IsMatrixSquare(InputRange)
arrTemp(11, 2) = IsMatrixIdentity(InputRange)
arrTemp(12, 2) = IsMatrixScalar(InputRange)
arrTemp(13, 2) = IsMatrixSymmetric(InputRange)
arrTemp(14, 2) = MatrixSkewSymmetric(InputRange)
arrTemp(15, 2) = MtrxTrace(InputRange)
arrTemp(16, 2) = IsMatrixOrthogonal(InputRange)
arrTemp(17, 2) = IsMatrixLowerTriangular(InputRange)
arrTemp(18, 2) = IsMatrixUpperTriangular(InputRange)
arrTemp(19, 2) = IsMatrixUnitaryInverse(InputRange)
arrTemp(20, 2) = IsMatrixUnitaryTransposeEqualInverse(InputRange)
arrTemp(21, 2) = IsMatrixDeterminantEqualToOne(InputRange)
arrTemp(22, 2) = CalculateMatrixDeterminant(InputRange)

'Defines output worksheet name
str_SheetName = "Matrix Analysis"
'---Create if output worksheet already exists. Delete it if it does and create a new sheet-----------
Call CheckExistenceAndDeleteOutputSheet(str_SheetName)
Call CreateOutputSheet(str_SheetName)
Sheets(str_SheetName).Range("A1").Resize(UBound(arrTemp, 1), UBound(arrTemp, 2)) = arrTemp

With Sheets(str_SheetName)
    With Columns("A")
        .ColumnWidth = 40
    End With
End With

End Sub

Function IsMatrixSquare(var_Input As Variant) As Boolean
'Determine whether matrix is square i.e. m = n
Application.Volatile

Dim arr_Input As Variant
If TypeName(var_Input) = "Range" Then
    ' Use the range shape, not .Value: a merged block can be several cells
    ' yet .Value is still a scalar, which makes UBound type-mismatch.
    IsMatrixSquare = (var_Input.Rows.Count = var_Input.Columns.Count)
    Exit Function
End If
arr_Input = ToMatrix2D(var_Input)
IsMatrixSquare = (UBound(arr_Input, 1) - LBound(arr_Input, 1) = UBound(arr_Input, 2) - LBound(arr_Input, 2))
End Function

' Personal Fn_MatricesRng names used by MatrixSkewSymmetric / MatrixMultDefined.
Private Function MatrixSquare(InputRange As Range) As Boolean
    MatrixSquare = IsMatrixSquare(InputRange)
End Function

Private Function MatrixRows(InputRange As Range) As Long
    MatrixRows = InputRange.Rows.Count
End Function

Private Function MatrixColumns(InputRange As Range) As Long
    MatrixColumns = InputRange.Columns.Count
End Function

' Range.Value of one cell is a scalar, not an array. UBound(..., 2) then raises
' type mismatch. Always return a 2D array (1-based when we wrap a scalar / 1D).
Public Function ToMatrix2D(ByVal var_Input As Variant) As Variant
    Dim out() As Variant
    Dim i As Long
    Dim n As Long
    Dim lb As Long
    Dim rng As Range
    Dim raw As Variant

    If TypeName(var_Input) = "Range" Then
        Set rng = var_Input
        raw = rng.Value2
        If IsArray(raw) Then
            ToMatrix2D = raw
            Exit Function
        End If
        ' Merged (or otherwise non-array) Value: keep the selection's row x col shape.
        ReDim out(1 To rng.Rows.Count, 1 To rng.Columns.Count)
        out(1, 1) = raw
        ToMatrix2D = out
        Exit Function
    End If

    If Not IsArray(var_Input) Then
        ReDim out(1 To 1, 1 To 1)
        out(1, 1) = var_Input
        ToMatrix2D = out
        Exit Function
    End If

    On Error Resume Next
    n = UBound(var_Input, 2)
    If Err.Number = 0 Then
        On Error GoTo 0
        ToMatrix2D = var_Input
        Exit Function
    End If
    Err.Clear
    On Error GoTo 0

    lb = LBound(var_Input)
    n = UBound(var_Input) - lb + 1
    ReDim out(1 To n, 1 To 1)
    For i = 1 To n
        out(i, 1) = var_Input(lb + i - 1)
    Next i
    ToMatrix2D = out
End Function

Function IsMatrixRectangular(var_Input As Variant) As Boolean
'Determine whether matrix is rectangular i.e. m <> n
Application.Volatile

Dim arr_Input As Variant
arr_Input = ToMatrix2D(var_Input)
IsMatrixRectangular = Not IsMatrixSquare(arr_Input)
End Function

Function IsMatrixIdentity(var_Input As Variant) As Boolean
'Determine whether the matrix is an identity matrix i.e. the
'   matrix is square
'   diagonal values = 1
'   all other values = 0

Application.Volatile

'---Find out if input is a range or array and determine sizes--------------------------
Dim arr_Input As Variant
'Defines the size whether the input is a range or an array
arr_Input = ToMatrix2D(var_Input)

Dim r As Long: r = 1
Dim c As Long: c = 1
Dim DiagonalValueSum As Double: DiagonalValueSum = 0
Dim NonDiagonalValueSum As Double: NonDiagonalValueSum = 0

'check if all non-diagonal elements = 0 and at least one diagonal value <> 0
For r = LBound(arr_Input, 1) To UBound(arr_Input, 1) Step 1
    For c = LBound(arr_Input, 2) To UBound(arr_Input, 2) Step 1
        'counts the number of zero diagonal elements. Should be at
        If r = c Then DiagonalValueSum = arr_Input(r, c) + DiagonalValueSum
        'counts the number of non zero non diagonal elements
        If r <> c Then NonDiagonalValueSum = arr_Input(r, c) + NonDiagonalValueSum
    Next c
Next r

If DiagonalValueSum = 1 * (UBound(arr_Input, 1) - LBound(arr_Input, 1) + 1) And NonDiagonalValueSum = 0 And IsMatrixSquare(arr_Input) = True Then
    IsMatrixIdentity = True
        Else: IsMatrixIdentity = False
End If

End Function

Private Function MatrixZero(InputRange As Range) As Boolean
'Determine whether all the elements in the matrix are equal to zero
Application.Volatile
Dim n As Long
Dim c As Range

n = 0

For Each c In InputRange
    If c.value = 0 Then
        n = n
            Else: n = n + 1
                End If
Next

    If n <> 0 Then
        MatrixZero = False
            Else: MatrixZero = True
                End If
                
End Function
Private Function MatrixOfOnes(InputRange As Range) As Boolean
'Determine whether all the elements in the matrix are equal to 1
Application.Volatile
Dim n As Long
Dim c As Range

n = 0

For Each c In InputRange
    If c.value = 1 Then
        n = n + 1
            Else: n = n
                End If
Next

    If n = InputRange.Cells.count Then
        MatrixOfOnes = True
            Else: MatrixOfOnes = False
                End If
                
End Function

Function IsMatrixDiagonal(var_Input As Variant) As Boolean
'Determine whether matrix is diagonal
Dim lng_Row, lng_Col As Long
Dim int_Flag As Integer: int_Flag = 1

Application.Volatile
'---Find out if input is a range or array and determine sizes--------------------------
Dim arr_Input As Variant
'Defines the size whether the input is a range or an array
arr_Input = ToMatrix2D(var_Input)

If IsMatrixSquare(arr_Input) = False Then
    IsMatrixDiagonal = False
    Else
        For lng_Row = LBound(arr_Input, 1) To UBound(arr_Input, 1) Step 1
            For lng_Col = LBound(arr_Input, 2) To UBound(arr_Input, 2) Step 1
                If lng_Row <> lng_Col And arr_Input(lng_Row, lng_Col) <> 0 Then
                    int_Flag = int_Flag * 0
                End If
            Next lng_Col
        Next lng_Row
        If int_Flag = 1 Then
            IsMatrixDiagonal = True
            Else: IsMatrixDiagonal = False
        End If
End If
            
End Function

Private Function MatrixIdentity(var_Input As Variant) As Boolean
'Determine whether the matrix is an identity matrix i.e. the
'   matrix is square
'   diagonal values = 1
'   all other values = 0
Application.Volatile
'---Find out if input is a range or array and determine sizes--------------------------
Dim arr_Input As Variant
Dim lng_NumRows, lng_NumCols As Long
'Defines the size whether the input is a range or an array
arr_Input = ToMatrix2D(var_Input)
lng_NumRows = UBound(arr_Input, 1) - LBound(arr_Input, 1) + 1
lng_NumCols = UBound(arr_Input, 2) - LBound(arr_Input, 2) + 1

Dim r As Long: r = 1
Dim c As Long: c = 1
Dim DiagonalValueSum As Double: DiagonalValueSum = 0
Dim NonDiagonalValueSum As Double: NonDiagonalValueSum = 0

'check if all non-diagonal elements = 0 and at least one diagonal value <> 0
For r = 1 To lng_NumRows Step 1
    For c = 1 To lng_NumCols Step 1
        'counts the number of zero diagonal elements. Should be at
        If r = c Then DiagonalValueSum = arr_Input(r, c) + DiagonalValueSum
        'counts the number of non zero non diagonal elements
        If r <> c Then NonDiagonalValueSum = arr_Input(r, c) + NonDiagonalValueSum
     
    Next c
Next r

If DiagonalValueSum = 1 * lng_NumRows And NonDiagonalValueSum = 0 And lng_NumRows = lng_NumCols Then
    MatrixIdentity = True
        Else: MatrixIdentity = False
End If

End Function

Function IsMatrixScalar(var_Input As Variant) As Boolean
'Determine whether the matrix is a scalar matrix i.e. the
'   matrix is square
'   diagonal values are equal
'   all other values = 0
Application.Volatile
'---Find out if input is a range or array and determine sizes--------------------------
Dim arr_Input As Variant
Dim lng_NumRows, lng_NumCols As Long
'Defines the size whether the input is a range or an array
arr_Input = ToMatrix2D(var_Input)
lng_NumRows = UBound(arr_Input, 1) - LBound(arr_Input, 1) + 1
lng_NumCols = UBound(arr_Input, 2) - LBound(arr_Input, 2) + 1

Dim arrDiagonalValues() As Variant
Dim r, c, DiagonalValueSum, NonDiagonalNonZeroCounter, EqualDiagonalValuesCount As Long

r = 1
c = 1
DiagonalValueSum = 0
NonDiagonalNonZeroCounter = 0

'Set dimensions for array to receive diagonal values
ReDim arrDiagonalValues(1 To lng_NumRows)

'check if all non-diagonal elements = 0 and at least one diagonal value <> 0
For r = 1 To lng_NumRows Step 1
    For c = 1 To lng_NumCols Step 1
        'writes the diagonal elements to a one dimensional array to be checked for equality later
        If r = c Then arrDiagonalValues(r) = arr_Input(r, c)
        'counts the number of non zero non diagonal elements
        If r <> c And arr_Input(r, c) <> 0 Then NonDiagonalNonZeroCounter = NonDiagonalNonZeroCounter + 1
     
    Next c
Next r

'reset counter r
'r = 0
EqualDiagonalValuesCount = 0

For r = 2 To lng_NumRows Step 1
    If arrDiagonalValues(r) = arrDiagonalValues(1) Then
        EqualDiagonalValuesCount = EqualDiagonalValuesCount + 1
    End If
Next r

If NonDiagonalNonZeroCounter = 0 And lng_NumRows = lng_NumCols And EqualDiagonalValuesCount + 1 = lng_NumRows Then
    IsMatrixScalar = True
        Else: IsMatrixScalar = False
End If

End Function
Function IsMatrixSymmetric(var_Input As Variant) As Boolean
'The entries of a symmetric matrix are symmetric with respect to the main diagonal, diagonal matrices
'are also symmetric. They must also be square.

Application.Volatile
'---Find out if input is a range or array and determine sizes--------------------------
Dim arr_Input As Variant
Dim int_NumRows, int_NumCols As Integer
'Defines the size whether the input is a range or an array
arr_Input = ToMatrix2D(var_Input)
int_NumRows = UBound(arr_Input, 1) - LBound(arr_Input, 1) + 1
int_NumCols = UBound(arr_Input, 2) - LBound(arr_Input, 2) + 1

'a symmetric matrix must be square
If IsMatrixSquare(arr_Input) = False Then
    IsMatrixSymmetric = False
    Else
        'a symmetric matrix must equal its transpose
        IsMatrixSymmetric = IsMatrixEqual(arr_Input, WorksheetFunction.Transpose(arr_Input))
End If

End Function

Private Function MatrixSkewSymmetric(InputRange As Range) As Boolean
'The entries of a skew symmetric matrix are symmetric, but negative, with respect to the main diagonal,
'All the diagonal elements must equal 0
'The matrix must also be square.

Application.Volatile

'Checks if matrix is square. If not, it cannot be symmetric.
If MatrixSquare(InputRange) = False Then
    MatrixSkewSymmetric = False
        Else
'----Code start for Else -------------
Dim arrTemp() As Variant
Dim r, c, NonDiagonalNonNegativeCounter, DiagonalNonZeroElementCounter As Long

r = 1
c = 1
DiagonalNonZeroElementCounter = 0
NonDiagonalNonNegativeCounter = 0

'Set dimensions for temporary array to receive the Input Range values
arrTemp = ToMatrix2D(InputRange)

'check if all non-diagonal elements are symmetric (equal) about the diagonal
For r = LBound(arrTemp, 1) To UBound(arrTemp, 1)
    For c = LBound(arrTemp, 2) To UBound(arrTemp, 2)
        'counts the number of zero diagonal elements. Should be at
        If r = c And arrTemp(r, c) <> 0 Then DiagonalNonZeroElementCounter = DiagonalNonZeroElementCounter + 1
        'counts the number of non zero non diagonal elements
        If r <> c And arrTemp(r, c) + arrTemp(c, r) <> 0 Then NonDiagonalNonNegativeCounter = NonDiagonalNonNegativeCounter + 1
     
    Next c
Next r

If MatrixSquare(InputRange) = True And NonDiagonalNonNegativeCounter = 0 And DiagonalNonZeroElementCounter = 0 Then
    MatrixSkewSymmetric = True
        Else: MatrixSkewSymmetric = False
End If
'----Code end for Else -------------
End If

End Function
Private Function MatrixNumeric(InputRange As Range) As Boolean
Dim NonNumericCounter As Long
Dim rngCell As Range

NonNumericCounter = 0

For Each rngCell In InputRange.Cells
    If IsNumeric(rngCell.value) Then
        NonNumericCounter = NonNumericCounter
            Else: NonNumericCounter = NonNumericCounter + 1
    End If
Next

If NonNumericCounter <> 0 Then
    MatrixNumeric = False
        Else: MatrixNumeric = True
End If

End Function
Private Function MatrixMultDefined(Matrix1 As Range, Matrix2 As Range) As Boolean
'Checks whether the multiplication of two matrices is defined by checking their sizes. More specifically
'whether the columns of Matrix1= rowsof Matrix2

Dim NumColsMatrix1, NumRowsMatrix2 As Long

NumColsMatrix1 = MatrixColumns(Matrix1)
NumRowsMatrix2 = MatrixRows(Matrix2)

If NumColsMatrix1 = NumRowsMatrix2 Then
    MatrixMultDefined = True
        Else
            MatrixMultDefined = False
End If

End Function
Function MtrxTrace(rng_Mtrx As Range) As Double
'tests whether a matrix satisfies the orthogonal criteria that each and every row has a magnitude = 1.
Dim lng_Rows, lng_Cols As Long
Dim arr_Mtrx As Variant
Dim dbl_DiagonalSum As Double
Application.Volatile

'---validation--------------------------------------------------------------
If rng_Mtrx.Rows.count <> rng_Mtrx.Columns.count Then
'    MsgBox "The number of rows must equal the number of columns. The program will end.", vbOKOnly, "Error"
    Exit Function
End If
'-----------------------------------------------------------------
ReDim arr_Mtrx(1 To rng_Mtrx.Rows.count, 1 To rng_Mtrx.Columns.count)
arr_Mtrx = ToMatrix2D(rng_Mtrx)
'-----------------------------------------------------------------
ReDim arr_MtrxDiagonalSum(1 To rng_Mtrx.Rows.count)
dbl_DiagonalSum = 0
For lng_Rows = 1 To UBound(arr_Mtrx, 1) Step 1
    dbl_DiagonalSum = dbl_DiagonalSum + arr_Mtrx(lng_Rows, lng_Rows)
Next
MtrxTrace = dbl_DiagonalSum
End Function

Function IsMatrixOrthogonal(var_Input As Variant) As Boolean
Dim int_Criteria As Integer: int_Criteria = 1

Application.Volatile
'---Find out if input is a range or array and determine sizes--------------------------
Dim arr_Input As Variant
'Defines the size whether the input is a range or an array
arr_Input = ToMatrix2D(var_Input)

'Check Matrix is square
If IsMatrixSquare(arr_Input) = True Then
    int_Criteria = 1 * int_Criteria
    Else:
        'If the matrix isn't square it cannot be orthogonal and the inverse is not defined and will cause
        'an error when calculating the inverse (following test), so stop the function ad return FALSE
        IsMatrixOrthogonal = False
        Exit Function
End If
'Matrix is unitary i.e. does AA-1=AAT=I
If IsMatrixUnitaryInverse(arr_Input) = True Then
    int_Criteria = 1 * int_Criteria
    Else: int_Criteria = 0 * int_Criteria
End If
'Determinant = 1 Or -1
If IsMatrixDeterminantEqualToOne(arr_Input) = True Then
    int_Criteria = 1 * int_Criteria
    Else: int_Criteria = 0 * int_Criteria
End If
'final reckoning. All criteria must be true i.e. int_Criteria=1
If int_Criteria = 1 Then
    IsMatrixOrthogonal = True
    Else: IsMatrixOrthogonal = False
End If

End Function

Private Function IsMatrixUnitaryTranspose(var_Input As Variant) As Boolean
'Matrix x Inverse = Identity
Dim arr_Transpose, arr_Multiply As Variant

Application.Volatile
'---Find out if input is a range or array and determine sizes--------------------------
Dim arr_Input As Variant
Dim int_NumRows, int_NumCols As Integer
'Defines the size whether the input is a range or an array
arr_Input = ToMatrix2D(var_Input)
int_NumRows = UBound(arr_Input, 1) - LBound(arr_Input, 1) + 1
int_NumCols = UBound(arr_Input, 2) - LBound(arr_Input, 2) + 1

arr_Transpose = WorksheetFunction.Transpose(arr_Input)
arr_Multiply = WorksheetFunction.MMult(arr_Input, arr_Transpose)

If IsMatrixIdentity(arr_Multiply) = True Then
    IsMatrixUnitaryTranspose = True
    Else: IsMatrixUnitaryTranspose = False
End If

End Function

Function IsMatrixLowerTriangular(var_Input As Variant) As Boolean
'Determine whether matrix is lower triangular
Dim lng_Row, lng_Col As Long
Dim dbl_Lower As Double: dbl_Lower = 0
Dim dbl_Diagonal As Double: dbl_Diagonal = 0

Application.Volatile
'---Find out if input is a range or array and determine sizes--------------------------
Dim arr_Input As Variant
Dim int_NumRows, int_NumCols As Integer
'Defines the size whether the input is a range or an array
arr_Input = ToMatrix2D(var_Input)
int_NumRows = UBound(arr_Input, 1) - LBound(arr_Input, 1) + 1
int_NumCols = UBound(arr_Input, 2) - LBound(arr_Input, 2) + 1

If IsMatrixSquare(arr_Input) = False Then
    IsMatrixLowerTriangular = False
    Else
        For lng_Row = LBound(arr_Input, 1) To UBound(arr_Input, 1) Step 1
            For lng_Col = LBound(arr_Input, 2) To UBound(arr_Input, 2) Step 1
                If lng_Row = lng_Col Then
                    dbl_Diagonal = dbl_Diagonal + arr_Input(lng_Row, lng_Col)
                End If
                If lng_Row > lng_Col Then
                    dbl_Lower = dbl_Lower + arr_Input(lng_Row, lng_Col)
                End If
            Next lng_Col
        Next lng_Row
End If

If dbl_Lower = 0 And dbl_Diagonal > 0 Then
    IsMatrixLowerTriangular = True
    Else: IsMatrixLowerTriangular = False
End If
            
End Function

Function IsMatrixUpperTriangular(var_Input As Variant) As Boolean
'Determine whether matrix is lower triangular
Dim lng_Row, lng_Col As Long
Dim dbl_Upper As Double: dbl_Upper = 0
Dim dbl_Diagonal As Double: dbl_Diagonal = 0

Application.Volatile
'---Find out if input is a range or array and determine sizes--------------------------
Dim arr_Input As Variant
'Defines the size whether the input is a range or an array
arr_Input = ToMatrix2D(var_Input)

If IsMatrixSquare(arr_Input) = False Then
    IsMatrixUpperTriangular = False
    Else
        For lng_Row = LBound(arr_Input, 1) To UBound(arr_Input, 1) Step 1
            For lng_Col = LBound(arr_Input, 2) To UBound(arr_Input, 2) Step 1
                If lng_Row = lng_Col Then
                    dbl_Diagonal = dbl_Diagonal + arr_Input(lng_Row, lng_Col)
                End If
                If lng_Row < lng_Col Then
                    dbl_Upper = dbl_Upper + arr_Input(lng_Row, lng_Col)
                End If
            Next lng_Col
        Next lng_Row
        If dbl_Upper = 0 And dbl_Diagonal > 0 Then
            IsMatrixUpperTriangular = True
            Else: IsMatrixUpperTriangular = False
        End If
End If
            
End Function

Function IsMatrixUnitaryTransposeEqualInverse(var_Input As Variant) As Boolean
'Matrix transpose = Matrix inverse
Dim arr_Inverse, arr_Transpose, arr_Multiply As Variant
Dim int_Criteria As Integer: int_Criteria = 1
Dim lng_CounterRow, lng_CounterCol As Long

Application.Volatile

'---Find out if input is a range or array and determine sizes--------------------------
Dim arr_Input As Variant
Dim int_NumRows, int_NumCols As Integer
'Defines the size whether the input is a range or an array
arr_Input = ToMatrix2D(var_Input)
int_NumRows = UBound(arr_Input, 1) - LBound(arr_Input, 1) + 1
int_NumCols = UBound(arr_Input, 2) - LBound(arr_Input, 2) + 1

arr_Transpose = WorksheetFunction.Transpose(arr_Input)
'if the matrix is square proceed with the check
If IsMatrixSquare(arr_Input) = True Then
    arr_Inverse = WorksheetFunction.MInverse(arr_Input)
    'Check each array element is identical
    For lng_CounterRow = LBound(arr_Transpose, 1) To UBound(arr_Transpose, 1) Step 1
        For lng_CounterCol = LBound(arr_Transpose, 2) To UBound(arr_Transpose, 2) Step 1
            If arr_Transpose(lng_CounterRow, lng_CounterCol) = arr_Inverse(lng_CounterRow, lng_CounterCol) Then
                int_Criteria = 1 * int_Criteria
                    Else: int_Criteria = 0 * int_Criteria
            End If
        Next lng_CounterCol
    Next lng_CounterRow

    If int_Criteria = 1 Then
        IsMatrixUnitaryTransposeEqualInverse = True
            Else: IsMatrixUnitaryTransposeEqualInverse = False
    End If
    
    'if the matrix isn't square the inverse doesn't exist so go to ELSE and return FALSE
    Else: IsMatrixUnitaryTransposeEqualInverse = False
End If
End Function

Function IsMatrixUnitaryInverse(var_Input As Variant) As Boolean
'Matrix x Inverse = Identity
Dim arr_Multiply As Variant

Application.Volatile

'---Find out if input is a range or array and determine sizes--------------------------
Dim arr_Input As Variant
'Defines the size whether the input is a range or an array
arr_Input = ToMatrix2D(var_Input)

If IsMatrixSquare(arr_Input) = True Then
    arr_Multiply = WorksheetFunction.MMult(arr_Input, WorksheetFunction.MInverse(arr_Input))
    If IsMatrixIdentity(arr_Multiply) = True Then
        IsMatrixUnitaryInverse = True
        Else: IsMatrixUnitaryInverse = False
    End If
    Else: IsMatrixUnitaryInverse = False
End If

End Function

Function IsMatrixDeterminantEqualToOne(var_Input As Variant) As Boolean
'Determine whether matrix determinant is +1 or -1, one of the criteria for orthogonal matrices
Application.Volatile

'---Find out if input is a range or array and determine sizes--------------------------
Dim arr_Input As Variant
'Defines the size whether the input is a range or an array
arr_Input = ToMatrix2D(var_Input)

'Check Matrix is square
If IsMatrixSquare(arr_Input) = False Then
    IsMatrixDeterminantEqualToOne = False
    Else
        If Abs(WorksheetFunction.MDeterm(arr_Input)) = 1 Then
            IsMatrixDeterminantEqualToOne = True
            Else: IsMatrixDeterminantEqualToOne = False
        End If
End If

End Function

Function CalculateMatrixDeterminant(var_Input As Variant) As Variant
'calculates the determinant if the matrix is square
Application.Volatile

'---Find out if input is a range or array and determine sizes--------------------------
Dim arr_Input As Variant
'Defines the size whether the input is a range or an array
arr_Input = ToMatrix2D(var_Input)

'Check Matrix is square
If IsMatrixSquare(arr_Input) = False Then
    CalculateMatrixDeterminant = "Not defined. Matrix isn't square."
    Else
        CalculateMatrixDeterminant = WorksheetFunction.MDeterm(arr_Input)
End If

End Function
