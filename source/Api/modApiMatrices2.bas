Attribute VB_Name = "Custom_Menu13_Matrices2"
Option Explicit

Sub MtrxMultFormulae()
Dim rng_Mtrx1, rng_Mtrx2, rng_Default, rng_Output, rng_DefaultOutput As Range
Dim obj_Row, obj_Col As Range
Dim lng_Mtrx1Row, lng_Mtrx1Col, lng_Mtrx2Row, lng_Mtrx2Col, lng_MtrxMultRow, lng_MtrxMultCol As Long
Dim lng_RowCounter, lng_ColCounter, lng_Element As Long
Dim lng_OutputRefRow, lng_OutputRefCol As Long
Dim str_Accumulator, str_TempRow, str_TempCol As String

Set rng_Default = Selection
'-----------------------------------------------------------------
'Enable error handling
On Error Resume Next
Set rng_Mtrx1 = Application.InputBox( _
        Prompt:="Select matrix 1", _
        Title:="Select range", _
        Default:=rng_Default.Address, _
        Type:=8)
'User cancel scenario
If rng_Mtrx1 Is Nothing Then
    MsgBox "Matrix 1 is empty", vbOKOnly, "Error!"
    Exit Sub
End If
'Disable error-handling
On Error GoTo 0
'-----------------------------------------------------------------
'Enable error handling
On Error Resume Next
Set rng_Mtrx2 = Application.InputBox( _
        Prompt:="Select matrix 2", _
        Title:="Select range", _
        Default:=rng_Default.Address, _
        Type:=8)

'User cancel scenario
If rng_Mtrx2 Is Nothing Then
    MsgBox "Matrix 2 is empty", vbOKOnly, "Error!"
    Exit Sub
End If
'Disable error-handling
On Error GoTo 0
'-----------------------------------------------------------------
lng_Mtrx1Row = rng_Mtrx1.Rows.count
lng_Mtrx1Col = rng_Mtrx1.Columns.count
lng_Mtrx2Row = rng_Mtrx2.Rows.count
lng_Mtrx2Col = rng_Mtrx2.Columns.count

If lng_Mtrx1Col <> lng_Mtrx2Row Then
    MsgBox "Columns of matrix 1 <> rows of matrix 2.", vbOKOnly, "Error!"
    Exit Sub
End If

lng_MtrxMultRow = lng_Mtrx1Row
lng_MtrxMultCol = lng_Mtrx2Col

str_Accumulator = ""
lng_RowCounter = 1
lng_ColCounter = 1

ReDim arr_TempRow(1, 1 To lng_Mtrx1Col)
ReDim arr_TempCol(1 To lng_Mtrx2Row, 1)
ReDim arr_TempRowTranspose(1 To lng_Mtrx1Col, 1)
ReDim arr_MtrxMult(1 To lng_Mtrx1Row, 1 To lng_Mtrx2Col)

For Each obj_Col In rng_Mtrx2.Columns
    For Each obj_Row In rng_Mtrx1.Rows
    
        For lng_Element = 1 To lng_Mtrx1Col Step 1
            str_TempRow = obj_Row.Cells(lng_Element).Address(RowAbsolute:=False, columnAbsolute:=False)
            str_TempCol = obj_Col.Cells(lng_Element).Address(RowAbsolute:=False, columnAbsolute:=False)
            
            str_Accumulator = str_Accumulator & "+" & str_TempRow & "*" & str_TempCol
        Next
        
        arr_MtrxMult(lng_RowCounter, lng_ColCounter) = "=" & str_Accumulator

        str_Accumulator = ""
        lng_RowCounter = lng_RowCounter + 1
    Next obj_Row
    str_Accumulator = ""
    lng_ColCounter = lng_ColCounter + 1
    lng_RowCounter = 1
Next obj_Col
'---set default output range-------------------------------------
lng_OutputRefRow = rng_Mtrx2.row
lng_OutputRefCol = rng_Mtrx2.Column + lng_Mtrx2Col + 1
Set rng_DefaultOutput = Cells(lng_OutputRefRow, lng_OutputRefCol)
'---capture output range-------------------------------------------
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
'---output results------------------------------------------------------

rng_Output.Resize(UBound(arr_MtrxMult, 1), UBound(arr_MtrxMult, 2)) = arr_MtrxMult

End Sub
Sub MtrxMultScalarFormulae()
Dim rng_Mtrx1, rng_Scalar, rng_MtrxMult, rng_Default, rng_DefaultOutput, rng_Output As Range
Dim obj_Row, obj_Col As Range
Dim lng_Mtrx1Row, lng_Mtrx1Col, lng_ScalarRow, lng_ScalarCol, lng_MtrxMultRow, lng_MtrxMultCol As Long
Dim lng_RowCounter, lng_ColCounter, lng_Element As Long
Dim lng_OutputRefRow, lng_OutputRefCol As Long
Dim str_Scalar As String

Set rng_Default = Selection
'-----------------------------------------------------------------
'Enable error handling
On Error Resume Next
Set rng_Mtrx1 = Application.InputBox( _
        Prompt:="Select matrix 1", _
        Title:="Select range", _
        Default:=rng_Default.Address, _
        Type:=8)
'User cancel scenario
If rng_Mtrx1 Is Nothing Then
    MsgBox "Matrix 1 is empty", vbOKOnly, "Error!"
    Exit Sub
End If
'Disable error-handling
On Error GoTo 0
'-----------------------------------------------------------------
'Enable error handling
On Error Resume Next
Set rng_Scalar = Application.InputBox( _
        Prompt:="Select scalar range", _
        Title:="Select range", _
        Default:=rng_Default.Address, _
        Type:=8)

'User cancel scenario
If rng_Scalar Is Nothing Then
    MsgBox "Scalar range is empty", vbOKOnly, "Error!"
    Exit Sub
End If
'Disable error-handling
On Error GoTo 0
'-----------------------------------------------------------------
lng_Mtrx1Row = rng_Mtrx1.Rows.count
lng_Mtrx1Col = rng_Mtrx1.Columns.count
lng_ScalarRow = rng_Scalar.Rows.count
lng_ScalarCol = rng_Scalar.Columns.count

str_Scalar = rng_Scalar.Address(RowAbsolute:=False, columnAbsolute:=False)

lng_RowCounter = 1
lng_ColCounter = 1

ReDim arr_MtrxMult(1 To lng_Mtrx1Row, 1 To lng_Mtrx1Col)

For lng_ColCounter = 1 To rng_Mtrx1.Columns.count
    For lng_RowCounter = 1 To rng_Mtrx1.Rows.count
    
        arr_MtrxMult(lng_RowCounter, lng_ColCounter) = "=" & rng_Mtrx1.Cells(lng_RowCounter, lng_ColCounter). _
            Address(RowAbsolute:=False, columnAbsolute:=False) & "*" & str_Scalar
    
    Next
Next
'---set default output range-------------------------------------
lng_OutputRefRow = rng_Scalar.row - 1
lng_OutputRefCol = rng_Scalar.Column + lng_ScalarCol + 1
Set rng_DefaultOutput = Cells(lng_OutputRefRow, lng_OutputRefCol)
'---capture output range-------------------------------------------
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
'---output results------------------------------------------------------
rng_Output.Resize(UBound(arr_MtrxMult, 1), UBound(arr_MtrxMult, 2)) = arr_MtrxMult

End Sub
Sub TransposeMatrixFormulae()
Dim rng_Default, rng_DefaultOutput, rng_Mtrx, rng_Output As Range
Dim lng_Rows, lng_Cols, lng_OutputRefRow, lng_OutputRefCol As Long
Dim arr_Temp, arr_Output As Variant

Set rng_Default = Selection
'-----------------------------------------------------------------
'Enable error handling
On Error Resume Next
Set rng_Mtrx = Application.InputBox( _
        Prompt:="Select square matrix to transpose", _
        Title:="Select range", _
        Default:=rng_Default.Address, _
        Type:=8)
'User cancel scenario
If rng_Mtrx Is Nothing Then
    MsgBox "Matrix is empty", vbOKOnly, "Error!"
    Exit Sub
End If
'Disable error-handling
On Error GoTo 0
'---------------------------------------------------------------------------
If rng_Mtrx.Rows.count <> rng_Mtrx.Columns.count Then
    MsgBox "Matrix MUST be square. This one isn't.", vbOKOnly, "Error!"
    Exit Sub
End If
'---compute default output location----------------------------------------------------
lng_OutputRefRow = rng_Mtrx.row
lng_OutputRefCol = rng_Mtrx.Column + rng_Mtrx.Columns.count + 1
Set rng_DefaultOutput = ActiveSheet.Cells(lng_OutputRefRow, lng_OutputRefCol)
'----capture output location-------------------------------------------------
On Error Resume Next
Set rng_Output = Application.InputBox( _
        Prompt:="Select output location", _
        Title:="Select range", _
        Default:=rng_DefaultOutput.Address, _
        Type:=8)
'User cancel scenario
If rng_Output.Cells.count <> 1 Then
    MsgBox "Select a single cell where output starts", vbOKOnly, "Error!"
    Exit Sub
End If
'Disable error-handling
On Error GoTo 0
'------------------------------------------------------------------------------
ReDim arr_Output(0 To rng_Mtrx.Rows.count, 0 To rng_Mtrx.Columns.count)

Dim lng_OutputStartRow, lng_OutputStartCol As Long
lng_OutputStartRow = rng_Output.row
lng_OutputStartCol = rng_Output.Column

Dim lng_InputStartRow, lng_InputStartCol As Long
lng_InputStartRow = rng_Mtrx.row
lng_InputStartCol = rng_Mtrx.Column

Dim lng_DiffOutputAndInputRow, lng_DiffOutputAndInputCol As Long
lng_DiffOutputAndInputRow = lng_OutputStartRow - lng_InputStartRow
lng_DiffOutputAndInputCol = lng_OutputStartCol - lng_InputStartCol

Dim lng_RowCounter, lng_ColCounter As Long
lng_RowCounter = 0
lng_ColCounter = 0

lng_Rows = 0
lng_Cols = 0

For lng_Cols = 0 To (rng_Mtrx.Columns.count - 1) Step 1
    
    For lng_Rows = 0 To (rng_Mtrx.Rows.count - 1) Step 1
        
        lng_RowCounter = lng_Cols - lng_Rows - lng_DiffOutputAndInputRow
        lng_ColCounter = lng_Rows - lng_Cols - lng_DiffOutputAndInputCol
        
        arr_Output(lng_Rows, lng_Cols) = "=R[" & lng_RowCounter & "]C[" & lng_ColCounter & "]"
    
    Next
Next
'------------------------------------------------------------------------------
rng_Output.Resize(UBound(arr_Output, 1), UBound(arr_Output, 2)) = arr_Output

End Sub
Sub MtrxInverse()
Dim rng_Default, rng_Mtrx, rng_Output, rng_DefaultOutput As Range
Dim str_OutputAddress As String
Dim lng_OutputRefRow, lng_OutputRefCol As Long

Set rng_Default = Selection
'-----------------------------------------------------------------
'Enable error handling
On Error Resume Next
Set rng_Mtrx = Application.InputBox( _
        Prompt:="Select matrix 1", _
        Title:="Select range", _
        Default:=rng_Default.Address, _
        Type:=8)
'User cancel scenario
If rng_Mtrx Is Nothing Then
    MsgBox "Matrix is empty. Program will end.", vbOKOnly, "Error!"
    Exit Sub
End If
'Disable error-handling
On Error GoTo 0
'-----------------------------------------------------------------
If rng_Mtrx.Rows.count <> rng_Mtrx.Columns.count Then
    MsgBox "Matrix must be square. Program will end.", vbOKOnly, "Error!"
    Exit Sub
End If
'---set default output range-------------------------------------
lng_OutputRefRow = rng_Mtrx.row
lng_OutputRefCol = rng_Mtrx.Column + rng_Mtrx.Columns.count + 1
Set rng_DefaultOutput = Cells(lng_OutputRefRow, lng_OutputRefCol)
'---capture output range-------------------------------------------
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
'---output results------------------------------------------------------

Set rng_Output = rng_Output.Resize(rng_Mtrx.Rows.count, rng_Mtrx.Columns.count)
str_OutputAddress = rng_Mtrx.Address(False, False)
rng_Output.FormulaArray = "=MINVERSE(" & str_OutputAddress & ")"
End Sub

Sub MtrxDotProduct()
'calculate the dot product of two matrices
Dim rng_Default, rng_Mtrx1, rng_Mtrx2, rng_Output, rng_DefaultOutput As Range
Dim lng_OutputRefRow, lng_OutputRefCol As Long
Dim arr_One, arr_Two As Variant
Dim str_OutputSheetName As String

Set rng_Default = Selection
'-----------------------------------------------------------------
'Enable error handling
On Error Resume Next
Set rng_Mtrx1 = Application.InputBox( _
        Prompt:="Select matrix 1", _
        Title:="Select range", _
        Default:=rng_Default.Address, _
        Type:=8)
'User cancel scenario
If rng_Mtrx1 Is Nothing Then
    MsgBox "Matrix 1 is empty", vbOKOnly, "Error!"
    Exit Sub
End If
'Disable error-handling
On Error GoTo 0
'-----------------------------------------------------------------
'Enable error handling
On Error Resume Next
Set rng_Mtrx2 = Application.InputBox( _
        Prompt:="Select matrix 2", _
        Title:="Select range", _
        Default:=rng_Default.Address, _
        Type:=8)

'User cancel scenario
If rng_Mtrx2 Is Nothing Then
    MsgBox "Matrix 2 is empty", vbOKOnly, "Error!"
    Exit Sub
End If
'Disable error-handling
On Error GoTo 0
'---check matrices are the same sizes------------------------------
If rng_Mtrx1.Rows.count = rng_Mtrx2.Rows.count And rng_Mtrx1.Columns.count = rng_Mtrx2.Columns.count Then
    '---set default output range-------------------------------------
    lng_OutputRefRow = rng_Mtrx2.row
    lng_OutputRefCol = rng_Mtrx2.Column + rng_Mtrx2.Columns.count + 1
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
    
    str_OutputSheetName = ActiveSheet.Name
    
    arr_One = rng_Mtrx1.value
    arr_Two = rng_Mtrx2.value
    
    With Sheets(str_OutputSheetName).Range(rng_Output.Address)
        .value = MatrixDotProduct(arr_One, arr_Two)
    End With
        Else
            MsgBox "Matrices MUST have identical number of rows and columns." & vbCrLf & "Ending procedure.", vbOKOnly, "Error!"
            Exit Sub
End If

End Sub
Private Function MatrixDotProduct(arr_One As Variant, arr_Two As Variant) As Variant
'returns the dot product of two matrices
Dim lng_Row, lng_Col As Long

MatrixDotProduct = 0

Application.Volatile
'---validation--------------------------------------------------------------
'check matrices areof same size
If UBound(arr_One, 1) - LBound(arr_One, 1) = UBound(arr_Two, 1) - LBound(arr_Two, 1) _
 And UBound(arr_One, 2) - LBound(arr_One, 2) = UBound(arr_Two, 2) - LBound(arr_Two, 2) Then
 
    For lng_Row = LBound(arr_One, 1) To UBound(arr_One, 1) Step 1
        For lng_Col = LBound(arr_One, 2) To UBound(arr_One, 2) Step 1
            MatrixDotProduct = arr_One(lng_Row, lng_Col) * arr_Two(lng_Row, lng_Col) + MatrixDotProduct
        Next lng_Col
    Next lng_Row
    
End If

End Function

Sub TransposeSquareMatrix()
Dim rng_Default, rng_DefaultOutput, rng_Mtrx, rng_Output As Range
Dim lng_Rows, lng_Cols, lng_OutputRefRow, lng_OutputRefCol As Long
Dim arr_Temp, arr_Output As Variant

Set rng_Default = Selection
'-----------------------------------------------------------------
'Enable error handling
On Error Resume Next
Set rng_Mtrx = Application.InputBox( _
        Prompt:="Select square matrix to transpose", _
        Title:="Select range", _
        Default:=rng_Default.Address, _
        Type:=8)
'User cancel scenario
If rng_Mtrx Is Nothing Then
    MsgBox "Matrix is empty", vbOKOnly, "Error!"
    Exit Sub
End If
'Disable error-handling
On Error GoTo 0
'---------------------------------------------------------------------------
If rng_Mtrx.Rows.count <> rng_Mtrx.Columns.count Then
    MsgBox "Columns of matrix 1 <> rows of matrix 2.", vbOKOnly, "Error!"
    Exit Sub
End If
'------------------------------------------------------------------------------
lng_OutputRefRow = rng_Mtrx.row
lng_OutputRefCol = rng_Mtrx.Column + rng_Mtrx.Columns.count + 1
Set rng_DefaultOutput = ActiveSheet.Cells(lng_OutputRefRow, lng_OutputRefCol)
'------------------------------------------------------------------------------
ReDim arr_Temp(1 To rng_Mtrx.Rows.count, 1 To rng_Mtrx.Columns.count)
ReDim arr_Output(1 To UBound(arr_Temp, 1), 1 To UBound(arr_Temp, 2))

arr_Temp = rng_Mtrx

For lng_Cols = 1 To rng_Mtrx.Columns.count Step 1
    For lng_Rows = 1 To rng_Mtrx.Rows.count Step 1
        arr_Output(lng_Rows, lng_Cols) = arr_Temp(lng_Cols, lng_Rows)
    Next
Next
'------------------------------------------------------------------------------
On Error Resume Next
Set rng_Output = Application.InputBox( _
        Prompt:="Select output location", _
        Title:="Select range", _
        Default:=rng_DefaultOutput.Address, _
        Type:=8)
'User cancel scenario
If rng_Output.Cells.count <> 1 Then
    MsgBox "Select a single cell where output starts", vbOKOnly, "Error!"
    Exit Sub
End If
'Disable error-handling
On Error GoTo 0

rng_Output.Resize(UBound(arr_Output, 1), UBound(arr_Output, 2)) = arr_Output

End Sub

Sub TransposeAnyMatrix()
Dim rng_Default, rng_DefaultOutput, rng_Mtrx, rng_Output As Range
Dim lng_Rows, lng_Cols, lng_OutputRefRow, lng_OutputRefCol As Long
Dim arr_Temp, arr_Output As Variant

Set rng_Default = Selection
'-----------------------------------------------------------------
'Enable error handling
On Error Resume Next
Set rng_Mtrx = Application.InputBox( _
        Prompt:="Select matrix to transpose", _
        Title:="Select range", _
        Default:=rng_Default.Address, _
        Type:=8)
'User cancel scenario
If rng_Mtrx Is Nothing Then
    MsgBox "Matrix is empty", vbOKOnly, "Error!"
    Exit Sub
End If
'Disable error-handling
On Error GoTo 0
'------------------------------------------------------------------------------
lng_OutputRefRow = rng_Mtrx.row
lng_OutputRefCol = rng_Mtrx.Column + rng_Mtrx.Columns.count + 1
Set rng_DefaultOutput = ActiveSheet.Cells(lng_OutputRefRow, lng_OutputRefCol)
'------------------------------------------------------------------------------
ReDim arr_Temp(1 To rng_Mtrx.Rows.count, 1 To rng_Mtrx.Columns.count)
ReDim arr_Output(1 To UBound(arr_Temp, 2), 1 To UBound(arr_Temp, 1))
'---readthe range to an array-------------------------------------------------
arr_Temp = rng_Mtrx.value
'---transpose the array ad write it to a new array-----------------------------
arr_Output = Application.Transpose(arr_Temp)
'---set the output location-------------------------------------------------------
On Error Resume Next
Set rng_Output = Application.InputBox( _
        Prompt:="Select output location", _
        Title:="Select range", _
        Default:=rng_DefaultOutput.Address, _
        Type:=8)
'User cancel scenario
If rng_Output.Cells.count <> 1 Then
    MsgBox "Select a single cell where output starts", vbOKOnly, "Error!"
    Exit Sub
End If
'Disable error-handling
On Error GoTo 0

rng_Output.Resize(UBound(arr_Output, 1), UBound(arr_Output, 2)) = arr_Output

End Sub
Sub MtrxScaling()
Dim rng_Default, rng_Output As Range
Dim arr_Scaling(1 To 5, 1 To 3) As Variant

Set rng_Default = Selection
'-----------------------------------------------------------------
'Enable error handling
On Error Resume Next
Set rng_Output = Application.InputBox( _
        Prompt:="Select output location", _
        Title:="Select range", _
        Default:=rng_Default.Address, _
        Type:=8)
'User cancel scenario
If rng_Output Is Nothing Then
    MsgBox "Output location is empty", vbOKOnly, "Error!"
    Exit Sub
End If
'Disable error-handling
On Error GoTo 0
'-----------------------------------------------------------------
arr_Scaling(1, 1) = 2
arr_Scaling(2, 1) = 0
arr_Scaling(3, 1) = 0
arr_Scaling(5, 1) = "Scaling"
arr_Scaling(1, 2) = 0
arr_Scaling(2, 2) = 2
arr_Scaling(3, 2) = 0
arr_Scaling(1, 3) = 0
arr_Scaling(2, 3) = 0
arr_Scaling(3, 3) = 2

rng_Output.Resize(UBound(arr_Scaling, 1), UBound(arr_Scaling, 2)) = arr_Scaling

End Sub
Sub MtrxStretch()
Dim rng_Default, rng_Output As Range
Dim arr_Stretch(1 To 5, 1 To 19) As Variant

Set rng_Default = Selection
'-----------------------------------------------------------------
'Enable error handling
On Error Resume Next
Set rng_Output = Application.InputBox( _
        Prompt:="Select output location", _
        Title:="Select range", _
        Default:=rng_Default.Address, _
        Type:=8)
'User cancel scenario
If rng_Output Is Nothing Then
    MsgBox "Output location is empty", vbOKOnly, "Error!"
    Exit Sub
End If
'Disable error-handling
On Error GoTo 0
'-----------------------------------------------------------------
arr_Stretch(1, 1) = "=R[4]C[1]"
arr_Stretch(2, 1) = 0
arr_Stretch(3, 1) = 0
arr_Stretch(5, 1) = "x-stretch"
arr_Stretch(1, 2) = 0
arr_Stretch(2, 2) = 1
arr_Stretch(3, 2) = 0
arr_Stretch(5, 2) = 2
arr_Stretch(1, 3) = 0
arr_Stretch(2, 3) = 0
arr_Stretch(3, 3) = 1
arr_Stretch(5, 3) = ""

arr_Stretch(1, 5) = 1
arr_Stretch(2, 5) = 0
arr_Stretch(3, 5) = 0
arr_Stretch(5, 5) = "Y-stretch"
arr_Stretch(1, 6) = 0
arr_Stretch(2, 6) = "=R[3]C"
arr_Stretch(3, 6) = 0
arr_Stretch(5, 6) = 2
arr_Stretch(1, 7) = 0
arr_Stretch(2, 7) = 0
arr_Stretch(3, 7) = 1
arr_Stretch(5, 7) = ""

arr_Stretch(5, 9) = "xy-stretch"

arr_Stretch(1, 13) = 1
arr_Stretch(2, 13) = 0
arr_Stretch(3, 13) = 0
arr_Stretch(5, 13) = "z-stretch"
arr_Stretch(1, 14) = 0
arr_Stretch(2, 14) = 1
arr_Stretch(3, 14) = 0
arr_Stretch(5, 14) = 2
arr_Stretch(1, 15) = 0
arr_Stretch(2, 15) = 0
arr_Stretch(3, 15) = "=R[2]C[-1]"
arr_Stretch(5, 15) = ""

arr_Stretch(5, 17) = "xyz-stretch"

rng_Output.Resize(UBound(arr_Stretch, 1), UBound(arr_Stretch, 2)) = arr_Stretch

End Sub
Sub MtrxRotX()
Dim rng_Default, rng_Mtrx As Range
Dim arr_Matrix(1 To 5, 1 To 3) As Variant

Set rng_Default = Selection
'-----------------------------------------------------------------
'Enable error handling
On Error Resume Next
Set rng_Mtrx = Application.InputBox( _
        Prompt:="Select output location", _
        Title:="Select range", _
        Default:=rng_Default.Address, _
        Type:=8)
'User cancel scenario
If rng_Mtrx Is Nothing Then
    MsgBox "Matrix is empty", vbOKOnly, "Error!"
    Exit Sub
End If
'Disable error-handling
On Error GoTo 0
'-----------------------------------------------------------------
arr_Matrix(1, 1) = 1
arr_Matrix(1, 2) = 0
arr_Matrix(1, 3) = 0
arr_Matrix(2, 1) = 0
arr_Matrix(2, 2) = "=COS(R[3]C*PI()/180)"
arr_Matrix(2, 3) = "=-SIN(R[3]C[-1]*PI()/180)"
arr_Matrix(3, 1) = 0
arr_Matrix(3, 2) = "=SIN(R[2]C*PI()/180)"
arr_Matrix(3, 3) = "=-COS(R[2]C[-1]*PI()/180)"
arr_Matrix(5, 1) = "Rot X"
arr_Matrix(5, 2) = 90
arr_Matrix(5, 3) = 0

Set rng_Mtrx = rng_Mtrx.Resize(UBound(arr_Matrix, 1), UBound(arr_Matrix, 2))
rng_Mtrx = arr_Matrix
End Sub
Sub MtrxRotY()
Dim rng_Default, rng_Mtrx As Range
Dim arr_Matrix(1 To 5, 1 To 3) As Variant

Set rng_Default = Selection
'-----------------------------------------------------------------
'Enable error handling
On Error Resume Next
Set rng_Mtrx = Application.InputBox( _
        Prompt:="Select output location", _
        Title:="Select range", _
        Default:=rng_Default.Address, _
        Type:=8)
'User cancel scenario
If rng_Mtrx Is Nothing Then
    MsgBox "Matrix is empty", vbOKOnly, "Error!"
    Exit Sub
End If
'Disable error-handling
On Error GoTo 0
'-----------------------------------------------------------------
arr_Matrix(1, 1) = "=COS(R[4]C[1]*PI()/180)"
arr_Matrix(1, 2) = 0
arr_Matrix(1, 3) = "=SIN(R[4]C[-1]*PI()/180)"
arr_Matrix(2, 1) = 0
arr_Matrix(2, 2) = 1
arr_Matrix(2, 3) = 0
arr_Matrix(3, 1) = "=-SIN(R[2]C[1]*PI()/180)"
arr_Matrix(3, 2) = 0
arr_Matrix(3, 3) = "=COS(R[2]C[-1]*PI()/180)"
arr_Matrix(5, 1) = "Rot Y"
arr_Matrix(5, 2) = 90
arr_Matrix(5, 3) = 0

Set rng_Mtrx = rng_Mtrx.Resize(UBound(arr_Matrix, 1), UBound(arr_Matrix, 2))
rng_Mtrx = arr_Matrix
End Sub
Sub MtrxRotZ()
Dim rng_Default, rng_Mtrx As Range
Dim arr_Matrix(1 To 5, 1 To 3) As Variant

Set rng_Default = Selection
'-----------------------------------------------------------------
'Enable error handling
On Error Resume Next
Set rng_Mtrx = Application.InputBox( _
        Prompt:="Select output location", _
        Title:="Select range", _
        Default:=rng_Default.Address, _
        Type:=8)
'User cancel scenario
If rng_Mtrx Is Nothing Then
    MsgBox "Matrix is empty", vbOKOnly, "Error!"
    Exit Sub
End If
'Disable error-handling
On Error GoTo 0
'-----------------------------------------------------------------
arr_Matrix(1, 1) = "=COS(R[4]C[1]*PI()/180)"
arr_Matrix(1, 2) = "=-SIN(R[4]C*PI()/180)"
arr_Matrix(1, 3) = 0
arr_Matrix(2, 1) = "=SIN(R[3]C[1]*PI()/180)"
arr_Matrix(2, 2) = "=COS(R[3]C*PI()/180)"
arr_Matrix(2, 3) = 0
arr_Matrix(3, 1) = 0
arr_Matrix(3, 2) = 0
arr_Matrix(3, 3) = 1
arr_Matrix(5, 1) = "Rot Z"
arr_Matrix(5, 2) = 90
arr_Matrix(5, 3) = 0

Set rng_Mtrx = rng_Mtrx.Resize(UBound(arr_Matrix, 1), UBound(arr_Matrix, 2))
rng_Mtrx = arr_Matrix
End Sub
Sub MtrxRotAll()
Dim rng_Default, rng_Mtrx As Range
Dim arr_Matrix(1 To 5, 1 To 27) As Variant

Set rng_Default = Selection
'-----------------------------------------------------------------
'Enable error handling
On Error Resume Next
Set rng_Mtrx = Application.InputBox( _
        Prompt:="Select output location", _
        Title:="Select range", _
        Default:=rng_Default.Address, _
        Type:=8)
'User cancel scenario
If rng_Mtrx Is Nothing Then
    MsgBox "Matrix is empty", vbOKOnly, "Error!"
    Exit Sub
End If
'Disable error-handling
On Error GoTo 0
'---Rotation x--------------------------------------------------------------
arr_Matrix(1, 1) = 1
arr_Matrix(1, 2) = 0
arr_Matrix(1, 3) = 0
arr_Matrix(2, 1) = 0
arr_Matrix(2, 2) = "=COS(R[3]C*PI()/180)"
arr_Matrix(2, 3) = "=-SIN(R[3]C[-1]*PI()/180)"
arr_Matrix(3, 1) = 0
arr_Matrix(3, 2) = "=SIN(R[2]C*PI()/180)"
arr_Matrix(3, 3) = "=-COS(R[2]C[-1]*PI()/180)"
arr_Matrix(5, 1) = "Rot X"
arr_Matrix(5, 2) = 90
arr_Matrix(5, 3) = 0
'---Rotation y--------------------------------------------------------------
arr_Matrix(1, 5) = "=COS(R[4]C[1]*PI()/180)"
arr_Matrix(1, 6) = 0
arr_Matrix(1, 7) = "=SIN(R[4]C[-1]*PI()/180)"
arr_Matrix(2, 5) = 0
arr_Matrix(2, 6) = 1
arr_Matrix(2, 7) = 0
arr_Matrix(3, 5) = "=-SIN(R[2]C[1]*PI()/180)"
arr_Matrix(3, 6) = 0
arr_Matrix(3, 7) = "=COS(R[2]C[-1]*PI()/180)"
arr_Matrix(5, 5) = "Rot Y"
arr_Matrix(5, 6) = 90
arr_Matrix(5, 7) = 0
'---Rotation z--------------------------------------------------------------
arr_Matrix(1, 13) = "=COS(R[4]C[1]*PI()/180)"
arr_Matrix(1, 14) = "=-SIN(R[4]C*PI()/180)"
arr_Matrix(1, 15) = 0
arr_Matrix(2, 13) = "=SIN(R[3]C[1]*PI()/180)"
arr_Matrix(2, 14) = "=COS(R[3]C*PI()/180)"
arr_Matrix(2, 15) = 0
arr_Matrix(3, 13) = 0
arr_Matrix(3, 14) = 0
arr_Matrix(3, 15) = 1
arr_Matrix(5, 13) = "Rot Z"
arr_Matrix(5, 14) = 90
arr_Matrix(5, 15) = 0

Set rng_Mtrx = rng_Mtrx.Resize(UBound(arr_Matrix, 1), UBound(arr_Matrix, 2))
rng_Mtrx = arr_Matrix
End Sub

Sub MtrxTriangularUpper()
Dim arr_Mtrx As Variant
Dim lng_Row, lng_Col, lng_MtrxSize As Long
Dim rng_Default, rng_Output As Range

'---capture size of matrix---------------------------------------
lng_MtrxSize = Application.InputBox( _
        Prompt:="How many rows / columns do you want", _
        Title:="Enter the number of rows / columns ", _
        Default:=4, _
        Type:=1)
'Error checking to capture divide by zero
If lng_MtrxSize = 0 Then 'Or lng_MatrixSize = "" Then
    MsgBox "Error! Number of rows/columns cannot be zero or empty. I'll use 4 instead."
    lng_MtrxSize = 4
End If
'---Populate the array-------------------------------------------
ReDim arr_Mtrx(1 To lng_MtrxSize, 1 To lng_MtrxSize)
For lng_Col = 1 To lng_MtrxSize Step 1
    For lng_Row = 1 To lng_MtrxSize Step 1
        If lng_Row < lng_Col Then
            arr_Mtrx(lng_Row, lng_Col) = 0
            Else
                arr_Mtrx(lng_Row, lng_Col) = WorksheetFunction.RandBetween(0, 100)
        End If
    Next
Next
'---Capture output location--------------------------------------
Set rng_Default = Selection
'Enable error handling
On Error Resume Next
Set rng_Output = Application.InputBox( _
        Prompt:="Select output location", _
        Title:="Select range", _
        Default:=rng_Default.Address, _
        Type:=8)
'User cancel scenario
If rng_Output Is Nothing Then
    MsgBox "Matrix 1 is empty", vbOKOnly, "Error!"
    Exit Sub
End If
'Disable error-handling
On Error GoTo 0
'---Output the array---------------------------------------------
Set rng_Output = rng_Output.Resize(UBound(arr_Mtrx, 1), UBound(arr_Mtrx, 2))
rng_Output = arr_Mtrx
End Sub

Sub MtrxTriangularLower()
Dim arr_Mtrx As Variant
Dim lng_Row, lng_Col, lng_MtrxSize As Long
Dim rng_Default, rng_Output As Range

'---capture size of matrix---------------------------------------
lng_MtrxSize = Application.InputBox( _
        Prompt:="How many rows / columns do you want", _
        Title:="Enter the number of rows / columns ", _
        Default:=4, _
        Type:=1)
'Error checking to capture divide by zero
If lng_MtrxSize = 0 Then 'Or lng_MatrixSize = "" Then
    MsgBox "Error! Number of rows/columns cannot be zero or empty. I'll use 4 instead."
    lng_MtrxSize = 4
End If
'---Populate the array-------------------------------------------
ReDim arr_Mtrx(1 To lng_MtrxSize, 1 To lng_MtrxSize)
For lng_Col = 1 To lng_MtrxSize Step 1
    For lng_Row = 1 To lng_MtrxSize Step 1
        If lng_Row > lng_Col Then
            arr_Mtrx(lng_Row, lng_Col) = 0
            Else
                arr_Mtrx(lng_Row, lng_Col) = WorksheetFunction.RandBetween(0, 100)
        End If
    Next
Next
'---Capture output location--------------------------------------
Set rng_Default = Selection
'Enable error handling
On Error Resume Next
Set rng_Output = Application.InputBox( _
        Prompt:="Select output location", _
        Title:="Select range", _
        Default:=rng_Default.Address, _
        Type:=8)
'User cancel scenario
If rng_Output Is Nothing Then
    MsgBox "Matrix 1 is empty", vbOKOnly, "Error!"
    Exit Sub
End If
'Disable error-handling
On Error GoTo 0
'---Output the array---------------------------------------------
Set rng_Output = rng_Output.Resize(UBound(arr_Mtrx, 1), UBound(arr_Mtrx, 2))
rng_Output = arr_Mtrx
End Sub

Sub CreateDiagonalMatrix()
'creates a diagonal matrix from a gven row or column vector
Dim rng_input As Range
Dim rng_Output As Range
Dim arr_DiagonalMatrix, arr_Input As Variant
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Set rng_input = GetRange(Selection, "Input row or column vector to be diagonalised.")
'---Validate range----------------------------------------------------
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'Validate all are true by multiplying and checking = 1
If rng_input.Cells.count = 1 Then
    MsgBox "Input is not valid. You selected a single cell. Ending procedure.", vbOKOnly + vbExclamation, "Error"
    Exit Sub
End If
'validate input array is numeric
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
If IsRangeNumeric(rng_input) = False Then
    MsgBox "Input range must either be numeric. Ending procedure.", vbOKOnly + vbExclamation, "Error!"
    Exit Sub
End If
'---assign range to an array------------------------------------------------------------------
arr_Input = rng_input
'---Validate array----------------------------------------------------
'validate the input array is either a row or column array. If not report error and end procedure
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
If IsRowVector(arr_Input) = False And IsColumnVector(arr_Input) = False Then
    MsgBox "Input range must either be a single row or single column. Ending procedure.", vbOKOnly + vbExclamation, "Error!"
    Exit Sub
End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
arr_DiagonalMatrix = GenerateDiagonalArray(arr_Input)
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'default range is shited 2 columns to the right
Set rng_Output = GetRange(Selection.Cells(1, 1).Offset(0, 2), "Select start of output range.")
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Call WriteArrayToWorksheetA1(arr_DiagonalMatrix, ActiveSheet.Name, rng_Output.Cells(1, 1).Address)

End Sub

