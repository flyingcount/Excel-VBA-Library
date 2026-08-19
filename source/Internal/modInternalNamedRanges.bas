Attribute VB_Name = "modInternalNamedRanges"
Option Explicit

' Internal: named ranges used by Api/feature modules
' (Personal Custom_Menu14_* plus Benford / plots sheet names).
' Called from modApiRanges and other Internal modules. Not the external API.

Public Const RangePropertiesSheetName As String = "Range properties"
Private Const HighlightColorIndex As Long = 36
Private Const PropColCount As Long = 8

Public Function CheckNamedRangeExists(ByVal RangeName As String) As Boolean
    Dim nm As Name
    CheckNamedRangeExists = False
    For Each nm In ActiveWorkbook.Names
        If nm.Name = RangeName Then
            CheckNamedRangeExists = True
            Exit Function
        End If
    Next nm
End Function

Public Sub CreateNamedRange(ByVal RangeName As String, ByVal RangeToName As Range)
    If CheckNamedRangeExists(RangeName) Then ActiveWorkbook.Names(RangeName).Delete
    ActiveWorkbook.Names.Add Name:=RangeName, RefersTo:=RangeToName, Visible:=True
    ActiveWorkbook.Names(RangeName).Comment = "Created " & Now()
End Sub

' Sheet-scoped name so Residuals / Order do not collide with other sheets.
Public Sub CreateSheetNamedRange(ByVal ws As Worksheet, ByVal RangeName As String, ByVal Target As Range, Optional ByVal makeVisible As Boolean = True, Optional ByVal Comment As String = "")
    On Error Resume Next
    ws.Names(RangeName).Delete
    On Error GoTo 0
    ws.Names.Add Name:=RangeName, RefersTo:=Target, Visible:=makeVisible
    If Len(Comment) > 0 Then ws.Names(RangeName).Comment = Comment
End Sub

' Excel name rules: letters, digits, underscore, period; not a cell address; max 255.
Public Function ValidRangeName(ByVal raw As String) As String
    Dim i As Long
    Dim code As Long
    Dim ch As String
    Dim cleaned As String
    Dim probe As Range

    cleaned = Replace(Trim$(raw), " ", ".")
    For i = 1 To Len(cleaned)
        ch = Mid$(cleaned, i, 1)
        code = Asc(ch)
        Select Case code
            Case 46, 48 To 57, 65 To 90, 95, 97 To 122
                ValidRangeName = ValidRangeName & ch
        End Select
    Next i

    If Len(ValidRangeName) = 0 Then
        ValidRangeName = "rng_Default"
        Exit Function
    End If
    If IsNumeric(Left$(ValidRangeName, 1)) Then ValidRangeName = "_" & ValidRangeName

    On Error Resume Next
    Set probe = ActiveSheet.Range(ValidRangeName)
    On Error GoTo 0
    If Not probe Is Nothing Then ValidRangeName = "rng_" & ValidRangeName

    ValidRangeName = Left$(ValidRangeName, 255)
End Function

Public Function LocalNameOf(ByVal nm As Name) As String
    Dim p As Long
    p = InStrRev(nm.Name, "!")
    If p > 0 Then
        LocalNameOf = Mid$(nm.Name, p + 1)
    Else
        LocalNameOf = nm.Name
    End If
End Function

' Workbook-scoped name first, then sheet-scoped on ws (active sheet if omitted).
Public Function FindName(ByVal wb As Workbook, ByVal LocalName As String, Optional ByVal ws As Worksheet) As Name
    Dim nm As Name
    Dim want As String
    want = Trim$(CStr(LocalName))
    If Len(want) = 0 Then Exit Function
    If ws Is Nothing Then Set ws = wb.ActiveSheet

    On Error Resume Next
    Set nm = wb.Names(want)
    On Error GoTo 0
    If Not nm Is Nothing Then
        If TypeOf nm.Parent Is Workbook Then
            Set FindName = nm
            Exit Function
        End If
    End If

    On Error Resume Next
    Set nm = Nothing
    Set nm = ws.Names(want)
    On Error GoTo 0
    If Not nm Is Nothing Then
        Set FindName = nm
        Exit Function
    End If

    For Each nm In wb.Names
        If StrComp(LocalNameOf(nm), want, vbTextCompare) = 0 Then
            Set FindName = nm
            Exit Function
        End If
    Next nm
End Function

Public Sub WriteNamedRangeInventory(ByVal wb As Workbook)
    Dim n As Long
    Dim arr As Variant
    Dim ws As Worksheet
    Dim nm As Name
    Dim r As Long
    Dim refers As Range

    n = wb.Names.Count
    If n = 0 Then
        MsgBox "There are no named ranges in " & wb.Name & ".", vbInformation, "Ranges"
        Exit Sub
    End If

    ReDim arr(1 To n + 1, 1 To PropColCount)
    arr(1, 1) = "Index"
    arr(1, 2) = "Name"
    arr(1, 3) = "Refers to"
    arr(1, 4) = "Address"
    arr(1, 5) = "Visible"
    arr(1, 6) = "Comment"
    arr(1, 7) = "Workbook parameter"
    arr(1, 8) = "Scope"

    r = 2
    For Each nm In wb.Names
        arr(r, 1) = r - 1
        arr(r, 2) = nm.Name
        arr(r, 3) = SafeNameValue(nm)
        Set refers = SafeRefersToRange(nm)
        If refers Is Nothing Then
            arr(r, 4) = ""
        Else
            arr(r, 4) = refers.Address(External:=False)
        End If
        arr(r, 5) = nm.Visible
        arr(r, 6) = SafeComment(nm)
        arr(r, 7) = SafeWorkbookParameter(nm)
        arr(r, 8) = ScopeLabel(nm)
        r = r + 1
    Next nm

    Call modApiSheets.CreateOutputSheet(RangePropertiesSheetName)
    Set ws = wb.Worksheets(RangePropertiesSheetName)
    ws.Range("A1").Value = "Named ranges in " & wb.Name & " (" & CStr(n) & ")"
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "Last updated " & Now
    ws.Range("A3").Resize(n + 1, PropColCount).Value = arr
    ws.Range("A3").Resize(1, PropColCount).Font.Bold = True
    ws.Columns("A").ColumnWidth = 8
    ws.Columns("B:H").AutoFit
    If ws.Columns("C").ColumnWidth > 50 Then ws.Columns("C").ColumnWidth = 50
    If ws.Columns("F").ColumnWidth > 40 Then ws.Columns("F").ColumnWidth = 40
    ws.Activate
    On Error Resume Next
    ActiveWindow.FreezePanes = False
    ActiveWindow.SplitColumn = 0
    ActiveWindow.SplitRow = 3
    ActiveWindow.FreezePanes = True
    On Error GoTo 0
End Sub

Public Sub SetAllNamesVisible(ByVal wb As Workbook, ByVal makeVisible As Boolean)
    Dim nm As Name
    For Each nm In wb.Names
        nm.Visible = makeVisible
    Next nm
End Sub

Public Sub SetSheetNamesVisible(ByVal ws As Worksheet, ByVal makeVisible As Boolean)
    Dim nm As Name
    For Each nm In ws.Names
        If TypeOf nm.Parent Is Worksheet Then nm.Visible = makeVisible
    Next nm
End Sub

' Returns a warning string for names that were not found (empty if all applied).
Public Function SetSpecifiedNamesVisible(ByVal wb As Workbook, ByVal nameCells As Range, ByVal makeVisible As Boolean) As String
    Dim cell As Range
    Dim nm As Name
    Dim missing As String
    Dim label As String
    For Each cell In nameCells.Cells
        label = Trim$(CStr(cell.Value))
        If Len(label) = 0 Then GoTo NextCell
        Set nm = FindName(wb, label)
        If nm Is Nothing Then
            missing = missing & "Name " & label & " does not exist" & vbCrLf
        Else
            nm.Visible = makeVisible
        End If
NextCell:
    Next cell
    SetSpecifiedNamesVisible = missing
End Function

Public Function DeleteSpecifiedNames(ByVal wb As Workbook, ByVal nameCells As Range) As String
    Dim cell As Range
    Dim nm As Name
    Dim missing As String
    Dim label As String
    For Each cell In nameCells.Cells
        label = Trim$(CStr(cell.Value))
        If Len(label) = 0 Then GoTo NextCell
        Set nm = FindName(wb, label)
        If nm Is Nothing Then
            missing = missing & "Name " & label & " does not exist" & vbCrLf
        Else
            nm.Delete
        End If
NextCell:
    Next cell
    DeleteSpecifiedNames = missing
End Function

Public Sub CreateHiddenRangeName(ByVal wb As Workbook, ByVal RangeName As String, ByVal Target As Range)
    Dim n As String
    n = ValidRangeName(RangeName)
    DeleteNameIfExists wb, n
    wb.Names.Add Name:=n, RefersTo:=Target, Visible:=False
    wb.Names(n).Comment = "Created " & Now()
End Sub

Public Sub CreateHiddenConstant(ByVal wb As Workbook, ByVal RangeName As String, ByVal Value As Variant, Optional ByVal ws As Worksheet)
    Dim n As String
    Dim refers As String
    n = ValidRangeName(RangeName)
    refers = RefersToFromValue(Value)
    If ws Is Nothing Then
        DeleteNameIfExists wb, n
        wb.Names.Add Name:=n, RefersTo:=refers, Visible:=False
        wb.Names(n).Comment = "Constant  Created " & Now()
    Else
        On Error Resume Next
        ws.Names(n).Delete
        On Error GoTo 0
        ws.Names.Add Name:=n, RefersTo:=refers, Visible:=False
        ws.Names(n).Comment = "Constant  Created " & Now()
    End If
End Sub

Public Sub CreateCommonConstantNames(ByVal wb As Workbook)
    Call CreateHiddenConstant(wb, "Divisor_Thousand", 1000)
    wb.Names("Divisor_Thousand").Visible = True
    Call CreateHiddenConstant(wb, "Divisor_Millions", 1000000)
    wb.Names("Divisor_Millions").Visible = True
End Sub

' Two-column block: name in column 1, value in column 2. Hidden constants.
Public Sub CreateHiddenConstantsFromRows(ByVal wb As Workbook, ByVal block As Range, Optional ByVal ws As Worksheet)
    Dim r As Range
    Dim nm As String
    If block.Columns.Count < 2 Then Err.Raise vbObjectError + 1401, "CreateHiddenConstantsFromRows", "Select two columns: names then values."
    For Each r In block.Rows
        nm = Trim$(CStr(r.Cells(1, 1).Value))
        If Len(nm) > 0 Then Call CreateHiddenConstant(wb, nm, r.Cells(1, 2).Value, ws)
    Next r
End Sub

' Header row names each column body (header excluded). Sheet-scoped, hidden.
Public Sub NameEachColumnFromHeader(ByVal headerBlock As Range)
    Dim col As Range
    Dim nRows As Long
    Dim nm As String
    Dim body As Range
    nRows = headerBlock.Rows.Count
    If nRows < 2 Then Err.Raise vbObjectError + 1402, "NameEachColumnFromHeader", "Need a header row plus at least one data row."
    For Each col In headerBlock.Columns
        If IsEmpty(col.Cells(1, 1).Value) Then Err.Raise vbObjectError + 1403, "NameEachColumnFromHeader", "The first row cannot contain empty cells."
        nm = ValidRangeName(CStr(col.Cells(1, 1).Value))
        Set body = col.Offset(1, 0).Resize(nRows - 1, 1)
        Call CreateSheetNamedRange(headerBlock.Worksheet, nm, body, False, "Created " & Now())
    Next col
End Sub

' Same local name on every worksheet, referring to the same address on that sheet.
Public Sub CreateLocalNameOnEverySheet(ByVal wb As Workbook, ByVal RangeName As String, ByVal Template As Range)
    Dim ws As Worksheet
    Dim n As String
    Dim addr As String
    n = ValidRangeName(RangeName)
    addr = Template.Address(External:=False)
    For Each ws In wb.Worksheets
        On Error Resume Next
        ws.Names(n).Delete
        On Error GoTo 0
        ws.Names.Add Name:=n, RefersTo:=ws.Range(addr), Visible:=True
        ws.Names(n).Comment = "Created " & Now()
    Next ws
End Sub

' Workbook-scoped names OutputAll and Output_1..n from each column. Optional header skip.
Public Sub SplitColumnsToWorkbookNames(ByVal src As Range, ByVal BaseName As String, ByVal HasHeader As Boolean)
    Dim col As Range
    Dim i As Long
    Dim n As String
    Dim body As Range
    Dim nRows As Long
    n = ValidRangeName(BaseName)
    nRows = src.Rows.Count
    If HasHeader Then
        If nRows < 2 Then Err.Raise vbObjectError + 1404, "SplitColumnsToWorkbookNames", "A header row needs at least one data row."
    End If
    Call CreateNamedRange(n & "All", src)
    i = 0
    For Each col In src.Columns
        i = i + 1
        If HasHeader Then
            Set body = col.Offset(1, 0).Resize(nRows - 1, 1)
        Else
            Set body = col
        End If
        Call CreateNamedRange(n & "_" & CStr(i), body)
    Next col
End Sub

Public Function ConvertLocalNamesToGlobal(ByVal wb As Workbook, ByVal nameCells As Range) As String
    Dim cell As Range
    Dim nm As Name
    Dim missing As String
    Dim localNm As String
    Dim refers As String
    Dim vis As Boolean
    Dim cmt As String
    Dim label As String
    For Each cell In nameCells.Cells
        label = Trim$(CStr(cell.Value))
        If Len(label) = 0 Then GoTo NextCell
        Set nm = FindName(wb, label)
        If nm Is Nothing Then
            missing = missing & "Name " & label & " does not exist" & vbCrLf
            GoTo NextCell
        End If
        If TypeOf nm.Parent Is Worksheet Then
            localNm = LocalNameOf(nm)
            refers = nm.RefersTo
            vis = nm.Visible
            cmt = SafeComment(nm)
            nm.Delete
            DeleteNameIfExists wb, localNm
            wb.Names.Add Name:=localNm, RefersTo:=refers, Visible:=vis
            wb.Names(localNm).Comment = AppendComment(cmt, "From local to global at " & Now())
        End If
NextCell:
    Next cell
    ConvertLocalNamesToGlobal = missing
End Function

Public Function ConvertGlobalNamesToLocal(ByVal wb As Workbook, ByVal nameCells As Range, ByVal destSheet As Worksheet) As String
    Dim cell As Range
    Dim nm As Name
    Dim missing As String
    Dim localNm As String
    Dim refers As String
    Dim vis As Boolean
    Dim cmt As String
    Dim label As String
    For Each cell In nameCells.Cells
        label = Trim$(CStr(cell.Value))
        If Len(label) = 0 Then GoTo NextCell
        Set nm = FindName(wb, label)
        If nm Is Nothing Then
            missing = missing & "Name " & label & " does not exist" & vbCrLf
            GoTo NextCell
        End If
        If TypeOf nm.Parent Is Workbook Then
            localNm = ValidRangeName(LocalNameOf(nm))
            refers = nm.RefersTo
            vis = nm.Visible
            cmt = SafeComment(nm)
            nm.Delete
            On Error Resume Next
            destSheet.Names(localNm).Delete
            On Error GoTo 0
            destSheet.Names.Add Name:=localNm, RefersTo:=refers, Visible:=vis
            destSheet.Names(localNm).Comment = AppendComment(cmt, "From global to local in " & destSheet.Name & " at " & Now())
        End If
NextCell:
    Next cell
    ConvertGlobalNamesToLocal = missing
End Function

Public Sub HighlightNamedRanges(ByVal names As Object, ByVal applyFill As Boolean)
    Dim nm As Name
    Dim rng As Range
    For Each nm In names
        Set rng = SafeRefersToRange(nm)
        If Not rng Is Nothing Then
            If applyFill Then
                rng.Interior.ColorIndex = HighlightColorIndex
            Else
                rng.Interior.ColorIndex = xlColorIndexNone
            End If
        End If
    Next nm
End Sub

Private Sub DeleteNameIfExists(ByVal wb As Workbook, ByVal RangeName As String)
    Dim nm As Name
    Set nm = FindName(wb, RangeName)
    If Not nm Is Nothing Then
        If TypeOf nm.Parent Is Workbook Then nm.Delete
    End If
End Sub

Private Function RefersToFromValue(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then
        RefersToFromValue = "="""""
    ElseIf (VarType(v) = vbString) Then
        RefersToFromValue = "=""" & Replace(CStr(v), """", """""") & """"
    ElseIf IsNumeric(v) Then
        RefersToFromValue = "=" & CStr(v)
    Else
        RefersToFromValue = "=""" & Replace(CStr(v), """", """""") & """"
    End If
End Function

Private Function SafeRefersToRange(ByVal nm As Name) As Range
    On Error Resume Next
    Set SafeRefersToRange = nm.RefersToRange
    On Error GoTo 0
End Function

Private Function SafeNameValue(ByVal nm As Name) As String
    On Error Resume Next
    SafeNameValue = CStr(nm.Value)
    On Error GoTo 0
End Function

Private Function SafeComment(ByVal nm As Name) As String
    On Error Resume Next
    SafeComment = CStr(nm.Comment)
    On Error GoTo 0
End Function

Private Function SafeWorkbookParameter(ByVal nm As Name) As Variant
    On Error Resume Next
    SafeWorkbookParameter = nm.WorkbookParameter
    On Error GoTo 0
End Function

Private Function ScopeLabel(ByVal nm As Name) As String
    If TypeOf nm.Parent Is Workbook Then
        ScopeLabel = "Workbook"
    Else
        ScopeLabel = nm.Parent.Name
    End If
End Function

Private Function AppendComment(ByVal existing As String, ByVal extra As String) As String
    If Len(existing) = 0 Then
        AppendComment = extra
    Else
        AppendComment = existing & vbCrLf & extra
    End If
End Function
