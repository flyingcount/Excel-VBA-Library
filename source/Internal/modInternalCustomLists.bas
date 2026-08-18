Attribute VB_Name = "modInternalCustomLists"
Option Explicit

' Internal: Excel custom lists and AutoCorrect entries
' (Personal Custom_Menu16_CustomLists / Custom_Menu16_Autocorrect).
' Custom lists and AutoCorrect replacements are application-level, not workbook-level.
' Called from modApiCustomLists. Do not document these as the external API.

Public Const CustomListsSheetName As String = "Custom List properties"
Public Const AutoCorrectSheetName As String = "Auto correct List"
' Excel built-in lists: Sun-Sat, Sunday-Saturday, Jan-Dec, January-December.
Public Const BuiltInCustomListCount As Long = 4

Public Function CustomListCountValue() As Long
    CustomListCountValue = Application.CustomListCount
End Function

Public Function PromptListRange(ByVal PromptText As String, Optional ByVal Title As String = "Custom lists") As Range
    Dim rng As Range
    Dim defAddr As String
    If TypeName(Selection) = "Range" Then defAddr = Selection.Address
    On Error Resume Next
    If Len(defAddr) > 0 Then
        Set rng = Application.InputBox(Prompt:=PromptText, Title:=Title, Default:=defAddr, Type:=8)
    Else
        Set rng = Application.InputBox(Prompt:=PromptText, Title:=Title, Type:=8)
    End If
    On Error GoTo 0
    Set PromptListRange = rng
End Function

' Intersect with UsedRange so entire-column selections do not scan a million empty cells.
Public Function UsedPart(ByVal rng As Range) As Range
    Dim u As Range
    If rng Is Nothing Then Exit Function
    On Error Resume Next
    Set u = rng.Worksheet.UsedRange
    On Error GoTo 0
    If u Is Nothing Then
        Set UsedPart = rng
        Exit Function
    End If
    Set UsedPart = Intersect(rng, u)
End Function

Public Function CustomListInventoryArray() As Variant
    Dim n As Long
    Dim i As Long
    Dim j As Long
    Dim maxItems As Long
    Dim items As Variant
    Dim lists() As Variant
    Dim arr As Variant
    Dim itemCount As Long
    Dim lb As Long

    n = Application.CustomListCount
    If n = 0 Then
        CustomListInventoryArray = Empty
        Exit Function
    End If

    ReDim lists(1 To n)
    maxItems = 0
    For i = 1 To n
        lists(i) = Application.GetCustomListContents(i)
        itemCount = UBound(lists(i)) - LBound(lists(i)) + 1
        If itemCount > maxItems Then maxItems = itemCount
    Next i

    ReDim arr(1 To n + 1, 1 To 3 + maxItems)
    arr(1, 1) = "Custom list number"
    arr(1, 2) = "Type"
    arr(1, 3) = "Item count"
    arr(1, 4) = "Custom list elements"

    For i = 1 To n
        items = lists(i)
        lb = LBound(items)
        itemCount = UBound(items) - lb + 1
        arr(i + 1, 1) = i
        If i <= BuiltInCustomListCount Then
            arr(i + 1, 2) = "Built-in"
        Else
            arr(i + 1, 2) = "User"
        End If
        arr(i + 1, 3) = itemCount
        For j = 0 To itemCount - 1
            arr(i + 1, 4 + j) = items(lb + j)
        Next j
    Next i
    CustomListInventoryArray = arr
End Function

Public Sub WriteCustomListInventory()
    Dim arr As Variant
    Dim ws As Worksheet
    Dim n As Long
    Dim cols As Long

    n = Application.CustomListCount
    If n = 0 Then
        MsgBox "There are no custom lists.", vbInformation, "Custom lists"
        Exit Sub
    End If

    arr = CustomListInventoryArray()
    cols = UBound(arr, 2)
    Call modApiSheets.CreateOutputSheet(CustomListsSheetName)
    Set ws = ActiveWorkbook.Worksheets(CustomListsSheetName)
    ws.Range("A1").Value = "Excel custom lists (application-wide; lists 1-" & _
        CStr(BuiltInCustomListCount) & " are built-in and cannot be deleted)"
    ws.Range("A1").Font.Bold = True
    ws.Range("A3").Resize(n + 1, cols).Value = arr
    ws.Range("A3").Resize(1, cols).Font.Bold = True
    ws.Range("A3").WrapText = True
    ws.Rows(3).RowHeight = 30
    ws.Columns(1).ColumnWidth = 12
    ws.Columns(2).ColumnWidth = 12
    ws.Columns(3).ColumnWidth = 12
    If cols > 3 Then ws.Columns(4).Resize(, cols - 3).AutoFit
    ws.Activate
End Sub

' byRow = True: one list per row; False: one list per column.
Public Sub AddCustomListsFromRange(ByVal rng As Range, ByVal byRow As Boolean, _
    ByRef added As Long, ByRef skippedExist As Long, ByRef tooShort As Long)

    Dim i As Long
    Dim n As Long
    Dim vec As Range
    Dim items As Variant

    added = 0
    skippedExist = 0
    tooShort = 0
    Set rng = UsedPart(rng)
    If rng Is Nothing Then Exit Sub

    If byRow Then
        n = rng.Rows.Count
    Else
        n = rng.Columns.Count
    End If

    For i = 1 To n
        If byRow Then
            Set vec = rng.Rows(i)
        Else
            Set vec = rng.Columns(i)
        End If
        items = NonBlankTexts(vec)
        If IsEmpty(items) Then
            tooShort = tooShort + 1
        ElseIf ExistingListNumber(items) > 0 Then
            skippedExist = skippedExist + 1
        Else
            Application.AddCustomList ListArray:=items
            added = added + 1
        End If
    Next i
End Sub

' Deletes user lists whose numbers appear in nums. Highest numbers first so remaining
' list numbers stay valid. Built-in lists 1-4 are never deleted.
Public Function DeleteCustomListsByNumbers(ByVal nums As Range, _
    ByRef skippedBuiltIn As Long, ByRef skippedInvalid As Long) As Long

    Dim cell As Range
    Dim dict As Object
    Dim keys As Variant
    Dim a() As Long
    Dim i As Long
    Dim n As Long
    Dim v As Long
    Dim maxList As Long
    Dim deleted As Long
    Dim rng As Range

    skippedBuiltIn = 0
    skippedInvalid = 0
    Set rng = UsedPart(nums)
    If rng Is Nothing Then Exit Function

    maxList = Application.CustomListCount
    Set dict = CreateObject("Scripting.Dictionary")
    For Each cell In rng.Cells
        v = ListNumberFromCell(cell.Value)
        If v <> 0 Then
            If Not dict.Exists(v) Then dict.Add v, v
        End If
    Next cell
    If dict.Count = 0 Then Exit Function

    keys = dict.Keys
    n = dict.Count
    ReDim a(1 To n)
    For i = 1 To n
        a(i) = CLng(keys(i - 1))
    Next i
    Call SortLongDesc(a, n)

    For i = 1 To n
        v = a(i)
        If v >= 1 And v <= BuiltInCustomListCount Then
            skippedBuiltIn = skippedBuiltIn + 1
        ElseIf v < 1 Or v > maxList Then
            skippedInvalid = skippedInvalid + 1
        Else
            Application.DeleteCustomList v
            deleted = deleted + 1
            maxList = maxList - 1
        End If
    Next i
    DeleteCustomListsByNumbers = deleted
End Function

Public Function AutoCorrectInventoryArray() As Variant
    Dim ace As Variant
    Dim n As Long
    Dim i As Long
    Dim arr As Variant
    Dim lb As Long

    ace = Application.AutoCorrect.ReplacementList
    If IsEmpty(ace) Then
        AutoCorrectInventoryArray = Empty
        Exit Function
    End If
    lb = LBound(ace, 1)
    n = UBound(ace, 1) - lb + 1
    ReDim arr(1 To n + 1, 1 To 2)
    arr(1, 1) = "Replace:"
    arr(1, 2) = "With:"
    For i = 0 To n - 1
        arr(i + 2, 1) = ace(lb + i, 1)
        arr(i + 2, 2) = ace(lb + i, 2)
    Next i
    AutoCorrectInventoryArray = arr
End Function

Public Sub WriteAutoCorrectInventory()
    Dim arr As Variant
    Dim ws As Worksheet
    Dim n As Long

    arr = AutoCorrectInventoryArray()
    If IsEmpty(arr) Then
        MsgBox "There are no AutoCorrect entries.", vbInformation, "Custom lists"
        Exit Sub
    End If

    n = UBound(arr, 1) - 1
    Call modApiSheets.CreateOutputSheet(AutoCorrectSheetName)
    Set ws = ActiveWorkbook.Worksheets(AutoCorrectSheetName)
    ws.Range("A1").Value = "Excel AutoCorrect replacements (application-wide)"
    ws.Range("A1").Font.Bold = True
    ws.Range("A3").Resize(n + 1, 2).Value = arr
    ws.Range("A3:B3").Font.Bold = True
    ws.Columns("A:B").AutoFit
    ws.Activate
End Sub

' Two-column range: Replace | With. Adds new pairs and overwrites matching codes.
' Does not remove existing AutoCorrect entries that are absent from the range.
Public Function AddAutoCorrectFromRange(ByVal rng As Range) As Long
    Dim r As Long
    Dim n As Long
    Dim startRow As Long
    Dim whatText As String
    Dim withText As String
    Dim added As Long
    Dim data As Variant

    Set rng = UsedPart(rng)
    If rng Is Nothing Then Exit Function
    If rng.Columns.Count < 2 Then Exit Function

    Set rng = rng.Resize(rng.Rows.Count, 2)
    If rng.Cells.Count = 2 Then
        ReDim data(1 To 1, 1 To 2)
        data(1, 1) = rng.Cells(1, 1).Value
        data(1, 2) = rng.Cells(1, 2).Value
        n = 1
    Else
        data = rng.Value
        n = UBound(data, 1)
    End If

    startRow = 1
    If LooksLikeAutoCorrectHeader(data(1, 1), data(1, 2)) Then startRow = 2

    For r = startRow To n
        If Not IsError(data(r, 1)) And Not IsError(data(r, 2)) Then
            If Not modInternalText.IsBlank(data(r, 1)) Then
                whatText = CStr(data(r, 1))
                If IsEmpty(data(r, 2)) Or IsNull(data(r, 2)) Then
                    withText = ""
                Else
                    withText = CStr(data(r, 2))
                End If
                Application.AutoCorrect.AddReplacement whatText, withText
                added = added + 1
            End If
        End If
    Next r
    AddAutoCorrectFromRange = added
End Function

Private Function NonBlankTexts(ByVal rng As Range) As Variant
    Dim cell As Range
    Dim col As Collection
    Dim i As Long
    Dim arr As Variant

    Set col = New Collection
    For Each cell In rng.Cells
        If Not IsError(cell.Value) Then
            If Not modInternalText.IsBlank(cell.Value) Then
                col.Add CStr(cell.Value)
            End If
        End If
    Next cell
    If col.Count < 2 Then
        NonBlankTexts = Empty
        Exit Function
    End If
    ReDim arr(1 To col.Count)
    For i = 1 To col.Count
        arr(i) = col(i)
    Next i
    NonBlankTexts = arr
End Function

Private Function ExistingListNumber(ByRef items As Variant) As Long
    Dim n As Long
    On Error Resume Next
    n = Application.GetCustomListNum(items)
    If Err.Number <> 0 Then
        Err.Clear
        n = 0
    End If
    On Error GoTo 0
    ExistingListNumber = n
End Function

Private Function ListNumberFromCell(ByVal Value As Variant) As Long
    If IsError(Value) Or IsEmpty(Value) Or IsNull(Value) Then Exit Function
    If VarType(Value) = vbBoolean Then Exit Function
    If Not IsNumeric(Value) Then Exit Function
    If CDbl(Value) <> CLng(Value) Then Exit Function
    ListNumberFromCell = CLng(Value)
End Function

Private Sub SortLongDesc(ByRef a() As Long, ByVal n As Long)
    Dim i As Long
    Dim j As Long
    Dim tmp As Long
    For i = 1 To n - 1
        For j = i + 1 To n
            If a(j) > a(i) Then
                tmp = a(i)
                a(i) = a(j)
                a(j) = tmp
            End If
        Next j
    Next i
End Sub

Private Function LooksLikeAutoCorrectHeader(ByVal a As Variant, ByVal b As Variant) As Boolean
    Dim sa As String
    Dim sb As String
    If IsError(a) Or IsError(b) Then Exit Function
    sa = LCase$(Replace(Trim$(CStr(a)), ":", ""))
    sb = LCase$(Replace(Trim$(CStr(b)), ":", ""))
    LooksLikeAutoCorrectHeader = (sa = "replace" And (sb = "with" Or sb = "replacement"))
End Function
