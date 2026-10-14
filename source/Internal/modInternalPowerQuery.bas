Attribute VB_Name = "modInternalPowerQuery"
Option Explicit

' Internal: Power Query connection and function-library helpers
' (Personal Custom_Menu29_PowerQuery / modPQ_*).
' Called from modApiPowerQuery and frmPQLibrary. Do not document these as the external API.

Private Const ConnPrefix As String = "Query - "
Public Const PqDocSheet As String = "PQ_Functions"
Public Const PqDocTable As String = "PQ_Functions"
Public Const PqLibraryTable As String = "tblPQ_Library"

Public Function HasQueries(ByVal wb As Workbook) As Boolean
    Dim n As Long
    If wb Is Nothing Then Exit Function
    On Error Resume Next
    n = wb.Queries.Count
    HasQueries = (Err.Number = 0)
    Err.Clear
End Function

' Flip BackgroundQuery on every OLEDB connection. nOn / nOff are counts after the flip.
Public Function ToggleOledbBackgroundQuery(ByVal wb As Workbook, ByRef nOn As Long, ByRef nOff As Long) As Long
    Dim i As Long
    Dim n As Long
    Dim conn As WorkbookConnection
    Dim ole As OLEDBConnection

    nOn = 0
    nOff = 0
    If wb Is Nothing Then Exit Function

    For i = 1 To wb.Connections.Count
        Set conn = wb.Connections(i)
        If conn.Type <> xlConnectionTypeOLEDB Then GoTo NextConn
        On Error Resume Next
        Set ole = conn.OLEDBConnection
        If Err.Number <> 0 Or ole Is Nothing Then
            Err.Clear
            On Error GoTo 0
            GoTo NextConn
        End If
        On Error GoTo 0
        ole.BackgroundQuery = Not ole.BackgroundQuery
        n = n + 1
        If ole.BackgroundQuery Then
            nOn = nOn + 1
        Else
            nOff = nOff + 1
        End If
NextConn:
        Set ole = Nothing
        Set conn = Nothing
    Next i
    ToggleOledbBackgroundQuery = n
End Function

' Returns the new FastCombine value.
Public Function ToggleFastCombine(ByVal wb As Workbook) As Boolean
    wb.Queries.FastCombine = Not wb.Queries.FastCombine
    ToggleFastCombine = wb.Queries.FastCombine
End Function

Public Function CountWorkbookTables(ByVal wb As Workbook) As Long
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim n As Long
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            n = n + 1
        Next lo
    Next ws
    CountWorkbookTables = n
End Function

' Create a connection-only query for each Excel table that does not already have one.
' When addToModel is True, the connection is loaded to the Data Model instead of
' adding a second connection with the same name (Personal did both and could fail).
Public Function AddConnectionsForAllTables(ByVal wb As Workbook, ByVal addToModel As Boolean) As Long
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim created As Long

    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If AddConnectionForTable(wb, lo.Name, addToModel) Then created = created + 1
        Next lo
    Next ws
    AddConnectionsForAllTables = created
End Function

Private Function AddConnectionForTable(ByVal wb As Workbook, ByVal tableName As String, ByVal addToModel As Boolean) As Boolean
    Dim src As String
    Dim formula As String
    Dim connName As String

    If Len(tableName) = 0 Then Exit Function
    src = TableSourceFormula(tableName)
    If QueryAlreadyCoversTable(wb, tableName, src) Then Exit Function

    connName = ConnPrefix & tableName
    If ConnectionExists(wb, connName) Then Exit Function

    formula = "let" & vbCrLf & _
              "    Source = " & src & vbCrLf & _
              "in" & vbCrLf & _
              "    Source"

    wb.Queries.Add Name:=tableName, Formula:=formula

    If addToModel Then
        wb.Connections.Add2 Name:=connName, _
            Description:="Connection to the '" & tableName & "' query in the workbook.", _
            ConnectionString:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & tableName & ";Extended Properties=""""", _
            CommandText:=tableName, _
            lCmdtype:=6, _
            CreateModelConnection:=True, _
            ImportRelationships:=False
    Else
        wb.Connections.Add2 Name:=connName, _
            Description:="Connection to the '" & tableName & "' query in the workbook.", _
            ConnectionString:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & tableName & ";Extended Properties=""""", _
            CommandText:="SELECT * FROM [" & tableName & "]", _
            lCmdtype:=2, _
            CreateModelConnection:=False, _
            ImportRelationships:=False
    End If
    AddConnectionForTable = True
End Function

Private Function TableSourceFormula(ByVal tableName As String) As String
    TableSourceFormula = "Excel.CurrentWorkbook(){[Name=" & QuoteM(tableName) & "]}[Content]"
End Function

Private Function QuoteM(ByVal s As String) As String
    QuoteM = """" & Replace$(s, """", """""") & """"
End Function

Private Function QueryAlreadyCoversTable(ByVal wb As Workbook, ByVal tableName As String, ByVal srcFormula As String) As Boolean
    Dim wq As WorkbookQuery
    For Each wq In wb.Queries
        If StrComp(wq.Name, tableName, vbTextCompare) = 0 Then
            QueryAlreadyCoversTable = True
            Exit Function
        End If
        If InStr(1, wq.Formula, srcFormula, vbTextCompare) > 0 Then
            QueryAlreadyCoversTable = True
            Exit Function
        End If
    Next wq
End Function

Private Function ConnectionExists(ByVal wb As Workbook, ByVal connName As String) As Boolean
    Dim conn As WorkbookConnection
    On Error Resume Next
    Set conn = wb.Connections(connName)
    On Error GoTo 0
    ConnectionExists = Not conn Is Nothing
End Function

Public Function PqNz(ByVal v As Variant, Optional ByVal defaultValue As String = "") As String
    If IsError(v) Then
        PqNz = defaultValue
    ElseIf IsNull(v) Or IsEmpty(v) Then
        PqNz = defaultValue
    ElseIf VarType(v) = vbString And Len(CStr(v)) = 0 Then
        PqNz = defaultValue
    Else
        PqNz = CStr(v)
    End If
End Function

' Master library ListObject tblPQ_Library in any open workbook, or one the user picks.
Public Function GetMasterLibraryTable() As ListObject
    Dim wb As Workbook
    Dim lo As ListObject
    Dim path As Variant

    For Each wb In Application.Workbooks
        Set lo = FindNamedTable(wb, PqLibraryTable)
        If Not lo Is Nothing Then
            Set GetMasterLibraryTable = lo
            Exit Function
        End If
    Next wb

    path = Application.GetOpenFilename( _
        FileFilter:="Excel files (*.xlsb;*.xlsm;*.xlsx),*.xlsb;*.xlsm;*.xlsx", _
        Title:="Open the workbook that contains tblPQ_Library")
    If VarType(path) = vbBoolean Then Exit Function
    If Len(CStr(path)) = 0 Then Exit Function

    On Error Resume Next
    Set wb = Application.Workbooks.Open(CStr(path))
    On Error GoTo 0
    If wb Is Nothing Then
        MsgBox "Could not open " & CStr(path) & ".", vbExclamation, "Power Query"
        Exit Function
    End If
    Set lo = FindNamedTable(wb, PqLibraryTable)
    If lo Is Nothing Then
        MsgBox "No table named " & PqLibraryTable & " in " & wb.Name & ".", vbExclamation, "Power Query"
        Exit Function
    End If
    Set GetMasterLibraryTable = lo
End Function

Public Function GetDocTable(ByVal wb As Workbook) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    On Error Resume Next
    Set ws = wb.Worksheets(PqDocSheet)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = PqDocSheet
    End If

    On Error Resume Next
    Set lo = ws.ListObjects(PqDocTable)
    On Error GoTo 0
    If lo Is Nothing Then
        ws.Range("A1").Resize(1, 6).Value = Array("FunctionName", "Category", "Version", "Dependencies", "Description", "MCode")
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").CurrentRegion, , xlYes)
        lo.Name = PqDocTable
    End If
    Set GetDocTable = lo
End Function

Public Function GetDependenciesFromMCode(ByVal mCode As String, ByVal selfName As String) As String
    Dim re As Object
    Dim matches As Object
    Dim m As Object
    Dim dict As Object
    Dim key As Variant
    Dim arr() As String
    Dim i As Long

    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "\bfn[A-Za-z0-9_]+\b"
    re.Global = True
    re.IgnoreCase = False
    Set dict = CreateObject("Scripting.Dictionary")

    If re.Test(mCode) Then
        Set matches = re.Execute(mCode)
        For Each m In matches
            If StrComp(CStr(m.Value), selfName, vbTextCompare) <> 0 Then
                If Not dict.Exists(m.Value) Then dict.Add m.Value, True
            End If
        Next m
    End If
    If dict.Count = 0 Then Exit Function

    ReDim arr(0 To dict.Count - 1)
    i = 0
    For Each key In dict.Keys
        arr(i) = CStr(key)
        i = i + 1
    Next key
    Call QuickSortStrings(arr, LBound(arr), UBound(arr))
    GetDependenciesFromMCode = Join(arr, ", ")
End Function

Public Function ExportWorkbookQueriesToPQFunctions(ByVal wb As Workbook) As Long
    Dim loDoc As ListObject
    Dim wsDoc As Worksheet
    Dim q As WorkbookQuery
    Dim newRow As ListRow
    Dim fnName As String
    Dim category As String
    Dim parts() As String
    Dim mCode As String
    Dim n As Long

    Set loDoc = GetDocTable(wb)
    Set wsDoc = loDoc.Parent
    If Not loDoc.DataBodyRange Is Nothing Then loDoc.DataBodyRange.Delete

    For Each q In wb.Queries
        mCode = q.Formula
        parts = Split(q.Name, "/")
        Select Case UBound(parts)
            Case 0
                category = ""
                fnName = parts(0)
            Case 1
                category = parts(0)
                fnName = parts(1)
            Case Else
                category = parts(UBound(parts) - 1)
                fnName = parts(UBound(parts))
        End Select
        Set newRow = loDoc.ListRows.Add
        With newRow.Range
            .Cells(1, 1).Value = fnName
            .Cells(1, 2).Value = category
            .Cells(1, 3).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
            .Cells(1, 4).Value = GetDependenciesFromMCode(mCode, fnName)
            .Cells(1, 5).Value = ""
            .Cells(1, 6).Value = mCode
        End With
        n = n + 1
    Next q
    wsDoc.Columns.AutoFit
    ExportWorkbookQueriesToPQFunctions = n
End Function

Public Sub ImportFunctionByName(ByVal targetFunctionName As String, ByVal loLib As ListObject, ByVal destWb As Workbook)
    Dim visited As Object
    Set visited = CreateObject("Scripting.Dictionary")
    Call ImportFunctionRecursive(targetFunctionName, loLib, destWb, visited)
End Sub

Private Sub ImportFunctionRecursive(ByVal fnName As String, ByVal loLib As ListObject, ByVal destWb As Workbook, ByVal visited As Object)
    Dim row As ListRow
    Dim deps As String
    Dim depArr() As String
    Dim i As Long

    If visited.Exists(fnName) Then Exit Sub
    visited.Add fnName, True
    Set row = FindFunctionRow(loLib, fnName)
    If row Is Nothing Then Exit Sub
    deps = PqNz(row.Range(1, 4).Value)
    If Len(deps) > 0 Then
        depArr = Split(deps, ",")
        For i = LBound(depArr) To UBound(depArr)
            depArr(i) = Trim$(depArr(i))
            If Len(depArr(i)) > 0 Then Call ImportFunctionRecursive(depArr(i), loLib, destWb, visited)
        Next i
    End If
    Call ImportSingleFunctionRow(row, destWb)
End Sub

Private Function FindFunctionRow(ByVal lo As ListObject, ByVal fnName As String) As ListRow
    Dim r As ListRow
    For Each r In lo.ListRows
        If StrComp(PqNz(r.Range(1, 1).Value), fnName, vbTextCompare) = 0 Then
            Set FindFunctionRow = r
            Exit Function
        End If
    Next r
End Function

Private Sub ImportSingleFunctionRow(ByVal r As ListRow, ByVal destWb As Workbook)
    Dim fnName As String
    Dim category As String
    Dim mCode As String
    Dim fullName As String
    Dim q As WorkbookQuery

    fnName = PqNz(r.Range(1, 1).Value)
    category = PqNz(r.Range(1, 2).Value)
    mCode = PqNz(r.Range(1, 6).Value)
    If Len(fnName) = 0 Or Len(mCode) = 0 Then Exit Sub
    If Len(category) = 0 Then
        fullName = fnName
    Else
        fullName = category & "/" & fnName
    End If
    On Error Resume Next
    Set q = destWb.Queries(fullName)
    On Error GoTo 0
    If q Is Nothing Then
        destWb.Queries.Add Name:=fullName, Formula:=mCode
    Else
        q.Formula = mCode
    End If
End Sub

Private Function FindNamedTable(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
                Set FindNamedTable = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

Private Sub QuickSortStrings(ByRef arr() As String, ByVal first As Long, ByVal last As Long)
    Dim i As Long
    Dim j As Long
    Dim pivot As String
    Dim temp As String
    i = first
    j = last
    pivot = arr((first + last) \ 2)
    Do While i <= j
        Do While arr(i) < pivot
            i = i + 1
        Loop
        Do While arr(j) > pivot
            j = j - 1
        Loop
        If i <= j Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop
    If first < j Then Call QuickSortStrings(arr, first, j)
    If i < last Then Call QuickSortStrings(arr, i, last)
End Sub

