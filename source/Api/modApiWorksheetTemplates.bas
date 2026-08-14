Attribute VB_Name = "modApiWorksheetTemplates"
Option Explicit

' Public API: worksheet templates from Personal123 Custom_Menu26_wkshtTmplt.
' Other workbooks / the add-in menu should call these names only.

''' @Description: Returns the Address (and SubAddress) of the first hyperlink on pRange.
''' @Example: =HyperLinkText(C4)
Public Function HyperLinkText(ByVal pRange As Range) As String
    Dim addr As String
    Dim subAddr As String

    If pRange.Hyperlinks.Count = 0 Then
        HyperLinkText = "not found"
        Exit Function
    End If

    addr = pRange.Hyperlinks(1).Address
    subAddr = pRange.Hyperlinks(1).SubAddress
    If Len(subAddr) > 0 Then addr = "[" & addr & "]" & subAddr
    HyperLinkText = addr
End Function

''' @Description: Links register (table Tbl_Links) with Category slicer and a sample Google hyperlink.
Public Sub CreateLinksTemplate()
    Dim ws As Worksheet
    Dim lo As ListObject

    On Error GoTo EH
    Set ws = modInternalWorksheetTemplates.PrepareTemplateSheet("Links", "Tbl_Links")
    If ws Is Nothing Then Exit Sub

    Call modInternalExcelApp.PushAppState
    Call WriteLinksLayout(ws)
    Set lo = modInternalWorksheetTemplates.CreateHeaderTable(ws, ws.Range("A3"), "Tbl_Links")
    If Not lo Is Nothing Then
        Call modInternalWorksheetTemplates.AddFieldSlicer(ws, lo, "Category", 20, 750, 150, 300)
    End If
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("CreateLinksTemplate")
End Sub

''' @Description: Actions tracker (table Tbl_Actions) with Live/Dead, Owner, and Category slicers.
Public Sub CreateActionsTemplate()
    Dim ws As Worksheet
    Dim lo As ListObject

    On Error GoTo EH
    Set ws = modInternalWorksheetTemplates.PrepareTemplateSheet("Actions", "Tbl_Actions")
    If ws Is Nothing Then Exit Sub

    Call modInternalExcelApp.PushAppState
    Call WriteActionsLayout(ws)
    Set lo = modInternalWorksheetTemplates.CreateHeaderTable(ws, ws.Range("A3"), "Tbl_Actions")
    If Not lo Is Nothing Then
        Call modInternalWorksheetTemplates.AddFieldSlicer(ws, lo, "Live / Dead", 20, 1000, 110, 70)
        Call modInternalWorksheetTemplates.AddFieldSlicer(ws, lo, "Owner", 100, 1000, 110, 130)
        Call modInternalWorksheetTemplates.AddFieldSlicer(ws, lo, "Action Category", 240, 1000, 110, 200)
    End If
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("CreateActionsTemplate")
End Sub

''' @Description: Force-field analysis table (drivers vs barriers).
Public Sub CreateForceFieldTemplate()
    Dim ws As Worksheet

    On Error GoTo EH
    Set ws = modInternalWorksheetTemplates.PrepareTemplateSheet("Force Field", "Tbl_ForceField")
    If ws Is Nothing Then Exit Sub

    Call modInternalExcelApp.PushAppState
    Call WriteForceFieldLayout(ws)
    Call modInternalWorksheetTemplates.CreateHeaderTable(ws, ws.Range("A3"), "Tbl_ForceField")
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("CreateForceFieldTemplate")
End Sub

''' @Description: Assumptions log (table Tbl_Assumptions) with a Status slicer.
Public Sub CreateAssumptionsTemplate()
    Dim ws As Worksheet
    Dim lo As ListObject

    On Error GoTo EH
    Set ws = modInternalWorksheetTemplates.PrepareTemplateSheet("Assumptions", "Tbl_Assumptions")
    If ws Is Nothing Then Exit Sub

    Call modInternalExcelApp.PushAppState
    Call WriteAssumptionsLayout(ws)
    Set lo = modInternalWorksheetTemplates.CreateHeaderTable(ws, ws.Range("A3"), "Tbl_Assumptions")
    If Not lo Is Nothing Then
        Call modInternalWorksheetTemplates.AddFieldSlicer(ws, lo, "Status", 20, 1000, 110, 70)
    End If
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("CreateAssumptionsTemplate")
End Sub

''' @Description: Questions log (table Tbl_Questions).
Public Sub CreateQuestionsTemplate()
    Dim ws As Worksheet

    On Error GoTo EH
    Set ws = modInternalWorksheetTemplates.PrepareTemplateSheet("Questions", "Tbl_Questions")
    If ws Is Nothing Then Exit Sub

    Call modInternalExcelApp.PushAppState
    Call WriteQuestionsLayout(ws)
    Call modInternalWorksheetTemplates.CreateHeaderTable(ws, ws.Range("A3"), "Tbl_Questions")
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("CreateQuestionsTemplate")
End Sub

''' @Description: New workbook with all five templates plus an Index. Prompts for a save path.
Public Sub CreateNotesWorkbook()
    Dim newWb As Workbook
    Dim savePath As Variant

    On Error GoTo EH
    Set newWb = Workbooks.Add
    newWb.Activate

    Call CreateLinksTemplate
    Call CreateActionsTemplate
    Call CreateForceFieldTemplate
    Call CreateAssumptionsTemplate
    Call CreateQuestionsTemplate
    Call modInternalWorksheetTemplates.BuildIndexSheet

    savePath = Application.GetSaveAsFilename( _
        InitialFileName:="Notes.xlsx", _
        FileFilter:="Excel Workbook (*.xlsx), *.xlsx", _
        Title:="Save Notes workbook")
    If savePath <> False Then
        newWb.SaveAs Filename:=CStr(savePath), FileFormat:=xlOpenXMLWorkbook
        MsgBox "Saved " & newWb.FullName, vbInformation, "Excel VBA Lib"
    End If
    Exit Sub
EH:
    Call modInternalError.RaiseCurrent("CreateNotesWorkbook")
End Sub

''' @Description: Sheet of Excel Python (=PY) import stubs, moved to the first sheet.
Public Sub ImportPythonPackages()
    Dim ws As Worksheet

    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    Set ws = modApiSheets.EnsureSheet("Python packages")
    ws.Cells.Clear
    If ws.Index > 1 Then ws.Move Before:=ActiveWorkbook.Worksheets(1)
    Call WritePythonPackageFormulas(ws)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("ImportPythonPackages")
End Sub

' =============================================================================
' Sheet layouts
' =============================================================================

Private Sub WriteLinksLayout(ByVal ws As Worksheet)
    With ws
        .Range("A3").Value = "Category"
        .Range("A3").ColumnWidth = 20
        .Range("B3").Value = "Description"
        .Range("B3").ColumnWidth = 30
        .Range("C3").Value = "Link"
        .Range("C3").ColumnWidth = 30
        .Range("D3").Value = "Archive Y/N"
        .Range("D3").ColumnWidth = 15
        .Range("E3").Value = "Sharepoint"
        .Range("E3").ColumnWidth = 15
        .Range("F3").Value = "Hyperlink"
        .Range("F3").ColumnWidth = 50

        .Range("A4").Value = "Browser"
        .Range("B4").Value = "Search engine"
        .Hyperlinks.Add Anchor:=.Range("C4"), Address:="https://www.google.com", TextToDisplay:="Google"
        ' UDF lives in this add-in; returns the hyperlink address from column C.
        .Range("F4").FormulaR1C1 = "=HyperLinkText(RC[-3])"
    End With
End Sub

Private Sub WriteActionsLayout(ByVal ws As Worksheet)
    Dim hdr As Range
    Set hdr = ws.Range("A3:L3")

    hdr.Cells(1).Value = "Ref"
    hdr.Cells(2).Value = "Action Category"
    hdr.Cells(3).Value = "Action"
    hdr.Cells(4).Value = "Owner"
    hdr.Cells(5).Value = "Due date"
    hdr.Cells(5).Offset(1, 0).NumberFormat = "ddd dd/mm/yyyy"
    hdr.Cells(6).Value = "Completion date"
    hdr.Cells(6).Offset(1, 0).NumberFormat = "ddd dd/mm/yyyy"
    hdr.Cells(7).Value = "Status"
    hdr.Cells(8).Value = "Comment"
    ' Col 9 Live/Dead from Status (col 7 = RC[-2]).
    hdr.Cells(9).Value = "Live / Dead"
    hdr.Cells(9).Offset(1, 0).FormulaR1C1 = _
        "=IF(OR(RC[-2]=""Closed"",RC[-2]=""Dropped"",RC[-2]=""Done""),""Dead"",""Live"")"
    ' Col 10 Timeliness = Completion (RC[-4]) - Due (RC[-5]).
    hdr.Cells(10).Value = "Timeliness"
    hdr.Cells(10).Offset(1, 0).FormulaR1C1 = _
        "=IF(AND(ISNUMBER(RC[-5]),ISNUMBER(RC[-4])),RC[-4]-RC[-5],"" - "")"
    ' Col 11 Days to go = Due (RC[-6]) - TODAY when Live (RC[-2]).
    hdr.Cells(11).Value = "Days to go"
    With hdr.Cells(11).Offset(1, 0)
        .FormulaR1C1 = "=IF(RC[-2]=""Live"",RC[-6]-TODAY(),""-"")"
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
        .FormatConditions(1).Interior.Color = vbRed
    End With
    hdr.Cells(12).Value = "Overdue"
    hdr.Cells(12).Offset(1, 0).FormulaR1C1 = "=IF(RC[-1]<0,""Overdue"",""Current"")"

    ws.Columns(1).ColumnWidth = 4
    ws.Columns(2).ColumnWidth = 24
    ws.Columns(3).ColumnWidth = 55
    ws.Columns(4).ColumnWidth = 16
    ws.Columns(5).ColumnWidth = 16
    ws.Columns(6).ColumnWidth = 16
    ws.Columns(7).ColumnWidth = 16
    ws.Columns(8).ColumnWidth = 40
    ws.Columns(9).ColumnWidth = 10
    ws.Columns(10).ColumnWidth = 10
    ws.Columns(11).ColumnWidth = 10
    ws.Columns(12).ColumnWidth = 10
End Sub

Private Sub WriteForceFieldLayout(ByVal ws As Worksheet)
    With ws
        .Range("A3").Value = "Drivers"
        .Range("A3").ColumnWidth = 40
        .Range("B3").Value = "Driver weight"
        .Range("B3").ColumnWidth = 10
        .Range("C3").Value = "Barrier weight"
        .Range("C3").ColumnWidth = 10
        .Range("D3").Value = "Barrier"
        .Range("D3").ColumnWidth = 40
    End With
End Sub

Private Sub WriteAssumptionsLayout(ByVal ws As Worksheet)
    With ws
        .Range("A3").Value = "Ref"
        .Range("B3").Value = "Assumption"
        .Range("C3").Value = "Status"
        .Range("D3").Value = "Archive"
        .Range("A4").Value = "A1"
        .Columns(1).ColumnWidth = 4
        .Columns(2).ColumnWidth = 80
        .Columns(3).ColumnWidth = 15
        .Columns(4).ColumnWidth = 15
    End With
End Sub

Private Sub WriteQuestionsLayout(ByVal ws As Worksheet)
    With ws
        .Range("A3").Value = "Ref"
        .Range("B3").Value = "Question"
        .Range("C3").Value = "Response"
        .Range("D3").Value = "Status"
        .Range("A4").Value = "Q1"
        .Columns(1).ColumnWidth = 4
        .Columns(2).ColumnWidth = 65
        .Columns(3).ColumnWidth = 65
        .Columns(4).ColumnWidth = 15
    End With
End Sub

Private Sub WritePythonPackageFormulas(ByVal ws As Worksheet)
    ' Excel Python cells: first argument is the code, second is a comment shown in the editor.
    With ws
        .Range("A3").Formula2R1C1 = "=PY(""import pandas as pd" & Chr(10) & """""Import pandas package as pd"""""",1)"
        .Range("A4").Formula2R1C1 = "=PY(""import numpy as np" & Chr(10) & """""Import numpy package as np"""""",1)"
        .Range("A5").Formula2R1C1 = "=PY(""import re" & Chr(10) & """""Import re regex package"""""",1)"
        .Range("A6").Formula2R1C1 = "=PY(""import seaborn as sns" & Chr(10) & """""Import seaborn package as sns"""""",1)"
        .Range("A7").Value = "matplotlib"
        .Range("A7").Font.Bold = True
        .Range("A8").Formula2R1C1 = "=PY(""import matplotlib.pyplot as plt" & Chr(10) & """""Import matplot package as plt"""""",1)"
        .Range("A9").Formula2R1C1 = "=PY(""import matplotlib.ticker as ticker" & Chr(10) & """""Import matplot ticker as ticker"""""",1)"
        .Range("A10").Formula2R1C1 = "=PY(""from matplotlib.patches import Rectangle" & Chr(10) & """""Import Rectangle for Gantt charts"""""",1)"
        .Range("A11").Value = "scipy"
        .Range("A11").Font.Bold = True
        .Range("A12").Formula2R1C1 = "=PY(""import scipy.stats as stats" & Chr(10) & """""Import scipy stats package as stats"""""",1)"
        .Range("A13").Formula2R1C1 = "=PY(""from scipy.linalg import svd" & Chr(10) & """""Import SVD"""""",1)"
        .Range("A14").Value = "sklearn"
        .Range("A14").Font.Bold = True
        .Range("A15").Formula2R1C1 = "=PY(""from sklearn.decomposition import PCA" & Chr(10) & """""Import PCA"""""",1)"
        .Range("A16").Formula2R1C1 = "=PY(""from sklearn.linear_model import LinearRegression" & Chr(10) & """""Import Linear regression"""""",1)"
        .Range("A17").Formula2R1C1 = "=PY(""from sklearn.metrics import mean_squared_error" & Chr(10) & """""Import rmse from sklearn"""""",1)"
        .Range("A18").Formula2R1C1 = "=PY(""from sklearn.preprocessing import MinMaxScaler" & Chr(10) & """""Import MinMaxScaler"""""",1)"
        .Range("A19").Formula2R1C1 = "=PY(""from sklearn.preprocessing import MaxAbsScaler" & Chr(10) & """""Import MaxAbsScaler"""""",1)"
        .Range("A20").Formula2R1C1 = "=PY(""from sklearn.preprocessing import RobustScaler" & Chr(10) & """""Import RobustScaler"""""",1)"
        .Range("A21").Formula2R1C1 = "=PY(""from sklearn.preprocessing import StandardScaler" & Chr(10) & """""Import StandardScaler"""""",1)"
        .Range("A22").Value = "statsmodel"
        .Range("A22").Font.Bold = True
        .Range("A23").Formula2R1C1 = "=PY(""from statsmodels.tsa.stattools import adfuller" & Chr(10) & """""Import adfuller from statsmodel"""""",1)"
        .Range("A24").Formula2R1C1 = "=PY(""from statsmodels.tsa.stattools import acf" & Chr(10) & """""Import acf from statsmodel"""""",1)"
        .Range("A25").Formula2R1C1 = "=PY(""from statsmodels.tsa.stattools import pacf" & Chr(10) & """""Import pacf from statsmodel"""""",1)"
        .Range("A26").Formula2R1C1 = "=PY(""from statsmodels.tsa.seasonal import STL" & Chr(10) & """""Import STL"""""",1)"
        .Range("A27").Formula2R1C1 = "=PY(""from statsmodels.tsa.arima.model import ARIMA" & Chr(10) & """""Import ARIMA"""""",1)"
        .Range("A28").Formula2R1C1 = "=PY(""import statsmodels.api as sm" & Chr(10) & """""Import statsmodel api as sm"""""",1)"
        .Range("A29").Formula2R1C1 = "=PY(""from statsmodels.multivariate.manova import MANOVA" & Chr(10) & """""Import MANOVA"""""",1)"
    End With
End Sub
