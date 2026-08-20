Attribute VB_Name = "modApiLogit"
Option Explicit

' Public API: Personal Custom_Menu5_Logit.
' Empty logit worksheet: coefficients (yellow), data block, and live logit / probability formulae.

Private Const SheetName As String = "Logit Input Template"
Private Const Yellow As Long = 65535

''' @Description: Create a logit input template with coefficient cells, a data block, and example logit / probability formulae.
''' @Example: CreateLogitInputTemplate
Public Sub CreateLogitInputTemplate()
    Dim ws As Worksheet
    Dim i As Long
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    Call modApiSheets.CreateOutputSheet(SheetName)
    Set ws = ActiveWorkbook.Worksheets(SheetName)
    With ws
        .Range("A3").Value = "Logit regression input template"
        .Range("A3").Font.Bold = True
        .Range("B6").Value = "Variable"
        .Range("C6").Value = "b0"
        .Range("D6").Value = "b1"
        .Range("E6").Value = "b2"
        .Range("F6").Value = "b3"
        .Range("G6").Value = "b4"
        .Range("H6").Value = "etc"
        .Range("B6:H6").Font.Bold = True
        .Range("B6:H6").HorizontalAlignment = xlCenter
        .Range("C7:G7").Value = 0.01
        .Range("C7:G7").Interior.Color = Yellow
        .Range("C7:G7").HorizontalAlignment = xlCenter
        Call modInternalNamedRanges.CreateSheetNamedRange(ws, "b0", .Range("C7"))
        Call modInternalNamedRanges.CreateSheetNamedRange(ws, "b1", .Range("D7"))
        Call modInternalNamedRanges.CreateSheetNamedRange(ws, "b2", .Range("E7"))
        Call modInternalNamedRanges.CreateSheetNamedRange(ws, "b3", .Range("F7"))
        Call modInternalNamedRanges.CreateSheetNamedRange(ws, "b4", .Range("G7"))
        .Range("B9").Value = "Data"
        .Range("C9").Value = "x0"
        .Range("D9").Value = "x1"
        .Range("E9").Value = "x2"
        .Range("F9").Value = "x3"
        .Range("G9").Value = "x4"
        .Range("H9").Value = "Logit"
        .Range("I9").Value = "Probability"
        .Range("B9:I9").Font.Bold = True
        .Range("B9:I9").HorizontalAlignment = xlCenter
        .Range("C10:G20").Interior.Color = Yellow
        For i = 10 To 20
            .Cells(i, 3).Value = 1
            .Cells(i, 8).FormulaR1C1 = "=b0*RC[-5]+b1*RC[-4]+b2*RC[-3]+b3*RC[-2]+b4*RC[-1]"
            .Cells(i, 9).FormulaR1C1 = "=1/(1+EXP(-RC[-1]))"
        Next i
        .Range("C10:C20").Interior.Color = RGB(226, 239, 218)
        .Range("A22").Value = "Yellow cells are inputs. x0 is the intercept column (1). Probability = 1 / (1 + EXP(−logit))."
        .Columns("A").ColumnWidth = 3
        .Columns("B").ColumnWidth = 12
        .Range("C:I").ColumnWidth = 12
        .Range("H10:I20").NumberFormat = "0.0000"
    End With
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("CreateLogitInputTemplate")
End Sub
