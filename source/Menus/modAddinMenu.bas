Attribute VB_Name = "modAddinMenu"
Option Explicit

' Add-in menu (Auto_Open / Auto_Close). Lives in the same .xlam as Api/Internal
' so OnAction names resolve without importing modules into the caller workbook.

Private Const TagMenu As String = "ExcelVbaLibMenu"

Public Sub Auto_Open()
    Call RemoveMenu
    Call BuildMenu
End Sub

Public Sub Auto_Close()
    Call RemoveMenu
End Sub

Private Sub BuildMenu()
    Dim pop As CommandBarPopup
    Dim tmpl As CommandBarPopup
    Dim ben As CommandBarPopup
    Dim samp As CommandBarPopup

    On Error Resume Next
    Set pop = Application.CommandBars("Worksheet Menu Bar").Controls.Add( _
        Type:=msoControlPopup, Temporary:=True)
    On Error GoTo 0
    If pop Is Nothing Then Exit Sub

    pop.Caption = "&Excel VBA Lib"
    pop.Tag = TagMenu

    Set ben = AddSubmenu(pop, "&Benford")
    Call AddButton(ben, "&First digit", "BenfordAnalysisFirstDigit")
    Call AddButton(ben, "&Second digit", "BenfordAnalysisSecondDigit")
    Call AddButton(ben, "&Third digit", "BenfordAnalysisThirdDigit")
    Call AddButton(ben, "First &two digits", "BenfordAnalysisTwoDigit")
    Call AddButton(ben, "First t&hree digits", "BenfordAnalysisThreeDigit")
    Call AddButton(ben, "&Last two digits", "BenfordAnalysisLastTwoDigit")

    Set samp = AddSubmenu(pop, "&Sampling")
    Call AddButton(samp, "Extract sample of &rows", "ExtractSample")

    Set tmpl = AddSubmenu(pop, "&Worksheet templates")
    Call AddButton(tmpl, "Notes workbook (all templates)", "CreateNotesWorkbook")
    Call AddButton(tmpl, "Links", "CreateLinksTemplate")
    Call AddButton(tmpl, "Actions", "CreateActionsTemplate")
    Call AddButton(tmpl, "Force field", "CreateForceFieldTemplate")
    Call AddButton(tmpl, "Assumptions", "CreateAssumptionsTemplate")
    Call AddButton(tmpl, "Questions", "CreateQuestionsTemplate")
    Call AddButton(tmpl, "Python packages", "ImportPythonPackages")
End Sub

Private Function AddSubmenu(ByVal Parent As CommandBarPopup, ByVal Caption As String) As CommandBarPopup
    Dim pop As CommandBarPopup
    Set pop = Parent.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    pop.Caption = Caption
    Set AddSubmenu = pop
End Function

Private Sub AddButton(ByVal Parent As CommandBarPopup, ByVal Caption As String, ByVal MacroName As String)
    Dim btn As CommandBarButton
    Set btn = Parent.Controls.Add(Type:=msoControlButton, Temporary:=True)
    btn.Caption = Caption
    btn.OnAction = MacroName
End Sub

Private Sub RemoveMenu()
    Dim c As CommandBarControl
    On Error Resume Next
    For Each c In Application.CommandBars("Worksheet Menu Bar").Controls
        If c.Tag = TagMenu Then c.Delete
    Next c
    On Error GoTo 0
End Sub
