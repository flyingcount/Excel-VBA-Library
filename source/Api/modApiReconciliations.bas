Attribute VB_Name = "modApiReconciliations"
Option Explicit

' Public API: two-column recon and range compare (Personal Custom_Menu20_*).
' Other workbooks / the add-in menu should call these names only.
' Personal OnAction names are kept.
' CompareTwoRangesForm is not included (CompareRanges uses InputBoxes).

''' @Description: Match two single-column numeric ranges 1-for-1 onto a Reconciliation sheet (matched flags, unmatched statement, totals).
''' @Example: ReconcileTwoColumns
Public Sub ReconcileTwoColumns()
    Dim rngA As Range
    Dim rngB As Range
    Dim nameA As Variant
    Dim nameB As Variant
    Set rngA = modInternalReconciliations.PromptReconRange("Select the first column of numbers.")
    If rngA Is Nothing Then Exit Sub
    If Not modInternalReconciliations.RequireSingleColumn(rngA, "First range") Then Exit Sub
    nameA = modInternalRanges.PromptString("Enter the first dataset name.", "First dataset", "Reconciliations")
    If IsEmpty(nameA) Then Exit Sub
    Set rngB = modInternalReconciliations.PromptReconRange("Select the second column of numbers.")
    If rngB Is Nothing Then Exit Sub
    If Not modInternalReconciliations.RequireSingleColumn(rngB, "Second range") Then Exit Sub
    nameB = modInternalRanges.PromptString("Enter the second dataset name.", "Second dataset", "Reconciliations")
    If IsEmpty(nameB) Then Exit Sub
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    Call modInternalReconciliations.ReconcileNumericColumns(rngA, CStr(nameA), rngB, CStr(nameB))
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("ReconcileTwoColumns")
End Sub

''' @Description: Match two single-column text ranges 1-for-1 onto a Reconciliation Strings sheet (matched flags; unmatched listed as not in the other set).
''' @Example: ReconcileTwoColumnsStrings
Public Sub ReconcileTwoColumnsStrings()
    Dim rngA As Range
    Dim rngB As Range
    Dim nameA As Variant
    Dim nameB As Variant
    Set rngA = modInternalReconciliations.PromptReconRange("Select the first column of text.")
    If rngA Is Nothing Then Exit Sub
    If Not modInternalReconciliations.RequireSingleColumn(rngA, "First range") Then Exit Sub
    nameA = modInternalRanges.PromptString("Enter the first dataset name.", "First dataset", "Reconciliations")
    If IsEmpty(nameA) Then Exit Sub
    Set rngB = modInternalReconciliations.PromptReconRange("Select the second column of text.")
    If rngB Is Nothing Then Exit Sub
    If Not modInternalReconciliations.RequireSingleColumn(rngB, "Second range") Then Exit Sub
    nameB = modInternalRanges.PromptString("Enter the second dataset name.", "Second dataset", "Reconciliations")
    If IsEmpty(nameB) Then Exit Sub
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    Call modInternalReconciliations.ReconcileStringColumns(rngA, CStr(nameA), rngB, CStr(nameB))
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("ReconcileTwoColumnsStrings")
End Sub

''' @Description: Compare two same-size ranges cell by cell. Differences are filled yellow on both ranges and listed on sheet Range Comparison.
''' @Example: CompareRanges
Public Sub CompareRanges()
    Dim rngA As Range
    Dim rngB As Range
    Dim nDiff As Long
    Set rngA = modInternalReconciliations.PromptReconRange("Select range 1.")
    If rngA Is Nothing Then Exit Sub
    Set rngB = modInternalReconciliations.PromptReconRange("Select range 2.")
    If rngB Is Nothing Then Exit Sub
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    nDiff = modInternalReconciliations.CompareTwoRanges(rngA, rngB)
    Call modInternalExcelApp.PopAppState
    If nDiff < 0 Then Exit Sub
    MsgBox CStr(nDiff) & " differences found.", vbInformation, "Reconciliations"
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("CompareRanges")
End Sub
