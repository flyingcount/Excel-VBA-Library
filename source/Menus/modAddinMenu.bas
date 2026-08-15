Attribute VB_Name = "modAddinMenu"
Option Explicit

' Add-in menu. Lives in the same .xlam as Api/Internal so OnAction names resolve
' without importing modules into the caller workbook.
'
' Excel does not run Auto_Open for an .xlam loaded as an add-in. Workbook_Open
' schedules InstallExcelVbaLibMenu via Application.OnTime (that name is not
' suppressed). Auto_Open remains as a manual entry point.

Private Const TagMenu As String = "ExcelVbaLibMenu"
Private Const MaxRetry As Long = 8
Private retryCount As Long

Public Sub Auto_Open()
    Call InstallExcelVbaLibMenu
End Sub

Public Sub Auto_Close()
    Call RemoveExcelVbaLibMenu
End Sub

' Must stay Public: Application.OnTime looks this up by name in the add-in.
Public Sub InstallExcelVbaLibMenu()
    On Error GoTo NeedRetry
    Call RemoveMenu
    If Not TryBuildMenu() Then GoTo NeedRetry
    On Error Resume Next
    Call Fn_MatricesRng.RegisterMatrixUdfs
    On Error GoTo 0
    retryCount = 0
    Exit Sub
NeedRetry:
    retryCount = retryCount + 1
    If retryCount > MaxRetry Then
        retryCount = 0
        Exit Sub
    End If
    On Error Resume Next
    Application.OnTime Now + TimeSerial(0, 0, 1), MenuProc("InstallExcelVbaLibMenu")
    On Error GoTo 0
End Sub

Public Sub RemoveExcelVbaLibMenu()
    retryCount = 0
    Call RemoveMenu
End Sub

' Run the installer when Excel is idle. CommandBars are often missing inside Workbook_Open.
Public Sub ScheduleMenuInstall()
    On Error Resume Next
    Application.OnTime Now, MenuProc("InstallExcelVbaLibMenu")
    If Err.Number <> 0 Then
        Err.Clear
        Call InstallExcelVbaLibMenu
    End If
    On Error GoTo 0
End Sub

Private Function MenuProc(ByVal proc As String) As String
    MenuProc = "'" & ThisWorkbook.Name & "'!" & proc
End Function

Private Function TryBuildMenu() As Boolean
    Dim bar As CommandBar
    Dim pop As CommandBarPopup
    Dim tmpl As CommandBarPopup
    Dim ben As CommandBarPopup
    Dim samp As CommandBarPopup
    Dim dataMenu As CommandBarPopup
    Dim dist As CommandBarPopup
    Dim mat As CommandBarPopup
    Dim matCreate As CommandBarPopup
    Dim matOps As CommandBarPopup
    Dim matDecomp As CommandBarPopup

    On Error Resume Next
    Set bar = Application.CommandBars("Worksheet Menu Bar")
    On Error GoTo 0
    If bar Is Nothing Then Exit Function

    On Error Resume Next
    Set pop = bar.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    On Error GoTo 0
    If pop Is Nothing Then Exit Function

    pop.Caption = "&Excel VBA Lib"
    pop.Tag = TagMenu

    Set ben = AddSubmenu(pop, "&Benford")
    Call AddButton(ben, "&First digit", "BenfordAnalysisFirstDigit")
    Call AddButton(ben, "&Second digit", "BenfordAnalysisSecondDigit")
    Call AddButton(ben, "&Third digit", "BenfordAnalysisThirdDigit")
    Call AddButton(ben, "First &two digits", "BenfordAnalysisTwoDigit")
    Call AddButton(ben, "First t&hree digits", "BenfordAnalysisThreeDigit")
    Call AddButton(ben, "&Last two digits", "BenfordAnalysisLastTwoDigit")

    Set dataMenu = AddSubmenu(pop, "&Data")
    Call AddButton(dataMenu, "&Combinations", "DataCombinations")
    Call AddButton(dataMenu, "Random &integers", "RandomIntegers")
    Call AddButton(dataMenu, "Random &numbers", "RandomNumbers")
    Call AddButton(dataMenu, "Random d&ates", "RandomDates")
    Call AddButton(dataMenu, "Random &strings", "RandomStrings")
    Call AddButton(dataMenu, "Random from &list", "RandomFromList")
    Call AddButton(dataMenu, "Random TRUE/&FALSE", "RandomTrueFalse")
    Call AddButton(dataMenu, "Random &Yes/No", "RandomYesNo")
    Call AddButton(dataMenu, "Random &1 or 0", "Random1or0")
    Call AddButton(dataMenu, "Yes/No &dataset", "CreateYesNoDataset")
    Call AddButton(dataMenu, "Random test data &types", "RandomTestDataTypes")
    Set dist = AddSubmenu(dataMenu, "&Probability distributions")
    Call AddButton(dist, "&Binomial", "RandomBinomialNumbers")
    Call AddButton(dist, "B&ernoulli", "RandomBernoulliNumbers")
    Call AddButton(dist, "&Normal", "RandomNormalNumbers")
    Call AddButton(dist, "&Poisson", "RandomPoissonNumbers")
    Call AddButton(dist, "&Exponential", "RandomExponentialNumbers")
    Call AddButton(dist, "&Gamma", "RandomGammaNumbers")
    Call AddButton(dist, "&Hypergeometric", "RandomHypergeometricNumbers")

    Set samp = AddSubmenu(pop, "&Sampling")
    Call AddButton(samp, "Extract sample of &rows", "ExtractSample")

    Set mat = AddSubmenu(pop, "&Matrices")
    Set matCreate = AddSubmenu(mat, "&Create")
    Call AddButton(matCreate, "&Identity", "MatrixCreateIdentity")
    Call AddButton(matCreate, "&Zeros", "MatrixCreateZeros")
    Call AddButton(matCreate, "&Ones", "MatrixCreateOnes")
    Call AddButton(matCreate, "&Diagonal from vector", "MatrixCreateDiagonal")
    Call AddButton(matCreate, "&Random", "MatrixCreateRandom")
    Call AddButton(matCreate, "&Hilbert", "MatrixCreateHilbert")
    Call AddButton(matCreate, "&Exchange", "MatrixCreateExchange")
    Call AddButton(matCreate, "&Toeplitz from vector", "MatrixCreateToeplitz")
    Call AddButton(matCreate, "&Vandermonde from vector", "MatrixCreateVandermonde")
    Call AddButton(matCreate, "&Companion from vector", "MatrixCreateCompanion")
    Set matOps = AddSubmenu(mat, "&Operations")
    Call AddButton(matOps, "&Transpose", "MatrixTranspose")
    Call AddButton(matOps, "&Add", "MatrixAdd")
    Call AddButton(matOps, "&Subtract", "MatrixSubtract")
    Call AddButton(matOps, "Sca&le", "MatrixScale")
    Call AddButton(matOps, "&Multiply", "MatrixMultiply")
    Call AddButton(matOps, "&Hadamard", "MatrixHadamard")
    Call AddButton(matOps, "&Kronecker", "MatrixKronecker")
    Call AddButton(matOps, "Outer &product", "MatrixOuter")
    Call AddButton(matOps, "D&ot product", "MatrixDot")
    Call AddButton(matOps, "&Inverse", "MatrixInverse")
    Call AddButton(matOps, "Ad&jugate", "MatrixAdjugate")
    Call AddButton(matOps, "Pseudoin&verse", "MatrixPseudoInverse")
    Call AddButton(matOps, "Po&wer", "MatrixPower")
    Call AddButton(matOps, "&Determinant", "MatrixDeterminant")
    Call AddButton(matOps, "T&race", "MatrixTrace")
    Call AddButton(matOps, "Extract d&iagonal", "MatrixDiagExtract")
    Call AddButton(matOps, "&Vec (column-major)", "MatrixVec")
    Call AddButton(matOps, "&Unvec", "MatrixUnvec")
    Call AddButton(matOps, "Ran&k", "MatrixRank")
    Call AddButton(matOps, "&Solve A X = B", "MatrixSolve")
    Call AddButton(matOps, "Frobenius &norm", "MatrixNorm")
    Call AddButton(matOps, "1-&norm", "MatrixNorm1")
    Call AddButton(matOps, "Infinit&y-norm", "MatrixNormInf")
    Call AddButton(matOps, "Co&factor matrix", "MatrixCofactor")
    Call AddButton(matOps, "M&inor", "MatrixMinor")
    Call AddButton(matOps, "Is s&ymmetric", "MatrixIsSymmetric")
    Call AddButton(matOps, "Is o&rthogonal", "MatrixIsOrthogonal")
    Set matDecomp = AddSubmenu(mat, "&Decompositions")
    Call AddButton(matDecomp, "&Cholesky", "MatrixCholesky")
    Call AddButton(matDecomp, "&Eigen (symmetric)", "MatrixEigen")
    Call AddButton(matDecomp, "&QR", "MatrixQR")
    Call AddButton(matDecomp, "&LU", "MatrixLU")

    Set tmpl = AddSubmenu(pop, "&Worksheet templates")
    Call AddButton(tmpl, "Notes workbook (all templates)", "CreateNotesWorkbook")
    Call AddButton(tmpl, "Links", "CreateLinksTemplate")
    Call AddButton(tmpl, "Actions", "CreateActionsTemplate")
    Call AddButton(tmpl, "Force field", "CreateForceFieldTemplate")
    Call AddButton(tmpl, "Assumptions", "CreateAssumptionsTemplate")
    Call AddButton(tmpl, "Questions", "CreateQuestionsTemplate")
    Call AddButton(tmpl, "Python packages", "ImportPythonPackages")

    TryBuildMenu = True
End Function

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
