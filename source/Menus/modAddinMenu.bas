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
    Call Fn_TimeSeries.RegisterTimeSeriesUdfs
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
    Dim anMenu As CommandBarPopup
    Dim dist As CommandBarPopup
    Dim prep As CommandBarPopup
    Dim clMenu As CommandBarPopup
    Dim rngMenu As CommandBarPopup
    Dim rngCreate As CommandBarPopup
    Dim rngAn As CommandBarPopup
    Dim rngCl As CommandBarPopup
    Dim dupMenu As CommandBarPopup
    Dim protMenu As CommandBarPopup
    Dim recMenu As CommandBarPopup
    Dim tblMenu As CommandBarPopup
    Dim filesMenu As CommandBarPopup
    Dim hlMenu As CommandBarPopup
    Dim pqMenu As CommandBarPopup
    Dim mat As CommandBarPopup
    Dim matCreate As CommandBarPopup
    Dim matOps As CommandBarPopup
    Dim matProps As CommandBarPopup
    Dim matValid As CommandBarPopup
    Dim matDecomp As CommandBarPopup
    Dim plots As CommandBarPopup
    Dim distPlots As CommandBarPopup
    Dim qqPlots As CommandBarPopup
    Dim tsMenu As CommandBarPopup

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

    Set prep = AddSubmenu(pop, "Data &Preprocessing")
    Call AddButton(prep, "&Standardise columns", "ScalingStandard")
    Call AddButton(prep, "&Normalise columns", "ScalingNormalise")
    Call AddButton(prep, "&Robust scale columns", "ScalingRobust")
    Call AddButton(prep, "Convert categorical data to dummy variable matrix", "DummyVariablesForMachineLearning")

    Set clMenu = AddSubmenu(pop, "&Custom lists")
    Call AddButton(clMenu, "&Count", "CountCustomLists")
    Call AddButton(clMenu, "&List on worksheet", "ShowCustomLists")
    Call AddButton(clMenu, "Create from &columns", "CreateCustomListByColumn")
    Call AddButton(clMenu, "Create from &rows", "CreateCustomListByRow")
    Call AddButton(clMenu, "&Delete by list number", "DeleteCustomList")
    Call AddButton(clMenu, "&AutoCorrect list", "AutoCorrectEntries_Display", True)
    Call AddButton(clMenu, "&Add AutoCorrect entries", "AutoCorrectEntries_Add")

    Set rngMenu = AddSubmenu(pop, "&Ranges")
    Call AddButton(rngMenu, "&List named range properties", "ListNamedRangeProperties")
    Call AddButton(rngMenu, "&Paste as picture", "PasteRangeAsPicture")
    Call AddButton(rngMenu, "&Colour named ranges", "HighlightRanges")
    Call AddButton(rngMenu, "&Remove named-range colour", "DeHighlightRanges")
    Call AddButton(rngMenu, "&Delete specified names", "DeleteNamedRanges", True)
    Set rngCreate = AddSubmenu(rngMenu, "&Create named ranges")
    Call AddButton(rngCreate, "Hidden workbook-scope &constants", "HiddenMasterDataWorkbookScope")
    Call AddButton(rngCreate, "Hidden &worksheet-scope constants", "HiddenMasterDataWorksheetScope")
    Call AddButton(rngCreate, "Name each &column from header", "NamedRangeIntoNamedRangeColumns")
    Call AddButton(rngCreate, "Hidden named &range", "CreateHiddenNamedRange")
    Call AddButton(rngCreate, "Hidden named &string", "CreateHiddenNamedString")
    Call AddButton(rngCreate, "Hidden named &number", "CreateHiddenNamedNumber")
    Call AddButton(rngCreate, "Common c&onstants (thousand / million)", "NameCreateConstantsAsNamedRanges")
    Call AddButton(rngCreate, "Local name on &every sheet", "CreateNamedRangeInAllWorksheets")
    Call AddButton(rngMenu, "&Hide all names (workbook)", "HideAllNamedRanges", True)
    Call AddButton(rngMenu, "&Unhide all names (workbook)", "UnhideAllNamedRanges")
    Call AddButton(rngMenu, "Unhide names on active &sheet", "UnhideNamedRangesWorksheet", True)
    Call AddButton(rngMenu, "Hide names on active s&heet", "HideNamedRangesWorksheet")
    Call AddButton(rngMenu, "Unhide &specified names", "UnhideSpecifiedNamedRanges", True)
    Call AddButton(rngMenu, "Hide spec&ified names", "HideSpecifiedNamedRanges")
    Call AddButton(rngMenu, "List &unique values", "ListUniqueValues")
    Call AddButton(rngMenu, "Scope: worksheet to &workbook", "ConvertNamedRangeLocalToGlobalScope", True)
    Call AddButton(rngMenu, "Scope: workbook to work&sheet", "ConvertNamedRangeGlodalToLocalScope")
    Set rngAn = AddSubmenu(rngMenu, "&Analyse range")
    Call AddButton(rngAn, "Analyse to a &sheet", "RangeAnalysis")
    Call AddButton(rngAn, "Analyse in a &message box", "RangeAnalysisMessage")
    Call AddButton(rngAn, "Character &frequency", "FrequencyAnalysis")
    Set rngCl = AddSubmenu(rngMenu, "C&leanse range")
    Call AddButton(rngCl, "Keep &digits", "Remove_AlphaCharactersFromString")
    Call AddButton(rngCl, "Keep &letters", "Remove_NumbersFromString")
    Call AddButton(rngCl, "Keep letters, digits, &spaces", "Remove_SpecialCharactersFromString")
    Call AddButton(rngCl, "&Trim whitespace", "ClearCellsThatOnlyContainWhitespaces")
    Call AddButton(rngCl, "Trim whitespace (&regex)", "RemoveWhiteSpaces")
    Call AddButton(rngCl, "List unique &characters", "ListCharactersAndCodesInRange")
    Call AddButton(rngMenu, "&Transpose to the right", "TransposeARange", True)
    Call AddButton(rngMenu, "S&plit into named columns", "SplitRangeAndNameEachColumn")

    Set dupMenu = AddSubmenu(pop, "D&uplicates")
    Call AddButton(dupMenu, "&Flag duplicates (single column)", "FlagDuplicates")
    Call AddButton(dupMenu, "&Unique references for duplicates", "ReferenceDuplicates")
    Call AddButton(dupMenu, "&Colour duplicate groups", "ColorDuplicates")
    Call AddButton(dupMenu, "Colour duplicates by &row", "ColourDuplicateValuesByRow", True)
    Call AddButton(dupMenu, "Colour duplicates by &column", "ColourDuplicateValuesByColumn")
    Call AddButton(dupMenu, "Colour duplicates in &selection", "ColourDuplicateValuesInSelection")
    Call AddButton(dupMenu, "C&ount duplicates in selection", "DuplicateCountFromSelection", True)

    Set protMenu = AddSubmenu(pop, "P&rotection")
    Call AddButton(protMenu, "&Limit scroll area to selection", "LimitScrollArea")
    Call AddButton(protMenu, "&Reset scroll area", "ResetScrollArea")
    Call AddButton(protMenu, "&Protect worksheet (default password)", "ProtectWorksheet", True)
    Call AddButton(protMenu, "&Unprotect worksheet (default password)", "UnProtectWorksheet")
    Call AddButton(protMenu, "&Show default password", "DisplayPassword")
    Call AddButton(protMenu, "Unhide all &worksheets", "UnHideAllSheets", True)
    Call AddButton(protMenu, "Unhide all &rows and columns", "UnhideAllRowsAndColumns")

    Set recMenu = AddSubmenu(pop, "Reconcili&ations")
    Call AddButton(recMenu, "Reconcile two columns of &numbers", "ReconcileTwoColumns")
    Call AddButton(recMenu, "Reconcile two columns of &strings", "ReconcileTwoColumnsStrings")
    Call AddButton(recMenu, "&Compare two ranges", "CompareRanges")

    Set tblMenu = AddSubmenu(pop, "&Tables")
    Call AddButton(tblMenu, "&List table properties", "ListTableProperties")
    Call AddButton(tblMenu, "List tables in a &message box", "ShowAllTablesInWorkbook")

    Set filesMenu = AddSubmenu(pop, "&Files")
    Call AddButton(filesMenu, "List files in &folder", "ReturnFilesInSelectedFolder")
    Call AddButton(filesMenu, "List files in folder and &subfolders", "BatchListAllFiles_FolderSubfolders")

    Set samp = AddSubmenu(pop, "&Sampling")
    Call AddButton(samp, "Extract sample of &rows", "ExtractSample")

    Set hlMenu = AddSubmenu(pop, "&Hyperlinks")
    Call AddButton(hlMenu, "&Inventory", "HyperlinkInventory")
    Call AddButton(hlMenu, "Create &index", "CreateIndex")
    Call AddButton(hlMenu, "&Update index", "UpdateIndex")
    Call AddButton(hlMenu, "List &worksheets", "ShowAllWorksheetsInWorkbook")
    Call AddButton(hlMenu, "Remove by &text to display", "RemovingHyperLink", True)
    Call AddButton(hlMenu, "&Open selected", "OpenHyperlink")
    Call AddButton(hlMenu, "&Back-links to this sheet A1", "AddHyperlinksToCurrentSheetA1")

    Set pqMenu = AddSubmenu(pop, "&Power Query")
    Call AddButton(pqMenu, "&Import or export queries and functions", "ShowPQLibraryForm")
    Call AddButton(pqMenu, "Toggle &background refresh", "BackgroundRefreshToggle")
    Call AddButton(pqMenu, "Toggle &privacy (Fast Combine)", "IgnorePrivacyToggle")
    Call AddButton(pqMenu, "&Connect all tables", "Add_Connection_All_Tables")

    Set mat = AddSubmenu(pop, "&Matrices")
    Call AddButton(mat, "Matrix &diagnostic", "MatrixDiagnosticMessage")
    Set matCreate = AddSubmenu(mat, "&Create")
    Call AddButton(matCreate, "&Identity", "MatrixCreateIdentity")
    Call AddButton(matCreate, "&Zeros", "MatrixCreateZeros")
    Call AddButton(matCreate, "&Ones", "MatrixCreateOnes")
    Call AddButton(matCreate, "&Hilbert", "MatrixCreateHilbert")
    Call AddButton(matCreate, "&Exchange", "MatrixCreateExchange")
    Call AddButton(matCreate, "&Anti-diagonal", "CreateAntiDiagonalMatrix")
    Call AddButton(matCreate, "Sca&ling matrix", "MtrxScaling")
    Call AddButton(matCreate, "S&tretch matrices", "MtrxStretch")
    Call AddButton(matCreate, "Random d&iagonal", "CreateRandomDiagonalMatrix", True)
    Call AddButton(matCreate, "Random s&ymmetric", "CreateRandomSymmetricMatrix")
    Call AddButton(matCreate, "&Random", "MatrixCreateRandom")
    Call AddButton(matCreate, "Random squar&e", "MatrixCreateRandomSquare")
    Call AddButton(matCreate, "&Upper triangular", "MtrxTriangularUpper", True)
    Call AddButton(matCreate, "&Lower triangular", "MtrxTriangularLower")
    Call AddButton(matCreate, "R&otation (all)", "MtrxRotAll", True)
    Call AddButton(matCreate, "Rotation &X", "MtrxRotX")
    Call AddButton(matCreate, "Rotation &Y", "MtrxRotY")
    Call AddButton(matCreate, "Rotation &Z", "MtrxRotZ")
    Call AddButton(matCreate, "&Diagonal from vector", "MatrixCreateDiagonal", True)
    Call AddButton(matCreate, "&Toeplitz from vector", "MatrixCreateToeplitz")
    Call AddButton(matCreate, "&Vandermonde from vector", "MatrixCreateVandermonde")
    Call AddButton(matCreate, "&Companion from vector", "MatrixCreateCompanion")
    Set matOps = AddSubmenu(mat, "&Operations")
    Call AddButton(matOps, "&Transpose", "MatrixTranspose")
    Call AddButton(matOps, "Transpose with &formulae", "TransposeMatrixFormulae")
    Call AddButton(matOps, "&Add", "MatrixAdd")
    Call AddButton(matOps, "&Subtract", "MatrixSubtract")
    Call AddButton(matOps, "Sca&le", "MatrixScale")
    Call AddButton(matOps, "Scale with f&ormulae", "MtrxMultScalarFormulae")
    Call AddButton(matOps, "&Multiply", "MatrixMultiply")
    Call AddButton(matOps, "Multiply with for&mulae", "MtrxMultFormulae")
    Call AddButton(matOps, "Multiplication-Hadamard", "MatrixMultiplicationHadamard")
    Call AddButton(matOps, "Multiplication-Kronecker", "MatrixMultiplicationKronecker")
    Call AddButton(matOps, "Outer &product", "MatrixOuter")
    Call AddButton(matOps, "D&ot product", "MatrixDot")
    Call AddButton(matOps, "&Inverse", "MatrixInverse")
    Call AddButton(matOps, "Inverse with form&ulae", "MtrxInverse")
    Call AddButton(matOps, "Ad&jugate", "MatrixAdjugate")
    Call AddButton(matOps, "Pseudoin&verse", "MatrixPseudoInverse")
    Call AddButton(matOps, "Po&wer", "MatrixPower")
    Call AddButton(matOps, "Extract d&iagonal", "MatrixDiagExtract")
    Call AddButton(matOps, "&Vec (column-major)", "MatrixVec")
    Call AddButton(matOps, "&Unvec", "MatrixUnvec")
    Call AddButton(matOps, "Co&factor matrix", "MatrixCofactor")
    Call AddButton(matOps, "M&inor", "MatrixMinor")
    Call AddButton(matOps, "&Solve A X = B", "MatrixSolve", True)
    Call AddButton(matOps, "&Covariance", "MatrixCovariance", True)
    Call AddButton(matOps, "&Standardised covariance", "MatrixCovarianceStandardise")
    Set matProps = AddSubmenu(mat, "&Properties")
    Call AddButton(matProps, "&All properties", "MatrixPropertiesAll")
    Call AddButton(matProps, "&Determinant", "MatrixDeterminant", True)
    Call AddButton(matProps, "T&race", "MatrixTrace")
    Call AddButton(matProps, "Ran&k", "MatrixRank")
    Call AddButton(matProps, "Frobenius &norm", "MatrixNorm")
    Call AddButton(matProps, "1-&norm", "MatrixNorm1")
    Call AddButton(matProps, "Infinit&y-norm", "MatrixNormInf")
    Call AddButton(matProps, "&Condition number", "MatrixConditionNumber")
    Call AddButton(matProps, "&Spectral radius", "MatrixSpectralRadius")
    Call AddButton(matProps, "&Eigenvalues", "MatrixEigenvalues", True)
    Call AddButton(matProps, "Eigen&vectors", "MatrixEigenvectors")
    Set matValid = AddSubmenu(mat, "&Validation")
    Call AddButton(matValid, "&Size (rows x columns)", "MatrixWriteSize")
    Call AddButton(matValid, "Is s&ymmetric", "MatrixIsSymmetric")
    Call AddButton(matValid, "Is o&rthogonal", "MatrixIsOrthogonal")
    Call AddButton(matValid, "Hadamard &Proof", "MatrixHadamardProof")
    Set matDecomp = AddSubmenu(mat, "&Decompositions")
    Call AddButton(matDecomp, "&Cholesky", "MatrixCholesky")
    Call AddButton(matDecomp, "Cholesky (&sheet)", "CholeskyDecompositionOutput")
    Call AddButton(matDecomp, "&Eigen (symmetric)", "MatrixEigen")
    Call AddButton(matDecomp, "Eigen (&sheet)", "EigenDecompositionSymmetricMatrix")
    Call AddButton(matDecomp, "Eigenvalue &diagonal", "DiagonalEigenvalueSymmetricMatrix")
    Call AddButton(matDecomp, "&QR", "MatrixQR")
    Call AddButton(matDecomp, "&LU", "MatrixLU")
    Call AddButton(matDecomp, "&SVD", "SVD")

    Set anMenu = AddSubmenu(pop, "Anal&ysis")
    Call AddButton(anMenu, "Calculate S&VD", "SVD")
    Call AddButton(anMenu, "Solve AX=&B", "LinearSystem_AXB_v2")
    Call AddButton(anMenu, "Confusion matrix (Yes/&No)", "ConfusionMatrix", True)
    Call AddButton(anMenu, "Confusion matrix (&1s and 0s)", "ConfusionMatrixOnesAndZeros")
    Call AddButton(anMenu, "Confusion matrix &template", "ConfusionMatrixTemplate")
    Call AddButton(anMenu, "&Variance-covariance", "CalculateVarianceCovarianceMatrix", True)
    Call AddButton(anMenu, "&Standardised covariance", "MatrixCovarianceStandardise")
    Call AddButton(anMenu, "&Correlation matrix", "CorrelationMatrix")
    Call AddButton(anMenu, "&Prove var-covar and correlation", "ProveVarCovarAndCorrel")
    Call AddButton(anMenu, "&Mean vector", "CalculateMeanVector", True)
    Call AddButton(anMenu, "Stdev (&population) vector", "CalculateStandardDeviationPopulationVector")
    Call AddButton(anMenu, "Stdev (s&ample) vector", "CalculateStandardDeviationSampleVector")
    Call AddButton(anMenu, "Stdev &product (population)", "CalculateStdDevProductMatrixPopulation")
    Call AddButton(anMenu, "&Residuals analysis", "ResidualsAnalysis", True)
    Call AddButton(anMenu, "&Bland-Altman plot", "BlandAltmanPlot", True)
    Call AddButton(anMenu, "&Deming regression", "CalculateDemingRegression")
    Call AddButton(anMenu, "Box-&Cox transformations", "BoxCox", True)
    Call AddButton(anMenu, "&Logit input template", "CreateLogitInputTemplate")

    Set plots = AddSubmenu(pop, "&Plots Charts")
    Call AddButton(plots, "&Histogram and data table", "HistogramTableAndPlot")
    Call AddButton(plots, "Histogram with &formulae", "HistogramFormulaeAndPlot")
    Call AddButton(plots, "&Linear regression", "LinearRegression", True)
    Call AddButton(plots, "Linear regression &v2", "LinearRegressionV2")
    Call AddButton(plots, "P&rocess capability", "ProcessCapabilityChart", True)
    Call AddButton(plots, "&Binomial", "GenerateBinomialPlot", True)
    Set distPlots = AddSubmenu(plots, "&Parametric data and plots")
    Call AddButton(distPlots, "&Normal", "GenerateNormalPlot")
    Call AddButton(distPlots, "&Log-normal", "GenerateLogNormalPlot")
    Call AddButton(distPlots, "&Poisson", "GeneratePoissonPlot")
    Call AddButton(distPlots, "&Weibull", "GenerateWeibullPlot")
    Call AddButton(distPlots, "&Gamma", "GenerateGammaPlot")
    Call AddButton(distPlots, "B&eta", "GenerateBetaPlot")
    Call AddButton(distPlots, "&Exponential", "GenerateExponentialPlot")
    Call AddButton(distPlots, "&Hypergeometric", "GenerateHypergeometricPlot")
    Call AddButton(distPlots, "L&ogistic curve", "GenerateLogisticCurve")
    Set qqPlots = AddSubmenu(plots, "&QQ plots")
    Call AddButton(qqPlots, "&Gaussian / normal", "QQPlotGaussianNormal")
    Call AddButton(qqPlots, "&Uniform", "QQPlotUniform")
    Call AddButton(plots, "Lorenz curve and &Gini", "GiniPlot", True)
    Call AddButton(plots, "&Autocorrelation (ACF)", "GenerateCorrelogram", True)
    Call AddButton(plots, "&Diebold-Mariano test", "DieboldMarianoTest")
    Call AddButton(plots, "&XmR chart", "XmR", True)
    Call AddButton(plots, "XmR &diagnostics", "GenerateXMRDiagnostics")
    Call AddButton(plots, "Line chart on a chart &sheet", "PlotLineChartSheet", True)

    Set tsMenu = AddSubmenu(pop, "&Time series")
    Call AddButton(tsMenu, "Time series &analysis", "TimeSeriesAnalysis")
    Call AddButton(tsMenu, "Time series analysis with &formulae", "TimeSeriesAnalysisFormula")
    Call AddButton(tsMenu, "&Date differencing", "DateDifferencing", True)
    Call AddButton(tsMenu, "Date differencing with f&ormulae", "DateDifferencingFormulae")

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

Private Sub AddButton(ByVal Parent As CommandBarPopup, ByVal Caption As String, ByVal MacroName As String, Optional ByVal BeginGroup As Boolean = False)
    Dim btn As CommandBarButton
    Set btn = Parent.Controls.Add(Type:=msoControlButton, Temporary:=True)
    btn.Caption = Caption
    btn.OnAction = MacroName
    If BeginGroup Then btn.BeginGroup = True
End Sub

Private Sub RemoveMenu()
    Dim c As CommandBarControl
    On Error Resume Next
    For Each c In Application.CommandBars("Worksheet Menu Bar").Controls
        If c.Tag = TagMenu Then c.Delete
    Next c
    On Error GoTo 0
End Sub
