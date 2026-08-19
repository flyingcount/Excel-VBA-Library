Attribute VB_Name = "modApiTimeSeries"
Option Explicit

' Public API: time-series analysis and lag differencing (Personal Custom_Menu27_*).
' Other workbooks / the add-in menu should call these names only.
' Personal OnAction names are kept.
' GeneratePACFList / GenerateAVCFList / GenerateAutoCovarianceMatrix are not on the Personal menu.

''' @Description: ACF, ACVF, PACF, Bartlett, Box-Pierce and Ljung-Box for a numeric column, written as values on sheet Time Series.
''' @Example: TimeSeriesAnalysis
Public Sub TimeSeriesAnalysis()
    Dim src As Range
    Set src = modInternalTimeSeries.PromptNumericColumn("Select a single column of time-series values.", 3)
    If src Is Nothing Then Exit Sub
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    Call modInternalTimeSeries.WriteTimeSeriesAnalysis(src, False)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("TimeSeriesAnalysis")
End Sub

''' @Description: Same layout as TimeSeriesAnalysis, with live worksheet formulas that call the add-in UDFs.
''' @Example: TimeSeriesAnalysisFormula
Public Sub TimeSeriesAnalysisFormula()
    Dim src As Range
    Set src = modInternalTimeSeries.PromptNumericColumn("Select a single column of time-series values.", 4)
    If src Is Nothing Then Exit Sub
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    Call Fn_TimeSeries.RegisterTimeSeriesUdfs
    Call modInternalTimeSeries.WriteTimeSeriesAnalysis(src, True)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("TimeSeriesAnalysisFormula")
End Sub

''' @Description: Lag-1, lag-2, ... differences of a numeric column on sheet Date Diff, with slope/R² and line charts.
''' @Example: DateDifferencing
Public Sub DateDifferencing()
    Dim src As Range
    Set src = modInternalTimeSeries.PromptNumericColumn("Select a single column of time-series values.", 2)
    If src Is Nothing Then Exit Sub
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    Call modInternalTimeSeries.WriteDateDifferencing(src, False)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("DateDifferencing")
End Sub

''' @Description: Same as DateDifferencing on sheet Date Diff Formula, with difference columns and an ACF table as live formulas. Lag-1 ACF < -0.5 flags over-differencing.
''' @Example: DateDifferencingFormulae
Public Sub DateDifferencingFormulae()
    Dim src As Range
    Set src = modInternalTimeSeries.PromptNumericColumn("Select a single column of time-series values.", 2)
    If src Is Nothing Then Exit Sub
    On Error GoTo EH
    Call modInternalExcelApp.PushAppState
    Call Fn_TimeSeries.RegisterTimeSeriesUdfs
    Call modInternalTimeSeries.WriteDateDifferencing(src, True)
    Call modInternalExcelApp.PopAppState
    Exit Sub
EH:
    Call modInternalExcelApp.PopAppState
    Call modInternalError.RaiseCurrent("DateDifferencingFormulae")
End Sub
