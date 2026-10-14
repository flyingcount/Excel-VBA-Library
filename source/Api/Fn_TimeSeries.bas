Attribute VB_Name = "Fn_TimeSeries"
Option Explicit

' Worksheet UDFs for Personal Menu27 (ACF, ACVF, PACF, Bartlett, Box-Pierce, Ljung-Box).
' Registered into Insert Function (category Excel VBA Lib) like Mat*.
' Unqualified =ACVF(...) is #NAME? until these are registered.

Private Const UdfCategory As String = "Excel VBA Lib"

''' @Description: Sample autocorrelation at lag k (population covariance, Var.P at lag 0).
''' @Example: =ACF(A1:A20, 1)
Public Function ACF(ByVal rng_Data As Range, ByVal int_Lag As Long) As Variant
    Application.Volatile
    ACF = modInternalTimeSeries.TsAcf(rng_Data, int_Lag)
End Function

''' @Description: Sample autocovariance at lag k (divisor n, same as Personal ACVF).
''' @Example: =ACVF(A1:A20, 1)
Public Function ACVF(ByVal rng_Data As Range, ByVal int_Lag As Long) As Variant
    Application.Volatile
    ACVF = modInternalTimeSeries.TsAcvf(rng_Data, int_Lag)
End Function

''' @Description: Partial autocorrelation at lag k (Yule-Walker).
''' @Example: =PACF(A1:A20, 2)
Public Function PACF(ByVal rng_Data As Range, ByVal int_Lag As Long) As Variant
    Application.Volatile
    PACF = modInternalTimeSeries.TsPacf(rng_Data, int_Lag)
End Function

''' @Description: Bartlett two-sided test text: |ACF| versus the normal critical value at alpha (default 0.05).
''' @Example: =Bartlett(A1:A20, 1)
Public Function Bartlett(ByVal rng_Data As Range, ByVal int_Lag As Long, Optional ByVal dbl_Alpha As Double = 0.05) As Variant
    Application.Volatile
    Bartlett = modInternalTimeSeries.TsBartlett(rng_Data, int_Lag, dbl_Alpha)
End Function

''' @Description: Box-Pierce Q statistic for lags 1..k (large-sample white-noise test).
''' @Example: =BoxPierce(A1:A20, 5)
Public Function BoxPierce(ByVal rng_Data As Range, ByVal int_Lag As Long) As Variant
    Application.Volatile
    BoxPierce = modInternalTimeSeries.TsBoxPierce(rng_Data, int_Lag)
End Function

''' @Description: Right-tail chi-square p-value of the Box-Pierce statistic.
''' @Example: =BoxPiercePVal(A1:A20, 5)
Public Function BoxPiercePVal(ByVal rng_Data As Range, ByVal int_Lag As Long, Optional ByVal dbl_Alpha As Double = 0.05) As Variant
    Application.Volatile
    BoxPiercePVal = modInternalTimeSeries.TsChiSqPVal(modInternalTimeSeries.TsBoxPierce(rng_Data, int_Lag), int_Lag)
End Function

''' @Description: Box-Pierce white-noise conclusion (reject H0 when p < alpha).
''' @Example: =BoxPierceTest(A1:A20, 5)
Public Function BoxPierceTest(ByVal rng_Data As Range, ByVal int_Lag As Long, Optional ByVal dbl_Alpha As Double = 0.05) As Variant
    Dim p As Variant
    Application.Volatile
    p = BoxPiercePVal(rng_Data, int_Lag, dbl_Alpha)
    BoxPierceTest = modInternalTimeSeries.TsWhiteNoiseTest(p, int_Lag, dbl_Alpha, "Box-Pierce")
End Function

''' @Description: Ljung-Box Q statistic for lags 1..k (small-sample white-noise test).
''' @Example: =LjungBox(A1:A20, 5)
Public Function LjungBox(ByVal rng_Data As Range, ByVal int_Lag As Long) As Variant
    Application.Volatile
    LjungBox = modInternalTimeSeries.TsLjungBox(rng_Data, int_Lag)
End Function

''' @Description: Right-tail chi-square p-value of the Ljung-Box statistic.
''' @Example: =LjungBoxPVal(A1:A20, 5)
Public Function LjungBoxPVal(ByVal rng_Data As Range, ByVal int_Lag As Long, Optional ByVal dbl_Alpha As Double = 0.05) As Variant
    Application.Volatile
    LjungBoxPVal = modInternalTimeSeries.TsChiSqPVal(modInternalTimeSeries.TsLjungBox(rng_Data, int_Lag), int_Lag)
End Function

''' @Description: Ljung-Box white-noise conclusion (reject H0 when p < alpha).
''' @Example: =LjungBoxTest(A1:A20, 5)
Public Function LjungBoxTest(ByVal rng_Data As Range, ByVal int_Lag As Long, Optional ByVal dbl_Alpha As Double = 0.05) As Variant
    Dim p As Variant
    Application.Volatile
    p = LjungBoxPVal(rng_Data, int_Lag, dbl_Alpha)
    LjungBoxTest = modInternalTimeSeries.TsWhiteNoiseTest(p, int_Lag, dbl_Alpha, "Ljung-Box")
End Function

Public Sub RegisterTimeSeriesUdfs()
    Call RegisterUdf("ACF", "Autocorrelation of a numeric column at lag k.", "data,lag")
    Call RegisterUdf("ACVF", "Autocovariance of a numeric column at lag k.", "data,lag")
    Call RegisterUdf("PACF", "Partial autocorrelation (Yule-Walker) at lag k.", "data,lag")
    Call RegisterUdf("Bartlett", "Bartlett test text for ACF at lag k.", "data,lag,alpha")
    Call RegisterUdf("BoxPierce", "Box-Pierce Q statistic for lags 1..k.", "data,lag")
    Call RegisterUdf("BoxPiercePVal", "Box-Pierce chi-square p-value.", "data,lag,alpha")
    Call RegisterUdf("BoxPierceTest", "Box-Pierce white-noise conclusion.", "data,lag,alpha")
    Call RegisterUdf("LjungBox", "Ljung-Box Q statistic for lags 1..k.", "data,lag")
    Call RegisterUdf("LjungBoxPVal", "Ljung-Box chi-square p-value.", "data,lag,alpha")
    Call RegisterUdf("LjungBoxTest", "Ljung-Box white-noise conclusion.", "data,lag,alpha")
End Sub

Private Sub RegisterUdf(ByVal procName As String, ByVal descr As String, ByVal argHelp As String)
    Dim qualified As String
    qualified = "'" & ThisWorkbook.Name & "'!" & procName
    On Error Resume Next
    Application.MacroOptions Macro:=qualified, Description:=descr, Category:=UdfCategory, ArgumentDescriptions:=Split(argHelp, ",")
    If Err.Number <> 0 Then
        Err.Clear
        Application.MacroOptions Macro:=qualified, Description:=descr, Category:=UdfCategory
    End If
    On Error GoTo 0
End Sub
