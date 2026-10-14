Attribute VB_Name = "Fn_MatricesArray"
Option Explicit

' Personal Fn_MatricesArray: constructor UDFs (no input range).

Public Function MatIdentity(ByVal n As Long) As Variant
    On Error GoTo Fail
    MatIdentity = modInternalMatrices.Identity(n)
    Exit Function
Fail:
    MatIdentity = CVErr(xlErrValue)
End Function

Public Function MatZeros(ByVal n As Long, Optional ByVal m As Variant) As Variant
    On Error GoTo Fail
    If IsMissing(m) Then
        MatZeros = modInternalMatrices.Filled(n, n, 0)
    Else
        MatZeros = modInternalMatrices.Filled(n, CLng(m), 0)
    End If
    Exit Function
Fail:
    MatZeros = CVErr(xlErrValue)
End Function

Public Function MatOnes(ByVal n As Long, Optional ByVal m As Variant) As Variant
    On Error GoTo Fail
    If IsMissing(m) Then
        MatOnes = modInternalMatrices.Filled(n, n, 1)
    Else
        MatOnes = modInternalMatrices.Filled(n, CLng(m), 1)
    End If
    Exit Function
Fail:
    MatOnes = CVErr(xlErrValue)
End Function

Public Function MatHilbert(ByVal n As Long) As Variant
    On Error GoTo Fail
    MatHilbert = modInternalMatrices.Hilbert(n)
    Exit Function
Fail:
    MatHilbert = CVErr(xlErrValue)
End Function

Public Function MatExchange(ByVal n As Long) As Variant
    On Error GoTo Fail
    MatExchange = modInternalMatrices.ExchangeMat(n)
    Exit Function
Fail:
    MatExchange = CVErr(xlErrValue)
End Function
