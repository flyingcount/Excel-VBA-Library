Attribute VB_Name = "Fn_Matrices2"
Option Explicit

' Personal Fn_Matrices2: decompositions, norms, rank, adjugate, cofactor, minor.

Public Function MatChol(ByVal rng As Range) As Variant
    On Error GoTo Fail
    MatChol = modInternalMatrices.Cholesky(modInternalMatrices.RangeToMatrix(rng))
    Exit Function
Fail:
    MatChol = CVErr(xlErrNum)
End Function

Public Function MatQR(ByVal rng As Range) As Variant
    On Error GoTo Fail
    MatQR = modInternalMatrices.QRFactors(modInternalMatrices.RangeToMatrix(rng))
    Exit Function
Fail:
    MatQR = CVErr(xlErrNum)
End Function

Public Function MatEigen(ByVal rng As Range) As Variant
    On Error GoTo Fail
    MatEigen = modInternalMatrices.EigenSymmetric(modInternalMatrices.RangeToMatrix(rng))
    Exit Function
Fail:
    MatEigen = CVErr(xlErrNum)
End Function

Public Function MatLU(ByVal rng As Range) As Variant
    On Error GoTo Fail
    MatLU = modInternalMatrices.LUFactors(modInternalMatrices.RangeToMatrix(rng))
    Exit Function
Fail:
    MatLU = CVErr(xlErrNum)
End Function

Public Function MatNormF(ByVal rng As Range) As Variant
    On Error GoTo Fail
    MatNormF = modInternalMatrices.FrobeniusNorm(modInternalMatrices.RangeToMatrix(rng))
    Exit Function
Fail:
    MatNormF = CVErr(xlErrValue)
End Function

Public Function MatNorm1(ByVal rng As Range) As Variant
    On Error GoTo Fail
    MatNorm1 = modInternalMatrices.Norm1(modInternalMatrices.RangeToMatrix(rng))
    Exit Function
Fail:
    MatNorm1 = CVErr(xlErrValue)
End Function

Public Function MatNormInf(ByVal rng As Range) As Variant
    On Error GoTo Fail
    MatNormInf = modInternalMatrices.NormInf(modInternalMatrices.RangeToMatrix(rng))
    Exit Function
Fail:
    MatNormInf = CVErr(xlErrValue)
End Function

Public Function MatRank(ByVal rng As Range) As Variant
    On Error GoTo Fail
    MatRank = modInternalMatrices.RankOf(modInternalMatrices.RangeToMatrix(rng))
    Exit Function
Fail:
    MatRank = CVErr(xlErrValue)
End Function

Public Function MatAdj(ByVal rng As Range) As Variant
    On Error GoTo Fail
    MatAdj = modInternalMatrices.Adjugate(modInternalMatrices.RangeToMatrix(rng))
    Exit Function
Fail:
    MatAdj = CVErr(xlErrNum)
End Function

Public Function MatPInv(ByVal rng As Range) As Variant
    On Error GoTo Fail
    MatPInv = modInternalMatrices.PseudoInverse(modInternalMatrices.RangeToMatrix(rng))
    Exit Function
Fail:
    MatPInv = CVErr(xlErrNum)
End Function

Public Function MatCofactor(ByVal rng As Range) As Variant
    On Error GoTo Fail
    MatCofactor = modInternalMatrices.CofactorMatrix(modInternalMatrices.RangeToMatrix(rng))
    Exit Function
Fail:
    MatCofactor = CVErr(xlErrNum)
End Function

Public Function MatMinor(ByVal rng As Range, ByVal r As Long, ByVal c As Long) As Variant
    On Error GoTo Fail
    MatMinor = modInternalMatrices.MinorMatrix(modInternalMatrices.RangeToMatrix(rng), r, c)
    Exit Function
Fail:
    MatMinor = CVErr(xlErrValue)
End Function

Public Function MatIsSymmetric(ByVal rng As Range) As Variant
    On Error GoTo Fail
    MatIsSymmetric = modInternalMatrices.IsSymmetric(modInternalMatrices.RangeToMatrix(rng))
    Exit Function
Fail:
    MatIsSymmetric = CVErr(xlErrValue)
End Function

Public Function MatIsOrthogonal(ByVal rng As Range) As Variant
    On Error GoTo Fail
    MatIsOrthogonal = modInternalMatrices.IsOrthogonal(modInternalMatrices.RangeToMatrix(rng))
    Exit Function
Fail:
    MatIsOrthogonal = CVErr(xlErrValue)
End Function
