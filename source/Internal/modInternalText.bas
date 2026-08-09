Attribute VB_Name = "modInternalText"
Option Explicit

' Internal: small string helpers shared across macros.

Public Function NzText(ByVal Value As Variant, Optional ByVal DefaultValue As String = "") As String
    If IsError(Value) Then
        NzText = DefaultValue
    ElseIf IsNull(Value) Or IsEmpty(Value) Then
        NzText = DefaultValue
    Else
        NzText = CStr(Value)
    End If
End Function

Public Function IsBlank(ByVal Value As Variant) As Boolean
    If IsError(Value) Or IsNull(Value) Or IsEmpty(Value) Then
        IsBlank = True
    ElseIf VarType(Value) = vbString Then
        IsBlank = (Len(Trim$(CStr(Value))) = 0)
    Else
        IsBlank = False
    End If
End Function
