VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "rcKeyValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Name: rcKeyValue
' Version: 0.8
' Author: Rafael Fillipe Silva
' Description: ...

Public Key As Variant
Public Value As Variant

Public Sub SetKV(Optional K As Variant, Optional V As Variant)
    If IsObject(K) Then
        Set Key = K
    Else
        Key = K
    End If

    If IsObject(V) Then
        Set Value = V
    Else
        Value = V
    End If
End Sub

Public Function GetKey() As Variant
    If IsObject(Key) Then
        Set GetKey = Key
    Else
        GetKey = Key
    End If
End Function

Public Function GetValue() As Variant
Attribute GetValue.VB_UserMemId = 0
    If IsObject(Value) Then
        Set GetValue = Value
    Else
        GetValue = Value
    End If
End Function

Public Sub Clear()
    Free
End Sub

Public Sub Free()
    If IsObject(Key) Then
        Set Key = Nothing
    Else
        Key = Empty
    End If

    If IsObject(Value) Then
        Set Value = Nothing
    Else
        Value = Empty
    End If
End Sub
