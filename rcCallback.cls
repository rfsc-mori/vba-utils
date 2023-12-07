VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "rcCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Name: rcCallback
' Version: 0.3
' Depends: rArrays
' Author: Rafael Fillipe Silva
' Description: ...

Public Obj As Object
Public Proc As String
Public CallType As VbCallType
Public Args As Variant
Public RetObj As Variant

Private Function CallByNameArgsHelper(Obj As Object, Proc As Variant, CallType As VbCallType, _
                                      Optional InArgs As Variant, Optional RetObj As Variant)
    Dim Size As Long
    Dim i As Long

    If IsArray1d(InArgs) Then
        i = LBound(InArgs)
        Size = UBound(InArgs) - i + 1

        If Size = 1 Then
            If RetObj Then
                Set CallByNameArgsHelper = CallByName(Obj, Proc, CallType, InArgs(i))
            Else
                CallByNameArgsHelper = CallByName(Obj, Proc, CallType, InArgs(i))
            End If
        ElseIf Size = 2 Then
            If RetObj Then
                Set CallByNameArgsHelper = CallByName(Obj, Proc, CallType, InArgs(i), InArgs(i + 1))
            Else
                CallByNameArgsHelper = CallByName(Obj, Proc, CallType, InArgs(i), InArgs(i + 1))
            End If
        ElseIf Size = 3 Then
            If RetObj Then
                Set CallByNameArgsHelper = CallByName(Obj, Proc, CallType, InArgs(i), InArgs(i + 1), InArgs(i + 2))
            Else
                CallByNameArgsHelper = CallByName(Obj, Proc, CallType, InArgs(i), InArgs(i + 1), InArgs(i + 2))
            End If
        ElseIf Size = 4 Then
            If RetObj Then
                Set CallByNameArgsHelper = CallByName(Obj, Proc, CallType, InArgs(i), InArgs(i + 1), InArgs(i + 2), InArgs(i + 3))
            Else
                CallByNameArgsHelper = CallByName(Obj, Proc, CallType, InArgs(i), InArgs(i + 1), InArgs(i + 2), InArgs(i + 3))
            End If
        ElseIf Size = 5 Then
            If RetObj Then
                Set CallByNameArgsHelper = CallByName(Obj, Proc, CallType, InArgs(i), InArgs(i + 1), InArgs(i + 2), InArgs(i + 3), InArgs(i + 4))
            Else
                CallByNameArgsHelper = CallByName(Obj, Proc, CallType, InArgs(i), InArgs(i + 1), InArgs(i + 2), InArgs(i + 3), InArgs(i + 4))
            End If
        End If
    Else
        If RetObj Then
            Set CallByNameArgsHelper = CallByName(Obj, Proc, CallType, InArgs)
        Else
            CallByNameArgsHelper = CallByName(Obj, Proc, CallType, InArgs)
        End If
    End If
End Function

Private Function RunArgsHelper(Optional InProc As Variant, Optional InArgs As Variant, Optional RetObj As Variant)
    Dim Size As Long
    Dim i As Long

    If IsArray1d(InArgs) Then
        i = LBound(InArgs)

        If IsMissing(InProc) Then
            InProc = InArgs(First)
            i = i + 1
        End If

        Size = UBound(InArgs) - i + 1

        If Size = 1 Then
            If RetObj Then
                Set RunArgsHelper = Application.Run(InProc, InArgs(i))
            Else
                RunArgsHelper = Application.Run(InProc, InArgs(i))
            End If
        ElseIf Size = 2 Then
            If RetObj Then
                Set RunArgsHelper = Application.Run(InProc, InArgs(i), InArgs(i + 1))
            Else
                RunArgsHelper = Application.Run(InProc, InArgs(i), InArgs(i + 1))
            End If
        ElseIf Size = 3 Then
            If RetObj Then
                Set RunArgsHelper = Application.Run(InProc, InArgs(i), InArgs(i + 1), InArgs(i + 2))
            Else
                RunArgsHelper = Application.Run(InProc, InArgs(i), InArgs(i + 1), InArgs(i + 2))
            End If
        ElseIf Size = 4 Then
            If RetObj Then
                Set RunArgsHelper = Application.Run(InProc, InArgs(i), InArgs(i + 1), InArgs(i + 2), InArgs(i + 3))
            Else
                RunArgsHelper = Application.Run(InProc, InArgs(i), InArgs(i + 1), InArgs(i + 2), InArgs(i + 3))
            End If
        ElseIf Size = 5 Then
            If RetObj Then
                Set RunArgsHelper = Application.Run(InProc, InArgs(i), InArgs(i + 1), InArgs(i + 2), InArgs(i + 3), InArgs(i + 4))
            Else
                RunArgsHelper = Application.Run(InProc, InArgs(i), InArgs(i + 1), InArgs(i + 2), InArgs(i + 3), InArgs(i + 4))
            End If
        End If
    Else
        If InProc <> "" Then
            If RetObj Then
                Set RunArgsHelper = Application.Run(InProc, InArgs)
            Else
                RunArgsHelper = Application.Run(InProc, InArgs)
            End If
        Else
            If RetObj Then
                Set RunArgsHelper = Application.Run(InArgs)
            Else
                RunArgsHelper = Application.Run(InArgs)
            End If
        End If
    End If
End Function

Public Function Run(Optional InObj As Variant, Optional InProc As Variant, _
                    Optional InArgs As Variant, Optional InRetObj As Variant, _
                    Optional InCallType As Variant) As Variant
    If Not IsMissing(InObj) Then
        If Not InObj Is Nothing Then
            Set Obj = InObj
        End If
    End If

    If Not IsMissing(InProc) Then
        Proc = InProc
    End If

    If IsMissing(InRetObj) Then
        If RetObj = Empty Then
            RetObj = False
        End If
    Else
         RetObj = InRetObj
    End If

    If IsMissing(InCallType) Then
        If CallType <> VbMethod And CallType <> VbGet And CallType <> VbLet And CallType <> VbSet Then
            CallType = VbMethod
        End If
    Else
        CallType = InCallType
    End If

    If Not IsMissing(InArgs) Then
        Args = InArgs
    End If

    If Not Obj Is Nothing Then
        If Proc <> "" Then
            ' CallByName
            If RetObj Then
                Set Run = CallByName(Obj, Proc, CallType, Args)
            Else
                Run = CallByName(Obj, Proc, CallType, Args)
            End If
        Else
            ' Default
            If RetObj Then
                Set Run = Obj!Args
            Else
                Run = Obj!Args
            End If
        End If
    Else
        ' Application.Run
        If Proc <> "" Then
            If RetObj Then
                Set Run = RunArgsHelper(Proc, Args, RetObj)
            Else
                Run = RunArgsHelper(Proc, Args, RetObj)
            End If
        Else
            If RetObj Then
                Set Run = RunArgsHelper(, Args, RetObj)
            Else
                Run = RunArgsHelper(, Args, RetObj)
            End If
        End If
    End If
End Function

Public Sub Clear()
    Free
End Sub

Public Sub Free()
    If IsObject(Obj) Then
        Set Obj = Nothing
    Else
        Obj = Empty
    End If

    If IsObject(Args) Then
        Set Args = Nothing
    Else
        Args = Empty
    End If
End Sub

