Attribute VB_Name = "rCallback"
Option Explicit

' Name: rCallback
' Version: 0.5
' Depends: rArray
' Author: Rafael Fillipe Silva
' Description: ...

Private Function CallByNameArgsHelper(ByRef Obj As Variant, ByRef Proc As Variant, Optional ByVal CallType As VbCallType = VbMethod, _
                                      Optional Args As Variant, Optional RetObj As Boolean = False) As Variant
    Dim Size As Long
    Dim i As Long

    If IsArray1d(Args) Then
        i = LBound(Args)
        Size = UBound(Args) - i + 1

        If Size = 1 Then
            If RetObj Then
                Set CallByNameArgsHelper = CallByName(Obj, Proc, CallType, CVar(Args(i)))
            Else
                CallByNameArgsHelper = CallByName(Obj, Proc, CallType, CVar(Args(i)))
            End If
        ElseIf Size = 2 Then
            If RetObj Then
                Set CallByNameArgsHelper = CallByName(Obj, Proc, CallType, CVar(Args(i)), CVar(Args(i + 1)))
            Else
                CallByNameArgsHelper = CallByName(Obj, Proc, CallType, CVar(Args(i)), CVar(Args(i + 1)))
            End If
        ElseIf Size = 3 Then
            If RetObj Then
                Set CallByNameArgsHelper = CallByName(Obj, Proc, CallType, CVar(Args(i)), CVar(Args(i + 1)), CVar(Args(i + 2)))
            Else
                CallByNameArgsHelper = CallByName(Obj, Proc, CallType, CVar(Args(i)), CVar(Args(i + 1)), CVar(Args(i + 2)))
            End If
        ElseIf Size = 4 Then
            If RetObj Then
                Set CallByNameArgsHelper = CallByName(Obj, Proc, CallType, CVar(Args(i)), CVar(Args(i + 1)), CVar(Args(i + 2)), CVar(Args(i + 3)))
            Else
                CallByNameArgsHelper = CallByName(Obj, Proc, CallType, CVar(Args(i)), CVar(Args(i + 1)), CVar(Args(i + 2)), CVar(Args(i + 3)))
            End If
        ElseIf Size = 5 Then
            If RetObj Then
                Set CallByNameArgsHelper = CallByName(Obj, Proc, CallType, CVar(Args(i)), CVar(Args(i + 1)), CVar(Args(i + 2)), CVar(Args(i + 3)), CVar(Args(i + 4)))
            Else
                CallByNameArgsHelper = CallByName(Obj, Proc, CallType, CVar(Args(i)), CVar(Args(i + 1)), CVar(Args(i + 2)), CVar(Args(i + 3)), CVar(Args(i + 4)))
            End If
        End If
    Else
        If RetObj Then
            Set CallByNameArgsHelper = CallByName(Obj, Proc, CallType, CVar(Args))
        Else
            CallByNameArgsHelper = CallByName(Obj, Proc, CallType, CVar(Args))
        End If
    End If
End Function

Private Function RunArgsHelper(Optional ByRef Proc As Variant, Optional ByRef Args As Variant, Optional ByRef RetObj As Boolean = False)
    Dim Size As Long
    Dim i As Long

    If IsMissing(Proc) Then
        Proc = ""
    End If

    If IsArray1d(Args) Then
        i = LBound(Args)

        If Proc = "" Then
            Proc = Args(i)
            i = i + 1
        End If

        Size = UBound(Args) - i + 1

        If Size = 1 Then
            If RetObj Then
                Set RunArgsHelper = Application.Run(Proc, CVar(Args(i)))
            Else
                RunArgsHelper = Application.Run(Proc, CVar(Args(i)))
            End If
        ElseIf Size = 2 Then
            If RetObj Then
                Set RunArgsHelper = Application.Run(Proc, CVar(Args(i)), CVar(Args(i + 1)))
            Else
                RunArgsHelper = Application.Run(Proc, CVar(Args(i)), CVar(Args(i + 1)))
            End If
        ElseIf Size = 3 Then
            If RetObj Then
                Set RunArgsHelper = Application.Run(Proc, CVar(Args(i)), CVar(Args(i + 1)), CVar(Args(i + 2)))
            Else
                RunArgsHelper = Application.Run(Proc, CVar(Args(i)), CVar(Args(i + 1)), CVar(Args(i + 2)))
            End If
        ElseIf Size = 4 Then
            If RetObj Then
                Set RunArgsHelper = Application.Run(Proc, CVar(Args(i)), CVar(Args(i + 1)), CVar(Args(i + 2)), CVar(Args(i + 3)))
            Else
                RunArgsHelper = Application.Run(Proc, CVar(Args(i)), CVar(Args(i + 1)), CVar(Args(i + 2)), CVar(Args(i + 3)))
            End If
        ElseIf Size = 5 Then
            If RetObj Then
                Set RunArgsHelper = Application.Run(Proc, CVar(Args(i)), CVar(Args(i + 1)), CVar(Args(i + 2)), CVar(Args(i + 3)), CVar(Args(i + 4)))
            Else
                RunArgsHelper = Application.Run(Proc, CVar(Args(i)), CVar(Args(i + 1)), CVar(Args(i + 2)), CVar(Args(i + 3)), CVar(Args(i + 4)))
            End If
        End If
    Else
        If Proc <> "" Then
            If RetObj Then
                Set RunArgsHelper = Application.Run(Proc, CVar(Args))
            Else
                RunArgsHelper = Application.Run(Proc, CVar(Args))
            End If
        Else
            If RetObj Then
                Set RunArgsHelper = Application.Run(Args)
            Else
                RunArgsHelper = Application.Run(Args)
            End If
        End If
    End If
End Function

Public Function RunCallback(Optional ByRef Obj As Variant, Optional ByRef Proc As Variant, _
                            Optional ByRef Args As Variant, Optional ByVal RetObj As Boolean = False, _
                            Optional ByVal CallType As Variant = VbMethod) As Variant
    If IsMissing(Obj) Then
        Set Obj = Nothing
    End If

    If IsMissing(Proc) Then
        Proc = ""
    End If

    If Not Obj Is Nothing Then
        If Proc <> "" Then
            ' CallByName
            If RetObj Then
                Set RunCallback = CallByNameArgsHelper(Obj, Proc, CallType, Args, RetObj)
            Else
                RunCallback = CallByNameArgsHelper(Obj, Proc, CallType, Args, RetObj)
            End If
        Else
            ' Default
            If RetObj Then
                Set RunCallback = Obj!Args
            Else
                RunCallback = Obj!Args
            End If
        End If
    Else
        ' Application.Run
        If Proc <> "" Then
            If RetObj Then
                Set RunCallback = RunArgsHelper(Proc, Args, RetObj)
            Else
                RunCallback = RunArgsHelper(Proc, Args, RetObj)
            End If
        Else
            If RetObj Then
                Set RunCallback = RunArgsHelper(, Args, RetObj)
            Else
                RunCallback = RunArgsHelper(, Args, RetObj)
            End If
        End If
    End If
End Function
