VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "rcProgressHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Name: rcProgressHelper
' Version: 0.24
' Author: Rafael Fillipe Silva
' Description: ...

' Auxiliares para manter o excel respondendo durante loops
' que potencialmente podem demorar muito
Private Loops As Long
Private Helper As Long
Private Step As Long
Private FirstStep As Date
Private Prefix As Variant
Private Suffix As Variant

Public Sub Prepare(Optional ByVal OptLoops As Long = 1000, Optional ByVal OptStep As Long = 1, _
                   Optional ByVal OptPrefix As String = Empty, Optional ByVal OptSuffix As String = Empty)
    Call Finish

    If OptLoops > 0 Then Loops = OptLoops
    If OptStep > 0 Then Step = OptStep

    FirstStep = DateTime.Now

    If OptPrefix <> "" Then Prefix = OptPrefix
    If OptSuffix <> "" Then Suffix = OptSuffix
End Sub

Public Function CheckDoStep() As Boolean
    CheckDoStep = False

    Helper = Helper + Step

    If Loops = 0 Then
        CheckDoStep = True
    ElseIf (Helper Mod Loops) = 0 Then
        CheckDoStep = True
    End If
End Function

Public Function DoStep(Optional ByVal Feedback As Boolean = True, _
                       Optional ByVal Total As Long = 0, _
                       Optional ByVal Current As Variant) As Boolean
    Dim Minutes As Long
    Dim Seconds As Long

    If IsMissing(Current) Then Current = Helper

    DoStep = CheckDoStep

    If DoStep Then
        If Feedback Then
            If Total > 0 And Current > 0 Then
                Minutes = 0
                Seconds = DateDiff("s", FirstStep, DateTime.Now)
                Seconds = (Total - Current) * (Seconds / Current)

                Do While Seconds > 60
                    Minutes = Minutes + 1
                    Seconds = Seconds - 60
                Loop

                Seconds = Round(Seconds)

                Application.StatusBar = Prefix & _
                                        Current & "/" & Total & Suffix & _
                                        " (" & Format(Current / Total, "Percent") & " - " & _
                                        Minutes & " minutos e " & _
                                        Seconds & " segundos restantes.)"
            Else
                Minutes = 0
                Seconds = DateDiff("s", FirstStep, DateTime.Now)

                Do While Seconds > 60
                    Minutes = Minutes + 1
                    Seconds = Seconds - 60
                Loop

                Seconds = Round(Seconds)

                Application.StatusBar = Prefix & Current & Suffix & _
                                        " (" & Minutes & " minutos e " & _
                                        Seconds & " segundos decorridos.)"
            End If
        End If

        DoEvents
    End If
End Function

Public Sub Finish(Optional ByVal Feedback As Boolean = True)
    Dim Minutes As Variant
    Dim Seconds As Variant

    If Feedback And FirstStep <> 0 Then
        Minutes = 0
        Seconds = DateDiff("s", FirstStep, DateTime.Now)

        Do While Seconds > 60
            Minutes = Minutes + 1
            Seconds = Seconds - 60
        Loop

        Seconds = Round(Seconds)

        Application.StatusBar = Prefix & "completo em " & _
                                Minutes & " minutos e " & Seconds & " segundos."
    End If

    ' Limpa a barra de status
    ClearStatusBar

    Helper = 0
    Loops = 100
    Step = 1
    Prefix = "Processando: "
    Suffix = "."
    FirstStep = Empty
End Sub

Public Sub Free()
End Sub
