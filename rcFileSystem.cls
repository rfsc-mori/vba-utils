VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "rcFileSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Name: rcFileSystem
' Version: 0.18
' Depends: rFiles,rcKeyValueCollection,rcProgressHelper,rcSetCollection
' Author: Rafael Fillipe Silva
' Description: ...

Public Root As Variant

Public Directories As rcKeyValueCollection
Public Files As rcSetCollection

Public Errors As rcSetCollection

Public DisableFeedback As Boolean

Public PH As rcProgressHelper
Public Status As rcKeyValue

Public Sub Traverse(ByVal Path As String, Optional ByVal SubDirectories As Boolean = False)
    Dim KV As rcKeyValue

    If PH Is Nothing Then
        If Not DisableFeedback Then
            Set PH = New rcProgressHelper
            Call PH.Prepare(1, 1, "Descobrindo pastas: ", "...")
        End If
    ElseIf DisableFeedback Then
        Call PH.Finish(True)
        Call rcFree(PH)
    End If

    If Not Directories Is Nothing Then
        Call rcFree(Directories)
    End If

    If Not Files Is Nothing Then
        Call rcFree(Files)
    End If

    Root = Path

    If Errors Is Nothing Then
        Set Errors = New rcSetCollection
    End If

    Call GetDirectories
    Call GetFiles

    If Status Is Nothing Then
        Set Status = New rcKeyValue

        Status.Key = 1
        Status.Value = Directories.Count
    Else
        Status.Key = Status.Key + 1
        Status.Value = Status.Value + Directories.Count
    End If

    If Not PH Is Nothing Then
        Call PH.DoStep(1, Status.Value, Status.Key)
    End If

    If SubDirectories Then
        For Each KV In Directories.List
            Set KV.Value = New rcFileSystem

            Set KV.Value.Errors = Errors

            Set KV.Value.PH = PH
            Set KV.Value.Status = Status

            Call KV.Value.Traverse(Path & "\" & KV.Key, SubDirectories)
        Next KV
    End If

    If Not PH Is Nothing Then
        If Status.Key = Status.Value Then
            Call PH.Finish(True)
        End If
    End If
End Sub

Public Sub ReplicateTree(ByVal Target As String)
    Dim KV As rcKeyValue

    If Directories Is Nothing Then
        Exit Sub
    End If

    For Each KV In Directories.List
        If Not DirectoryExists(Target & "\" & KV.Key) Then
            Call MkDir(Target & "\" & KV.Key)
        End If

        If DirectoryExists(Target & "\" & KV.Key) Then
            If Not KV.Value Is Nothing Then
                Call KV.Value.ReplicateTree(Target & "\" & KV.Key)
            End If
        End If
    Next KV
End Sub

Public Function Count(Optional ByVal SubDirectories As Boolean = False) As Long
    Count = CountDirectories(SubDirectories) + CountFiles(SubDirectories)
End Function

Public Function CountDirectories(Optional ByVal SubDirectories As Boolean = False) As Long
    Dim KV As rcKeyValue

    If Not Directories Is Nothing Then
        CountDirectories = Directories.List.Count

        If SubDirectories Then
            For Each KV In Directories.List
                CountDirectories = CountDirectories + KV.Value.CountDirectories(SubDirectories)
            Next KV
        End If
    Else
        CountDirectories = 0
    End If
End Function

Public Function CountFiles(Optional ByVal SubDirectories As Boolean = False) As Long
    Dim KV As rcKeyValue

    If Not Files Is Nothing Then
        CountFiles = Files.List.Count

        If SubDirectories Then
            For Each KV In Directories.List
                If Not KV.Value Is Nothing Then
                    CountFiles = CountFiles + KV.Value.CountFiles(SubDirectories)
                End If
            Next KV
        End If
    Else
        CountFiles = 0
    End If
End Function

Public Function FindFile(ByVal Pattern As String, Optional ByVal Full As Boolean = True) As String
    Dim File As Variant

    For Each File In Files
        If File Like Pattern Then
            If Full Then
                FindFile = FullName(File)
            Else
                FindFile = File
            End If

            Exit Function
        End If
    Next File
End Function

Public Function FullName(Optional ByVal Path As String = Empty) As String
    If Path = "" Then
        FullName = Root
    Else
        FullName = Root & "\" & Path
    End If
End Function

Public Function FileName(Optional ByVal Path As String = Empty) As String
    If Path = "" Then
        FileName = GetFileFromPath(Root)
    Else
        FileName = GetFileFromPath(Path)
    End If
End Function

Public Function FilePath(Optional ByVal Path As String = Empty) As Variant
    If Path = "" Then
        FilePath = GetPathFromFile(Root)
    Else
        FilePath = GetPathFromFile(Path)
    End If
End Function

Public Function FullyQualifiedDirectories() As rcSetCollection
    Dim KV As rcKeyValue

    Set FullyQualifiedDirectories = Nothing

    If Directories Is Nothing Then
        Set FullyQualifiedDirectories = New rcSetCollection

        For Each File In Directories.List
            Call FullyQualifiedDirectories.Add(FullName(KV.Key))
        Next File
    End If
End Function

Public Function FullyQualifiedFiles() As rcSetCollection
    Dim File As Variant

    Set FullyQualifiedFiles = Nothing

    If Files Is Nothing Then
        Set FullyQualifiedFiles = New rcSetCollection

        For Each File In Files.List
            Call FullyQualifiedFiles.Add(FullName(File))
        Next File
    End If
End Function

Private Sub GetDirectories()
    Dim Directory As Variant

    On Error Resume Next

    Set Directories = New rcKeyValueCollection

    Directory = Dir(Root & "\*", vbDirectory)

    Do While Directory <> ""
        If Directory <> "." And Directory <> ".." And (GetAttr(Root & "\" & Directory) And vbDirectory) Then
            If Err.Number = 0 Then
                Call Directories.Add(Directory, Nothing)
            Else
                Call Errors.Add(Root & "\" & Directory)
                Call Err.Clear
            End If
        End If

        Directory = Dir(Attributes:=vbDirectory)
    Loop

    Call Directories.Sort
End Sub

Private Sub GetFiles()
    Dim File As Variant

    On Error Resume Next

    Set Files = New rcSetCollection

    File = Dir(Root & "\*")

    If Err.Number <> 0 Then
        Call Errors.Add(Root & "\*")
        Err.Clear
    End If

    Do While File <> ""
        Call Files.Add(File)
        File = Dir
    Loop

    Call Files.Sort
End Sub

Public Sub Free()
    Call rcFree(Directories)
    Call rcFree(Files)
End Sub
