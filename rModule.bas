Attribute VB_Name = "rModule"
Option Explicit

' Name: rModule
' Version: 0.58
' Depends: rCommon
' Author: Rafael Fillipe Silva
' Description: ...

' NOTA: Para utilizar as funções de módulo é necessário
' ativar a opção de confiar no modelo de objeto do VBA.
' Estas são apenas utilizadas no desenvolvimento do relatório
' através da função executar Sub (F5) ou automaticamente quando ativadas.

' Caminho para onde os módulos são salvos. PS: Necessário uma contrabarra "\" no final.
Private Const Path As String = "D:\Profiles\s-rafael.silva\Documents\Modules\"
Private Const HistoryPath As String = "D:\Profiles\s-rafael.silva\Documents\ModuleHistory\"

Private Const vbext_ct_ClassModule = 2 ' Classe
Private Const vbext_ct_StdModule = 1 ' módulo

Private Enum reModuleReplaceStatus
    reModuleSameVersion = 0
    reModuleDontReplace = 1
    reModuleCancelled = -1
    reModuleReplaced = 2
End Enum

' Importa o módulo de gerenciamento
Private Sub rModMngmt_Import()
    Call ImportModules("rModMngmt")
End Sub

Public Function ExportModules(Optional ByVal Overwrite As Boolean = False, _
                              Optional ByRef Modules As Variant, _
                              Optional ByRef Book As Workbook, _
                              Optional ByVal Silently As Boolean = False) As Long
    Dim VB As Object

    Dim File As String
    Dim History As String

    Dim Script As Object
    Dim NewScript As Object

    Dim Aux As Variant

    Dim LV, RV As String

    If Book Is Nothing Then Set Book = ThisWorkbook

    On Error Resume Next
    Set VB = Book.VBProject
    On Error GoTo 0

    If VB Is Nothing Then
        Exit Function
    End If

    For Each Script In VB.VBComponents
        If Not IsMissing(Modules) Then
            File = ""

            For Each Aux In Modules
                If StrComp(Script.Name, Aux, vbTextCompare) = 0 Then
                    File = Script.Name
                    Exit For
                End If
            Next Aux

            If File = "" Then
                GoTo Next_Module
            End If
        End If

        File = GetScriptPath(Script)

        If Len(File) = 0 Then
            GoTo Next_Module
        End If

        If Dir(File) <> "" Then
            If Overwrite Then
                Set NewScript = VB.VBComponents.Import(File)
                History = GetHistoryPath(Script.Name, GetModuleVersion(NewScript, True))

                If History = "" Then
                    MsgBox "Impossivel fazer backup do script: " & vbNewLine & Quote(File)
                    GoTo CleanUp
                End If

                LV = GetModuleVersion(Script, True)
                RV = GetModuleVersion(NewScript, True)

                If StrComp(LV, RV, vbBinaryCompare) <> 0 Then
                    If Not Silently Then
                        Aux = MsgBox(Script.Name & ": Versão diferente do disco, deseja mesmo salvar?" & vbNewLine & vbNewLine & _
                                     "Antigo:" & vbTab & RV & vbNewLine & "Novo:" & vbTab & LV, vbYesNoCancel)
                    Else
                        Aux = vbYes
                    End If

                    VB.VBComponents.Remove NewScript
                    Set NewScript = Nothing

                    If Aux = vbNo Then
                        GoTo Next_Module
                    ElseIf Aux = vbCancel Then
                        GoTo CleanUp
                    End If
                Else
                    VB.VBComponents.Remove NewScript
                    Set NewScript = Nothing

                    GoTo Next_Module
                End If

                SetAttr File, vbNormal
                FileCopy File, History
                Kill File
            Else
                GoTo Next_Module
            End If
        End If

        Script.Export File
        ExportModules = ExportModules + 1

Next_Module:
    Next Script

CleanUp:
End Function

Public Function ImportModules(ByRef Modules As Variant, _
                              Optional ByRef Book As Workbook, _
                              Optional ByVal Silently As Boolean = False, _
                              Optional ByVal ImportDependencies As Boolean = True) As Long
    Dim Module As Variant
    Dim Aux As Variant

    Dim VB As Object

    Dim File As String

    Dim Script As Object
    Dim NewScript As Object

    Dim Status As reModuleReplaceStatus

    Dim Dependencies As Collection
    Dim Imported As Collection

    If Book Is Nothing Then Set Book = ThisWorkbook

    On Error Resume Next
    Set VB = Book.VBProject
    On Error GoTo 0

    If VB Is Nothing Then
        Exit Function
    End If

    Set Dependencies = New Collection

    If IsObject(Modules) Then
        For Each Module In Modules
            Dependencies.Add Module
        Next Module
    ElseIf IsArray(Modules) Then
        For Module = LBound(Modules) To UBound(Modules)
            Dependencies.Add Modules(Module)
        Next Module
    Else
        Dependencies.Add Modules
    End If

    If ImportDependencies Then
        Set Imported = New Collection
    End If

Load_Dependencies:
    For Each Module In Dependencies
        If ImportDependencies Then
            For Each Aux In Imported
                If StrComp(Module, Aux, vbTextCompare) = 0 Then
                    Module = Empty
                    Exit For
                End If
            Next Aux
        End If

        If Module = "" Then
            If ImportDependencies Then Imported.Add Module
            GoTo Next_Module
        End If

        File = GetScriptPath(Module)

        If Len(File) = 0 Then
            If ImportDependencies Then Imported.Add Module
            GoTo Next_Module
        End If

        If Dir(File) <> "" Then
            Set NewScript = VB.VBComponents.Import(File)

            If StrComp(NewScript.Name, Module, vbTextCompare) <> 0 Then
                Status = ReplaceScript(VB.VBComponents(Module), NewScript, Book, Silently)
            Else
                Status = reModuleReplaced
            End If

            If ImportDependencies Then
                If Status = reModuleReplaced Then
                    GetModuleDependencies NewScript, Dependencies
                End If

                Imported.Add Module
            End If

            If Status = reModuleReplaced Then
                ImportModules = ImportModules + 1
            Else
                VB.VBComponents.Remove NewScript
            End If

            Set NewScript = Nothing
        Else
            If ImportDependencies Then Imported.Add Module
        End If

Next_Module:
    Next Module

    If ImportDependencies Then
        For Each Module In Imported
            For Aux = 1 To Dependencies.Count
                If StrComp(Dependencies(Aux), Module) = 0 Then
                    Dependencies.Remove Aux
                    Exit For
                End If
            Next Aux
        Next Module

        If Dependencies.Count > 0 Then
            GoTo Load_Dependencies
        End If
    End If
End Function

Public Function IsModuleLoaded(ByVal Module As String, Optional ByRef Book As Workbook) As Boolean
    Dim VB As Object

    Dim File As String

    If Book Is Nothing Then Set Book = ThisWorkbook

    On Error Resume Next
    Set VB = Book.VBProject
    On Error GoTo 0

    If VB Is Nothing Then
        Exit Function
    End If

    For Each Script In VB.VBComponents
        If StrComp(Script.Name, Module, vbTextCompare) = 0 Then
            IsModuleLoaded = True
            GoTo CleanUp
        End If
    Next Script

CleanUp:
End Function

Public Function SafeModule(ByVal Module As String, Optional ByRef Book As Workbook) As Object
    Dim VB As Object

    Dim Script As Object
    Dim File As String

    If Book Is Nothing Then Set Book = ThisWorkbook

    On Error Resume Next
    Set VB = Book.VBProject
    On Error GoTo 0

    If VB Is Nothing Then
        Exit Function
    End If

    For Each Script In VB.VBComponents
        If StrComp(Script.Name, Module, vbTextCompare) = 0 Then
            Set SafeModule = Script
            GoTo CleanUp
        End If
    Next Script

CleanUp:
End Function

Public Function RefreshModules(Optional ByRef Book As Workbook, _
                               Optional ByVal Silently As Boolean = False, _
                               Optional ByVal ImportDependencies As Boolean = True) As Long
    Dim VB As Object

    Dim File As Variant
    Dim Aux As Variant

    Dim Script As Object
    Dim NewScript As Object

    Dim Status As reModuleReplaceStatus

    Dim Dependencies As Collection
    Dim Imported As Collection

    If Book Is Nothing Then Set Book = ThisWorkbook

    On Error Resume Next
    Set VB = Book.VBProject
    On Error GoTo 0

    If VB Is Nothing Then
        Exit Function
    End If

    If ImportDependencies Then
        Set Imported = New Collection
    End If

    Set Dependencies = New Collection

    For Each Script In VB.VBComponents
        Dependencies.Add Script.Name
    Next Script

    For Each File In Dependencies
        Set Script = SafeModule(File, Book)

        If Not Script Is Nothing Then
            File = GetScriptPath(Script)

            If Len(File) = 0 Then
                If ImportDependencies Then Imported.Add Script.Name
                GoTo Next_Module
            End If

            If Dir(File) <> "" Then
                Set NewScript = VB.VBComponents.Import(File)
                Status = ReplaceScript(Script, NewScript, Book, Silently)

                If Status <> reModuleSameVersion Then
                    If Status = reModuleReplaced Then
                        RefreshModules = RefreshModules + 1

                        If ImportDependencies Then
                            Imported.Add NewScript.Name
                            GetModuleDependencies NewScript, Dependencies
                        End If
                    Else
                        VB.VBComponents.Remove NewScript

                        If Status = reModuleDontReplace Then
                            If ImportDependencies Then Imported.Add Script.Name
                            GoTo Next_Module
                        ElseIf Status = reModuleCancelled Then
                            GoTo CleanUp
                        End If
                    End If
                Else
                    VB.VBComponents.Remove NewScript
                    If ImportDependencies Then Imported.Add Script.Name
                End If

                Set NewScript = Nothing
            Else
                If ImportDependencies Then Imported.Add Script.Name
            End If
        End If
Next_Module:
    Next File

    If ImportDependencies Then
        For Each File In Imported
            For Aux = 1 To Dependencies.Count
                If StrComp(Dependencies(Aux), File) = 0 Then
                    Dependencies.Remove Aux
                    Exit For
                End If
            Next Aux
        Next File

        If Dependencies.Count > 0 Then
            ImportModules Dependencies, Book, Silently, ImportDependencies
        End If
    End If

CleanUp:
End Function

' Pega o caminho para um script apenas se ele faz parte da biblioteca "r".
Private Function GetScriptPath(ByRef Script As Variant) As String
    Dim ScriptPath As String

    If Right$(Path, 1) <> "\" And Right$(Path, 1) <> "/" Then
        ScriptPath = Path & "\"
    Else
        ScriptPath = Path
    End If

    If Dir(ScriptPath, vbDirectory) = "" Then
        Exit Function
    End If

    If IsObject(Script) And TypeName(Script) = "VBComponent" Then
        If StrComp(Left$(Script.Name, 2), "rc", vbBinaryCompare) = 0 Then
            If Script.Type = vbext_ct_ClassModule Then
                GetScriptPath = ScriptPath & Script.Name & ".cls"
            End If
        ElseIf StrComp(Left$(Script.Name, 1), "r", vbBinaryCompare) = 0 Then
            If Script.Type = vbext_ct_StdModule Then
                GetScriptPath = ScriptPath & Script.Name & ".bas"
            End If
        End If
    ElseIf TypeName(Script) = "String" Then
        If StrComp(Left$(Script, 2), "rc", vbBinaryCompare) = 0 Then
            GetScriptPath = ScriptPath & Script & ".cls"
        ElseIf StrComp(Left$(Script, 1), "r", vbBinaryCompare) = 0 Then
            GetScriptPath = ScriptPath & Script & ".bas"
        End If
    End If
End Function

' Pega o caminho para o histórico de um script apenas se ele faz parte da biblioteca "r".
Private Function GetHistoryPath(ByVal Script As String, ByVal Version As String) As String
    Dim History As String
    Dim Obj As Object
    Dim Aux As String

    GetHistoryPath = Empty

    If Right$(HistoryPath, 1) <> "\" And Right$(HistoryPath, 1) <> "/" Then
        History = HistoryPath & "\"
    Else
        History = HistoryPath
    End If

    If Dir(History, vbDirectory) = "" Then
        Exit Function
    End If

    Aux = Replace(Version, "/", "_")
    Aux = Replace(Version, "\", "_")

    If StrComp(Left$(Script, 2), "rc", vbBinaryCompare) = 0 Then
        GetHistoryPath = History & Script & "_" & Aux & ".cls"
    ElseIf StrComp(Left$(Script, 1), "r", vbBinaryCompare) = 0 Then
        GetHistoryPath = History & Script & "_" & Aux & ".cls"
    End If
End Function

' Remove o antigo script e adiciona o novo se as versões forem diferentes
Private Function ReplaceScript(ByRef Script As Object, ByRef NewScript As Object, _
                               Optional ByRef Book As Workbook, Optional ByVal Silently As Boolean = False) As reModuleReplaceStatus
    Dim VB As Object
    Dim Aux As String
    Dim LV, RV As String

    If Book Is Nothing Then Set Book = ThisWorkbook

    On Error Resume Next
    Set VB = Book.VBProject
    On Error GoTo 0

    If VB Is Nothing Then
        Exit Function
    End If

    LV = GetModuleVersion(Script, True)
    RV = GetModuleVersion(NewScript, True)

    If StrComp(LV, RV, vbBinaryCompare) <> 0 Then
        If Not Silently Then
            Aux = MsgBox(Script.Name & ": Versão diferente do disco, deseja mesmo recarregar?" & vbNewLine & vbNewLine & _
                         "Antigo:" & vbTab & LV & vbNewLine & "Novo:" & vbTab & RV, vbYesNoCancel)
        Else
            Aux = vbYes
        End If

        If Aux = vbNo Then
            ReplaceScript = reModuleDontReplace
            Exit Function
        ElseIf Aux = vbCancel Then
            ReplaceScript = reModuleCancelled
            Exit Function
        Else
            ReplaceScript = reModuleReplaced
        End If

        Aux = Script.Name
        Script.Name = Script.Name & "_tmp"
        VB.VBComponents.Remove Script
        NewScript.Name = Aux
    Else
        ReplaceScript = reModuleSameVersion
    End If
End Function

Private Function SameVersion(ByRef Script As Object, ByRef NewScript As Object, Optional Checksum As Boolean = False) As Boolean
    SameVersion = (StrComp(GetModuleVersion(Script, Checksum), GetModuleVersion(NewScript, Checksum), vbBinaryCompare) = 0)
End Function

' Lê o campo comentado Version: nos módulos da biblioteca "r"
Private Function GetModuleVersion(ByRef Script As Object, Optional Checksum As Boolean = False) As String
    Dim VStr As Variant

    Dim SL As Long
    Dim EL As Long
    Dim SC As Long
    Dim EC As Long

    Dim CK As Long

    Dim Found As Boolean

    VStr = "' Version: "

    With Script.CodeModule
        SL = 1 ' Linha inicial
        EL = .CountOfLines ' Linha final
        SC = 1 ' Coluna inicial
        EC = 255 ' Coluna final

        ' Find atualiza as variáveis
        Found = .Find(Target:=VStr, StartLine:=SL, StartColumn:=SC, EndLine:=EL, EndColumn:=EC)

        If Found Then
            VStr = Mid$(.Lines(SL, EL - SL + 1), SC + 11)

            If Checksum Then
                CK = .CountOfLines + .CountOfDeclarationLines + Len(.Lines(1, .CountOfLines))
                VStr = VStr & "!" & CK
            End If

            GetModuleVersion = VStr
            Exit Function
        End If
    End With

    GetModuleVersion = Empty
End Function

' Lê o campo comentado Depends: nos módulos da biblioteca "r"
Private Sub GetModuleDependencies(ByRef Script As Object, ByRef Dependencies As Collection, Optional ByVal IgnoredLoaded As Boolean = False)
    Dim DStr As Variant
    Dim Aux As Variant
    Dim Helper As Variant
    Dim Mods
    Dim Module As Variant

    Dim SL As Long
    Dim EL As Long
    Dim SC As Long
    Dim EC As Long

    Dim Found As Boolean

    If Dependencies Is Nothing Then
        Exit Sub
    End If

    DStr = "' Depends: "

    With Script.CodeModule
        SL = 1 ' Linha inicial
        EL = .CountOfLines ' Linha final
        SC = 1 ' Coluna inicial
        EC = 255 ' Coluna final

        ' Find atualiza as variáveis
        Found = .Find(Target:=DStr, StartLine:=SL, StartColumn:=SC, EndLine:=EL, EndColumn:=EC)

        Do While Found
            Module = Mid$(.Lines(SL, EL - SL + 1), SC + 11)
            Mods = Split(Module, ",")

            For Aux = LBound(Mods) To UBound(Mods)
                Module = Mods(Aux)

                For Each Helper In Dependencies
                    If StrComp(Module, Helper, vbTextCompare) = 0 Then
                        Module = Empty
                        Exit For
                    End If
                Next Helper

                If Module <> "" And DStr <> "" Then
                    If IgnoredLoaded Then
                        If Not IsModuleLoaded(Module) Then
                            Dependencies.Add Trim$(Module)
                        End If
                    Else
                        Dependencies.Add Trim$(Module)
                    End If
                End If
            Next Aux

            SL = SL + 1 ' Linha inicial
            EL = .CountOfLines ' Linha final
            SC = 1 ' Coluna inicial
            EC = 255 ' Coluna final

            Found = .Find(Target:=DStr, StartLine:=SL, StartColumn:=SC, EndLine:=EL, EndColumn:=EC)
        Loop
    End With
End Sub

' Copia módulos que não sejam da biblioteca "r" de um arquivo para outro
' Ex: Atualizar o módulo controle em um arquivo sem alterar seus dados.
Public Function CopyModulesTo(ByRef Target As Workbook, _
                              Optional ByVal Insert As Boolean = False, _
                              Optional ByVal Overwrite As Boolean = False, _
                              Optional ByRef Source As Workbook, _
                              Optional ByRef Modules As Variant)
    Dim VBSource As Object
    Dim VBTarget As Object

    Dim Module As Variant

    Dim SourceScript As Object
    Dim TargetScript As Object

    Dim Aux As Variant

    If Source Is Nothing Then Set Source = ThisWorkbook

    On Error Resume Next
    Set VBSource = Source.VBProject
    Set VBTarget = Target.VBProject
    On Error GoTo 0

    If VBSource Is Nothing Or VBTarget Is Nothing Then
        Exit Function
    End If

    If IsMissing(Modules) Then
        Set Modules = Nothing
    End If

    For Each SourceScript In VBSource.VBComponents
        Module = GetScriptPath(SourceScript)

        If Len(Module) <> 0 Then
            GoTo Next_SourceScript
        End If

        If SourceScript.Type <> vbext_ct_ClassModule And SourceScript.Type <> vbext_ct_StdModule Then
            GoTo Next_SourceScript
        End If

        If Not Modules Is Nothing Then
            For Each Module In Modules
                If StrComp(SourceScript.Name, Module, vbTextCompare) <> 0 Then
                    GoTo Next_SourceScript
                End If
            Next Module
        End If

        Set TargetScript = Nothing

        For Each Aux In VBTarget.VBComponents
            If StrComp(SourceScript.Name, Aux.Name, vbTextCompare) = 0 Then
                Set TargetScript = Aux
                Exit For
            End If
        Next Aux

        If TargetScript Is Nothing Then
            If Insert Then
                If SourceScript.Type = vbext_ct_ClassModule Then
                    Set TargetScript = VBTarget.VBComponents.Add(vbext_ct_ClassModule)
                ElseIf SourceScript.Type = vbext_ct_StdModule Then
                    Set TargetScript = VBTarget.VBComponents.Add(vbext_ct_StdModule)
                End If
            End If
        Else
            If Not Overwrite Then
                Set TargetScript = Nothing
            End If
        End If

        If TargetScript Is Nothing Then
            GoTo Next_SourceScript
        End If

        TargetScript.Name = SourceScript.Name

        With TargetScript.CodeModule
            .DeleteLines 1, .CountOfLines
        End With

        With SourceScript.CodeModule
            TargetScript.CodeModule.AddFromString .Lines(1, .CountOfLines)
        End With

        Set TargetScript = Nothing

Next_SourceScript:
    Next SourceScript
End Function

