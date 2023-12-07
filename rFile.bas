Attribute VB_Name = "rFile"
Option Explicit

' Name: rFile
' Version: 0.53
' Depends: rArray, rCommon
' Author: Rafael Fillipe Silva
' Description: ...

Public Function AskForFile(Optional ByVal Title As String = Empty, Optional ByVal Path As String = Empty) As Variant
    If Title = "" Then Title = ThisWorkbook.Name
    If Path = "" Then Path = ThisWorkbook.Path

    ' Abre dialogo de escolha para abrir um arquivo
    Set AskForFile = AskForFiles(Title, True, Path)

    If Not AskForFile Is Nothing Then
        If AskForFile.Count > 0 Then
            AskForFile = AskForFile.Item(1)
        Else
            AskForFile = False
        End If
    Else
        AskForFile = False
    End If
End Function

Public Function AskForDirectory(Optional ByVal Title As String = Empty, Optional ByVal Path As String = Empty) As Variant
    Dim Choice As Long

    If Title = "" Then Title = ThisWorkbook.Name
    If Path = "" Then Path = ThisWorkbook.Path

    Set AskForDirectory = Nothing

    Application.FileDialog(msoFileDialogFolderPicker).InitialFileName = Path

    ' Abre dialogo de escolha para selecionar uma pasta
    Application.FileDialog(msoFileDialogFolderPicker).AllowMultiSelect = False
    Application.FileDialog(msoFileDialogFolderPicker).Title = Title

    Choice = Application.FileDialog(msoFileDialogFolderPicker).Show

    ' Clicou em cancelar
    If Choice = 0 Then
        AskForDirectory = False
        Exit Function
    End If

    Set AskForDirectory = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems

    If Not AskForDirectory Is Nothing Then
        If AskForDirectory.Count > 0 Then
            AskForDirectory = AskForDirectory.Item(1)
        Else
            AskForDirectory = False
        End If
    Else
        AskForDirectory = False
    End If
End Function

Public Function AskForFiles(Optional ByVal Title As String = Empty, _
                            Optional ByVal SingleFile As Boolean = False, _
                            Optional ByVal Path As String = Empty) As FileDialogSelectedItems
    Dim Choice As Long

    If Title = "" Then Title = ThisWorkbook.Name
    If Path = "" Then Path = ThisWorkbook.Path

    Set AskForFiles = Nothing

    Application.FileDialog(msoFileDialogOpen).InitialFileName = Path

    ' Abre dialogo de escolha para abrir varios arquivos
    Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = (Not SingleFile)
    Application.FileDialog(msoFileDialogOpen).Title = Title

    Choice = Application.FileDialog(msoFileDialogOpen).Show

    ' Clicou em cancelar
    If Choice = 0 Then
        Exit Function
    End If

    Set AskForFiles = Application.FileDialog(msoFileDialogOpen).SelectedItems
End Function

Public Function AskSaveFile(Optional ByVal Title As String = Empty, _
                            Optional ByVal Path As String = Empty, _
                            Optional ByVal FilterId As Variant) As Variant
    Dim Choice As Long
    Dim Items As FileDialogSelectedItems

    If Title = "" Then Title = ThisWorkbook.Name
    If Path = "" Then Path = ThisWorkbook.Path

    AskSaveFile = False

    Application.FileDialog(msoFileDialogSaveAs).Title = Title
    Application.FileDialog(msoFileDialogSaveAs).InitialFileName = Path

    If Not IsMissing(FilterId) Then
        Application.FileDialog(msoFileDialogSaveAs).FilterIndex = FilterId
    End If

    Application.FileDialog(msoFileDialogSaveAs).AllowMultiSelect = False

    Choice = Application.FileDialog(msoFileDialogSaveAs).Show

    ' Clicou em cancelar
    If Choice = 0 Then
        Exit Function
    End If

    Set Items = Application.FileDialog(msoFileDialogSaveAs).SelectedItems

    If Items Is Nothing Then
        Exit Function
    End If

    If Items.Count > 0 Then
        AskSaveFile = Items(1)
    End If
End Function

' Retorna Nothing ao invés de erro
Public Function SafeOpen(ByVal File As String, Optional ByVal RW As Boolean = False, Optional ByVal ForceLocal As Boolean = False) As Workbook
    Set SafeOpen = Nothing

    If StrComp(File, ThisWorkbook.FullName, vbTextCompare) = 0 Then
        Exit Function
    End If

    On Error GoTo ErrHandler
    Set SafeOpen = Workbooks.Open(FileName:=File, UpdateLinks:=0, ReadOnly:=Not RW, Local:=ForceLocal)

ErrHandler:
End Function

Public Function SafeWorkbook(ByVal File As String) As Workbook
    On Error Resume Next
    Set SafeWorkbook = Workbooks(GetFileFromPath(File))
End Function

Public Function SafeFile(ByVal File As String, Optional ByVal ForceLocal As Boolean = False) As Workbook
    Set SafeFile = SafeWorkbook(GetFileFromPath(File))

    If SafeFile Is Nothing Then
        Set SafeFile = SafeOpen(File, True, ForceLocal)
    End If
End Function

Public Function SafeSave(ByVal File As String, Optional ByVal Format As Variant, Optional ByRef Book As Workbook) As Boolean
    SafeSave = False

    If Book Is Nothing Then Set Book = ThisWorkbook

    On Error GoTo ErrHandler

    If IsMissing(Format) Then
        Book.SaveAs File
    Else
        Book.SaveAs File, Format
    End If

    SafeSave = True

ErrHandler:
End Function

' Retorna Nothing ao invés de erro
' não mostra alertas e perguntas sobre arquivos abertos etc
Public Function SilentlyOpen(ByVal File As String, Optional ByVal RW As Boolean = False, Optional ByVal ForceLocal As Boolean = False) As Workbook
    Set SilentlyOpen = Nothing

    If StrComp(File, ThisWorkbook.FullName, vbTextCompare) = 0 Then
        Exit Function
    End If

    On Error GoTo CleanUp
    Application.DisplayAlerts = False

    Set SilentlyOpen = Workbooks.Open(FileName:=File, UpdateLinks:=0, ReadOnly:=Not RW, Local:=ForceLocal)

CleanUp:
    On Error Resume Next
    Application.DisplayAlerts = True
End Function

' Lista de arquivos abertos que buscam informações deste arquivo através de links
Public Function LinkedToMe() As Collection
    Dim Source As String
    Dim Conflicts As Collection

    Dim Book As Workbook
    Dim Refs As Variant

    Dim i As Variant

    Set Conflicts = New Collection
    Set LinkedToMe = Nothing

    Source = ThisWorkbook.Name

    For Each Book In Application.Workbooks
        Refs = Book.LinkSources(xlExcelLinks)

        If Not IsEmpty(Refs) Then
            For i = 1 To UBound(Refs)
                If InStr(1, Refs(i), Source, vbTextCompare) <> 0 Then
                    Conflicts.Add Book.Name
                End If
            Next i
        End If
    Next Book

    Set LinkedToMe = Conflicts
End Function

' Quebra links a arquivos externos
Public Sub BreakLinks(Optional ByRef Book As Workbook = Nothing)
    Dim Refs As Variant
    Dim i As Variant

    If Book Is Nothing Then
        Set Book = ThisWorkbook
    End If

    Refs = Book.LinkSources(xlExcelLinks)

    If Not IsEmpty(Refs) Then
        For i = 1 To UBound(Refs)
            Call Book.BreakLink(Refs(i), xlLinkTypeExcelLinks)
        Next i
    End If
End Sub

' Checa se arquivo existe
Public Function FileExists(ByVal File As String) As Boolean
    FileExists = (Dir(File) <> "")
End Function

' Checa se pasta existe
Public Function DirectoryExists(ByVal File As String) As Boolean
    DirectoryExists = (Dir(File, vbDirectory) <> "")
End Function

Public Function TryFiles(ByRef Files As Variant, Optional Ext As Variant) As String
    Dim FList As Variant
    Dim EList As Variant
    Dim File As Variant
    Dim x, y As Long

    FList = SMakeArray1d(Files, 0)
    EList = SMakeArray1d(Ext, 0)

    For x = LBound(FList) To UBound(FList)
        For y = LBound(EList) To UBound(EList)
            File = FList(x) & EList(y)

            If FileExists(File) Then
                TryFiles = File
                Exit Function
            End If
        Next y
    Next x
End Function

' Apaga um arquivo
Public Function DeleteFile(ByVal File As String) As Boolean
    DeleteFile = False

    If FileExists(File) Then
        SetAttr File, vbNormal
        Kill File

        DeleteFile = True
    End If
End Function

' Separa o caminho da pasta de um nome de arquivo
Public Function GetPathFromPath(ByVal Path As String) As String
    Dim Pos As Variant

    Pos = FindLast(Path, "\")

    If Pos <> 0 Then
        GetPathFromPath = Left$(Path, Pos - 1)
    Else
        GetPathFromPath = Empty
    End If
End Function

' Separa o nome de um arquivo através de um caminho
Public Function GetFileFromPath(ByVal Path As String) As String
    Dim Pos As Variant

    Pos = FindLast(Path, "\")

    If Pos <> 0 Then
        GetFileFromPath = Mid$(Path, Pos + 1)
    Else
        GetFileFromPath = Path
    End If
End Function

' Separa o nome de um arquivo sem extensão através de um caminho
Public Function GetFileWithoutExtFromPath(ByVal Path As String) As String
    Dim Pos As Variant

    GetFileWithoutExtFromPath = GetFileFromPath(Path)

    Pos = FindLast(GetFileWithoutExtFromPath, ".")

    If Pos <> 0 Then
        GetFileWithoutExtFromPath = Left$(GetFileWithoutExtFromPath, Pos - 1)
    End If
End Function

' Separa a extensão de um arquivo através de um caminho
Public Function GetFileExt(ByVal Path As String) As String
    Dim Pos As Variant

    Pos = FindLast(Path, ".")

    If Pos <> 0 Then
        GetFileExt = Mid$(Path, Pos + 1)
    Else
        GetFileExt = Empty
    End If
End Function

' Altera o nome de um arquivo independente do caminho e da extensão
Public Function ChangeFileName(ByVal FullPath As String, ByVal FileName As String)
    Dim Path As Variant
    Dim Ext As Variant

    Path = GetPathFromPath(FullPath)

    If Path <> "" Then
        Path = Path & "\"
    End If

    Ext = GetExt(FullPath)

    If Ext <> "" Then
        Ext = "." & Ext
    End If

    ChangeFileName = Path & FileName & Ext
End Function

' Obtém o endereço UNC de uma pasta de trabalho aberta
Public Function GetUNCPath(Optional ByRef Book As Workbook) As Variant
    Dim Controls As CommandBarControls
    Dim Control As CommandBarControl
    Dim Current As Workbook

    If Book Is Nothing Then Set Book = ThisWorkbook

    Set Current = ActiveWorkbook

    Book.Activate

    If Book.CommandBars Is Nothing Then
        Set Controls = Application.CommandBars("Web").Controls
    Else
        Set Controls = Book.CommandBars("Web").Controls
    End If

    For Each Control In Controls
        If Control.ID = 1740 Then
            GetUNCPath = Control.Text
            Exit For
        End If
    Next Control

    Current.Activate
End Function

Public Sub ChangeLinkSource(ByVal FileNamePattern As String, ByVal NewFile As String, Optional ByRef Book As Workbook = Nothing)
    Dim Refs As Variant
    Dim File As Variant
    Dim Aux As Variant
    Dim Source As Workbook
    Dim Target As Workbook
    Dim Sheet As Worksheet
    Dim R As Range
    Dim F As Range
    Dim x, y, z, i As Long

    If Book Is Nothing Then
        Set Book = ThisWorkbook
    End If

    Refs = Book.LinkSources(xlExcelLinks)

    If Not IsEmpty(Refs) Then
        Set Target = SafeWorkbook(NewFile)

        If Target Is Nothing Then
            Set Target = SafeOpen(NewFile)
        End If

        NewFile = GetPathFromPath(NewFile) & "\[" & GetFileFromPath(NewFile) & "]"

        If Not Target Is Nothing Then
            For x = 1 To UBound(Refs)
                File = GetFileFromPath(Refs(x))

                If File <> Target.Name Then
                    If File Like FileNamePattern Then
                        Set Source = SafeWorkbook(Refs(x))

                        If Source Is Nothing Then
                            Set Source = SafeOpen(Refs(x))
                        End If

                        File = "[" & File & "]"

                        On Error Resume Next

                        For Each Sheet In Book.Worksheets
                            Set R = GetUsedRange(Sheet)

                            If Not Source Is Nothing Then
                                Call R.Replace(File, NewFile, xlPart, , False, False, False, False)
                            Else
                                Set F = R.Find(File, , xlFormulas, xlPart, , , False, False, False)

                                If Not F Is Nothing Then
                                    Aux = F.Formula

                                    y = InStr(1, Aux, File, vbTextCompare)

                                    If y <> 0 Then
                                        i = InStr(y, Aux, "]", vbTextCompare)

                                        z = y

                                        Do While z > 1
                                            z = InStrRev(Aux, "'", z, vbTextCompare)

                                            If z <> 0 Then
                                                If z > 1 Then
                                                    If Mid$(Aux, z - 1, 1) <> "'" Then
                                                        Exit Do
                                                    End If
                                                End If
                                            End If
                                        Loop

                                        If z > 0 And z <> y Then
                                            Call R.Replace(Mid$(Aux, z, y - z + Len(File)), "'" & NewFile, xlPart, , False, False, False, False)
                                        End If
                                    End If
                                End If
                            End If
                        Next Sheet

                        On Error GoTo -1

                        If Not Source Is Nothing Then
                            Call Source.Close(False)
                            Set Source = Nothing
                        End If
                    End If
                End If
            Next x

            Call Target.Close(False)
            Set Target = Nothing
        End If
    End If
End Sub
