Attribute VB_Name = "rModMngmt"
Option Explicit

' Name: rModMngmt
' Version: 0.39
' Depends: rCommon,rModule
' Author: Rafael Fillipe Silva
' Description: ...

' Arquivo DEV_FLAG
Private Const DevFlag = "D:\Profiles\s-rafael.silva\Documents\Modules\DEV_FLAG"

' Checa se o gerenciamento de módulos será ativado
Private Function rModMngmt_IsDevUser() As Boolean
    On Error GoTo ErrHandler

    rModMngmt_IsDevUser = (StrComp(Environ("UserName"), "s-Rafael.Silva", vbTextCompare) = 0)  ' Usuário

    If rModMngmt_IsDevUser Then
        If (Dir(DevFlag) <> "") Then ' Arquivo DEV_FLAG
            rModMngmt_IsDevUser = (Not ThisWorkbook.ReadOnly) ' Desativa para arquivos somente leitura
            Exit Function
        End If
    End If

ErrHandler: ' Se ocorrer erros desativa o gerenciamento de módulos
    rModMngmt_IsDevUser = False
End Function

' Atualiza todos os módulos com versões diferentes ao abrir a planilha
Private Sub Auto_Open()
    Dim Saved As Boolean

    If rModMngmt_IsDevUser Then
        Saved = ThisWorkbook.Saved

        If rModMngmt_RefreshAll = 0 And Saved = True Then
            Application.OnTime DateTime.Now, "rModMngmt_ForceSaved"
        End If
    End If
End Sub

' Salva todos os módulos com versões diferentes ao fechar a planilha
Private Sub Auto_Close()
    Dim Saved As Boolean

    If rModMngmt_IsDevUser Then
        Saved = ThisWorkbook.Saved
        rModMngmt_ExportAll
        ThisWorkbook.Saved = Saved
    End If
End Sub

' Força o status salvo no arquivo
Private Sub rModMngmt_ForceSaved()
    ThisWorkbook.Saved = True
End Sub

' Salva todos os módulos
Private Function rModMngmt_ExportAll() As Variant
    rModMngmt_ExportAll = ExportModules(True)
End Function

' Atualiza todos os módulos
Private Function rModMngmt_RefreshAll() As Variant
    rModMngmt_RefreshAll = RefreshModules
End Function

' Importa todos os módulos
Private Sub rModMngmt_ImportAll()
    Dim Mods As New Collection

    Call Mods.Add("rArray")
    Call Mods.Add("rcArrayTable")
    Call Mods.Add("rCallback")
    Call Mods.Add("rChart")
    Call Mods.Add("rCommon")
    Call Mods.Add("rFile")
    Call Mods.Add("rModMngmt")
    Call Mods.Add("rModule")
    Call Mods.Add("rSAP")
    Call Mods.Add("rShape")
    Call Mods.Add("rSheet")
    Call Mods.Add("rTable")
    Call Mods.Add("rUtils")
    Call Mods.Add("rcFileSystem")
    Call Mods.Add("rcKeyValue")
    Call Mods.Add("rcKeyValueCollection")
    Call Mods.Add("rcLookupHelper")
    Call Mods.Add("rcProductDescription")
    Call Mods.Add("rcProgressHelper")
    Call Mods.Add("rcSetCollection")
    Call Mods.Add("rcTreeCollection")
    Call Mods.Add("rcZCO093")
    Call Mods.Add("rcZPL287")
    Call Mods.Add("rcZPL350")
    Call Mods.Add("rcZPLQ010")

    Call ImportModules(Mods)

    Set Mods = Nothing
End Sub

' Importa módulos específicos [...]
Private Sub rArray_Import()
    Call ImportModules("rArray")
End Sub

Private Sub rcArrayTable_Import()
    Call ImportModules("rcArrayTable")
End Sub

Private Sub rCallback_Import()
    Call ImportModules("rCallback")
End Sub

Private Sub rChart_Import()
    Call ImportModules("rChart")
End Sub

Private Sub rCommon_Import()
    Call ImportModules("rCommon")
End Sub

Private Sub rFile_Import()
    Call ImportModules("rFile")
End Sub

Private Sub rSAP_Import()
    Call ImportModules("rSAP")
End Sub

Private Sub rShape_Import()
    Call ImportModules("rShape")
End Sub

Private Sub rSheet_Import()
    Call ImportModules("rSheet")
End Sub

Private Sub rTable_Import()
    Call ImportModules("rTable")
End Sub

Private Sub rUtils_Import()
    Call ImportModules("rUtils")
End Sub

Private Sub rcFileSystem_Import()
    Call ImportModules("rcFileSystem")
End Sub

Private Sub rcKeyValue_Import()
    Call ImportModules("rcKeyValue")
End Sub

Private Sub rcKeyValueCollection_Import()
    Call ImportModules("rcKeyValueCollection")
End Sub

Private Sub rcLookupHelper_Import()
    Call ImportModules("rcLookupHelper")
End Sub

Private Sub rcProductDescription_Import()
    Call ImportModules("rcProductDescription")
End Sub

Private Sub rcProgressHelper_Import()
    Call ImportModules("rcProgressHelper")
End Sub

Private Sub rcSetCollection_Import()
    Call ImportModules("rcSetCollection")
End Sub

Private Sub rcTreeCollection_Import()
    Call ImportModules("rcTreeCollection")
End Sub

Private Sub rcZCO093_Import()
    Call ImportModules("rcZCO093")
End Sub

Private Sub rcZPL287_Import()
    Call ImportModules("rcZPL287")
End Sub

Private Sub rcZPL350_Import()
    Call ImportModules("rcZPL350")
End Sub

Private Sub rcZPLQ010_Import()
    Call ImportModules("rcZPLQ010")
End Sub
