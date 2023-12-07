Attribute VB_Name = "rUtils"
Option Explicit

' Name: rUtils
' Version: 0.11
' Depends: rFile
' Author: Rafael Fillipe Silva
' Description: ...

' NOTA: Utilidades para serem executadas com a função Executar Sub (F5).

' Copia módulos que não sejam da biblioteca "r" de um arquivo para outro
' Ex: Atualizar o módulo controle em um arquivo sem alterar seus dados.
Private Sub CopyModules()
    Dim File As Variant
    Dim Book As Workbook

    File = AskForFile

    If File = False Then
        Exit Sub
    End If

    Set Book = SafeOpen(File, RW:=True)

    If Book Is Nothing Then
        Exit Sub
    End If

    File = AskSaveFile(Book.Name, Book.Path & "\" & Book.Name & ".xlsm", 2)

    If File <> False Then
        Call CopyModulesTo(Book, True, True, ThisWorkbook)
        Call Book.SaveAs(File, xlOpenXMLWorkbookMacroEnabled)
    End If

    Call Book.Close(False)
    Set Book = Nothing
End Sub

Private Sub ResetStatusBar()
    Application.StatusBar = Empty
End Sub

Private Sub ResetApplicationFlags()
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.IgnoreRemoteRequests = False
    Application.Interactive = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Private Sub InfoCell()
    Debug.Print ActiveCell.Font.ColorIndex
End Sub

