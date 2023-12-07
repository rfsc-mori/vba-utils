Attribute VB_Name = "rSAP"
Option Explicit

' Name: rSAP
' Version: 0.11
' Depends: rSheet
' Author: Rafael Fillipe Silva
' Description: ...

' Conserta as colunas movidas durante a exportação de planilhas do SAP
' KeyCols: Colunas referências que sempre terão valores ao menos nas primeiras linhas
Public Sub SAP_FixSheetColumns(ByRef Sheet As Worksheet, Optional ByRef KeyCols As Range)
    Dim Used As Range

    Dim Column As Range
    Dim Columns As Range

    Dim MinRow As Range
    Dim Aux As Range

    Dim i As Variant

    Dim Left
    Dim Right
    Dim Row As Variant

    Dim Helper As Boolean

    Set Used = GetUsedRange(Sheet)

    If Used.Rows.Count = 1 Then
        Exit Sub
    End If

    If Not IsMissing(KeyCols) Then
        Set Columns = KeyCols.Rows(1).Columns

        For Each Column In Columns
            If Column.Offset(1).Value <> "" Then
                GoTo Fix_Columns
            Else
                Set Aux = Column.End(xlDown)

                If MinRow Is Nothing Then
                    Set MinRow = Aux
                ElseIf Aux.Row < MinRow.Row Then
                    Set MinRow = Aux
                End If
            End If
        Next Column

        Set Columns = Used.Rows(1).Columns

        For Each Column In Columns
            Set Aux = Column.End(xlDown)

            If Aux.Row < MinRow.Row Then
                For i = 1 To MinRow.Row - Aux.Row
                    Column.Value = Trim$(Trim$(Column.Value) & " " & Trim$(Column.Offset(i).Value))
                    Column.Offset(i).ClearContents
                Next i
            ElseIf Aux.Row > MinRow.Row Then
                For i = 1 To MinRow.Row - 2
                    Column.Value = Trim$(Trim$(Column.Value) & " " & Trim$(Column.Offset(i).Value))
                    Column.Offset(i).ClearContents
                Next i
            End If
        Next Column

        For i = 1 To (MinRow.Row - 1) - Used.Row
            Used.Rows(i + 1).Delete
        Next i
    End If

Fix_Columns:
    Set Aux = Nothing

    For i = 1 To Used.Columns.Count
        Set Column = Used.Columns(i)

        With Column.Rows(1)
            If .Value <> "" Then
                Do While i + 1 <= Used.Columns.Count And Used.Columns(i + 1).Rows(1).Value = ""
                    Left = Application.Transpose(Used.Columns(i).Rows.Value)
                    Right = Application.Transpose(Used.Columns(i + 1).Rows.Value)

                    Helper = True

                    If Application.CountA(Right) > 1 Then
                        Helper = False
                    End If

                    If Helper Then
                        For Row = LBound(Left) To UBound(Left)
                            If IsEmpty(Left(Row)) And Not IsEmpty(Right(Row)) Then
                                Used.Columns(i).Rows(Row).Value = Right(Row)
                            End If
                        Next Row

                        If Aux Is Nothing Then
                            Set Aux = Used.Columns(i + 1)
                        Else
                            Set Aux = Union(Aux, Used.Columns(i + 1))
                        End If

                        i = i + 1
                    Else
                        Exit Do
                    End If
                Loop
            ElseIf i - 1 >= 1 And Used.Columns(i).Rows(1).Value = "" Then
                .Value = Used.Columns(i + 1).Rows(1).Value

                If Aux Is Nothing Then
                    Set Aux = Used.Columns(i + 1)
                Else
                    Set Aux = Union(Aux, Used.Columns(i + 1))
                End If
            End If
        End With
Next_I:
    Next i

    If Not Aux Is Nothing Then
        Aux.Delete
    End If
End Sub

' Se várias colunas são exportadas pelo SAP como apenas uma com espaços
' ao invés de novas colunas, trata-os gerando novas colunas.
Public Sub SAP_FixColumnsBySpaces(ByRef Sheet As Worksheet, ByRef KeyCols As Range)
    Dim Column As Range
    Dim Columns As Range

    Dim Row As Range
    Dim Used As Range

    Dim i As Variant
    Dim lvl As Variant

    Set Used = GetUsedRange(Sheet)
    Set Columns = KeyCols.Rows(1).Columns

    lvl = 0

    For Each Column In Columns
        Set Column = ModifyRangeColumn(Used, Column.Column)

        For Each Row In Column.Rows
            i = (InStr(Row.Value, Left$(Trim$(Row.Value), 1)) - 1)

            If i > 0 Then
                If i > lvl Then
                    Row.Offset(0, i).EntireColumn.Insert xlShiftToRight
                    lvl = i
                    Column.Offset(0, i).Rows(1).Value = Left$(Column.Rows(1).Value, 3) & " " & (i + 1)
                End If

                Row.Offset(0, i).Value = Trim$(Row.Value)
                Row.Value = ""
            End If
        Next Row
    Next Column
End Sub
