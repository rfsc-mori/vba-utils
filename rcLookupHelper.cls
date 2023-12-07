VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "rcLookupHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Name: rcLookupHelper
' Version: 0.16
' Depends: rCommon,rSheet,rTable,rcKeyValue,rcKeyValueCollection
' Author: Rafael Fillipe Silva
' Description: ...

Public Columns As rcKeyValueCollection
Public Cache As rcKeyValueCollection

' Cria a lista de colunas
Public Sub Process(Optional ByRef Used As Range, Optional ByRef Sheet As Worksheet)
    Dim R As Range

    If Used Is Nothing Then
        If Not Sheet Is Nothing Then
            Set R = GetUsedRange(Sheet)
        End If
    Else
        Set R = Used
    End If

    If Not R Is Nothing Then
        Set Columns = GetColumns(R)
    Else
        Set Columns = Nothing
    End If
End Sub

' Busca a coluna especificada pelo título no cabeçário
Public Function Column(ByVal Header As String) As Range
Attribute Column.VB_UserMemId = 0
    Dim KP As rcKeyValue

    Set Column = Nothing

    If Columns Is Nothing Then
        Exit Function
    End If

    For Each KP In Columns.List
        If TrimComp(KP.Key, Header) Then
            Set Column = KP.Value
            Exit For
        End If
    Next KP
End Function

' Copia um Range para outro na coluna especificada pelo título no cabeçário
Public Function CopyColumn(ByRef Source As Range, ByRef Target As Range, Optional ByVal Header As String, Optional Parse As Boolean = False) As Range
    Dim KP As rcKeyValue

    Set CopyColumn = Nothing

    If Columns Is Nothing Or Source Is Nothing Or Target Is Nothing Then
        Exit Function
    End If

    If Header <> "" Then
        For Each KP In Columns.List
            If TrimComp(KP.Key, Header) Then
                Set CopyColumn = ModifyRangeColumn(Source, KP.Value.Column)
                CopyValues CopyColumn, Target, Parse
                Exit For
            End If
        Next KP
    Else
        CopyValues Source, Target, Parse
    End If
End Function

' Altera um Range para a coluna especificada pelo título no cabeçário
Public Function ChangeColumn(ByRef Source As Range, ByVal Header As String) As Range
    Dim KP As rcKeyValue

    Set ChangeColumn = Nothing

    If Columns Is Nothing Then
        Exit Function
    End If

    For Each KP In Columns.List
        If TrimComp(KP.Key, Header) Then
            Set ChangeColumn = ModifyRangeColumn(Source, KP.Value.Column)
            Exit For
        End If
    Next KP
End Function

' Busca o número da coluna especificada pelo titulo no cabeçário
Public Function ColumnNumber(ByVal Header As String) As Long
    Dim KP As rcKeyValue

    ColumnNumber = 0

    If Columns Is Nothing Then
        Exit Function
    End If

    For Each KP In Columns.List
        If TrimComp(KP.Key, Header) Then
            ColumnNumber = KP.Value.Column
            Exit For
        End If
    Next KP
End Function

' Calcula a diferença para a coluna especificada pelo titulo no cabeçário
Public Function ColumnOffset(ByRef Source As Range, ByVal Header As String) As Long
    Dim KP As rcKeyValue

    ColumnOffset = 0

    If Columns Is Nothing Or Source Is Nothing Then
        Exit Function
    End If

    For Each KP In Columns.List
        If TrimComp(KP.Key, Header) Then
            ColumnOffset = KP.Value.Column - Source.Column + 1
            Exit For
        End If
    Next KP
End Function

' Checa se a coluna especificada pelo titulo existe no cabeçário
Public Function ColumnExists(ByVal Header As String) As Boolean
    Dim KP As rcKeyValue

    ColumnExists = False

    If Columns Is Nothing Then
        Exit Function
    End If

    For Each KP In Columns.List
        If TrimComp(KP.Key, Header) Then
            ColumnExists = True
            Exit For
        End If
    Next KP
End Function

' Retorna uma célula na mesma linha da coluna e valor pesquisado, como no VLookup
Public Function Lookup(ByVal LColumn As String, ByRef LValue As Variant, ByVal RColumn As String, _
                       Optional ByVal ForceIfNothing As Boolean = False) As Range
    Dim CColumn As rcKeyValueCollection
    Dim Row As Range

    Dim Aux As Range

    Dim FirstAddr As String

    Set Lookup = Nothing
    Set Row = Nothing

    If Cache Is Nothing Then
        Set Cache = New rcKeyValueCollection
    End If

    If Cache.Exists(LColumn) Then
        Set CColumn = Cache.Value(LColumn)
    Else
        Set CColumn = Cache.Add(LColumn, New rcKeyValueCollection).Value
    End If

    If CColumn.Exists(LValue) Then
        Set Row = CColumn.Value(LValue)

        If Row Is Nothing And ForceIfNothing = False Then
            GoTo CleanUp
        End If
    End If

    If Row Is Nothing Then
        Set Aux = Column(LColumn)

        If Aux Is Nothing Then
            GoTo CleanUp
        End If

        Set Row = Aux.Find(What:=Trim$(LValue), LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlRows)

        If Not Row Is Nothing Then
            If TrimComp(Row.Value, LValue) Then
                CColumn.Add LValue, Row
                Set Aux = Nothing
            Else
                FirstAddr = Row.Address

                Do
                    If TrimComp(Row.Value, LValue) Then
                        CColumn.Add LValue, Row
                        Set Aux = Nothing
                        Exit Do
                    End If

                    Set Row = Aux.FindNext(Row)
                Loop While FirstAddr <> Row.Address
            End If

            If Not Aux Is Nothing Then
                Set Row = Nothing
            End If
        End If
    End If

    If Not Row Is Nothing Then
        Set Aux = Column(RColumn)

        If Aux Is Nothing Then
            GoTo CleanUp
        End If

        Set Row = ModifyRangeColumn(Row, Aux.Column)
        Set Lookup = Row
    End If

CleanUp:
End Function

' Libera a memória das variáveis
Public Sub Free()
    rcFree Columns
    rcFree Cache
End Sub

