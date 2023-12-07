Attribute VB_Name = "rTable"
Option Explicit

' Name: rTable
' Version: 0.39
' Depends: rCommon,rSheet,rcKeyValueCollection,rcLookupHelper,rcProgressHelper,rcSetCollection,rcTreeCollection
' Author: Rafael Fillipe Silva
' Description: ...

' Checa se tabela existe, caso contrário retorna Nothing
Public Function SafeTable(ByVal Name As String, ByRef Where As Worksheet) As ListObject
    Dim Obj As ListObject

    Set SafeTable = Nothing

    For Each Obj In Where.ListObjects
        If StrComp(Obj.Name, Name, vbTextCompare) = 0 Then
            Set SafeTable = Obj
            Exit Function
        End If
    Next Obj
End Function

' Cria uma tabela no intervalo
Public Function AsTable(ByRef Source As Range, Optional ByVal Name As String, _
                        Optional ByRef Target As Worksheet, _
                        Optional ByVal Overwrite As Boolean = False) As ListObject
    Dim Obj As ListObject

    Set AsTable = Nothing

    If Target Is Nothing Then
        If Source Is Nothing Then
            Exit Function
        End If

        Set Target = Source.Worksheet
    End If

    For Each Obj In Target.ListObjects
        If StrComp(Obj.Name, Name, vbTextCompare) = 0 Then
            If Overwrite Then
                Obj.Delete
            Else
                Set AsTable = Obj
                Exit Function
            End If
        End If
    Next Obj

    Set AsTable = Target.ListObjects.Add(xlSrcRange, Source, , xlYes)

    If Name <> "" Then
        AsTable.Name = Name
    End If
End Function

' Gera um banco de dados a partir de colunas e linhas selecionadas e filtradas
Public Function GetDataTree(ByRef Used As Range, _
                            ByRef InFields As Variant, _
                            ByRef InSums As Variant, _
                            Optional ByRef PFilters As rcKeyValueCollection, _
                            Optional ByRef NFilters As rcKeyValueCollection, _
                            Optional ByVal MergeAction As reMergeAction = reMergeKeep) As rcTreeCollection
    Dim Table As Range
    Dim LH As rcLookupHelper
    Dim Aux As Long
    Dim Fields As Variant
    Dim Sums As Variant
    Dim FieldCount As Long
    Dim SumCount As Long
    Dim KV As rcKeyValue
    Dim PH As rcProgressHelper
    Dim Data As Variant
    Dim Total As Long
    Dim Helper As Long
    Dim Keys() As Variant
    Dim Values() As Variant
    Dim Test As Long
    Dim PTest() As Variant
    Dim NTest() As Variant

    Set Table = Used

    If Table.Rows.Count < 1 Then
        Exit Function
    End If

    Set LH = New rcLookupHelper
    LH.Process Table

    Fields = SMakeArray1d(InFields, 1)
    Sums = SMakeArray1d(InSums, 1)

    Set GetDataTree = New rcTreeCollection
    GetDataTree.Setup Fields, Sums, MergeAction

    For Aux = LBound(Fields) To UBound(Fields)
        If LH.ColumnExists(Fields(Aux)) Then
            Fields(Aux) = LH.ColumnOffset(Table.Cells(1, 1), Fields(Aux))
        Else
            Exit Function
        End If

        FieldCount = FieldCount + 1
    Next Aux

    If FieldCount <= 0 Then
        Exit Function
    End If

    For Aux = LBound(Sums) To UBound(Sums)
        If LH.ColumnExists(Sums(Aux)) Then
            Sums(Aux) = LH.ColumnOffset(Table.Cells(1, 1), Sums(Aux))
        Else
            Exit Function
        End If

        SumCount = SumCount + 1
    Next Aux

    If SumCount <= 0 Then
        Exit Function
    End If

    Data = RangeMid(Table, 1).Value
    Total = UBound(Data, 1)

    ReDim Keys(1 To FieldCount)
    ReDim Values(1 To SumCount)

    If Not PFilters Is Nothing Then
        ReDim PTest(1 To FieldCount + SumCount)

        For Each KV In PFilters
            If LH.ColumnExists(KV.Key) Then
                Aux = LH.ColumnOffset(Table.Cells(1, 1), KV.Key)
                PTest(Aux) = CriteriaTest(Data, KV.Value, Aux)
            Else
                Exit Function
            End If
        Next KV
    End If

    If Not NFilters Is Nothing Then
        ReDim NTest(1 To NFilters.List.Count)

        For Each KV In NFilters
            If LH.ColumnExists(KV.Key) Then
                Aux = LH.ColumnOffset(Table.Cells(1, 1), KV.Key)
                NTest(Aux) = CriteriaTest(Data, KV.Value, Aux)
            Else
                Exit Function
            End If
        Next KV
    End If

    Set PH = New rcProgressHelper
    PH.Prepare 1000

    For Aux = LBound(Data, 1) To UBound(Data, 1)
        PH.DoStep True, Total, Aux

        If Not PFilters Is Nothing Then
            For Helper = LBound(PTest) To UBound(PTest)
                If Not IsEmpty(PTest(Helper)) Then
                    If Not PTest(Helper)(LBound(PTest(Helper)) + Aux - LBound(Data, 1)) Then
                        GoTo Next_Line
                    End If
                End If
            Next Helper
        End If

        If Not NFilters Is Nothing Then
            For Helper = LBound(NTest) To UBound(NTest)
                If Not IsEmpty(NTest(Helper)) Then
                    If Not NTest(Helper)(LBound(NTest(Helper)) + Aux - LBound(Data, 1)) Then
                        GoTo Next_Line
                    End If
                End If
            Next Helper
        End If

        For Helper = LBound(Fields) To UBound(Fields)
            Keys(Helper) = Data(Aux, Fields(Helper))
        Next Helper

        For Helper = LBound(Sums) To UBound(Sums)
            Values(Helper) = Data(Aux, Sums(Helper))
        Next Helper

        GetDataTree.Add Keys, Values

Next_Line:
    Next Aux

    PH.Finish

    GetDataTree.Root.Sort
End Function

' Consolida uma tabela consolidando todos os valores das colunas informadas
' se as colunas "Criteria" são iguais.
Public Function ConsolidateTable(ByRef Table As Range, Sum As Variant, _
                                 Optional Criteria As Variant = Empty, _
                                 Optional Ignore As Variant = Empty) As Variant
    Dim Columns As rcKeyValueCollection
    Dim KV As rcKeyValue

    Dim Aux As Variant
    Dim Helper As Variant
    Dim Status As Variant
    Dim Total As Variant

    Dim Base As Range
    Dim Rows As Range
    Dim Row As Range

    Dim Cursor As Range
    Dim Remove As Variant

    Dim CSum
    Dim CIgnore
    Dim CCriteria

    Dim PH As rcProgressHelper

    ConsolidateTable = 0

    Set Columns = GetColumns(Table)

    If Not IsEmpty(Ignore) Then
        Aux = 0

        For Each Helper In Ignore
            If Columns.Exists(Helper) Then
                Aux = Aux + 1
            End If
        Next Helper

        If Aux > 0 Then
            ReDim CIgnore(1 To Aux)

            Aux = LBound(CIgnore)

            For Each Helper In Ignore
                If Columns.Exists(Helper) Then
                    CIgnore(Aux) = Helper
                    Aux = Aux + 1
                End If
            Next Helper
        End If
    End If

    If Not IsEmpty(Criteria) Then
        Aux = 0

        For Each Helper In Criteria
            If Columns.Exists(Helper) Then
                If Not IsEmpty(Ignore) Then
                    For Status = LBound(CIgnore) To UBound(CIgnore)
                        If Helper = CIgnore(Status) Then
                            Helper = Empty
                            Exit For
                        End If
                    Next Status
                End If

                If Not IsEmpty(Helper) Then
                    Aux = Aux + 1
                End If
            End If
        Next Helper

        If Aux > 0 Then
            ReDim CCriteria(1 To Aux)

            Aux = LBound(CCriteria)

            For Each Helper In Criteria
                If Columns.Exists(Helper) Then
                    If Not IsEmpty(Ignore) Then
                        For Status = LBound(CIgnore) To UBound(CIgnore)
                            If Helper = CIgnore(Status) Then
                                Helper = Empty
                                Exit For
                            End If
                        Next Status
                    End If

                    If Not IsEmpty(Helper) Then
                        CCriteria(Aux) = Helper
                        Aux = Aux + 1
                    End If
                End If
            Next Helper
        End If
    End If

    Aux = 0

    For Each Helper In Sum
        If Columns.Exists(Helper) Then
            If Not IsEmpty(Ignore) Then
                For Status = LBound(CIgnore) To UBound(CIgnore)
                    If Helper = CIgnore(Status) Then
                        Helper = Empty
                        Exit For
                    End If
                Next Status
            End If

            If Not IsEmpty(Helper) Then
                Aux = Aux + 1
            End If
        End If
    Next Helper

    If Aux <= 0 Or IsEmpty(Sum) Then
        Exit Function
    Else
        ReDim CSum(1 To Aux)

        Aux = LBound(CSum)

        For Each Helper In Sum
            If Columns.Exists(Helper) Then
                If Not IsEmpty(Ignore) Then
                    For Status = LBound(CIgnore) To UBound(CIgnore)
                        If Helper = CIgnore(Status) Then
                            Helper = Empty
                            Exit For
                        End If
                    Next Status
                End If

                If Not IsEmpty(Helper) Then
                    CSum(Aux) = Helper
                    Aux = Aux + 1
                End If
            End If
        Next Helper
    End If

    If IsEmpty(Criteria) Then
        Aux = 0

        For Each KV In Columns
            Helper = KV.Key

            If Not IsEmpty(Ignore) Then
                For Status = LBound(CIgnore) To UBound(CIgnore)
                    If Helper = CIgnore(Status) Then
                        Helper = Empty
                        Exit For
                    End If
                Next Status
            End If

            If Not IsEmpty(Helper) Then
                Aux = Aux + 1
            End If
        Next KV

        ReDim CCriteria(1 To Aux)
        Aux = LBound(CCriteria)

        For Each KV In Columns
            Helper = KV.Key

            If Not IsEmpty(Ignore) Then
                For Status = LBound(CIgnore) To UBound(CIgnore)
                    If Helper = CIgnore(Status) Then
                        Helper = Empty
                        Exit For
                    End If
                Next Status
            End If

            If Not IsEmpty(Helper) Then
                If Base Is Nothing Then
                    Set Base = KV.Value
                End If

                CCriteria(Aux) = Helper
                Aux = Aux + 1
            End If
        Next KV
    Else
        Set Base = Columns.Value(CCriteria(1))
    End If

    For Aux = LBound(CCriteria) To UBound(CCriteria)
        CCriteria(Aux) = Columns.Value(CCriteria(Aux)).Column - Base.Column
    Next Aux

    For Aux = LBound(CSum) To UBound(CSum)
        CSum(Aux) = Columns.Value(CSum(Aux)).Column - Base.Column
    Next Aux

    Set Rows = Base.Rows

    Status = Rows.Value
    Total = Rows.Count

    Set PH = New rcProgressHelper
    PH.Prepare 1000

    For Aux = Base.Row + 1 To Total
        Helper = Status(Aux, 1)

        If IsEmpty(Helper) Then
            Set Cursor = Rows(Aux).End(xlDown)
            Aux = Cursor.Row

            If Aux > Total Then
                Exit For
            End If
        End If

        PH.DoStep True, Total, Aux

        Set Row = Rows(Aux)
        Set Cursor = Rows(Aux - 1)

        If Helper <> Cursor.Value Then
            Set Cursor = Base.Find(Helper, Cursor, xlFormulas, xlWhole, xlByRows, xlPrevious)
        End If

        Do While Not Cursor Is Nothing
            If Cursor.Row < Aux Then
                Remove = 1

                For Helper = LBound(CCriteria) To UBound(CCriteria)
                    If CCriteria(Helper) <> 0 Then
                        If Cursor.Cells(ColumnIndex:=1 + CCriteria(Helper)).Value <> _
                           Row.Cells(ColumnIndex:=1 + CCriteria(Helper)).Value Then
                            Remove = 0
                            Exit For
                        End If
                    End If
                Next Helper

                If Remove = 1 Then
                    For Remove = LBound(CSum) To UBound(CSum)
                        If IsNumeric(Cursor.Cells(ColumnIndex:=1 + CSum(Remove)).Value) Then
                            Cursor.Cells(ColumnIndex:=1 + CSum(Remove)).Value = _
                                Cursor.Cells(ColumnIndex:=1 + CSum(Remove)).Value + _
                                Row.Cells(ColumnIndex:=1 + CSum(Remove)).Value
                        End If
                    Next Remove

                    Rows(Aux).ClearContents

                    Exit Do
                Else
                    Set Cursor = Base.FindPrevious(Cursor)
                End If
            Else
                Set Cursor = Nothing
            End If
        Loop
    Next Aux

    Set Rows = SafeSpecialCells(Rows, xlCellTypeBlanks)

    If Not Rows Is Nothing Then
        ConsolidateTable = Rows.Count
        Rows.EntireRow.Delete
    End If

    PH.Finish True
End Function

' Converte um range com várias áreas em uma array de valores
Public Function TableToSet(ByRef Table As Range) As rcSetCollection
    Dim Column As Range
    Dim Area As Range

    Set TableToSet = New rcSetCollection

    For Each Area In Table.Areas
        For Each Column In Area.Columns
            TableToSet.Add Application.Transpose(Column.Value)
        Next Column
    Next Area
End Function
