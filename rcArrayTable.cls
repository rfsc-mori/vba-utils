VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "rcArrayTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Name: rcArrayTable
' Version: 0.94
' Depends: rArray,rcKeyValue,rcKeyValueCollection,rcSetCollection
' Author: Rafael Fillipe Silva
' Description: ...

Public Data, Filtered

Public Dimensions As Long

Public yLower, yUpper, ySize As Long
Public xLower, xUpper, xSize As Long

Public yFLower, yFUpper, yFSize As Long
Public xFLower, xFUpper, xFSize As Long

Public UseHeaders As Boolean
Public Columns As rcKeyValueCollection

Public ColumnFilter As rcKeyValueCollection
Public PositiveFilter As rcKeyValueCollection
Public NegativeFilter As rcKeyValueCollection
Public ValueFilter As rcKeyValueCollection
Public RankFilter As rcKeyValueCollection
Public RankFilterOnColumns As rcKeyValueCollection

Public ColumnReorder As rcKeyValueCollection
Public EnsureValues As rcKeyValueCollection
Public EnsureValuesOnColumns As rcKeyValueCollection
Public RenameColumn As rcKeyValueCollection

Public MergeFilter As rcKeyValueCollection
Public MergeActions As rcKeyValueCollection

Public ColumnFilterMask As Variant
Public RowFilterMask As Variant
Public DataFilterMask As Variant
Public ValueFilterMask As Variant
Public RankFilterMask As Variant
Public RankFilterColumnsMask As Variant

Public ColumnReorderMask As Variant
Public RowReorderMask As Variant
Public EnsureValuesMask As Variant
Public EnsureValuesColumnsMask As Variant
Public EnsureValuesRowsMask As Variant
Public RenameColumnMask As Variant

Public MergeFilterMask As Variant

Public PositiveAnd As Boolean
Public PositiveByRow As Boolean
Public NegativeAnd As Boolean
Public NegativeByRow As Boolean

Private Sub Class_Initialize()
    PositiveAnd = False
    PositiveByRow = False
    NegativeAnd = False
    NegativeByRow = False
End Sub

Public Function Init(ByRef Arr, Optional ParseHeaders As Boolean = True, _
                     Optional InitFilters As Boolean = True, _
                     Optional InitNormalizers As Boolean = True, _
                     Optional InitMerge As Boolean = True) As Boolean
    Dim Aux As Long
    Dim Col As Variant
    Dim Counter As Long

    Dimensions = ArrayDimensions(Arr)

    If Dimensions <> 2 Then
        Exit Function
    End If

    yLower = LBound(Arr, 1)
    yUpper = UBound(Arr, 1)
    ySize = (yUpper - yLower) + 1

    xLower = LBound(Arr, 2)
    xUpper = UBound(Arr, 2)
    xSize = (xUpper - xLower) + 1

    yFLower = 0
    yFUpper = 0
    yFSize = 0

    xFLower = 0
    xFUpper = 0
    xFSize = 0

    Data = Arr
    Filtered = Data

    UseHeaders = ParseHeaders

    If UseHeaders Then
        Set Columns = New rcKeyValueCollection

        For Aux = xLower To xUpper
            If Not IsEmpty(Data(1, Aux)) Then
                Col = Data(1, Aux)

                If VarType(Col) = vbString Then
                    Col = NormalizedString(Col)
                End If
            Else
                Col = ""
            End If

            If Not Columns.Exists(Col) And Col <> "" Then
                Call Columns.Add(Col, Aux, True)
            Else
                Counter = 1

                Do While Columns.Exists(Col) Or Col = ""
                    Col = Data(1, Aux)

                    If VarType(Col) = vbString Then
                        Col = NormalizedString(Col)
                    End If

                    Col = Col & "[" & Counter & "]"

                    If Not Columns.Exists(Col) And Col <> "" Then
                        Call Columns.Add(Col, Aux, True)
                        Exit Do
                    Else
                        Counter = Counter + 1
                    End If
                Loop
            End If
        Next Aux
    End If

    If InitFilters Then
        Set ColumnFilter = New rcKeyValueCollection
        Set PositiveFilter = New rcKeyValueCollection
        Set NegativeFilter = New rcKeyValueCollection
        Set ValueFilter = New rcKeyValueCollection
        Set RankFilter = New rcKeyValueCollection
        Set RankFilterOnColumns = New rcKeyValueCollection
    End If

    If InitNormalizers Then
        Set ColumnReorder = New rcKeyValueCollection
        RowReorderMask = Empty
        Set EnsureValues = New rcKeyValueCollection
        Set EnsureValuesOnColumns = New rcKeyValueCollection
        Set RenameColumn = New rcKeyValueCollection
    End If

    If InitMerge Then
        Set MergeFilter = New rcKeyValueCollection
        Set MergeActions = New rcKeyValueCollection
    End If
End Function

Public Function ColumnExists(ByRef Column As Variant) As Boolean
    ColumnExists = Columns.Exists(NormalizedString(Column))
End Function

Public Function GetColumnIndex(ByRef Column As Variant, Optional ByVal index As Variant) As Variant
    GetColumnIndex = Columns(NormalizedString(Column), index)
End Function

Public Sub SetSortColumns()
    Dim x, z As Long
    Dim Aux As Long

    If Not UseHeaders Then
        Exit Sub
    End If

    ReDim ColumnReorderMask(xLower To xUpper) As Long

    For x = xLower To xUpper
        ColumnReorderMask(x) = x
    Next x

    ' Sort
    For x = 1 To Columns.List.Count
        For z = x + 1 To Columns.List.Count
            If Columns.Item(x).Key > Columns.Item(z).Key Then
                Aux = ColumnReorderMask(z)
                ColumnReorderMask(z) = ColumnReorderMask(x)
                ColumnReorderMask(x) = Aux
            End If
        Next z
    Next x

    ' Store
    If Not ColumnReorder Is Nothing Then
        Call ColumnReorder.Clear

        For x = xLower To xUpper
            For z = xLower To xUpper
                If x = ColumnReorderMask(z) Then
                    ColumnReorder.SetValue Columns.Item(x).Key, z
                    Exit For
                End If
            Next z
        Next x
    End If
End Sub

Public Sub SetColumnOrder(Columns As Variant, Optional Filter As Boolean = False)
    Dim Aux As Variant
    Dim Cols As Variant
    Dim x, z As Long

    If Filter Then
        If Not ColumnFilter Is Nothing Then
            Cols = SMakeArray1d(Columns, 1)

            For z = LBound(Cols) To UBound(Cols)
                Call ColumnFilter.SetPairs(Cols(z), True)
            Next z
        End If
    End If

    If Not ColumnReorder Is Nothing Then
        Aux = SMakeArray1d(Columns, 1)

        For x = LBound(Aux) To UBound(Aux)
            Cols = SMakeArray1d(Aux(x))

            For z = LBound(Cols) To UBound(Cols)
                Call ColumnReorder.SetValue(Cols(z), x)
            Next z
        Next x
    End If
End Sub

Public Sub SetSortRows(Optional ByRef InColumns As Variant, Optional ByRef InOrders As Variant)
    Dim x, y, z As Long
    Dim Aux As Variant
    Dim Cols As Variant
    Dim Orders As Variant

    If Not IsMissing(InColumns) Then
        Cols = SMakeArray1d(InColumns)

        If IsEmpty(Cols) Then
            Exit Sub
        End If

        If Not IsMissing(InOrders) Then
            Orders = SMakeArray1d(InOrders)

            If IsArrayInvalid(Orders) Then
                ReDim Orders(LBound(Cols) To UBound(Cols)) As Variant

                For x = LBound(Orders) To UBound(Orders)
                    Orders(x) = True
                Next x
            Else
                Aux = SArrayCount(Orders)

                If SArrayCount(Orders) < SArrayCount(Cols) Then
                    ReDim Preserve Orders(LBound(Cols) To UBound(Cols)) As Variant

                    For x = LBound(Orders) + Aux To UBound(Orders)
                        Orders(x) = Orders(Aux - 1)
                    Next x
                End If
            End If
        Else
            ReDim Orders(LBound(Cols) To UBound(Cols)) As Variant

            For x = LBound(Orders) To UBound(Orders)
                Orders(x) = True
            Next x
        End If

        For x = LBound(Cols) To UBound(Cols)
            If ColumnExists(Cols(x)) Then
                Cols(x) = GetColumnIndex(Cols(x))
            Else
                Exit Sub
            End If
        Next x
    Else
        Cols = Columns.Values.ToArray

        ReDim Orders(LBound(Cols) To UBound(Cols)) As Variant

        For x = LBound(Orders) + Aux + 1 To UBound(Orders)
            Orders(x) = True
        Next x
    End If

    ReDim RowReorderMask(yLower To yUpper) As Long

    For y = yLower To yUpper
        RowReorderMask(y) = y
    Next y

    y = yLower
    z = yUpper

    If UseHeaders Then
        y = y + 1
    End If

    'Sort
    SetSortRows_Helper Cols, Orders, y, z
End Sub

Private Sub SetSortRows_Helper(ByRef Cols As Variant, ByRef Orders As Variant, Optional ByVal lower As Variant, _
                               Optional ByVal upper As Variant)
    Dim x, h, y, L, R As Long
    Dim Aux As Variant

    For x = LBound(Cols) To UBound(Cols)
        If x = LBound(Cols) Then
            Call Quicksort2d(Data, Cols(x), lower, upper, Orders(x), Mask:=RowReorderMask)
        Else
            L = lower
            R = upper

            Aux = Empty

            Do While L <= upper
                For h = LBound(Cols) To x - 1
                    If R - L >= 1 Then
                        For y = L To R
                            If IsEmpty(Aux) Then
                                L = y

                                Aux = Data(RowReorderMask(y), Cols(h))
                            Else
                                If Data(RowReorderMask(y), Cols(h)) <> Aux Then
                                    R = y - 1
                                    Exit For
                                End If
                            End If
                        Next y

                        Aux = Empty
                    Else
                        Exit For
                    End If
                Next h

                If R - L >= 1 Then
                    Call Quicksort2d(Data, Cols(x), L, R, Orders(x), Mask:=RowReorderMask)
                End If

                L = R + 1
                R = upper

                Aux = Empty
            Loop
        End If
    Next x
End Sub

Public Sub GenFilterMask(Optional ByRef DataMask As Variant)
    Dim HeaderOffset As Long
    Dim KV As rcKeyValue
    Dim Aux, Helper, Ensurer As Variant
    Dim Column As Variant
    Dim y, x As Long
    Dim z, i As Long
    Dim MergeHelper
    Dim MergeStack

    ReDim ColumnFilterMask(xLower To xUpper) As Boolean
    ReDim RowFilterMask(yLower To yUpper) As Boolean
    ReDim DataFilterMask(yLower To yUpper, xLower To xUpper) As Boolean
    ReDim ValueFilterMask(yLower To yUpper, xLower To xUpper) As Variant
    ReDim RankFilterMask(yLower To yUpper, xLower To xUpper) As Variant
    ReDim RankFilterColumnsMask(xLower To xUpper) As Variant

    ReDim ColumnReorderMask(xLower To xUpper) As Long
    ReDim EnsureValuesMask(yLower To yUpper, xLower To xUpper) As Variant
    ReDim EnsureValuesColumnsMask(xLower To xUpper) As Variant
    ReDim EnsureValuesRowsMask(yLower To yUpper) As Variant
    ReDim RenameColumnMask(xLower To xUpper) As Variant

    ReDim MergeFilterMask(yLower To yUpper) As Long

    If Not IsMissing(DataMask) Then
        Call ArrayCopy2d(DataMask, DataFilterMask)
    End If

    If UseHeaders Then
        HeaderOffset = 1
    Else
        HeaderOffset = 0
    End If

    If Not ColumnFilter Is Nothing Then
        Aux = False

        For Each KV In ColumnFilter
            If KV.Value = True Then
                Aux = True
                Exit For
            End If
        Next KV

        If Aux Then
            For Aux = xLower To xUpper
                ColumnFilterMask(Aux) = True
            Next Aux
        End If

        For Each KV In ColumnFilter
            Column = Empty

            Helper = SMakeArray1d(KV.Key)

            For x = LBound(Helper) To UBound(Helper)
                If UseHeaders Then
                    If ColumnExists(Helper(x)) Then
                        Column = GetColumnIndex(Helper(x))

                        If Not DataFilterMask(yLower, Column) Then
                            If Not KV.Value Then
                                DataFilterMask(yLower, Column) = True
                                ColumnFilterMask(Column) = True
                                Exit For
                            Else
                                DataFilterMask(yLower, Column) = False
                                ColumnFilterMask(Column) = False
                            End If
                        Else
                            ColumnFilterMask(Column) = True
                            Exit For
                        End If
                    End If
                Else
                    Column = Helper(x)

                    If Not DataFilterMask(yLower, Column) Then
                        If Not KV.Value Then
                            DataFilterMask(yLower, Column) = True
                            ColumnFilterMask(Column) = True
                            Exit For
                        Else
                            DataFilterMask(yLower, Column) = False
                            ColumnFilterMask(Column) = False
                        End If
                    Else
                        ColumnFilterMask(Column) = True
                        Exit For
                    End If
                End If
            Next x
        Next KV
    End If

    If Not PositiveFilter Is Nothing Then
        For y = yLower + HeaderOffset To yUpper
            Helper = PositiveByRow

            For Each KV In PositiveFilter
                Column = Empty

                If UseHeaders Then
                    If ColumnExists(KV.Key) Then
                        Column = GetColumnIndex(KV.Key)
                    End If
                Else
                    Column = KV.Key
                End If

                If Not IsEmpty(Column) Then
                    If SArrayCheck(CriteriaTest(Data(y, Column), KV.Value, , PositiveAnd), True) <> 1 Or DataFilterMask(y, Column) Then
                        If Not PositiveByRow Then
                            DataFilterMask(y, Column) = True
                            RowFilterMask(y) = True
                        End If
                    Else
                        If PositiveByRow Then
                            Helper = False
                        End If
                    End If
                End If
            Next KV

            If PositiveByRow And Helper = True Then
                For Each KV In PositiveFilter
                    Column = Empty

                    If UseHeaders Then
                        If ColumnExists(KV.Key) Then
                            Column = GetColumnIndex(KV.Key)
                        End If
                    Else
                        Column = KV.Key
                    End If

                    If Not IsEmpty(Column) Then
                        DataFilterMask(y, Column) = True
                    End If
                Next KV

                RowFilterMask(y) = True
            End If
        Next y
    End If

    If Not NegativeFilter Is Nothing Then
        For y = yLower + HeaderOffset To yUpper
            Helper = NegativeByRow

            For Each KV In NegativeFilter
                Column = Empty

                If UseHeaders Then
                    If ColumnExists(KV.Key) Then
                        Column = GetColumnIndex(KV.Key)
                    End If
                Else
                    Column = KV.Key
                End If

                If Not IsEmpty(Column) Then
                    If SArrayCheck(CriteriaTest(Data(y, Column), KV.Value, , NegativeAnd), False) <> 1 Or DataFilterMask(y, Column) Then
                        If Not NegativeByRow Then
                            DataFilterMask(y, Column) = True
                            RowFilterMask(y) = True
                        End If
                    Else
                        If NegativeByRow Then
                            Helper = False
                        End If
                    End If
                End If
            Next KV

            If NegativeByRow And Helper = True Then
                For Each KV In NegativeFilter
                    Column = Empty

                    If UseHeaders Then
                        If ColumnExists(KV.Key) Then
                            Column = GetColumnIndex(KV.Key)
                        End If
                    Else
                        Column = KV.Key
                    End If

                    If Not IsEmpty(Column) Then
                        DataFilterMask(y, Column) = True
                    End If
                Next KV

                RowFilterMask(y) = True
            End If
        Next y
    End If

    If Not ValueFilter Is Nothing Then
        For Each KV In ValueFilter
            Column = Empty

            If UseHeaders Then
                If ColumnExists(KV.Key) Then
                    Column = GetColumnIndex(KV.Key)
                End If
            Else
                Column = KV.Key
            End If

            If Not IsEmpty(Column) Then
                For y = yLower + HeaderOffset To yUpper
                    If IsEmpty(ValueFilterMask(y, Column)) Then
                        If Not RowFilterMask(y) And Not DataFilterMask(y, Column) Then
                            Aux = SMakeArray1d(KV.Value, 1)

                            If Not IsArray(Aux) Then
                                Aux = Array(Aux)
                            End If

                            If IsArray(Aux(LBound(Aux))) Then
                                For x = LBound(Aux) To UBound(Aux)
                                    If VariantComp(Aux(x)(LBound(Aux(x))), Data(y, Column)) Then
                                        ValueFilterMask(y, Column) = Aux(x)(LBound(Aux(x)) + 1)
                                        Exit For
                                    ElseIf SingleCriteriaTest(Data(y, Column), Aux(x)(LBound(Aux(x)))) Then
                                        ValueFilterMask(y, Column) = Aux(x)(LBound(Aux(x)) + 1)
                                        Exit For
                                    End If
                                Next x
                            Else
                                If VariantComp(Aux(1), Data(y, Column)) Then
                                    ValueFilterMask(y, Column) = Aux(2)
                                ElseIf SingleCriteriaTest(Data(y, Column), Aux(1)) Then
                                    ValueFilterMask(y, Column) = Aux(2)
                                ElseIf IsEmpty(Data(y, Column)) Then
                                    ValueFilterMask(y, Column) = KV.Value
                                End If
                            End If
                        End If
                    End If
                Next y
            End If
        Next KV
    End If

    If Not RankFilterOnColumns Is Nothing Then
        Aux = False

        For Each KV In RankFilterOnColumns
            If KV.Value = True Then
                Aux = True
                Exit For
            End If
        Next KV

        If Aux Then
            For Aux = xLower To xUpper
                RankFilterColumnsMask(Aux) = True
            Next Aux
        End If

        For Each KV In RankFilterOnColumns
            Column = Empty

            If UseHeaders Then
                If ColumnExists(KV.Key) Then
                    Column = GetColumnIndex(KV.Key)
                End If
            Else
                Column = KV.Key
            End If

            If Not IsEmpty(Column) Then
                If Not KV.Value Then
                    RankFilterColumnsMask(Column) = True
                Else
                    RankFilterColumnsMask(Column) = False
                End If
            End If
        Next KV

        For x = xLower To xUpper
            If ColumnFilterMask(x) Then
                RankFilterColumnsMask(x) = True
            End If
        Next x
    End If

    If Not RankFilter Is Nothing Then
        For Each KV In RankFilter
            Column = Empty

            If UseHeaders Then
                If ColumnExists(KV.Key) Then
                    Column = GetColumnIndex(KV.Key)
                End If
            Else
                Column = KV.Key
            End If

            If Not IsEmpty(Column) Then
                Aux = SMakeArray1d(KV.Value)

                RankFilterColumnsMask(Column) = True

                For z = LBound(Aux) To UBound(Aux)
                    For y = yLower + HeaderOffset To yUpper
                        If Not RowFilterMask(y) And Not DataFilterMask(y, Column) Then
                            If IsEmpty(RankFilterMask(y, Column)) Or RankFilterMask(y, Column) < UBound(Aux) - z Then
                                If VariantComp(Data(y, Column), Aux(z)) Then
                                    RankFilterMask(y, Column) = UBound(Aux) - z
                                End If
                            End If
                        End If
                    Next y
                Next z
            End If
        Next KV

        If RankFilter.Count > 0 Then
            For y = yLower + HeaderOffset To yUpper
                If Not RowFilterMask(y) Then
                    For z = yLower + HeaderOffset To yUpper
                        If Not RowFilterMask(z) Then
                            Aux = True

                            For x = xLower To xUpper
                                If Not RankFilterColumnsMask(x) Then
                                    If Not VariantComp(Data(y, x), Data(z, x)) Then
                                        Aux = False
                                        Exit For
                                    End If
                                End If
                            Next x

                            If Aux = True Then
                                For x = xLower To xUpper
                                    If RankFilterMask(z, x) > RankFilterMask(y, x) Then
                                        RowFilterMask(y) = True
                                        DataFilterMask(y, x) = True
                                    End If
                                Next x
                            End If
                        End If
                    Next z
                End If
            Next y
        End If
    End If

    If Not ColumnReorder Is Nothing Then
        For x = xLower To xUpper
            ColumnReorderMask(x) = x
        Next x

        For Each KV In ColumnReorder
            Column = Empty

            If UseHeaders Then
                If ColumnExists(KV.Key) Then
                    Column = GetColumnIndex(KV.Key)
                End If
            Else
                Column = KV.Key
            End If

            If Not IsEmpty(Column) Then
                For x = xLower To xUpper
                    If Not ColumnFilterMask(x) Then
                        If ColumnReorderMask(x) = Column Then
                            If KV.Value < xLower Then
                                z = xLower
                            ElseIf KV.Value > xUpper Then
                                z = xUpper
                            Else
                                z = KV.Value
                            End If

                            If ColumnReorderMask(x) <> z Then
                               Aux = ColumnReorderMask(z)
                               ColumnReorderMask(z) = ColumnReorderMask(x)
                               ColumnReorderMask(x) = Aux
                            End If

                            Exit For
                        End If
                    End If
                Next x
            End If
        Next KV
    End If

    If Not RenameColumn Is Nothing Then
        For x = xLower To xUpper
            RenameColumnMask(x) = Empty
        Next x

        For Each KV In RenameColumn
            Helper = SMakeArray1d(KV.Key)

            For x = LBound(Helper) To UBound(Helper)
                Column = Empty

                If UseHeaders Then
                    If ColumnExists(Helper(x)) Then
                        Column = GetColumnIndex(Helper(x))
                    End If
                Else
                    Column = Helper(x)
                End If

                If Not IsEmpty(Column) Then
                    RenameColumnMask(Column) = KV.Value
                    Exit For
                End If
            Next x
        Next KV
    End If

    If Not MergeFilter Is Nothing Then
        MergeHelper = MergeFilter.UniqueKeys.ToArray

        If Not IsEmpty(MergeHelper) Then
            For Aux = LBound(MergeHelper) To UBound(MergeHelper)
                MergeHelper(Aux) = Columns.Value(MergeHelper(Aux), Empty)
            Next Aux

            Helper = ArraySelfMatch2d(Data, MergeHelper, yLower + HeaderOffset, yUpper)

            If Not IsEmpty(Helper) Then
                MergeFilterMask = Helper

                For y = yLower To yUpper
                    Aux = MergeFilterMask(y)

                    If Aux <> 0 Then
                        If RowFilterMask(Aux) Then
                            For z = y + 1 To yUpper
                                If MergeFilterMask(z) = Aux And Not RowFilterMask(z) Then
                                    For i = yLower To yUpper
                                        If MergeFilterMask(i) = Aux Then
                                            MergeFilterMask(i) = z
                                        End If
                                    Next i

                                    Exit For
                                End If
                            Next z
                        End If
                    End If
                Next y
            End If
        End If
    End If

    If Not EnsureValuesOnColumns Is Nothing Then
        Aux = False

        For Each KV In EnsureValuesOnColumns
            If KV.Value = True Then
                Aux = True
                Exit For
            End If
        Next KV

        If Aux Then
            For Aux = xLower To xUpper
                EnsureValuesColumnsMask(Aux) = True
            Next Aux
        End If

        For Each KV In EnsureValuesOnColumns
            Helper = SMakeArray1d(KV.Key)

            For x = LBound(Helper) To UBound(Helper)
                Column = Empty

                If UseHeaders Then
                    If ColumnExists(Helper(x)) Then
                        Column = GetColumnIndex(Helper(x))
                    End If
                Else
                    Column = Helper(x)
                End If

                If Not IsEmpty(Column) Then
                    If Not KV.Value Then
                        EnsureValuesColumnsMask(Column) = True
                    Else
                        EnsureValuesColumnsMask(Column) = False
                    End If

                    Exit For
                End If
            Next x
        Next KV
    End If

    If Not EnsureValues Is Nothing Then
        ''' FIX ME '''
        If EnsureValues.Count > 1 Then
            Stop ' Função incompleta, erro ao gerenciar mais de uma coluna.
            End
        End If
        ''' FIX ME '''

        For Each KV In EnsureValues
            Column = Empty

            If UseHeaders Then
                If ColumnExists(KV.Key) Then
                    Column = GetColumnIndex(KV.Key)
                End If
            Else
                Column = KV.Key
            End If

            If Not IsEmpty(Column) Then
                EnsureValuesColumnsMask(Column) = True

                Helper = SMakeArray1d(KV.Value)

                For y = yLower + HeaderOffset To yUpper
                    If EnsureValuesMask(y, Column) <> -1 And RowFilterMask(y) = False Then
                        Ensurer = Helper

                        For z = y To yUpper
                            If RowFilterMask(z) = False Then
                                i = (xUpper - xLower + 1)

                                For x = xLower To xUpper
                                    If Not ColumnFilter Is Nothing Then
                                        If ColumnFilterMask(x) Then
                                            i = i - 1
                                        ElseIf EnsureValuesColumnsMask(x) Then
                                            i = i - 1
                                        ElseIf VariantComp(Data(y, x), Data(z, x)) Then
                                            i = i - 1
                                        Else
                                            Exit For
                                        End If
                                    ElseIf EnsureValuesColumnsMask(x) Then
                                        i = i - 1
                                    ElseIf VariantComp(Data(y, x), Data(z, x)) Then
                                        i = i - 1
                                    Else
                                        Exit For
                                    End If
                                Next x

                                If i = 0 Then
                                    For i = LBound(Helper) To UBound(Helper)
                                        If VariantComp(Helper(i), Data(z, Column)) Then
                                            Ensurer(i) = Empty

                                            If IsEmpty(EnsureValuesMask(z, Column)) Then
                                                EnsureValuesMask(z, Column) = -1
                                            End If
                                        End If
                                    Next i
                                End If
                            End If
                        Next z

                        EnsureValuesRowsMask(y) = EnsureValuesRowsMask(y) + (SArrayCount(Ensurer) - SArrayCheck(Ensurer, Empty))
                        EnsureValuesMask(y, Column) = Ensurer
                    End If
                Next y
            End If
        Next KV
    End If
End Sub

Public Sub Apply(Optional HideHeaders As Boolean = False)
    Call ApplyFilterMask(HideHeaders, True)
End Sub

Public Sub ApplyFilterMask(Optional HideHeaders As Boolean = False, Optional GenerateMask As Boolean = True)
    Dim ColumnCount, Rows As Long
    Dim y, x, yF, xF, yE, xE, yG, xG As Long
    Dim Column, Aux As Variant
    Dim z As Long
    Dim Masked
    Dim RemoveRow As Boolean
    Dim MergeFiltered() As Boolean

    ReDim MergeFiltered(yLower To yUpper) As Boolean

    If GenerateMask Then
        Call GenFilterMask
    End If

    For x = xLower To xUpper
        If Not ColumnFilterMask(x) Then
            ColumnCount = ColumnCount + 1
        End If
    Next x

    For y = yLower To yUpper
        If Not RowFilterMask(y) Then
            Rows = Rows + 1
        End If
    Next y

    If UseHeaders And HideHeaders Then
        If Not RowFilterMask(1) Then
            Rows = Rows - 1
        End If
    End If

    If MergeActions.Count = 0 Then
        For y = yLower To yUpper
            If Not UseHeaders Or y > 1 Then
                If Not RowFilterMask(y) Then
                    For x = xLower To xUpper
                        If Not ColumnFilterMask(x) Then
                            z = VarType(Data(y, x))

                            If z = vbInteger Or z = vbLong Or z = vbSingle Or _
                               z = vbDouble Or z = vbCurrency Or z = vbDecimal Then
                                If UseHeaders Then
                                    MergeActions.SetValue Data(yLower, x), reMergeSum
                                End If
                            End If
                        End If
                    Next x

                    Exit For
                End If
            End If
        Next y
    End If

    For x = xLower To xUpper
        If Not ColumnFilterMask(x) Then
            For y = yLower To yUpper
                If Not RowFilterMask(y) Then
                    If Not IsEmpty(ValueFilterMask(y, x)) Then
                        If IsEmpty(Masked) Then
                            Masked = Data
                        End If

                        Masked(y, x) = ValueFilterMask(y, x)
                    End If
                End If
            Next y
        End If
    Next x

    If MergeActions.Count > 0 Then
        Aux = MergeActions.ToArray

        For x = LBound(Aux, 1) To UBound(Aux, 1)
            If UseHeaders Then
                If ColumnExists(Aux(x, 1)) Then
                    Aux(x, 1) = GetColumnIndex(Aux(x, 1))
                End If
            End If
        Next x

        For y = yLower To yUpper
            If Not RowFilterMask(y) Then
                z = MergeFilterMask(y)

                If z <> 0 And z <> y Then
                    If Not RowFilterMask(z) Then
                        RemoveRow = False

                        If IsEmpty(Masked) Then
                            Masked = Data
                        End If

                        If Not MergeFiltered(y) Then
                            For x = LBound(Aux, 1) To UBound(Aux, 1)
                                If Not IsError(Masked(z, Aux(x, 1))) And Not IsError(Masked(y, Aux(x, 1))) Then
                                    If Aux(x, 2) = reMergeSum Or IsEmpty(Aux(x, 2)) Or Aux(x, 2) = True Then
                                        Masked(z, Aux(x, 1)) = Masked(z, Aux(x, 1)) + Masked(y, Aux(x, 1))
                                        MergeFiltered(y) = True
                                        RemoveRow = True
                                    ElseIf Aux(x, 2) = reMergeSubtract Then
                                        Masked(z, Aux(x, 1)) = Masked(z, Aux(x, 1)) - Masked(y, Aux(x, 1))
                                        MergeFiltered(y) = True
                                        RemoveRow = True
                                    ElseIf Aux(x, 2) = reMergeConcat Then
                                        Masked(z, Aux(x, 1)) = Masked(z, Aux(x, 1)) & Masked(y, Aux(x, 1))
                                        MergeFiltered(y) = True
                                        RemoveRow = True
                                    ElseIf Aux(x, 2) = reMergeOverwrite Then
                                        Masked(z, Aux(x, 1)) = Masked(y, Aux(x, 1))
                                        MergeFiltered(y) = True
                                        RemoveRow = True
                                    ElseIf Aux(x, 2) = reMergeLower Then
                                        If Masked(z, Aux(x, 1)) > Masked(y, Aux(x, 1)) Then
                                            Masked(z, Aux(x, 1)) = Masked(y, Aux(x, 1))
                                        End If

                                        MergeFiltered(y) = True
                                        RemoveRow = True
                                    ElseIf Aux(x, 2) = reMergeGreater Then
                                        If Masked(z, Aux(x, 1)) < Masked(y, Aux(x, 1)) Then
                                            Masked(z, Aux(x, 1)) = Masked(y, Aux(x, 1))
                                        End If

                                        MergeFiltered(y) = True
                                        RemoveRow = True
                                    ElseIf Aux(x, 2) = reMergeIgnore Or Aux(x, 2) = False Then
                                        MergeFiltered(y) = True
                                        RemoveRow = True
                                    ElseIf Aux(x, 2) = reMergeKeep Then
                                        MergeFiltered(y) = False
                                        RemoveRow = False
                                        Exit For
                                    End If
                                End If
                            Next x
                        End If

                        If RemoveRow Then
                            Rows = Rows - 1
                        End If
                    End If
                End If
            End If
        Next y
    End If

    If EnsureValues.Count > 0 Then
        For y = yLower To yUpper
            If EnsureValuesRowsMask(y) > 0 Then
                If Not RowFilterMask(y) And Not MergeFiltered(y) Then
                    Rows = Rows + EnsureValuesRowsMask(y)
                End If
            End If
        Next y
    End If

    If Rows = 0 Or ColumnCount = 0 Then
        xFLower = 0
        xFUpper = -1
        xFSize = 0

        yFLower = 0
        yFUpper = -1
        yFSize = 0

        Filtered = SMakeArray2d(Empty)

        Exit Sub
    End If

    ReDim Filtered(1 To Rows, 1 To ColumnCount) As Variant

    xF = 0
    xG = 0

    For x = xLower To xUpper
        If Not ColumnReorder Is Nothing Then
            xE = ColumnReorderMask(x)
        Else
            xE = x
        End If

        If UseHeaders Then
            If ColumnFilterMask(xE) Then
                xF = xF + 1
                GoTo Next_X
            End If

            Aux = RenameColumnMask(xE)

            If Not IsEmpty(Aux) Then
                Filtered(1, x - xF + xG) = Aux
            Else
                If Not IsEmpty(Masked) Then
                    Filtered(1, x - xF + xG) = Masked(1, xE)
                Else
                    Filtered(1, x - xF + xG) = Data(1, xE)
                End If
            End If
        End If

        yF = 0
        yG = 0

        For y = yLower To yUpper
            If Not IsEmpty(RowReorderMask) Then
                yE = RowReorderMask(y)
            Else
                yE = y
            End If

            If yE = yLower Then
                If UseHeaders And HideHeaders Then
                    If Not RowFilterMask(yE) Then
                        yF = yF + 1
                        GoTo Next_Y
                    End If
                End If
            End If

            If RowFilterMask(yE) Then
                yF = yF + 1
                GoTo Next_Y
            End If

            If MergeFiltered(yE) Then
                yF = yF + 1
                GoTo Next_Y
            End If

            Aux = RenameColumnMask(xE)

            If UseHeaders And Not HideHeaders And yE = 1 And Not IsEmpty(Aux) Then
                Filtered(yE, x - xF + xG) = Aux
            Else
                If Not IsEmpty(Masked) Then
                    Filtered(y - yF + yG, x - xF + xG) = Masked(yE, xE)
                Else
                    Filtered(y - yF + yG, x - xF + xG) = Data(yE, xE)
                End If
            End If

            If EnsureValues.Count > 0 Then
                If EnsureValuesRowsMask(yE) > 0 Then
                    If Not IsEmpty(EnsureValuesMask(yE, xE)) And IsArray(EnsureValuesMask(yE, xE)) Then
                        For z = LBound(EnsureValuesMask(yE, xE)) To UBound(EnsureValuesMask(yE, xE))
                            If Not IsEmpty(EnsureValuesMask(yE, xE)(z)) Then
                                yG = yG + 1
                                Filtered(y - yF + yG, x - xF + xG) = EnsureValuesMask(yE, xE)(z)
                            End If
                        Next z
                    Else
                        For z = 1 To EnsureValuesRowsMask(yE)
                            yG = yG + 1

                            If Not EnsureValuesColumnsMask(xE) Then
                                If Not IsEmpty(Masked) Then
                                    Filtered(y - yF + yG, x - xF + xG) = Masked(yE, xE)
                                Else
                                    Filtered(y - yF + yG, x - xF + xG) = Data(yE, xE)
                                End If
                            End If
                        Next z
                    End If
                End If
            End If
Next_Y:
        Next y
Next_X:
    Next x

    xFLower = 1
    xFUpper = ColumnCount
    xFSize = ColumnCount

    yFLower = 1
    yFUpper = Rows
    yFSize = Rows
End Sub

Public Function GetColumn(ByRef Column As Variant, Optional UseFiltered As Boolean = True, Optional HideHeaders As Boolean = True) As Variant
    Dim Arr As Variant
    Dim Aux As Long
    Dim Code As Variant
    Dim HeaderOffset As Long

    If UseHeaders And HideHeaders Then
        HeaderOffset = 1
    End If

    If UseFiltered Then
        ReDim Arr(yFLower To yFUpper - HeaderOffset)

        If UseHeaders Then
            For Aux = xFLower To xFUpper
                If VariantComp(Column, Filtered(yFLower, Aux)) Then
                    Code = Aux
                    Exit For
                End If
            Next Aux
        Else
            Code = Column
        End If

        If IsEmpty(Code) Then
            Exit Function
        End If

        For Aux = yFLower To yFUpper - HeaderOffset
            Arr(Aux) = Filtered(Aux + HeaderOffset, Code)
        Next Aux
    Else
        ReDim Arr(yLower To yUpper - HeaderOffset)

        If UseHeaders Then
            For Aux = xLower To xUpper
                If VariantComp(Column, Data(yLower, Aux)) Then
                    Code = Aux
                    Exit For
                End If
            Next Aux
        Else
            Code = Column
        End If

        If IsEmpty(Code) Then
            Exit Function
        End If

        For Aux = yLower To yUpper - HeaderOffset
            Arr(Aux) = Data(Aux + HeaderOffset, Code)
        Next Aux
    End If

    GetColumn = Arr
End Function

Public Sub Free()
    Data = Empty
    Filtered = Empty

    Set Columns = Nothing

    Set ColumnFilter = Nothing
    Set PositiveFilter = Nothing
    Set NegativeFilter = Nothing
    Set ValueFilter = Nothing

    Set ColumnReorder = Nothing
    Set EnsureValues = Nothing

    Set MergeFilter = Nothing
    Set MergeActions = Nothing

    ColumnFilterMask = Empty
    RowFilterMask = Empty
    DataFilterMask = Empty
    ValueFilterMask = Empty

    ColumnReorderMask = Empty
    EnsureValuesRowsMask = Empty

    MergeFilterMask = Empty
End Sub
