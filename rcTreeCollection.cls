VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "rcTreeCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Name: rcTreeCollection
' Version: 0.31
' Depends: rCommon,rcKeyValue,rcKeyValueCollection,rcSetCollection
' Author: Rafael Fillipe Silva
' Description: ...

Public KeyTree
Public KeyCount As Long
Public ValueList
Public ValueCount As Long
Public Count As Variant
Public MergeAction As reMergeAction
Public Root As rcKeyValueCollection

Public Sub Setup(Optional Keys = Empty, Optional Values = Empty, Optional ByVal Action As reMergeAction = reMergeKeep)
    Dim Counter As Long
    Dim Aux As Long

    KeyCount = SArrayCount(Keys)
    ValueCount = SArrayCount(Values)

    If KeyCount <= 0 Or ValueCount <= 0 Then
        Exit Sub
    End If

    Set Root = New rcKeyValueCollection

    KeyTree = SMakeArray1d(Keys)
    ValueList = SMakeArray1d(Values)

    ReDim Count(LBound(KeyTree) To UBound(KeyTree)) As Long

    MergeAction = Action
End Sub

Private Function Tree(ByRef Keys As Variant, Optional Add As Boolean = False) As rcKeyValueCollection
    Dim KeyList
    Dim Counter As Long
    Dim Aux As Long

    Set Tree = Nothing

    Counter = SArrayCount(Keys)

    If Counter <> KeyCount Or Counter <= 0 Then
        Exit Function
    End If

    KeyList = SMakeArray1d(Keys)

    Set Tree = Root

    For Aux = LBound(KeyList) To UBound(KeyList)
        If Not Tree.Exists(KeyList(Aux)) Or (Add And Aux = UBound(KeyList) - 1) Then
            Set Tree = Tree.Add(KeyList(Aux), New rcKeyValueCollection, True).Value
            Count(Aux) = Count(Aux) + 1
        Else
            Set Tree = Tree.Value(KeyList(Aux))
        End If
    Next Aux

    If Tree.List.Count = 0 Then
        If ValueCount > 0 Then
            For Aux = LBound(ValueList) To UBound(ValueList)
                If Not Tree.Exists(ValueList(Aux)) Then
                    Tree.Add ValueList(Aux), Empty
                End If
            Next Aux
        End If
    End If
End Function

Public Function GetKeys(ByRef Key As Variant) As rcSetCollection
    Dim Tree As rcSetCollection
    Dim Branch As rcKeyValue
    Dim Branches As rcSetCollection
    Dim Child As rcKeyValue
    Dim Aux As Long
    Dim Column As Variant

    Column = Empty

    Set GetKeys = New rcSetCollection

    For Aux = LBound(KeyTree) To UBound(KeyTree)
        If VariantComp(KeyTree(Aux), Key) Then
            Column = Aux
            Exit For
        End If
    Next Aux

    If IsEmpty(Column) Then
        For Aux = LBound(ValueList) To UBound(ValueList)
            If VariantComp(ValueList(Aux), Key) Then
                Column = KeyCount + Aux
                Exit For
            End If
        Next Aux
    End If

    If IsEmpty(Column) Then
        Exit Function
    End If

    Set Branches = New rcSetCollection
    Set Tree = New rcSetCollection

    Aux = LBound(KeyTree)

    For Each Branch In Root
        Tree.Add Branch, True
    Next Branch

Continue:
    For Each Branch In Tree
        Branches.Add Branch, True
    Next Branch

    Tree.Clear
    Aux = Aux + 1

    If Aux = Column Then
        For Each Branch In Branches
            If TypeOf Branch.Value Is rcKeyValueCollection Then
                For Each Child In Branch.Value
                    GetKeys.Add Child.Key
                Next Child
            End If
        Next Branch

        Exit Function
    End If

    For Each Branch In Branches
        If TypeOf Branch.Value Is rcKeyValueCollection Then
            For Each Child In Branch.Value
                If TypeOf Child Is rcKeyValue Then
                    Tree.Add Child, True
                End If
            Next Child
        End If
    Next Branch

    Branches.Clear

    If Tree.Count > 0 Then
        GoTo Continue
    End If
End Function

Public Sub EnsureKeys(ByRef InValues As Variant, ByRef Key As Variant)
    Dim Tree As rcSetCollection
    Dim Branch As rcKeyValue
    Dim Branches As rcSetCollection
    Dim Child As rcKeyValue
    Dim Aux As Long
    Dim Column As Variant
    Dim Values As Variant
    Dim Value As Variant

    Column = Empty

    For Aux = LBound(KeyTree) To UBound(KeyTree)
        If VariantComp(KeyTree(Aux), Key) Then
            Column = Aux
            Exit For
        End If
    Next Aux

    If IsEmpty(Column) Then
        For Aux = LBound(ValueList) To UBound(ValueList)
            If VariantComp(ValueList(Aux), Key) Then
                Column = KeyCount + Aux
                Exit For
            End If
        Next Aux
    End If

    If IsEmpty(Column) Then
        Exit Sub
    End If

    Set Branches = New rcSetCollection
    Set Tree = New rcSetCollection

    Aux = LBound(KeyTree)

    For Each Branch In Root
        Tree.Add Branch, True
    Next Branch

Continue:
    For Each Branch In Tree
        Branches.Add Branch, True
    Next Branch

    Tree.Clear
    Aux = Aux + 1

    If Aux = Column Then
        For Each Branch In Branches
            If TypeOf Branch.Value Is rcKeyValueCollection Then
                Values = SMakeArray1d(InValues)

                For Each Value In Values
                    If Not Branch.Value.Exists(Value) Then
                        Branch.Value.Add Value, Empty
                        Count(Aux) = Count(Aux) + 1
                    End If
                Next Value
            End If
        Next Branch

        Exit Sub
    End If

    For Each Branch In Branches
        If TypeOf Branch.Value Is rcKeyValueCollection Then
            For Each Child In Branch.Value
                If TypeOf Child Is rcKeyValue Then
                    Tree.Add Child, True
                End If
            Next Child
        End If
    Next Branch

    Branches.Clear

    If Tree.Count > 0 Then
        GoTo Continue
    End If
End Sub

Private Sub SetValues(ByRef Branch As rcKeyValueCollection, ByRef Values As Variant)
    Dim Vals
    Dim Counter As Long
    Dim Aux As Long
    Dim Val As Variant

    Counter = SArrayCount(Values)

    If Counter <> ValueCount Or Counter <= 0 Then
        Exit Sub
    End If

    Vals = SMakeArray1d(Values)
    Counter = LBound(ValueList)

    For Aux = LBound(Vals) To UBound(Vals)
        Val = ValueList(Aux - LBound(Vals) + Counter)

        If Branch.Value(Val) = "" Or MergeAction = reMergeOverwrite Then
            Branch.KeyValue(Val).Value = Vals(Aux)
        Else
            If MergeAction = reMergeSum Then
                Branch.KeyValue(Val).Value = Branch.Value(Val) + Vals(Aux)
            ElseIf MergeAction = reMergeSubtract Then
                Branch.KeyValue(Val).Value = Branch.Value(Val) - Vals(Aux)
            ElseIf MergeAction = reMergeConcat Then
                Branch.KeyValue(Val).Value = Branch.Value(Val) & Vals(Aux)
            End If
        End If
    Next Aux
End Sub

Public Function Add(Keys As Variant, Values As Variant) As Boolean
    Dim Branch As rcKeyValueCollection
    Dim AddNew As Boolean

    Add = False
    AddNew = (MergeAction = reMergeKeep)

    Set Branch = Tree(Keys, AddNew)

    If Branch Is Nothing Then
        Exit Function
    End If

    SetValues Branch, Values

    Add = True
End Function

Public Function TotalRows(Optional Headers As Boolean = True) As Long
    Dim Aux As Long

    For Aux = LBound(Count) To UBound(Count)
        If TotalRows < Count(Aux) Then
            TotalRows = Count(Aux)
        End If
    Next Aux

    If Headers Then
        TotalRows = TotalRows + 1
    End If
End Function

Public Function TotalColumns() As Long
    TotalColumns = KeyCount + ValueCount
End Function

Public Function Remove(Keys As Variant) As Boolean
    Dim KeyList
    Dim Counter As Long
    Dim Aux As Long
    Dim Key As Variant
    Dim TreeList() As rcKeyValueCollection

    Set Tree = Nothing

    Counter = ArrayCount(Keys)

    If Counter <> KeyCount Or Counter <= 0 Then
        Exit Function
    End If

    KeyList = MakeArray(Keys, 0)

    ReDim TreeList(LBound(KeyList) To UBound(KeyList))

    Set Tree = Root

    For Aux = LBound(KeyList) To UBound(KeyList)
        If Tree.Exists(KeyList(Aux)) Then
            Set Tree = Tree.Value(KeyList(Aux))
            Set TreeList(Aux) = Tree
        End If
    Next Aux

    Tree.Free

    For Aux = UBound(TreeList) To LBound(TreeList)
        If TreeList(Aux).List Is Nothing Then
            TreeList(Aux).Free
            Count(Aux) = Count(Aux) - 1
        ElseIf TreeList(Aux).List.Count = 0 Then
            TreeList(Aux).Free
            Count(Aux) = Count(Aux) - 1
        Else
            Counter = 1

            For Each Key In TreeList(Aux).List
                If Not Key Is Nothing Then
                    Counter = 0
                    Exit For
                End If
            Next Key

            If Counter = 1 Then
                TreeList(Aux).Free
                Count(Aux) = Count(Aux) - 1
            End If
        End If
    Next Aux
End Function

Public Sub Normalize(Optional Sort As Boolean = True)
    If Root Is Nothing Then
        Exit Sub
    End If

    Normalize_Helper Root, LBound(KeyTree)

    If Sort Then
        Root.SortTree
    End If
End Sub

Private Sub Normalize_Helper(ByRef Root As rcKeyValueCollection, ByVal Level As Long)
    Dim Tree As rcSetCollection
    Dim Branch As rcKeyValue
    Dim Branches As rcSetCollection
    Dim Child As rcKeyValue
    Dim Children As rcKeyValueCollection
    Dim Key As rcKeyValue
    Dim Keys As rcKeyValueCollection
    Dim Aux As Long

    Set Branches = New rcSetCollection
    Set Keys = New rcKeyValueCollection

    Set Tree = New rcSetCollection

    Aux = Level

    For Each Branch In Root
        Tree.Add Branch, True
    Next Branch

Continue:
    For Each Branch In Tree
        Branches.Add Branch, True
    Next Branch

    Tree.Clear
    Aux = Aux + 1

    For Each Branch In Branches
        If TypeOf Branch.Value Is rcKeyValueCollection Then
            For Each Child In Branch.Value
                If Not Keys.Exists(Child.Key) Then
                    Keys.Add Child.Key, Child.Value
                End If
            Next Child
        End If
    Next Branch

    For Each Branch In Branches
        If TypeOf Branch.Value Is rcKeyValueCollection Then
            Set Children = Branch.Value

            For Each Key In Keys
                If Not Children.Exists(Key.Key) Then
                    If TypeOf Key.Value Is rcKeyValueCollection Then
                        Children.Add Key.Key, New rcKeyValueCollection
                        Count(Aux) = Count(Aux) + 1
                    Else
                        Children.Add Key.Key, Empty
                    End If
                End If
            Next Key

            For Each Child In Children
                If TypeOf Child Is rcKeyValue Then
                    Tree.Add Child, True
                End If
            Next Child
        End If
    Next Branch

    Branches.Clear
    Keys.Clear

    If Tree.Count > 0 Then
        GoTo Continue
    End If
End Sub

Public Function ToArray(Optional Headers As Boolean = True, _
                        Optional ByRef PFilters As rcKeyValueCollection, _
                        Optional ByRef NFilters As rcKeyValueCollection) As Variant
    Dim Arr As Variant
    Dim Aux As Long
    Dim KV As rcKeyValue
    Dim Code As Variant
    Dim PTest As Variant
    Dim NTest As Variant
    Dim RowIndex As Long

    If Not PFilters Is Nothing Then
        ReDim PTest(LBound(KeyTree) To LBound(KeyTree) + TotalColumns)

        For Each KV In PFilters
            Code = Empty

            For Aux = LBound(KeyTree) To UBound(KeyTree)
                If KV.Key = KeyTree(Aux) Then
                    Code = Aux
                    Exit For
                End If
            Next Aux

            If IsEmpty(Code) Then
                For Aux = LBound(ValueList) To UBound(ValueList)
                    If KV.Key = ValueList(Aux) Then
                        Code = UBound(KeyTree) + 1 + Aux - LBound(ValueList)
                        Exit For
                    End If
                Next Aux
            End If

            If Not IsEmpty(Code) Then
                PTest(Code) = KV.Value
            Else
                Exit Function
            End If
        Next KV
    End If

    If Not NFilters Is Nothing Then
        ReDim NTest(LBound(KeyTree) To LBound(KeyTree) + TotalColumns)

        For Each KV In NFilters
            Code = Empty

            For Aux = LBound(KeyTree) To UBound(KeyTree)
                If KV.Key = KeyTree(Aux) Then
                    Code = Aux
                    Exit For
                End If
            Next Aux

            If IsEmpty(Code) Then
                For Aux = LBound(ValueList) To UBound(ValueList)
                    If KV.Key = ValueList(Aux) Then
                        Code = UBound(KeyTree) + 1 + Aux - LBound(ValueList)
                        Exit For
                    End If
                Next Aux
            End If

            If Not IsEmpty(Code) Then
                NTest(Code) = KV.Value
            Else
                Exit Function
            End If
        Next KV
    End If

    If Headers Then
        ReDim Arr(1 To TotalRows, 1 To TotalColumns)

        For Aux = LBound(KeyTree) To UBound(KeyTree)
            Arr(1, LBound(Arr, 2) + Aux - LBound(KeyTree)) = KeyTree(Aux)
        Next Aux

        For Aux = LBound(ValueList) To UBound(ValueList)
            Arr(1, LBound(Arr, 2) + KeyCount + Aux - LBound(ValueList)) = ValueList(Aux)
        Next Aux

        RowIndex = 2
        ToArray_Helper Arr, Root, RowIndex, 1, PTest, NTest
    Else
        ReDim Arr(1 To TotalRows, 1 To TotalColumns)

        RowIndex = 1
        ToArray_Helper Arr, Root, RowIndex, 1, PTest, NTest
    End If

    ResizeArray2d Arr, LLast:=RowIndex, RLast:=TotalColumns

    ToArray = Arr
End Function

Private Sub ToArray_Helper(ByRef Arr As Variant, ByRef Tree As rcKeyValueCollection, _
                           ByRef RowIndex As Long, ByVal ColumnIndex As Long, _
                           ByRef PFilters As Variant, ByRef NFilters As Variant)
    Dim Branch As rcKeyValue
    Dim AddLine As Boolean
    Dim Filtered As Boolean
    Dim Aux As Long
    Dim Column As Variant

    AddLine = False

    For Each Branch In Tree.List
        Filtered = False

        If ColumnIndex <= KeyCount Then
            If Not IsEmpty(PFilters) Then
                If Not IsEmpty(PFilters(ColumnIndex - 1)) Then
                    If Not SingleCriteriaTest(Branch.Key, PFilters(ColumnIndex - 1)) Then
                        Filtered = True
                    End If
                End If
            End If

            If Not IsEmpty(NFilters) Then
                If Not IsEmpty(NFilters(ColumnIndex - 1)) Then
                    If SingleCriteriaTest(Branch.Key, NFilters(ColumnIndex - 1)) Then
                        Filtered = True
                    End If
                End If
            End If

            If Not Filtered Then
                If AddLine Then
                    RowIndex = RowIndex + 1

                    If ColumnIndex > LBound(Arr, 2) Then
                        For Aux = LBound(Arr, 2) To ColumnIndex - 1
                            Arr(RowIndex, Aux) = Arr(RowIndex - 1, Aux)
                        Next Aux
                    End If
                End If

                Arr(RowIndex, ColumnIndex) = Branch.Key

                If Not IsEmpty(Branch.Value) Then
                    ToArray_Helper Arr, Branch.Value, RowIndex, ColumnIndex + 1, PFilters, NFilters
                End If
            End If
        Else
            Column = Empty

            For Aux = LBound(ValueList) To UBound(ValueList)
                If VariantComp(ValueList(Aux), Branch.Key) Then
                    Column = KeyCount + Aux + 1
                    Exit For
                End If
            Next Aux

            If Not IsEmpty(Column) Then
                If Not IsEmpty(PFilters) Then
                    If Not IsEmpty(PFilters(Column)) Then
                        If Not SingleCriteriaTest(Branch.Key, PFilters(Column)) Then
                            Filtered = True
                        End If
                    End If
                End If

                If Not IsEmpty(NFilters) Then
                    If Not IsEmpty(NFilters(Column)) Then
                        If SingleCriteriaTest(Branch.Key, NFilters(Column)) Then
                            Filtered = True
                        End If
                    End If
                End If

                If Not Filtered Then
                    Arr(RowIndex, Column) = Branch.Value
                End If
            End If

            ColumnIndex = ColumnIndex + 1
        End If

        If Not Filtered Then
            AddLine = True
        End If
    Next Branch
End Sub

Public Sub Free()
    rcFree Root
End Sub
