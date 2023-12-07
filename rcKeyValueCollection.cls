VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "rcKeyValueCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Name: rcKeyValueCollection
' Version: 0.66
' Depends: rArray,rCommon,rcKeyValue,rcSetCollection
' Author: Rafael Fillipe Silva
' Description: ...

Public List As Collection
Public TextMode As Boolean
Public NormalizeMode As Boolean

Public Enum reMergeAction
    reMergeKeep = 1
    reMergeSum = 2
    reMergeConcat = 3
    reMergeOverwrite = 4
    reMergeIgnore = 5
    reMergeSubtract = 6
    reMergeLower = 7
    reMergeGreater = 8
    reMergeCustom = 9
End Enum

' Initializer
Private Sub Class_Initialize()
    Set List = New Collection
    TextMode = True
    NormalizeMode = True
End Sub

Public Function Add(ByRef InKey As Variant, ByRef Value As Variant, Optional ByVal AllowDuplicates As Boolean = False) As rcKeyValue
    Dim Key As Variant

    If TextMode And NormalizeMode Then
        Key = NormalizedString(InKey)
    Else
        Key = InKey
    End If

    If Not Exists(Key) Or AllowDuplicates Then
        Set Add = New rcKeyValue
        Call Add.SetKV(Key, Value)
        Call List.Add(Add)
    End If
End Function

Public Function AddPairs(ByRef InKeys As Variant, ByRef InValues As Variant, Optional ByVal AllowDuplicates As Boolean = False) As rcSetCollection
    Dim Keys, Values As Variant
    Dim Value As Variant
    Dim Aux As Long

    Set AddPairs = New rcSetCollection

    Keys = SMakeArray1d(InKeys)
    Values = SMakeArray1d(InValues)

    For Aux = LBound(Keys) To UBound(Keys)
        If Aux > UBound(Values) Then
            Value = Values(UBound(Values))
        Else
            Value = Values(Aux)
        End If

        Call AddPairs.Add(Add(Keys(Aux), Value, AllowDuplicates))
    Next Aux
End Function

Public Sub Clear()
    Set List = New Collection
End Sub

Public Function SetValue(ByRef InKey As Variant, Optional ByRef Value As Variant, Optional ByVal CanAdd As Boolean = True) As rcKeyValue
    Dim Key As Variant

    If TextMode And NormalizeMode Then
        Key = NormalizedString(InKey)
    Else
        Key = InKey
    End If

    If IsMissing(Value) Then
        Call Remove(Key)
        Exit Function
    End If

    Set SetValue = KeyValue(Key)

    If Not SetValue Is Nothing Then
        If IsObject(Value) Then
            Set SetValue.Value = Value
        Else
            SetValue.Value = Value
        End If
    ElseIf CanAdd Then
        Set SetValue = Add(Key, Value, False)
    End If
End Function

Public Function SetPairs(ByRef InKeys As Variant, Optional ByRef InValues As Variant, Optional ByVal CanAdd As Boolean = True) As rcSetCollection
    Dim Keys, Values As Variant
    Dim Value As Variant
    Dim Aux As Long

    If IsMissing(InValues) Then
        Call Remove(InKeys)
        Exit Function
    End If

    Set SetPairs = New rcSetCollection

    Keys = SMakeArray1d(InKeys)
    Values = SMakeArray1d(InValues)

    For Aux = LBound(Keys) To UBound(Keys)
        If Aux > UBound(Values) Then
            Value = Values(UBound(Values))
        Else
            Value = Values(Aux)
        End If

        Call SetPairs.Add(SetValue(Keys(Aux), Value, CanAdd))
    Next Aux
End Function

Public Function Count(Optional ByRef InKey As Variant) As Long
    Dim KV As rcKeyValue
    Dim Key As Variant

    If TextMode And NormalizeMode Then
        Key = NormalizedString(InKey)
    Else
        Key = InKey
    End If

    If Not IsMissing(InKey) Then
        For Each KV In List
            If TextMode Then
                If StrComp(TypeName(KV.Key), "String", vbTextCompare) = 0 And _
                   StrComp(TypeName(Key), "String", vbTextCompare) = 0 Then
                    If TrimComp(KV.Key, Key) Then
                        Count = Count + 1
                    End If
                ElseIf VariantComp(KV.Key, Key) Then
                    Count = Count + 1
                End If
            ElseIf VariantComp(KV.Key, Key) Then
                Count = Count + 1
            End If

            If VariantComp(KV.Key, Key) Then
                Count = Count + 1
            End If
        Next KV
    Else
        Count = List.Count
    End If
End Function

Public Function Exists(ByRef InKey As Variant) As Boolean
    Dim KV As rcKeyValue
    Dim Key As Variant

    If TextMode And NormalizeMode Then
        Key = NormalizedString(InKey)
    Else
        Key = InKey
    End If

    On Error Resume Next

    For Each KV In List
        If TextMode Then
            If StrComp(TypeName(KV.Key), "String", vbTextCompare) = 0 And _
               StrComp(TypeName(Key), "String", vbTextCompare) = 0 Then
                If TrimComp(KV.Key, Key) Then
                    Exists = True
                    Exit For
                End If
            ElseIf VariantComp(KV.Key, Key) Then
                Exists = True
                Exit For
            End If
        ElseIf VariantComp(KV.Key, Key) Then
            Exists = True
            Exit For
        End If
    Next KV
End Function

Public Function Keys() As rcSetCollection
    Dim KV As rcKeyValue

    Set Keys = New rcSetCollection

    For Each KV In List
        Call Keys.Add(KV.Key, True)
    Next KV
End Function

Public Function Item(ByRef Index As Variant) As Variant
    If IsObject(List.Item(Index)) Then
        Set Item = List.Item(Index)
    Else
        Item = List.Item(Index)
    End If
End Function

Public Function UniqueKeys() As rcSetCollection
    Dim KV As rcKeyValue

    Set UniqueKeys = New rcSetCollection

    For Each KV In List
        Call UniqueKeys.Add(KV.Key)
    Next KV
End Function

Public Function Values(Optional ByRef InKey As Variant) As rcSetCollection
    Dim KV As rcKeyValue
    Dim Key As Variant


    If TextMode And NormalizeMode And Not IsMissing(InKey) Then
        Key = NormalizedString(InKey)
    Else
        Key = InKey
    End If

    Set Values = New rcSetCollection

    If Not IsMissing(Key) Then
        For Each KV In List
            If TextMode Then
                If StrComp(TypeName(KV.Key), "String", vbTextCompare) = 0 And _
                   StrComp(TypeName(Key), "String", vbTextCompare) = 0 Then
                    If TrimComp(KV.Key, Key) Then
                        Call Values.Add(KV.Value, True)
                    End If
                ElseIf VariantComp(KV.Key, Key) Then
                    Call Values.Add(KV.Value, True)
                End If
            ElseIf VariantComp(KV.Key, Key) Then
                Call Values.Add(KV.Value, True)
            End If
        Next KV
    Else
        For Each KV In List
            Call Values.Add(KV.Value, True)
        Next KV
    End If
End Function

Public Function ValuesLike(ByRef InKeys As Variant) As rcSetCollection
    Dim KV As rcKeyValue
    Dim Keys As Variant
    Dim i As Long

    Set ValuesLike = New rcSetCollection

    Keys = SMakeArray1d(InKeys)

    If IsArrayInvalid(Keys) Then
        Exit Function
    End If

    For Each KV In List
        For i = LBound(Keys) To UBound(Keys)
            If SingleCriteriaTest(KV.Key, Keys(i)) Then
                Call ValuesLike.Add(KV.Value, True)
            End If
        Next i
    Next KV
End Function

Public Function UniqueValues() As rcSetCollection
    Dim KV As rcKeyValue

    Set Values = New rcSetCollection

    For Each KV In List
        Call Values.Add(KV.Value)
    Next KV
End Function

Public Function KeyValue(ByRef InKey As Variant) As rcKeyValue
    Dim KV As rcKeyValue
    Dim Key As Variant

    If TextMode And NormalizeMode Then
        Key = NormalizedString(InKey)
    Else
        Key = InKey
    End If

    Set KeyValue = Nothing

    For Each KV In List
        If TextMode Then
            If StrComp(TypeName(KV.Key), "String", vbTextCompare) = 0 And _
               StrComp(TypeName(Key), "String", vbTextCompare) = 0 Then
                If TrimComp(KV.Key, Key) Then
                    Set KeyValue = KV
                    Exit For
                End If
            ElseIf VariantComp(KV.Key, Key) Then
                Set KeyValue = KV
                Exit For
            End If
        ElseIf VariantComp(KV.Key, Key) Then
            Set KeyValue = KV
            Exit For
        End If
    Next KV
End Function

Public Function KeyValuesLike(ByRef InKeys As Variant) As rcSetCollection
    Dim KV As rcKeyValue
    Dim Keys As Variant
    Dim i As Long

    Set KeyValuesLike = New rcSetCollection

    Keys = SMakeArray1d(InKeys)

    If IsArrayInvalid(Keys) Then
        Exit Function
    End If

    For Each KV In List
        For i = LBound(Keys) To UBound(Keys)
            If SingleCriteriaTest(KV.Key, Keys(i)) Then
                Call KeyValuesLike.Add(KV, True)
            End If
        Next i
    Next KV
End Function

Public Function Value(ByRef InKey As Variant, Optional ByRef Default As Variant) As Variant
Attribute Value.VB_UserMemId = 0
    Dim KV As rcKeyValue
    Dim Key As Variant

    If TextMode And NormalizeMode Then
        Key = NormalizedString(InKey)
    Else
        Key = InKey
    End If

    Set Value = Nothing

    For Each KV In List
        If TextMode Then
            If StrComp(TypeName(KV.Key), "String", vbTextCompare) = 0 And _
               StrComp(TypeName(Key), "String", vbTextCompare) = 0 Then
                If TrimComp(KV.Key, Key) Then
                    If IsObject(KV.Value) Then
                        Set Value = KV.Value
                    Else
                        Value = KV.Value
                    End If

                    Exit Function
                End If
            ElseIf VariantComp(KV.Key, Key) Then
                If IsObject(KV.Value) Then
                    Set Value = KV.Value
                Else
                    Value = KV.Value
                End If

                Exit Function
            End If
        ElseIf VariantComp(KV.Key, Key) Then
            If IsObject(KV.Value) Then
                Set Value = KV.Value
            Else
                Value = KV.Value
            End If

            Exit Function
        End If
    Next KV

    If Not IsMissing(Default) Then
        If IsObject(Default) Then
            Set Value = Default
        Else
            Value = Default
        End If
    End If
End Function

Public Function SortHelper(ByRef Param As Variant) As Boolean
    Dim Left As rcKeyValue
    Dim Right As rcKeyValue
    Dim Order As Variant

    Set Left = Param(0)
    Set Right = Param(1)

    Order = Param(2)

    If Order = True Then
        SortHelper = Left.Key < Right.Key
    Else
        SortHelper = Left.Key > Right.Key
    End If
End Function

' Ordena o dicionário
Public Function Sort() As rcKeyValueCollection
    Dim Arr() As rcKeyValue
    Dim Aux As Variant
    Dim x As Variant
    Dim y As Variant

    Set Sort = Me

    If List.Count = 0 Then
        Exit Function
    End If

    ReDim Arr(0 To List.Count - 1)

    For x = 0 To List.Count - 1
        Set Arr(x) = List(x + 1)
    Next x

    Call Quicksort1d(Arr, Callback:="SortHelper", CallbackObj:=Me)

    Set List = New Collection

    For Each Aux In Arr
        Call List.Add(Aux)
    Next Aux
End Function

' Ordena todos os dicionários (incluindo os filhos) caso seja uma árvores de dicionários
Public Function SortTree() As rcKeyValueCollection
    Dim Cursor As rcKeyValueCollection
    Dim KV As rcKeyValue

    Set Cursor = Me
    Set SortTree = Me

    Call Cursor.Sort

    For Each KV In Cursor.List
        If TypeName(KV.Value) = TypeName(Me) Then
            Call KV.Value.SortTree
        End If
    Next KV
End Function

' Mescla dois dicionários
Public Function Merge(ByRef KVC As rcKeyValueCollection) As rcKeyValueCollection
    Dim KV As rcKeyValue

    For Each KV In KVC.List
        Add KV.Key, KV.Value, True
    Next KV
End Function

' Consolida os dicionários filhos em novos objetos ao invés de referências
' NOTA: Previne loop infinito em MergeTree
Public Function ConsolidateTree() As rcKeyValueCollection
    Dim i As Variant
    Dim KV As rcKeyValue
    Dim KC As rcKeyValue
    Dim KVC As rcKeyValueCollection

    For i = 1 To List.Count
        Set KV = List(i)

        If TypeName(KV.Value) = TypeName(Me) Then
            Set KVC = New rcKeyValueCollection

            For Each KC In KV.Value.List
                Call KVC.Add(KC.Key, KC.Value, True)
            Next KC

            Call KVC.ConsolidateTree

            Set KV.Value = KVC
        End If
    Next i
End Function

' Mescla os dicionários filhos duplicados
Public Function MergeTree(Optional ByVal Action As reMergeAction = reMergeKeep, Optional ByVal MaxDepth As Variant, _
                          Optional ByRef Depth As Long) As rcKeyValueCollection
    Dim KV As rcKeyValue
    Dim Col As Collection
    Dim K As rcKeyValue
    Dim i As Variant
    Dim x As Variant
    Dim y As Variant
    Dim z As Variant
    Dim Removed As Boolean

    For i = 1 To List.Count
        Set KV = List(i)

Restart:
        If TypeName(KV.Value) = TypeName(Me) Then
            If Not IsMissing(MaxDepth) Then
                If Depth >= MaxDepth Then
                    Exit Function
                End If
            End If

            Depth = Depth + 1

            Call KV.Value.MergeTree(Action, MaxDepth, Depth)

            Depth = Depth - 1
        End If

        If Count(KV.Key) > 1 Then
            Set Col = List

            For x = i To Col.Count
                If Col(x).Key = KV.Key Then
                    For y = Col.Count To x + 1 Step -1
                        If Col(y).Key = KV.Key Then
                            Removed = False

                            If TypeName(Col(x).Value) = TypeName(Me) Then
                                For z = 1 To Col(y).Value.List.Count
                                    Set K = Col(y).Value.List(z)
                                    Col(x).Value.Add K.Key, K.Value, True
                                Next z

                                Call Col.Remove(y)
                                Removed = True
                            Else
                                If Action = reMergeKeep Then
                                    Removed = False
                                ElseIf Action = reMergeSum Then
                                    Col(x).Value = Col(x).Value + Col(y).Value
                                    Removed = True
                                ElseIf Action = reMergeConcat Then
                                    Col(x).Value = Col(x).Value & Col(y).Value
                                    Removed = True
                                ElseIf Action = reMergeOverwrite Then
                                    Col(x).Value = Col(y).Value
                                    Removed = True
                                ElseIf Action = reMergeLower Then
                                    If Col(x).Value > Col(y).Value Then
                                        Col(x).Value = Col(y).Value
                                        Removed = True
                                    End If
                                ElseIf Action = reMergeGreater Then
                                    If Col(x).Value < Col(y).Value Then
                                        Col(x).Value = Col(y).Value
                                        Removed = True
                                    End If
                                ElseIf Action = reMergeIgnore Then
                                    Removed = True
                                End If

                                If Removed Then
                                    Call Col.Remove(y)
                                End If
                            End If
                        End If
                    Next y

                    If Removed Then
                        GoTo Restart
                    End If
                End If
            Next x
        End If

        If i >= List.Count Then
            Exit For
        End If
    Next i
End Function

Public Function ToArray() As Variant
    Dim Arr As Variant
    Dim Aux As Long
    Dim KV As rcKeyValue

    If List.Count = 0 Then
        ToArray = Empty
        Exit Function
    End If

    ReDim Arr(1 To List.Count, 1 To 2)
    Aux = LBound(Arr, 1)

    For Each KV In List
        Arr(Aux, 1) = KV.Key
        Arr(Aux, 2) = KV.Value

        Aux = Aux + 1
    Next KV

    ToArray = Arr
End Function

Public Function Remove(ByRef InKeys As Variant) As Long
    Dim i As Long
    Dim Aux As Long
    Dim Keys As Variant

    Keys = SMakeArray1d(InKeys)

    For Aux = LBound(Keys) To UBound(Keys)
        For i = 1 To List.Count
            If TextMode Then
                If StrComp(TypeName(List(i).Key), "String", vbTextCompare) = 0 And _
                   StrComp(TypeName(Keys(Aux)), "String", vbTextCompare) = 0 Then
                    If TrimComp(List(i).Key, Keys(Aux)) Then
                        List.Remove i
                        Remove = Remove + 1
                        Exit For
                    End If
                ElseIf VariantComp(List(i).Key, Keys(Aux)) Then
                    List.Remove i
                    Remove = Remove + 1
                    Exit For
                End If
            ElseIf VariantComp(List(i).Key, Keys(Aux)) Then
                List.Remove i
                Remove = Remove + 1
                Exit For
            End If
        Next i
    Next Aux
End Function

Public Sub Free()
    Set List = Nothing
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = List.[_NewEnum]
End Property
