VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "rcSetCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Name: rcSetCollection
' Version: 0.41
' Depends: rArray,rCommon
' Author: Rafael Fillipe Silva
' Description: ...

Public List As Collection

' Initializer
Private Sub Class_Initialize()
    Set List = New Collection
End Sub

Public Function Add(ByRef Value As Variant, Optional ByVal AllowDuplicates As Boolean = False) As Variant
    Dim CanAdd As Boolean

    If IsObject(Value) Then
        Set Add = Nothing
    Else
        Add = 0
    End If

    CanAdd = AllowDuplicates

    If Not CanAdd Then
        CanAdd = Not Exists(Value)
    End If

    If CanAdd Then
        Call List.Add(Value)

        If IsObject(Value) Then
            Set Add = Value
        Else
            Add = Value
        End If
    End If
End Function

Public Function Exists(ByRef Value As Variant) As Boolean
    Dim Test As Variant

    On Error Resume Next

    For Each Test In List
        If VariantComp(Test, Value) Then
            Exists = True
            Exit For
        End If
    Next Test
End Function

Public Sub Clear()
    Set List = New Collection
End Sub

Public Function Count() As Long
    If Not List Is Nothing Then
        Count = List.Count
    Else
        Count = 0
    End If
End Function

Public Function Item(ByVal index As Long) As Variant
Attribute Item.VB_UserMemId = 0
    Item = Empty

    If index > 0 And index <= List.Count Then
        If IsObject(List(index)) Then
            Set Item = List(index)
        Else
            Item = List(index)
        End If
    End If
End Function

Public Function TakeItem(ByVal index As Long) As Variant
    TakeItem = Empty

    If index > 0 And index <= List.Count Then
        If IsObject(List(index)) Then
            Set TakeItem = List(index)
        Else
            TakeItem = List(index)
        End If

        Call List.Remove(index)
    End If
End Function

Public Function SetValue(ByVal index As Long, ByRef Value As Variant, Optional ByVal AllowDuplicates As Boolean = False) As Boolean
    Dim CanAdd As Boolean

    CanAdd = AllowDuplicates

    If Not CanAdd Then
        CanAdd = Not Exists(Value)
    End If

    If CanAdd Then
        If index > 0 And index <= List.Count Then
            Call List.Add(Value, Before:=index)
            Call List.Remove(index + 1)
            SetValue = True
        End If
    End If
End Function

Public Function AddFrom(ByRef Values As Variant) As rcSetCollection
    Dim V As Variant

    If IsObject(Values) Then
        For Each V In Values
            Call Add(V)
        Next V
    ElseIf IsArray1d(Values) Then
        For V = LBound(Values) To UBound(Values)
            Call Add(Values(V))
        Next V
    ElseIf IsArray2d(Values) Then
        For V = LBound(Values) To UBound(Values)
            Call Add(Values(V, 1))
        Next V
    Else
        Call Add(Values)
    End If

    Set AddFrom = Me
End Function

Public Function Unique() As rcSetCollection
    Dim Aux As Variant
    Set Unique = New rcSetCollection

    For Each Aux In List
        Call Unique.Add(Aux)
    Next Aux
End Function

Public Function ItemsLike(ByRef InValues As Variant) As rcSetCollection
    Dim Aux As Variant
    Dim Values As Variant
    Dim i As Long

    Set ItemsLike = New rcSetCollection

    Values = SMakeArray1d(InValues)

    If IsArrayInvalid(Values) Then
        Exit Function
    End If

    For Each Aux In List
        For i = LBound(Values) To UBound(Values)
            If SingleCriteriaTest(Aux, Values(i)) Then
                Call ItemsLike.Add(Aux, True)
            End If
        Next i
    Next Aux
End Function

Public Function ToArray(Optional ByVal Vertical As Boolean = True) As Variant
    Dim Arr As Variant
    Dim Aux As Variant
    Dim Counter As Variant

    If List.Count = 0 Then
        ToArray = Empty
        Exit Function
    End If

    If Vertical Then
        ReDim Arr(1 To List.Count)
        Counter = LBound(Arr, 1)

        For Each Aux In List
            Arr(Counter) = Aux
            Counter = Counter + 1
        Next Aux
    Else
        ReDim Arr(1 To 1, 1 To List.Count)
        Counter = LBound(Arr, 2)

        For Each Aux In List
            Arr(1, Counter) = Aux
            Counter = Counter + 1
        Next Aux
    End If

    ToArray = Arr
End Function

Public Function Sort() As rcSetCollection
    Dim Arr() As Variant
    Dim Aux As Variant
    Dim x As Variant
    Dim y As Variant

    Set Sort = Me

    If List.Count = 0 Then
        Exit Function
    End If

    ReDim Arr(0 To List.Count - 1)

    For x = 0 To List.Count - 1
        If IsObject(List(x + 1)) Then
            Set Arr(x) = List(x + 1)
        Else
            Arr(x) = List(x + 1)
        End If
    Next x

    Call Quicksort1d(Arr)

    Set List = New Collection

    For Each Aux In Arr
        Call List.Add(Aux)
    Next Aux
End Function

Public Function Remove(ByRef Value As Variant) As Boolean
    Dim i As Variant

    Remove = False

    For i = 1 To List.Count
        If VariantComp(List(i), Value) Then
            Call List.Remove(i)
            Remove = True
            Exit For
        End If
    Next i
End Function

Public Function First() As Long
    First = 1
End Function

Public Function Last() As Long
    Last = List.Count
End Function

Public Function Invalid() As Long
    Invalid = 0
End Function

Public Sub Free()
    Set List = Nothing
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = List.[_NewEnum]
End Property
