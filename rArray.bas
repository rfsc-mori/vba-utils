Attribute VB_Name = "rArray"
Option Explicit

' Name: rArray
' Version: 0.68
' Depends: rCallback
' Author: Rafael Fillipe Silva
' Description: ...

Public Function ArrayDimensions(ByRef Arr As Variant) As Long
    Dim Aux As Long

    If Not IsArray(Arr) Then
        Exit Function
    End If

    ArrayDimensions = 1

    On Error Resume Next

    Do While True
        Aux = LBound(Arr, ArrayDimensions)

        If Err.Number <> 0 Then
            ArrayDimensions = ArrayDimensions - 1
            Err.Clear
            Exit Do
        End If

        ArrayDimensions = ArrayDimensions + 1
    Loop
End Function

Public Function IsArrayInvalid(ByRef Arr As Variant, Optional ByVal Dimension As Long = 1) As Boolean
    On Error Resume Next
    IsArrayInvalid = (UBound(Arr, Dimension) < LBound(Arr, Dimension))
End Function

Public Function IsArray1d(ByRef Arr As Variant) As Boolean
    IsArray1d = (ArrayDimensions(Arr) = 1)
End Function

Public Function IsArray2d(ByRef Arr As Variant) As Boolean
    IsArray2d = (ArrayDimensions(Arr) = 2)
End Function

Public Function SafeIndex(ByRef Arr As Variant, ByVal index As Long, Optional ByVal Col As Long = 1) As Variant
    If Not IsArray(Arr) Or IsArrayInvalid(Arr) Then
        Exit Function
    End If

    If index < LBound(Arr) Then
        index = LBound(Arr)
    ElseIf index > UBound(Arr) Then
        index = UBound(Arr)
    End If

    On Error Resume Next

    If IsArray2d(Arr) Then
        If IsObject(Arr(index)) Then
            Set SafeIndex = Arr(index, Col)
        Else
            SafeIndex = Arr(index, Col)
        End If
    ElseIf IsArray1d(Arr) Then
        If IsObject(Arr(index)) Then
            Set SafeIndex = Arr(index)
        Else
            SafeIndex = Arr(index)
        End If
    End If
End Function

Public Sub ArrayCopy(ByRef Source As Variant, ByRef Target As Variant)
    If IsArray2d(Source) And IsArray2d(Target) Then
        ArrayCopy2d Source, Target
    Else
        ArrayCopy1d Source, Target
    End If
End Sub

Public Sub ArrayCopy1d(ByRef Source As Variant, ByRef Target As Variant, _
                       Optional ByVal SourceBase As Long = 1, Optional ByVal TargetBase As Long = 1)
    Dim T As Long
    Dim S As Long

    For T = LBound(Target, TargetBase) To UBound(Target, TargetBase)
        S = LBound(Source, SourceBase) + T - LBound(Target, TargetBase)

        If S >= LBound(Source, SourceBase) And S <= UBound(Source, SourceBase) Then
            Target(T) = Source(S)
        Else
            Exit For
        End If
    Next Aux
End Sub

Public Sub ArrayCopy2d(ByRef Source As Variant, ByRef Target As Variant)
    Dim y, x As Long
    Dim ys, xs As Long

    For y = LBound(Target, 1) To UBound(Target, 1)
        ys = LBound(Source, 1) + y - LBound(Target, 1)

        If ys >= LBound(Source) And ys <= UBound(Source) Then
            For x = LBound(Target, 2) To UBound(Target, 2)
                xs = LBound(Source, 2) + x - LBound(Target, 2)

                If xs >= LBound(Source, 2) And xs <= UBound(Source, 2) Then
                    Target(y, x) = Source(ys, xs)
                Else
                    Exit For
                End If
            Next x
        Else
            Exit For
        End If
    Next Aux
End Sub

Public Sub Swap(ByRef Arr As Variant, ByVal Source As Long, ByVal Target As Long)
    Dim Aux As Variant

    If IsObject(Arr(Source)) Then
        Set Aux = Arr(Source)
    Else
        Aux = Arr(Source)
    End If

    If IsObject(Arr(Target)) Then
        Set Arr(Source) = Arr(Target)
    Else
        Arr(Source) = Arr(Target)
    End If

    If IsObject(Aux) Then
        Set Arr(Target) = Aux
    Else
        Arr(Target) = Aux
    End If
End Sub

Public Sub Swap2d(ByRef Arr As Variant, ByVal Source As Long, ByVal Target As Long, _
                  Optional ByVal SColumn As Variant, Optional ByVal TColumn As Variant)
    Dim Aux As Variant

    If IsMissing(SColumn) Then
        SColumn = LBound(Arr, 2)
    End If

    If IsMissing(TColumn) Then
        TColumn = SColumn
    End If

    If IsObject(Arr(Source, SColumn)) Then
        Set Aux = Arr(Source, SColumn)
    Else
        Aux = Arr(Source, SColumn)
    End If

    If IsObject(Arr(Target, TColumn)) Then
        Set Arr(Source, SColumn) = Arr(Target, TColumn)
    Else
        Arr(Source, SColumn) = Arr(Target, TColumn)
    End If

    If IsObject(Aux) Then
        Set Arr(Target, TColumn) = Aux
    Else
        Arr(Target, TColumn) = Aux
    End If
End Sub

Public Function SArrayCheck(ByRef Arr As Variant, Optional ByRef Value As Variant) As Long
    Dim i As Long
    Dim j As Long

    If Not IsEmpty(Arr) Then
        If IsArray1d(Arr) Then
            If IsMissing(Value) Then
                Value = Arr(LBound(Arr))
            End If

            If UBound(Arr) < LBound(Arr) Then
                Exit Function
            End If

            j = 0

            For i = LBound(Arr) To UBound(Arr)
                If Arr(i) = Value Then
                    j = j + 1
                End If
            Next i

            SArrayCheck = j
        End If
    End If
End Function

' Quantidade de itens em uma array
Public Function SArrayCount(ByRef Arr As Variant) As Long
    Dim Aux As Variant
    Dim Count As Variant
    Dim index As Variant

    SArrayCount = 0
    On Error GoTo ErrHandler

    If Not IsEmpty(Arr) Then
        If IsObject(Arr) Then
            For Each Aux In Arr
                SArrayCount = SArrayCount + 1
            Next Aux
        Else
            If IsArray(Arr) Then
                Aux = ArrayDimensions(Arr)

                For index = 1 To Aux
                    Count = UBound(Arr, Aux) - LBound(Arr, Aux) + 1

                    If Count > SArrayCount Then
                        SArrayCount = Count
                    End If
                Next index
            Else
                SArrayCount = 1
            End If
        End If
    End If

ErrHandler:
End Function

Public Function ResizeArray1d(ByRef Arr As Variant, Optional ByVal First As Variant, _
                              Optional ByVal Last As Variant, Optional ByVal Base As Variant) As Variant
    Dim Aux As Long
    Dim NewArr As Variant

    If IsMissing(First) Then
        First = LBound(Arr)
    End If

    If IsMissing(Last) Then
        Last = UBound(Arr)
    End If

    If IsMissing(Base) Then
        Base = LBound(Arr)
    End If

    ReDim NewArr(Base To Base + Last - First) As Variant

    For Aux = First To Last
        NewArr(Base + Aux - First) = Arr(Aux)
    Next Aux

    ResizeArray1d = NewArr
End Function

Public Function ResizeArray2d(ByRef Arr As Variant, _
                              Optional ByVal LFirst As Variant, Optional ByVal LLast As Variant, _
                              Optional ByVal RFirst As Variant, Optional ByVal RLast As Variant, _
                              Optional ByVal LBase As Variant, Optional ByVal RBase As Variant, _
                              Optional Placeholder As Variant = Empty) As Variant
    Dim y As Long
    Dim x As Long
    Dim Aux As Long
    Dim NewArr As Variant

    If IsMissing(LFirst) Then
        LFirst = LBound(Arr, 1)
    End If

    If IsMissing(LLast) Then
        LLast = UBound(Arr, 1)
    End If

    If IsMissing(RFirst) Then
        RFirst = LBound(Arr, 2)
    End If

    If IsMissing(RLast) Then
        RLast = UBound(Arr, 2)
    End If

    If IsMissing(LBase) Then
        LBase = LFirst
    End If

    If IsMissing(RBase) Then
        RBase = RFirst
    End If

    ReDim NewArr(LBase To LBase + LLast - LFirst, RBase To RBase + RLast - RFirst) As Variant

    For y = LFirst To LLast
        Aux = (LBase + y - LFirst)

        For x = RFirst To RLast
            If y >= LBound(Arr, 1) And y <= UBound(Arr, 1) And x >= LBound(Arr, 2) And x <= UBound(Arr, 2) Then
                NewArr(Aux, RBase + x - RFirst) = Arr(y, x)
            Else
                NewArr(Aux, RBase + x - RFirst) = Placeholder
            End If
        Next x
    Next y

    ResizeArray2d = NewArr
End Function

' Concatena duas colunas em uma
Public Function ConcatToColumn2d(ByRef InArr As Variant, ByRef InColumns As Variant, ByRef Separator As Variant, _
                                 Optional Target As Variant, Optional First As Variant, Optional Last As Variant, _
                                 Optional Placeholder As Variant = Empty, Optional Remove As Boolean = True) As Variant
    Dim Arr As Variant
    Dim Cols As Variant
    Dim Aux As Variant
    Dim RCols As Variant
    Dim i As Long
    Dim x As Long

    Arr = InArr
    Cols = SMakeArray1d(InColumns)

    If IsMissing(Target) Then
        Target = Empty
    ElseIf Target < LBound(Arr, 2) Or Target > UBound(Arr, 2) Then
        Target = Empty
    End If

    If IsEmpty(Target) Then
        ReDim Preserve Arr(LBound(Arr, 1) To UBound(Arr, 1), LBound(Arr, 2) To UBound(Arr, 2) + 1)
        Target = UBound(Arr, 2)
    End If

    If IsMissing(First) Then
        First = LBound(Arr, 1)
    End If

    If IsMissing(Last) Then
        Last = UBound(Arr, 1)
    End If

    For i = LBound(Arr, 1) To UBound(Arr, 1)
        If i < First Then
            Arr(i, Target) = Placeholder
        ElseIf i > Last Then
            Arr(i, Target) = Placeholder
        Else
            Aux = Arr(i, Target)
            Arr(i, Target) = Empty

            For x = LBound(Cols, 1) To UBound(Cols, 1)
                If Not IsEmpty(Arr(i, Target)) Then
                    If Cols(x) <> Target Then
                        Arr(i, Target) = Arr(i, Target) & Separator & Arr(i, Cols(x))
                    Else
                        Arr(i, Target) = Arr(i, Target) & Separator & Aux
                    End If
                Else
                    If Cols(x) <> Target Then
                        Arr(i, Target) = Arr(i, Cols(x))
                    Else
                        Arr(i, Target) = Aux
                    End If
                End If
            Next x
        End If
    Next i

    If Remove Then
        ReDim RCols(LBound(Cols) To UBound(Cols)) As Variant

        i = LBound(Cols) - 1

        For x = LBound(Cols) To UBound(Cols)
            If Cols(x) <> Target Then
                If i < LBound(Cols) Or i > UBound(Cols) Then
                    i = LBound(Cols)
                End If

                RCols(i) = Cols(x)
                i = i + 1
            End If
        Next x

        If i > LBound(Cols) - 1 Then
            ReDim Preserve RCols(LBound(Cols) To LBound(Cols) + i - 1) As Variant

            ConcatToColumn2d = RemoveColumns2d(Arr, RCols)
        Else
            ConcatToColumn2d = Arr
        End If
    Else
        ConcatToColumn2d = Arr
    End If
End Function

' Ordena uma array de uma dimensão
Public Sub Quicksort1d(ByRef Arr As Variant, Optional ByVal lower As Variant, _
                        Optional ByVal upper As Variant, Optional ByRef Order As Variant = True, _
                        Optional ByRef Callback As Variant, Optional ByRef CallbackObj As Variant, _
                        Optional ByRef Mask As Variant)
    Call Quicksort2d(Arr, , lower, upper, Order, Callback, CallbackObj, Mask)
End Sub

' Ordena uma coluna de uma array de duas dimensões
Public Sub Quicksort2d(ByRef Arr As Variant, Optional ByVal Col As Variant, Optional ByVal lower As Variant, _
                        Optional ByVal upper As Variant, Optional ByRef Order As Variant = True, _
                        Optional ByRef Callback As Variant, Optional ByRef CallbackObj As Variant, _
                        Optional ByRef Mask As Variant)
    Dim L As Long
    Dim R As Long
    Dim p As Long
    Dim y As Long
    Dim o As Long
    Dim x As Variant
    Dim c As Boolean
    Dim ArrIs2d As Boolean
    Dim MaskIs2d As Boolean
    Dim lstack As rcSetCollection
    Dim rstack As rcSetCollection

    Set lstack = New rcSetCollection
    Set rstack = New rcSetCollection

    If IsMissing(lower) Then
        lower = LBound(Arr, 1)
    End If

    If IsMissing(upper) Then
        upper = UBound(Arr, 1)
    End If

    ArrIs2d = IsArray2d(Arr)
    MaskIs2d = IsArray2d(Mask)

    If IsMissing(Col) Then
        x = Empty
    ElseIf ArrIs2d Then
        If Col < LBound(Arr, 2) Or Col > UBound(Arr, 2) Then
            x = Empty
        Else
            x = Col
        End If
    Else
        x = Empty
    End If

    Call lstack.Add(New rcKeyValue, True).SetKV(lower, upper)

    Do While True
        If lstack.Count > 0 Then
            With lstack(lstack.Last)
                L = .Key
                R = .Value
            End With

            Call lstack.List.Remove(lstack.Last)
        ElseIf rstack.Count > 0 Then
            With rstack(rstack.Last)
                L = .Key
                R = .Value
            End With

            Call rstack.List.Remove(rstack.Last)
        Else
            Exit Do
        End If

        If (R - L + 1) >= 2 Then
            p = L + ((R - L + 1) / 2)

            If IsMissing(Mask) Then
                If ArrIs2d Then
                    Call Swap2d(Arr, p, L, x)
                Else
                    Call Swap(Arr, p, L)
                End If
            Else
                If IsArray2d(Mask) Then
                    Call Swap2d(Mask, p, L, x)
                Else
                    Call Swap(Mask, p, L)
                End If
            End If

            p = L
            o = L - 1
            L = L + 1

            For y = L To R
                If o < p Or o > R Then
                    o = y
                End If

                c = False

                If IsMissing(Callback) Then
                    If IsMissing(Mask) Then
                        If ArrIs2d Then
                            If Order = True Then
                                If Not IsError(Arr(y, x)) And Not IsError(Arr(p, x)) Then
                                    c = Arr(y, x) < Arr(p, x)
                                Else
                                    c = IsError(Arr(p, x))
                                End If
                            Else
                                If Not IsError(Arr(y, x)) And Not IsError(Arr(p, x)) Then
                                    c = Arr(y, x) > Arr(p, x)
                                Else
                                    c = IsError(Arr(y, x))
                                End If
                            End If
                        Else
                            If Order = True Then
                                If Not IsError(Arr(y)) And Not IsError(Arr(p)) Then
                                    c = Arr(y) < Arr(p)
                                Else
                                    c = IsError(Arr(p))
                                End If
                            Else
                                If Not IsError(Arr(y)) And Not IsError(Arr(p)) Then
                                    c = Arr(y) > Arr(p)
                                Else
                                    c = IsError(Arr(y))
                                End If
                            End If
                        End If
                    Else
                        If MaskIs2d Then
                            If ArrIs2d Then
                                If Order = True Then
                                    If Not IsError(Arr(Mask(y, x), x)) And Not IsError(Arr(Mask(p, x), x)) Then
                                        c = Arr(Mask(y, x), x) < Arr(Mask(p, x), x)
                                    Else
                                        c = IsError(Arr(Mask(p, x), x))
                                    End If
                                Else
                                    If Not IsError(Arr(Mask(y, x), x)) And Not IsError(Arr(Mask(p, x), x)) Then
                                        c = Arr(Mask(y, x), x) > Arr(Mask(p, x), x)
                                    Else
                                        c = IsError(Arr(Mask(y, x), x))
                                    End If
                                End If
                            Else
                                If Order = True Then
                                    If Not IsError(Arr(Mask(y, x))) And Not IsError(Arr(Mask(p, x))) Then
                                        c = Arr(Mask(y, x)) < Arr(Mask(p, x))
                                    Else
                                        c = IsError(Arr(Mask(p, x)))
                                    End If
                                Else
                                    If Not IsError(Arr(Mask(y, x))) And Not IsError(Arr(Mask(p, x))) Then
                                        c = Arr(Mask(y, x)) > Arr(Mask(p, x))
                                    Else
                                        c = IsError(Arr(Mask(y, x)))
                                    End If
                                End If
                            End If
                        Else
                            If ArrIs2d Then
                                If Order = True Then
                                    If Not IsError(Arr(Mask(y), x)) And Not IsError(Arr(Mask(p), x)) Then
                                        c = Arr(Mask(y), x) < Arr(Mask(p), x)
                                    Else
                                        c = IsError(Arr(Mask(p), x))
                                    End If
                                Else
                                    If Not IsError(Arr(Mask(y), x)) And Not IsError(Arr(Mask(p), x)) Then
                                        c = Arr(Mask(y), x) > Arr(Mask(p), x)
                                    Else
                                        c = IsError(Arr(Mask(y), x))
                                    End If
                                End If
                            Else
                                If Order = True Then
                                    If Not IsError(Arr(Mask(y))) And Not IsError(Arr(Mask(p))) Then
                                        c = Arr(Mask(y)) < Arr(Mask(p))
                                    Else
                                        c = IsError(Arr(Mask(p)))
                                    End If
                                Else
                                    If Not IsError(Arr(Mask(y))) And Not IsError(Arr(Mask(p))) Then
                                        c = Arr(Mask(y)) > Arr(Mask(p))
                                    Else
                                        c = IsError(Arr(Mask(y)))
                                    End If
                                End If
                            End If
                        End If
                    End If
                Else
                    If IsMissing(Mask) Then
                        If ArrIs2d Then
                            c = RunCallback(CallbackObj, Callback, Array(Array(Arr(y, x), Arr(p, x), Order)), False, VbMethod)
                        Else
                            c = RunCallback(CallbackObj, Callback, Array(Array(Arr(y), Arr(p), Order)), False, VbMethod)
                        End If
                    Else
                        If MaskIs2d Then
                            If ArrIs2d Then
                                c = RunCallback(CallbackObj, Callback, Array(Array(Arr(Mask(y, x), x), Arr(Mask(p, x), x), Order)), False, VbMethod)
                            Else
                                c = RunCallback(CallbackObj, Callback, Array(Array(Arr(Mask(y, x)), Arr(Mask(p, x)), Order)), False, VbMethod)
                            End If
                        Else
                            If ArrIs2d Then
                                c = RunCallback(CallbackObj, Callback, Array(Array(Arr(Mask(y), x), Arr(Mask(p), x), Order)), False, VbMethod)
                            Else
                                c = RunCallback(CallbackObj, Callback, Array(Array(Arr(Mask(y)), Arr(Mask(p)), Order)), False, VbMethod)
                            End If
                        End If
                    End If
                End If

                If c Then
                    If o <> y Then
                        If IsMissing(Mask) Then
                            If IsArray2d(Arr) Then
                                Call Swap2d(Arr, y, o, x)
                            Else
                                Call Swap(Arr, y, o)
                            End If
                        Else
                            If IsArray2d(Mask) Then
                                Call Swap2d(Mask, y, o, x)
                            Else
                                Call Swap(Mask, y, o)
                            End If
                        End If
                    End If

                    o = o + 1
                End If
            Next y

            If o > lower Then
                o = o - 1
            End If

            If o <> p Then
                If IsMissing(Mask) Then
                    If IsArray2d(Arr) Then
                        Call Swap2d(Arr, p, o, x)
                    Else
                        Call Swap(Arr, p, o)
                    End If
                Else
                    If IsArray2d(Mask) Then
                        Call Swap2d(Mask, p, o, x)
                    Else
                        Call Swap(Mask, p, o)
                    End If
                End If
            End If

            If p < o Then
                Call lstack.Add(New rcKeyValue, True).SetKV(p, o - 1)
            End If

            If o < R Then
                Call rstack.Add(New rcKeyValue, True).SetKV(o + 1, R)
            End If
        End If
    Loop
End Sub

' Remove linhas
Public Function SRemoveRows(ByRef InArr As Variant, ByRef InRows As Variant) As Variant
    Dim Arr As Variant
    Dim Rows As Variant
    Dim y, V, x As Long
    Dim R, S As Long
    Dim Last As Boolean

    Arr = InArr
    Rows = SMakeArray1d(InRows)

    If IsEmpty(Rows) Then
        SRemoveRows = Arr
        Exit Function
    End If

    R = 0

    Do
        Last = False

        For y = LBound(Rows) To UBound(Rows)
            If Rows(y) <> LBound(Arr, 1) - 1 Then
                If Rows(y) = UBound(Arr, 1) - R Then
                    R = R + 1
                    Last = True

                    Rows(y) = LBound(Arr, 1) - 1

                    Exit For
                End If
            End If
        Next y
    Loop While Last = True

    Call Quicksort1d(Rows)

    S = 0
    Last = False

    If R < UBound(Rows) - LBound(Rows) + 1 Then
        For y = UBound(Rows) To LBound(Rows) Step -1
            If Rows(y) <> LBound(Arr, 1) - 1 Then
                If Rows(y) < UBound(Arr, 1) - R Then
                    For V = Rows(y) To UBound(Arr, 1) - R - 1
                        For x = LBound(Arr, 2) To UBound(Arr, 2)
                            Arr(V, x) = Arr(V + 1, x)
                        Next x
                    Next V
                End If
            End If

            S = S + 1
        Next y
    End If

    If UBound(Arr, 1) - LBound(Arr, 1) + 1 - R <= 0 Then
        SRemoveRows = Empty
    Else
        SRemoveRows = ResizeArray2d(Arr, LBound(Arr, 1), UBound(Arr, 1) - R - S, LBound(Arr, 2), UBound(Arr, 2))
    End If
End Function

' Remove colunas
Public Function SRemoveColumns(ByRef InArr As Variant, ByRef InColumns As Variant) As Variant
    Dim Arr As Variant
    Dim Cols As Variant
    Dim x, h, y As Long
    Dim R, S As Long
    Dim Last As Boolean

    Arr = InArr
    Cols = SMakeArray1d(InColumns)

    If IsEmpty(Cols) Then
        SRemoveColumns = Arr
        Exit Function
    End If

    R = 0

    Do
        Last = False

        For x = LBound(Cols) To UBound(Cols)
            If Cols(x) <> LBound(Arr, 2) - 1 Then
                If Cols(x) = UBound(Arr, 2) - R Then
                    R = R + 1
                    Last = True

                    Cols(x) = LBound(Arr, 2) - 1

                    Exit For
                End If
            End If
        Next x
    Loop While Last = True

    Call Quicksort1d(Cols)

    S = 0
    Last = False

    If R < UBound(Cols) - LBound(Cols) + 1 Then
        For x = UBound(Cols) To LBound(Cols) Step -1
            If Cols(x) <> LBound(Arr, 2) - 1 Then
                If Cols(x) < UBound(Arr, 2) - R Then
                    For h = Cols(x) To UBound(Arr, 2) - R - 1
                        For y = LBound(Arr, 1) To UBound(Arr, 1)
                            Arr(y, h) = Arr(y, h + 1)
                        Next y
                    Next h
                End If
            End If

            S = S + 1
        Next x
    End If

    If UBound(Arr, 2) - LBound(Arr, 2) + 1 - R <= 0 Then
        SRemoveColumns = Empty
    Else
        ReDim Preserve Arr(LBound(Arr, 1) To UBound(Arr, 1), LBound(Arr, 2) To UBound(Arr, 2) - R - S)
        SRemoveColumns = Arr
    End If
End Function

' Combina duas arrays em uma
Public Function CombineArray1d(ByRef Left As Variant, ByRef Delimiter As Variant, ByRef Right As Variant) As Variant
    Dim LB As Variant
    Dim LBL As Variant
    Dim LBR As Variant

    Dim UB As Variant
    Dim UBL As Variant
    Dim UBR As Variant

    Dim Aux As Variant

    LBL = LBound(Left)
    LBR = LBound(Right)
    UBL = UBound(Left)
    UBR = UBound(Right)

    If LBL < LBR Then
        LB = LBL
    Else
        LB = LBR
    End If

    If UBL < UBR Then
        UB = UBL
    Else
        UB = UBR
    End If

    ReDim CombineArray1d(LB To UB)

    For Aux = LB To UB
        If Aux >= LBL And Aux <= UBL Then
            CombineArray1d(Aux) = Left(Aux)
        End If

        If Aux >= LBR And Aux <= UBR Then
            If Not IsEmpty(CombineArray1d(Aux)) Then
                CombineArray1d(Aux) = CombineArray1d(Aux) & Delimiter & Right(Aux)
            Else
                CombineArray1d(Aux) = Right(Aux)
            End If
        End If
    Next Aux
End Function

Public Function ArrayMatch2d(ByRef Arr As Variant, ByRef InColumns As Variant, ByRef InValues As Variant, _
                             Optional ByRef InFirst As Variant, Optional ByRef InLast As Variant) As Variant
    Dim Rows() As Boolean
    Dim Cols As Variant
    Dim Values As Variant
    Dim Col As Variant
    Dim Row As Variant
    Dim Found As Boolean
    Dim First As Long
    Dim Last As Long

    Cols = SMakeArray1d(InColumns, 1)
    Values = SMakeArray1d(InValues, 1)

    If IsEmpty(Cols) Then
        Exit Function
    End If

    If UBound(Cols) <> UBound(Values) Then
        Exit Function
    End If

    ReDim Rows(LBound(Arr) To UBound(Arr)) As Boolean

    If IsMissing(InFirst) Then
        First = LBound(Rows)
    ElseIf InFirst = 0 Then
        First = LBound(Rows)
    Else
        First = InFirst
    End If

    If IsMissing(InLast) Then
        Last = UBound(Rows)
    ElseIf InLast = 0 Then
        Last = UBound(Rows)
    Else
        Last = InLast
    End If

    For Row = First To Last
        Rows(Row) = True
    Next Row

    For Col = LBound(Cols) To UBound(Cols)
        Found = False

        For Row = First To Last
            If Rows(Row) Then
                If Not IsEmpty(Cols(Col)) And Not IsError(Cols(Col)) Then
                    If Not IsError(Arr(Row, Cols(Col))) Then
                        If Arr(Row, Cols(Col)) <> Values(Col) Then
                            Rows(Row) = False
                        Else
                            If Not Found Then
                                Found = True

                                If Row > First Then
                                    First = Row
                                End If
                            Else
                                Last = Row
                            End If
                        End If
                    End If
                End If
            End If
        Next Row

        If Not Found Then
            ArrayMatch2d = Empty
            Exit Function
        End If
    Next Col

    ArrayMatch2d = Rows

    If Not IsMissing(InFirst) Then
        InFirst = First
    End If

    If Not IsMissing(InLast) Then
        InLast = Last
    End If
End Function

Public Function ArraySelfMatch2d(ByRef Arr As Variant, ByRef InColumns As Variant, _
                                 Optional ByVal InFirst As Variant, Optional ByVal InLast As Variant) As Variant
    Dim Matches() As Long
    Dim Cols As Variant
    Dim Values() As Variant
    Dim Col As Variant
    Dim Row As Variant
    Dim First As Long
    Dim Last As Long
    Dim Helper As Variant
    Dim Aux As Long

    Cols = SMakeArray1d(InColumns, 1)

    If IsEmpty(Cols) Then
        Exit Function
    End If

    ReDim Matches(LBound(Arr) To UBound(Arr)) As Long
    ReDim Values(LBound(Cols) To UBound(Cols)) As Variant

    If IsMissing(InFirst) Or InFirst = 0 Then
        First = LBound(Matches)
    Else
        First = InFirst
    End If

    If IsMissing(InLast) Or InLast = 0 Then
        Last = UBound(Matches)
    Else
        Last = InLast
    End If

    For Row = First To Last
        If Matches(Row) = 0 Then
            For Col = LBound(Cols) To UBound(Cols)
                If Not IsEmpty(Cols(Col)) Then
                    Values(Col) = Arr(Row, Cols(Col))
                End If
            Next Col

            Helper = ArrayMatch2d(Arr, Cols, Values, Row, Last + 0)

            If Not IsEmpty(Helper) Then
                For Aux = LBound(Helper) To UBound(Helper)
                    If Helper(Aux) Then
                        Matches(Aux) = Row
                    End If
                Next Aux
            End If
        End If
    Next Row

    ArraySelfMatch2d = Matches
End Function

' Copia uma array ou variável de forma inteligente
Public Function SMakeArray1d(ByRef Arr As Variant, Optional ByRef Base As Variant, Optional ByVal ArrIndex As Variant = 1) As Variant
    Dim Helper As Variant
    Dim Count As Long
    Dim Aux As Variant

    Count = 0
    On Error GoTo ErrHandler

    If Not IsEmpty(Arr) Then
        If IsObject(Arr) Then
            For Each Aux In Arr
                Count = Count + 1
            Next Aux

            If IsMissing(Base) Then
                Base = 0
            End If

            ReDim Helper(Base To Base + Count - 1)
            Count = Base

            For Each Aux In Arr
                Helper(Count) = Aux
                Count = Count + 1
            Next Aux
        Else
            If IsArray(Arr) Then
                Count = UBound(Arr, ArrIndex) - LBound(Arr, ArrIndex) + 1

                If IsMissing(Base) Then
                    If IsArray1d(Arr) Then
                        ReDim Helper(LBound(Arr) To UBound(Arr))

                        For Aux = LBound(Arr) To UBound(Arr)
                            Helper(Aux) = Arr(Aux)
                        Next Aux
                    Else
                        ReDim Helper(LBound(Arr, ArrIndex) To UBound(Arr, ArrIndex))

                        For Aux = LBound(Arr, ArrIndex) To UBound(Arr, ArrIndex)
                            If ArrIndex = 1 Then
                                Helper(Aux) = Arr(Aux, 1)
                            Else
                                Helper(Aux) = Arr(1, Aux)
                            End If
                        Next Aux
                    End If
                Else
                    ReDim Helper(Base To Base + Count - 1)

                    If IsArray1d(Arr) Then
                        For Aux = LBound(Arr) To UBound(Arr)
                            Helper(Aux - LBound(Arr) + Base) = Arr(Aux)
                        Next Aux
                    Else
                        For Aux = LBound(Arr, ArrIndex) To UBound(Arr, ArrIndex)
                            If ArrIndex = 1 Then
                                Helper(Aux - LBound(Arr, ArrIndex) + Base) = Arr(Aux, 1)
                            Else
                                Helper(Aux - LBound(Arr, ArrIndex) + Base) = Arr(1, Aux)
                            End If
                        Next Aux
                    End If
                End If
            Else
                If IsMissing(Base) Then
                    ReDim Helper(0 To 0)
                    Base = 0
                Else
                    ReDim Helper(Base To Base)
                End If

                Helper(Base) = Arr
            End If
        End If
    Else
        Helper = Array()
    End If

    SMakeArray1d = Helper

ErrHandler:
End Function

Public Function SMakeArray2d(ByRef Arr As Variant, Optional ByRef Base As Variant, _
                             Optional ByVal ArrIndex As Variant = 1, Optional ByVal ArrCol As Variant = 1) As Variant
    Dim Helper As Variant
    Dim Count As Long
    Dim Aux As Variant

    Count = 0
    On Error GoTo ErrHandler

    If Not IsEmpty(Arr) Then
        If IsObject(Arr) Then
            For Each Aux In Arr
                Count = Count + 1
            Next Aux

            If IsMissing(Base) Then
                Base = 0
            End If

            ReDim Helper(Base To Base + Count - 1, Base To Base)
            Count = Base

            For Each Aux In Arr
                Helper(Count, Base) = Aux
                Count = Count + 1
            Next Aux
        Else
            If IsArray(Arr) Then
                Count = UBound(Arr, ArrIndex) - LBound(Arr, ArrIndex) + 1

                If IsMissing(Base) Then
                    If IsArray1d(Arr) Then
                        ReDim Helper(LBound(Arr) To UBound(Arr), Base To Base)

                        For Aux = LBound(Arr) To UBound(Arr)
                            Helper(Aux, Base) = Arr(Aux)
                        Next Aux
                    Else
                        ReDim Helper(LBound(Arr, ArrIndex) To UBound(Arr, ArrIndex), Base To Base)

                        For Aux = LBound(Arr, ArrIndex) To UBound(Arr, ArrIndex)
                            If ArrIndex = 1 Then
                                Helper(Aux, Base) = Arr(Aux, ArrCol)
                            Else
                                Helper(Aux, Base) = Arr(ArrCol, Aux)
                            End If
                        Next Aux
                    End If
                Else
                    ReDim Helper(Base To Base + Count - 1, Base To Base)

                    If IsArray1d(Arr) Then
                        For Aux = LBound(Arr) To UBound(Arr)
                            Helper(Aux - LBound(Arr) + Base, Base) = Arr(Aux)
                        Next Aux
                    Else
                        For Aux = LBound(Arr, ArrIndex) To UBound(Arr, ArrIndex)
                            If ArrIndex = 1 Then
                                Helper(Aux - LBound(Arr, ArrIndex) + Base, Base) = Arr(Aux, ArrCol)
                            Else
                                Helper(Aux - LBound(Arr, ArrIndex) + Base, Base) = Arr(ArrCol, Aux)
                            End If
                        Next Aux
                    End If
                End If
            Else
                If IsMissing(Base) Then
                    ReDim Helper(0 To 0, 0 To 0)
                    Base = 0
                Else
                    ReDim Helper(Base To Base, Base To Base)
                End If

                Helper(Base, Base) = Arr
            End If
        End If
    Else
        Helper = Array()
    End If

    SMakeArray2d = Helper

ErrHandler:
End Function
