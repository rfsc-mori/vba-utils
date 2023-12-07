Attribute VB_Name = "rCommon"
Option Explicit
Option Compare Text

' Name: rCommon
' Version: 0.78
' Depends: rArray
' Author: Rafael Fillipe Silva
' Description: ...

Public gStatusBar As Date

Public Sub rcFree(ByRef Obj As Variant)
    If IsObject(Obj) Then
        If Not Obj Is Nothing Then
            On Error Resume Next
            Obj.Free
            On Error GoTo -1
            Set Obj = Nothing
        End If
    End If
End Sub

Public Function VariantComp(ByRef String1 As Variant, ByRef String2 As Variant, Optional ByVal Compare As VbCompareMethod = vbTextCompare) As Boolean
    If VarType(String1) = vbString And VarType(String2) = vbString Then
        VariantComp = (TrimComp(String1, String2))
    ElseIf VarType(String1) = VarType(String2) Or TypeName(String1) = TypeName(String2) Then
        VariantComp = (String1 = String2)
    Else
        On Error Resume Next
        VariantComp = (String1 = String2)

        If Err.Number <> 0 Then
            VariantComp = False
            Call Err.Clear
        End If
    End If
End Function

Public Function GetRangeSeparator(Optional ByVal Fallback As String = ",") As String
    Dim Separators
    Dim Aux As Long
    Dim R As Range

    Separators = Array(";", _
                       ",", _
                       Application.International(xlRowSeparator), _
                       Application.International(xlListSeparator), _
                       Application.International(xlColumnSeparator))

    On Error Resume Next
    Set R = Nothing

    For Aux = LBound(Separators) To UBound(Separators)
        Set R = Application.Range("A1" & Separators(Aux) & "A2")

        If Not R Is Nothing Then
            GetRangeSeparator = Separators(Aux)
            Exit For
        End If
    Next Aux

    If GetRangeSeparator = "" Then GetRangeSeparator = Fallback
End Function

Public Function FixRangeSeparator(ByVal Addr As String, Optional ByVal Separator As String = ",") As String
    FixRangeSeparator = Replace(Addr, Separator, GetRangeSeparator)
End Function

Public Function SingleCriteriaTest(ByRef Value As Variant, ByVal Criteria As String) As Boolean
    Dim Pos As Long
    Dim Code As String
    Dim Invert As Boolean
    Dim DateMode As Boolean

    On Error Resume Next

    Do While True
        If Left$(Criteria, 1) = "!" Then
            Code = Left$(Criteria, 1)
            Criteria = Mid$(Criteria, 2)

            Invert = Not Invert
        ElseIf Left$(Criteria, 1) = "#" Then
            Code = Left$(Criteria, 1)
            Criteria = Mid$(Criteria, 2)

            DateMode = Not DateMode
        Else
            Exit Do
        End If
    Loop

    Pos = 1
    Code = Left$(Criteria, Pos)

    If Code = "<" Or Code = "=" Or Code = ">" Then ' < or = or >
        If DateMode Then
            Code = Mid$(Criteria, Pos, 1)

            Do While Code = "<" Or Code = "=" Or Code = ">"
                Pos = Pos + 1
                Code = Mid$(Criteria, Pos, 1)
            Loop

            Code = Left$(Criteria, Pos - 1)
            Criteria = Mid$(Criteria, Pos)

            If Code = "<" Then
                SingleCriteriaTest = CDate(Value) < CDate(Criteria)
            ElseIf Code = "=" Then
                SingleCriteriaTest = CDate(Value) = CDate(Criteria)
            ElseIf Code = ">" Then
                SingleCriteriaTest = CDate(Value) > CDate(Criteria)
            ElseIf Code = "<=" Then
                SingleCriteriaTest = CDate(Value) <= CDate(Criteria)
            ElseIf Code = ">=" Then
                SingleCriteriaTest = CDate(Value) >= CDate(Criteria)
            ElseIf Code = "<>" Then
                SingleCriteriaTest = CDate(Value) <> CDate(Criteria)
            End If
        Else
            SingleCriteriaTest = Application.Evaluate(Value & Criteria)

            If Err.Number <> 0 Then
                Err.Clear

                SingleCriteriaTest = Application.Evaluate(Quote(Value) & Criteria)

                If Err.Number <> 0 Then
                    Err.Clear

                    Code = Mid$(Criteria, Pos, 1)

                    Do While Code = "<" Or Code = "=" Or Code = ">"
                        Pos = Pos + 1
                        Code = Mid$(Criteria, Pos, 1)
                    Loop

                    SingleCriteriaTest = Application.Evaluate(Quote(Value) & Mid$(Criteria, Pos - 1, 1) & Quote(Mid$(Criteria, Pos)))
                End If
            End If
        End If
    ElseIf Code = "~" Then ' ~ (like)
        Code = Mid$(Criteria, Pos, 1)

        Do While Code = "~"
            Pos = Pos + 1
            Code = Mid$(Criteria, Pos, 1)
        Loop

        If DateMode Then
            SingleCriteriaTest = (CDate(Value) Like Mid$(Criteria, Pos))
        Else
            SingleCriteriaTest = (Value Like Mid$(Criteria, Pos))
        End If
    Else
        If DateMode Then
            SingleCriteriaTest = (CDate(Value) = Criteria)
        ElseIf VarType(Value) = vbString And VarType(Criteria) = vbString Then
            SingleCriteriaTest = (StrComp(Value, Criteria, vbTextCompare) = 0)
        Else
            SingleCriteriaTest = (Value = Criteria)
        End If
    End If

    If Invert Then
        SingleCriteriaTest = Not SingleCriteriaTest
    End If
End Function

Public Function CriteriaTest(ByRef InValues As Variant, ByRef InCriteria As Variant, Optional ValuesIndex As Variant, Optional ByVal CheckAll As Boolean = True) As Variant
    Dim Values As Variant
    Dim Criteria As Variant
    Dim IValue As Long
    Dim ICriteria As Long
    Dim Result As Variant

    Values = SMakeArray1d(InValues, 0, ValuesIndex)
    Criteria = SMakeArray1d(InCriteria, 0)

    If IsArrayInvalid(Values) Or IsEmpty(Values) Then
        Values = Array(Empty)
    End If

    If IsArrayInvalid(Criteria) Then
        Criteria = Array(Empty)
    End If

    Result = Values

    For IValue = LBound(Values) To UBound(Values)
        Result(IValue) = False

        For ICriteria = LBound(Criteria) To UBound(Criteria)
            If SingleCriteriaTest(Values(IValue), Criteria(ICriteria)) Then
                Result(IValue) = True

                If Not CheckAll Then
                    GoTo Next_Value
                End If
            ElseIf CheckAll Then
                Result(IValue) = False
                Exit For
            End If
        Next ICriteria
Next_Value:
    Next IValue

    CriteriaTest = Result
End Function

' Converte caracteres de quebra de linha em espaços
Public Function RemoveCrLf(ByVal Str As String) As String
    RemoveCrLf = Replace(Str, vbNewLine, " ")
    RemoveCrLf = Replace(RemoveCrLf, vbCr, " ")
    RemoveCrLf = Replace(RemoveCrLf, vbLf, " ")
End Function

Public Function Extract(ByRef Source As Variant, Optional What As String = "abcdefghijklmnopqrstuvwxyz0123456789") As String
    Dim Target As String
    Dim Count As Long
    Dim x As Long

    On Error GoTo ErrHandler

    Count = Len(Source)

    If Count > 0 Then
        For x = 1 To Count
            If InStr(1, What, Mid$(Source, x, 1), vbTextCompare) >= 1 Then
                Target = Target & Mid$(Source, x, 1)
            End If
        Next x
    End If

    Extract = Target
ErrHandler:
End Function

Public Function CustomTrim(ByRef Source As Variant, Optional InWhat As Variant = " ") As String
    Dim L As Long
    Dim R As Long
    Dim x As Long
    Dim y As Long
    Dim What As Variant

    CustomTrim = Source

    If Not IsArray(InWhat) Or IsEmpty(InWhat) Then
        If InWhat = "" Then
            Exit Function
        End If
    End If

    What = SMakeArray1d(InWhat)

    For y = LBound(What) To UBound(What)
        L = 1
        x = Len(What(y))

        Do While Mid$(CustomTrim, L, x) = What(y)
            L = L + x
        Loop

        CustomTrim = Mid$(CustomTrim, L)

        R = Len(CustomTrim)

        Do While InStrRev(CustomTrim, What(y), R, vbTextCompare) = R
            R = R - x
        Loop

        CustomTrim = Mid$(CustomTrim, 1, R)
    Next y
End Function

' Remove e simplifica os espaços extras e remove caracteres de quebra de linha
Public Function NormalizedString(ByRef Str As Variant, Optional InWhat As Variant = " ") As String
    Dim Pos As Variant
    Dim y As Long
    Dim What As Variant

    If IsEmpty(Str) Then
        Exit Function
    ElseIf IsMissing(Str) Then
        Exit Function
    ElseIf Str = "" Then
        Exit Function
    End If

    If Not IsArray(InWhat) Or IsEmpty(InWhat) Then
        If InWhat = "" Then
            Exit Function
        ElseIf InWhat = " " Then
            NormalizedString = Trim$(RemoveCrLf(Str))
        Else
            NormalizedString = CustomTrim(RemoveCrLf(Str), InWhat)
        End If
    Else
        NormalizedString = CustomTrim(RemoveCrLf(Str), InWhat)
    End If

    What = SMakeArray1d(InWhat)

    For y = LBound(What) To UBound(What)
        Pos = InStr(NormalizedString, What(y) & What(y))

        Do While Pos <> 0
            NormalizedString = Replace(NormalizedString, What(y) & What(y), What(y))
            Pos = InStr(NormalizedString, What(y) & What(y))
        Loop
    Next y
End Function

Public Function NormalizedComp(ByVal L As String, ByVal R As String) As Boolean
    NormalizedComp = (StrComp(NormalizedString(L), NormalizedString(R), vbTextCompare) = 0)
End Function

Public Function NormalizedMonths(ByRef Months As Variant) As Variant
    Dim Normalized As Variant
    Dim First As Date
    Dim Last As Date
    Dim Count As Long
    Dim i As Long

    Normalized = SMakeArray1d(Months)

    If IsArrayInvalid(Normalized) Then
        Exit Function
    End If

    Call Quicksort1d(Normalized)

    Last = DateTime.Now

    For i = LBound(Normalized) To UBound(Normalized)
        If Not IsError(Normalized(i)) Then
            If Abs(DateDiff("yyyy", CDate(Normalized(i)), Last)) <= 5 Then
                First = Normalized(i)
                Exit For
            End If
        End If
    Next i

    Last = Empty

    For i = UBound(Normalized) To LBound(Normalized) Step -1
        If Not IsError(Normalized(i)) Then
            If Abs(DateDiff("yyyy", First, CDate(Normalized(i)))) <= 5 Then
                Last = Normalized(i)
                Exit For
            End If
        End If
    Next i

    Count = DateDiff("m", First, Last)

    If UBound(Normalized) - LBound(Normalized) + 1 <> Count Then
        ReDim Normalized(LBound(Normalized) To LBound(Normalized) + Count) As Date

        For i = LBound(Normalized) To UBound(Normalized)
            Normalized(i) = DateAdd("m", i - LBound(Normalized), First)
        Next i
    End If

    NormalizedMonths = Normalized
End Function

Public Function SafeDate(ByRef InDate As Variant) As Date
    On Error Resume Next
    SafeDate = DateValue(InDate)
End Function

Public Function StrArrayComp(ByRef InLeft As Variant, ByRef InRight As Variant, Optional Compare As VbCompareMethod = vbTextCompare) As Long
    Dim Left As Variant
    Dim Right As Variant
    Dim x, y As Long

    StrArrayComp = 0

    If StrComp(TypeName(InLeft), "string", vbTextCompare) = 0 And StrComp(TypeName(InRight), "string", vbTextCompare) = 0 Then
        If StrComp(InLeft, InRight, Compare) = 0 Then
            StrArrayComp = 1
        End If

        Exit Function
    End If

    Left = SMakeArray1d(InLeft)
    Right = SMakeArray1d(InRight)

    For x = LBound(Left) To UBound(Left)
        For y = LBound(Right) To UBound(Right)
            If StrComp(Left(x), Right(y), Compare) = 0 Then
                StrArrayComp = StrArrayComp + 1
            End If
        Next y
    Next x
End Function

Public Function InStrArray(ByRef InLeft As Variant, ByRef InRight As Variant, Optional Compare As VbCompareMethod = vbTextCompare) As Long
    Dim Left As Variant
    Dim Right As Variant
    Dim x, y As Long

    InStrArray = 0

    If StrComp(TypeName(InLeft), "string", vbTextCompare) = 0 And StrComp(TypeName(InRight), "string", vbTextCompare) = 0 Then
        If InStr(1, InLeft, InRight, Compare) >= 1 Then
            InStrArray = 1
        End If

        Exit Function
    End If

    Left = SMakeArray1d(InLeft)
    Right = SMakeArray1d(InRight)

    For x = LBound(Left) To UBound(Left)
        For y = LBound(Right) To UBound(Right)
            If InStr(1, Left(x), Right(y), Compare) >= 1 Then
                InStrArray = InStrArray + 1
            End If
        Next y
    Next x
End Function

' Remove os espaços e compara duas strings
Public Function TrimComp(ByVal L As String, ByVal R As String) As Boolean
    TrimComp = (StrComp(Trim$(L), Trim$(R), vbTextCompare) = 0)
End Function

Public Function Quote(ByVal Str As String, Optional ByVal Char As String = """") As String
    Quote = Char & Str & Char
End Function

Public Function GetMonthByName(ByRef Mon As Variant) As Variant
    If IsEmpty(Mon) Then
        Exit Function
    End If

    On Error Resume Next
    GetMonthByName = Month(Mon & " 1")
End Function

Public Function FindLast(ByVal Str As String, ByVal What As String) As Variant
    Dim Pos As Variant
    Dim LPos As Variant

    Pos = InStr(1, Str, What, vbTextCompare)
    LPos = Pos

    If Pos <> 0 Then
        Do
            Pos = InStr(LPos + 1, Str, What, vbTextCompare)

            If Pos <> 0 Then
                LPos = Pos
            Else
                Exit Do
            End If
        Loop While Pos <> 0
    End If

    FindLast = LPos
End Function

Public Function DateReplace(ByVal Str As String, Optional ByVal Source As Date = 0, Optional ByVal Pattern As String = "yyyy-mm-dd") As String
    If Source = 0 Then
        Source = Date
    End If

    Str = Replace(Str, "$DATE", Format(Source, Pattern))
    Str = Replace(Str, "$DAY", Day(Source))
    Str = Replace(Str, "$MON", Month(Source))
    Str = Replace(Str, "$YEAR", Year(Source))

    DateReplace = Str
End Function

' Limpa o texto a barra de status
Public Function ClearStatusBar_Helper()
    Application.StatusBar = False
End Function

Public Sub ClearStatusBar(Optional ByVal When As Date)
    If gStatusBar > DateTime.Now Then
        Application.OnTime gStatusBar, "ClearStatusBar_Helper", gStatusBar + TimeValue("00:00:10"), False
        gStatusBar = Empty
    End If

    If When > DateTime.Now Then
        gStatusBar = When
        Application.OnTime gStatusBar, "ClearStatusBar_Helper", gStatusBar + TimeValue("00:00:10"), True
    Else
        gStatusBar = Empty
        ClearStatusBar_Helper
    End If
End Sub
