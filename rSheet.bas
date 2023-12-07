Attribute VB_Name = "rSheet"
Option Explicit

' Name: rSheet
' Version: 0.76
' Depends: rArray,rCommon,rcKeyValueCollection,rcProgressHelper
' Author: Rafael Fillipe Silva
' Description: ...

' Remove todos os filtros
Public Sub RemoveFilters(ByRef Sheet As Worksheet)
    With Sheet
        If .AutoFilterMode = True Then
            .Cells.AutoFilter
        End If
    End With
End Sub

' Exibe todas as colunas ocultas
Public Sub UnhideColumns(ByRef Sheet As Worksheet)
    Sheet.Cells.EntireColumn.Hidden = False
End Sub

' Exibe todas as linhas ocultas
Public Sub UnhideRows(ByRef Sheet As Worksheet)
    Sheet.Cells.EntireRow.Hidden = False
End Sub

' Exibe todas as células ocultas
Public Sub UnhideCells(ByRef Sheet As Worksheet)
    Call UnhideColumns(Sheet)
    Call UnhideRows(Sheet)
End Sub

' Remove todas as linhas e colunas em branco
Public Sub ClearBlanks(ByRef Sheet As Worksheet)
    Call ClearBlankColumns(Sheet)
    Call ClearBlankRows(Sheet)
End Sub

' Remove todas as colunas em branco
Public Sub ClearBlankColumns(ByRef Sheet As Worksheet)
    Dim Used As Range
    Dim Row As Range

    Dim Begin As Variant
    Dim Last As Long

    Dim Current
    Dim Merged
    Dim Remove As rcSetCollection

    Dim Aux As Variant
    Dim LCol As Variant
    Dim Addr As String
    Dim Size As Variant

    Dim PH As rcProgressHelper

    Dim Separator As Variant
    Separator = GetRangeSeparator

    Set Used = GetUsedRange(Sheet)
    Set Used = Sheet.Range("A1", LastCell(Used))

    Begin = Used.Column
    Last = Used.Column + Used.Columns.Count - 1

    Current = Empty
    Merged = Empty

    Set PH = New rcProgressHelper
    PH.Prepare 1000

    ' Checa todos os valores de todas as linhas para detectar colunas em branco no intervalo
    For Each Row In Used.Rows
        PH.DoStep True, Used.Rows.Count, Row.Row

        If IsEmpty(Merged) Then
            Merged = Row.Value

            If IsArray(Merged) Then
                If Not IsArray1d(Merged) Then
                    Merged = SMakeArray1d(Merged, 1, 2)
                End If
            Else
                ReDim Merged(Begin To Last)
                Merged(Begin) = Row.Value
            End If
        Else
            Current = Row.Value

            If IsArray(Current) Then
                If Not IsArray1d(Current) Then
                    Current = SMakeArray1d(Current, 1, 2)
                End If
            Else
                ReDim Current(Begin To Last)
                Current(Begin) = Row.Value
            End If

            For Aux = Begin To Last
                If IsEmpty(Merged(Aux)) And Not IsEmpty(Current(Aux)) Then
                    Merged(Aux) = 1
                End If
            Next Aux
        End If
    Next Row

    PH.Finish True

    Set Used = Nothing

    Addr = Empty
    Size = 0

    PH.Prepare 1000

    Set Remove = New rcSetCollection

    ' Coleta o endereço de todos os intervalos de colunas em branco
    For Aux = Begin To Last
        PH.DoStep True, Last, Aux
        LCol = ColumnLetterFromNumber(Aux)

        If IsEmpty(Merged(Aux)) Then
            If Addr <> "" Then
                If Mid$(Addr, Size, 1) <> ":" Then
                    Addr = Addr & Separator & LCol & ":"
                    Size = Size + Len(Separator) + Len(LCol) + 1
                End If
            Else
                Addr = LCol & ":"
                Size = Len(LCol) + 1
            End If

            If Size >= 200 Then ' Tamanho máximo para função .Range() é 255 caracteres
                Addr = Addr & LCol
                Remove.Add Addr
                Addr = ""
                Size = 0
            End If
        Else
            If Addr <> "" Then
                If Mid$(Addr, Size, 1) = ":" Then
                    LCol = ColumnLetterFromNumber(Aux - 1)

                    Addr = Addr & LCol
                    Size = Size + Len(LCol)
                End If

                If Size >= 200 Then ' Tamanho máximo para função .Range() é 255 caracteres
                    Remove.Add Addr
                    Addr = ""
                    Size = 0
                End If
            End If
        End If
    Next Aux

    If Addr <> "" Then
        If Mid$(Addr, Size, 1) = ":" Then
            LCol = ColumnLetterFromNumber(Aux - 1)

            Addr = Addr & LCol
            Size = Size + Len(LCol)
        End If

        Remove.Add Addr
        Addr = ""
        Size = 0
    End If

    PH.Finish True

    PH.Prepare

    For Aux = Remove.Last To Remove.First Step -1
        PH.DoStep True, Remove.Last, (Remove.Last - Aux)

        Sheet.Range(Remove.List(Aux)).Delete
    Next Aux

    PH.Finish True

    rcFree PH
End Sub

' Remove todas as linhas em branco
Public Sub ClearBlankRows(ByRef Sheet As Worksheet)
    Dim Used As Range
    Dim Col As Range

    Dim Begin As Variant
    Dim Last As Long

    Dim Current
    Dim Merged
    Dim Remove As rcSetCollection

    Dim Aux As Variant
    Dim Addr As String
    Dim Size As Variant

    Dim PH As rcProgressHelper

    Dim Separator As Variant
    Separator = GetRangeSeparator

    Set Used = GetUsedRange(Sheet)
    Set Used = Sheet.Range("A1", LastCell(Used))

    Begin = Used.Row
    Last = Used.Row + Used.Rows.Count - 1

    Current = Empty
    Merged = Empty

    Set PH = New rcProgressHelper
    PH.Prepare 1000

    ' Checa todos os valores de todas as colunas para detectar linhas em branco no intervalo
    For Each Col In Used.Columns
        PH.DoStep True, Used.Columns.Count, Col.Column

        If IsEmpty(Merged) Then
            Merged = Col.Value

            If IsArray(Merged) Then
                If Not IsArray1d(Merged) Then
                    Merged = SMakeArray1d(Merged, 1)
                End If
            Else
                ReDim Merged(Begin To Last)
                Merged(Begin) = Col.Value
            End If
        Else
            Current = Col.Value

            If IsArray(Current) Then
                If Not IsArray1d(Current) Then
                    Current = SMakeArray1d(Current, 1)
                End If
            Else
                ReDim Current(Begin To Last)
                Current(Begin) = Col.Value
            End If

            For Aux = Begin To Last
                If IsEmpty(Merged(Aux)) And Not IsEmpty(Current(Aux)) Then
                    Merged(Aux) = 1
                End If
            Next Aux
        End If
    Next Col

    PH.Finish True

    Set Used = Nothing

    Addr = Empty
    Size = 0

    PH.Prepare 1000

    Set Remove = New rcSetCollection

    ' Coleta o endereço de todos os intervalos de colunas em branco
    For Aux = Begin To Last
        PH.DoStep True, Last, Aux

        If IsEmpty(Merged(Aux)) Then
            If Addr <> "" Then
                If Mid$(Addr, Size, 1) <> ":" Then
                    Addr = Addr & Separator & Aux & ":"
                    Size = Size + Len(Separator) + Len(Aux) + 1
                End If
            Else
                Addr = Aux & ":"
                Size = Len(Aux) + 1
            End If

            If Size >= 200 Then ' Tamanho máximo para função .Range() é 255 caracteres
                Addr = Addr & Aux
                Remove.Add Addr
                Addr = ""
                Size = 0
            End If
        Else
            If Addr <> "" Then
                If Mid$(Addr, Size, 1) = ":" Then
                    Addr = Addr & (Aux - 1)
                    Size = Size + Len(Aux - 1)
                End If
            End If

            If Size >= 200 Then ' Tamanho máximo para função .Range() é 255 caracteres
                Remove.Add Addr
                Addr = ""
                Size = 0
            End If
        End If
    Next Aux

    If Addr <> "" Then
        If Mid$(Addr, Size, 1) = ":" Then
            Addr = Addr & (Aux - 1)
            Size = Size + Len(Aux - 1)
        End If

        Remove.Add Addr
        Addr = ""
        Size = 0
    End If

    PH.Finish True

    PH.Prepare

    For Aux = Remove.Last To Remove.First Step -1
        PH.DoStep True, Remove.Last, (Remove.Last - Aux)
        Sheet.Range(Remove.List(Aux)).Delete
    Next Aux

    PH.Finish True

    rcFree Remove
    rcFree PH
End Sub

' Converte um número de coluna para a respectiva letra
Public Function ColumnLetterFromNumber(ByVal Col As Variant) As String
    ColumnLetterFromNumber = Split(Cells(1, Col).Address(1, 0), "$")(0)
End Function

' Gera referências para intervalos através dos titulos em uma linha
Public Function GetColumns(ByRef Table As Range) As rcKeyValueCollection
    Dim Header As Range
    Dim Columns As rcKeyValueCollection

    Dim Key As String
    Dim i As Long

'    If Table.Rows.Count = 1 Then
'        Set Table = Table.Resize(2)
'    End If

    Set Columns = New rcKeyValueCollection

    For Each Header In Table.Columns
        If Not IsError(Header.Rows(1).Value) Then
            Key = NormalizedString(Header.Rows(1).Value)
        Else
            Key = ""
        End If

        If Not Columns.Exists(Key) And Key <> "" Then
            Call Columns.Add(Key, Header, True)
        Else
            i = 1

            Do While Columns.Exists(Key) Or Key = ""
                Key = NormalizedString(Header.Rows(1).Value) & "[" & i & "]"

                If Not Columns.Exists(Key) And Key <> "" Then
                    Call Columns.Add(Key, Header, True)
                    Exit Do
                Else
                    i = i + 1
                End If
            Loop
        End If
    Next Header

    Set GetColumns = Columns
End Function

Public Function ArrayToRange(ByRef Where As Range, ByRef Arr As Variant, Optional ByVal ClearRange As Boolean = True) As Range
    If Where Is Nothing Then
        Exit Function
    End If

    If ClearRange Then
        Where.UnMerge
        Where.Clear
    End If

    If IsArray2d(Arr) Then
        Set ArrayToRange = Where.Cells(1, 1).Resize(UBound(Arr, 1) - LBound(Arr, 1) + 1, UBound(Arr, 2) - LBound(Arr, 2) + 1)
        ArrayToRange.Value = Arr
    ElseIf IsArray1d(Arr) Then
        Set ArrayToRange = Where.Cells(1, 1).Resize(UBound(Arr, 1) - LBound(Arr, 1) + 1, 1)
        ArrayToRange.Value = SMakeArray2d(Arr, 1)
    ElseIf Not IsArray(Arr) Then
        Set ArrayToRange = Where.Cells(1, 1)
        ArrayToRange.Value = Arr
    Else
        Exit Function
    End If
End Function

Public Function CropRange(ByRef Where As Range, ByRef Mask As Range) As Range
    Dim R As Range

    If Where Is Nothing Or Mask Is Nothing Then
        Exit Function
    End If

    If Application.Intersect(Where, Mask) Is Nothing Then
        Exit Function
    End If

    ' Fix, for each Area
    For Each R In Where
        If Application.Intersect(R, Mask) Is Nothing Then
            If CropRange Is Nothing Then
                Set CropRange = R
            Else
                Set CropRange = Union(CropRange, R)
            End If
        End If
    Next R
End Function

Public Sub BorderAroundGroup(ByRef Where As Range, Optional ByRef InIndexes As Variant, Optional ByVal Style As XlLineStyle = xlContinuous)
    Dim Aux As Variant
    Dim Indexes As Variant
    Dim i, y, z As Long
    Dim Helper, Check As Boolean

    If Where Is Nothing Then
        Exit Sub
    End If

    Aux = Where.Value

    If Not IsMissing(InIndexes) Then
        Indexes = SMakeArray1d(InIndexes, 1)

        If IsArrayInvalid(Indexes) Or IsEmpty(Indexes) Then
            ReDim Indexes(LBound(Aux, 2) To UBound(Aux, 2)) As Variant

            For i = LBound(Indexes, 2) To UBound(Indexes, 2)
                Indexes(i) = i
            Next i
        End If
    Else
        Indexes = Array(LBound(Aux, 2))
    End If

    For y = LBound(Aux) To UBound(Aux)
        Helper = False

        For z = y + 1 To UBound(Aux)
            Check = True

            For i = LBound(Indexes, 1) To UBound(Indexes, 1)
                If Aux(y, Indexes(i)) <> Aux(z, Indexes(i)) Then
                    Check = False

                    Exit For
                End If
            Next i

            If Not Check Then
                Call Where.Offset(y - LBound(Aux)).Resize(z - y).BorderAround(Style)

                Helper = True
                y = z - 1

                Exit For
            End If
        Next z

        If Not Helper Then
            Call Where.Offset(y - LBound(Aux)).Resize(UBound(Aux) - y + 1).BorderAround(Style)

            Exit For
        End If
    Next y
End Sub

' Executa a função .End() em um intervalo de célula para retornar um resultado comum entre estas
Public Function CommonEnd(ByRef Where As Range, Optional ByVal Direction As XlDirection = xlDown) As Range
    Dim Aux As Range
    Dim Min As Range
    Dim At As Range
    Dim R As Range

    If Direction = xlDown Or Direction = xlUp Then
        For Each Aux In Where.Rows(1).Columns
            Set R = Aux.End(Direction)

            Do While R.Row > 1 And _
                     R.Row < Where.Worksheet.Rows.Count And _
                     WorksheetFunction.CountBlank(Where.Worksheet.Range(Aux, R)) = 0
                Set Aux = R
                Set R = R.End(Direction)

                If Min Is Nothing Then
                    Set Min = Aux
                Else
                    If Direction = xlDown Then
                        If Min.Row < Aux.Row Then
                            Set Min = Aux
                        End If
                    Else
                        If Min.Row > Aux.Row Then
                            Set Min = Aux
                        End If
                    End If
                End If
            Loop

            If Direction = xlDown Then
                If (R.Row - Aux.Row) = 1 Then
                    Set At = R
                    Exit For
                End If

                If At Is Nothing Then
                    Set At = R
                Else
                    If Min Is Nothing Then
                        If R.Row < At.Row Then
                            Set At = R
                        End If
                    Else
                        If R.Row > Min.Row And R.Row < At.Row Then
                            Set At = R
                        End If

                        If At.Row < Min.Row Then
                            Set At = Min
                        End If
                    End If
                End If
            Else
                If (R.Row - Aux.Row) = -1 Then
                    Set At = R
                    Exit For
                End If

                If At Is Nothing Then
                    Set At = R
                Else
                    If Min Is Nothing Then
                        If R.Row > At.Row Then
                            Set At = R
                        End If
                    Else
                        If R.Row < Min.Row And R.Row > At.Row Then
                            Set At = R
                        End If

                        If At.Row > Min.Row Then
                            Set At = Min
                        End If
                    End If
                End If
            End If
        Next Aux

        Set CommonEnd = ModifyRangeRow(Where, At.Row)
    ElseIf Direction = xlToLeft Or Direction = xlToRight Then
        For Each Aux In Where.Columns(1).Rows
            Set R = Aux.End(Direction)

            Do While R.Column > 1 And _
                     R.Column < Where.Worksheet.Columns.Count And _
                     WorksheetFunction.CountBlank(Where.Worksheet.Range(Aux, R)) = 0
                Set Aux = R
                Set R = R.End(Direction)

                If Min Is Nothing Then
                    Set Min = Aux
                Else
                    If Direction = xlToRight Then
                        If Min.Column < Aux.Column Then
                            Set Min = Aux
                        End If
                    Else
                        If Min.Column > Aux.Column Then
                            Set Min = Aux
                        End If
                    End If
                End If
            Loop

            If Direction = xlToRight Then
                If (R.Column - Aux.Column) = 1 Then
                    Set At = R
                    Exit For
                End If

                If At Is Nothing Then
                    Set At = R
                Else
                    If Min Is Nothing Then
                        If R.Column < At.Column Then
                            Set At = R
                        End If
                    Else
                        If R.Column > Min.Column And R.Column < At.Column Then
                            Set At = R
                        End If

                        If At.Column < Min.Column Then
                            Set At = Min
                        End If
                    End If
                End If
            Else
                If (R.Column - Aux.Column) = -1 Then
                    Set At = R
                    Exit For
                End If

                If At Is Nothing Then
                    Set At = R
                Else
                    If Min Is Nothing Then
                        If R.Column > At.Column Then
                            Set At = R
                        End If
                    Else
                        If R.Column < Min.Column And R.Column > At.Column Then
                            Set At = R
                        End If

                        If At.Column > Min.Column Then
                            Set At = R
                        End If
                    End If
                End If
            End If
        Next Aux

        Set CommonEnd = ModifyRangeColumn(Where, At.Column)
    End If
End Function

Public Function CopyValues(ByRef Source As Range, ByRef Target As Range, Optional Parse As Boolean = False) As Range
    Set CopyValues = Nothing

    If Source Is Nothing Or Target Is Nothing Then
        Exit Function
    End If

    On Error GoTo ErrHandler

    Set CopyValues = Target.Resize(Source.Rows.Count)

    If Not Parse Then
        Source.Copy
        CopyValues.PasteSpecial xlPasteValues

        Application.CutCopyMode = False
    Else
        CopyValues.Value = Source.Value
    End If

    GoTo CleanUp

ErrHandler:
    Set CopyValues = Nothing

CleanUp:
    Exit Function
End Function

' Aplica uma formulaarray maior que 255 caracteres a um intervalo
' INVALID TODO
Public Sub SetBigArrayFormula(ByRef Where As Range, ByRef FormulaLocal As Variant)
    Const Prefix = "0+"
    Const Suffix = "-0"
    Dim Size As Long
    Dim LP As Long
    Dim FP As Long
    Dim RP As Long
    Dim Placeholder As String
    Dim Stack As Variant
    Dim x As Long
    Dim y As Long
    Dim Test As Boolean

    On Error GoTo CleanUp

Start:
    Debug.Assert False

    Size = Len(FormulaLocal)
    LP = Size - Len(Replace(FormulaLocal, "(", ""))
    RP = Size - Len(Replace(FormulaLocal, ")", ""))

    If LP <> RP Then
        ' Invalid / unsupported formula (i.e. invalid, parenthesis inside quotes)
        GoTo CleanUp
    End If

    ReDim Stack(1 To LP) As Variant
    x = LBound(Stack)

    Placeholder = FormulaLocal

    For y = LBound(Stack) To UBound(Stack)
        RP = InStr(Placeholder, ")")

        If RP = 0 Then
            ' Not found
            GoTo CleanUp
        End If

        LP = InStrRev(Placeholder, "(", RP)

        If LP = 0 Then
            GoTo CleanUp
        End If

        FP = LP

        Stack(x) = Mid(Placeholder, LP, RP - LP + 1)
        Placeholder = Replace(Placeholder, Stack(x), Prefix & x & Suffix)
        x = x + 1
    Next y

    Placeholder = Replace(Placeholder, Prefix, "(T(" & Prefix)
    Placeholder = Replace(Placeholder, Suffix, Suffix & "))")

    Debug.Assert False

    Where.FormulaArray = "=" & 0
    Where.Replace "=0", Placeholder, xlWhole, xlByColumns, False, False, False, False

    Debug.Assert False

    For y = UBound(Stack) To LBound(Stack) Step -1
        Stack(y) = Replace(Stack(y), Prefix, "(T(" & Prefix)
        Stack(y) = Replace(Stack(y), Suffix, Suffix & "))")
        Test = Where.Replace("(T(" & Prefix & y & Suffix & "))", Stack(y), xlPart, xlByColumns, False, False, False, False)
    Next y

    Debug.Assert False

    GoTo CleanUp

    Exit Sub
CleanUp:
    Where.Formula = FormulaLocal
    GoTo Start
End Sub

' Retorna todos os resultados encontrados com um certo texto em um determinado intervalo
Public Function FindAll(ByRef Where As Range, ByVal txt As Variant) As Range
    Dim First As Range ' FIXME
    Dim Aux As Range

    Set FindAll = Nothing

    Set First = Where.Find(What:=txt, LookIn:=xlFormulas, LookAt:=xlPart, SearchDirection:=xlNext)

    Do While True
        If Not Aux Is Nothing Then
            If First.Row = Aux.Row And First.Column = Aux.Column Then
                Exit Do
            End If
        End If

        If Aux Is Nothing Then
            Set Aux = First
        End If

        If FindAll Is Nothing Then
            Set FindAll = Aux
        Else
            Set FindAll = Union(FindAll, Aux)
        End If

        Set Aux = Where.FindNext(Aux)

        If Aux Is Nothing Then
            Set FindAll = Nothing
            Exit Function
        End If
    Loop
End Function

' Checa se planilha existe, caso contrário retorna Nothing
Public Function SafeSheet(ByVal Name As String, Optional ByRef Book As Workbook) As Worksheet
    Dim Sheet As Worksheet

    If Book Is Nothing Then Set Book = ThisWorkbook

    Set SafeSheet = Nothing

    For Each Sheet In Book.Worksheets
        If StrComp(Sheet.Name, Name, vbTextCompare) = 0 Then
            Set SafeSheet = Sheet
            Exit Function
        End If
    Next Sheet
End Function

' Executa a função .SpecialCells sem gerar erros, mas retornando Nothing
Public Function SafeSpecialCells(ByRef Where As Range, ByRef InType As Variant, Optional InValue As Variant) As Range
    Dim Types As Variant
    Dim Values As Variant
    Dim R As Range
    Dim i As Long

    If Where Is Nothing Then
        Exit Function
    End If

    Types = SMakeArray1d(InType)
    Values = SMakeArray1d(InValue)

    If IsArrayInvalid(Types) Then
        Exit Function
    End If

    If IsArrayInvalid(Values) Then
        Values = Empty
    End If

    On Error Resume Next

    For i = LBound(Types) To UBound(Types)
        If Not IsEmpty(Values) Then
            Set R = Where.SpecialCells(Types(i), SafeIndex(Values, i))
        Else
            Set R = Where.SpecialCells(Types(i))
        End If

        If Not R Is Nothing Then
            If SafeSpecialCells Is Nothing Then
                Set SafeSpecialCells = R
            Else
                Set SafeSpecialCells = Union(SafeSpecialCells, R)
            End If

            Set R = Nothing
        End If
    Next i
End Function

Public Sub UnMergeCells(ByRef Where As Range)
    Dim R As Range
    Dim V As Variant

    Application.FindFormat.Clear
    Application.FindFormat.MergeCells = True

    Do
        Set R = Where.Find("*", SearchFormat:=True)

        If R Is Nothing Then
            Set R = Where.Find("", SearchFormat:=True)
        End If

        If Not R Is Nothing Then
            Set R = R.MergeArea
            V = R.Cells(1, 1).Value
            R.UnMerge
            R.Value = V
        End If
    Loop Until R Is Nothing

    Application.FindFormat.Clear
End Sub

' Cria planilha com opção de apagar a antiga
Public Function CreateSheet(ByVal Name As String, Optional ByVal Overwrite As Boolean = False, Optional ByRef Book As Workbook) As Worksheet
    Set CreateSheet = Nothing

    If Book Is Nothing Then Set Book = ThisWorkbook

    If SheetExists(Name, Book) Then
        If Overwrite Then
            If Not DeleteSheet(Name, Book) Then
                Exit Function
            End If
        Else
            Exit Function
        End If
    End If

    Set CreateSheet = Book.Worksheets.Add(After:=Book.Worksheets(Book.Worksheets.Count))

    If Not CreateSheet Is Nothing Then
        CreateSheet.Name = Name
    End If
End Function

' Testa se planilha existe
Public Function SheetExists(ByVal Name As String, Optional ByRef Book As Workbook) As Boolean
    If Book Is Nothing Then Set Book = ThisWorkbook
    SheetExists = (Not SafeSheet(Name, Book) Is Nothing)
End Function

' Apaga uma planilha sem mostrar alertas
Public Function SilentlyDeleteSheet(ByRef Sheet As Worksheet) As Boolean
    Dim Active As Worksheet
    Dim Book As Workbook
    Dim N As String

    SilentlyDeleteSheet = False

    On Error GoTo CleanUp
    Application.DisplayAlerts = False

    Set Active = Application.ActiveSheet

    If Active.Name = Sheet.Name Then
        Set Active = Nothing
    End If

    Set Book = Sheet.Parent
    N = Sheet.Name

    If Sheet.Visible <> xlSheetVisible Then
        Sheet.Visible = xlSheetVisible
    End If

    Sheet.Activate ' Workaround para bug (erro de runtime)
    Sheet.Delete

    SilentlyDeleteSheet = Not SheetExists(N, Book)

    If Not Active Is Nothing Then
        Active.Activate
    End If

CleanUp:
    On Error Resume Next
    Application.DisplayAlerts = True
End Function

' Apaga uma planilha pelo nome
Public Function DeleteSheet(ByVal Name As String, Optional ByRef Book As Workbook) As Boolean
    Dim Sheet As Worksheet

    If Book Is Nothing Then Set Book = ThisWorkbook

    DeleteSheet = False

    For Each Sheet In Book.Worksheets
        If StrComp(Sheet.Name, Name, vbTextCompare) = 0 Then
            DeleteSheet = SilentlyDeleteSheet(Sheet)
            Exit Function
        End If
    Next Sheet
End Function

' Copia uma planilha para o fim do arquivo atual, sem alertas sobre nomes na planilha
Public Function SilentlyImportSheet(ByRef Source As Worksheet, Optional ByVal Name As String, _
                                    Optional ByVal Overwrite As Boolean = False, Optional ByRef Target As Workbook) As Worksheet
    Dim Active As Worksheet
    Dim Aux As XlSheetVisibility

    If IsMissing(Name) Then Name = Source.Name

    If Target Is Nothing Then Set Target = ThisWorkbook

    Set SilentlyImportSheet = Nothing

    If SheetExists(Name, Target) Then
        If Overwrite Then
            DeleteSheet Name, Target
        Else
            Exit Function
        End If
    End If

    On Error GoTo CleanUp
    Application.DisplayAlerts = False

    Set Active = Application.ActiveSheet

    If Source.Visible <> xlSheetVisible Then
        Aux = Source.Visible
        Source.Visible = xlSheetVisible
        Source.Copy After:=Target.Worksheets(Target.Worksheets.Count)
        Set SilentlyImportSheet = Application.ActiveSheet
        Source.Visible = Aux
    Else
        Source.Copy After:=Target.Worksheets(Target.Worksheets.Count)
        Set SilentlyImportSheet = Application.ActiveSheet
    End If

    SilentlyImportSheet.Name = Name

    If Not Active Is Nothing Then
        Active.Activate
    End If

CleanUp:
    On Error Resume Next
    Application.DisplayAlerts = True
End Function

' Retorna um Nome (objeto) sem gerar erros, mas nothing em caso de erro
Public Function SafeName(ByVal RN As String, Optional ByRef Book As Workbook) As Name
    Dim N As Name

    If Book Is Nothing Then
        Set Book = ThisWorkbook
    End If

    Set SafeName = Nothing

    For Each N In Book.Names
        If StrComp(N.Name, RN, vbTextCompare) = 0 Then
            Set SafeName = N
            Exit For
        End If
    Next N
End Function

' Retorna Nothing ao invés de erro ao ler Range de um Nome (objeto) sem Range
Public Function SafeNameRange(ByVal RN As String, Optional ByRef Book As Workbook) As Range
    Dim N As Name

    If Book Is Nothing Then
        Set Book = ThisWorkbook
    End If

    Set SafeNameRange = Nothing

    On Error Resume Next

    For Each N In Book.Names
        If StrComp(N.Name, RN, vbTextCompare) = 0 Then
                Set SafeNameRange = N.RefersToRange
            Exit For
        End If
    Next N
End Function

Public Function SafeNameValue(ByVal RN As String, Optional ByRef Book As Workbook, Optional ByRef Default As Variant) As Variant
    Dim N As Name

    If Book Is Nothing Then
        Set Book = ThisWorkbook
    End If

    SafeNameValue = Default

    On Error Resume Next

    For Each N In Book.Names
        If StrComp(N.Name, RN, vbTextCompare) = 0 Then
                SafeNameValue = N.RefersToRange.Value
            Exit For
        End If
    Next N
End Function

' Executa a função .Range() sem retornar erros
Public Function SafeRange(ByVal Addr As String, Optional ByRef Base As Range) As Range
    Set SafeRange = Nothing

    On Error GoTo Err_Handler

    If Base Is Nothing Then
        Set SafeRange = Range(Addr)
    Else
        Set SafeRange = Base.Range(Addr)
    End If

Err_Handler:
End Function

' Checa se um Nome (objeto) existe no arquivo
Public Function NameExists(ByVal RN As String, Optional ByRef Book As Workbook) As Boolean
    If Book Is Nothing Then Set Book = ThisWorkbook
    NameExists = (Not SafeName(RN, Book) Is Nothing)
End Function

' Atualiza a Range de um Nome (objeto)
Public Function UpdateName(ByVal RN As String, ByRef R As Range, Optional ByRef Book As Workbook) As Name
    Dim N As Name

    If Book Is Nothing Then Set Book = ThisWorkbook

    Set N = SafeName(RN, Book)

    If N Is Nothing Then
        Set N = Book.Names.Add(RN, R.Address(True, True, xlA1), True)
    Else
        R.Name = N.Name
    End If

    Set UpdateName = N
End Function

' Limpa os nomes de intervalos no arquivo
Public Sub ClearNames(Optional ByRef Book As Workbook, Optional Only As Variant = Empty, Optional Except As Variant = Empty)
    Dim Name As Name
    Dim Names As Collection

    Dim Aux As Variant
    Dim Clear As Boolean

    If Book Is Nothing Then Set Book = ThisWorkbook

    If IsObject(Except) Then
        If Except Is Nothing Then
            Except = Empty
        End If
    End If

    If IsObject(Only) Then
        If Only Is Nothing Then
            Only = Empty
        End If
    End If

    Set Names = New Collection

    For Each Name In Book.Names
        Clear = True

        If Not IsEmpty(Only) Then
            If IsArray(Only) Then
                For Each Aux In Only
                    If Not Name.Name Like Aux Then
                        Clear = False
                        GoTo Next_Name
                        Exit For
                    End If
                Next Aux
            Else
                If Not Name.Name Like Only Then
                    Clear = False
                    GoTo Next_Name
                End If
            End If
        End If

        If Not IsEmpty(Except) Then
            If IsArray(Except) Then
                For Each Aux In Except
                    If Name.Name Like Aux Then
                        Clear = False
                        GoTo Next_Name
                        Exit For
                    End If
                Next Aux
            Else
                If Name.Name Like Except Then
                    Clear = False
                    GoTo Next_Name
                End If
            End If
        End If

        If Clear Then
            Names.Add Name
        End If

Next_Name:
    Next Name

    On Error Resume Next

    For Each Name In Names
        Name.Delete
    Next Name

    Set Names = Nothing
End Sub

' Última coluna utilizada
Public Function LastColumn(ByRef R As Range) As Range
    Set LastColumn = R.Cells(1, R.Columns.Count).EntireColumn
End Function

' Última linha utilizada
Public Function LastRow(ByRef R As Range) As Range
    Set LastRow = R.Cells(R.Rows.Count, 1).EntireRow
End Function

' Última célula utilizada
Public Function LastCell(ByRef R As Range) As Range
    Set LastCell = R.Cells(R.Rows.Count, R.Columns.Count)
End Function

' Intervalo de células utilizado (UsedRange normal não é confiável)
Public Function GetUsedRange(ByRef Sheet As Worksheet) As Range
    Dim First As Range
    Dim Last As Range
    Dim Aux As Range

    Set GetUsedRange = Nothing

    With Sheet
        Set First = .UsedRange.Cells(1, 1)

        Set Aux = .Cells.Find("*", First, xlFormulas, xlPart, xlByRows, xlPrevious)
        Set Last = .Cells.Find("*", First, xlFormulas, xlPart, xlByColumns, xlPrevious)

        If Last Is Nothing Then
            Set Last = LastCell(.UsedRange)
        Else
            Set Last = Sheet.Cells(Aux.Row, Last.Column)
        End If

        Set GetUsedRange = .Range(First, Last)
    End With
End Function

Public Function GetSheetWindow(ByRef Sheet As Worksheet) As Window
    Sheet.Activate
    Set GetSheetWindow = Application.ActiveWindow

    Exit Function
End Function

Public Function ExtendRange(ByRef From As Range, ByRef Target As Range) As Range
    Set ExtendRange = From.Resize(Target.Row - From.Row + 1, Target.Column - From.Column + 1)
End Function


' Redimensiona um intervalo de células
Public Function RangeMid(ByRef Source As Range, _
                         Optional ByVal MoveRows As Long = 0, Optional ByVal MoveColumns As Long = 0, _
                         Optional ByVal AddRows As Long = 0, Optional ByVal AddColumns As Long = 0) As Range
    Dim TotalRows As Variant
    Dim TotalColumns As Variant

    Set RangeMid = Source
    Set RangeMid = RangeMid.Offset(MoveRows, MoveColumns)

    TotalRows = RangeMid.Rows.Count - MoveRows + AddRows
    TotalColumns = RangeMid.Columns.Count - MoveColumns + AddColumns

    If TotalRows <= 0 Then
        TotalRows = 1
    End If

    If TotalColumns <= 0 Then
        TotalColumns = 1
    End If

    Set RangeMid = RangeMid.Resize(TotalRows, TotalColumns)
End Function

' Modifica a coluna de um intervalo
Public Function ModifyRangeColumn(ByRef R As Range, ByRef Column As Variant) As Range
    With R.Worksheet
        Set ModifyRangeColumn = .Cells(LastRow(R).Row, Column)
        Set ModifyRangeColumn = .Range(.Cells(R.Row, Column), ModifyRangeColumn)
    End With
End Function

' Modifica a linha de um intervalo
Public Function ModifyRangeRow(ByRef R As Range, ByRef Row As Variant) As Range
    With R.Worksheet
        Set ModifyRangeRow = .Cells(Row, LastColumn(R).Column)
        Set ModifyRangeRow = .Range(ModifyRangeRow, .Cells(Row, R.Column))
    End With
End Function

' Clear formulas
Public Sub ClearFormulas(ByRef Sheet As Worksheet)
    With GetUsedRange(Sheet)
        .Copy
        .PasteSpecial xlPasteValues
    End With

    Application.CutCopyMode = False
End Sub

' Desbloqueia planilhas via bruteforce
Public Function UnprotectWorkbook(ByRef Book As Workbook) As Variant
    Dim i, i1, i2, i3, i4, i5, i6 As Integer, j As Integer, K As Integer, L As Integer, m As Integer, N As Integer
    Dim Aux As Variant

    Dim PH As rcProgressHelper

    UnprotectWorkbook = False

    If Not Book.ProtectStructure And Not Book.ProtectWindows Then
        UnprotectWorkbook = ""
        Exit Function
    End If

    On Error Resume Next

    Set PH = New rcProgressHelper
    PH.Prepare 100, 1, "Testando: ", " senhas."

    For i = 65 To 66
        For j = 65 To 66
            For K = 65 To 66
                For L = 65 To 66
                    For m = 65 To 66
                        For i1 = 65 To 66
                            For i2 = 65 To 66
                                For i3 = 65 To 66
                                    For i4 = 65 To 66
                                        For i5 = 65 To 66
                                            For i6 = 65 To 66
                                                For N = 32 To 126
                                                    PH.DoStep True

                                                    Aux = Chr(i) & Chr(j) & Chr(K) & Chr(L) & _
                                                    Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
                                                    Chr(i4) & Chr(i5) & Chr(i6) & Chr(N)

                                                    Book.Unprotect Aux

                                                    If Not Book.ProtectStructure And Not Book.ProtectWindows Then
                                                        UnprotectWorkbook = Aux
                                                        GoTo CleanUp
                                                    End If
                                                Next
                                            Next
                                        Next
                                    Next
                                Next
                            Next
                        Next
                    Next
                Next
            Next
        Next
    Next

CleanUp:
    PH.Finish True
    rcFree PH
End Function

Public Function UnprotectAllSheets(ByRef Book As Workbook) As Variant
    Dim Password As Variant
    Dim Sheet As Worksheet

    Password = UnprotectSheet(Book.Worksheets(1))

    If Password <> False Then
        For Each Sheet In Book.Worksheets
            Sheet.Unprotect Password
        Next Sheet
    End If

    UnprotectAllSheets = Password
End Function

Public Function UnprotectSheet(ByRef Sheet As Worksheet) As Variant
    Dim i, i1, i2, i3, i4, i5, i6 As Integer, j As Integer, K As Integer, L As Integer, m As Integer, N As Integer
    Dim Aux As Variant

    Dim PH As rcProgressHelper

    UnprotectSheet = False

    If Not Sheet.ProtectContents Then
        UnprotectSheet = ""
        Exit Function
    End If

    On Error Resume Next

    Set PH = New rcProgressHelper
    PH.Prepare 100, 1, "Testando: ", " senhas."

    For i = 65 To 66
        For j = 65 To 66
            For K = 65 To 66
                For L = 65 To 66
                    For m = 65 To 66
                        For i1 = 65 To 66
                            For i2 = 65 To 66
                                For i3 = 65 To 66
                                    For i4 = 65 To 66
                                        For i5 = 65 To 66
                                            For i6 = 65 To 66
                                                For N = 32 To 126
                                                    PH.DoStep True

                                                    Aux = Chr(i) & Chr(j) & Chr(K) & Chr(L) & _
                                                    Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
                                                    Chr(i4) & Chr(i5) & Chr(i6) & Chr(N)

                                                    Sheet.Unprotect Aux

                                                    If Not Sheet.ProtectContents Then
                                                        UnprotectSheet = Aux
                                                        GoTo CleanUp
                                                    End If
                                                Next
                                            Next
                                        Next
                                    Next
                                Next
                            Next
                        Next
                    Next
                Next
            Next
        Next
    Next

CleanUp:
    PH.Finish True
    rcFree PH
End Function
