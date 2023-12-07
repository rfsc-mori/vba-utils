Attribute VB_Name = "rShape"
Option Explicit

' Name: rShape
' Version: 0.9
' Author: Rafael Fillipe Silva
' Description: ...

Public Function FindShape(ByRef Sheet As Worksheet, Optional ShapeType As MsoShapeType = msoAutoShape) As Shape
    Dim x As Long

    On Error GoTo ErrHandler

    With Sheet.Shapes
        For x = 1 To .Count
            If .Item(x).Type = ShapeType Then
                Set FindShape = .Item(x)
                Exit Function
            End If
        Next x
    End With

ErrHandler:
End Function

Public Function FindShapes(ByRef Sheet As Worksheet, Optional ShapeType As MsoShapeType = msoAutoShape) As ShapeRange
    Dim List() As Variant
    Dim x, y As Long

    On Error GoTo ErrHandler

    With Sheet.Shapes
        ReDim List(0 To .Count) As Variant

        y = LBound(List) - 1

        For x = 1 To .Count
            If .Item(x).Type = ShapeType Then
                y = y + 1
                List(y) = .Item(x).Name
            End If
        Next x

        If y >= LBound(List) Then
            ReDim Preserve List(0 To y) As Variant
            Set FindShapes = .Range(List)
        End If
    End With

ErrHandler:
End Function

' Procura um objeto Shape que contenha o texto "What".
Public Function FindInShapes(ByRef Sheet As Worksheet, ByVal What As String) As Shape
    Dim S As Shape
    Dim Aux As String

    For Each S In Sheet.Shapes
        Aux = SafeGetShapeText(S)

        If Aux <> "" Then
            If InStr(1, Aux, What, vbTextCompare) >= 1 Then
                Set FindInShapes = S
                Exit Function
            End If
        End If
    Next S
End Function

Public Function SafeGetShapeText(ByRef S As Shape) As Variant
    On Error Resume Next
    SafeGetShapeText = S.TextFrame.Characters.Text
End Function
