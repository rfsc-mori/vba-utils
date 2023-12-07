Attribute VB_Name = "rChart"
Option Explicit

' Name: rChart
' Version: 0.15
' Depends: rArray
' Author: Rafael Fillipe Silva
' Description: ...

Public Function CreateChart(ByRef Sheet As Worksheet, Optional Title As Variant, _
                            Optional ByVal ChartType As XlChartType, Optional Labels As Boolean = True, _
                            Optional ByRef TopLeft As Range, Optional ByRef BottomRight As Range, _
                            Optional ByRef Source As Range, Optional ByVal PlotBy As Variant) As Shape
    Dim S As Shape
    Dim c As Chart
    Dim Left As Variant
    Dim Top As Variant
    Dim Width As Variant
    Dim Height As Variant

    If Sheet Is Nothing Then
        Exit Function
    End If

    If Not TopLeft Is Nothing Then
        Top = TopLeft.Top
        Left = TopLeft.Left

        If Not BottomRight Is Nothing Then
            Width = BottomRight.Left + BottomRight.Width - Left
            Height = BottomRight.Top + BottomRight.Height - Top
        End If
    End If

    Set S = Sheet.Shapes.AddChart(ChartType, Left, Top, Width, Height)

    If S Is Nothing Then
        Exit Function
    End If

    Set c = S.Chart

    If c Is Nothing Then
        S.Delete
        Exit Function
    End If

    If Not IsMissing(Title) Then
        Call SetChartTitle(c, Title)
    End If

    If Not Source Is Nothing Then
        Call c.SetSourceData(Source, PlotBy)
    End If

    If Labels Then
        If c.SeriesCollection.Count > 0 Then
            Call c.SetElement(msoElementDataLabelCenter)
        End If
    End If

    Set CreateChart = S
End Function

Public Sub SetChartTitle(ByRef c As Chart, ByRef Title As Variant)
    If c Is Nothing Then
        Exit Sub
    End If

    On Error Resume Next

    Call c.SetElement(msoElementChartTitleAboveChart)
    c.ChartTitle.Text = Title
End Sub

Public Sub SetChartValueAxisUnit(ByRef c As Chart, ByVal Unit As Long)
    If c Is Nothing Then
        Exit Sub
    End If

    On Error Resume Next

    c.Axes(xlValue).MajorUnit = Unit
End Sub

Public Sub SetChartSeriesName(ByRef c As Chart, ByRef InNames As Variant)
    Dim Names As Variant
    Dim i As Long

    If c Is Nothing Then
        Exit Sub
    End If

    Names = SMakeArray1d(InNames, 1)

    If IsArrayInvalid(Names) Then
        Exit Sub
    End If

    For i = LBound(Names) To UBound(Names)
        If Not IsEmpty(Names(i)) And c.SeriesCollection.Count >= i Then
            c.SeriesCollection(i).Name = Names(i)
        End If
    Next i
End Sub

Public Sub SetChartSeriesColor(ByRef c As Chart, ByRef InColors As Variant)
    Dim Colors As Variant
    Dim i As Long

    If c Is Nothing Then
        Exit Sub
    End If

    Colors = SMakeArray1d(InColors, 1)

    If IsArrayInvalid(Colors) Then
        Exit Sub
    End If

    For i = LBound(Colors) To UBound(Colors)
        If Not IsEmpty(Colors(i)) And c.SeriesCollection.Count >= i Then
            c.SeriesCollection(i).Format.Fill.ForeColor.RGB = Colors(i)
        End If
    Next i
End Sub

Public Sub SetChartSeriesNumberFormat(ByRef c As Chart, ByRef InFormats As Variant)
    Dim Formats As Variant
    Dim i As Long

    If c Is Nothing Then
        Exit Sub
    End If

    Formats = SMakeArray1d(InFormats, 1)

    If IsArrayInvalid(Formats) Then
        Exit Sub
    End If

    For i = LBound(Formats) To UBound(Formats)
        If Not IsEmpty(Formats(i)) And c.SeriesCollection.Count >= i Then
            c.SeriesCollection(i).DataLabels.NumberFormat = Formats(i)
        End If
    Next i
End Sub

Public Sub SetChartSeries3d(ByRef c As Chart, Optional ByRef InSeries As Variant, _
                            Optional ByVal ThreeD As Boolean = True, Optional ByVal Value As Double = 6#)
    Dim Series As Variant
    Dim i As Long

    If c Is Nothing Then
        Exit Sub
    End If

    If Not IsMissing(InSeries) Then
        Series = SMakeArray1d(InSeries, 1)
    Else
        i = c.SeriesCollection.Count

        If i > 0 Then
            ReDim Series(1 To i) As Variant

            For i = LBound(Series) To UBound(Series)
                Series(i) = i
            Next i
        End If
    End If

    If IsArrayInvalid(Series) Then
        Exit Sub
    End If

    For i = LBound(Series) To UBound(Series)
        With c.SeriesCollection(Series(i)).Format.ThreeD
            If ThreeD Then
                .BevelTopInset = Value
                .BevelTopDepth = Value
                .BevelBottomInset = Value
                .BevelBottomDepth = Value
            Else
                .BevelTopInset = 0
                .BevelTopDepth = 0
                .BevelBottomInset = 0
                .BevelBottomDepth = 0
            End If
        End With
    Next i
End Sub

Public Sub SetChartSeriesLine(ByRef c As Chart, Optional ByRef InSeries As Variant, _
                              Optional ByVal Visible As Boolean = True, Optional ByVal Label As Boolean = True, _
                              Optional Style As Variant, Optional ByVal RGB As Variant, _
                              Optional ByVal Marker As Boolean = False, _
                              Optional ByVal MarkerBG As Variant, Optional ByVal MarkerFG As Variant)
    Dim Series As Variant
    Dim i As Long

    If c Is Nothing Then
        Exit Sub
    End If

    If Not IsMissing(InSeries) Then
        Series = SMakeArray1d(InSeries, 1)
    Else
        i = c.SeriesCollection.Count
        ReDim Series(1 To i) As Variant

        For i = LBound(Series) To UBound(Series)
            Series(i) = i
        Next i
    End If

    If IsArrayInvalid(Series) Then
        Exit Sub
    End If

    For i = LBound(Series) To UBound(Series)
        With c.SeriesCollection(Series(i))
            .ChartType = xlLineStacked

            If Not IsMissing(Style) Then
                .Format.Line.DashStyle = Style
            End If

            .Format.Line.Visible = False
            .Format.Line.Visible = Visible

            If Not IsMissing(RGB) Then
                .Format.Line.ForeColor.RGB = RGB
            End If

            If Label Then
                .HasDataLabels = True
                .DataLabels.Position = xlLabelPositionAbove
            Else
                .HasDataLabels = False
            End If

            If Marker Then
                .MarkerStyle = xlMarkerStyleCircle
                .MarkerSize = 5

                If Not IsMissing(MarkerBG) Then
                    If MarkerBG <> xlColorIndexAutomatic And MarkerBG <> xlColorIndexNone Then
                        .MarkerBackgroundColor = MarkerBG
                    Else
                        .MarkerBackgroundColorIndex = MarkerBG
                    End If
                End If

                If Not IsMissing(MarkerFG) Then
                    If MarkerFG <> xlColorIndexAutomatic And MarkerFG <> xlColorIndexNone Then
                        .MarkerForegroundColor = MarkerFG
                    Else
                        .MarkerForegroundColorIndex = MarkerFG
                    End If
                End If
            Else
                .MarkerStyle = xlMarkerStyleNone
            End If
        End With
    Next i
End Sub

Public Sub DeleteChartLegendEntries(ByRef c As Chart, Optional ByRef InEntries As Variant)
    Dim Entries As Variant
    Dim i As Long

    If c Is Nothing Then
        Exit Sub
    End If

    If Not IsMissing(InEntries) Then
        Entries = SMakeArray1d(InEntries, 1)
    Else
        i = c.Legend.LegendEntries.Count
        ReDim Entries(1 To i) As Variant

        For i = LBound(Entries) To UBound(Entries)
            Entries(i) = i
        Next i
    End If

    If IsArrayInvalid(Entries) Then
        Exit Sub
    End If

    For i = LBound(Entries) To UBound(Entries)
        c.Legend.LegendEntries(Entries(i)).Delete
    Next i
End Sub
