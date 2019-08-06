Attribute VB_Name = "ModuleReGra"
'''그래프의 원점을 x,y데이타 최소값-alpha로 한다.
'''alpha=(max-min)/10
Sub ScatterPlot(OutSheetName, Left, Top, Width, Height, _
    xRn, yRn, xname, yname, Optional CorrTest As Boolean = False, _
    Optional myTitle As String = "적합선그림")
    
    Dim plot As ChartObject: Dim SampCorr, TStat, significance As Double
    Dim adjTitle As String: Dim temp1, temp2 As Double
    Dim PointPos1, PointPos2 As String
    
    adjTitle = myTitle
    If CorrTest = True Then
        SampCorr = Application.Correl(xRn, yRn)
        If SampCorr <> 1 Then
            TStat = Sqr(yRn.count - 2) * SampCorr / Sqr(1 - SampCorr ^ 2)
            significance = Application.TDist(Abs(TStat), yRn.count - 2, 2)
        Else: significance = 0
        End If
        adjTitle = adjTitle & Chr(10) & _
            "r=" & Format(SampCorr, "0.00") & Chr(10) & _
            "H0:ρ=0 ; " & "유의확률=" & Format(significance, "0.0000")
    End If
    
    Set plot = Worksheets(OutSheetName).ChartObjects.Add(Left, Top, Width, Height)
    plot.Chart.ChartWizard Source:=yRn, _
        Gallery:=xlXYScatter, Format:=1, HasLegend:=False, _
        CategoryTitle:=xname, ValueTitle:=yname, _
        title:=adjTitle
    
    With plot.Chart.ChartTitle
        .Font.Size = 10
        .Font.Bold = True
        If CorrTest = True Then
            .Characters(19).Font.Bold = False
        End If
    End With
    
    With plot.Chart.PlotArea.Border
        .Weight = xlThin
        .LineStyle = xlAutomatic
        .ColorIndex = 16
    End With

    With plot.Chart.SeriesCollection(1)
        .XValues = xRn
        '.MarkerBackgroundColorIndex = 3
        '.MarkerForegroundColorIndex = 3
        .MarkerStyle = xlCircle
        .MarkerSize = 3
    End With
    
    temp1 = (Application.Max(xRn) - Application.Min(xRn)) / 10
    If temp1 <> 0 Then
        PointPos1 = CStrNumPoint(temp1 * 10, xRn.count)
        With plot.Chart.Axes(xlCategory)
            .MinimumScale = Application.Min(xRn) - temp1
            .MaximumScale = Application.Max(xRn) + temp1
            .TickLabels.NumberFormat = PointPos1
        End With
    End If
    
    temp2 = (Application.Max(yRn) - Application.Min(yRn)) / 10
    If temp2 <> 0 Then
        PointPos2 = CStrNumPoint(temp2 * 10, yRn.count)
        With plot.Chart.Axes(xlValue)
            .MinimumScale = Application.Min(yRn) - temp2
            .MaximumScale = Application.Max(yRn) + temp2
            .TickLabels.NumberFormat = PointPos2
        End With
    End If
    
    With plot.Chart.Axes(xlValue)
        .TickLabels.Font.Size = 8
        .AxisTitle.Orientation = xlVertical
        .AxisTitle.Font.Size = 8
        .TickLabelPosition = xlLow
        .Border.LineStyle = xlNone  '필요시 삭제
    End With
    
    With plot.Chart.Axes(xlCategory)
        .TickLabels.Font.Size = 8
        .AxisTitle.Orientation = xlHorizontal
        .AxisTitle.Font.Size = 8
        .Border.Weight = xlHairline
        '.Border.LineStyle = xlNone
        .MajorTickMark = xlNone
        .MinorTickMark = xlNone
        .TickLabelPosition = xlLow
        .HasMajorGridlines = False
        .TickLabelPosition = xlLow
        .TickLabels.NumberFormatLinked = True
        '.Crosses = xlCustom
        '.CrossesAt = .MinimumScale
    End With
    
    plot.Chart.SeriesCollection(1).Trendlines.Add
End Sub

'''관측순서별 그래프 그리기
Sub OrderScatterPlot(OutSheetName, Left, Top, Width, Height, _
    Rn, Rnname, RefLine As Double)
    
    Dim plot As ChartObject
    Dim temp1 As Double
    Dim a, b As Double
    
    
    
    
    Set plot = Worksheets(OutSheetName).ChartObjects.Add(Left, Top, Width, Height)
    plot.Chart.ChartWizard Source:=Rn, _
        Gallery:=xlLine, Format:=4, HasLegend:=False, _
        CategoryTitle:="관측순서", ValueTitle:=Rnname, _
        title:=Rnname & " vs. " & "관측순서"
    
    With plot.Chart.ChartTitle
        .Font.Size = 10
        .Font.Bold = True
    End With

    With plot.Chart.PlotArea.Border
        .Weight = xlThin
        .LineStyle = xlAutomatic
        .ColorIndex = 16
    End With

    '''y축의 눈금 조절
    temp1 = (Application.Max(Rn) - Application.Min(Rn)) / 10
    If temp1 <> 0 Then
        With plot.Chart.Axes(xlValue)
            .MinimumScale = Application.Min(Rn) - temp1
            .MaximumScale = Application.Max(Rn) + temp1
            .TickLabels.NumberFormat = CStrNumPoint(temp1 * 10, Rn.count)
        End With
    End If
    a = plot.Chart.Axes(xlValue).MinimumScale
    b = plot.Chart.Axes(xlValue).MaximumScale
    
    With plot.Chart.SeriesCollection(1)
        .Border.Weight = xlThin
        .Border.LineStyle = xlNone
        .MarkerBackgroundColorIndex = 11
        .MarkerForegroundColorIndex = 11
        .MarkerStyle = xlCircle
        .MarkerSize = 3
    End With
    With plot.Chart.Axes(xlCategory)
        .TickLabels.Font.Size = 8
        .AxisTitle.Orientation = xlHorizontal
        .AxisBetweenCategories = True
        .AxisTitle.Font.Size = 8
        .MajorTickMark = xlNone
        .MinorTickMark = xlNone
        .TickLabelPosition = xlLow
        '.Border.LineStyle = xlNone  '필요시 삭제       'y=0 선그리기 위해
        .TickLabels.Orientation = xlHorizontal
    End With
    With plot.Chart.Axes(xlValue)
        .HasMajorGridlines = False
        .TickLabels.Font.Size = 8
        .AxisTitle.Orientation = xlVertical
        .AxisTitle.Font.Size = 8
        .TickLabelPosition = xlLow
        .Border.LineStyle = xlNone  '필요시 삭제
    End With
        
    If RefLine <> 0 Then
        plot.Chart.SeriesCollection.NewSeries
        With plot.Chart.SeriesCollection(2)
            .XValues = Array(0, Rn.count + 1)
            .Values = Array(RefLine, RefLine)
            .ChartType = xlXYScatterSmoothNoMarkers
            .Border.ColorIndex = 3
        End With
        plot.Chart.SeriesCollection.NewSeries
        With plot.Chart.SeriesCollection(3)
            .XValues = Array(0, Rn.count + 1)
            .Values = Array(-RefLine, -RefLine)
            .ChartType = xlXYScatterSmoothNoMarkers
            .Border.ColorIndex = 3
        End With
    
        With plot.Chart
            .Axes(xlValue, xlSecondary).MinimumScale = a
            .Axes(xlValue, xlSecondary).MaximumScale = b
            .Axes(xlCategory, xlSecondary).MinimumScale = 0
            .Axes(xlCategory, xlSecondary).MaximumScale = Rn.count + 0.5
        End With
    
        With plot.Chart.Axes(xlValue, xlSecondary)
            .MajorTickMark = xlNone
            .MinorTickMark = xlNone
            .TickLabelPosition = xlNone
        End With
    
        With plot.Chart.Axes(xlCategory, xlSecondary)
            .MajorTickMark = xlNone
            .MinorTickMark = xlNone
            .TickLabelPosition = xlNone
        End With
    End If
    plot.Chart.SeriesCollection(1).Trendlines.Add
    
End Sub

'''그래프의 원점을 x,y데이타 최소값-alpha로 한다.
'''alpha=(max-min)/10
Function ScatterPlotprv(OutSheetName, Left, Top, Width, Height, _
    xRn, yRn, xname, yname, Optional CorrTest As Boolean = False, _
    Optional myTitle As String = "적합선그림") As String
    
    
    Dim plot As ChartObject: Dim SampCorr, TStat, significance As Double
    Dim adjTitle As String: Dim temp1, temp2 As Double
    Dim PointPos1, PointPos2 As String
    
    adjTitle = myTitle
    If CorrTest = True Then
        SampCorr = Application.Correl(xRn, yRn)
        If SampCorr <> 1 Then
            TStat = Sqr(yRn.count - 2) * SampCorr / Sqr(1 - SampCorr ^ 2)
            significance = Application.TDist(Abs(TStat), yRn.count - 2, 2)
        Else: significance = 0
        End If
        adjTitle = adjTitle & Chr(10) & _
            "r=" & Format(SampCorr, "0.00") & Chr(10) & _
            "H0:ρ=0 ; " & "유의확률=" & Format(significance, "0.0000")
    End If
    
    
    Worksheets(OutSheetName).Activate
    Worksheets(OutSheetName).Cells(Top + 5, 1).Select
    Worksheets(OutSheetName).Cells(Top + 5, 1).Activate
    
    Set plot = Worksheets(OutSheetName).ChartObjects.Add(Left, Top, Width, Height)
    plot.Chart.ChartWizard Source:=yRn, _
        Gallery:=xlXYScatter, Format:=1, HasLegend:=False, _
        CategoryTitle:=xname, ValueTitle:=yname, _
        title:=adjTitle
    
    With plot.Chart.ChartTitle
        .Font.Size = 10
        .Font.Bold = True
        If CorrTest = True Then
            .Characters(19).Font.Bold = False
        End If
    End With
    
    With plot.Chart.PlotArea.Border
        .Weight = xlThin
        .LineStyle = xlAutomatic
        .ColorIndex = 16
    End With

    With plot.Chart.SeriesCollection(1)
        .XValues = xRn
        '.MarkerBackgroundColorIndex = 3
        '.MarkerForegroundColorIndex = 3
        .MarkerStyle = xlCircle
        .MarkerSize = 3
    End With
    
    temp1 = (Application.Max(xRn) - Application.Min(xRn)) / 10
    If temp1 <> 0 Then
        PointPos1 = CStrNumPoint(temp1 * 10, xRn.count)
        With plot.Chart.Axes(xlCategory)
            .MinimumScale = Application.Min(xRn) - temp1
            .MaximumScale = Application.Max(xRn) + temp1
            .TickLabels.NumberFormat = PointPos1
        End With
    End If
    
    temp2 = (Application.Max(yRn) - Application.Min(yRn)) / 10
    If temp2 <> 0 Then
        PointPos2 = CStrNumPoint(temp2 * 10, yRn.count)
        With plot.Chart.Axes(xlValue)
            .MinimumScale = Application.Min(yRn) - temp2
            .MaximumScale = Application.Max(yRn) + temp2
            .TickLabels.NumberFormat = PointPos2
        End With
    End If
    
    With plot.Chart.Axes(xlValue)
        .TickLabels.Font.Size = 8
        .AxisTitle.Orientation = xlVertical
        .AxisTitle.Font.Size = 8
        .TickLabelPosition = xlLow
        .Border.LineStyle = xlNone  '필요시 삭제
    End With
    
    With plot.Chart.Axes(xlCategory)
        .TickLabels.Font.Size = 8
        .AxisTitle.Orientation = xlHorizontal
        .AxisTitle.Font.Size = 8
        .Border.Weight = xlHairline
        '.Border.LineStyle = xlNone
        .MajorTickMark = xlNone
        .MinorTickMark = xlNone
        .TickLabelPosition = xlLow
        .HasMajorGridlines = False
        .TickLabelPosition = xlLow
        .TickLabels.NumberFormatLinked = True
        '.Crosses = xlCustom
        '.CrossesAt = .MinimumScale
    End With

'
'    temp1.Chart.SeriesCollection(2).Trendlines.Add Type:=xlLinear
'
'    With temp1.Chart.SeriesCollection(2).Trendlines(1).Border
'        .ColorIndex = 3
'        .Weight = xlThin
'        .LineStyle = xlContinuous
'    End With
plot.Chart.SeriesCollection(1).Trendlines.Add


 ScatterPlotprv = plot.Name

End Function
Function OrderScatterPlotprv(OutSheetName, Left, Top, Width, Height, Rn, Rnname, RefLine As Double) As String
    
    
    Dim plot As ChartObject
    Dim temp1 As Double
    Dim a, b As Double
    
    Set plot = Worksheets(OutSheetName).ChartObjects.Add(Left, Top, Width, Height)
    plot.Chart.ChartWizard Source:=Rn, _
        Gallery:=xlLine, Format:=4, HasLegend:=False, _
        CategoryTitle:="관측순서", ValueTitle:=Rnname, _
        title:=Rnname & " vs. " & "관측순서"
    
    With plot.Chart.ChartTitle
        .Font.Size = 10
        .Font.Bold = True
    End With

    With plot.Chart.PlotArea.Border
        .Weight = xlThin
        .LineStyle = xlAutomatic
        .ColorIndex = 16
    End With
    
    
    Worksheets(OutSheetName).Activate
    Worksheets(OutSheetName).Cells(activePt + 5, 1).Select
    Worksheets(OutSheetName).Cells(activePt + 5, 1).Activate

    '''y축의 눈금 조절
    temp1 = (Application.Max(Rn) - Application.Min(Rn)) / 10
    If temp1 <> 0 Then
        With plot.Chart.Axes(xlValue)
            .MinimumScale = Application.Min(Rn) - temp1
            .MaximumScale = Application.Max(Rn) + temp1
            .TickLabels.NumberFormat = CStrNumPoint(temp1 * 10, Rn.count)
        End With
    End If
    a = plot.Chart.Axes(xlValue).MinimumScale
    b = plot.Chart.Axes(xlValue).MaximumScale
    
    With plot.Chart.SeriesCollection(1)
        .Border.Weight = xlThin
        .Border.LineStyle = xlNone
        .MarkerBackgroundColorIndex = 11
        .MarkerForegroundColorIndex = 11
        .MarkerStyle = xlCircle
        .MarkerSize = 3
    End With
    With plot.Chart.Axes(xlCategory)
        .TickLabels.Font.Size = 8
        .AxisTitle.Orientation = xlHorizontal
        .AxisBetweenCategories = True
        .AxisTitle.Font.Size = 8
        .MajorTickMark = xlNone
        .MinorTickMark = xlNone
        .TickLabelPosition = xlLow
        '.Border.LineStyle = xlNone  '필요시 삭제       'y=0 선그리기 위해
        .TickLabels.Orientation = xlHorizontal
    End With
    With plot.Chart.Axes(xlValue)
        .HasMajorGridlines = False
        .TickLabels.Font.Size = 8
        .AxisTitle.Orientation = xlVertical
        .AxisTitle.Font.Size = 8
        .TickLabelPosition = xlLow
        .Border.LineStyle = xlNone  '필요시 삭제
    End With
        
    If RefLine <> 0 Then
        plot.Chart.SeriesCollection.NewSeries
        With plot.Chart.SeriesCollection(2)
            .XValues = Array(0, Rn.count + 1)
            .Values = Array(RefLine, RefLine)
            .ChartType = xlXYScatterSmoothNoMarkers
            .Border.ColorIndex = 3
        End With
        plot.Chart.SeriesCollection.NewSeries
        With plot.Chart.SeriesCollection(3)
            .XValues = Array(0, Rn.count + 1)
            .Values = Array(-RefLine, -RefLine)
            .ChartType = xlXYScatterSmoothNoMarkers
            .Border.ColorIndex = 3
        End With
    
        With plot.Chart
            .Axes(xlValue, xlSecondary).MinimumScale = a
            .Axes(xlValue, xlSecondary).MaximumScale = b
            .Axes(xlCategory, xlSecondary).MinimumScale = 0
            .Axes(xlCategory, xlSecondary).MaximumScale = Rn.count + 0.5
        End With
    
        With plot.Chart.Axes(xlValue, xlSecondary)
            .MajorTickMark = xlNone
            .MinorTickMark = xlNone
            .TickLabelPosition = xlNone
        End With
    
        With plot.Chart.Axes(xlCategory, xlSecondary)
            .MajorTickMark = xlNone
            .MinorTickMark = xlNone
            .TickLabelPosition = xlNone
        End With
    End If
    
        plot.Chart.SeriesCollection(1).Trendlines.Add
    OrderScatterPlotprv = plot.Name
    
End Function
