Attribute VB_Name = "scatterModule"
'''그래프의 원점을 x,y데이타 최소값-alpha로 한다.
'''alpha=(max-min)/10
Sub ScatterPlot(OutSheetName, Left, Top, Width, Height, _
    xRn, yRn, xnames, xname, yname, Optional CorrTest As Boolean = False, _
    Optional myTitle As String = "산점도(Scatter Plot)")
    
    Dim plot As ChartObject: Dim SampCorr, TStat, significance As Double
    Dim adjTitle As String: Dim temp1, temp2 As Double
    Dim PointPos1, PointPos2 As String
    
    adjTitle = myTitle
    If CorrTest = True Then
        SampCorr = Application.Correl(xRn, yRn)
        If SampCorr <> 1 Then
            TStat = Sqr(yRn.count - 2) * SampCorr / Sqr(1 - SampCorr ^ 2)     '유의확률 계산.
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
        Title:=adjTitle
    
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
    
    'temp1 = (Application.Max(xRn) - Application.Min(xRn)) / 10
    'If temp1 <> 0 Then
    '    PointPos1 = CStrNumPoint(temp1 * 10, xRn.count)
    '    With plot.Chart.Axes(xlCategory)
    '        .MinimumScale = Format(Application.Min(xRn) - temp1, "0.00")
    '        .MaximumScale = Format(Application.Max(xRn) + temp1, "0.00")
    '        .TickLabels.NumberFormat = Format(PointPos1, "0.00")
   '
   '     End With
   ' End If


    
    
    
    temp2 = (Application.Max(yRn) - Application.Min(yRn)) / 10
    If temp2 <> 0 Then
        PointPos2 = CStrNumPoint(temp2 * 10, yRn.count)
        With plot.Chart.Axes(xlValue)
            .MinimumScale = Format(Application.Min(yRn) - temp2, "0.00")
            .MaximumScale = Format(Application.Max(yRn) + temp2, "0.00")
            .TickLabels.NumberFormat = Format(PointPos2, "0.00")
         End With
    End If
    
    With plot.Chart.Axes(xlValue)
        .TickLabels.Font.Size = 8
        .AxisTitle.Orientation = xlUpward
        .AxisTitle.Font.Size = 8
        .TickLabelPosition = xlLow
        .Border.LineStyle = xlNone  '필요시 삭제
    End With
    
    With plot.Chart.Axes(xlCategory)
        .TickLabels.Font.Size = 8
        .AxisTitle.Orientation = xlHorizontal
        .AxisTitle.Font.Size = 8
        .Border.Weight = xlHairline
        '.Border.LineStyle = xlNone                     'y=0 선그리기 위해
        .MajorTickMark = xlNone                         'y=0 선그리기 위해
        .MinorTickMark = xlNone                         'y=0 선그리기 위해
        .TickLabelPosition = xlLow
        .HasMajorGridlines = False
        .TickLabelPosition = xlLow
        .TickLabels.NumberFormatLinked = True           '부분회귀산점도 위해
        '.Crosses = xlCustom
        '.CrossesAt = .MinimumScale
    End With


'    temp1.Chart.SeriesCollection(2).Trendlines.Add Type:=xlLinear
'
'    With temp1.Chart.SeriesCollection(2).Trendlines(1).Border
'        .ColorIndex = 3
'        .Weight = xlThin
'        .LineStyle = xlContinuous
'    End With

End Sub
Sub ScatterPlot1(OutSheetName, Left, Top, Width, Height, _
    xRn1, xRn2, yRn, xname1, rname, xname2, yname, cr, N, h, p, cnt, Optional CorrTest As Boolean = False, _
    Optional myTitle As String = "산점도")
    
    Dim plot As ChartObject: Dim SampCorr, TStat, significance As Double
    Dim adjTitle As String: Dim temp1, temp2, temp3, temp4 As Variant
    Dim PointPos1, PointPos2 As String
    Dim sum As Double
    Dim xdate1, ydate1, dd As Range
    Dim a As Double
    Dim addr, ttemp As Range
    
    
    temp4 = (Application.Max(yRn) - Application.Min(yRn)) / 10
    temp3 = (Application.Max(xRn2) - Application.Min(xRn2)) / 10
    
    adjTitle = myTitle
    
    
    sum = 0
    
    For i = 1 To cr
    ActiveSheet.Rows(1).Cells(1, h).Offset(sum + 1, 0).Select
    Set xdate1 = Range(Selection, Selection.Offset(cnt(i) - 1, 0))
    
    ActiveSheet.Rows(1).Cells(1, p).Offset(sum + 1, 0).Select
    Set ydate1 = Range(Selection, Selection.Offset(cnt(i) - 1, 0))
        
    sum = sum + cnt(i)
    
   
    If i <> 1 Then
    Left = Left + 216
    End If
    Set plot = Worksheets(OutSheetName).ChartObjects.Add(Left, Top, Width, Height)
    plot.Chart.ChartWizard Source:=ydate1, _
        Gallery:=xlXYScatter, Format:=1, HasLegend:=False, _
        CategoryTitle:=xname2, ValueTitle:=yname, _
        Title:=adjTitle & "(Group " & rname(i) & ")"
    
    
        
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
        .XValues = xdate1
        '.MarkerBackgroundColorIndex = 3
        '.MarkerForegroundColorIndex = 3
        .MarkerStyle = xlCircle
        .MarkerSize = 3
    End With
    
     'xlDownward
'xlHorizontal
'xlUpward
'xlVertical
     temp1 = (Application.Max(xdate1) - Application.Min(xdate1)) / 10
    If temp3 <> 0 Then
        PointPos1 = CStrNumPoint(temp1 * 10, xRn2.count)
        With plot.Chart.Axes(xlCategory)
            .MinimumScale = Format(Application.Min(xRn2) - temp3, "0.00")
            .MaximumScale = Format(Application.Max(xRn2) + temp3, "0.00")
            .TickLabels.NumberFormat = Format(PointPos1, "0.00")
        End With
    End If
   
    
    temp2 = (Application.Max(ydate1) - Application.Min(ydate1)) / 10
    If temp4 <> 0 Then
        PointPos2 = CStrNumPoint(temp2 * 10, yRn.count)
        With plot.Chart.Axes(xlValue)
            .MinimumScale = Format(Application.Min(yRn) - temp4, "0.00")
            .MaximumScale = Format(Application.Max(yRn) + temp4, "0.00")
            .TickLabels.NumberFormat = Format(PointPos2, "0.00")
        End With
    End If
   
     With plot.Chart.Axes(xlCategory)
        .TickLabels.Font.Size = 8
        .AxisTitle.Orientation = xlHorizontal
        .AxisTitle.Font.Size = 8
        .Border.Weight = xlHairline
        '.Border.LineStyle = xlNone                     'y=0 선그리기 위해
        .MajorTickMark = xlNone                         'y=0 선그리기 위해
        .MinorTickMark = xlNone                         'y=0 선그리기 위해
        .TickLabelPosition = xlLow
        .HasMajorGridlines = False
        .TickLabelPosition = xlLow
        .TickLabels.NumberFormatLinked = True           '부분회귀산점도 위해.
        '.Crosses = xlCustom
        '.CrossesAt = .MinimumScale
    End With
    
    With plot.Chart.Axes(xlValue)
        .TickLabels.Font.Size = 8
        .AxisTitle.Orientation = xlUpward
        .AxisTitle.Font.Size = 8
        .TickLabelPosition = xlLow
        .Border.LineStyle = xlNone  '필요시 삭제
    End With
    
   plot.Chart.SeriesCollection(1).Trendlines.Add Type:=xlLinear
    
    With plot.Chart.SeriesCollection(1).Trendlines(1).Border
        .ColorIndex = 3
        .Weight = xlThin
        .LineStyle = xlContinuous
    End With
    
    
   Next i
    
End Sub



'''관측순서별 그래프 그리기
Sub OrderScatterPlot(OutSheetName, Left, Top, Width, Height, _
    rn, Rnname, RefLine As Double)
    
    Dim plot As ChartObject
    Dim temp1 As Double
    Dim a, b As Double
    
    Set plot = Worksheets(OutSheetName).ChartObjects.Add(Left, Top, Width, Height)
    plot.Chart.ChartWizard Source:=rn, _
        Gallery:=xlLine, Format:=4, HasLegend:=False, _
        CategoryTitle:="관측순서", ValueTitle:=Rnname, _
        Title:=Rnname & " vs. " & "관측순서"
    
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
    temp1 = (Application.Max(rn) - Application.Min(rn)) / 10
    If temp1 <> 0 Then
        With plot.Chart.Axes(xlValue)
            .MinimumScale = Application.Min(rn) - temp1
            .MaximumScale = Application.Max(rn) + temp1
            .TickLabels.NumberFormat = CStrNumPoint(temp1 * 10, rn.count)
        End With
    End If
    a = plot.Chart.Axes(xlValue).MinimumScale
    b = plot.Chart.Axes(xlValue).MaximumScale
    
    With plot.Chart.SeriesCollection(1)
        .Border.Weight = xlThin
        .Border.LineStyle = xlNone
        .MarkerStyle = xlCircle
        .MarkerSize = 3
    End With
    With plot.Chart.Axes(xlCategory)
        .TickLabels.Font.Size = 8
        .AxisTitle.Orientation = xlHorizontal
        .AxisTitle.Font.Size = 8
        .Border.Weight = xlHairline
        '.Border.LineStyle = xlNone                     'y=0 선그리기 위해
        .MajorTickMark = xlNone                         'y=0 선그리기 위해
        .MinorTickMark = xlNone                         'y=0 선그리기 위해
        .TickLabelPosition = xlLow
        .HasMajorGridlines = False
        .TickLabelPosition = xlLow
        .TickLabels.NumberFormatLinked = True           '부분회귀산점도 위해.
        .Crosses = xlCustom
       ' .CrossesAt = .MinimumScale
    End With
    With plot.Chart.Axes(xlValue)
        .HasMajorGridlines = False
        .TickLabels.Font.Size = 8
        .AxisTitle.Orientation = xlUpward
        .AxisTitle.Font.Size = 8
        .TickLabelPosition = xlLow
        .Border.LineStyle = xlNone  '필요시 삭제
    End With
    
         
        
    If RefLine <> 0 Then
        plot.Chart.SeriesCollection.NewSeries
        With plot.Chart.SeriesCollection(2)
            .XValues = Array(0, rn.count + 1)
            .Values = Array(RefLine, RefLine)
            .ChartType = xlXYScatterSmoothNoMarkers
            .Border.ColorIndex = 3
        End With
        plot.Chart.SeriesCollection.NewSeries
        With plot.Chart.SeriesCollection(3)
            .XValues = Array(0, rn.count + 1)
            .Values = Array(-RefLine, -RefLine)
            .ChartType = xlXYScatterSmoothNoMarkers
            .Border.ColorIndex = 3
        End With
    
        With plot.Chart
            .Axes(xlValue, xlSecondary).MinimumScale = a
            .Axes(xlValue, xlSecondary).MaximumScale = b
            .Axes(xlCategory, xlSecondary).MinimumScale = 0
            .Axes(xlCategory, xlSecondary).MaximumScale = rn.count + 0.5
        End With
    
        With plot.Chart.Axes(xlValue, xlSecondary)
            .MajorTickMark = xlNone
            .MinorTickMark = xlNone
            '.TickLabelPosition = xlNone
        End With
    
        With plot.Chart.Axes(xlCategory, xlSecondary)
            .MajorTickMark = xlNone
            .MinorTickMark = xlNone
            .TickLabelPosition = xlNone
        End With
    End If

End Sub


'''축조절을 위한 함수
'''숫자 자리수만큼의 스트링을 만듬(음수일 경우만)
Function CStrNumPoint(DataWid, DataCount) As String
    
    Dim i As Integer: Dim LogScale As Double
    Dim temp As String
    
    i = 0: temp = "0."
    LogScale = Application.Power(10, _
             Int(Application.Log10(DataWid / DataCount)))
    If LogScale >= 1 Then
        CStrNumPoint = "0"
    Else
        Do
            temp = temp & "0": i = i + 1
            If LogScale = 10 ^ (-i) Then Exit Do
        Loop While (1)
        CStrNumPoint = CStr(temp)
    End If

End Function
