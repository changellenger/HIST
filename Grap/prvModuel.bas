Attribute VB_Name = "prvModuel"

Function ParetoResultprv(OutputSheet, name1, name2, ClassRange, FreqRange, num) As String


    Dim sumvalue As Double
    Dim myRange As Range, SortingRange As Range
    Dim i As Integer
    Dim j As Integer
    Dim a() As Double
    Dim b As Double
    Dim ParetoChart As ChartObject
    Dim Addr, ttemp As Range
    Dim posi(0 To 1) As Long
    
    If IsEmpty(OutputSheet.Range("a1")) = True Then
        Set ttemp = OutputSheet.Cells(2, 1)
        Set Addr = OutputSheet.Range("a1")

    Else: Set Addr = OutputSheet.Range("a1")
         Set ttemp = OutputSheet.Cells(Addr.Value, 1)
    End If
    
    PublicModule.Title1 OutputSheet, "분석결과"
    Set ttemp = ttemp.Offset(5, 1)

    
    ParetoModule.DesignOutPutCell ttemp.Resize(, 4), xlEdgeTop, xlContinuous, xlMedium, xlAutomatic
    ParetoModule.DesignOutPutCell ttemp.Resize(, 4), xlEdgeBottom, xlContinuous, xlMedium, xlAutomatic
    
    ttemp.Value = name1
    ttemp.Offset(0, 1) = name2
    ttemp.Offset(0, 2) = "누적비율"
    ttemp.Offset(0, 3) = "상대비율"
    
     ParetoModule.DesignOutPutCell ttemp.Offset(num + 1, 0).Resize(, 4), xlEdgeTop, xlContinuous, xlMedium, xlAutomatic
     ParetoModule.DesignOutPutCell ttemp.Offset(num + 1, 0).Resize(, 4), xlEdgeBottom, xlContinuous, xlMedium, xlAutomatic
    
    ttemp.Offset(num + 1, 0) = "합"
    
    
    For i = 1 To num
        ttemp.Offset(i, 0) = ClassRange.Cells(i)
        ttemp.Offset(i, 1) = FreqRange.Cells(i)
    Next i

    Set myRange = OutputSheet.Range(ttemp, ttemp.Offset(num + 1, 3))
    myRange.Sort key1:=myRange.Cells(2, 2), order1:=xlDescending
    sumvalue = Application.Sum(FreqRange)
    myRange.Cells(num + 2, 2) = sumvalue

    For j = 1 To num
        myRange.Cells(j + 1, 4) = myRange.Cells(j + 1, 2) / sumvalue
    Next j

    ReDim a(num - 1)
    
    For j = 1 To num
        a(j - 1) = myRange.Cells(j + 1, 4)
    Next j


    b = 0
    For j = 1 To num
        b = b + a(j - 1)
        myRange.Cells(j + 1, 3) = b
    Next j

    
    Set ttemp = ttemp.Offset(num + 3)
    Set ParetoChart = OutputSheet.ChartObjects.Add(ttemp.Left, ttemp.Top, 400, 250)
    ParetoChart.Chart.ChartWizard Source:=OutputSheet.Range(myRange.Cells(1, 1), myRange.Cells(num + 1, 2)), _
        PlotBy:=xlColumns, title:="파레토그림(Pareto Chart)"
    
    With ParetoChart.Chart
        .SeriesCollection(1).XValues = OutputSheet.Range(myRange.Cells(2, 1), myRange.Cells(num + 1, 1))
        .SeriesCollection(1).Values = OutputSheet.Range(myRange.Cells(2, 2), myRange.Cells(num + 1, 2))
        .SeriesCollection(1).Name = name1
        .SeriesCollection(1).Border.ColorIndex = 16
        .SeriesCollection.NewSeries
        .PlotArea.Border.ColorIndex = 16
        .ChartTitle.Font.Size = 11
        .HasLegend = True
        .Legend.position = xlLegendPositionBottom
        .Legend.Font.Size = 8
    End With
    With ParetoChart.Chart.SeriesCollection(2)
        .Values = OutputSheet.Range(myRange.Cells(2, 3), myRange.Cells(num + 1, 3))
        .Name = name2
        .ChartType = xlLine
        .Border.ColorIndex = 25
        .Border.Weight = xlThin
        .MarkerBackgroundColorIndex = 25
        .MarkerForegroundColorIndex = 25
        .MarkerStyle = xlCircle
        .MarkerSize = 5
        .AxisGroup = 2
    End With
    
    
    With ParetoChart.Chart
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = name1
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = name2
        .Axes(xlValue, xlSecondary).HasTitle = True
        .Axes(xlValue, xlSecondary).AxisTitle.Characters.Text = "백분율"
        .Axes(xlCategory, xlSecondary).HasTitle = False
        .Axes(xlCategory).TickLabels.Font.Size = 8
        .Axes(xlValue).AxisTitle.Font.Size = 9
        .Axes(xlValue).AxisTitle.Orientation = xlVertical
        .Axes(xlCategory).AxisTitle.Font.Size = 9
        .Axes(xlValue, xlSecondary).AxisTitle.Font.Size = 9
        .Axes(xlValue, xlSecondary).AxisTitle.Orientation = xlVertical
    End With
    
    With ParetoChart.Chart.Axes(xlValue)
        .HasMajorGridlines = True
        .HasMinorGridlines = False
        .MinimumScale = 0
        .MaximumScale = sumvalue
        .MajorUnitIsAuto = True
        .ScaleType = xlLinear
    End With

    With ParetoChart.Chart.ChartGroups(1)
        .Overlap = 0
        .GapWidth = 0
    End With
       
    
    With ParetoChart.Chart.Axes(xlValue, xlSecondary)
        .MinimumScale = 0
        .MaximumScale = 1
        .MajorUnit = 0.2
        .ScaleType = xlLinear
        .TickLabels.NumberFormatLocal = "0%"
    End With

    With ParetoChart.Chart.Axes(xlValue).MajorGridlines.Border
        .ColorIndex = 24
        .Weight = xlHairline
        .LineStyle = xlDashDot
    End With
        
    Range(myRange.Cells(2, 3), myRange.Cells(num + 1, 4)).NumberFormatLocal = "0.0000_ "
    Set ttemp = ttemp.Offset(24, -1)
    Addr.Value = ttemp.Cells.row
    
    ParetoResultprv = ParetoChart.Name

        
End Function
