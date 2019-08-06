Attribute VB_Name = "Agraph"
Sub makeGraph(myRange, outputsheet As Worksheet)
    Dim ttemp As Range
    Dim addr As Range
    Dim Mychart As ChartObject
    'If IsEmpty(outputsheet.Range("a1")) = True Then
    '   Set ttemp = outputsheet.Cells(2, 1)
    '   Set addr = outputsheet.Range("a1")
    'Else: Set addr = outputsheet.Range("a1")
    '     Set ttemp = outputsheet.Range(addr.Value)
    'End If
    
    Set addr = outputsheet.Range("a1")
    Set ttemp = outputsheet.Range("a" & addr.Value)
    
    yp = ttemp.Offset(0, 1).Top
    xp = ttemp.Offset(0, 1).Left
    Set Mychart = outputsheet.ChartObjects.Add(xp, yp, 188, 191.25)
    Mychart.Chart.ChartType = xlLineMarkers
    Mychart.Chart.SetSourceData Source:=myRange, PlotBy:=xlRows
    Mychart.Chart.Legend.position = xlLegendPositionBottom
    With Mychart.Chart
        .HasTitle = True
        .ChartTitle.Characters.Text = "교호작용도"
        .ChartArea.Font.Size = 9
        .ChartArea.Font.name = "굴림"
        .ChartTitle.Font.Size = 10
        .ChartTitle.Font.Bold = True
        .PlotArea.Border.ColorIndex = 16
    End With
    
    With Mychart.Chart.Axes(xlValue).MajorGridlines.Border
        .Weight = xlHairline
        .LineStyle = xlDashDotDot
    End With
    
    Set ttemp = ttemp.Offset(20, 0)
    '''addr.Value = ttemp.Address
    addr.Value = Right(ttemp.Address, Len(ttemp.Address) - 3)
End Sub
