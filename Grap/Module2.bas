Attribute VB_Name = "Module2"
Sub MakeTable(rngTable As Range)
'   Making a table

    Dim strTable As String
    Dim Asht As Worksheet
    
    Set Asht = ActiveSheet
    strTable = "GRRAtable" & Asht.ListObjects.count + 1
    Asht.ListObjects.Add(xlSrcRange, rngTable, , xlYes).Name = strTable
    Asht.ListObjects(strTable).TableStyle = "TableStyleLight20"
    Asht.ListObjects(strTable).ShowAutoFilter = False
    Asht.ListObjects(strTable).Unlist
    rngTable.ColumnWidth = 8.38
    
End Sub
Sub Contour_Plot(rngIn As Range, rngCht As Range)

    Dim cht As Chart
    
    strName = ActiveSheet.Name
    
    Set cht = Charts.Add
    Set cht = cht.Location(where:=xlLocationAsObject, Name:=strName)
    With cht
        .SetSourceData Source:=rngIn
        .ChartType = xlSurfaceTopView  'xlSurfaceTopViewWireframe
    End With

    With cht
        .HasLegend = True
        .HasTitle = True
        With .ChartTitle
            .AutoScaleFont = False
            .Characters.Text = "등고선도(Contour Plot)"
            .Characters.Font.Size = 11
        End With
        
    End With
    
    Call chtLocation(cht, rngCht)
    
End Sub


Sub Surface_Plot(rngIn As Range, rngCht As Range)
'Sub Surface_Plot(x As Range, y As Range, z As Range, rngCht As Range)
    Dim cht As Chart
    
    strName = ActiveSheet.Name
    
    Set cht = Charts.Add
    Set cht = cht.Location(where:=xlLocationAsObject, Name:=strName)
    With cht
        .SetSourceData Source:=rngIn
        .ChartType = xlSurface
    End With

    With cht
        .HasLegend = True
        .HasTitle = True
        With .ChartTitle
            .AutoScaleFont = False
            .Characters.Text = "표면도(Surface Plot)"
            .Characters.Font.Size = 11
        End With
        
    End With
    
    'Charts(1).SetSourceData Source:=Sheets(1).Range("a1:a10"), PlotBy:=xlColumns
    
    With cht.Axes(xlValue)
        .CrossesAt = .MinimumScale
    End With
    
    Call chtLocation(cht, rngCht)
    
End Sub
Sub chtLocation(cht As Chart, rng As Range)

    With cht.Parent
        .Left = rng.Left
        .Top = rng.Top
        .Width = rng.Width
        .Height = rng.Height
    End With

End Sub


Sub BarchartGrap(OutSheetName, Left, Top, Width, Height, _
    xRn, yRn, xname, yname, Optional CorrTest As Boolean = False, _
    Optional myTitle As String = "산점도(Scatter Plot)")
    
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

End Sub
