Attribute VB_Name = "BoxPlotModule"
Option Explicit

''상자그림을 그리기 위한 값을 임시로 프린트할 시트 만들기
Private TempWorksheet As Worksheet

''상자그림을 그리기 위한 값 구하기
Sub FindingBoxPlotValue(s_Input, xOutliers, mOutliers, adj, q, OutliersCheck)
    Dim N, xcount, mcount As Long: Dim tmpindex As Boolean
    Dim mini, maxi, IQR, UnitScale As Single
    Dim LIfence, LOfence, UIfence, UOfence As Single
    Dim c As Range
    
    xcount = 0: mcount = 0: N = s_Input.count
    mini = Application.Min(s_Input): maxi = Application.Max(s_Input)
    q(1) = Application.Quartile(s_Input, 1)
    q(2) = Application.Quartile(s_Input, 2)
    q(3) = Application.Quartile(s_Input, 3)
    IQR = q(3) - q(1): UnitScale = 1.5 * IQR
    LIfence = q(1) - UnitScale: LOfence = LIfence - UnitScale
    UIfence = q(3) + UnitScale: UOfence = UIfence + UnitScale
    
    tmpindex = False
    For Each c In s_Input
        If c.Value < LOfence Then
            xcount = xcount + 1: ReDim Preserve xOutliers(1 To xcount)
            xOutliers(xcount) = c.Value
        ElseIf c.Value < LIfence Then
            mcount = mcount + 1: ReDim Preserve mOutliers(1 To mcount)
            mOutliers(mcount) = c.Value
        ElseIf c.Value <= UIfence Then
            If tmpindex = False Then
                adj(1) = c.Value
                adj(2) = c.Value
                tmpindex = True
            Else
                adj(1) = Application.Min(adj(1), c.Value)
                adj(2) = Application.Max(adj(2), c.Value)
            End If
        ElseIf c.Value <= UOfence Then
            mcount = mcount + 1: ReDim Preserve mOutliers(1 To mcount)
            mOutliers(mcount) = c.Value
        Else
            xcount = xcount + 1: ReDim Preserve xOutliers(1 To xcount)
            xOutliers(xcount) = c.Value
        End If
    Next c
    
    OutliersCheck(1) = True: OutliersCheck(2) = True
    If xcount = 0 Then OutliersCheck(1) = False
    If mcount = 0 Then OutliersCheck(2) = False
    
End Sub

''상자그림을 그리기 위한 값을 숨겨진 임시시트에 프린트
Sub PrintingBoxPlotValue(xOutliers, mOutliers, adj, q, OutliersCheck, xValueOption)
    
    Static xValue As Integer
    Dim prtindex, i As Long
    
'    xvalueoption 값이 거짓이면 x축을 처음부터 시작하고 _
'    임시시트를 한개 만들든지 이전 것을 사용한다. _
'    값이 참인 경우는 두개 이상의 상자그림을 그릴 경우이다.
    If xValueOption = False Then
        xValue = 0
        ModuleControl.openTempWorkSheet TempWorksheet, "_TempBoxplot_"
    End If
    
    prtindex = TempWorksheet.Cells(1, 1)
    
    If OutliersCheck(1) = True Then
        For i = 1 To UBound(xOutliers)
            TempWorksheet.Cells(prtindex + 2, i) = xOutliers(i)
            TempWorksheet.Cells(prtindex + 1, i) = 2 + 3 * Int(xValue / 14)
        Next i
    End If
    
    If OutliersCheck(2) = True Then
        For i = 1 To UBound(mOutliers)
            TempWorksheet.Cells(prtindex + 4, i) = mOutliers(i)
            TempWorksheet.Cells(prtindex + 3, i) = 2 + 3 * Int(xValue / 14)
        Next i
    End If
    
    With TempWorksheet
        .Cells(prtindex + 6, 1) = adj(1)
        .Cells(prtindex + 6, 2) = q(1)
        .Cells(prtindex + 5, 1) = 2 + 3 * Int(xValue / 14)
        .Cells(prtindex + 5, 2) = 2 + 3 * Int(xValue / 14)
        .Cells(prtindex + 8, 1) = q(3)
        .Cells(prtindex + 8, 2) = adj(2)
        .Cells(prtindex + 7, 1) = 2 + 3 * Int(xValue / 14)
        .Cells(prtindex + 7, 2) = 2 + 3 * Int(xValue / 14)
        
        .Cells(prtindex + 9, 1) = 2 + 3 * Int(xValue / 14)
        .Cells(prtindex + 9, 2) = 3 + 3 * Int(xValue / 14)
        .Cells(prtindex + 9, 3) = 3 + 3 * Int(xValue / 14)
        .Cells(prtindex + 9, 4) = 1 + 3 * Int(xValue / 14)
        .Cells(prtindex + 9, 5) = 1 + 3 * Int(xValue / 14)
        .Cells(prtindex + 9, 6) = 2 + 3 * Int(xValue / 14)
        
        .Cells(prtindex + 10, 1) = q(1)
        .Cells(prtindex + 10, 2) = q(1)
        .Cells(prtindex + 10, 3) = q(3)
        .Cells(prtindex + 10, 4) = q(3)
        .Cells(prtindex + 10, 5) = q(1)
        .Cells(prtindex + 10, 6) = q(1)
        
        .Cells(prtindex + 11, 1) = 1 + 3 * Int(xValue / 14)
        .Cells(prtindex + 11, 2) = 3 + 3 * Int(xValue / 14)
        .Cells(prtindex + 12, 1) = q(2)
        .Cells(prtindex + 12, 2) = q(2)

        .Cells(prtindex + 13, 1) = 2 + 3 * Int(xValue / 14)
        .Cells(prtindex + 13, 2) = 2 + 3 * Int(xValue / 14)
        .Cells(prtindex + 14, 1) = adj(1)
        .Cells(prtindex + 14, 2) = adj(2)
    End With
    
    TempWorksheet.Cells(1, 1) = TempWorksheet.Cells(1, 1) + 14
    xValue = xValue + 14
    
End Sub

''상자그림 그리기
''분산형 차트를 상자그림으로 변화시키기
Function DrawingBoxplot(xOutliers, mOutliers, OutliersCheck, s_max, s_min, outputsheet, ChartAdd, xp, yp) As String
    
    Static co As ChartObject
    Dim i, AddingSeries, NowSeriesCount, prtindex As Integer
    Dim tempRec, ir1, ir2 As String
    
    prtindex = TempWorksheet.Cells(1, 1) - 14
    If ChartAdd = False Then
        Set co = outputsheet.ChartObjects.Add(xp, yp, 200, 200)
        co.Chart.ChartWizard Source:=TempWorksheet.Range("A11"), _
                Gallery:=xlXYScatter, Format:=1, _
                HasLegend:=1, Title:="수준별 상자그림(Box Plot)"
        ChartAdd = True
    End If
        
    NowSeriesCount = co.Chart.SeriesCollection.count
    If NowSeriesCount = 1 Then
        NowSeriesCount = 0
    Else: co.Chart.SeriesCollection.NewSeries
    End If
    
    AddingSeries = -1
    If OutliersCheck(1) = True Then
        ir1 = TempWorksheet.Cells(1 + prtindex, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        ir2 = TempWorksheet.Cells(1 + prtindex, UBound(xOutliers)).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        co.Chart.SeriesCollection(1 + NowSeriesCount).XValues = TempWorksheet.Range(ir1 & ":" & ir2)
    
        ir1 = TempWorksheet.Cells(2 + prtindex, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        ir2 = TempWorksheet.Cells(2 + prtindex, UBound(xOutliers)).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        co.Chart.SeriesCollection(1 + NowSeriesCount).Values = TempWorksheet.Range(ir1 & ":" & ir2)
        AddingSeries = 0
    End If
    
    If OutliersCheck(2) = True Then
        If AddingSeries = 0 Then co.Chart.SeriesCollection.NewSeries
        ir1 = TempWorksheet.Cells(3 + prtindex, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        ir2 = TempWorksheet.Cells(3 + prtindex, UBound(mOutliers)).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        co.Chart.SeriesCollection(2 + AddingSeries + NowSeriesCount).XValues = TempWorksheet.Range(ir1 & ":" & ir2)
        
        ir1 = TempWorksheet.Cells(4 + prtindex, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        ir2 = TempWorksheet.Cells(4 + prtindex, UBound(mOutliers)).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        co.Chart.SeriesCollection(2 + AddingSeries + NowSeriesCount).Values = TempWorksheet.Range(ir1 & ":" & ir2)
        AddingSeries = 0
    End If
    
    For i = 1 To 5 + AddingSeries
        co.Chart.SeriesCollection.NewSeries
    Next i
    
    If OutliersCheck(1) = False Then
        If OutliersCheck(2) = False Then
            AddingSeries = -2
            co.Chart.SeriesCollection(5 + NowSeriesCount).name = "=""인접값"""
        Else
            AddingSeries = -1
            co.Chart.SeriesCollection(1 + NowSeriesCount).name = "=""보통이상점"""
            co.Chart.SeriesCollection(6 + NowSeriesCount).name = "=""인접값"""
        End If
    Else
        If OutliersCheck(2) = False Then
            AddingSeries = -1
            co.Chart.SeriesCollection(1 + NowSeriesCount).name = "=""극단이상점"""
            co.Chart.SeriesCollection(6 + NowSeriesCount).name = "=""인접값"""
        Else
            AddingSeries = 0
            co.Chart.SeriesCollection(1 + NowSeriesCount).name = "=""극단이상점"""
            co.Chart.SeriesCollection(2 + NowSeriesCount).name = "=""보통이상점"""
            co.Chart.SeriesCollection(7 + NowSeriesCount).name = "=""인접값"""
        End If
    End If
  
    For i = 0 To 4
        If i <> 2 Then
            ir1 = TempWorksheet.Cells(5 + i * 2 + prtindex, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
            ir2 = TempWorksheet.Cells(5 + i * 2 + prtindex, 2).Address(RowAbsolute:=False, ColumnAbsolute:=False)
            co.Chart.SeriesCollection(3 + i + AddingSeries + NowSeriesCount).XValues = TempWorksheet.Range(ir1 & ":" & ir2)
            
            ir1 = TempWorksheet.Cells(6 + i * 2 + prtindex, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
            ir2 = TempWorksheet.Cells(6 + i * 2 + prtindex, 2).Address(RowAbsolute:=False, ColumnAbsolute:=False)
            co.Chart.SeriesCollection(3 + i + AddingSeries + NowSeriesCount).Values = TempWorksheet.Range(ir1 & ":" & ir2)
        End If
    Next i
    
    ir1 = TempWorksheet.Cells(9 + prtindex, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    ir2 = TempWorksheet.Cells(9 + prtindex, 6).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    co.Chart.SeriesCollection(5 + AddingSeries + NowSeriesCount).XValues = TempWorksheet.Range(ir1 & ":" & ir2)
            
    ir1 = TempWorksheet.Cells(10 + prtindex, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    ir2 = TempWorksheet.Cells(10 + prtindex, 6).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    co.Chart.SeriesCollection(5 + AddingSeries + NowSeriesCount).Values = TempWorksheet.Range(ir1 & ":" & ir2)
  
    co.Chart.Legend.position = xlBottom
    co.Chart.PlotArea.Interior.ColorIndex = 2
    
    co.Chart.Axes(xlValue).MinimumScale = s_min - 0.1 * Abs(s_max - s_min)
    co.Chart.Axes(xlValue).MaximumScale = s_max + 0.1 * Abs(s_max - s_min)
    co.Chart.Axes(xlValue).TickLabels.NumberFormat = "0.00"
    co.Chart.Axes(xlValue, xlPrimary).HasTitle = False
    'co.Chart.Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = ""
    
    co.Chart.PlotArea.Border.ColorIndex = 16
    If OutliersCheck(1) = True Then
        With co.Chart.SeriesCollection(1 + NowSeriesCount)
            .Border.LineStyle = xlNone
            .MarkerBackgroundColorIndex = 3
            .MarkerForegroundColorIndex = 1
            .MarkerStyle = xlCircle
            .ApplyDataLabels Type:=xlDataLabelsShowValue
        End With
    End If

    If OutliersCheck(2) = True Then
        With co.Chart.SeriesCollection(2 + AddingSeries + NowSeriesCount)
            .Border.LineStyle = xlNone
            .MarkerBackgroundColorIndex = 6
            .MarkerForegroundColorIndex = 1
            .MarkerStyle = xlCircle
            .ApplyDataLabels Type:=xlDataLabelsShowValue
        End With
    End If
    
    
    For i = 3 To 6
        With co.Chart.SeriesCollection(i + AddingSeries + NowSeriesCount)
            .Border.ColorIndex = 55
            .Border.Weight = xlThin
            .Border.LineStyle = xlContinuous
            .MarkerStyle = xlNone
            .Smooth = False
        End With
    Next i
    
    With co.Chart.SeriesCollection(7 + AddingSeries + NowSeriesCount)
        .Border.LineStyle = xlNone
        .MarkerBackgroundColorIndex = xlNone
        .MarkerForegroundColorIndex = 55
        .MarkerStyle = xlX
    End With
    DrawingBoxplot = co.name
    
    'co.Chart.CopyPicture Appearance:=xlScreen, _
                            Size:=xlScreen, _
                            Format:=xlPicture
    'ActiveWindow.Visible = False
    'co.Chart.Delete
    'ActiveSheet.Pictures.Paste
    
End Function

''변수이름과 범례를 상자그림에 맞게 고치기
''극단이상점, 보통이상점에 대한 범례 나타내기
Sub AdjustingChart(outputsheet, coName, ValueNum, VarName)
    
    Dim index As Integer: Dim s As Series
    Dim flag1, flag2 As Boolean: Dim xValueStr As String
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    xValueStr = VarName(1)
    If ValueNum > 1 Then
       For i = 2 To ValueNum
          xValueStr = xValueStr & " , " & VarName(i)
       Next i
    End If
    With outputsheet.ChartObjects(coName).Chart
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = xValueStr
    End With
    
    For Each s In outputsheet.ChartObjects(coName).Chart.SeriesCollection
        index = index + 1
        If s.name = "극단이상점" Then
            If flag1 = False Then
                flag1 = True
            Else
                outputsheet.ChartObjects(coName).Chart.Legend.LegendEntries(index).Delete
                index = index - 1
            End If
        ElseIf s.name = "보통이상점" Then
            If flag2 = False Then
                flag2 = True
            Else
                outputsheet.ChartObjects(coName).Chart.Legend.LegendEntries(index).Delete
                index = index - 1
            End If

        Else
            outputsheet.ChartObjects(coName).Chart.Legend.LegendEntries(index).Delete
            index = index - 1
        End If
    Next s
    Exit Sub
    
ErrorHandler:
    outputsheet.ChartObjects(coName).Chart.HasLegend = False
    Exit Sub
    
End Sub

''상자그림을 그리기 위한 Main 프로시저
''인수 설명:
'' ArrayofRange : 변수영역을 담고 있는 배열(하한 1부터 시작)
'' CountofRange : 변수의 개수 즉, 배열 크기
'' xPos : 상자그림을 그릴 x 좌표
'' yPos : 상자그림을 그릴 y 좌표
'' outputsheet : 상자그림을 그릴 시트 개체
Sub MainBoxPlot(ArrayofRn, CountofRange, xPos, yPos, _
    outputsheet, VarName, Optional IsArray As Boolean = True)
    
    Dim xOutliers(), mOutliers(), adj(1 To 2), q(1 To 3) As Single
    Dim OutliersCheck(1 To 2), xValueOption, ChartAdd As Boolean
    Dim yAxisMin, yAxisMax As Single: Dim tmpRange(1 To 1) As Range
    Dim TempCoName, SelectionString As String: Dim i As Integer
    Dim AnySelection, ArrayofRange As Variant
    
    If IsArray = False Then
        Set tmpRange(1) = ArrayofRn
        ArrayofRange = tmpRange
    Else: ArrayofRange = ArrayofRn
    End If
    
    '선택된 차트가 있을 경우 에러가 발생한다. 이상하다.
    Set AnySelection = Selection: SelectionString = TypeName(AnySelection)
    If SelectionString = "ChartArea" Then ActiveWindow.Visible = False

    ChartAdd = False: xValueOption = False
        
    yAxisMax = Application.Max(ArrayofRange(1)): yAxisMin = Application.Min(ArrayofRange(1))

    For i = 1 To CountofRange
        yAxisMax = Application.Max(ArrayofRange(i), yAxisMax): yAxisMin = Application.Min(ArrayofRange(i), yAxisMin)
        FindingBoxPlotValue ArrayofRange(i), xOutliers, mOutliers, adj, q, OutliersCheck
        PrintingBoxPlotValue xOutliers, mOutliers, adj, q, OutliersCheck, xValueOption
        TempCoName = DrawingBoxplot(xOutliers, mOutliers, OutliersCheck, yAxisMax, yAxisMin, outputsheet, ChartAdd, xPos, yPos)
        Erase xOutliers, mOutliers
        ChartAdd = True: xValueOption = True
    Next i
    
    outputsheet.ChartObjects(TempCoName).Chart.Axes(xlCategory).MinimumScale = 0
    outputsheet.ChartObjects(TempCoName).Chart.Axes(xlCategory).MaximumScale = 3 * CountofRange + 1
    outputsheet.ChartObjects(TempCoName).Chart.ChartArea.Font.Size = 9
    outputsheet.ChartObjects(TempCoName).Chart.HasAxis(xlCategory, xlPrimary) = False
    With outputsheet.ChartObjects(TempCoName).Chart.ChartTitle
            .Font.Size = 10
            .Font.Bold = True
    End With

    AdjustingChart outputsheet, TempCoName, CountofRange, VarName
 
End Sub
