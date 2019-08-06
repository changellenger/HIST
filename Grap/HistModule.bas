Attribute VB_Name = "HistModule"
Option Explicit

''히스토그램을 그리기 위한 값을 임시로 프린트할 시트
Private TempWorksheet As Worksheet

Sub CountingFreq(s_Input, class, freq, NofClasses, s_MiniUnit)
    
    Dim WidthofItv, mini, maxi, rng As Double
    Dim n As Long: Dim temp() As Double
    Dim PossibleDataN, MaxiNofClass, i, j As Integer
    Dim c As Range
    
    n = s_Input.count: ReDim temp(1 To n): i = 0
    For Each c In s_Input
        i = i + 1
        temp(i) = Format(c.Value, "###0.0000000000")
    Next c
    mini = Application.Min(temp)
    maxi = Application.Max(temp)
    rng = maxi - mini
    PossibleDataN = Int(rng / s_MiniUnit + 1)
    MaxiNofClass = Application.RoundUp(PossibleDataN / NofClasses, 0)
    WidthofItv = s_MiniUnit * MaxiNofClass
    ReDim class(0 To NofClasses + 2)
    class(1) = mini - 0.5 * (NofClasses * MaxiNofClass - PossibleDataN) * s_MiniUnit
    
    If (class(1) / s_MiniUnit) * 10 = Int(class(1) / s_MiniUnit) * 10 Then
        class(1) = class(1) - 0.5 * s_MiniUnit
    End If
    
    class(0) = class(1) - WidthofItv
    For i = 2 To NofClasses + 2
        class(i) = class(i - 1) + WidthofItv
    Next i
    
    ReDim freq(0 To NofClasses + 1)
    For i = 0 To NofClasses + 1: freq(i) = 0: Next i
    For j = 1 To n: For i = 1 To NofClasses
        If temp(j) >= class(i) And temp(j) < class(i + 1) Then freq(i) = freq(i) + 1
    Next i: Next j
    
End Sub

Function FindingMiniUnit(s_Input)
    
    Dim temp, MiniUnit, zeroindi As Double
    Dim c As Range: Dim zerocount As Integer
    
    MiniUnit = 0
    For Each c In s_Input
        zerocount = 0
        Do
            temp = c.Value * (10 ^ zerocount)
            zeroindi = temp - Fix(temp)
            zerocount = zerocount + 1
        Loop Until (zeroindi = 0 Or zerocount = 11)
        MiniUnit = Application.Max(MiniUnit, zerocount - 1)
    Next c
    
    FindingMiniUnit = MiniUnit
        
End Function

Function FindingNofClasses(Obs) As Integer

    If 1 <= Obs < 100 Then
        FindingNofClasses = -Int(-Sqr(Obs))
    ElseIf 100 <= Obs <= 400 Then
        FindingNofClasses = Int(Sqr(Obs))
    ElseIf Obs > 400 Then
        FindingNofClasses = 20
    Else: FindingNofClasses = 0
    End If
    
End Function

Function MainHistogram(Rn, xPos, yPos, OutputSheet, _
         Optional CustomN As Integer = 0, Optional VarName As String = "") As String

    Dim class(), freq() As Single
    Dim s_Input, s_source0, s_source1 As Range
    Dim temp1 As ChartObject: Dim i, prtindex, tmpCount As Integer
    Dim TempVarName As String
    
    If CustomN = 0 Then
        tmpCount = FindingNofClasses(Rn.count)
    Else: tmpCount = CustomN
    End If
    
    CountingFreq Rn, class, freq, tmpCount, 10 ^ (-FindingMiniUnit(Rn))
    
    PublicModule.OpenTempWorkSheet TempWorksheet, "_TempHistogram_"
    prtindex = TempWorksheet.Cells(1, 1).Value
    TempWorksheet.Cells(prtindex + 1, 1).Value = "하한"
    TempWorksheet.Cells(prtindex + 1, 2).Value = "상한"
    TempWorksheet.Cells(prtindex + 1, 3).Value = "중앙값"
    TempWorksheet.Cells(prtindex + 1, 4).Value = "도수"

    For i = 0 To UBound(freq)
        TempWorksheet.Cells(prtindex + i + 2, 1).Value = class(i)
        TempWorksheet.Cells(prtindex + i + 2, 2).Value = class(i + 1)
        TempWorksheet.Cells(prtindex + i + 2, 3).Value = (class(i) + class(i + 1)) / 2
        TempWorksheet.Cells(prtindex + i + 2, 4).Value = freq(i)
    Next i
    TempWorksheet.Cells(1, 1).Value = prtindex + UBound(freq) + 2
    Set s_source1 = Range(TempWorksheet.Cells(prtindex + 2, 4), _
                          TempWorksheet.Cells(prtindex + UBound(freq) + 2, 4))
    Set s_source0 = Range(TempWorksheet.Cells(prtindex + 2, 3), _
                          TempWorksheet.Cells(prtindex + UBound(freq) + 2, 3))
    Set temp1 = OutputSheet.ChartObjects.Add(xPos + 45, yPos + 30, 200, 200)
    
    If VarName = "" Then
        TempVarName = "히스토그램"
    Else: TempVarName = "히스토그램" & ": " & VarName
    End If
    
    temp1.Chart.ChartWizard Source:=s_source1, _
        Gallery:=xlColumn, Format:=8, _
        HasLegend:=False, title:=TempVarName, _
        CategoryTitle:="계급값", ValueTitle:="도수"
    
    temp1.Chart.SeriesCollection(1).XValues = s_source0
    
    temp1.Chart.Axes(xlCategory).TickLabels.Orientation = xlHorizontal
    temp1.Chart.Axes(xlCategory).TickLabels.NumberFormat = "0.00"
    temp1.Chart.Axes(xlValue).AxisTitle.Orientation = xlVertical
    temp1.Chart.Axes(xlValue).MajorTickMark = xlOutside
    temp1.Chart.ChartArea.Font.Size = 9
    temp1.Chart.PlotArea.Border.LineStyle = xlNone
    temp1.Chart.PlotArea.Interior.ColorIndex = xlNone
    With temp1.Chart.ChartTitle
        .Font.Size = 10
        .Font.Bold = True
        
    End With
    temp1.Chart.SeriesCollection(1).Border.Color = RGB(70, 70, 255)
  '  temp1.Chart.SeriesCollection(1).Trendlines.Add                                  ' 추세선 추가
    'temp1.Chart.SeriesCollection(1).Trendlines.Add.DisplayEquation = True
    'temp1.Chart.SeriesCollection(1).Trendlines.Add.DisplayRSquared = True          'R
    
    MainHistogram = temp1.Name
    
End Function
