Attribute VB_Name = "ParetoModule"
Sub ParetoShow()

 Dim ErrSignforDataSheet As Integer
 
 ErrSignforDataSheet = InitializeDlg(frmPareto)
 
 Select Case ErrSignforDataSheet
 Case 0: frmPareto.Show
 Case -1
  MsgBox "시트가 보호 상태에 있습니다." & Chr(10) & _
         "데이터를 읽을 수 없습니다.", _
                vbExclamation, "HIST"
         
 Case 1
  MsgBox "시트에 데이터가 있는지 확인하십시요." & Chr(10) & _
         "1행 1열부터 변수명을 입력해야 합니다.", _
                vbExclamation, "HIST"
         
 Case Else
 
  End Select
  
 
End Sub


' 유저폼의 목록상자에 쉬트의 첫행에 있는 변수들을 받아 들인다.
Function InitializeDlg(ParentDlg) As Integer

   Dim myRange As Range
   Dim Cnt As Long
   Dim myArray() As String
   Dim i As Integer
   On Error GoTo ErrorFlag
   
   Set myRange = ActiveSheet.Cells.CurrentRegion
   If myRange.count = 1 Or myRange.Cells(1, 1) = "" Then
        InitializeDlg = 1: Exit Function
   End If
   Set myRange = ActiveSheet.Cells.CurrentRegion.Rows(1)
   ParentDlg.ListBox1.Clear: ParentDlg.ListBox2.Clear
   Cnt = myRange.Cells.count
   
   ReDim myArray(0 To Cnt - 1)
   For i = 1 To Cnt
     myArray(i - 1) = myRange.Cells(i)
   Next i
   ParentDlg.ListBox1.List() = myArray
   InitializeDlg = 0
   Exit Function
   
ErrorFlag:
   InitializeDlg = -1
  
End Function



'숫자형 변수에 문자가 있는지 체크한다.

Function FindingRangeError2(Rn, num) As Boolean
  
  Dim i As Integer
  
  For i = 1 To num
   If IsNumeric(Rn.Cells(i)) = False Then
    FindingRangeError = True
   End If
  Next i
    
End Function

Sub DesignOutPutCell(TargetCell, Direction, myLineStyle, _
    myWeight, myColorIndex)
    
    With TargetCell.Borders(Direction)
        .LineStyle = myLineStyle
        .Weight = myWeight
        .ColorIndex = myColorIndex
    End With

End Sub
Sub MakeTempWorkbook(i)
    
    Dim sin As Integer
    Dim t As Workbook
    
    sin = Application.SheetsInNewWorkbook
    Application.SheetsInNewWorkbook = i
    Set t = Workbooks.Add
    Application.SheetsInNewWorkbook = sin
    ActiveWindow.Visible = True
    TempControlBook = t.Name

End Sub

'
' 메인 알고리즘 ...
'

Sub ParetoResult(OutputSheet, name1, name2, ClassRange, FreqRange, num)

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

    
    DesignOutPutCell ttemp.Resize(, 4), xlEdgeTop, xlContinuous, xlMedium, xlAutomatic
    DesignOutPutCell ttemp.Resize(, 4), xlEdgeBottom, xlContinuous, xlMedium, xlAutomatic
    
    ttemp.Value = name1
    ttemp.Offset(0, 1) = name2
    ttemp.Offset(0, 2) = "누적비율"
    ttemp.Offset(0, 3) = "상대비율"
    
    DesignOutPutCell ttemp.Offset(num + 1, 0).Resize(, 4), xlEdgeTop, xlContinuous, xlMedium, xlAutomatic
    DesignOutPutCell ttemp.Offset(num + 1, 0).Resize(, 4), xlEdgeBottom, xlContinuous, xlMedium, xlAutomatic
    
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
        
End Sub
