Attribute VB_Name = "QQmodule"
Option Explicit

''정규확률그림을 그리기 위한 값을 임시로 프린트할 시트 만들기
Private TempWorksheet As Worksheet

Function MainNormPlot(rn, xp, yp, outputsheet, Optional VarName As String = "", Optional NTest As Boolean = False) As String
    
    Dim c, DSource, DSource0, ASource, ac As Range
    Dim i, Obs As Long: Dim prtindex As Long
    Dim qq_pdata(), qq_qdata() As Double
    Dim mean, stan As Double: Dim tmpstring As String
    Dim temp1 As ChartObject
    Dim significance As Double
    PublicModule.openTempWorkSheet TempWorksheet, "_TempQQPlot_"
    prtindex = TempWorksheet.Cells(1, 1).Value
    On Error Resume Next
    Obs = rn.count
    
    ReDim qq_pdata(1 To Obs): ReDim qq_qdata(1 To Obs)
     
     
     '분위수 구하기
     For i = 1 To Obs
        qq_qdata(i) = Application.NormSInv((i - 3 / 8) / (Obs + 1 / 4))
     Next i
     
     i = 1
     mean = Application.Average(rn): stan = Application.StDev(rn)
     For Each c In rn
            qq_pdata(i) = (c.Value - mean) / stan
            i = i + 1
     Next c
        
     StemModule.procSort1D qq_pdata, 1, Obs
    
     'QQ_Plot 그리기

     For i = 1 To Obs
        TempWorksheet.Cells(i, prtindex + 1).Value = qq_qdata(i)
        TempWorksheet.Cells(i, prtindex + 2).Value = qq_pdata(i)
     Next i
     
     Set ASource = Range(TempWorksheet.Cells(1, prtindex + 1), TempWorksheet.Cells(Obs, prtindex + 2))
     Set DSource = Range(TempWorksheet.Cells(1, prtindex + 2), TempWorksheet.Cells(Obs, prtindex + 2))
     Set DSource0 = Range(TempWorksheet.Cells(1, prtindex + 1), TempWorksheet.Cells(Obs, prtindex + 1))
     TempWorksheet.Cells(1, 1) = TempWorksheet.Cells(1, 1) + 2
    
    Set temp1 = outputsheet.ChartObjects.Add(xp + 30, yp + 40, 200, 200)
    If VarName <> "" Then
        tmpstring = "정규확률그림" & ": " & VarName
    Else
        tmpstring = "정규확률그림"
    End If
    
    '첨가된 부분
    If NTest = True Then
        significance = kolmo(rn, Obs)
        tmpstring = tmpstring & Chr(10) & "정규성검정" & " 유의확률=" & _
              Format(significance, "0.0000")
    End If
    ' 첨가된 부분 끝
    
    temp1.Chart.ChartWizard Source:=ASource, _
        Gallery:=xlXYScatter, Format:=1, _
        HasLegend:=False, title:=tmpstring, _
        CategoryTitle:="정규점수", ValueTitle:="표준화된값"
    temp1.Chart.Axes(xlCategory).Name = VarName
    temp1.Chart.Axes(xlValue).AxisTitle.Orientation = xlVertical
    temp1.Chart.SeriesCollection(1).XValues = DSource0
    temp1.Chart.SeriesCollection(1).Values = DSource
    temp1.Chart.SeriesCollection(2).XValues = DSource0
    temp1.Chart.SeriesCollection(2).Values = DSource0
    temp1.Chart.ChartArea.Font.Size = 9
   
    'temp1.Chart.PlotArea.Border.ColorIndex = 16

    With temp1.Chart.SeriesCollection(1)
        .MarkerStyle = xlCircle
        .MarkerSize = 5
    End With
    
    With temp1.Chart.SeriesCollection(2)
        .Border.LineStyle = xlNone
        .MarkerBackgroundColorIndex = xlAutomatic
        .MarkerForegroundColorIndex = xlAutomatic
        .MarkerStyle = xlNone
        .Smooth = True
        .MarkerSize = 5
        .Shadow = False
    End With
    
    temp1.Chart.SeriesCollection(2).Trendlines.Add Type:=xlLinear
    
    With temp1.Chart.SeriesCollection(2).Trendlines(1).Border
        .ColorIndex = 3
        .Weight = xlThin
        .LineStyle = xlContinuous
    End With
        Erase qq_pdata, qq_qdata
    With temp1.Chart.ChartTitle
        .Font.Size = 10
        .Font.Bold = True
        .Characters(Start:=7, Length:=33).Font.Bold = False
        .Characters(Start:=7, Length:=33).Font.Size = 9
    End With

    MainNormPlot = temp1.Name
    
End Function

Function kolmo(data, cnt)
  Dim X() As Double
  Dim i, J, nm1, L, xn, il, im As Integer
  Dim temp As Double
  Dim PI As Double
  Dim fi, fs, u, s, dn, ei, es, z, t, d, y, zs As Double
  PI = 3.1415926
  u = Application.Average(data)
  s = Application.StDev(data)
  ReDim X(1 To cnt)
  For i = 1 To cnt
     X(i) = data.Cells(i).Value
  Next i
  ' Bubble Sorting
  
  For i = 1 To cnt - 1
      For J = cnt To i + 1 Step -1
        If X(J - 1) > X(J) Then
           temp = X(J - 1)
           X(J - 1) = X(J)
           X(J) = temp
        End If
      Next J
  Next i
' compute Dn=max|fn(x)-f(x)|
  
  nm1 = cnt - 1
  xn = cnt
  dn = 0
  fs = 0
  il = 1
6:
  For i = il To nm1
      J = i
      If X(J) < X(J + 1) Then
         GoTo 9
     End If
  Next i
8:
  J = cnt
9:
  il = J + 1
  fi = fs
  fs = J / xn
  z = (X(J) - u) / s
  t = 1 / (1 + 0.2316419 * Abs(z))
  d = exp(-z * z / 2) / Sqr(2 * PI)
  y = 1 - d * t * ((((1.330274 * t - 1.821256) * t + 1.781478) * t - 0.3565638) * t + 0.3193815)
  If z < 0 Then
     y = 1 - y
  End If
  ei = Abs(y - fi)
  es = Abs(y - fs)

  dn = maxi(dn, es, ei)

  If il - cnt < 0 Then
     GoTo 6
  ElseIf il = cnt Then
     GoTo 8
  End If
  
  'compute z=dn*sqrt and prob
  zs = dn * Sqr(xn)
  kolmo = 1 - smirn(zs)
End Function

Function smirn(X)
 Dim PI As Double
 Dim c1, c2, c4, c8 As Double
 PI = 3.1415926
 If X <= 0.27 Then
    smirn = 0
 End If
 If X < 1 And X > 0.27 Then
    c1 = exp(-PI ^ 2 / (8 * (X ^ 2)))
    c2 = c1 * c1
    c4 = c2 * c2
    c8 = c4 * c4
    If c8 < 10 ^ -25 Then
       c8 = 0
    End If
    smirn = (Sqr(2 * PI) / X) * c1 * (1 + c8 * (1 + c8 * c8))
 End If
 If X >= 3.1 Then
    smirn = 1
 End If
 If X < 3.1 And X >= 1 Then
       c1 = exp(-2 * X * X)
       c2 = c1 * c1
       c4 = c2 * c2
       c8 = c4 * c4
       smirn = 1 - 2 * (c1 - c4 + c8 * (c1 - c8))
 End If
 
End Function
Function maxi(X, y, z)
 If X >= y And X >= z Then
       maxi = X
 End If
 If y >= X And y >= z Then
       maxi = y
 End If
If z >= y And z >= X Then
      maxi = z
 End If
 End Function
