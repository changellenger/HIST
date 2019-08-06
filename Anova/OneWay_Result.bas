Attribute VB_Name = "OneWay_Result"
Sub aResult(RnArray, VarName, st, se, td, ed, outputsheet)
    Dim sst As Double: Dim ttemp, addr, qq As Range
    Dim Comment1, Comment2 As String
    Dim val3535 As Long '초기위치 저장할 공간'
    Dim s3535 As Worksheet
    val3535 = 2
    
    On Error GoTo Err_delete
   
    For Each s3535 In ActiveWorkbook.Sheets
        If s3535.name = RstSheet Then
            val3535 = Sheets(RstSheet).Cells(1, 1).Value
        End If
    Next s3535  '시트가 이미있으면 출력 위치 저장을하고, 없으면 1을 저장한다.
    
    sst = st + se
    'If IsEmpty(outputsheet.Range("a1")) = True Then
    '   Set ttemp = outputsheet.Cells(2, 1)
    '   Set addr = outputsheet.Range("a1")
    'Else: Set addr = outputsheet.Range("a1")
    '     Set ttemp = outputsheet.Range(addr.Value)
    'End If
    
    Set addr = outputsheet.Range("a1")
    Set ttemp = outputsheet.Range("a" & addr.Value)
    
    
        Dim fvalue As Double
        fvalue = (st / td) / (se / ed)
        pvalue = Application.FDist(fvalue, td, ed)
        Set ttemp = ttemp.Offset(0, 1)
        yp = ttemp.Top
        
        Set Title = outputsheet.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 250, 22#)
        Title.Shadow.Type = msoShadow17
        With Title.Fill
            .ForeColor.SchemeColor = 1
            .Visible = msoTrue
            .Solid
        End With
        Title.TextFrame.Characters.Text = "분산분석표"
        With Title.TextFrame.Characters.Font
            .name = "굴림"
            .FontStyle = "굵게"
            .Size = 11
            .ColorIndex = xlAutomatic
        End With
        Title.TextFrame.HorizontalAlignment = xlCenter
        Set ttemp = ttemp.Offset(2, 0)
        Set qq = ttemp.Offset(3, 0)
        With ttemp.Resize(, 6).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With ttemp.Resize(, 6).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With qq.Resize(, 6).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        
        With qq.Resize(, 6).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
    
        ttemp.Value = "요인"
        ttemp.Offset(0, 1) = "제곱합"
        ttemp.Offset(0, 2) = "자유도"
        ttemp.Offset(0, 3) = "평균제곱"
        ttemp.Offset(0, 4) = "F값"
        ttemp.Offset(0, 5) = "유의확률"
        Set ttemp = ttemp.Offset(1, 0)
        ttemp.Value = "처리"
        ttemp.Offset(0, 1) = Format(st, "0.0000")
        ttemp.Offset(0, 2) = Format(td, "0.0000")
        ttemp.Offset(0, 3) = Format(st / td, "0.0000")
        ttemp.Offset(0, 4) = Format(fvalue, "0.0000")
        ttemp.Offset(0, 5) = Format(pvalue, "0.0000")
        Set ttemp = ttemp.Offset(1, 0)
        ttemp.Value = "잔차"
        ttemp.Offset(0, 1) = Format(se, "0.0000")
        ttemp.Offset(0, 2) = Format(ed, "0.0000")
        ttemp.Offset(0, 3) = Format(se / ed, "0.0000")
        Set ttemp = ttemp.Offset(1, 0)
        ttemp.Value = "계"
        ttemp.Offset(0, 1) = Format(sst, "0.0000")
        ttemp.Offset(0, 2) = Format(td + ed, "0.0000")
        
        Set ttemp = ttemp.Offset(1, 0)
        If se <> 0 Then
        If pvalue <= 0.01 Then
            Comment1 = """H0:모평균들이 서로 같다.""" & "를 유의수준 α=0.01에서 기각한다."
            Comment2 = "즉, 표본평균들이 아주 뚜렷한(p<0.01) 차이가 있다."
        ElseIf pvalue <= 0.05 Then
            Comment1 = """H0:모평균들이 서로 같다.""" & "를 유의수준 α=0.05에서 기각한다."
            Comment2 = "즉, 표본평균들이 뚜렷한(p<0.05) 차이가 있다."
        Else
            Comment1 = """H0:모평균들이 서로 같다.""" & "를 유의수준 α=0.05에서 기각할 수 없다."
            Comment2 = "즉, 표본평균들이 차이가 있다(p<0.05)고 할 수 없다."
        End If
        With ttemp
            .Value = Comment1
            .Font.Size = 9
            .HorizontalAlignment = xlLeft
        End With
        Set ttemp = ttemp.Offset(1, 0)
        With ttemp
            .Value = Comment2
            .Font.Size = 9
            .HorizontalAlignment = xlLeft
        End With
End If
        Set ttemp = ttemp.Offset(3, -1)
        
        addr.Value = Right(ttemp.Address, Len(ttemp.Address) - 3)
Exit Sub
Err_delete:

For Each s3535 In ActiveWorkbook.Sheets
        If s3535.name = RstSheet Then
Sheets(RstSheet).Range(Cells(val3535, 1), Cells(10000, 10000)).Select
Selection.Delete
Sheets(RstSheet).Cells(val3535, 1).Select

End If
Next s3535
    
End Sub
Sub dResult(ave, st, ct, fact, fn, outputsheet)
    Dim ttemp As Range
    Dim addr As Range
    Dim yp As Double
    '    If IsEmpty(outputsheet.Range("a1")) = True Then
    '   Set ttemp = outputsheet.Cells(2, 1)
    '   Set addr = outputsheet.Range("a1")
    'Else: Set addr = outputsheet.Range("a1")
    '     Set ttemp = outputsheet.Range(addr.Value)
    'End If
    
    Set addr = outputsheet.Range("a1")
    Set ttemp = outputsheet.Range("a" & addr.Value)
    
    yp = ttemp.Top
    Set Title = outputsheet.Shapes.AddShape(msoShapeRectangle, 3.75, yp + 2.25, 400, 25#)
    Title.TextFrame.Characters.Text = "일원배치 분산분석 결과"
    With Title
        .Fill.ForeColor.SchemeColor = 57
        .Line.DashStyle = msoLineSolid
        .Line.Style = msoLineSingle
        .Line.ForeColor.SchemeColor = 8
        .Line.Weight = 1
        .Line.Visible = msoTrue
       ' .Shadow.Type = msoShadow1
    End With
    
    With Title.TextFrame.Characters.Font
        .Size = 14
        .ColorIndex = 2
    End With
    Title.TextFrame.HorizontalAlignment = xlCenter
    Set ttemp = ttemp.Offset(3, 1)
    yp = ttemp.Top
    Set Title = outputsheet.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 250, 22#)
    Title.Shadow.Type = msoShadow17
    With Title.Fill
         .ForeColor.SchemeColor = 1
         .Visible = msoTrue
         .Solid
    End With
    Title.TextFrame.Characters.Text = "기술 통계량"
    With Title.TextFrame.Characters.Font
        .Size = 11
        .ColorIndex = xlAutomatic
    End With
    Title.TextFrame.HorizontalAlignment = xlCenter
    Set ttemp = ttemp.Offset(2, 0)
    Set qq = ttemp.Offset(fn, 0)
    With ttemp.Resize(, 4).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With ttemp.Resize(, 4).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With qq.Resize(, 4).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    
    ttemp.Offset(0, 1).Value = "개수"
    ttemp.Offset(0, 2).Value = "평균"
    ttemp.Offset(0, 3).Value = "표준편차"
    For i = 1 To fn
        Set ttemp = ttemp.Offset(1, 0)
        ttemp.Value = fact(i)
        ttemp.Offset(0, 1).Value = ct(i)
        ttemp.Offset(0, 2).Value = Format(ave(i), "0.0000")
        ttemp.Offset(0, 3).Value = Format(st(i), "0.0000")
    Next i
    Set ttemp = ttemp.Offset(3, -1)
    '''addr.Value = ttemp.Address
    addr.Value = Right(ttemp.Address, Len(ttemp.Address) - 3)
End Sub
Sub cResult(list, ave, fact, ct, se, td, ed, fn, alpha, outputsheet, Check4 As Boolean, check8 As Boolean, check5 As Boolean)
'다중비교 sub
    Dim ttemp, addr As Range
    Dim dave(), temp(), temp1(), temp2(), q(), names(), fact1, ave1, c As Double
    Dim tvalue(), pvalue(), pvalue1() As Double
    Dim N, z, a, b, d, count As Integer
     Set addr = outputsheet.Range("a1")
    Set ttemp = outputsheet.Range("a" & addr.Value)
    If fn <= 2 Then
    Set ttemp = ttemp.Offset(0, 1)
        yp = ttemp.Top
        Set Title = outputsheet.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 250, 20#)
        Title.Shadow.Type = msoShadow17
        With Title.Fill
            .ForeColor.SchemeColor = 1
            .Visible = msoTrue
            .Solid
        End With
        Title.TextFrame.Characters.Text = "다중비교 결과"
        With Title.TextFrame.Characters.Font
            .name = "굴림"
            .FontStyle = "굵게"
            .Size = 11
            .ColorIndex = xlAutomatic
        End With
        Title.TextFrame.HorizontalAlignment = xlCenter
        Set ttemp = ttemp.Offset(2, 0)
    Set ttemp = ttemp.Offset(1, 0)
            Comment2 = "인자의 수준수가 둘이하이므로 " & list & " 인자에 대한 다중비교를 수행할수 없습니다."
        With ttemp
            .Value = Comment2
            .Font.Size = 9
            .HorizontalAlignment = xlLeft
        End With
        Set ttemp = ttemp.Offset(3, -1)
   
        addr.Value = Right(ttemp.Address, Len(ttemp.Address) - 3)
             
    Else
     
      N = 0
      '평균을 통한 Sorting
    For i = 1 To fn
    For J = i + 1 To fn
    If ave(i) > ave(J) Then
    ave1 = ave(J)
    fact1 = fact(J)
    ave(J) = ave(i)
    fact(J) = fact(i)
    ave(i) = ave1
    fact(i) = fact1
    End If
    Next J
    Next i
    
    'Q-statistic을 구하기 위한 For문
        For i = 1 To fn - 1
        For J = i + 1 To fn
        N = N + 1: ReDim Preserve dave(1 To N): ReDim Preserve q(1 To N): ReDim Preserve names(1 To N)
        dave(N) = ave(i) - ave(J)
        names(N) = ave(J)
        q(N) = Abs(ave(i) - ave(J)) / (((se / ed) / ct(J)) ^ 0.5)
        
        Next J
        Next i
        
    'If IsEmpty(outputsheet.Range("a1")) = True Then
    '   Set ttemp = outputsheet.Cells(2, 1)
    '   Set addr = outputsheet.Range("a1")
    'Else: Set addr = outputsheet.Range("a1")
    '     Set ttemp = outputsheet.Range(addr.Value)
    'End If

    'Fisher LSD Method
    If Check4 = True Then
        ReDim pvalue(1 To N): ReDim tvalue(1 To N)
        'T-값과 P-value을 구하기 위한 for문
        For i = 1 To N
        tvalue(i) = Abs(dave(i)) / (((2 * se / ed) / ct(1)) ^ (0.5))
        pvalue(i) = Application.TDist(tvalue(i), ed, 2)
        Next i
        
        Set ttemp = ttemp.Offset(0, 1)
        yp = ttemp.Top
        Set Title = outputsheet.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 250, 20#)
        Title.Shadow.Type = msoShadow17
        With Title.Fill
            .ForeColor.SchemeColor = 1
            .Visible = msoTrue
            .Solid
        End With
        Title.TextFrame.Characters.Text = "다중비교 결과"
        With Title.TextFrame.Characters.Font
            .name = "굴림"
            .FontStyle = "굵게"
            .Size = 11
            .ColorIndex = xlAutomatic
        End With
        Title.TextFrame.HorizontalAlignment = xlCenter
        Set ttemp = ttemp.Offset(2, 0)
        
        'Fisher LSD table을 만드는 구간
        z = 0
        
        ttemp.Value = "Fisher's LSD"
        ttemp.Offset(0, 3) = "유의수준 = " & alpha / 100 & " 에 대한 그룹"
        ttemp.Offset(1, 0) = list
        ttemp.Offset(1, 1) = "자료수"
        'For i = 1 To fn
        'z = z + 1
        'ttemp.Offset(1, 1 + i) = z & " 그룹 "
        'Next i
        Set ttemp = ttemp.Offset(1, 0)
         

       Set ttemp = ttemp.Offset(1, 0)
       ttemp.Value = fact(1)
       ttemp.Offset(0, 1).Value = ct(1)
       ttemp.Offset(0, 2).Value = Format(ave(1), "0.0000")
       For i = 2 To fn
        ttemp.Offset(i - 1, 0).Value = fact(i)
        ttemp.Offset(i - 1, 1).Value = ct(i)
       Next i
       
       b = 0: c = 0
       Do Until b >= N
       If b = 0 Then
       a = b
       Do Until pvalue(a + 1) < alpha / 100
       a = a + 1
       If a >= fn - 1 Then Exit Do
       Loop
       For J = b + 1 To a
       ttemp.Offset(J, 2 + c).Value = Format(names(J), "0.0000")
       Next J
       If Format(ttemp.Offset(a, 2 + c).Value, "0.0000") = Format(names(N), "0.0000") Then Exit Do
       b = b + 1
       a = b + fn - 1 * b
       c = c + 1
       
       Else
       ttemp.Offset(b, 2 + c).Value = Format(names(b), "0.0000")
       d = 0
       Do Until pvalue(a + d) < alpha / 100
       d = d + 1
       If d >= fn - 1 - b Then Exit Do
       Loop
       
       If d = 0 And Format(names(b), "0.0000") <> Format(ttemp.Offset(b, 2 + c - 1).Value, "0.0000") Then
       ElseIf d = 0 And Format(names(b), "0.0000") = Format(ttemp.Offset(b, 2 + c - 1).Value, "0.0000") Then
       ttemp.Offset(b, 2 + c).Value = Empty
       c = c - 1
       Else
       For J = b + 1 To b + d
       If Format(names(J), "0.0000") = Format(ttemp.Offset(J, 2 + c - 1).Value, "0.0000") Then
       ttemp.Offset(b, 2 + c).Value = Empty
       c = c - 1 / d
       Else
       ttemp.Offset(J, 2 + c).Value = Format(names(J), "0.0000")
       End If
       Next J
       End If
       If d = 0 And a = N And Format(names(b), "0.0000") = Format(ttemp.Offset(b, 2 + c).Value, "0.0000") Then
       ttemp.Offset(b + 1, 2 + c + 1).Value = Format(names(a), "0.0000")
       End If
       If a >= N Then Exit Do
       If Format(ttemp.Offset(b + d, 2 + c).Value, "0.0000") = Format(names(N), "0.0000") Then Exit Do
       b = b + 1
       a = a + fn - 1 * b
       c = c + 1
       
       End If
       If c >= fn - 1 Then Exit Do
       Loop
        
       Set ttemp = ttemp.Offset(-2, 0)
       Set qq = ttemp.Offset(fn + 1, 0)
       For i = 1 To b + 1
       ttemp.Offset(1, 1 + i) = " 그룹 " & i
       Next i
      With ttemp.Resize(, 2 + b + 1).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    Set ttemp = ttemp.Offset(1, 0)
  
    With ttemp.Resize(, 2 + b + 1).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
       
    With qq.Resize(, 2 + b + 1).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    
   
    
    
    
       Set ttemp = ttemp.Offset(fn, 0)
       Set ttemp = ttemp.Offset(1, 0)
            Comment1 = " 같은 그룹에 속한 경우 유의수준 α= " & alpha / 100 & " 에서 처리평균에 차이가 없는 것으로 판단한다."
        With ttemp
            .Value = Comment1
            .Font.Size = 9
            .HorizontalAlignment = xlLeft
        End With
        Set ttemp = ttemp.Offset(1, 0)
        With ttemp
            .Value = Comment2
            .Font.Size = 9
            .HorizontalAlignment = xlLeft
        End With
       Set ttemp = ttemp.Offset(3, -1)
        '''addr.Value = ttemp.Address
        addr.Value = Right(ttemp.Address, Len(ttemp.Address) - 3)
              
    End If
    
    'Duncan과 Tukey를 구하기 위한 사전작업(Studentized range distribution 만들기)
If check5 = True Or check8 = True Then
  
        
        ReDim pvalue1(1 To N)
        For k = 1 To N
        z = 0
        For i = 0 To 10 Step 0.1
        z = z + 1: ReDim Preserve temp(1 To z)
        For J = -10 To 10 Step 0.1
        If J = -10 Then
        temp(z) = fn * (1 / (Application.PI() * 2) ^ 0.5) * exp(-(J ^ 2) / 2) _
                * ((Application.NormSDist(J) - Application.NormSDist(J - q(k) * i)) ^ (fn - 1)) _
                * 0.1
        Else
        temp(z) = temp(z) + fn * (1 / (Application.PI() * 2) ^ 0.5) * exp(-(J ^ 2) / 2) _
                * ((Application.NormSDist(J) - Application.NormSDist(J - q(k) * i)) ^ (fn - 1)) _
                * 0.1
        End If
        Next J
        Next i
        z = 0
        For t = 0 To 10 Step 0.1
        z = z + 1: ReDim Preserve temp1(1 To z)
        If t = 0 Then
        temp1(z) = temp(z) * Application.GammaDist(t, ed / 2, 2, 0) * 2 * (ed ^ (ed / 2)) _
                   * (t ^ (ed / 2)) * exp(t / 2) * exp(-ed * (t ^ 2) / 2) * 0.1
        Else
        temp1(z) = temp1(z - 1) + temp(z) * Application.GammaDist(t, ed / 2, 2, 0) * 2 _
                 * (ed ^ (ed / 2)) * (t ^ (ed / 2)) * exp(t / 2) _
                 * exp(-ed * (t ^ 2) / 2) * 0.1
        End If
        Next t
        pvalue1(k) = 1 - temp1(101)
        Next k
        
        'Tukey HSD Method
        If check8 = True Then
        Set ttemp = ttemp.Offset(0, 1)
'        yp = ttemp.Top
'        Set Title = outputsheet.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 90, 22#)
'        Title.Shadow.Type = msoShadow17
'        With Title.Fill
'            .ForeColor.SchemeColor = 22
'            .Visible = msoTrue
'            .Solid
'        End With
'        Title.TextFrame.Characters.Text = "다중비교 결과"
'        With Title.TextFrame.Characters.Font
'            .name = "굴림"
'            .FontStyle = "굵게"
'            .Size = 11
'            .ColorIndex = xlAutomatic
'        End With
'        Title.TextFrame.HorizontalAlignment = xlCenter
'        Set ttemp = ttemp.Offset(2, 0)

        'Tukey HSD table을 만드는 구간
        z = 0
        ttemp.Value = "Tukey HSD"
        ttemp.Offset(0, 3) = "유의수준 = " & alpha / 100 & " 에 대한 그룹"
        ttemp.Offset(1, 0) = list
        ttemp.Offset(1, 1) = "자료수"
        
        
        Set ttemp = ttemp.Offset(1, 0)
         
       
       Set ttemp = ttemp.Offset(1, 0)
       ttemp.Value = fact(1)
       ttemp.Offset(0, 1).Value = ct(1)
       ttemp.Offset(0, 2).Value = Format(ave(1), "0.0000")
       For i = 2 To fn
        ttemp.Offset(i - 1, 0).Value = fact(i)
        ttemp.Offset(i - 1, 1).Value = ct(i)
       Next i
       
       'ReDim temp2(1 To N)
       b = 0: c = 0
       Do Until b >= N
       If b = 0 Then
       a = b
       Do Until pvalue1(a + 1) < alpha / 100
       a = a + 1
       If a >= fn - 1 Then Exit Do
       Loop
       For J = b + 1 To a
       ttemp.Offset(J, 2 + c).Value = Format(names(J), "0.0000")
       Next J
       
       'If a = 0 Then
       'temp2(c + 1) = 1
       'Else
       'temp2(c + 1) = pvalue1(a)
       'End If
       If Format(ttemp.Offset(a, 2 + c).Value, "0.0000") = Format(names(N), "0.0000") Then Exit Do
       b = b + 1
       a = b + fn - 1 * b
       c = c + 1
              
       Else
       ttemp.Offset(b, 2 + c).Value = Format(names(b), "0.0000")
       d = 0
       Do Until pvalue1(a + d) < alpha / 100
       d = d + 1
       If d >= fn - 1 - b Then Exit Do
       Loop
       
       
       If d = 0 And Format(names(b), "0.0000") <> Format(ttemp.Offset(b, 2 + c - 1).Value, "0.0000") Then
       'temp2(c + 1) = 1
       ElseIf d = 0 And Format(names(b), "0.0000") = Format(ttemp.Offset(b, 2 + c - 1).Value, "0.0000") Then
       ttemp.Offset(b, 2 + c).Value = Empty
       c = c - 1
       Else
       For J = b + 1 To b + d
       If Format(names(J), "0.0000") = Format(ttemp.Offset(J, 2 + c - 1).Value, "0.0000") Then
       ttemp.Offset(b, 2 + c).Value = Empty
       c = c - 1 / d
       Else
       ttemp.Offset(J, 2 + c).Value = Format(names(J), "0.0000")
       'temp2(c + 1) = pvalue1(a + d - 1)
       End If
       Next J
       End If
       If d = 0 And a = N And Format(names(b), "0.0000") = Format(ttemp.Offset(b, 2 + c).Value, "0.0000") Then
       ttemp.Offset(b + 1, 2 + c + 1).Value = Format(names(a), "0.0000")
       'temp2(c + 2) = 1
       c = c + 1
       End If
       
       If a >= N Then Exit Do
       If Format(ttemp.Offset(b + d, 2 + c).Value, "0.0000") = Format(names(N), "0.0000") Then Exit Do
       b = b + 1
       a = a + fn - 1 * b
       c = c + 1
       End If
       If c >= fn - 1 Then Exit Do
       Loop
       
        Set ttemp = ttemp.Offset(-2, 0)
       Set qq = ttemp.Offset(fn + 1, 0)
       For i = 1 To b + 1
       ttemp.Offset(1, 1 + i) = " 그룹 " & i
       Next i
      With ttemp.Resize(, 2 + b + 1).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Set ttemp = ttemp.Offset(1, 0)
    With ttemp.Resize(, 2 + b + 1).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    
    With qq.Resize(, 2 + b + 1).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    
    
       
       Set ttemp = ttemp.Offset(fn, 0)
       Set ttemp = ttemp.Offset(1, 0)
            Comment1 = " 같은 그룹에 속한 경우 유의수준 α= " & alpha / 100 & " 에서 처리평균에 차이가 없는 것으로 판단한다."
       
       
        With ttemp
            .Value = Comment1
            .Font.Size = 9
            .HorizontalAlignment = xlLeft
        End With
        Set ttemp = ttemp.Offset(1, 0)
        With ttemp
            .Value = Comment2
            .Font.Size = 9
            .HorizontalAlignment = xlLeft
        End With
       Set ttemp = ttemp.Offset(3, -1)
        '''addr.Value = ttemp.Address
        addr.Value = Right(ttemp.Address, Len(ttemp.Address) - 3)
        End If
        
        'Duncan Method
        If check5 = True Then
        Set ttemp = ttemp.Offset(0, 1)
'        yp = ttemp.Top
'        Set Title = outputsheet.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 90, 22#)
'        Title.Shadow.Type = msoShadow17
'        With Title.Fill
'            .ForeColor.SchemeColor = 22
'            .Visible = msoTrue
'            .Solid
'        End With
'        Title.TextFrame.Characters.Text = "다중비교 결과"
'        With Title.TextFrame.Characters.Font
'            .name = "굴림"
'            .FontStyle = "굵게"
'            .Size = 11
'            .ColorIndex = xlAutomatic
'        End With
'        Title.TextFrame.HorizontalAlignment = xlCenter
'        Set ttemp = ttemp.Offset(2, 0)

        
        'Tukey와 달리 Ducan은 comparisonwise error rate을 사용
        '이를 이용해서 다시 Ducan에 관한 P-value를 구하는 For 문
          For k = 1 To N
        pvalue1(k) = 1 - ((1 - pvalue1(k)) ^ (1 / 4))
        Next k
        'Ducan table을 만드는 구간
        z = 0
        ttemp.Value = "Duncan"
        ttemp.Offset(0, 3) = "유의수준 = " & alpha / 100 & " 에 대한 그룹"
        ttemp.Offset(1, 0) = list
        ttemp.Offset(1, 1) = "자료수"
        
        Set ttemp = ttemp.Offset(1, 0)
         
       
       Set ttemp = ttemp.Offset(1, 0)
       ttemp.Value = fact(1)
       ttemp.Offset(0, 1).Value = ct(1)
       ttemp.Offset(0, 2).Value = Format(ave(1), "0.0000")
       For i = 2 To fn
        ttemp.Offset(i - 1, 0).Value = fact(i)
        ttemp.Offset(i - 1, 1).Value = ct(i)
       Next i
       
       b = 0: c = 0
       Do Until b >= N
       If b = 0 Then
       a = b
       Do Until pvalue1(a + 1) < alpha / 100
       a = a + 1
       If a >= fn - 1 Then Exit Do
       Loop
       For J = b + 1 To a
       ttemp.Offset(J, 2 + c).Value = Format(names(J), "0.0000")
       Next J
       If Format(ttemp.Offset(a, 2 + c).Value, "0.0000") = Format(names(N), "0.0000") Then Exit Do
       b = b + 1
       a = b + fn - 1 * b
       c = c + 1
       
       Else
       ttemp.Offset(b, 2 + c).Value = Format(names(b), "0.0000")
       d = 0
       Do Until pvalue1(a + d) < alpha / 100
       d = d + 1
       If d >= fn - 1 - b Then Exit Do
       Loop
       
       If d = 0 And Format(names(b), "0.0000") <> Format(ttemp.Offset(b, 2 + c - 1).Value, "0.0000") Then
       ElseIf d = 0 And Format(names(b), "0.0000") = Format(ttemp.Offset(b, 2 + c - 1).Value, "0.0000") Then
       ttemp.Offset(b, 2 + c).Value = Empty
       c = c - 1
       Else
       For J = b + 1 To b + d
       If Format(names(J), "0.0000") = Format(ttemp.Offset(J, 2 + c - 1).Value, "0.0000") Then
       ttemp.Offset(b, 2 + c).Value = Empty
       c = c - 1 / d
       Else
       ttemp.Offset(J, 2 + c).Value = Format(names(J), "0.0000")
       End If
       Next J
       End If
       If d = 0 And a = N And Format(names(b), "0.0000") = Format(ttemp.Offset(b, 2 + c).Value, "0.0000") Then
       ttemp.Offset(b + 1, 2 + c + 1).Value = Format(names(a), "0.0000")
       End If
       If a >= N Then Exit Do
       If Format(ttemp.Offset(b + d, 2 + c).Value, "0.0000") = Format(names(N), "0.0000") Then Exit Do
       b = b + 1
       a = a + fn - 1 * b
       c = c + 1
       
       End If
       If c >= fn - 1 Then Exit Do
       Loop
         Set ttemp = ttemp.Offset(-2, 0)
       Set qq = ttemp.Offset(fn + 1, 0)
       
       For i = 1 To b + 1
       ttemp.Offset(1, 1 + i) = " 그룹 " & i
       Next i
       
      With ttemp.Resize(, 2 + b + 1).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Set ttemp = ttemp.Offset(1, 0)
    With ttemp.Resize(, 2 + b + 1).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    
    With qq.Resize(, 2 + b + 1).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    
  
       
       Set ttemp = ttemp.Offset(fn, 0)
       Set ttemp = ttemp.Offset(1, 0)
            Comment1 = " 같은 그룹에 속한 경우 유의수준 α= " & alpha / 100 & " 에서 처리평균에 차이가 없는 것으로 판단한다."
        With ttemp
            .Value = Comment1
            .Font.Size = 9
            .HorizontalAlignment = xlLeft
        End With
        Set ttemp = ttemp.Offset(1, 0)
        With ttemp
            .Value = Comment2
            .Font.Size = 9
            .HorizontalAlignment = xlLeft
        End With
       Set ttemp = ttemp.Offset(3, -1)
        '''addr.Value = ttemp.Address
        addr.Value = Right(ttemp.Address, Len(ttemp.Address) - 3)
    End If
    End If
    End If

End Sub
Sub eResult(ave, vrn, st, ct, fn, outputsheet)
'등분산 검정 sub (Levene's test)
    Dim ttemp, addr As Range
    Dim temp, temp1 As Double
    Dim s1, xsq, xsq1, xsq2, w As Double
    Dim z(), s() As Double
    Dim pvalue As Double
    Dim N, k As Long
    Dim one, ind As Integer
    
    temp = 0
    one = 0 '수준수가 3이하인 인자의 갯수
    ind = 1 '수준수가 3이하인 인자들은 지우기 위한 변수
    
    For i = 1 To fn
        If ct(i) < 3 Then
            one = one + 1
        End If
        temp = temp + ct(i)
        N = temp
    Next i

    Set addr = outputsheet.Range("a1")
    Set ttemp = outputsheet.Range("a" & addr.Value)
    
    If one = 0 Then
    'Levene's test 식 참고
    k = 0:
    ReDim z(1 To N - one)
    ReDim s(1 To fn - one)
    
    For i = 1 To fn
        temp1 = 0
        s(i) = 0
            For J = 0 To ct(i) - 1
                k = k + 1
                z(k) = Abs(vrn.Cells(k, 1) - ave(i))
                temp1 = temp1 + z(k)
            Next J
        s(i) = temp1
    Next i
    
    ind = 1 '수준수가 3이하인 인자들은 지우기 위한 변수 초기화
    
    For i = 1 To fn
        s(i) = s(i) / ct(i)
    Next i
    
    s1 = Application.Average(s)
    xsq = 0: xsq1 = 0: xsq2 = 0:
    For i = 1 To N - one
    xsq = xsq + (z(i) ^ 2)
    Next i
   
    For i = 1 To fn
    xsq1 = xsq1 + (s(i) ^ 2) * ct(i)
    xsq2 = xsq2 + ((s(i) - s1) ^ 2) * ct(i)
    Next i
    
    
    w = ((N - fn) / (fn - 1)) * (xsq2 / (xsq - xsq1))
    pvalue = Application.FDist(w, fn - 1, N - fn)
    
        Set ttemp = ttemp.Offset(0, 1)
        yp = ttemp.Top
        Set Title = outputsheet.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 250, 20#)
        Title.Shadow.Type = msoShadow17
        With Title.Fill
            .ForeColor.SchemeColor = 1
            .Visible = msoTrue
            .Solid
        End With
        Title.TextFrame.Characters.Text = "등분산 검정"
        With Title.TextFrame.Characters.Font
            .name = "굴림"
            .FontStyle = "굵게"
            .Size = 11
            .ColorIndex = xlAutomatic
        End With
        
        Title.TextFrame.HorizontalAlignment = xlCenter
        Set ttemp = ttemp.Offset(2, 0)
        Set qq = ttemp.Offset(3, 0)
        
        With ttemp.Resize(, 6).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With ttemp.Resize(, 6).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
          
        ttemp.Value = "Levene's test"
        ttemp.Offset(0, 1) = "제곱합"
        ttemp.Offset(0, 2) = "자유도"
        ttemp.Offset(0, 3) = "평균제곱"
        ttemp.Offset(0, 4) = "F값"
        ttemp.Offset(0, 5) = "유의확률"
      
        Set ttemp = ttemp.Offset(1, 0)
        ttemp.Value = "처리"
        ttemp.Offset(0, 1) = Format(xsq2, "0.0000")
        ttemp.Offset(0, 2) = Format(fn - 1, "0.0000")
        ttemp.Offset(0, 3) = Format(xsq2 / (fn - 1), "0.0000")
        ttemp.Offset(0, 4) = Format(w, "0.0000")
        ttemp.Offset(0, 5) = Format(pvalue, "0.0000")
         
        Set ttemp = ttemp.Offset(1, 0)
        ttemp.Value = "잔차"
        ttemp.Offset(0, 1) = Format(xsq - xsq1, "0.0000")
        ttemp.Offset(0, 2) = Format(N - fn, "0.0000")
        ttemp.Offset(0, 3) = Format((xsq - xsq1) / (N - fn), "0.0000")
                
        With ttemp.Resize(, 6).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
            End With
        Set ttemp = ttemp.Offset(1, 0)
            Comment1 = " 유의확률 값이 유의수준 α 보다 작으면 등분산 가정이 만족하지 않음을 의미한다."
        With ttemp
            .Value = Comment1
            .Font.Size = 9
            .HorizontalAlignment = xlLeft
        End With
        Set ttemp = ttemp.Offset(1, 0)
        With ttemp
            .Value = Comment2
            .Font.Size = 9
            .HorizontalAlignment = xlLeft
        End With
        Set ttemp = ttemp.Offset(3, -1)
        addr.Value = Right(ttemp.Address, Len(ttemp.Address) - 3)
        
        
    Else
    Set ttemp = ttemp.Offset(0, 1)
        yp = ttemp.Top
        Set Title = outputsheet.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 250, 20#)
        Title.Shadow.Type = msoShadow17
        With Title.Fill
            .ForeColor.SchemeColor = 1
            .Visible = msoTrue
            .Solid
        End With
        Title.TextFrame.Characters.Text = "등분산 검정"
        With Title.TextFrame.Characters.Font
            .name = "굴림"
            .FontStyle = "굵게"
            .Size = 11
            .ColorIndex = xlAutomatic
        End With
        
        Title.TextFrame.HorizontalAlignment = xlCenter
        Set ttemp = ttemp.Offset(2, 0)
        Set qq = ttemp.Offset(3, 0)
        
        With ttemp.Resize(, 6).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With ttemp.Resize(, 6).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
          
        ttemp.Offset(0, 4) = "수준수가 1인 인자가 있어서 Levene's test를 할수 없습니다."
        Set ttemp = ttemp.Offset(3, -1)
        addr.Value = Right(ttemp.Address, Len(ttemp.Address) - 3)
     End If
        
End Sub
