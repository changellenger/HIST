Attribute VB_Name = "ratiotest"
Sub ratioresult(Z, hatp, p, n, s, c, Str, OutputSheet, choice)
    Dim ttemp As Range
    Dim addr As Range
    Dim yp, ul, ll, za As Double
    Dim Hyp As Integer
    'If IsEmpty(outputsheet.Range("a1")) = True Then
    '   Set ttemp = outputsheet.Cells(2, 1)
    '   Set addr = outputsheet.Range("a1")
    'Else: Set addr = outputsheet.Range("a1")
    '     Set ttemp = outputsheet.Range(addr.Value)
    'End If
    
    Set addr = OutputSheet.Range("a1")
    Set ttemp = OutputSheet.Range("a" & addr.Value + 2)
    
    
    Hyp = choice(3)
    
    za = Application.NormSInv((100 - c) / 200)
    ul = hatp - za * Sqr(hatp * (1 - hatp) / n)
    ll = hatp + za * Sqr(hatp * (1 - hatp) / n)
    yp = ttemp.Top
    Set title = OutputSheet.Shapes.AddShape(msoShapeRectangle, 3.75, yp + 2.25, 400, 25#)
    
    title.TextFrame.Characters.Text = "모비율에 대한 z-검정"

    With title
        .Fill.ForeColor.SchemeColor = 57
        .Line.DashStyle = msoLineSolid
        .Line.Style = msoLineSingle
        .Line.Weight = 1
        .Line.Visible = msoTrue
'        .Shadow.Type = msoShadow1
    End With
    
    With title.TextFrame.Characters.Font
        .size = 14
        .ColorIndex = 2
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
    Set ttemp = ttemp.Offset(3, 1)
     yp = ttemp.Top
    Set title = OutputSheet.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 250, 22#)
    title.Shadow.Type = msoShadow17
    With title.Fill
         .ForeColor.SchemeColor = 1
         .Visible = msoTrue
         .Solid
    End With
    title.TextFrame.Characters.Text = "모비율 추정 "
    With title.TextFrame.Characters.Font
        .size = 12
        .ColorIndex = xlAutomatic
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
    Set ttemp = ttemp.Offset(3, 0)
    ttemp.Value = "시행횟수"
    ttemp.Offset(0, 1) = "성공횟수"
    ttemp.Offset(0, 3) = "비율의 추정치"
    ttemp.Offset(1, 0) = n
    ttemp.Offset(1, 1) = s
    ttemp.Offset(1, 3) = Format(hatp, "##0.000")
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
    With ttemp.Offset(1, 0).Resize(, 4).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
 
    
    
    
    
    
    
    Set ttemp = ttemp.Offset(4, 0)
     yp = ttemp.Top
    Set title = OutputSheet.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 250, 22#)
    title.Shadow.Type = msoShadow17
    With title.Fill
         .ForeColor.SchemeColor = 1
         .Visible = msoTrue
         .Solid
    End With
    title.TextFrame.Characters.Text = "H0 : p=" & p & " 의 가설검정"
    With title.TextFrame.Characters.Font
        .size = 12
        .ColorIndex = xlAutomatic
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
    Set ttemp = ttemp.Offset(3, 2)
    Select Case Hyp
    Case 1
       ttemp.Value = " H0 : p = " & p & " vs. H1 : p ≠ " & p
    Case 2
       ttemp.Value = " H0 : p = " & p & " vs. H1 : p > " & p
    Case 3
       ttemp.Value = " H0 : p = " & p & " vs. H1 : p < " & p
    End Select
        
    
    Set ttemp = ttemp.Offset(3, -2)
    ttemp.Value = "검정통계량"
    ttemp.Offset(1, 0).Value = Format(Z, "0.0000")
    With ttemp.Resize(, 1).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With ttemp.Resize(, 1).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With ttemp.Offset(1, 0).Resize(, 1).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
 
  '  ttemp.Offset(1, 0) = "유의확률"
  '  ttemp.Offset(1, 1).Value = Format(2 * (1 - WorksheetFunction.NormSDist(Abs(Z))), "0.0000")
    Set ttemp = ttemp.Offset(4, 0)
    yp = ttemp.Top
    Set title = OutputSheet.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 250, 22#)
    title.Shadow.Type = msoShadow17
    With title.Fill
         .ForeColor.SchemeColor = 1
         .Visible = msoTrue
         .Solid
    End With
       
    
    
    
    
    title.TextFrame.Characters.Text = "모비율에 대한 " & c & "% 신뢰구간"
    With title.TextFrame.Characters.Font
        .size = 12
        .ColorIndex = xlAutomatic
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
    Set ttemp = ttemp.Offset(3, 0)
    ttemp.Value = c & "% 신뢰구간"
    ttemp.Offset(0, 1) = "하한"
    ttemp.Offset(0, 2) = "상한"
    ttemp.Offset(1, 0) = " "
    ttemp.Offset(1, 1) = Format(ll, "0.0000")
    ttemp.Offset(1, 2) = Format(ul, "0.0000")
    With ttemp.Resize(, 3).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With ttemp.Resize(, 3).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With ttemp.Offset(1, 0).Resize(, 3).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
     Set ttemp = ttemp.Offset(4, 0)
    yp = ttemp.Top
    Set title = OutputSheet.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 250, 22#)
    title.Shadow.Type = msoShadow17
    With title.Fill
         .ForeColor.SchemeColor = 1
         .Visible = msoTrue
         .Solid
    End With
   
    title.TextFrame.Characters.Text = "요약 및 결론"
    With title.TextFrame.Characters.Font
        .size = 12
        .ColorIndex = xlAutomatic
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
     Set ttemp = ttemp.Offset(3, 0)
   ttemp.Value = "n = " & n & ", p0= " & Format(hatp, "##0.000") & "이고, np0= " & n * hatp & ", n(1 - p0) = " & n * (1 - hatp) & "로 모두 5보다 정규근사 검정을 시행할 수 있다."
   ttemp.Offset(1, 0) = "검정통계량은 " & Format(Z, "0.0000") & "이다. "
    With ttemp
        .HorizontalAlignment = xlLeft
        ttemp.Offset(1, 0).HorizontalAlignment = xlLeft
    End With
    
    Set ttemp = ttemp.Offset(2, 0)
    With ttemp
        
        .HorizontalAlignment = xlLeft
    End With
    Set ttemp = ttemp.Offset(2, -1)
    '''addr.Value = ttemp.Address
    addr.Value = Right(ttemp.Address, Len(ttemp.Address) - 3)
End Sub

Sub ratio2result(hatp1, hatp2, Z, n1, n2, s1, s2, c, OutputSheet, choice)
    Dim ttemp As Range
    Dim addr As Range
    Dim yp, ul, ll, za As Double
    Dim Hyp As Integer
    'If IsEmpty(outputsheet.Range("a1")) = True Then
    '   Set ttemp = outputsheet.Cells(2, 1)
    '   Set addr = outputsheet.Range("a1")
    'Else: Set addr = outputsheet.Range("a1")
    '     Set ttemp = outputsheet.Range(addr.Value)
    'End If
    
    Set addr = OutputSheet.Range("a1")
    Set ttemp = OutputSheet.Range("a" & addr.Value + 2)
    
    
    Hyp = choice(3)
    
    za = Application.NormSInv((100 - c) / 200)
    ul = hatp1 - hatp2 - za * Sqr(hatp1 * (1 - hatp1) / n1 + hatp2 * (1 - hatp2) / n2)
    ll = hatp1 - hatp2 + za * Sqr(hatp1 * (1 - hatp1) / n1 + hatp2 * (1 - hatp2) / n2)
    yp = ttemp.Top
    Set title = OutputSheet.Shapes.AddShape(msoShapeRectangle, 3.75, yp + 2.25, 400, 25#)
    
    With title
        .Fill.ForeColor.SchemeColor = 57
        .Line.DashStyle = msoLineSolid
        .Line.Style = msoLineSingle
        .Line.Weight = 1
        .Line.Visible = msoTrue
  '      .Shadow.Type = msoShadow1
    End With
    
    title.TextFrame.Characters.Text = "두 모비율에 대한 z-검정"
    With title.TextFrame.Characters.Font
        .size = 14
        .ColorIndex = 2
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
    Set ttemp = ttemp.Offset(3, 1)
    yp = ttemp.Top
    Set title = OutputSheet.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 250, 22#)
    title.Shadow.Type = msoShadow17
    With title.Fill
         .ForeColor.SchemeColor = 1
         .Visible = msoTrue
         .Solid
    End With
    title.TextFrame.Characters.Text = "모비율 추정 "
    With title.TextFrame.Characters.Font
        .size = 12
        .ColorIndex = xlAutomatic
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
    
    Set ttemp = ttemp.Offset(3, 0)
    ttemp.Value = ""
    ttemp.Offset(1, 0) = "집단1"
    ttemp.Offset(2, 0) = "집단2"
    ttemp.Offset(0, 1) = "시행횟수"
    ttemp.Offset(0, 2) = "성공횟수"
    ttemp.Offset(0, 3) = "비율추정치"
    ttemp.Offset(1, 1) = n1
    ttemp.Offset(2, 1) = n2
    ttemp.Offset(1, 2) = s1
    ttemp.Offset(2, 2) = s2
    ttemp.Offset(1, 3) = Format(hatp1, "##0.000")
    ttemp.Offset(2, 3) = Format(hatp2, "##0.000")
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
    With ttemp.Offset(2, 0).Resize(, 4).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Set ttemp = ttemp.Offset(4, 0)
    
    yp = ttemp.Top
    Set title = OutputSheet.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 250, 22#)
    title.Shadow.Type = msoShadow17
    With title.Fill
         .ForeColor.SchemeColor = 1
         .Visible = msoTrue
         .Solid
    End With

    
      title.TextFrame.Characters.Text = "H0 : p1 = p2의가설검정"
    With title.TextFrame.Characters.Font
        .size = 12
        .ColorIndex = xlAutomatic
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
    Set ttemp = ttemp.Offset(3, 2)
    Select Case Hyp
    Case 1
       ttemp.Value = " H0 : p1 = p2  vs. H1 : p1 ≠ p2"
    Case 2
       ttemp.Value = " H0 : p1 = p2  vs. H1 : p1 > p2"
    Case 3
       ttemp.Value = " H0 : p1 = p2  vs. H1 : p1 < p2"
    End Select
        
    
    title.TextFrame.HorizontalAlignment = xlCenter
    Set ttemp = ttemp.Offset(3, -2)
    ttemp.Value = "검정통계량"
    ttemp.Offset(0, 1).Value = "유의확률"
    ttemp.Offset(1, 0) = Format(Z, "0.0000")
    ttemp.Offset(1, 1).Value = Format(2 * (1 - Application.NormSDist(Abs(Z))), "0.0000")
        With ttemp.Resize(, 2).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With ttemp.Resize(, 2).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With ttemp.Offset(1, 0).Resize(, 2).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    
    Set ttemp = ttemp.Offset(4, 0)
    yp = ttemp.Top
    Set title = OutputSheet.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 250, 25#)
    title.Shadow.Type = msoShadow17
    With title.Fill
         .ForeColor.SchemeColor = 1
         .Visible = msoTrue
         .Solid
    End With
    
    title.TextFrame.Characters.Text = "p1-p2에 대한 " & c & "% 신뢰구간"
    With title.TextFrame.Characters.Font
        .size = 12
        .ColorIndex = xlAutomatic
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
    Set ttemp = ttemp.Offset(3, 0)
    ttemp.Value = c & "% 신뢰구간"
    ttemp.Offset(0, 1) = "하한"
    ttemp.Offset(0, 2) = "상한"
    ttemp.Offset(1, 1) = Format(ll, "0.0000")
    ttemp.Offset(1, 2) = Format(ul, "0.0000")
    With ttemp.Resize(, 3).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With ttemp.Resize(, 3).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With ttemp.Offset(1, 0).Resize(, 3).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With

    Set ttemp = ttemp.Offset(4, 0)
    yp = ttemp.Top
    Set title = OutputSheet.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 250, 22#)
    title.Shadow.Type = msoShadow17
    With title.Fill
         .ForeColor.SchemeColor = 1
         .Visible = msoTrue
         .Solid
    End With
   
    title.TextFrame.Characters.Text = "요약 및 결론"
    With title.TextFrame.Characters.Font
        .size = 12
        .ColorIndex = xlAutomatic
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
     Set ttemp = ttemp.Offset(3, 0)
   ttemp.Value = "집단1의 추정값은 X1/n1 =" & s1 & "/" & n1 & "=" & Format(s1 / n1, "##0.000") & "이고,"
   ttemp.Offset(1, 0) = "집단2의 추정값은 X2/n2 =" & s2 & "/" & n2 & "=" & Format(s2 / n2, "##0.000") & "이다."
   ttemp.Offset(2, 0) = "공통 모비율에 대한 추정값은 (X1+X2)/(n1+n2) = " & Format((s1 + s2) / (n1 + n2), "##0.000") & "임을 알 수 있다."
   ttemp.Offset(3, 0) = "유의수준이" & (100 - c) / 100 & " 이고 유의확률이 " & Format(2 * (1 - Application.NormSDist(Abs(Z))), "0.0000") & "이다."
  
    If 2 * (1 - Application.NormSDist(Abs(Z))) < (100 - c) / 100 Then
     ttemp.Offset(4, 0).Value = "유의확률(p-value)가 유의수준α= " & (100 - c) / 100 & "보다 작기때문에 귀무가설H。을 기각한다."
    Else
     ttemp.Offset(4, 0).Value = "유의확률(p-value)가 유의수준α= " & (100 - c) / 100 & "보다 크기때문에 귀무가설H。을 기각하지 못한다."
    
    End If
    With ttemp
        .HorizontalAlignment = xlLeft
         ttemp.Offset(1, 0).HorizontalAlignment = xlLeft
         ttemp.Offset(2, 0).HorizontalAlignment = xlLeft
         ttemp.Offset(3, 0).HorizontalAlignment = xlLeft
         ttemp.Offset(4, 0).HorizontalAlignment = xlLeft
    End With
   

    
    
   
    Set ttemp = ttemp.Offset(4, -1)
    '''addr.Value = ttemp.Address
    addr.Value = Right(ttemp.Address, Len(ttemp.Address) - 3)
End Sub

Sub resultOneChi(Z, hatp, p, n, s, c, Str, OutputSheet, choice)
    Dim ttemp As Range
    Dim addr As Range
    Dim yp, ul, ll, za, za1, za2 As Double
    'If IsEmpty(outputsheet.Range("a1")) = True Then
    '   Set ttemp = outputsheet.Cells(2, 1)
    '   Set addr = outputsheet.Range("a1")
    'Else: Set addr = outputsheet.Range("a1")
    '     Set ttemp = outputsheet.Range(addr.Value)
    'End If
    Dim Hyp As Integer
    
    Set addr = OutputSheet.Range("a1")
    Set ttemp = OutputSheet.Range("a" & addr.Value + 2)
    
    za = (100 - c) / 100 '유의수준
    za1 = WorksheetFunction.ChiInv(za / 2, n - 1) '카이제곱(알파/2) '하한
    za2 = WorksheetFunction.ChiInv(1 - (za / 2), n - 1) '카이제곱(1-(알파/2))'상한

    ll = ((n - 1) * s) / za1
    ul = ((n - 1) * s) / za2
    
    Hyp = choice(3)
    
    TModulePrint.Title1 "모분산σ²에 대한 Χ²검정"

     ''''
    '''''기초통계량
    '''''
    
    Set ttemp = ttemp.Offset(3, 1)
     
     yp = ttemp.Top
    Set title = OutputSheet.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 250, 22#)
    title.Shadow.Type = msoShadow17
    With title.Fill
         .ForeColor.SchemeColor = 1
         .Visible = msoTrue
         .Solid
    End With
    title.TextFrame.Characters.Text = "모분산 추정 "
    With title.TextFrame.Characters.Font
        .size = 12
        .ColorIndex = xlAutomatic
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
    
   
    Set ttemp = ttemp.Offset(3, 0)
    ttemp.Value = "표본의 개수"
    ttemp.Offset(0, 1) = "표본의 분산"
    ttemp.Offset(0, 2) = "모분산"
    ttemp.Offset(1, 0) = n
    ttemp.Offset(1, 1) = s
    ttemp.Offset(1, 2) = p
    With ttemp.Resize(, 3).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With ttemp.Resize(, 3).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With ttemp.Offset(1, 0).Resize(, 3).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
 
    ''''
    '''''가설검정
    '''''
    Set ttemp = ttemp.Offset(4, 0)
     yp = ttemp.Top
    Set title = OutputSheet.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 250, 22#)
    title.Shadow.Type = msoShadow17
    With title.Fill
         .ForeColor.SchemeColor = 1
         .Visible = msoTrue
         .Solid
    End With
   
    
    title.TextFrame.Characters.Text = "H0 : σ²= " & p & " 에 대한 가설 검정"
    With title.TextFrame.Characters.Font
        .size = 12
        .ColorIndex = xlAutomatic
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
    
    Set ttemp = ttemp.Offset(3, 2)
    Select Case Hyp
    Case 1
       ttemp.Value = " H0 : σ² =σ²。vs. H1 : σ² ≠σ²。     (σ²。= " & p & " )"
    Case 2
       ttemp.Value = " H0 : σ² =σ²。vs. H1 : σ² > σ²。     (σ²。= " & p & " )"
    Case 3
    ttemp.Value = " H0 : σ² =σ²。vs. H1 : σ² < σ²。     (σ²。= " & p & " )"
    End Select
    
    'ttemp.Select
   ' With Selection.Font
             '.FontStyle = "Bold"
             '.size = 10
   ' End With
    
    
    Set ttemp = ttemp.Offset(3, -2)
    ttemp.Value = "유의수준"
    ttemp.Offset(0, 1) = "자유도"
    ttemp.Offset(0, 2) = "검정통계량"
    ttemp.Offset(0, 3) = "기각역"
    ttemp.Offset(1, 0) = za
    ttemp.Offset(1, 1) = n - 1
    ttemp.Offset(1, 2) = Format((n - 1) * s / p, "##0.0000")
    ttemp.Offset(1, 3) = Format(WorksheetFunction.ChiInv(ttemp.Offset(1, 0), n - 1), "##0.0000")
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
    With ttemp.Offset(1, 0).Resize(, 4).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
 
    
    
   '''''
    '''''신뢰구간
    '''''
    Set ttemp = ttemp.Offset(4, 0)
    yp = ttemp.Top
    Set title = OutputSheet.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 250, 22#)
    title.Shadow.Type = msoShadow17
    With title.Fill
         .ForeColor.SchemeColor = 1
         .Visible = msoTrue
         .Solid
    End With
    
    title.TextFrame.Characters.Text = "Χ²검정에 대한 " & c & "% 신뢰구간"
    With title.TextFrame.Characters.Font
        .size = 12
        .ColorIndex = xlAutomatic
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
    Set ttemp = ttemp.Offset(3, 0)
    ttemp.Value = c & "% 신뢰구간"
    ttemp.Offset(0, 1) = c & "% 하한"
    ttemp.Offset(0, 2) = c & "% 상한"
    ttemp.Offset(1, 1) = Format(ll, "0.0000")
    ttemp.Offset(1, 2) = Format(ul, "0.0000")
    With ttemp.Resize(, 3).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With ttemp.Resize(, 3).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With ttemp.Offset(1, 0).Resize(, 3).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
 
    '''''결과값해석
    '''''
    Set ttemp = ttemp.Offset(4, 0)
    yp = ttemp.Top
    Set title = OutputSheet.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 250, 22#)
    title.Shadow.Type = msoShadow17
    With title.Fill
         .ForeColor.SchemeColor = 1
         .Visible = msoTrue
         .Solid
    End With
    
    title.TextFrame.Characters.Text = "요약 및 결론"
    With title.TextFrame.Characters.Font
        .size = 12
        .ColorIndex = xlAutomatic
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
    Set ttemp = ttemp.Offset(3, 0)
     ttemp.Value = "유의수준 α=" & za & "이고 자유도가" & n & "- 1 = " & n - 1 & "이다."
     ttemp.Offset(1, 0) = "검정통계량을 구하면" & Format((n - 1) * s / p, "##0.000") & "이다."
     ttemp.Offset(2, 0) = "해당 X²값을 카이제곱표에서 찾으면 기각역 = " & Format(WorksheetFunction.ChiInv(za, n - 1), "##0.0000") & "가 나온다."
    If (n - 1) * s / p > Format(WorksheetFunction.ChiInv(za, n - 1), "##0.0000") Then
        ttemp.Offset(4, 0) = " 검정통계량 > 기각역이므로 귀무가설이 기각된다. "
    Else
         ttemp.Offset(4, 0) = " 검정통계량 < 기각역이므로 귀무가설이 채택된다. "
    End If
        With ttemp

        .HorizontalAlignment = xlLeft
         ttemp.Offset(1, 0).HorizontalAlignment = xlLeft
         ttemp.Offset(2, 0).HorizontalAlignment = xlLeft
         ttemp.Offset(3, 0).HorizontalAlignment = xlLeft
         ttemp.Offset(4, 0).HorizontalAlignment = xlLeft
        End With
        

      Set ttemp = ttemp.Offset(5, 0)
    With ttemp
        .Value = Str
        .HorizontalAlignment = xlCenter
    End With
    Set ttemp = ttemp.Offset(2, -1)
    '''addr.Value = ttemp.Address
    addr.Value = Right(ttemp.Address, Len(ttemp.Address) - 3)

End Sub
