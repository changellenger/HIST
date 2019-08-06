Attribute VB_Name = "TwoWay1_Result"
Sub aResult(sta, stb, se, dfa, dfb, dfe, outputsheet)
    Dim sst, sm As Double: Dim ttemp, addr As Range
    Dim Comment0, Comment1, Comment2 As String
    Dim fvalue, fvalue1, fvalue2 As Double
    Dim p_value, pvalue, pvalue1, pvalue2 As Double
    Dim i As Integer
    
    sm = sta + stb
    sst = sta + stb + se
    'If IsEmpty(outputsheet.Range("a1")) = True Then
    '   Set ttemp = outputsheet.Cells(2, 1)
    '   Set addr = outputsheet.Range("a1")
    'Else: Set addr = outputsheet.Range("a1")
    '     Set ttemp = outputsheet.Range(addr.Value)
    'End If
    
    Set addr = outputsheet.Range("a1")
    Set ttemp = outputsheet.Range("a" & addr.Value)
    
    
    fvalue = (sm / (dfa + dfb)) / (se / dfe)
    pvalue = Application.FDist(fvalue, dfa + dfb, dfe)
    fvalue1 = (sta / dfa) / (se / dfe)
    pvalue1 = Application.FDist(fvalue1, dfa, dfe)
    fvalue2 = (stb / dfb) / (se / dfe)
    pvalue2 = Application.FDist(fvalue2, dfb, dfe)
    Set ttemp = ttemp.Offset(1, 1)
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
         .Size = 11
         .ColorIndex = xlAutomatic
    End With
    Title.TextFrame.HorizontalAlignment = xlCenter
    Set ttemp = ttemp.Offset(2, 0)
    Set qq = ttemp.Offset(4, 0)
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
    If Frm2_way2.MultiPage1.Value = 1 Then
        ttemp.Value = Frm2_way2.ListBox2.list(0)
    Else
        ttemp.Value = Frm2_way2.TextBox2.Text
    End If
    ttemp.Offset(0, 1) = Format(sta, "0.0000")
    ttemp.Offset(0, 2) = Format(dfa, "0.0000")
    ttemp.Offset(0, 3).Value = Format(sta / dfa, "0.0000")
    ttemp.Offset(0, 4) = Format(fvalue1, "0.0000")
    ttemp.Offset(0, 5) = Format(pvalue1, "0.0000")
    Set ttemp = ttemp.Offset(1, 0)
    If Frm2_way2.MultiPage1.Value = 1 Then
        ttemp.Value = Frm2_way2.ListBox3.list(0)
    Else
        ttemp.Value = Frm2_way2.TextBox3.Text
    End If
    ttemp.Offset(0, 1) = Format(stb, "0.0000")
    ttemp.Offset(0, 2) = Format(dfb, "0.0000")
    ttemp.Offset(0, 3).Value = Format(stb / dfb, "0.0000")
    ttemp.Offset(0, 4) = Format(fvalue2, "0.0000")
    ttemp.Offset(0, 5) = Format(pvalue2, "0.0000")
    Set ttemp = ttemp.Offset(1, 0)
    ttemp.Value = "잔차"
    ttemp.Offset(0, 1) = Format(se, "0.0000")
    ttemp.Offset(0, 2) = Format(dfe, "0.0000")
    ttemp.Offset(0, 3) = Format(se / dfe, "0.0000")
    Set ttemp = ttemp.Offset(1, 0)
    ttemp.Value = "계"
    ttemp.Offset(0, 1) = Format(sst, "0.0000")
    ttemp.Offset(0, 2) = dfa + dfb + dfe
    
    If Frm2_outoption.CheckBox1.Value = True Or Frm2_outoption.CheckBox2.Value = True Then
    Set ttemp = ttemp.Offset(1, 0)
    Comment1 = "반복이 없는 경우의 이원배치에서는 제곱합들의 값이 일치합니다."
    End If
           With ttemp
            .Value = Comment1
            .Font.Size = 9
            .HorizontalAlignment = xlLeft
        End With
'    p_value = pvalue1: Comment0 = "행인자의 "
'    For i = 1 To 2
'        If p_value <= 0.01 Then
'            Comment1 = """H0:" & Comment0 & "모평균들이 서로 같다.""" & "를 유의수준 α=0.01에서 기각한다."
'            Comment2 = "즉, " & Comment0 & "표본평균들이 아주 뚜렷한(p<0.01) 차이가 있다."
'        ElseIf p_value <= 0.05 Then
'            Comment1 = """H0:" & Comment0 & "모평균들이 서로 같다.""" & "를 유의수준 α=0.05에서 기각한다."
'            Comment2 = "즉, " & Comment0 & "표본평균들이 뚜렷한(p<0.05) 차이가 있다."
'        Else
'            Comment1 = """H0:" & Comment0 & "모평균들이 서로 같다.""" & "를 유의수준 α=0.05에서 기각할 수 없다."
'            Comment2 = "즉, " & Comment0 & "표본평균들이 차이가 있다(p<0.05)고 할 수 없다."
'        End If
'        With ttemp
'            .Value = Comment1
'            .Font.Size = 9
'            .HorizontalAlignment = xlLeft
'        End With
'        Set ttemp = ttemp.Offset(1, 0)
'        With ttemp
'            .Value = Comment2
'            .Font.Size = 9
'            .HorizontalAlignment = xlLeft
'        End With
'        p_value = pvalue2: Comment0 = "열인자의 "
'        Set ttemp = ttemp.Offset(1, 0)
'    Next i
    
    Set ttemp = ttemp.Offset(3, -1)
    '''addr.Value = ttemp.Address
    addr.Value = Right(ttemp.Address, Len(ttemp.Address) - 3)
End Sub
Sub dResult(rave, cave, rst, cst, rcnt, ccnt, rv, cv, outputsheet)
   Dim ttemp As Range
   Dim addr As Range
   Dim yp As Double
   'If IsEmpty(outputsheet.Range("a1")) = True Then
   '    Set ttemp = outputsheet.Cells(2, 1)
   '    Set addr = outputsheet.Range("a1")
   'Else: Set addr = outputsheet.Range("a1")
   '      Set ttemp = outputsheet.Range(addr.Value)
   'End If
   
   
   Set addr = outputsheet.Range("a1")
   Set ttemp = outputsheet.Range("a" & addr.Value)
   yp = ttemp.Top
   Set Title = outputsheet.Shapes.AddShape(msoShapeRectangle, 3.75, yp + 2.25, 400, 25#)
   Title.TextFrame.Characters.Text = "반복이 없는 이원배치 분산분석 결과"
    With Title
        .Fill.ForeColor.SchemeColor = 57
        .Line.DashStyle = msoLineSolid
        .Line.Style = msoLineSingle
        .Line.Weight = 1
        .Line.Visible = msoTrue
     '   .Shadow.Type = msoShadow1
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
   Set qq = ttemp.Offset(rcnt + ccnt + 1, 0)
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
   
   ttemp.Offset(0, 1).Value = "관측도수"
   ttemp.Offset(0, 2).Value = "평균"
   ttemp.Offset(0, 3).Value = "표준편차"
   For i = 1 To rcnt
       Set ttemp = ttemp.Offset(1, 0)
       ttemp.Value = rv(i)
       ttemp.Offset(0, 1).Value = ccnt
       ttemp.Offset(0, 2).Value = Format(rave(i), "0.0000")
       ttemp.Offset(0, 3).Value = Format(rst(i), "0.0000")
   Next i
   Set ttemp = ttemp.Offset(1, 0)
   For i = 1 To ccnt
       Set ttemp = ttemp.Offset(1, 0)
       ttemp.Value = cv(i)
       ttemp.Offset(0, 1).Value = rcnt
       ttemp.Offset(0, 2).Value = Format(cave(i), "0.0000")
       ttemp.Offset(0, 3).Value = Format(cst(i), "0.0000")
   Next i
   Set ttemp = ttemp.Offset(2, -1)
   '''addr.Value = ttemp.Address
   addr.Value = Right(ttemp.Address, Len(ttemp.Address) - 3)
End Sub
