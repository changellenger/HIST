Attribute VB_Name = "TwoWay2_Result"
Sub sResult(sst, sta, stb, stab, se, dfa, dfb, dfab, dfe, outputsheet, ListRow, ListCol, a)
    Dim sm As Double: Dim ttemp, addr As Range
    Dim Comment0, Comment1, Comment2 As String
    Dim fvalue, fvalue1, fvalue2, fvalue3 As Double
    Dim p_value, pvalue, pvalue1, pvalue2, pvalue3 As Double
    Dim i, chflag As Integer
    
    sm = sta + stb + stab
    
    Set addr = outputsheet.Range("a1")
    Set ttemp = outputsheet.Range("a" & addr.Value)
    
    fvalue = (sm / (dfa + dfb + dfab)) / (se / dfe)
    pvalue = Application.FDist(fvalue, dfa + dfb + dfab, dfe)
    fvalue1 = (sta / dfa) / (se / dfe)
    pvalue1 = Application.FDist(fvalue1, dfa, dfe)
    fvalue2 = (stb / dfb) / (se / dfe)
    pvalue2 = Application.FDist(fvalue2, dfb, dfe)
    
        Set ttemp = ttemp.Offset(1, 1)
        yp = ttemp.Top
        Set Title = outputsheet.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 65, 20#)
        Title.Shadow.Type = msoShadow17
        With Title.Fill
          .ForeColor.SchemeColor = 22
          .Visible = msoTrue
          .Solid
        End With
        Title.TextFrame.Characters.Text = "Type " & a & "  SS"
        With Title.TextFrame.Characters.Font
            .Size = 11
            .ColorIndex = xlAutomatic
        End With
        Title.TextFrame.HorizontalAlignment = xlCenter
        Set ttemp = ttemp.Offset(2, 0)
        
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
        If Frm2_way2.CheckBox2 = False Then
           chflag = 1
        Else: chflag = 0
        End If
        
        ttemp.Value = "요인"
        ttemp.Offset(0, 1) = "제곱합"
        ttemp.Offset(0, 2) = "자유도"
        ttemp.Offset(0, 3) = "평균제곱"
        ttemp.Offset(0, 4) = "F값"
        ttemp.Offset(0, 5) = "유의확률"
        
        If Unbal = False Then
            Set ttemp = ttemp.Offset(1, 0)
            ttemp.Value = ListRow
            ttemp.Offset(0, 1) = Format(sta, "0.0000")
            ttemp.Offset(0, 2) = Format(dfa, "0.0000")
            ttemp.Offset(0, 3).Value = Format(sta / dfa, "0.0000")
            ttemp.Offset(0, 4) = Format(fvalue1, "0.0000")
            ttemp.Offset(0, 5) = Format(pvalue1, "0.0000")
            Set ttemp = ttemp.Offset(1, 0)
            ttemp.Value = ListCol
            ttemp.Offset(0, 1) = Format(stb, "0.0000")
            ttemp.Offset(0, 2) = Format(dfb, "0.0000")
            ttemp.Offset(0, 3).Value = Format(stb / dfb, "0.0000")
            ttemp.Offset(0, 4) = Format(fvalue2, "0.0000")
            ttemp.Offset(0, 5) = Format(pvalue2, "0.0000")
            If Frm2_way2.CheckBox2 = False Then
                fvalue3 = (stab / dfab) / (se / dfe)
                pvalue3 = Application.FDist(fvalue3, dfab, dfe)
                Set ttemp = ttemp.Offset(1, 0)
                ttemp.Value = "교호작용"
                ttemp.Offset(0, 1) = Format(stab, "0.0000")
                ttemp.Offset(0, 2) = Format(dfab, "0.0000")
                ttemp.Offset(0, 3).Value = Format(stab / dfab, "0.0000")
                ttemp.Offset(0, 4) = Format(fvalue3, "0.0000")
                ttemp.Offset(0, 5) = Format(pvalue3, "0.0000")
            End If
        End If
        
        If Unbal = True Then
            Set ttemp = ttemp.Offset(1, 0)
            ttemp.Value = "model"
            ttemp.Offset(0, 1) = Format(sta, "0.0000")
            ttemp.Offset(0, 2) = Format(dfa, "0.0000")
            ttemp.Offset(0, 3).Value = Format(sta / dfa, "0.0000")
            ttemp.Offset(0, 4) = Format(fvalue1, "0.0000")
            ttemp.Offset(0, 5) = Format(pvalue1, "0.0000")
        End If
        
        Set qq = ttemp
        With qq.Resize(, 6).Borders(xlEdgeBottom)
             .LineStyle = xlContinuous
             .Weight = xlMedium
             .ColorIndex = xlAutomatic
        End With

    Set ttemp = ttemp.Offset(3, -1)
    '''addr.Value = ttemp.Address
    addr.Value = Right(ttemp.Address, Len(ttemp.Address) - 3)
End Sub
Sub a1Result(sta, stb, stab, se, dfa, dfb, dfab, dfe, outputsheet)
    Dim sst, sm As Double: Dim ttemp, addr As Range
    Dim Comment0, Comment1, Comment2 As String
    Dim fvalue, fvalue1, fvalue2, fvalue3 As Double
    Dim p_value, pvalue, pvalue1, pvalue2, pvalue3 As Double
    Dim i, chflag As Integer
    
    sm = sta + stb + stab
    sst = sta + stb + stab + se
    'If IsEmpty(outputsheet.Range("a1")) = True Then
    '   Set ttemp = outputsheet.Cells(2, 1)
    '   Set addr = outputsheet.Range("a1")
    'Else: Set addr = outputsheet.Range("a1")
    '     Set ttemp = outputsheet.Range(addr.Value)
    'End If
    
    Set addr = outputsheet.Range("a1")
    Set ttemp = outputsheet.Range("a" & addr.Value)
    
    fvalue = (sm / (dfa + dfb + dfab)) / (se / dfe)
    pvalue = Application.FDist(fvalue, dfa + dfb + dfab, dfe)
    fvalue1 = (sta / dfa) / (se / dfe)
    pvalue1 = Application.FDist(fvalue1, dfa, dfe)
    fvalue2 = (stb / dfb) / (se / dfe)
    pvalue2 = Application.FDist(fvalue2, dfb, dfe)
    If Frm2_way2.TextBox1.Value = True Then
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
        If Frm2_way2.chkinteract = True Then
           chflag = 1
        Else: chflag = 0
        End If
        Set qq = ttemp.Offset(4 + chflag, 0)
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
        If Frm2_way2.chkinteract = True Then
            fvalue3 = (stab / dfab) / (se / dfe)
            pvalue3 = Application.FDist(fvalue3, dfab, dfe)
            Set ttemp = ttemp.Offset(1, 0)
            ttemp.Value = "교호작용"
            ttemp.Offset(0, 1) = Format(stab, "0.0000")
            ttemp.Offset(0, 2) = Format(dfab, "0.0000")
            ttemp.Offset(0, 3).Value = Format(stab / dfab, "0.0000")
            ttemp.Offset(0, 4) = Format(fvalue3, "0.0000")
            ttemp.Offset(0, 5) = Format(pvalue3, "0.0000")
        End If
        Set ttemp = ttemp.Offset(1, 0)
        ttemp.Value = "잔차"
        ttemp.Offset(0, 1) = Format(se, "0.0000")
        ttemp.Offset(0, 2) = Format(dfe, "0.0000")
        ttemp.Offset(0, 3) = Format(se / dfe, "0.0000")
        Set ttemp = ttemp.Offset(1, 0)
        ttemp.Value = "계"
        ttemp.Offset(0, 1) = Format(sst, "0.0000")
        ttemp.Offset(0, 2) = Format(dfa + dfb + dfab + dfe, "0.0000")
    End If
    

    Set ttemp = ttemp.Offset(3, -1)
    '''addr.Value = ttemp.Address
    addr.Value = Right(ttemp.Address, Len(ttemp.Address) - 3)
End Sub
Sub aResult(sst, N, sta, stb, stab, se, dfa, dfb, dfab, dfe, outputsheet, ListRow, ListCol, a, Optional Unbal As Boolean = True)
    Dim sm As Double: Dim ttemp, addr As Range
    Dim Comment0, Comment1, Comment2 As String
    Dim fvalue, fvalue1, fvalue2, fvalue3 As Double
    Dim p_value, pvalue, pvalue1, pvalue2, pvalue3 As Double
    Dim i, chflag As Integer
    
    sm = sta + stb + stab

    Set addr = outputsheet.Range("a1")
    Set ttemp = outputsheet.Range("a" & addr.Value)
    
    fvalue = (sm / (dfa + dfb + dfab)) / (se / dfe)
    pvalue = Application.FDist(fvalue, dfa + dfb + dfab, dfe)
    fvalue1 = (sta / dfa) / (se / dfe)
    pvalue1 = Application.FDist(fvalue1, dfa, dfe)
    fvalue2 = (stb / dfb) / (se / dfe)
    pvalue2 = Application.FDist(fvalue2, dfb, dfe)
    
'    If Frm2_way2.TextBox1.Value = True Then
        Set ttemp = ttemp.Offset(1, 1)
        yp = ttemp.Top
        Set Title = outputsheet.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 250, 20#)
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
        If Frm2_way2.CheckBox2 = False Then
           chflag = 1
        Else: chflag = 0
        End If
        
        ttemp.Value = "요인"
        ttemp.Offset(0, 1) = "제곱합"
        ttemp.Offset(0, 2) = "자유도"
        ttemp.Offset(0, 3) = "평균제곱"
        ttemp.Offset(0, 4) = "F값"
        ttemp.Offset(0, 5) = "유의확률"
        
        If Unbal = False Then
            Set ttemp = ttemp.Offset(1, 0)
            ttemp.Value = ListRow
            ttemp.Offset(0, 1) = Format(sta, "0.0000")
            ttemp.Offset(0, 2) = Format(dfa, "0.0000")
            ttemp.Offset(0, 3).Value = Format(sta / dfa, "0.0000")
            ttemp.Offset(0, 4) = Format(fvalue1, "0.0000")
            ttemp.Offset(0, 5) = Format(pvalue1, "0.0000")
            Set ttemp = ttemp.Offset(1, 0)
            ttemp.Value = ListCol
            ttemp.Offset(0, 1) = Format(stb, "0.0000")
            ttemp.Offset(0, 2) = Format(dfb, "0.0000")
            ttemp.Offset(0, 3).Value = Format(stb / dfb, "0.0000")
            ttemp.Offset(0, 4) = Format(fvalue2, "0.0000")
            ttemp.Offset(0, 5) = Format(pvalue2, "0.0000")
            If Frm2_way2.CheckBox2 = False Then
                fvalue3 = (stab / dfab) / (se / dfe)
                pvalue3 = Application.FDist(fvalue3, dfab, dfe)
                Set ttemp = ttemp.Offset(1, 0)
                ttemp.Value = "교호작용"
                ttemp.Offset(0, 1) = Format(stab, "0.0000")
                ttemp.Offset(0, 2) = Format(dfab, "0.0000")
                ttemp.Offset(0, 3).Value = Format(stab / dfab, "0.0000")
                ttemp.Offset(0, 4) = Format(fvalue3, "0.0000")
                ttemp.Offset(0, 5) = Format(pvalue3, "0.0000")
            End If
        End If
        
        If Unbal = True Then
            Set ttemp = ttemp.Offset(1, 0)
            ttemp.Value = "model"
            ttemp.Offset(0, 1) = Format(sta, "0.0000")
            ttemp.Offset(0, 2) = Format(dfa, "0.0000")
            ttemp.Offset(0, 3).Value = Format(sta / dfa, "0.0000")
            ttemp.Offset(0, 4) = Format(fvalue1, "0.0000")
            ttemp.Offset(0, 5) = Format(pvalue1, "0.0000")
        End If
        
        Set ttemp = ttemp.Offset(1, 0)
        ttemp.Value = "잔차"
        ttemp.Offset(0, 1) = Format(se, "0.0000")
        ttemp.Offset(0, 2) = Format(dfe, "0.0000")
        ttemp.Offset(0, 3) = Format(se / dfe, "0.0000")
        
        Set ttemp = ttemp.Offset(1, 0)
        ttemp.Value = "계"
        ttemp.Offset(0, 1) = Format(sst, "0.0000")
        ttemp.Offset(0, 2) = Format(N - 1, "0.0000")
                
        Set qq = ttemp
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


    
    Set ttemp = ttemp.Offset(3, -1)
    '''addr.Value = ttemp.Address
    addr.Value = Right(ttemp.Address, Len(ttemp.Address) - 3)
End Sub

Sub dResult(ave, st, cccnt, mrave, mrst, mcave, mcst, rname, cname, rcnt, ccnt, tmean, tstd, outputsheet, list)
    Dim ttemp As Range
    Dim addr As Range
    Dim qq, qqq As Range
    Dim yp As Double
    Dim sum As Integer
    
    'If IsEmpty(outputsheet.Range("a1")) = True Then
    '   Set ttemp = outputsheet.Cells(2, 1)
    '   Set addr = outputsheet.Range("a1")
    'Else: Set addr = outputsheet.Range("a1")
    '      Set ttemp = outputsheet.Range(addr.Value)
    'End If
    
    Set addr = outputsheet.Range("a1")
    Set ttemp = outputsheet.Range("a" & addr.Value)
    
    yp = ttemp.Top
    Set Title = outputsheet.Shapes.AddShape(msoShapeRectangle, 3.75, yp + 2.25, 300, 25#)
    Title.TextFrame.Characters.Text = "반복이 있는 이원배치 분산분석 결과"
    With Title
        .Fill.ForeColor.SchemeColor = 9
        .Line.DashStyle = msoLineSolid
        .Line.Style = msoLineSingle
        .Line.Weight = 1
        .Line.Visible = msoTrue
        .Shadow.Type = msoShadow1
    End With
    With Title.TextFrame.Characters.Font
         .Size = 14
         .ColorIndex = 41
    End With
    Title.TextFrame.HorizontalAlignment = xlCenter
    
    Set ttemp = ttemp.Offset(3, 1)
    yp = ttemp.Top
    Set Title = outputsheet.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 80, 22#)
    Title.Shadow.Type = msoShadow17
    With Title.Fill
        .ForeColor.SchemeColor = 22
        .Visible = msoTrue
        .Solid
    End With
    Title.TextFrame.Characters.Text = "기술 통계량"
    With Title.TextFrame.Characters.Font
         .Size = 11
         .ColorIndex = xlAutomatic
    End With
    Title.TextFrame.HorizontalAlignment = xlCenter
    
    Set qq = ttemp.Offset(2, 0)
    For i = 1 To ccnt
        qq.Offset(0, i).Value = cname(i)
        With qq.Resize((rcnt + 1) * 4 + 1, 1).Borders(xlEdgeRight)
             .LineStyle = xlContinuous
             .Weight = xlMedium
             .ColorIndex = xlAutomatic
        End With
        With qq.Offset(0, i).Resize((rcnt + 1) * 4 + 1, 1).Borders(xlEdgeRight)
             .LineStyle = xlContinuous
             .Weight = xlThin
             .ColorIndex = xlAutomatic
        End With
    Next i
    With qq.Resize(, ccnt + 2).Borders(xlEdgeBottom)
           .LineStyle = xlContinuous
           .Weight = xlMedium
           .ColorIndex = xlAutomatic
    End With
    With qq.Offset(0, ccnt + 1)
        .Value = "계"
        .Interior.Color = 16049112
        .Interior.Pattern = xlSolid
    End With
    For i = 1 To rcnt + 1
        sum = 0
        Set ttemp = ttemp.Offset(4, 0)
        Set qq = ttemp.Offset(2, 0)
    '    With ttemp.Resize(, 4).Borders(xlEdgeBottom)
    '         .LineStyle = xlContinuous
    '         .Weight = xlMedium
    '         .ColorIndex = xlAutomatic
    '    End With
        With qq.Resize(, ccnt + 2).Borders(xlEdgeBottom)
             .LineStyle = xlContinuous
             .Weight = xlMedium
             .ColorIndex = xlAutomatic
        End With
        Set aa = ttemp.Offset(-1, 0)
        If i <> rcnt + 1 Then
           aa.Value = rname(i)
        Else: aa.Value = "계"
              Set qqq = aa.Resize(4, 1)
              With qqq.Interior
                  .Color = 16049112
                  .Pattern = xlSolid
              End With
        End If
        With aa.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        ttemp.Value = "관측도수"
        ttemp.Offset(1, 0).Value = "평균"
        ttemp.Offset(2, 0).Value = "표준편차"
        If i <> rcnt + 1 Then
           For J = 1 To ccnt + 1
              If J <> ccnt + 1 Then
                 Set qq = ttemp.Offset(0, J)
                 If list = 1 Then
                 qq.Value = cccnt
                 sum = sum + cccnt
                 Else
                 qq.Value = cccnt(i, J)
                 sum = sum + cccnt(i, J)
                 End If
                 
                 qq.Offset(1, 0).Value = Format(ave(i, J), "0.0000")
                 qq.Offset(2, 0).Value = Format(st(i, J), "0.0000")
              Else: Set qq = ttemp.Offset(0, ccnt + 1)
              If list = 1 Then
                    qq.Value = sum
              Else
                    qq.Value = cccnt(i, J)
              End If
                    qq.Offset(1, 0).Value = Format(mrave(i), "0.0000")
                    qq.Offset(2, 0).Value = Format(mrst(i), "0.0000")
                    Set qqq = qq.Offset(-1, 0).Resize(4, 1)
                    With qqq.Interior
                      .Color = 16049112
                      .Pattern = xlSolid
                    End With
              End If
           Next J
        Else: For J = 1 To ccnt
                 Set qq = ttemp.Offset(0, J)
                 sum = 0
                 If list = 1 Then
                 For k = 1 To rcnt
                 sum = sum + ttemp.Offset(-4 * k, 1)
                 Next k
                 qq.Value = sum
                 Else
                 qq.Value = cccnt(i, J)
                 End If
                 qq.Offset(1, 0).Value = Format(mcave(J), "0.0000")
                 qq.Offset(2, 0).Value = Format(mcst(J), "0.0000")
                 Set qqq = qq.Offset(-1, 0).Resize(4, 1)
                 With qqq.Interior
                      .Color = 16049112
                      .Pattern = xlSolid
                 End With
              Next J
        End If
    Next i
    Set qq = ttemp.Offset(0, ccnt + 1)
    Dim cou As Integer
    cou = 0
    
    For i = 1 To ccnt
     cou = cou + qq.Offset(0, -i).Value
    Next i
    qq.Value = cou
    qq.Offset(1, 0).Value = Format(tmean, "0.0000")
    qq.Offset(2, 0).Value = Format(tstd, "0.0000")
    Set qqq = qq.Offset(-1, 0).Resize(4, 1)
    With qqq.Interior
         .Color = 16049112
         .Pattern = xlSolid
    End With
    Set ttemp = ttemp.Offset(4, -1)
    '''addr.Value = ttemp.Address
    addr.Value = Right(ttemp.Address, Len(ttemp.Address) - 3)
End Sub

Sub cResult(list, ave, name, ct, se, td, ed, fn, alpha, outputsheet, Check4 As Boolean, check8 As Boolean, check5 As Boolean)
'다중비교 sub
    Dim ttemp, addr As Range
    Dim dave(), temp(), temp1(), temp2(), q(), names(), fact1, ave1, c As Double
    Dim tvalue(), pvalue(), pvalue1() As Double
    Dim N, z, a, b, d, count As Integer
    Dim fact() As String
    Dim pave() As Double
    
    Set addr = outputsheet.Range("a1")
    Set ttemp = outputsheet.Range("a" & addr.Value)
    
    If fn <= 2 Then
    Set ttemp = ttemp.Offset(0, 1)
        yp = ttemp.Top
        Set Title = outputsheet.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 90, 20#)
        Title.Shadow.Type = msoShadow17
        With Title.Fill
            .ForeColor.SchemeColor = 22
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
            Comment2 = "인자의 수준수가 둘이하이므로 " & list & " 인자 에 대한 다중비교를 수행할수 없습니다."
        With ttemp
            .Value = Comment2
            .Font.Size = 9
            .HorizontalAlignment = xlLeft
        End With
        Set ttemp = ttemp.Offset(3, -1)
   
        addr.Value = Right(ttemp.Address, Len(ttemp.Address) - 3)
             
    Else
    
    ReDim fact(1 To fn)
    For i = 1 To fn
        fact(i) = name(i)
    Next i
    
    ReDim pave(1 To fn)
    For i = 1 To fn
        pave(i) = ave(i)
    Next i

      
      N = 0
      '평균을 통한 Sorting
    For i = 1 To fn
    For J = i + 1 To fn
    If pave(i) > pave(J) Then
    ave1 = pave(J)
    fact1 = fact(J)
    pave(J) = pave(i)
    fact(J) = fact(i)
    pave(i) = ave1
    fact(i) = fact1
    End If
    Next J
    Next i
    
    'Q-statistic과 t-statistic을 구하기 위한 For문
        For i = 1 To fn - 1
        For J = i + 1 To fn
        N = N + 1: ReDim Preserve dave(1 To N): ReDim Preserve q(1 To N): ReDim Preserve names(1 To N)
        ReDim Preserve pvalue(1 To N): ReDim Preserve tvalue(1 To N)
        dave(N) = pave(i) - pave(J)
        names(N) = pave(J)
        q(N) = Abs(pave(i) - pave(J)) / (((se / ed) * (1 / ct(i) + 1 / ct(J)) / 2) ^ 0.5)
        tvalue(N) = Abs(dave(N)) / (((se / ed) * (1 / ct(i) + 1 / ct(J))) ^ (0.5))
        pvalue(N) = Application.TDist(tvalue(N), ed, 2)
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
        Set ttemp = ttemp.Offset(0, 1)
        yp = ttemp.Top
        Set Title = outputsheet.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 95, 20#)
        Title.Shadow.Type = msoShadow17
        With Title.Fill
            .ForeColor.SchemeColor = 22
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
       ttemp.Offset(0, 2).Value = Format(pave(1), "0.0000")
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
       If a >= N Then
       b = b + 1
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
 '       yp = ttemp.Top
 '       Set Title = outputsheet.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 135, 22#)
 '       Title.Shadow.Type = msoShadow17
 '       With Title.Fill
 '           .ForeColor.SchemeColor = 22
 '           .Visible = msoTrue
 '           .Solid
 '       End With
 '       Title.TextFrame.Characters.Text = "다중비교 - " & list
 '       With Title.TextFrame.Characters.Font
 '           .name = "굴림"
 '           .FontStyle = "굵게"
 '           .Size = 11
 '           .ColorIndex = xlAutomatic
 '       End With
 '       Title.TextFrame.HorizontalAlignment = xlCenter
 '       Set ttemp = ttemp.Offset(2, 0)
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
       ttemp.Offset(0, 2).Value = Format(pave(1), "0.0000")
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
       
       If a >= N Then
       b = b + 1
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
'        Set Title = outputsheet.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 95, 22#)
'        Title.Shadow.Type = msoShadow17
'        With Title.Fill
'            .ForeColor.SchemeColor = 22
'            .Visible = msoTrue
'            .Solid
'        End With
'        Title.TextFrame.Characters.Text = "다중비교 - " & list
'        With Title.TextFrame.Characters.Font
'            .name = "굴림"
'            .FontStyle = "굵게"
'            .Size = 11
'            .ColorIndex = xlAutomatic
'        End With
 '       Title.TextFrame.HorizontalAlignment = xlCenter
 '       Set ttemp = ttemp.Offset(2, 0)
        
        'Tukey와 달리 Ducan은 comparisonwise error rate을 사용
        '이를 이용해서 다시 Ducan에 관한 P-value를 구하는 For 문
          For k = 1 To N
        pvalue1(k) = 1 - ((1 - pvalue1(k)) ^ (1 / fn)) + 0.01
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
       ttemp.Offset(0, 2).Value = Format(pave(1), "0.0000")
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
       c = c + 1
       End If
       If a >= N Then
       b = b + 1
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
