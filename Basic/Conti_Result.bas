Attribute VB_Name = "Conti_Result"
Sub cResult(rcnt, ccnt, table, exp, rt, ct, chi, rn, cn, outputsheet)
    Dim ttemp As Range
    Dim addr As Range
    Dim total As Long
    Dim warn As Long
    total = 0
    warn = 0
    'If IsEmpty(outputsheet.Range("a1")) = True Then
    '   Set ttemp = outputsheet.Cells(2, 1)
    '   Set addr = outputsheet.Range("a1")
    'Else: Set addr = outputsheet.Range("a1")
    '     Set ttemp = outputsheet.Range(addr.Value)
    'End If
    
    
    Set addr = outputsheet.Range("a1")
    Set ttemp = outputsheet.Range("a" & addr.Value + 1)
    
    yp = ttemp.Top
    Set title = outputsheet.Shapes.AddShape(msoShapeRectangle, 3.75, yp + 2.25, 400, 25#)
    title.TextFrame.Characters.Text = "�����м� ���"
    With title
        .Fill.ForeColor.SchemeColor = 57
        .Line.DashStyle = msoLineSolid
        .Line.Style = msoLineSingle
        .Line.Weight = 1
        .Line.Visible = msoTrue
     '   .Shadow.Type = msoShadow1
    End With
    With title.TextFrame.Characters.Font
        .Size = 14
        .ColorIndex = 2
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
    Set ttemp = ttemp.Offset(3, 1)
    yp = ttemp.Top
    Set title = outputsheet.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 250, 22#)
    title.Shadow.Type = msoShadow17
    With title.Fill
         .ForeColor.SchemeColor = 1
         .Visible = msoTrue
         .Solid
    End With
    title.TextFrame.Characters.Text = "�����м�ǥ"
    With title.TextFrame.Characters.Font
         .Size = 11
         .ColorIndex = xlAutomatic
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
    Set ttemp = ttemp.Offset(2, 0)
    Set qq = ttemp.Offset(1, 0)
    If frameCrfre.Expect.Value = True Then
       ee = 3
    Else: ee = 2
    End If
    For i = 1 To ccnt
        If ee = 3 Then
           z = -1
        Else: z = 0
        End If
        qq.Offset(0, i).Value = cn(i)
        With qq.Resize((rcnt + 1) * ee + z, 1).Borders(xlEdgeRight)
             .LineStyle = xlContinuous
             .Weight = xlMedium
             .ColorIndex = xlAutomatic
        End With
        With qq.Offset(0, i).Resize((rcnt + 1) * ee + z, 1).Borders(xlEdgeRight)
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
    qq.Offset(0, ccnt + 1).Value = "��"
    For i = 1 To rcnt + 1
        If i <> rcnt + 1 And ee = 3 Then
           Set ttemp = ttemp.Offset(3, 0)
        Else: Set ttemp = ttemp.Offset(2, 0)
        End If
        If i = rcnt + 1 Then
           s = 0
        Else: s = 1
        End If
        
        If ee = 3 Then
           Set aa = ttemp.Offset(-1, 0)
        Else: Set aa = ttemp
        End If
        With aa.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        If i <> rcnt + 1 Then
           aa.Value = rn(i)
        ElseIf ee = 3 Then
              aa.Offset(1, 0).Value = "��"
        ElseIf ee = 2 Then
              aa.Value = "��"
        
        End If
        
        Set qq = ttemp.Offset(s, 0)
        With qq.Resize(, ccnt + 2).Borders(xlEdgeBottom)
             .LineStyle = xlContinuous
             .Weight = xlMedium
             .ColorIndex = xlAutomatic
        End With
        
        If i <> rcnt + 1 And ee = 3 Then
           ttemp.Value = "��������"
        ElseIf i <> rcnt + 1 And ee = 2 Then
           ttemp.Offset(1, 0) = "��������"
        End If
        If ee = 3 And i <> rcnt + 1 Then
          ttemp.Offset(1, 0).Value = "��뵵��"
        End If
        If i <> rcnt + 1 Then
           For J = 1 To ccnt + 1
              If J <> ccnt + 1 Then
                 If ee = 2 Then
                    k = 1
                 Else: k = 0
                 End If
                 Set qq = ttemp.Offset(k, J)
                 qq.Value = table.Cells(i, J).Value
                 If exp(i, J) < 5 Then
                    warn = warn + 1
                 End If
                 If ee = 3 Then
                    qq.Offset(1, 0).Value = Format(exp(i, J), "0.0000")
                 End If
              Else: Set qq = ttemp.Offset(0, ccnt + 1)
                    qq.Value = rt(i)
              End If
           Next J
        Else: For J = 1 To ccnt + 1
                 Set qq = ttemp.Offset(0, J)
                 If J <> ccnt + 1 Then
                    qq.Value = ct(J)
                    total = total + ct(J)
                 Else: qq.Value = total
                 End If
              Next J
        End If
    Next i
    Dim pvalue As Double
    Set ttemp = ttemp.Offset(2, 0)
    ttemp.Value = "ī������ ��跮 : " & Format(chi, "0.0000")
    ttemp.HorizontalAlignment = xlGeneral
    Set ttemp = ttemp.Offset(1, 0)
    pvalue = Application.WorksheetFunction.ChiDist(chi, (rcnt - 1) * (ccnt - 1))
    ttemp.Value = "����Ȯ�� : " & Format(pvalue, "0.00000")
    ttemp.HorizontalAlignment = xlGeneral
    If warn <> 0 Then
       Set ttemp = ttemp.Offset(1, 0)
       Dim per As Double
       per = warn / (rcnt * ccnt)
       ttemp.Value = Format(per * 100, "0.0000") & "%�� ���� ��뵵���� 5���� �۽��ϴ�."
       ttemp.HorizontalAlignment = xlGeneral
    End If
    Set ttemp = ttemp.Offset(4, -1)
    '''addr.Value = ttemp.Address
    
    addr.Value = right(ttemp.Address, Len(ttemp.Address) - 3)
End Sub
