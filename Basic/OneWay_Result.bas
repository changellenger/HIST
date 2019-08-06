Attribute VB_Name = "OneWay_Result"

Sub eResult(ave, vrn, st, ct, fn, outputsheet)
'��л� ���� sub (Levene's test)
    Dim ttemp, addr As Range
    Dim temp, temp1 As Double
    Dim s1, xsq, xsq1, xsq2, w As Double
    Dim z(), s() As Double
    Dim pvalue As Double
    Dim N, k As Long
    Dim one, ind As Integer
    
    temp = 0
    one = 0 '���ؼ��� 3������ ������ ����
    ind = 1 '���ؼ��� 3������ ���ڵ��� ����� ���� ����
    
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
    'Levene's test �� ����
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
    
    ind = 1 '���ؼ��� 3������ ���ڵ��� ����� ���� ���� �ʱ�ȭ
    
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
    
        
        
        
        
    
    Set addr = outputsheet.Range("a1")
    Set ttemp = outputsheet.Range("a" & addr.Value + 1)
    
    yp = ttemp.Top
    Set title = outputsheet.Shapes.AddShape(msoShapeRectangle, 3.75, yp + 2.25, 400, 25#)
    title.TextFrame.Characters.Text = "��л���� ���"
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
        Set title = outputsheet.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 250, 20#)
        title.Shadow.Type = msoShadow17
        With title.Fill
            .ForeColor.SchemeColor = 1
            .Visible = msoTrue
            .Solid
        End With
        title.TextFrame.Characters.Text = "��л� ����"
        With title.TextFrame.Characters.Font
            .Name = "����"
            .FontStyle = "����"
            .Size = 11
            .ColorIndex = xlAutomatic
        End With
        
        title.TextFrame.HorizontalAlignment = xlCenter
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
        ttemp.Offset(0, 1) = "������"
        ttemp.Offset(0, 2) = "������"
        ttemp.Offset(0, 3) = "�������"
        ttemp.Offset(0, 4) = "F��"
        ttemp.Offset(0, 5) = "����Ȯ��"
      
        Set ttemp = ttemp.Offset(1, 0)
        ttemp.Value = "ó��"
        ttemp.Offset(0, 1) = Format(xsq2, "0.0000")
        ttemp.Offset(0, 2) = Format(fn - 1, "0.0000")
        ttemp.Offset(0, 3) = Format(xsq2 / (fn - 1), "0.0000")
        ttemp.Offset(0, 4) = Format(w, "0.0000")
        ttemp.Offset(0, 5) = Format(pvalue, "0.0000")
         
        Set ttemp = ttemp.Offset(1, 0)
        ttemp.Value = "����"
        ttemp.Offset(0, 1) = Format(xsq - xsq1, "0.0000")
        ttemp.Offset(0, 2) = Format(N - fn, "0.0000")
        ttemp.Offset(0, 3) = Format((xsq - xsq1) / (N - fn), "0.0000")
                
        With ttemp.Resize(, 6).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
            End With
        Set ttemp = ttemp.Offset(1, 0)
            Comment1 = " ����Ȯ�� ���� ���Ǽ��� �� ���� ������ ��л� ������ �������� ������ �ǹ��Ѵ�."
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
        addr.Value = right(ttemp.Address, Len(ttemp.Address) - 3)
        
        
    Else
    yp = ttemp.Top
    Set title = outputsheet.Shapes.AddShape(msoShapeRectangle, 3.75, yp + 2.25, 400, 25#)
    title.TextFrame.Characters.Text = "��л���� ���"
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
        Set title = outputsheet.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 250, 20#)
        title.Shadow.Type = msoShadow17
        With title.Fill
            .ForeColor.SchemeColor = 1
            .Visible = msoTrue
            .Solid
        End With
        title.TextFrame.Characters.Text = "��л� ����"
        With title.TextFrame.Characters.Font
            .Name = "����"
            .FontStyle = "����"
            .Size = 11
            .ColorIndex = xlAutomatic
        End With
        
        title.TextFrame.HorizontalAlignment = xlCenter
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
        
        ttemp.Offset(0, 4) = "���ؼ��� 1�� ���ڰ� �־ Levene's test�� �Ҽ� �����ϴ�."
        Set ttemp = ttemp.Offset(3, -1)
        addr.Value = right(ttemp.Address, Len(ttemp.Address) - 3)
     End If
        
End Sub
