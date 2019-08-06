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
    
    title.TextFrame.Characters.Text = "������� ���� z-����"

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
    title.TextFrame.Characters.Text = "����� ���� "
    With title.TextFrame.Characters.Font
        .size = 12
        .ColorIndex = xlAutomatic
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
    Set ttemp = ttemp.Offset(3, 0)
    ttemp.Value = "����Ƚ��"
    ttemp.Offset(0, 1) = "����Ƚ��"
    ttemp.Offset(0, 3) = "������ ����ġ"
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
    title.TextFrame.Characters.Text = "H0 : p=" & p & " �� ��������"
    With title.TextFrame.Characters.Font
        .size = 12
        .ColorIndex = xlAutomatic
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
    Set ttemp = ttemp.Offset(3, 2)
    Select Case Hyp
    Case 1
       ttemp.Value = " H0 : p = " & p & " vs. H1 : p �� " & p
    Case 2
       ttemp.Value = " H0 : p = " & p & " vs. H1 : p > " & p
    Case 3
       ttemp.Value = " H0 : p = " & p & " vs. H1 : p < " & p
    End Select
        
    
    Set ttemp = ttemp.Offset(3, -2)
    ttemp.Value = "������跮"
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
 
  '  ttemp.Offset(1, 0) = "����Ȯ��"
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
       
    
    
    
    
    title.TextFrame.Characters.Text = "������� ���� " & c & "% �ŷڱ���"
    With title.TextFrame.Characters.Font
        .size = 12
        .ColorIndex = xlAutomatic
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
    Set ttemp = ttemp.Offset(3, 0)
    ttemp.Value = c & "% �ŷڱ���"
    ttemp.Offset(0, 1) = "����"
    ttemp.Offset(0, 2) = "����"
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
   
    title.TextFrame.Characters.Text = "��� �� ���"
    With title.TextFrame.Characters.Font
        .size = 12
        .ColorIndex = xlAutomatic
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
     Set ttemp = ttemp.Offset(3, 0)
   ttemp.Value = "n = " & n & ", p0= " & Format(hatp, "##0.000") & "�̰�, np0= " & n * hatp & ", n(1 - p0) = " & n * (1 - hatp) & "�� ��� 5���� ���Աٻ� ������ ������ �� �ִ�."
   ttemp.Offset(1, 0) = "������跮�� " & Format(Z, "0.0000") & "�̴�. "
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
    
    title.TextFrame.Characters.Text = "�� ������� ���� z-����"
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
    title.TextFrame.Characters.Text = "����� ���� "
    With title.TextFrame.Characters.Font
        .size = 12
        .ColorIndex = xlAutomatic
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
    
    Set ttemp = ttemp.Offset(3, 0)
    ttemp.Value = ""
    ttemp.Offset(1, 0) = "����1"
    ttemp.Offset(2, 0) = "����2"
    ttemp.Offset(0, 1) = "����Ƚ��"
    ttemp.Offset(0, 2) = "����Ƚ��"
    ttemp.Offset(0, 3) = "��������ġ"
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

    
      title.TextFrame.Characters.Text = "H0 : p1 = p2�ǰ�������"
    With title.TextFrame.Characters.Font
        .size = 12
        .ColorIndex = xlAutomatic
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
    Set ttemp = ttemp.Offset(3, 2)
    Select Case Hyp
    Case 1
       ttemp.Value = " H0 : p1 = p2  vs. H1 : p1 �� p2"
    Case 2
       ttemp.Value = " H0 : p1 = p2  vs. H1 : p1 > p2"
    Case 3
       ttemp.Value = " H0 : p1 = p2  vs. H1 : p1 < p2"
    End Select
        
    
    title.TextFrame.HorizontalAlignment = xlCenter
    Set ttemp = ttemp.Offset(3, -2)
    ttemp.Value = "������跮"
    ttemp.Offset(0, 1).Value = "����Ȯ��"
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
    
    title.TextFrame.Characters.Text = "p1-p2�� ���� " & c & "% �ŷڱ���"
    With title.TextFrame.Characters.Font
        .size = 12
        .ColorIndex = xlAutomatic
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
    Set ttemp = ttemp.Offset(3, 0)
    ttemp.Value = c & "% �ŷڱ���"
    ttemp.Offset(0, 1) = "����"
    ttemp.Offset(0, 2) = "����"
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
   
    title.TextFrame.Characters.Text = "��� �� ���"
    With title.TextFrame.Characters.Font
        .size = 12
        .ColorIndex = xlAutomatic
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
     Set ttemp = ttemp.Offset(3, 0)
   ttemp.Value = "����1�� �������� X1/n1 =" & s1 & "/" & n1 & "=" & Format(s1 / n1, "##0.000") & "�̰�,"
   ttemp.Offset(1, 0) = "����2�� �������� X2/n2 =" & s2 & "/" & n2 & "=" & Format(s2 / n2, "##0.000") & "�̴�."
   ttemp.Offset(2, 0) = "���� ������� ���� �������� (X1+X2)/(n1+n2) = " & Format((s1 + s2) / (n1 + n2), "##0.000") & "���� �� �� �ִ�."
   ttemp.Offset(3, 0) = "���Ǽ�����" & (100 - c) / 100 & " �̰� ����Ȯ���� " & Format(2 * (1 - Application.NormSDist(Abs(Z))), "0.0000") & "�̴�."
  
    If 2 * (1 - Application.NormSDist(Abs(Z))) < (100 - c) / 100 Then
     ttemp.Offset(4, 0).Value = "����Ȯ��(p-value)�� ���Ǽ��إ�= " & (100 - c) / 100 & "���� �۱⶧���� �͹�����H���� �Ⱒ�Ѵ�."
    Else
     ttemp.Offset(4, 0).Value = "����Ȯ��(p-value)�� ���Ǽ��إ�= " & (100 - c) / 100 & "���� ũ�⶧���� �͹�����H���� �Ⱒ���� ���Ѵ�."
    
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
    
    za = (100 - c) / 100 '���Ǽ���
    za1 = WorksheetFunction.ChiInv(za / 2, n - 1) 'ī������(����/2) '����
    za2 = WorksheetFunction.ChiInv(1 - (za / 2), n - 1) 'ī������(1-(����/2))'����

    ll = ((n - 1) * s) / za1
    ul = ((n - 1) * s) / za2
    
    Hyp = choice(3)
    
    TModulePrint.Title1 "��л����� ���� �֩�����"

     ''''
    '''''������跮
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
    title.TextFrame.Characters.Text = "��л� ���� "
    With title.TextFrame.Characters.Font
        .size = 12
        .ColorIndex = xlAutomatic
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
    
   
    Set ttemp = ttemp.Offset(3, 0)
    ttemp.Value = "ǥ���� ����"
    ttemp.Offset(0, 1) = "ǥ���� �л�"
    ttemp.Offset(0, 2) = "��л�"
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
    '''''��������
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
   
    
    title.TextFrame.Characters.Text = "H0 : ���= " & p & " �� ���� ���� ����"
    With title.TextFrame.Characters.Font
        .size = 12
        .ColorIndex = xlAutomatic
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
    
    Set ttemp = ttemp.Offset(3, 2)
    Select Case Hyp
    Case 1
       ttemp.Value = " H0 : ��� =�����vs. H1 : ��� �������     (�����= " & p & " )"
    Case 2
       ttemp.Value = " H0 : ��� =�����vs. H1 : ��� > �����     (�����= " & p & " )"
    Case 3
    ttemp.Value = " H0 : ��� =�����vs. H1 : ��� < �����     (�����= " & p & " )"
    End Select
    
    'ttemp.Select
   ' With Selection.Font
             '.FontStyle = "Bold"
             '.size = 10
   ' End With
    
    
    Set ttemp = ttemp.Offset(3, -2)
    ttemp.Value = "���Ǽ���"
    ttemp.Offset(0, 1) = "������"
    ttemp.Offset(0, 2) = "������跮"
    ttemp.Offset(0, 3) = "�Ⱒ��"
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
    '''''�ŷڱ���
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
    
    title.TextFrame.Characters.Text = "�֩������� ���� " & c & "% �ŷڱ���"
    With title.TextFrame.Characters.Font
        .size = 12
        .ColorIndex = xlAutomatic
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
    Set ttemp = ttemp.Offset(3, 0)
    ttemp.Value = c & "% �ŷڱ���"
    ttemp.Offset(0, 1) = c & "% ����"
    ttemp.Offset(0, 2) = c & "% ����"
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
 
    '''''������ؼ�
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
    
    title.TextFrame.Characters.Text = "��� �� ���"
    With title.TextFrame.Characters.Font
        .size = 12
        .ColorIndex = xlAutomatic
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
    Set ttemp = ttemp.Offset(3, 0)
     ttemp.Value = "���Ǽ��� ��=" & za & "�̰� ��������" & n & "- 1 = " & n - 1 & "�̴�."
     ttemp.Offset(1, 0) = "������跮�� ���ϸ�" & Format((n - 1) * s / p, "##0.000") & "�̴�."
     ttemp.Offset(2, 0) = "�ش� X������ ī������ǥ���� ã���� �Ⱒ�� = " & Format(WorksheetFunction.ChiInv(za, n - 1), "##0.0000") & "�� ���´�."
    If (n - 1) * s / p > Format(WorksheetFunction.ChiInv(za, n - 1), "##0.0000") Then
        ttemp.Offset(4, 0) = " ������跮 > �Ⱒ���̹Ƿ� �͹������� �Ⱒ�ȴ�. "
    Else
         ttemp.Offset(4, 0) = " ������跮 < �Ⱒ���̹Ƿ� �͹������� ä�õȴ�. "
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
