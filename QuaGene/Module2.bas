Attribute VB_Name = "Module2"



Sub makebtn(str As String)
  Dim btn As Button
  Application.ScreenUpdating = False
  ActiveSheet.Buttons.Delete
    i = 9
    Dim stn As Integer
    
  Dim t As Range
  Dim LResult As String
  
  
  stn = ActiveSheet.Cells(1, 1).Value
  LResult = Left(str, 4)
  
  If LResult = "'btn" Or LResult = "btnI" Then
     k = stn + 44
   
   
   If Left(str, 6) = "'btnNP" Then
    k = stn + 35
    End If
   

   
  Else: k = stn + 35
   End If
    
    Set t = ActiveSheet.Range(Cells(k, i), Cells(k, i)) ' ���� ��ġ
    Set btn = ActiveSheet.Buttons.Add(t.Left, t.Top, t.Width, t.Height)
    With btn
      .OnAction = str
      .Caption = "��"
      .Name = "Btn" & i
          End With
  
    
  
  Application.ScreenUpdating = True
End Sub

Sub btnS2()
 MsgBox " check"
End Sub

Sub btnP()

 Dim a As String
 
 Rinterface.StartRServer
 Rinterface.RRun "install.packages (" & Chr(34) & "qcc" & Chr(34) & ")"
    Rinterface.RRun "require (qcc)"



Application.ScreenUpdating = False
    Dim stname As String
    Dim stn As Integer
    
    
     stname = "�����ϱ� ������"
    Module3.OpenOutSheet stname, True
    Worksheets(stname).Activate
    stn = ActiveSheet.Cells(1, 1).Value
    ActiveSheet.Cells(stn + 2, 1).Value = xlist
   ActiveSheet.Cells(stn + 2, 2).Value = xlist2
   



a = " p <- qcc(arraytest, type= " & Chr(34) & "p" & Chr(34) & ", size= arraytest2, plot = FALSE)"
Rinterface.RRun a


Rinterface.RRun "l <- limits.p(p$center,p$std.dev,p$sizes,3) "
Rinterface.RRun "s <- stats.p(p$data, p$sizes)"
Rinterface.RRun "ss <- as.data.frame(s)"
Rinterface.RRun "d <- data.frame(l,ss)"

Rinterface.RRun "aa <- as.data.frame(arraytest)"
Rinterface.RRun "bb <- as.data.frame(arraytest2)"
Rinterface.RRun "cc <- data.frame(aa,bb)"

Rinterface.RRun "dd <- cc[-which(d$UCL <= d$statistics), ]"
Rinterface.GetArray "dd", Range(Cells(stn + 3, 1), Cells(stn + 3, 1))

Rinterface.RRun "pp <- qcc(dd$arraytest, type = " & Chr(34) & "p" & Chr(34) & ", sizes = dd$arraytest2, title= " & Chr(34) & "P������" & Chr(34) & ")"


Rinterface.InsertCurrentRPlot Range(Cells(stn + 1, 4), Cells(stn + 1, 4)), widthrescale:=0.6, heightrescale:=0.6, closergraph:=True





Rinterface.RRun "NofSG <- length(dd$arraytest)" '�κб���'
Rinterface.RRun "MSofSG <- mean(dd$arraytest2)" '��� �κб� ũ��'
Rinterface.RRun "NofD <- sum(dd$arraytest)" '�ҷ�ǰ ��'
Rinterface.RRun "ALL <- sum(dd$arraytest2)" '�� �׸��'
Rinterface.RRun "PofD <- (NofD/ALL)*100" '�ҷ���'

Range(Cells(stn + 1, 8), Cells(stn + 1, 8)).Value = "�κб� ��"
Range(Cells(stn + 1, 8), Cells(stn + 1, 8)).Cells.ColumnWidth = 15
Range(Cells(stn + 2, 8), Cells(stn + 2, 8)).Value = "��� �κб� ũ��"
Range(Cells(stn + 2, 8), Cells(stn + 2, 8)).Cells.ColumnWidth = 15
Range(Cells(stn + 3, 8), Cells(stn + 3, 8)).Value = "�ҷ�ǰ ��"
Range(Cells(stn + 3, 8), Cells(stn + 3, 8)).Cells.ColumnWidth = 15
Range(Cells(stn + 4, 8), Cells(stn + 4, 8)).Value = "�� �׸��"
Range(Cells(stn + 4, 8), Cells(stn + 4, 8)).Cells.ColumnWidth = 15
Range(Cells(stn + 5, 8), Cells(stn + 5, 8)).Value = "�ҷ���"
Range(Cells(stn + 5, 8), Cells(stn + 5, 8)).Cells.ColumnWidth = 15



Range(Cells(stn + 1, 8), Cells(stn + 1, 8)).Font.Bold = True
Range(Cells(stn + 1, 8), Cells(stn + 1, 8)).Interior.Color = RGB(220, 238, 130)
Range(Cells(stn + 2, 8), Cells(stn + 2, 8)).Font.Bold = True
Range(Cells(stn + 2, 8), Cells(stn + 2, 8)).Interior.Color = RGB(220, 238, 130)
Range(Cells(stn + 3, 8), Cells(stn + 3, 8)).Font.Bold = True
Range(Cells(stn + 3, 8), Cells(stn + 3, 8)).Interior.Color = RGB(220, 238, 130)
Range(Cells(stn + 4, 8), Cells(stn + 4, 8)).Font.Bold = True
Range(Cells(stn + 4, 8), Cells(stn + 4, 8)).Interior.Color = RGB(220, 238, 130)
Range(Cells(stn + 5, 8), Cells(stn + 5, 8)).Font.Bold = True
Range(Cells(stn + 5, 8), Cells(stn + 5, 8)).Interior.Color = RGB(220, 238, 130)



Rinterface.GetArray "NofSG", Range(Cells(stn + 1, 9), Cells(stn + 1, 9))
Rinterface.GetArray "MSofSG", Range(Cells(stn + 2, 9), Cells(stn + 2, 9))
Rinterface.GetArray "NofD", Range(Cells(stn + 3, 9), Cells(stn + 3, 9))
Rinterface.GetArray "ALL", Range(Cells(stn + 4, 9), Cells(stn + 4, 9))
Rinterface.GetArray "PofD", Range(Cells(stn + 5, 9), Cells(stn + 5, 9))

  Range(Cells(stn + 1, 8), Cells(stn + 5, 8)).Borders(xlEdgeLeft).LineStyle = xlContinuous
Range(Cells(stn + 1, 8), Cells(stn + 5, 8)).Borders(xlEdgeLeft).Color = RGB(34, 116, 34)
 Range(Cells(stn + 1, 8), Cells(stn + 5, 8)).Borders(xlEdgeLeft).Weight = 3
 
 
 Range(Cells(stn + 1, 9), Cells(stn + 5, 9)).Borders(xlEdgeRight).LineStyle = xlContinuous  '���� ������ �׵θ� ����
   Range(Cells(stn + 1, 9), Cells(stn + 5, 9)).Borders(xlEdgeRight).Color = RGB(34, 116, 34)
  Range(Cells(stn + 1, 9), Cells(stn + 5, 9)).Borders(xlEdgeRight).Weight = 3
 
 
 
 Range(Cells(stn + 1, 8), Cells(stn + 1, 9)).Borders(xlEdgeTop).LineStyle = xlContinuous '���� ���� �׵θ� ����
    Range(Cells(stn + 1, 8), Cells(stn + 1, 9)).Borders(xlEdgeTop).Color = RGB(34, 116, 34)
  Range(Cells(stn + 1, 8), Cells(stn + 1, 9)).Borders(xlEdgeTop).Weight = 3
 
 Range(Cells(stn + 5, 8), Cells(stn + 5, 9)).Borders(xlEdgeBottom).LineStyle = xlContinuous  '���� �Ʒ��� �׵θ� ����
   Range(Cells(stn + 5, 8), Cells(stn + 5, 9)).Borders(xlEdgeBottom).Color = RGB(34, 116, 34)
Range(Cells(stn + 5, 8), Cells(stn + 5, 9)).Borders(xlEdgeBottom).Weight = 3
 
 
 Range(Cells(stn + 1, 9), Cells(stn + 5, 9)).Borders(xlEdgeLeft).LineStyle = xlContinuous '���� ���� �׵θ�
 Range(Cells(stn + 1, 9), Cells(stn + 5, 9)).Borders(xlEdgeLeft).Color = RGB(34, 116, 34)
 Range(Cells(stn + 1, 9), Cells(stn + 5, 9)).Borders(xlEdgeLeft).Weight = 3
 


  Range(Cells(stn + 28, 1), Cells(stn + 28, 25)).Borders(xlEdgeBottom).LineStyle = xlContinuous  '���� �Ʒ��� �׵θ� ����
   Range(Cells(stn + 28, 1), Cells(stn + 28, 25)).Borders(xlEdgeBottom).Color = vbBlack
 Range(Cells(stn + 28, 1), Cells(stn + 28, 25)).Borders(xlEdgeBottom).Weight = 1



If N > 28 Then
 ActiveSheet.Cells(1, 1).Value = stn + N + 2
 Else
ActiveSheet.Cells(1, 1).Value = stn + 30
 
 End If
   

End Sub

Sub btnNP(no As Integer)

 Dim a As String
'Dim no As Integer
 
 Rinterface.StartRServer
 Rinterface.RRun "install.packages (" & Chr(34) & "qcc" & Chr(34) & ")"
    Rinterface.RRun "require (qcc)"
 
 'no = GqcNP.TextBox1.Value
 
 no = -no

'MsgBox no
Application.ScreenUpdating = False
    Dim stname As String
    Dim stn As Integer
    
    
     stname = "�����ϱ� ������"
    Module3.OpenOutSheet stname, True
    Worksheets(stname).Activate
    stn = ActiveSheet.Cells(1, 1).Value
    ActiveSheet.Cells(stn + 2, 1).Value = xlist
    

a = " np <- qcc(arraytest, type= " & Chr(34) & "np" & Chr(34) & ", size= " & no & ", plot = FALSE)"
Rinterface.RRun a


Rinterface.RRun "l <- limits.np(np$center,np$std.dev,np$sizes,3) "
Rinterface.RRun "s <- stats.np(np$data, np$sizes)"
Rinterface.RRun "ss <- as.data.frame(s)"
Rinterface.RRun "d <- data.frame(l,ss)"

Rinterface.RRun "aa <- as.data.frame(arraytest)"

Rinterface.RRun "bb <- aa[-which(d$UCL <= d$statistics), ]"

Rinterface.GetArray "bb", Range(Cells(stn + 3, 1), Cells(stn + 3, 1))

Rinterface.RRun "npnp <- qcc(bb, type = " & Chr(34) & "np" & Chr(34) & ", sizes = " & no & ", title= " & Chr(34) & "NP������" & Chr(34) & ")"

Rinterface.InsertCurrentRPlot Range(Cells(stn + 1, 3), Cells(stn + 1, 3)), widthrescale:=0.6, heightrescale:=0.6, closergraph:=True




Rinterface.RRun "NofSG <- length(bb)" '�κб���'
Rinterface.RRun "N <- " & no & " " '�κб� ũ��'
Rinterface.RRun "NofD <- sum(bb)" '�ҷ�ǰ ��'
Rinterface.RRun "ALL <- N*NofSG" '�� �׸��'
Rinterface.RRun "PofD <- (NofD/ALL)*100" '�ҷ���'

Range(Cells(stn + 1, 7), Cells(stn + 1, 7)).Value = "�κб� ��"
Range(Cells(stn + 1, 7), Cells(stn + 1, 7)).Cells.ColumnWidth = 15
Range(Cells(stn + 2, 7), Cells(stn + 2, 7)).Value = "�κб� ũ��"
Range(Cells(stn + 2, 7), Cells(stn + 2, 7)).Cells.ColumnWidth = 15
Range(Cells(stn + 3, 7), Cells(stn + 3, 7)).Value = "�ҷ�ǰ ��"
Range(Cells(stn + 3, 7), Cells(stn + 3, 7)).Cells.ColumnWidth = 15
Range(Cells(stn + 4, 7), Cells(stn + 4, 7)).Value = "�� �׸��"
Range(Cells(stn + 4, 7), Cells(stn + 4, 7)).Cells.ColumnWidth = 15
Range(Cells(stn + 5, 7), Cells(stn + 5, 7)).Value = "�ҷ���"
Range(Cells(stn + 5, 7), Cells(stn + 5, 7)).Cells.ColumnWidth = 15

Rinterface.GetArray "NofSG", Range(Cells(stn + 1, 8), Cells(stn + 1, 8))
Rinterface.GetArray "N", Range(Cells(stn + 2, 8), Cells(stn + 2, 8))
Rinterface.GetArray "NofD", Range(Cells(stn + 3, 8), Cells(stn + 3, 8))
Rinterface.GetArray "ALL", Range(Cells(stn + 4, 8), Cells(stn + 4, 8))
Rinterface.GetArray "PofD", Range(Cells(stn + 5, 8), Cells(stn + 5, 8))

Range(Cells(stn + 1, 7), Cells(stn + 1, 7)).Font.Bold = True
Range(Cells(stn + 1, 7), Cells(stn + 1, 7)).Interior.Color = RGB(220, 238, 130)
Range(Cells(stn + 2, 7), Cells(stn + 2, 7)).Font.Bold = True
Range(Cells(stn + 2, 7), Cells(stn + 2, 7)).Interior.Color = RGB(220, 238, 130)
Range(Cells(stn + 3, 7), Cells(stn + 3, 7)).Font.Bold = True
Range(Cells(stn + 3, 7), Cells(stn + 3, 7)).Interior.Color = RGB(220, 238, 130)
Range(Cells(stn + 4, 7), Cells(stn + 4, 7)).Font.Bold = True
Range(Cells(stn + 4, 7), Cells(stn + 4, 7)).Interior.Color = RGB(220, 238, 130)
Range(Cells(stn + 5, 7), Cells(stn + 5, 7)).Font.Bold = True
Range(Cells(stn + 5, 7), Cells(stn + 5, 7)).Interior.Color = RGB(220, 238, 130)


 
 
 Range(Cells(stn + 1, 7), Cells(stn + 5, 7)).Borders(xlEdgeLeft).LineStyle = xlContinuous
Range(Cells(stn + 1, 7), Cells(stn + 5, 7)).Borders(xlEdgeLeft).Color = RGB(34, 116, 34)
 Range(Cells(stn + 1, 7), Cells(stn + 5, 7)).Borders(xlEdgeLeft).Weight = 3
 
 
 Range(Cells(stn + 1, 8), Cells(stn + 5, 8)).Borders(xlEdgeRight).LineStyle = xlContinuous  '���� ������ �׵θ� ����
   Range(Cells(stn + 1, 8), Cells(stn + 5, 8)).Borders(xlEdgeRight).Color = RGB(34, 116, 34)
  Range(Cells(stn + 1, 8), Cells(stn + 5, 8)).Borders(xlEdgeRight).Weight = 3
 
 
 
 Range(Cells(stn + 1, 7), Cells(stn + 1, 8)).Borders(xlEdgeTop).LineStyle = xlContinuous '���� ���� �׵θ� ����
    Range(Cells(stn + 1, 7), Cells(stn + 1, 8)).Borders(xlEdgeTop).Color = RGB(34, 116, 34)
  Range(Cells(stn + 1, 7), Cells(stn + 1, 8)).Borders(xlEdgeTop).Weight = 3
 
 Range(Cells(stn + 5, 7), Cells(stn + 5, 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous  '���� �Ʒ��� �׵θ� ����
   Range(Cells(stn + 5, 7), Cells(stn + 5, 8)).Borders(xlEdgeBottom).Color = RGB(34, 116, 34)
Range(Cells(stn + 5, 7), Cells(stn + 5, 8)).Borders(xlEdgeBottom).Weight = 3
 
 
 Range(Cells(stn + 1, 8), Cells(stn + 5, 8)).Borders(xlEdgeLeft).LineStyle = xlContinuous '���� ���� �׵θ�
 Range(Cells(stn + 1, 8), Cells(stn + 5, 8)).Borders(xlEdgeLeft).Color = RGB(34, 116, 34)
 Range(Cells(stn + 1, 8), Cells(stn + 5, 8)).Borders(xlEdgeLeft).Weight = 3
 


  Range(Cells(stn + 28, 1), Cells(stn + 28, 25)).Borders(xlEdgeBottom).LineStyle = xlContinuous  '���� �Ʒ��� �׵θ� ����
   Range(Cells(stn + 28, 1), Cells(stn + 28, 25)).Borders(xlEdgeBottom).Color = vbBlack
 Range(Cells(stn + 28, 1), Cells(stn + 28, 25)).Borders(xlEdgeBottom).Weight = 1



If N > 28 Then
 ActiveSheet.Cells(1, 1).Value = stn + N + 2
 Else
ActiveSheet.Cells(1, 1).Value = stn + 30
 
 End If

End Sub
Sub btnU()

 Dim a As String
 
 Rinterface.StartRServer
 Rinterface.RRun "install.packages (" & Chr(34) & "qcc" & Chr(34) & ")"
    Rinterface.RRun "require (qcc)"
    
    
 stname = "�����ϱ� ������"
    Module3.OpenOutSheet stname, True
    Worksheets(stname).Activate
    stn = ActiveSheet.Cells(1, 1).Value
    ActiveSheet.Cells(stn + 2, 1).Value = xlist
   ActiveSheet.Cells(stn + 2, 2).Value = xlist2
   



a = " u <- qcc(arraytest, type= " & Chr(34) & "u" & Chr(34) & ", size= arraytest2, plot = FALSE)"
Rinterface.RRun a


Rinterface.RRun "l <- limits.u(u$center,u$std.dev,u$sizes,3) "
Rinterface.RRun "s <- stats.u(u$data, u$sizes)"
Rinterface.RRun "ss <- as.data.frame(s)"
Rinterface.RRun "d <- data.frame(l,ss)"

Rinterface.RRun "aa <- as.data.frame(arraytest)"
Rinterface.RRun "bb <- as.data.frame(arraytest2)"
Rinterface.RRun "cc <- data.frame(aa,bb)"

Rinterface.RRun "dd <- cc[-which(d$UCL <= d$statistics), ]"

Rinterface.RRun "pp <- qcc(dd$arraytest, type = " & Chr(34) & "u" & Chr(34) & ", sizes = dd$arraytest2, title= " & Chr(34) & "U������" & Chr(34) & ")"


Rinterface.InsertCurrentRPlot Range(Cells(stn + 1, 4), Cells(stn + 1, 4)), widthrescale:=0.6, heightrescale:=0.6, closergraph:=True




Rinterface.RRun "NofSG <- length(dd$arraytest)" '�κб���'
Rinterface.RRun "MSofSG <- mean(dd$arraytest2)" '��� �κб� ũ��'
Rinterface.RRun "SNMSG <- NofSG*MSofSG" '�� ������'
Rinterface.RRun "SD <- sum(dd$arraytest)" '�� ������'
Rinterface.RRun "PofD <- SD/SNMSG" '������ ���� ��'

Range(Cells(stn + 1, 8), Cells(stn + 1, 8)).Value = "�κб� ��"
Range(Cells(stn + 1, 8), Cells(stn + 1, 8)).Cells.ColumnWidth = 15
Range(Cells(stn + 2, 8), Cells(stn + 2, 8)).Value = "��� �κб� ũ��"
Range(Cells(stn + 2, 8), Cells(stn + 2, 8)).Cells.ColumnWidth = 15
Range(Cells(stn + 3, 8), Cells(stn + 3, 8)).Value = "�� ������"
Range(Cells(stn + 3, 8), Cells(stn + 3, 8)).Cells.ColumnWidth = 15
Range(Cells(stn + 4, 8), Cells(stn + 4, 8)).Value = "�� ������"
Range(Cells(stn + 4, 8), Cells(stn + 4, 8)).Cells.ColumnWidth = 15
Range(Cells(stn + 5, 8), Cells(stn + 5, 8)).Value = "������ ���� ��"
Range(Cells(stn + 5, 8), Cells(stn + 5, 8)).Cells.ColumnWidth = 15


Rinterface.GetArray "NofSG", Range(Cells(stn + 1, 9), Cells(stn + 1, 9))
Rinterface.GetArray "MSofSG", Range(Cells(stn + 2, 9), Cells(stn + 2, 9))
Rinterface.GetArray "SNMSG", Range(Cells(stn + 3, 9), Cells(stn + 3, 9))
Rinterface.GetArray "SD", Range(Cells(stn + 4, 9), Cells(stn + 4, 9))
Rinterface.GetArray "PofD", Range(Cells(stn + 5, 9), Cells(stn + 5, 9))

Range(Cells(stn + 1, 8), Cells(stn + 1, 8)).Font.Bold = True
Range(Cells(stn + 1, 8), Cells(stn + 1, 8)).Interior.Color = RGB(220, 238, 130)
Range(Cells(stn + 2, 8), Cells(stn + 2, 8)).Font.Bold = True
Range(Cells(stn + 2, 8), Cells(stn + 2, 8)).Interior.Color = RGB(220, 238, 130)
Range(Cells(stn + 3, 8), Cells(stn + 3, 8)).Font.Bold = True
Range(Cells(stn + 3, 8), Cells(stn + 3, 8)).Interior.Color = RGB(220, 238, 130)
Range(Cells(stn + 4, 8), Cells(stn + 4, 8)).Font.Bold = True
Range(Cells(stn + 4, 8), Cells(stn + 4, 8)).Interior.Color = RGB(220, 238, 130)
Range(Cells(stn + 5, 8), Cells(stn + 5, 8)).Font.Bold = True
Range(Cells(stn + 5, 8), Cells(stn + 5, 8)).Interior.Color = RGB(220, 238, 130)



 
  Range(Cells(stn + 1, 8), Cells(stn + 5, 8)).Borders(xlEdgeLeft).LineStyle = xlContinuous
Range(Cells(stn + 1, 8), Cells(stn + 5, 8)).Borders(xlEdgeLeft).Color = RGB(34, 116, 34)
 Range(Cells(stn + 1, 8), Cells(stn + 5, 8)).Borders(xlEdgeLeft).Weight = 3
 
 
 Range(Cells(stn + 1, 9), Cells(stn + 5, 9)).Borders(xlEdgeRight).LineStyle = xlContinuous  '���� ������ �׵θ� ����
   Range(Cells(stn + 1, 9), Cells(stn + 5, 9)).Borders(xlEdgeRight).Color = RGB(34, 116, 34)
  Range(Cells(stn + 1, 9), Cells(stn + 5, 9)).Borders(xlEdgeRight).Weight = 3
 
 
 
 Range(Cells(stn + 1, 8), Cells(stn + 1, 9)).Borders(xlEdgeTop).LineStyle = xlContinuous '���� ���� �׵θ� ����
    Range(Cells(stn + 1, 8), Cells(stn + 1, 9)).Borders(xlEdgeTop).Color = RGB(34, 116, 34)
  Range(Cells(stn + 1, 8), Cells(stn + 1, 9)).Borders(xlEdgeTop).Weight = 3
 
 Range(Cells(stn + 5, 8), Cells(stn + 5, 9)).Borders(xlEdgeBottom).LineStyle = xlContinuous  '���� �Ʒ��� �׵θ� ����
   Range(Cells(stn + 5, 8), Cells(stn + 5, 9)).Borders(xlEdgeBottom).Color = RGB(34, 116, 34)
Range(Cells(stn + 5, 8), Cells(stn + 5, 9)).Borders(xlEdgeBottom).Weight = 3
 
 
 Range(Cells(stn + 1, 9), Cells(stn + 5, 9)).Borders(xlEdgeLeft).LineStyle = xlContinuous '���� ���� �׵θ�
 Range(Cells(stn + 1, 9), Cells(stn + 5, 9)).Borders(xlEdgeLeft).Color = RGB(34, 116, 34)
 Range(Cells(stn + 1, 9), Cells(stn + 5, 9)).Borders(xlEdgeLeft).Weight = 3
 


  Range(Cells(stn + 28, 1), Cells(stn + 28, 25)).Borders(xlEdgeBottom).LineStyle = xlContinuous  '���� �Ʒ��� �׵θ� ����
   Range(Cells(stn + 28, 1), Cells(stn + 28, 25)).Borders(xlEdgeBottom).Color = vbBlack
 Range(Cells(stn + 28, 1), Cells(stn + 28, 25)).Borders(xlEdgeBottom).Weight = 1



If N > 28 Then
 ActiveSheet.Cells(1, 1).Value = stn + N + 2
 Else
ActiveSheet.Cells(1, 1).Value = stn + 30
 
 End If
   
End Sub
Sub btnC()

 Dim a As String
'Dim no As Integer
 
 Rinterface.StartRServer
 Rinterface.RRun "install.packages (" & Chr(34) & "qcc" & Chr(34) & ")"
    Rinterface.RRun "require (qcc)"
 
 'no = GqcNP.TextBox1.Value
 
 'no = -no

'MsgBox no



     stname = "�����ϱ� ������"
    Module3.OpenOutSheet stname, True
    Worksheets(stname).Activate
    stn = ActiveSheet.Cells(1, 1).Value
    ActiveSheet.Cells(stn + 2, 1).Value = xlist
    


a = " c <- qcc(arraytest, type= " & Chr(34) & "c" & Chr(34) & ", size= 1000, plot = FALSE)"
Rinterface.RRun a


Rinterface.RRun "l <- limits.c(c$center,c$std.dev,c$sizes,3) "
Rinterface.RRun "s <- stats.c(c$data, c$sizes)"
Rinterface.RRun "ss <- as.data.frame(s)"
Rinterface.RRun "d <- data.frame(l,ss)"

Rinterface.RRun "aa <- as.data.frame(arraytest)"

Rinterface.RRun "bb <- aa[-which(d$UCL < d$statistics), ]"



Rinterface.RRun "cc <- qcc(bb, type = " & Chr(34) & "c" & Chr(34) & ", sizes = 1000,title= " & Chr(34) & "C ������" & Chr(34) & ")"

Rinterface.InsertCurrentRPlot Range(Cells(stn + 1, 3), Cells(stn + 1, 3)), widthrescale:=0.6, heightrescale:=0.6, closergraph:=True



'Rinterface.RRun "NofSG <- length(arraytest)" '�κб���'
'Rinterface.RRun "MSofSG <- mean(arraytest2)" '�κб� ũ��'
Rinterface.RRun "AA <- length(bb)" '�� �˻� ������
Rinterface.RRun "BB <- sum(bb)" '�� ������'
Rinterface.RRun "CC <- BB/AA" '������ ���� ��'

Range(Cells(stn + 1, 7), Cells(stn + 1, 7)).Value = "�� �˻� ������"
Range(Cells(stn + 1, 7), Cells(stn + 1, 7)).Cells.ColumnWidth = 15
Range(Cells(stn + 2, 7), Cells(stn + 2, 7)).Value = "�� ������"
Range(Cells(stn + 2, 7), Cells(stn + 2, 7)).Cells.ColumnWidth = 15
Range(Cells(stn + 3, 7), Cells(stn + 3, 7)).Value = "������ ���� ��"
Range(Cells(stn + 3, 7), Cells(stn + 3, 7)).Cells.ColumnWidth = 15
'Range("g28").Value = "�� ������"
'Range("g29").Value = "������ ���� ��"


Rinterface.GetArray "AA", Range(Cells(stn + 1, 8), Cells(stn + 1, 8))
Rinterface.GetArray "BB", Range(Cells(stn + 2, 8), Cells(stn + 2, 8))
Rinterface.GetArray "CC", Range(Cells(stn + 3, 8), Cells(stn + 3, 8))
'Rinterface.GetArray "SD", Range("sheet1!i28")
'Rinterface.GetArray "PofD", Range("sheet1!i29")

Range(Cells(stn + 1, 7), Cells(stn + 1, 7)).Font.Bold = True
Range(Cells(stn + 1, 7), Cells(stn + 1, 7)).Interior.Color = RGB(220, 238, 130)
Range(Cells(stn + 2, 7), Cells(stn + 2, 7)).Font.Bold = True
Range(Cells(stn + 2, 7), Cells(stn + 2, 7)).Interior.Color = RGB(220, 238, 130)
Range(Cells(stn + 3, 7), Cells(stn + 3, 7)).Font.Bold = True
Range(Cells(stn + 3, 7), Cells(stn + 3, 7)).Interior.Color = RGB(220, 238, 130)




 
 
Range(Cells(stn + 1, 7), Cells(stn + 3, 7)).Borders(xlEdgeLeft).LineStyle = xlContinuous
Range(Cells(stn + 1, 7), Cells(stn + 3, 7)).Borders(xlEdgeLeft).Color = RGB(34, 116, 34)
 Range(Cells(stn + 1, 7), Cells(stn + 3, 7)).Borders(xlEdgeLeft).Weight = 3
 
 
 Range(Cells(stn + 1, 8), Cells(stn + 3, 8)).Borders(xlEdgeRight).LineStyle = xlContinuous  '���� ������ �׵θ� ����
   Range(Cells(stn + 1, 8), Cells(stn + 3, 8)).Borders(xlEdgeRight).Color = RGB(34, 116, 34)
  Range(Cells(stn + 1, 8), Cells(stn + 3, 8)).Borders(xlEdgeRight).Weight = 3
 
 
 
 Range(Cells(stn + 1, 7), Cells(stn + 1, 8)).Borders(xlEdgeTop).LineStyle = xlContinuous '���� ���� �׵θ� ����
    Range(Cells(stn + 1, 7), Cells(stn + 1, 8)).Borders(xlEdgeTop).Color = RGB(34, 116, 34)
  Range(Cells(stn + 1, 7), Cells(stn + 1, 8)).Borders(xlEdgeTop).Weight = 3
 
 Range(Cells(stn + 3, 7), Cells(stn + 3, 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous  '���� �Ʒ��� �׵θ� ����
   Range(Cells(stn + 3, 7), Cells(stn + 3, 8)).Borders(xlEdgeBottom).Color = RGB(34, 116, 34)
Range(Cells(stn + 3, 7), Cells(stn + 3, 8)).Borders(xlEdgeBottom).Weight = 3
 
 
 Range(Cells(stn + 1, 8), Cells(stn + 3, 8)).Borders(xlEdgeLeft).LineStyle = xlContinuous '���� ���� �׵θ�
 Range(Cells(stn + 1, 8), Cells(stn + 3, 8)).Borders(xlEdgeLeft).Color = RGB(34, 116, 34)
 Range(Cells(stn + 1, 8), Cells(stn + 3, 8)).Borders(xlEdgeLeft).Weight = 3
 


  Range(Cells(stn + 28, 1), Cells(stn + 28, 25)).Borders(xlEdgeBottom).LineStyle = xlContinuous  '���� �Ʒ��� �׵θ� ����
   Range(Cells(stn + 28, 1), Cells(stn + 28, 25)).Borders(xlEdgeBottom).Color = vbBlack
 Range(Cells(stn + 28, 1), Cells(stn + 28, 25)).Borders(xlEdgeBottom).Weight = 1



If N > 28 Then
 ActiveSheet.Cells(1, 1).Value = stn + N + 2
 Else
ActiveSheet.Cells(1, 1).Value = stn + 30
 
 End If

   

End Sub
Sub btnR(no As Integer)

 Dim a As String
 Dim b As String
 Dim c As String
 
'Dim no As Integer
 
 Rinterface.StartRServer
 Rinterface.RRun "install.packages (" & Chr(34) & "qcc" & Chr(34) & ")"
    Rinterface.RRun "require (qcc)"
 
 'no = GqcNP.TextBox1.Value
 
 no = -no

'MsgBox no
a = "x1 <- matrix(data= arraytest, ncol= " & no & ", byrow = TRUE)"
Rinterface.RRun a

Application.ScreenUpdating = False
    Dim stname As String
    Dim stn As Integer
    
    
     stname = "�����ϱ� ������"
    Module3.OpenOutSheet stname, True
    Worksheets(stname).Activate
    stn = ActiveSheet.Cells(1, 1).Value
    ActiveSheet.Cells(stn + 2, 1).Value = xlist
   
    

b = " xbar <- qcc(x1, type= " & Chr(34) & "xbar" & Chr(34) & ", nsigmas=3, plot=FALSE) "
Rinterface.RRun b

Rinterface.RRun "xl<- limits.xbar(xbar$center, xbar$std.dev, xbar$sizes,3)"
Rinterface.RRun "xs<- stats.xbar(xbar$data, xbar$sizes)"
Rinterface.RRun "xss <- as.data.frame(xs)"
Rinterface.RRun "xd <- data.frame(xl,xss)"

Rinterface.RRun "v <- x1[-which(xd$UCL < xd$statistics | xd$LCL > xd$statistics), ]"
Rinterface.RRun "vdata <-as.data.frame(v)"

'Dim vstn As String
'For q = 1 To no
'vstn = "vdata$v" & q
'Rinterface.GetArray vstn, Range(Cells(stn + 3, 1), Cells(stn + 3, 1))
'Next q

 Rinterface.GetArray vstn, Range(Cells(stn + 3, 1), Cells(stn + 3, 1))

Rinterface.RRun "cc <- qcc(v, type = " & Chr(34) & "xbar" & Chr(34) & ", nsigmas=3, title= " & Chr(34) & "Xbar ������" & Chr(34) & ")"

Rinterface.InsertCurrentRPlot Range(Cells(stn + 1, 3), Cells(stn + 1, 3)), widthrescale:=0.6, heightrescale:=0.6, closergraph:=True




c = " r <- qcc(x1, type= " & Chr(34) & "R" & Chr(34) & ", nsigmas=3, plot=FALSE) "
Rinterface.RRun c

Rinterface.RRun "rl<- limits.R(r$center, r$std.dev, r$sizes,3)"
Rinterface.RRun "rs<- stats.R(r$data, r$sizes)"
Rinterface.RRun "rss <- as.data.frame(rs)"
Rinterface.RRun "rd <- data.frame(rl,rss)"

Rinterface.RRun "vv <- x1[-which(rd$UCL < rd$statistics | rd$LCL > rd$statistics), ]"

Rinterface.RRun "win.graph()"

Rinterface.RRun "rrr <- qcc(vv, type = " & Chr(34) & "R" & Chr(34) & ", nsigmas=3, title= " & Chr(34) & "R ������" & Chr(34) & ")"

Rinterface.InsertCurrentRPlot Range(Cells(stn + 1, 8), Cells(stn + 1, 8)), widthrescale:=0.6, heightrescale:=0.6, closergraph:=True



'Rinterface.RRun "NofSG <- length(arraytest)" '�κб���'
'Rinterface.RRun "MSofSG <- mean(arraytest2)" '�κб� ũ��'
Rinterface.RRun "AA <- nrow(vv)" '�κб�'
Rinterface.RRun "BB <- mean(vv)" '���'
Rinterface.RRun "CC <- sd(vv)" 'ǥ������'

Range(Cells(stn + 1, 14), Cells(stn + 1, 14)).Value = "�κб�"
Range(Cells(stn + 2, 14), Cells(stn + 2, 14)).Value = "���"
Range(Cells(stn + 3, 14), Cells(stn + 3, 14)).Value = "ǥ������"
'Range("g28").Value = "�� ������"
'Range("g29").Value = "������ ���� ��"


Rinterface.GetArray "AA", Range(Cells(stn + 1, 15), Cells(stn + 1, 15))
Rinterface.GetArray "BB", Range(Cells(stn + 2, 15), Cells(stn + 2, 15))
Rinterface.GetArray "CC", Range(Cells(stn + 3, 15), Cells(stn + 3, 15))
'Rinterface.GetArray "SD", Range("sheet1!i28")
'Rinterface.GetArray "PofD", Range("sheet1!i29")

Range(Cells(stn + 1, 14), Cells(stn + 1, 14)).Font.Bold = True
Range(Cells(stn + 1, 14), Cells(stn + 1, 14)).Interior.Color = RGB(220, 238, 130)
Range(Cells(stn + 2, 14), Cells(stn + 2, 14)).Font.Bold = True
Range(Cells(stn + 2, 14), Cells(stn + 2, 14)).Interior.Color = RGB(220, 238, 130)
Range(Cells(stn + 3, 14), Cells(stn + 3, 14)).Font.Bold = True
Range(Cells(stn + 3, 14), Cells(stn + 3, 14)).Interior.Color = RGB(220, 238, 130)

 
 
 Range(Cells(stn + 1, 14), Cells(stn + 3, 14)).Borders(xlEdgeLeft).LineStyle = xlContinuous
Range(Cells(stn + 1, 14), Cells(stn + 3, 14)).Borders(xlEdgeLeft).Color = RGB(34, 116, 34)
 Range(Cells(stn + 1, 14), Cells(stn + 3, 14)).Borders(xlEdgeLeft).Weight = 3
 
 
 Range(Cells(stn + 1, 15), Cells(stn + 3, 15)).Borders(xlEdgeRight).LineStyle = xlContinuous  '���� ������ �׵θ� ����
   Range(Cells(stn + 1, 15), Cells(stn + 3, 15)).Borders(xlEdgeRight).Color = RGB(34, 116, 34)
  Range(Cells(stn + 1, 15), Cells(stn + 3, 15)).Borders(xlEdgeRight).Weight = 3
 
 
 
 Range(Cells(stn + 1, 14), Cells(stn + 1, 15)).Borders(xlEdgeTop).LineStyle = xlContinuous '���� ���� �׵θ� ����
    Range(Cells(stn + 1, 14), Cells(stn + 1, 15)).Borders(xlEdgeTop).Color = RGB(34, 116, 34)
  Range(Cells(stn + 1, 14), Cells(stn + 1, 15)).Borders(xlEdgeTop).Weight = 3
 
 Range(Cells(stn + 3, 14), Cells(stn + 3, 15)).Borders(xlEdgeBottom).LineStyle = xlContinuous  '���� �Ʒ��� �׵θ� ����
   Range(Cells(stn + 3, 14), Cells(stn + 3, 15)).Borders(xlEdgeBottom).Color = RGB(34, 116, 34)
Range(Cells(stn + 3, 14), Cells(stn + 3, 15)).Borders(xlEdgeBottom).Weight = 3
 
 
 Range(Cells(stn + 1, 15), Cells(stn + 3, 15)).Borders(xlEdgeLeft).LineStyle = xlContinuous '���� ���� �׵θ�
 Range(Cells(stn + 1, 15), Cells(stn + 3, 15)).Borders(xlEdgeLeft).Color = RGB(34, 116, 34)
 Range(Cells(stn + 1, 15), Cells(stn + 3, 15)).Borders(xlEdgeLeft).Weight = 3
 


  Range(Cells(stn + 28, 1), Cells(stn + 28, 25)).Borders(xlEdgeBottom).LineStyle = xlContinuous  '���� �Ʒ��� �׵θ� ����
   Range(Cells(stn + 28, 1), Cells(stn + 28, 25)).Borders(xlEdgeBottom).Color = vbBlack
 Range(Cells(stn + 28, 1), Cells(stn + 28, 25)).Borders(xlEdgeBottom).Weight = 1
 
 If N > 28 Then
 ActiveSheet.Cells(1, 1).Value = stn + N + 2
 Else
ActiveSheet.Cells(1, 1).Value = stn + 30
 
 End If
 


End Sub

 
Sub btnI()

 'Dim a As String
 Dim b As String
 'Dim c As String
 
'Dim no As Integer
 
 Rinterface.StartRServer
 Rinterface.RRun "install.packages (" & Chr(34) & "qcc" & Chr(34) & ")"
    Rinterface.RRun "require (qcc)"
 
 'no = GqcNP.TextBox1.Value
 
 'no = -no

'MsgBox no

Application.ScreenUpdating = False
    Dim stname As String
    Dim stn As Integer
    
    
     stname = "�����ϱ� ������"
    Module3.OpenOutSheet stname, True
    Worksheets(stname).Activate
    stn = ActiveSheet.Cells(1, 1).Value
    ActiveSheet.Cells(stn + 2, 1).Value = xlist

b = " xbar <- qcc(arraytest, type= " & Chr(34) & "xbar.one" & Chr(34) & ", nsigmas=3, plot=FALSE) "
Rinterface.RRun b

Rinterface.RRun "xl<- limits.xbar.one(xbar$center, xbar$std.dev, xbar$sizes,3)"
Rinterface.RRun "xs<- stats.xbar.one(xbar$data, xbar$sizes)"
Rinterface.RRun "xss <- as.data.frame(xs)"
Rinterface.RRun "xd <- data.frame(xl,xss)"

Rinterface.RRun "ff <- as.data.frame(arraytest)"
Rinterface.RRun "v <- ff[-which(xd$UCL < xd$statistics | xd$LCL > xd$statistics), ] "
 
Rinterface.GetArray "v", Range(Cells(stn + 3, 1), Cells(stn + 3, 1))

Rinterface.RRun "cc <- qcc(v, type = " & Chr(34) & "xbar.one" & Chr(34) & ", nsigmas=3,  title= " & Chr(34) & "I ������" & Chr(34) & ")"
Rinterface.InsertCurrentRPlot Range(Cells(stn + 1, 3), Cells(stn + 1, 3)), widthrescale:=0.6, heightrescale:=0.6, closergraph:=True



Rinterface.RRun "imr <- matrix(cbind(arraytest[1:length(arraytest)-1],arraytest[2:length(arraytest)]), ncol=2)"
Rinterface.RRun " mr <- qcc(imr, type=  " & Chr(34) & "R" & Chr(34) & ", nsigmas=3, plot=FALSE)"


Rinterface.RRun "rl<- limits.R(mr$center, mr$std.dev, mr$sizes,3)"
Rinterface.RRun "rs<- stats.R(mr$data, mr$sizes)"
Rinterface.RRun "rss <- as.data.frame(rs)"
Rinterface.RRun "rd <- data.frame(rl,rss)"

Rinterface.RRun "aa <- as.data.frame(imr)"
Rinterface.RRun "vv <- aa[-which(rd$UCL < rd$statistics | rd$LCL > rd$statistics), ]"



Rinterface.RRun "win.graph()"


Rinterface.RRun "iii <- qcc(vv , type = " & Chr(34) & "R" & Chr(34) & ", nsigmas=3, title= " & Chr(34) & "MR ������" & Chr(34) & ")"

Rinterface.InsertCurrentRPlot Range(Cells(stn + 1, 8), Cells(stn + 1, 8)), widthrescale:=0.6, heightrescale:=0.6, closergraph:=True



'Rinterface.RRun "NofSG <- length(arraytest)" '�κб���'
'Rinterface.RRun "MSofSG <- mean(arraytest2)" '�κб� ũ��'
Rinterface.RRun "AA <- length(v)" 'N'
Rinterface.RRun "BB <- mean(v)" '���'
Rinterface.RRun "CC <- sd(v)" 'ǥ������'

Range(Cells(stn + 1, 14), Cells(stn + 1, 14)).Value = "N"
Range(Cells(stn + 2, 14), Cells(stn + 2, 14)).Value = "���"
Range(Cells(stn + 3, 14), Cells(stn + 3, 14)).Value = "ǥ������"
'Range("g28").Value = "�� ������"
'Range("g29").Value = "������ ���� ��"


Range(Cells(stn + 1, 14), Cells(stn + 1, 14)).Font.Bold = True
Range(Cells(stn + 1, 14), Cells(stn + 1, 14)).Interior.Color = RGB(220, 238, 130)
Range(Cells(stn + 2, 14), Cells(stn + 2, 14)).Font.Bold = True
Range(Cells(stn + 2, 14), Cells(stn + 2, 14)).Interior.Color = RGB(220, 238, 130)
Range(Cells(stn + 3, 14), Cells(stn + 3, 14)).Font.Bold = True
Range(Cells(stn + 3, 14), Cells(stn + 3, 14)).Interior.Color = RGB(220, 238, 130)


Rinterface.GetArray "AA", Range(Cells(stn + 1, 15), Cells(stn + 1, 15))
Rinterface.GetArray "BB", Range(Cells(stn + 2, 15), Cells(stn + 2, 15))
Rinterface.GetArray "CC", Range(Cells(stn + 3, 15), Cells(stn + 3, 15))
'Rinterface.GetArray "SD", Range("sheet1!i28")
'Rinterface.GetArray "PofD", Range("sheet1!i29")





 
Range(Cells(stn + 1, 14), Cells(stn + 3, 14)).Borders(xlEdgeLeft).LineStyle = xlContinuous
Range(Cells(stn + 1, 14), Cells(stn + 3, 14)).Borders(xlEdgeLeft).Color = RGB(34, 116, 34)
 Range(Cells(stn + 1, 14), Cells(stn + 3, 14)).Borders(xlEdgeLeft).Weight = 3
 
 
 Range(Cells(stn + 1, 15), Cells(stn + 3, 15)).Borders(xlEdgeRight).LineStyle = xlContinuous  '���� ������ �׵θ� ����
   Range(Cells(stn + 1, 15), Cells(stn + 3, 15)).Borders(xlEdgeRight).Color = RGB(34, 116, 34)
  Range(Cells(stn + 1, 15), Cells(stn + 3, 15)).Borders(xlEdgeRight).Weight = 3
 
 
 
 Range(Cells(stn + 1, 14), Cells(stn + 1, 15)).Borders(xlEdgeTop).LineStyle = xlContinuous '���� ���� �׵θ� ����
    Range(Cells(stn + 1, 14), Cells(stn + 1, 15)).Borders(xlEdgeTop).Color = RGB(34, 116, 34)
  Range(Cells(stn + 1, 14), Cells(stn + 1, 15)).Borders(xlEdgeTop).Weight = 3
 
 Range(Cells(stn + 3, 14), Cells(stn + 3, 15)).Borders(xlEdgeBottom).LineStyle = xlContinuous  '���� �Ʒ��� �׵θ� ����
   Range(Cells(stn + 3, 14), Cells(stn + 3, 15)).Borders(xlEdgeBottom).Color = RGB(34, 116, 34)
Range(Cells(stn + 3, 14), Cells(stn + 3, 15)).Borders(xlEdgeBottom).Weight = 3
 
 
 Range(Cells(stn + 1, 15), Cells(stn + 3, 15)).Borders(xlEdgeLeft).LineStyle = xlContinuous '���� ���� �׵θ�
 Range(Cells(stn + 1, 15), Cells(stn + 3, 15)).Borders(xlEdgeLeft).Color = RGB(34, 116, 34)
 Range(Cells(stn + 1, 15), Cells(stn + 3, 15)).Borders(xlEdgeLeft).Weight = 3
 


  Range(Cells(stn + 28, 1), Cells(stn + 28, 25)).Borders(xlEdgeBottom).LineStyle = xlContinuous  '���� �Ʒ��� �׵θ� ����
   Range(Cells(stn + 28, 1), Cells(stn + 28, 25)).Borders(xlEdgeBottom).Color = vbBlack
 Range(Cells(stn + 28, 1), Cells(stn + 28, 25)).Borders(xlEdgeBottom).Weight = 1
 

If N > 28 Then
 ActiveSheet.Cells(1, 1).Value = stn + N + 2
 Else
ActiveSheet.Cells(1, 1).Value = stn + 30
 
 End If


 
End Sub
Sub btnS(no As Integer)

 Dim a As String
 Dim b As String
 Dim c As String
 
'Dim no As Integer
 
 Rinterface.StartRServer
 Rinterface.RRun "install.packages (" & Chr(34) & "qcc" & Chr(34) & ")"
    Rinterface.RRun "require (qcc)"
 
 'no = GqcNP.TextBox1.Value
 
 no = -no

'MsgBox no
a = "x1 <- matrix(data= arraytest, ncol= " & no & ", byrow = TRUE)"
Rinterface.RRun a


Application.ScreenUpdating = False
    Dim stname As String
    Dim stn As Integer
    
    
     stname = "�����ϱ� ������"
    Module3.OpenOutSheet stname, True
    Worksheets(stname).Activate
    stn = ActiveSheet.Cells(1, 1).Value
    ActiveSheet.Cells(stn + 2, 1).Value = xlist


b = " xbar <- qcc(x1, type= " & Chr(34) & "xbar" & Chr(34) & ", nsigmas=3, plot=FALSE) "
Rinterface.RRun b

Rinterface.RRun "xl<- limits.xbar(xbar$center, xbar$std.dev, xbar$sizes,3)"
Rinterface.RRun "xs<- stats.xbar(xbar$data, xbar$sizes)"
Rinterface.RRun "xss <- as.data.frame(xs)"
Rinterface.RRun "xd <- data.frame(xl,xss)"

Rinterface.RRun "v <- x1[-which(xd$UCL < xd$statistics | xd$LCL > xd$statistics), ]"
Rinterface.GetArray "v", Range(Cells(stn + 3, 1), Cells(stn + 3, 1))


Rinterface.RRun "cc <- qcc(v, type = " & Chr(34) & "xbar" & Chr(34) & ", nsigmas=3, title= " & Chr(34) & "Xbar ������" & Chr(34) & ")"
Rinterface.InsertCurrentRPlot Range(Cells(stn + 1, 3), Cells(stn + 1, 3)), widthrescale:=0.6, heightrescale:=0.6, closergraph:=True


c = " s <- qcc(x1, type= " & Chr(34) & "S" & Chr(34) & ", nsigmas=3, plot=FALSE) "
Rinterface.RRun c

Rinterface.RRun "sl<- limits.S(s$center, s$std.dev, s$sizes,3)"
Rinterface.RRun "ss<- stats.S(s$data, s$sizes)"
Rinterface.RRun "sss <- as.data.frame(ss)"
Rinterface.RRun "sd <- data.frame(sl,sss)"

Rinterface.RRun "vv <- x1[-which(sd$UCL < sd$statistics | sd$LCL > sd$statistics), ]"

Rinterface.RRun "win.graph()"

Rinterface.RRun "sdsd <- qcc(vv, type = " & Chr(34) & "S" & Chr(34) & ", nsigmas=3, title= " & Chr(34) & "S ������" & Chr(34) & ")"
Rinterface.InsertCurrentRPlot Range(Cells(stn + 1, 8), Cells(stn + 1, 8)), widthrescale:=0.6, heightrescale:=0.6, closergraph:=True


'Rinterface.RRun "NofSG <- length(arraytest)" '�κб���'
'Rinterface.RRun "MSofSG <- mean(arraytest2)" '�κб� ũ��'
Rinterface.RRun "AA <- nrow(vv)" '�κб�'
Rinterface.RRun "BB <- mean(vv)" '���'
Rinterface.RRun "CC <- sd(vv)" 'ǥ������'

Range(Cells(stn + 1, 14), Cells(stn + 1, 14)).Value = "�κб�"
Range(Cells(stn + 2, 14), Cells(stn + 2, 14)).Value = "���"
Range(Cells(stn + 3, 14), Cells(stn + 3, 14)).Value = "ǥ������"
'Range("g28").Value = "�� ������"
'Range("g29").Value = "������ ���� ��"


Range(Cells(stn + 1, 14), Cells(stn + 1, 14)).Font.Bold = True
Range(Cells(stn + 1, 14), Cells(stn + 1, 14)).Interior.Color = RGB(220, 238, 130)
Range(Cells(stn + 2, 14), Cells(stn + 2, 14)).Font.Bold = True
Range(Cells(stn + 2, 14), Cells(stn + 2, 14)).Interior.Color = RGB(220, 238, 130)
Range(Cells(stn + 3, 14), Cells(stn + 3, 14)).Font.Bold = True
Range(Cells(stn + 3, 14), Cells(stn + 3, 14)).Interior.Color = RGB(220, 238, 130)

 


Rinterface.GetArray "AA", Range(Cells(stn + 1, 15), Cells(stn + 1, 15))
Rinterface.GetArray "BB", Range(Cells(stn + 2, 15), Cells(stn + 2, 15))
Rinterface.GetArray "CC", Range(Cells(stn + 3, 15), Cells(stn + 3, 15))
'Rinterface.GetArray "SD", Range("sheet1!i28")
'Rinterface.GetArray "PofD", Range("sheet1!i29")


 
 
 
 Range(Cells(stn + 1, 14), Cells(stn + 3, 14)).Borders(xlEdgeLeft).LineStyle = xlContinuous
Range(Cells(stn + 1, 14), Cells(stn + 3, 14)).Borders(xlEdgeLeft).Color = RGB(34, 116, 34)
 Range(Cells(stn + 1, 14), Cells(stn + 3, 14)).Borders(xlEdgeLeft).Weight = 3
 
 
 Range(Cells(stn + 1, 15), Cells(stn + 3, 15)).Borders(xlEdgeRight).LineStyle = xlContinuous  '���� ������ �׵θ� ����
   Range(Cells(stn + 1, 15), Cells(stn + 3, 15)).Borders(xlEdgeRight).Color = RGB(34, 116, 34)
  Range(Cells(stn + 1, 15), Cells(stn + 3, 15)).Borders(xlEdgeRight).Weight = 3
 
 
 
 Range(Cells(stn + 1, 14), Cells(stn + 1, 15)).Borders(xlEdgeTop).LineStyle = xlContinuous '���� ���� �׵θ� ����
    Range(Cells(stn + 1, 14), Cells(stn + 1, 15)).Borders(xlEdgeTop).Color = RGB(34, 116, 34)
  Range(Cells(stn + 1, 14), Cells(stn + 1, 15)).Borders(xlEdgeTop).Weight = 3
 
 Range(Cells(stn + 3, 14), Cells(stn + 3, 15)).Borders(xlEdgeBottom).LineStyle = xlContinuous  '���� �Ʒ��� �׵θ� ����
   Range(Cells(stn + 3, 14), Cells(stn + 3, 15)).Borders(xlEdgeBottom).Color = RGB(34, 116, 34)
Range(Cells(stn + 3, 14), Cells(stn + 3, 15)).Borders(xlEdgeBottom).Weight = 3
 
 
 Range(Cells(stn + 1, 15), Cells(stn + 3, 15)).Borders(xlEdgeLeft).LineStyle = xlContinuous '���� ���� �׵θ�
 Range(Cells(stn + 1, 15), Cells(stn + 3, 15)).Borders(xlEdgeLeft).Color = RGB(34, 116, 34)
 Range(Cells(stn + 1, 15), Cells(stn + 3, 15)).Borders(xlEdgeLeft).Weight = 3
 


  Range(Cells(stn + 28, 1), Cells(stn + 28, 25)).Borders(xlEdgeBottom).LineStyle = xlContinuous  '���� �Ʒ��� �׵θ� ����
   Range(Cells(stn + 28, 1), Cells(stn + 28, 25)).Borders(xlEdgeBottom).Color = vbBlack
 Range(Cells(stn + 28, 1), Cells(stn + 28, 25)).Borders(xlEdgeBottom).Weight = 1
 
If N > 28 Then
 ActiveSheet.Cells(1, 1).Value = stn + N + 2
 Else
ActiveSheet.Cells(1, 1).Value = stn + 30
 
 End If

End Sub
