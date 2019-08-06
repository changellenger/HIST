Attribute VB_Name = "ModuleExp"
Sub showRRRR()
   rinterface.RRun "install.packages (" & Chr(34) & "FrF2" & Chr(34) & ")" ' : R ��Ű�� �ʿ����:
    rinterface.RRun "require (FrF2)"
    rinterface.RRun "install.packages (" & Chr(34) & "qualityTools" & Chr(34) & ")" ' : R ��Ű�� �ʿ����:
    rinterface.RRun "require (qualityTools)"
    rinterface.RRun "arrayfrac <- fracChoose()"
    rinterface.RRun "output<-as.data.frame(arrayfrac)" 'ok
    
    rinterface.GetDataframe "output", Range("Sheet11!A1")
    Dim lastColumn As Integer
    lastColumn = ActiveCell.Worksheet.UsedRange.Columns.count
    
   ' lastColumn = Sheet1.Cells(1, Columns.Count).End(xlToLeft).Column
 '   MsgBox lastColumn
   ' MsgBox Worksheets("Sheet11").Range("G1").value
    MsgBox Worksheets("Sheet11").Cells(1, lastColumn).value
    
    'Worksheets("Sheet1").Range("A1").Value = 100
    If Worksheets("Sheet11").Cells(1, lastColumn).value = "y" Then
    Worksheets("Sheet11").Cells(1, lastColumn).value = "Response"
        
    End If

End Sub

Sub showR2()

rinterface.RRun "require(qualityTools)"
rinterface.RRun "require(FrF2)"


rinterface.RRun "Response = c(580, 1090, 1392, 568)"
rinterface.RRun "Response(arrayfrac) = Response"
  rinterface.RRun "DesignC<-as.data.frame(arrayfrac)" 'ok
' ȸ��
rinterface.RRun "lm.1 =lm(Response ~ A+B+C+A*B+C*A+B*C+A*B*C, data = arrayfrac)"
rinterface.RRun "summary(lm.1)"

'�л�м�
rinterface.RRun "AnovaREAL<- aov(lm(Response ~ A+B+C+A*B+C*A+B*C+A*B*C, data = arrayfrac))"
rinterface.RRun "anova(AnovaREAL)"

'��ȿ����
'rinterface.RRun "MEPlot(AnovaREAL, main = paste(" & Chr(34) & "��ȿ����" & Chr(34) & "), ylab=" & Chr(34) & "���" & Chr(34) & " , pch = 15, mgp.ylab = 4, cex.title = 1.5, cex.main = par(" & Chr(34) & "cex.main" & Chr(34) & "), lwd = par(" & Chr(34) & "lwd" & Chr(34) & "), abbrev = 3, select = NULL)"
rinterface.RRun "interactionPlot(arrayfrac, response(arrayfrac), fun = mean, main= " & Chr(34) & "ck" & Chr(34) & " , col = 1:2)"
End Sub

Sub showR3()
'rinterface.RRun "paretoPlot(arrayfrac, main = paste(" & Chr(34) & "ǥ��ȭ�� ȿ���� Pareto��Ʈ" & Chr(34) & ") )"
rinterface.RRun "normalPlot(arrayfrac, main = paste(" & Chr(34) & "ǥ��ȭ ȿ���� ����Ȯ����" & Chr(34) & ") )"

rinterface.RRun "par (mfrow = c(1, 3))"
rinterface.RRun "class(par(mfrow=c(1,3)))"


End Sub

Sub showR4()
'#�����׷���1��- ������ġ ok
rinterface.RRun "plot(residuals(AnovaREAL) ~ fitted(AnovaREAL), xlab= " & Chr(34) & "����ġ" & Chr(34) & ", ylab= " & Chr(34) & "����" & Chr(34) & ",main= " & Chr(34) & "�� ����ġ" & Chr(34) & ")"
rinterface.RRun "abline(h=0,lty=1,col= " & Chr(34) & "red" & Chr(34) & ")"

'#�����׷���2��-����Ȯ���� ok
'rinterface.RRun "qqnorm(resid(AnovaREAL),xlab=" & Chr(34) & "����" & Chr(34) & ", ylab=" & Chr(34) & "�����" & Chr(34) & ", main=" & Chr(34) & "����Ȯ����" & Chr(34) & ")"
'rinterface.RRun "qqline(resid(AnovaREAL),lty=1,col=" & Chr(34) & "red" & Chr(34) & ")"

'#�����׷���3��-������׷�
'rinterface.RRun "hist(resid(AnovaREAL), breaks= 9, xlab= " & Chr(34) & "����" & Chr(34) & ",ylab= " & Chr(34) & "��" & Chr(34) & ", main= " & Chr(34) & "���� ������׷�" & Chr(34) & ", border= " & Chr(34) & "red" & Chr(34) & ", col= " & Chr(34) & "black" & Chr(34) & ")"
'rinterface.RRun "lines(c(min(AnovaREAL$breaks), AnovaREAL$mids, mas(AnovaREAL$breaks)), c(0,AnovaREAL$counts,0),type = " & Chr(34) & "l" & Chr(34) & ")"
'rinterface.RRun "Lines(density(AnovaREAL))"

End Sub
