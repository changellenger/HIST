Attribute VB_Name = "Packinstall"
Sub packinstall()
Rinterface.StartRServer

 Rinterface.RRun "install.packages(" & Chr(34) & "cluster" & Chr(34) & ")"  ' kmeans
 Rinterface.RRun "install.packages (" & Chr(34) & "forecast" & Chr(34) & ")" '������Ȱ ,�Ƹ���
 Rinterface.RRun "install.packages (" & Chr(34) & "tree" & Chr(34) & ")"    '�ǻ��������
 Rinterface.RRun "install.packages (" & Chr(34) & "arules" & Chr(34) & ")"      'APRIORI
 Rinterface.RRun "install.packages (" & Chr(34) & "arulesViz" & Chr(34) & ")"    'APRIORI
 
Rinterface.RRun "install.packages (" & Chr(34) & "qualityTools" & Chr(34) & ")" '�����ȹ ���μ��� , �����ɷ�
Rinterface.RRun "install.packages (" & Chr(34) & "FrF2" & Chr(34) & ")" '�׷���
Rinterface.RRun "install.packages (" & Chr(34) & "qcc" & Chr(34) & ")" '������
 
End Sub
