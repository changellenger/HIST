Attribute VB_Name = "Packinstall"
Sub packinstall()
Rinterface.StartRServer

 Rinterface.RRun "install.packages(" & Chr(34) & "cluster" & Chr(34) & ")"  ' kmeans
 Rinterface.RRun "install.packages (" & Chr(34) & "forecast" & Chr(34) & ")" '지수평활 ,아리마
 Rinterface.RRun "install.packages (" & Chr(34) & "tree" & Chr(34) & ")"    '의사결정나무
 Rinterface.RRun "install.packages (" & Chr(34) & "arules" & Chr(34) & ")"      'APRIORI
 Rinterface.RRun "install.packages (" & Chr(34) & "arulesViz" & Chr(34) & ")"    'APRIORI
 
Rinterface.RRun "install.packages (" & Chr(34) & "qualityTools" & Chr(34) & ")" '실험계획 요인설계 , 공정능력
Rinterface.RRun "install.packages (" & Chr(34) & "FrF2" & Chr(34) & ")" '그래프
Rinterface.RRun "install.packages (" & Chr(34) & "qcc" & Chr(34) & ")" '관리도
 
End Sub
