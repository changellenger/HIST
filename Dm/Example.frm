VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Example 
   OleObjectBlob   =   "Example.frx":0000
   Caption         =   "Example"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2895
   StartUpPosition =   1  '소유자 가운데
   TypeInfoVer     =   3
End
Attribute VB_Name = "Example"
Attribute VB_Base = "0{5E60A1B2-EF1C-4FE8-8562-30349B24C829}{A6637961-C2BD-46DC-9D9D-4260D6BF720A}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False


Private Sub CommandButton1_Click()
 
    rinterface.StartRServer
rinterface.RRun "install.packages (" & Chr(34) & "arules" & Chr(34) & ")"
rinterface.RRun "install.packages (" & Chr(34) & "arulesViz" & Chr(34) & ")"
rinterface.RRun "require (arules)"
rinterface.RRun "require (arulesViz)"




rinterface.RRun "Assex<-read.csv(file = " & Chr(34) & "C:/Users/Administrator/Desktop/Association.csv" & Chr(34) & ", header = TRUE, sep=" & Chr(34) & "," & Chr(34) & ")"



rinterface.RRun "Assex2 <-as.matrix(Assex)"


rinterface.RRun "trans = as(Assex2, " & Chr(34) & "transactions" & Chr(34) & ")"

rinterface.RRun "rules <- apriori(trans, parameter = list(supp = 0.2, conf = 0.3, target = " & Chr(34) & "rules" & Chr(34) & "))"
rinterface.RRun "ins<-inspect(rules)"
rinterface.RRun "aa<-sort(rules[1:20],by=" & Chr(34) & "lift" & Chr(34) & ")"
rinterface.RRun "plot(aa, method = " & Chr(34) & "graph" & Chr(34) & ")"

rinterface.RRun "win.graph()"
rinterface.RRun "plot(aa, method = " & Chr(34) & "grouped" & Chr(34) & ")"

End Sub


Private Sub CommandButton2_Click()

Unload Me

End Sub
Private Sub CommandButton3_Click()

rinterface.StopRServer True

End Sub








Private Sub UserForm_Click()

fin

End Sub
