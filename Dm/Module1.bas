Attribute VB_Name = "Module1"
Sub ShowDm01()
    Example.Show
End Sub

Sub ShowDm06()
frameDMtree.OptionButton1.Value = True
frameDMtree.Show
End Sub


Sub rinstall()
rinterface.StartRServer
rinterface.RRun "install.packages(" & Chr(34) & " kohmap" & Chr(34) & ")"
End Sub

Sub ShowDm02()
    framekmeanscl.OptionButton1.Value = True
    framekmeanscl.CheckBox6.Value = True
    framekmeanscl.Show
End Sub
Sub ShowDm03()
    framehericl.OptionButton1.Value = True
    framehericl.Show
End Sub
Sub ShowDm04()
    frameES.OptionButton1.Value = True
    frameES.Show
End Sub
Sub ShowDm05()
    frameArima.OptionButton1.Value = True
    frameArima.Show
End Sub
