VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameStNor 
   OleObjectBlob   =   "frameStNor.frx":0000
   Caption         =   "표준정규분포"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12720
   StartUpPosition =   1  '소유자 가운데
   TypeInfoVer     =   50
End
Attribute VB_Name = "frameStNor"
Attribute VB_Base = "0{35252424-963F-4254-A468-5D14C0BBC892}{987225F3-F09D-4563-AE4F-A6896F7EE778}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False


Sub CommandButton1_Click()
    zv = Val(TextBox1.Text)
    p = Application.WorksheetFunction.NormSDist(zv)
    TextBox2.Text = Format((1 - p), "0.00000")
End Sub


Sub CommandButton2_Click()
    zp = Val(TextBox4.Text)
    zv = Application.WorksheetFunction.NormSInv(1 - zp)
    TextBox3.Text = Format(zv, "0.00000")
End Sub
Private Sub CommandButton29_Click()
    ChartOut 4, _
    "표준정규분포(df=" & TextBox5.Value & ")"
End Sub

Private Sub Image4_Click()

End Sub

Private Sub TextBox3_Change()

End Sub

Private Sub TextBox4_Change()

End Sub
