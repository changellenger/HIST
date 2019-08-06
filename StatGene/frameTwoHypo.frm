VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameTwoHypo 
   OleObjectBlob   =   "frameTwoHypo.frx":0000
   Caption         =   "따라하기 : 두 표본 가설검정"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13245
   StartUpPosition =   1  '소유자 가운데
   TypeInfoVer     =   115
End
Attribute VB_Name = "frameTwoHypo"
Attribute VB_Base = "0{69444B8A-049A-428C-910A-9B6656EBA29C}{F4248851-8C86-4F80-8761-AA5DF9AD9568}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False


Private Sub Cancel_Click()
    Unload Me
    
End Sub

Private Sub CommandButton1_Click()
    frameDType.Show
End Sub

Private Sub CommandButton10_Click()
Unload Me
frameTwoZ.OptionButton1.Value = True

    frameTwoZ.Show
End Sub

Private Sub CommandButton11_Click()
    frameDType4.Show
End Sub

Private Sub CommandButton2_Click()
    frameDType2.Show
End Sub

Private Sub CommandButton3_Click()
    frameDType5.Show
    
End Sub

Private Sub CommandButton5_Click()
Unload Me
frameTPaired.OptionButton1.Value = True

    frameTPaired.Show
    
End Sub

Private Sub CommandButton6_Click()
Unload Me
frameTwoF.OptionButton1.Value = True

    frameTwoF.Show
    
End Sub

Private Sub CommandButton7_Click()
    frameTwopZ.Show
End Sub

Private Sub CommandButton8_Click()
Unload Me
frameTwoT.OptionButton1.Value = True

    frameTwoT.Show
End Sub

Private Sub Label12_Click()

End Sub

Private Sub Label22_Click()

End Sub

Private Sub Label32_Click()

End Sub
