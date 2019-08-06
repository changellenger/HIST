VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameOneHypo 
   OleObjectBlob   =   "frameOneHypo.frx":0000
   Caption         =   "따라하기 : 한 표본 가설검정"
   ClientHeight    =   8820
   ClientLeft      =   7125
   ClientTop       =   375
   ClientWidth     =   11640
   StartUpPosition =   1  '소유자 가운데
   TypeInfoVer     =   44
End
Attribute VB_Name = "frameOneHypo"
Attribute VB_Base = "0{8B748A72-B847-4FC1-B362-4333EE64E460}{4C4194E8-CDF2-4238-8282-F69D0C2466D6}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False



Private Sub Cancel_Click()
    Unload Me
    frameHypo1.Show
    
End Sub

Private Sub CommandButton1_Click()
    frameDType.Show
    
End Sub

Private Sub CommandButton2_Click()
    frameDType2.Show
    
End Sub

Private Sub CommandButton3_Click()
    frameDType3.Show
End Sub

Private Sub CommandButton4_Click()
Unload Me
frameOneZ.OptionButton1.Value = True
    frameOneZ.Show
    
End Sub

Private Sub CommandButton5_Click()
Unload Me
 frameOneTtest.OptionButton1.Value = True
    frameOneTtest.Show
End Sub

Private Sub CommandButton6_Click()
Unload Me
    frameOneChi.Show
End Sub

Private Sub CommandButton7_Click()
Unload Me
    frameOnepZ.Show
    
End Sub

Private Sub CommandButton8_Click()
    frameDType.Show
End Sub

Private Sub Label27_Click()

End Sub

Private Sub Label46_Click()

End Sub

Private Sub UserForm_Click()

End Sub
