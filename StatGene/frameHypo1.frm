VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameHypo1 
   OleObjectBlob   =   "frameHypo1.frx":0000
   Caption         =   "따라하기 : 가설검정"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11475
   StartUpPosition =   1  '소유자 가운데
   TypeInfoVer     =   106
End
Attribute VB_Name = "frameHypo1"
Attribute VB_Base = "0{C7C0B548-D12B-42E8-BA55-5F167AC0354E}{355C8D18-919C-4781-A154-2C8CD832C813}"
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
    Unload Me
    frameOneHypo.Show
End Sub

Private Sub CommandButton6_Click()
    Unload Me
    frameTwoHypo.Show
End Sub

Private Sub CommandButton7_Click()

    frameDTsize.Show
    
End Sub

Private Sub Label15_Click()

frameFHypo.Show

End Sub

Private Sub OB2_Click()

End Sub

Private Sub OK_Click()

' 한표본
If OptionButton10.Value = True Then
Unload Me
frameOneZ.OptionButton1.Value = True

    frameOneZ.Show
End If

If OptionButton11.Value = True Then
 Unload Me
 frameOneTtest.OptionButton1.Value = True
      frameOneTtest.Show
    
End If

If OptionButton12.Value = True Then
    frameOneChi.Show
End If

If OptionButton9.Value = True Then
    frameOnepZ.Show
End If
'두 표본
If OB2.Value = True Then
frameTwoZ.OptionButton1.Value = True
    frameTwoZ.Show
End If

If OptionButton1.Value = True Then
frameTwoT.OptionButton1.Value = True
    frameTwoT.Show
End If

If OptionButton2.Value = True Then
frameTPaired.OptionButton1.Value = True
    frameTPaired.Show
End If

If OptionButton3.Value = True Then
frameTwoF.OptionButton1.Value = True
    frameTwoF.Show
End If

If OptionButton13.Value = True Then
frameTwoZ.OptionButton1.Value = True
    frameTwopZ.Show
End If
    Unload Me
End Sub

Private Sub OptionButton1_Click()

End Sub

Private Sub OptionButton10_Click()

End Sub

Private Sub OptionButton11_Click()
 frameOneTtest.OptionButton1.Value = True
End Sub

Private Sub OptionButton12_Click()

End Sub

Private Sub OptionButton13_Click()

End Sub

Private Sub OptionButton2_Click()

End Sub

Private Sub OptionButton3_Click()

End Sub

Private Sub OptionButton9_Click()

End Sub

Private Sub UserForm_Click()

End Sub
