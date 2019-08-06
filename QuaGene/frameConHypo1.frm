VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameConHypo1 
   OleObjectBlob   =   "frameConHypo1.frx":0000
   Caption         =   "따라하기 : 관리도"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12045
   StartUpPosition =   1  '소유자 가운데
   TypeInfoVer     =   162
End
Attribute VB_Name = "frameConHypo1"
Attribute VB_Base = "0{3104E733-752F-4F81-8173-C84038F2B387}{21FA452C-617E-423E-B130-6AFF299CEBA5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False


Private Sub Label16_Click()

End Sub

Private Sub Cancel_Click()
Unload Me

End Sub

Private Sub CommandButton10_Click()
    frameConDType2.Show
End Sub

Private Sub CommandButton11_Click()
    frameConDType3.Show
End Sub

Private Sub CommandButton12_Click()
    frameConDType4.Show
End Sub

Private Sub CommandButton13_Click()
GqcP.OptionButton1.Value = False
GqcP.OptionButton1.Value = True
GqcP.Show

End Sub

Private Sub CommandButton14_Click()
GqcU.OptionButton1.Value = False
GqcU.OptionButton1.Value = True
GqcU.Show

End Sub

Private Sub CommandButton15_Click()
GqcC.OptionButton1.Value = False
GqcC.OptionButton1.Value = True
GqcC.Show

End Sub

Private Sub CommandButton16_Click()
GqcI.OptionButton1.Value = False
GqcI.OptionButton1.Value = True
GqcI.Show

End Sub

Private Sub CommandButton17_Click()
GqcS.OptionButton1.Value = False
GqcS.OptionButton1.Value = True
GqcS.Show

End Sub

Private Sub CommandButton18_Click()
GqcR.OptionButton1.Value = False
GqcR.OptionButton1.Value = True
GqcR.Show

End Sub

Private Sub CommandButton6_Click()
GqcNP.OptionButton1.Value = False
GqcNP.OptionButton1.Value = True
GqcNP.Show

End Sub

Private Sub CommandButton8_Click()
    frameConDType1.Show
End Sub

Private Sub CommandButton9_Click()
    frameConDType5.Show
End Sub

Private Sub Label73_Click()

End Sub

Private Sub UserForm_Click()

End Sub
