VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameReHypo1 
   OleObjectBlob   =   "frameReHypo1.frx":0000
   Caption         =   "따라하기 : 회귀분석"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10080
   StartUpPosition =   1  '소유자 가운데
   TypeInfoVer     =   72
End
Attribute VB_Name = "frameReHypo1"
Attribute VB_Base = "0{CB13F39B-F779-496F-8D97-D1289B99D2D9}{01D69539-1374-461F-B6AE-092C8B08FB6D}"
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
    frameReDType2.Show
    
End Sub

Private Sub CommandButton4_Click()
'Unload Me
'frameRelogi.OptionButton.value = 1
'    frameRelogi.Show
    
End Sub

Private Sub CommandButton7_Click()
    frameReDType1.Show
    
End Sub

Private Sub CommandButton8_Click()
Unload Me

frameReMulti.OptionButton1.value = 1
frameReMulti.Show
    
End Sub

Private Sub CommandButton9_Click()
Unload Me
frameReSimple1.OptionButton1.value = 1
frameReSimple1.Show
    
End Sub

Private Sub frameMain_Click()

End Sub

Private Sub Label11_Click()

End Sub

Private Sub Label21_Click()

End Sub

Private Sub Label25_Click()

End Sub

Private Sub Label26_Click()
frameFRe.Show
End Sub

Private Sub UserForm_Click()

End Sub
