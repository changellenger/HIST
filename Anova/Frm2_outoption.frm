VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm2_outoption 
   OleObjectBlob   =   "Frm2_outoption.frx":0000
   Caption         =   "免仿可记"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2595
   StartUpPosition =   1  '家蜡磊 啊款单
   TypeInfoVer     =   8
End
Attribute VB_Name = "Frm2_outoption"
Attribute VB_Base = "0{DC706E5F-D321-4EC9-B4F2-A3245A4EF368}{17C33075-9C58-4787-A7B3-F7F881670B4A}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private Sub CheckBox6_Click()
If CheckBox6.Value = True Then
    Frm2_way2.CheckBox3.Enabled = True
Else
    Frm2_way2.CheckBox3.Value = False
    Frm2_way2.CheckBox3.Enabled = False
End If
End Sub

Private Sub CommandButton7_Click()
    Unload Me
End Sub

Private Sub CommandButton8_Click()
    Me.Hide
End Sub
