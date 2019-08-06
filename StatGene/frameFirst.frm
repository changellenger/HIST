VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameFirst 
   OleObjectBlob   =   "frameFirst.frx":0000
   Caption         =   "따라하기 구성"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11490
   StartUpPosition =   1  '소유자 가운데
   TypeInfoVer     =   18
End
Attribute VB_Name = "frameFirst"
Attribute VB_Base = "0{E00AEFDE-DB9C-48EC-8B1A-3F03826A8E9B}{4279BEFE-8479-4BC2-A451-0A093609FF5C}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False



Private Sub CommandButton1_Click()
    Unload Me
    frameFHypo.Show
    
End Sub
