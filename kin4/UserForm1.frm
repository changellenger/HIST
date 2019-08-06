VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   OleObjectBlob   =   "UserForm1.frx":0000
   Caption         =   "변수선택"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   StartUpPosition =   1  '소유자 가운데
   TypeInfoVer     =   6
End
Attribute VB_Name = "UserForm1"
Attribute VB_Base = "0{59C59FC8-808E-4D91-942E-F663BF3F0459}{C9FF908A-ED73-4194-A138-475DE90ADC59}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Dim cnt
Dim nGroup As Long, N As Long
Dim arrName As Variant
Dim rngFirst As Range

Private Sub CheckBox1_Click()
    Dim ac As Range, tc As Range
    Dim c As Long, r As Long

    On Error Resume Next
    If Me.CheckBox1 Then
        Set ac = Range(Me.RefEdit1) 'ActiveCell
        Set tc = ac.CurrentRegion

        c = ac.Column - tc.Column
        r = ac.Row - tc.Row

        tc.Offset(r, c).Resize(tc.Rows.Count - r, ac.Columns.Count).Select
        Me.RefEdit1 = selection.Address
      
    Else
        Range(Me.RefEdit1).Resize(1, 1).Select
        Me.RefEdit1 = selection.Address
        Me.RefEdit1.SetFocus
    End If
End Sub

Private Sub CommandButton1_Click()
    Dim rngData As Range
    Set rngData = Range(Me.RefEdit1)
    ReDim arrName(rngData.Columns.Count)
' Reading Data
    For i = 1 To rngData.Columns.Count
        arrName(i) = rngData.Cells(1, i)
    Next i
    
    Set rngData = rngData.Offset(1, 0).Resize(rngData.Rows.Count - 1, rngData.Columns.Count)
    

   Dim myRange As Range
   Dim myArray()
   
   Me.ListBox1.Clear
 
   ReDim myArray(rngData.Columns.Count)
   
   'For i = 1 To rngData.Columns.Count
  'If arrName(i) = "" Then
   ' arrName(i) = "***empty***"
  ' Next i
   
   a = 0
   For i = 1 To rngData.Columns.Count
   If arrName(i) <> "" Then                     '빈칸제거
   myArray(a) = arrName(i)
   a = a + 1
   
   Else:
   End If
   Next i
   Me.ListBox1.List() = myArray
   
   For i = 1 To rngData.Columns.Count
    rngFirst.Offset(i, 1) = myArray(i - 1)
  Next i
  
   ' rngFirst.Offset(2, 2) = "test"
 
End Sub

Private Sub UserForm_Click()

End Sub
