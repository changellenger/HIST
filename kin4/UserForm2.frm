VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   OleObjectBlob   =   "UserForm2.frx":0000
   Caption         =   "UserForm2"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   StartUpPosition =   1  '家蜡磊 啊款单
   TypeInfoVer     =   1
End
Attribute VB_Name = "UserForm2"
Attribute VB_Base = "0{FD0142E1-1C38-419E-B427-68967D49011C}{B44790B1-9E6C-43D4-B254-6C8A40E861F9}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub UserForm_Click()
    
    
    Dim TempSheet As Worksheet
       a = ActiveSheet.UsedRange.Columns.Count
       
   Set TempSheet = ActiveCell.Worksheet
    
       Worksheets("己利").Copy after:=Worksheets("己利")
    
    
   Dim myRange As Range
   Dim myArray()
   
 
    
    For j = 1 To a
    If TempSheet.Cells(1, j).Value = "" Then
    TempSheet.Columns(j).Delete
    End If
    Next j
    
    b = TempSheet.UsedRange.Columns.Count
    MsgBox b
   
   Me.ListBox1.Clear
   Set myRange = ActiveSheet.Rows(1)
   cnt = ActiveSheet.UsedRange.Columns.Count
   ReDim myArray(cnt - 1)
   For i = 1 To cnt
     myArray(i - 1) = myRange.Cells(i)
   Next i
   Me.ListBox1.List() = myArray
   
   




End Sub
