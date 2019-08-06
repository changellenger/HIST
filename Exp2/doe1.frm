VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} doe1 
   OleObjectBlob   =   "doe1.frx":0000
   Caption         =   "요인 설계 생성(2수준)"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3750
   StartUpPosition =   1  '소유자 가운데
   TypeInfoVer     =   280
End
Attribute VB_Name = "doe1"
Attribute VB_Base = "0{AF198A7A-4BF2-42C4-8023-0C6830ABE6DB}{89B0F771-886C-45F9-ADF9-45A9FEA70FF0}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False


Private Sub CommandButton3_Click()
ShellExecute 0, "open", "hh.exe", ThisWorkbook.Path + "\KESS%202013.chm::/요인%20설계.htm", "", 1
End Sub

Private Sub CommandButton4_Click()

    Dim nfact1, a1, p, q, value, co As Integer
    Dim wsheet2 As Worksheet
    Dim resultsheet, s As Worksheet
    
    Dim a As String
    Dim b As String
    Dim c As String
    
    Dim Drng As Range
   
    Dim blocksN As Integer
    Dim runsN As Integer
    Dim replicationsN As Integer
    Dim factorsN As Integer
    Dim centersN As Integer
    Dim numberK As Integer
    Dim sheetname As String
    
    
     
    factorsN = TextBox1.value ' 요인수
    centersN = ComboBox4.value '중심점 추가
    blocksN = ComboBox3.value ' 블록수
    replicationsN = ComboBox2.value ' 반복수
    
 rinterface.RRun "install.packages (" & Chr(34) & "qualityTools" & Chr(34) & ")"
rinterface.RRun "install.packages (" & Chr(34) & "FrF2" & Chr(34) & ")"
    
rinterface.RRun "require(qualityTools)"
rinterface.RRun "require(FrF2)"


Dim Rstr As String
Dim lastColumn As Integer

Rstr = "arrayfrac<- fracDesign(k = " & factorsN & " , p = 0, gen = NULL, replicates = " & replicationsN & " , blocks = " & blocksN & ", centerCube = " & centersN & " , random.seed = 10000)"
rinterface.RRun Rstr
'rinterface.RRun " Please2<- fracDesign(k = 3, p = 0, gen = NULL, replicates = 3, blocks = 1, centerCube = 0, random.seed = 10000)"
rinterface.RRun "output<-as.data.frame(arrayfrac)" 'ok
rinterface.RRun "factorNs<- " & factorsN





'dr1 = sheet.name

ActiveSheet.Cells(1, 1).Select

rinterface.GetDataframe "output", Range("A1")

 lastColumn = ActiveCell.Worksheet.UsedRange.Columns.count



If ActiveSheet.Cells(1, lastColumn).value = "y" Then
ActiveSheet.Cells(1, lastColumn).value = "Response"

End If


'If Worksheets(ActiveSheet).Cells(1, lastColumn).value = "y" Then
'Worksheets(ActiveSheet).Cells(1, lastColumn).value = "Response"

'End If


 
Unload Me
    
End Sub



Private Sub SpinButton1_Change()
    Me.TextBox1.value = SpinButton1.value
    Dim a, b, c As Integer
    
    
        If doe1.TextBox1.value = 2 Then
         '   doe1.ListBox1.list() = Array("완전요인설계")
            a = 2 ^ doe1.TextBox1.value
           ' doe1.ListBox2.list() = Array(a)
        ElseIf doe1.TextBox1 = 3 Or doe1.TextBox1 = 4 Then
          '  doe1.ListBox1.list() = Array("완전요인설계", "1/2 부분요인설계")
            a = 2 ^ doe1.TextBox1.value
            b = (1 / 2) * (2 ^ doe1.TextBox1.value)
          '  doe1.ListBox2.list() = Array(a, b)
        Else
           ' doe1.ListBox1.list() = Array("완전요인설계", "1/2 부분요인설계", "1/4 부분요인설계")
            a = 2 ^ doe1.TextBox1.value
            b = (1 / 2) * (2 ^ doe1.TextBox1.value)
            c = (1 / 4) * (2 ^ doe1.TextBox1.value)
           ' doe1.ListBox2.list() = Array(a, b, c)
        End If
End Sub

Private Sub ComboBox2_Change()
    Dim MyArray3 As Variant

    ComboBox3.value = "선택"
    '블록수
    If doe1.TextBox1.value = 2 Or (doe1.TextBox1.value = 5 And doe1.ListBox1.Selected(2) = True) Then
        If doe1.ComboBox2.value = 1 Then
            MyArray3 = [{1;2}]
        ElseIf doe1.ComboBox2.value = 2 Then
            MyArray3 = [{1;2;4}]
        ElseIf doe1.ComboBox2.value = 3 Then
            MyArray3 = [{1;2;3}]
        ElseIf doe1.ComboBox2.value = 4 Then
            MyArray3 = [{1;2;4}]
        Else
            MyArray3 = [{1;2;5}]
        End If
    ElseIf (doe1.TextBox1.value = 3 And doe1.ListBox1.Selected(1) = True) Then
        If doe1.ComboBox2.value = 1 Then
            MyArray3 = [{1}]
        ElseIf doe1.ComboBox2.value = 2 Then
            MyArray3 = [{1;2}]
        ElseIf doe1.ComboBox2.value = 3 Then
            MyArray3 = [{1;3}]
        ElseIf doe1.ComboBox2.value = 4 Then
            MyArray3 = [{1;2;4}]
        Else
            MyArray3 = [{1;5}]
        End If
    Else
        If doe1.ComboBox2.value = 1 Then
            MyArray3 = [{1;2;4}]
        ElseIf doe1.ComboBox2.value = 2 Then
            MyArray3 = [{1;2;4}]
        ElseIf doe1.ComboBox2.value = 3 Then
            MyArray3 = [{1;2;3;4}]
        ElseIf doe1.ComboBox2.value = 4 Then
            MyArray3 = [{1;2;4}]
        Else
            MyArray3 = [{1;2;4;5}]
        End If
    End If
    ComboBox3.list = MyArray3
End Sub

Private Sub btnCancel_Click()
Unload Me
End Sub
