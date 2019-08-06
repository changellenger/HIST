VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} doe1 
   OleObjectBlob   =   "doe1.frx":0000
   Caption         =   "요인 설계 생성(2수준)"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4725
   StartUpPosition =   1  '소유자 가운데
   TypeInfoVer     =   246
End
Attribute VB_Name = "doe1"
Attribute VB_Base = "0{CB2DE3B0-59F2-431B-B90A-EFAA50F4BD09}{C523AC7D-45CD-4B9E-949C-195FF2655E30}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Dim Aname As String
Dim Ahigh, Alow As Integer

Dim Bname As String
Dim Bhigh, Blow As Integer

 

' 요인설정 페이지로 넘겨야 함'
Private Sub CommandButton1_Click()
    Dim a, b, c As Integer
    'If OptionButton1.value = True Then
        If doe1.TextBox1.value = 2 Then
            doe1.ListBox1.List() = Array("완전요인설계")
            a = 2 ^ doe1.TextBox1.value
            doe1.ListBox2.List() = Array(a)
        ElseIf doe1.TextBox1 = 3 Or doe1.TextBox1 = 4 Then
            doe1.ListBox1.List() = Array("완전요인설계", "1/2 부분요인설계")
            a = 2 ^ doe1.TextBox1.value
            b = (1 / 2) * (2 ^ doe1.TextBox1.value)
            doe1.ListBox2.List() = Array(a, b)
        Else
            doe1.ListBox1.List() = Array("완전요인설계", "1/2 부분요인설계", "1/4 부분요인설계")
            a = 2 ^ doe1.TextBox1.value
            b = (1 / 2) * (2 ^ doe1.TextBox1.value)
            c = (1 / 4) * (2 ^ doe1.TextBox1.value)
            doe1.ListBox2.List() = Array(a, b, c)
        End If
    'End If
    doe1.Show
End Sub



Private Sub ComboBox3_Change()

End Sub

Private Sub CommandButton3_Click()
ShellExecute 0, "open", "hh.exe", ThisWorkbook.Path + "\KESS%202013.chm::/요인%20설계.htm", "", 1
End Sub

Private Sub CommandButton4_Click()
    Dim nfact1, a1, p, q, value, co As Integer
    Dim wsheet2 As Worksheet
    Dim resultsheet, s As Worksheet
    
    Dim a As String
    Dim b As String
    Dim Drng As Range
   
    Dim blocksN As Integer
    Dim runsN As Integer
    Dim replicationsN As Integer
    Dim factorsN As Integer
    Dim centersN As Integer
    Dim numberK As Integer

    
    
 


        For I = 0 To doe1.ListBox1.ListCount - 1
            If doe1.ListBox1.Selected(I) = True Then
                co = co + 1
            End If
        Next I
        
        If co = 0 Then
            MsgBox "설계 창에서 요인설계방법을 선택하여주세요."
            Exit Sub
        ElseIf doe1.ComboBox2.value = "선택" Then
            MsgBox "설계 창에서 요인의 반복수를 선택하여주세요."
            Exit Sub
        ElseIf doe1.ComboBox3.value = "선택" Then
            MsgBox "설계 창에서 블록의 수를 선택하여 주세요."
            Exit Sub
        End If
        
        
        
        
    Set resultsheet = OpenOutSheet("_통계분석결과_", True)
    activePt = resultsheet.Cells(1, 1).value
    
    blocksN = ComboBox3.value ' 블록수
    replicationsN = ComboBox2.value ' 반복수
    factorsN = TextBox1.value ' 요인수
    centersN = ComboBox4.value '중심점 추가
    
    nfact1 = TextBox1.value  '요인개수
    co = 0
    Set wsheet2 = Worksheets.Add
    For Each s In ActiveWorkbook.Sheets
        If Left(s.name, 4) = "요인분석" Then
           If Right(s.name, 1) > co Then
           co = Right(s.name, 1)
           End If
        End If
    Next s
    wsheet2.name = "요인분석입니다" & co + 1            '오류 발생시 시트 이름 겹침 ( 시트 삭제 요망)
    With ActiveWindow.Application.Cells
         .Font.name = "맑은 고딕"
         .Font.Size = 11
         .HorizontalAlignment = xlRight
         
    End With
    wsheet2.Range("A1").Select
    Selection.value = "블록"
    For I = 1 To nfact1
        Selection.Offset(0, I).value = "요인" & I
    Next I
    co = 0
    '2수준 요인설계
    'If OptionButton1.value = True Then
        a11 = 2 ^ Me.TextBox1.value   '자료 개수
        
        
        
        '완전요인설계
        If doe1.ListBox1.Selected(0) = True Then
            'TwoLevelFD.fullFD a11, nfact1, wsheet1
            'TwoLevel_Result.CFDresult resultsheet, "완전", nfact1, 2 ^ nfact1, doe1.ComboBox3.value, _
                                        'doe1.ComboBox2.value, doe1.ComboBox4.value, wsheet1.name
                  
                'Design.5 <- fac.design(nfactors= 2 ,replications= 3 ,repeat.only= TRUE ,blocks= 1 ,randomize= TRUE ,seed= 13692 ,nlevels=c(3,3), factor.names=list( A=c(100,125,150), B=c(1,2,3) ))
  
          Rinterface.StartRServer
        ' Rinterface.PutArray "arraytest", Range(rng)
          Rinterface.RRun "install.packages (" & Chr(34) & "DoE.base" & Chr(34) & ")"
          Rinterface.RRun "require (DoE.base)"
          
          numberK = doe1.TextBox2.value
         a = " Design <- fac.design(nfactors= " & factorsN & " ,replications= " & replicationsN & " ,repeat.only= True  ,blocks= " & blocksN & " ,randomize=  TRUE ,seed= 13692 ,nlevels=c(2,2) , factor.names = list( " & Aname & "=c(" & Alow & " , " & Ahigh & " ) , " & Bname & "=c(" & Blow & " , " & Bhigh & ") ))"
         'a = " Design <- fac.design(nfactors= 2, replications= 4 ,repeat.only= TRUE , blocks= 1 ,randomize= TRUE ,seed= 13692 ,nlevels=c(2,2), factor.names=list( A=c(10,20), B=c(30,40) ))"

        'a = "Design" & numberK & " <- fac.design(nfactors= " & factorsN & " ,replications= " & replicationsN & " ,repeat.only= True  ,blocks= " & blocksN & " ,randomize=  TRUE ,seed= 13692 ,nlevels=c(2,2) , factor.names = list( " & Aname & "=c(" & Alow & " , " & Ahigh & " ) , " & Bname & "=c(" & Blow & " , " & Bhigh & ") ))"

        MsgBox a
        
         Rinterface.RRun a
         
         Dim dsn As String
         dsn = "Design" & numberK
      '      Rinterface.GetArray "dsn", Range("B2")
      
      Rinterface.GetArray "Design", Range(" B2 ")
       'Rinterface.RRun "qcc(x1, type= " & Chr(34) & s"xbar" & Chr(34) & ", nsigmas=3)"
        
        End If
        
        '1/2 부분요인설계
        If doe1.ListBox1.Selected(1) = True Then
         
         Rinterface.RRun "install.packages (" & Chr(34) & "FrF2" & Chr(34) & ")"
          Rinterface.RRun "require (FrF2)"
         
         b = "Design6 <- FrF2(nruns= " & runsN & " ,nfactors= " & factorsN & " , blocks= " & blocksN & " , alias.block.2fis = FALSE , ncenter= " & centersN & " , MaxC2 = FALSE , resolution = NULL ,replications= 1 , repeat.only= FALSE ,randomize= TRUE ,seed= 17524 , factor.names=list( " & Aname & "=c(" & Alow & " , " & Ahigh & " ) , " & Bname & "=c(" & Blow & " , " & Bhigh & ") )"
         Rinterface.RRun b
            'TwoLevelFD.halfFD a11, nfact1, wsheet1
            'TwoLevel_Result.CFDresult resultsheet, "1/2 부분", nfact1, (1 / 2) * (2 ^ nfact1), doe1.ComboBox3.value, _
                                        doe1.ComboBox2.value, doe1.ComboBox4.value, wsheet1.name
        End If
        
        '1/4 부분요인설계
        If doe1.ListBox1.Selected(2) = True Then
         
         Rinterface.RRun "install.packages (" & Chr(34) & "FrF2" & Chr(34) & ")"
          Rinterface.RRun "require (FrF2)"
          b = "Design.6 <- FrF2(nruns= " & runsN & " ,nfactors= " & factorsN & " , blocks= " & blocksN & " , alias.block.2fis = FALSE , ncenter= " & centersN & " , MaxC2 = FALSE , resolution = NULL ,replications= 1 , repeat.only= FALSE ,randomize= TRUE ,seed= 17524 , factor.names=list( A=c(-1,1),B=c(-1,1),C=c(-1,1),D=c(-1,1) ) )"
         Rinterface.RRun b
            
            
            'TwoLevelFD.quarterFD a11, nfact1, wsheet1
            'TwoLevel_Result.CFDresult resultsheet, "1/4 부분", nfact1, (1 / 4) * (2 ^ nfact1), doe1.ComboBox3.value, _
                                        doe1.ComboBox2.value, doe1.ComboBox4.value, wsheet1.name
        End If
        
        Me.Hide
    'End If

    '3수준 요인설계
    'If OptionButton2.value = True Then
        
    'End If

    '완전요인설계
    'If OptionButton3.value = True Then
    '
    'End If
    'TwoLevel_Result.dResult

    ''resultsheet.Activate
    wsheet2.Activate
    doe1.TextBox2.value = numberK + 1
    
    Unload Me
    
End Sub

Private Sub CommandButton5_Cslick()
    Unload Me
End Sub

Private Sub doe2showbtn_Click() ' 요인설정 버튼 클릭시 수준값과 이름을 설정할 수 있는 화면이 뜬다

doe2.Show

End Sub

Public Sub passing()

Aname = doe2.TextBox3.value
Alow = doe2.TextBox4.value
Ahigh = doe2.TextBox5.value


Bname = doe2.TextBox6.value
Blow = doe2.TextBox7.value
Bhigh = doe2.TextBox8.value

End Sub

Private Sub ListBox1_Click()
    If ListBox1.Selected(0) = True Then
        ListBox2.Selected(0) = True
        ComboBox2.value = "선택"
        ComboBox3.value = "선택"
        ComboBox3.Clear
    End If
    If ListBox1.Selected(1) = True Then
        ListBox2.Selected(1) = True
        ComboBox2.value = "선택"
        ComboBox3.value = "선택"
        ComboBox3.Clear
    End If
    If ListBox1.Selected(2) = True Then
        ListBox2.Selected(2) = True
        ComboBox2.value = "선택"
        ComboBox3.value = "선택"
        ComboBox3.Clear
    End If
End Sub

Private Sub ListBox2_Click()
    If ListBox2.Selected(0) = True Then
        ListBox1.Selected(0) = True
        ComboBox2.value = "선택"
        ComboBox3.value = "선택"
        ComboBox3.Clear
    End If
    If ListBox2.Selected(1) = True Then
        ListBox1.Selected(1) = True
        ComboBox2.value = "선택"
        ComboBox3.value = "선택"
        ComboBox3.Clear
    End If
    If ListBox2.Selected(2) = True Then
        ListBox1.Selected(2) = True
        ComboBox2.value = "선택"
        ComboBox3.value = "선택"
        ComboBox3.Clear
    End If
End Sub



Private Sub SpinButton1_Change()
    Me.TextBox1.value = SpinButton1.value
    Dim a, b, c As Integer
    
    
        If doe1.TextBox1.value = 2 Then
            doe1.ListBox1.List() = Array("완전요인설계")
            a = 2 ^ doe1.TextBox1.value
            doe1.ListBox2.List() = Array(a)
        ElseIf doe1.TextBox1 = 3 Or doe1.TextBox1 = 4 Then
            doe1.ListBox1.List() = Array("완전요인설계", "1/2 부분요인설계")
            a = 2 ^ doe1.TextBox1.value
            b = (1 / 2) * (2 ^ doe1.TextBox1.value)
            doe1.ListBox2.List() = Array(a, b)
        Else
            doe1.ListBox1.List() = Array("완전요인설계", "1/2 부분요인설계", "1/4 부분요인설계")
            a = 2 ^ doe1.TextBox1.value
            b = (1 / 2) * (2 ^ doe1.TextBox1.value)
            c = (1 / 4) * (2 ^ doe1.TextBox1.value)
            doe1.ListBox2.List() = Array(a, b, c)
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
    ComboBox3.List = MyArray3
End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub UserForm_Click()

End Sub
