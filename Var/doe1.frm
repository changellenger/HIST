VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} doe1 
   OleObjectBlob   =   "doe1.frx":0000
   Caption         =   "���� ���� ����(2����)"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4725
   StartUpPosition =   1  '������ ���
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

 

' ���μ��� �������� �Ѱܾ� ��'
Private Sub CommandButton1_Click()
    Dim a, b, c As Integer
    'If OptionButton1.value = True Then
        If doe1.TextBox1.value = 2 Then
            doe1.ListBox1.List() = Array("�������μ���")
            a = 2 ^ doe1.TextBox1.value
            doe1.ListBox2.List() = Array(a)
        ElseIf doe1.TextBox1 = 3 Or doe1.TextBox1 = 4 Then
            doe1.ListBox1.List() = Array("�������μ���", "1/2 �κп��μ���")
            a = 2 ^ doe1.TextBox1.value
            b = (1 / 2) * (2 ^ doe1.TextBox1.value)
            doe1.ListBox2.List() = Array(a, b)
        Else
            doe1.ListBox1.List() = Array("�������μ���", "1/2 �κп��μ���", "1/4 �κп��μ���")
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
ShellExecute 0, "open", "hh.exe", ThisWorkbook.Path + "\KESS%202013.chm::/����%20����.htm", "", 1
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
            MsgBox "���� â���� ���μ������� �����Ͽ��ּ���."
            Exit Sub
        ElseIf doe1.ComboBox2.value = "����" Then
            MsgBox "���� â���� ������ �ݺ����� �����Ͽ��ּ���."
            Exit Sub
        ElseIf doe1.ComboBox3.value = "����" Then
            MsgBox "���� â���� ����� ���� �����Ͽ� �ּ���."
            Exit Sub
        End If
        
        
        
        
    Set resultsheet = OpenOutSheet("_���м����_", True)
    activePt = resultsheet.Cells(1, 1).value
    
    blocksN = ComboBox3.value ' ��ϼ�
    replicationsN = ComboBox2.value ' �ݺ���
    factorsN = TextBox1.value ' ���μ�
    centersN = ComboBox4.value '�߽��� �߰�
    
    nfact1 = TextBox1.value  '���ΰ���
    co = 0
    Set wsheet2 = Worksheets.Add
    For Each s In ActiveWorkbook.Sheets
        If Left(s.name, 4) = "���κм�" Then
           If Right(s.name, 1) > co Then
           co = Right(s.name, 1)
           End If
        End If
    Next s
    wsheet2.name = "���κм��Դϴ�" & co + 1            '���� �߻��� ��Ʈ �̸� ��ħ ( ��Ʈ ���� ���)
    With ActiveWindow.Application.Cells
         .Font.name = "���� ���"
         .Font.Size = 11
         .HorizontalAlignment = xlRight
         
    End With
    wsheet2.Range("A1").Select
    Selection.value = "���"
    For I = 1 To nfact1
        Selection.Offset(0, I).value = "����" & I
    Next I
    co = 0
    '2���� ���μ���
    'If OptionButton1.value = True Then
        a11 = 2 ^ Me.TextBox1.value   '�ڷ� ����
        
        
        
        '�������μ���
        If doe1.ListBox1.Selected(0) = True Then
            'TwoLevelFD.fullFD a11, nfact1, wsheet1
            'TwoLevel_Result.CFDresult resultsheet, "����", nfact1, 2 ^ nfact1, doe1.ComboBox3.value, _
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
        
        '1/2 �κп��μ���
        If doe1.ListBox1.Selected(1) = True Then
         
         Rinterface.RRun "install.packages (" & Chr(34) & "FrF2" & Chr(34) & ")"
          Rinterface.RRun "require (FrF2)"
         
         b = "Design6 <- FrF2(nruns= " & runsN & " ,nfactors= " & factorsN & " , blocks= " & blocksN & " , alias.block.2fis = FALSE , ncenter= " & centersN & " , MaxC2 = FALSE , resolution = NULL ,replications= 1 , repeat.only= FALSE ,randomize= TRUE ,seed= 17524 , factor.names=list( " & Aname & "=c(" & Alow & " , " & Ahigh & " ) , " & Bname & "=c(" & Blow & " , " & Bhigh & ") )"
         Rinterface.RRun b
            'TwoLevelFD.halfFD a11, nfact1, wsheet1
            'TwoLevel_Result.CFDresult resultsheet, "1/2 �κ�", nfact1, (1 / 2) * (2 ^ nfact1), doe1.ComboBox3.value, _
                                        doe1.ComboBox2.value, doe1.ComboBox4.value, wsheet1.name
        End If
        
        '1/4 �κп��μ���
        If doe1.ListBox1.Selected(2) = True Then
         
         Rinterface.RRun "install.packages (" & Chr(34) & "FrF2" & Chr(34) & ")"
          Rinterface.RRun "require (FrF2)"
          b = "Design.6 <- FrF2(nruns= " & runsN & " ,nfactors= " & factorsN & " , blocks= " & blocksN & " , alias.block.2fis = FALSE , ncenter= " & centersN & " , MaxC2 = FALSE , resolution = NULL ,replications= 1 , repeat.only= FALSE ,randomize= TRUE ,seed= 17524 , factor.names=list( A=c(-1,1),B=c(-1,1),C=c(-1,1),D=c(-1,1) ) )"
         Rinterface.RRun b
            
            
            'TwoLevelFD.quarterFD a11, nfact1, wsheet1
            'TwoLevel_Result.CFDresult resultsheet, "1/4 �κ�", nfact1, (1 / 4) * (2 ^ nfact1), doe1.ComboBox3.value, _
                                        doe1.ComboBox2.value, doe1.ComboBox4.value, wsheet1.name
        End If
        
        Me.Hide
    'End If

    '3���� ���μ���
    'If OptionButton2.value = True Then
        
    'End If

    '�������μ���
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

Private Sub doe2showbtn_Click() ' ���μ��� ��ư Ŭ���� ���ذ��� �̸��� ������ �� �ִ� ȭ���� ���

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
        ComboBox2.value = "����"
        ComboBox3.value = "����"
        ComboBox3.Clear
    End If
    If ListBox1.Selected(1) = True Then
        ListBox2.Selected(1) = True
        ComboBox2.value = "����"
        ComboBox3.value = "����"
        ComboBox3.Clear
    End If
    If ListBox1.Selected(2) = True Then
        ListBox2.Selected(2) = True
        ComboBox2.value = "����"
        ComboBox3.value = "����"
        ComboBox3.Clear
    End If
End Sub

Private Sub ListBox2_Click()
    If ListBox2.Selected(0) = True Then
        ListBox1.Selected(0) = True
        ComboBox2.value = "����"
        ComboBox3.value = "����"
        ComboBox3.Clear
    End If
    If ListBox2.Selected(1) = True Then
        ListBox1.Selected(1) = True
        ComboBox2.value = "����"
        ComboBox3.value = "����"
        ComboBox3.Clear
    End If
    If ListBox2.Selected(2) = True Then
        ListBox1.Selected(2) = True
        ComboBox2.value = "����"
        ComboBox3.value = "����"
        ComboBox3.Clear
    End If
End Sub



Private Sub SpinButton1_Change()
    Me.TextBox1.value = SpinButton1.value
    Dim a, b, c As Integer
    
    
        If doe1.TextBox1.value = 2 Then
            doe1.ListBox1.List() = Array("�������μ���")
            a = 2 ^ doe1.TextBox1.value
            doe1.ListBox2.List() = Array(a)
        ElseIf doe1.TextBox1 = 3 Or doe1.TextBox1 = 4 Then
            doe1.ListBox1.List() = Array("�������μ���", "1/2 �κп��μ���")
            a = 2 ^ doe1.TextBox1.value
            b = (1 / 2) * (2 ^ doe1.TextBox1.value)
            doe1.ListBox2.List() = Array(a, b)
        Else
            doe1.ListBox1.List() = Array("�������μ���", "1/2 �κп��μ���", "1/4 �κп��μ���")
            a = 2 ^ doe1.TextBox1.value
            b = (1 / 2) * (2 ^ doe1.TextBox1.value)
            c = (1 / 4) * (2 ^ doe1.TextBox1.value)
            doe1.ListBox2.List() = Array(a, b, c)
        End If
End Sub

Private Sub ComboBox2_Change()
    Dim MyArray3 As Variant

    ComboBox3.value = "����"
    '��ϼ�
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
