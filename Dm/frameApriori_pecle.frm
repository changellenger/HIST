VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameApriori_pecle 
   OleObjectBlob   =   "frameApriori_pecle.frx":0000
   Caption         =   "APIRIOI"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7665
   StartUpPosition =   1  '������ ���
   TypeInfoVer     =   94
End
Attribute VB_Name = "frameApriori_pecle"
Attribute VB_Base = "0{748C4333-8BC9-48E8-825F-080FC79C653D}{2669D3D0-884F-4089-B4F7-784DB25C1830}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private Sub Label3_Click()

End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim i As Integer
     Dim j As Integer
     
    i = 0
    If Me.ListBox2.ListCount = 0 Then
        Do While i <= Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(i) = True Then
               Me.ListBox2.AddItem Me.ListBox1.List(i)
               Me.ListBox1.RemoveItem (i)
               Me.CB1.Visible = False
               Me.CB2.Visible = True
               Exit Sub
            End If
            i = i + 1
        Loop
    End If
    
     j = 0
    If Me.ListBox3.ListCount = 0 Then
        Do While j <= Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(j) = True Then
               Me.ListBox3.AddItem Me.ListBox1.List(j)
               Me.ListBox1.RemoveItem (j)
               Me.CB1.Visible = False
               Me.CB3.Visible = True
               Exit Sub
            End If
            j = j + 1
        Loop
    End If
    
    
End Sub
Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.ListBox2.ListCount <> 0 Then
        Me.ListBox1.AddItem Me.ListBox2.List(0)
        Me.ListBox2.RemoveItem (0)
        Me.CB1.Visible = True
        Me.CB2.Visible = False
    End If
End Sub
Private Sub ListBox3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.ListBox3.ListCount <> 0 Then
        Me.ListBox1.AddItem Me.ListBox3.List(0)
        Me.ListBox3.RemoveItem (0)
        Me.CB1.Visible = True
        Me.CB3.Visible = False
    End If

End Sub


Private Sub CB1_Click()
Dim i As Integer
    i = 0
    If Me.ListBox2.ListCount = 0 Then
        Do While i <= Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(i) = True Then
               Me.ListBox2.AddItem Me.ListBox1.List(i)
               Me.ListBox1.RemoveItem (i)
              
               Exit Sub
            End If
            i = i + 1
        Loop
    End If
End Sub
Private Sub CB2_Click()
If Me.ListBox2.ListCount <> 0 Then
        Me.ListBox1.AddItem ListBox2.List(0)
        Me.ListBox2.RemoveItem (0)
     
    End If
End Sub
Private Sub CB3_Click()
Dim i As Integer
    i = 0
    If Me.ListBox3.ListCount = 0 Then
        Do While i <= Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(i) = True Then
               Me.ListBox3.AddItem Me.ListBox1.List(i)
               Me.ListBox1.RemoveItem (i)
               Me.CB1.Visible = False
               Me.CB2.Visible = True
               Me.CB3.Visible = True
               Exit Sub
            End If
            i = i + 1
        Loop
    End If
End Sub

Private Sub CB4_Click()
If Me.ListBox3.ListCount <> 0 Then
        Me.ListBox1.AddItem ListBox3.List(0)
        Me.ListBox3.RemoveItem (0)
     
    End If

End Sub
Private Sub okbtn_Click()
       Dim dataRange As Range
    Dim i As Integer
    Dim activePt As Long                                '��� �м��� ���۵Ǵ� �κ��� �����ֱ� ����
    Dim rng As Range
    Dim a As String ' string  a  ����
    Dim b As String
    
   Dim j As Integer
      Dim dataRange2 As Range

    
    '''
    '''������ �������� �ʾ��� ���
    '''
    If Me.ListBox2.ListCount = 0 Then
        MsgBox "������ ������ �ֽñ� �ٶ��ϴ�.", vbExclamation, "HIST"
        Exit Sub
    End If
    
    '''
    '''public ���� ���� xlist, DataSheet, RstSheet, m, k1, n
    '''
    xlist = Me.ListBox2.List(0)
     xlist2 = Me.ListBox3.List(0)
    DataSheet = ActiveSheet.Name                        'DataSheet : Data�� �ִ� Sheet �̸�
    RstSheet = "_���м����_"                       'RstSheet  : ����� �����ִ� Sheet �̸�
    
    
    
    '������ �Է�
'On Error GoTo Err_delete
Dim val3535 As Long '�ʱ���ġ ������ ����'
Dim s3535 As Worksheet
val3535 = 2
    For Each s3535 In ActiveWorkbook.Sheets
        If s3535.Name = RstSheet Then
val3535 = Sheets(RstSheet).Cells(1, 1).Value
End If
Next s3535  '��Ʈ�� �̹������� ��� ��ġ �������ϰ�, ������ 2�� �����Ѵ�.


    
    Set dataRange = ActiveSheet.Cells.CurrentRegion
    m = dataRange.Columns.Count                         'm         : dataSheet�� �ִ� ���� ����
    
    tmp = 0
    For i = 1 To m
        If xlist = ActiveSheet.Cells(1, i) Then
            k1 = i  'k1                                 : k1 : ���õ� ������ ���° ���� �ִ���
            tmp = tmp + 1
        End If
    Next i
    
    n = ActiveSheet.Cells(1, k1).End(xlDown).Row - 1    'n         : ���õ� ������ ����Ÿ ����
    
   ' rng = Range(Cells(2, k1), Cells(N + 1, k1))
    
    
    

    '''
    ''' �������� ���� ��� - ������ ���� �ִ� ������ �ԷµǹǷ� ����ó���Ѵ�.
    '''
    If tmp > 1 Then
        MsgBox xlist & "�� ���� �������� �ֽ��ϴ�. " & vbCrLf & "�������� �ٲ��ֽñ� �ٶ��ϴ�.", vbExclamation, "HIST"
        Exit Sub
    End If
    
     Set dataRange2 = ActiveSheet.Cells.CurrentRegion
    m2 = dataRange2.Columns.Count
    
    
    tmp2 = 0
    For j = 1 To m2
        If xlist2 = ActiveSheet.Cells(1, j) Then
            k2 = j  'k1                                 : k1 : ���õ� ������ ���° ���� �ִ���
            tmp2 = tmp2 + 1
        End If
    Next j
    
    n2 = ActiveSheet.Cells(1, k2).End(xlDown).Row - 1    'n         : ���õ� ������ ����Ÿ ����
    
   ' rng = Range(Cells(2, k1), Cells(N + 1, k1))
    
    
    

    '''
    ''' �������� ���� ��� - ������ ���� �ִ� ������ �ԷµǹǷ� ����ó���Ѵ�.
    '''
    If tmp2 > 1 Then
        MsgBox xlist2 & "�� ���� �������� �ֽ��ϴ�. " & vbCrLf & "�������� �ٲ��ֽñ� �ٶ��ϴ�.", vbExclamation, "HIST"
        Exit Sub
    End If
    
    
    
    
      
    rinterface.StartRServer
    'Rinterface.PutArray "arraytest", Range(rng)
    rinterface.RRun "install.packages (" & Chr(34) & "arules" & Chr(34) & ")"
        rinterface.RRun "install.packages (" & Chr(34) & "arulesViz" & Chr(34) & ")"
    rinterface.RRun "require (arules)"
     rinterface.RRun "require (arulesViz)"
    
    
      rinterface.PutArray "arraytest", Range(Cells(2, k1), Cells(n + 1, k1))
      rinterface.PutArray "arraytest2", Range(Cells(2, k2), Cells(n2 + 1, k2))
      
      rinterface.PutDataframe "money", Range(Cells(1, k1), Cells(n2 + 1, k2))
      
    a = "list(money$Item, money$ID)" 'ok
    ' a = "list(arraytest, arraytest2)" 'ok
    rinterface.RRun a
                MsgBox a
    b = "alist <- split(money$Item, money$ID)" 'ok
    ' b = "alist <- split(arraytest, arraytest2)" 'ok
    rinterface.RRun b
    MsgBox b
        
                'money.trans <- as(money.list,"transactions")
      c = "trans<-as(alist," & Chr(34) & " transactions " & Chr(34) & " )"
      MsgBox c
      
      rinterface.RRun "trans <- as(alist," & Chr(34) & "transactions" & Chr(34) & " )"
      
      rinterface.RRun "rules <- apriori(trans) "
      rinterface.RRun "result <- inspect(rules)"
'      Rinterface.GetArray "result", Range("f1") '����

      
      
  
'a = " p <- qcc(arraytest, type= " & Chr(34) & "p" & Chr(34) & ", size= arraytest2, plot = FALSE)"
'Rinterface.RRun "qcc.options(" & Chr(34) & "beyond.limits" & Chr(34) & " = list(pch=15, col= " & Chr(34) & "orangered" & Chr(34) & "))"
'Rinterface.RRun " plot(p, title = " & Chr(34) & "P ������" & Chr(34) & ")"

'Rinterface.RRun a



'Rinterface.RRun "l <- limits.p(p$center,p$std.dev,p$sizes,3) "
'Rinterface.RRun "s <- stats.p(p$data, p$sizes)"
'Rinterface.RRun "ss <- as.data.frame(s)"
'Rinterface.RRun "d <- data.frame(l,ss)"
'Rinterface.RRun "result <- row.names(d[which(d$UCL <= d$statistics), ])"

'Rinterface.RRun "result3 <- t(result) "


'Rinterface.GetArray "result3", Range("sheet1!o3")

'Range("n1").Value = "P������ ����ؼ�"
'Range("n3").Value = "P�������Ѽ��� ����� �κб�:"




'If Range("o3").Value = "" Then
'Range("o7").Value = "������ �������¿� �ִ� ������ ������ �� �ֽ��ϴ�."
'Else
'Range("o4").Value = "��° �κб��� '�������Ѽ�'�� ������ϴ�. ���� ������ �̻������ �ִ� ������ �����˴ϴ�."
'End If




'Rinterface.InsertCurrentRPlot Range("sheet1!i3"), widthrescale:=0.5, heightrescale:=0.5, closergraph:=True

      'Rinterface.RRun "x1 <- matrix(data= arraytest, nrow = 4, ncol=5, byrow = FALSE, dimnames = NULL)"
'no2 = Me.TextBox.Value
'Rinterface.RRun "qcc(arraytest, type= " & Chr(34) & "p" & Chr(34) & ", size=3)"
'b = " qcc(x1, type= " & Chr(34) & "p" & Chr(34) & ", size=" & no2 & ")  "
'Rinterface.RRun b
   
      
   Unload Me


 
End Sub
Private Sub OptionButton1_Click()
   
   Dim myRange As Range
   Dim myArray()
   Dim arrName As Variant
   Dim TempSheet As Worksheet
   Set TempSheet = ActiveCell.Worksheet
   
    ReDim arrName(TempSheet.UsedRange.Columns.Count)
' Reading Data
    For i = 1 To TempSheet.UsedRange.Columns.Count
        arrName(i) = TempSheet.Cells(1, i)
    Next i
   
   Me.ListBox1.Clear
'-------------
  'Set myRange = Cells.CurrentRegion.Rows(1)
   'cnt = myRange.Cells.Count
   'ReDim myArray(cnt - 1)
  ' For i = 1 To cnt
  '   myArray(i - 1) = myRange.Cells(i)
  ' Next i
   'Me.ListBox1.List() = myArray
'-----------
    ReDim myArray(TempSheet.UsedRange.Columns.Count - 1)
    a = 0
   For i = 1 To TempSheet.UsedRange.Columns.Count
   If arrName(i) <> "" Then                     '��ĭ����
   myArray(a) = arrName(i)
   a = a + 1
   
   Else:
   End If
   Next i
   
   
   
   Me.ListBox1.List() = myArray
   
 '  For i = 1 To TempSheet.UsedRange.Columns.Count
 '   rngFirst.Offset(i, 1) = myArray(i - 1)
 ' Next i
  
End Sub

Private Sub UserForm_Click()

End Sub
