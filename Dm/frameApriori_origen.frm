VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameApriori_origen 
   OleObjectBlob   =   "frameApriori_origen.frx":0000
   Caption         =   "APIRIORI"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7545
   StartUpPosition =   1  '������ ���
   TypeInfoVer     =   94
End
Attribute VB_Name = "frameApriori_origen"
Attribute VB_Base = "0{7C43E65C-D395-4688-95D4-EFE159121546}{6C7DE8B1-01F4-4C04-8054-7236DB3004DE}"
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
      Dim xlist As String
      Dim xlist2 As String
      

    
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
rinterface.RRun "install.packages (" & Chr(34) & "arules" & Chr(34) & ")"
rinterface.RRun "install.packages (" & Chr(34) & "arulesViz" & Chr(34) & ")"
rinterface.RRun "install.packages (" & Chr(34) & "grid" & Chr(34) & ")"

rinterface.RRun "require(grid)"
rinterface.RRun "library (arules)"

rinterface.RRun "require (arulesViz) "
rinterface.RRun "require (arules)"

'=====================��Ű�� ��ġ �Ϸ�================
    
      rinterface.PutDataframe xlist, Range(Cells(1, k1), Cells(n + 1, k1))
      rinterface.PutDataframe xlist2, Range(Cells(1, k2), Cells(n2 + 1, k2))
      
      rinterface.RRun " frmset <- cbind(" & xlist & "," & xlist2 & " )"
    ' rinterface.PutDataframe "money", Range(Cells(1, k1), Cells(N2 + 1, k2))
      rinterface.RRun " attach(frmset)"
 '===================== ������ �Է� �� ���� ==================
    
      
      a = "list(frmset$" & xlist & ", frmset$" & xlist2 & ")" 'ok
    ' a = "list(arraytest, arraytest2)" 'ok
      rinterface.RRun a
                MsgBox a
      b = "frmset.list <- split(frmset$" & xlist & ", frmset$" & xlist2 & ")" 'ok
    ' b = "alist <- split(arraytest, arraytest2)" 'ok
      rinterface.RRun b
      MsgBox b
      
            
      rinterface.RRun "frmset.trans<- as(frmset.list, " & Chr(34) & "transactions" & Chr(34) & ") "
     ' money.trans <- as(money.list,"transactions")
      rinterface.RRun "frmset.rules<-apriori(frmset.trans)"
      rinterface.RRun "ro<-as(frmset.rules, " & Chr(34) & "data.frame" & Chr(34) & ")"  ' rules ���, �� ��°����� Ÿ��
      rinterface.RRun "r <-inspect(frmset.rules)"
      
      rinterface.RRun "top.frmset.rules<- head(sort(frmset.rules, decreasing = TRUE, by=" & Chr(34) & "lift" & Chr(34) & "))"
      rinterface.RRun "inspect(top.frmset.rules)"
      
      
      
      rinterface.RRun "plot(frmset.rules, method= " & Chr(34) & "grouped" & Chr(34) & " ) "
       '������Ģ�� ���ǰ� ����� �������� �׷����� ������. ������ ���ϱ� ���. ����ũ��=������. ����(LHS)���� ���ڴ� �� �������� �Ǿ� �ִ� ������Ģ�� ��. +���ڴ� ������ ��ǰ'
     rinterface.RRun "plot(frmset.rules, method= " & Chr(34) & "graph" & Chr(34) & ") "  '��ǰ�� ���� ������. ȭ��ǥ �β� = ������ , ȭ��ǥ ���ϱ� = ���.
   
    
      'rinterface.RRun "plot(frmset.rules, measure = c(" & Chr(34) & "support" & Chr(34) & "," & Chr(34) & " lift" & Chr(34) & "), shading = " & Chr(34) & "confidence" & Chr(34) & ")"

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
