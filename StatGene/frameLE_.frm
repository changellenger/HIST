VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameLE_ 
   OleObjectBlob   =   "frameLE_.frx":0000
   Caption         =   "��л����"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7200
   StartUpPosition =   1  '������ ���
   TypeInfoVer     =   37
End
Attribute VB_Name = "frameLE_"
Attribute VB_Base = "0{A967C9EC-D076-4F9E-AFBF-3C74A6E5F48B}{9D925C18-8298-4C10-8597-29D4A0DD5E43}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False




Private Sub Cancel_Click()
    Unload Me
End Sub


Private Sub CB3_Click()
    Dim i As Integer
    i = 0
    If Me.ListBox2.ListCount = 0 Then
        Do While i <= Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(i) = True Then
               Me.ListBox2.AddItem Me.ListBox1.List(i)
               Me.ListBox1.RemoveItem (i)
               Me.CB3.Visible = False
               Me.CB4.Visible = True
               Exit Sub
            End If
            i = i + 1
        Loop
    End If
End Sub

Private Sub CB4_Click()
    If Me.ListBox2.ListCount <> 0 Then
        Me.ListBox1.AddItem ListBox2.List(0)
        Me.ListBox2.RemoveItem (0)
        Me.CB3.Visible = True
        Me.CB4.Visible = False
    End If
End Sub

Private Sub CB5_Click()
    Dim i As Integer
    i = 0
    If Me.ListBox5.ListCount = 0 Then
        Do While i <= Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(i) = True Then
               Me.ListBox5.AddItem Me.ListBox1.List(i)
               Me.ListBox1.RemoveItem (i)
               Me.CB5.Visible = False
               Me.CB6.Visible = True
               Exit Sub
            End If
            i = i + 1
        Loop
    End If
End Sub

Private Sub CB6_Click()
    If Me.ListBox5.ListCount <> 0 Then
        Me.ListBox1.AddItem ListBox5.List(0)
        Me.ListBox5.RemoveItem (0)
        Me.CB5.Visible = True
        Me.CB6.Visible = False
    End If
End Sub

Private Sub CommandButton1_Click()

       
    Dim choice(3) As Variant                            '�ѱ� ������ �븳����, �ŷڱ���,������� 3���ϱ�
    Dim dataRange As Range
    Dim i As Integer
    Dim activePt As Long                                '��� �м��� ���۵Ǵ� �κ��� �����ֱ� ����
    
    '''
    '''������ �������� �ʾ��� ���
    '''
    If Me.ListBox2.ListCount + Me.ListBox2.ListCount <> 2 Then
        MsgBox "2���� ������ ������ �ֽñ� �ٶ��ϴ�.", vbExclamation, "HIST"
        Exit Sub
    End If
    
    '''
    '''public ���� ���� xlist2, DataSheet, RstSheet, m, k2, n2
    '''
    ReDim xlist2(2)
    xlist2(1) = Me.ListBox2.List(0)
    
    MsgBox xlist2(1), vbExclamation, "HIST"
    xlist2(2) = Me.ListBox5.List(0)
     MsgBox xlist2(2), vbExclamation, "HIST"
    
    DataSheet = ActiveSheet.Name                        'DataSheet : Data�� �ִ� Sheet �̸�
    RstSheet = "_���м����_"                       'RstSheet  : ����� �����ִ� Sheet �̸�
    
    
    '������ �Է�
On Error GoTo Err_delete
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
    
    tmp1 = 2
    ReDim xlist2(tmp1)                                  '�����̸���
    ReDim k2(tmp1)                                      '���°���� ��������
    ReDim n2(tmp1)                                      '����Ÿ �������
    ReDim tmp(tmp1)
    
     i = 1
        tmp(1) = 0
        tmp(2) = 0
        For j = 1 To m
            If Me.ListBox2.List(0) = ActiveSheet.Cells(1, j) Then
                xlist2(1) = ActiveSheet.Cells(1, j)
                k2(1) = j
                n2(1) = ActiveSheet.Cells(1, j).End(xlDown).row - 1
            '    tmp(i) = tmp(i) + 1
            End If
               If Me.ListBox5.List(0) = ActiveSheet.Cells(1, j) Then
                xlist2(2) = ActiveSheet.Cells(1, j)
                k2(2) = j
                n2(2) = ActiveSheet.Cells(1, j).End(xlDown).row - 1
             '   tmp(i) = tmp(i) + 1
            End If
    Next j
    tmp(1) = 1
    tmp(2) = 1
    
    
    '''
    ''' �������� ���� ��� - ������ ���� �ִ� ������ �ԷµǹǷ� ����ó���Ѵ�.
    '''
    For i = 1 To tmp1
    If tmp(i) > 1 Then
        MsgBox xlist2(i) & "�� ���� �������� �ֽ��ϴ�. " & vbCrLf & "�������� �ٲ��ֽñ� �ٶ��ϴ�.", vbExclamation, "HIST"
        Exit Sub
    End If
    Next i
    
         
    '''
    '''���ڿ� ���ڰ� ȥ�յǾ� ���� ���
    '''
    For i = 1 To tmp1
       If TModuleControl.FindingRangeError(xlist2(i)) = True Then
           MsgBox "������ �м������� ���ڳ� ������ �ֽ��ϴ�." & Chr(10) & _
                    ": " & xlist2(i), vbExclamation, "HIST"
            Exit Sub
        End If
    Next i
        
    '''
    '''�������� ���ð�� �Է� - choice(1)
    '''
    choice(1) = 1
  
    If choice(1) = 2 Then
        If n2(1) <> n2(2) Then
        MsgBox "������ ������ ����Ÿ������ �ٸ��ϴ�. �����񱳸� �� �� �����ϴ�.", vbExclamation, "HIST": Exit Sub
        Exit Sub
        End If
    End If
    
    '''
    ''' ����Ÿ ������ �Ѱ��� ���
    '''
    If n2(1) = 1 Or n2(2) = 1 Then
        MsgBox "�� ���� ����Ÿ�� ������ ������ �� �����ϴ�.", vbExclamation, "HIST"
        Exit Sub
    End If

    '''
    '''�ŷڱ����� �߸� �Է��� ���
    '''
   ' If Me.ChB2.Value = True Then
   '     If IsNumeric(Me.TextBox4.Value) = False Then
   '         MsgBox "����� �ŷڱ����� �Է��� �ֽñ� �ٶ��ϴ�.", vbExclamation, "HIST"
   '         Exit Sub
   '     ElseIf Me.TextBox4.Value < 0 Or Me.TextBox4.Value > 100 Then
   ''         MsgBox "����� �ŷڱ����� %������ �Է��� �ֽñ� �ٶ��ϴ�.", vbExclamation, "HIST"
   '         Exit Sub
  '      End If
  '  End If
    
    '''
    '''�ŷڱ��� �Է� - choice(2)
    '''
    If Me.ChB1.Value = True Then choice(2) = Me.TextBox4.Value
   ' If Me.ChB2.Value = False Then choice(2) = -1
    
    '''
    '''�͹����� ���ð�� �Է� - choice(3)
    '''
    If Me.OB4 = True Then choice(3) = 1
    If Me.OB5 = True Then choice(3) = 2
    If Me.OB6 = True Then choice(3) = 3
                
      '''
    '''��� ó��
    '''
    TModuleControl.SettingStatusBar True, "��ǥ�� f-�������Դϴ�."
    Application.ScreenUpdating = False
    TModulePrint.makeOutputSheet (RstSheet)
    'Worksheets(RstSheet).Unprotect "prophet"
    activePt = Worksheets(RstSheet).Cells(1, 1).Value
    
    TModuleControl.analle choice, 1
    
    
    TModuleControl.SettingStatusBar False
    Application.ScreenUpdating = True
    Unload Me
    
    Worksheets(RstSheet).Activate
    
    '���� ���� üũ �� �񱳰� ����
    Dim Cmp_Value As Long
    
    If PublicModule.ChkVersion(ActiveWorkbook.Name) = True Then
        Cmp_Value = 1048000
    Else
        Cmp_Value = 65000
    End If
    
    If Worksheets(RstSheet).Cells(1, 1).Value > Cmp_Value Then
        MsgBox "[_���м����_]��Ʈ�� ���� ��� ����Ͽ����ϴ�." & vbCrLf & "�� ��Ʈ�� �̸��� �ٲٰų� ������ �ּ���", vbExclamation, "HIST"
        Exit Sub
    End If

    Worksheets(RstSheet).Cells(activePt + 10, 1).Select
    Worksheets(RstSheet).Cells(activePt + 10, 1).Activate
                            '��� �м��� ���۵Ǵ� �κ��� �����ָ� ��ģ��.
                                           




'�ǵڿ� ���̱�
Exit Sub
Err_delete:

For Each s3535 In ActiveWorkbook.Sheets
        If s3535.Name = RstSheet Then
Sheets(RstSheet).Range(Cells(val3535, 1), Cells(5000, 1000)).Select
Selection.Delete
Sheets(RstSheet).Cells(1, 1) = val3535
Sheets(RstSheet).Cells(val3535, 1).Select

If val3535 = 2 Then
Application.DisplayAlerts = False
Sheets(RstSheet).Delete
End If

End If


Next s3535

MsgBox ("���α׷��� ������ �ֽ��ϴ�.")
 'End sub �տ��� ���δ�.

''�ؼ�, ������ ���� Err_delete�� �ͼ� ù�����ķ� �����. ���� ù���� 2�� ��Ʈ�� �����.�׸��� �����޽��� ���
'rSTsheet����⵵ ���� �������� ��쿡�� �ƹ� ���۵� ���� �ʰ�, �����޽����� ����.
        
End Sub


Private Sub CommandButton2_Click()

End Sub

Private Sub Label16_Click()
    frameDEx4.Show
    
End Sub

Private Sub ListBox3_Click()

End Sub

Private Sub Label7_Click()

End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim i As Integer
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


   ElseIf Me.ListBox5.ListCount = 0 Then
        Do While j <= Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(j) = True Then
               Me.ListBox5.AddItem Me.ListBox1.List(j)
               Me.ListBox1.RemoveItem (j)
               Me.CB1.Visible = False
               Me.CB2.Visible = True
               Exit Sub
            End If
            j = j + 1
        Loop
    End If
    
End Sub

Private Sub ListBox4_Click()

End Sub

Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.ListBox2.ListCount <> 0 Then
        Me.ListBox1.AddItem Me.ListBox2.List(0)
        Me.ListBox2.RemoveItem (0)
        Me.CB1.Visible = True
        Me.CB2.Visible = False
    End If
End Sub

Private Sub ListBox5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.ListBox5.ListCount <> 0 Then
        Me.ListBox1.AddItem Me.ListBox5.List(0)
        Me.ListBox5.RemoveItem (0)
        Me.CB1.Visible = True
        Me.CB2.Visible = False
    End If
End Sub

Private Sub OK_Click()

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
    
    
    Dim j As Integer
    j = 0
    If Me.ListBox5.ListCount = 0 Then
        Do While j <= Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(j) = True Then
               Me.ListBox5.AddItem Me.ListBox1.List(j)
               Me.ListBox1.RemoveItem (j)

               Exit Sub
            End If
            j = j + 1
        Loop
    End If
    
    
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
  


    For j = 1 To TempSheet.UsedRange.Columns.Count
        arrName(j) = TempSheet.Cells(1, j)
    Next j
   
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
   For j = 1 To TempSheet.UsedRange.Columns.Count
   If arrName(j) <> "" Then                     '��ĭ����
   myArray(a) = arrName(j)
   a = a + 1
   
   Else:
   End If
   Next j
   
   
   
   Me.ListBox1.List() = myArray
   
 '  For i = 1 To TempSheet.UsedRange.Columns.Count
 '   rngFirst.Offset(i, 1) = myArray(i - 1)
 ' Next i
  




End Sub



Private Sub UserForm_Click()

End Sub
