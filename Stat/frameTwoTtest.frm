VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameTwoTtest 
   OleObjectBlob   =   "frameTwoTtest.frx":0000
   Caption         =   "����ǥ�� t ����"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5295
   StartUpPosition =   1  '������ ���
   TypeInfoVer     =   55
End
Attribute VB_Name = "frameTwoTtest"
Attribute VB_Base = "0{03E6A100-3F54-40DD-AAD6-474E8D85D54A}{EA102E55-B4AA-4D6A-B9D5-6AF940465899}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private Sub Cancel1_Click()
    Unload Me
End Sub

Private Sub Cancel2_Click()
    Unload Me
End Sub

Private Sub CB1_Click()
    Dim i As Integer
    i = 0
    If Me.ListBox1.ListCount > 0 Then
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
        Me.ListBox1.AddItem Me.ListBox2.List(0)
        Me.ListBox2.RemoveItem (0)
    End If
End Sub

Private Sub CB3_Click()
    Dim i As Integer
    i = 0
    If Me.ListBox4.ListCount = 0 Then
        Do While i <= Me.ListBox3.ListCount - 1
            If Me.ListBox3.Selected(i) = True Then
               Me.ListBox4.AddItem Me.ListBox3.List(i)
               Me.ListBox3.RemoveItem (i)
               Me.CB3.Visible = False
               Me.CB4.Visible = True
               Exit Sub
            End If
            i = i + 1
        Loop
    End If
End Sub

Private Sub CB4_Click()
    If Me.ListBox4.ListCount <> 0 Then
        Me.ListBox3.AddItem ListBox4.List(0)
        Me.ListBox4.RemoveItem (0)
        Me.CB3.Visible = True
        Me.CB4.Visible = False
    End If
End Sub

Private Sub CB5_Click()
    Dim i As Integer
    i = 0
    If Me.ListBox5.ListCount = 0 Then
        Do While i <= Me.ListBox3.ListCount - 1
            If Me.ListBox3.Selected(i) = True Then
               Me.ListBox5.AddItem Me.ListBox3.List(i)
               Me.ListBox3.RemoveItem (i)
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
        Me.ListBox3.AddItem ListBox5.List(0)
        Me.ListBox5.RemoveItem (0)
        Me.CB5.Visible = True
        Me.CB6.Visible = False
    End If
End Sub



Private Sub ChB1_Click()
    If Me.ChB1.Value = True Then Me.TextBox2.Enabled = True
    If Me.ChB1.Value = False Then Me.TextBox2.Enabled = False
End Sub

Private Sub ChB2_Click()
    If Me.ChB2.Value = True Then Me.TextBox4.Enabled = True
    If Me.ChB2.Value = False Then Me.TextBox4.Enabled = False
End Sub


Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim i As Integer
    i = 0
    If Me.ListBox1.ListCount > 0 Then
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

Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.ListBox2.ListCount <> 0 Then
        Me.ListBox1.AddItem Me.ListBox2.List(0)
        Me.ListBox2.RemoveItem (0)
    End If
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub ListBox3_Click()

End Sub

Private Sub MultiPage1_Change()
    
    TModuleControl.InitializeDlg2 Me
        
End Sub

Private Sub OK1_Click()
        
    Dim choice(3) As Variant                            '�ѱ� ������ �븳����, �ŷڱ���,������� 3���ϱ�
    Dim dataRange As Range
    Dim i As Integer
    Dim activePt As Long                                '��� �м��� ���۵Ǵ� �κ��� �����ֱ� ����
    
    '''
    '''������ �������� �ʾ��� ���
    '''
    If Me.ListBox4.ListCount + Me.ListBox4.ListCount <> 2 Then
        MsgBox "2���� ������ ������ �ֽñ� �ٶ��ϴ�.", vbExclamation, "HIST"
        Exit Sub
    End If
    
    '''
    '''public ���� ���� xlist2, DataSheet, RstSheet, m, k2, n2
    '''
    ReDim xlist2(2)
    xlist2(1) = Me.ListBox4.List(0)
    
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
            If Me.ListBox4.List(0) = ActiveSheet.Cells(1, j) Then
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
    'If Me.OB7.Value = True Then choice(1) = 1
    'If Me.OB8.Value = True Then choice(1) = 2
    
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
    If Me.ChB2.Value = True Then
        If IsNumeric(Me.TextBox4.Value) = False Then
            MsgBox "����� �ŷڱ����� �Է��� �ֽñ� �ٶ��ϴ�.", vbExclamation, "HIST"
            Exit Sub
        ElseIf Me.TextBox4.Value < 0 Or Me.TextBox4.Value > 100 Then
            MsgBox "����� �ŷڱ����� %������ �Է��� �ֽñ� �ٶ��ϴ�.", vbExclamation, "HIST"
            Exit Sub
        End If
    End If
    
    '''
    '''�ŷڱ��� �Է� - choice(2)
    '''
    If Me.ChB2.Value = True Then choice(2) = Me.TextBox4.Value
    If Me.ChB2.Value = False Then choice(2) = -1
    
    '''
    '''�͹����� ���ð�� �Է� - choice(3)
    '''
    If Me.OB4 = True Then choice(3) = 1
    If Me.OB5 = True Then choice(3) = 2
    If Me.OB6 = True Then choice(3) = 3
    
    '''
    '''��� ó��
    '''
    TModuleControl.SettingStatusBar True, "��ǥ�� t-�������Դϴ�."
    Application.ScreenUpdating = False
    TModulePrint.makeOutputSheet (RstSheet)
    'Worksheets(RstSheet).Unprotect "prophet"
    activePt = Worksheets(RstSheet).Cells(1, 1).Value
    
    TModuleControl.TTest2 choice, 1
    
    
    'Worksheets(RstSheet).Protect Password:="prophet", DrawingObjects:=False, _
    '                                contents:=True, Scenarios:=True
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

Private Sub OK2_Click()
    Dim choice(3) As Variant                            '�ѱ� ������ �븳����, �ŷڱ���,������� 3���ϱ�
    Dim dataRange As Range
    Dim i As Integer
    Dim activePt As Long                                '��� �м��� ���۵Ǵ� �κ��� �����ֱ� ����
    
    '''
    '''������ �������� �ʾ��� ���
    '''
    If Me.ListBox4.ListCount = 0 Then
        MsgBox "�з������� ������ �ֽñ� �ٶ��ϴ�.", vbExclamation, "HIST"
        Exit Sub
    End If
    If Me.ListBox5.ListCount = 0 Then
        MsgBox "�м������� ������ �ֽñ� �ٶ��ϴ�.", vbExclamation, "HIST"
        Exit Sub
    End If
    
    '''
    '''public ���� ���� xlist2, DataSheet, RstSheet, m, k2, n2
    '''
    ReDim xlist2(2)
    xlist2(1) = Me.ListBox4.List(0)
    xlist2(2) = Me.ListBox5.List(0)

    DataSheet = ActiveSheet.Name                        'DataSheet : Data�� �ִ� Sheet �̸�
    RstSheet = "_���м����_"                       'RstSheet  : ����� �����ִ� Sheet �̸�
    Set dataRange = ActiveSheet.Cells.CurrentRegion
    m = dataRange.Columns.Count                         'm         : dataSheet�� �ִ� ���� ����
    
    tmp1 = 2
    ReDim xlist2(tmp1)                                  '�����̸���
    ReDim k2(tmp1)                                      '���°���� ��������
    ReDim n2(tmp1)                                      '����Ÿ �������
    ReDim tmp(tmp1)
    
    
    
    For i = 1 To tmp1
        tmp(i) = 0
        If i = 1 Then tmpList = Me.ListBox4.List(0)
        If i = 2 Then tmpList = Me.ListBox5.List(0)
        For j = 1 To m
            If tmpList = ActiveSheet.Cells(1, j) Then                   ' j ��
                xlist2(i) = ActiveSheet.Cells(1, j)                     '   �����̸����� , ����
                k2(i) = j
                n2(i) = ActiveSheet.Cells(1, j).End(xlDown).row - 1
                tmp(i) = tmp(i) + 1
            End If
    Next j
    Next i
    
        
                    ''''
                    '''' �ӽ÷� �迭�� 1���� �з����� 2���� �м�����
                    ''''

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
    '''�м� ������ ���ڿ� ���ڰ� ȥ�յǾ� ���� ���
    '''
    'For i = 1 To tmp1
        If TModuleControl.FindingRangeError(Me.ListBox5.List(0)) = True Then
            MsgBox "������ �м������� ���ڳ� ������ �ֽ��ϴ�." & Chr(10) & _
                    ": " & xlist2(i), vbExclamation, "HIST"
            Exit Sub
        End If
    'Next i
        
        
    '''
    ''' ����Է��̹Ƿ� �з�����, �м��������� �����ϱ�
    '''
        
    tmpList = 0
    tmp(1) = ActiveSheet.Cells(2, k2(1)).Value
    For i = 2 To n2(1)
        If tmpList = 1 And tmp(1) <> ActiveSheet.Cells(i + 1, k2(1)).Value And tmp(2) <> ActiveSheet.Cells(i + 1, k2(1)).Value Then
            MsgBox "�з������� 2������ ������ ������ �־�� �մϴ�." & Chr(10), vbExclamation, "HIST"
            Exit Sub
        End If
        If tmpList = 0 And tmp(1) <> ActiveSheet.Cells(i + 1, k2(1)).Value Then
            tmp(2) = ActiveSheet.Cells(i + 1, k2(1)).Value
            tmpList = tmpList + 1
        End If
        
    Next i
    
    If n2(1) <> n2(2) Then
        MsgBox "�з������� �м������� ����Ÿ ���� ���ƾ� �մϴ�." & Chr(10), vbExclamation, "HIST"
        Exit Sub
    End If
    
    tmpList = n2(1)
    n2(1) = 0: n2(2) = 0
    For i = 1 To tmpList
        If ActiveSheet.Cells(i + 1, k2(1)) = tmp(1) Then n2(1) = n2(1) + 1
        If ActiveSheet.Cells(i + 1, k2(1)) = tmp(2) Then n2(2) = n2(2) + 1
    Next i



    '''
    '''�������� ���ð�� �Է� - choice(1), choice(2)
    '''
    
    choice(1) = 1
    '''If Me.OB9.Value = True Then choice(1) = 1
    '''If Me.OB10.Value = True Then choice(1) = 2
    
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
    If Me.ChB2.Value = True Then
        If IsNumeric(Me.TextBox4.Value) = False Then
            MsgBox "����� �ŷڱ����� �Է��� �ֽñ� �ٶ��ϴ�.", vbExclamation, "HIST"
            Exit Sub
        ElseIf Me.TextBox4.Value < 0 Or Me.TextBox4.Value > 100 Then
            MsgBox "����� �ŷڱ����� %������ �Է��� �ֽñ� �ٶ��ϴ�.", vbExclamation, "HIST"
            Exit Sub
        End If
    End If
    
    '''
    '''�ŷڱ��� �Է� - choice(2)
    '''
    If Me.ChB2.Value = True Then choice(2) = Me.TextBox4.Value
    If Me.ChB2.Value = False Then choice(2) = -1
    
    '''
    '''�͹����� ���ð�� �Է� - choice(3)
    '''
    If Me.OB4 = True Then choice(3) = 1
    If Me.OB5 = True Then choice(3) = 2
    If Me.OB6 = True Then choice(3) = 3
    
    '''
    '''��� ó��
    '''
    TModuleControl.SettingStatusBar True, "��ǥ�� t-�������Դϴ�."
    Application.ScreenUpdating = False
    TModulePrint.makeOutputSheet (RstSheet)
    'Worksheets(RstSheet).Unprotect "prophet"
    activePt = Worksheets(RstSheet).Cells(1, 1).Value
    
    TModuleControl.TTest2 choice, 2
    
    
    'Worksheets(RstSheet).Protect Password:="prophet", DrawingObjects:=False, _
    '                                contents:=True, Scenarios:=True
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
   
   Me.ListBox3.Clear
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
   
   
   
   Me.ListBox3.List() = myArray
   
 '  For i = 1 To TempSheet.UsedRange.Columns.Count
 '   rngFirst.Offset(i, 1) = myArray(i - 1)
 ' Next i
  
End Sub

Private Sub UserForm_Click()

End Sub
