VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameReMulti 
   OleObjectBlob   =   "frameReMulti.frx":0000
   Caption         =   "�����ϱ� : ���߼���ȸ�ͺм�"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9105
   StartUpPosition =   1  '������ ���
   TypeInfoVer     =   143
End
Attribute VB_Name = "frameReMulti"
Attribute VB_Base = "0{6B82FC11-AFE7-46C1-A2CB-B03AE4CE3E5A}{EC1A1963-7EC8-44BE-81B7-318BF52A4660}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private Sub Cancel_Click()
    Unload Me
End Sub
Private Sub CB1_Click()

    Dim i As Integer
    i = 0
    If Me.ListBox2.ListCount = 0 Then
        Do While i <= Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(i) = True Then
               Me.ListBox2.AddItem Me.ListBox1.list(i)
               Me.ListBox1.RemoveItem (i)

               Exit Sub
            End If
            i = i + 1
        Loop
    End If
    
    
    Dim j As Integer
    j = 0
    If Me.ListBox3.ListCount = 0 Then
        Do While j <= Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(j) = True Then
               Me.ListBox3.AddItem Me.ListBox1.list(j)
               Me.ListBox1.RemoveItem (j)

               Exit Sub
            End If
            j = j + 1
        Loop
    End If
    
    
End Sub

Private Sub CB4_Click()

    Dim i As Integer
    Dim FromLNum, ToLNum
    
    i = 0
    FromLNum = "ListBox3": ToLNum = "ListBox1"
    Do While i <= Me.Controls(FromLNum).ListCount - 1
        If Me.Controls(FromLNum).Selected(i) = True Then
           Me.Controls(ToLNum).AddItem Me.Controls(FromLNum).list(i)
           Me.Controls(FromLNum).RemoveItem i
           Me.Controls(FromLNum).Selected(i) = False
          'i = i - 1
        End If
        i = i + 1
    Loop

End Sub


Private Sub CB3_Click()

    Dim i As Integer
    Dim FromLNum, ToLNum
    
    i = 0
    FromLNum = "ListBox1": ToLNum = "ListBox3"
    Do While i <= Me.Controls(FromLNum).ListCount - 1
        If Me.Controls(FromLNum).Selected(i) = True Then
           Me.Controls(ToLNum).AddItem Me.Controls(FromLNum).list(i)
           Me.Controls(FromLNum).RemoveItem i
           Me.Controls(FromLNum).Selected(i) = False
          ' i = i - 1
        End If
        i = i + 1
    Loop
End Sub


Private Sub Label15_Click()
    frameReEx2.Show
    
End Sub
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim i As Integer
    i = 0
    If Me.ListBox2.ListCount = 0 Then
        Do While i <= Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(i) = True Then
               Me.ListBox2.AddItem Me.ListBox1.list(i)
               Me.ListBox1.RemoveItem (i)
               Me.CB1.Visible = False
               Me.CB2.Visible = True
               Exit Sub
            End If
            i = i + 1
        Loop
    End If


   ElseIf Me.ListBox3.ListCount = 0 Then
        Do While j <= Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(j) = True Then
               Me.ListBox3.AddItem Me.ListBox1.list(j)
               Me.ListBox1.RemoveItem (j)
               Me.CB1.Visible = False
               Me.CB2.Visible = True
               Exit Sub
            End If
            j = j + 1
        Loop
    End If
    
End Sub
Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.ListBox2.ListCount <> 0 Then
        Me.ListBox1.AddItem Me.ListBox2.list(0)
        Me.ListBox2.RemoveItem (0)
        Me.CB1.Visible = True
        Me.CB2.Visible = False
    End If
End Sub
Private Sub ListBox3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.ListBox3.ListCount <> 0 Then
        Me.ListBox1.AddItem Me.ListBox3.list(0)
        Me.ListBox3.RemoveItem (0)
        Me.CB1.Visible = True
        Me.CB2.Visible = False
    End If
End Sub

Private Sub OK_Click()

    Dim choice(3) As Variant                            '�ѱ� ������ ������,�ŷڱ���,�븳���� 3���ϱ�
    Dim dataRange As Range
    Dim i As Integer
    Dim activePt As Long                                '��� �м��� ���۵Ǵ� �κ��� �����ֱ� ����
    
    '''
    '''������ �������� �ʾ��� ���
    '''
    If Me.ListBox2.ListCount = 0 Then
        MsgBox "������ ������ �ֽñ� �ٶ��ϴ�.", vbExclamation, "HIST"
        Exit Sub
    End If
     Me.ChB1.value = True
    '''
    '''public ���� ���� xlist, DataSheet, RstSheet, m, k1, n
    '''
    xlist = Me.ListBox2.list(0)
    DataSheet = ActiveSheet.Name                        'DataSheet : Data�� �ִ� Sheet �̸�
    RstSheet = "_���м����_"                       'RstSheet  : ����� �����ִ� Sheet �̸�
    
    
    
    '������ �Է�
On Error GoTo Err_delete
Dim val3535 As Long '�ʱ���ġ ������ ����'
Dim s3535 As Worksheet
val3535 = 2
    For Each s3535 In ActiveWorkbook.Sheets
        If s3535.Name = RstSheet Then
val3535 = Sheets(RstSheet).Cells(1, 1).value
End If
Next s3535  '��Ʈ�� �̹������� ��� ��ġ �������ϰ�, ������ 2�� �����Ѵ�.


    
    Set dataRange = ActiveSheet.Cells.CurrentRegion
    m = dataRange.Columns.count                         'm         : dataSheet�� �ִ� ���� ����
    
    tmp = 0
    For i = 1 To m
        If xlist = ActiveSheet.Cells(1, i) Then
            k1 = i  'k1                                 : k1 : ���õ� ������ ���° ���� �ִ���
            tmp = tmp + 1
        End If
    Next i
    N = ActiveSheet.Cells(1, k1).End(xlDown).row - 1    'n         : ���õ� ������ ����Ÿ ����

    '''
    ''' �������� ���� ��� - ������ ���� �ִ� ������ �ԷµǹǷ� ����ó���Ѵ�.
    '''
    If tmp > 1 Then
        MsgBox xlist & "�� ���� �������� �ֽ��ϴ�. " & vbCrLf & "�������� �ٲ��ֽñ� �ٶ��ϴ�.", vbExclamation, "HIST"
        Exit Sub
    End If
    
    '''
    '''���ڿ� ���ڰ� ȥ�յǾ� ���� ���
    '''
    If TModuleControl.FindingRangeError(xlist) = True Then
        MsgBox "������ �м������� ���ڳ� ������ �ֽ��ϴ�." & Chr(10) & _
               ": " & xlist, vbExclamation, "HIST"
        Exit Sub
    End If
    
    '''
    '''�������� �Է����� ���� ���
    '''
    If IsNumeric(Me.TextBox1.value) = False Then
        MsgBox "����� �������� �Է��� �ֽñ� �ٶ��ϴ�.", vbExclamation, "HIST"
        Exit Sub
    End If
    
    '''
    '''�ŷڱ����� �߸� �Է��� ���
    '''
  '  If Me.ChB1.Value = True Then
        If IsNumeric(Me.TextBox2.value) = False Then
            MsgBox "����� �ŷڱ����� �Է��� �ֽñ� �ٶ��ϴ�.", vbExclamation, "HIST"
            Exit Sub
        ElseIf Me.TextBox2.value < 0 Or Me.TextBox2.value > 100 Then
            MsgBox "����� �ŷڱ����� %������ �Է��� �ֽñ� �ٶ��ϴ�.", vbExclamation, "HIST"
            Exit Sub
        End If
  '  End If
    
    '''
    ''' ����Ÿ ������ �Ѱ��� ���
    '''
    If N = 1 Then
        MsgBox "�� ���� ����Ÿ�� ������ ������ �� �����ϴ�.", vbExclamation, "HIST"
        Exit Sub
    End If
    
    '''
    '''������ ���ð�� �Է� - choice(1)
    '''
    choice(1) = Me.TextBox1.value
    
    '''
    '''�ŷڱ��� �Է� - choice(2)
    '''
    If Me.ChB1.value = True Then choice(2) = Me.TextBox2.value
    If Me.ChB1.value = False Then choice(2) = -1
    
    '''
    '''�͹����� ���ð�� �Է� - choice(3)
    '''
    If Me.OB1 = True Then choice(3) = 1
    If Me.OB2 = True Then choice(3) = 2
    If Me.OB3 = True Then choice(3) = 3
    
    
    '''
    '''��� ó��
    '''
    TModuleControl.SettingStatusBar True, "��ǥ�� t-�������Դϴ�."
    Application.ScreenUpdating = False
    TModulePrint.MakeOutputSheet (RstSheet)
    'Worksheets(RstSheet).Unprotect "prophet"
    activePt = Worksheets(RstSheet).Cells(1, 1).value
    
    TModuleControl.TTestR choice
    
    
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
    
    If Worksheets(RstSheet).Cells(1, 1).value > Cmp_Value Then
        MsgBox "[_���м����_]��Ʈ�� ���� ��� ����Ͽ����ϴ�." & vbCrLf & "�� ��Ʈ�� �̸��� �ٲٰų� ������ �ּ���", vbExclamation, "HIST"
        Exit Sub
    End If
    
    Worksheets(RstSheet).Cells(activePt + 10, 1).Select
    Worksheets(RstSheet).Cells(activePt + 10, 1).Activate
                            '��� �м��� ���۵Ǵ� �κ��� �����ָ� ��ģ��.
                                
        
    Unload Me



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

MsgBox ("���α׷��� ������ �ֽ��ϴ� .")
 'End sub �տ��� ���δ�.

''�ؼ�, ������ ���� Err_delete�� �ͼ� ù�����ķ� �����. ���� ù���� 2�� ��Ʈ�� �����.�׸��� �����޽��� ���
'rSTsheet����⵵ ���� �������� ��쿡�� �ƹ� ���۵� ���� �ʰ�, �����޽����� ����.
        
End Sub
Private Sub OK1_Click()                                                '''  "_ȸ�ͺм����_"

    Dim intercept As Boolean
    Dim ci As Boolean
    Dim Alpha As Single
    Dim ScatterPlot As Boolean, PIgraph As Boolean
    Dim resi(18) As Boolean         'resi(0)��� '��' �Ұ���.
    Dim simple(3)                   'simple(0)��� '��' �Ұ���.
    Dim method As Integer
    Dim addlevel As Double, rmlevel As Double
    Dim criteria(2)
    Dim sign As Boolean, errsign1 As Boolean, errsign2 As Boolean
    Dim errString As String
    Dim activePt As Long            '��� �м��� ���۵Ǵ� �κ��� �����ֱ� ����
    
    Dim ws As Worksheet
    Dim check1 As Integer, check2 As Integer
    
    '''
    '''���� ó�� �κ� 1
    '''
    If Me.ListBox2.ListCount = 0 Or Me.ListBox3.ListCount = 0 Then
        MsgBox "���� ������ �������� �ʽ��ϴ�.", vbExclamation
        Exit Sub
    End If
    If IsNumeric(Me.TextBox1.value) = False Then
        MsgBox "�ŷ�Ȯ���� �ùٸ��� �ʽ��ϴ�.", vbExclamation
        Exit Sub
    Else
        If Me.TextBox1.value <= 0 Or Me.TextBox1.value >= 100 Then
            MsgBox "�ŷ�Ȯ���� �ùٸ��� �ʽ��ϴ�.", vbExclamation
            Exit Sub
        End If
    End If


    '''
    '''�Է¹��� ���� �����ϱ�
        
    '������� MdControl ���� ����� Public ����
    '���⼭ �ѹ��� �������ش�
    
    DataSheet = ActiveSheet.Name        'Data�� �ִ� Sheet �̸�
    RstSheet = "_���м����_"         '����� �����ִ� Sheet �̸�
    '����ϴ� �ش� ��⿡ �� ���� ����'
'������ �Է�
On Error GoTo Err_delete
Dim val3535 As Long '�ʱ���ġ ������ ����'
Dim s3535 As Worksheet
val3535 = 2
    For Each s3535 In ActiveWorkbook.Sheets
        If s3535.Name = RstSheet Then
val3535 = Sheets(RstSheet).Cells(1, 1).value
End If
Next s3535  '��Ʈ�� �̹������� ��� ��ġ �������ϰ�, ������ 2�� �����Ѵ�.

                                        
    '''
    ylist = Me.ListBox2.list(0)            '���õ� ���Ӻ����̸�
    p = Me.ListBox3.ListCount              '���õ� �������� ����
    
    ReDim xlist(p - 1)
    For i = 0 To p - 1
        xlist(i) = ListBox3.list(i)         '���õ� �������� �̸�
    Next i
    
    Set dataRange = ActiveSheet.Cells.CurrentRegion
    N = dataRange.Cells(1, 1).End(xlDown).row - 1       'Data����
    m = dataRange.Cells(1, 1).End(xlToRight).Column - 1 '�������� ����
    
    '������� MdControl ���� ����� Public ����
    '���⼭ �ѹ��� �������ش�. �ٸ� ������ �ٲ��� �ʴ´�
    'DataSheet, RstSheet, ylist, xlist, N, M, p
    
    
    
    '''
    '''
    '�������� �Է����� �����ϱ�
    intercept = CheckBox1.value
    ci = CheckBox2.value
    Alpha = TextBox1.value
    
    simple(1) = CheckBox3.value     '������
    simple(2) = CheckBox4.value     '�ŷڴ� �׷���
    simple(3) = CheckBox5.value     'vs�������� �׷���
    
    method = -1

    
    check1 = 0
    check2 = 0
    For Each ws In Worksheets
        If ws.Name = RstSheet Then check1 = 1
        If ws.Name = "_#TmpHIST1#_" Then check2 = 1
    Next ws
    
    Application.DisplayAlerts = False
    If check1 = 0 And check2 = 1 Then Worksheets("_#TmpHIST1#_").Delete
    Application.DisplayAlerts = True

    '''
    '''�������� �������� ����
    '''
    If N <> Modulecontrol.FindVarCount(ylist) Then errsign1 = True
    For i = 0 To p - 1
        If N <> Modulecontrol.FindVarCount(xlist(i)) Then errsign1 = True
    Next i
    '''
    '''���ڿ� ���ڰ� ȥ�յǾ� ���� ���
    '''
    If Modulecontrol.FindingRangeError(ylist) = True Then
        errsign2 = True: errString = Me.ListBox2.list(0)
    End If
    
    For i = 0 To p - 1
        If Modulecontrol.FindingRangeError(xlist(i)) = True Then
            errsign2 = True
            If errString <> "" Then
                errString = errString & "," & xlist(i)
            Else: errString = xlist(i)
            End If
        End If
    Next i
    '''
    '''������ ���� ��� ���� �޽��� ���
    '''
    If errsign1 = True Then
        MsgBox "�������� �������� �ٸ��ϴ�.", _
                vbExclamation, "HIST"
        Exit Sub
    End If
    If errsign2 = True Then
        MsgBox "������ �м������� ���ڳ� ������ �ֽ��ϴ�." & Chr(10) & _
               ": " & errString, vbExclamation, "HIST"
        Exit Sub
    End If
                                                           
    '''
    '''������ ó���ϴ� �κ�
    '''
    Modulecontrol.SettingStatusBar True, "ȸ�� �м����Դϴ�."
    Application.ScreenUpdating = False
    
    ModulePrint.MakeOutputSheet RstSheet
    'Worksheets(RstSheet).Unprotect "prophet"
    activePt = Worksheets(RstSheet).Range("a1").value
    
   
    Modulecontrol.Reg intercept
    
    If p > 1 Then Modulecontrol.VarSel method, addlevel, rmlevel, criteria, intercept, resi, ci, Alpha, simple
    
    If method <= 0 Or p = 1 Then ModuleResi.Diagnosis00 resi, intercept, ci, Alpha, simple
   
    Modulecontrol.SettingStatusBar False
    Application.ScreenUpdating = True
   
    Unload Me
    
    '��� �м��� ���۵Ǵ� �κп��� ���� �Ʒ� ���� �����ָ� ��ģ��.
    Worksheets(RstSheet).Activate
    
    '���� ���� üũ �� �񱳰� ����
    Dim Cmp_Value As Long
    
    If Modulecontrol.ChkVersion(ActiveWorkbook.Name) = True Then
        Cmp_Value = 1048000
    Else
        Cmp_Value = 65000
    End If
    
    If Worksheets(RstSheet).Cells(1, 1).value > Cmp_Value Then
        MsgBox "[_���м����_]��Ʈ�� ���� ��� ����Ͽ����ϴ�." & vbCrLf & "�� ��Ʈ�� �̸��� �ٲٰų� ������ �ּ���", vbExclamation, "HIST"
        Exit Sub
    End If
    
    Worksheets(RstSheet).Range("a" & activePt + 10).Select
    Worksheets(RstSheet).Range("a" & activePt + 10).Activate
    
    Unload Me
    
    
Exit Sub
'�ǵڿ� ���̱�
Err_delete:

For Each s3535 In ActiveWorkbook.Sheets
        If s3535.Name = RstSheet Then
Sheets(RstSheet).Range(Cells(val3535, 1), Cells(10000, 10000)).Select
Selection.Delete
Sheets(RstSheet).Cells(1, 1) = val3535
Sheets(RstSheet).Cells(val3535, 1).Select

End If
Next s3535
If val3535 = 2 Then
Application.DisplayAlerts = False
Sheets(RstSheet).Delete
End If
MsgBox ("���α׷��� ������ �ֽ��ϴ�.")

 'End sub �տ��� ���δ�.

''�ؼ�, ������ ���� Err_delete�� �ͼ� �����. Rstsheet�� ������ �������. RSTsheet����⵵ ���� ''�������� ��.. ����� ���� �� ������.
End Sub



Private Sub OptionButton1_Click()

 
   Dim myRange As Range
   Dim myArray()
   Dim arrName As Variant
   Dim TempSheet As Worksheet
   Set TempSheet = ActiveCell.Worksheet
   
    ReDim arrName(TempSheet.UsedRange.Columns.count)
' Reading Data
    For i = 1 To TempSheet.UsedRange.Columns.count
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
    ReDim myArray(TempSheet.UsedRange.Columns.count - 1)
    a = 0
   For i = 1 To TempSheet.UsedRange.Columns.count
   If arrName(i) <> "" Then                     '��ĭ����
   myArray(a) = arrName(i)
   a = a + 1
   
   Else:
   End If
   Next i
   
   
   
   Me.ListBox1.list() = myArray
   
 '  For i = 1 To TempSheet.UsedRange.Columns.Count
 '   rngFirst.Offset(i, 1) = myArray(i - 1)
 ' Next i
  


    For j = 1 To TempSheet.UsedRange.Columns.count
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
    ReDim myArray(TempSheet.UsedRange.Columns.count - 1)
    a = 0
   For j = 1 To TempSheet.UsedRange.Columns.count
   If arrName(j) <> "" Then                     '��ĭ����
   myArray(a) = arrName(j)
   a = a + 1
   
   Else:
   End If
   Next j
   
   
   
   Me.ListBox1.list() = myArray
   
 '  For i = 1 To TempSheet.UsedRange.Columns.Count
 '   rngFirst.Offset(i, 1) = myArray(i - 1)
 ' Next i
  

End Sub

Private Sub UserForm_Click()

End Sub
