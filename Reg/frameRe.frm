VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameRe 
   OleObjectBlob   =   "frameRe.frx":0000
   Caption         =   "ȸ�ͺм�"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5370
   StartUpPosition =   1  '������ ���
   TypeInfoVer     =   47
End
Attribute VB_Name = "frameRe"
Attribute VB_Base = "0{2E49CCFA-88BE-4882-AAA7-71E9D5213812}{4418A814-02ED-48CC-9A83-9544F6CF3D6C}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

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



End Sub
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
               Me.CB1.Visible = False
               Me.CB2.Visible = True
               Exit Sub
            End If
            i = i + 1
        Loop
    End If
End Sub

Private Sub CB2_Click()
    If Me.ListBox2.ListCount <> 0 Then
        Me.ListBox1.AddItem ListBox2.list(0)
        Me.ListBox2.RemoveItem (0)
        Me.CB1.Visible = True
        Me.CB2.Visible = False
    End If
End Sub

Private Sub CB3_Click()
    Dim i As Integer
    i = 0
         Do While i <= Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(i) = True Then
               Me.ListBox3.AddItem Me.ListBox1.list(i)
               Me.ListBox1.RemoveItem (i)
               Exit Sub
            End If
            i = i + 1
        Loop
 
End Sub

Private Sub CB4_Click()
    If Me.ListBox3.ListCount <> 0 Then
        Me.ListBox1.AddItem ListBox3.list(0)
        Me.ListBox3.RemoveItem (0)
   
    End If
End Sub


Private Sub CheckBox2_Click()
   ' If CheckBox2.Value = True Then
  '      TextBox1.Enabled = True
  '  Else
   '     TextBox1.Enabled = False
  '  End If
End Sub

Private Sub HelpBtn_Click()
ShellExecute 0, "open", "hh.exe", ThisWorkBook.Path + "\HIST%202013.chm::/ȸ�ͺм�.htm", "", 1
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
    Else
        Do While i <= Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(i) = True Then
               Me.ListBox3.AddItem Me.ListBox1.list(i)
               Me.ListBox1.RemoveItem (i)
               Exit Do
            End If
            i = i + 1
        Loop
    End If
    
    If Me.ListBox3.ListCount = 1 Then
        Me.Frame2.Enabled = True
        Me.CheckBox3.Enabled = True
        Me.CheckBox4.Enabled = True
        Me.CheckBox5.Enabled = True
        Me.Label5.Enabled = True
    Else
        Me.Frame2.Enabled = False
        Me.CheckBox3.Enabled = False
        Me.CheckBox4.Enabled = False
        Me.CheckBox5.Enabled = False
        Me.Label5.Enabled = False
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
    Application.Run "HIST.xlam!publicmodule.MoveBtwnListBox", _
        Me, "ListBox3", "ListBox1"
    If Me.ListBox3.ListCount = 1 Then
        Me.Frame2.Enabled = True
        Me.CheckBox3.Enabled = True
        Me.CheckBox4.Enabled = True
        Me.CheckBox5.Enabled = True
        Me.Label5.Enabled = True
    Else
        Me.Frame2.Enabled = False
        Me.CheckBox3.Enabled = False
        Me.CheckBox4.Enabled = False
        Me.CheckBox5.Enabled = False
        Me.Label5.Enabled = False
    End If
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
    Dim Sign As Boolean, errSign1 As Boolean, ErrSign2 As Boolean
    Dim errString As String
    Dim activePt As Long            '��� �м��� ���۵Ǵ� �κ��� �����ֱ� ����
    
    Dim WS As Worksheet
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
    n = dataRange.Cells(1, 1).End(xlDown).row - 1       'Data����
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
 '   For i = 0 To 3
 '       If frmVarSel.Controls("OptionButton" & (i + 1)).value = True Then
 '       method = i + 1
 '       End If
 '   Next i
    
'    If frmVarSel.Controls("Label1").Enabled = True Then
'        addlevel = frmVarSel.Controls("TextBox1").value / 100
'    Else: addlevel = -1
'    End If
    
'    If frmVarSel.Controls("Label2").Enabled = True Then
'        rmlevel = frmVarSel.Controls("TextBox2").value / 100
'    Else: rmlevel = -1
'    End If
          
'    For i = 1 To 3
'        criteria(i - 1) = 0
'        If frmVarSel.Controls("CheckBox" & i).value = True Then
'        criteria(i - 1) = 1
'        End If
'    Next i
                                                       
    'intercept : ����� ���� ���� : Boolean
    'alpha : �������� �ŷڱ��� , '''�Ҽ��ý� -1 : double
    'ScatterPlot, PIgraph : Boolean '''�̸��ٲ� simple()����
    'method : ���� ���� ��� 1~4, �Ҽ��ý� -1 : integer
    'addlevel : �߰� ����(%), �Ҽ��ý� -1  : double
    'rmlevel : ���� ����(%), �Ҽ��ý� -1   : double
    'criteria : (��簡����) ���� ���� ���� 5~7, �Ҽ��ý� -1 : integer
    
    
    '''
    '''
    'ȸ������ �Է����� �����ϱ�
        
  '  For i = 1 To 18
  '      resi(i) = frmResid.Controls("CheckBox" & i).value
  '  Next i
    
    '������跮
    'resi(1)    : ���
    'resi(2)    : vs �������� �׷���
    'resi(3)    : vs ����ġ �׷���
    'resi(4)    : ������׷�
    'resi(5)    : ����Ȯ���׸�
    
    'resi(1)~(5) ����
    'resi(6)~(10) ǥ��ȭ ����
    'resi(11)~(15) ǥ��ȭ ���� ����
    
    'resi(16)   :���������
    'resi(17)   :���߰�����
    'resi(18)   :�κ�ȸ�ͻ�����
    
    '���ο����� ��� �ӽý�Ʈ �����ϱ�
    
    check1 = 0
    check2 = 0
    For Each WS In Worksheets
        If WS.Name = RstSheet Then check1 = 1
        If WS.Name = "_#TmpHIST1#_" Then check2 = 1
    Next WS
    
    Application.DisplayAlerts = False
    If check1 = 0 And check2 = 1 Then Worksheets("_#TmpHIST1#_").Delete
    Application.DisplayAlerts = True

    '''
    '''�������� �������� ����
    '''
    If n <> ModuleControl.FindVarCount(ylist) Then errSign1 = True
    For i = 0 To p - 1
        If n <> ModuleControl.FindVarCount(xlist(i)) Then errSign1 = True
    Next i
    '''
    '''���ڿ� ���ڰ� ȥ�յǾ� ���� ���
    '''
    If ModuleControl.FindingRangeError(ylist) = True Then
        ErrSign2 = True: errString = Me.ListBox2.list(0)
    End If
    
    For i = 0 To p - 1
        If ModuleControl.FindingRangeError(xlist(i)) = True Then
            ErrSign2 = True
            If errString <> "" Then
                errString = errString & "," & xlist(i)
            Else: errString = xlist(i)
            End If
        End If
    Next i
    '''
    '''������ ���� ��� ���� �޽��� ���
    '''
    If errSign1 = True Then
        MsgBox "�������� �������� �ٸ��ϴ�.", _
                vbExclamation, "HIST"
        Exit Sub
    End If
    If ErrSign2 = True Then
        MsgBox "������ �м������� ���ڳ� ������ �ֽ��ϴ�." & Chr(10) & _
               ": " & errString, vbExclamation, "HIST"
        Exit Sub
    End If
                                                           
    '''
    '''������ ó���ϴ� �κ�
    '''
    ModuleControl.SettingStatusBar True, "ȸ�� �м����Դϴ�."
    Application.ScreenUpdating = False
    
    ModulePrint.MakeOutputSheet RstSheet
    'Worksheets(RstSheet).Unprotect "prophet"
    activePt = Worksheets(RstSheet).Range("a1").value
    
   
    ModuleControl.Reg intercept
    
    If p > 1 Then ModuleControl.VarSel method, addlevel, rmlevel, criteria, intercept, resi, ci, Alpha, simple
    
    If method <= 0 Or p = 1 Then ModuleResi.Diagnosis00 resi, intercept, ci, Alpha, simple
    
    'Worksheets(RstSheet).Protect password:="prophet", DrawingObjects:=False, _
    '                                contents:=True, Scenarios:=True
    ModuleControl.SettingStatusBar False
    Application.ScreenUpdating = True
   
    Unload Me
    
    '��� �м��� ���۵Ǵ� �κп��� ���� �Ʒ� ���� �����ָ� ��ģ��.
    Worksheets(RstSheet).Activate
    
    '���� ���� üũ �� �񱳰� ����
    Dim Cmp_Value As Long
    
    If ModuleControl.ChkVersion(ActiveWorkbook.Name) = True Then
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
