VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameReSimple1 
   OleObjectBlob   =   "frameReSimple1.frx":0000
   Caption         =   "�����ϱ� : �ܼ�����ȸ�ͺм�"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9480
   StartUpPosition =   1  '������ ���
   TypeInfoVer     =   136
End
Attribute VB_Name = "frameReSimple1"
Attribute VB_Base = "0{C511A6D1-9AD2-4A89-A4C7-385AFF09A851}{336531F5-9169-46C2-9D10-954B922C93FF}"
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
    If Me.ListBox3.ListCount = 0 Then
        Do While i <= Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(i) = True Then
               Me.ListBox3.AddItem Me.ListBox1.list(i)
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
    If Me.ListBox3.ListCount <> 0 Then
        Me.ListBox1.AddItem ListBox3.list(0)
        Me.ListBox3.RemoveItem (0)
        Me.CB3.Visible = True
        Me.CB4.Visible = False
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
ShellExecute 0, "open", "hh.exe", ThisWorkbook.Path + "\HIST%202013.chm::/ȸ�ͺм�.htm", "", 1
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