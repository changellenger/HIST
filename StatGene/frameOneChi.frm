VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameOneChi 
   OleObjectBlob   =   "frameOneChi.frx":0000
   Caption         =   "�����ϱ� : ����� ���� �֩�����"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8535
   StartUpPosition =   1  '������ ���
   TypeInfoVer     =   46
End
Attribute VB_Name = "frameOneChi"
Attribute VB_Base = "0{609917EC-EB46-40BC-80D0-87481267F7E5}{A7A280AA-2CC2-4E3F-99C2-14F8F7D8D5E8}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private Sub Cancel_Click()
    Unload Me
End Sub

Private Sub Label11_Click()

End Sub

Private Sub Label15_Click()
    frameDEx2.Show
    
End Sub

Private Sub Label7_Click()

End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub ListBox2_Click()

End Sub

Private Sub OB2_Click()

End Sub

Private Sub OB3_Click()

End Sub

Private Sub OK_Click()
    
    Dim WarningMsg As String
    Dim choice(3) As Variant
    
    If IsNumeric(TextBox3.Value) = False Then
       MsgBox ("ǥ���� ������ �ùٸ��� �ʽ��ϴ�.")
       Exit Sub
    ElseIf IsNumeric(TextBox4.Value) = False Then
       MsgBox ("ǥ���� �л��̿ùٸ��� �ʽ��ϴ�.")
       Exit Sub
    ElseIf IsNumeric(TextBox1.Value) = False Then
       MsgBox ("��л��� ���� �ùٸ��� �ʽ��ϴ�.")
       Exit Sub
    ElseIf IsNumeric(TextBox2.Value) = False Then
       MsgBox ("�ŷڼ����� �ùٸ��� �ʽ��ϴ�.")
       Exit Sub
    End If
    t = CDbl(TextBox3.Value)
    s = CDbl(TextBox4.Value)
    r = CDbl(TextBox1.Value)
    L = CDbl(TextBox2.Value)
    If Int(t) - t <> 0 Or t <= 0 Then
       MsgBox ("�ڿ����� �Է��ϼ���.")
       Exit Sub
    End If
    If Int(s) - s <> 0 Or s < 0 Then
       MsgBox ("�ڿ����� �Է��ϼ���.")
       Exit Sub
    End If
    If L <= 0 Or L >= 100 Then
       MsgBox ("�ŷڼ����� �������� ������ϴ�.")
       Exit Sub
    End If
'''����Ȯ�κ�Ź
  '  If t <= s Or s = 0 Then
  '     MsgBox ("�л��� 0�� ����Դϴ�1.")
  '     Exit Sub
   ' End If
    'If r <= s Or s = 0 Then
    '   MsgBox ("�л��� 0�� ����Դϴ�2.")
    '   Exit Sub
   ' End If
 
    'PHat = s / t
    'lim1 = t * PHat
    'lim2 = t * (1 - PHat)
    
    'If lim1 < 5 Or lim2 < 5 Then
    '   WarningMsg = "*����: ǥ���� ũ�Ⱑ �۽��ϴ�."
   ' End If
    zstat = (((t - 1) * s) / r)
    Set resultsheet = OpenOutSheet2("_���м����_", True)
    '''
    '''�͹����� ���ð�� �Է� - choice(3)
    '''
    If Me.OB1 = True Then choice(3) = 1
    If Me.OB2 = True Then choice(3) = 2
    If Me.OB3 = True Then choice(3) = 3
    
    '''
    '''
    '''
    '''
    '''
    '''
    RstSheet = "_���м����_"
    
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


    
    'Worksheets(RstSheet).Unprotect "prophet"
    activePt = Worksheets(RstSheet).Cells(1, 1).Value
    
    '''Worksheets(RstSheet).Cells(1, 1).Value = "$A$" & Worksheets(RstSheet).Cells(1, 1).Value
 
    
    
    'resultsheet.Unprotect "prophet"

    ratiotest.resultOneChi zstat, PHat, r, t, s, L, WarningMsg, resultsheet, choice
    'resultsheet.Protect "prophet"
    
    '''Worksheets(RstSheet).Cells(1, 1).Value = Right(Worksheets(RstSheet).Cells(1, 1).Value, Len(Worksheets(RstSheet).Cells(1, 1).Value) - 3)
    
    



    '���� ���� üũ �� �񱳰� ����
    Dim Cmp_Value As Long
    
    If PublicModule.ChkVersion1(ActiveWorkbook.Name) = True Then
        Cmp_Value = 1048000
    Else
        Cmp_Value = 65000
    End If
    
    If Worksheets(RstSheet).Cells(1, 1).Value > Cmp_Value Then
        MsgBox "[_���м����_]��Ʈ�� ���� ��� ����Ͽ����ϴ�." & vbCrLf & "�� ��Ʈ�� �̸��� �ٲٰų� ������ �ּ���", vbExclamation, "HIST"
        Exit Sub
    End If
    
    Worksheets(RstSheet).Activate
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

MsgBox ("���α׷��� ������ �ֽ��ϴ�.")
 'End sub �տ��� ���δ�.

''�ؼ�, ������ ���� Err_delete�� �ͼ� ù�����ķ� �����. ���� ù���� 2�� ��Ʈ�� �����.�׸��� �����޽��� ���
'rSTsheet����⵵ ���� �������� ��쿡�� �ƹ� ���۵� ���� �ʰ�, �����޽����� ����.
End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub TextBox2_Change()

End Sub

Private Sub UserForm_Click()

End Sub
