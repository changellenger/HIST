VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameOnepZ 
   OleObjectBlob   =   "frameOnepZ.frx":0000
   Caption         =   "�����ϱ� : p�� ���� z-����"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8790
   StartUpPosition =   1  '������ ���
   TypeInfoVer     =   50
End
Attribute VB_Name = "frameOnepZ"
Attribute VB_Base = "0{4E12898B-740A-4391-B3FC-752A25DCDC30}{A624648B-5432-4631-8591-4B681B3BB7CC}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False


Private Sub BtnOK_Click()
    
    Dim WarningMsg As String
    Dim choice(3) As Variant
    
    If IsNumeric(trial.Value) = False Then
       MsgBox ("����Ƚ���� �ùٸ��� �ʽ��ϴ�.")
       Exit Sub
    ElseIf IsNumeric(success.Value) = False Then
       MsgBox ("����Ƚ���� �ùٸ��� �ʽ��ϴ�.")
       Exit Sub
    ElseIf IsNumeric(ratio.Value) = False Then
       MsgBox ("���������� �ùٸ��� �ʽ��ϴ�.")
       Exit Sub
    ElseIf IsNumeric(level.Value) = False Then
       MsgBox ("�ŷڼ����� �ùٸ��� �ʽ��ϴ�.")
       Exit Sub
    End If
    t = CDbl(trial.Value)
    s = CDbl(success.Value)
    r = CDbl(ratio.Value)
    L = CDbl(level.Value)
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
    If t <= s Or s = 0 Then
       MsgBox ("�л��� 0�� ����Դϴ�.")
       Exit Sub
    End If
    If r >= 1 Or r <= 0 Then
       MsgBox ("������� 1���� ũ�ų� 0���� �۽��ϴ�.")
       Exit Sub
    End If
    PHat = s / t
    lim1 = t * PHat
    lim2 = t * (1 - PHat)
    
    If lim1 < 5 Or lim2 < 5 Then
       WarningMsg = "*����: ǥ���� ũ�Ⱑ �۽��ϴ�."
    End If
    zstat = (PHat - r) / Sqr((r - r ^ 2) / t)
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
    

    
    'resultsheet.Unprotect "prophet"
  
  ratiotest.ratioresult zstat, PHat, r, t, s, L, WarningMsg, resultsheet, choice
  

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

Private Sub Label7_Click()

End Sub

Private Sub level_Change()

End Sub

Private Sub ratio_Change()

End Sub

Private Sub success_Change()

End Sub

Private Sub UserForm_Click()

End Sub
