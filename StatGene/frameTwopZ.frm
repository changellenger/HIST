VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameTwopZ 
   OleObjectBlob   =   "frameTwopZ.frx":0000
   Caption         =   "�����ϱ� : p��-p���� ���� z-����"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10980
   StartUpPosition =   1  '������ ���
   TypeInfoVer     =   102
End
Attribute VB_Name = "frameTwopZ"
Attribute VB_Base = "0{27AA43E3-4EA8-4B17-B517-3E438C105E34}{36C3D5BB-ED95-406B-A1A7-809C0B289CEF}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False


Private Sub BtnOK_Click()                                       '''"_������������_"
 
    Dim choice(3) As Variant
    
    If IsNumeric(trial1.Value) = False Then
       MsgBox ("����1�� ����Ƚ���� �ùٸ��� �ʽ��ϴ�.")
       Exit Sub
    ElseIf IsNumeric(success1.Value) = False Then
       MsgBox ("����1�� ����Ƚ���� �ùٸ��� �ʽ��ϴ�.")
       Exit Sub
    ElseIf IsNumeric(trial2.Value) = False Then
       MsgBox ("����2�� ����Ƚ���� �ùٸ��� �ʽ��ϴ�.")
       Exit Sub
    ElseIf IsNumeric(success2.Value) = False Then
       MsgBox ("����2�� ����Ƚ���� �ùٸ��� �ʽ��ϴ�.")
       Exit Sub
     ElseIf IsNumeric(level.Value) = False Then
       MsgBox ("�ŷڼ����� �ùٸ��� �ʽ��ϴ�.")
       Exit Sub
    End If
    t1 = CDbl(trial1.Value)
    t2 = CDbl(trial2.Value)
    s1 = CDbl(success1.Value)
    s2 = CDbl(success2.Value)
    L = CDbl(level.Value)
    If Int(t1) - t1 <> 0 Or t1 <= 0 Then
       MsgBox ("�ڿ����� �Է��ϼ���")
       Exit Sub
    End If
    
    If Int(s1) - s1 <> 0 Or s1 < 0 Then
       MsgBox ("�ڿ����� �Է��ϼ���")
       Exit Sub
    End If
    If L <= 0 Or L >= 100 Then
       MsgBox ("�ŷڼ����� �������� ������ϴ�.")
       Exit Sub
    End If
    If t1 < s1 Then
       MsgBox ("�Է������� ���� �ʽ��ϴ�.")
       Exit Sub
    End If
    If Int(t2) - t2 <> 0 Or t2 <= 0 Then
       MsgBox ("�ڿ����� �Է��ϼ���")
       Exit Sub
    End If
    If Int(s2) - s2 <> 0 Or s2 < 0 Then
       MsgBox ("�ڿ����� �Է��ϼ���")
       Exit Sub
    End If
    
    If t2 < s2 Then
       MsgBox ("�Է������� ���� �ʽ��ϴ�.")
       Exit Sub
    End If
    
    If t1 = s1 And t2 = s2 Then
       MsgBox ("�л��� 0�� ����Դϴ�.")
       Exit Sub
    End If
    
    If s1 = 0 And s2 = 0 Then
       MsgBox ("�л��� 0�� ����Դϴ�.")
       Exit Sub
    End If
        
    PHat1 = s1 / t1
    PHat2 = s2 / t2
    PHat = (s1 + s2) / (t1 + t2)
    zstat = (PHat1 - PHat2) / Sqr((PHat - PHat ^ 2)) / Sqr((1 / trial1.Value + 1 / trial2.Value))
    Set resultsheet = OpenOutSheet2("_���м����_", True)
    
    '''
    '''
    
    '''
    '''�͹����� ���ð�� �Է� - choice(3)
    '''
    If Me.OB4 = True Then choice(3) = 1
    If Me.OB5 = True Then choice(3) = 2
    If Me.OB6 = True Then choice(3) = 3
    
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
    ratiotest.ratio2result PHat1, PHat2, zstat, t1, t2, s1, s2, L, resultsheet, choice
    'resultsheet.Protect "prophet"
    ''''Worksheets(RstSheet).Cells(1, 1).Value = Right(Worksheets(RstSheet).Cells(1, 1).Value, Len(Worksheets(RstSheet).Cells(1, 1).Value) - 3)
    
    
    'Worksheets(RstSheet).Protect Password:="prophet", DrawingObjects:=False, _
                                    contents:=True, Scenarios:=True             '''
    


    Worksheets(RstSheet).Activate

    '���� ���� üũ �� �񱳰� ����,
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
