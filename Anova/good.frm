VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} good 
   OleObjectBlob   =   "good.frx":0000
   Caption         =   "���յ� ����"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3840
   StartUpPosition =   1  '������ ���
   TypeInfoVer     =   21
End
Attribute VB_Name = "good"
Attribute VB_Base = "0{0DB8CB45-3E4C-413D-B4EE-357CF9C1606B}{C39DA695-6419-4E2C-AC7F-BEEA630A8F64}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False


Private Sub BtnCan_Click()
  Unload Me
End Sub

Private Sub ok_btn_Click()                                          ''''"_�������ڷ�м����_"
  Dim TempSheet As Worksheet:  Dim resultsheet As Worksheet
  Dim temp, NumRn As Range: Dim VarName() As String
  Dim raw() As Integer
  Dim Expect() As Double
  Dim cnt As Integer
  Dim i As Integer
  Dim total1, total2, chis As Double
  Set TempSheet = ActiveCell.Worksheet
  total1 = 0
  total2 = 0
  chis = 0
  Set temp = TempSheet.Cells.CurrentRegion
  '''���� üũ
  Set NumRn = temp.Offset(1, 0).Resize(temp.Rows.count - 1, temp.Columns.count)
  If FindingRangeError(NumRn) = True Then
        MsgBox "�м������� ���ڳ� ������ �ֽ��ϴ�.", vbExclamation, "HIST"
        Exit Sub
  End If

  If column_btn.Value = True Then
    Set NumRn = temp.Offset(1, 0).Resize(temp.Rows.count - 1, temp.Columns.count)
    If FindingRangeError(NumRn) = True Then
          MsgBox "�м������� ���ڳ� ������ �ֽ��ϴ�.", vbExclamation, "HIST"
          Exit Sub
    End If
     cnt = temp.Columns.count
     ReDim VarName(1 To cnt)
     ReDim Expect(1 To cnt)
     ReDim raw(1 To cnt)
     For i = 1 To cnt
         VarName(i) = temp.Cells(1, i).Value
         raw(i) = temp.Cells(2, i).Value
         total1 = total1 + temp.Cells(2, i).Value
         total2 = total2 + temp.Cells(3, i).Value
     Next i
     For i = 1 To cnt
         Expect(i) = total1 * (temp.Cells(3, i).Value / total2)
         chis = chis + (temp.Cells(2, i) - Expect(i)) ^ 2 / Expect(i)
     Next i
   Else: cnt = temp.Rows.count
    Set NumRn = temp.Offset(0, 1).Resize(temp.Rows.count, temp.Columns.count - 1)
    If FindingRangeError(NumRn) = True Then
          MsgBox "�м������� ���ڳ� ������ �ֽ��ϴ�.", vbExclamation, "HIST"
          Exit Sub
    End If
     ReDim VarName(1 To cnt)
     ReDim Expect(1 To cnt)
     ReDim raw(1 To cnt)
     For i = 1 To cnt
         VarName(i) = temp.Cells(i, 1).Value
         raw(i) = temp.Cells(i, 2).Value
         total1 = total1 + temp.Cells(i, 2).Value
         total2 = total2 + temp.Cells(i, 3).Value
     Next i
     For i = 1 To cnt
         Expect(i) = total1 * (temp.Cells(i, 3).Value / total2)
         chis = chis + (temp.Cells(i, 2) - Expect(i)) ^ 2 / Expect(i)
     Next i
   End If
   Set resultsheet = OpenOutSheet("_���м����_", True)
   
   '''
    '''
    '''
    RstSheet = "_���м����_"
    '����ϴ� �ش� ��⿡ �� ���� ����'
'������ �Է�
On Error GoTo Err_delete
Dim val3535 As Long '�ʱ���ġ ������ ����'
Dim s3535 As Worksheet
val3535 = 2
    For Each s3535 In ActiveWorkbook.Sheets
        If s3535.name = RstSheet Then
val3535 = Sheets(RstSheet).Cells(1, 1).Value
End If
Next s3535  '��Ʈ�� �̹������� ��� ��ġ �������ϰ�, ������ 2�� �����Ѵ�.
   ' Worksheets(RstSheet).Unprotect "prophet"
    activePt = Worksheets(RstSheet).Cells(1, 1).Value
   
    
   'resultsheet.Unprotect "prophet"
   Good_result.gresult chis, VarName, raw, Expect, total1, cnt, resultsheet
   'resultsheet.Protect Password:="prophet", DrawingObjects:=False, contents:=True, Scenarios:=True
   'resultsheet.Unprotect "prophet"
    ''''ratiotest.ratioresult zstat, phat, r, T, s, L, WarningMsg, resultsheet
    '''' ���� ���� ���� �ǹ����� �𸣰ھ �ϴ� ���� ��Ŵ   �� ������ �ԷµǴ� ���� ����.
   'resultsheet.Protect "prophet"
    
    '''Worksheets(RstSheet).Cells(1, 1).Value = Right(Worksheets(RstSheet).Cells(1, 1).Value, Len(Worksheets(RstSheet).Cells(1, 1).Value) - 3)



    Worksheets(RstSheet).Activate

    '���� ���� üũ �� �񱳰� ����
    Dim Cmp_Value As Long
    
    If ModuleControl.ChkVersion(ActiveWorkbook.name) = True Then
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
        If s3535.name = RstSheet Then
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

Private Sub UserForm_Terminate()
     Unload Me
End Sub
