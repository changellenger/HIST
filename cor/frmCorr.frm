VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCorr 
   OleObjectBlob   =   "frmCorr.frx":0000
   Caption         =   "����м�"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7080
   StartUpPosition =   1  '������ ���
   TypeInfoVer     =   18
End
Attribute VB_Name = "frmCorr"
Attribute VB_Base = "0{DB2617FE-3D63-41C5-8A79-4D74C77F553B}{88404FE6-1331-4E88-9D1C-C27D0924086C}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Sub MoveBtwnListBox(ParentD, FromLNum, ToLNum)

    Dim i As Integer
    i = 0
    Do While i <= ParentD.Controls(FromLNum).ListCount - 1
        If ParentD.Controls(FromLNum).Selected(i) = True Then
           ParentD.Controls(ToLNum).AddItem ParentD.Controls(FromLNum).list(i)
           ParentD.Controls(FromLNum).RemoveItem i
            Exit Do
        End If
        i = i + 1
    Loop

End Sub
Private Sub Cancel_Click()

    Unload Me
    
End Sub
Private Sub CB1_Click()

    MoveBtwnListBox Me, "ListBox1", "ListBox2"

End Sub
Private Sub CB2_Click()

    MoveBtwnListBox Me, "ListBox2", "ListBox1"
    
End Sub
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    MoveBtwnListBox Me, "ListBox1", "ListBox2"
    
End Sub
Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    MoveBtwnListBox Me, "ListBox2", "ListBox1"
    
End Sub
Private Sub OK_Click()

    Dim dataRange As Range
    Dim i As Integer, J As Integer, errSign As Boolean, errTmp As Integer, sheetApproxRowNum
    Dim activePt As Long, ErrString As String                                   '' ��� �м��� ���۵Ǵ� �κ��� �����ֱ� ����
    
    '''
    ''' ���� ó�� �κ� 0: 1�� 1��
    '''
    If ActiveSheet.Cells(1, 1) = "" Then
        MsgBox "1�� 1���� �������� �ʿ��մϴ�.", vbExclamation                  '' 1�� 1���� ��� ���� ���� �߻�.
        Exit Sub
    End If
    
    '''
    ''' ���� ó�� �κ� 1: ���� ���ÿ��� Ȯ��
    '''
    If Me.ListBox2.ListCount = 0 Then
        MsgBox "������ �������� �ʾҽ��ϴ�.", vbExclamation
        Exit Sub
    End If

    '''
    ''' �Է¹��� ���� �����ϱ�
    '''
    ''' ������� ModeuleControl ���� ����� Public ����
    ''' ���⼭ �ѹ��� �������ش�
    ''' sheetRowNum, sheetColNum, DataSheet, RstSheet, xlist, n, m, p
    '''
    If right(ActiveWorkbook.Name, 4) = ".xls" Or right(ActiveWorkbook.Name, 4) = ".XLS" Then
        sheetRowNum = 2 ^ 16            '65536
        sheetColNum = 2 ^ 8             '256
        sheetApproxRowNum = 65000
    Else
        sheetRowNum = 2 ^ 20            '1048576
        sheetColNum = 2 ^ 14            '16384
        sheetApproxRowNum = 1048000
    End If
    
    DataSheet = ActiveSheet.Name                                                '' Data�� �ִ� Sheet �̸�
    rstSheet = "_���м����_"                                                 '' ����� �����ִ� Sheet �̸�
    '����ϴ� �ش� ��⿡ �� ���� ����'
'������ �Է�
On Error GoTo Err_delete
Dim val3535 As Long '�ʱ���ġ ������ ����'
Dim s3535 As Worksheet
val3535 = 2
    For Each s3535 In ActiveWorkbook.Sheets
        If s3535.Name = rstSheet Then
val3535 = Sheets(rstSheet).Cells(1, 1).Value
End If
Next s3535  '��Ʈ�� �̹������� ��� ��ġ �������ϰ�, ������ 2�� �����Ѵ�.

    
        
    Set dataRange = ActiveSheet.Cells.CurrentRegion
    m = dataRange.Cells(1, 1).End(xlToRight).Column                             '' ��ü �������� ����
    
    p = Me.ListBox2.ListCount                                                   '' ���õ� �������� ����
    ReDim xlist(p - 1)
    For i = 0 To p - 1
        xlist(i) = ListBox2.list(i)                                             '' ���õ� �������� �̸�
    Next i
    
    N = ModuleControl.FindDataCount(xlist) - 1                                  '' ���õ� ������ Data����
    
    '''
    ''' ���� ó�� �κ� 2: �������� �������� ����
    '''
    errSign = False
    For i = 0 To p - 1
        If N <> ModuleControl.FindColDataCount(xlist(i)) Then errSign = True
    Next i
    
    If errSign = True Then
        MsgBox "���õ� �׸�鰣�� �������� �ٸ��ϴ�.", vbExclamation, "HIST"
        Exit Sub
    End If
    errSign = False
    
    '''
    '''���� ó�� �κ� 3: ���ڿ� ���ڰ� ȥ�յǾ� ���� ���
    '''
    For i = 0 To p - 1
        If ModuleControl.FindingRangeError(xlist(i)) = True Then
            errSign = True
            If ErrString <> "" Then
                ErrString = ErrString & "," & xlist(i)
            Else: ErrString = xlist(i)
            End If
        End If
    Next i
    
    If errSign = True Then
        MsgBox "������ �м������� ���ڳ� ������ �ֽ��ϴ�." & vbCrLf & vbCrLf & ErrString, vbExclamation, "HIST"
        Exit Sub
    End If
    errSign = False
    
    '''
    ''' ���� ó�� �κ� 4: �������� ���� ��� - ������ ���� �ִ� ������ �ԷµǹǷ� ����ó���Ѵ�.
    '''
    For i = 1 To p
        errTmp = 0
        For J = 1 To m
            If Me.ListBox2.list(i - 1) = ActiveSheet.Cells(1, J) Then
                errTmp = errTmp + 1
            End If
        Next J
        If errTmp > 1 Then
            MsgBox xlist(i - 1) & vbCrLf & vbCrLf & "���� �м������� ���� �������� �ֽ��ϴ�. " & vbCrLf & "�������� �ٲ��ֽñ� �ٶ��ϴ�.", vbExclamation, "HIST"
            Exit Sub
        End If
    Next i
    
    '''
    '''��� ó��
    '''
    ModuleControl.SettingStatusBar True, "����м����Դϴ�."
    Application.ScreenUpdating = False
    
    ModulePrint.makeOutputSheet rstSheet
    activePt = Worksheets(rstSheet).Cells(1, 1).Value
    
    ModuleControl.CorrAnal
    ModuleControl.SettingStatusBar False
    Application.ScreenUpdating = True
    Unload Me
    
    Worksheets(rstSheet).Activate
    Worksheets(rstSheet).Cells(activePt, 1).Select
    Worksheets(rstSheet).Cells(activePt, 1).Activate                            '��� �м��� ���۵Ǵ� �κ��� �����ָ� ��ģ��.
    
    Worksheets(rstSheet).Activate
    If Worksheets(rstSheet).Cells(1, 1).Value > sheetApproxRowNum Then
        MsgBox "[_���м����_]��Ʈ�� ���� ��� ����Ͽ����ϴ�." & vbCrLf & "�� ��Ʈ�� �̸��� �ٲٰų� ������ �ּ���", vbExclamation, "HIST"
        Exit Sub
    End If


'�ǵڿ� ���̱�
Exit Sub
Err_delete:

For Each s3535 In ActiveWorkbook.Sheets
        If s3535.Name = rstSheet Then
Sheets(rstSheet).Range(Cells(val3535, 1), Cells(5000, 1000)).Select
Selection.Delete
Sheets(rstSheet).Cells(1, 1) = val3535
Sheets(rstSheet).Cells(val3535, 1).Select

If val3535 = 2 Then
Application.DisplayAlerts = False
Sheets(rstSheet).Delete
End If

End If


Next s3535

MsgBox ("���α׷��� ������ �ֽ��ϴ�.")
 'End sub �տ��� ���δ�.

''�ؼ�, ������ ���� Err_delete�� �ͼ� ù�����ķ� �����. ���� ù���� 2�� ��Ʈ�� �����.�׸��� �����޽��� ���
'rSTsheet����⵵ ���� �������� ��쿡�� �ƹ� ���۵� ���� �ʰ�, �����޽����� ����.
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

Private Sub UserForm_Click()

End Sub
