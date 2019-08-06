Attribute VB_Name = "TModuleControl"
Option Base 1
Public DataSheet As String, RstSheet As String      'sheet�̸� �� ��
Public m As Long                                    'dataSheet�� �ִ� ������ ����

Public xlist As String                              '���õ� ������ �̸�     : ��ǥ���� ����
Public n As Long                                    '���õ� ������ ����Ÿ ����  : ��ǥ���� ����
Public k1 As Long                                   '���õ� ������ ���° ���� �ִ���   : ��ǥ���� ����

Public xlist2() As String                           '���õ� �������� �̸�   : ��ǥ���� ����
Public n2() As Long                                 '���õ� �������� ����Ÿ ����    : ��ǥ���� ����
Public k2() As Integer                              '���õ� �������� ���° ���� �ִ��� : ��ǥ���� ����



Sub TShow1()

    Dim ErrSignforDataSheet As Integer
    
    ErrSignforDataSheet = InitializeDlg1(frmTtest1)
                                    
    Select Case ErrSignforDataSheet
    Case 0: frmTtest1.Show
    Case -1
        MsgBox "��Ʈ�� ��ȣ���¿� �ֽ��ϴ�." & Chr(10) & _
               "����Ÿ�� ���� �� �����ϴ�.", _
                vbExclamation, "HIST"
    Case 1
        MsgBox "��Ʈ�� ����Ÿ�� �ִ��� Ȯ���Ͻʽÿ�." & Chr(10) & _
               "1��1������ �����̸��� �Է��ؾ� �մϴ�.", _
               vbExclamation, "HIST"
    Case Else
    End Select
    
End Sub

Function InitializeDlg1(ParentDlg) As Integer
   
   Dim myRange As Range: Dim Cnt As Long
   Dim myArray() As String
   
   On Error GoTo ErrorFlag
   
   Set myRange = ActiveSheet.Cells.CurrentRegion
   If myRange.count = 1 And myRange.Cells(1, 1) = "" Then
        InitializeDlg1 = 1: Exit Function
   End If
   Set myRange = ActiveSheet.Cells.CurrentRegion.Rows(1)
   ParentDlg.ListBox1.Clear: ParentDlg.ListBox2.Clear
   Cnt = myRange.Cells.count
   
   ReDim myArray(0 To Cnt - 1)
   For i = 1 To Cnt
     myArray(i - 1) = myRange.Cells(i)
   Next i
   ParentDlg.ListBox1.List() = myArray
   InitializeDlg1 = 0
   Exit Function
   
ErrorFlag:
   InitializeDlg1 = -1
   
End Function


Sub TShow2()

    Dim ErrSignforDataSheet As Integer
    
    ErrSignforDataSheet = InitializeDlg1(frmTtest2)
                                    
    Select Case ErrSignforDataSheet
    Case 0: frmTtest2.Show
    Case -1
        MsgBox "��Ʈ�� ��ȣ���¿� �ֽ��ϴ�." & Chr(10) & _
               "����Ÿ�� ���� �� �����ϴ�.", _
                vbExclamation, "HIST"
    Case 1
        MsgBox "��Ʈ�� ����Ÿ�� �ִ��� Ȯ���Ͻʽÿ�." & Chr(10) & _
               "1��1������ �����̸��� �Է��ؾ� �մϴ�.", _
               vbExclamation, "HIST"
    Case Else
    End Select
    
    
End Sub

Function InitializeDlg2(ParentDlg) As Integer
   
   Dim myRange As Range: Dim Cnt As Long
   Dim myArray() As String
   
   On Error GoTo ErrorFlag
   
   Set myRange = ActiveSheet.Cells.CurrentRegion
   If myRange.count = 1 And myRange.Cells(1, 1) = "" Then
        InitializeDlg2 = 1: Exit Function
   End If
   Set myRange = ActiveSheet.Cells.CurrentRegion.Rows(1)
   ParentDlg.ListBox1.Clear: ParentDlg.ListBox2.Clear
   ParentDlg.ListBox3.Clear: ParentDlg.ListBox4.Clear: ParentDlg.ListBox5.Clear
   Cnt = myRange.Cells.count
   
   ReDim myArray(0 To Cnt - 1)
   For i = 1 To Cnt
     myArray(i - 1) = myRange.Cells(i)
   Next i
   ParentDlg.ListBox1.List() = myArray
   ParentDlg.ListBox3.List() = myArray
   InitializeDlg2 = 0
   Exit Function
   
ErrorFlag:
   InitializeDlg2 = -1
   
End Function

Function FindingRangeError(ListVar) As Boolean
    
    Dim temp, m2, m3, j As Long
    Dim TempSheet As Worksheet
    Dim tmp As Range, tmp11 As Range, tmp1 As Range, tmp2 As Range, tmp3 As Range
    
    Set TempSheet = Worksheets(DataSheet)
    
   Dim Chk_Ver As Boolean   '���� ���� üũ
   Dim Cmp_R As Long        '���� ������ ���� �� ���� ��
   
   '���� ������ ���� ��� ���� �񱳰� ����
   Chk_Ver = PublicModule.ChkVersion(ActiveWorkbook.Name)
   If Chk_Ver = True Then
        Cmp_R = 1048576
    Else
        Cmp_R = 65536
    End If
    
    For j = 1 To m
       If StrComp(ListVar, TempSheet.Cells(1, j).Value, 1) = 0 Then
          Set tmp11 = TempSheet.Columns(j)
          m2 = tmp11.Cells(1, 1).End(xlDown).row
          If m2 <> Cmp_R Then
             m3 = tmp11.Cells(m2, 1).End(xlDown).row
             If m3 <> Cmp_R Then m2 = m3
          End If
          Set tmp = tmp11.Range(Cells(2, 1), Cells(m2, 1))
       End If
    Next j
    
    On Error Resume Next
    
    If Application.CountBlank(tmp) >= 1 Then
        FindingRangeError = True
        Exit Function
    End If
    Set tmp1 = tmp.SpecialCells(xlCellTypeConstants, 22)
    Set tmp2 = tmp.SpecialCells(xlCellTypeFormulas, 22)
    Set tmp3 = tmp.SpecialCells(xlCellTypeBlanks)
    
    If tmp.count = 1 And IsNumeric(tmp.Cells(1, 1)) = True Then
        FindingRangeError = False
    Else
        If tmp1 Is Nothing And tmp2 Is Nothing And tmp3 Is Nothing Then
            FindingRangeError = False
        Else: FindingRangeError = True
        End If
    End If
    
End Function

Sub SettingStatusBar(SettingChoice As Boolean, _
        Optional NewString As String = "")

    Static oldStatusBar As String
    
    If SettingChoice = True Then
        oldStatusBar = Application.DisplayStatusBar
        Application.DisplayStatusBar = True
        Application.StatusBar = NewString
    Else
        Application.StatusBar = False
        Application.DisplayStatusBar = oldStatusBar
    End If
    
End Sub



'''
''' xlist, n, k1 : public ���� ����Ǿ� ����.
'''
Sub TTest1(choice)
    
    Dim dataArray(), rstArray()
    Dim theta0 As Single, CI As Single, Hyp As Integer
    Dim mySheet As Worksheet
    Dim titleTmp() As String, title As String
    
    '''
    '''����Ÿ �迭
    '''
    Worksheets(DataSheet).Activate
    dataArray = Worksheets(DataSheet).Range(Cells(2, k1), Cells(n + 1, k1)).Value
    
    '''
    '''������ theta0, �ŷڱ��� CI , �͹����� Hyp
    '''
    theta0 = Format(choice(1), "##0.0000")
    CI = Format(choice(2), "##0.0000")
    Hyp = choice(3)
    
    '''
    ''' �����Ʈ Ȱ��ȭ
    '''
    Set mySheet = Worksheets(RstSheet)
    mySheet.Activate
    
    '''
    ''' ���� ����
    '''
    TModulePrint.Title1 "t-���� �м����"
    TModulePrint.Title3 "��ǥ�� ����"


    '''
    ''' ����ϰ� ����ϱ�
    '''
    ReDim rstArray(2, 4)
    rstArray(1, 1) = "������": rstArray(1, 2) = "����": rstArray(1, 3) = "���": rstArray(1, 4) = "ǥ������"
    rstArray(2, 1) = xlist: rstArray(2, 2) = n
    rstArray(2, 3) = Format(WorksheetFunction.Average(dataArray), "##0.0000")
    rstArray(2, 4) = Format(WorksheetFunction.StDev(dataArray), "##0.0000")
    
    TModulePrint.printRst "", rstArray
    
    
    
    ReDim rstArray(2, 3)
    
    rstArray(1, 1) = "t-��跮": rstArray(1, 2) = "������": rstArray(1, 3) = "����Ȯ��"
    rstArray(2, 1) = Format((WorksheetFunction.Average(dataArray) - theta0) / WorksheetFunction.StDev(dataArray) * Sqr(n), "##0.0000")
    rstArray(2, 2) = n - 1
    
    Select Case Hyp
    Case 1
    title = " H : ��=�졣vs. K : ����졣     (�졣= " & theta0 & " )"
    rstArray(2, 3) = Format(WorksheetFunction.TDist(Abs(rstArray(2, 1)), n - 1, 2), "##0.0000")
    
    Case 2
    title = " H : ��=�졣vs. K : �� > �졣     (�졣= " & theta0 & " )"
    rstArray(2, 3) = Format(WorksheetFunction.TDist(Abs(rstArray(2, 1)), n - 1, 1), "##0.0000")
    If rstArray(2, 1) < 0 Then rstArray(2, 3) = Format(1 - WorksheetFunction.TDist(Abs(rstArray(2, 1)), n - 1, 1), "##0.0000")
    
    Case 3
    title = " H : ��=�졣vs. K : �� < �졣     (�졣= " & theta0 & " )"
    rstArray(2, 3) = Format(1 - WorksheetFunction.TDist(Abs(rstArray(2, 1)), n - 1, 1), "##0.0000")
    If rstArray(2, 1) < 0 Then rstArray(2, 3) = Format(WorksheetFunction.TDist(Abs(rstArray(2, 1)), n - 1, 1), "##0.0000")
    End Select
       
    TModulePrint.printRst title, rstArray
 
 
    If CI > 0 Then
    ReDim rstArray(2, 3)

    rstArray(1, 1) = CI & "% �ŷڱ���": rstArray(1, 2) = "����": rstArray(1, 3) = "����"
        
    tmp = WorksheetFunction.TInv(1 - CI / 100, n - 1) * WorksheetFunction.StDev(dataArray) / WorksheetFunction.Power(n, 0.5)
    rstArray(2, 2) = Format(WorksheetFunction.Average(dataArray) - tmp, "##0.0000")
    rstArray(2, 3) = Format(WorksheetFunction.Average(dataArray) + tmp, "##0.0000")

    TModulePrint.printRst "", rstArray
    End If
    
    
End Sub



Sub TTest2(choice, op)
    
    Dim dataArray1(), dataArray2(), data_X(), data_Y(), dataArray()
    Dim CI As Single, Hyp As Integer
    Dim mySheet As Worksheet
    Dim titleTmp() As String
    Dim gpChk(2)
    Dim list_X As String, list_Y As String
    
    Dim avg_X As Single, avg_Y As Single, s_Xsq As Single, s_Ysq As Single
    Dim paTeq As Single, paTuneq As Single
    Dim pvalueTeq As Single, pvalueTuneq As Single
    Dim num_df, dem_df As Integer
    '''
    '''�Ϲ������� ��������, ����Ÿ����, ����Ÿ�迭
    '''
If op = 1 Then                                              'ǥ���Է�
    
    ReDim dataArray1(n2(1))
    ReDim dataArray2(n2(2))

    With Worksheets(DataSheet)
        For j = 1 To n2(1)
            dataArray1(j) = .Cells(j + 1, k2(1)).Value
        Next j
        For j = 1 To n2(2)
            dataArray2(j) = .Cells(j + 1, k2(2)).Value
        Next j
    End With
    
End If

If op = 2 Then                                             '����Է�
    
    Set mySheet = Worksheets(DataSheet)
    
    
    gpChk(1) = mySheet.Cells(2, k2(1)).Value
    For i = 2 To n2(1) + n2(2)
        If gpChk(1) <> mySheet.Cells(i + 1, k2(1)).Value Then
            gpChk(2) = mySheet.Cells(i + 1, k2(1)).Value
            Exit For
        End If
    Next i
    
    
    ReDim dataArray1(n2(1))
    ReDim dataArray2(n2(2))

        j1 = 1: j2 = 1
        
        For j = 1 To n2(1) + n2(2)
            If mySheet.Cells(j + 1, k2(1)).Value = gpChk(1) Then
                dataArray1(j1) = mySheet.Cells(j + 1, k2(2))
                j1 = j1 + 1
            End If
            If mySheet.Cells(j + 1, k2(1)).Value = gpChk(2) Then
                dataArray2(j2) = mySheet.Cells(j + 1, k2(2))
                j2 = j2 + 1
            End If
        Next j
    
    
    For i = 1 To 2
    xlist2(i) = gpChk(i)
    Next i
    
End If


If choice(1) = 1 Then
    '''
    '''�������choice(1), �ŷڱ��� CI , �͹����� Hyp
    '''
    CI = Format(choice(2), "##0.0000")
    Hyp = choice(3)
    
    '''
    ''' ���� ����
    '''
    TModulePrint.Title1 "t-���� �м����"
    TModulePrint.Title3 "��ǥ�� ���� (������)"
        
    If n2(1) <= n2(2) Then
        'i = 1
        list_X = xlist2(2)                  '����Ÿ ������ ���� ���� X
        list_Y = xlist2(1)
        num_X = n2(2)
        num_Y = n2(1)
        ReDim data_X(n2(2))
        ReDim data_Y(n2(1))
        data_X = dataArray2
        data_Y = dataArray1
    Else
        'i = 2
        list_X = xlist2(1)
        list_Y = xlist2(2)
        num_X = n2(1)
        num_Y = n2(2)
        ReDim data_X(n2(1))
        ReDim data_Y(n2(2))
        data_X = dataArray1
        data_Y = dataArray2
    End If
       
    ReDim rstArray(3, 4)
    
    rstArray(1, 1) = "������"
    rstArray(1, 2) = "����"
    rstArray(1, 3) = "���"
    rstArray(1, 4) = "ǥ������"
    rstArray(2, 1) = list_Y
    rstArray(2, 1) = msgWord(rstArray(2, 1))
    rstArray(3, 1) = list_X
    rstArray(3, 1) = msgWord(rstArray(3, 1))
    
    rstArray(2, 2) = num_Y
    rstArray(3, 2) = num_X
    
    rstArray(2, 3) = Format(Application.WorksheetFunction.Average(data_Y), "##0.0000")
    rstArray(3, 3) = Format(Application.WorksheetFunction.Average(data_X), "##0.0000")
    
    rstArray(2, 4) = Format(Application.WorksheetFunction.StDev(data_Y), "##0.0000")
    rstArray(3, 4) = Format(Application.WorksheetFunction.StDev(data_X), "##0.0000")
    
    TModulePrint.printRst "", rstArray
    
    
    
    avg_X = Application.WorksheetFunction.Average(data_X)
    avg_Y = Application.WorksheetFunction.Average(data_Y)
    s_Xsq = Application.WorksheetFunction.var(data_X)
    s_Ysq = Application.WorksheetFunction.var(data_Y)
    spsq = ((num_Y - 1) * s_Ysq + (num_X - 1) * s_Xsq) / (num_Y + num_X - 2)
    df2 = (s_Ysq / num_Y + s_Xsq / num_X) ^ 2 / ((s_Ysq / num_Y) ^ 2 / (num_Y - 1) + (s_Xsq / num_X) ^ 2 / (num_X - 1))
    df22 = Application.RoundDown(df2, 0)
    paTeq = (avg_Y - avg_X) / Sqr(spsq) / Sqr(1 / num_Y + 1 / num_X)
    paTuneq = (avg_Y - avg_X) / Sqr(s_Ysq / num_Y + s_Xsq / num_X)
    
    '''fstat = s_Ysq / s_Xsq
    '''If fstat > 1 Then
    '''pValueF = 2 * Application.FDist(fstat, num_Y - 1, num_X - 1)
    '''Else: pValueF = 2 * Application.FDist(1 / fstat, num_X - 1, num_Y - 1)
    '''End If
    
    fstat = WorksheetFunction.Max(s_Ysq, s_Xsq) / WorksheetFunction.Min(s_Ysq, s_Xsq)
    If WorksheetFunction.Max(s_Ysq, s_Xsq) = s_Ysq Then
    dem_df = num_Y
    num_df = num_X
    Else
    dem_df = num_X
    num_df = num_Y
    End If
    
    pvaluef = 2 * Application.FDist(fstat, dem_df - 1, num_df - 1)
 
    
    pvalueTeq = Application.WorksheetFunction.TDist(Abs(paTeq), num_Y + num_X - 2, 1)
    pvalueTuneq = Application.WorksheetFunction.TDist(Abs(paTuneq), df22, 1)
    tmp = df2 - df22
    pvalueTuneq = (1 - tmp) * Application.WorksheetFunction.TDist(Abs(paTuneq), df22, 1) + tmp * Application.WorksheetFunction.TDist(Abs(paTuneq), df22 + 1, 1)
        
    Select Case choice(3)
    Case 1
        pvalueTeq = Application.WorksheetFunction.Min(pvalueTeq, 1 - pvalueTeq) * 2
        pvalueTuneq = Application.WorksheetFunction.Min(pvalueTuneq, 1 - pvalueTuneq) * 2
        title = " H : ���= ���vs. K : ����� ���     (���: " & list_Y & ", ���: " & list_X & " )"
    Case 2
        If paTeq < 0 Then pvalueTeq = 1 - pvalueTeq
        If paTuneq < 0 Then pvalueTuneq = 1 - pvalueTuneq
        title = " H : ���= ���vs. K : ���> ���      (���: " & list_Y & ", ���: " & list_X & " )"
    Case 3
        If paTeq > 0 Then pvalueTeq = 1 - pvalueTeq
        If paTuneq > 0 Then pvalueTuneq = 1 - pvalueTuneq
        title = " H : ���= ���vs. K : ���< ���      (���: " & list_Y & ", ���: " & list_X & " )"
    End Select
    
    TT = "��л� ����"
    ReDim rstArray(2, 3)
    rstArray(1, 1) = "������"
    rstArray(1, 2) = "F ��"
    rstArray(1, 3) = "����Ȯ��"
    rstArray(2, 1) = "( " & dem_df - 1 & " , " & num_df - 1 & " )"
    rstArray(2, 2) = Format(fstat, "##0.0000")
    rstArray(2, 3) = Format(pvaluef, "##0.0000")
    TModulePrint.printRst TT, rstArray
    
    
    If pvaluef <= 0.01 Then
            Comment1 = """H0:�� ǥ���� �л���� ���� ����.""" & "�� ���Ǽ��� ��=0.01���� �Ⱒ�Ѵ�."
            
        ElseIf pvaluef <= 0.05 Then
            Comment1 = """H0:�� ǥ���� �л���� ���� ����.""" & "�� ���Ǽ��� ��=0.05���� �Ⱒ�Ѵ�."
            
        Else
            Comment1 = """H0:�� ǥ���� �л���� ���� ����.""" & "�� ���Ǽ��� ��=0.05���� �Ⱒ�� �� ����."
            
        End If
    Comment2 = "������Ȯ���� ���Ǽ��غ��� ū ��쿡�� ��л� ����� ����ϴ� ���� ����. "
    
    
    
    Set mySheet = Worksheets(RstSheet)
    mySheet.Activate ''''''''''''
    mySheet.Cells(1, 1) = mySheet.Cells(1, 1) + 1
    
        mySheet.Cells(1, 1) = mySheet.Cells(1, 1) + 2
        Flag = mySheet.Cells(1, 1).Value
        mySheet.Cells(Flag - 2, 2) = Comment1
        mySheet.Cells(Flag - 2, 2).Font.Size = 10
        mySheet.Cells(Flag - 2, 2).Font.Bold = True
        mySheet.Cells(Flag - 2, 2).HorizontalAlignment = xlGeneral
        mySheet.Cells(Flag - 1, 2) = Comment2
        mySheet.Cells(Flag - 1, 2).Font.Size = 10
        mySheet.Cells(Flag - 1, 2).Font.Bold = True
        mySheet.Cells(Flag - 1, 2).HorizontalAlignment = xlGeneral
    
    
    ReDim rstArray(3, 4)
    rstArray(1, 1) = "�л�"
    rstArray(2, 1) = "��л�"
    rstArray(3, 1) = "�̺л�"
    
    rstArray(1, 2) = "t-��跮"
    rstArray(1, 3) = "������"
    rstArray(1, 4) = "����Ȯ��"
    rstArray(2, 2) = Format(paTeq, "##0.0000")
    rstArray(2, 3) = num_Y + num_X - 2
    rstArray(2, 4) = Format(pvalueTeq, "##0.0000")
    rstArray(3, 2) = Format(paTuneq, "##0.0000")
    rstArray(3, 3) = Format(df2, "##0.0000")
    rstArray(3, 4) = Format(pvalueTuneq, "##0.0000")
    
    TModulePrint.printRst title, rstArray
    
    If CI > 0 Then
    ReDim rstArray(3, 3)

    rstArray(1, 1) = CI & "% �ŷڱ���": rstArray(1, 2) = "����": rstArray(1, 3) = "����"
    rstArray(2, 1) = "��л�": rstArray(3, 1) = "�̺л�"
    
    Alpha = 1 - CI / 100
    var = ((num_X - 1) * s_Xsq + (num_Y - 1) * s_Ysq) / (num_X + num_Y - 2) * (1 / num_X + 1 / num_Y)
    df = num_X + num_Y - 2
    tmp = cilength(var, df, Alpha)
    rstArray(2, 2) = Format(avg_Y - avg_X - tmp, "##0.0000")
    rstArray(2, 3) = Format(avg_Y - avg_X + tmp, "##0.0000")

    var = s_Xsq / num_X + s_Ysq / num_Y
    tmp = cilength(var, df2, Alpha)
    rstArray(3, 2) = Format(avg_Y - avg_X - tmp, "##0.0000")
    rstArray(3, 3) = Format(avg_Y - avg_X + tmp, "##0.0000")

    TModulePrint.printRst "", rstArray
    End If
    
End If

If choice(1) = 2 Then

    '''
    '''�������choice(1), �ŷڱ��� CI , �͹����� Hyp
    '''
    CI = Format(choice(2), "##0.0000")
    Hyp = choice(3)
    
    '''
    ''' ���� ����
    '''
    TModulePrint.Title1 "t-���� �м����"
    TModulePrint.Title3 "��ǥ�� ���� (������)"
        
    ReDim rstArray(3, 4)
    
    rstArray(1, 1) = "������"
    rstArray(1, 2) = "����"
    rstArray(1, 3) = "���"
    rstArray(1, 4) = "ǥ������"
    rstArray(2, 1) = xlist2(1)
    rstArray(2, 1) = msgWord(rstArray(2, 1))
    rstArray(3, 1) = xlist2(2)
    rstArray(3, 1) = msgWord(rstArray(3, 1))
    
    rstArray(2, 2) = n2(1)
    rstArray(3, 2) = n2(2)
    
    rstArray(2, 3) = Format(Application.WorksheetFunction.Average(dataArray1), "##0.0000")
    rstArray(3, 3) = Format(Application.WorksheetFunction.Average(dataArray2), "##0.0000")
    
    rstArray(2, 4) = Format(Application.WorksheetFunction.StDev(dataArray1), "##0.0000")
    rstArray(3, 4) = Format(Application.WorksheetFunction.StDev(dataArray2), "##0.0000")
    
    TModulePrint.printRst "", rstArray

    ReDim dataArray(n2(1))
    
    For i = 1 To n2(1)
    dataArray(i) = dataArray1(i) - dataArray2(i)
    Next i
    
    ReDim rstArray(2, 3)
    
    rstArray(1, 1) = "t-��跮": rstArray(1, 2) = "������": rstArray(1, 3) = "����Ȯ��"
    rstArray(2, 1) = Format((WorksheetFunction.Average(dataArray)) / WorksheetFunction.StDev(dataArray) * Sqr(n2(1)), "##0.0000")
    rstArray(2, 2) = n2(1) - 1
    
    Select Case Hyp
    Case 1
    title = " H : ���- ���= 0 vs. K : ���- �����0      (���: " & xlist2(1) & ", ���: " & xlist2(2) & " )"
    rstArray(2, 3) = Format(WorksheetFunction.TDist(Abs(rstArray(2, 1)), n2(1) - 1, 2), "##0.0000")
    
    Case 2
    title = " H : ���- ���= 0 vs. K : ���- ��� > 0     (���: " & xlist2(1) & ", ���: " & xlist2(2) & " )"
    rstArray(2, 3) = Format(WorksheetFunction.TDist(Abs(rstArray(2, 1)), n2(1) - 1, 1), "##0.0000")
    If rstArray(2, 1) < 0 Then rstArray(2, 3) = Format(1 - WorksheetFunction.TDist(Abs(rstArray(2, 1)), n2(1) - 1, 1), "##0.0000")
    
    Case 3
    title = " H : ���- ���= 0 vs. K : ���- ��� < 0     (���: " & xlist2(1) & ", ���: " & xlist2(2) & " )"
    rstArray(2, 3) = Format(1 - WorksheetFunction.TDist(Abs(rstArray(2, 1)), n2(1) - 1, 1), "##0.0000")
    If rstArray(2, 1) < 0 Then rstArray(2, 3) = Format(WorksheetFunction.TDist(Abs(rstArray(2, 1)), n2(1) - 1, 1), "##0.0000")
    End Select
       
    TModulePrint.printRst title, rstArray
 
 
    If CI > 0 Then
    ReDim rstArray(2, 3)

    rstArray(1, 1) = CI & "% �ŷڱ���": rstArray(1, 2) = "����": rstArray(1, 3) = "����"
        
    tmp = WorksheetFunction.TInv(1 - CI / 100, n2(1) - 1) * WorksheetFunction.StDev(dataArray) / WorksheetFunction.Power(n2(1), 0.5)
    rstArray(2, 2) = Format(WorksheetFunction.Average(dataArray) - tmp, "##0.0000")
    rstArray(2, 3) = Format(WorksheetFunction.Average(dataArray) + tmp, "##0.0000")

    TModulePrint.printRst "", rstArray
    End If
    

End If
    
    
    
    
    
    
    
   
    
End Sub




Function msgWord(word) As String                    '������ word �� string

    If Len(word) > 9 Then
        msgWord = Mid(word, 1, 6) & vbLf & Mid(word, 7)
    Else
        msgWord = word
    End If
    
End Function



Function cilength(var, df, Alpha)
    If df = Application.RoundDown(df, 0) Then
       cilength = Application.TInv(Alpha, df) * Application.Power(var, 0.5)
    Else: df_d = Application.RoundDown(df, 0)
       t_d = Application.TInv(Alpha, df_d)
       t_u = Application.TInv(Alpha, df_d + 1)
       W = df - df_d
       t_value = (1 - W) * t_d + W * t_u
       cilength = t_value * Application.Power(var, 0.5)
    End If
End Function


'�ڽ��÷�
Private Sub boxgraph()

    Dim dataArray(), rstArray()
    Dim theta0 As Single, CI As Single, Hyp As Integer
    Dim mySheet As Worksheet
    Dim titleTmp() As String, title As String


    Worksheets(DataSheet).Activate
    dataArray = Worksheets(DataSheet).Range(Cells(2, k1), Cells(n + 1, k1)).Value
    
    Dim i As Long, nCol As Long, nRow As Long
    Dim nGrp As Integer, nPage As Integer
    Dim IDcht As Chart
    Dim strName As String
    Dim rngData As Range, rngName As Range, rngTitle As Range, rngTmp() As Range
    Dim rngFirst As Range
    Dim cc1(10) As Double, cc2 As Double
    Dim arrTlt() As String

    On Error GoTo ErrEnd
    Application.ScreenUpdating = False
    
'   Read data
  '  Set rngData = Range(Me.RefEdit1)
     Set rngData = Range(Cells(2, k1), Cells(n + 1, k1))
    
'   Count the # of rows and columns
   ' nGrp = rngData.Columns.Count
     nGrp = 1
    nRow = rngData.Rows.count
    
'   Resize data
    ReDim arrTlt(nGrp)

            arrTlt(1) = frameInterval.ListBox2.List(0)
           ' arrTlt(2) = "��ǥ ���"
            

'''
''' Err Check           �Ѱ��� �־ �� ������ �ϳ��� ��հ����� ���
'''

  '  If nGrp < 2 Then
 '       MsgBox "2�� �̻��� �׷��� �����ؾ� �մϴ�."
 '       Exit Sub
  '  End If
    
    ReDim rngTmp(nGrp)
    For i = 1 To nGrp
        Set rngTmp(i) = rngData.Columns(i)
    Next i

'   Add a Sheet
    For i = 1 To Sheets.count
        If Sheets(i).Name = "_���м����_" Then
            GoTo 31
        Else
            GoTo 32
        End If
        
32: Next i
    Worksheets.Add Before:=Worksheets(1)
    ActiveSheet.Name = "_���м����_"
    ActiveWindow.DisplayGridlines = False
    Cells(1, 1) = 1

31: Sheets("_���м����_").Activate
    Application.ScreenUpdating = False

'   Current worksheet's name
    strName = ActiveSheet.Name

'   Chart Location
    Set rngFirst = Cells(Cells(1, 1) + 3, 1)

Set IDcht = Charts.Add
Set IDcht = IDcht.Location(where:=xlLocationAsObject, Name:=strName)

With IDcht
    .ChartType = xlLineMarkers
    .HasLegend = False
   ' .HasTitle = True

 '   With .ChartTitle
 '       .Characters.Text = "Individual Values Plot"
 '       .Font.Size = 12
 '       .Font.Bold = True
 '       .AutoScaleFont = False
 '   End With
    
    For i = 1 To nRow
        .SeriesCollection.NewSeries
    Next i
    On Error Resume Next
    If .SeriesCollection(nRow) Then .SeriesCollection(nCol + 1).Delete
    .SeriesCollection(1).XValues = rngTitle
    .SeriesCollection(1).XValues = arrTlt
    
    For i = 1 To nRow
    
    With .SeriesCollection(i)
        .Values = rngData.Rows(i)
        .Border.LineStyle = xlNone
        .MarkerStyle = xlMarkerStyleCircle
        '.MarkerSize = 3
        .MarkerForegroundColorIndex = 5
        .MarkerBackgroundColorIndex = 6
    End With
    Next i

    With .Axes(xlValue, xlPrimary)
        .HasTitle = False
        .HasMajorGridlines = False
        .HasMinorGridlines = False
        .MinimumScaleIsAuto = True

'   Ajusting Y axis value
        .MinimumScaleIsAuto = False
        For i = 1 To 10
            If WorksheetFunction.Min(rngData) > .MinimumScale + .MajorUnit * 2 Then
                .MinimumScale = .MinimumScale + .MajorUnit
            End If
        Next i
    End With

    With .Parent
        .Top = rngFirst.Offset(2, 1).Top
        .Left = rngFirst.Offset(2, 1).Left
        .Width = 240
        .Height = 180
    End With

End With

ErrEnd:

'   Page number reset
    nPage = 25
   ' rngFirst.Offset = "Created at " & Now()
    Application.Goto rngFirst, Scroll:=True
    Range("A1") = Range("A1") + nPage

    Application.ScreenUpdating = True
   ' Unload Me

End Sub
