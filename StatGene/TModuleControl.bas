Attribute VB_Name = "TModuleControl"
Option Base 1
Public DataSheet As String, RstSheet As String      'sheet이름 두 개
Public m As Long                                    'dataSheet에 있는 변수의 개수

Public xlist As String                              '선택된 변수의 이름     : 일표본용 변수
Public n As Long                                    '선택된 변수의 데이타 갯수  : 일표본용 변수
Public k1 As Long                                   '선택된 변수가 몇번째 열에 있는지   : 일표본용 변수

Public xlist2() As String                           '선택된 변수들의 이름   : 이표본용 변수
Public n2() As Long                                 '선택된 변수들의 데이타 갯수    : 이표본용 변수
Public k2() As Integer                              '선택된 변수들이 몇번째 열에 있는지 : 이표본용 변수

Sub TShow1()

    Dim ErrSignforDataSheet As Integer
    
    ErrSignforDataSheet = InitializeDlg1(frmTtest1)
                                    
    Select Case ErrSignforDataSheet
    Case 0: frmTtest1.Show
    Case -1
        MsgBox "시트가 보호상태에 있습니다." & Chr(10) & _
               "데이타를 읽을 수 없습니다.", _
                vbExclamation, "HIST"
    Case 1
        MsgBox "시트에 데이타가 있는지 확인하십시오." & Chr(10) & _
               "1행1열부터 변수이름을 입력해야 합니다.", _
               vbExclamation, "HIST"
    Case Else
    End Select
    
End Sub

Function InitializeDlg1(ParentDlg) As Integer
   
   Dim myRange As Range: Dim cnt As Long
   Dim myArray() As String
   
   On Error GoTo ErrorFlag
   
   Set myRange = ActiveSheet.Cells.CurrentRegion
   If myRange.Count = 1 And myRange.Cells(1, 1) = "" Then
        InitializeDlg1 = 1: Exit Function
   End If
   Set myRange = ActiveSheet.Cells.CurrentRegion.Rows(1)
   ParentDlg.ListBox1.Clear: ParentDlg.ListBox2.Clear
   cnt = myRange.Cells.Count
   
   ReDim myArray(0 To cnt - 1)
   For i = 1 To cnt
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
        MsgBox "시트가 보호상태에 있습니다." & Chr(10) & _
               "데이타를 읽을 수 없습니다.", _
                vbExclamation, "HIST"
    Case 1
        MsgBox "시트에 데이타가 있는지 확인하십시오." & Chr(10) & _
               "1행1열부터 변수이름을 입력해야 합니다.", _
               vbExclamation, "HIST"
    Case Else
    End Select
    
    
End Sub

Function InitializeDlg2(ParentDlg) As Integer
   
   Dim myRange As Range: Dim cnt As Long
   Dim myArray() As String
   
   On Error GoTo ErrorFlag
   
   Set myRange = ActiveSheet.Cells.CurrentRegion
   If myRange.Count = 1 And myRange.Cells(1, 1) = "" Then
        InitializeDlg2 = 1: Exit Function
   End If
   Set myRange = ActiveSheet.Cells.CurrentRegion.Rows(1)
   ParentDlg.ListBox1.Clear: ParentDlg.ListBox2.Clear
   ParentDlg.ListBox3.Clear: ParentDlg.ListBox4.Clear: ParentDlg.ListBox5.Clear
   cnt = myRange.Cells.Count
   
   ReDim myArray(0 To cnt - 1)
   For i = 1 To cnt
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
    
    Dim temp, M2, m3, j As Long
    Dim TempSheet As Worksheet
    Dim tmp As Range, tmp11 As Range, tmp1 As Range, tmp2 As Range, tmp3 As Range
    
    Set TempSheet = Worksheets(DataSheet)
    
   Dim Chk_Ver As Boolean   '파일 버전 체크
   Dim Cmp_R As Long        '파일 버전에 따른 비교 행의 값
   
   '파일 버전에 따른 행과 열의 비교값 정의
   Chk_Ver = PublicModule.ChkVersion(ActiveWorkbook.Name)
   If Chk_Ver = True Then
        Cmp_R = 1048576
    Else
        Cmp_R = 65536
    End If
    
    For j = 1 To m
       If StrComp(ListVar, TempSheet.Cells(1, j).Value, 1) = 0 Then
          Set tmp11 = TempSheet.Columns(j)
          M2 = tmp11.Cells(1, 1).End(xlDown).row
          If M2 <> Cmp_R Then
             m3 = tmp11.Cells(M2, 1).End(xlDown).row
             If m3 <> Cmp_R Then M2 = m3
          End If
          Set tmp = tmp11.Range(Cells(2, 1), Cells(M2, 1))
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
    
    If tmp.Count = 1 And IsNumeric(tmp.Cells(1, 1)) = True Then
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
''' xlist, n, k1 : public 변수 선언되어 있음.
'''
Sub TTest1(choice)
    
    Dim dataArray(), rstArray()
    Dim theta0 As Single, CI As Single, Hyp As Integer
    Dim mySheet As Worksheet
    Dim titleTmp() As String, title As String
    
    '''
    '''데이타 배열
    '''
    Worksheets(DataSheet).Activate
    dataArray = Worksheets(DataSheet).Range(Cells(2, k1), Cells(n + 1, k1)).Value
    
    '''
    '''검정값 theta0, 신뢰구간 CI , 귀무가설 Hyp
    '''
    theta0 = Format(choice(1), "##0.0000")
    CI = Format(choice(2), "##0.0000")
    Hyp = choice(3)
    
    '''
    ''' 결과시트 활성화
    '''
    Set mySheet = Worksheets(RstSheet)
    mySheet.Activate
    
    '''
    ''' 제목 쓰기
    '''
    TModulePrint.Title1 "t-검정 분석결과"
    TModulePrint.Title3 "단일표본 t검정"


    '''
    ''' 계산하고 출력하기
    '''
    ReDim rstArray(2, 4)
    rstArray(1, 1) = "변수명": rstArray(1, 2) = "개수": rstArray(1, 3) = "평균": rstArray(1, 4) = "표준편차"
    rstArray(2, 1) = xlist: rstArray(2, 2) = n
    rstArray(2, 3) = Format(WorksheetFunction.Average(dataArray), "##0.0000")
    rstArray(2, 4) = Format(WorksheetFunction.StDev(dataArray), "##0.0000")
    
    TModulePrint.printRst "", rstArray
    
    ReDim rstArray(2, 3)
    
    rstArray(1, 1) = "t-통계량": rstArray(1, 2) = "자유도": rstArray(1, 3) = "유의확률"
    rstArray(2, 1) = Format((WorksheetFunction.Average(dataArray) - theta0) / WorksheetFunction.StDev(dataArray) * Sqr(n), "##0.0000")
    rstArray(2, 2) = n - 1
    
    Select Case Hyp
    Case 1
    title = " H : μ=μ。vs. K : μ≠μ。     (μ。= " & theta0 & " )"
    rstArray(2, 3) = Format(WorksheetFunction.TDist(Abs(rstArray(2, 1)), n - 1, 2), "##0.0000")
    
    Case 2
    title = " H : μ=μ。vs. K : μ > μ。     (μ。= " & theta0 & " )"
    rstArray(2, 3) = Format(WorksheetFunction.TDist(Abs(rstArray(2, 1)), n - 1, 1), "##0.0000")
    If rstArray(2, 1) < 0 Then rstArray(2, 3) = Format(1 - WorksheetFunction.TDist(Abs(rstArray(2, 1)), n - 1, 1), "##0.0000")
    
    Case 3
    title = " H : μ=μ。vs. K : μ < μ。     (μ。= " & theta0 & " )"
    rstArray(2, 3) = Format(1 - WorksheetFunction.TDist(Abs(rstArray(2, 1)), n - 1, 1), "##0.0000")
    If rstArray(2, 1) < 0 Then rstArray(2, 3) = Format(WorksheetFunction.TDist(Abs(rstArray(2, 1)), n - 1, 1), "##0.0000")
    End Select
       
    TModulePrint.printRst title, rstArray
 
 
    If CI > 0 Then
    ReDim rstArray(2, 3)

    rstArray(1, 1) = CI & "% 신뢰구간": rstArray(1, 2) = "하한": rstArray(1, 3) = "상한"
        
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
    Dim Pvalue As Single, Pvalue2 As Single
    
    
    '''
    '''일반적으로 변수개수, 데이타개수, 데이타배열
    '''
If op = 1 Then                                              '표준입력
    
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

If op = 2 Then                                             '고급입력
    
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
    '''검정방법choice(1), 신뢰구간 CI , 귀무가설 Hyp
    '''
    CI = Format(choice(2), "##0.0000")
    Hyp = choice(3)
    
    '''
    ''' 제목 쓰기
    '''
    TModulePrint.Title1 "독립표본 t 검정"

        
    If n2(1) <= n2(2) Then
        'i = 1
        list_X = xlist2(2)                  '데이타 개수가 많은 것이 X
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
    
    TModulePrint.Title3 list_Y & "과 " & list_X & "의 기초통계량"
    ReDim rstArray(3, 4)
    rstArray(1, 1) = "변수명"
    rstArray(1, 2) = "개수"
    rstArray(1, 3) = "평균"
    rstArray(1, 4) = "표준편차"
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
    s_Xsq = Application.WorksheetFunction.Var(data_X)
    s_Ysq = Application.WorksheetFunction.Var(data_Y)
    spsq = ((num_Y - 1) * s_Ysq + (num_X - 1) * s_Xsq) / (num_Y + num_X - 2)
    DF2 = (s_Ysq / num_Y + s_Xsq / num_X) ^ 2 / ((s_Ysq / num_Y) ^ 2 / (num_Y - 1) + (s_Xsq / num_X) ^ 2 / (num_X - 1))
    df22 = Application.RoundDown(DF2, 0)
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
    tmp = DF2 - df22
    pvalueTuneq = (1 - tmp) * Application.WorksheetFunction.TDist(Abs(paTuneq), df22, 1) + tmp * Application.WorksheetFunction.TDist(Abs(paTuneq), df22 + 1, 1)
        
    Select Case choice(3)
    Case 1
        pvalueTeq = Application.WorksheetFunction.Min(pvalueTeq, 1 - pvalueTeq) * 2
        pvalueTuneq = Application.WorksheetFunction.Min(pvalueTuneq, 1 - pvalueTuneq) * 2
        title = " H0 : μ₁= μ₂vs. H1 : μ₁≠ μ₂     (μ₁: " & list_Y & ", μ₂: " & list_X & " )"
    Case 2
        If paTeq < 0 Then pvalueTeq = 1 - pvalueTeq
        If paTuneq < 0 Then pvalueTuneq = 1 - pvalueTuneq
        title = " H0 : μ₁= μ₂vs. H1 : μ₁> μ₂      (μ₁: " & list_Y & ", μ₂: " & list_X & " )"
    Case 3
        If paTeq > 0 Then pvalueTeq = 1 - pvalueTeq
        If paTuneq > 0 Then pvalueTuneq = 1 - pvalueTuneq
        title = " H0 : μ₁= μ₂vs. H2 : μ₁< μ₂      (μ₁: " & list_Y & ", μ₂: " & list_X & " )"
    End Select
    
    TT = "등분산 검정"
    ReDim rstArray(2, 3)
    rstArray(1, 1) = "자유도"
    rstArray(1, 2) = "F 통계량"
    rstArray(1, 3) = "유의확률"
    rstArray(2, 1) = "( " & dem_df - 1 & " , " & num_df - 1 & " )"
    rstArray(2, 2) = Format(fstat, "##0.0000")
    rstArray(2, 3) = Format(pvaluef, "##0.0000")
    
    
    TModulePrint.printRst TT, rstArray
    
   '======
   ' If pvaluef <= 0.01 Then
    '        Comment1 = """H0:두 표본의 분산들이 서로 같다.""" & "를 유의수준 α=0.01에서 기각한다."
     '
      '  ElseIf pvaluef <= 0.05 Then
       '     Comment1 = """H0:두 표본의 분산들이 서로 같다.""" & "를 유의수준 α=0.05에서 기각한다."
       '
      '  Else
      '      Comment1 = """H0:두 표본의 분산들이 서로 같다.""" & "를 유의수준 α=0.05에서 기각할 수 없다."
      '
      '  End If
    'Comment2 = "※유의확률이 유의수준보다 큰 경우에는 등분산 결과를 사용하는 것이 좋다. "
    
    
    
   ' Set mySheet = Worksheets(RstSheet)
    'mySheet.Activate ''''''''''''
   ' mySheet.Cells(1, 1) = mySheet.Cells(1, 1) + 1
    
    '    mySheet.Cells(1, 1) = mySheet.Cells(1, 1) + 2
     '   Flag = mySheet.Cells(1, 1).Value
     '   mySheet.Cells(Flag - 2, 2) = Comment1
     '   mySheet.Cells(Flag - 2, 2).Font.size = 10
     '   mySheet.Cells(Flag - 2, 2).Font.Bold = True
     '   mySheet.Cells(Flag - 2, 2).HorizontalAlignment = xlGeneral
     '   mySheet.Cells(Flag - 1, 2) = Comment2
     '   mySheet.Cells(Flag - 1, 2).Font.size = 10
     '   mySheet.Cells(Flag - 1, 2).Font.Bold = True
     '   mySheet.Cells(Flag - 1, 2).HorizontalAlignment = xlGeneral
    '=====
    
    TModulePrint.Title3 list_Y & "과 " & list_X & "의 가설검정"
    
    ReDim rstArray(3, 4)
    rstArray(1, 1) = "분산"
    rstArray(2, 1) = "등분산"
    rstArray(3, 1) = "이분산"
    
    rstArray(1, 2) = "t-통계량"
    rstArray(1, 3) = "자유도"
    rstArray(1, 4) = "유의확률"
    rstArray(2, 2) = Format(paTeq, "##0.0000")
    rstArray(2, 3) = num_Y + num_X - 2
    rstArray(2, 4) = Format(pvalueTeq, "##0.0000")
    rstArray(3, 2) = Format(paTuneq, "##0.0000")
    rstArray(3, 3) = Format(DF2, "##0.0000")
    rstArray(3, 4) = Format(pvalueTuneq, "##0.0000")
    
       Pvalue = rstArray(3, 4)
    TModulePrint.printRst title, rstArray
    
 
    TModulePrint.Title3 list_Y & "과 " & list_X & "의 신뢰구간"
    If CI > 0 Then
    ReDim rstArray(3, 3)

    rstArray(1, 1) = CI & "% 신뢰구간": rstArray(1, 2) = "하한": rstArray(1, 3) = "상한"
    rstArray(2, 1) = "등분산": rstArray(3, 1) = "이분산"
    
    Alpha = 1 - CI / 100
    Var = ((num_X - 1) * s_Xsq + (num_Y - 1) * s_Ysq) / (num_X + num_Y - 2) * (1 / num_X + 1 / num_Y)
    df = num_X + num_Y - 2
    tmp = cilength(Var, df, Alpha)
    rstArray(2, 2) = Format(avg_Y - avg_X - tmp, "##0.0000")
    rstArray(2, 3) = Format(avg_Y - avg_X + tmp, "##0.0000")

    Var = s_Xsq / num_X + s_Ysq / num_Y
    tmp = cilength(Var, DF2, Alpha)
    rstArray(3, 2) = Format(avg_Y - avg_X - tmp, "##0.0000")
    rstArray(3, 3) = Format(avg_Y - avg_X + tmp, "##0.0000")

    TModulePrint.printRst "", rstArray
    End If
    
       '''''해설
    TModulePrint.Title2 "요약 및 결론"

    ReDim rstArray(5, 8)
    
        
     rstArray(1, 1) = "유의수준 α=" & (100 - CI) / 100 & "이고 p-value가 " & Pvalue & "이다."
        If Pvalue < (100 - CI) / 100 Then
        rstArray(3, 1) = "유의확률(p-value)가 유의수준α= " & (100 - CI) / 100 & "보다 작기때문에 귀무가설H。을 기각한다."
        Else
        rstArray(3, 1) = "유의확률(p-value)가 유의수준α= " & (100 - CI) / 100 & "보다 크기때문에 귀무가설H。을 기각하지 못한다."
        End If

            rstArray(4, 1) = "※유의확률이 유의수준보다 큰 경우에는 등분산 결과를 사용하는 것이 좋다. "
    
    
  
    TModulePrint.printRstResult title, rstArray
    
    
End If

If choice(1) = 2 Then  '대응

    '''
    '''검정방법choice(1), 신뢰구간 CI , 귀무가설 Hyp
    '''
    CI = Format(choice(2), "##0.0000")
    Hyp = choice(3)
    
    '''
    ''' 제목 쓰기
    '''
    TModulePrint.Title1 "대응표본 t 검정"
        
    TModulePrint.Title3 xlist2(1) & "과 " & xlist2(2) & "의 기초통계량"
    ReDim rstArray(3, 4)
    
    rstArray(1, 1) = "변수명"
    rstArray(1, 2) = "개수"
    rstArray(1, 3) = "평균"
    rstArray(1, 4) = "표준편차"
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
    
    TModulePrint.Title3 xlist2(1) & "과 " & xlist2(2) & "의 가설검정"
    ReDim rstArray(2, 3)
    rstArray(1, 1) = "t-통계량": rstArray(1, 2) = "자유도": rstArray(1, 3) = "유의확률"
    rstArray(2, 1) = Format((WorksheetFunction.Average(dataArray)) / WorksheetFunction.StDev(dataArray) * Sqr(n2(1)), "##0.0000")
    rstArray(2, 2) = n2(1) - 1
    Pvalue2 = rstArray(2, 1)
    
    Select Case Hyp
    Case 1
    title = " H0 : μ₁= μ₂ vs. H1 : μ₁≠ μ₂      (μ₁: " & xlist2(1) & ", μ₂: " & xlist2(2) & " )"
    rstArray(2, 3) = Format(WorksheetFunction.TDist(Abs(rstArray(2, 1)), n2(1) - 1, 2), "##0.0000")
    
    Case 2
    title = " H0 : μ₁= μ₂ vs. H1 : μ₁> μ₂     (μ₁: " & xlist2(1) & ", μ₂: " & xlist2(2) & " )"
    rstArray(2, 3) = Format(WorksheetFunction.TDist(Abs(rstArray(2, 1)), n2(1) - 1, 1), "##0.0000")
    If rstArray(2, 1) < 0 Then rstArray(2, 3) = Format(1 - WorksheetFunction.TDist(Abs(rstArray(2, 1)), n2(1) - 1, 1), "##0.0000")
    
    Case 3
    title = " H0 : μ₁= μ₂ vs. H1 : μ₁< μ₂   (μ₁: " & xlist2(1) & ", μ₂: " & xlist2(2) & " )"
    rstArray(2, 3) = Format(1 - WorksheetFunction.TDist(Abs(rstArray(2, 1)), n2(1) - 1, 1), "##0.0000")
    If rstArray(2, 1) < 0 Then rstArray(2, 3) = Format(WorksheetFunction.TDist(Abs(rstArray(2, 1)), n2(1) - 1, 1), "##0.0000")
    End Select
       
    TModulePrint.printRst title, rstArray
 
    TModulePrint.Title3 xlist2(1) & "과 " & xlist2(2) & "의 신뢰구간"
    If CI > 0 Then
    ReDim rstArray(2, 3)

    rstArray(1, 1) = CI & "% 신뢰구간": rstArray(1, 2) = "하한": rstArray(1, 3) = "상한"
        
    tmp = WorksheetFunction.TInv(1 - CI / 100, n2(1) - 1) * WorksheetFunction.StDev(dataArray) / WorksheetFunction.Power(n2(1), 0.5)
    rstArray(2, 2) = Format(WorksheetFunction.Average(dataArray) - tmp, "##0.0000")
    rstArray(2, 3) = Format(WorksheetFunction.Average(dataArray) + tmp, "##0.0000")

    TModulePrint.printRst "", rstArray
    End If
    
       '''''해설
    TModulePrint.Title2 "요약 및 결론"

    ReDim rstArray(8, 8)
    
        
     rstArray(1, 1) = "유의수준 α=" & (100 - CI) / 100 & "이고 p-value가 " & Pvalue2 & "이다."
     If Pvalue2 < (100 - CI) / 100 Then
     rstArray(3, 1) = "유의확률(p-value)가 유의수준α= " & (100 - CI) / 100 & "보다 작기때문에 귀무가설H。을 기각한다."
     Else
     rstArray(3, 1) = "유의확률(p-value)가 유의수준α= " & (100 - CI) / 100 & "보다 크기때문에 귀무가설H。을 기각하지 못한다."
     End If

          If Pvalue2 < (100 - CI) / 100 Then
   rstArray(5, 1) = CI & "% 신뢰구간은" & "(" & Format(WorksheetFunction.Average(dataArray) - tmp, "##0.0000") & "," & Format(WorksheetFunction.Average(dataArray) + tmp, "##0.0000") & ") 이다."
   rstArray(6, 1) = "μ= " & Format(Application.WorksheetFunction.Average(dataArray1), "##0.0000") - Format(Application.WorksheetFunction.Average(dataArray2), "##0.0000") & "가 " & CI & "% 신뢰구간에 속하지 않으므로 유의수준" & 100 - CI & "%에서 귀무가설을 기각한다."
       Else
   rstArray(5, 1) = CI & "% 신뢰구간은" & "(" & Format(WorksheetFunction.Average(dataArray) - tmp, "##0.0000") & "," & Format(WorksheetFunction.Average(dataArray) + tmp, "##0.0000") & ") 이다."
   rstArray(6, 1) = "μ= " & Format(Application.WorksheetFunction.Average(dataArray1), "##0.0000") - Format(Application.WorksheetFunction.Average(dataArray2), "##0.0000") & "가 " & CI & "% 신뢰구간에 속하므로 유의수준" & 100 - CI & "%에서 귀무가설을 기각하지 못한다."
       End If
    
  
    TModulePrint.printRstResult title, rstArray

End If
    
End Sub
Sub FTest1(choice, op)
    
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
    '''일반적으로 변수개수, 데이타개수, 데이타배열
    '''
If op = 1 Then                                              '표준입력
    
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


If choice(1) = 1 Then
    '''
    '''검정방법choice(1), 신뢰구간 CI , 귀무가설 Hyp
    '''
    CI = Format(choice(2), "##0.0000")
    Hyp = choice(3)
    
    '''
    ''' 제목 쓰기
    '''
    TModulePrint.Title1 "두 모분산σ²에 대한 F-검정"

    If n2(1) <= n2(2) Then
        'i = 1
        list_X = xlist2(2)                  '데이타 개수가 많은 것이 X
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
    TModulePrint.Title3 list_Y & "과 " & list_X & "의 기초통계량"
    ReDim rstArray(3, 4)
    
    rstArray(1, 1) = "변수명"
    rstArray(1, 2) = "개수"
    rstArray(1, 3) = "평균"
    rstArray(1, 4) = "표준편차"
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
    s_Xsq = Application.WorksheetFunction.Var(data_X)
    s_Ysq = Application.WorksheetFunction.Var(data_Y)
    spsq = ((num_Y - 1) * s_Ysq + (num_X - 1) * s_Xsq) / (num_Y + num_X - 2)
    DF2 = (s_Ysq / num_Y + s_Xsq / num_X) ^ 2 / ((s_Ysq / num_Y) ^ 2 / (num_Y - 1) + (s_Xsq / num_X) ^ 2 / (num_X - 1))
    df22 = Application.RoundDown(DF2, 0)
    paTeq = (avg_Y - avg_X) / Sqr(spsq) / Sqr(1 / num_Y + 1 / num_X)
    paTuneq = (avg_Y - avg_X) / Sqr(s_Ysq / num_Y + s_Xsq / num_X)
    
   
    
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
    tmp = DF2 - df22
    pvalueTuneq = (1 - tmp) * Application.WorksheetFunction.TDist(Abs(paTuneq), df22, 1) + tmp * Application.WorksheetFunction.TDist(Abs(paTuneq), df22 + 1, 1)
        
    Select Case choice(3)
    Case 1
        pvalueTeq = Application.WorksheetFunction.Min(pvalueTeq, 1 - pvalueTeq) * 2
        pvalueTuneq = Application.WorksheetFunction.Min(pvalueTuneq, 1 - pvalueTuneq) * 2
        title = " H0 : σ₁= σ₂vs. H1 : σ₁≠ σ₂     (σ₁: " & list_Y & ", σ₂: " & list_X & " )"
    Case 2
        If paTeq < 0 Then pvalueTeq = 1 - pvalueTeq
        If paTuneq < 0 Then pvalueTuneq = 1 - pvalueTuneq
        title = " H0 : σ₁= σ₂vs. H1 : σ₁> σ₂      (σ₁: " & list_Y & ", σ₂: " & list_X & " )"
    Case 3
        If paTeq > 0 Then pvalueTeq = 1 - pvalueTeq
        If paTuneq > 0 Then pvalueTuneq = 1 - pvalueTuneq
        title = " H0 : σ₁= σ₂vs. H1 : σ₁< σ₂      (σ₁: " & list_Y & ", σ₂: " & list_X & " )"
    End Select
    
    TModulePrint.Title3 list_Y & "과 " & list_X & "의 가설검정"
    ReDim rstArray(2, 3)
    rstArray(1, 1) = "자유도"
    rstArray(1, 2) = "F - 통계량"
    rstArray(1, 3) = "유의확률"
    rstArray(2, 1) = "( " & dem_df - 1 & " , " & num_df - 1 & " )"
    rstArray(2, 2) = Format(fstat, "##0.0000")
    rstArray(2, 3) = Format(pvaluef, "##0.0000")
    TModulePrint.printRstOne title, rstArray
    
 
  
    
    TModulePrint.Title3 list_Y & "과 " & list_X & "의 신뢰구간"
    If CI > 0 Then
    ReDim rstArray(2, 3)

    rstArray(1, 1) = CI & "% 신뢰구간": rstArray(1, 2) = "하한": rstArray(1, 3) = "상한"
    rstArray(2, 1) = " "
    
    Alpha = (100 - CI) / 100 '유의수준
    Var = WorksheetFunction.FInv((Alpha / 2), num_X, num_Y) 'f-2검정
    df = WorksheetFunction.FInv((Alpha / 2), num_Y, num_X) 'f-1검정
    
   ' tmp = cilength(Var, df, Alpha)
    
    rstArray(2, 2) = Format((s_Ysq / s_Xsq) * (1 / df), "##0.0000")
    rstArray(2, 3) = Format((s_Ysq / s_Xsq) * Var, "##0.0000")

    
    TModulePrint.printRst "", rstArray
    End If
    
       '''''해설
    TModulePrint.Title2 "요약 및 결론"

    ReDim rstArray(8, 8)
    
        
    rstArray(1, 1) = "유의수준 α=" & (100 - CI) / 100 & "이고 p-value가 " & Format(pvaluef, "##0.0000") & "이다."
       If pvaluef < (100 - CI) / 100 Then
       rstArray(4, 1) = "유의확률(p-value)가 유의수준α= " & (100 - CI) / 100 & "보다 작기때문에 귀무가설H。을 기각한다."
       Else
       rstArray(4, 1) = "유의확률(p-value)가 유의수준α= " & (100 - CI) / 100 & "보다 크기때문에 귀무가설H。을 기각하지 못한다."
       End If
    
   ' rstArray(5, 7) = "귀무가설H。하에서 검정통계량의 관측값은" & Zvalue & "이고"
    If pvaluef < (100 - CI) / 100 Then
        rstArray(6, 1) = CI & "% 신뢰구간은" & "(" & Format((s_Ysq / s_Xsq) * (1 / df), "##0.0000") & "," & Format((s_Ysq / s_Xsq) * Var, "##0.0000") & ") 이다."
        rstArray(7, 1) = "σ₁/ σ₂= " & s_Ysq / s_Xsq & "가 " & CI & "% 신뢰구간에 속하지 않으므로 유의수준" & 100 - CI & "%에서 귀무가설을 기각한다."
    Else
    rstArray(6, 1) = CI & "% 신뢰구간은" & "(" & Format((s_Ysq / s_Xsq) * (1 / df), "##0.0000") & "," & Format((s_Ysq / s_Xsq) * Var, "##0.0000") & ") 이다."
    rstArray(7, 1) = "σ₁/ σ₂= " & s_Ysq / s_Xsq & "가 " & CI & "% 신뢰구간에 속하므로 유의수준" & 100 - CI & "%에서 귀무가설을 기각하지 못한다."
    End If
        
    rstArray(8, 1) = "※유의확률이 유의수준보다 큰 경우에는 등분산 결과를 사용하는 것이 좋다. "
    TModulePrint.printRstResult title, rstArray
    
End If
   
End Sub


Function msgWord(word) As String

    If Len(word) > 9 Then
        msgWord = Mid(word, 1, 6) & vbLf & Mid(word, 7)
    Else
        msgWord = word
    End If
    
End Function

Function cilength(Var, df, Alpha)
    If df = Application.RoundDown(df, 0) Then
       cilength = Application.TInv(Alpha, df) * Application.Power(Var, 0.5)
    Else: df_d = Application.RoundDown(df, 0)
       t_d = Application.TInv(Alpha, df_d)
       t_u = Application.TInv(Alpha, df_d + 1)
       W = df - df_d
       t_value = (1 - W) * t_d + W * t_u
       cilength = t_value * Application.Power(Var, 0.5)
    End If
End Function

'''
''' xlist, n, k1 : public 변수 선언되어 있음.
'''
Sub TTestR(choice)
    
    Dim dataArray(), rstArray()
    Dim theta0 As Single, CI As Single, Hyp As Integer
    Dim mySheet As Worksheet
    Dim titleTmp() As String, title As String

    
    '''
    '''데이타 배열
    '''
    Worksheets(DataSheet).Activate
    dataArray = Worksheets(DataSheet).Range(Cells(2, k1), Cells(n + 1, k1)).Value
    
    '''
    '''검정값 theta0, 신뢰구간 CI , 귀무가설 Hyp
    '''
    theta0 = Format(choice(1), "##0.0000")
    CI = Format(choice(2), "##0.0000")
    Hyp = choice(3)
    
    '''
    ''' 결과시트 활성화
    '''
    Set mySheet = Worksheets(RstSheet)
    mySheet.Activate
    
    '''
    ''' 제목 쓰기
    '''
    TModulePrint.Title1 "모평균μ에 대한 t-검정"
    TModulePrint.Title3 xlist & "의 기초통계량"



    '''
    ''' 계산하고 출력하기
    '''
    Dim Tvalue
    Dim Pvalue
    Tvalue = Format((WorksheetFunction.Average(dataArray) - theta0) / WorksheetFunction.StDev(dataArray) * Sqr(n), "##0.0000")
    
    ReDim rstArray(2, 4)
    rstArray(1, 1) = "변수명": rstArray(1, 2) = "표본크기": rstArray(1, 3) = "평균": rstArray(1, 4) = "표준편차"
    rstArray(2, 1) = xlist: rstArray(2, 2) = n
    rstArray(2, 3) = Format(WorksheetFunction.Average(dataArray), "##0.0000")
    rstArray(2, 4) = Format(WorksheetFunction.StDev(dataArray), "##0.0000")
    
    TModulePrint.printRst "", rstArray
     
    TModulePrint.Title3 xlist & "의 가설검정"
    
    ReDim rstArray(2, 4)
    rstArray(1, 1) = "자유도": rstArray(1, 2) = "목표값": rstArray(1, 3) = "t-통계량": rstArray(1, 4) = "유의확률"
    rstArray(2, 1) = n - 1
    rstArray(2, 2) = theta0
    rstArray(2, 3) = Format((WorksheetFunction.Average(dataArray) - theta0) / WorksheetFunction.StDev(dataArray) * Sqr(n), "##0.0000")

    
    

    Select Case Hyp
    Case 1
    title = " H : μ=μ。vs. K : μ≠μ。     (μ。= " & theta0 & " )"
    rstArray(2, 4) = Format(WorksheetFunction.TDist(Abs(rstArray(2, 3)), n - 1, 2), "##0.0000")
    
    Case 2
    title = " H : μ=μ。vs. K : μ > μ。     (μ。= " & theta0 & " )"
    rstArray(2, 4) = Format(WorksheetFunction.TDist(Abs(rstArray(2, 3)), n - 1, 1), "##0.0000")
    If rstArray(2, 1) < 0 Then rstArray(2, 4) = Format(1 - WorksheetFunction.TDist(Abs(rstArray(2, 3)), n - 1, 1), "##0.0000")
    
    Case 3
    title = " H : μ=μ。vs. K : μ < μ。     (μ。= " & theta0 & " )"
    rstArray(2, 4) = Format(1 - WorksheetFunction.TDist(Abs(rstArray(2, 3)), n - 1, 1), "##0.0000")
    If rstArray(2, 3) < 0 Then rstArray(2, 4) = Format(WorksheetFunction.TDist(Abs(rstArray(2, 3)), n - 1, 1), "##0.0000")
    End Select
          Pvalue = rstArray(2, 4)
    TModulePrint.printRst title, rstArray

    TModulePrint.Title3 xlist & "의 신뢰구간 그래프"
    
    If CI > 0 Then
    ReDim rstArray(2, 3)

    rstArray(1, 1) = CI & "% 신뢰구간": rstArray(1, 2) = "하한": rstArray(1, 3) = "상한"
        
    tmp = WorksheetFunction.TInv(1 - CI / 100, n - 1) * WorksheetFunction.StDev(dataArray) / WorksheetFunction.Power(n, 0.5)
    rstArray(2, 2) = Format(WorksheetFunction.Average(dataArray) - tmp, "##0.0000")
    rstArray(2, 3) = Format(WorksheetFunction.Average(dataArray) + tmp, "##0.0000")

    TModulePrint.printRst "", rstArray
    End If

     TModuleGraph.Ciplot '신뢰구간그래프
    
       '''''해설
    TModulePrint.Title2 "요약 및 결론"

    ReDim rstArray(8, 8)
    
    If n > 30 Then
        rstArray(1, 1) = "n = " & n & "으로, 30 이상이므로 z-검정을 실시해야한다."
    Else
        rstArray(1, 1) = "n = " & n & "이고, 모분산을 모르기 때문에 t-검정을 실시한다."
    
        rstArray(3, 1) = "유의수준 α=" & (100 - CI) / 100 & "이고 p-value가 " & Pvalue & "이다."
        If Pvalue < (100 - CI) / 100 Then
        rstArray(4, 1) = "유의확률(p-value)가 유의수준α= " & (100 - CI) / 100 & "보다 작기때문에 귀무가설H。을 기각한다."
        Else
        rstArray(4, 1) = "유의확률(p-value)가 유의수준α= " & (100 - CI) / 100 & "보다 크기때문에 귀무가설H。을 기각하지 못한다."
        End If
    
   ' rstArray(5, 7) = "귀무가설H。하에서 검정통계량의 관측값은" & Tvalue & "이고"
      If Pvalue < (100 - CI) / 100 Then
   rstArray(6, 1) = CI & "% 신뢰구간은" & "(" & Format(WorksheetFunction.Average(dataArray) - tmp, "##0.0000") & "," & Format(WorksheetFunction.Average(dataArray) + tmp, "##0.0000") & ") 이다."
  ' rstArray(7, 1) = "μ= " & theta0 & "가 " & CI & "% 신뢰구간에 속하지 않으므로 유의수준" & 100 - CI & "%에서 귀무가설을 기각한다."
       Else
   rstArray(6, 1) = CI & "% 신뢰구간은" & "(" & Format(WorksheetFunction.Average(dataArray) - tmp, "##0.0000") & "," & Format(WorksheetFunction.Average(dataArray) + tmp, "##0.0000") & ") 이다."
  ' rstArray(7, 1) = "μ= " & theta0 & "가 " & CI & "% 신뢰구간에 속하므로 유의수준" & 100 - CI & "%에서 귀무가설을 기각하지 못한다."
       End If
     
    End If
    TModulePrint.printRstResult title, rstArray
    'TModulePrint.printRstResult title, rstArray

End Sub
'박스플롯
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
    nRow = rngData.Rows.Count
    
'   Resize data
    ReDim arrTlt(nGrp)

            arrTlt(1) = frameOneTtest.ListBox2.List(0)
           ' arrTlt(2) = "목표 평균"
            

'''
''' Err Check           한개만 있어도 됌 나머지 하나는 평균값으로 계산
'''

  '  If nGrp < 2 Then
 '       MsgBox "2개 이상의 그룹을 선택해야 합니다."
 '       Exit Sub
  '  End If
    
    ReDim rngTmp(nGrp)
    For i = 1 To nGrp
        Set rngTmp(i) = rngData.Columns(i)
    Next i

'   Add a Sheet
    For i = 1 To Sheets.Count
        If Sheets(i).Name = "_통계분석결과_" Then
            GoTo 31
        Else
            GoTo 32
        End If
        
32: Next i
    Worksheets.Add Before:=Worksheets(1)
    ActiveSheet.Name = "_통계분석결과_"
    ActiveWindow.DisplayGridlines = False
    Cells(1, 1) = 1

31: Sheets("_통계분석결과_").Activate
    Application.ScreenUpdating = False

'   Current worksheet's name
    strName = ActiveSheet.Name

'   Chart Location
    Set rngFirst = Cells(Cells(1, 1), 1)

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

'''OneZ 계산완료

'''
''' xlist, n, k1 : public 변수 선언되어 있음.
'''
Sub ZTest1(choice)
    
    Dim dataArray(), rstArray()
    Dim theta0 As Single, CI As Single, Hyp As Integer, zSD As Single
    Dim mySheet As Worksheet
    Dim titleTmp() As String, title As String
    Dim posi(0 To 1) As Long '그래프설정으로 추가
    Dim Pvalue As Single, Zvalue As Single
    
    
    '''
    '''데이타 배열
    '''
    Worksheets(DataSheet).Activate
    dataArray = Worksheets(DataSheet).Range(Cells(2, k1), Cells(n + 1, k1)).Value
    
    '''
    '''검정값 theta0, 신뢰구간 CI , 귀무가설 Hyp
    '''
    theta0 = Format(choice(1), "##0.0000")
    CI = choice(2)
    Hyp = choice(3)
    zSD = choice(4)
    
    
    
    '''
    ''' 결과시트 활성화
    '''
    Set mySheet = Worksheets(RstSheet)
    mySheet.Activate
    
    '''
    ''' 제목 쓰기
    '''
    TModulePrint.Title1 "모평균μ에 대한 Z-검정 분석결과"
    TModulePrint.Title3 xlist & "의 기초통계량"


    '''
    ''' 계산하고 출력하기
    '''
    ReDim rstArray(2, 4)
    rstArray(1, 1) = "변수명": rstArray(1, 2) = "개수": rstArray(1, 3) = "평균": rstArray(1, 4) = "표준편차"
    rstArray(2, 1) = xlist: rstArray(2, 2) = n
    rstArray(2, 3) = Format(WorksheetFunction.Average(dataArray), "##0.0000")
    rstArray(2, 4) = Format(WorksheetFunction.StDev(dataArray), "##0.0000")
    TModulePrint.printRst "", rstArray
    
    
    TModulePrint.Title3 xlist & "의 데이터 분포에 대한 그래프"
    
    
    '그래프 출력 해석
    ReDim rstArray(5, 3)
    rstArray(1, 1) = xlist & "의 데이터 분포 그래프"
    rstArray(2, 1) = xlist & "의 데이터 분포 최소값"
    rstArray(2, 3) = WorksheetFunction.Min(dataArray)
    rstArray(3, 1) = xlist & "의 데이터 분포 최대값"
    rstArray(3, 3) = WorksheetFunction.Max(dataArray)
     
    
    TModulePrint.printRstOne title, rstArray
     '데이터의 분포 그래프출력
    'TModuleGraph.Ciplot
    boxgraph1
    
   '   TModulePrint.printRst "", rstArray
     
   

    '수정해야함 z통계량 아웃풋
    
    TModulePrint.Title3 xlist & "의 가설검정"

    ReDim rstArray(2, 3)
    rstArray(1, 1) = "모표준편차"
    rstArray(1, 2) = "z-통계량": rstArray(1, 3) = "유의확률"
    rstArray(2, 1) = zSD
    rstArray(2, 2) = Format((WorksheetFunction.Average(dataArray) - theta0) / zSD * Sqr(n), "##0.0000")
    rstArray(2, 3) = Format(1 - (WorksheetFunction.Erf(Abs(rstArray(2, 1)) / Sqr(2))), "##0.0000")
    
    Zvalue = Format((WorksheetFunction.Average(dataArray) - theta0) / zSD * Sqr(n), "##0.0000")
    Pvalue = Format(1 - (WorksheetFunction.Erf(Abs(rstArray(2, 1)) / Sqr(2))), "##0.0000")
    
    Select Case Hyp
    Case 1
    title = " H0 : μ=μ。vs. H1 : μ≠μ。     (μ。= " & theta0 & " )"
    rstArray(2, 2) = Format((WorksheetFunction.Average(dataArray) - theta0) / zSD * Sqr(n), "##0.0000")
    
    Case 2
    title = " H0 : μ=μ。vs. H1 : μ > μ。     (μ。= " & theta0 & " )"
    rstArray(2, 2) = Format((WorksheetFunction.Average(dataArray) - theta0) / zSD * Sqr(n), "##0.0000")
    If rstArray(2, 2) < 0 Then rstArray(2, 3) = Format(1 - (WorksheetFunction.Erf(Abs(rstArray(2, 1)) / Sqr(2))), "##0.0000")
    
    Case 3
    title = " H0 : μ=μ。vs. H1 : μ < μ。     (μ。= " & theta0 & " )"
     rstArray(2, 2) = Format((WorksheetFunction.Average(dataArray) - theta0) / zSD * Sqr(n), "##0.0000")
   If rstArray(2, 2) < 0 Then rstArray(2, 3) = Format(1 - (WorksheetFunction.Erf(Abs(rstArray(2, 1)) / Sqr(2))), "##0.0000")
    End Select
    
      ' If rstArray(2, 3) > 0.05 Then
      ' rstArray(4, 1) = "유의확률 a =  " & rstArray(2, 3) & " 으로 0.05 이상이므로 귀무가설H。을 기각하지 못함"
      ' Else
     '  rstArray(4, 1) = "유의확률 a =  " & rstArray(2, 3) & " 으로 0.05 이하이므로 귀무가설H。을 기각하고 대립가설H₁을 채택함"
      ' End If
       
    TModulePrint.printRstOne title, rstArray
 
    TModulePrint.Title3 xlist & "의 신뢰구간"
    If CI > 0 Then
    ReDim rstArray(2, 3)

    rstArray(1, 1) = CI & "%신뢰구간": rstArray(1, 2) = CI & "%하한": rstArray(1, 3) = CI & "%상한"
        
    tmp = WorksheetFunction.NormSInv((1 - (CI / 100)) / 2) * (zSD / Sqr(n))
   
    rstArray(2, 2) = Format(WorksheetFunction.Average(dataArray) + tmp, "##0.0000")
    rstArray(2, 3) = Format(WorksheetFunction.Average(dataArray) - tmp, "##0.0000")

    TModulePrint.printRst "", rstArray
   '''''해설
    TModulePrint.Title2 "요약 및 결론"

    ReDim rstArray(8, 8)
    
    If n > 30 Then
        rstArray(1, 1) = "n = " & n & "로 충분히 크므로 검정통계량은 Z 검정을 따르고, Z는 근사적으로 정규분포를 따른다."
    Else
        rstArray(1, 1) = "n = " & n & "로 충분히 크지 않은 경우에 모평균에 대한 가설검정을 위해서는 모집단의 분포가 정규분포라는 가정이 추가로 필요하게 된다."
    End If
    
     rstArray(3, 1) = "유의수준 α=" & (100 - CI) / 100 & "이고 p-value가 " & Pvalue & "이다."
        If Pvalue < (100 - CI) / 100 Then
        rstArray(4, 1) = "유의확률(p-value)가 유의수준α= " & (100 - CI) / 100 & "보다 작기때문에 귀무가설H。을 기각한다."
        Else
        rstArray(4, 1) = "유의확률(p-value)가 유의수준α= " & (100 - CI) / 100 & "보다 크기때문에 귀무가설H。을 기각하지 못한다."
        End If
    
   ' rstArray(5, 7) = "귀무가설H。하에서 검정통계량의 관측값은" & Zvalue & "이고"
      If Pvalue < (100 - CI) / 100 Then
   rstArray(6, 1) = CI & "% 신뢰구간은" & "(" & Format(WorksheetFunction.Average(dataArray) + tmp, "##0.0000") & "," & Format(WorksheetFunction.Average(dataArray) - tmp, "##0.0000") & ") 이다."
   rstArray(7, 1) = "μ= " & theta0 & "가 " & CI & "% 신뢰구간에 속하지 않으므로 유의수준" & 100 - CI & "%에서 귀무가설을 기각한다."
       Else
   rstArray(6, 1) = CI & "% 신뢰구간은" & "(" & Format(WorksheetFunction.Average(dataArray) + tmp, "##0.0000") & "," & Format(WorksheetFunction.Average(dataArray) - tmp, "##0.0000") & ") 이다."
   rstArray(7, 1) = "μ= " & theta0 & "가 " & CI & "% 신뢰구간에 속하므로 유의수준" & 100 - CI & "%에서 귀무가설을 기각하지 못한다."
       End If
       
    TModulePrint.printRstResult title, rstArray
   End If
    
End Sub

'''''OneZ 데이터분포그래프 완료

Private Sub boxgraph1()

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
    nRow = rngData.Rows.Count
    
'   Resize data
    ReDim arrTlt(nGrp)

            arrTlt(1) = frameOneZ.ListBox2.List(0)
           ' arrTlt(2) = "목표 평균"
            

'''
''' Err Check           한개만 있어도 됌 나머지 하나는 평균값으로 계산
'''

  '  If nGrp < 2 Then
 '       MsgBox "2개 이상의 그룹을 선택해야 합니다."
 '       Exit Sub
  '  End If
    
    ReDim rngTmp(nGrp)
    For i = 1 To nGrp
        Set rngTmp(i) = rngData.Columns(i)
    Next i

'   Add a Sheet
    For i = 1 To Sheets.Count
        If Sheets(i).Name = "_통계분석결과_" Then
            GoTo 31
        Else
            GoTo 32
        End If
        
32: Next i
    Worksheets.Add Before:=Worksheets(1)
    ActiveSheet.Name = "_통계분석결과_"
    ActiveWindow.DisplayGridlines = False
    Cells(1, 1) = 1

31: Sheets("_통계분석결과_").Activate
    Application.ScreenUpdating = False

'   Current worksheet's name
    strName = ActiveSheet.Name

'   Chart Location
    Set rngFirst = Cells(Cells(1, 1) - 2, 1)

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
    nPage = 15
   ' rngFirst.Offset = "Created at " & Now()
    Application.Goto rngFirst, Scroll:=True
    Range("A1") = Range("A1") + nPage

    Application.ScreenUpdating = True
   ' Unload Me

End Sub

Sub ZTest2(choice, op)
    
    Dim dataArray1(), dataArray2(), data_X(), data_Y(), dataArray()
    Dim CI As Single, Hyp As Integer, zSD1 As Single, zSD2 As Single
    Dim mySheet As Worksheet
    Dim titleTmp() As String
    Dim gpChk(2)
    Dim list_X As String, list_Y As String
    
    Dim Zvalue As Single
    Dim Pvalue As Single
    

    Dim avg_X As Single, avg_Y As Single, s_Xsq As Single, s_Ysq As Single
    Dim paTeq As Single, paTuneq As Single
    Dim pvalueTeq As Single, pvalueTuneq As Single
    Dim num_df, dem_df As Integer
    '''
    '''일반적으로 변수개수, 데이타개수, 데이타배열
    '''
If op = 1 Then                                              '표준입력
    
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


If choice(1) = 1 Then
    '''
    '''검정방법choice(1), 신뢰구간 CI , 귀무가설 Hyp
    '''
    CI = Format(choice(2), "##0.0000")
    Hyp = choice(3)
    zSD1 = choice(4)
    zSD2 = choice(5)
    '''
    ''' 제목 쓰기
    '''
    TModulePrint.Title1 "두 모집단μ에 대한 z-검정"

        
    If n2(1) <= n2(2) Then
        'i = 1
        list_X = xlist2(2)                  '데이타 개수가 많은 것이 X
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
    TModulePrint.Title3 list_Y & "과 " & list_X & "의 기초통계량"
    ReDim rstArray(3, 5)
    
    rstArray(1, 1) = "변수명"
    rstArray(1, 2) = "개수"
    rstArray(1, 3) = "평균"
    rstArray(1, 4) = "표준편차"
    rstArray(1, 5) = "모표준편차"
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
    
    rstArray(2, 5) = zSD1
    rstArray(3, 5) = zSD2
    
    TModulePrint.printRst "", rstArray

    
  
    avg_X = Application.WorksheetFunction.Average(data_X)
    avg_Y = Application.WorksheetFunction.Average(data_Y)
    s_Xsq = Application.WorksheetFunction.Var(data_X)
    s_Ysq = Application.WorksheetFunction.Var(data_Y)
    
    
    spsq = ((num_Y - 1) * s_Ysq + (num_X - 1) * s_Xsq) / (num_Y + num_X - 2)
    DF2 = (s_Ysq / num_Y + s_Xsq / num_X) ^ 2 / ((s_Ysq / num_Y) ^ 2 / (num_Y - 1) + (s_Xsq / num_X) ^ 2 / (num_X - 1))
    df22 = Application.RoundDown(DF2, 0)
    paTeq = (avg_Y - avg_X) / Sqr(spsq) / Sqr(1 / num_Y + 1 / num_X)
    paTuneq = (avg_Y - avg_X) / Sqr(s_Ysq / num_Y + s_Xsq / num_X)

    Zvalue = Format((avg_X - avg_Y) / Sqr(((s_Xsq / num_X) + (s_Ysq / num_Y))), "##0.0000")
    
    'fstat = WorksheetFunction.Max(s_Ysq, s_Xsq) / WorksheetFunction.Min(s_Ysq, s_Xsq)
    If WorksheetFunction.Max(s_Ysq, s_Xsq) = s_Ysq Then
    dem_df = num_Y
    num_df = num_X
    Else
    dem_df = num_X
    num_df = num_Y
    End If
    
   ' pvaluef = 2 * Application.FDist(fstat, dem_df - 1, num_df - 1)
    
    'pvalueTeq = Application.WorksheetFunction.TDist(Abs(paTeq), num_Y + num_X - 2, 1)
    'pvalueTuneq = Application.WorksheetFunction.TDist(Abs(paTuneq), df22, 1)
    'tmp = DF2 - df22
    'pvalueTuneq = (1 - tmp) * Application.WorksheetFunction.TDist(Abs(paTuneq), df22, 1) + tmp * Application.WorksheetFunction.TDist(Abs(paTuneq), df22 + 1, 1)
    pvaluez = Format(1 - (WorksheetFunction.Erf(Abs(rstArray(2, 2)) / Sqr(2))), "##0.0000")
    
    TModulePrint.Title3 list_Y & "과 " & list_X & "의 가설검정"
   
     ReDim rstArray(2, 3)
    rstArray(1, 1) = "자유도"
    rstArray(1, 2) = "z-통계량"
    rstArray(1, 3) = "유의확률"
    rstArray(2, 1) = "( " & dem_df & " , " & num_df & " )"
    rstArray(2, 2) = Format(Zvalue, "##0.0000")
    rstArray(2, 3) = Format(1 - (WorksheetFunction.Erf(Abs(rstArray(2, 2)) / Sqr(2))), "##0.0000")
   ' Comment2 = "※유의확률이 유의수준보다 큰 경우에는 등분산 결과를 사용하는 것이 좋다. "
    
    Pvalue = Format(1 - (WorksheetFunction.Erf(Abs(rstArray(2, 2)) / Sqr(2))), "##0.0000")
    
    Select Case choice(3)
    Case 1
     title = " H0 : μ₁= μ₂vs. H1 : μ₁≠ μ₂     (μ₁: " & list_Y & ", μ₂: " & list_X & " )"
    rstArray(2, 2) = Format(Zvalue, "##0.0000")
    
    Case 2
    title = " H0 : μ₁= μ₂vs. H1 : μ₁> μ₂      (μ₁: " & list_Y & ", μ₂: " & list_X & " )"
    rstArray(2, 2) = Format(Zvalue, "##0.0000")
    If rstArray(2, 2) < 0 Then rstArray(2, 3) = Format(1 - (WorksheetFunction.Erf(Abs(rstArray(2, 2)) / Sqr(2))), "##0.0000")
    
    Case 3
     title = " H0 : μ₁= μ₂vs. H1 : μ₁< μ₂      (μ₁: " & list_Y & ", μ₂: " & list_X & " )"
     rstArray(2, 2) = Format(Zvalue, "##0.0000")
     If rstArray(2, 2) < 0 Then rstArray(2, 3) = Format(1 - (WorksheetFunction.Erf(Abs(rstArray(2, 2)) / Sqr(2))), "##0.0000")
    End Select
     
    
  
    TModulePrint.printRstOne title, rstArray


    TModulePrint.Title3 list_Y & "과 " & list_X & "의 신뢰구간"
    
    '신뢰구간
    If CI > 0 Then
    ReDim rstArray(2, 3)

    rstArray(1, 1) = CI & "% 신뢰구간": rstArray(1, 2) = CI & "% 하한": rstArray(1, 3) = CI & "% 상한"
    rstArray(2, 1) = " "
    
    Alpha = 1 - CI / 100 ' 유의수준
   ' Var = ((num_X - 1) * s_Xsq + (num_Y - 1) * s_Ysq) / (num_X + num_Y - 2) * (1 / num_X + 1 / num_Y)
   ' df = num_X + num_Y - 2
    tmp = WorksheetFunction.NormSInv((1 - (CI / 100)) / 2) * Sqr(((s_Xsq / num_X) + (s_Ysq / num_Y)))
    
    rstArray(2, 2) = Format(avg_Y - avg_X + tmp, "##0.0000")
    rstArray(2, 3) = Format(avg_Y - avg_X - tmp, "##0.0000")
    TModulePrint.printRst "", rstArray
    End If
    
       '''''해설
    TModulePrint.Title2 "요약 및 결론"

    ReDim rstArray(8, 8)
    
        
     rstArray(1, 1) = "유의수준 α=" & (100 - CI) / 100 & "이고 p-value가 " & Pvalue & "이다."
        If Pvalue < (100 - CI) / 100 Then
        rstArray(4, 1) = "유의확률(p-value)가 유의수준α= " & (100 - CI) / 100 & "보다 작기때문에 귀무가설H。을 기각한다."
        Else
        rstArray(4, 1) = "유의확률(p-value)가 유의수준α= " & (100 - CI) / 100 & "보다 크기때문에 귀무가설H。을 기각하지 못한다."
        End If
    
   ' rstArray(5, 7) = "귀무가설H。하에서 검정통계량의 관측값은" & Zvalue & "이고"
      If Pvalue < (100 - CI) / 100 Then
   rstArray(6, 1) = CI & "% 신뢰구간은" & "(" & Format(avg_Y - avg_X + tmp, "##0.0000") & "," & Format(avg_Y - avg_X - tmp, "##0.0000") & ") 이다."
   rstArray(7, 1) = "μ= " & avg_Y - avg_X & "가 " & CI & "% 신뢰구간에 속하지 않으므로 유의수준" & 100 - CI & "%에서 귀무가설을 기각한다."
       Else
   rstArray(6, 1) = CI & "% 신뢰구간은" & "(" & Format(avg_Y - avg_X + tmp, "##0.0000") & "," & Format(avg_Y - avg_X - tmp, "##0.0000") & ") 이다."
   rstArray(7, 1) = "μ= " & avg_Y - avg_X & "가 " & CI & "% 신뢰구간에 속하므로 유의수준" & 100 - CI & "%에서 귀무가설을 기각하지 못한다."
       End If
       
    TModulePrint.printRstResult title, rstArray
   End If
    

End Sub
Sub analle(choice, op)
    
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
    '''일반적으로 변수개수, 데이타개수, 데이타배열
    '''
If op = 1 Then                                              '표준입력
    
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


If choice(1) = 1 Then
    '''
    '''검정방법choice(1), 신뢰구간 CI , 귀무가설 Hyp
    '''
    CI = Format(choice(2), "##0.0000")
    Hyp = choice(3)
    
    '''
    ''' 제목 쓰기
    '''
    TModulePrint.Title1 "등분산검정"

    If n2(1) <= n2(2) Then
        'i = 1
        list_X = xlist2(2)                  '데이타 개수가 많은 것이 X
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

    
    
    avg_X = Application.WorksheetFunction.Average(data_X)
    avg_Y = Application.WorksheetFunction.Average(data_Y)
    s_Xsq = Application.WorksheetFunction.Var(data_X)
    s_Ysq = Application.WorksheetFunction.Var(data_Y)
    spsq = ((num_Y - 1) * s_Ysq + (num_X - 1) * s_Xsq) / (num_Y + num_X - 2)
    DF2 = (s_Ysq / num_Y + s_Xsq / num_X) ^ 2 / ((s_Ysq / num_Y) ^ 2 / (num_Y - 1) + (s_Xsq / num_X) ^ 2 / (num_X - 1))
    df22 = Application.RoundDown(DF2, 0)
    paTeq = (avg_Y - avg_X) / Sqr(spsq) / Sqr(1 / num_Y + 1 / num_X)
    paTuneq = (avg_Y - avg_X) / Sqr(s_Ysq / num_Y + s_Xsq / num_X)
    
   
    
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
    tmp = DF2 - df22
    pvalueTuneq = (1 - tmp) * Application.WorksheetFunction.TDist(Abs(paTuneq), df22, 1) + tmp * Application.WorksheetFunction.TDist(Abs(paTuneq), df22 + 1, 1)
        
    Select Case choice(3)
    Case 1
        pvalueTeq = Application.WorksheetFunction.Min(pvalueTeq, 1 - pvalueTeq) * 2
        pvalueTuneq = Application.WorksheetFunction.Min(pvalueTuneq, 1 - pvalueTuneq) * 2
        title = " H0 : σ₁= σ₂vs. H1 : σ₁≠ σ₂     (σ₁: " & list_Y & ", σ₂: " & list_X & " )"
    Case 2
        If paTeq < 0 Then pvalueTeq = 1 - pvalueTeq
        If paTuneq < 0 Then pvalueTuneq = 1 - pvalueTuneq
        title = " H0 : σ₁= σ₂vs. H1 : σ₁> σ₂      (σ₁: " & list_Y & ", σ₂: " & list_X & " )"
    Case 3
        If paTeq > 0 Then pvalueTeq = 1 - pvalueTeq
        If paTuneq > 0 Then pvalueTuneq = 1 - pvalueTuneq
        title = " H0 : σ₁= σ₂vs. H1 : σ₁< σ₂      (σ₁: " & list_Y & ", σ₂: " & list_X & " )"
    End Select
    
    TModulePrint.Title3 list_Y & "과 " & list_X & "의 등분산검정"
    ReDim rstArray(2, 3)
    rstArray(1, 1) = "자유도"
    rstArray(1, 2) = "F - 통계량"
    rstArray(1, 3) = "유의확률"
    rstArray(2, 1) = "( " & dem_df - 1 & " , " & num_df - 1 & " )"
    rstArray(2, 2) = Format(fstat, "##0.0000")
    rstArray(2, 3) = Format(pvaluef, "##0.0000")
    
        TModulePrint.printRst "", rstArray
 '   TModulePrint.printRstOne title, rstArray
    
 

    


End If

End Sub
