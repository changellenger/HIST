Attribute VB_Name = "ModuleControl"
Public DataSheet As String, RstSheet As String      'sheet이름 두 개
Public ylist As String, xlist() As String           '변수명 두 개
Public N As Long, m As Integer, p As Integer        '갯수 세 개
                                                    '이상 Public변수 7개
                                                    '모두 frmRegression 에서 한번만 지정되고
                                                    '다른 곳에서는 바꾸지 않는다.
'''

Public Declare Function ShellExecute _
 Lib "shell32.dll" _
 Alias "ShellExecuteA" ( _
 ByVal hwnd As Long, _
 ByVal lpOperation As String, _
 ByVal lpFile As String, _
 ByVal lpParameters As String, _
 ByVal lpDirectory As String, _
 ByVal nShowCmd As Long) _
 As Long
'''

Sub RegressionShow()

    Dim ErrSignforDataSheet As Integer
    
    ErrSignforDataSheet = InitializeDlg(frmRegression)
                                    
    Select Case ErrSignforDataSheet
    Case 0: frmRegression.Show
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

Function InitializeDlg(ParentDlg) As Integer
   
   Dim myRange As Range: Dim cnt As Long
   Dim myArray() As String
   
   On Error GoTo ErrorFlag
   
   Set myRange = ActiveSheet.Cells.CurrentRegion
   If myRange.count = 1 And myRange.Cells(1, 1) = "" Then
        InitializeDlg = 1: Exit Function
   End If
   Set myRange = ActiveSheet.Cells.CurrentRegion.Rows(1)
   ParentDlg.ListBox1.Clear: ParentDlg.ListBox2.Clear
   ParentDlg.ListBox3.Clear: cnt = myRange.Cells.count
   
   ReDim myArray(0 To cnt - 1)
   For i = 1 To cnt
     myArray(i - 1) = myRange.Cells(i)
   Next i
   ParentDlg.ListBox1.list() = myArray
   InitializeDlg = 0
   Exit Function
   
ErrorFlag:
   InitializeDlg = -1
   
End Function

Function FindVarCount(ListVar) As Long
   
    Dim temp, M2, m3, j As Long
    Dim TempSheet As Worksheet
    Dim tmp2, tmp As Range
    
    Set TempSheet = Worksheets(DataSheet)
    temp = Cells.CurrentRegion.Columns.count
    
   Dim Chk_Ver As Boolean   '파일 버전 체크
   Dim Cmp_R As Long        '파일 버전에 따른 비교 행의 값
   
   '파일 버전에 따른 행과 열의 비교값 정의
   Chk_Ver = ChkVersion(ActiveWorkbook.Name)
   If Chk_Ver = True Then
        Cmp_R = 1048576
    Else
        Cmp_R = 65536
    End If


    For j = 1 To temp
       If StrComp(ListVar, TempSheet.Cells(1, j).value, 1) = 0 Then
          Set tmp2 = TempSheet.Columns(j)
          M2 = tmp2.Cells(1, 1).End(xlDown).row
          If M2 <> Cmp_R Then
             m3 = tmp2.Cells(M2, 1).End(xlDown).row
             If m3 <> Cmp_R Then M2 = m3
          End If
          Set tmp = tmp2.Range(Cells(2, 1), Cells(M2, 1))
       End If
    Next j
    
    FindVarCount = tmp.count
    
End Function

Function FindingRangeError(ListVar) As Boolean
    
    Dim temp, M2, m3, j As Long
    Dim TempSheet As Worksheet
    Dim tmp As Range, tmp11 As Range, tmp1 As Range, tmp2 As Range, tmp3 As Range
    
    Set TempSheet = Worksheets(DataSheet)
    temp = Cells.CurrentRegion.Columns.count
    
   Dim Chk_Ver As Boolean   '파일 버전 체크
   Dim Cmp_R As Long        '파일 버전에 따른 비교 행의 값
   
   '파일 버전에 따른 행과 열의 비교값 정의
   Chk_Ver = ChkVersion(ActiveWorkbook.Name)
   If Chk_Ver = True Then
        Cmp_R = 1048576
    Else
        Cmp_R = 65536
    End If


    For j = 1 To temp
       If StrComp(ListVar, TempSheet.Cells(1, j).value, 1) = 0 Then
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
    
    If tmp.count = 1 And IsNumeric(tmp.Cells(1, 1)) = True Then
        FindingRangeError = False
    Else
        If tmp1 Is Nothing And tmp2 Is Nothing And tmp3 Is Nothing Then
            FindingRangeError = False
        Else: FindingRangeError = True
        End If
    End If
    
End Function

'DataSheet : frmRegression에서만 1회 지정
'RstSheet : "_회귀분석결과_" frmRegression에서만 1회 지정
'ylist, xlist(), N, M, p : frmRegression에서만 1회 지정

    'intercept : 상수항 포함 여부 : Boolean
    'PILevel : 예측값의 신뢰구간 , 불선택시 -1 : double
    'ScatterPlot, PIgraph : Boolean
    'method : 변수 선택 방법 1~4, 불선택시 -1 : integer
    'addlevel : 추가 기준(%), 불선택시 -1  : double
    'rmlevel : 제거 기준(%), 불선택시 -1   : double
    'criteria : (모든가능한) 변수 선택 기준 5~7, 불선택시 -1 : integer
    
''
'' 회귀 분석 결과
''
Sub Reg(intercept As Boolean)
    
    Dim Index()
    
    ModulePrint.Title1 ("선형 회귀분석 결과")
    
    'index 만들기, p개의 변수 모두 1로 하면 된다.
    ReDim Index(p - 1)
    Index = ModuleMatrix.makeIndex(p - 1, 1)
    
    '분산분석, 모수 추정을 출력한다.
    ModulePrint.ANOVA Index, intercept
    ModulePrint.beta Index, intercept

End Sub

''
''변수 선택 결과
''
Sub VarSel(method As Integer, addlevel As Double, rmlevel As Double, criteria, intercept As Boolean, _
                            resi, ci, Alpha, simple)
    
    Select Case method
    
        Case 0
        
        Case 1
        Forward addlevel, intercept, resi, ci, Alpha, simple
        
        Case 2
        Backward rmlevel, intercept, resi, ci, Alpha, simple
        
        Case 3
        Stepwise addlevel, rmlevel, intercept, resi, ci, Alpha, simple
        
        Case 4
        Allpossible criteria, intercept, resi, ci, Alpha, simple
        
    End Select
    
    Worksheets(RstSheet).Activate
    'Columns("B:B").EntireColumn.AutoFit              '변수명 칸 맞추기
    
End Sub
''
''Application.WorksheetFunction.LinEst(y, x, intercept=1, 1) 의 결과
''
'' b_p,   b_(p-1), ... , b_1,    a          -> 배열x의 입력순 i.e. b_1은 x(0)의 계수추정치
''se_p,  se_(p-1), ... ,se_1, se_0
'' r^2,       se
''   F,       df
'' SSR,      SSE
''
Function stat(Index, Y, x, intercept, MSEf)                 'index가 1인 변수들의 통계량 계산
                                                            'index가 2인 변수의 F,P 반환
    Dim k1 As Integer, k2 As Integer
    Dim tmpx()
    Dim b As Double, s As Double, ssr As Double, sse As Double
    Dim rst()
    Dim DFssr As Long, DFsse As Long
    
    k1 = 0
    For j = 0 To p - 1
        If Index(j) = 1 Then k1 = k1 + 1                  'k1 : index가 0이 아닌 것의 개수
        If Index(j) = 2 Then k1 = k1 + 1: k2 = k1         'k2 : 그 중 몇번째 변수의 index가 2인지
    Next j
    
    If k1 = 0 Then Exit Function
    
    ReDim tmpx(N - 1, k1 - 1)
    
    tmpx = ModuleMatrix.selectedX(Index, x)
    
    rst = Application.WorksheetFunction.LinEst(Y, tmpx, intercept, 1)
    
    '변수와 관련된 index는 0~(p-1)까지
    'index(p)     index(p+1)   index(p+2)    index(p+3)   index(p+4)    index(p+5)    index(p+6)
    '변수명(j)    SSR          SSE           DFssr        DFsse         F-value       P-value
    'index(p+7)   index(p+8)   index(p+9)    index(p+10)
    'R-square     adjR-square  Cp            AIC
    
    If intercept <> 0 Then
        ssr = rst(5, 1)
        sse = rst(5, 2)
        DFssr = N - 1 - rst(4, 2)
        DFsse = rst(4, 2)
    Else
        ssr = ModuleMatrix.noIntSSR(Y, tmpx)
        sse = rst(5, 2)
        DFssr = N - rst(4, 2)
        DFsse = rst(4, 2)
    End If
    
    Index(p + 1) = ssr
    Index(p + 2) = sse
    Index(p + 3) = DFssr
    Index(p + 4) = DFsse
    If k2 > 0 Then
    b = Application.WorksheetFunction.Index(rst, 1, k1 - k2 + 1)        'index가 2인 변수의 것
    s = Application.WorksheetFunction.Index(rst, 2, k1 - k2 + 1)
    Index(p + 5) = b ^ 2 / s ^ 2
    Index(p + 6) = Application.WorksheetFunction.FDist(b ^ 2 / s ^ 2, 1, DFsse)
    Else
    Index(p + 5) = 1
    Index(p + 6) = 1                                                    '의미없음 에러대비
    End If

    Index(p + 7) = ssr / (ssr + sse)
    Index(p + 8) = 1 - (sse / DFsse) * ((DFssr + DFsse) / (sse + ssr))
    Index(p + 9) = DFssr + 1 + ((sse / DFsse - MSEf) / MSEf) * (N - DFssr - 1)
    Index(p + 10) = N * Log(sse / N) + 2 * (DFssr + 1)
    
    stat = Index
    
End Function

Sub Forward(addlevel As Double, intercept As Boolean, resi, ci, Alpha, simple)
    
    Dim j As Integer, K As Integer
    Dim Y(), x(), tmpx(), Index(), summary()
    Dim max(10)
    Dim MSEf As Double
    
    ReDim Index(p + 10)
    Index = ModuleMatrix.makeIndex(p + 10, 0)
     
    ReDim Y(N - 1, 0)
    Y = ModuleMatrix.pureY
    
    ReDim x(N - 1, p - 1)
    x = ModuleMatrix.pureX
    
    MSEf = ModuleMatrix.fullModelMSE(Y, x, intercept)
    
    ModulePrint.Title1 "변수선택 결과"
    
    '요약 통계량을 제시할 때 사용하기 위함.
    max(0) = -1: max(5) = -1
    ReDim summary(p - 1, 10)
        
    For K = 0 To p - 1
    For j = 0 To p - 1
        If Index(j) = 0 Then
            Index(j) = 2
            Index = stat(Index, Y, x, intercept, MSEf)
            If max(5) < Index(p + 5) Then
                max(0) = j
                
                For i = 1 To 10
                max(i) = Index(p + i)
                Next i
                
            End If
            Index(j) = 0
        End If
    Next j
        If max(6) > addlevel Then GoTo LastLine
        Index(max(0)) = 1
        
        For i = 0 To 10
            summary(K, i) = max(i)
        Next i
        
        '단계수 만큼의 행이 기록되게 된다.
        
        ModulePrint.Title2 "변수추가 " & K + 1 & "단계"
        ModulePrint.Comment "변수 " & xlist(max(0)) & " 진입 : 결정계수 = " _
            & Format(summary(K, 7), "0.0000") & " ,Cp = " & Format(summary(K, 9), "0.0000")

        ModulePrint.ANOVA Index, intercept
        ModulePrint.beta Index, intercept
        
        max(0) = -1: max(5) = -1
    Next K
        
LastLine:

    ModulePrint.Title2 "변수추가 요약"
    If K = 0 Then ModulePrint.Comment "추가되는 변수가 없습니다."
    ModulePrint.summaryAdd summary, K
    
    For i = 1 To 18
        If resi(i) = True Then
            x = ModuleMatrix.selectedX(Index, x)
            ModuleResi.Diagnosis resi, intercept, ci, Alpha, simple, x, Index
            Exit For
        End If
    Next i
    
End Sub

Sub Backward(rmlevel As Double, intercept As Boolean, resi, ci, Alpha, simple)
    
    Dim j As Integer, K As Integer
    Dim Y(), x(), tmpx(), Index(), summary(), rst()
    Dim min(10)
    Dim MSEf As Double
    
    ReDim Index(p + 10)
    Index = ModuleMatrix.makeIndex(p + 10, 1)
     
    ReDim Y(N - 1, 0)
    Y = ModuleMatrix.pureY
    
    ReDim x(N - 1, p - 1)
    x = ModuleMatrix.pureX
    
    MSEf = ModuleMatrix.fullModelMSE(Y, x, intercept)
    
    '제목 쓰기
    ModulePrint.Title1 "변수선택 결과"
    
    '요약 통계량을 제시할 때 사용하기 위함
    min(0) = 99999: min(5) = 99999
    ReDim summary(p - 1, 10)
        
        ModulePrint.Title2 "변수제거 0 단계"
        ModulePrint.Comment "변수제거없음"
        ModulePrint.ANOVA Index, intercept
        ModulePrint.beta Index, intercept
        
    For K = 0 To p - 1
    For j = 0 To p - 1
        If Index(j) = 1 Then
            Index(j) = 2
            Index = stat(Index, Y, x, intercept, MSEf)
            If min(5) > Index(p + 5) Then
                min(0) = j
                
                For i = 1 To 10
                min(i) = Index(p + i)
                Next i
                
            End If
            Index(j) = 1
        End If
    Next j
        If min(6) < rmlevel Then GoTo LastLine
                
        Index(min(0)) = 0                   '제거
        summary(K, 0) = min(0)
        '변수 제거된 후의 index를 사용해야 한다
        tmpx = ModuleMatrix.selectedX(Index, x)
        Index = stat(Index, Y, x, intercept, MSEf)
        For i = 1 To 10
            summary(K, i) = Index(p + i)
        Next i
        'F-value, P-value 기록
        summary(K, 5) = min(5): summary(K, 6) = min(6)

        ModulePrint.Title2 "변수제거 " & K + 1 & "단계"
        ModulePrint.Comment "변수 " & xlist(min(0)) & " 제거 : 결정계수 = " _
            & Format(summary(K, 7), "0.0000") & " ,Cp = " & Format(summary(K, 9), "0.0000")
        ModulePrint.ANOVA Index, intercept
        ModulePrint.beta Index, intercept
        
        min(0) = 99999: min(5) = 99999
    Next K
        
LastLine:

    ModulePrint.Title2 "변수제거 요약"
    If K = 0 Then ModulePrint.Comment "제거되는 변수가 없습니다."
    ModulePrint.summaryRm summary, K
    
    For i = 1 To 18
        If resi(i) = True Then
            x = ModuleMatrix.selectedX(Index, x)
            ModuleResi.Diagnosis resi, intercept, ci, Alpha, simple, x, Index
            Exit For
        End If
    Next i

    
End Sub

Sub Stepwise(addlevel, rmlevel, intercept, resi, ci, Alpha, simple)

    Dim j As Integer, K As Integer, numInModel As Integer, p1 As Integer, i As Integer
    Dim Y(), x(), tmpx(), Index(), summary()
    Dim max(10), min(10)
    Dim MSEf As Double
    Dim stepNum As Long
    
    ReDim Index(p + 10)
    Index = ModuleMatrix.makeIndex(p + 10, 0)
     
    ReDim Y(N - 1, 0)
    Y = ModuleMatrix.pureY
    
    ReDim x(N - 1, p - 1)
    x = ModuleMatrix.pureX
    
    MSEf = ModuleMatrix.fullModelMSE(Y, x, intercept)
    
    '제목 쓰기
    ModulePrint.Title1 "변수선택 결과"
    
    
    '요약 통계량을 제시할 때 사용하기 위함.
    max(0) = -1: max(5) = -1: min(0) = 99999: min(5) = 99999
    ReDim summary(2 * p - 1, 11)                            '11에 진입:1 제거:-1 제거실패:0
       
    stepNum = 0
    K = 0
    Do While K < 2 * p + 1
    
    'Forward
    For j = 0 To p - 1
        If Index(j) = 0 Then
            Index(j) = 2
            Index = stat(Index, Y, x, intercept, MSEf)
            If max(5) < Index(p + 5) Then
                max(0) = j
                
                For i = 1 To 10
                max(i) = Index(p + i)
                Next i
                
            End If
            Index(j) = 0
        End If
    Next j
    
        If max(6) > addlevel Or K = 2 * p Then GoTo LastLine                'k=2*p중요-모든변수진입후중단
        Index(max(0)) = 1
        
        p1 = 0
        For i = 0 To p - 1
            If Index(i) <> 0 Then p1 = p1 + 1
        Next i
        
        'Forward  결과 기록
        For i = 0 To 10
            summary(K, i) = max(i)
        Next i
        summary(K, 11) = 1
    
    
        numInModel = 0
        For j = 0 To p - 1
            If Index(j) <> 0 Then numInModel = numInModel + 1
        Next j
            
        stepNum = stepNum + 1
        ModulePrint.Title2 "변수증감 " & stepNum & "단계"
        
        ModulePrint.Comment "변수 " & xlist(max(0)) & " 추가 : 결정계수 = " _
                                & Format(summary(K, 7), "0.0000") & "" _
                                & " ,Cp = " & Format(summary(K, 9), "0.0000")
        ModulePrint.ANOVA Index, intercept
        ModulePrint.beta Index, intercept
    K = K + 1
    
    'Backward
    For j = 0 To p - 1
        If Index(j) = 1 Then
            Index(j) = 2
            Index = stat(Index, Y, x, intercept, MSEf)
            If min(5) > Index(p + 5) Then
                min(0) = j
                
                For i = 1 To 10
                min(i) = Index(p + i)
                Next i
                
            End If
            Index(j) = 1
        End If
    Next j
    
        If min(6) < rmlevel Then        'NoRemove의 경우
            summary(K, 11) = 0
        Else                            'Remove의 경우
            'Backward 결과 기록
            '변수명 기록
            summary(K, 0) = min(0)
        
            Index(min(0)) = 0                   '제거
            summary(K, 11) = -1
            
            '변수 제거된 후의 index를 사용해야 한다
            tmpx = ModuleMatrix.selectedX(Index, x)
            Index = stat(Index, Y, x, intercept, MSEf)
            For i = 1 To 10
                summary(K, i) = Index(p + i)
            Next i
            summary(K, 5) = min(5): summary(K, 6) = min(6)
            stepNum = stepNum + 1
        End If
       
        If summary(K, 11) = -1 Then
            ModulePrint.Title2 "변수증감 " & stepNum & "단계"
            ModulePrint.Comment "변수 " & xlist(min(0)) & " 제거 : 결정계수 = " _
                                    & Format(summary(K, 7), "0.0000") & "" _
                                    & " ,Cp = " & Format(summary(K, 9), "0.0000")
            ModulePrint.ANOVA Index, intercept
            ModulePrint.beta Index, intercept
        End If
        
    K = K + 1
    max(0) = -1: max(5) = -1: min(0) = 99999: min(5) = 99999
    Loop
    
LastLine:

    ModulePrint.Title2 "변수증감 요약"
    
    ModulePrint.summaryStep summary, K
    
    For i = 1 To 18
        If resi(i) = True Then
            x = ModuleMatrix.selectedX(Index, x)
            ModuleResi.Diagnosis resi, intercept, ci, Alpha, simple, x, Index
            Exit For
        End If
    Next i

    
End Sub

Sub Allpossible(criteria, intercept, resi, ci, Alpha, simple)
    Dim Index(), rst(), Y(), x(), tmpx()
    Dim num As Long, i As Long
    Dim col As Integer, j As Integer, numInModel As Integer, K As Integer
    Dim tmpstr As String
    Dim MSEf As Double
    Dim varInModel As String
    
    ReDim Index(p + 10)
    num = 2 ^ p - 1
    col = UBound(criteria) + 1
    ReDim rst(num, col + 3)
    
    col = 0
    rst(0, 0) = "변수이름"
    rst(0, 1) = "변수개수"
    rst(0, 2) = "결정계수"
    If criteria(0) = 1 Then rst(0, col + 3) = "수정결정계수": col = col + 1
    If criteria(1) = 1 Then rst(0, col + 3) = "Cp": col = col + 1
    If criteria(2) = 1 Then rst(0, col + 3) = "AIC": col = col + 1                  '^^

    'data 잡아주기
    '계산횟수를 줄이려면 여기서 계산해 함수인수로 넘긴다
    ReDim Y(N - 1, 0)
    Y = ModuleMatrix.pureY
    
    ReDim x(N - 1, p - 1)
    x = ModuleMatrix.pureX
    
    MSEf = ModuleMatrix.fullModelMSE(Y, x, intercept)
    
    '제목 쓰기
    ModulePrint.Title1 "변수선택 결과"
    ModulePrint.Title2 "모든 가능한 회귀"
    ModulePrint.tableAll col
    
        For i = 0 To num - 1
        
        tmpstr = ModuleMatrix.binStr(i + 1)
        col = Len(tmpstr)                       '변수 또 만들지 않으려고 의미없이 빌려사용 col
        For j = 1 To p - col
            tmpstr = "0" & tmpstr
        Next j
        
            numInModel = 0
            varInModel = ""
            For j = 0 To p - 1
                If Mid(tmpstr, j + 1, 1) <> "1" Then
                Index(j) = 0
                Else
                Index(j) = 1: numInModel = numInModel + 1: varInModel = varInModel & xlist(j) & " "
                End If
            Next j                                  'index 잡아주기
        
        Index = stat(Index, Y, x, intercept, MSEf)
        col = 0
        rst(i + 1, 0) = varInModel
        rst(i + 1, 1) = numInModel
        rst(i + 1, 2) = Index(p + 7)
        If criteria(0) = 1 Then rst(i + 1, col + 3) = Index(p + 8): col = col + 1
        If criteria(1) = 1 Then rst(i + 1, col + 3) = Index(p + 9): col = col + 1
        If criteria(2) = 1 Then rst(i + 1, col + 3) = Index(p + 10): col = col + 1
        Next i
        
    ModulePrint.All rst
    
    For i = 1 To 18
        If resi(i) = True Then
            x = ModuleMatrix.selectedX(Index, x)
            ModuleResi.Diagnosis resi, intercept, ci, Alpha, simple, x, Index
            Exit For
        End If
    Next i

End Sub

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


Function msgWord(word) As String                    '마사지 word 는 string

    If Len(word) > 9 Then
        msgWord = Mid(word, 1, 6) & vbLf & Mid(word, 7)
    Else
        msgWord = word
    End If
    
End Function


'파일 버전 체크
Function ChkVersion(File_Name) As Boolean
    
    If Right(File_Name, 4) = ".xls" Or Right(File_Name, 4) = ".XLS" Then
        ChkVersion = False
    Else
        ChkVersion = True
    End If
End Function
