VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameReMulti 
   OleObjectBlob   =   "frameReMulti.frx":0000
   Caption         =   "따라하기 : 다중선형회귀분석"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9105
   StartUpPosition =   1  '소유자 가운데
   TypeInfoVer     =   143
End
Attribute VB_Name = "frameReMulti"
Attribute VB_Base = "0{6B82FC11-AFE7-46C1-A2CB-B03AE4CE3E5A}{EC1A1963-7EC8-44BE-81B7-318BF52A4660}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

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

               Exit Sub
            End If
            i = i + 1
        Loop
    End If
    
    
    Dim j As Integer
    j = 0
    If Me.ListBox3.ListCount = 0 Then
        Do While j <= Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(j) = True Then
               Me.ListBox3.AddItem Me.ListBox1.list(j)
               Me.ListBox1.RemoveItem (j)

               Exit Sub
            End If
            j = j + 1
        Loop
    End If
    
    
End Sub

Private Sub CB4_Click()

    Dim i As Integer
    Dim FromLNum, ToLNum
    
    i = 0
    FromLNum = "ListBox3": ToLNum = "ListBox1"
    Do While i <= Me.Controls(FromLNum).ListCount - 1
        If Me.Controls(FromLNum).Selected(i) = True Then
           Me.Controls(ToLNum).AddItem Me.Controls(FromLNum).list(i)
           Me.Controls(FromLNum).RemoveItem i
           Me.Controls(FromLNum).Selected(i) = False
          'i = i - 1
        End If
        i = i + 1
    Loop

End Sub


Private Sub CB3_Click()

    Dim i As Integer
    Dim FromLNum, ToLNum
    
    i = 0
    FromLNum = "ListBox1": ToLNum = "ListBox3"
    Do While i <= Me.Controls(FromLNum).ListCount - 1
        If Me.Controls(FromLNum).Selected(i) = True Then
           Me.Controls(ToLNum).AddItem Me.Controls(FromLNum).list(i)
           Me.Controls(FromLNum).RemoveItem i
           Me.Controls(FromLNum).Selected(i) = False
          ' i = i - 1
        End If
        i = i + 1
    Loop
End Sub


Private Sub Label15_Click()
    frameReEx2.Show
    
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
    End If


   ElseIf Me.ListBox3.ListCount = 0 Then
        Do While j <= Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(j) = True Then
               Me.ListBox3.AddItem Me.ListBox1.list(j)
               Me.ListBox1.RemoveItem (j)
               Me.CB1.Visible = False
               Me.CB2.Visible = True
               Exit Sub
            End If
            j = j + 1
        Loop
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
    If Me.ListBox3.ListCount <> 0 Then
        Me.ListBox1.AddItem Me.ListBox3.list(0)
        Me.ListBox3.RemoveItem (0)
        Me.CB1.Visible = True
        Me.CB2.Visible = False
    End If
End Sub

Private Sub OK_Click()

    Dim choice(3) As Variant                            '넘길 정보는 검정값,신뢰구간,대립가설 3개니까
    Dim dataRange As Range
    Dim i As Integer
    Dim activePt As Long                                '결과 분석이 시작되는 부분을 보여주기 위함
    
    '''
    '''변수를 선택하지 않았을 경우
    '''
    If Me.ListBox2.ListCount = 0 Then
        MsgBox "변수를 선택해 주시기 바랍니다.", vbExclamation, "HIST"
        Exit Sub
    End If
     Me.ChB1.value = True
    '''
    '''public 변수 선언 xlist, DataSheet, RstSheet, m, k1, n
    '''
    xlist = Me.ListBox2.list(0)
    DataSheet = ActiveSheet.Name                        'DataSheet : Data가 있는 Sheet 이름
    RstSheet = "_통계분석결과_"                       'RstSheet  : 결과를 보여주는 Sheet 이름
    
    
    
    '맨위에 입력
On Error GoTo Err_delete
Dim val3535 As Long '초기위치 저장할 공간'
Dim s3535 As Worksheet
val3535 = 2
    For Each s3535 In ActiveWorkbook.Sheets
        If s3535.Name = RstSheet Then
val3535 = Sheets(RstSheet).Cells(1, 1).value
End If
Next s3535  '시트가 이미있으면 출력 위치 저장을하고, 없으면 2을 저장한다.


    
    Set dataRange = ActiveSheet.Cells.CurrentRegion
    m = dataRange.Columns.count                         'm         : dataSheet에 있는 변수 개수
    
    tmp = 0
    For i = 1 To m
        If xlist = ActiveSheet.Cells(1, i) Then
            k1 = i  'k1                                 : k1 : 선택된 변수가 몇번째 열에 있는지
            tmp = tmp + 1
        End If
    Next i
    N = ActiveSheet.Cells(1, k1).End(xlDown).row - 1    'n         : 선택된 변수의 데이타 갯수

    '''
    ''' 변수명이 같은 경우 - 마지막 열에 있는 변수만 입력되므로 에러처리한다.
    '''
    If tmp > 1 Then
        MsgBox xlist & "와 같은 변수명이 있습니다. " & vbCrLf & "변수명을 바꿔주시기 바랍니다.", vbExclamation, "HIST"
        Exit Sub
    End If
    
    '''
    '''숫자와 문자가 혼합되어 있을 경우
    '''
    If TModuleControl.FindingRangeError(xlist) = True Then
        MsgBox "다음의 분석변수에 문자나 공백이 있습니다." & Chr(10) & _
               ": " & xlist, vbExclamation, "HIST"
        Exit Sub
    End If
    
    '''
    '''검정값을 입력하지 않은 경우
    '''
    If IsNumeric(Me.TextBox1.value) = False Then
        MsgBox "사용자 검정값을 입력해 주시기 바랍니다.", vbExclamation, "HIST"
        Exit Sub
    End If
    
    '''
    '''신뢰구간을 잘못 입력한 경우
    '''
  '  If Me.ChB1.Value = True Then
        If IsNumeric(Me.TextBox2.value) = False Then
            MsgBox "사용자 신뢰구간을 입력해 주시기 바랍니다.", vbExclamation, "HIST"
            Exit Sub
        ElseIf Me.TextBox2.value < 0 Or Me.TextBox2.value > 100 Then
            MsgBox "사용자 신뢰구간을 %단위로 입력해 주시기 바랍니다.", vbExclamation, "HIST"
            Exit Sub
        End If
  '  End If
    
    '''
    ''' 데이타 개수가 한개일 경우
    '''
    If N = 1 Then
        MsgBox "한 개의 데이타로 검정을 시행할 수 없습니다.", vbExclamation, "HIST"
        Exit Sub
    End If
    
    '''
    '''검정값 선택결과 입력 - choice(1)
    '''
    choice(1) = Me.TextBox1.value
    
    '''
    '''신뢰구간 입력 - choice(2)
    '''
    If Me.ChB1.value = True Then choice(2) = Me.TextBox2.value
    If Me.ChB1.value = False Then choice(2) = -1
    
    '''
    '''귀무가설 선택결과 입력 - choice(3)
    '''
    If Me.OB1 = True Then choice(3) = 1
    If Me.OB2 = True Then choice(3) = 2
    If Me.OB3 = True Then choice(3) = 3
    
    
    '''
    '''결과 처리
    '''
    TModuleControl.SettingStatusBar True, "일표본 t-검정중입니다."
    Application.ScreenUpdating = False
    TModulePrint.MakeOutputSheet (RstSheet)
    'Worksheets(RstSheet).Unprotect "prophet"
    activePt = Worksheets(RstSheet).Cells(1, 1).value
    
    TModuleControl.TTestR choice
    
    
    'Worksheets(RstSheet).Protect Password:="prophet", DrawingObjects:=False, _
    '                                contents:=True, Scenarios:=True
    TModuleControl.SettingStatusBar False
    Application.ScreenUpdating = True
    Unload Me
    

    Worksheets(RstSheet).Activate
    
    '파일 버전 체크 후 비교값 정의
    Dim Cmp_Value As Long
    
    If PublicModule.ChkVersion(ActiveWorkbook.Name) = True Then
        Cmp_Value = 1048000
    Else
        Cmp_Value = 65000
    End If
    
    If Worksheets(RstSheet).Cells(1, 1).value > Cmp_Value Then
        MsgBox "[_통계분석결과_]시트를 거의 모두 사용하였습니다." & vbCrLf & "이 시트의 이름을 바꾸거나 삭제해 주세요", vbExclamation, "HIST"
        Exit Sub
    End If
    
    Worksheets(RstSheet).Cells(activePt + 10, 1).Select
    Worksheets(RstSheet).Cells(activePt + 10, 1).Activate
                            '결과 분석이 시작되는 부분을 보여주며 마친다.
                                
        
    Unload Me



'맨뒤에 붙이기
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

MsgBox ("프로그램에 문제가 있습니다 .")
 'End sub 앞에다 붙인다.

''해석, 에러가 나면 Err_delete로 와서 첫셀이후로 지운다. 만약 첫셀이 2면 시트를 지운다.그리고 에러메시지 출력
'rSTsheet만들기도 전에 에러나는 경우에는 아무 동작도 하지 않고, 에러메시지만 띄운다.
        
End Sub
Private Sub OK1_Click()                                                '''  "_회귀분석결과_"

    Dim intercept As Boolean
    Dim ci As Boolean
    Dim Alpha As Single
    Dim ScatterPlot As Boolean, PIgraph As Boolean
    Dim resi(18) As Boolean         'resi(0)사용 '안' 할것임.
    Dim simple(3)                   'simple(0)사용 '안' 할것임.
    Dim method As Integer
    Dim addlevel As Double, rmlevel As Double
    Dim criteria(2)
    Dim sign As Boolean, errsign1 As Boolean, errsign2 As Boolean
    Dim errString As String
    Dim activePt As Long            '결과 분석이 시작되는 부분을 보여주기 위함
    
    Dim ws As Worksheet
    Dim check1 As Integer, check2 As Integer
    
    '''
    '''에러 처리 부분 1
    '''
    If Me.ListBox2.ListCount = 0 Or Me.ListBox3.ListCount = 0 Then
        MsgBox "변수 선택이 완전하지 않습니다.", vbExclamation
        Exit Sub
    End If
    If IsNumeric(Me.TextBox1.value) = False Then
        MsgBox "신뢰확률이 올바르지 않습니다.", vbExclamation
        Exit Sub
    Else
        If Me.TextBox1.value <= 0 Or Me.TextBox1.value >= 100 Then
            MsgBox "신뢰확률이 올바르지 않습니다.", vbExclamation
            Exit Sub
        End If
    End If


    '''
    '''입력받은 정보 정리하기
        
    '여기부터 MdControl 에서 선언된 Public 변수
    '여기서 한번만 지정해준다
    
    DataSheet = ActiveSheet.Name        'Data가 있는 Sheet 이름
    RstSheet = "_통계분석결과_"         '결과를 보여주는 Sheet 이름
    '출력하는 해당 모듈에 덧 붙일 내용'
'맨위에 입력
On Error GoTo Err_delete
Dim val3535 As Long '초기위치 저장할 공간'
Dim s3535 As Worksheet
val3535 = 2
    For Each s3535 In ActiveWorkbook.Sheets
        If s3535.Name = RstSheet Then
val3535 = Sheets(RstSheet).Cells(1, 1).value
End If
Next s3535  '시트가 이미있으면 출력 위치 저장을하고, 없으면 2을 저장한다.

                                        
    '''
    ylist = Me.ListBox2.list(0)            '선택된 종속변수이름
    p = Me.ListBox3.ListCount              '선택된 독립변수 개수
    
    ReDim xlist(p - 1)
    For i = 0 To p - 1
        xlist(i) = ListBox3.list(i)         '선택된 독립변수 이름
    Next i
    
    Set dataRange = ActiveSheet.Cells.CurrentRegion
    N = dataRange.Cells(1, 1).End(xlDown).row - 1       'Data개수
    m = dataRange.Cells(1, 1).End(xlToRight).Column - 1 '독립변수 개수
    
    '여기까지 MdControl 에서 선언된 Public 변수
    '여기서 한번만 지정해준다. 다른 곳에서 바꾸지 않는다
    'DataSheet, RstSheet, ylist, xlist, N, M, p
    
    
    
    '''
    '''
    '변수선택 입력정보 정리하기
    intercept = CheckBox1.value
    ci = CheckBox2.value
    Alpha = TextBox1.value
    
    simple(1) = CheckBox3.value     '산점도
    simple(2) = CheckBox4.value     '신뢰대 그래프
    simple(3) = CheckBox5.value     'vs독립변수 그래프
    
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
    '''변수들의 관측수의 대응
    '''
    If N <> Modulecontrol.FindVarCount(ylist) Then errsign1 = True
    For i = 0 To p - 1
        If N <> Modulecontrol.FindVarCount(xlist(i)) Then errsign1 = True
    Next i
    '''
    '''숫자와 문자가 혼합되어 있을 경우
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
    '''에러가 있을 경우 에러 메시지 출력
    '''
    If errsign1 = True Then
        MsgBox "변수들의 관측수가 다릅니다.", _
                vbExclamation, "HIST"
        Exit Sub
    End If
    If errsign2 = True Then
        MsgBox "다음의 분석변수에 문자나 공백이 있습니다." & Chr(10) & _
               ": " & errString, vbExclamation, "HIST"
        Exit Sub
    End If
                                                           
    '''
    '''실제로 처리하는 부분
    '''
    Modulecontrol.SettingStatusBar True, "회귀 분석중입니다."
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
    
    '결과 분석이 시작되는 부분에서 조금 아래 쪽을 보여주며 마친다.
    Worksheets(RstSheet).Activate
    
    '파일 버전 체크 후 비교값 정의
    Dim Cmp_Value As Long
    
    If Modulecontrol.ChkVersion(ActiveWorkbook.Name) = True Then
        Cmp_Value = 1048000
    Else
        Cmp_Value = 65000
    End If
    
    If Worksheets(RstSheet).Cells(1, 1).value > Cmp_Value Then
        MsgBox "[_통계분석결과_]시트를 거의 모두 사용하였습니다." & vbCrLf & "이 시트의 이름을 바꾸거나 삭제해 주세요", vbExclamation, "HIST"
        Exit Sub
    End If
    
    Worksheets(RstSheet).Range("a" & activePt + 10).Select
    Worksheets(RstSheet).Range("a" & activePt + 10).Activate
    
    Unload Me
    
    
Exit Sub
'맨뒤에 붙이기
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
MsgBox ("프로그램에 문제가 있습니다.")

 'End sub 앞에다 붙인다.

''해석, 에러가 나면 Err_delete로 와서 지운다. Rstsheet가 없으면 안지운다. RSTsheet만들기도 전에 ''에러나면 뭐.. 상관은 없을 것 같지만.
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
'-------------
  'Set myRange = Cells.CurrentRegion.Rows(1)
   'cnt = myRange.Cells.Count
   'ReDim myArray(cnt - 1)
  ' For i = 1 To cnt
  '   myArray(i - 1) = myRange.Cells(i)
  ' Next i
   'Me.ListBox1.List() = myArray
'-----------
    ReDim myArray(TempSheet.UsedRange.Columns.count - 1)
    a = 0
   For i = 1 To TempSheet.UsedRange.Columns.count
   If arrName(i) <> "" Then                     '빈칸제거
   myArray(a) = arrName(i)
   a = a + 1
   
   Else:
   End If
   Next i
   
   
   
   Me.ListBox1.list() = myArray
   
 '  For i = 1 To TempSheet.UsedRange.Columns.Count
 '   rngFirst.Offset(i, 1) = myArray(i - 1)
 ' Next i
  


    For j = 1 To TempSheet.UsedRange.Columns.count
        arrName(j) = TempSheet.Cells(1, j)
    Next j
   
   Me.ListBox1.Clear
'-------------
  'Set myRange = Cells.CurrentRegion.Rows(1)
   'cnt = myRange.Cells.Count
   'ReDim myArray(cnt - 1)
  ' For i = 1 To cnt
  '   myArray(i - 1) = myRange.Cells(i)
  ' Next i
   'Me.ListBox1.List() = myArray
'-----------
    ReDim myArray(TempSheet.UsedRange.Columns.count - 1)
    a = 0
   For j = 1 To TempSheet.UsedRange.Columns.count
   If arrName(j) <> "" Then                     '빈칸제거
   myArray(a) = arrName(j)
   a = a + 1
   
   Else:
   End If
   Next j
   
   
   
   Me.ListBox1.list() = myArray
   
 '  For i = 1 To TempSheet.UsedRange.Columns.Count
 '   rngFirst.Offset(i, 1) = myArray(i - 1)
 ' Next i
  

End Sub

Private Sub UserForm_Click()

End Sub
