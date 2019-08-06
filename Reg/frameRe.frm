VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameRe 
   OleObjectBlob   =   "frameRe.frx":0000
   Caption         =   "회귀분석"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5370
   StartUpPosition =   1  '소유자 가운데
   TypeInfoVer     =   47
End
Attribute VB_Name = "frameRe"
Attribute VB_Base = "0{2E49CCFA-88BE-4882-AAA7-71E9D5213812}{4418A814-02ED-48CC-9A83-9544F6CF3D6C}"
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
   If arrName(i) <> "" Then                     '빈칸제거
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
         Do While i <= Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(i) = True Then
               Me.ListBox3.AddItem Me.ListBox1.list(i)
               Me.ListBox1.RemoveItem (i)
               Exit Sub
            End If
            i = i + 1
        Loop
 
End Sub

Private Sub CB4_Click()
    If Me.ListBox3.ListCount <> 0 Then
        Me.ListBox1.AddItem ListBox3.list(0)
        Me.ListBox3.RemoveItem (0)
   
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
ShellExecute 0, "open", "hh.exe", ThisWorkBook.Path + "\HIST%202013.chm::/회귀분석.htm", "", 1
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
    Dim Sign As Boolean, errSign1 As Boolean, ErrSign2 As Boolean
    Dim errString As String
    Dim activePt As Long            '결과 분석이 시작되는 부분을 보여주기 위함
    
    Dim WS As Worksheet
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
    n = dataRange.Cells(1, 1).End(xlDown).row - 1       'Data개수
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
 '   For i = 0 To 3
 '       If frmVarSel.Controls("OptionButton" & (i + 1)).value = True Then
 '       method = i + 1
 '       End If
 '   Next i
    
'    If frmVarSel.Controls("Label1").Enabled = True Then
'        addlevel = frmVarSel.Controls("TextBox1").value / 100
'    Else: addlevel = -1
'    End If
    
'    If frmVarSel.Controls("Label2").Enabled = True Then
'        rmlevel = frmVarSel.Controls("TextBox2").value / 100
'    Else: rmlevel = -1
'    End If
          
'    For i = 1 To 3
'        criteria(i - 1) = 0
'        If frmVarSel.Controls("CheckBox" & i).value = True Then
'        criteria(i - 1) = 1
'        End If
'    Next i
                                                       
    'intercept : 상수항 포함 여부 : Boolean
    'alpha : 예측값의 신뢰구간 , '''불선택시 -1 : double
    'ScatterPlot, PIgraph : Boolean '''이름바꿈 simple()으로
    'method : 변수 선택 방법 1~4, 불선택시 -1 : integer
    'addlevel : 추가 기준(%), 불선택시 -1  : double
    'rmlevel : 제거 기준(%), 불선택시 -1   : double
    'criteria : (모든가능한) 변수 선택 기준 5~7, 불선택시 -1 : integer
    
    
    '''
    '''
    '회귀진단 입력정보 정리하기
        
  '  For i = 1 To 18
  '      resi(i) = frmResid.Controls("CheckBox" & i).value
  '  Next i
    
    '진단통계량
    'resi(1)    : 출력
    'resi(2)    : vs 관측순서 그래프
    'resi(3)    : vs 예측치 그래프
    'resi(4)    : 히스토그램
    'resi(5)    : 정규확률그림
    
    'resi(1)~(5) 잔차
    'resi(6)~(10) 표준화 잔차
    'resi(11)~(15) 표준화 제외 잔차
    
    'resi(16)   :영향관측점
    'resi(17)   :다중공선성
    'resi(18)   :부분회귀산점도
    
    '새로열었을 경우 임시시트 삭제하기
    
    check1 = 0
    check2 = 0
    For Each WS In Worksheets
        If WS.Name = RstSheet Then check1 = 1
        If WS.Name = "_#TmpHIST1#_" Then check2 = 1
    Next WS
    
    Application.DisplayAlerts = False
    If check1 = 0 And check2 = 1 Then Worksheets("_#TmpHIST1#_").Delete
    Application.DisplayAlerts = True

    '''
    '''변수들의 관측수의 대응
    '''
    If n <> ModuleControl.FindVarCount(ylist) Then errSign1 = True
    For i = 0 To p - 1
        If n <> ModuleControl.FindVarCount(xlist(i)) Then errSign1 = True
    Next i
    '''
    '''숫자와 문자가 혼합되어 있을 경우
    '''
    If ModuleControl.FindingRangeError(ylist) = True Then
        ErrSign2 = True: errString = Me.ListBox2.list(0)
    End If
    
    For i = 0 To p - 1
        If ModuleControl.FindingRangeError(xlist(i)) = True Then
            ErrSign2 = True
            If errString <> "" Then
                errString = errString & "," & xlist(i)
            Else: errString = xlist(i)
            End If
        End If
    Next i
    '''
    '''에러가 있을 경우 에러 메시지 출력
    '''
    If errSign1 = True Then
        MsgBox "변수들의 관측수가 다릅니다.", _
                vbExclamation, "HIST"
        Exit Sub
    End If
    If ErrSign2 = True Then
        MsgBox "다음의 분석변수에 문자나 공백이 있습니다." & Chr(10) & _
               ": " & errString, vbExclamation, "HIST"
        Exit Sub
    End If
                                                           
    '''
    '''실제로 처리하는 부분
    '''
    ModuleControl.SettingStatusBar True, "회귀 분석중입니다."
    Application.ScreenUpdating = False
    
    ModulePrint.MakeOutputSheet RstSheet
    'Worksheets(RstSheet).Unprotect "prophet"
    activePt = Worksheets(RstSheet).Range("a1").value
    
   
    ModuleControl.Reg intercept
    
    If p > 1 Then ModuleControl.VarSel method, addlevel, rmlevel, criteria, intercept, resi, ci, Alpha, simple
    
    If method <= 0 Or p = 1 Then ModuleResi.Diagnosis00 resi, intercept, ci, Alpha, simple
    
    'Worksheets(RstSheet).Protect password:="prophet", DrawingObjects:=False, _
    '                                contents:=True, Scenarios:=True
    ModuleControl.SettingStatusBar False
    Application.ScreenUpdating = True
   
    Unload Me
    
    '결과 분석이 시작되는 부분에서 조금 아래 쪽을 보여주며 마친다.
    Worksheets(RstSheet).Activate
    
    '파일 버전 체크 후 비교값 정의
    Dim Cmp_Value As Long
    
    If ModuleControl.ChkVersion(ActiveWorkbook.Name) = True Then
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
