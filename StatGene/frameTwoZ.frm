VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameTwoZ 
   OleObjectBlob   =   "frameTwoZ.frx":0000
   Caption         =   "따라하기 : μ₁-μ₂에 대한 z-검정"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10905
   StartUpPosition =   1  '소유자 가운데
   TypeInfoVer     =   33
End
Attribute VB_Name = "frameTwoZ"
Attribute VB_Base = "0{84C41D9A-49D1-43BC-9093-E27F6D3A7E4D}{9E7947F3-58C4-4A4B-991D-678AB4B93911}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False


Private Sub Cancel_Click()
    Unload Me
End Sub

Private Sub ChB1_Click()

End Sub

Private Sub Label16_Click()
    frameDEx4.Show
    
End Sub
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim i As Integer
    i = 0
    If Me.ListBox2.ListCount = 0 Then
        Do While i <= Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(i) = True Then
               Me.ListBox2.AddItem Me.ListBox1.List(i)
               Me.ListBox1.RemoveItem (i)
               Me.CB1.Visible = False
               Me.CB2.Visible = True
               Exit Sub
            End If
            i = i + 1
        Loop
    End If


   ElseIf Me.ListBox5.ListCount = 0 Then
        Do While j <= Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(j) = True Then
               Me.ListBox5.AddItem Me.ListBox1.List(j)
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
        Me.ListBox1.AddItem Me.ListBox2.List(0)
        Me.ListBox2.RemoveItem (0)
        Me.CB1.Visible = True
        Me.CB2.Visible = False
    End If
End Sub

Private Sub ListBox5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.ListBox5.ListCount <> 0 Then
        Me.ListBox1.AddItem Me.ListBox5.List(0)
        Me.ListBox5.RemoveItem (0)
        Me.CB1.Visible = True
        Me.CB2.Visible = False
    End If
End Sub



Private Sub OK_Click()

    Dim i As Integer
    i = 0
    If Me.ListBox2.ListCount = 0 Then
        Do While i <= Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(i) = True Then
               Me.ListBox2.AddItem Me.ListBox1.List(i)
               Me.ListBox1.RemoveItem (i)

               Exit Sub
            End If
            i = i + 1
        Loop
    End If
    
    
    Dim j As Integer
    j = 0
    If Me.ListBox5.ListCount = 0 Then
        Do While j <= Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(j) = True Then
               Me.ListBox5.AddItem Me.ListBox1.List(j)
               Me.ListBox1.RemoveItem (j)

               Exit Sub
            End If
            j = j + 1
        Loop
    End If
    
    
End Sub


Private Sub OK1_Click()
                               ''''"_두표본z-검정분석결과_"
    
    Dim choice(6) As Variant                            '넘길 정보는 검정값,신뢰구간,대립가설,모표준편차1.2 5개니까
    Dim dataRange As Range
    Dim i As Integer
    Dim activePt As Long                                '결과 분석이 시작되는 부분을 보여주기 위함
    
    '''
    '''변수를 선택하지 않았을 경우
    '''
    If Me.ListBox2.ListCount + Me.ListBox5.ListCount <> 2 Then
        MsgBox "2개의 변수를 선택해 주시기 바랍니다.", vbExclamation, "HIST"
        Exit Sub
    End If
    '''
    '''public 변수 선언 xlist2, DataSheet, RstSheet, m, k2, n2
    '''
    ReDim xlist2(2)
    xlist2(1) = Me.ListBox2.List(0)
    
    MsgBox xlist2(1), vbExclamation, "HIST"
    xlist2(2) = Me.ListBox5.List(0)
     MsgBox xlist2(2), vbExclamation, "HIST"
    
    DataSheet = ActiveSheet.Name                        'DataSheet : Data가 있는 Sheet 이름
    RstSheet = "_통계분석결과_"                       'RstSheet  : 결과를 보여주는 Sheet 이름
    
    '맨위에 입력
On Error GoTo Err_delete
Dim val3535 As Long '초기위치 저장할 공간'
Dim s3535 As Worksheet
val3535 = 2
    For Each s3535 In ActiveWorkbook.Sheets
        If s3535.Name = RstSheet Then
val3535 = Sheets(RstSheet).Cells(1, 1).Value
End If
Next s3535  '시트가 이미있으면 출력 위치 저장을하고, 없으면 2을 저장한다.

    Set dataRange = ActiveSheet.Cells.CurrentRegion
    m = dataRange.Columns.Count                         'm         : dataSheet에 있는 변수 개수
    
    tmp1 = 2
    ReDim xlist2(tmp1)                                  '변수이름들
    ReDim k2(tmp1)                                      '몇번째열의 변수인지
    ReDim n2(tmp1)                                      '데이타 몇개씩인지
    ReDim tmp(tmp1)
    
     i = 1
        tmp(1) = 0
        tmp(2) = 0
        For j = 1 To m
            If Me.ListBox2.List(0) = ActiveSheet.Cells(1, j) Then
                xlist2(1) = ActiveSheet.Cells(1, j)
                k2(1) = j
                n2(1) = ActiveSheet.Cells(1, j).End(xlDown).row - 1
            '    tmp(i) = tmp(i) + 1
            End If
               If Me.ListBox5.List(0) = ActiveSheet.Cells(1, j) Then
                xlist2(2) = ActiveSheet.Cells(1, j)
                k2(2) = j
                n2(2) = ActiveSheet.Cells(1, j).End(xlDown).row - 1
             '   tmp(i) = tmp(i) + 1
            End If
    Next j
    tmp(1) = 1
    tmp(2) = 1
   
    '''
    ''' 변수명이 같은 경우 - 마지막 열에 있는 변수만 입력되므로 에러처리한다.
    '''
    For i = 1 To tmp1
    If tmp(i) > 1 Then
        MsgBox xlist2(i) & "와 같은 변수명이 있습니다. " & vbCrLf & "변수명을 바꿔주시기 바랍니다.", vbExclamation, "HIST"
        Exit Sub
    End If
    Next i
    
    '''
    '''숫자와 문자가 혼합되어 있을 경우
    '''
    For i = 1 To tmp1
       If TModuleControl.FindingRangeError(xlist2(i)) = True Then
           MsgBox "다음의 분석변수에 문자나 공백이 있습니다." & Chr(10) & _
                    ": " & xlist2(i), vbExclamation, "HIST"
            Exit Sub
        End If
    Next i
            

    '''
    '''검정값을 입력하지 않은 경우
    '''
   ' If IsNumeric(Me.TextBox1.Value) = False Then
    '    MsgBox "사용자 검정값을 입력해 주시기 바랍니다.", vbExclamation, "HIST"
     '   Exit Sub
   ' End If
    
    
    '''
    '''신뢰구간을 잘못 입력한 경우
    '''
  '  If Me.ChB1.Value = True Then
        If IsNumeric(Me.TextBox4.Value) = False Then
            MsgBox "사용자 신뢰구간을 입력해 주시기 바랍니다.", vbExclamation, "HIST"
            Exit Sub
        ElseIf Me.TextBox4.Value < 0 Or Me.TextBox4.Value > 100 Then
            MsgBox "사용자 신뢰구간을 %단위로 입력해 주시기 바랍니다.", vbExclamation, "HIST"
            Exit Sub
        End If
  '  End If
    '''
    ''' 데이타 개수가 한개일 경우
    '''
    If n2(1) = 1 Or n2(2) = 1 Then
        MsgBox "한 개의 데이타로 검정을 시행할 수 없습니다.", vbExclamation, "HIST"
        Exit Sub
    End If

    '''
    '''검정값 선택결과 입력 - choice(1)
    '''
    choice(1) = 1
    'choice(6) = Me.TextBox5.Value
    
    
    '''
    '''신뢰구간 입력 - choice(2)
    '''
    If Me.ChB1.Value = True Then choice(2) = Me.TextBox4.Value
   ' If Me.ChB1.Value = False Then choice(2) = -1
    
    '''
    '''귀무가설 선택결과 입력 - choice(3)
    '''
    If Me.OB5 = True Then choice(3) = 1
    If Me.OB6 = True Then choice(3) = 2
    If Me.OB4 = True Then choice(3) = 3
    '''
    '''검정값 선택결과 입력 - choice(4,5)
    '''
    choice(4) = Me.TextBox3.Value
    choice(5) = Me.TextBox5.Value
    
    
    '''
    '''결과 처리
    '''
    TModuleControl.SettingStatusBar True, "이표본 z-검정중입니다."
    Application.ScreenUpdating = False
    TModulePrint.makeOutputSheet (RstSheet)
    'Worksheets(RstSheet).Unprotect "prophet"
    activePt = Worksheets(RstSheet).Cells(1, 1).Value
    
    TModuleControl.ZTest2 choice, 1
    
    
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
    
    If Worksheets(RstSheet).Cells(1, 1).Value > Cmp_Value Then
        MsgBox "[_통계분석결과_]시트를 거의 모두 사용하였습니다." & vbCrLf & "이 시트의 이름을 바꾸거나 삭제해 주세요", vbExclamation, "HIST"
        Exit Sub
    End If
    
    Worksheets(RstSheet).Cells(activePt + 10, 1).Select
    Worksheets(RstSheet).Cells(activePt + 10, 1).Activate
                            '결과 분석이 시작되는 부분을 보여주며 마친다.
                                
        
  '  Unload Me



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

Private Sub OptionButton1_Click()
  
   Dim myRange As Range
   Dim myArray()
   Dim arrName As Variant
   Dim TempSheet As Worksheet
   Set TempSheet = ActiveCell.Worksheet
   
    ReDim arrName(TempSheet.UsedRange.Columns.Count)
' Reading Data
    For i = 1 To TempSheet.UsedRange.Columns.Count
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
    ReDim myArray(TempSheet.UsedRange.Columns.Count - 1)
    a = 0
   For i = 1 To TempSheet.UsedRange.Columns.Count
   If arrName(i) <> "" Then                     '빈칸제거
   myArray(a) = arrName(i)
   a = a + 1
   
   Else:
   End If
   Next i
   
   
   
   Me.ListBox1.List() = myArray
   
 '  For i = 1 To TempSheet.UsedRange.Columns.Count
 '   rngFirst.Offset(i, 1) = myArray(i - 1)
 ' Next i
  


    For j = 1 To TempSheet.UsedRange.Columns.Count
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
    ReDim myArray(TempSheet.UsedRange.Columns.Count - 1)
    a = 0
   For j = 1 To TempSheet.UsedRange.Columns.Count
   If arrName(j) <> "" Then                     '빈칸제거
   myArray(a) = arrName(j)
   a = a + 1
   
   Else:
   End If
   Next j
   
   
   
   Me.ListBox1.List() = myArray
   
 '  For i = 1 To TempSheet.UsedRange.Columns.Count
 '   rngFirst.Offset(i, 1) = myArray(i - 1)
 ' Next i
  




End Sub

Private Sub UserForm_Click()

End Sub
