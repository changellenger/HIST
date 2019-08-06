VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameOneChi 
   OleObjectBlob   =   "frameOneChi.frx":0000
   Caption         =   "따라하기 : σ²에 대한 Χ²검정"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8535
   StartUpPosition =   1  '소유자 가운데
   TypeInfoVer     =   46
End
Attribute VB_Name = "frameOneChi"
Attribute VB_Base = "0{609917EC-EB46-40BC-80D0-87481267F7E5}{A7A280AA-2CC2-4E3F-99C2-14F8F7D8D5E8}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private Sub Cancel_Click()
    Unload Me
End Sub

Private Sub Label11_Click()

End Sub

Private Sub Label15_Click()
    frameDEx2.Show
    
End Sub

Private Sub Label7_Click()

End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub ListBox2_Click()

End Sub

Private Sub OB2_Click()

End Sub

Private Sub OB3_Click()

End Sub

Private Sub OK_Click()
    
    Dim WarningMsg As String
    Dim choice(3) As Variant
    
    If IsNumeric(TextBox3.Value) = False Then
       MsgBox ("표본의 개수가 올바르지 않습니다.")
       Exit Sub
    ElseIf IsNumeric(TextBox4.Value) = False Then
       MsgBox ("표본의 분산이올바르지 않습니다.")
       Exit Sub
    ElseIf IsNumeric(TextBox1.Value) = False Then
       MsgBox ("모분산의 값이 올바르지 않습니다.")
       Exit Sub
    ElseIf IsNumeric(TextBox2.Value) = False Then
       MsgBox ("신뢰수준이 올바르지 않습니다.")
       Exit Sub
    End If
    t = CDbl(TextBox3.Value)
    s = CDbl(TextBox4.Value)
    r = CDbl(TextBox1.Value)
    L = CDbl(TextBox2.Value)
    If Int(t) - t <> 0 Or t <= 0 Then
       MsgBox ("자연수를 입력하세요.")
       Exit Sub
    End If
    If Int(s) - s <> 0 Or s < 0 Then
       MsgBox ("자연수를 입력하세요.")
       Exit Sub
    End If
    If L <= 0 Or L >= 100 Then
       MsgBox ("신뢰수준의 범위에서 벗어났습니다.")
       Exit Sub
    End If
'''조건확인부탁
  '  If t <= s Or s = 0 Then
  '     MsgBox ("분산이 0인 경우입니다1.")
  '     Exit Sub
   ' End If
    'If r <= s Or s = 0 Then
    '   MsgBox ("분산이 0인 경우입니다2.")
    '   Exit Sub
   ' End If
 
    'PHat = s / t
    'lim1 = t * PHat
    'lim2 = t * (1 - PHat)
    
    'If lim1 < 5 Or lim2 < 5 Then
    '   WarningMsg = "*주의: 표본의 크기가 작습니다."
   ' End If
    zstat = (((t - 1) * s) / r)
    Set resultsheet = OpenOutSheet2("_통계분석결과_", True)
    '''
    '''귀무가설 선택결과 입력 - choice(3)
    '''
    If Me.OB1 = True Then choice(3) = 1
    If Me.OB2 = True Then choice(3) = 2
    If Me.OB3 = True Then choice(3) = 3
    
    '''
    '''
    '''
    '''
    '''
    '''
    RstSheet = "_통계분석결과_"
    
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


    
    'Worksheets(RstSheet).Unprotect "prophet"
    activePt = Worksheets(RstSheet).Cells(1, 1).Value
    
    '''Worksheets(RstSheet).Cells(1, 1).Value = "$A$" & Worksheets(RstSheet).Cells(1, 1).Value
 
    
    
    'resultsheet.Unprotect "prophet"

    ratiotest.resultOneChi zstat, PHat, r, t, s, L, WarningMsg, resultsheet, choice
    'resultsheet.Protect "prophet"
    
    '''Worksheets(RstSheet).Cells(1, 1).Value = Right(Worksheets(RstSheet).Cells(1, 1).Value, Len(Worksheets(RstSheet).Cells(1, 1).Value) - 3)
    
    



    '파일 버전 체크 후 비교값 정의
    Dim Cmp_Value As Long
    
    If PublicModule.ChkVersion1(ActiveWorkbook.Name) = True Then
        Cmp_Value = 1048000
    Else
        Cmp_Value = 65000
    End If
    
    If Worksheets(RstSheet).Cells(1, 1).Value > Cmp_Value Then
        MsgBox "[_통계분석결과_]시트를 거의 모두 사용하였습니다." & vbCrLf & "이 시트의 이름을 바꾸거나 삭제해 주세요", vbExclamation, "HIST"
        Exit Sub
    End If
    
    Worksheets(RstSheet).Activate
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

MsgBox ("프로그램에 문제가 있습니다.")
 'End sub 앞에다 붙인다.

''해석, 에러가 나면 Err_delete로 와서 첫셀이후로 지운다. 만약 첫셀이 2면 시트를 지운다.그리고 에러메시지 출력
'rSTsheet만들기도 전에 에러나는 경우에는 아무 동작도 하지 않고, 에러메시지만 띄운다.
End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub TextBox2_Change()

End Sub

Private Sub UserForm_Click()

End Sub
