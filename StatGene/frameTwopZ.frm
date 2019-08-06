VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameTwopZ 
   OleObjectBlob   =   "frameTwopZ.frx":0000
   Caption         =   "따라하기 : p₁-p₂에 대한 z-검정"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10980
   StartUpPosition =   1  '소유자 가운데
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


Private Sub BtnOK_Click()                                       '''"_모비율검정결과_"
 
    Dim choice(3) As Variant
    
    If IsNumeric(trial1.Value) = False Then
       MsgBox ("집단1의 시행횟수가 올바르지 않습니다.")
       Exit Sub
    ElseIf IsNumeric(success1.Value) = False Then
       MsgBox ("집단1의 성공횟수가 올바르지 않습니다.")
       Exit Sub
    ElseIf IsNumeric(trial2.Value) = False Then
       MsgBox ("집단2의 시행횟수가 올바르지 않습니다.")
       Exit Sub
    ElseIf IsNumeric(success2.Value) = False Then
       MsgBox ("집단2의 성공횟수가 올바르지 않습니다.")
       Exit Sub
     ElseIf IsNumeric(level.Value) = False Then
       MsgBox ("신뢰수준이 올바르지 않습니다.")
       Exit Sub
    End If
    t1 = CDbl(trial1.Value)
    t2 = CDbl(trial2.Value)
    s1 = CDbl(success1.Value)
    s2 = CDbl(success2.Value)
    L = CDbl(level.Value)
    If Int(t1) - t1 <> 0 Or t1 <= 0 Then
       MsgBox ("자연수를 입력하세요")
       Exit Sub
    End If
    
    If Int(s1) - s1 <> 0 Or s1 < 0 Then
       MsgBox ("자연수를 입력하세요")
       Exit Sub
    End If
    If L <= 0 Or L >= 100 Then
       MsgBox ("신뢰수준의 범위에서 벗어났습니다.")
       Exit Sub
    End If
    If t1 < s1 Then
       MsgBox ("입력형식이 맞지 않습니다.")
       Exit Sub
    End If
    If Int(t2) - t2 <> 0 Or t2 <= 0 Then
       MsgBox ("자연수를 입력하세요")
       Exit Sub
    End If
    If Int(s2) - s2 <> 0 Or s2 < 0 Then
       MsgBox ("자연수를 입력하세요")
       Exit Sub
    End If
    
    If t2 < s2 Then
       MsgBox ("입력형식이 맞지 않습니다.")
       Exit Sub
    End If
    
    If t1 = s1 And t2 = s2 Then
       MsgBox ("분산이 0인 경우입니다.")
       Exit Sub
    End If
    
    If s1 = 0 And s2 = 0 Then
       MsgBox ("분산이 0인 경우입니다.")
       Exit Sub
    End If
        
    PHat1 = s1 / t1
    PHat2 = s2 / t2
    PHat = (s1 + s2) / (t1 + t2)
    zstat = (PHat1 - PHat2) / Sqr((PHat - PHat ^ 2)) / Sqr((1 / trial1.Value + 1 / trial2.Value))
    Set resultsheet = OpenOutSheet2("_통계분석결과_", True)
    
    '''
    '''
    
    '''
    '''귀무가설 선택결과 입력 - choice(3)
    '''
    If Me.OB4 = True Then choice(3) = 1
    If Me.OB5 = True Then choice(3) = 2
    If Me.OB6 = True Then choice(3) = 3
    
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
    ratiotest.ratio2result PHat1, PHat2, zstat, t1, t2, s1, s2, L, resultsheet, choice
    'resultsheet.Protect "prophet"
    ''''Worksheets(RstSheet).Cells(1, 1).Value = Right(Worksheets(RstSheet).Cells(1, 1).Value, Len(Worksheets(RstSheet).Cells(1, 1).Value) - 3)
    
    
    'Worksheets(RstSheet).Protect Password:="prophet", DrawingObjects:=False, _
                                    contents:=True, Scenarios:=True             '''
    


    Worksheets(RstSheet).Activate

    '파일 버전 체크 후 비교값 정의,
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
