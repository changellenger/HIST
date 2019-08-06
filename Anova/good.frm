VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} good 
   OleObjectBlob   =   "good.frx":0000
   Caption         =   "적합도 검정"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3840
   StartUpPosition =   1  '소유자 가운데
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

Private Sub ok_btn_Click()                                          ''''"_범주형자료분석결과_"
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
  '''에러 체크
  Set NumRn = temp.Offset(1, 0).Resize(temp.Rows.count - 1, temp.Columns.count)
  If FindingRangeError(NumRn) = True Then
        MsgBox "분석변수에 문자나 공백이 있습니다.", vbExclamation, "HIST"
        Exit Sub
  End If

  If column_btn.Value = True Then
    Set NumRn = temp.Offset(1, 0).Resize(temp.Rows.count - 1, temp.Columns.count)
    If FindingRangeError(NumRn) = True Then
          MsgBox "분석변수에 문자나 공백이 있습니다.", vbExclamation, "HIST"
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
          MsgBox "분석변수에 문자나 공백이 있습니다.", vbExclamation, "HIST"
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
   Set resultsheet = OpenOutSheet("_통계분석결과_", True)
   
   '''
    '''
    '''
    RstSheet = "_통계분석결과_"
    '출력하는 해당 모듈에 덧 붙일 내용'
'맨위에 입력
On Error GoTo Err_delete
Dim val3535 As Long '초기위치 저장할 공간'
Dim s3535 As Worksheet
val3535 = 2
    For Each s3535 In ActiveWorkbook.Sheets
        If s3535.name = RstSheet Then
val3535 = Sheets(RstSheet).Cells(1, 1).Value
End If
Next s3535  '시트가 이미있으면 출력 위치 저장을하고, 없으면 2을 저장한다.
   ' Worksheets(RstSheet).Unprotect "prophet"
    activePt = Worksheets(RstSheet).Cells(1, 1).Value
   
    
   'resultsheet.Unprotect "prophet"
   Good_result.gresult chis, VarName, raw, Expect, total1, cnt, resultsheet
   'resultsheet.Protect Password:="prophet", DrawingObjects:=False, contents:=True, Scenarios:=True
   'resultsheet.Unprotect "prophet"
    ''''ratiotest.ratioresult zstat, phat, r, T, s, L, WarningMsg, resultsheet
    '''' 위의 줄은 무슨 의미인줄 모르겠어서 일단 정지 시킴   각 변수에 입력되는 값이 없음.
   'resultsheet.Protect "prophet"
    
    '''Worksheets(RstSheet).Cells(1, 1).Value = Right(Worksheets(RstSheet).Cells(1, 1).Value, Len(Worksheets(RstSheet).Cells(1, 1).Value) - 3)



    Worksheets(RstSheet).Activate

    '파일 버전 체크 후 비교값 정의
    Dim Cmp_Value As Long
    
    If ModuleControl.ChkVersion(ActiveWorkbook.name) = True Then
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

MsgBox ("프로그램에 문제가 있습니다.")
 'End sub 앞에다 붙인다.

''해석, 에러가 나면 Err_delete로 와서 첫셀이후로 지운다. 만약 첫셀이 2면 시트를 지운다.그리고 에러메시지 출력
'rSTsheet만들기도 전에 에러나는 경우에는 아무 동작도 하지 않고, 에러메시지만 띄운다.
End Sub

Private Sub UserForm_Terminate()
     Unload Me
End Sub
