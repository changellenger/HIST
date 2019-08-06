VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameCrfre 
   OleObjectBlob   =   "frameCrfre.frx":0000
   Caption         =   "교차분석"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4140
   StartUpPosition =   1  '소유자 가운데
   TypeInfoVer     =   71
End
Attribute VB_Name = "frameCrfre"
Attribute VB_Base = "0{DD0ADB77-4F06-4BB9-AD52-68C01514899B}{36FABEC2-E16A-490A-9139-89D13AFA0C56}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False





Private Sub BtnOK_Click()                               ''''"_범주형자료분석결과_"
  Dim TempSheet As Worksheet
  Dim resultsheet As Worksheet
  Dim temp As Range
  Dim tm As Range
  Dim rname() As String
  Dim cname() As String
  Dim rtotal() As Long
  Dim ctotal() As Long
  Dim Expect() As Double
  Dim total As Long
  Dim chisq As Double
  Set TempSheet = ActiveCell.Worksheet
  Set temp = TempSheet.Cells.CurrentRegion
  c = temp.Columns.count - 1
  r = temp.Rows.count - 1
      '''에러 체크
    Set tm = temp.Offset(1, 1).Resize(r, c)
    If FindingRangeError2(tm) = True Then
          MsgBox "분석변수에 문자나 공백이 있습니다.", vbExclamation, "HIST"
          Exit Sub
    End If
  Set tm = temp
  ReDim rname(1 To r)
  ReDim cname(1 To c)
  ReDim rtotal(1 To r)
  ReDim ctotal(1 To c)
  Set tp = temp.Columns(1)
  For i = 2 To r + 1
      rname(i - 1) = tp.Cells(i, 1)
  Next i
  Set tp = temp.Rows(1)
  For i = 2 To c + 1
      cname(i - 1) = tp.Cells(1, i)
  Next i
  Set temp = temp.Offset(1, 1)
  total = Application.sum(temp)
  For i = 1 To r
      rtotal(i) = Application.sum(temp.Rows(i))
  Next i
  For i = 1 To c
      ctotal(i) = Application.sum(temp.Columns(i))
  Next i
  ReDim Expect(1 To r, 1 To c)
  chisq = 0
  For i = 1 To r
      For J = 1 To c
          Expect(i, J) = rtotal(i) * ctotal(J) / total
          chisq = chisq + (temp.Cells(i, J).Value - Expect(i, J)) ^ 2 / Expect(i, J)
      Next J
  Next i
  Set resultsheet = OpenOutSheet2("_통계분석결과_", True)
  
  
    '''
    '''
    '''
    rstSheet = "_통계분석결과_"
    
    '맨위에 입력
On Error GoTo Err_delete
Dim val3535 As Long '초기위치 저장할 공간'
Dim s3535 As Worksheet
val3535 = 2
    For Each s3535 In ActiveWorkbook.Sheets
        If s3535.Name = rstSheet Then
val3535 = Sheets(rstSheet).Cells(1, 1).Value
End If
Next s3535  '시트가 이미있으면 출력 위치 저장을하고, 없으면 2을 저장한다.
    
    'Worksheets(RstSheet).Unprotect "prophet"
    activePt = Worksheets(rstSheet).Cells(1, 1).Value
    

  
  'resultsheet.Unprotect "prophet"
  Conti_Result.cResult r, c, temp, Expect, rtotal, ctotal, chisq, rname, cname, resultsheet
  'resultsheet.Protect Password:="prophet", DrawingObjects:=False, contents:=True, Scenarios:=True
    
    
    
    '''Worksheets(RstSheet).Cells(1, 1).Value = Right(Worksheets(RstSheet).Cells(1, 1).Value, Len(Worksheets(RstSheet).Cells(1, 1).Value) - 3)

    


    Worksheets(rstSheet).Activate

    '파일 버전 체크 후 비교값 정의
    Dim Cmp_Value As Long
    
    If PublicModule.ChkVersion(ActiveWorkbook.Name) = True Then
        Cmp_Value = 1048000
    Else
        Cmp_Value = 65000
    End If
    
    If Worksheets(rstSheet).Cells(1, 1).Value > Cmp_Value Then
        MsgBox "[_통계분석결과_]시트를 거의 모두 사용하였습니다." & vbCrLf & "이 시트의 이름을 바꾸거나 삭제해 주세요", vbExclamation, "HIST"
        Exit Sub
    End If

    Worksheets(rstSheet).Cells(activePt + 10, 1).Select
    Worksheets(rstSheet).Cells(activePt + 10, 1).Activate
                            '결과 분석이 시작되는 부분을 보여주며 마친다.
                            
    Unload Me
    
'맨뒤에 붙이기
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

MsgBox ("프로그램에 문제가 있습니다.")
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
   
    ReDim arrName(TempSheet.UsedRange.Columns.count)
' Reading Data
    For i = 1 To TempSheet.UsedRange.Columns.count
        arrName(i) = TempSheet.Cells(1, i)
    Next i
   
   Me.Listbox1.Clear

    ReDim myArray(TempSheet.UsedRange.Columns.count - 1)
    a = 0
   For i = 1 To TempSheet.UsedRange.Columns.count
   If arrName(i) <> "" Then                     '빈칸제거
   myArray(a) = arrName(i)
   a = a + 1
   
   Else:
   End If
   Next i
  
   Me.Listbox1.list() = myArray



End Sub

Private Sub UserForm_Click()

End Sub
