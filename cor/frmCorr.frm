VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCorr 
   OleObjectBlob   =   "frmCorr.frx":0000
   Caption         =   "상관분석"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7080
   StartUpPosition =   1  '소유자 가운데
   TypeInfoVer     =   18
End
Attribute VB_Name = "frmCorr"
Attribute VB_Base = "0{DB2617FE-3D63-41C5-8A79-4D74C77F553B}{88404FE6-1331-4E88-9D1C-C27D0924086C}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Sub MoveBtwnListBox(ParentD, FromLNum, ToLNum)

    Dim i As Integer
    i = 0
    Do While i <= ParentD.Controls(FromLNum).ListCount - 1
        If ParentD.Controls(FromLNum).Selected(i) = True Then
           ParentD.Controls(ToLNum).AddItem ParentD.Controls(FromLNum).list(i)
           ParentD.Controls(FromLNum).RemoveItem i
            Exit Do
        End If
        i = i + 1
    Loop

End Sub
Private Sub Cancel_Click()

    Unload Me
    
End Sub
Private Sub CB1_Click()

    MoveBtwnListBox Me, "ListBox1", "ListBox2"

End Sub
Private Sub CB2_Click()

    MoveBtwnListBox Me, "ListBox2", "ListBox1"
    
End Sub
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    MoveBtwnListBox Me, "ListBox1", "ListBox2"
    
End Sub
Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    MoveBtwnListBox Me, "ListBox2", "ListBox1"
    
End Sub
Private Sub OK_Click()

    Dim dataRange As Range
    Dim i As Integer, J As Integer, errSign As Boolean, errTmp As Integer, sheetApproxRowNum
    Dim activePt As Long, ErrString As String                                   '' 결과 분석이 시작되는 부분을 보여주기 위함
    
    '''
    ''' 에러 처리 부분 0: 1행 1열
    '''
    If ActiveSheet.Cells(1, 1) = "" Then
        MsgBox "1행 1열에 변수명이 필요합니다.", vbExclamation                  '' 1행 1열이 비면 각종 오류 발생.
        Exit Sub
    End If
    
    '''
    ''' 에러 처리 부분 1: 변수 선택여부 확인
    '''
    If Me.ListBox2.ListCount = 0 Then
        MsgBox "변수를 선택하지 않았습니다.", vbExclamation
        Exit Sub
    End If

    '''
    ''' 입력받은 정보 정리하기
    '''
    ''' 여기부터 ModeuleControl 에서 선언된 Public 변수
    ''' 여기서 한번만 지정해준다
    ''' sheetRowNum, sheetColNum, DataSheet, RstSheet, xlist, n, m, p
    '''
    If right(ActiveWorkbook.Name, 4) = ".xls" Or right(ActiveWorkbook.Name, 4) = ".XLS" Then
        sheetRowNum = 2 ^ 16            '65536
        sheetColNum = 2 ^ 8             '256
        sheetApproxRowNum = 65000
    Else
        sheetRowNum = 2 ^ 20            '1048576
        sheetColNum = 2 ^ 14            '16384
        sheetApproxRowNum = 1048000
    End If
    
    DataSheet = ActiveSheet.Name                                                '' Data가 있는 Sheet 이름
    rstSheet = "_통계분석결과_"                                                 '' 결과를 보여주는 Sheet 이름
    '출력하는 해당 모듈에 덧 붙일 내용'
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

    
        
    Set dataRange = ActiveSheet.Cells.CurrentRegion
    m = dataRange.Cells(1, 1).End(xlToRight).Column                             '' 전체 독립변수 개수
    
    p = Me.ListBox2.ListCount                                                   '' 선택된 독립변수 개수
    ReDim xlist(p - 1)
    For i = 0 To p - 1
        xlist(i) = ListBox2.list(i)                                             '' 선택된 독립변수 이름
    Next i
    
    N = ModuleControl.FindDataCount(xlist) - 1                                  '' 선택된 변수의 Data개수
    
    '''
    ''' 에러 처리 부분 2: 변수들의 관측수의 대응
    '''
    errSign = False
    For i = 0 To p - 1
        If N <> ModuleControl.FindColDataCount(xlist(i)) Then errSign = True
    Next i
    
    If errSign = True Then
        MsgBox "선택된 항목들간의 관측수가 다릅니다.", vbExclamation, "HIST"
        Exit Sub
    End If
    errSign = False
    
    '''
    '''에러 처리 부분 3: 숫자와 문자가 혼합되어 있을 경우
    '''
    For i = 0 To p - 1
        If ModuleControl.FindingRangeError(xlist(i)) = True Then
            errSign = True
            If ErrString <> "" Then
                ErrString = ErrString & "," & xlist(i)
            Else: ErrString = xlist(i)
            End If
        End If
    Next i
    
    If errSign = True Then
        MsgBox "다음의 분석변수에 문자나 공백이 있습니다." & vbCrLf & vbCrLf & ErrString, vbExclamation, "HIST"
        Exit Sub
    End If
    errSign = False
    
    '''
    ''' 에러 처리 부분 4: 변수명이 같은 경우 - 마지막 열에 있는 변수만 입력되므로 에러처리한다.
    '''
    For i = 1 To p
        errTmp = 0
        For J = 1 To m
            If Me.ListBox2.list(i - 1) = ActiveSheet.Cells(1, J) Then
                errTmp = errTmp + 1
            End If
        Next J
        If errTmp > 1 Then
            MsgBox xlist(i - 1) & vbCrLf & vbCrLf & "위의 분석변수와 같은 변수명이 있습니다. " & vbCrLf & "변수명을 바꿔주시기 바랍니다.", vbExclamation, "HIST"
            Exit Sub
        End If
    Next i
    
    '''
    '''결과 처리
    '''
    ModuleControl.SettingStatusBar True, "상관분석중입니다."
    Application.ScreenUpdating = False
    
    ModulePrint.makeOutputSheet rstSheet
    activePt = Worksheets(rstSheet).Cells(1, 1).Value
    
    ModuleControl.CorrAnal
    ModuleControl.SettingStatusBar False
    Application.ScreenUpdating = True
    Unload Me
    
    Worksheets(rstSheet).Activate
    Worksheets(rstSheet).Cells(activePt, 1).Select
    Worksheets(rstSheet).Cells(activePt, 1).Activate                            '결과 분석이 시작되는 부분을 보여주며 마친다.
    
    Worksheets(rstSheet).Activate
    If Worksheets(rstSheet).Cells(1, 1).Value > sheetApproxRowNum Then
        MsgBox "[_통계분석결과_]시트를 거의 모두 사용하였습니다." & vbCrLf & "이 시트의 이름을 바꾸거나 삭제해 주세요", vbExclamation, "HIST"
        Exit Sub
    End If


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

Private Sub UserForm_Click()

End Sub
