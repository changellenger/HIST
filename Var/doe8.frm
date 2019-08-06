VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} doe8 
   OleObjectBlob   =   "doe8.frx":0000
   Caption         =   "요인설계분석"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6705
   StartUpPosition =   1  '소유자 가운데
   TypeInfoVer     =   197
End
Attribute VB_Name = "doe8"
Attribute VB_Base = "0{7A3AC40B-F261-484C-AA83-543ABC62A345}{05882093-785B-4960-85EE-1159FD7B6531}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private Sub btnCancel_Click()
Unload Me
End Sub

Private Sub CB1_Click()
Dim I As Integer
    I = 0
    If Me.ListBox2.ListCount = 0 Then
        Do While I <= Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(I) = True Then
               Me.ListBox2.AddItem Me.ListBox1.List(I)
               Me.ListBox1.RemoveItem (I)
               Me.CB1.Visible = False
               Me.CB2.Visible = True
               Exit Sub
            End If
            I = I + 1
        Loop
    End If
End Sub

Private Sub CB2_Click()
If Me.ListBox2.ListCount <> 0 Then
        Me.ListBox1.AddItem ListBox2.List(0)
        Me.ListBox2.RemoveItem (0)
        Me.CB1.Visible = True
        Me.CB2.Visible = False
    End If
End Sub



Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim I As Integer
    I = 0
    If Me.ListBox2.ListCount = 0 Then
        Do While I <= Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(I) = True Then
               Me.ListBox2.AddItem Me.ListBox1.List(I)
               Me.ListBox1.RemoveItem (I)
               Me.CB1.Visible = False
               Me.CB2.Visible = True
               Exit Sub
            End If
            I = I + 1
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

Private Sub OptionButton1_Click()
   Dim myRange As Range
   Dim MyArray()
   Dim arrName As Variant
   Dim TempSheet As Worksheet
   Set TempSheet = ActiveCell.Worksheet
   
    ReDim arrName(TempSheet.UsedRange.Columns.count)
' Reading Data
    For I = 1 To TempSheet.UsedRange.Columns.count
        arrName(I) = TempSheet.Cells(1, I)
    Next I
   
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
    ReDim MyArray(TempSheet.UsedRange.Columns.count - 1)
    a = 0
   For I = 1 To TempSheet.UsedRange.Columns.count
   If arrName(I) <> "" Then                     '빈칸제거
   MyArray(a) = arrName(I)
   a = a + 1
   
   Else:
   End If
   Next I
   
   
   
   Me.ListBox1.List() = MyArray
   
 '  For i = 1 To TempSheet.UsedRange.Columns.Count
 '   rngFirst.Offset(i, 1) = myArray(i - 1)
 ' Next i
  
   
End Sub

Private Sub ToggleButton1_Click()
                           '넘길 정보는 검정값,신뢰구간,대립가설 3개니까
    Dim dataRange As Range
    Dim I As Integer
    Dim activePt As Long                                '결과 분석이 시작되는 부분을 보여주기 위함
    Dim rng As Range
    
    
    '''
    '''변수를 선택하지 않았을 경우
    '''
    If Me.ListBox2.ListCount = 0 Then
        MsgBox "변수를 선택해 주시기 바랍니다.", vbExclamation, "HIST"
        Exit Sub
    End If
    
    '''
    '''public 변수 선언 xlist, DataSheet, RstSheet, m, k1, n
    '''
    xlist = Me.ListBox2.List(0)
    DataSheet = ActiveSheet.name                        'DataSheet : Data가 있는 Sheet 이름
    RstSheet = "_통계분석결과_"                       'RstSheet  : 결과를 보여주는 Sheet 이름
    
    
    
    '맨위에 입력
'On Error GoTo Err_delete
Dim val3535 As Long '초기위치 저장할 공간'
Dim s3535 As Worksheet
val3535 = 2
    For Each s3535 In ActiveWorkbook.Sheets
        If s3535.name = RstSheet Then
val3535 = Sheets(RstSheet).Cells(1, 1).value
End If
Next s3535  '시트가 이미있으면 출력 위치 저장을하고, 없으면 2을 저장한다.


    
    Set dataRange = ActiveSheet.Cells.CurrentRegion
    m = dataRange.Columns.count                         'm         : dataSheet에 있는 변수 개수
    
    tmp = 0
    For I = 1 To m
        If xlist = ActiveSheet.Cells(1, I) Then
            k1 = I  'k1                                 : k1 : 선택된 변수가 몇번째 열에 있는지
            tmp = tmp + 1
        End If
    Next I
    
    n = ActiveSheet.Cells(1, k1).End(xlDown).Row - 1    'n         : 선택된 변수의 데이타 갯수
    
   ' rng = Range(Cells(2, k1), Cells(N + 1, k1))
    
    
    

    '''
    ''' 변수명이 같은 경우 - 마지막 열에 있는 변수만 입력되므로 에러처리한다.
    '''
    If tmp > 1 Then
        MsgBox xlist & "와 같은 변수명이 있습니다. " & vbCrLf & "변수명을 바꿔주시기 바랍니다.", vbExclamation, "HIST"
        Exit Sub
    End If

Rinterface.StartRServer
    'Rinterface.PutArray "Design.5", Range(rng)
    'Rinterface.RRun "install.packages (" & Chr(34) & "" & Chr(34) & ")"  : R 패키지 필요없음:
    'Rinterface.RRun "require ()"
    
    
Rinterface.PutArray "Design.5", Range(Cells(2, k1), Cells(n + 1, k1))

Rinterface.RRun "AnovaModel.5 <- aov(lm(C ~ A+B+A*B, data=Design.5))"
Rinterface.RRun "anova(Anovamodel.5)"
Rinterface.RRun "par(mfrow = c(2, 2))"
Rinterface.RRun "class (par(mfrow = c(2, 2)))"
Rinterface.RRun "plot(residuals(AnovaModel.5) ~ fitted(AnovaModel.5), xlab=" & Chr(34) & "적합치" & Chr(34) & ", ylab=" & Chr(34) & "잔차" & Chr(34) & ",main=" & Chr(34) & "대 적합치" & Chr(34) & ")"
Rinterface.RRun "abline(h=0,lty=1,col=" & Chr(34) & "red" & Chr(34) & ")"
Rinterface.RRun "qqnorm(resid(AnovaModel.5),xlab=" & Chr(34) & "잔차" & Chr(34) & ", ylab=" & Chr(34) & "백분율" & Chr(34) & ", main=" & Chr(34) & "정규확률도" & Chr(34) & ")"
Rinterface.RRun "qqline(resid(AnovaModel.5),lty=1,col=" & Chr(34) & "red" & Chr(34) & ")"
Rinterface.RRun "hist(resid(AnovaModel.5), breaks=9, xlab=" & Chr(34) & "잔차" & Chr(34) & ", ylab=" & Chr(34) & "빈도" & Chr(34) & ", main=" & Chr(34) & "잔차 히스토그램" & Chr(34) & ", border=" & Chr(34) & "blue" & Chr(34) & ", col=" & Chr(34) & "green" & Chr(34) & ")"
Rinterface.RRun "lines(c(min(AnovaModel.5$breaks), AnovaModel.5$mids, mas(AnovaModel.5$breaks)), c(0,AnovaModel.5$counts,0),type = " & Chr(34) & "l" & Chr(34) & ")"
Rinterface.RRun "lines(density(AnovaModel.5))"


End Sub

Private Sub UserForm_Click()

End Sub
