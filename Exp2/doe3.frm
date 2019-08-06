VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} doe3 
   OleObjectBlob   =   "doe3.frx":0000
   Caption         =   "요인설계분석"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9420
   StartUpPosition =   1  '소유자 가운데
   TypeInfoVer     =   238
End
Attribute VB_Name = "doe3"
Attribute VB_Base = "0{57E8107D-CF87-4994-9BF0-C2E869BFEAF5}{482E5331-94C0-4194-924E-E9EDCA2AC989}"
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
  MoveBtwnListBox Me, "ListBox1", "ListBox2"

End Sub

Private Sub CB2_Click()
  MoveBtwnListBox Me, "ListBox2", "ListBox1"
End Sub



Private Sub CheckBox1_Click()

End Sub

Private Sub ComboBox1_Change()      ' 콤보박스 바꿨을대 리스트박스 수정
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

    ReDim myArray(TempSheet.UsedRange.Columns.count - 5)
    a = 0
   For i = 4 To TempSheet.UsedRange.Columns.count
   If arrName(i) <> Me.ComboBox1.value Then
   If arrName(i) <> "" Then                     '빈칸제거
   myArray(a) = arrName(i)
   a = a + 1
   
   Else:
   End If
   End If
  
   Next i

   Me.ListBox3.list = myArray
   
    element = myArray
    'element = Array("가", "나", "다", "라")
    combinationModule.comb (element)

 
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
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
End Sub

Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.ListBox2.ListCount <> 0 Then
        Me.ListBox1.AddItem Me.ListBox2.list(0)
        Me.ListBox2.RemoveItem (0)
       
    End If
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
  
   
End Sub

Private Sub ToggleButton1_Click()

                           '넘길 정보는 검정값,신뢰구간,대립가설 3개니까
    Dim dataRange As Range
    Dim i As Integer
    Dim activePt As Long                                '결과 분석이 시작되는 부분을 보여주기 위함
    Dim rng As Range
    Dim xlist()
    Dim xstrlist()
    Dim nol As Integer
    Dim k1() As Integer
    Dim rssheet As Worksheet
    Dim myResultSheet As Worksheet
    
      nol = Me.ListBox2.ListCount 'ListBox2에 있는 변수갯수
      
    rinterface.RRun "require (FrF2)"
    rinterface.RRun "require (qualityTools)"
    'RInterface.StartRServer
    
    rinterface.StartRServer

    
    If nol = 0 Then
        MsgBox "변수를 선택해 주시기 바랍니다.", vbExclamation, "HIST"
'    ElseIf nol >= 21 Then
'
'        MxgBox "분석변수는 20개 이하로 지정해야 합니다.", vbExclamation, "HIST"
        Exit Sub
'    Else
    
    End If
    

    
    
    '''
    '''public 변수 선언 xlist, DataSheet, RstSheet, m, k1, n
    '''

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


  
    ReDim k1(nol) As Integer
     ReDim xstrlist(nol - 1)
    ReDim xlist(nol)                       'ListBox2에 있는 List(j)번째 변수명을 xlist(j)에 할당
    
    For j = 0 To nol - 1
    xstrlist(j) = ListBox2.list(j)
    
        xlist(j) = ListBox2.list(j)
    Next j
        xlist(j) = doe3.ComboBox1.value
    
    Set dataRange = ActiveSheet.Cells.CurrentRegion
    m = dataRange.Columns.count                         'm         : dataSheet에 있는 변수 개수
    
    lastColumn = ActiveCell.Worksheet.UsedRange.Columns.count - 1
    
    
    tmp = 0
        For j = 0 To nol
            For i = 1 To m
                If xlist(j) = ActiveSheet.Cells(1, i).value Then
                    k1(j) = i  'k1                                 : k1 : 선택된 변수가 몇번째 열에 있는지
               
                End If
            Next i
            
            n = ActiveSheet.Cells(1, 1).End(xlDown).row - 1    'n         : 선택된 변수의 데이타 갯수
        Next j



    Dim checkarray As String
    Dim anovastr As String
    Dim anovastr2 As String
    
    Dim temp As String
    
    Dim cbindstr As String
    
      For j = 0 To nol - 1
        checkarray = xlist(j)
'        rinterface.PutArray checkarray, Range(Cells(2, k1(j)), Cells(n + 1, k1(j)))
            If j = 0 Then
                cbindstr = checkarray
                anovastr = checkarray
                anovastr2 = Left(checkarray, 1)
                Else
                cbindstr = cbindstr & "," & checkarray
                anovastr = anovastr & "+" & checkarray
                anovastr2 = anovastr2 & "+" & Left(checkarray, 1)
            End If
         
    Next j
    checkarray = xlist(j)
   '  rinterface.PutArray checkarray, Range(Cells(2, k1(j)), Cells(N + 1, k1(j)))
   '  rinterface.PutDataframe "Response2", Range(Cells(1, k1(j)), Cells(N + 1, k1(j)))
     
     Dim strc As String
     
     For q = 1 To n
     If q = 1 Then
     strc = Cells(q + 1, k1(j)).value
     Else
     
     strc = strc & ", " & Cells(q + 1, k1(j)).value
     End If
     
     Next q
    ' MsgBox q - 1
     
     
     
    strc = "Response = c(" & strc & ")"
    ' MsgBox strc
    '  Rinterface.RRun "Responset = c(580, 1090, 1392, 568, 1087, 1380, 570, 1085, 1386, 550, 1070, 1328, 530, 1035, 1312, 579)"
     rinterface.RRun strc
     rinterface.RRun "response(arrayfrac) = Response"
    'Rinterface.RRun "install.packages (" & Chr(34) & "FrF2" & Chr(34) & ")" ' : R 패키지 필요없음:
    
    'Rinterface.RRun "install.packages (" & Chr(34) & "qualityTools" & Chr(34) & ")" ' : R 패키지 필요없음:

    
    'Rinterface.rrun "arraytest<-cbind(AP,BP,CP,Response)" 'ok
    'MsgBox "arraytest<-cbind(AP,BP,CP,Response)"
   ' MsgBox "arraytest<-cbind(" & cbindstr & "," & xlist(j) & " )"
    'rinterface.RRun "arraytest<-cbind(" & cbindstr & "," & checkarray & ")"
   
    'temp = combinationModule.combstr(xstrlist)
    MsgBox anovastr2
    rinterface.RRun "WQ<-as.data.frame(arrayfrac)" 'ok
    rinterface.RRun "lm.5 =lm(" & checkarray & " ~" & anovastr & ", data = arrayfrac)"
   ' rinterface.RRun "lm.3 =lm(" & checkarray & " ~ " & temp & ", data = arrayfrac)"
    j = Me.ListBox2.ListCount
    'MsgBox j & " N"
    
    rinterface.RRun "summary(lm.5)" 'ok
    'MsgBox anovastr
    rinterface.RRun "AnovaModeQ <- aov(lm(" & checkarray & " ~" & anovastr & ", data = arrayfrac))"
    rinterface.RRun "anova(AnovaModeQ)"     'ok
    rinterface.RRun "ogx <- anova(AnovaModeQ)"
    rinterface.RRun "cgx <-as.data.frame(ogx)"
    
    
    
    Application.ScreenUpdating = False
    Dim stname As String
    Dim lastCol, lastRow As Integer
    
    stname = Me.Caption
    PublicModule.OpenOutSheet stname, True
    Worksheets(stname).Activate
   
   
     
    rinterface.GetDataframe "cgx", Range("B2"), True
    
     ActiveSheet.Cells(2, 2).value = "분산분석 결과"
     ActiveSheet.Cells(2, 2).Font.Bold = True
     ActiveSheet.Cells(2, 2).Interior.Color = RGB(220, 238, 130)
     ActiveSheet.Cells(2, 2).Cells.ColumnWidth = 20
    
    lastCol = ActiveCell.Worksheet.UsedRange.Columns.count
    lastRow = ActiveCell.Worksheet.UsedRange.rows.count
    
   ' MsgBox " col:" & lastCol & " row: " & lastRow
    
    
    
  
    
    
    Range(Cells(2, 2), Cells(lastRow + 1, 2)).Borders(xlEdgeLeft).LineStyle = xlContinuous '셀의 왼쪽 테두리 설정
    Range(Cells(2, 2), Cells(lastRow + 1, 2)).Borders(xlEdgeLeft).Color = RGB(34, 116, 34)
    Range(Cells(2, 2), Cells(lastRow + 1, 2)).Borders(xlEdgeLeft).Weight = 3
 
 
    
    Range(Cells(2, lastCol + 1), Cells(lastRow + 1, lastCol + 1)).Borders(xlEdgeRight).LineStyle = xlContinuous '셀의 오른쪽 테두리 설정
    Range(Cells(2, lastCol + 1), Cells(lastRow + 1, lastCol + 1)).Borders(xlEdgeRight).Color = RGB(34, 116, 34)
    Range(Cells(2, lastCol + 1), Cells(lastRow + 1, lastCol + 1)).Borders(xlEdgeRight).Weight = 3
 
    

    Range(Cells(2, 2), Cells(2, lastCol + 1)).Borders(xlEdgeTop).LineStyle = xlContinuous '셀의 위쪽 테두리 설정
    Range(Cells(2, 2), Cells(2, lastCol + 1)).Borders(xlEdgeTop).Color = RGB(34, 116, 34)
    Range(Cells(2, 2), Cells(2, lastCol + 1)).Borders(xlEdgeTop).Weight = 3
 
    Range(Cells(lastRow + 1, 2), Cells(lastRow + 1, lastCol + 1)).Borders(xlEdgeBottom).LineStyle = xlContinuous '셀의 아래쪽 테두리 설정
    Range(Cells(lastRow + 1, 2), Cells(lastRow + 1, lastCol + 1)).Borders(xlEdgeBottom).Color = RGB(34, 116, 34)
    Range(Cells(lastRow + 1, 2), Cells(lastRow + 1, lastCol + 1)).Borders(xlEdgeBottom).Weight = 3
    

    Range(Cells(2, 2), Cells(2, lastCol + 1)).Borders(xlEdgeBottom).LineStyle = xlContinuous '셀의 아래쪽 테두리 설정
    Range(Cells(2, 2), Cells(2, lastCol + 1)).Borders(xlEdgeBottom).Color = RGB(34, 116, 34)
    Range(Cells(2, 2), Cells(2, lastCol + 1)).Borders(xlEdgeBottom).Weight = 3
    
        
    Range(Cells(2, 2), Cells(lastRow + 1, 2)).Borders(xlEdgeRight).LineStyle = xlContinuous '셀의 오른쪽 테두리 설정
    Range(Cells(2, 2), Cells(lastRow + 1, 2)).Borders(xlEdgeRight).Color = RGB(34, 116, 34)
    Range(Cells(2, 2), Cells(lastRow + 1, 2)).Borders(xlEdgeRight).Weight = 3
 
 
    ActiveSheet.Cells(lastRow + 3, 2).value = "잔차 그래프"
     ActiveSheet.Cells(lastRow + 3, 2).Font.Bold = True
     ActiveSheet.Cells(lastRow + 3, 2).Interior.Color = RGB(220, 238, 130)
     ActiveSheet.Cells(lastRow + 3, 2).Cells.ColumnWidth = 20


'#잔차그래프1번- 대적합치 ok
rinterface.RRun "plot(residuals(AnovaModeQ) ~ fitted(AnovaModeQ), xlab= " & Chr(34) & " 적합치 " & Chr(34) & ", ylab= " & Chr(34) & " 잔차 " & Chr(34) & " ,main= " & Chr(34) & " 대 적합치 " & Chr(34) & ")"
rinterface.RRun "abline(h=0,lty=1,col= " & Chr(34) & " red " & Chr(34) & " )"
rinterface.InsertCurrentRPlot Range(Cells(lastRow + 4, 2), Cells(lastRow + 4, 2)), widthrescale:=0.6, heightrescale:=0.6, closergraph:=True
'#잔차그래프2번-정규확률도 ok
rinterface.RRun "qqnorm(resid(AnovaModeQ),xlab=" & Chr(34) & " 잔차 " & Chr(34) & ", ylab=" & Chr(34) & " 백분율 " & Chr(34) & ", main=" & Chr(34) & " 정규확률도 " & Chr(34) & ")"
rinterface.RRun "qqline(resid(AnovaModeQ),lty=1,col=" & Chr(34) & " red " & Chr(34) & ")"
rinterface.InsertCurrentRPlot Range(Cells(lastRow + 4, 7), Cells(lastRow + 4, 8)), widthrescale:=0.6, heightrescale:=0.6, closergraph:=True
'#잔차그래프3번-히스토그램 ok
rinterface.RRun "hist(resid(AnovaModeQ), breaks= 9, xlab= " & Chr(34) & " 잔차 " & Chr(34) & ",ylab= " & Chr(34) & " 빈도 " & Chr(34) & ", main= " & Chr(34) & " 잔차 히스토그램 " & Chr(34) & ", border= " & Chr(34) & " black " & Chr(34) & ", col= " & Chr(34) & " sky blue " & Chr(34) & ")"
rinterface.RRun "lines(c(min(AnovaModeQ$breaks), AnovaModeQ$mids, mas(AnovaModeQ$breaks)), c(0,AnovaModeQ$counts,0),type = " & Chr(34) & " l " & Chr(34) & ")" '에러메세지 출력
rinterface.RRun "lines(density(AnovaModeQ))"
rinterface.InsertCurrentRPlot Range(Cells(lastRow + 4, 13), Cells(lastRow + 4, 13)), widthrescale:=0.6, heightrescale:=0.6, closergraph:=True

ActiveSheet.Cells(lastRow + 33, 2).value = "표준화된 효과의 그래프"
ActiveSheet.Cells(lastRow + 33, 2).Font.Bold = True
ActiveSheet.Cells(lastRow + 33, 2).Interior.Color = RGB(220, 238, 130)
ActiveSheet.Cells(lastRow + 33, 2).Cells.ColumnWidth = 20

rinterface.RRun "paretoPlot(arrayfrac, main = paste(" & Chr(34) & " 표준화된 효과의 Pareto차트 " & Chr(34) & ") )"
rinterface.InsertCurrentRPlot Range(Cells(lastRow + 34, 2), Cells(lastRow + 34, 2)), widthrescale:=0.6, heightrescale:=0.6, closergraph:=True

rinterface.RRun "normalPlot(arrayfrac, main = paste(" & Chr(34) & " 표준화 효과의 정규확률도 " & Chr(34) & ") )"
rinterface.InsertCurrentRPlot Range(Cells(lastRow + 34, 7), Cells(lastRow + 34, 7)), widthrescale:=0.6, heightrescale:=0.6, closergraph:=True

Unload Me

End Sub

Private Sub UserForm_Click()

End Sub
