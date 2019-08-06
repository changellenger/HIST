VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} doe5 
   OleObjectBlob   =   "doe5.frx":0000
   Caption         =   "교호작용도"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4635
   StartUpPosition =   1  '소유자 가운데
   TypeInfoVer     =   74
End
Attribute VB_Name = "doe5"
Attribute VB_Base = "0{58688D27-BDD2-47AC-9BC3-849A381E226E}{4250754F-7B0D-4EF3-9A4C-281CE07C1A16}"
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
 Me.ListBox2.Clear


    ReDim myArray(TempSheet.UsedRange.Columns.count - 2)
    a = 0
   For i = 1 To TempSheet.UsedRange.Columns.count
   If arrName(i) <> Me.ComboBox1.value Then
   If arrName(i) <> "" Then                     '빈칸제거
   myArray(a) = arrName(i)
   a = a + 1
   
   Else:
   End If
   End If
   Next i
      
   Me.ListBox1.list() = myArray
     
   
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
    
      nol = Me.ListBox2.ListCount 'ListBox2에 있는 변수갯수
      
    rinterface.RRun "require (FrF2)"
    rinterface.RRun "require (qualityTools)"
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
        xlist(j) = doe5.ComboBox1.value
    
    Set dataRange = ActiveSheet.Cells.CurrentRegion
    m = dataRange.Columns.count                         'm         : dataSheet에 있는 변수 개수
    
    tmp = 0
        For j = 0 To nol
            For i = 1 To m
                If xlist(j) = ActiveSheet.Cells(1, i) Then
                    k1(j) = i  'k1                                 : k1 : 선택된 변수가 몇번째 열에 있는지
               
                End If
            Next i
            
            n = ActiveSheet.Cells(1, k1(0)).End(xlDown).row - 1    'n         : 선택된 변수의 데이타 갯수
        Next j



    Dim checkarray As String
    Dim anovastr As String
    Dim temp As String
    
    Dim cbindstr As String
    
    
    
   
    
      For j = 0 To nol - 1
        checkarray = xlist(j)
        If checkarray = "C" Then
           rinterface.PutArray "Cc", Range(Cells(2, k1(j)), Cells(n + 1, k1(j)))
        Else
        rinterface.PutArray checkarray, Range(Cells(2, k1(j)), Cells(n + 1, k1(j)))
        End If
    
            If j = 0 Then
                cbindstr = checkarray
                anovastr = checkarray
                Else
                cbindstr = cbindstr & "," & checkarray
                anovastr = anovastr & "+" & checkarray
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
   
    temp = combinationModule.combstr(xstrlist)
   ' MsgBox temp
    rinterface.RRun "QWE<-as.data.frame(arrayfrac)" 'ok
    rinterface.RRun "lm.3 =lm(" & checkarray & " ~ " & temp & ", data = arrayfrac)"
     rinterface.RRun "lmdata <- as.data.frame(lm.3)"
    j = Me.ListBox2.ListCount
    'MsgBox j & " N"
    
    rinterface.RRun "summary(lm.3)" 'ok
    'MsgBox anovastr
    rinterface.RRun "AnovaModeQ <- aov(lm(" & checkarray & " ~" & temp & ", data = arrayfrac))"
    rinterface.RRun "anova(AnovaModeQ)"     'ok
    'rinterface.RRun "MEPlot(AnovaModeQ, main = paste(" & Chr(34) & "주효과도" & Chr(34) & "), ylab=" & Chr(34) & "평균" & Chr(34) & ", pch = 15, mgp.ylab = 4, cex.title = 1.5, cex.main = par(" & Chr(34) & "cex.main" & Chr(34) & "), lwd = par(" & Chr(34) & "lwd" & Chr(34) & "), abbrev = " & j & " , select = NULL)" 'abbrev 요인수랑 맞춰야함
    'rinterface.RRun "MEPlot(AnovaModeQ, main = paste(" & Chr(34) & "주효과도" & Chr(34) & "), ylab=" & Chr(34) & "평균" & Chr(34) & ", pch = 15, mgp.ylab = 4, cex.title = 1.5, cex.main = par(" & Chr(34) & "cex.main" & Chr(34) & "), lwd = par(" & Chr(34) & "lwd" & Chr(34) & "), abbrev =  3 , select = NULL)" 'abbrev 요인수랑 맞춰야함
    rinterface.RRun "interactionPlot(arrayfrac, response(arrayfrac), fun = mean, main= " & Chr(34) & "교호작용도" & Chr(34) & ", col = 1:2 )"
    rinterface.InsertCurrentRPlot Range("I13"), widthrescale:=0.5, heightrescale:=0.5, closergraph:=True
  '  Rinterface.rrun "checker<-cbind(" & cbindstr & ")"
Unload Me
End Sub


Private Sub UserForm_Click()

End Sub
