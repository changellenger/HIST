VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameApriori_origen 
   OleObjectBlob   =   "frameApriori_origen.frx":0000
   Caption         =   "APIRIORI"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7545
   StartUpPosition =   1  '소유자 가운데
   TypeInfoVer     =   94
End
Attribute VB_Name = "frameApriori_origen"
Attribute VB_Base = "0{7C43E65C-D395-4688-95D4-EFE159121546}{6C7DE8B1-01F4-4C04-8054-7236DB3004DE}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private Sub Label3_Click()

End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim i As Integer
     Dim j As Integer
     
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
    
     j = 0
    If Me.ListBox3.ListCount = 0 Then
        Do While j <= Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(j) = True Then
               Me.ListBox3.AddItem Me.ListBox1.List(j)
               Me.ListBox1.RemoveItem (j)
               Me.CB1.Visible = False
               Me.CB3.Visible = True
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
Private Sub ListBox3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.ListBox3.ListCount <> 0 Then
        Me.ListBox1.AddItem Me.ListBox3.List(0)
        Me.ListBox3.RemoveItem (0)
        Me.CB1.Visible = True
        Me.CB3.Visible = False
    End If

End Sub


Private Sub CB1_Click()
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
End Sub
Private Sub CB2_Click()
If Me.ListBox2.ListCount <> 0 Then
        Me.ListBox1.AddItem ListBox2.List(0)
        Me.ListBox2.RemoveItem (0)
     
    End If
End Sub
Private Sub CB3_Click()
Dim i As Integer
    i = 0
    If Me.ListBox3.ListCount = 0 Then
        Do While i <= Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(i) = True Then
               Me.ListBox3.AddItem Me.ListBox1.List(i)
               Me.ListBox1.RemoveItem (i)
               Me.CB1.Visible = False
               Me.CB2.Visible = True
               Me.CB3.Visible = True
               Exit Sub
            End If
            i = i + 1
        Loop
    End If
End Sub

Private Sub CB4_Click()
If Me.ListBox3.ListCount <> 0 Then
        Me.ListBox1.AddItem ListBox3.List(0)
        Me.ListBox3.RemoveItem (0)
     
    End If

End Sub
Private Sub okbtn_Click()
       Dim dataRange As Range
    Dim i As Integer
    Dim activePt As Long                                '결과 분석이 시작되는 부분을 보여주기 위함
    Dim rng As Range
    Dim a As String ' string  a  지정
    Dim b As String
    
   Dim j As Integer
      Dim dataRange2 As Range
      Dim xlist As String
      Dim xlist2 As String
      

    
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
     xlist2 = Me.ListBox3.List(0)
    DataSheet = ActiveSheet.Name                        'DataSheet : Data가 있는 Sheet 이름
    RstSheet = "_통계분석결과_"                       'RstSheet  : 결과를 보여주는 Sheet 이름
    
    
    
    '맨위에 입력
'On Error GoTo Err_delete
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
    
    tmp = 0
    For i = 1 To m
        If xlist = ActiveSheet.Cells(1, i) Then
            k1 = i  'k1                                 : k1 : 선택된 변수가 몇번째 열에 있는지
            tmp = tmp + 1
        End If
    Next i
    
    n = ActiveSheet.Cells(1, k1).End(xlDown).Row - 1    'n         : 선택된 변수의 데이타 갯수
    
   ' rng = Range(Cells(2, k1), Cells(N + 1, k1))
    
    
    

    '''
    ''' 변수명이 같은 경우 - 마지막 열에 있는 변수만 입력되므로 에러처리한다.
    '''
    If tmp > 1 Then
        MsgBox xlist & "와 같은 변수명이 있습니다. " & vbCrLf & "변수명을 바꿔주시기 바랍니다.", vbExclamation, "HIST"
        Exit Sub
    End If
    
     Set dataRange2 = ActiveSheet.Cells.CurrentRegion
    m2 = dataRange2.Columns.Count
    
    
    tmp2 = 0
    For j = 1 To m2
        If xlist2 = ActiveSheet.Cells(1, j) Then
            k2 = j  'k1                                 : k1 : 선택된 변수가 몇번째 열에 있는지
            tmp2 = tmp2 + 1
        End If
    Next j
    
    n2 = ActiveSheet.Cells(1, k2).End(xlDown).Row - 1    'n         : 선택된 변수의 데이타 갯수
    
   ' rng = Range(Cells(2, k1), Cells(N + 1, k1))
    

    '''
    ''' 변수명이 같은 경우 - 마지막 열에 있는 변수만 입력되므로 에러처리한다.
    '''
    If tmp2 > 1 Then
        MsgBox xlist2 & "와 같은 변수명이 있습니다. " & vbCrLf & "변수명을 바꿔주시기 바랍니다.", vbExclamation, "HIST"
        Exit Sub
    End If
    
    
rinterface.StartRServer
rinterface.RRun "install.packages (" & Chr(34) & "arules" & Chr(34) & ")"
rinterface.RRun "install.packages (" & Chr(34) & "arulesViz" & Chr(34) & ")"
rinterface.RRun "install.packages (" & Chr(34) & "grid" & Chr(34) & ")"

rinterface.RRun "require(grid)"
rinterface.RRun "library (arules)"

rinterface.RRun "require (arulesViz) "
rinterface.RRun "require (arules)"

'=====================패키지 설치 완료================
    
      rinterface.PutDataframe xlist, Range(Cells(1, k1), Cells(n + 1, k1))
      rinterface.PutDataframe xlist2, Range(Cells(1, k2), Cells(n2 + 1, k2))
      
      rinterface.RRun " frmset <- cbind(" & xlist & "," & xlist2 & " )"
    ' rinterface.PutDataframe "money", Range(Cells(1, k1), Cells(N2 + 1, k2))
      rinterface.RRun " attach(frmset)"
 '===================== 데이터 입력 및 통합 ==================
    
      
      a = "list(frmset$" & xlist & ", frmset$" & xlist2 & ")" 'ok
    ' a = "list(arraytest, arraytest2)" 'ok
      rinterface.RRun a
                MsgBox a
      b = "frmset.list <- split(frmset$" & xlist & ", frmset$" & xlist2 & ")" 'ok
    ' b = "alist <- split(arraytest, arraytest2)" 'ok
      rinterface.RRun b
      MsgBox b
      
            
      rinterface.RRun "frmset.trans<- as(frmset.list, " & Chr(34) & "transactions" & Chr(34) & ") "
     ' money.trans <- as(money.list,"transactions")
      rinterface.RRun "frmset.rules<-apriori(frmset.trans)"
      rinterface.RRun "ro<-as(frmset.rules, " & Chr(34) & "data.frame" & Chr(34) & ")"  ' rules 행렬, 즉 출력가능한 타입
      rinterface.RRun "r <-inspect(frmset.rules)"
      
      rinterface.RRun "top.frmset.rules<- head(sort(frmset.rules, decreasing = TRUE, by=" & Chr(34) & "lift" & Chr(34) & "))"
      rinterface.RRun "inspect(top.frmset.rules)"
      
      
      
      rinterface.RRun "plot(frmset.rules, method= " & Chr(34) & "grouped" & Chr(34) & " ) "
       '연관규칙의 조건과 결과를 기준으로 그래프를 보여줌. 색상의 진하기 향상도. 원의크기=지지도. 조건(LHS)앞의 숫자는 그 조건으로 되어 있는 연관규칙의 수. +숫자는 생략된 물품'
     rinterface.RRun "plot(frmset.rules, method= " & Chr(34) & "graph" & Chr(34) & ") "  '물품들 간의 연관성. 화살표 두께 = 지지도 , 화살표 진하기 = 향상도.
   
    
      'rinterface.RRun "plot(frmset.rules, measure = c(" & Chr(34) & "support" & Chr(34) & "," & Chr(34) & " lift" & Chr(34) & "), shading = " & Chr(34) & "confidence" & Chr(34) & ")"

Unload Me


 
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
  
End Sub

Private Sub UserForm_Click()

End Sub
