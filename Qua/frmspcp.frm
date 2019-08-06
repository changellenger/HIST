VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmspcp 
   OleObjectBlob   =   "frmspcp.frx":0000
   Caption         =   "계수형 관리도  P"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   StartUpPosition =   1  '소유자 가운데
   TypeInfoVer     =   61
End
Attribute VB_Name = "frmspcp"
Attribute VB_Base = "0{590B81BF-3603-4CF8-970A-F10B893E6999}{F4B6E405-6191-4991-9E3F-9ED36D63D467}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False


Private Declare Function ShellExecute _
 Lib "shell32.dll" _
 Alias "ShellExecuteA" ( _
 ByVal hwnd As Long, _
 ByVal lpOperation As String, _
 ByVal lpFile As String, _
 ByVal lpParameters As String, _
 ByVal lpDirectory As String, _
 ByVal nShowCmd As Long) _
 As Long
Private Sub btnCancel_Click()
Unload Me
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

Private Sub Frame2_Click()

End Sub

Private Sub HlpBtn_Click()
ShellExecute 0, "open", "hh.exe", ThisWorkbook.Path + "\hist_help_V.2.5.1.chm::/P1.htm", "", 1
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



Private Sub TextBox1_Change()

End Sub
Sub CleanCharts()
    Dim chrt As Picture
    On Error Resume Next
    For Each chrt In ActiveSheet.Pictures
        chrt.Delete
    Next chrt
End Sub


Private Sub ToggleButton1_Click()
                           '넘길 정보는 검정값,신뢰구간,대립가설 3개니까
    Dim dataRange As Range
    Dim i As Integer
    Dim activePt As Long                                '결과 분석이 시작되는 부분을 보여주기 위함
    Dim rng As Range
    Dim a As String ' string  a  지정
   Dim j As Integer
      Dim dataRange2 As Range

    
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
    
    N = ActiveSheet.Cells(1, k1).End(xlDown).Row - 1    'n         : 선택된 변수의 데이타 갯수
    
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
    
    N2 = ActiveSheet.Cells(1, k2).End(xlDown).Row - 1    'n         : 선택된 변수의 데이타 갯수
    
   ' rng = Range(Cells(2, k1), Cells(N + 1, k1))
    
    
    

    '''
    ''' 변수명이 같은 경우 - 마지막 열에 있는 변수만 입력되므로 에러처리한다.
    '''
    If tmp2 > 1 Then
        MsgBox xlist2 & "와 같은 변수명이 있습니다. " & vbCrLf & "변수명을 바꿔주시기 바랍니다.", vbExclamation, "HIST"
        Exit Sub
    End If
      
    Rinterface.StartRServer
    'Rinterface.PutArray "arraytest", Range(rng)
    Rinterface.RRun "install.packages (" & Chr(34) & "qcc" & Chr(34) & ")"
    Rinterface.RRun "require (qcc)"
    
    
      Rinterface.PutArray "arraytest", Range(Cells(2, k1), Cells(N + 1, k1))
      Rinterface.PutArray "arraytest2", Range(Cells(2, k2), Cells(N2 + 1, k2))
      
      
      
      
     Application.ScreenUpdating = False
    Dim stname As String
    'Dim lastCol, lastRow As Integer
    
    stname = "관리도"
    Module2.OpenOutSheet stname, True
    Worksheets(stname).Activate
   
   Dim stn As Integer
   stn = Sheets(stname).Cells(1, 1).Value
    ActiveSheet.Cells(stn + 1, 1).Value = "데이터"
     ActiveSheet.Cells(stn + 1, 1).Font.Bold = True
     ActiveSheet.Cells(stn + 1, 1).Interior.Color = RGB(220, 238, 130)
     ActiveSheet.Cells(stn + 1, 1).Cells.ColumnWidth = 20
        
     ActiveSheet.Cells(stn + 1, 2).Interior.Color = RGB(220, 238, 130)
     

    ActiveSheet.Cells(stn + 2, 1).Value = xlist
    ActiveSheet.Cells(stn + 2, 2).Value = xlist2
    Rinterface.GetArray "arraytest", Range(Cells(stn + 3, 1), Cells(stn + 3, 1))
    Rinterface.GetArray "arraytest2", Range(Cells(stn + 3, 2), Cells(stn + 3, 2))
    

    'lastCol = ActiveCell.Worksheet.UsedRange.Columns.Count
    'lastRow = ActiveCell.Worksheet.UsedRange.Rows.Count
    
    'MsgBox " col:" & lastCol & " row: " & lastRow
    
    



 ActiveSheet.Cells(stn + 1, 4).Value = "관리도 그래프"
      ActiveSheet.Cells(stn + 1, 4).Font.Bold = True
ActiveSheet.Cells(stn + 1, 4).Interior.Color = RGB(220, 238, 130)

 
      
      
      
      
  
a = " p <- qcc(arraytest, type= " & Chr(34) & "p" & Chr(34) & ", size= arraytest2, title= " & Chr(34) & "P관리도" & Chr(34) & ")"
'Rinterface.RRun "qcc.options(" & Chr(34) & "beyond.limits" & Chr(34) & " = list(pch=15, col= " & Chr(34) & "orangered" & Chr(34) & "))"
'Rinterface.RRun " plot(p, title = " & Chr(34) & "P 관리도" & Chr(34) & ")"

Rinterface.RRun a


Rinterface.InsertCurrentRPlot Range(Cells(stn + 3, 4), Cells(stn + 3, 4)), widthrescale:=0.6, heightrescale:=0.6, closergraph:=True



Rinterface.RRun "l <- limits.p(p$center,p$std.dev,p$sizes,3) "
Rinterface.RRun "s <- stats.p(p$data, p$sizes)"
Rinterface.RRun "ss <- as.data.frame(s)"
Rinterface.RRun "d <- data.frame(l,ss)"
Rinterface.RRun "result <- row.names(d[which(d$UCL <= d$statistics), ])"

Rinterface.RRun "result3 <- t(result) "


Rinterface.GetArray "result3", Range(Cells(stn + 32, 5), Cells(stn + 32, 5))

Range(Cells(stn + 30, 4), Cells(stn + 30, 4)).Value = "P관리도 결과해석"
Range(Cells(stn + 30, 4), Cells(stn + 30, 4)).Cells.ColumnWidth = 15
Range(Cells(stn + 30, 4), Cells(stn + 30, 4)).Font.Bold = True
Range(Cells(stn + 30, 4), Cells(stn + 30, 4)).Interior.Color = RGB(220, 238, 130)

Range(Cells(stn + 32, 4), Cells(stn + 32, 4)).Value = "P관리상한선을 벗어나는 부분군:"
Range(Cells(stn + 32, 4), Cells(stn + 32, 4)).Cells.ColumnWidth = 28
Range(Cells(stn + 32, 4), Cells(stn + 32, 4)).Font.Color = vbBlack
Range(Cells(stn + 32, 4), Cells(stn + 32, 4)).Font.Bold = True

If Range(Cells(stn + 32, 5), Cells(stn + 32, 5)).Value = "" Then
Range(Cells(stn + 34, 5), Cells(stn + 34, 5)).Value = "공정이 관리상태에 있는 것으로 판정할 수 있습니다."
Else
Range(Cells(stn + 33, 5), Cells(stn + 33, 5)).Value = "번째 부분군이 '관리상한선'을 벗어났습니다. 따라서 공정에 이상원인이 있는 것으로 추정됩니다."
End If






Range(Cells(stn + 30, 4), Cells(stn + 34, 4)).Borders(xlEdgeLeft).LineStyle = xlContinuous '셀의 왼쪽 테두리 설정
   Range(Cells(stn + 30, 4), Cells(stn + 34, 4)).Borders(xlEdgeLeft).Color = RGB(34, 116, 34)
 Range(Cells(stn + 30, 4), Cells(stn + 34, 4)).Borders(xlEdgeLeft).Weight = 3
 
 Range(Cells(stn + 30, 14), Cells(stn + 34, 14)).Borders(xlEdgeRight).LineStyle = xlContinuous  '셀의 오른쪽 테두리 설정
 Range(Cells(stn + 30, 14), Cells(stn + 34, 14)).Borders(xlEdgeRight).Color = RGB(34, 116, 34)
  Range(Cells(stn + 30, 14), Cells(stn + 34, 14)).Borders(xlEdgeRight).Weight = 3
 
Range(Cells(stn + 30, 4), Cells(stn + 30, 14)).Borders(xlEdgeTop).LineStyle = xlContinuous  '셀의 위쪽 테두리 설정
 Range(Cells(stn + 30, 4), Cells(stn + 30, 14)).Borders(xlEdgeTop).Color = RGB(34, 116, 34)
Range(Cells(stn + 30, 4), Cells(stn + 30, 14)).Borders(xlEdgeTop).Weight = 3
 
 
  Range(Cells(stn + 34, 4), Cells(stn + 34, 14)).Borders(xlEdgeBottom).LineStyle = xlContinuous  '셀의 아래쪽 테두리 설정
  Range(Cells(stn + 34, 4), Cells(stn + 34, 14)).Borders(xlEdgeBottom).Color = RGB(34, 116, 34)
Range(Cells(stn + 34, 4), Cells(stn + 34, 14)).Borders(xlEdgeBottom).Weight = 3
 

 
 Range(Cells(stn + 30, 4), Cells(stn + 30, 14)).Borders(xlEdgeBottom).LineStyle = xlContinuous  '셀의 아래쪽 테두리 설정
   Range(Cells(stn + 30, 4), Cells(stn + 30, 14)).Borders(xlEdgeBottom).Color = RGB(34, 116, 34)
 Range(Cells(stn + 30, 4), Cells(stn + 30, 14)).Borders(xlEdgeBottom).Weight = 3
 
   Range(Cells(stn + 35, 1), Cells(stn + 35, 25)).Borders(xlEdgeBottom).LineStyle = xlContinuous  '셀의 아래쪽 테두리 설정
   Range(Cells(stn + 35, 1), Cells(stn + 35, 25)).Borders(xlEdgeBottom).Color = vbBlack
 Range(Cells(stn + 35, 1), Cells(stn + 35, 25)).Borders(xlEdgeBottom).Weight = 1
 

If N > 35 Then
 ActiveSheet.Cells(1, 1).Value = stn + N + 2
 Else
ActiveSheet.Cells(1, 1).Value = stn + 37
 
 
End If




'Rinterface.InsertCurrentRPlot Range("sheet1!i3"), widthrescale:=0.5, heightrescale:=0.5, closergraph:=True

      'Rinterface.RRun "x1 <- matrix(data= arraytest, nrow = 4, ncol=5, byrow = FALSE, dimnames = NULL)"
'no2 = Me.TextBox.Value
'Rinterface.RRun "qcc(arraytest, type= " & Chr(34) & "p" & Chr(34) & ", size=3)"
'b = " qcc(x1, type= " & Chr(34) & "p" & Chr(34) & ", size=" & no2 & ")  "
'Rinterface.RRun b
Unload Me

End Sub


Private Sub UserForm_Click()

End Sub
