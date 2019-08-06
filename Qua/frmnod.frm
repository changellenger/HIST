VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmnod 
   OleObjectBlob   =   "frmnod.frx":0000
   Caption         =   "정규분포"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   StartUpPosition =   1  '소유자 가운데
   TypeInfoVer     =   64
End
Attribute VB_Name = "frmnod"
Attribute VB_Base = "0{BAFFB040-C97B-4DE5-ABEF-1DA58D39DCFA}{B5C02616-BA47-4CA7-9302-1F7391B8DB5E}"
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
               Me.CB1.Visible = False
               Me.CB2.Visible = True
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
        Me.CB1.Visible = True
        Me.CB2.Visible = False
    End If
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub Frame2_Click()

End Sub

Private Sub HlpBtn_Click()
 ShellExecute 0, "open", "hh.exe", ThisWorkbook.Path + "\hist_help_V.2.5.1.chm::/정규분포_공정능력분석1.htm", "", 1
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
 Dim b As String ' string  a  지정
Dim no2 As Integer '부분군
   Dim no3 As Integer
   
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
   
    DataSheet = ActiveSheet.Name                'DataSheet : Data가 있는 Sheet 이름
   
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
  
  
  
  
  
   ' rng = Range(Cells(2, k1), Cells(N + 1, k1))
    
    
    

    '''
    ''' 변수명이 같은 경우 - 마지막 열에 있는 변수만 입력되므로 에러처리한다.
    '''
   
      
    Rinterface.StartRServer
    'Rinterface.PutArray "arraytest", Range(rng)
    Rinterface.RRun "install.packages (" & Chr(34) & "qualityTools" & Chr(34) & ")"
    Rinterface.RRun "require (qualityTools)"
    
    
      Rinterface.PutArray "arraytest", Range(Cells(2, k1), Cells(N + 1, k1))
      
      
      Application.ScreenUpdating = False
    Dim stname As String
    'Dim lastCol, lastRow As Integer
    
    stname = "공정능력분석"
    Module2.OpenOutSheet stname, True
    Worksheets(stname).Activate
   
   Dim stn As Integer
   stn = Sheets(stname).Cells(1, 1).Value
    ActiveSheet.Cells(stn + 1, 1).Value = "데이터"
     ActiveSheet.Cells(stn + 1, 1).Font.Bold = True
     ActiveSheet.Cells(stn + 1, 1).Interior.Color = RGB(220, 238, 130)
     ActiveSheet.Cells(stn + 1, 1).Cells.ColumnWidth = 20

    ActiveSheet.Cells(stn + 2, 1).Value = xlist
    Rinterface.GetArray "arraytest", Range(Cells(stn + 3, 1), Cells(stn + 3, 1))
    

    'lastCol = ActiveCell.Worksheet.UsedRange.Columns.Count
    'lastRow = ActiveCell.Worksheet.UsedRange.Rows.Count
    
    'MsgBox " col:" & lastCol & " row: " & lastRow
    
    


no2 = Me.TextBox2.Value
no3 = Me.TextBox3.Value
no4 = Me.TextBox4.Value


'Rinterface.RRun "qcc(x1, type= " & Chr(34) & "xbar" & Chr(34) & ", nsigmas=3)"


'a = 30 # N
'b = 5 #넘버수

'#  stringmerge(s1,s2)
'c = a / b
'Group = c(rep(1, 2), rep(2, 2), rep(3, 2), rep(4, 2), rep(5, 2), rep(6, 2), rep(7, 2), rep(8, 2), rep(9, 2), rep(10, 2))

 Dim Na As Integer
' Dim number As Integer
 Dim x, y As Integer
 Dim strZ As String
 
Na = N
y = Me.TextBox1.Value
x = Na / y
strZ = ""

For ii = 1 To x
If ii = 1 Then
strZ = "rep(" & ii & " , " & y & ")"

Else: strZ = strZ & ",rep(" & ii & " , " & y & ")"

End If

Next ii
strZ = "c(" & strZ & ")"
'MsgBox strZ



b = " rst <- pcr(arraytest, distribution = " & Chr(34) & "normal" & Chr(34) & ", lsl= " & no3 & ", usl = " & no2 & ", target= " & no4 & ", grouping=  " & strZ & ", main= " & Chr(34) & "정규분포 공정능력분석" & Chr(34) & ")"
'MsgBox strZ
  Rinterface.RRun b
Rinterface.InsertCurrentRPlot Range(Cells(stn + 3, 3), Cells(stn + 3, 3)), widthrescale:=0.9, heightrescale:=0.9, closergraph:=True

Rinterface.RRun " cpvalue <- rst$cp"
Rinterface.GetArray "cpvalue", Range(Cells(stn + 44, 4), Cells(stn + 44, 4))
Range(Cells(stn + 44, 3), Cells(stn + 44, 3)).Value = "공정능력지수(Cp): "
Range(Cells(stn + 44, 3), Cells(stn + 44, 3)).Font.Bold = True
Range(Cells(stn + 44, 3), Cells(stn + 44, 3)).Interior.Color = RGB(220, 238, 130)
Range(Cells(stn + 44, 3), Cells(stn + 44, 3)).Cells.ColumnWidth = 15

    
    
    
    
   Range(Cells(stn + 44, 3), Cells(stn + 45, 3)).Borders(xlEdgeLeft).LineStyle = xlContinuous  '셀의 왼쪽 테두리 설정
   Range(Cells(stn + 44, 3), Cells(stn + 45, 3)).Borders(xlEdgeLeft).Color = RGB(34, 116, 34)
 Range(Cells(stn + 44, 3), Cells(stn + 45, 3)).Borders(xlEdgeLeft).Weight = 3
 
  Range(Cells(stn + 44, 6), Cells(stn + 45, 6)).Borders(xlEdgeRight).LineStyle = xlContinuous  '셀의 오른쪽 테두리 설정
 Range(Cells(stn + 44, 6), Cells(stn + 45, 6)).Borders(xlEdgeRight).Color = RGB(34, 116, 34)
  Range(Cells(stn + 44, 6), Cells(stn + 45, 6)).Borders(xlEdgeRight).Weight = 3
 
 Range(Cells(stn + 44, 3), Cells(stn + 44, 6)).Borders(xlEdgeTop).LineStyle = xlContinuous  '셀의 위쪽 테두리 설정
 Range(Cells(stn + 44, 3), Cells(stn + 44, 6)).Borders(xlEdgeTop).Color = RGB(34, 116, 34)
Range(Cells(stn + 44, 3), Cells(stn + 44, 6)).Borders(xlEdgeTop).Weight = 3
 
 
   Range(Cells(stn + 45, 3), Cells(stn + 45, 6)).Borders(xlEdgeBottom).LineStyle = xlContinuous  '셀의 아래쪽 테두리 설정
   Range(Cells(stn + 45, 3), Cells(stn + 45, 6)).Borders(xlEdgeBottom).Color = RGB(34, 116, 34)
  Range(Cells(stn + 45, 3), Cells(stn + 45, 6)).Borders(xlEdgeBottom).Weight = 3
 
 
 
   Range(Cells(stn + 47, 1), Cells(stn + 47, 25)).Borders(xlEdgeBottom).LineStyle = xlContinuous  '셀의 아래쪽 테두리 설정
   Range(Cells(stn + 47, 1), Cells(stn + 47, 25)).Borders(xlEdgeBottom).Color = vbBlack
  Range(Cells(stn + 47, 1), Cells(stn + 47, 25)).Borders(xlEdgeBottom).Weight = 2
 

    
    

If Range(Cells(stn + 44, 4), Cells(stn + 44, 4)).Value >= 1.33 Then

Range(Cells(stn + 45, 4), Cells(stn + 45, 4)).Value = "공정능력이 충분합니다. "

ElseIf Range(Cells(stn + 44, 4), Cells(stn + 44, 4)).Value >= 1 And Range(Cells(stn + 44, 4), Cells(stn + 44, 4)).Value < 1.33 Then

Range(Cells(stn + 45, 4), Cells(stn + 45, 4)).Value = "공정능력이 있습니다. "

ElseIf Range(Cells(stn + 44, 4), Cells(stn + 44, 4)).Value >= 0.67 And Range(Cells(stn + 44, 4), Cells(stn + 44, 4)).Value < 1 Then

Range(Cells(stn + 45, 4), Cells(stn + 45, 4)).Value = "공정능력이 부족합니다. "


ElseIf Range(Cells(stn + 44, 4), Cells(stn + 44, 4)).Value < 0.67 Then

Range(Cells(stn + 45, 4), Cells(stn + 45, 4)).Value = "공정능력이 매우 부족합니다. "

End If


 If N > 47 Then
 ActiveSheet.Cells(1, 1).Value = stn + N + 2
 Else
ActiveSheet.Cells(1, 1).Value = stn + 48
 
 End If

Unload Me
End Sub

Private Sub ToggleButton2_Click()

End Sub

Private Sub UserForm_Click()

End Sub
