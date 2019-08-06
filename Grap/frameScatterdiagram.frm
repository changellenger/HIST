VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameScatterdiagram 
   OleObjectBlob   =   "frameScatterdiagram.frx":0000
   Caption         =   "산점도"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6735
   StartUpPosition =   1  '소유자 가운데
   TypeInfoVer     =   27
End
Attribute VB_Name = "frameScatterdiagram"
Attribute VB_Base = "0{D5AC2F89-0ECB-4923-853C-149BCA25A49F}{B995388C-C5AF-44C2-91E0-628F8828E256}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private Sub CheckBox1_Click()
    If Me.CheckBox1.Value = True Then
        Me.CheckBox2.Enabled = False
        Me.CheckBox2.Value = False
    Else: Me.CheckBox2.Enabled = True
    End If
End Sub

Private Sub CommandButton1_Click()
    
    Dim i As Integer
    i = 0
    Do While i <= Me.ListBox1.ListCount - 1
        If Me.ListBox1.Selected(i) = True Then
           Me.ListBox2.AddItem Me.ListBox1.List(i)
           Me.ListBox1.RemoveItem (i)
           Me.CommandButton1.Visible = False
           Me.CommandButton7.Visible = True
           Exit Sub
        End If
        i = i + 1
    Loop

End Sub

Private Sub CommandButton2_Click()
    
    Dim i As Integer
    i = 0
    Do While i <= Me.ListBox1.ListCount - 1
        If Me.ListBox1.Selected(i) = True Then
           Me.ListBox3.AddItem Me.ListBox1.List(i)
           Me.ListBox1.RemoveItem (i)
           Me.CommandButton2.Visible = False
           Me.CommandButton3.Visible = True
           Exit Sub
        End If
        i = i + 1
    Loop

End Sub

Private Sub CommandButton3_Click()
    
    Me.ListBox1.AddItem Me.ListBox3.List(0)
    Me.ListBox3.RemoveItem (0)
    Me.CommandButton3.Visible = False
    Me.CommandButton2.Visible = True

End Sub
Private Sub CommandButton5_Click()
    Unload Me
End Sub

Private Sub CommandButton6_Click()
   ShellExecute 0, "open", "hh.exe", ThisWorkBook.Path + "\HIST%202013.chm::/산점도.htm", "", 1
End Sub

Private Sub CommandButton7_Click()

    Me.ListBox1.AddItem Me.ListBox2.List(0)
    Me.ListBox2.RemoveItem (0)
    Me.CommandButton7.Visible = False
    Me.CommandButton1.Visible = True

End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
    Dim i As Integer
    
    i = 0
    If Me.ListBox2.ListCount = 0 Then
        Do While i <= Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(i) = True Then
               Me.ListBox2.AddItem Me.ListBox1.List(i)
               Me.ListBox1.RemoveItem (i)
               Me.CommandButton1.Visible = False
               Me.CommandButton7.Visible = True
               Exit Sub
            End If
            i = i + 1
        Loop
    ElseIf Me.ListBox3.ListCount = 0 Then
        Do While i <= Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(i) = True Then
               Me.ListBox3.AddItem Me.ListBox1.List(i)
               Me.ListBox1.RemoveItem (i)
               Me.CommandButton2.Visible = False
               Me.CommandButton3.Visible = True
               Exit Sub
            End If
            i = i + 1
        Loop
    Else
    End If

End Sub

Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.ListBox2.ListCount <> 0 Then
        Me.ListBox1.AddItem Me.ListBox2.List(0)
        Me.ListBox2.RemoveItem (0)
        Me.CommandButton7.Visible = False
        Me.CommandButton1.Visible = True
    End If
End Sub

Private Sub ListBox3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.ListBox3.ListCount <> 0 Then
        Me.ListBox1.AddItem Me.ListBox3.List(0)
        Me.ListBox3.RemoveItem 0
        Me.CommandButton3.Visible = False
        Me.CommandButton2.Visible = True
    End If
End Sub

Private Sub okbtn_Click()                       ''''"_그래프출력_"
    
    Dim x As Range: Dim y As Range: Dim ErrSign, ErrSign2 As Boolean
    Dim posi(0 To 1) As Long: Dim Vname(1 To 2) As String
    
    If Me.CheckBox1.Value = False And Me.ListBox2.ListCount = 1 And _
       Me.ListBox3.ListCount = 1 Then
        ErrSign2 = False
    ElseIf Me.CheckBox1.Value = True And Me.ListBox2.ListCount = 1 Then
        ErrSign2 = False
    Else: ErrSign2 = True
    End If
    
    If ErrSign2 = True Then
        MsgBox "변수의 선택이 불완전합니다.", vbExclamation
        Exit Sub
    End If
    
    If Me.CheckBox1.Value = False Then
        Vname(1) = PublicModule.SelectedVariable(Me.ListBox3.List(0), x, True)
    End If
    Vname(2) = PublicModule.SelectedVariable(Me.ListBox2.List(0), y, True)

    If PublicModule.FindingRangeError(y) Then
        MsgBox "분석변수에 문자나 공백이 있습니다.", vbExclamation
        Exit Sub
    End If
    If Me.CheckBox1.Value = False Then
        If PublicModule.FindingRangeError(x) Then
            MsgBox "분석변수에 문자나 공백이 있습니다.", vbExclamation
            Exit Sub
        End If
    End If
    
    If Me.CheckBox1.Value = False And Me.ListBox2.ListCount = 1 And _
       Me.ListBox3.ListCount = 1 Then
       If x.count <> y.count Then
            MsgBox "X-Y변수의 개수가 서로 같아야 합니다.", vbExclamation
            Exit Sub
       End If
    End If

    Me.Hide

    ChartOutControl posi, True

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
    TModulePrint.Title1 "그래프출력"
    activePt = Worksheets(RstSheet).Cells(1, 1).Value
      TModulePrint.Title3 "산점도"
    
    If Me.CheckBox1.Value = False Then
        ModuleScatter.ScatterPlot "_통계분석결과_", posi(0) + 45, posi(1) + 30, 200, 200, x, y, Vname(1), Vname(2), Me.CheckBox2.Value
    Else
        ModuleScatter.OrderScatterPlot "_통계분석결과_", posi(0), posi(1), 200, 200, y, Vname(2), 0
    End If
    ChartOutControl 200, False
    
    'Worksheets(RstSheet).Protect Password:="prophet", DrawingObjects:=False, _
    '                                contents:=True, Scenarios:=True             ''

    Worksheets("_통계분석결과_").Activate
    
    '파일 버전 체크 후 비교값 정의
    Dim Cmp_Value As Long
    
    If PublicModule.ChkVersion(ActiveWorkbook.Name) = True Then
        Cmp_Value = 1048000
    Else
        Cmp_Value = 65000
    End If
    
    If Worksheets(RstSheet).Cells(1, 1).Value > Cmp_Value Then
        MsgBox "[_통계분석결과_]시트를 거의 모두 사용하였습니다." & vbCrLf & "이 시트의 이름을 바꾸거나 삭제해 주세요", vbExclamation, "HIST"
        Exit Sub
    End If

    Worksheets(RstSheet).Activate
    Worksheets(RstSheet).Cells(activePt + 5, 1).Select
    Worksheets(RstSheet).Cells(activePt + 5, 1).Activate
                            '결과 분석이 시작되는 부분을 보여주며 마친다.
                            


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
   
   
   
   Me.ListBox1.List() = myArray
   
 '  For i = 1 To TempSheet.UsedRange.Columns.Count
 '   rngFirst.Offset(i, 1) = myArray(i - 1)
 ' Next i
  
End Sub

Private Sub prv_Click()
 Dim tempchartO As String
        Dim x As Range: Dim y As Range: Dim ErrSign, ErrSign2 As Boolean
    Dim posi(0 To 1) As Long: Dim Vname(1 To 2) As String
    Dim nowsheet As String
    
    nowsheet = ActiveSheet.Name
    
    
    If Me.CheckBox1.Value = False And Me.ListBox2.ListCount = 1 And _
       Me.ListBox3.ListCount = 1 Then
        ErrSign2 = False
    ElseIf Me.CheckBox1.Value = True And Me.ListBox2.ListCount = 1 Then
        ErrSign2 = False
    Else: ErrSign2 = True
    End If
    
    If ErrSign2 = True Then
        MsgBox "변수의 선택이 불완전합니다.", vbExclamation
        Exit Sub
    End If
    
    If Me.CheckBox1.Value = False Then
        Vname(1) = PublicModule.SelectedVariable(Me.ListBox3.List(0), x, True)
    End If
    Vname(2) = PublicModule.SelectedVariable(Me.ListBox2.List(0), y, True)

    If PublicModule.FindingRangeError(y) Then
        MsgBox "분석변수에 문자나 공백이 있습니다.", vbExclamation
        Exit Sub
    End If
    If Me.CheckBox1.Value = False Then
        If PublicModule.FindingRangeError(x) Then
            MsgBox "분석변수에 문자나 공백이 있습니다.", vbExclamation
            Exit Sub
        End If
    End If
    
    If Me.CheckBox1.Value = False And Me.ListBox2.ListCount = 1 And _
       Me.ListBox3.ListCount = 1 Then
       If x.count <> y.count Then
            MsgBox "X-Y변수의 개수가 서로 같아야 합니다.", vbExclamation
            Exit Sub
       End If
    End If

    'Me.Hide

    ChartOutControl posi, True
    
    If Me.CheckBox1.Value = False And Me.ListBox2.ListCount = 1 And _
       Me.ListBox3.ListCount = 1 Then
        ErrSign2 = False
    ElseIf Me.CheckBox1.Value = True And Me.ListBox2.ListCount = 1 Then
        ErrSign2 = False
    Else: ErrSign2 = True
    End If
    

       '------ 에러 검사
 '   If PublicModule.FindingRangeError(SelVar) = True Then
 '       MsgBox "분석변수에 문자나 공백이 있습니다.", _
 '           vbExclamation, "HIST"
 '      Exit Sub
 '   End If
        '-------
    ChartOutControl posi, True
    
        
    If Me.CheckBox1.Value = False Then
      tempchartO = ModuleScatter.ScatterPlotprv("_통계분석결과_", posi(0), posi(1), 200, 200, x, y, Vname(1), Vname(2), Me.CheckBox2.Value)
      '  tempchartO = ModuleScatter.OrderScatterPlotprv("_통계분석결과_", posi(0), posi(1), 200, 200, y, Vname(2), 0)
    Else
     tempchartO = ModuleScatter.OrderScatterPlotprv("_통계분석결과_", posi(0), posi(1), 200, 200, y, Vname(2), 0)
    End If
    
    
    
 '   If Me.AutoClass = True Then
 '       tempchartO = HistModule.MainHistogram(SelVar, 100, 100, ActiveSheet, VarName:=VarName)
 '   Else
 '       temp = Val(Me.TextBox1.Value)
 '       tempchartO = HistModule.MainHistogram(SelVar, 100, 100, ActiveSheet, temp, VarName)
 '   End If

    ActiveSheet.ChartObjects(tempchartO).Chart.Export Filename:="hist.tmp", FilterName:="GIF"
    ActiveSheet.ChartObjects(tempchartO).Delete
    Me.Image1.Picture = LoadPicture("hist.tmp")
    Kill "hist.tmp"
    
    Worksheets(nowsheet).Activate

End Sub

Private Sub UserForm_Terminate()
    Unload Me
End Sub
Function StorageForStatic(ChartName As String, _
    ChartNum As Integer, Output As Boolean) As String
    
    Static NewChartName(1 To 6) As String
    
    If Output = False Then
        NewChartName(ChartNum) = ChartName
        StorageForStatic = ""
    Else
        StorageForStatic = NewChartName(ChartNum)
    End If
    
End Function
