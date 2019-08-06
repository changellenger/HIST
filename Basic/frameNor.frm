VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameNor 
   OleObjectBlob   =   "frameNor.frx":0000
   Caption         =   "정규성검정"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7620
   StartUpPosition =   1  '소유자 가운데
   TypeInfoVer     =   31
End
Attribute VB_Name = "frameNor"
Attribute VB_Base = "0{638BFC97-302B-443A-853A-E561E02D1479}{6C7CF219-7DAD-458E-9183-D8FBA8E69599}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False



Private Sub CB1_Click()
    Dim i As Integer
    i = 0
    If Me.ListBox2.ListCount = 0 Then
        Do While i <= Me.Listbox1.ListCount - 1
            If Me.Listbox1.Selected(i) = True Then
               Me.ListBox2.AddItem Me.Listbox1.list(i)
               Me.Listbox1.RemoveItem (i)
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
        Me.Listbox1.AddItem ListBox2.list(0)
        Me.ListBox2.RemoveItem (0)
        Me.CB1.Visible = True
        Me.CB2.Visible = False
    End If
End Sub

Private Sub CommandButton1_Click()
    
    Dim temp As Integer
    Dim tempchartO As String
    Dim VarName As String: Dim selvar As Range
    
    VarName = SelectedVariable(Me.Listbox1.Value, selvar, Me.OptionButton1.Value)
    If VarName = "" Then
        If Me.Listbox1.list(0) = "" Then
            MsgBox "변수를 찾을 수 없습니다.", vbExclamation, "HIST"
        Else: MsgBox "분석변수를 선택하시오.", vbExclamation, "HIST"
        End If
        Exit Sub
    End If
        
    If PublicModule.FindingRangeError(selvar) = True Then
        MsgBox "분석변수에 문자나 공백이 있습니다.", _
            vbExclamation, "HIST"
        Exit Sub
    End If
    tempchartO = QQmodule.MainNormPlot(selvar, 100, 100, ActiveSheet, VarName:=VarName, NTest:=True)
    ActiveSheet.ChartObjects(tempchartO).Chart.Export _
        Filename:="qplot.tmp", FilterName:="GIF"
    ActiveSheet.ChartObjects(tempchartO).Delete
    Me.Image1.Picture = LoadPicture("qplot.tmp")
    Kill "qplot.tmp"

End Sub

Private Sub HistOk_Click()                      ''''"_그래프출력_"
    
    Dim temp As Integer: Dim ErrSign As Boolean
    Dim VarName As String: Dim selvar As Range
    Dim posi(0 To 1) As Long
  
    VarName = SelectedVariable(Me.Listbox1.Value, selvar, Me.OptionButton1.Value)
    If VarName = "" Then
        If Me.Listbox1.list(0) = "" Then
            MsgBox "변수를 찾을 수 없습니다.", vbExclamation, "HIST"
        Else: MsgBox "분석변수를 선택하시오.", vbExclamation, "HIST"
        End If
        Exit Sub
    End If
        
    If PublicModule.FindingRangeError(selvar) = True Then
        MsgBox "분석변수에 문자나 공백이 있습니다.", _
            vbExclamation, "HIST"
        Exit Sub
    End If
    
    Me.Hide
    PublicModule.SettingStatusBar True, "그래프 출력 중입니다."
    Application.ScreenUpdating = False
    
    ChartOutControl posi, True
    
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
    ModulePrint.Title1 "정규성검정 결과 "
    ModulePrint.Title3 "정규성검정"
    activePt = Worksheets(rstSheet).Cells(1, 1).Value
 

    QQmodule.MainNormPlot selvar, posi(0), posi(1), Worksheets("_통계분석결과_"), VarName:=VarName, NTest:=True
    ChartOutControl 192, False

  
  
    Application.ScreenUpdating = True
    PublicModule.SettingStatusBar False
    Worksheets("_통계분석결과_").Activate
    
    
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

    
    Worksheets(rstSheet).Activate
    Worksheets(rstSheet).Cells(activePt + 5, 1).Select
    Worksheets(rstSheet).Cells(activePt + 5, 1).Activate
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


Private Sub HistOk_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim i As Integer
    
    i = 0
    
    
    
    If Me.ListBox2.ListCount = 0 Then
        Do While i <= Me.Listbox1.ListCount - 1
            If Me.Listbox1.Selected(i) = True Then
               Me.ListBox2.AddItem Me.Listbox1.list(i)
               Me.Listbox1.RemoveItem (i)
               Me.CB1.Visible = False
               Me.CB2.Visible = True
               Exit Sub
            End If
            i = i + 1
        Loop
    Else
        Do While i <= Me.Listbox1.ListCount - 1
            If Me.Listbox1.Selected(i) = True Then
               Me.ListBox3.AddItem Me.Listbox1.list(i)
               Me.Listbox1.RemoveItem (i)
               Exit Do
            End If
            i = i + 1
        Loop
    End If
    
    If Me.ListBox3.ListCount = 1 Then
        Me.Frame2.Enabled = True
        Me.CheckBox3.Enabled = True
        Me.CheckBox4.Enabled = True
        Me.CheckBox5.Enabled = True
        Me.Label5.Enabled = True
    Else
        Me.Frame2.Enabled = False
        Me.CheckBox3.Enabled = False
        Me.CheckBox4.Enabled = False
        Me.CheckBox5.Enabled = False
        Me.Label5.Enabled = False
    End If

    
End Sub
Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.ListBox2.ListCount <> 0 Then
        Me.Listbox1.AddItem Me.ListBox2.list(0)
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
