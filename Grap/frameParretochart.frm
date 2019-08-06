VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameParretochart 
   OleObjectBlob   =   "frameParretochart.frx":0000
   Caption         =   "파레토그림"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6900
   StartUpPosition =   1  '소유자 가운데
   TypeInfoVer     =   26
End
Attribute VB_Name = "frameParretochart"
Attribute VB_Base = "0{1C3799CF-B6D5-48F3-98F8-E07DD12AC174}{10A47699-9FF3-4E84-B84D-B841CB8636DB}"
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

Private Sub CB1_Click()
    
    Dim i As Integer
    i = 0
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

End Sub

Private Sub CB2_Click()
    
    Me.ListBox1.AddItem Me.ListBox2.List(0)
    Me.ListBox2.RemoveItem (0)
    Me.CB2.Visible = False
    Me.CB1.Visible = True

End Sub

Private Sub CB3_Click()

   Dim i As Integer
    i = 0
    Do While i <= Me.ListBox1.ListCount - 1
        If Me.ListBox1.Selected(i) = True Then
           Me.ListBox3.AddItem Me.ListBox1.List(i)
           Me.ListBox1.RemoveItem (i)
           Me.CB3.Visible = False
           Me.CB4.Visible = True
           Exit Sub
        End If
        i = i + 1
    Loop

End Sub
Private Sub CB4_Click()
    
    
    Me.ListBox1.AddItem Me.ListBox3.List(0)
    Me.ListBox3.RemoveItem (0)
    Me.CB4.Visible = False
    Me.CB3.Visible = True


End Sub

Private Sub CommandButton1_Click()

    Dim i, cnt1, cnt2, Cnt3 As Integer
    Dim M1, m2, m3, M4 As Long
    
    Dim ErrSign As Boolean
    Dim MyColumnRange1 As Range, MyColumnRange2 As Range
   
    Dim ErrString As String
    Dim Resultsheet As Worksheet
    
    Dim nowsheet As String
    
     nowsheet = ActiveSheet.Name
    
  '폼에서의 입력에러 체크
    
    cnt1 = Me.ListBox1.ListCount
    cnt2 = Me.ListBox2.ListCount
    Cnt3 = Me.ListBox3.ListCount
    
    
   Dim Chk_Ver As Boolean
   Dim Cmp_R As Long
  
   Chk_Ver = PublicModule.ChkVersion(ActiveWorkbook.Name)
   If Chk_Ver = True Then
        Cmp_R = 1048576
    Else
        Cmp_R = 65536
    End If
    
    If cnt2 = 0 Then
        MsgBox "항목변수가 없습니다.", vbExclamation
        Exit Sub
    End If
    
    If Cnt3 = 0 Then
      MsgBox "분석변수가 없습니다.", vbExclamation
    Exit Sub
    End If
   
      
    For i = 1 To cnt1 + 2    ''''''빈칸 체크..
      
      If Me.ListBox2.List(0) = ActiveCell.Worksheet.Cells(1, i) Then
       M1 = ActiveSheet.Columns(i).Cells(1, 1).End(xlDown).row
       m2 = ActiveSheet.Columns(i).Cells(M1, 1).End(xlDown).row
         If m2 <> Cmp_R Then
            MsgBox ("항목변수란에 빈칸이 있습니다.")
         Exit Sub
         End If
      
       Set MyColumnRange1 = ActiveSheet.Columns(i).Range(Cells(2, 1), Cells(M1, 1))
      
       End If
         
       If Me.ListBox3.List(0) = ActiveCell.Worksheet.Cells(1, i) Then
       m3 = ActiveSheet.Columns(i).Cells(1, 1).End(xlDown).row
       M4 = ActiveSheet.Columns(i).Cells(m3, 1).End(xlDown).row
         If M4 <> Cmp_R Then
            MsgBox ("분석변수란에 빈칸이 있습니다.")
         Exit Sub
         End If
      
       Set MyColumnRange2 = ActiveSheet.Columns(i).Range(Cells(2, 1), Cells(m3, 1))
      
       End If
       
    Next i
    
    If M1 <> m3 Then
      MsgBox ("항목변수와 분석변수의 개수가 다릅니다.")
      Exit Sub
    End If
      
  
    num = MyColumnRange2.count
    
  '분석변수에 문자가 있는지 체크함
  '  ErrSign = ParetoModule.FindingRangeError2(MyColumnRange2, num)
   
    If ErrSign = True Then
        MsgBox "다음의 분석변수에 문자나 공백이 있습니다."
        Exit Sub
    End If
    
    
    Application.ScreenUpdating = False

   ' PublicModule.OpenOutSheet2 "_통계분석결과_", False
    PublicModule.OpenOutSheet2 "_TempPareto_", False
    Set Resultsheet = Worksheets("_TempPareto_")
    'Resultsheet.Unprotect "prophet"
    activePt = Resultsheet.Cells(1, 1).Value
    
    Dim tempchart0 As String
    
   tempchart0 = prvModuel.ParetoResultprv(Resultsheet, Me.ListBox2.List(0), Me.ListBox3.List(0), MyColumnRange1, MyColumnRange2, num)
            
   ' Resultsheet.Protect Password:="prophet", DrawingObjects:=False, contents:=True, Scenarios:=True
 
    
    
    Resultsheet.ChartObjects(tempchart0).Chart.Export Filename:="hist.tmp", FilterName:="GIF"
    Resultsheet.ChartObjects(tempchart0).Delete
    Me.Image1.Picture = LoadPicture("hist.tmp")
    Kill "hist.tmp"
    
    Worksheets(nowsheet).Activate

    

    
    

    
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


Private Sub OK_Click()

    Dim i, cnt1, cnt2, Cnt3 As Integer
    Dim M1, m2, m3, M4 As Long
    
    Dim ErrSign As Boolean
    Dim MyColumnRange1 As Range, MyColumnRange2 As Range
   
    Dim ErrString As String
    Dim Resultsheet As Worksheet
    
    
  '폼에서의 입력에러 체크
    
    cnt1 = Me.ListBox1.ListCount
    cnt2 = Me.ListBox2.ListCount
    Cnt3 = Me.ListBox3.ListCount
    
    
   Dim Chk_Ver As Boolean
   Dim Cmp_R As Long
  
   Chk_Ver = PublicModule.ChkVersion(ActiveWorkbook.Name)
   If Chk_Ver = True Then
        Cmp_R = 1048576
    Else
        Cmp_R = 65536
    End If
    
    If cnt2 = 0 Then
        MsgBox "항목변수가 없습니다.", vbExclamation
        Exit Sub
    End If
    
    If Cnt3 = 0 Then
      MsgBox "분석변수가 없습니다.", vbExclamation
    Exit Sub
    End If
   
      
    For i = 1 To cnt1 + 2    ''''''빈칸 체크..
      
      If Me.ListBox2.List(0) = ActiveCell.Worksheet.Cells(1, i) Then
       M1 = ActiveSheet.Columns(i).Cells(1, 1).End(xlDown).row
       m2 = ActiveSheet.Columns(i).Cells(M1, 1).End(xlDown).row
         If m2 <> Cmp_R Then
            MsgBox ("항목변수란에 빈칸이 있습니다.")
         Exit Sub
         End If
      
       Set MyColumnRange1 = ActiveSheet.Columns(i).Range(Cells(2, 1), Cells(M1, 1))
      
       End If
         
       If Me.ListBox3.List(0) = ActiveCell.Worksheet.Cells(1, i) Then
       m3 = ActiveSheet.Columns(i).Cells(1, 1).End(xlDown).row
       M4 = ActiveSheet.Columns(i).Cells(m3, 1).End(xlDown).row
         If M4 <> Cmp_R Then
            MsgBox ("분석변수란에 빈칸이 있습니다.")
         Exit Sub
         End If
      
       Set MyColumnRange2 = ActiveSheet.Columns(i).Range(Cells(2, 1), Cells(m3, 1))
      
       End If
       
    Next i
    
    If M1 <> m3 Then
      MsgBox ("항목변수와 분석변수의 개수가 다릅니다.")
      Exit Sub
    End If
      
  
    num = MyColumnRange2.count
    
  '분석변수에 문자가 있는지 체크함
  '  ErrSign = ParetoModule.FindingRangeError2(MyColumnRange2, num)
   
    If ErrSign = True Then
        MsgBox "다음의 분석변수에 문자나 공백이 있습니다."
        Exit Sub
    End If
    
    Me.Hide
    Application.ScreenUpdating = False

    PublicModule.OpenOutSheet2 "_통계분석결과_", False
    Set Resultsheet = Worksheets("_통계분석결과_")
    'Resultsheet.Unprotect "prophet"
    activePt = Resultsheet.Cells(1, 1).Value
    
    ParetoModule.ParetoResult Resultsheet, Me.ListBox2.List(0), Me.ListBox3.List(0), MyColumnRange1, MyColumnRange2, num
            
   ' Resultsheet.Protect Password:="prophet", DrawingObjects:=False, contents:=True, Scenarios:=True
  
    Unload Me
    

  
    Dim Cmp_Value As Long
    
    If PublicModule.ChkVersion(ActiveWorkbook.Name) = True Then
        Cmp_Value = 1048000
    Else
        Cmp_Value = 65000
    End If
    
    If Resultsheet.Cells(1, 1).Value > Cmp_Value Then
        MsgBox "[_통계분석결과_]시트를 거의 모두 사용하였습니다." & vbCrLf & "이 시트의 이름을 바꾸거나 삭제해 주세요", vbExclamation, "HIST"
        Exit Sub
    End If
    
    '결과 분석이 시작되는 부분을 보여주며 마친다.
    Resultsheet.Activate
    Resultsheet.Range("a" & activePt).Select
    Resultsheet.Range("a" & activePt).Activate
    
End Sub  ' 확인 버튼 프로시져 끝!!


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
   
   
   
   Me.ListBox1.List() = myArray
   
 '  For i = 1 To TempSheet.UsedRange.Columns.Count
 '   rngFirst.Offset(i, 1) = myArray(i - 1)
 ' Next i
  
End Sub

Private Sub UserForm_Click()

End Sub
