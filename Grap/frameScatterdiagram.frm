VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameScatterdiagram 
   OleObjectBlob   =   "frameScatterdiagram.frx":0000
   Caption         =   "������"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6735
   StartUpPosition =   1  '������ ���
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
   ShellExecute 0, "open", "hh.exe", ThisWorkBook.Path + "\HIST%202013.chm::/������.htm", "", 1
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

Private Sub okbtn_Click()                       ''''"_�׷������_"
    
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
        MsgBox "������ ������ �ҿ����մϴ�.", vbExclamation
        Exit Sub
    End If
    
    If Me.CheckBox1.Value = False Then
        Vname(1) = PublicModule.SelectedVariable(Me.ListBox3.List(0), x, True)
    End If
    Vname(2) = PublicModule.SelectedVariable(Me.ListBox2.List(0), y, True)

    If PublicModule.FindingRangeError(y) Then
        MsgBox "�м������� ���ڳ� ������ �ֽ��ϴ�.", vbExclamation
        Exit Sub
    End If
    If Me.CheckBox1.Value = False Then
        If PublicModule.FindingRangeError(x) Then
            MsgBox "�м������� ���ڳ� ������ �ֽ��ϴ�.", vbExclamation
            Exit Sub
        End If
    End If
    
    If Me.CheckBox1.Value = False And Me.ListBox2.ListCount = 1 And _
       Me.ListBox3.ListCount = 1 Then
       If x.count <> y.count Then
            MsgBox "X-Y������ ������ ���� ���ƾ� �մϴ�.", vbExclamation
            Exit Sub
       End If
    End If

    Me.Hide

    ChartOutControl posi, True

    '''
    '''
    '''
    RstSheet = "_���м����_"
    
    '������ �Է�
On Error GoTo Err_delete
Dim val3535 As Long '�ʱ���ġ ������ ����'
Dim s3535 As Worksheet
val3535 = 2
    For Each s3535 In ActiveWorkbook.Sheets
        If s3535.Name = RstSheet Then
val3535 = Sheets(RstSheet).Cells(1, 1).Value
End If
Next s3535  '��Ʈ�� �̹������� ��� ��ġ �������ϰ�, ������ 2�� �����Ѵ�.




    'Worksheets(RstSheet).Unprotect "prophet"
    TModulePrint.Title1 "�׷������"
    activePt = Worksheets(RstSheet).Cells(1, 1).Value
      TModulePrint.Title3 "������"
    
    If Me.CheckBox1.Value = False Then
        ModuleScatter.ScatterPlot "_���м����_", posi(0) + 45, posi(1) + 30, 200, 200, x, y, Vname(1), Vname(2), Me.CheckBox2.Value
    Else
        ModuleScatter.OrderScatterPlot "_���м����_", posi(0), posi(1), 200, 200, y, Vname(2), 0
    End If
    ChartOutControl 200, False
    
    'Worksheets(RstSheet).Protect Password:="prophet", DrawingObjects:=False, _
    '                                contents:=True, Scenarios:=True             ''

    Worksheets("_���м����_").Activate
    
    '���� ���� üũ �� �񱳰� ����
    Dim Cmp_Value As Long
    
    If PublicModule.ChkVersion(ActiveWorkbook.Name) = True Then
        Cmp_Value = 1048000
    Else
        Cmp_Value = 65000
    End If
    
    If Worksheets(RstSheet).Cells(1, 1).Value > Cmp_Value Then
        MsgBox "[_���м����_]��Ʈ�� ���� ��� ����Ͽ����ϴ�." & vbCrLf & "�� ��Ʈ�� �̸��� �ٲٰų� ������ �ּ���", vbExclamation, "HIST"
        Exit Sub
    End If

    Worksheets(RstSheet).Activate
    Worksheets(RstSheet).Cells(activePt + 5, 1).Select
    Worksheets(RstSheet).Cells(activePt + 5, 1).Activate
                            '��� �м��� ���۵Ǵ� �κ��� �����ָ� ��ģ��.
                            


'�ǵڿ� ���̱�
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

MsgBox ("���α׷��� ������ �ֽ��ϴ�.")
 'End sub �տ��� ���δ�.

''�ؼ�, ������ ���� Err_delete�� �ͼ� ù�����ķ� �����. ���� ù���� 2�� ��Ʈ�� �����.�׸��� �����޽��� ���
'rSTsheet����⵵ ���� �������� ��쿡�� �ƹ� ���۵� ���� �ʰ�, �����޽����� ����.
                            
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
   If arrName(i) <> "" Then                     '��ĭ����
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
        MsgBox "������ ������ �ҿ����մϴ�.", vbExclamation
        Exit Sub
    End If
    
    If Me.CheckBox1.Value = False Then
        Vname(1) = PublicModule.SelectedVariable(Me.ListBox3.List(0), x, True)
    End If
    Vname(2) = PublicModule.SelectedVariable(Me.ListBox2.List(0), y, True)

    If PublicModule.FindingRangeError(y) Then
        MsgBox "�м������� ���ڳ� ������ �ֽ��ϴ�.", vbExclamation
        Exit Sub
    End If
    If Me.CheckBox1.Value = False Then
        If PublicModule.FindingRangeError(x) Then
            MsgBox "�м������� ���ڳ� ������ �ֽ��ϴ�.", vbExclamation
            Exit Sub
        End If
    End If
    
    If Me.CheckBox1.Value = False And Me.ListBox2.ListCount = 1 And _
       Me.ListBox3.ListCount = 1 Then
       If x.count <> y.count Then
            MsgBox "X-Y������ ������ ���� ���ƾ� �մϴ�.", vbExclamation
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
    

       '------ ���� �˻�
 '   If PublicModule.FindingRangeError(SelVar) = True Then
 '       MsgBox "�м������� ���ڳ� ������ �ֽ��ϴ�.", _
 '           vbExclamation, "HIST"
 '      Exit Sub
 '   End If
        '-------
    ChartOutControl posi, True
    
        
    If Me.CheckBox1.Value = False Then
      tempchartO = ModuleScatter.ScatterPlotprv("_���м����_", posi(0), posi(1), 200, 200, x, y, Vname(1), Vname(2), Me.CheckBox2.Value)
      '  tempchartO = ModuleScatter.OrderScatterPlotprv("_���м����_", posi(0), posi(1), 200, 200, y, Vname(2), 0)
    Else
     tempchartO = ModuleScatter.OrderScatterPlotprv("_���м����_", posi(0), posi(1), 200, 200, y, Vname(2), 0)
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
