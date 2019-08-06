VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameHistogram 
   OleObjectBlob   =   "frameHistogram.frx":0000
   Caption         =   "������׷�"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6150
   StartUpPosition =   1  '������ ���
   TypeInfoVer     =   57
End
Attribute VB_Name = "frameHistogram"
Attribute VB_Base = "0{085DE886-91EC-4429-B503-028D15EF73EA}{B231C558-E081-4316-8E87-217CBB45E733}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Obs As Integer

Private Sub AutoClass_Click()
    Me.TextBox1.Visible = False
    Me.SpinButton1.Visible = False
End Sub

Private Sub CommandButton1_Click()
    
    Dim temp As Integer
    Dim tempchartO As String
    Dim VarName As String: Dim SelVar As Range
    
    VarName = SelectedVariable(Me.ListBox1.Value, SelVar, Me.OptionButton1.Value)
    If VarName = "" Then
        If Me.ListBox1.List(0) = "" Then
            MsgBox "������ ã�� �� �����ϴ�.", vbExclamation, "HIST"
        Else: MsgBox "�м������� �����Ͻÿ�.", vbExclamation, "HIST"
        End If
        Exit Sub
    End If
       '------ ���� �˻�
    If PublicModule.FindingRangeError(SelVar) = True Then
        MsgBox "�м������� ���ڳ� ������ �ֽ��ϴ�.", _
            vbExclamation, "HIST"
       Exit Sub
    End If
        '-------
    
    If Me.AutoClass = True Then
        tempchartO = HistModule.MainHistogram(SelVar, 100, 100, ActiveSheet, VarName:=VarName)
    Else
        temp = Val(Me.TextBox1.Value)
        tempchartO = HistModule.MainHistogram(SelVar, 100, 100, ActiveSheet, temp, VarName)
    End If

    ActiveSheet.ChartObjects(tempchartO).Chart.Export _
        Filename:="hist.tmp", FilterName:="GIF"
    ActiveSheet.ChartObjects(tempchartO).Delete
    Me.Image1.Picture = LoadPicture("hist.tmp")
    Kill "hist.tmp"
    
    If Me.AutoClass = True Then
        Me.CustomClassN.Enabled = True
        Obs = SelVar.count
        Me.TextBox1.Value = HistModule.FindingNofClasses(Obs)
    Else
    End If

End Sub


Private Sub CommandButton2_Click()
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


Private Sub CommandButton3_Click()
    If Me.ListBox2.ListCount <> 0 Then
        Me.ListBox1.AddItem ListBox2.List(0)
        Me.ListBox2.RemoveItem (0)

End Sub

Private Sub CustomClassN_Click()
    Me.TextBox1.Visible = True
    Me.SpinButton1.Visible = True
    Me.SpinButton1.Enabled = True
End Sub

Private Sub HistCancel_Click()
    Unload Me
End Sub

Private Sub HistOk_Click()              ''''"_�׷������_"
    
    Dim temp As Integer: Dim ErrSign As Boolean
    Dim VarName As String: Dim SelVar As Range
    Dim posi(0 To 1) As Long
  
    VarName = PublicModule.SelectedVariable(Me.ListBox1.Value, SelVar, Me.OptionButton1.Value)
    If VarName = "" Then
        If Me.ListBox1.List(0) = "" Then
            MsgBox "������ ã�� �� �����ϴ�.", vbExclamation, "HIST"
        Else: MsgBox "�м������� �����Ͻÿ�.", vbExclamation, "HIST"
        End If
        Exit Sub
    End If
        
    If PublicModule.FindingRangeError(SelVar) = True Then
        MsgBox "�м������� ���ڳ� ������ �ֽ��ϴ�.", _
            vbExclamation, "HIST"
        Exit Sub
    End If
    
    Me.Hide
    PublicModule.SettingStatusBar True, "�׷��� ��� ���Դϴ�."
    Application.ScreenUpdating = False
    
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


    
    TModulePrint.TitleN "�׷������"
    activePt = Worksheets(RstSheet).Cells(1, 1).Value
      TModulePrint.Title3 "������׷�"



    If Me.AutoClass = True Then
        HistModule.MainHistogram SelVar, posi(0), posi(1), Worksheets("_���м����_"), VarName:=VarName
    Else
        temp = Me.TextBox1.Value
        HistModule.MainHistogram SelVar, posi(0), posi(1), Worksheets("_���м����_"), temp, VarName
    End If
    ChartOutControl 210, False

    Application.ScreenUpdating = True
    PublicModule.SettingStatusBar False
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
      Unload Me

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

Private Sub ListBox1_Change()
    Me.AutoClass.Value = True
    Me.CustomClassN.Enabled = False
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

    ReDim myArray(TempSheet.UsedRange.Columns.count)
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

Private Sub OptionButton2_Click()

   Dim myRange As Range
   Dim myArray()
   
   Me.ListBox1.Clear
   Set myRange = Cells.CurrentRegion.Columns(1)
   Cnt = myRange.Cells.count
   ReDim myArray(Cnt - 1)
   For i = 1 To Cnt
     myArray(i - 1) = myRange.Cells(i)
   Next i
   Me.ListBox1.List() = myArray

End Sub

Private Sub SpinButton1_SpinDown()
    If Me.TextBox1.Value >= 2 Then
        Me.TextBox1.Value = Me.TextBox1.Value - 1
    End If
End Sub

Private Sub SpinButton1_SpinUp()
    If Me.TextBox1.Value < 30 _
       And Me.TextBox1.Value < 2 * Int(Sqr(Obs)) Then
        Me.TextBox1.Value = Me.TextBox1.Value + 1
    End If
End Sub

Private Sub UserForm_Terminate()
    Unload Me
End Sub
