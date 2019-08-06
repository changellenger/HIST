VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFrequency 
   OleObjectBlob   =   "frmFrequency.frx":0000
   Caption         =   "�� �м�"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7050
   StartUpPosition =   1  '������ ���
   TypeInfoVer     =   75
End
Attribute VB_Name = "frmFrequency"
Attribute VB_Base = "0{E354677E-A2E6-4755-BF54-C698A21903D6}{7E1E991A-9913-455A-B62E-B8FA468510C9}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False





Sub MoveBtwnListBox(ParentD, FromLNum, ToLNum)

    Dim i As Integer
    i = 0
    Do While i <= ParentD.Controls(FromLNum).ListCount - 1
        If ParentD.Controls(FromLNum).Selected(i) = True Then
           ParentD.Controls(ToLNum).AddItem ParentD.Controls(FromLNum).list(i)
           ParentD.Controls(FromLNum).RemoveItem i
            Exit Do
        End If
        i = i + 1
    Loop

End Sub



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

Private Sub CommandButton6_Click()
ShellExecute 0, "open", "hh.exe", ThisWorkbook.Path + "\HIST%202013.chm::/�󵵺м�.htm", "", 1
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



Private Sub CommandButton1_Click()
       
    MoveBtwnListBox Me, "ListBox1", "ListBox2"
    
End Sub

Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.ListBox2.ListCount <> 0 Then
        Me.Listbox1.AddItem Me.ListBox2.list(0)
        Me.ListBox2.RemoveItem (0)
        Me.CB1.Visible = True
        Me.CB2.Visible = False
    End If
End Sub


Private Sub CommandButton3_Click()

    MoveBtwnListBox Me, "ListBox2", "ListBox1"
      
End Sub


Private Sub Listbox3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    If Me.ListBox3.ListCount <> 0 Then
        Me.Listbox1.AddItem Me.ListBox3.list(0)
        Me.ListBox3.RemoveItem (0)
        Me.CommandButton4.Visible = False
        Me.CommandButton2.Visible = True
    End If
      
End Sub


Private Sub CommandButton2_Click()

    Dim i As Integer
    i = 0
    Do While i <= Me.Listbox1.ListCount - 1
    If Me.Listbox1.Selected(i) = True Then
           Me.ListBox3.AddItem Me.Listbox1.list(i)
           Me.Listbox1.RemoveItem (i)
           Me.CommandButton2.Visible = False
           Me.CommandButton4.Visible = True
           Exit Sub
    End If
    i = i + 1
    Loop
    
End Sub

Private Sub CommandButton4_Click()

    Me.Listbox1.AddItem Me.ListBox3.list(0)
    Me.ListBox3.RemoveItem (0)
    Me.CommandButton4.Visible = False
    Me.CommandButton2.Visible = True

End Sub


Private Sub BoxCancel_Click()

    Unload Me
    
End Sub


Private Sub CheckBox1_Click()

    If Me.CheckBox1.Value = True Then
        Me.Label5.Enabled = True
        Me.ListBox3.Enabled = True
        Me.CommandButton2.Enabled = True
    Else
        Me.Label5.Enabled = False
        Me.ListBox3.Enabled = False
        Me.CommandButton2.Enabled = False
    End If
    
End Sub
















Private Sub BoxOk_Click()
   
    Dim dataRange As Variant, Position1 As Range, Position2 As Range
    Dim MyFreqName() As String, MyFreqStringName() As String
    Dim VarLen() As Long, KK As Integer, LevelVariable As String
    Dim tempstr() As String, ErrString As String
    Dim i As Integer, Obscount As Integer
    Dim activePt As Long                                                        '' ��� �м��� ���۵Ǵ� �κ��� �����ֱ� ����
    Dim X(), list(), cal()
    Dim tit As Integer
    '''
    ''' ���� ó�� �κ� 0: 1�� 1��
    '''
    If ActiveSheet.Cells(1, 1) = "" Then
        MsgBox "1�� 1���� �������� �ʿ��մϴ�.", vbExclamation                  '' 1�� 1���� ��� ���� ���� �߻�.
        Exit Sub
    End If
    
    '''
    ''' ���� ó�� �κ� 1: ���� ���ÿ��� Ȯ��
    '''
    If Me.ListBox2.ListCount = 0 Then
        MsgBox "������ �������� �ʾҽ��ϴ�.", vbExclamation                     '' ���� ������ �������� �ʽ��ϴ�.  ����  ��������.
        Exit Sub
    End If


    '''
    ''' �Է¹��� ���� �����ϱ�
    '''
    ''' ������� ModeuleControl ���� ����� Public ����
    ''' ���⼭ �ѹ��� �������ش�
    ''' sheetRowNum, sheetColNum, DataSheet, RstSheet, xlist, n, m, p
    '''
    If right(ActiveWorkbook.Name, 4) = ".xls" Or right(ActiveWorkbook.Name, 4) = ".XLS" Then
        sheetRowNum = 2 ^ 16            '65536
        sheetColNum = 2 ^ 8             '256
        sheetApproxRowNum = 65000
    Else
        sheetRowNum = 2 ^ 20            '1048576
        sheetColNum = 2 ^ 14            '16384
        sheetApproxRowNum = 1048000
    End If
    
    DataSheet = ActiveSheet.Name                                                '' Data�� �ִ� Sheet �̸�
    rstSheet = "_���м����_"                                                 '' ����� �����ִ� Sheet �̸�
    '����ϴ� �ش� ��⿡ �� ���� ����'
'������ �Է�
On Error GoTo Err_delete
Dim val3535 As Long '�ʱ���ġ ������ ����'
Dim s3535 As Worksheet
val3535 = 2
    For Each s3535 In ActiveWorkbook.Sheets
        If s3535.Name = rstSheet Then
val3535 = Sheets(rstSheet).Cells(1, 1).Value
End If
Next s3535  '��Ʈ�� �̹������� ��� ��ġ �������ϰ�, ������ 2�� �����Ѵ�.

        
    Set dataRange = ActiveSheet.Cells.CurrentRegion
    m = dataRange.Cells(1, 1).End(xlToRight).Column                             '' ��ü �������� ����
    
    p = Me.ListBox2.ListCount                                                   '' ���õ� �������� ����
    ReDim xlist(p - 1)
    For i = 0 To p - 1
        xlist(i) = ListBox2.list(i)                                             '' ���õ� �������� �̸�
    Next i
    
    N = ModuleControl.FindDataCount(xlist) - 1                                  '' ���õ� ������ Data����
    
    '''
    ''' ���� ó�� �κ� 2: �������� �������� ����
    '''
    ErrSign = False
    For i = 0 To p - 1
        If N <> ModuleControl.FindColDataCount(xlist(i)) Then ErrSign = True
    Next i
    
    If ErrSign = True Then
        MsgBox "���õ� �׸�鰣�� �������� �ٸ��ϴ�.", vbExclamation, "HIST"
        Exit Sub
    End If
    ErrSign = False
    
    '''
    ''' ���� ó�� �κ� 4: �������� ���� ��� - ������ ���� �ִ� ������ �ԷµǹǷ� ����ó���Ѵ�.
    '''
    For i = 1 To p
        errTmp = 0
        For J = 1 To m
            If Me.ListBox2.list(i - 1) = ActiveSheet.Cells(1, J) Then
                errTmp = errTmp + 1
            End If
        Next J
        If errTmp > 1 Then
            MsgBox xlist(i - 1) & vbCrLf & vbCrLf & "���� �м������� ���� �������� �ֽ��ϴ�. " & vbCrLf & "�������� �ٲ��ֽñ� �ٶ��ϴ�.", vbExclamation, "HIST"
            Exit Sub
        End If
    Next i
    
'    If Me.ListBox3.ListCount = 1 Then
'        LevelVariable = Me.ListBox3.List(0)
'    Else
'        LevelVariable = ""
'    End If
    ReDim X(N, m)
    X = ActiveSheet.Range(Cells(1, 1), Cells(N + 1, m)).Value


    ReDim list(1 To m) '���� ��� �ۼ� �������� ��������'
    For i = 1 To m
    list(i) = X(1, i)
    Next i

    ReDim cal(1 To N, 1 To m - 1)
    For i = 1 To N
        For J = 1 To m - 1
            cal(i, J) = X(i + 1, J + 1)
        Next J
    Next i
    
    tit = 1
    '''
    '''��� ó��
    '''
    ModuleControl.SettingStatusBar True, "�� �м����Դϴ�."
    Application.ScreenUpdating = False
    
    ModulePrint.makeOutputSheet rstSheet
    activePt = Worksheets(rstSheet).Cells(1, 1).Value
    
    
    ModuleControl.FreqAnalysis cal, list, tit '���ڿ�, �󵵺м��ϳ��� ��ġ��
    
  
    
    ModuleControl.SettingStatusBar False
    Application.ScreenUpdating = True
    Unload Me
    
    Worksheets(rstSheet).Activate
    Worksheets(rstSheet).Cells(activePt, 1).Select
    Worksheets(rstSheet).Cells(activePt, 1).Activate                            '��� �м��� ���۵Ǵ� �κ��� �����ָ� ��ģ��.
    
    Worksheets(rstSheet).Activate
    If Worksheets(rstSheet).Cells(1, 1).Value > sheetApproxRowNum Then
        MsgBox "[_���м����_]��Ʈ�� ���� ��� ����Ͽ����ϴ�." & vbCrLf & "�� ��Ʈ�� �̸��� �ٲٰų� ������ �ּ���", vbExclamation, "HIST"
        Exit Sub
    End If
    
'�ǵڿ� ���̱�
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
   
   Me.Listbox1.Clear

    ReDim myArray(TempSheet.UsedRange.Columns.count - 1)
    a = 0
   For i = 1 To TempSheet.UsedRange.Columns.count
   If arrName(i) <> "" Then                     '��ĭ����
   myArray(a) = arrName(i)
   a = a + 1
   
   Else:
   End If
   Next i
  
   Me.Listbox1.list() = myArray



End Sub

Private Sub UserForm_Click()

End Sub
