VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameBoxchart 
   OleObjectBlob   =   "frameBoxchart.frx":0000
   Caption         =   "���ڱ׸�"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6510
   StartUpPosition =   1  '������ ���
   TypeInfoVer     =   36
End
Attribute VB_Name = "frameBoxchart"
Attribute VB_Base = "0{F25B25B4-7470-4840-8AD1-175D83BA1D7F}{1FE8C013-B09C-4743-88B1-33D146EBE727}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False




Private Sub BoxOk_Click()                               ''''"_�׷������_"
    
    Dim i, Cnt As Integer: Dim ErrSign As Boolean
    Dim RnArray() As Range: Dim TempArray(1 To 1) As Range
    Dim VarName() As String: Dim posi(0 To 1) As Long
    Dim TempVarName(1 To 1), ErrString As String
    
    Cnt = Me.ListBox2.ListCount
    If Cnt = 0 Then
        MsgBox "�м������� �����ϴ�.", vbExclamation, "HIST"
        Exit Sub
    Else
        ReDim RnArray(1 To Cnt): ReDim VarName(1 To Cnt)
        SelectMultiRange Me, RnArray, VarName
    End If
        
    For i = 1 To Cnt
        If PublicModule.FindingRangeError2(RnArray(i)) = True Then
            ErrSign = True
            If ErrString <> "" Then
                ErrString = ErrString & "," & VarName(i)
            Else: ErrString = VarName(i)
            End If
        End If
    Next i
    If ErrSign = True Then
        MsgBox "������ �м������� ���ڳ� ������ �ֽ��ϴ�." & Chr(10) & _
               ": " & ErrString, vbExclamation, "HIST"
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

    
   
    TModulePrint.Title1 "�׷��� ���"
    activePt = Worksheets(RstSheet).Cells(1, 1).Value
   
    TModulePrint.Title3 "���ڱ׸�"
    
    If Me.optBoxCommonYes = True Then
        BoxPlotModule.MainBoxPlot RnArray, Cnt, _
            posi(0), posi(1), Worksheets("_���м����_"), VarName
        ChartOutControl 200, False
    Else
        For i = 1 To Cnt
            Set TempArray(1) = RnArray(i)
            TempVarName(1) = VarName(i)
            BoxPlotModule.MainBoxPlot TempArray, _
               1, posi(0) + 10 * (i - 1), posi(1) + 10 * (i - 1), _
                Worksheets("_���м����_"), TempVarName
        Next i
        ChartOutControl 200 + 10 * (i - 1), False
    End If
 
    
    
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
                            
End Sub

Private Sub ComBtn1_Click()
    MoveBtwnListBox Me, "ListBox1", "ListBox2"
End Sub
Private Sub ComBtn2_Click()
    MoveBtwnListBox Me, "ListBox2", "ListBox1"
End Sub


Private Sub CommandButton1_Click()

End Sub
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    MoveBtwnListBox Me, "ListBox1", "ListBox2"
End Sub

Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    MoveBtwnListBox Me, "ListBox2", "ListBox1"
End Sub
Private Sub OptionButton1_Click()
    OptBtn12Click Me, True
End Sub

Private Sub OptionButton2_Click()
    OptBtn12Click Me, False
End Sub
Private Sub UserForm_Terminate()
    Unload Me
End Sub
