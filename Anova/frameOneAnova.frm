VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameOneAnova 
   OleObjectBlob   =   "frameOneAnova.frx":0000
   Caption         =   "�Ͽ���ġ��"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7815
   StartUpPosition =   1  '������ ���
   TypeInfoVer     =   42
End
Attribute VB_Name = "frameOneAnova"
Attribute VB_Base = "0{689C6689-4F63-4749-8668-56BC6C9B7313}{D9EA435E-1B6D-4F43-8084-CC11F4292DAB}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CheckBox3_Click()
If Me.CheckBox3.Value = True Then
    Me.CheckBox4.Enabled = True
Else
    Me.CheckBox4.Value = False
    Me.CheckBox4.Enabled = False
End If
End Sub

Private Sub CheckBox4_Click()

End Sub

Private Sub CommandButton11_Click()
    
    Dim i As Integer
    i = 0
    Do While i <= Me.ListBox1.ListCount - 1
        If Me.ListBox1.Selected(i) = True Then
           Me.ListBox2.AddItem Me.ListBox1.list(i)
           Me.ListBox1.RemoveItem (i)
           Me.CommandButton11.Visible = False
           Me.CommandButton14.Visible = True
           Exit Sub
        End If
        i = i + 1
    Loop

End Sub
Private Sub CommandButton12_Click()
    Dim i As Integer
    i = 0
    Do While i <= Me.ListBox1.ListCount - 1
        If Me.ListBox1.Selected(i) = True Then
           Me.ListBox3.AddItem Me.ListBox1.list(i)
           Me.ListBox1.RemoveItem (i)
           Me.CommandButton12.Visible = False
           Me.CommandButton13.Visible = True
           Exit Sub
        End If
        i = i + 1
    Loop
End Sub
Private Sub CommandButton13_Click()
    Me.ListBox1.AddItem Me.ListBox3.list(0)
    Me.ListBox3.RemoveItem (0)
    Me.CommandButton13.Visible = False
    Me.CommandButton12.Visible = True
End Sub

Private Sub CommandButton14_Click()
    Me.ListBox1.AddItem Me.ListBox2.list(0)
    Me.ListBox2.RemoveItem (0)
    Me.CommandButton14.Visible = False
    Me.CommandButton11.Visible = True
End Sub
Private Sub CommandButton15_Click()
    Dim resultsheet, TempSheet As Worksheet
    Dim cr As Long
    Dim N As Long
    Dim tmean As Double
    Dim tsum As Double
    Dim tisum As Double
    Dim tisumsq As Double
    Dim SSE As Double
    Dim st As Double
    Dim ct As Integer
    Dim xsq As Double
    Dim d As Range
    Dim Sa As Double
    Dim es As Boolean
    Dim res As Worksheet
    Dim xnames()
    
    Dim Colname, valueName, factor() As String
    Dim cRn, vrn, temp As Range: Dim sRn() As Range
    Dim cnt() As Long: Dim mean() As Double: Dim std()
        
    Dim M1, M2 As Long
    Dim fitted(), resi() As Double
    Dim posi(0 To 1) As Long
    Dim fit, X, y As Range
    Dim selvar As Range
    
    
    If Me.ListBox2.ListCount = 0 Or Me.ListBox3.ListCount = 0 Then
        MsgBox "������ ������ �ҿ����մϴ�.", vbExclamation, "HIST"
        Exit Sub
    End If
    
    Colname = ModuleControl.SelectedVariable(Me.ListBox2.list(0), cRn, True)
    valueName = ModuleControl.SelectedVariable(Me.ListBox3.list(0), vrn, True)
    
    If FindingRangeError(vrn) Then
        MsgBox "�з������� �м������� ���ڳ� ������ �ֽ��ϴ�.", vbExclamation, "HIST"
        Exit Sub
    End If
    
    If cRn.count <> vrn.count Then
            MsgBox "�з������� �м��������� ������ �߸��Ǿ����ϴ�.", vbExclamation, "HIST"
            Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Set TempSheet = ActiveCell.Worksheet
    Set temp = TempSheet.Cells.CurrentRegion
    ModuleControl.PivotMakerforOneWay temp, Colname, valueName, cnt, mean, std, factor
    cr = UBound(cnt) - 1
    
    st = 0: SSE = 0: N = 0: tisumsq = 0: xsq = 0
    tmean = Application.Average(vrn): tsum = Application.sum(vrn)
    
     For Each d In vrn
        xsq = xsq + d.Value ^ 2
    Next d
    For i = 1 To cr
        N = N + cnt(i)
        tisum = cnt(i) * mean(i)
        tisumsq = tisumsq + tisum ^ 2 / cnt(i)
    Next i
    tsum = tsum ^ 2 / N
    Sa = xsq - tsum
    st = tisumsq - tsum
    SSE = Sa - st
    tdf = cr - 1
    edf = N - cr
    
    '���հ� ���ؼ� �迭�� ������ ����
    ReDim fitted(0 To N - 1)
    J = 1
    For i = 1 To N
        Do While cRn(i) <> factor(J)
            J = J + 1
        Loop
        fitted(i - 1) = mean(J)
        fitted(i - 1) = Application.Round(fitted(i - 1), 4)
        J = 1
    Next i

'���� ���ؼ� �迭�� ������ ����
    ReDim resi(0 To N - 1)
    J = 1
    For i = 1 To N
        Do While cRn(i) <> factor(J)
            J = J + 1
        Loop
        resi(i - 1) = vrn(i) - mean(J)
        resi(i - 1) = Application.Round(resi(i - 1), 4)
        J = 1
    Next i
           
'����, ���԰� ������ ��Ʈ�� �Ѹ���
Dim count As Integer
count = 0
     If Me.CheckBox3 = True Then
        M1 = ActiveSheet.Rows(1).Cells(1, 1).End(xlToRight).Column
        Set ttemp1 = ActiveSheet.Cells(1, M1 + 1)
        For i = 1 To M1
        If Left(ActiveSheet.Cells(1, i).Value, 3) = "���հ�" Then
            count = count + 1
        End If
        Next i
        If count = 0 Then
            ttemp1.Value = "���հ�"
        Else
            ttemp1.Value = "���հ�" & count
        End If
        
        
        
        
        For i = 1 To N
            ttemp1.Offset(i, 0) = fitted(i - 1)
        Next i
        
        
        Set ttemp2 = ActiveSheet.Cells(1, M1 + 2)
        If ttemp1.Value = "���հ�" Then
        ttemp2.Value = "����"
        Else
        ttemp2.Value = "����" & count
        End If
        For i = 1 To N
           ttemp2.Offset(i, 0) = resi(i - 1)
        Next i
    End If

    
    
    
    
    
    Set TempSheet = ModuleControl.TransClassVar(cnt, cRn, vrn, sRn)
    Set resultsheet = OpenOutSheet("_���м����_", True)
    
   
    'resultsheet.Unprotect "prophet"
    
    '''
    '''
    '''
    RstSheet = "_���м����_"
    '����ϴ� �ش� ��⿡ �� ���� ����'
'������ �Է�
On Error GoTo Err_delete
Dim val3535 As Long '�ʱ���ġ ������ ����'
Dim s3535 As Worksheet
val3535 = 2
    For Each s3535 In ActiveWorkbook.Sheets
        If s3535.name = RstSheet Then
val3535 = Sheets(RstSheet).Cells(1, 1).Value
End If
Next s3535  '��Ʈ�� �̹������� ��� ��ġ �������ϰ�, ������ 2�� �����Ѵ�.
    'Worksheets(RstSheet).Unprotect "prophet"
    activePt = Worksheets(RstSheet).Cells(1, 1).Value
    
    '''Worksheets(RstSheet).Cells(1, 1).Value = "$A$" & Worksheets(RstSheet).Cells(1, 1).Value

    
    
    '�����跮�� ���ϱ� ���� �Լ�
    OneWay_Result.dResult mean, std, cnt, factor, cr, resultsheet
    '��л����
    OneWay_Result.eResult mean, vrn, std, cnt, cr, resultsheet
    '�л�м�
    OneWay_Result.aResult sRn, factor, st, SSE, tdf, edf, resultsheet
    '���ߺ�
    If (Frm_Multicom.Controls("CheckBox1").Value = True) Or _
        (Frm_Multicom.Controls("CheckBox2").Value = True) Or _
        (Frm_Multicom.Controls("CheckBox3").Value = True) Then
        OneWay_Result.cResult Me.ListBox2.list(0), mean, factor, cnt, SSE, tdf, edf, cr, _
        Frm_Multicom.Controls("TextBox1").Value, resultsheet, Frm_Multicom.Controls("CheckBox1").Value, _
        Frm_Multicom.Controls("CheckBox2").Value, Frm_Multicom.Controls("CheckBox3").Value
    End If
    
     M2 = ActiveSheet.Rows(1).Cells(1, 1).End(xlToRight).Column
    For i = 1 To M2
        If ActiveSheet.Rows(1).Cells(1, i).Value = Me.ListBox2.list(0) Then
            k = i
        End If
    Next i
    
    For i = 1 To M2
        If ActiveSheet.Rows(1).Cells(1, i).Value = Me.ListBox3.list(0) Then
            p = i
        End If
    Next i
    
    ActiveSheet.Rows(1).Cells(1, p).Offset(1, 0).Select
    Set y = Range(Selection, Selection.End(xlDown))

    ActiveSheet.Rows(1).Cells(1, k).Offset(1, 0).Select
    Set X = Range(Selection, Selection.End(xlDown))
    
 '====================================
    '������ �׸� ����
 '============================
    Set addr = resultsheet.Range("a1")                  'a1�� ��� �� �� ��ȣ�� �����
    Set ttemp3 = resultsheet.Range("a" & addr.Value)     '���� ��� ���� ��ġ
    
    'BoxPlot �׸���
    If CheckBox2.Value = True Then
        BoxPlotModule.MainBoxPlot sRn, _
        UBound(sRn), ttemp3.Offset(0, 1).Left, ttemp3.Top, resultsheet, VarName
        Set ttemp3 = ttemp3.Offset(21, 0)
        addr.Value = Right(ttemp3.Address, Len(ttemp3.Address) - 3)   '"a1"�� ���� ��µ� ��ġ ����
    End If
    
    
     If CheckBox4.Value = True Then
        
        '���հ�, ���� ��Ʈ
        Set res = Worksheets.Add
        res.Range("A1").Select
        For i = 1 To N
            Selection.Offset(i - 1, 0).Value = fitted(i - 1)
            Selection.Offset(i - 1, 1).Value = resi(i - 1)
        Next i
        Set fit = Range(Selection, Selection.End(xlDown))
        Set selvar = Range(Selection.Offset(0, 1), Selection.Offset(0, 1).End(xlDown))
        res.Visible = xlSheetHidden
    '���� ����Ȯ���� �׸���
        ChartOutControl posi, True
       ' Worksheets(RstSheet).Unprotect "prophet"
'        activePt = Worksheets(Rstsheet).Cells(1, 1).Value

        QQmodule.MainNormPlot selvar, posi(0), posi(1), Worksheets("_���м����_"), VarName:="����", NTest:=True
        
'        ChartOutControl 192, False
       ' Worksheets(RstSheet).Protect Password:="prophet", DrawingObjects:=False, _
                                    contents:=True, Scenarios:=True
                                    
    '���� ������ �׸���
'        ChartOutControl posi, True
       ' Worksheets(RstSheet).Unprotect "prophet"
        activePt = Worksheets(RstSheet).Cells(1, 1).Value

        scatterModule.OrderScatterPlot "_���м����_", Worksheets("_���м����_").Cells(activePt, 2).Offset(0, 4).Left, _
        Worksheets("_���м����_").Cells(activePt, 2).Offset(0, 4).Top, 200, 200, selvar, "����", 0

'        ChartOutControl 200, False
        'Worksheets(RstSheet).Protect Password:="prophet", DrawingObjects:=False, _
                                            contents:=True, Scenarios:=True

        '���� vs ���հ� ������ �׸���
        'ChartOutControl posi, True
       ' Worksheets(RstSheet).Unprotect "prophet"
        activePt = Worksheets(RstSheet).Cells(1, 1).Value

        scatterModule.ScatterPlot "_���м����_", Worksheets("_���м����_").Cells(activePt, 2).Offset(0, 8).Left, _
        Worksheets("_���м����_").Cells(activePt, 2).Offset(0, 8).Top, 200, 200, fit, selvar, "", "���հ�", "����", 0

        ChartOutControl 200, False
        'Worksheets(RstSheet).Protect Password:="prophet", DrawingObjects:=False, _
                                            contents:=True, Scenarios:=True
        
        Worksheets("_���м����_").Activate
        Worksheets(RstSheet).Cells(activePt + 5, 1).Select
        Worksheets(RstSheet).Cells(activePt + 5, 1).Activate
    End If
    
    Application.DisplayAlerts = False
    TempSheet.Delete
    Application.DisplayAlerts = True

    'resultsheet.Protect Password:="prophet", DrawingObjects:=False, contents:=True, Scenarios:=True
    
    
    Application.ScreenUpdating = False
   
    '''Worksheets(RstSheet).Cells(1, 1).Value = Right(Worksheets(RstSheet).Cells(1, 1).Value, Len(Worksheets(RstSheet).Cells(1, 1).Value) - 3)
    


    Worksheets(RstSheet).Activate

    '���� ���� üũ �� �񱳰� ����
    Dim Cmp_Value As Long
    
    If ModuleControl.ChkVersion(ActiveWorkbook.name) = True Then
        Cmp_Value = 1048000
    Else
        Cmp_Value = 65000
    End If
    
    If Worksheets(RstSheet).Cells(1, 1).Value > Cmp_Value Then
        MsgBox "[_���м����_]��Ʈ�� ���� ��� ����Ͽ����ϴ�." & vbCrLf & "�� ��Ʈ�� �̸��� �ٲٰų� ������ �ּ���", vbExclamation, "HIST"
        Exit Sub
    End If

    Worksheets(RstSheet).Cells(activePt + 10, 1).Select
    Worksheets(RstSheet).Cells(activePt + 10, 1).Activate
                            '��� �м��� ���۵Ǵ� �κ��� �����ָ� ��ģ��.
    Unload Frm_Multicom
    Unload Me
    


'�ǵڿ� ���̱�
Exit Sub
Err_delete:

For Each s3535 In ActiveWorkbook.Sheets
        If s3535.name = RstSheet Then
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

    ReDim myArray(TempSheet.UsedRange.Columns.count - 1)
    a = 0
   For i = 1 To TempSheet.UsedRange.Columns.count
   If arrName(i) <> "" Then                     '��ĭ����
   myArray(a) = arrName(i)
   a = a + 1
   
   Else:
   End If
   Next i
  
   Me.ListBox1.list() = myArray



End Sub
