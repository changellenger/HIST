VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameLe 
   OleObjectBlob   =   "frameLe.frx":0000
   Caption         =   "��л����"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7035
   StartUpPosition =   1  '������ ���
   TypeInfoVer     =   27
End
Attribute VB_Name = "frameLe"
Attribute VB_Base = "0{8E4303E1-E709-49AB-9A61-76DB9D9F0934}{56A30C33-3021-4E4A-9B0F-39DC587530EA}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private Sub CheckBox4_Click()

End Sub

Private Sub CommandButton11_Click()
    
    Dim i As Integer
    i = 0
    Do While i <= Me.Listbox1.ListCount - 1
        If Me.Listbox1.Selected(i) = True Then
           Me.ListBox2.AddItem Me.Listbox1.list(i)
           Me.Listbox1.RemoveItem (i)
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
    Do While i <= Me.Listbox1.ListCount - 1
        If Me.Listbox1.Selected(i) = True Then
           Me.ListBox3.AddItem Me.Listbox1.list(i)
           Me.Listbox1.RemoveItem (i)
           Me.CommandButton12.Visible = False
           Me.CommandButton13.Visible = True
           Exit Sub
        End If
        i = i + 1
    Loop
End Sub
Private Sub CommandButton13_Click()
    Me.Listbox1.AddItem Me.ListBox3.list(0)
    Me.ListBox3.RemoveItem (0)
    Me.CommandButton13.Visible = False
    Me.CommandButton12.Visible = True
End Sub
Private Sub CommandButton14_Click()
    Me.Listbox1.AddItem Me.ListBox2.list(0)
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
    Dim cnt() As Long: Dim mean() As Double: Dim Std()
        
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
    
    If FindingRangeError2(vrn) Then
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
    ModuleControl.PivotMakerforOneWay temp, Colname, valueName, cnt, mean, Std, factor
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
           

    
    
    
    
    Set TempSheet = ModuleControl.TransClassVar(cnt, cRn, vrn, sRn)
    Set resultsheet = OpenOutSheet2("_���м����_", True)
    
   
    
    '''
    '''
    '''
    rstSheet = "_���м����_"
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
    'Worksheets(RstSheet).Unprotect "prophet"
    activePt = Worksheets(rstSheet).Cells(1, 1).Value
    
   

    
    
   
    '��л����
    OneWay_Result.eResult mean, vrn, Std, cnt, cr, resultsheet

   
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
    

    
    If VarType(X(1)) < 2 Or VarType(X(1)) > 5 Then  '�з������� �����ϰ��
    Dim prtindex As Long
    Dim m3 As Integer
    Dim nx

    
    
    If TempSheet.Rows(1).Cells(1, 1).End(xlToRight).Column = 16384 Then
    m3 = 1
    Else
    m3 = TempSheet.Rows(1).Cells(1, 1).End(xlToRight).Column
    End If
    
    ReDim xnames(0 To cr - 1, 0 To 1)
    
    For i = 0 To cr - 1
        xnames(i, 0) = i + 1
        xnames(i, 1) = factor(i + 1)
    Next i
    ReDim nx(0 To N)
    TempSheet.Cells(1, m3) = X(0)

    For i = 1 To N
       For J = 1 To cr
           If X(i) = factor(J) Then
               TempSheet.Cells(i, 1) = xnames(J - 1, 0)
               nx(i) = TempSheet.Cells(i, 1)
                J = cr
            End If
        Next J
    Next i
        

    
    Else
    'scatterModule.ScatterPlot "_���м����_", posi(0), posi(1), 200, 200, X, y, xnames, "���� ����", "������ ��", 0
    End If
    
 
    
    Set addr = resultsheet.Range("a1")                  'a1�� ��� �� �� ��ȣ�� �����''''
    Set ttemp3 = resultsheet.Range("a" & addr.Value)     '���� ��� ���� ��ġ
    

 
     If CheckBox4.Value = True Then
        
    
        Set res = Worksheets.Add
        res.Range("A1").Select
        For i = 1 To N
            Selection.Offset(i - 1, 0).Value = fitted(i - 1)
            Selection.Offset(i - 1, 1).Value = resi(i - 1)
        Next i
        Set fit = Range(Selection, Selection.End(xlDown))
        Set selvar = Range(Selection.Offset(0, 1), Selection.Offset(0, 1).End(xlDown))
        res.Visible = xlSheetHidden

       ' Worksheets(RstSheet).Unprotect "prophet"
        activePt = Worksheets(rstSheet).Cells(1, 1).Value

    End If
    
    Application.DisplayAlerts = False
    TempSheet.Delete
    Application.DisplayAlerts = True

   
    
    Application.ScreenUpdating = False
   



    Worksheets(rstSheet).Activate

    '���� ���� üũ �� �񱳰� ����
    Dim Cmp_Value As Long
    
    If PublicModule.ChkVersion(ActiveWorkbook.Name) = True Then
        Cmp_Value = 1048000
    Else
        Cmp_Value = 65000
    End If
    
    If Worksheets(rstSheet).Cells(1, 1).Value > Cmp_Value Then
        MsgBox "[_���м����_]��Ʈ�� ���� ��� ����Ͽ����ϴ�." & vbCrLf & "�� ��Ʈ�� �̸��� �ٲٰų� ������ �ּ���", vbExclamation, "HIST"
        Exit Sub
    End If

    Worksheets(rstSheet).Cells(activePt + 10, 1).Select
    Worksheets(rstSheet).Cells(activePt + 10, 1).Activate
                            '��� �м��� ���۵Ǵ� �κ��� �����ָ� ��ģ��.
  
    Unload Me
    


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

Private Sub Frame4_Click()

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
