Attribute VB_Name = "ModuleControl"
Public DataSheet As String, RstSheet As String      'sheet�̸� �� ��

                                                    '�̻� Public���� 7��
                                                    '��� frmRegression ���� �ѹ��� �����ǰ�
                                                    '�ٸ� �������� �ٲ��� �ʴ´�.
'������ ���� ���� �Լ�    Winchm�� ���� �ο�
Public Declare Function ShellExecute _
 Lib "shell32.dll" _
 Alias "ShellExecuteA" ( _
 ByVal hwnd As Long, _
 ByVal lpOperation As String, _
 ByVal lpFile As String, _
 ByVal lpParameters As String, _
 ByVal lpDirectory As String, _
 ByVal nShowCmd As Long) _
 As Long

Sub FactDe()
    
    InitializeDlg doe1
    doe1.TextBox1.value = doe1.SpinButton1.value
    Dim a, b, c As Integer
    Dim MyArray1, MyArray2, MyArray3 As Variant

        
    doe1.ComboBox4.ColumnCount = 1
    doe1.ComboBox2.ColumnCount = 1
    doe1.ComboBox3.ColumnCount = 1
    
    MyArray1 = [{0;1;2;3;4;5}]  '�߽��� ����
    MyArray2 = [{1;2;3;4;5}]    '�ݺ���
    MyArray3 = [{1}]
        
    doe1.ComboBox4.list = MyArray1
    doe1.ComboBox2.list = MyArray2
    doe1.ComboBox3.list = MyArray3
    doe1.Show

End Sub


Sub FactAnal()
    If CheckSheetError = False Then
        InitializeDlg doe33
        For i = 0 To doe33.ListBox1.ListCount - 1
            If doe33.ListBox1.list(i) = "Block" Or Left(doe33.ListBox1.list(i), 2) = "����" Then
                e = e + 1
            End If
        Next i

        For i = 0 To e - 1
            doe33.ListBox3.AddItem doe33.ListBox1.list(i + 2)
        Next i
        
        
        doe33.Show
     
    End If
End Sub
Function FindVarCount(ListVar) As Long
   
    Dim temp, m2, m3, j As Long
    Dim TempSheet As Worksheet
    Dim tmp2, tmp As Range
    
    Set TempSheet = Worksheets(DataSheet)
    temp = Cells.CurrentRegion.Columns.count
    
   Dim Chk_Ver As Boolean   '���� ���� üũ, 2009.01.02 ���ο� �߰�
   Dim Cmp_R As Long        '���� ������ ���� �� ���� ��, 2009.01.02 ���ο� �߰�
   
   '���� ������ ���� ��� ���� �񱳰� ����, 2009.01.02 ���ο� �߰�
   Chk_Ver = ChkVersion(ActiveWorkbook.name)
   If Chk_Ver = True Then
        Cmp_R = 1048576
    Else
        Cmp_R = 65536
    End If


    For j = 1 To temp
       If StrComp(ListVar, TempSheet.Cells(1, j).value, 1) = 0 Then
          Set tmp2 = TempSheet.Columns(j)
          m2 = tmp2.Cells(1, 1).End(xlDown).row
          If m2 <> Cmp_R Then
             m3 = tmp2.Cells(m2, 1).End(xlDown).row
             If m3 <> Cmp_R Then m2 = m3
          End If
          Set tmp = tmp2.Range(Cells(2, 1), Cells(m2, 1))
       End If
    Next j
    
    FindVarCount = tmp.count
    
End Function
Function ChkVersion(File_Name) As Boolean
    
    If Right(File_Name, 4) = ".xls" Or Right(File_Name, 4) = ".XLS" Then
        ChkVersion = False
    Else
        ChkVersion = True
    End If
End Function
Function CheckSheetError() As Boolean
    
    On Error GoTo ErrorFlag
    
    Set myRange = ActiveSheet.Cells.CurrentRegion
    If myRange.count = 1 And myRange.Cells(1, 1) = "" Then
        MsgBox "��Ʈ�� ����Ÿ�� �ִ��� Ȯ���Ͻʽÿ�." & Chr(10) & _
               "1��1������ �����̸��� �Է��ؾ� �մϴ�.", _
               vbExclamation, "SQI"
        CheckSheetError = True: Exit Function
    End If
    CheckSheetError = False: Exit Function
ErrorFlag:
    MsgBox "��Ʈ�� ��ȣ���¿� �ֽ��ϴ�." & Chr(10) & _
           "����Ÿ�� ���� �� �����ϴ�.", _
            vbExclamation, "SQI"
    CheckSheetError = True
End Function

Function InitializeDlg(ParentDlg) As Integer
   
   Dim myRange As Range: Dim cnt As Long
   Dim myArray() As String
   
   On Error GoTo ErrorFlag
   
   Set myRange = ActiveSheet.Cells.CurrentRegion
   If myRange.count = 1 And myRange.Cells(1, 1) = "" Then
        InitializeDlg = 1: Exit Function
   End If
   Set myRange = ActiveSheet.Cells.CurrentRegion.rows(1)
   ParentDlg.ListBox1.Clear
': ParentDlg.ListBox2.Clear: ParentDlg.ListBox3.Clear
   cnt = myRange.Cells.count
   
   ReDim myArray(0 To cnt - 1)
   For i = 1 To cnt
     myArray(i - 1) = myRange.Cells(i)
   Next i
   ParentDlg.ListBox1.list() = myArray
   InitializeDlg = 0
   Exit Function
   
ErrorFlag:
   InitializeDlg = -1
   
End Function

Function OpenOutSheet(sheetname, Optional IsAddress As Boolean = False) As Worksheet
    
    Dim s, CurS As Worksheet
    
    Application.ScreenUpdating = False
    For Each s In ActiveWorkbook.Sheets
        If s.name = sheetname Then
            Set OpenOutSheet = s
            Exit Function
        End If
    Next s
    Set CurS = ActiveSheet: Set s = Worksheets.Add
    With ActiveWindow
        .DisplayGridlines = False
'        .DisplayHeadings = False
    End With
    
    With ActiveWindow.Application.Cells
         .Font.name = "����"
         .Font.Size = 9
         .HorizontalAlignment = xlRight
    End With

    s.name = sheetname: CurS.Activate
    With Worksheets(sheetname).Range("a1")
        .value = 2
        '''If IsAddress = True Then .Value = "A2"
        .Font.ColorIndex = 2
    End With
    Worksheets(sheetname).rows(1).Hidden = True

    's.Protect Password:="prophet", DrawingObjects:=False, contents:=True, Scenarios:=True
    Application.ScreenUpdating = True
    
    Set OpenOutSheet = s
    
End Function

Function FindingRangeError(Rn) As Boolean
    
    Dim tmp1 As Range: Dim tmp2 As Range
    Dim tmp3 As Range
    
    On Error Resume Next
    
    If Application.CountBlank(Rn) >= 1 Then
        FindingRangeError = True
        Exit Function
    End If
    Set tmp1 = Rn.SpecialCells(xlCellTypeConstants, 22)
    Set tmp2 = Rn.SpecialCells(xlCellTypeFormulas, 22)
    Set tmp3 = Rn.SpecialCells(xlCellTypeBlanks)
    
    If Rn.count = 1 And IsNumeric(Rn.Cells(1, 1)) = True Then
        FindingRangeError = False
    Else
        If tmp1 Is Nothing And tmp2 Is Nothing And tmp3 Is Nothing Then
            FindingRangeError = False
        Else: FindingRangeError = True
        End If
    End If
    
End Function



Function SelectedVariable(ParentDlgLbxValue, SelVar, _
         IsRowData As Boolean) As String
   
   Dim temp, m2, m3 As Long
   Dim TempSheet As Worksheet
   Dim tmp2, tmp As Range
   
   Set TempSheet = ActiveCell.Worksheet
   

   Dim Chk_Ver As Boolean   '���� ���� üũ, 2009.01.02 ���ο� �߰�
   Dim Cmp_R As Long        '���� ������ ���� �� ���� ��, 2009.01.02 ���ο� �߰�
   Dim Cmp_C As Integer     '���� ������ ���� �� ���� ��, 2009.01.02 ���ο� �߰�
   
   '���� ������ ���� ��� ���� �񱳰� ����, 2009.01.02 ���ο� �߰�
   Chk_Ver = ChkVersion(ActiveWorkbook.name)
   If Chk_Ver = True Then
        Cmp_R = 1048576
        Cmp_C = 16384
    Else
        Cmp_R = 65536
        Cmp_C = 256
    End If
    
   If IsRowData = True Then
        temp = Cells.CurrentRegion.Columns.count
        For j = 1 To temp
           If StrComp(ParentDlgLbxValue, TempSheet.Cells(1, j).value, 1) = 0 Then
              Set tmp2 = TempSheet.Columns(j)
              m2 = tmp2.Cells(1, 1).End(xlDown).row
              If m2 <> Cmp_R Then
                 m3 = tmp2.Cells(m2, 1).End(xlDown).row
                 If m3 <> Cmp_R Then m2 = m3
              End If
              Set tmp = tmp2.Range(Cells(2, 1), Cells(m2, 1))
           End If
        Next j
   Else
        temp = Cells.CurrentRegion.rows.count
        For j = 1 To temp
           If StrComp(ParentDlgLbxValue, TempSheet.Cells(j, 1).value, 1) = 0 Then
              Set tmp2 = TempSheet.rows(j)
              m2 = tmp2.Cells(1, 1).End(xlToRight).Column
              If m2 <> Cmp_C Then
                 m3 = tmp2.Cells(1, m2).End(xlToRight).Column
                 If m3 <> Cmp_C Then m2 = m3
              End If
              Set tmp = tmp2.Range(Cells(1, 2), Cells(1, m2))
           End If
        Next j
   End If
    
   Set SelVar = tmp
   
   If IsNull(ParentDlgLbxValue) = True Then
        SelectedVariable = ""
   Else: SelectedVariable = ParentDlgLbxValue
   End If

End Function



Function TransClassVar(cnt, classRn, ValueRn, sRn) As Worksheet
     
    Dim TempSheet As Worksheet: Dim i, mystart, myend As Long
    
    Set TempSheet = Worksheets.Add: TempSheet.Visible = xlSheetHidden
    classRn.Copy: TempSheet.Paste TempSheet.Cells(1, 1)
    ValueRn.Copy: TempSheet.Paste TempSheet.Cells(1, 2)
    TempSheet.Cells(1, 1).Sort _
        Key1:=TempSheet.Cells(1, 1), _
        Order1:=xlAscending, Header:=xlGuess
    mystart = 1: myend = cnt(1)
    ReDim sRn(1 To UBound(cnt) - 1)
    For j = 1 To UBound(sRn)
        Set sRn(j) = Range(TempSheet.Cells(mystart, 2), TempSheet.Cells(myend, 2))
        If j = UBound(sRn) Then Exit For
        mystart = myend + 1: myend = mystart + cnt(j + 1) - 1
    Next j
    
    Set TransClassVar = TempSheet
    
End Function

Sub PivotMaker(DataRn, ColVn, DataVn, _
    cnt, ave, st, factor, count, t, nn, max1)
        
    Dim actSh, tmpSh As Worksheet
    Dim StartCell As String: Dim i, n As Long
    Dim temp As Range
    
    Set actSh = ActiveSheet
    Set tmpSh = Worksheets.Add
    actSh.Select
    StartCell = tmpSh.name & "!R1C1"
    ActiveSheet.PivotTableWizard SourceType:=xlDatabase, SourceData:= _
        DataRn, TableDestination:=StartCell, TableName:="�ǹ� ���̺�1"
    
    ActiveSheet.PivotTables("�ǹ� ���̺�1").AddFields ColumnFields:=ColVn
    ActiveSheet.PivotTables("�ǹ� ���̺�1").PivotFields(DataVn).Orientation = _
        xlDataField
    ActiveSheet.PivotTables("�ǹ� ���̺�1").PivotFields(tmpSh.Cells(1, 1).value).Function = xlCount
    ActiveSheet.PivotTables("�ǹ� ���̺�1").PivotSelect "", xlDataOnly
    n = Selection.Columns.count
    If count = 0 Then
    max1 = n
    End If
    If n > max1 Then
    max1 = n
    ReDim Preserve cnt(t - 1, n - 1)
    ReDim Preserve ave(t - 1, n - 1)
    ReDim Preserve st(t - 1, n - 1)
    ReDim Preserve factor(t - 1, n - 1)
    Else
    ReDim Preserve cnt(t - 1, max1 - 1)
    ReDim Preserve ave(t - 1, max1 - 1)
    ReDim Preserve st(t - 1, max1 - 1)
    ReDim Preserve factor(t - 1, max1 - 1)
    End If
    For i = 1 To n
        cnt(count, i - 1) = Selection.Cells(i)
    Next i
    ActiveSheet.PivotTables("�ǹ� ���̺�1").PivotFields(tmpSh.Cells(1, 1).value).Function = xlAverage
    ActiveSheet.PivotTables("�ǹ� ���̺�1").PivotSelect "", xlDataOnly
    For i = 1 To n
        ave(count, i - 1) = Selection.Cells(i)
    Next i
    
    Set temp = Selection.Offset(-1, 0)
    For i = 1 To n
        factor(count, i - 1) = temp.Cells(i)
    Next i
    
    ActiveSheet.PivotTables("�ǹ� ���̺�1").PivotFields(tmpSh.Cells(1, 1).value).Function = xlStDev
    ActiveSheet.PivotTables("�ǹ� ���̺�1").PivotSelect "", xlDataOnly
    For i = 1 To n
        st(count, i - 1) = Selection.Cells(i)
    Next i
    nn(count + 1) = n
    Application.DisplayAlerts = False
    tmpSh.Delete
    Application.DisplayAlerts = True
    
End Sub
