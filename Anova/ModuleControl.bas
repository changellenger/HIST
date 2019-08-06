Attribute VB_Name = "ModuleControl"
Public RstSheet As String
Public DataSheet As String     'sheet이름 두 개

                                                    '이상 Public변수 7개
                                                    '모두 frmRegression 에서 한번만 지정되고
                                                    '다른 곳에서는 바꾸지 않는다.
'도움말을 쓰기 위한 함수    Winchm의 도움말 인용
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

Function SelectedVariable(ParentDlgLbxValue, selvar, _
         IsRowData As Boolean) As String
   
   Dim temp, M2, m3 As Long
   Dim TempSheet As Worksheet
   Dim tmp2, tmp As Range
   
   Set TempSheet = ActiveCell.Worksheet
   

   Dim Chk_Ver As Boolean   '파일 버전 체크
   Dim Cmp_R As Long        '파일 버전에 따른 비교 행의 값
   Dim Cmp_C As Integer     '파일 버전에 따른 비교 열의 값
   
   '파일 버전에 따른 행과 열의 비교값 정의
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
        For J = 1 To temp
           If StrComp(ParentDlgLbxValue, TempSheet.Cells(1, J).Value, 1) = 0 Then
              Set tmp2 = TempSheet.Columns(J)
              M2 = tmp2.Cells(1, 1).End(xlDown).Row
              If M2 <> Cmp_R Then
                 m3 = tmp2.Cells(M2, 1).End(xlDown).Row
                 If m3 <> Cmp_R Then M2 = m3
              End If
              Set tmp = tmp2.Range(Cells(2, 1), Cells(M2, 1))
           End If
        Next J
   Else
        temp = Cells.CurrentRegion.Rows.count
        For J = 1 To temp
           If StrComp(ParentDlgLbxValue, TempSheet.Cells(J, 1).Value, 1) = 0 Then
              Set tmp2 = TempSheet.Rows(J)
              M2 = tmp2.Cells(1, 1).End(xlToRight).Column
              If M2 <> Cmp_C Then
                 m3 = tmp2.Cells(1, M2).End(xlToRight).Column
                 If m3 <> Cmp_C Then M2 = m3
              End If
              Set tmp = tmp2.Range(Cells(1, 2), Cells(1, M2))
           End If
        Next J
   End If
    
   Set selvar = tmp
   
   If IsNull(ParentDlgLbxValue) = True Then
        SelectedVariable = ""
   Else: SelectedVariable = ParentDlgLbxValue
   End If

End Function

Function FindingRangeError(rn) As Boolean
    
    Dim tmp1 As Range: Dim tmp2 As Range
    Dim tmp3 As Range
    
    On Error Resume Next
    
    If Application.CountBlank(rn) >= 1 Then
        FindingRangeError = True
        Exit Function
    End If
    Set tmp1 = rn.SpecialCells(xlCellTypeConstants, 22)
    Set tmp2 = rn.SpecialCells(xlCellTypeFormulas, 22)
    Set tmp3 = rn.SpecialCells(xlCellTypeBlanks)
    
    If rn.count = 1 And IsNumeric(rn.Cells(1, 1)) = True Then
        FindingRangeError = False
    Else
        If tmp1 Is Nothing And tmp2 Is Nothing And tmp3 Is Nothing Then
            FindingRangeError = False
        Else: FindingRangeError = True
        End If
    End If
    
End Function

Function CheckSheetError() As Boolean
    
    On Error GoTo ErrorFlag
    
    Set myRange = ActiveSheet.Cells.CurrentRegion
    If myRange.count = 1 And myRange.Cells(1, 1) = "" Then
        MsgBox "시트에 데이타가 있는지 확인하십시오." & Chr(10) & _
               "1행1열부터 변수이름을 입력해야 합니다.", _
               vbExclamation, "HIST"
        CheckSheetError = True: Exit Function
    End If
    CheckSheetError = False: Exit Function
ErrorFlag:
    MsgBox "시트가 보호상태에 있습니다." & Chr(10) & _
           "데이타를 읽을 수 없습니다.", _
            vbExclamation, "HIST"
    CheckSheetError = True
End Function
Function CheckSheetError1() As Boolean
    
    On Error GoTo ErrorFlag
    
    Set myRange = ActiveSheet.Cells.CurrentRegion
    If myRange.count = 1 Or myRange.Cells(1, 1) = "" Then
        MsgBox "시트에 데이타가 있는지 확인하십시오." & Chr(10) & _
               "1행1열부터 변수이름을 입력해야 합니다.", _
               vbExclamation, "HIST"
        CheckSheetError1 = True: Exit Function
    End If
    CheckSheetError1 = False: Exit Function
ErrorFlag:
    MsgBox "시트가 보호상태에 있습니다." & Chr(10) & _
           "데이타를 읽을 수 없습니다.", _
            vbExclamation, "HIST"
    CheckSheetError1 = True
End Function




Sub Showbratio1()
    bratio1.Show
End Sub
Sub Showbratio2()
    bratio2.Show
End Sub
Sub TwoWay_Anova1()
    If CheckSheetError = False Then
        Frm2_way1.MultiPage1.Value = 0
        Frm2_way1.Show
    End If
End Sub
Sub OneWay_Anova()
    If CheckSheetError1 = False Then
        Frm1_way.MultiPage1.Value = 0
        Frm1_way.Show
    End If
End Sub
Sub conti()
    If CheckSheetError = False Then
        Conti_Frm.Show
    End If
End Sub
Sub loglinear()
    If CheckSheetError = False Then
        InitializeDlgl Frm_loglinear
        Frm_loglinear.Show
    End If
End Sub
Sub loglinear1()
    If CheckSheetError = False Then
        Frm_loglinear1.MultiPage1.Value = 0
        InitializeDlg2 Frm_loglinear1
        Frm_loglinear1.Show
    End If
End Sub
Sub good_initialize()
    If CheckSheetError = False Then
        good.Show
    End If
End Sub
Sub TwoWay_Anova2()
    If CheckSheetError = False Then
        Frm2_way2.MultiPage1.Value = 0
        Frm2_way2.Show
    End If
End Sub

'''SheetName의 시트를 만든다.
'''이미 만들어져 있을 경우는 다시 만들지 않는다.
'''A1셀에 인쇄할 장소를 적어둔다.
Function OpenOutSheet1(SheetName, Optional IsAddress As Boolean = False) As Worksheet
    
    Dim s, CurS As Worksheet
    
    Application.ScreenUpdating = False
    For Each s In ActiveWorkbook.Sheets
        If s.name = SheetName Then
            Set OpenOutSheet1 = s
            Exit Function
        End If
    Next s
    Set CurS = ActiveSheet: Set s = Worksheets.Add
    s.name = SheetName: CurS.Activate
    
    Set OpenOutSheet1 = s
    
End Function
Function OpenOutSheet(SheetName, Optional IsAddress As Boolean = False) As Worksheet
    
    Dim s, CurS As Worksheet
    
    Application.ScreenUpdating = False
    For Each s In ActiveWorkbook.Sheets
        If s.name = SheetName Then
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
         .Font.name = "굴림"
         .Font.Size = 9
         .HorizontalAlignment = xlRight
    End With

    s.name = SheetName: CurS.Activate
    With Worksheets(SheetName).Range("a1")
        .Value = 2
        '''If IsAddress = True Then .Value = "A2"
        .Font.ColorIndex = 2
    End With
    Worksheets(SheetName).Rows(1).Hidden = True
    
    's.Protect Password:="prophet", DrawingObjects:=False, contents:=True, Scenarios:=True
    Application.ScreenUpdating = True
    
    Set OpenOutSheet = s
    
End Function

Function myStdev(rn)
    On Error GoTo errcontrol
    myStdev = Application.StDev(rn)
    Exit Function
errcontrol:
    myStdev = "#N/A"
End Function
Sub ShowfrmTTest2()
    On Error Resume Next
    DlgShow frmTtest2, 6
End Sub



Function InitializeDlg(ParentDlg) As Integer
   
   Dim myRange As Range: Dim cnt As Long
   Dim myArray() As String
   
   On Error GoTo ErrorFlag
   
   Set myRange = ActiveSheet.Cells.CurrentRegion
   If myRange.count = 1 And myRange.Cells(1, 1) = "" Then
        InitializeDlg = 1: Exit Function
   End If
   Set myRange = ActiveSheet.Cells.CurrentRegion.Rows(1)
   ParentDlg.ListBox4.Clear: ParentDlg.ListBox5.Clear
   cnt = myRange.Cells.count
   
   ReDim myArray(0 To cnt - 1)
   For i = 1 To cnt
     myArray(i - 1) = myRange.Cells(i)
   Next i
   ParentDlg.ListBox4.list() = myArray
   InitializeDlg = 0
   Exit Function
   
ErrorFlag:
   InitializeDlg = -1
   
End Function
Function InitializeDlgl(ParentDlg) As Integer
   
   Dim myRange As Range: Dim cnt As Long
   Dim myArray() As String
   
   On Error GoTo ErrorFlag
   
   Set myRange = ActiveSheet.Cells.CurrentRegion
   If myRange.count = 1 And myRange.Cells(1, 1) = "" Then
        InitializeDlgl = 1: Exit Function
   End If
   Set myRange = ActiveSheet.Cells.CurrentRegion.Rows(1)
   ParentDlg.ListBox1.Clear
   ParentDlg.ListBox2.Clear
   
   cnt = myRange.Cells.count
   
   ReDim myArray(0 To cnt - 1)
   For i = 1 To cnt
     myArray(i - 1) = myRange.Cells(i)
   Next i
   ParentDlg.ListBox2.list() = myArray
   InitializeDlgl = 0
   Exit Function
   
ErrorFlag:
   InitializeDlgl = -1
   
End Function
Function InitializeDlg2(ParentDlg) As Integer
   
   Dim myRange As Range: Dim cnt As Long
   Dim myArray() As String
   
   On Error GoTo ErrorFlag
   
   Set myRange = ActiveSheet.Cells.CurrentRegion
   If myRange.count = 1 And myRange.Cells(1, 1) = "" Then
        InitializeDlg2 = 1: Exit Function
   End If
   Set myRange = ActiveSheet.Cells.CurrentRegion.Rows(1)
   ParentDlg.ListBox3.Clear
   ParentDlg.ListBox4.Clear
   ParentDlg.ListBox5.Clear
   ParentDlg.CommandButton10.Visible = True
   ParentDlg.CommandButton11.Visible = False
   cnt = myRange.Cells.count
   
   ReDim myArray(0 To cnt - 1)
   For i = 1 To cnt
     myArray(i - 1) = myRange.Cells(i)
   Next i
   ParentDlg.ListBox4.list() = myArray
   InitializeDlg2 = 0
   Exit Function
   
ErrorFlag:
   InitializeDlg2 = -1
   
End Function
Sub SetUpforPage3(ParentDlg, opt As Integer)

   Dim myRange As Range: Dim cnt As Long
   Dim myArray() As String
   
   Set myRange = ActiveSheet.Cells.CurrentRegion.Rows(1)
   ParentDlg.ListBox1.Clear
   ParentDlg.ListBox2.Clear
   
   
   cnt = myRange.Cells.count
   
   ReDim myArray(0 To cnt - 1)
   For i = 1 To cnt
     If myRange.Cells(i) = "" Then
        MsgBox "자료의 입력방식을 확인하시기 바랍니다.", vbExclamation, "HIST"
        Exit Sub
     End If
     myArray(i - 1) = myRange.Cells(i)
   Next i
   
   ParentDlg.ListBox2.list() = myArray
   
End Sub

Sub SetUpforPage2(ParentDlg, opt As Integer)

   Dim myRange As Range: Dim cnt As Long
   Dim myArray() As String
   
   Set myRange = ActiveSheet.Cells.CurrentRegion.Rows(1)
   ParentDlg.ListBox1.Clear: ParentDlg.ListBox2.Clear: ParentDlg.ListBox3.Clear
   If opt = 2 Then
        ParentDlg.ListBox4.Clear
        ParentDlg.CheckBox3.Value = True
        ParentDlg.CommandButton19.Visible = False
        ParentDlg.CommandButton18.Visible = True
   End If
   If opt <> 3 Then ParentDlg.CheckBox2.Value = True
   ParentDlg.CheckBox3.Value = True
   If opt = 3 Then ParentDlg.CheckBox3.Value = True
   ParentDlg.CommandButton13.Visible = False
   ParentDlg.CommandButton12.Visible = True
   ParentDlg.CommandButton14.Visible = False
   ParentDlg.CommandButton11.Visible = True
   cnt = myRange.Cells.count
   
   ReDim myArray(0 To cnt - 1)
   For i = 1 To cnt
     If myRange.Cells(i) = "" Then
        MsgBox "변수명에 공백이 있습니다.", vbExclamation, "HIST"
        Exit Sub
     End If
     myArray(i - 1) = myRange.Cells(i)
   Next i
   
   ParentDlg.ListBox1.list() = myArray
   
End Sub

Function PivotMakerforTwoWay(DataRn, RowVn, ColVn, DataVn, _
    cnt, ave, st, Colname, Rowname, Optional DoClose As Boolean = True) _
    As Worksheet
        
    Dim actSh, tmpSh As Worksheet
    Dim StartCell As String: Dim i, J, m, N As Long
    Dim temp As Range
    
    Set actSh = ActiveSheet
    Set tmpSh = Worksheets.Add
    actSh.Select
    StartCell = tmpSh.name & "!R1C1"
    ActiveSheet.PivotTableWizard SourceType:=xlDatabase, SourceData:= _
        DataRn, TableDestination:=StartCell, TableName:="피벗 테이블1"
    
    ActiveSheet.PivotTables("피벗 테이블1").AddFields RowFields:=RowVn, _
        ColumnFields:=ColVn
    ActiveSheet.PivotTables("피벗 테이블1").PivotFields(DataVn).Orientation = _
        xlDataField
                                                        ''' "합계 : " & DataVn
    ActiveSheet.PivotTables("피벗 테이블1").PivotFields(tmpSh.Cells(1, 1).Value).Function = xlCount
    ActiveSheet.PivotTables("피벗 테이블1").PivotSelect "", xlDataOnly
    m = Selection.Rows.count: N = Selection.Columns.count
    ReDim cnt(1 To m, 1 To N): ReDim ave(1 To m, 1 To N): ReDim st(1 To m, 1 To N)
    ReDim Colname(1 To N): ReDim Rowname(1 To m)
    For i = 1 To m: For J = 1 To N
        cnt(i, J) = Selection.Cells(i, J)
    Next J: Next i
    Set temp = Selection.Offset(-1, -1)
    For i = 1 To m - 1: Rowname(i) = temp.Cells(i + 1, 1): Next i
    For J = 1 To N - 1: Colname(J) = temp.Cells(1, J + 1): Next J
                                                        '''  "개수 : " & DataVn
    ActiveSheet.PivotTables("피벗 테이블1").PivotFields(tmpSh.Cells(1, 1).Value).Function = xlAverage
    ActiveSheet.PivotTables("피벗 테이블1").PivotSelect "", xlDataOnly
    For i = 1 To m: For J = 1 To N
        ave(i, J) = Selection.Cells(i, J)
    Next J: Next i
                                                        '''  "평균 : " & DataVn
    ActiveSheet.PivotTables("피벗 테이블1").PivotFields(tmpSh.Cells(1, 1).Value).Function = xlStDev
    ActiveSheet.PivotTables("피벗 테이블1").PivotSelect "", xlDataOnly
    For i = 1 To m: For J = 1 To N
        st(i, J) = Selection.Cells(i, J)
    Next J: Next i
    
    If DoClose = True Then
        Application.DisplayAlerts = False
        tmpSh.Delete
        Application.DisplayAlerts = True
    End If
    
    Set PivotMakerforTwoWay = tmpSh
    
End Function
Function FindVarCount(ListVar) As Long
   
Dim temp, M2, m3, J As Long
Dim TempSheet As Worksheet
Dim tmp2, tmp As Range
    
Set TempSheet = Worksheets(DataSheet)
temp = Cells.CurrentRegion.Columns.count
    
Dim Chk_Ver As Boolean   '파일 버전 체크
Dim Cmp_R As Long        '파일 버전에 따른 비교 행의 값
   
'파일 버전에 따른 행과 열의 비교값 정의
Chk_Ver = ChkVersion(ActiveWorkbook.name)
If Chk_Ver = True Then
    Cmp_R = 1048576
Else
    Cmp_R = 65536
End If

For J = 1 To temp
    If StrComp(ListVar, TempSheet.Cells(1, J).Value, 1) = 0 Then
        Set tmp2 = TempSheet.Columns(J)
        M2 = tmp2.Cells(1, 1).End(xlDown).Row
        If M2 <> Cmp_R Then
            m3 = tmp2.Cells(M2, 1).End(xlDown).Row
            If m3 <> Cmp_R Then M2 = m3
        End If
        Set tmp = tmp2.Range(Cells(2, 1), Cells(M2, 1))
    End If
Next J
    
FindVarCount = tmp.count
    
End Function
Sub PivotMaker(DataRn, ColVn, DataVn, _
    cnt, st, factor, count, t, nn, max1, es)
        
Dim actSh, tmpSh As Worksheet
Dim StartCell As String: Dim i, N As Long
Dim temp As Range
    
cc = 0: es = True
Set actSh = ActiveSheet
Set tmpSh = Worksheets.Add
actSh.Select
StartCell = tmpSh.name & "!R1C1"
ActiveSheet.PivotTableWizard SourceType:=xlDatabase, SourceData:= _
    DataRn, TableDestination:=StartCell, TableName:="피벗 테이블1"
    
ActiveSheet.PivotTables("피벗 테이블1").AddFields ColumnFields:=ColVn
ActiveSheet.PivotTables("피벗 테이블1").PivotFields(DataVn).Orientation = _
        xlDataField
ActiveSheet.PivotTables("피벗 테이블1").PivotFields(tmpSh.Cells(1, 1).Value).Function = xlCount
ActiveSheet.PivotTables("피벗 테이블1").PivotSelect "", xlDataOnly
    
N = Selection.Columns.count
    
If count = 0 Then
    max1 = N
End If

If N > max1 Then
    max1 = N
    ReDim Preserve cnt(t - 1, N - 1)
    ReDim Preserve factor(t - 1, N - 1)
Else
    ReDim Preserve cnt(t - 1, max1 - 1)
    ReDim Preserve factor(t - 1, max1 - 1)
End If
    
For i = 1 To N
    cnt(count, i - 1) = Selection.Cells(i)
Next i
    
Set temp = Selection.Offset(-1, 0)
For i = 1 To N
    factor(count, i - 1) = temp.Cells(i)
Next i

nn(count + 1) = N
Application.DisplayAlerts = False
tmpSh.Delete
Application.DisplayAlerts = True
    
End Sub
Sub PivotMaker1(DataRn, xlist, DataVn, _
    cnt, st, factor, count, t, nn, max1, nmul, es)
        
Dim actSh, tmpSh As Worksheet
Dim StartCell As String: Dim i, N As Long
Dim temp As Range
    
cc = 0: es = True
Set actSh = ActiveSheet
Set tmpSh = Worksheets.Add
actSh.Select
StartCell = tmpSh.name & "!R1C1"
ActiveSheet.PivotTableWizard SourceType:=xlDatabase, SourceData:= _
    DataRn, TableDestination:=StartCell, TableName:="피벗 테이블1"
    
ActiveSheet.PivotTables("피벗 테이블1").AddFields ColumnFields:=xlist
ActiveSheet.PivotTables("피벗 테이블1").PivotFields(DataVn).Orientation = _
    xlDataField
ActiveSheet.PivotTables("피벗 테이블1").PivotFields(tmpSh.Cells(1, 1).Value).Function = xlCount
ActiveSheet.PivotTables("피벗 테이블1").PivotSelect "", xlDataOnly
    
N = Selection.Columns.count
    
If count = 0 Then
    max1 = N
End If

If N > max1 Then
    max1 = N
    ReDim cnt(t, N - 1)
    ReDim factor(t - 1, N - 1)
Else
    ReDim cnt(t, max1 - 1)
    ReDim factor(t - 1, max1 - 1)
End If

For i = 1 To N
    cnt(0, i - 1) = Selection.Cells(i)
Next i
    
For J = 0 To t - 1
Set temp = Selection.Offset(-(J + 1), 0)
For i = 1 To N
    factor(J, i - 1) = temp.Cells(i)
Next i
Next J

For i = 1 To N
    If factor(0, i - 1) = "" Then
        cnt(0, i - 1) = ""
    End If
Next i
    
nmul = 1
For i = 1 To t
    nmul = nmul * (nn(i) - 1)
Next i
    
For i = 1 To nmul
    For J = 1 To N - 1
        If cnt(0, J - 1) = "" Then
            cnt(0, J - 1) = cnt(0, J)
            cnt(0, J) = ""
        End If
    Next J
Next i
    
Application.DisplayAlerts = False
tmpSh.Delete
Application.DisplayAlerts = True
    
End Sub

Sub PivotMaker2(DataRn, ColVn, DataVn, cnt, st, factor, count, t, nn, max1, nmul, es)
        
Dim actSh, tmpSh As Worksheet
Dim StartCell As String: Dim i, N As Long
Dim temp As Range
    
cc = 0: es = True
Set actSh = ActiveSheet
Set tmpSh = Worksheets.Add
actSh.Select
StartCell = tmpSh.name & "!R1C1"
ActiveSheet.PivotTableWizard SourceType:=xlDatabase, SourceData:= _
    DataRn, TableDestination:=StartCell, TableName:="피벗 테이블1"
    
ActiveSheet.PivotTables("피벗 테이블1").AddFields ColumnFields:=ColVn
ActiveSheet.PivotTables("피벗 테이블1").PivotFields(DataVn).Orientation = _
        xlDataField
ActiveSheet.PivotTables("피벗 테이블1").PivotFields(tmpSh.Cells(1, 1).Value).Function = xlCount
ActiveSheet.PivotTables("피벗 테이블1").PivotSelect "", xlDataOnly
    
N = Selection.Columns.count

If count = 1 Then
ReDim cnt(t - 1, nmul - 1)
ReDim factor(t - 1, nmul - 1)
Else
ReDim Preserve cnt(t - 1, nmul - 1)
ReDim Preserve factor(t - 1, nmul - 1)
End If

For i = 1 To nmul
    cnt(count - 1, i - 1) = Selection.Cells(i)
Next i
    
Set temp = Selection.Offset(-1, 0)
For i = 1 To nmul
    factor(count - 1, i - 1) = temp.Cells(i)
Next i

Application.DisplayAlerts = False
tmpSh.Delete
Application.DisplayAlerts = True
    
End Sub
Sub PivotMakerforOneWay(DataRn, ColVn, DataVn, _
    cnt, ave, st, factor)
        
    Dim actSh, tmpSh As Worksheet
    Dim StartCell As String: Dim i, N As Long
    Dim temp As Range
    
    Set actSh = ActiveSheet
    Set tmpSh = Worksheets.Add
    actSh.Select
    StartCell = tmpSh.name & "!R1C1"
    ActiveSheet.PivotTableWizard SourceType:=xlDatabase, SourceData:= _
        DataRn, TableDestination:=StartCell, TableName:="피벗 테이블1"
    
    ActiveSheet.PivotTables("피벗 테이블1").AddFields ColumnFields:=ColVn
    ActiveSheet.PivotTables("피벗 테이블1").PivotFields(DataVn).Orientation = _
        xlDataField
    ActiveSheet.PivotTables("피벗 테이블1").PivotFields(tmpSh.Cells(1, 1).Value).Function = xlCount
    ActiveSheet.PivotTables("피벗 테이블1").PivotSelect "", xlDataOnly
    N = Selection.Columns.count
    ReDim cnt(1 To N): ReDim ave(1 To N): ReDim st(1 To N): ReDim factor(1 To N)
    For i = 1 To N
        cnt(i) = Selection.Cells(i)
    Next i
    ActiveSheet.PivotTables("피벗 테이블1").PivotFields(tmpSh.Cells(1, 1).Value).Function = xlAverage
    ActiveSheet.PivotTables("피벗 테이블1").PivotSelect "", xlDataOnly
    For i = 1 To N
        ave(i) = Selection.Cells(i)
    Next i
    
    Set temp = Selection.Offset(-1, 0)
    For i = 1 To N
        factor(i) = temp.Cells(i)
    Next i
    
    ActiveSheet.PivotTables("피벗 테이블1").PivotFields(tmpSh.Cells(1, 1).Value).Function = xlStDev
    ActiveSheet.PivotTables("피벗 테이블1").PivotSelect "", xlDataOnly
    For i = 1 To N
    If cnt(i) = 1 Then
        st(i) = "."
    Else
        st(i) = Selection.Cells(i)
    End If
    Next i
    
    Application.DisplayAlerts = False
    tmpSh.Delete
    Application.DisplayAlerts = True
    
End Sub

Sub test()
    
    Dim c(), m(), s() As Double
    
    PivotMaker ActiveSheet.name & "!A1:C35", "class1", "class2", "value", c, m, s
    
End Sub

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
    For J = 1 To UBound(sRn)
        Set sRn(J) = Range(TempSheet.Cells(mystart, 2), TempSheet.Cells(myend, 2))
        If J = UBound(sRn) Then Exit For
        mystart = myend + 1: myend = mystart + cnt(J + 1) - 1
    Next J
    
    Set TransClassVar = TempSheet
    
End Function

Sub Title1(contents As String)
    Dim Flag As Long
    Dim mySheet As Worksheet
    Dim tmpSign
    
    '''
    tmpSign = 0
    Set mySheet = Worksheets(RstSheet)
    If Left(mySheet.Range("a1"), 1) = "$" Then
        mySheet.Cells(1, 1) = Right(mySheet.Cells(1, 1).Value, Len(mySheet.Cells(1, 1).Value) - 3)
        tmpSign = 1
    End If
  
    
    Flag = mySheet.Cells(1, 1).Value
    yp = mySheet.Cells(Flag + 2, 1).Top
    
    On Error Resume Next
    
    Set Title = mySheet.Shapes.AddShape(msoShapeRectangle, 3.75, yp + 2.25, 400, 25#)
    With Title
        .Fill.ForeColor.SchemeColor = 57
        .Line.DashStyle = msoLineSolid
        .Line.Style = msoLineSingle
        .Line.Weight = 1
        .Line.Visible = msoTrue
      '  .Shadow.Type = msoShadow1
    End With
   
    With Title.TextFrame.Characters
        .Text = contents
        .Font.name = "굴림"
        .Font.FontStyle = "굵게"
        .Font.Size = 14
        .Font.ColorIndex = 2
    End With
    Title.TextFrame.HorizontalAlignment = xlCenter
    
    mySheet.Cells(1, 1) = Flag + 4
    

    
End Sub


'파일 버전 체크
Function ChkVersion(File_Name) As Boolean
    
    If Right(File_Name, 4) = ".xls" Or Right(File_Name, 4) = ".XLS" Then
        ChkVersion = False
    Else
        ChkVersion = True
    End If
End Function
Function openTempWorkSheet(tmpWS As Worksheet, _
    WSName As String, Optional StartNum As Integer = 1) As Boolean
    
    Dim Flag As Boolean: Dim ws As Worksheet
    
    For Each ws In Worksheets
        If ws.name = WSName Then
            Flag = True
            Set tmpWS = ws
            Exit For
        End If
    Next ws
    
    If Flag = False Then
        Set tmpWS = Worksheets.Add
        tmpWS.name = WSName
        tmpWS.Cells(1, 1) = StartNum
        tmpWS.Visible = xlSheetHidden
    End If
    
    openTempWorkSheet = True
        
End Function
Function Trans(list, xnames, cr, t, loco, c, outputsheet)
    Dim ttemp3 As Range
    Dim addr As Range
    Dim yp As Double
    
    'Set addr = outputsheet.Range("a1")
    'addr.Value = addr.Value
    'Set ttemp = outputsheet.Range("a" & addr.Value)
    
    Set addr = outputsheet.Range("a1") 'a1에 출력 될 행 번호가 저장됨
    Set ttemp3 = outputsheet.Range("a" & addr.Value)     '다음 출력 시작 위치
        If loco = 0 And t = 0 And c = 2 Then
            t = 1
            loco = loco + 1
            t = t * 220
            loco = loco * 4
        Else
            t = t * 320
            loco = loco * 6
        End If
    
    
    yp = ttemp3.Top
    Set Title = outputsheet.Shapes.AddShape(msoShapeRectangle, 270.25 + t, yp, 250, 20#)
    Title.Shadow.Type = msoShadow17
    With Title.Fill
         .ForeColor.SchemeColor = 1
         .Visible = msoTrue
         .Solid
    End With
    Title.TextFrame.Characters.Text = "변환표"
    With Title.TextFrame.Characters.Font
        .Size = 11
        .ColorIndex = xlAutomatic
    End With
    Title.TextFrame.HorizontalAlignment = xlCenter
    Set ttemp3 = ttemp3.Offset(2, 5 + loco)
    Set qq = ttemp3.Offset(cr, 0)
    With ttemp3.Resize(, 2).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With ttemp3.Resize(, 2).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With qq.Resize(, 2).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    
    ttemp3.Value = list
    ttemp3.Offset(0, 1).Value = "변환결과"
    Set ttemp3 = ttemp3.Offset(1, 0)
    For i = 0 To cr - 1
        ttemp3.Value = xnames(i, 1)
        ttemp3.Offset(0, 1).Value = xnames(i, 0)
        Set ttemp3 = ttemp3.Offset(1, 0)
    Next i
    'Set ttemp3 = ttemp3.Offset(3, -1)
    '''addr.Value = ttemp.Address
    'addr.Value = Right(ttemp.Address, Len(ttemp.Address) - 3)
End Function
