Attribute VB_Name = "PublicModule"

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

Function SelectedVariable(ParentDlgLbxValue, SelVar, _
         IsRowData As Boolean) As String
   
   Dim temp, m2, m3 As Long
   Dim TempSheet As Worksheet
   Dim tmp2, tmp As Range
   
   Set TempSheet = ActiveCell.Worksheet
   
   Dim Chk_Ver As Boolean
   Dim Cmp_R As Long        '
   Dim Cmp_C As Integer
   
  
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

Sub SelectMultiRange(ParentDlg, Rn, Vname, _
        Optional ColumnNum As Integer = 0)
   
   Dim cnt, temp, m2, m3, i, j As Long
   Dim TempSheet As Worksheet
   Dim tmp2 As Range
   
   cnt = ParentDlg.ListBox2.ListCount
   Set TempSheet = ActiveSheet

   Dim Chk_Ver As Boolean
   Dim Cmp_R As Long
   Dim Cmp_C As Integer     '
   
   '���� ������ ���� ��� ���� �񱳰� ����
   Chk_Ver = ChkVersion(ActiveWorkbook.name)
   If Chk_Ver = True Then
        Cmp_R = 1048576
        Cmp_C = 16384
    Else
        Cmp_R = 65536
        Cmp_C = 256
    End If
    
   'If ParentDlg.OptionButton1.Value = True Then
        temp = Cells.CurrentRegion.Columns.count
        For i = 1 To cnt
            For j = 1 To temp
               If StrComp(ParentDlg.ListBox2.list(i - 1, ColumnNum), TempSheet.Cells(1, j).value, 1) = 0 Then
                  Set tmp2 = TempSheet.Columns(j)
                  m2 = tmp2.Cells(1, 1).End(xlDown).row
                  If m2 <> Cmp_R Then
                     m3 = tmp2.Cells(m2, 1).End(xlDown).row
                     If m3 <> Cmp_R Then m2 = m3
                  End If
                  Set Rn(i) = tmp2.Range(Cells(2, 1), Cells(m2, 1))
                  Vname(i) = ParentDlg.ListBox2.list(i - 1, ColumnNum)
               End If
            Next j
        Next i
   'Else
   '     temp = Cells.CurrentRegion.Rows.count
   '     For i = 1 To cnt
   '         For j = 1 To temp
   '            If StrComp(ParentDlg.ListBox2.List(i - 1, ColumnNum), TempSheet.Cells(j, 1).Value, 1) = 0 Then
   '               Set tmp2 = TempSheet.Rows(j)
   '               m2 = tmp2.Cells(1, 1).End(xlToRight).Column
   '               If m2 <> cmp_c Then
   '                  m3 = tmp2.Cells(1, m2).End(xlToRight).Column
   '                  If m3 <> cmp_c Then m2 = m3
   '               End If
   '               Set Rn(i) = tmp2.Range(Cells(1, 2), Cells(1, m2))
   '               Vname(i) = ParentDlg.ListBox2.List(i - 1, ColumnNum)
   '            End If
   '         Next j
   '     Next i
   'End If
   
End Sub

'''SheetName�� ��Ʈ�� �����.
'''�̹� ������� ���� ���� �ٽ� ������ �ʴ´�.
'''A1���� �μ��� ��Ҹ� ����д�.
Sub OpenOutSheet(sheetname, Optional IsAddress As Boolean = False)
    
    Dim s, CurS As Worksheet
    
    Application.ScreenUpdating = False
    For Each s In ActiveWorkbook.Sheets
        If s.name = sheetname Then Exit Sub
    Next s
    Set CurS = ActiveSheet: Set s = Worksheets.Add
    With ActiveWindow
        .DisplayGridlines = False
'        .DisplayHeadings = False
    End With
    
    With Cells
         .Font.name = "����"
         .Font.Size = 9
         .HorizontalAlignment = xlRight
    End With

    s.name = sheetname: CurS.Activate
   ' With Worksheets(sheetname).Range("a1")
   '     .value = 2
        '''If IsAddress = True Then .Value = "A2"
  '      .Font.ColorIndex = 2
 '   End With
   ' Worksheets(sheetname).rows(1).Hidden = True
      

    's.Protect Password:="prophet", DrawingObjects:=False, contents:=True, Scenarios:=True
    Application.ScreenUpdating = True
    
End Sub

Sub ChartOutControl(PrintPosi, StartIndex As Boolean)               ''''"_�׷������_"

    Static s As Worksheet
    Static position As Range
    
    On Error GoTo sbcError
    If StartIndex = True Then
        OpenOutSheet "_���м����_"
        Set s = Worksheets("_���м����_")
        Set position = s.Range("a1")
        PrintPosi(0) = s.Cells(position.value + 6, 2).Left          '''
        PrintPosi(1) = s.Cells(position.value + 6, 2).Top
    Else
        's.Unprotect "prophet"
        '''�̶��� PrintPosi�� ��Ʈ�� ���α��̸� ��Ÿ���� ������.
        position.value = position.value + Int(PrintPosi / s.Range("a2").Height) + 4
        's.Protect Password:="prophet", DrawingObjects:=False, contents:=True, Scenarios:=True
    End If
    Exit Sub

sbcError:
    MsgBox "��½�Ʈ�� ���� �� �����ϴ�." & Chr(10) & _
    "[_���м����_]�̶�� �̸��� ��Ʈ�� ������ �ֽʽÿ�.", vbExclamation, title:="��� ����"

End Sub

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

Sub ShowfrmStat()
    On Error Resume Next
    DlgShow frmDisc, 7
End Sub
Sub Showfrmboxplot()
    On Error Resume Next
    DlgShow frmBoxplot, 3
End Sub
Sub Showfrmhistogram()
    On Error Resume Next
    DlgShow frmhistogram, 1
End Sub
Sub showfrmQQplot()
    On Error Resume Next
    DlgShow frmQQplot, 1
End Sub
Sub Showfrmstemandleaf()
    On Error Resume Next
    DlgShow frmstemandleaf, 2
End Sub
Sub ShowScatterPlot()
    On Error Resume Next
    DlgShow frmScatter, 4
End Sub
Sub ShowfrmTTest1()
    On Error Resume Next
    DlgShow frmTtest1, 5
End Sub
Sub ShowfrmTTest2()
    On Error Resume Next
    DlgShow1 frmTtest2
End Sub
Sub DlgShow1(ParentDlg As Object)
    
    Dim ErrSignforDataSheet As Integer
    
    ErrSignforDataSheet = InitializeDlg1(ParentDlg)
    
    Select Case ErrSignforDataSheet
    Case 0: ParentDlg.Show
    Case -1
        MsgBox "��Ʈ�� ��ȣ���¿� �ֽ��ϴ�." & Chr(10) & _
               "����Ÿ�� ���� �� �����ϴ�.", _
                vbExclamation, "HIST"
    Case 1
        MsgBox "��Ʈ�� ����Ÿ�� �ִ��� Ȯ���Ͻʽÿ�." & Chr(10) & _
               "1��1������ �����̸��� �Է��ؾ� �մϴ�.", _
               vbExclamation, "HIST"
    Case Else
    End Select

End Sub
Function InitializeDlg1(ParentDlg) As Integer
   
   Dim myRange As Range: Dim cnt As Long
   Dim myArray() As String
   
   On Error GoTo ErrorFlag
   
   Set myRange = ActiveSheet.Cells.CurrentRegion
   If myRange.count = 1 And myRange.Cells(1, 1) = "" Then
        InitializeDlg1 = 1: Exit Function
   End If
   Set myRange = ActiveSheet.Cells.CurrentRegion.rows(1)
   ParentDlg.ListBox4.Clear: ParentDlg.ListBox5.Clear
   cnt = myRange.Cells.count
   
   ReDim myArray(0 To cnt - 1)
   For i = 1 To cnt
     myArray(i - 1) = myRange.Cells(i)
   Next i
   ParentDlg.ListBox4.list() = myArray
   InitializeDlg1 = 0
   Exit Function
   
ErrorFlag:
   InitializeDlg1 = -1
   
End Function
Sub SetUpforPage2(ParentDlg, opt As Integer)

   Dim myRange As Range: Dim cnt As Long
   Dim myArray() As String
   
   Set myRange = ActiveSheet.Cells.CurrentRegion.rows(1)
   ParentDlg.ListBox1.Clear: ParentDlg.ListBox2.Clear: ParentDlg.ListBox3.Clear
   If opt = 2 Then
        ParentDlg.ListBox4.Clear
        ParentDlg.CommandButton19.Visible = False
        ParentDlg.CommandButton18.Visible = True
   End If
   'If opt <> 3 Then ParentDlg.CheckBox2.Value = True
   'ParentDlg.CheckBox3.Value = True
   'If opt = 3 Then ParentDlg.CheckBox4.Value = True
   'ParentDlg.CommandButton13.Visible = False
   'ParentDlg.CommandButton12.Visible = True
   'ParentDlg.CommandButton14.Visible = False
   'ParentDlg.CommandButton11.Visible = True
   cnt = myRange.Cells.count
   
   
   ReDim myArray(0 To cnt - 1)
   For i = 1 To cnt
     myArray(i - 1) = myRange.Cells(i)
   Next i
   ParentDlg.ListBox1.list() = myArray
   
End Sub



'''DlgOpt=1:������׷��� ���
'''DlgOpt=2:�ٱ��ٱ׸��� ���
'''DlgOpt=3:���ڱ׸��� ���
'''DlgOpt=4:�������� ���
'''DlgOpt=5,6:t-������ ���
'''DlgOpt=7:������跮 ���
Sub DlgShow(ParentDlg As Object, DlgOpt As Integer)
    
    Dim ErrSignforDataSheet As Integer
    
    ErrSignforDataSheet = InitializeDlg(ParentDlg, DlgOpt)
    
    Select Case ErrSignforDataSheet
    Case 0: ParentDlg.Show
    Case -1
        MsgBox "��Ʈ�� ��ȣ���¿� �ֽ��ϴ�." & Chr(10) & _
               "����Ÿ�� ���� �� �����ϴ�.", _
                vbExclamation, "HIST"
    Case 1
        MsgBox "��Ʈ�� ����Ÿ�� �ִ��� Ȯ���Ͻʽÿ�." & Chr(10) & _
               "1��1������ �����̸��� �Է��ؾ� �մϴ�.", _
               vbExclamation, "HIST"
    Case Else
    End Select

End Sub
Function InitializeDlg(ParentDlg, DlgOpt As Integer) As Integer
   
   Dim myRange As Range: Dim cnt As Long
   Dim myArray() As String
   
   On Error GoTo ErrorFlag
   
   Set myRange = ActiveSheet.Cells.CurrentRegion
   If myRange.count = 1 And myRange.Cells(1, 1) = "" Then
        InitializeDlg = 1: Exit Function
   End If
   Set myRange = ActiveSheet.Cells.CurrentRegion.rows(1)
   ParentDlg.ListBox1.Clear
   If DlgOpt = 1 Then
        ParentDlg.OptionButton1 = True
        ParentDlg.Image1.Picture = LoadPicture("")
   ElseIf DlgOpt = 2 Then
        ParentDlg.ListBox1.Clear: ParentDlg.Previewtxt.text = ""
        ParentDlg.OptionButton1 = True
   ElseIf DlgOpt = 4 Then
        ParentDlg.ListBox2.Clear: ParentDlg.ListBox3.Clear
        ParentDlg.CheckBox1.value = False
        ParentDlg.CheckBox2.value = True
        ParentDlg.CommandButton3.Visible = False
        ParentDlg.CommandButton2.Visible = True
        ParentDlg.CommandButton7.Visible = False
        ParentDlg.CommandButton1.Visible = True
   ElseIf DlgOpt = 5 Then
        ParentDlg.ListBox2.Clear
   ElseIf DlgOpt = 3 Or DlgOpt = 7 Then
        ParentDlg.ListBox2.Clear
        ParentDlg.OptionButton1 = True
   ElseIf DlgOpt = 6 Then
        ParentDlg.CommandButton3.Visible = False
        ParentDlg.CommandButton2.Visible = True
        ParentDlg.CommandButton7.Visible = False
        ParentDlg.CommandButton1.Visible = True
        ParentDlg.ListBox2.Clear
        ParentDlg.ListBox3.Clear
   End If
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

''�ӽý�Ʈ �����
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

'''����ǥ������ �̿��Ͽ� ���α׷� ���࿡ ���� �����ֱ�
Sub SettingStatusBar(SettingChoice As Boolean, _
        Optional NewString As String = "")

    Static oldStatusBar As String
    
    If SettingChoice = True Then
        oldStatusBar = Application.DisplayStatusBar
        Application.DisplayStatusBar = True
        Application.StatusBar = NewString
    Else
        Application.StatusBar = False
        Application.DisplayStatusBar = oldStatusBar
    End If
    
End Sub

Sub MoveBtwnListBox(ParentD, FromLNum, ToLNum)

    Dim i As Integer
    i = 0
    Do While i <= ParentD.Controls(FromLNum).ListCount - 1
        If ParentD.Controls(FromLNum).Selected(i) = True Then
           ParentD.Controls(ToLNum).AddItem ParentD.Controls(FromLNum).list(i)
           ParentD.Controls(FromLNum).RemoveItem i
           i = i - 1
        End If
        i = i + 1
    Loop

End Sub

Sub OptBtn12Click(ParentD, IsColumn As Boolean)
   
   Dim myRange As Range
   Dim myArray()
   
   ParentD.ListBox1.Clear: ParentD.ListBox2.Clear
   If IsColumn = True Then
      Set myRange = Cells.CurrentRegion.rows(1)
   Else
      Set myRange = Cells.CurrentRegion.Columns(1)
   End If
   cnt = myRange.Cells.count
   ReDim myArray(cnt - 1)
   For i = 1 To cnt
     myArray(i - 1) = myRange.Cells(i)
   Next i
   ParentD.ListBox1.list() = myArray
   
End Sub

'''���� ������ �׸��� ����.
Sub DesignOutPutCell(TargetCell, Direction, myLineStyle, _
    myWeight, myColorIndex)
    
    With TargetCell.Borders(Direction)
        .LineStyle = myLineStyle
        .Weight = myWeight
        .ColorIndex = myColorIndex
    End With

End Sub

'''�������� ���� �Լ�
'''���� �ڸ�����ŭ�� ��Ʈ���� ����(������ ��츸)
Function CStrNumPoint(DataWid, DataCount) As String
    
    Dim i As Integer: Dim LogScale As Double
    Dim temp As String
    
    i = 0: temp = "0."
    LogScale = Application.Power(10, _
             Int(Application.Log10(DataWid / DataCount)))
    If LogScale >= 1 Then
        CStrNumPoint = "0"
    Else
        Do
            temp = temp & "0": i = i + 1
            If LogScale = 10 ^ (-i) Then Exit Do
        Loop While (1)
        CStrNumPoint = CStr(temp)
    End If

End Function


Function ChkVersion(File_Name) As Boolean
    
    If Right(File_Name, 4) = ".xls" Or Right(File_Name, 4) = ".XLS" Then
        ChkVersion = False
    Else
        ChkVersion = True
    End If
End Function
