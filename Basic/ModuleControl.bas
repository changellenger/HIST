Attribute VB_Name = "ModuleControl"
Public sheetRowNum As Long, sheetColNum As Integer  'excel version�� ���� sheet ��� �� ����
Public DataSheet As String, rstSheet As String      'sheet�̸� �� ��
Public xlist() As String                            '������ �� ��
Public N As Long, m As Long, p As Long              '���õ� ������ ����Ÿ ����, ��ü��������, ���õ� ��������
                                                    '��� frmSurvey ���� �ѹ��� �����ǰ� �ٸ� �������� �ٲ��� �ʴ´�.
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
                                                    

                                                    
Sub AlphaShow()

    Dim ErrSignforDataSheet As Integer
    
    ErrSignforDataSheet = InitializeDlg(frmAlpha)
                                    
    Select Case ErrSignforDataSheet
    Case 0: frmAlpha.Show
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


Function InitializeDlg(ParentDlg) As Integer
   
   Dim myRange As Range: Dim cnt As Long
   Dim myArray() As String
   
   On Error GoTo ErrorFlag
   
   Set myRange = ActiveSheet.Cells.CurrentRegion
   If myRange.count = 1 And myRange.Cells(1, 1) = "" Then
        InitializeDlg = 1: Exit Function
   End If
   Set myRange = ActiveSheet.Cells.CurrentRegion.Rows(1)
   ParentDlg.Listbox1.Clear: ParentDlg.ListBox2.Clear
   cnt = myRange.Cells.count
   
   ReDim myArray(0 To cnt - 1)
   For i = 1 To cnt
     myArray(i - 1) = myRange.Cells(i)
   Next i
   ParentDlg.Listbox1.list() = myArray
   InitializeDlg = 0
   Exit Function
   
ErrorFlag:
   InitializeDlg = -1
   
End Function


Function FindDataCount(xlist)

    Set tmpSheet = Worksheets(DataSheet)
    
    For i = 0 To p - 1
        For J = 1 To m
            If StrComp(xlist(i), tmpSheet.Cells(1, J).Value, 1) = 0 Then
                If i = 0 Then
                    N = tmpSheet.Columns(J).Cells(1, 1).End(xlDown).row
                Else
                    N = WorksheetFunction.max(N, N = tmpSheet.Columns(J).Cells(1, 1).End(xlDown).row)
                End If
            End If
        Next J
    Next i
    
    FindDataCount = N
    
End Function
    
    
Function FindColDataCount(ListVar) As Long
   
    Dim J As Integer, M2 As Long
    Dim tmpSheet As Worksheet
    Dim tmpRange As Range
    
    Set tmpSheet = Worksheets(DataSheet)
    
    For J = 1 To m
        If StrComp(ListVar, tmpSheet.Cells(1, J).Value, 1) = 0 Then
            M2 = tmpSheet.Columns(J).Cells(1, 1).End(xlDown).row
            FindColDataCount = tmpSheet.Range(Cells(2, 1), Cells(M2, 1)).count
        End If
    Next J
    
End Function


Function FindingRangeError(ListVar) As Boolean
    
    Dim tmpSheet As Worksheet
    Dim tmpCount As Integer
    Dim tmpCol As Range, tmp1 As Range, tmp2 As Range, tmp3 As Range
    
    Set tmpSheet = Worksheets(DataSheet)

    For J = 1 To m
       If StrComp(ListVar, tmpSheet.Cells(1, J).Value, 1) = 0 Then
          Set tmpCol = tmpSheet.Columns(J)
          M2 = tmpCol.Cells(1, 1).End(xlDown).row
          If M2 <> sheetRowNum Then
             m3 = tmpCol.Cells(M2, 1).End(xlDown).row
             If m3 <> sheetRowNum Then M2 = m3
          End If
          Set tmpCol = tmpCol.Range(Cells(2, 1), Cells(M2, 1))
       End If
    Next J
    
    On Error Resume Next
    
    If Application.CountBlank(tmpCol) >= 1 Then
        FindingRangeError = True
        Exit Function
    End If
    
    Set tmp1 = tmpCol.SpecialCells(xlCellTypeConstants, 22)
    Set tmp2 = tmpCol.SpecialCells(xlCellTypeFormulas, 22)
    Set tmp3 = tmpCol.SpecialCells(xlCellTypeBlanks)
    
    If tmpCol.count = 1 And IsNumeric(tmpCol.Cells(1, 1)) = True Then
        FindingRangeError = False
    Else
        If tmp1 Is Nothing And tmp2 Is Nothing And tmp3 Is Nothing Then
            FindingRangeError = False
        Else: FindingRangeError = True
        End If
    End If
        
End Function

   
Sub SettingStatusBar(SettingChoice As Boolean, Optional NewString As String = "")

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


Sub SurveyAnal()
        
    Dim dataX(), rst1(), rst2()
    Dim row As Integer, col As Integer
    Dim index As Integer, Flag As Long
    Dim mySheet As Worksheet, tmpSheet As Worksheet, WS As Worksheet
    Dim pt As Range, tmp_pt As Range
    Dim alpha1, alpha2
    Dim alpha11(), alpha22(), corr11(), corr22()
    Dim tmpCheck As Integer
    
    dataX = arrayX(xlist, p)
    

    ModulePrint.Title1 "�ŷڵ� �м����"
    
    
    
    tmpCheck = 0
    
    For Each WS In Worksheets
        If WS.Name = "_#TempSurvey#_" Then
        tmpCheck = 1
        End If
    Next WS
    
    If tmpCheck = 0 Then
        Worksheets.Add.Name = "_#TempSurvey#_"
    End If
    
    Worksheets("_#TempSurvey#_").Visible = False
   
    
    Set tmpSheet = Worksheets("_#TempSurvey#_")
    tmpSheet.Cells(1, 1).Value = 2
    Flag = tmpSheet.Cells(1, 1).Value
            
    'If p > 1 Then
    '    ModulePrint.TABLE n + 3, 1 + p + 1 + p
    'Else
    '    ModulePrint.TABLE n + 3, 1 + p + 1
    'End If
    
    Set tmp_pt = tmpSheet.Cells(Flag, 2)
    tmp_pt(1, 1) = "��ȣ"
    tmp_pt(1 + N + 1, 1) = "�л�"
    tmp_pt(1 + N + 2, 1) = "���"
    With tmp_pt(1 + N + 1, 1).Resize(, 1 + p + 1).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    For row = 1 To N
        tmp_pt(1 + row, 1) = row
    Next row
    Set tmp_pt = tmpSheet.Cells(Flag, 3)
    
    Range(tmp_pt(1, 1), tmp_pt(N + 1, p)) = dataX
    
    tmp_pt(1, p + 1) = "������"
    
    For row = 1 To N
        tmp_pt(1 + row, p + 1).FormulaR1C1 = WorksheetFunction.sum(Range(tmp_pt(row + 1, 1), tmp_pt(row + 1, p)))
    Next row
    For col = 1 To p + 1
        tmp_pt(1 + N + 1, col) = WorksheetFunction.var(Range(tmp_pt(1 + 1, col), tmp_pt(1 + N, col)))
        tmp_pt(1 + N + 2, col) = WorksheetFunction.Average(Range(tmp_pt(1 + 1, col), tmp_pt(1 + N, col)))
    Next col
    
    If p > 1 Then
        tmp_pt(1, 1) = Chr(10) & Chr(10) & "��ȣ"
        tmp_pt(1, p + 1) = Chr(10) & Chr(10) & "������"
        For col = 1 To p
            tmp_pt(1, col) = Chr(10) & Chr(10) & xlist(col - 1)
            tmp_pt(1, p + 1 + col) = xlist(col - 1) & Chr(10) & "����" & Chr(10) & "������"
        Next col
        For row = 1 To N
            tmp_pt(1 + row, p + 2).FormulaR1C1 = WorksheetFunction.sum(Range(tmp_pt(row + 1, 2), tmp_pt(row + 1, p)))
            tmp_pt(1 + row, p + 1 + p).FormulaR1C1 = WorksheetFunction.sum(Range(tmp_pt(row + 1, 1), tmp_pt(row + 1, p - 1)))
        Next row
        For col = 2 To p - 1
            For row = 1 To N
                tmp_pt(1 + row, p + 1 + col).FormulaR1C1 = WorksheetFunction.sum(Range(tmp_pt(row + 1, 1), tmp_pt(row + 1, col - 1)), Range(tmp_pt(row + 1, col + 1), tmp_pt(row + 1, p)))
            Next row
        Next col
        For col = 1 To p
            tmp_pt(1 + N + 1, p + 1 + col) = WorksheetFunction.var(Range(tmp_pt(1 + 1, p + 1 + col), tmp_pt(1 + N, p + 1 + col)))
            tmp_pt(1 + N + 2, p + 1 + col) = WorksheetFunction.Average(Range(tmp_pt(1 + 1, p + 1 + col), tmp_pt(1 + N, p + 1 + col)))
        Next col
        With tmp_pt(1 + N + 1, 1).Resize(, p + 1 + p).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    End If
    
    tmp_pt.Cells(1 + N + 1, 1).Resize(2, p + 1 + p).NumberFormatLocal = "0.000_ "
    
    tmpSheet.Cells(1, 1) = Flag + N + 3 + 2
    
    
    Flag = tmpSheet.Cells(1, 1).Value
        
    'If p > 1 Then
    '    ModulePrint.TABLE 4, 1 + p + 1 + p
    'Else
    '    ModulePrint.TABLE 4, 1 + p + 1
    'End If
    
    Set tmp_pt = tmpSheet.Cells(Flag, 2)
    tmp_pt(2, 1) = "���׺� �л���"
    tmp_pt(3, 1) = "ũ�й� ��"
    tmp_pt(4, 1) = "�κ��հ� ������"
    
    tmp_pt(1, 1 + p + 1) = "������"
    tmp_pt(2, 1 + p + 1).FormulaR1C1 = WorksheetFunction.sum(Range(tmp_pt(-3, 2), tmp_pt(-3, p + 1)))
    If p > 1 Then
        tmp_pt(3, 1 + p + 1) = p / (p - 1) * (1 - tmp_pt(2, 1 + p + 1) / tmp_pt(-3, 1 + p + 1))
        alpha1 = p / (p - 1) * (1 - tmp_pt(2, 1 + p + 1) / tmp_pt(-3, 1 + p + 1))
    End If
    
    If p > 1 Then
        tmp_pt(1, 1 + p + 1) = Chr(10) & "������"
        For col = 1 To p
            tmp_pt(1, 1 + p + 1 + col) = xlist(col - 1) & Chr(10) & "����"
            tmp_pt(1, 1 + col) = Chr(10) & xlist(col - 1)
        Next col
        tmp_pt(2, 1 + p + 1 + 1).FormulaR1C1 = WorksheetFunction.sum(Range(tmp_pt(-3, 3), tmp_pt(-3, 1 + p)))
        tmp_pt(2, 1 + p + 1 + p).FormulaR1C1 = WorksheetFunction.sum(Range(tmp_pt(-3, 2), tmp_pt(-3, p)))
        For col = 2 To p - 1
            tmp_pt(2, 1 + p + 1 + col).FormulaR1C1 = WorksheetFunction.sum(Range(tmp_pt(-3, 2), tmp_pt(-3, col)), Range(tmp_pt(-3, col + 2), tmp_pt(-3, 1 + p)))
        Next col
        If p > 2 Then
        ReDim alpha11(p)
        For col = 1 To p
            tmp_pt(3, 1 + p + 1 + col) = (p - 1) / (p - 2) * (1 - tmp_pt(2, 1 + p + 1 + col) / tmp_pt(-3, 1 + p + 1 + col))
            alpha11(col) = (p - 1) / (p - 2) * (1 - tmp_pt(2, 1 + p + 1 + col) / tmp_pt(-3, 1 + p + 1 + col))
        Next col
        End If
    End If
    
    ReDim corr11(p)
    If p > 1 Then
        For col = 1 To p
            tmp_pt(4, 1 + col) = WorksheetFunction.Correl(Range(tmp_pt(-4, 1 + col), tmp_pt(-4 - N + 1, 1 + col)), Range(tmp_pt(-4, 1 + col + 1 + p), tmp_pt(-4 - N + 1, 1 + col + 1 + p)))
            corr11(col) = WorksheetFunction.Correl(Range(tmp_pt(-4, 1 + col), tmp_pt(-4 - N + 1, 1 + col)), Range(tmp_pt(-4, 1 + col + 1 + p), tmp_pt(-4 - N + 1, 1 + col + 1 + p)))
        Next col
    Else
        For col = 1 To p
            tmp_pt(4, 1 + col) = WorksheetFunction.Correl(Range(tmp_pt(-4, 1 + col), tmp_pt(-4 - N + 1, 1 + col)), Range(tmp_pt(-4, 1 + col), tmp_pt(-4 - N + 1, 1 + col)))
            corr11(col) = WorksheetFunction.Correl(Range(tmp_pt(-4, 1 + col), tmp_pt(-4 - N + 1, 1 + col)), Range(tmp_pt(-4, 1 + col), tmp_pt(-4 - N + 1, 1 + col)))
        Next col
    End If
    
    tmp_pt.Cells(2, 2).Resize(3, p + 1 + p).NumberFormatLocal = "0.000_ "
    
    tmpSheet.Cells(1, 1) = Flag + 4 + 4
    
    
    
    
    
    
    Flag = tmpSheet.Cells(1, 1).Value
    
    'If p > 1 Then
    '    ModulePrint.TABLE n + 3, 1 + p + 1 + p
    'Else
    '    ModulePrint.TABLE n + 3, 1 + p + 1
    'End If
    
    Set tmp_pt = tmpSheet.Cells(Flag, 2)
    tmp_pt(1, 1) = "��ȣ"
    tmp_pt(1 + N + 1, 1) = "�л�"
    tmp_pt(1 + N + 2, 1) = "���"
    With tmp_pt(1 + N + 1, 1).Resize(, 1 + p + 1).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    For row = 1 To N
        tmp_pt(1 + row, 1) = row
    Next row
    Set tmp_pt = tmpSheet.Cells(Flag, 3)
    
    
    For col = 1 To p
        For row = 1 To N
            tmp_pt(1 + row, col) = WorksheetFunction.Standardize(tmp_pt(-11 - N + row - 1, col), WorksheetFunction.Average(Range(tmp_pt(-11 - N, col), tmp_pt(-11 - 1, col))), WorksheetFunction.StDev(Range(tmp_pt(-11 - N, col), tmp_pt(-11 - 1, col))))
        Next row
    Next col
    
    
    tmp_pt(1, p + 1) = "������"
    
    For row = 1 To N
        tmp_pt(1 + row, p + 1).FormulaR1C1 = WorksheetFunction.sum(Range(tmp_pt(row + 1, 1), tmp_pt(row + 1, p)))
    Next row
    For col = 1 To p + 1
        tmp_pt(1 + N + 1, col) = WorksheetFunction.var(Range(tmp_pt(1 + 1, col), tmp_pt(1 + N, col)))
        tmp_pt(1 + N + 2, col) = WorksheetFunction.Average(Range(tmp_pt(1 + 1, col), tmp_pt(1 + N, col)))
    Next col
    
    If p > 1 Then
        tmp_pt(1, 1) = Chr(10) & Chr(10) & "��ȣ"
        tmp_pt(1, p + 1) = Chr(10) & Chr(10) & "������"
        For col = 1 To p
            tmp_pt(1, col) = Chr(10) & Chr(10) & xlist(col - 1)
            tmp_pt(1, p + 1 + col) = xlist(col - 1) & Chr(10) & "����" & Chr(10) & "������"
        Next col
        For row = 1 To N
            tmp_pt(1 + row, p + 2).FormulaR1C1 = WorksheetFunction.sum(Range(tmp_pt(row + 1, 2), tmp_pt(row + 1, p)))
            tmp_pt(1 + row, p + 1 + p).FormulaR1C1 = WorksheetFunction.sum(Range(tmp_pt(row + 1, 1), tmp_pt(row + 1, p - 1)))
        Next row
        For col = 2 To p - 1
            For row = 1 To N
                tmp_pt(1 + row, p + 1 + col).FormulaR1C1 = WorksheetFunction.sum(Range(tmp_pt(row + 1, 1), tmp_pt(row + 1, col - 1)), Range(tmp_pt(row + 1, col + 1), tmp_pt(row + 1, p)))
            Next row
        Next col
        For col = 1 To p
            tmp_pt(1 + N + 1, p + 1 + col) = WorksheetFunction.var(Range(tmp_pt(1 + 1, p + 1 + col), tmp_pt(1 + N, p + 1 + col)))
            tmp_pt(1 + N + 2, p + 1 + col) = WorksheetFunction.Average(Range(tmp_pt(1 + 1, p + 1 + col), tmp_pt(1 + N, p + 1 + col)))
        Next col
        With tmp_pt(1 + N + 1, 1).Resize(, p + 1 + p).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    End If
    
    tmp_pt.Cells(1, 1).Resize(N + 3, p + 1 + p).NumberFormatLocal = "0.000_ "
    
    tmpSheet.Cells(1, 1) = Flag + N + 3 + 2
    
    
    Flag = tmpSheet.Cells(1, 1).Value
        
    'If p > 1 Then
    '    ModulePrint.TABLE 4, 1 + p + 1 + p
    'Else
    '    ModulePrint.TABLE 4, 1 + p + 1
    'End If
    
    Set tmp_pt = tmpSheet.Cells(Flag, 2)
    tmp_pt(2, 1) = "���׺� �л���"
    tmp_pt(3, 1) = "ũ�й� ��"
    tmp_pt(4, 1) = "�κ��հ� ������"
    
    tmp_pt(1, 1 + p + 1) = "������"
    tmp_pt(2, 1 + p + 1).FormulaR1C1 = WorksheetFunction.sum(Range(tmp_pt(-3, 2), tmp_pt(-3, p + 1)))
    If p > 1 Then
        tmp_pt(3, 1 + p + 1) = p / (p - 1) * (1 - tmp_pt(2, 1 + p + 1) / tmp_pt(-3, 1 + p + 1))
        alpha2 = p / (p - 1) * (1 - tmp_pt(2, 1 + p + 1) / tmp_pt(-3, 1 + p + 1))
    End If
    
    If p > 1 Then
        tmp_pt(1, 1 + p + 1) = Chr(10) & "������"
        For col = 1 To p
            tmp_pt(1, 1 + p + 1 + col) = xlist(col - 1) & Chr(10) & "����"
            tmp_pt(1, 1 + col) = Chr(10) & xlist(col - 1)
        Next col
        tmp_pt(2, 1 + p + 1 + 1).FormulaR1C1 = WorksheetFunction.sum(Range(tmp_pt(-3, 3), tmp_pt(-3, 1 + p)))
        tmp_pt(2, 1 + p + 1 + p).FormulaR1C1 = WorksheetFunction.sum(Range(tmp_pt(-3, 2), tmp_pt(-3, p)))
        For col = 2 To p - 1
            tmp_pt(2, 1 + p + 1 + col).FormulaR1C1 = WorksheetFunction.sum(Range(tmp_pt(-3, 2), tmp_pt(-3, col)), Range(tmp_pt(-3, col + 2), tmp_pt(-3, 1 + p)))
        Next col
        If p > 2 Then
        ReDim alpha22(p)
        For col = 1 To p
            tmp_pt(3, 1 + p + 1 + col) = (p - 1) / (p - 2) * (1 - tmp_pt(2, 1 + p + 1 + col) / tmp_pt(-3, 1 + p + 1 + col))
            alpha22(col) = (p - 1) / (p - 2) * (1 - tmp_pt(2, 1 + p + 1 + col) / tmp_pt(-3, 1 + p + 1 + col))
        Next col
        End If
    End If
    
    ReDim corr22(p)
    If p > 1 Then
        For col = 1 To p
            tmp_pt(4, 1 + col) = WorksheetFunction.Correl(Range(tmp_pt(-4, 1 + col), tmp_pt(-4 - N + 1, 1 + col)), Range(tmp_pt(-4, 1 + col + 1 + p), tmp_pt(-4 - N + 1, 1 + col + 1 + p)))
            corr22(col) = WorksheetFunction.Correl(Range(tmp_pt(-4, 1 + col), tmp_pt(-4 - N + 1, 1 + col)), Range(tmp_pt(-4, 1 + col + 1 + p), tmp_pt(-4 - N + 1, 1 + col + 1 + p)))
        Next col
    Else
        For col = 1 To p
            tmp_pt(4, 1 + col) = WorksheetFunction.Correl(Range(tmp_pt(-4, 1 + col), tmp_pt(-4 - N + 1, 1 + col)), Range(tmp_pt(-4, 1 + col), tmp_pt(-4 - N + 1, 1 + col)))
            corr22(col) = WorksheetFunction.Correl(Range(tmp_pt(-4, 1 + col), tmp_pt(-4 - N + 1, 1 + col)), Range(tmp_pt(-4, 1 + col), tmp_pt(-4 - N + 1, 1 + col)))
        Next col
    End If
    tmp_pt.Cells(2, 2).Resize(3, p + 1 + p).NumberFormatLocal = "0.000_ "
    
    tmpSheet.Cells(1, 1) = Flag + 4 + 4
    
    
    
    
    
    Set mySheet = Worksheets(rstSheet)
    Flag = mySheet.Cells(1, 1).Value + 2
    Set pt = mySheet.Cells(Flag, 2)
    
    Set tmp_pt = tmpSheet.Cells(tmpSheet.Cells(1, 1).Value, 2)
    
    pt(1, 1) = "���׺� ������ ���"
    pt(1, 1).HorizontalAlignment = xlLeft
    With pt.Resize(1, 2).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    
    
    
    For col = 1 To p
        pt(3 + col, 1) = xlist(col - 1)
        pt(3, 1 + col) = xlist(col - 1)
    Next col
        
    For col = 1 To p
        For row = 1 To p
            pt(3 + col, row + 1) = WorksheetFunction.Correl(Range(tmp_pt(-24 - N - N, 1 + col), tmp_pt(-24 - N - 1, 1 + col)), Range(tmp_pt(-24 - N - N, 1 + row), tmp_pt(-24 - N - 1, 1 + row)))
        Next row
    Next col
    pt.Cells(4, 2).Resize(p, p).NumberFormatLocal = "0.000_ "
    
    With pt(2, 1).Resize(1, 1 + p).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With pt(3, 1).Resize(1, 1 + p).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With pt(3 + p, 1).Resize(1, 1 + p).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    With pt(3, 1).Resize(1 + p, 1).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With pt(3, 1 + p).Resize(1 + p, 1).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With pt(3, 1).Resize(1 + p, 1).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With pt(3, 1 + p).Resize(1 + p, 1).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    mySheet.Cells(1, 1) = Flag + 1 + p + 4
    
    
    
    
    
    
    
    
    
    
    Flag = mySheet.Cells(1, 1).Value
    Set pt = mySheet.Cells(Flag, 2)
    
    pt(1, 1) = "Cronbach ���İ��"
    pt(1, 1).HorizontalAlignment = xlLeft
    With pt.Resize(1, 2).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    
    
    
    With pt(2, 1).Resize(1, 2).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With pt(3, 1).Resize(1, 2).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With pt(5, 1).Resize(1, 2).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    pt(3, 1) = "����"
    pt(3, 2) = "���İ��"
    pt(4, 1) = "�� ������"
    pt(5, 1) = "ǥ��ȭ ������"
    pt(4, 2) = alpha1
    pt(5, 2) = alpha2
    pt.Cells(4, 2).Resize(2, 1).NumberFormatLocal = "0.000_ "
    mySheet.Cells(1, 1) = Flag + 3 + 4
    
    
    
    Flag = mySheet.Cells(1, 1).Value
    Set pt = mySheet.Cells(Flag, 2)
    
    pt(1, 1) = "������ ������ ���� Cronbach ���İ��"
    pt(1, 1).HorizontalAlignment = xlLeft
    With pt.Resize(1, 3).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With

    
    
    With pt(2, 1).Resize(1, 5).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With pt(3, 2).Resize(1, 4).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With pt(5, 1).Resize(1, 5).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With pt(5 + p, 1).Resize(1, 5).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With



    With pt(3, 1).Resize(p + 3, 1).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With pt(3, 1).Resize(p + 3, 1).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With pt(4, 2).Resize(p + 2, 1).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With pt(3, 3).Resize(p + 3, 1).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With pt(4, 4).Resize(p + 2, 1).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With pt(3, 5).Resize(p + 3, 1).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With


    pt(4, 1) = "������"
    pt(5, 1) = "����"
    pt(3, 2) = "��"
    pt(3, 3) = " ������"
    pt(3, 3).HorizontalAlignment = xlLeft
    pt(3, 4) = "ǥ��ȭ"
    pt(3, 5) = " ������"
    pt(3, 5).HorizontalAlignment = xlLeft
    pt(4, 2) = "�κ��հ�"
    pt(5, 2) = "������"
    pt(5, 3) = "���İ��"
    pt(4, 4) = "�κ��հ�"
    pt(5, 4) = "������"
    pt(5, 5) = "���İ��"


If p = 2 Then
    For row = 1 To p
        pt(6 + row - 1, 1) = xlist(row - 1)
        pt(6 + row - 1, 2) = corr11(row)
        pt(6 + row - 1, 4) = corr22(row)
    Next row

    pt.Cells(6, 2).Resize(p, 4).NumberFormatLocal = "0.000_ "
End If

If p > 2 Then
    For row = 1 To p
        pt(6 + row - 1, 1) = xlist(row - 1)
        pt(6 + row - 1, 2) = corr11(row)
        pt(6 + row - 1, 3) = alpha11(row)
        pt(6 + row - 1, 4) = corr22(row)
        pt(6 + row - 1, 5) = alpha22(row)
    Next row

    pt.Cells(6, 2).Resize(p, 4).NumberFormatLocal = "0.000_ "
End If

    mySheet.Cells(1, 1) = Flag + p + 4 + 4



End Sub


''dataX(0,0)~dataX(n,p-1)�� ������ + �ڷ�
''varList�� ��޵� ���������� ����Ÿ���� ������� �Բ� ������ �迭�� ��ȯ

Function arrayX(varList, pp) As Variant

    Dim dataRange As Range
    Dim i As Long, J As Integer, k As Integer
    Dim X()
    
    Set dataRange = Worksheets(DataSheet).Cells.CurrentRegion
    ReDim X(N, pp - 1)
   
    For J = 0 To pp - 1
         For k = 0 To m
            If varList(J) = dataRange.Cells(1, k + 1) Then
                For i = 0 To N
                    X(i, J) = dataRange.Cells(i + 1, k + 1).Value
                Next i
            End If
        Next k
    Next J
    
    arrayX = X
    
End Function


Sub FreqShow()

    Dim ErrSignforDataSheet As Integer
    'On Error Resume Next
    ErrSignforDataSheet = InitializeDlg(frmFrequency)
    
    Select Case ErrSignforDataSheet
    Case 0: frmFrequency.Show
    Case -1
        MsgBox "��Ʈ�� ��ȣ���¿� �ֽ��ϴ�." & Chr(10) & _
               "����Ÿ�� ���� �� �����ϴ�.", _
                vbExclamation, "DB �м�"
    Case 1
        MsgBox "��Ʈ�� ����Ÿ�� �ִ��� Ȯ���Ͻʽÿ�." & Chr(10) & _
               "1��1������ �����̸��� �Է��ؾ� �մϴ�.", _
               vbExclamation, "DB �м�"
    Case Else
    End Select

End Sub


Sub FreqAnalysis(myRange, list, tit)
    
    Dim temp As Worksheet
    Dim MyNewRange As Range
    Dim i As Long, M2 As Long, num As Integer
    Dim yp As Double
    Dim k As Integer
    Dim ttemp, addr As Range
    Dim qq As Range
    Set temp = Worksheets(rstSheet)
    Set addr = temp.Range("a1")
    Set ttemp = temp.Range("a" & addr.Value)
    
    Dim cou As Integer
    Dim test()
    Dim sum, aa As Integer
    
    aa = temp.Cells(1, 1).Value
    'Set ttemp = ttemp.Offset(i, 0)
    If NumOfTable = "" Then NumOfTable = 0
        
  
    If tit = 1 Then
    ModulePrint.Title1 "�󵵺м����"
    Else
    ModulePrint.Title1 "���ڿ�������"
    End If

   
   
    
    For k = 1 To p
      
        ModulePrint.Title3 xlist(k - 1)
       
        'vartype(�ڷ�) �Է¿��� ������..
        For J = 1 To UBound(list, 1)
            If xlist(k - 1) = list(J) Then
                cou = UBound(ratio(myRange, test, J - 1), 1)
                J = UBound(list, 1)
'        Call cswap(MyRange, k, UBound(MyRange, 2))
                i = temp.Cells(1, 1).Value
                
         
                Set ttemp = ttemp.Offset(i - aa, 2)
         
                
                With ttemp.Resize(, 3).Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
       
                With ttemp.Resize(, 3).Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                    .ColorIndex = xlAutomatic
                End With
                
                With ttemp.Resize(cou + 2, 1).Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                    .ColorIndex = xlAutomatic
                End With
                                
            End If
        Next J
        
   
        ttemp.Value = "������ ����"
        ttemp.Offset(0, 1) = "�󵵼�"
        ttemp.Offset(0, 2) = "����"
        sum = 0: sum1 = 0
        
        For w = 1 To cou
            ttemp.Offset(w, 0) = test(w, 1)
            ttemp.Offset(w, 1) = test(w, 3)
            sum = test(w, 3) + sum
            ttemp.Offset(w, 2) = test(w, 2) * 100 & "%"
        Next w
        
        ttemp.Offset(cou + 1, 0) = "���հ�"
        ttemp.Offset(cou + 1, 1) = sum
        ttemp.Offset(cou + 1, 2) = "100%"
        
        temp.Cells(1, 1) = i + 10
        
       Set ttemp = ttemp.Offset(cou + 1, 0)
       
       With ttemp.Resize(, 3).Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
       End With
       Set ttemp = ttemp.Offset(-i + aa - cou - 1, -2)
       temp.Cells(1, 1) = i + cou + 1
       
    Next k
   ' temp.Cells(1, 1) = i + 8
    
    
        
End Sub



Function InitializeDlg1(ParentDlg) As Integer
   
   Dim myRange As Range
   Dim cnt As Long
   Dim myArray() As String
   Dim i As Integer
   
   On Error GoTo ErrorFlag
   
   Set myRange = ActiveSheet.Cells.CurrentRegion
   If myRange.count = 1 Or myRange.Cells(1, 1) = "" Then
        InitializeDlg1 = 1: Exit Function
   End If
   Set myRange = ActiveSheet.Cells.CurrentRegion.Rows(1)
   ParentDlg.Listbox1.Clear: ParentDlg.ListBox2.Clear
   cnt = myRange.Cells.count
   
   ReDim myArray(0 To cnt - 1)
   For i = 1 To cnt
     myArray(i - 1) = myRange.Cells(i)
   Next i
   ParentDlg.Listbox1.list() = myArray
   InitializeDlg1 = 0
   Exit Function
   
ErrorFlag:
   InitializeDlg1 = -1
   
End Function

Sub CharShow()

    Dim ErrSignforDataSheet As Integer
    
    On Error Resume Next
    ErrSignforDataSheet = InitializeDlg1(frmChar)
    
    Select Case ErrSignforDataSheet
    Case 0: frmChar.Show
    Case -1
        MsgBox "��Ʈ�� ��ȣ���¿� �ֽ��ϴ�." & _
               "����Ÿ�� ���� �� �����ϴ�.", _
                vbExclamation, "DB �м�"
    Case 1
        MsgBox "��Ʈ�� ����Ÿ�� �ִ��� Ȯ���Ͻʽÿ�." & _
               "1��1������ �����̸��� �Է��ؾ� �մϴ�.", _
                vbExclamation, "DB �м�"
    Case Else
    End Select

End Sub



Sub CharAnalysis(myRange, list)
    
    Dim temp As Worksheet
    Dim MyNewRange As Range
    Dim i As Long, M2 As Long, num As Integer
    Dim yp As Double
    Dim k As Integer
    Dim MakePivot As PivotTable
    Dim a As Long
    Dim cou As Integer
    Dim test()
    Dim sum, aa As Integer
    Set temp = Worksheets(rstSheet)
        
   
    ModulePrint.Title1 "���ڿ�������"
    
    NumOfTable = temp.Cells(1, 2).Value
    If NumOfTable = "" Then NumOfTable = 0
        
    For k = 1 To p
      
        ModulePrint.Title3 xlist(k - 1)
        
        i = temp.Cells(1, 1).Value
    
        num = NumOfTable + k
          
          
        temp.PivotTableWizard SourceType:=xlDatabase, SourceData:=myRange, TableDestination:=temp.Cells(i, 3), TableName:="Table" & num
          
        Set MakePivot = ActiveSheet.PivotTables("Table" & num)
        MakePivot.SmallGrid = True
        MakePivot.AddFields RowFields:=xlist(k - 1)
          
        With MakePivot.PivotFields(xlist(k - 1))
           .Orientation = xlDataField
           .Caption = "�󵵼�" & " " & xlist(k - 1)
           .position = 1
           .Function = xlCount
        End With
         
        MakePivot.PivotSelect ""
        Selection.Font.Name = "����ü"
        Selection.Font.Size = 9
        
        Application.CommandBars("PivotTable").Visible = False
       
        i = ActiveCell.SpecialCells(xlLastCell).row
        ActiveSheet.Cells(i + 1, 1).Select
        temp.Cells(1, 1) = i + 4
            
    Next k
     
    temp.Cells(1, 2) = NumOfTable + p
        
End Sub


Sub DescriptiveShow()

    Dim ErrSignforDataSheet As Integer
    
    ErrSignforDataSheet = InitializeDlg(frmDescriptive)
    On Error Resume Next
    Select Case ErrSignforDataSheet
    Case 0: frmDescriptive.Show
    Case -1
        MsgBox "��Ʈ�� ��ȣ���¿� �ֽ��ϴ�." & Chr(10) & _
               "����Ÿ�� ���� �� �����ϴ�.", _
                vbExclamation, "DB �м�"
    Case 1
        MsgBox "��Ʈ�� ����Ÿ�� �ִ��� Ȯ���Ͻʽÿ�." & Chr(10) & _
               "1��1������ �����̸��� �Է��ؾ� �մϴ�.", _
                vbExclamation, "DB �м�"
    Case Else
    End Select

End Sub
Function DBFindingRangeError(rn) As Boolean
    
    Dim i As Long

    For i = 1 To rn.Cells.count - 1
    
    If IsNumeric(rn.Cells(1 + i, 1)) = False Then
        DBFindingRangeError = True
       Exit Function
    End If
   Next i
   
End Function

Function DBFindVarRange(ListVar) As Range
  
    Dim TempSheet As Worksheet
    Dim TempRange As Range
    Dim cnt As Long
   
    Set TempSheet = ActiveCell.Worksheet
    Set TempRange = TempSheet.Cells.CurrentRegion
       cnt = TempRange.Columns.count
    For J = 1 To cnt
       If StrComp(ListVar, TempSheet.Cells(1, J).Value, 1) = 0 Then
          Set DBFindVarRange = TempRange.Columns(J)
          Exit Function
       End If
    Next J
     
End Function
Sub DescriptiveAnalysis(myRange, MyVarName, LevelVarRange, LevelVariable, StartRow, NumofVariable, OutPutSheetName, cnt)

''' HIST1�� �����ϰ�, �� sub�� ������� �ʱ�� ��.
    
    Dim temp As Worksheet, temp2 As Worksheet, Temp2Cells As Range
    Dim LevelOutPutRange As Range, EachLevelRange() As Range
    Dim OutPutRange As Range, ValuesOfLevel() As String, TrueLevel() As String
    Dim CompareValue As String, k As Long, NumOfTable As Integer
    Dim i As Integer, LevelNumber As Integer, J As Long, ii As Long, KK As Long, jj As Long, kkk As Integer
    Dim yp As Double, iii As Integer, jjj As Long
    Dim EachLevelVar() As Variant, number As Long
    Dim IsBlank As Boolean, OutPutArray() As Variant, EachLevelArray() As Variant
    Dim title As Shape, a As Integer, aa As Integer, b As Integer, BB As Integer
    
    Set temp = Worksheets(OutPutSheetName)
    Application.DisplayAlerts = False
    On Error Resume Next
  ''''''''''''''''''1)
    If LevelVariable <> "" Then
     
       number = LevelVarRange.Cells.count - 1
       
       MakeTempWorkSheet
       Set Temp2Cells = Worksheets("_#Temp#_").Range(Cells(1, 1), Cells(number, 1))
       
       For jj = 1 To number
         If LevelVarRange.Cells(jj + 1) = "" Then
            Temp2Cells.Cells(jj, 1) = "zzz"
            IsBlank = True
         Else
            Temp2Cells.Cells(jj, 1) = LevelVarRange.Cells(jj + 1)
         End If
      
       Next jj
       
       Temp2Cells.Sort Key1:=Temp2Cells.Cells(1, 1), Order1:=xlAscending, Header:=xlNo
       
       ReDim ValuesOfLevel(0 To number - 1)
       
       ValuesOfLevel(0) = Temp2Cells.Cells(1, 1)
       
       CompareValue = ValuesOfLevel(0)
        
       J = 0
      
         For KK = 2 To number
            If Temp2Cells.Cells(KK) <> CompareValue Then
               J = J + 1
               ValuesOfLevel(J) = Temp2Cells.Cells(KK)
               CompareValue = ValuesOfLevel(J)
             End If
         Next KK
      
        ReDim Preserve ValuesOfLevel(J)
      
        LevelNumber = J + 1
    
   End If
    
    
        
    For k = 1 To cnt
              
        If LevelVariable = "" Then
              
          i = temp.Cells(1, 2)
          yp = temp.Cells(i + 3, 1).Top
          Set title = temp.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 60, 20#)
          title.Shadow.Type = msoShadow17
          With title.Fill
             .ForeColor.SchemeColor = 22
             .Visible = msoTrue
             .Solid
          End With
                 
          title.TextFrame.Characters.Text = MyVarName(k - 1)
          With title.TextFrame.Characters.Font
             .Name = "����"
             .FontStyle = "����"
             .Size = 11
             .ColorIndex = xlAutomatic
          End With
          title.TextFrame.HorizontalAlignment = xlCenter
    
        Else
    
          i = temp.Cells(1, 2)
          yp = temp.Cells(i + 3, 1).Top
          Set title = temp.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 90, 20#)
          title.Shadow.Type = msoShadow17
          With title.Fill
             .ForeColor.SchemeColor = 22
             .Visible = msoTrue
             .Solid
          End With
                 
          title.TextFrame.Characters.Text = MyVarName(k - 1) & "*" & LevelVariable
          With title.TextFrame.Characters.Font
             .Name = "����"
             .FontStyle = "����"
             .Size = 11
             .ColorIndex = xlAutomatic
          End With
          title.TextFrame.HorizontalAlignment = xlCenter
    
       End If
    
    
          '��½�Ʈ�� ��µ� ������
         ' Num = NumofVariable + k
    
        ' ReDim OutPutRange(0 To Cnt - 1)
     
          temp.Activate
        ' Set OutPutRange(k - 1) = temp.Range(Cells(i + 5, 3), Cells(i + 20, 4))
        
        ReDim OutPutArray(0 To 1, 0 To 14)
          
        OutPutArray(0, 0) = MyVarName(k - 1)
        OutPutArray(0, 2) = "���"
        OutPutArray(0, 3) = "ǥ�ؿ���"
        OutPutArray(0, 4) = "�߾Ӱ�"
        OutPutArray(0, 5) = "�ֺ�"
        OutPutArray(0, 6) = "ǥ������"
        OutPutArray(0, 7) = "�л�"
        OutPutArray(0, 8) = "÷��"
        OutPutArray(0, 9) = "�ֵ�"
        OutPutArray(0, 10) = "����"
        OutPutArray(0, 11) = "�ּҰ�"
        OutPutArray(0, 12) = "�ִ밪"
        OutPutArray(0, 13) = "��"
        OutPutArray(0, 14) = "������"
         
        OutPutArray(1, 0) = ""
        OutPutArray(1, 1) = ""
        OutPutArray(1, 2) = Application.Average(myRange(k - 1))
        OutPutArray(1, 4) = Application.median(myRange(k - 1))
        OutPutArray(1, 5) = Application.mode(myRange(k - 1))
        OutPutArray(1, 6) = Application.StDev(myRange(k - 1))
        OutPutArray(1, 7) = OutPutArray(1, 6) ^ 2
        OutPutArray(1, 8) = Application.Kurt(myRange(k - 1))
        OutPutArray(1, 9) = Application.Skew(myRange(k - 1))
        OutPutArray(1, 11) = Application.min(myRange(k - 1))
        OutPutArray(1, 12) = Application.max(myRange(k - 1))
        OutPutArray(1, 10) = OutPutArray(1, 12) - OutPutArray(1, 11)
        OutPutArray(1, 13) = Application.sum(myRange(k - 1))
        OutPutArray(1, 14) = Application.count(myRange(k - 1))
        OutPutArray(1, 3) = OutPutArray(1, 6) / Sqr(OutPutArray(1, 14))
                    
     
     If LevelVariable <> "" Then
     
        ReDim EachLevelArray(0 To LevelNumber - 1, 0 To 15)
        ReDim EachLevelVar(0 To number)
       
        Worksheets("_#Temp#_").Delete
       
        On Error Resume Next
      
      If IsBlank = True Then ValuesOfLevel(LevelNumber - 1) = ""
          
      For iii = 1 To LevelNumber
            
            jjj = 0
            For kkk = 2 To number + 1
               If LevelVarRange.Cells(kkk) = ValuesOfLevel(iii - 1) Then
                  EachLevelVar(jjj) = myRange(k - 1).Cells(kkk)
               End If
                  jjj = jjj + 1
            Next kkk
        
         ReDim Preserve EachLevelVar(jjj)
     
          EachLevelArray(iii - 1, 0) = ValuesOfLevel(iii - 1)
          EachLevelArray(iii - 1, 1) = ""
          EachLevelArray(iii - 1, 2) = Application.Average(EachLevelVar)
          EachLevelArray(iii - 1, 4) = Application.median(EachLevelVar)
          EachLevelArray(iii - 1, 5) = Application.mode(EachLevelVar)
          EachLevelArray(iii - 1, 6) = Application.StDev(EachLevelVar)
          EachLevelArray(iii - 1, 7) = EachLevelArray(iii - 1, 6) ^ 2
          EachLevelArray(iii - 1, 8) = Application.Kurt(EachLevelVar)
          EachLevelArray(iii - 1, 9) = Application.Skew(EachLevelVar)
          EachLevelArray(iii - 1, 11) = Application.min(EachLevelVar)
          EachLevelArray(iii - 1, 12) = Application.max(EachLevelVar)
          EachLevelArray(iii - 1, 10) = EachLevelArray(iii - 1, 12) - EachLevelArray(iii - 1, 11)
          EachLevelArray(iii - 1, 13) = Application.sum(EachLevelVar)
          EachLevelArray(iii - 1, 14) = Application.count(EachLevelVar)
          EachLevelArray(iii - 1, 3) = EachLevelArray(iii - 1, 6) / Sqr(EachLevelArray(iii - 1, 14))
          EachLevelArray(iii - 1, 15) = ""
                    
        
        ReDim EachLevelVar(0 To number)
     
     Next iii
  
  End If
  
    Set OutPutRange = temp.Range(Cells(i + 5, 3), Cells(i + 20, 4))

       For aa = 1 To 2
          For a = 1 To 15
            OutPutRange.Cells(a, aa) = OutPutArray(aa - 1, a - 1)
          Next a
       Next aa
    
        OutPutRange.Columns(1).HorizontalAlignment = xlLeft
        OutPutRange.Font.Name = "����"
        OutPutRange.Font.Size = 9
        With OutPutRange.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With OutPutRange.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    
        With OutPutRange.Rows(1).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThick
            .ColorIndex = xlAutomatic
        End With
        
  If LevelVariable <> "" Then
   
     Set LevelOutPutRange = temp.Range(Cells(i + 5, 6), Cells(i + 20, 5 + LevelNumber))
    
       For BB = 1 To LevelNumber
          For b = 1 To 16
            LevelOutPutRange.Cells(b, BB) = EachLevelArray(BB - 1, b - 1)
          Next b
       Next BB
       
       LevelOutPutRange.Font.Name = "����"
       LevelOutPutRange.Font.Size = 9
       
       With LevelOutPutRange.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
       End With
       
       With LevelOutPutRange.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
       End With
        
       With LevelOutPutRange.Rows(1).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThick
            .ColorIndex = xlAutomatic
       End With
           
  End If
  
     i = ActiveCell.SpecialCells(xlLastCell).row
     ActiveSheet.Cells(i + 1, 1).Select
     temp.Cells(1, 2) = i
     ReDim EachLevelArray(0)
  Next k
    
     temp.Cells(1, 2) = i + 2
     temp.Cells(2, 2) = NumOfTable + cnt
     
     temp.Cells(1, 1) = i + 2
     
End Sub

''�ӽý�Ʈ �����
Sub MakeTempWorkSheet()
    
    Dim WS As Worksheet
    
    For Each WS In Worksheets
        If WS.Name = "_#Temp#_" Then
           Exit Sub
        End If
    Next WS
    
    Worksheets.Add.Name = "_#Temp#_"
    Worksheets("_#Temp#_").Visible = True
   
        
End Sub
                                     

Sub CorrShow()

    Dim ErrSignforDataSheet As Integer
    
    ErrSignforDataSheet = InitializeDlg(frmCorr)
                                    
    Select Case ErrSignforDataSheet
    Case 0: frmCorr.Show
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



Sub CorrAnal()
        
    Dim dataX(), rst1(), rst2()
    Dim row As Integer, col As Integer
    Dim index As Integer, Flag As Long
    Dim mySheet As Worksheet, tmpSheet As Worksheet, WS As Worksheet
    Dim pt As Range, tmp_pt As Range
    Dim alpha1, alpha2
    Dim alpha11(), alpha22(), corr11(), corr22()
    Dim tmpCheck As Integer
    Dim rr
    
    
    dataX = arrayX(xlist, p)
    
    ModulePrint.Title1 "����м����"
    ModulePrint.Title3 "����м�"
    
    
    
    tmpCheck = 0
    
    For Each WS In Worksheets
        If WS.Name = "_#TempSurvey#_" Then
        tmpCheck = 1
        End If
    Next WS
    
    If tmpCheck = 0 Then
        Worksheets.Add.Name = "_#TempSurvey#_"
    End If
    
    Worksheets("_#TempSurvey#_").Visible = False
   
    
    Set tmpSheet = Worksheets("_#TempSurvey#_")
    tmpSheet.Cells(1, 1).Value = 2
    Flag = tmpSheet.Cells(1, 1).Value
            
    'If p > 1 Then
    '    ModulePrint.TABLE n + 3, 1 + p + 1 + p
    'Else
    '    ModulePrint.TABLE n + 3, 1 + p + 1
    'End If
    
    Set tmp_pt = tmpSheet.Cells(Flag, 2)
    tmp_pt(1, 1) = "��ȣ"
    tmp_pt(1 + N + 1, 1) = "�л�"
    tmp_pt(1 + N + 2, 1) = "���"
    With tmp_pt(1 + N + 1, 1).Resize(, 1 + p + 1).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    For row = 1 To N
        tmp_pt(1 + row, 1) = row
    Next row
    Set tmp_pt = tmpSheet.Cells(Flag, 3)
    
    Range(tmp_pt(1, 1), tmp_pt(N + 1, p)) = dataX
    
    tmp_pt(1, p + 1) = "������"
    
    For row = 1 To N
        tmp_pt(1 + row, p + 1).FormulaR1C1 = WorksheetFunction.sum(Range(tmp_pt(row + 1, 1), tmp_pt(row + 1, p)))
    Next row
    For col = 1 To p + 1
        tmp_pt(1 + N + 1, col) = WorksheetFunction.var(Range(tmp_pt(1 + 1, col), tmp_pt(1 + N, col)))
        tmp_pt(1 + N + 2, col) = WorksheetFunction.Average(Range(tmp_pt(1 + 1, col), tmp_pt(1 + N, col)))
    Next col
    
    If p > 1 Then
        tmp_pt(1, 1) = Chr(10) & Chr(10) & "��ȣ"
        tmp_pt(1, p + 1) = Chr(10) & Chr(10) & "������"
        For col = 1 To p
            tmp_pt(1, col) = Chr(10) & Chr(10) & xlist(col - 1)
            tmp_pt(1, p + 1 + col) = xlist(col - 1) & Chr(10) & "����" & Chr(10) & "������"
        Next col
        For row = 1 To N
            tmp_pt(1 + row, p + 2).FormulaR1C1 = WorksheetFunction.sum(Range(tmp_pt(row + 1, 2), tmp_pt(row + 1, p)))
            tmp_pt(1 + row, p + 1 + p).FormulaR1C1 = WorksheetFunction.sum(Range(tmp_pt(row + 1, 1), tmp_pt(row + 1, p - 1)))
        Next row
        For col = 2 To p - 1
            For row = 1 To N
                tmp_pt(1 + row, p + 1 + col).FormulaR1C1 = WorksheetFunction.sum(Range(tmp_pt(row + 1, 1), tmp_pt(row + 1, col - 1)), Range(tmp_pt(row + 1, col + 1), tmp_pt(row + 1, p)))
            Next row
        Next col
        For col = 1 To p
            tmp_pt(1 + N + 1, p + 1 + col) = WorksheetFunction.var(Range(tmp_pt(1 + 1, p + 1 + col), tmp_pt(1 + N, p + 1 + col)))
            tmp_pt(1 + N + 2, p + 1 + col) = WorksheetFunction.Average(Range(tmp_pt(1 + 1, p + 1 + col), tmp_pt(1 + N, p + 1 + col)))
        Next col
        With tmp_pt(1 + N + 1, 1).Resize(, p + 1 + p).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    End If
    
    tmp_pt.Cells(1 + N + 1, 1).Resize(2, p + 1 + p).NumberFormatLocal = "0.000_ "
    
    tmpSheet.Cells(1, 1) = Flag + N + 3 + 2
    
    
    Flag = tmpSheet.Cells(1, 1).Value
        
    'If p > 1 Then
    '    ModulePrint.TABLE 4, 1 + p + 1 + p
    'Else
    '    ModulePrint.TABLE 4, 1 + p + 1
    'End If
    
    Set tmp_pt = tmpSheet.Cells(Flag, 2)
    tmp_pt(2, 1) = "���׺� �л���"
    tmp_pt(3, 1) = "ũ�й� ��"
    tmp_pt(4, 1) = "�κ��հ� ������"
    
    tmp_pt(1, 1 + p + 1) = "������"
    tmp_pt(2, 1 + p + 1).FormulaR1C1 = WorksheetFunction.sum(Range(tmp_pt(-3, 2), tmp_pt(-3, p + 1)))
    If p > 1 Then
        tmp_pt(3, 1 + p + 1) = p / (p - 1) * (1 - tmp_pt(2, 1 + p + 1) / tmp_pt(-3, 1 + p + 1))
        alpha1 = p / (p - 1) * (1 - tmp_pt(2, 1 + p + 1) / tmp_pt(-3, 1 + p + 1))
    End If
    
    If p > 1 Then
        tmp_pt(1, 1 + p + 1) = Chr(10) & "������"
        For col = 1 To p
            tmp_pt(1, 1 + p + 1 + col) = xlist(col - 1) & Chr(10) & "����"
            tmp_pt(1, 1 + col) = Chr(10) & xlist(col - 1)
        Next col
        tmp_pt(2, 1 + p + 1 + 1).FormulaR1C1 = WorksheetFunction.sum(Range(tmp_pt(-3, 3), tmp_pt(-3, 1 + p)))
        tmp_pt(2, 1 + p + 1 + p).FormulaR1C1 = WorksheetFunction.sum(Range(tmp_pt(-3, 2), tmp_pt(-3, p)))
        For col = 2 To p - 1
            tmp_pt(2, 1 + p + 1 + col).FormulaR1C1 = WorksheetFunction.sum(Range(tmp_pt(-3, 2), tmp_pt(-3, col)), Range(tmp_pt(-3, col + 2), tmp_pt(-3, 1 + p)))
        Next col
        If p > 2 Then
        ReDim alpha11(p)
        For col = 1 To p
            tmp_pt(3, 1 + p + 1 + col) = (p - 1) / (p - 2) * (1 - tmp_pt(2, 1 + p + 1 + col) / tmp_pt(-3, 1 + p + 1 + col))
            alpha11(col) = (p - 1) / (p - 2) * (1 - tmp_pt(2, 1 + p + 1 + col) / tmp_pt(-3, 1 + p + 1 + col))
        Next col
        End If
    End If
    
    ReDim corr11(p)
    If p > 1 Then
        For col = 1 To p
            tmp_pt(4, 1 + col) = WorksheetFunction.Correl(Range(tmp_pt(-4, 1 + col), tmp_pt(-4 - N + 1, 1 + col)), Range(tmp_pt(-4, 1 + col + 1 + p), tmp_pt(-4 - N + 1, 1 + col + 1 + p)))
            corr11(col) = WorksheetFunction.Correl(Range(tmp_pt(-4, 1 + col), tmp_pt(-4 - N + 1, 1 + col)), Range(tmp_pt(-4, 1 + col + 1 + p), tmp_pt(-4 - N + 1, 1 + col + 1 + p)))
        Next col
    Else
        For col = 1 To p
            tmp_pt(4, 1 + col) = WorksheetFunction.Correl(Range(tmp_pt(-4, 1 + col), tmp_pt(-4 - N + 1, 1 + col)), Range(tmp_pt(-4, 1 + col), tmp_pt(-4 - N + 1, 1 + col)))
            corr11(col) = WorksheetFunction.Correl(Range(tmp_pt(-4, 1 + col), tmp_pt(-4 - N + 1, 1 + col)), Range(tmp_pt(-4, 1 + col), tmp_pt(-4 - N + 1, 1 + col)))
        Next col
    End If
    
    tmp_pt.Cells(2, 2).Resize(3, p + 1 + p).NumberFormatLocal = "0.000_ "
    
    tmpSheet.Cells(1, 1) = Flag + 4 + 4
    
    
    
    
    
    
    Flag = tmpSheet.Cells(1, 1).Value
    
    'If p > 1 Then
    '    ModulePrint.TABLE n + 3, 1 + p + 1 + p
    'Else
    '    ModulePrint.TABLE n + 3, 1 + p + 1
    'End If
    
    Set tmp_pt = tmpSheet.Cells(Flag, 2)
    tmp_pt(1, 1) = "��ȣ"
    tmp_pt(1 + N + 1, 1) = "�л�"
    tmp_pt(1 + N + 2, 1) = "���"
    With tmp_pt(1 + N + 1, 1).Resize(, 1 + p + 1).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    For row = 1 To N
        tmp_pt(1 + row, 1) = row
    Next row
    Set tmp_pt = tmpSheet.Cells(Flag, 3)
    
    
    For col = 1 To p
        For row = 1 To N
            tmp_pt(1 + row, col) = WorksheetFunction.Standardize(tmp_pt(-11 - N + row - 1, col), WorksheetFunction.Average(Range(tmp_pt(-11 - N, col), tmp_pt(-11 - 1, col))), WorksheetFunction.StDev(Range(tmp_pt(-11 - N, col), tmp_pt(-11 - 1, col))))
        Next row
    Next col
    
    
    tmp_pt(1, p + 1) = "������"
    
    For row = 1 To N
        tmp_pt(1 + row, p + 1).FormulaR1C1 = WorksheetFunction.sum(Range(tmp_pt(row + 1, 1), tmp_pt(row + 1, p)))
    Next row
    For col = 1 To p + 1
        tmp_pt(1 + N + 1, col) = WorksheetFunction.var(Range(tmp_pt(1 + 1, col), tmp_pt(1 + N, col)))
        tmp_pt(1 + N + 2, col) = WorksheetFunction.Average(Range(tmp_pt(1 + 1, col), tmp_pt(1 + N, col)))
    Next col
    
    If p > 1 Then
        tmp_pt(1, 1) = Chr(10) & Chr(10) & "��ȣ"
        tmp_pt(1, p + 1) = Chr(10) & Chr(10) & "������"
        For col = 1 To p
            tmp_pt(1, col) = Chr(10) & Chr(10) & xlist(col - 1)
            tmp_pt(1, p + 1 + col) = xlist(col - 1) & Chr(10) & "����" & Chr(10) & "������"
        Next col
        For row = 1 To N
            tmp_pt(1 + row, p + 2).FormulaR1C1 = WorksheetFunction.sum(Range(tmp_pt(row + 1, 2), tmp_pt(row + 1, p)))
            tmp_pt(1 + row, p + 1 + p).FormulaR1C1 = WorksheetFunction.sum(Range(tmp_pt(row + 1, 1), tmp_pt(row + 1, p - 1)))
        Next row
        For col = 2 To p - 1
            For row = 1 To N
                tmp_pt(1 + row, p + 1 + col).FormulaR1C1 = WorksheetFunction.sum(Range(tmp_pt(row + 1, 1), tmp_pt(row + 1, col - 1)), Range(tmp_pt(row + 1, col + 1), tmp_pt(row + 1, p)))
            Next row
        Next col
        For col = 1 To p
            tmp_pt(1 + N + 1, p + 1 + col) = WorksheetFunction.var(Range(tmp_pt(1 + 1, p + 1 + col), tmp_pt(1 + N, p + 1 + col)))
            tmp_pt(1 + N + 2, p + 1 + col) = WorksheetFunction.Average(Range(tmp_pt(1 + 1, p + 1 + col), tmp_pt(1 + N, p + 1 + col)))
        Next col
        With tmp_pt(1 + N + 1, 1).Resize(, p + 1 + p).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    End If
    
    tmp_pt.Cells(1, 1).Resize(N + 3, p + 1 + p).NumberFormatLocal = "0.000_ "
    
    tmpSheet.Cells(1, 1) = Flag + N + 3 + 2
    
    
    Flag = tmpSheet.Cells(1, 1).Value
        
    'If p > 1 Then
    '    ModulePrint.TABLE 4, 1 + p + 1 + p
    'Else
    '    ModulePrint.TABLE 4, 1 + p + 1
    'End If
    
    Set tmp_pt = tmpSheet.Cells(Flag, 2)
    tmp_pt(2, 1) = "���׺� �л���"
    tmp_pt(3, 1) = "ũ�й� ��"
    tmp_pt(4, 1) = "�κ��հ� ������"
    
    tmp_pt(1, 1 + p + 1) = "������"
    tmp_pt(2, 1 + p + 1).FormulaR1C1 = WorksheetFunction.sum(Range(tmp_pt(-3, 2), tmp_pt(-3, p + 1)))
    If p > 1 Then
        tmp_pt(3, 1 + p + 1) = p / (p - 1) * (1 - tmp_pt(2, 1 + p + 1) / tmp_pt(-3, 1 + p + 1))
        alpha2 = p / (p - 1) * (1 - tmp_pt(2, 1 + p + 1) / tmp_pt(-3, 1 + p + 1))
    End If
    
    If p > 1 Then
        tmp_pt(1, 1 + p + 1) = Chr(10) & "������"
        For col = 1 To p
            tmp_pt(1, 1 + p + 1 + col) = xlist(col - 1) & Chr(10) & "����"
            tmp_pt(1, 1 + col) = Chr(10) & xlist(col - 1)
        Next col
        tmp_pt(2, 1 + p + 1 + 1).FormulaR1C1 = WorksheetFunction.sum(Range(tmp_pt(-3, 3), tmp_pt(-3, 1 + p)))
        tmp_pt(2, 1 + p + 1 + p).FormulaR1C1 = WorksheetFunction.sum(Range(tmp_pt(-3, 2), tmp_pt(-3, p)))
        For col = 2 To p - 1
            tmp_pt(2, 1 + p + 1 + col).FormulaR1C1 = WorksheetFunction.sum(Range(tmp_pt(-3, 2), tmp_pt(-3, col)), Range(tmp_pt(-3, col + 2), tmp_pt(-3, 1 + p)))
        Next col
        If p > 2 Then
        ReDim alpha22(p)
        For col = 1 To p
            tmp_pt(3, 1 + p + 1 + col) = (p - 1) / (p - 2) * (1 - tmp_pt(2, 1 + p + 1 + col) / tmp_pt(-3, 1 + p + 1 + col))
            alpha22(col) = (p - 1) / (p - 2) * (1 - tmp_pt(2, 1 + p + 1 + col) / tmp_pt(-3, 1 + p + 1 + col))
        Next col
        End If
    End If
    
    ReDim corr22(p)
    If p > 1 Then
        For col = 1 To p
            tmp_pt(4, 1 + col) = WorksheetFunction.Correl(Range(tmp_pt(-4, 1 + col), tmp_pt(-4 - N + 1, 1 + col)), Range(tmp_pt(-4, 1 + col + 1 + p), tmp_pt(-4 - N + 1, 1 + col + 1 + p)))
            corr22(col) = WorksheetFunction.Correl(Range(tmp_pt(-4, 1 + col), tmp_pt(-4 - N + 1, 1 + col)), Range(tmp_pt(-4, 1 + col + 1 + p), tmp_pt(-4 - N + 1, 1 + col + 1 + p)))
        Next col
    Else
        For col = 1 To p
            tmp_pt(4, 1 + col) = WorksheetFunction.Correl(Range(tmp_pt(-4, 1 + col), tmp_pt(-4 - N + 1, 1 + col)), Range(tmp_pt(-4, 1 + col), tmp_pt(-4 - N + 1, 1 + col)))
            corr22(col) = WorksheetFunction.Correl(Range(tmp_pt(-4, 1 + col), tmp_pt(-4 - N + 1, 1 + col)), Range(tmp_pt(-4, 1 + col), tmp_pt(-4 - N + 1, 1 + col)))
        Next col
    End If
    tmp_pt.Cells(2, 2).Resize(3, p + 1 + p).NumberFormatLocal = "0.000_ "
    
    tmpSheet.Cells(1, 1) = Flag + 4 + 4
    
    
    
    
    
    Set mySheet = Worksheets(rstSheet)
    Flag = mySheet.Cells(1, 1).Value
    Set pt = mySheet.Cells(Flag, 2)
    
    Set tmp_pt = tmpSheet.Cells(tmpSheet.Cells(1, 1).Value, 2)
    
    pt(1, 1) = "������" & Chr(10) & "(����Ȯ��)"
    pt(1, 1).HorizontalAlignment = xlLeft
    With pt.Resize(1, 1).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    
    
    
    For col = 1 To p
        pt(3 + 2 * col - 1, 1) = xlist(col - 1)
        pt(3 + 2 * col, 1) = "(����Ȯ��)"
        pt(3, 1 + col) = xlist(col - 1)
    Next col
        
    For col = 1 To p
        For row = 1 To p
            rr = WorksheetFunction.Correl(Range(tmp_pt(-24 - N - N, 1 + col), tmp_pt(-24 - N - 1, 1 + col)), Range(tmp_pt(-24 - N - N, 1 + row), tmp_pt(-24 - N - 1, 1 + row)))
            pt(3 + 2 * col - 1, row + 1) = WorksheetFunction.Round(rr, 4)
            If col = row Then
                pt(3 + 2 * col, row + 1) = "."
            Else
                pt(3 + 2 * col, row + 1) = WorksheetFunction.Round(WorksheetFunction.TDist(rr / WorksheetFunction.ImSqrt((1 - rr ^ 2) / (N - 2)), N - 2, 2), 4)
 '               pt(3 + 2 * col, row + 1).NumberFormatLocal = "@"
 '               pt(3 + 2 * col, row + 1) = "(" & pt(3 + 2 * col, row + 1).Value & ")"
            End If
        Next row
    Next col
    ''pt.Cells(4, 2).Resize(2 * p, p).NumberFormatLocal = "0.000_ "
    ''' n=1 or 2 ����ó�� �ʿ�.
       
    
    With pt(2, 1).Resize(1, 1 + p).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With pt(3, 1).Resize(1, 1 + p).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With pt(3 + 2 * p, 1).Resize(1, 1 + p).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    With pt(3, 1).Resize(1 + 2 * p, 1).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With pt(3, 1 + p).Resize(1 + 2 * p, 1).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With pt(3, 1).Resize(1 + 2 * p, 1).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With pt(3, 1 + p).Resize(1 + 2 * p, 1).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    mySheet.Cells(1, 1) = Flag + 1 + 2 * p + 4
    
    
    
    
    


End Sub

Function ratio(X As Variant, test, order) As Variant() '������ �̸��� ������ ����ؼ� ǥ�� ������ش�'
Dim fac() As String
Dim �з��̸�() As String
Dim temp() As Long
Dim number() As Long
Dim �� As Long
Dim �� As Long
Dim i As Long
Dim J As Long
Dim ���� As Long
Dim ����ī���� As Long
Dim ���� As Double
Dim ����� As Double
Dim ���() As Variant



�� = UBound(X) '���'
�� = UBound(X, 2) '����'
Call cswap(X, 1, order) '1�� �� ���� �ٲٴ� �����ƾ'
Call QuickSort(X)     '1�� �������� �����ź��� �����ϴ� �����ƾ'
Call cswap(X, 1, order)  '��������� �����Ѱ� ���������� ���ݺ��� ���� �����ϰ� ��������'
'-------------------'

ReDim fac(1 To ��)  '������ �̸��� �����ϱ� ���� �������'
ReDim temp(1 To ��)
fac(1) = X(1, order)
���� = 1
'����1����, �� �ڷᰡ �Ѱ���? �����Ϸ� ������'
If �� = 1 Then
  ReDim ���(2, 1)
  ���(1, 1) = X(1, order)
  ���(2, 1) = 1
Else

For i = 2 To ��
If X(i, order) <> X(i - 1, order) Then
���� = ���� + 1
fac(����) = X(i, order)
End If
Next i

ReDim �з��̸�(1 To ����) '�з��� �̸� ����'
For i = 1 To ����
�з��̸�(i) = fac(i)
Next i
'�̸����� �����ϱ�'
'-----------------------'
ReDim number(1 To ����) ' �� �з��� ���� ������ ����'

���� = 1
����ī���� = 1 '�ʱ�ȭ'
'���� �� ���� ��������'
For i = 2 To ��      '�տ��Ͱ� ������ ī��Ʈ �ø��� �ƴϸ� ���� �ϴ� ������ ����'
 If X(i, order) = X(i - 1, order) Then
 ����ī���� = ����ī���� + 1

  End If

 If X(i, order) <> X(i - 1, order) Then
 number(����) = ����ī����
 ����ī���� = 1
   ���� = ���� + 1
   End If
   number(����) = ����ī���� '������ ������ ���������� �־���. ��ȿ���� �������'
Next i
ReDim ���(����, 3)
For i = 1 To ����
���(i, 1) = �з��̸�(i)
���(i, 2) = number(i) / ��
���(i, 3) = number(i)
Next i

End If

ReDim ratio(����, 3)
ReDim test(����, 3)
ratio = ���
test = ���
End Function
' i���� j���� �ٲٴ� sub '
Sub cswap(ByRef Values As Variant, ByVal i As Long, ByVal J As Long)
  Dim temp()
  Dim temp2 As Double
  Dim row As Long
  Dim �� As Long
  
If i = J Then
Exit Sub
End If

row = UBound(Values, 1) '�ళ�� ���Ⱑ ���� �ļ�'
  
ReDim temp(1 To row)
  
For �� = 1 To row
temp(��) = Values(��, i)
Next ��
' I ���� temp�� �����ߴ�'

For �� = 1 To row
Values(��, i) = Values(��, J)
Next ��

For �� = 1 To row
Values(��, J) = temp(��)
Next ��
' i �İ� j���� �ٲ��'

  
End Sub
Sub Swap(ByRef Values As Variant, ByVal i As Long, ByVal J As Long) 'i�ٰ� j�� �� �ٲ۴�'
  Dim temp()
  Dim temp2 As Double
  Dim col As Long
  Dim �� As Long
  
col = UBound(Values, 2) '������ ���Ⱑ ���� �ļ�'
  
ReDim temp(1 To col)
  
For �� = 1 To col
temp(��) = Values(i, ��)
Next ��
' I ���� temp�� �����ߴ�'

For �� = 1 To col
Values(i, ��) = Values(J, ��)
Next ��

For �� = 1 To col
Values(J, ��) = temp(��)
Next ��
' i �ٰ� j���� �ٲ��'

  
End Sub

'1���� �������� �����ϴ� �����ƾ'
 Sub QuickSort(ByRef Values As Variant, Optional ByVal leftt As Long, Optional ByVal right As Long)
  Dim i As Long
  Dim J As Long
  Dim Item1 As Variant
  Dim Item2 As Variant

  On Error GoTo Catch
  If IsMissing(leftt) Or leftt = 0 Then leftt = LBound(Values)
  If IsMissing(right) Or right = 0 Then right = UBound(Values)
  i = leftt
  J = right


  Item1 = Values((leftt + right) \ 2, 1)
  Do While i < J
    Do While Values(i, 1) < Item1 And i < right
      i = i + 1
    Loop
    Do While Values(J, 1) > Item1 And J > leftt
      J = J - 1
    Loop
    If i < J Then
      Call Swap(Values, i, J)
    End If
    If i <= J Then
      i = i + 1
      J = J - 1
    End If
  Loop
  If J > leftt Then Call QuickSort(Values, leftt, J)
  If i < right Then Call QuickSort(Values, i, right)
    Exit Sub

Catch:
  MsgBox Err.Description, vbCritical
  
End Sub

''''''��������
Function FindingRangeError2(rn) As Boolean
    
    Dim tmp1 As Range: Dim tmp2 As Range
    Dim tmp3 As Range
    
    On Error Resume Next
    
    If Application.CountBlank(rn) >= 1 Then
        FindingRangeError2 = True
        Exit Function
    End If
    Set tmp1 = rn.SpecialCells(xlCellTypeConstants, 22)
    Set tmp2 = rn.SpecialCells(xlCellTypeFormulas, 22)
    Set tmp3 = rn.SpecialCells(xlCellTypeBlanks)
    
    If rn.count = 1 And IsNumeric(rn.Cells(1, 1)) = True Then
        FindingRangeError2 = False
    Else
        If tmp1 Is Nothing And tmp2 Is Nothing And tmp3 Is Nothing Then
            FindingRangeError2 = False
        Else: FindingRangeError2 = True
        End If
    End If
    
End Function

Function OpenOutSheet2(SheetName, Optional IsAddress As Boolean = False) As Worksheet
    
    Dim s, CurS As Worksheet
    
    Application.ScreenUpdating = False
    For Each s In ActiveWorkbook.Sheets
        If s.Name = SheetName Then
            Set OpenOutSheet2 = s
            Exit Function
        End If
    Next s
    Set CurS = ActiveSheet: Set s = Worksheets.Add
    With ActiveWindow
        .DisplayGridlines = False
'        .DisplayHeadings = False
    End With
    
    With ActiveWindow.Application.Cells
         .Font.Name = "����"
         .Font.Size = 9
         .HorizontalAlignment = xlRight
    End With

    s.Name = SheetName: CurS.Activate
    With Worksheets(SheetName).Range("a1")
        .Value = 2
        '''If IsAddress = True Then .Value = "A2"
        .Font.ColorIndex = 2
    End With
    Worksheets(SheetName).Rows(1).Hidden = True
    
    's.Protect Password:="prophet", DrawingObjects:=False, contents:=True, Scenarios:=True
    Application.ScreenUpdating = True
    
    Set OpenOutSheet2 = s
    
End Function
Function SelectedVariable(ParentDlgLbxValue, selvar, _
         IsRowData As Boolean) As String   '��л�
   
   Dim temp, M2, m3 As Long
   Dim TempSheet As Worksheet
   Dim tmp2, tmp As Range
   
   Set TempSheet = ActiveCell.Worksheet
   

   Dim Chk_Ver As Boolean   '���� ���� üũ
   Dim Cmp_R As Long        '���� ������ ���� �� ���� ��
   Dim Cmp_C As Integer     '���� ������ ���� �� ���� ��
   
   '���� ������ ���� ��� ���� �񱳰� ����
   Chk_Ver = ChkVersion(ActiveWorkbook.Name)
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
              M2 = tmp2.Cells(1, 1).End(xlDown).row
              If M2 <> Cmp_R Then
                 m3 = tmp2.Cells(M2, 1).End(xlDown).row
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
''��л�
Sub PivotMakerforOneWay(DataRn, ColVn, DataVn, _
    cnt, ave, st, factor)
        
    Dim actSh, tmpSh As Worksheet
    Dim StartCell As String: Dim i, N As Long
    Dim temp As Range
    
    Set actSh = ActiveSheet
    Set tmpSh = Worksheets.Add
    actSh.Select
    StartCell = tmpSh.Name & "!R1C1"
    ActiveSheet.PivotTableWizard SourceType:=xlDatabase, SourceData:= _
        DataRn, TableDestination:=StartCell, TableName:="�ǹ� ���̺�1"
    
    ActiveSheet.PivotTables("�ǹ� ���̺�1").AddFields ColumnFields:=ColVn
    ActiveSheet.PivotTables("�ǹ� ���̺�1").PivotFields(DataVn).Orientation = _
        xlDataField
    ActiveSheet.PivotTables("�ǹ� ���̺�1").PivotFields(tmpSh.Cells(1, 1).Value).Function = xlCount
    ActiveSheet.PivotTables("�ǹ� ���̺�1").PivotSelect "", xlDataOnly
    N = Selection.Columns.count
    ReDim cnt(1 To N): ReDim ave(1 To N): ReDim st(1 To N): ReDim factor(1 To N)
    For i = 1 To N
        cnt(i) = Selection.Cells(i)
    Next i
    ActiveSheet.PivotTables("�ǹ� ���̺�1").PivotFields(tmpSh.Cells(1, 1).Value).Function = xlAverage
    ActiveSheet.PivotTables("�ǹ� ���̺�1").PivotSelect "", xlDataOnly
    For i = 1 To N
        ave(i) = Selection.Cells(i)
    Next i
    
    Set temp = Selection.Offset(-1, 0)
    For i = 1 To N
        factor(i) = temp.Cells(i)
    Next i
    
    ActiveSheet.PivotTables("�ǹ� ���̺�1").PivotFields(tmpSh.Cells(1, 1).Value).Function = xlStDev
    ActiveSheet.PivotTables("�ǹ� ���̺�1").PivotSelect "", xlDataOnly
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
''��л�
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
