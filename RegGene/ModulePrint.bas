Attribute VB_Name = "ModulePrint"
Sub MakeOutputSheet(SheetName, Optional IsAddress As Boolean = False)
    
    Dim s, CurS As Worksheet
    
    For Each s In ActiveWorkbook.Sheets
        If s.Name = SheetName Then Exit Sub
    Next s
    Set CurS = ActiveSheet: Set s = Worksheets.add
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
        .value = 2
        If IsAddress = True Then .value = "A2"
        .Font.ColorIndex = 2
    End With
    Worksheets(SheetName).Rows(1).Hidden = True
    Worksheets(SheetName).Activate
    Cells.Select
    Selection.RowHeight = 13.5
    
    
    's.Protect password:="prophet", DrawingObjects:=False, contents:=True, Scenarios:=True
    
End Sub

Sub Title1(contents As String)
    Dim Flag As Long
    Dim mySheet As Worksheet
    
    Set mySheet = Worksheets(RstSheet)
    Flag = mySheet.Cells(1, 1).value
    yp = mySheet.Cells(Flag + 2, 1).Top
    
    On Error Resume Next
    
    
    
    Set Title = mySheet.Shapes.AddShape(msoShapeRectangle, 3.75, yp + 2.5, 440, 25#)
    With Title
        .Fill.ForeColor.SchemeColor = 57
        .Line.DashStyle = msoLineSolid
        .Line.Style = msoLineSingle
        .Line.ForeColor.SchemeColor = 8
        .Line.Weight = 1
        .Line.Visible = msoTrue
        '.Shadow.Type = msoShadow1
    End With
   
    With Title.TextFrame.Characters
        .text = contents
        .Font.Name = "����"
        .Font.FontStyle = "����"
        .Font.Size = 14
        .Font.ColorIndex = 2
    End With
    Title.TextFrame.HorizontalAlignment = xlCenter
    
    mySheet.Cells(1, 1) = Flag + 4
End Sub

Sub Title2(contents As String)
    Dim Flag As Long
    Dim mySheet As Worksheet
    
    Set mySheet = Worksheets(RstSheet)
    Flag = mySheet.Cells(1, 1).value
    yp = mySheet.Cells(Flag + 2, 1).Top
    
    On Error Resume Next
    
    Set Title = mySheet.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 240, 25#)
    With Title
        .Fill.ForeColor.SchemeColor = 9
        .Solid
        .Line.Visible = msoTrue
        .Shadow.Type = msoShadow17
    End With
   
    With Title.TextFrame.Characters
        .text = contents
        .Font.Name = "����"
        .Font.FontStyle = "����"
        .Font.Size = 11
        .Font.ColorIndex = 41
    End With
    Title.TextFrame.HorizontalAlignment = xlCenter
    
    mySheet.Cells(1, 1) = Flag + 4
End Sub

Sub Title3(contents As String)

    Dim Flag As Long
    Dim mySheet As Worksheet
    
    Set mySheet = Worksheets(RstSheet)
    Flag = mySheet.Cells(1, 1).value
    yp = mySheet.Cells(Flag + 2, 1).Top
    
    On Error Resume Next
    
    Set Title = mySheet.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 240, 22#)
    With Title
        .Fill.ForeColor.SchemeColor = 9
        .Solid
        .Line.Visible = msoTrue
        .Shadow.Type = msoShadow17
    End With
   
    With Title.TextFrame.Characters
        .text = contents
        .Font.Name = "����"
        .Font.FontStyle = "����"
        .Font.Size = 11
        .Font.ColorIndex = xlAutomatic
    End With
    Title.TextFrame.HorizontalAlignment = xlCenter
    
    mySheet.Cells(1, 1) = Flag + 4
    
End Sub

Sub Comment(contents As String)

    Dim Flag As Long
    Dim mySheet As Worksheet
    
    Set mySheet = Worksheets(RstSheet)
    Flag = mySheet.Cells(1, 1).value - 1
    Set yp = mySheet.Cells(Flag + 2, 1)
    
    yp.Cells(1, 5) = contents
    
    mySheet.Cells(1, 1) = Flag + 2
    
End Sub

Sub TABLE(row, col, total)
                                            'Flag�� ��ȭ����. ���� �׷���
                                            '(Flag,2)���� (row,col)��ŭ total<>0�̸� ���ٴ� ���� �׸�
    Dim mySheet As Worksheet
    
    Set mySheet = Worksheets(RstSheet)
    Flag = mySheet.Cells(1, 1).value
    Set pt = mySheet.Cells(Flag, 2)
    ''
    ''
    With pt.Resize(, col).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    pt.Resize(row + total, col).HorizontalAlignment = xlRight
    
    
    With pt.Resize(, col).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    ''
    ''
    With pt.Cells(row, 1).Resize(, col).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    ''
    ''
    If total > 0 Then
    With pt.Cells(row + total, 1).Resize(, col).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    End If
    ''
    ''
End Sub

Sub tableAll(col)
    Dim mySheet As Worksheet
    Dim i As Long
    
    Set mySheet = Worksheets(RstSheet)
    Flag = mySheet.Cells(1, 1).value
    Set pt = mySheet.Cells(Flag + 1, 2)
    
    col = col + 3
    
    With pt.Resize(, col).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    
    pt.Resize(2 ^ p + 1, col).HorizontalAlignment = xlRight
    
    yp = 1
    For i = 1 To p
        yp = yp + Application.WorksheetFunction.Combin(p, i)
        With pt.Resize(yp, col).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    Next i
End Sub

Sub ANOVA(Index, intercept)
    
    Dim Flag As Long
    Dim mySheet As Worksheet
    Dim rst(), x(), Y(), tmpx()
    Dim pValue As Double, ssr As Double
    
    '���� ���
    Title3 ("�л�м�ǥ")
    
    '����Ÿ ����ֱ�
    Y = ModuleMatrix.pureY
    tmpx = ModuleMatrix.pureX
    x = ModuleMatrix.selectedX(Index, tmpx)
    
    Set mySheet = Worksheets(RstSheet)
    Flag = mySheet.Cells(1, 1).value
    Set pt = mySheet.Cells(Flag, 2)
    
    '���߱�
    TABLE 3, 6, 1
    
    'ù �� ���
    pt.Cells(1, 1) = "����"
    pt.Cells(1, 2) = "������"
    pt.Cells(1, 3) = "������"
    pt.Cells(1, 4) = "�������"
    pt.Cells(1, 5) = "F ��"
    pt.Cells(1, 6) = "����Ȯ��"
    
    'ù �� ���
    pt.Cells(2, 1) = "ȸ��"
    pt.Cells(3, 1) = "����"
    pt.Cells(4, 1) = "��"
     
    '������ �ִ� ���, ���� Ȯ�� ��� ���� ����
    If intercept <> 0 Then
    
        rst = Application.WorksheetFunction.LinEst(Y, x, 1, 1)
    
        pt.Cells(2, 2) = rst(5, 1): pt.Cells(3, 2) = rst(5, 2): pt.Cells(4, 2) = rst(5, 1) + rst(5, 2)
        pt.Cells(2, 3) = N - 1 - rst(4, 2): pt.Cells(3, 3) = rst(4, 2): pt.Cells(4, 3) = N - 1
        pt.Cells(2, 4) = rst(5, 1) / (N - 1 - rst(4, 2)): pt.Cells(3, 4) = rst(5, 2) / rst(4, 2)
        pt.Cells(2, 5) = rst(4, 1)
    
        pValue = Application.WorksheetFunction.FDist(rst(4, 1), N - 1 - rst(4, 2), rst(4, 2))
        
        '�Ʒ��� ����������� ��跮�� ���
        pt.Cells(6, 1) = "Root MSE": pt.Cells(6, 2) = Sqr(rst(5, 2) / rst(4, 2))
        pt.Cells(7, 1) = "�������": pt.Cells(7, 2) = rst(5, 1) / (rst(5, 1) + rst(5, 2))
        pt.Cells(8, 1) = "�����������"
        pt.Cells(8, 2) = 1 - (N - 1) * rst(5, 2) / (N - p - 1) / (rst(5, 2) + rst(5, 1))
               
    '������ ���� ���,���� Ȯ�� ��� ���� ����
    Else
    
        rst = Application.WorksheetFunction.LinEst(Y, x, 0, 1)
        ssr = ModuleMatrix.noIntSSR(Y, x)
    
        pt.Cells(2, 2) = ssr: pt.Cells(3, 2) = rst(5, 2): pt.Cells(4, 2) = ssr + rst(5, 2)
        pt.Cells(2, 3) = N - rst(4, 2): pt.Cells(3, 3) = rst(4, 2): pt.Cells(4, 3) = N
        pt.Cells(2, 4) = ssr / (N - rst(4, 2)): pt.Cells(3, 4) = rst(5, 2) / rst(4, 2)
        pt.Cells(2, 5) = (ssr / (N - rst(4, 2))) / (rst(5, 2) / rst(4, 2))
    
        pValue = Application.WorksheetFunction.FDist _
            ((ssr / (N - rst(4, 2))) / (rst(5, 2) / rst(4, 2)), N - rst(4, 2), rst(4, 2))
        
        '�Ʒ��� ����������� ��跮�� ���
        pt.Cells(6, 1) = "Root MSE": pt.Cells(6, 2) = Sqr(rst(5, 2) / rst(4, 2))
        pt.Cells(7, 1) = "�������": pt.Cells(7, 2) = ssr / (ssr + rst(5, 2))
        pt.Cells(8, 1) = "�����������"
        pt.Cells(8, 2) = 1 - N * rst(5, 2) / (N - p) / (ssr + rst(5, 2))
        
    End If
    
    '����Ȯ�� ���
    printPvalue pValue, pt.Cells(2, 6)
    
    '��°� ���� ����
    pt.Cells(2, 2).Resize(3, 1).NumberFormatLocal = "0.0000_ "
    pt.Cells(2, 4).Resize(2, 1).NumberFormatLocal = "0.0000_ "
    pt.Cells(2, 5).NumberFormatLocal = "0.000_ "
    pt.Cells(2, 6).NumberFormatLocal = "0.0000_ "
    
    pt.Cells(6, 2).Resize(3, 1).NumberFormatLocal = "0.0000_ "
        
    '�ڸ� ����
    mySheet.Cells(1, 1) = Flag + 8
    
End Sub

'stat function ����ϸ� �� ª�� �ڵ������ϳ� �̹� ���ؼ� ��������

Sub beta(Index, intercept)
        
    Dim Flag As Long
    Dim mySheet As Worksheet
    Dim rst(), x(), Y(), tmpx(), varName()
    Dim pValue As Double
    Dim p1 As Integer, j As Integer, K As Integer
    
    '���� ���
    Title3 ("��� ����")
    
    '����Ÿ ����ֱ�
    Y = ModuleMatrix.pureY
    tmpx = ModuleMatrix.pureX
    x = ModuleMatrix.selectedX(Index, tmpx)
    p1 = UBound(x, 2) + 1
    
    Set mySheet = Worksheets(RstSheet)
    Flag = mySheet.Cells(1, 1).value
    Set pt = mySheet.Cells(Flag, 2)
    
    '�������� �̸� ����ֱ�
    ReDim varName(p1 - 1)
    K = 0
    For j = 0 To p - 1
        If Index(j) <> 0 Then varName(K) = xlist(j): K = K + 1
    Next j
    
    'ù �� ���
    pt.Cells(1, 1) = "������"
    pt.Cells(1, 2) = "������"
    pt.Cells(1, 3) = "ǥ�ؿ���"
    pt.Cells(1, 4) = "t-��跮"
    pt.Cells(1, 5) = "����Ȯ��"
  
    '������ �ִ� ���,
    If intercept <> 0 Then
    
        '���߱�
        TABLE p1 + 2, 5, 0
        
        '��� �� ���
        rst = Application.WorksheetFunction.LinEst(Y, x, 1, 1)
        For j = 0 To p1
        
            If j = 0 Then
                pt.Cells(2, 1) = "����"
                pt.Cells(2, 2) = rst(1, p1 + 1)
                pt.Cells(2, 3) = rst(2, p1 + 1)
                pt.Cells(2, 4) = rst(1, p1 + 1) / rst(2, p1 + 1)
                pValue = Application.WorksheetFunction. _
                            TDist(Abs(rst(1, p1 + 1) / rst(2, p1 + 1)), rst(4, 2), 2)
                printPvalue pValue, pt.Cells(2, 5)
                
            Else
                pt.Cells(j + 2, 1) = varName(j - 1)
                pt.Cells(j + 2, 2) = rst(1, p1 - j + 1)
                pt.Cells(j + 2, 3) = rst(2, p1 - j + 1)
                pt.Cells(j + 2, 4) = rst(1, p1 - j + 1) / rst(2, p1 - j + 1)
            
                pValue = Application.WorksheetFunction. _
                            TDist(Abs(rst(1, p1 - j + 1) / rst(2, p1 - j + 1)), rst(4, 2), 2)
                printPvalue pValue, pt.Cells(j + 2, 5)
                
            End If
        Next j
        
        '��°� ���� ����
        pt.Cells(2, 2).Resize(p1 + 1, 2).NumberFormatLocal = "0.00000_ "
        pt.Cells(2, 4).Resize(p1 + 1, 1).NumberFormatLocal = "0.000_ "
        pt.Cells(2, 5).Resize(p1 + 1, 1).NumberFormatLocal = "0.0000_ "
    
    '������ ���� ���
    Else
    
        '���߱�
        TABLE p1 + 1, 5, 0
    
        '��� �� ���
        rst = Application.WorksheetFunction.LinEst(Y, x, 0, 1)
        For j = 1 To p1
        
            pt.Cells(j + 1, 1) = varName(j - 1)
            pt.Cells(j + 1, 2) = rst(1, p1 - j + 1)
            pt.Cells(j + 1, 3) = rst(2, p1 - j + 1)
            pt.Cells(j + 1, 4) = rst(1, p1 - j + 1) / rst(2, p1 - j + 1)
            
            pValue = Application.WorksheetFunction. _
                            TDist(Abs(rst(1, p1 - j + 1) / rst(2, p1 - j + 1)), rst(4, 2), 2)
            printPvalue pValue, pt.Cells(j + 1, 5)
            
        Next j
        
        '��°� ���� ����
        pt.Cells(2, 2).Resize(p1, 2).NumberFormatLocal = "0.00000_ "
        pt.Cells(2, 4).Resize(p1, 1).NumberFormatLocal = "0.000_ "
        pt.Cells(2, 5).Resize(p1, 1).NumberFormatLocal = "0.0000_ "
        
    End If
    If pt.Cells(4, 1).value = "" Then
        pt.Cells(5, 1) = " ȸ�͹�����"
        pt.Cells(5, 3) = "y = " & Format(pt.Cells(2, 2), "##0.00") & " + " & Format(pt.Cells(3, 2), "##0.00") & " x "
      End If
    
    
    '�ڸ� ����
    mySheet.Cells(1, 1) = Flag + p1 + 4
        
End Sub


'���� Ȯ�� ����ϴ� �Լ�
Sub printPvalue(rst, pt As Range)

    If rst > 0.0001 Then
        pt = rst
    Else: pt = "< 0.0001"
    End If
    
End Sub

Sub summaryAdd(summary, K)
        
    Dim Flag As Long
    Dim mySheet As Worksheet
    Dim j As Integer
    Dim tmpRsq As Double
       
    Set mySheet = Worksheets(RstSheet)
    mySheet.Cells(1, 1).value = mySheet.Cells(1, 1).value + 1       '�� �� ����.
    
    '���߱�
    TABLE K + 1, 8, 0
    Flag = mySheet.Cells(1, 1).value
    Set pt = mySheet.Cells(Flag, 2)
    
    'ù �� ���
    pt.Cells(1, 1) = "Step"
    pt.Cells(1, 2) = "Var Entered"
    pt.Cells(1, 3) = "Num Vars In"
    pt.Cells(1, 4) = "P R-sq"
    pt.Cells(1, 5) = "M R-sq"
    pt.Cells(1, 6) = "C_p"
    pt.Cells(1, 7) = "F ��"
    pt.Cells(1, 8) = "����Ȯ��"
        
    '��� ��跮�� ���
    tmpRsq = 0
    For j = 0 To K - 1
        pt.Cells(j + 2, 1) = j + 1
        pt.Cells(j + 2, 2) = xlist(summary(j, 0))
        pt.Cells(j + 2, 3) = j + 1
        pt.Cells(j + 2, 4) = summary(j, 7) - tmpRsq
            tmpRsq = summary(j, 7)
        pt.Cells(j + 2, 5) = summary(j, 7)
        pt.Cells(j + 2, 6) = summary(j, 9)
        pt.Cells(j + 2, 7) = summary(j, 5)
        printPvalue summary(j, 6), pt.Cells(j + 2, 8)
    Next j
    
    '��°� ���� ����
        pt.Cells(2, 4).Resize(K, 5).NumberFormatLocal = "0.0000_ "
        pt.Cells(2, 7).Resize(K, 1).NumberFormatLocal = "0.000_ "
    
    '�ڸ� ����
    mySheet.Cells(1, 1) = Flag + K + 2
    
End Sub

Sub summaryRm(summary, K)
        
    Dim Flag As Long
    Dim mySheet As Worksheet
    Dim j As Integer
    Dim tmpRsq As Double
       
    Set mySheet = Worksheets(RstSheet)
    mySheet.Cells(1, 1).value = mySheet.Cells(1, 1).value + 1       '�� �� ����.
    
    If K = 0 Then GoTo LastLine
    
    '���߱�
    TABLE K + 1, 8, 0
    Flag = mySheet.Cells(1, 1).value
    Set pt = mySheet.Cells(Flag, 2)
    
    'ù �� ���
    pt.Cells(1, 1) = "Step"
    pt.Cells(1, 2) = "Var Removed"
    pt.Cells(1, 3) = "Num Vars In"
    pt.Cells(1, 4) = "P R-sq"
    pt.Cells(1, 5) = "M R-sq"
    pt.Cells(1, 6) = "C_p"
    pt.Cells(1, 7) = "F ��"
    pt.Cells(1, 8) = "����Ȯ��"
        
    '��� ��跮�� ���
    tmpRsq = summary(0, 7)
    For j = 0 To K - 1
        pt.Cells(j + 2, 1) = j + 1
        pt.Cells(j + 2, 2) = xlist(summary(j, 0))
        pt.Cells(j + 2, 3) = p - j - 1
        pt.Cells(j + 2, 4) = tmpRsq - summary(j, 7)
            tmpRsq = summary(j, 7)
        pt.Cells(j + 2, 5) = summary(j, 7)
        pt.Cells(j + 2, 6) = summary(j, 9)
        pt.Cells(j + 2, 7) = summary(j, 5)
        printPvalue summary(j, 6), pt.Cells(j + 2, 8)
    Next j
    
    '��°� ���� ����
        pt.Cells(2, 4).Resize(K, 5).NumberFormatLocal = "0.0000_ "
        pt.Cells(2, 7).Resize(K, 1).NumberFormatLocal = "0.000_ "
        
    '�ڸ� ����
    mySheet.Cells(1, 1) = Flag + K + 2
    
LastLine:
End Sub

Sub summaryStep(summary, K)
        
    Dim Flag As Long
    Dim mySheet As Worksheet
    Dim j As Integer, cnt As Integer, numInModel As Integer
    Dim tmpRsq As Double
       
    Set mySheet = Worksheets(RstSheet)
    mySheet.Cells(1, 1).value = mySheet.Cells(1, 1).value + 1       '�� �� ����.
    
    '����� �� ���� - �����̳� ���Ű� �Ͼ Ƚ��, summary(,5)<>0 �ΰ͵�
    cnt = 0
    For j = 0 To 2 * p - 2
        If summary(j, 11) <> 0 Then cnt = cnt + 1
    Next j
    
    '���߱�
    TABLE cnt + 1, 8, 0
    Flag = mySheet.Cells(1, 1).value
    Set pt = mySheet.Cells(Flag, 2)
    
    'ù �� ���
    pt.Cells(1, 1) = "Step"
    pt.Cells(1, 2) = "Var Entered"
    pt.Cells(1, 2) = "Var Changed"
    pt.Cells(1, 3) = "Num Vars In"
    pt.Cells(1, 4) = "P R-sq"
    pt.Cells(1, 5) = "M R-sq"
    pt.Cells(1, 6) = "C_p"
    pt.Cells(1, 7) = "F ��"
    pt.Cells(1, 8) = "����Ȯ��"
    
    
    
    '��� ��跮�� ���
    j = 0: numInModel = 0: tmpRsq = 0
    Do While j < cnt
    For K = 0 To UBound(summary, 1)
    
        numInModel = numInModel + summary(K, 11)
      
        Select Case summary(K, 11)
        
        Case 1
        pt.Cells(j + 2, 1) = j + 1
        pt.Cells(j + 2, 2) = xlist(summary(K, 0)) & " (�߰�)"
        pt.Cells(j + 2, 2) = msgWord(pt.Cells(j + 2, 2))
        pt.Cells(j + 2, 3) = numInModel
        pt.Cells(j + 2, 4) = summary(K, 7) - tmpRsq
            tmpRsq = summary(K, 7)
        pt.Cells(j + 2, 5) = summary(K, 7)
        pt.Cells(j + 2, 6) = summary(K, 9)
        pt.Cells(j + 2, 7) = summary(K, 5)
        printPvalue summary(K, 6), pt.Cells(j + 2, 8)
        j = j + 1
        
        Case -1
        pt.Cells(j + 2, 1) = j + 1
        pt.Cells(j + 2, 2) = xlist(summary(K, 0)) & " (����)"
        pt.Cells(j + 2, 2) = msgWord(pt.Cells(j + 2, 2))
        pt.Cells(j + 2, 3) = numInModel
        pt.Cells(j + 2, 4) = tmpRsq - summary(K, 7)
            tmpRsq = summary(K, 7)
        pt.Cells(j + 2, 5) = summary(K, 7)
        pt.Cells(j + 2, 6) = summary(K, 9)
        pt.Cells(j + 2, 7) = summary(K, 5)
        printPvalue summary(K, 6), pt.Cells(j + 2, 8)
        j = j + 1
        
        Case 0
                
        End Select
        
    Next K
    Loop
    
    '��°� ���� ����
        pt.Cells(2, 4).Resize(K, 5).NumberFormatLocal = "0.0000_ "
        pt.Cells(2, 7).Resize(K, 1).NumberFormatLocal = "0.000_ "
        
    '�ڸ� ����
    mySheet.Cells(1, 1) = Flag + K + 2
    
End Sub

Function msgWord(word) As String                    '������ word �� string

    If Len(word) > 9 Then
        msgWord = Mid(word, 1, 6) & vbLf & Mid(word, 7)
    Else
        msgWord = word
    End If
    
End Function

Sub All(rst)
    Dim num As Long
    Dim mySheet As Worksheet
        
    Set mySheet = Worksheets(RstSheet)
    Flag = mySheet.Cells(1, 1).value
    Set pt = mySheet.Cells(Flag + 1, 2)
    
    num = 2 ^ p
    col = UBound(rst, 2)
    
    mySheet.Range(pt.Cells(1, 1), pt.Cells(num, col + 1)) = rst
    mySheet.Activate
    
    Range(pt.Cells(2, 1), pt.Cells(num, col + 1)).Select
    
    Selection.Sort Key1:=Range("C3"), Order1:=xlAscending, Key2:=Range("D3") _
        , Order2:=xlDescending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom
        
    '��°� ���� ����
        pt.Cells(2, 3).Resize(num, col + 1).NumberFormatLocal = "0.0000_ "
        'Columns("E:E").EntireColumn.AutoFit
        
    '�ڸ� ����
    mySheet.Cells(1, 1) = Flag + num + 2
End Sub


''�ӽý�Ʈ �����
Sub MakeTmpSheet(SheetName As String)
    
    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
        If ws.Name = SheetName Then Exit Sub
        
    Next ws
    
    
    Worksheets.add.Name = SheetName
    Worksheets(SheetName).Cells.Select
    Selection.NumberFormatLocal = "0.0_ "
    Worksheets(SheetName).Visible = False
        
End Sub
