VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDisc 
   OleObjectBlob   =   "frmDisc.frx":0000
   Caption         =   "������"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   StartUpPosition =   1  '������ ���
   TypeInfoVer     =   463
End
Attribute VB_Name = "frmDisc"
Attribute VB_Base = "0{DB2EA74A-6CD4-49B3-BE91-DA94CC386AFA}{DF18EBD5-2FB7-419C-8DF9-6437AB47BED0}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False


Private TempWorksheet As Worksheet
Private FlagTmp As Boolean

Private Sub back_k_Click()
    If Me.back_k.Value = True Then Me.TextBox3.Enabled = True
    If Me.back_k.Value = False Then Me.TextBox3.Enabled = False
End Sub

Private Sub front_k_Click()
    If Me.front_k.Value = True Then Me.TextBox2.Enabled = True
    If Me.front_k.Value = False Then Me.TextBox2.Enabled = False
End Sub



Private Sub �ű��ư_Click()
    MoveBtwnListBox Me, "ListBox1", "ListBox2"
End Sub

Private Sub ���ڸ���ư_Click()
    MoveBtwnListBox Me, "ListBox2", "ListBox1"
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    MoveBtwnListBox Me, "ListBox1", "ListBox2"
End Sub

Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    MoveBtwnListBox Me, "ListBox2", "ListBox1"
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
   
 '  For i = 1 To TempSheet.UsedRange.Columns.Count
 '   rngFirst.Offset(i, 1) = myArray(i - 1)
 ' Next i
  
End Sub

Private Sub OptionButton2_Click()
   ' OptBtn12Click Me, False
End Sub

Private Sub �߰�����_Click()
myToggle �߰�����, �߰���跮, �߰�����.Name
End Sub

Private Sub ��ҹ�ư_Click()
      Unload Me
End Sub

Private Sub Ȯ�ι�ư_Click()                                    ''''"_������м����_"
   
    Dim i, cnt As Integer: Dim ErrSign As Boolean
    Dim RnArray(), PrintStartCell As Range
    Dim VarName(), ErrString As String
    Dim myResultSheet As Worksheet
    cnt = Me.ListBox2.ListCount
    If cnt = 0 Then
        MsgBox "�м������� �����ϴ�.", vbExclamation, "HIST"
        Exit Sub
    Else
        ReDim RnArray(1 To cnt): ReDim VarName(1 To cnt)
        SelectMultiRange Me, RnArray, VarName
    End If
        
    For i = 1 To cnt
        If PublicModule.FindingRangeError(RnArray(i)) = True Then
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
    
    PublicModule.SettingStatusBar True, "���� ��� �м� ���Դϴ�."
    Application.ScreenUpdating = False
    PublicModule.OpenOutSheet "_���м����_", True
    Set myResultSheet = Worksheets("_���м����_")
    'myResultSheet.Unprotect "prophet"
    
    '''
    '''
    '''
    rstSheet = "_���м����_"
    ModulePrint.Title1 "������м� ���"
    activePt = Worksheets(rstSheet).Cells(1, 1).Value
    
    For i = 1 To cnt
        With Worksheets("_���м����_").Range("a1")
            If Left(.Value, 1) = "$" Then
                Set PrintStartCell = myResultSheet.Range(myResultSheet.Range("a1").Value).Offset(1, 0)
            Else
                tmpStr = "$A$" & myResultSheet.Range("a1").Value
                Set PrintStartCell = myResultSheet.Range(tmpStr).Offset(1, 0)
            End If
        End With
        desresult PrintStartCell, RnArray(i), VarName(i), myResultSheet
    Next i
    
    '=== ����
    '    For i = 1 To cnt
     '   With Worksheets("_���м����_").Range("a1")
     '       If Left(.Value, 1) = "$" Then
     '           Set PrintStartCell = myResultSheet.Range(myResultSheet.Range("a1").Value).Offset(1, 0)
     '       Else
    '            tmpStr = "$A$" & myResultSheet.Range("a1").Value
     '           Set PrintStartCell = myResultSheet.Range(tmpStr).Offset(1, 0)
     '       End If
    '    End With
    '    desresult PrintStartCell, RnArray(i), VarName(i), myResultSheet
   ' Next i
    
   ' ===================
    
    
    
    
    If Left(Worksheets(rstSheet).Range("a1"), 1) = "$" Then
        Worksheets(rstSheet).Cells(1, 1) = right(Worksheets(rstSheet).Cells(1, 1).Value, Len(Worksheets(rstSheet).Cells(1, 1).Value) - 3)
    End If
    

                                    
    Application.ScreenUpdating = True
    PublicModule.SettingStatusBar False
    Unload Me
    Worksheets("_���м����_").Activate
    Worksheets("_���м����_").Columns(10).Delete              ' J�� ����
      
    

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
    
    Worksheets(rstSheet).Activate
    Worksheets(rstSheet).Cells(activePt + 5, 1).Select
    Worksheets(rstSheet).Cells(activePt + 5, 1).Activate
                            '��� �м��� ���۵Ǵ� �κ��� �����ָ� ��ģ��.
End Sub

Private Sub CommandButton1_Click()
  ShellExecute 0, "open", "hh.exe", ThisWorkbook.Path + "\HIST%202013.chm::/������.htm", "", 1
End Sub

Function PrintStat(myControl, TargetCell, StatName, StatValue, chknum) As Integer
    
    If myControl.Value = True Then
        Set TargetCell = TargetCell.Offset(1, 0)
        With TargetCell
            .ColumnWidth = 13
            .HorizontalAlignment = xlLeft
        End With
        TargetCell.Value = StatName
        If IsNumeric(StatValue) = True Then
            TargetCell.Offset(0, 1).Value = Format(StatValue, "0.0000")
        Else: TargetCell.Offset(0, 1).Value = "#DIV/0!"
        End If
        PrintStat = chknum + 1: Exit Function
    End If
    PrintStat = chknum
    
End Function

Function calMode(ra) As Long
     On Error Resume Next
     calMode = Application.mode(ra)
     If Err.number <> 0 Then
          calMode = ra.Cells(1).Value
     End If
End Function

Sub desresult(temp, tmp1, lv, dResultsheet)
    Dim oc As Integer
    Dim yyj, qq, ack As Range
    Dim M1 As String
    Dim M2, xp, yp, yp1 As Double
    Dim title As Shape
    Dim vv()
    yp1 = temp.Top
    Dim chknum, an, bn As Integer
    
    chknum = 0
    ReDim vv(1 To 1)
    Set yyj = temp.Offset(3, 1)
    Set temp = temp.Offset(3, 1)
    yp = temp.Top
    xp = temp.Left
    Set temp = temp.Offset(1, 0)
    Set ack = temp
    '''
    chknum = PrintStat(mean, temp, mean.Caption, Application.Average(tmp1), chknum)
    chknum = PrintStat(median, temp, median.Caption, Application.median(tmp1), chknum)
    chknum = PrintStat(mode, temp, mode.Caption, calMode(tmp1), chknum)
    chknum = PrintStat(trimmean, temp, trimmean.Caption, Application.trimmean((tmp1), 0.05), chknum)
    
    If chknum = 0 Then
       Set temp = yyj
    Else: MakeSubTitle dResultsheet, xp, yp, "�߽ɿ� ���� ����"
         Set temp = yyj.Offset(0, 3)
         yp = temp.Top
         xp = temp.Left
    End If
          
    dResultsheet.Columns(5).ColumnWidth = 13
    Set temp = temp.Offset(1, 0)
    
    an = 0
    
    '''
    an = PrintStat(Variance, temp, Variance.Caption, Application.var(tmp1), an)
    an = PrintStat(Std, temp, Std.Caption, Application.StDev(tmp1), an)
    If Application.Average(tmp1) <> 0 And tmp1.count > 1 Then
       an = PrintStat(Cv, temp, Cv.Caption, Application.StDev(tmp1) / Application.Average(tmp1), an)
    Else
       an = PrintStat(Cv, temp, Cv.Caption, "����", an)
    End If
    an = PrintStat(IQR, temp, IQR.Caption, Application.Quartile(tmp1, 3) - Application.Quartile(tmp1, 1), an)
    an = PrintStat(�������1, temp, �������1.Caption, Application.Quartile(tmp1, 1), an)
    an = PrintStat(�������3, temp, �������3.Caption, Application.Quartile(tmp1, 3), an)
    
    Dim yoon As Range
    Set yoon = temp
    If an = 0 Then
       If chknum <> 0 Then                                                  'an 0 ���ٰ� chknum ����
            With yoon.Offset(1, -3).Resize(, 2).Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
        End If
    Else
          MakeSubTitle dResultsheet, xp, yp, "������ ���� ����"
          Set yoon = temp.CurrentRegion
          Set qq = yoon.Cells(1, 2)
          Set qq = qq.Offset(0, -4).Resize(, 5)
          With qq.Borders(xlEdgeTop)
               .LineStyle = xlContinuous
               .Weight = xlMedium
               .ColorIndex = xlAutomatic
          End With
          Dim toon As Range
          
          Set toon = yoon.Offset(-2, 3) '��ġ�缳��
          yp = toon.Top
          xp = toon.Left
          
          Dim joon As Range
          
          Set joon = ack
          Set joon = joon.Offset(0, 6)
          
          yp = joon.Top
          xp = joon.Left
          
          
     End If
     
    'temp  ���ľ���
        
    dResultsheet.Columns(8).ColumnWidth = 13
    If an <> 0 Then
    Set temp = joon.Offset(0, 0)
    Else: Set temp = yoon.Offset(1, 0)
    End If
    
    
    bn = 0
    
    '''
    bn = PrintStat(max, temp, max.Caption, Application.max(tmp1), bn)
    bn = PrintStat(min, temp, min.Caption, Application.min(tmp1), bn)
    bn = PrintStat(�ֵ�, temp, �ֵ�.Caption, Application.Skew(tmp1), bn)
    bn = PrintStat(Kurtosis, temp, Kurtosis.Caption, Application.Kurt(tmp1), bn)
    bn = PrintStat(num, temp, num.Caption, Application.count(tmp1), bn)
    bn = PrintStat(����, temp, ����.Caption, Application.max(tmp1) - Application.min(tmp1), bn)
    bn = PrintStat(summ, temp, summ.Caption, Application.sum(tmp1), bn)
    If tmp1.count > 1 Then
        bn = PrintStat(SE, temp, SE.Caption, Application.StDev(tmp1) / Sqr(tmp1.count), bn)
    Else
        bn = PrintStat(SE, temp, SE.Caption, "����", bn)
    End If
    bn = PrintStat(front_k, temp, front_k, Application.Small(tmp1, Me.TextBox2.Value), bn)
    bn = PrintStat(back_k, temp, back_k, Application.Large(tmp1, Me.TextBox3.Value), bn)
    
   ' Dim yoon As Range
    'Set yoon = temp
    If bn = 0 Then
       If chknum <> 0 Then
            With joon.Offset(1, -3).Resize(, 2).Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
        End If
    Else
        If an <> 0 Then
          MakeSubTitle dResultsheet, xp, yp - 12, "��Ÿ ��跮"
          Set joon = temp.CurrentRegion
          Set qq = joon.Cells(1, 2)                         ' �� ��ġ
          Set qq = qq.Offset(0, -7).Resize(, 8)             ' ������ �������� ���ڿ� �ٽ� ����ĭ�ڷ�
          With qq.Borders(xlEdgeTop)
               .LineStyle = xlContinuous
               .Weight = xlMedium
               .ColorIndex = 2
          End With
        
        Else: MakeSubTitle dResultsheet, xp, yp - 12, "��Ÿ ��跮"
          Set joon = temp.CurrentRegion
          Set qq = joon.Cells(1, 2)
          Set qq = qq.Offset(0, -4).Resize(, 5)
          With qq.Borders(xlEdgeTop)
               .LineStyle = xlContinuous
               .Weight = xlMedium
               .ColorIndex = 2
          End With
         End If
          
     End If

    If chknum <> 0 Or an <> 0 Or bn <> 0 Then
       
       Set title = dResultsheet.Shapes.AddShape(msoShapeRectangle, 60.75, yp1, 150, 22#)  '3.75 + 20, yp1 + 2,150, 22#
       title.Fill.ForeColor.SchemeColor = 9
       title.Line.Weight = 1
       'title.Shadow.Type = msoShadow1
       title.TextFrame.Characters.Text = lv
       With title.TextFrame.Characters.Font
            .Name = "���� ���"
           .Size = 11
           .ColorIndex = xlAutomatic
       End With
       title.TextFrame.HorizontalAlignment = xlCenter
    End If
    
    
    
    Dim cnt1, cnt2, cnt3 As Integer
    Dim ct1, ct2, ct3 As Integer
    
    '' chk ������ cnt1 �� ��������
    
    ct1 = yoon.Rows.count          '���
    ct2 = yoon.Offset(0, -4).CurrentRegion.Rows.count  '
    ct3 = toon.Rows.count
    
    If chknum = 0 Then
       cnt1 = yoon.Rows.count
       cnt2 = 0
    Else: cnt1 = yoon.Rows.count
          cnt2 = yoon.Offset(0, -4).CurrentRegion.Rows.count
    End If
    
    
    
    Dim diff As Integer
    If cnt1 < cnt2 Then
       diff = cnt2 - cnt1
       Set temp = yoon.Rows(yoon.Rows.count + diff).Cells(1)
       Set temp = temp.Offset(1, -4)
    ElseIf chknum = 0 Then
       Set temp = yoon.Rows(yoon.Rows.count).Cells(1)
       Set qq = temp.Resize(, 2)
        With qq.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
    Else:
         Set temp = joon.Rows(yoon.Rows.count).Cells(1)
         Set temp = temp.Offset(1, -4)
    End If
    
    
    
    'If an = 0 Then
    '  With temp.Offset(0, 1).Resize(, 2).Borders(xlEdgeTop)
    '    .LineStyle = xlContinuous
    '    .Weight = xlMedium
    '    .ColorIndex = 14
   ' End With
    
  '  ElseIf chknum <> 0 Then                                         ' ��� ���� ���� �׸���
           'With temp.Offset(0, -2).Resize(, 8).Borders(xlEdgeTop)
              '  .LineStyle = xlContinuous
               ' .Weight = xlMedium
               ' .ColorIndex = 14
      '     End With
   ' End If
  
   ' Set temp = temp.Offset(-8, 0)
    
    
    
    dResultsheet.Range("a1").Value = temp.Address
    ad = temp.Address
    dResultsheet.Range("a1").Value = ad
   
 
    
End Sub

Sub MakeSubTitle(rsheet, xpoint, ypoint, t)
    Dim title As Shape
    Set title = rsheet.Shapes.AddShape(msoShapeRectangle, xpoint, ypoint, 135, 22)
    With title
        .Fill.ForeColor.SchemeColor = 1
        .Fill.Visible = msoTrue
        .Fill.Solid
        ' .Shadow.Type = msoShadow1
    End With
    title.TextFrame.Characters.Text = t ' &"�� ���� ������跮"
    With title.TextFrame.Characters.Font
        .Size = 12
        .ColorIndex = xlAutomatic
    End With
    title.TextFrame.HorizontalAlignment = xlCenter

End Sub

Private Sub UserForm_Terminate()
      Unload Me
End Sub
