VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmspcimr 
   OleObjectBlob   =   "frmspcimr.frx":0000
   Caption         =   "I-MR"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6720
   StartUpPosition =   1  '������ ���
   TypeInfoVer     =   37
End
Attribute VB_Name = "frmspcimr"
Attribute VB_Base = "0{6A5539AB-2A87-48A3-AE7F-413816A40A04}{9073A2CA-DFFE-40E4-8AA8-5748DAE6B773}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False


Private Declare Function ShellExecute _
 Lib "shell32.dll" _
 Alias "ShellExecuteA" ( _
 ByVal hwnd As Long, _
 ByVal lpOperation As String, _
 ByVal lpFile As String, _
 ByVal lpParameters As String, _
 ByVal lpDirectory As String, _
 ByVal nShowCmd As Long) _
 As Long
Private Sub btnCancel_Click()
Unload Me
End Sub

Private Sub HlpBtn_Click()
 ShellExecute 0, "open", "hh.exe", ThisWorkbook.Path + "\hist_help_V.2.5.1.chm::/I-MR1.htm", "", 1
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim i As Integer
    i = 0
    If Me.ListBox2.ListCount = 0 Then
        Do While i <= Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(i) = True Then
               Me.ListBox2.AddItem Me.ListBox1.List(i)
               Me.ListBox1.RemoveItem (i)
               Me.CB1.Visible = False
               Me.CB2.Visible = True
               Exit Sub
            End If
            i = i + 1
        Loop
    End If
End Sub

Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.ListBox2.ListCount <> 0 Then
        Me.ListBox1.AddItem Me.ListBox2.List(0)
        Me.ListBox2.RemoveItem (0)
        Me.CB1.Visible = True
        Me.CB2.Visible = False
    End If
End Sub
Private Sub CB1_Click()
Dim i As Integer
    i = 0
    If Me.ListBox2.ListCount = 0 Then
        Do While i <= Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(i) = True Then
               Me.ListBox2.AddItem Me.ListBox1.List(i)
               Me.ListBox1.RemoveItem (i)
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
        Me.ListBox1.AddItem ListBox2.List(0)
        Me.ListBox2.RemoveItem (0)
        Me.CB1.Visible = True
        Me.CB2.Visible = False
    End If
End Sub
Private Sub OptionButton1_Click()
   Dim myRange As Range
   Dim myArray()
   Dim arrName As Variant
   Dim TempSheet As Worksheet
   Set TempSheet = ActiveCell.Worksheet
   
    ReDim arrName(TempSheet.UsedRange.Columns.Count)
' Reading Data
    For i = 1 To TempSheet.UsedRange.Columns.Count
        arrName(i) = TempSheet.Cells(1, i)
    Next i
   
   Me.ListBox1.Clear
'-------------
  'Set myRange = Cells.CurrentRegion.Rows(1)
   'cnt = myRange.Cells.Count
   'ReDim myArray(cnt - 1)
  ' For i = 1 To cnt
  '   myArray(i - 1) = myRange.Cells(i)
  ' Next i
   'Me.ListBox1.List() = myArray
'-----------
    ReDim myArray(TempSheet.UsedRange.Columns.Count - 1)
    a = 0
   For i = 1 To TempSheet.UsedRange.Columns.Count
   If arrName(i) <> "" Then                     '��ĭ����
   myArray(a) = arrName(i)
   a = a + 1
   
   Else:
   End If
   Next i
   
   
   
   Me.ListBox1.List() = myArray
   
 '  For i = 1 To TempSheet.UsedRange.Columns.Count
 '   rngFirst.Offset(i, 1) = myArray(i - 1)
 ' Next i
  
   


End Sub
Sub CleanCharts()
    Dim chrt As Picture
    On Error Resume Next
    For Each chrt In ActiveSheet.Pictures
        chrt.Delete
    Next chrt
End Sub

Private Sub ToggleButton1_Click()
                           '�ѱ� ������ ������,�ŷڱ���,�븳���� 3���ϱ�
    Dim dataRange As Range
    Dim i As Integer
    Dim activePt As Long                                '��� �м��� ���۵Ǵ� �κ��� �����ֱ� ����
    Dim rng As Range
      Dim b As String
    Dim no2 As Integer
    
    '''
    '''������ �������� �ʾ��� ���
    '''
    If Me.ListBox2.ListCount = 0 Then
        MsgBox "������ ������ �ֽñ� �ٶ��ϴ�.", vbExclamation, "HIST"
        Exit Sub
    End If
    
    '''
    '''public ���� ���� xlist, DataSheet, RstSheet, m, k1, n
    '''
    xlist = Me.ListBox2.List(0)
    DataSheet = ActiveSheet.Name                        'DataSheet : Data�� �ִ� Sheet �̸�
    RstSheet = "_���м����_"                       'RstSheet  : ����� �����ִ� Sheet �̸�
    
    
    
    '������ �Է�
'On Error GoTo Err_delete
Dim val3535 As Long '�ʱ���ġ ������ ����'
Dim s3535 As Worksheet
val3535 = 2
    For Each s3535 In ActiveWorkbook.Sheets
        If s3535.Name = RstSheet Then
val3535 = Sheets(RstSheet).Cells(1, 1).Value
End If
Next s3535  '��Ʈ�� �̹������� ��� ��ġ �������ϰ�, ������ 2�� �����Ѵ�.


    
    Set dataRange = ActiveSheet.Cells.CurrentRegion
    m = dataRange.Columns.Count                         'm         : dataSheet�� �ִ� ���� ����
    
    tmp = 0
    For i = 1 To m
        If xlist = ActiveSheet.Cells(1, i) Then
            k1 = i  'k1                                 : k1 : ���õ� ������ ���° ���� �ִ���
            tmp = tmp + 1
        End If
    Next i
    
    N = ActiveSheet.Cells(1, k1).End(xlDown).Row - 1    'n         : ���õ� ������ ����Ÿ ����
    
   ' rng = Range(Cells(2, k1), Cells(N + 1, k1))
    
    
    

    '''
    ''' �������� ���� ��� - ������ ���� �ִ� ������ �ԷµǹǷ� ����ó���Ѵ�.
    '''
    If tmp > 1 Then
        MsgBox xlist & "�� ���� �������� �ֽ��ϴ�. " & vbCrLf & "�������� �ٲ��ֽñ� �ٶ��ϴ�.", vbExclamation, "HIST"
        Exit Sub
    End If
    
    
      
    Rinterface.StartRServer
    'Rinterface.PutArray "arraytest", Range(rng)
    Rinterface.RRun "install.packages (" & Chr(34) & "qcc" & Chr(34) & ")"
    Rinterface.RRun "require (qcc)"
    
    
      Rinterface.PutArray "arraytest", Range(Cells(2, k1), Cells(N + 1, k1))
      
      'Rinterface.RRun "x1 <- matrix(data= arraytest, nrow = 20, ncol=5, byrow = FALSE, dimnames = NULL)"
      
      
      
      Application.ScreenUpdating = False
    Dim stname As String
    'Dim lastCol, lastRow As Integer
    
    stname = "������"
    Module2.OpenOutSheet stname, True
    Worksheets(stname).Activate
   
   Dim stn As Integer
   stn = Sheets(stname).Cells(1, 1).Value
    ActiveSheet.Cells(stn + 1, 1).Value = "������"
     ActiveSheet.Cells(stn + 1, 1).Font.Bold = True
     ActiveSheet.Cells(stn + 1, 1).Interior.Color = RGB(220, 238, 130)
     ActiveSheet.Cells(stn + 1, 1).Cells.ColumnWidth = 20

    ActiveSheet.Cells(stn + 2, 1).Value = xlist
    Rinterface.GetArray "arraytest", Range(Cells(stn + 3, 1), Cells(stn + 3, 1))
    

    'lastCol = ActiveCell.Worksheet.UsedRange.Columns.Count
    'lastRow = ActiveCell.Worksheet.UsedRange.Rows.Count
    
    'MsgBox " col:" & lastCol & " row: " & lastRow
    
    



 ActiveSheet.Cells(stn + 1, 3).Value = "������ �׷���"
      ActiveSheet.Cells(stn + 1, 3).Font.Bold = True
ActiveSheet.Cells(stn + 1, 3).Interior.Color = RGB(220, 238, 130)

      
      
      
      
      
      
      
no2 = Me.TextBox2.Value
'Rinterface.RRun "qcc(x1, type= " & Chr(34) & "xbar.one" & Chr(34) & ", nsigmas=3)"
b = " xbar <- qcc(arraytest, type= " & Chr(34) & "xbar.one" & Chr(34) & ", nsigmas=" & no2 & ", title= " & Chr(34) & "I ������" & Chr(34) & ")  "
'Rinterface.RRun "qcc.options(" & Chr(34) & "beyond.limits" & Chr(34) & " =list(pch=15, col= " & Chr(34) & "orangered" & Chr(34) & "))"
'Rinterface.RRun " plot(xbar, title = " & Chr(34) & "I ������" & Chr(34) & ")"

Rinterface.RRun b
Rinterface.InsertCurrentRPlot Range(Cells(stn + 3, 3), Cells(stn + 3, 3)), widthrescale:=0.6, heightrescale:=0.6, closergraph:=True

Rinterface.RRun "imr <- matrix(cbind(arraytest[1:length(arraytest)-1],arraytest[2:length(arraytest)]), ncol=2)"
Rinterface.RRun "win.graph()"


Rinterface.RRun " mr <- qcc(imr, type=  " & Chr(34) & "R" & Chr(34) & ", title= " & Chr(34) & "MR ������" & Chr(34) & ")"
'Rinterface.RRun "qcc.options(" & Chr(34) & "beyond.limits" & Chr(34) & " =list(pch=15, col= " & Chr(34) & "orangered" & Chr(34) & "))"
'Rinterface.RRun " plot(mr, title = " & Chr(34) & "MR ������" & Chr(34) & ")"
Rinterface.InsertCurrentRPlot Range(Cells(stn + 3, 8), Cells(stn + 3, 8)), widthrescale:=0.6, heightrescale:=0.6, closergraph:=True


'i������ ������ؼ�

Rinterface.RRun "xl<- limits.xbar.one(xbar$center, xbar$std.dev, xbar$sizes,3)"
Rinterface.RRun "xs<- stats.xbar.one(xbar$data, xbar$sizes)"
Rinterface.RRun "xss <- as.data.frame(xs)"
Rinterface.RRun "xd <- data.frame(xl,xss)"

Rinterface.RRun "result <- row.names(xd[which(xd$UCL < xd$statistics), ])"
Rinterface.RRun "result2 <-  row.names(xd[which(xd$LCL > xd$statistics), ])"
Rinterface.RRun "result3 <- t(result) "
Rinterface.RRun "result4 <- t(result2)"

Rinterface.GetArray "result3", Range(Cells(stn + 32, 4), Cells(stn + 32, 4))
Rinterface.GetArray "result4", Range(Cells(stn + 34, 4), Cells(stn + 34, 4))

Range(Cells(stn + 30, 3), Cells(stn + 30, 3)).Value = "I������ ����ؼ�"
Range(Cells(stn + 30, 3), Cells(stn + 30, 3)).Cells.ColumnWidth = 15
Range(Cells(stn + 30, 3), Cells(stn + 30, 3)).Font.Bold = True
Range(Cells(stn + 30, 3), Cells(stn + 30, 3)).Interior.Color = RGB(220, 238, 130)

Range(Cells(stn + 32, 3), Cells(stn + 32, 3)).Value = "I�������Ѽ��� ����� �κб�:"
Range(Cells(stn + 32, 3), Cells(stn + 32, 3)).Cells.ColumnWidth = 28
Range(Cells(stn + 32, 3), Cells(stn + 32, 3)).Font.Color = vbBlack
Range(Cells(stn + 32, 3), Cells(stn + 32, 3)).Font.Bold = True


Range(Cells(stn + 34, 3), Cells(stn + 34, 3)).Value = "I�������Ѽ��� ����� �κб�:"
Range(Cells(stn + 34, 3), Cells(stn + 34, 3)).Cells.ColumnWidth = 28
Range(Cells(stn + 34, 3), Cells(stn + 34, 3)).Font.Color = vbBlack
Range(Cells(stn + 34, 3), Cells(stn + 34, 3)).Font.Bold = True


If Range(Cells(stn + 32, 4), Cells(stn + 32, 4)).Value & Range(Cells(stn + 34, 4), Cells(stn + 34, 4)).Value = "" Then
Range(Cells(stn + 36, 4), Cells(stn + 36, 4)).Value = "������ �������¿� �ִ� ������ ������ �� �ֽ��ϴ�."
ElseIf Range(Cells(stn + 32, 4), Cells(stn + 32, 4)).Value <> "" Then
Range(Cells(stn + 33, 4), Cells(stn + 33, 4)).Value = "��° �κб��� '�������Ѽ�'�� ������ϴ�. ���� ������ �̻������ �ִ� ������ �����˴ϴ�."
ElseIf Range(Cells(stn + 34, 4), Cells(stn + 34, 4)).Value <> "" Then
Range(Cells(stn + 35, 4), Cells(stn + 35, 4)).Value = "��° �κб��� '�������Ѽ�'�� ������ϴ�. ���� ������ �̻������ �ִ� ������ �����˴ϴ�."
End If


Range(Cells(stn + 37, 3), Cells(stn + 37, 3)).Value = "MR������ ����ؼ�"
Range(Cells(stn + 37, 3), Cells(stn + 37, 3)).Cells.ColumnWidth = 15
Range(Cells(stn + 37, 3), Cells(stn + 37, 3)).Font.Bold = True
Range(Cells(stn + 37, 3), Cells(stn + 37, 3)).Interior.Color = RGB(220, 238, 130)



Range(Cells(stn + 39, 3), Cells(stn + 39, 3)).Value = "MR�������Ѽ��� ����� �κб�:"
Range(Cells(stn + 39, 3), Cells(stn + 39, 3)).Cells.ColumnWidth = 28
Range(Cells(stn + 39, 3), Cells(stn + 39, 3)).Font.Color = vbBlack
Range(Cells(stn + 39, 3), Cells(stn + 39, 3)).Font.Bold = True



Range(Cells(stn + 41, 3), Cells(stn + 41, 3)).Value = "MR�������Ѽ��� ����� �κб�:"
Range(Cells(stn + 41, 3), Cells(stn + 41, 3)).Cells.ColumnWidth = 28
Range(Cells(stn + 41, 3), Cells(stn + 41, 3)).Font.Color = vbBlack
Range(Cells(stn + 41, 3), Cells(stn + 41, 3)).Font.Bold = True



'r������ ������ؼ�

Rinterface.RRun "rl<- limits.R(mr$center, mr$std.dev, mr$sizes,3)"
Rinterface.RRun "rs<- stats.R(mr$data, mr$sizes)"
Rinterface.RRun "rss <- as.data.frame(rs)"
Rinterface.RRun "rd <- data.frame(rl,rss)"

Rinterface.RRun "rresult <- row.names(rd[which(rd$UCL < rd$statistics), ])"
Rinterface.RRun "rresult2 <-  row.names(rd[which(rd$LCL > rd$statistics), ])"
Rinterface.RRun "rresult3 <- t(rresult) "
Rinterface.RRun "rresult4 <- t(rresult2)"

Rinterface.GetArray "rresult3", Range(Cells(stn + 39, 4), Cells(stn + 39, 4))
Rinterface.GetArray "rresult4", Range(Cells(stn + 41, 4), Cells(stn + 41, 4))


If Range(Cells(stn + 39, 4), Cells(stn + 39, 4)).Value & Range(Cells(stn + 41, 4), Cells(stn + 41, 4)).Value = "" Then
Range(Cells(stn + 43, 4), Cells(stn + 43, 4)).Value = "������ �������¿� �ִ� ������ ������ �� �ֽ��ϴ�."
ElseIf Range(Cells(stn + 39, 4), Cells(stn + 39, 4)).Value <> "" Then
Range(Cells(stn + 40, 4), Cells(stn + 40, 4)).Value = "��° �κб��� '�������Ѽ�'�� ������ϴ�. ���� ������ �̻������ �ִ� ������ �����˴ϴ�."
ElseIf Range(Cells(stn + 41, 4), Cells(stn + 41, 4)).Value <> "" Then
Range(Cells(stn + 42, 4), Cells(stn + 42, 4)).Value = "��° �κб��� '�������Ѽ�'�� ������ϴ�. ���� ������ �̻������ �ִ� ������ �����˴ϴ�."
End If






Range(Cells(stn + 30, 3), Cells(stn + 43, 3)).Borders(xlEdgeLeft).LineStyle = xlContinuous  '���� ���� �׵θ� ����
  Range(Cells(stn + 30, 3), Cells(stn + 43, 3)).Borders(xlEdgeLeft).Color = RGB(34, 116, 34)
 Range(Cells(stn + 30, 3), Cells(stn + 43, 3)).Borders(xlEdgeLeft).Weight = 3
 
  Range(Cells(stn + 30, 13), Cells(stn + 43, 13)).Borders(xlEdgeRight).LineStyle = xlContinuous  '���� ������ �׵θ� ����
  Range(Cells(stn + 30, 13), Cells(stn + 43, 13)).Borders(xlEdgeRight).Color = RGB(34, 116, 34)
   Range(Cells(stn + 30, 13), Cells(stn + 43, 13)).Borders(xlEdgeRight).Weight = 3
   
    
Range(Cells(stn + 30, 3), Cells(stn + 30, 13)).Borders(xlEdgeTop).LineStyle = xlContinuous  '���� ���� �׵θ� ����
 Range(Cells(stn + 30, 3), Cells(stn + 30, 13)).Borders(xlEdgeTop).Color = RGB(34, 116, 34)
Range(Cells(stn + 30, 3), Cells(stn + 30, 13)).Borders(xlEdgeTop).Weight = 3

  Range(Cells(stn + 43, 3), Cells(stn + 43, 13)).Borders(xlEdgeBottom).LineStyle = xlContinuous  '���� �Ʒ��� �׵θ� ����
   Range(Cells(stn + 43, 3), Cells(stn + 43, 13)).Borders(xlEdgeBottom).Color = RGB(34, 116, 34)
 Range(Cells(stn + 43, 3), Cells(stn + 43, 13)).Borders(xlEdgeBottom).Weight = 3
 
 
  Range(Cells(stn + 30, 3), Cells(stn + 30, 13)).Borders(xlEdgeBottom).LineStyle = xlContinuous  '���� �Ʒ��� �׵θ� ����
   Range(Cells(stn + 30, 3), Cells(stn + 30, 13)).Borders(xlEdgeBottom).Color = RGB(34, 116, 34)
 Range(Cells(stn + 30, 3), Cells(stn + 30, 13)).Borders(xlEdgeBottom).Weight = 3
 
 Range(Cells(stn + 37, 3), Cells(stn + 37, 13)).Borders(xlEdgeBottom).LineStyle = xlContinuous  '���� �Ʒ��� �׵θ� ����
  Range(Cells(stn + 37, 3), Cells(stn + 37, 13)).Borders(xlEdgeBottom).Color = RGB(34, 116, 34)
 Range(Cells(stn + 37, 3), Cells(stn + 37, 13)).Borders(xlEdgeBottom).Weight = 3
 
  Range(Cells(stn + 44, 1), Cells(stn + 44, 25)).Borders(xlEdgeTop).LineStyle = xlContinuous  '���� ���� �׵θ� ����
  Range(Cells(stn + 44, 1), Cells(stn + 44, 25)).Borders(xlEdgeTop).Color = RGB(34, 116, 34)
Range(Cells(stn + 44, 1), Cells(stn + 44, 25)).Borders(xlEdgeTop).Weight = 3
 
 
If N > 44 Then
 ActiveSheet.Cells(1, 1).Value = stn + N + 2
 Else
ActiveSheet.Cells(1, 1).Value = stn + 46
 
 
End If

'If Range("o3") & Range("o5") & Range("o12") & Range("o15") = "" Then
'Range("o19").Value = ""
'Else

'Dim btnI As String
'btnI = Chr(39) & "btnI ""( " & no & ")" & Chr(39)


'Module2.makebtn btnI



'Range("o19").Value = "������Ż���� �����Ͻð� �������� �ٽ� �׸��ðڽ��ϱ�?"
'End If





Unload Me



End Sub

Private Sub UserForm_Click()

End Sub