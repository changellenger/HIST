VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GqcC 
   OleObjectBlob   =   "GqcC.frx":0000
   Caption         =   "�����ϱ� : C������"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10785
   StartUpPosition =   1  '������ ���
   TypeInfoVer     =   24
End
Attribute VB_Name = "GqcC"
Attribute VB_Base = "0{905AD38C-6723-439C-94AC-32C0E566CBAD}{11B00A06-16BE-4276-B543-F33EE4308365}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private Sub CB1_Click()
    Dim i As Integer
    i = 0
    If Me.ListBox2.ListCount = 0 Then
        Do While i <= Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(i) = True Then
               Me.ListBox2.AddItem Me.ListBox1.List(i)
               Me.ListBox1.RemoveItem (i)

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
      
    End If
End Sub
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim i As Integer
    i = 0
    If Me.ListBox2.ListCount = 0 Then
        Do While i <= Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(i) = True Then
               Me.ListBox2.AddItem Me.ListBox1.List(i)
               Me.ListBox1.RemoveItem (i)
        
               Exit Sub
            End If
            i = i + 1
        Loop
    End If
End Sub
Private Sub Cancel_Click()


    frameConHypo1.Show
        Unload Me
    
End Sub



Private Sub OK_Click()
    Dim i As Integer
    i = 0
    If Me.ListBox2.ListCount = 0 Then
        Do While i <= Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(i) = True Then
               Me.ListBox2.AddItem Me.ListBox1.List(i)
               Me.ListBox1.RemoveItem (i)

               Exit Sub
            End If
            i = i + 1
        Loop
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
Private Sub ToggleButton1_Click()
                           '�ѱ� ������ ������,�ŷڱ���,�븳���� 3���ϱ�
    Dim dataRange As Range
    Dim i As Integer
    Dim activePt As Long                                '��� �м��� ���۵Ǵ� �κ��� �����ֱ� ����
    Dim rng As Range
    Dim b As String ' string  a  ����
    Dim no2 As Integer '�κб�
  
    
    
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
    
     'no = Me.TextBox1.Value
      Rinterface.PutArray "arraytest", Range(Cells(2, k1), Cells(N + 1, k1))
   
    'a = " qcc(arraytest, type= " & Chr(34) & "c" & Chr(34) & ", size=" & no & ")  "
'Rinterface.RRun a
      'Rinterface.RRun "x1 <- matrix(data= arraytest, nrow = 20, ncol=5, byrow = FALSE, dimnames = NULL)"

'Rinterface.RRun "qcc(x1, type= " & Chr(34) & "c" & Chr(34) & ", size=5)"

  Application.ScreenUpdating = False
    Dim stname As String
    'Dim lastCol, lastRow As Integer
    
    stname = "�����ϱ� ������"
    Module3.OpenOutSheet stname, True
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
b = " c <- qcc(arraytest, type= " & Chr(34) & "c" & Chr(34) & ", size=1000, title = " & Chr(34) & "C������" & Chr(34) & ")"
Rinterface.RRun b

Rinterface.InsertCurrentRPlot Range(Cells(stn + 3, 3), Cells(stn + 3, 3)), widthrescale:=0.6, heightrescale:=0.6, closergraph:=True


Rinterface.RRun "cl <- limits.c(c$center,c$std.dev,c$sizes,3) "
Rinterface.RRun "cs <- stats.c(c$data, c$sizes)"
Rinterface.RRun "css <- as.data.frame(cs)"
Rinterface.RRun "cd <- data.frame(cl,css)"
Rinterface.RRun "cresult <- row.names(cd[which(cd$UCL < cd$statistics), ])"
Rinterface.RRun "cresult3 <- t(cresult) "


Rinterface.GetArray "cresult3", Range(Cells(stn + 32, 4), Cells(stn + 32, 4))

Range(Cells(stn + 32, 4), Cells(stn + 32, 4)).Font.Color = vbRed
Range(Cells(stn + 32, 4), Cells(stn + 32, 4)).Font.Bold = True





Range(Cells(stn + 30, 3), Cells(stn + 30, 3)).Value = "C������ ����ؼ�"
Range(Cells(stn + 30, 3), Cells(stn + 30, 3)).Cells.ColumnWidth = 15
Range(Cells(stn + 30, 3), Cells(stn + 30, 3)).Font.Bold = True
Range(Cells(stn + 30, 3), Cells(stn + 30, 3)).Interior.Color = RGB(220, 238, 130)


Range(Cells(stn + 32, 3), Cells(stn + 32, 3)).Value = "C�������Ѽ��� ����� �κб�:"
Range(Cells(stn + 32, 3), Cells(stn + 32, 3)).Cells.ColumnWidth = 28
Range(Cells(stn + 32, 3), Cells(stn + 32, 3)).Font.Color = vbBlack
Range(Cells(stn + 32, 3), Cells(stn + 32, 3)).Font.Bold = True




If Range(Cells(stn + 32, 4), Cells(stn + 32, 4)).Value = "" Then
Range(Cells(stn + 34, 4), Cells(stn + 34, 4)).Value = "������ �������¿� �ִ� ������ ������ �� �ֽ��ϴ�."
Else
Range(Cells(stn + 33, 4), Cells(stn + 33, 4)).Value = "��° �κб��� '�������Ѽ�'�� ������ϴ�. ���� ������ �̻������ �ִ� ������ �����˴ϴ�."
End If



Module2.makebtn "btnC"




If Range(Cells(stn + 32, 4), Cells(stn + 32, 4)).Value = "" Then
Range(Cells(stn + 35, 4), Cells(stn + 35, 4)).Value = ""
Else
Range(Cells(stn + 35, 4), Cells(stn + 35, 4)).Value = "������Ż���� �����Ͻð� �������� �ٽ� �׸��ðڽ��ϱ�?"
End If




 
Range(Cells(stn + 30, 3), Cells(stn + 35, 3)).Borders(xlEdgeLeft).LineStyle = xlContinuous '���� ���� �׵θ� ����
   Range(Cells(stn + 30, 3), Cells(stn + 35, 3)).Borders(xlEdgeLeft).Color = RGB(34, 116, 34)
 Range(Cells(stn + 30, 3), Cells(stn + 35, 3)).Borders(xlEdgeLeft).Weight = 3
 
 Range(Cells(stn + 30, 13), Cells(stn + 35, 13)).Borders(xlEdgeRight).LineStyle = xlContinuous  '���� ������ �׵θ� ����
 Range(Cells(stn + 30, 13), Cells(stn + 35, 13)).Borders(xlEdgeRight).Color = RGB(34, 116, 34)
  Range(Cells(stn + 30, 13), Cells(stn + 35, 13)).Borders(xlEdgeRight).Weight = 3
 
Range(Cells(stn + 30, 3), Cells(stn + 30, 13)).Borders(xlEdgeTop).LineStyle = xlContinuous  '���� ���� �׵θ� ����
 Range(Cells(stn + 30, 3), Cells(stn + 30, 13)).Borders(xlEdgeTop).Color = RGB(34, 116, 34)
Range(Cells(stn + 30, 3), Cells(stn + 30, 13)).Borders(xlEdgeTop).Weight = 3
 
 
  Range(Cells(stn + 35, 3), Cells(stn + 35, 13)).Borders(xlEdgeBottom).LineStyle = xlContinuous  '���� �Ʒ��� �׵θ� ����
  Range(Cells(stn + 35, 3), Cells(stn + 35, 13)).Borders(xlEdgeBottom).Color = RGB(34, 116, 34)
Range(Cells(stn + 35, 3), Cells(stn + 35, 13)).Borders(xlEdgeBottom).Weight = 3
 

 
 Range(Cells(stn + 30, 3), Cells(stn + 30, 13)).Borders(xlEdgeBottom).LineStyle = xlContinuous  '���� �Ʒ��� �׵θ� ����
   Range(Cells(stn + 30, 3), Cells(stn + 30, 13)).Borders(xlEdgeBottom).Color = RGB(34, 116, 34)
 Range(Cells(stn + 30, 3), Cells(stn + 30, 13)).Borders(xlEdgeBottom).Weight = 3
 
   Range(Cells(stn + 36, 1), Cells(stn + 36, 25)).Borders(xlEdgeBottom).LineStyle = xlContinuous  '���� �Ʒ��� �׵θ� ����
   Range(Cells(stn + 36, 1), Cells(stn + 36, 25)).Borders(xlEdgeBottom).Color = vbBlack
 Range(Cells(stn + 36, 1), Cells(stn + 36, 25)).Borders(xlEdgeBottom).Weight = 1
 

 
 

'Rinterface.RRun "NofSG <- length(arraytest)" '�κб���'
'Rinterface.RRun "MSofSG <- mean(arraytest2)" '�κб� ũ��'
Rinterface.RRun "AA <- length(arraytest)" '�� �˻� ������
Rinterface.RRun "BB <- sum(arraytest)" '�� ������'
Rinterface.RRun "CC <- BB/AA" '������ ���� ��'

Range(Cells(stn + 4, 7), Cells(stn + 4, 7)).Value = "�� �˻� ������"
Range(Cells(stn + 4, 7), Cells(stn + 4, 7)).Cells.ColumnWidth = 15
Range(Cells(stn + 5, 7), Cells(stn + 5, 7)).Value = "�� ������"
Range(Cells(stn + 5, 7), Cells(stn + 5, 7)).Cells.ColumnWidth = 15
Range(Cells(stn + 6, 7), Cells(stn + 6, 7)).Value = "������ ���� ��"
Range(Cells(stn + 6, 7), Cells(stn + 6, 7)).Cells.ColumnWidth = 15
'Range("g28").Value = "�� ������"
'Range("g29").Value = "������ ���� ��"

Range(Cells(stn + 4, 7), Cells(stn + 4, 7)).Font.Bold = True
Range(Cells(stn + 4, 7), Cells(stn + 4, 7)).Interior.Color = RGB(220, 238, 130)
Range(Cells(stn + 5, 7), Cells(stn + 5, 7)).Font.Bold = True
Range(Cells(stn + 5, 7), Cells(stn + 5, 7)).Interior.Color = RGB(220, 238, 130)
Range(Cells(stn + 6, 7), Cells(stn + 6, 7)).Font.Bold = True
Range(Cells(stn + 6, 7), Cells(stn + 6, 7)).Interior.Color = RGB(220, 238, 130)

Rinterface.GetArray "AA", Range(Cells(stn + 4, 8), Cells(stn + 4, 8))
Rinterface.GetArray "BB", Range(Cells(stn + 5, 8), Cells(stn + 5, 8))
Rinterface.GetArray "CC", Range(Cells(stn + 6, 8), Cells(stn + 6, 8))
'Rinterface.GetArray "SD", Range("sheet1!i28")
'Rinterface.GetArray "PofD", Range("sheet1!i29")



 
 
 

Range(Cells(stn + 4, 7), Cells(stn + 6, 7)).Borders(xlEdgeLeft).LineStyle = xlContinuous  '���� ���� �׵θ� ����
   Range(Cells(stn + 4, 7), Cells(stn + 6, 7)).Borders(xlEdgeLeft).Color = RGB(34, 116, 34)
 Range(Cells(stn + 4, 7), Cells(stn + 6, 7)).Borders(xlEdgeLeft).Weight = 3
 
 Range(Cells(stn + 4, 8), Cells(stn + 6, 8)).Borders(xlEdgeRight).LineStyle = xlContinuous  '���� ������ �׵θ� ����
 Range(Cells(stn + 4, 8), Cells(stn + 6, 8)).Borders(xlEdgeRight).Color = RGB(34, 116, 34)
 Range(Cells(stn + 4, 8), Cells(stn + 6, 8)).Borders(xlEdgeRight).Weight = 3
 
 Range(Cells(stn + 4, 7), Cells(stn + 4, 8)).Borders(xlEdgeTop).LineStyle = xlContinuous  '���� ���� �׵θ� ����
 Range(Cells(stn + 4, 7), Cells(stn + 4, 8)).Borders(xlEdgeTop).Color = RGB(34, 116, 34)
 Range(Cells(stn + 4, 7), Cells(stn + 4, 8)).Borders(xlEdgeTop).Weight = 3
 
 
 Range(Cells(stn + 6, 7), Cells(stn + 6, 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous  '���� �Ʒ��� �׵θ� ����
  Range(Cells(stn + 6, 7), Cells(stn + 6, 8)).Borders(xlEdgeBottom).Color = RGB(34, 116, 34)
Range(Cells(stn + 6, 7), Cells(stn + 6, 8)).Borders(xlEdgeBottom).Weight = 3
 

  Range(Cells(stn + 4, 8), Cells(stn + 6, 8)).Borders(xlEdgeLeft).LineStyle = xlContinuous  '���� ���� �׵θ�
   Range(Cells(stn + 4, 8), Cells(stn + 6, 8)).Borders(xlEdgeLeft).Color = RGB(34, 116, 34)
  Range(Cells(stn + 4, 8), Cells(stn + 6, 8)).Borders(xlEdgeLeft).Weight = 3

   If N > 35 Then
 ActiveSheet.Cells(1, 1).Value = stn + N + 2
 Else
ActiveSheet.Cells(1, 1).Value = stn + 37
 
 
End If
   
Unload Me
End Sub

Private Sub UserForm_Click()

End Sub
