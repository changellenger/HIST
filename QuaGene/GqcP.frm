VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GqcP 
   OleObjectBlob   =   "GqcP.frx":0000
   Caption         =   "�����ϱ� : P������"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10800
   StartUpPosition =   1  '������ ���
   TypeInfoVer     =   26
End
Attribute VB_Name = "GqcP"
Attribute VB_Base = "0{AEDE5554-9A77-42B6-88B6-5C76BD20836D}{AAFA9642-CCA7-4967-A848-93B1DA2C5E27}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private Sub ListBox5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.ListBox5.ListCount <> 0 Then
        Me.ListBox1.AddItem Me.ListBox5.List(0)
        Me.ListBox5.RemoveItem (0)
     
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

               Exit Sub
            End If
            i = i + 1
        Loop
    End If
    
     Dim j As Integer
    j = 0
    If Me.ListBox5.ListCount = 0 Then
        Do While j <= Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(j) = True Then
               Me.ListBox5.AddItem Me.ListBox1.List(j)
               Me.ListBox1.RemoveItem (j)

               Exit Sub
            End If
            j = j + 1
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
    
     If Me.ListBox5.ListCount = 0 Then
        Do While j <= Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(j) = True Then
               Me.ListBox5.AddItem Me.ListBox1.List(j)
               Me.ListBox1.RemoveItem (j)
              
               Exit Sub
            End If
            j = j + 1
        Loop
    End If
End Sub
Private Sub Cancel_Click()

    Unload Me
    frameConHypo1.Show
    
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
    
    Dim j As Integer
    j = 0
    If Me.ListBox5.ListCount = 0 Then
        Do While j <= Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(j) = True Then
               Me.ListBox5.AddItem Me.ListBox1.List(j)
               Me.ListBox1.RemoveItem (j)

               Exit Sub
            End If
            j = j + 1
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
    Dim a As String ' string  a  ����
   Dim j As Integer
      Dim dataRange2 As Range

    
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
     xlist2 = Me.ListBox5.List(0)
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
    
     Set dataRange2 = ActiveSheet.Cells.CurrentRegion
    m2 = dataRange2.Columns.Count
    
    
    tmp2 = 0
    For j = 1 To m2
        If xlist2 = ActiveSheet.Cells(1, j) Then
            k2 = j  'k1                                 : k1 : ���õ� ������ ���° ���� �ִ���
            tmp2 = tmp2 + 1
        End If
    Next j
    
    N2 = ActiveSheet.Cells(1, k2).End(xlDown).Row - 1    'n         : ���õ� ������ ����Ÿ ����
    
   ' rng = Range(Cells(2, k1), Cells(N + 1, k1))
    
    
    

    '''
    ''' �������� ���� ��� - ������ ���� �ִ� ������ �ԷµǹǷ� ����ó���Ѵ�.
    '''
    If tmp2 > 1 Then
        MsgBox xlist2 & "�� ���� �������� �ֽ��ϴ�. " & vbCrLf & "�������� �ٲ��ֽñ� �ٶ��ϴ�.", vbExclamation, "HIST"
        Exit Sub
    End If
      
    Rinterface.StartRServer
    'Rinterface.PutArray "arraytest", Range(rng)
    Rinterface.RRun "install.packages (" & Chr(34) & "qcc" & Chr(34) & ")"
    Rinterface.RRun "require (qcc)"
    
    
      Rinterface.PutArray "arraytest", Range(Cells(2, k1), Cells(N + 1, k1))
      Rinterface.PutArray "arraytest2", Range(Cells(2, k2), Cells(N2 + 1, k2))
      
      
      
      
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
     ActiveSheet.Cells(stn + 2, 2).Value = xlist2
    Rinterface.GetArray "arraytest", Range(Cells(stn + 3, 1), Cells(stn + 3, 1))
      Rinterface.GetArray "arraytest2", Range(Cells(stn + 3, 2), Cells(stn + 3, 2))
    
    
      ActiveSheet.Cells(stn + 1, 4).Value = "������ �׷���"
      ActiveSheet.Cells(stn + 1, 4).Font.Bold = True
ActiveSheet.Cells(stn + 1, 4).Interior.Color = RGB(220, 238, 130)
      
      

  
a = " p <- qcc(arraytest, type= " & Chr(34) & "p" & Chr(34) & ", size= arraytest2,title = " & Chr(34) & "P ������" & Chr(34) & ")"
Rinterface.RRun a

Rinterface.InsertCurrentRPlot Range(Cells(stn + 3, 4), Cells(stn + 3, 4)), widthrescale:=0.6, heightrescale:=0.6, closergraph:=True

Rinterface.RRun "l <- limits.p(p$center,p$std.dev,p$sizes,3) "
Rinterface.RRun "s <- stats.p(p$data, p$sizes)"
Rinterface.RRun "ss <- as.data.frame(s)"
Rinterface.RRun "d <- data.frame(l,ss)"
Rinterface.RRun "result <- row.names(d[which(d$UCL <= d$statistics), ])"

Rinterface.RRun "result3 <- t(result) "


Rinterface.GetArray "result3", Range(Cells(stn + 32, 5), Cells(stn + 32, 5))



Range(Cells(stn + 32, 5), Cells(stn + 32, 5)).Font.Color = vbRed
Range(Cells(stn + 32, 5), Cells(stn + 32, 5)).Font.Bold = True




Range(Cells(stn + 30, 4), Cells(stn + 30, 4)).Value = "P������ ����ؼ�"
Range(Cells(stn + 30, 4), Cells(stn + 30, 4)).Cells.ColumnWidth = 15
Range(Cells(stn + 30, 4), Cells(stn + 30, 4)).Font.Bold = True
Range(Cells(stn + 30, 4), Cells(stn + 30, 4)).Interior.Color = RGB(220, 238, 130)

Range(Cells(stn + 32, 4), Cells(stn + 32, 4)).Value = "P�������Ѽ��� ����� �κб�:"
Range(Cells(stn + 32, 4), Cells(stn + 32, 4)).Cells.ColumnWidth = 28
Range(Cells(stn + 32, 4), Cells(stn + 32, 4)).Font.Color = vbBlack
Range(Cells(stn + 32, 4), Cells(stn + 32, 4)).Font.Bold = True




If Range(Cells(stn + 32, 5), Cells(stn + 32, 5)).Value = "" Then
Range(Cells(stn + 34, 5), Cells(stn + 34, 5)).Value = "������ �������¿� �ִ� ������ ������ �� �ֽ��ϴ�."
Else
Range(Cells(stn + 33, 5), Cells(stn + 33, 5)).Value = "��° �κб��� '�������Ѽ�'�� ������ϴ�. ���� ������ �̻������ �ִ� ������ �����˴ϴ�."
End If




Module2.makebtn "btnP"




If Range(Cells(stn + 32, 5), Cells(stn + 32, 5)).Value = "" Then
Range(Cells(stn + 35, 5), Cells(stn + 35, 5)).Value = ""
Else
Range(Cells(stn + 35, 5), Cells(stn + 35, 5)).Value = "������Ż���� �����Ͻð� �������� �ٽ� �׸��ðڽ��ϱ�?"
End If





 
 
 Range(Cells(stn + 30, 4), Cells(stn + 35, 4)).Borders(xlEdgeLeft).LineStyle = xlContinuous '���� ���� �׵θ� ����
   Range(Cells(stn + 30, 4), Cells(stn + 35, 4)).Borders(xlEdgeLeft).Color = RGB(34, 116, 34)
 Range(Cells(stn + 30, 4), Cells(stn + 35, 4)).Borders(xlEdgeLeft).Weight = 3
 
 Range(Cells(stn + 30, 14), Cells(stn + 35, 14)).Borders(xlEdgeRight).LineStyle = xlContinuous  '���� ������ �׵θ� ����
 Range(Cells(stn + 30, 14), Cells(stn + 35, 14)).Borders(xlEdgeRight).Color = RGB(34, 116, 34)
  Range(Cells(stn + 30, 14), Cells(stn + 35, 14)).Borders(xlEdgeRight).Weight = 3
 
Range(Cells(stn + 30, 4), Cells(stn + 30, 14)).Borders(xlEdgeTop).LineStyle = xlContinuous  '���� ���� �׵θ� ����
 Range(Cells(stn + 30, 4), Cells(stn + 30, 14)).Borders(xlEdgeTop).Color = RGB(34, 116, 34)
Range(Cells(stn + 30, 4), Cells(stn + 30, 14)).Borders(xlEdgeTop).Weight = 3
 
 
  Range(Cells(stn + 35, 4), Cells(stn + 35, 14)).Borders(xlEdgeBottom).LineStyle = xlContinuous  '���� �Ʒ��� �׵θ� ����
  Range(Cells(stn + 35, 4), Cells(stn + 35, 14)).Borders(xlEdgeBottom).Color = RGB(34, 116, 34)
Range(Cells(stn + 35, 4), Cells(stn + 35, 14)).Borders(xlEdgeBottom).Weight = 3
 

 
 Range(Cells(stn + 30, 4), Cells(stn + 30, 14)).Borders(xlEdgeBottom).LineStyle = xlContinuous  '���� �Ʒ��� �׵θ� ����
   Range(Cells(stn + 30, 4), Cells(stn + 30, 14)).Borders(xlEdgeBottom).Color = RGB(34, 116, 34)
 Range(Cells(stn + 30, 4), Cells(stn + 30, 14)).Borders(xlEdgeBottom).Weight = 3
 
   Range(Cells(stn + 36, 1), Cells(stn + 36, 25)).Borders(xlEdgeBottom).LineStyle = xlContinuous  '���� �Ʒ��� �׵θ� ����
   Range(Cells(stn + 36, 1), Cells(stn + 36, 25)).Borders(xlEdgeBottom).Color = vbBlack
 Range(Cells(stn + 36, 1), Cells(stn + 36, 25)).Borders(xlEdgeBottom).Weight = 1
 
 
 

 
 
 
 


Rinterface.RRun "NofSG <- length(arraytest)" '�κб���'
Rinterface.RRun "MSofSG <- mean(arraytest2)" '��� �κб� ũ��'
Rinterface.RRun "NofD <- sum(arraytest)" '�ҷ�ǰ ��'
Rinterface.RRun "ALL <- sum(arraytest2)" '�� �׸��'
Rinterface.RRun "PofD <- (NofD/ALL)*100" '�ҷ���'

Range(Cells(stn + 4, 8), Cells(stn + 4, 8)).Value = "�κб� ��"
Range(Cells(stn + 4, 8), Cells(stn + 4, 8)).Cells.ColumnWidth = 15
Range(Cells(stn + 5, 8), Cells(stn + 5, 8)).Value = "��� �κб� ũ��"
Range(Cells(stn + 5, 8), Cells(stn + 5, 8)).Cells.ColumnWidth = 15
Range(Cells(stn + 6, 8), Cells(stn + 6, 8)).Value = "�ҷ�ǰ ��"
Range(Cells(stn + 6, 8), Cells(stn + 6, 8)).Cells.ColumnWidth = 15
Range(Cells(stn + 7, 8), Cells(stn + 7, 8)).Value = "�� �׸��"
Range(Cells(stn + 7, 8), Cells(stn + 7, 8)).Cells.ColumnWidth = 15
Range(Cells(stn + 8, 8), Cells(stn + 8, 8)).Value = "�ҷ���"
Range(Cells(stn + 8, 8), Cells(stn + 8, 8)).Cells.ColumnWidth = 15

Range(Cells(stn + 4, 8), Cells(stn + 4, 8)).Font.Bold = True
Range(Cells(stn + 4, 8), Cells(stn + 4, 8)).Interior.Color = RGB(220, 238, 130)
Range(Cells(stn + 5, 8), Cells(stn + 5, 8)).Font.Bold = True
Range(Cells(stn + 5, 8), Cells(stn + 5, 8)).Interior.Color = RGB(220, 238, 130)
Range(Cells(stn + 6, 8), Cells(stn + 6, 8)).Font.Bold = True
Range(Cells(stn + 6, 8), Cells(stn + 6, 8)).Interior.Color = RGB(220, 238, 130)
Range(Cells(stn + 7, 8), Cells(stn + 7, 8)).Font.Bold = True
Range(Cells(stn + 7, 8), Cells(stn + 7, 8)).Interior.Color = RGB(220, 238, 130)
Range(Cells(stn + 8, 8), Cells(stn + 8, 8)).Font.Bold = True
Range(Cells(stn + 8, 8), Cells(stn + 8, 8)).Interior.Color = RGB(220, 238, 130)


Rinterface.GetArray "NofSG", Range(Cells(stn + 4, 9), Cells(stn + 4, 9))
Rinterface.GetArray "MSofSG", Range(Cells(stn + 5, 9), Cells(stn + 5, 9))
Rinterface.GetArray "NofD", Range(Cells(stn + 6, 9), Cells(stn + 6, 9))
Rinterface.GetArray "ALL", Range(Cells(stn + 7, 9), Cells(stn + 7, 9))
Rinterface.GetArray "PofD", Range(Cells(stn + 8, 9), Cells(stn + 8, 9))






Range(Cells(stn + 4, 8), Cells(stn + 8, 8)).Borders(xlEdgeLeft).LineStyle = xlContinuous  '���� ���� �׵θ� ����
   Range(Cells(stn + 4, 8), Cells(stn + 8, 8)).Borders(xlEdgeLeft).Color = RGB(34, 116, 34)
 Range(Cells(stn + 4, 8), Cells(stn + 8, 8)).Borders(xlEdgeLeft).Weight = 3
 
 Range(Cells(stn + 4, 9), Cells(stn + 8, 9)).Borders(xlEdgeRight).LineStyle = xlContinuous  '���� ������ �׵θ� ����
 Range(Cells(stn + 4, 9), Cells(stn + 8, 9)).Borders(xlEdgeRight).Color = RGB(34, 116, 34)
 Range(Cells(stn + 4, 9), Cells(stn + 8, 9)).Borders(xlEdgeRight).Weight = 3
 
 Range(Cells(stn + 4, 8), Cells(stn + 4, 9)).Borders(xlEdgeTop).LineStyle = xlContinuous  '���� ���� �׵θ� ����
 Range(Cells(stn + 4, 8), Cells(stn + 4, 9)).Borders(xlEdgeTop).Color = RGB(34, 116, 34)
 Range(Cells(stn + 4, 8), Cells(stn + 4, 9)).Borders(xlEdgeTop).Weight = 3
 
 
 Range(Cells(stn + 8, 8), Cells(stn + 8, 9)).Borders(xlEdgeBottom).LineStyle = xlContinuous  '���� �Ʒ��� �׵θ� ����
  Range(Cells(stn + 8, 8), Cells(stn + 8, 9)).Borders(xlEdgeBottom).Color = RGB(34, 116, 34)
Range(Cells(stn + 8, 8), Cells(stn + 8, 9)).Borders(xlEdgeBottom).Weight = 3
 

  Range(Cells(stn + 4, 9), Cells(stn + 8, 9)).Borders(xlEdgeLeft).LineStyle = xlContinuous  '���� ���� �׵θ�
   Range(Cells(stn + 4, 9), Cells(stn + 8, 9)).Borders(xlEdgeLeft).Color = RGB(34, 116, 34)
  Range(Cells(stn + 4, 9), Cells(stn + 8, 9)).Borders(xlEdgeLeft).Weight = 3



If N > 35 Then
 ActiveSheet.Cells(1, 1).Value = stn + N + 2
 Else
ActiveSheet.Cells(1, 1).Value = stn + 37
 
 
End If


      'Rinterface.RRun "x1 <- matrix(data= arraytest, nrow = 4, ncol=5, byrow = FALSE, dimnames = NULL)"
'no2 = Me.TextBox.Value
'Rinterface.RRun "qcc(arraytest, type= " & Chr(34) & "p" & Chr(34) & ", size=3)"
'b = " qcc(x1, type= " & Chr(34) & "p" & Chr(34) & ", size=" & no2 & ")  "
'Rinterface.RRun b
Unload Me
End Sub

Private Sub UserForm_Click()

End Sub
