VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameES 
   OleObjectBlob   =   "frameES.frx":0000
   Caption         =   "������Ȱ"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7995
   StartUpPosition =   1  '������ ���
   TypeInfoVer     =   89
End
Attribute VB_Name = "frameES"
Attribute VB_Base = "0{C9E1C79A-1318-4F37-A3A8-DC1F9B0772FC}{4A40338F-A724-4F95-8DCA-3F38A42C69CE}"
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

Private Sub Frame1_Click()

End Sub

Private Sub TextBox5_Change()

End Sub

Private Sub UserForm_Initialize()
Dim myArray As Variant
ComboBox1.ColumnCount = 1
myArray = [{"����","�б⺰"}]
ComboBox1.List = myArray
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

Private Sub okbtn_Click()
  Dim dataRange As Range
    Dim i As Integer
    Dim activePt As Long                                '��� �м��� ���۵Ǵ� �κ��� �����ֱ� ����
    Dim rng As Range
    Dim a As String ' string  a  ����
    Dim no As Integer '�κб�
    Dim b As String
    Dim no2 As Integer
     Dim c As String
     Dim comvar As String
     

    
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
    
    n = ActiveSheet.Cells(1, k1).End(xlDown).Row - 1    'n         : ���õ� ������ ����Ÿ ����
    
   ' rng = Range(Cells(2, k1), Cells(N + 1, k1))
    
    
    

    '''
    ''' �������� ���� ��� - ������ ���� �ִ� ������ �ԷµǹǷ� ����ó���Ѵ�.
    '''
    If tmp > 1 Then
        MsgBox xlist & "�� ���� �������� �ֽ��ϴ�. " & vbCrLf & "�������� �ٲ��ֽñ� �ٶ��ϴ�.", vbExclamation, "HIST"
        Exit Sub
    End If
    
    
 rinterface.StartRServer
 
    rinterface.RRun "install.packages (" & Chr(34) & "forecast" & Chr(34) & ")"
    rinterface.RRun "require (forecast)"
    
     'Rinterface.PutDataframe "arraytest", Range(Cells(1, k1), Cells(N + 1, k1))
     rinterface.PutArray "arraytest", Range(Cells(2, k1), Cells(n + 1, k1))
     no = Me.TextBox1.Value
     no2 = Me.TextBox2.Value
     no3 = Me.TextBox3.Value
     no4 = Me.TextBox4.Value
     no5 = Me.TextBox5.Value
    
    comvar = Me.ComboBox1.Value
    'MsgBox comvar
              
              If comvar = "����" Then
    
   a = "essr <- ts(arraytest,start=c(" & no & "," & no2 & "),freq = 12)"
   'MsgBox a
  rinterface.RRun a
  End If
  
    
   
              If comvar = "�б⺰" Then
    
   a = "essr <- ts(arraytest,start=c(" & no & "," & no2 & "),freq = 4)"
   'MsgBox a & " alskdjflkasdf"
  rinterface.RRun a
  End If
  
        Range("h3").Value = "������Ȱ �׷���"
               ActiveSheet.Cells(3, 8).Font.Bold = True
     ActiveSheet.Cells(3, 8).Interior.Color = RGB(220, 238, 130)
     ActiveSheet.Cells(3, 8).Cells.ColumnWidth = 17
     ActiveSheet.Cells(3, 8).HorizontalAlignment = xlCenter
  If CheckBox16.Value = True Then
    rinterface.RRun "es.hw <- HoltWinters(essr,alpha=" & no3 & ", beta = FALSE, gamma = FALSE, l.start = " & no4 & ")"

    rinterface.RRun "plot(es.hw)"
    rinterface.InsertCurrentRPlot Range("h4"), widthrescale:=0.5, heightrescale:=0.5, closergraph:=True
 
        If CheckBox17.Value = True Then
            rinterface.RRun " pre <- forecast.HoltWinters(es.hw, h=" & no5 & ")"
        
         'Rinterface.GetArray "pre", Range("o3"), False, False, False, True, True
         
         
         rinterface.RRun "tq<-data.frame(pre$mean,pre$lower[,2],pre$upper[,2])"
         
         
         rinterface.GetDataframe "tq", Range("o3"), True
          
          Range("o3").Value = "������"
               ActiveSheet.Cells(3, 15).Font.Bold = True
     ActiveSheet.Cells(3, 15).Interior.Color = RGB(220, 238, 130)
     ActiveSheet.Cells(3, 15).Cells.ColumnWidth = 17
     ActiveSheet.Cells(3, 15).HorizontalAlignment = xlCenter
           Range("p3").Value = "������"
            Range("q3").Value = "95% �ŷڼ���(����)"
  
     Range("q3").Cells.ColumnWidth = 17
      Range("q3").HorizontalAlignment = xlCenter
             Range("r3").Value = "95% �ŷڼ���(����)"
              Range("r3").Cells.ColumnWidth = 17
      Range("r3").HorizontalAlignment = xlCenter
        End If

        'If CheckBox12.Value = True Then
        'Range("o9").Value = "SSE:"
        'rinterface.RRun "sse <- es.hw$SSE"
        'rinterface.GetArray "sse", Range("o10")
        'End If
        
        '============================
         If CheckBox12.Value = True Then
        Range("t3").Value = "������Ȯ��"
        Range("t3").Cells.ColumnWidth = 15
        Range("t3").Font.Bold = True
        Range("t3").Interior.Color = RGB(220, 238, 130)
        Range("t3").HorizontalAlignment = xlRight
        rinterface.RRun "sse <- t<-accuracy(pre)"
        rinterface.RRun "ssee<-t(sse)"
        rinterface.GetArray "ssee", Range("u4"), False, False, False, True, False
        
        End If
        '============================
        
        If CheckBox14.Value = True Then
       rinterface.RRun "plot.forecast(pre)"
     '  Rinterface.RRun "win.graph()"
       rinterface.InsertCurrentRPlot Range("h22"), widthrescale:=0.5, heightrescale:=0.5, closergraph:=True
        End If
        
        If CheckBox18.Value = True Then
        Range("L21").Value = "���� ������׷�"
        Range("L21").Font.Bold = True
        Range("L21").Cells.ColumnWidth = 15
        Range("L21").Interior.Color = RGB(220, 238, 130)
        Range("L21").HorizontalAlignment = xlLeft
 
 rinterface.RRun "hist(pre$residuals,col = " & Chr(34) & "lightblue" & Chr(34) & ")"
 rinterface.InsertCurrentRPlot Range("L22"), widthrescale:=0.5, heightrescale:=0.5, closergraph:=True
 End If
 
 If CheckBox19.Value = True Then
 
 
 Range("p21").Value = "���� ���� Ȯ����"
        Range("p21").Font.Bold = True
        Range("p21").Cells.ColumnWidth = 15
        Range("p21").Interior.Color = RGB(220, 238, 130)
        Range("p21").HorizontalAlignment = xlRight
 
 rinterface.RRun "qqnorm(pre$residuals, col=" & Chr(34) & "blue" & Chr(34) & ")"
rinterface.RRun "qqline(pre$residuals, col=" & Chr(34) & "red" & Chr(34) & ")"
rinterface.InsertCurrentRPlot Range("p22"), widthrescale:=0.5, heightrescale:=0.5, closergraph:=True
 End If
        

    Else
        rinterface.RRun "es.hw2 <- HoltWinters(essr,alpha=" & no3 & ", beta = FALSE, gamma = FALSE)"
        rinterface.RRun "plot(es.hw2)"
        rinterface.InsertCurrentRPlot Range("h5"), widthrescale:=0.5, heightrescale:=0.5, closergraph:=True
  
  
         If CheckBox17.Value = True Then
           rinterface.RRun " pre2 <- forecast.HoltWinters(es.hw2, h=" & no5 & ")"
      
           rinterface.RRun "tq<-data.frame(pre2$mean,pre2$lower[,2],pre2$upper[,2])"
         
         
         rinterface.GetDataframe "tq", Range("o3"), True
         
          
          Range("o3").Value = "������"
               ActiveSheet.Cells(3, 15).Font.Bold = True
     ActiveSheet.Cells(3, 15).Interior.Color = RGB(220, 238, 130)
     ActiveSheet.Cells(3, 15).Cells.ColumnWidth = 17
     ActiveSheet.Cells(3, 15).HorizontalAlignment = xlCenter
           Range("p3").Value = "������"
            Range("q3").Value = "95% �ŷڼ���(����)"
               Range("q3").Cells.ColumnWidth = 18.5
             Range("r3").Value = "95% �ŷڼ���(����)"
                 Range("r3").Cells.ColumnWidth = 18.5
             
         End If

    '    If CheckBox12.Value = True Then
    '        Range("t3").Value = "SSE:"
    '        Range("t3").Font.Bold = True
    '        Range("t3").Interior.Color = RGB(220, 238, 130)
    '        Range("t3").HorizontalAlignment = xlRight
    '        rinterface.RRun "ssee <- es.hw2$SSE"
    '        rinterface.GetArray "ssee", Range("u3")
    '    End If
        
        
        '==============================
         If CheckBox12.Value = True Then
        Range("u3").Value = "������Ȯ��:"
        Range("u3").Font.Bold = True
        Range("u3").Interior.Color = RGB(220, 238, 130)
        Range("u3").HorizontalAlignment = xlRight
        rinterface.RRun "sse <- t<-accuracy(pre2)"
        rinterface.RRun "ssee<-t(sse)"
        rinterface.GetArray "ssee", Range("u3"), False, False, False, True, False
        End If
        
        
        '==============================
        
          If CheckBox14.Value = True Then
            rinterface.RRun "plot.forecast(pre2)"
            '  Rinterface.RRun "win.graph()"
        rinterface.InsertCurrentRPlot Range("h21"), widthrescale:=0.5, heightrescale:=0.5, closergraph:=True
        End If
        
                If CheckBox18.Value = True Then
                Range("m20").Value = "���� �׷���"
Range("m20").Font.Bold = True
Range("m20").Interior.Color = RGB(220, 238, 130)
Range("m20").Cells.ColumnWidth = 17
Range("m20").HorizontalAlignment = xlCenter

 rinterface.RRun "hist(pre2$residuals,col = " & Chr(34) & " lightblue " & Chr(34) & ")"
 rinterface.InsertCurrentRPlot Range("m21"), widthrescale:=0.5, heightrescale:=0.5, closergraph:=True
 End If
 
 If CheckBox19.Value = True Then
 
 rinterface.RRun "qqnorm(pre2$residuals, col=" & Chr(34) & " blue " & Chr(34) & ")"
 rinterface.RRun "qqline(pre2$residuals, col=" & Chr(34) & " red " & Chr(34) & ")"
 rinterface.InsertCurrentRPlot Range("p21"), widthrescale:=0.5, heightrescale:=0.5, closergraph:=True
 End If
        
 End If


 



'If CheckBox14.Value = True Then
'Rinterface.RRun "es.hw.forecast <- forecast.HoltWinters(es.hw, h=8)"
'Rinterface.RRun "win.graph()"
'Rinterface.RRun "plot.Forecast (es.hw.Forecast)"
'End If




End Sub


Sub CleanCharts()
    Dim chrt As Picture
    On Error Resume Next
    For Each chrt In ActiveSheet.Pictures
        chrt.Delete
    Next chrt
End Sub

Private Sub TextBox2_Change()

End Sub

Private Sub UserForm_Click()

End Sub
