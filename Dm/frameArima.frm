VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameArima 
   OleObjectBlob   =   "frameArima.frx":0000
   Caption         =   "ARIMA"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7950
   StartUpPosition =   1  '������ ���
   TypeInfoVer     =   66
End
Attribute VB_Name = "frameArima"
Attribute VB_Base = "0{38044FDA-5606-4D99-B96A-FDFFAD6D3926}{65EFD58B-E644-4714-92FF-850CB506DD3C}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private Sub CheckBox7_Click()

End Sub

Private Sub Label4_Click()

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

Private Sub TextBox1_Change()

End Sub

Private Sub ToggleButton1_Click()
                           '�ѱ� ������ ������,�ŷڱ���,�븳���� 3���ϱ�
    Dim dataRange As Range
    Dim i As Integer
    Dim activePt As Long                                '��� �м��� ���۵Ǵ� �κ��� �����ֱ� ����
    Dim rng As Range
    Dim a As String ' string  a  ����
    Dim no As Integer '�κб�
    Dim b As String
    Dim no2 As Integer
     Dim c As String
    
    
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
    'Rinterface.PutArray "arraytest", Range(rng)
    rinterface.RRun "install.packages (" & Chr(34) & "forecast" & Chr(34) & ")"
    rinterface.RRun "require (forecast)"
 '   a = "require (qcc)"
  '  Rinterface.RRun a
  
  
      rinterface.PutArray "arraytest", Range(Cells(2, k1), Cells(n + 1, k1))
     
     no1 = Me.TextBox1.Value
     no2 = Me.TextBox2.Value
     no3 = Me.TextBox3.Value
     
     
     
a = "ar <-arima(arraytest, order=c(" & no1 & "," & no2 & "," & no3 & ")) "
rinterface.RRun a
rinterface.RRun "tsdiag(ar)"
rinterface.InsertCurrentRPlot Range("h20"), widthrescale:=0.5, heightrescale:=0.5, closergraph:=True





If CheckBox7.Value = True Then

Range("N3").Value = "�м� ���"

rinterface.RRun "mm<-forecast.Arima(ar, h=" & no4 & ")"

       rinterface.RRun "tq<-data.frame(mm$mean,mm$lower[,2],mm$upper[,2])"
       rinterface.GetDataframe "tq", Range("n4"), True
         ActiveSheet.Cells(3, 14).Font.Bold = True
     ActiveSheet.Cells(3, 14).Interior.Color = RGB(220, 238, 130)
     ActiveSheet.Cells(3, 14).Cells.ColumnWidth = 17
     ActiveSheet.Cells(3, 14).HorizontalAlignment = xlCenter
           Range("o4").Value = "������"
            Range("p4").Value = "95% �ŷڼ���(����)"
             Range("q4").Value = "95% �ŷڼ���(����)"
        End If
    




If CheckBox14.Value = True Then

 rinterface.RRun "plot.forecast(mm)"
     '  Rinterface.RRun "win.graph()"
       rinterface.InsertCurrentRPlot Range("h4"), widthrescale:=0.5, heightrescale:=0.5, closergraph:=True
       
     Range("h3").Value = "���� �׷���"
     Range("h3").Font.Bold = True
     Range("h3").Interior.Color = RGB(220, 238, 130)
     Range("h3").Cells.ColumnWidth = 17
     Range("h3").HorizontalAlignment = xlCenter
       
       
        End If
        
        If CheckBox18.Value = True Then
 rinterface.RRun "hist(mm$residuals,col = " & Chr(34) & "lightblue" & Chr(34) & ")"
 rinterface.InsertCurrentRPlot Range("n20"), widthrescale:=0.5, heightrescale:=0.5, closergraph:=True
 
Range("n19").Value = "���� ������׷�"
Range("n19").Font.Bold = True
Range("n19").Interior.Color = RGB(220, 238, 130)
Range("n19").Cells.ColumnWidth = 17
Range("n19").HorizontalAlignment = xlCenter

 
 End If
    
 If CheckBox19.Value = True Then
 
 Range("s19").Value = "���� ���� Ȯ����"
Range("s19").Font.Bold = True
Range("s19").Interior.Color = RGB(220, 238, 130)
Range("s19").Cells.ColumnWidth = 17
Range("s19").HorizontalAlignment = xlCenter
 
rinterface.RRun "qqnorm(mm$residuals, col=" & Chr(34) & " blue " & Chr(34) & ")"
rinterface.RRun "qqline(mm$residuals, col=" & Chr(34) & " red " & Chr(34) & ")"
rinterface.InsertCurrentRPlot Range("s20"), widthrescale:=0.5, heightrescale:=0.5, closergraph:=True
 End If

'If CheckBox8.Value = True Then
'Rinterface.RRun "win.graph()"
'Rinterface.RRun "plot(forecast(ar))"
'Rinterface.InsertCurrentRPlot Range("sheet9!o100"), widthrescale:=0.5, heightrescale:=0.5, closergraph:=True
'End If

'Rinterface.InsertCurrentRPlot Range("sheet9!o28"), widthrescale:=0.5, heightrescale:=0.5, closergraph:=True


Unload Me


End Sub


Private Sub UserForm_Click()

End Sub
