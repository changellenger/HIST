VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameArima_back 
   OleObjectBlob   =   "frameArima_back.frx":0000
   Caption         =   "ARIMA"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8985
   StartUpPosition =   1  '������ ���
   TypeInfoVer     =   58
End
Attribute VB_Name = "frameArima_back"
Attribute VB_Base = "0{35B3392E-062F-40B0-8182-B302054AA211}{6CB53F66-C5D7-418C-91E0-8C5742ECA88E}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private Sub CheckBox7_Click()

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
rinterface.InsertCurrentRPlot Range("o23"), widthrescale:=0.5, heightrescale:=0.5, closergraph:=True


If CheckBox7.Value = True Then

Range("n3").Value = "���� ��:"

rinterface.RRun "mm <- forecast.Arima(ar, h=" & no4 & ")"
rinterface.RRun "mmm <- as.data.frame(mm)"
rinterface.GetArray "mmm", Range("o3")

End If



If CheckBox6.Value = True Then

Range("n10").Value = "�ŷڱ���:"

rinterface.RRun "conf <- confint(ar)"
'Rinterface.RRun "conff <- as.data.frame(conf)"


rinterface.GetDataframe "conf", Range("o10")
    
End If

If CheckBox8.Value = True Then
b = " auto <- auto.arima(arraytest)"
rinterface.RRun b

rinterface.GetArray "auto", Range("o50")
End If




If CheckBox8.Value = True Then
rinterface.RRun "win.graph()"
rinterface.RRun "plot(forecast(ar))"
rinterface.InsertCurrentRPlot Range("o10"), widthrescale:=0.5, heightrescale:=0.5, closergraph:=True
End If

'Rinterface.InsertCurrentRPlot Range("sheet9!o28"), widthrescale:=0.5, heightrescale:=0.5, closergraph:=True


Unload Me


End Sub


Private Sub UserForm_Click()

End Sub
