VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} doe5 
   OleObjectBlob   =   "doe5.frx":0000
   Caption         =   "��ȣ�ۿ뵵"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4635
   StartUpPosition =   1  '������ ���
   TypeInfoVer     =   74
End
Attribute VB_Name = "doe5"
Attribute VB_Base = "0{58688D27-BDD2-47AC-9BC3-849A381E226E}{4250754F-7B0D-4EF3-9A4C-281CE07C1A16}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private Sub btnCancel_Click()
Unload Me
End Sub

Private Sub CB1_Click()
  MoveBtwnListBox Me, "ListBox1", "ListBox2"
 
End Sub
 
Private Sub CB2_Click()
  MoveBtwnListBox Me, "ListBox2", "ListBox1"
End Sub


Private Sub ComboBox1_Change()      ' �޺��ڽ� �ٲ����� ����Ʈ�ڽ� ����
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
   
   Me.ListBox1.Clear
 Me.ListBox2.Clear


    ReDim myArray(TempSheet.UsedRange.Columns.count - 2)
    a = 0
   For i = 1 To TempSheet.UsedRange.Columns.count
   If arrName(i) <> Me.ComboBox1.value Then
   If arrName(i) <> "" Then                     '��ĭ����
   myArray(a) = arrName(i)
   a = a + 1
   
   Else:
   End If
   End If
   Next i
      
   Me.ListBox1.list() = myArray
     
   
End Sub
Private Sub ToggleButton1_Click()

                           '�ѱ� ������ ������,�ŷڱ���,�븳���� 3���ϱ�
    Dim dataRange As Range
    Dim i As Integer
    Dim activePt As Long                                '��� �м��� ���۵Ǵ� �κ��� �����ֱ� ����
    Dim rng As Range
    Dim xlist()
    Dim xstrlist()
    Dim nol As Integer
    Dim k1() As Integer
    
      nol = Me.ListBox2.ListCount 'ListBox2�� �ִ� ��������
      
    rinterface.RRun "require (FrF2)"
    rinterface.RRun "require (qualityTools)"
    rinterface.StartRServer
    
    
    If nol = 0 Then
        MsgBox "������ ������ �ֽñ� �ٶ��ϴ�.", vbExclamation, "HIST"
'    ElseIf nol >= 21 Then
'
'        MxgBox "�м������� 20�� ���Ϸ� �����ؾ� �մϴ�.", vbExclamation, "HIST"
        Exit Sub
'    Else
    
    End If
    

    
    
    '''
    '''public ���� ���� xlist, DataSheet, RstSheet, m, k1, n
    '''

        DataSheet = ActiveSheet.name                        'DataSheet : Data�� �ִ� Sheet �̸�
        RstSheet = "_���м����_"                       'RstSheet  : ����� �����ִ� Sheet �̸�

    
    '������ �Է�
'On Error GoTo Err_delete
Dim val3535 As Long '�ʱ���ġ ������ ����'
Dim s3535 As Worksheet
            val3535 = 2
    For Each s3535 In ActiveWorkbook.Sheets
        If s3535.name = RstSheet Then
            val3535 = Sheets(RstSheet).Cells(1, 1).value
        End If
    Next s3535  '��Ʈ�� �̹������� ��� ��ġ �������ϰ�, ������ 2�� �����Ѵ�.


  
    ReDim k1(nol) As Integer
     ReDim xstrlist(nol - 1)
    ReDim xlist(nol)                       'ListBox2�� �ִ� List(j)��° �������� xlist(j)�� �Ҵ�
    
    For j = 0 To nol - 1
    xstrlist(j) = ListBox2.list(j)
    
        xlist(j) = ListBox2.list(j)
    Next j
        xlist(j) = doe5.ComboBox1.value
    
    Set dataRange = ActiveSheet.Cells.CurrentRegion
    m = dataRange.Columns.count                         'm         : dataSheet�� �ִ� ���� ����
    
    tmp = 0
        For j = 0 To nol
            For i = 1 To m
                If xlist(j) = ActiveSheet.Cells(1, i) Then
                    k1(j) = i  'k1                                 : k1 : ���õ� ������ ���° ���� �ִ���
               
                End If
            Next i
            
            n = ActiveSheet.Cells(1, k1(0)).End(xlDown).row - 1    'n         : ���õ� ������ ����Ÿ ����
        Next j



    Dim checkarray As String
    Dim anovastr As String
    Dim temp As String
    
    Dim cbindstr As String
    
    
    
   
    
      For j = 0 To nol - 1
        checkarray = xlist(j)
        If checkarray = "C" Then
           rinterface.PutArray "Cc", Range(Cells(2, k1(j)), Cells(n + 1, k1(j)))
        Else
        rinterface.PutArray checkarray, Range(Cells(2, k1(j)), Cells(n + 1, k1(j)))
        End If
    
            If j = 0 Then
                cbindstr = checkarray
                anovastr = checkarray
                Else
                cbindstr = cbindstr & "," & checkarray
                anovastr = anovastr & "+" & checkarray
            End If
         
    Next j
    checkarray = xlist(j)
   '  rinterface.PutArray checkarray, Range(Cells(2, k1(j)), Cells(N + 1, k1(j)))
   '  rinterface.PutDataframe "Response2", Range(Cells(1, k1(j)), Cells(N + 1, k1(j)))
     
     Dim strc As String
     
     For q = 1 To n
     If q = 1 Then
     strc = Cells(q + 1, k1(j)).value
     Else
     
     strc = strc & ", " & Cells(q + 1, k1(j)).value
     End If
     
     Next q
    ' MsgBox q - 1
     
     
     
    strc = "Response = c(" & strc & ")"
    ' MsgBox strc
    '  Rinterface.RRun "Responset = c(580, 1090, 1392, 568, 1087, 1380, 570, 1085, 1386, 550, 1070, 1328, 530, 1035, 1312, 579)"
     rinterface.RRun strc
     rinterface.RRun "response(arrayfrac) = Response"
    'Rinterface.RRun "install.packages (" & Chr(34) & "FrF2" & Chr(34) & ")" ' : R ��Ű�� �ʿ����:
    
    'Rinterface.RRun "install.packages (" & Chr(34) & "qualityTools" & Chr(34) & ")" ' : R ��Ű�� �ʿ����:

    
    'Rinterface.rrun "arraytest<-cbind(AP,BP,CP,Response)" 'ok
    'MsgBox "arraytest<-cbind(AP,BP,CP,Response)"
   ' MsgBox "arraytest<-cbind(" & cbindstr & "," & xlist(j) & " )"
    'rinterface.RRun "arraytest<-cbind(" & cbindstr & "," & checkarray & ")"
   
    temp = combinationModule.combstr(xstrlist)
   ' MsgBox temp
    rinterface.RRun "QWE<-as.data.frame(arrayfrac)" 'ok
    rinterface.RRun "lm.3 =lm(" & checkarray & " ~ " & temp & ", data = arrayfrac)"
     rinterface.RRun "lmdata <- as.data.frame(lm.3)"
    j = Me.ListBox2.ListCount
    'MsgBox j & " N"
    
    rinterface.RRun "summary(lm.3)" 'ok
    'MsgBox anovastr
    rinterface.RRun "AnovaModeQ <- aov(lm(" & checkarray & " ~" & temp & ", data = arrayfrac))"
    rinterface.RRun "anova(AnovaModeQ)"     'ok
    'rinterface.RRun "MEPlot(AnovaModeQ, main = paste(" & Chr(34) & "��ȿ����" & Chr(34) & "), ylab=" & Chr(34) & "���" & Chr(34) & ", pch = 15, mgp.ylab = 4, cex.title = 1.5, cex.main = par(" & Chr(34) & "cex.main" & Chr(34) & "), lwd = par(" & Chr(34) & "lwd" & Chr(34) & "), abbrev = " & j & " , select = NULL)" 'abbrev ���μ��� �������
    'rinterface.RRun "MEPlot(AnovaModeQ, main = paste(" & Chr(34) & "��ȿ����" & Chr(34) & "), ylab=" & Chr(34) & "���" & Chr(34) & ", pch = 15, mgp.ylab = 4, cex.title = 1.5, cex.main = par(" & Chr(34) & "cex.main" & Chr(34) & "), lwd = par(" & Chr(34) & "lwd" & Chr(34) & "), abbrev =  3 , select = NULL)" 'abbrev ���μ��� �������
    rinterface.RRun "interactionPlot(arrayfrac, response(arrayfrac), fun = mean, main= " & Chr(34) & "��ȣ�ۿ뵵" & Chr(34) & ", col = 1:2 )"
    rinterface.InsertCurrentRPlot Range("I13"), widthrescale:=0.5, heightrescale:=0.5, closergraph:=True
  '  Rinterface.rrun "checker<-cbind(" & cbindstr & ")"
Unload Me
End Sub


Private Sub UserForm_Click()

End Sub
