VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} framehericl 
   OleObjectBlob   =   "framehericl.frx":0000
   Caption         =   "�����������м�"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8040
   StartUpPosition =   1  '������ ���
   TypeInfoVer     =   59
End
Attribute VB_Name = "framehericl"
Attribute VB_Base = "0{2E7285A5-522A-49CD-B8EF-2ADA59AF3C6B}{9D76CCE9-C567-48F4-8A93-8D27D7C35EF4}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False



Private Sub CB1_Click()
    MoveBtwnListBox Me, "ListBox1", "ListBox3"
    Me.CB1.Visible = False
    Me.CB2.Visible = True
End Sub
Private Sub CB2_Click()
    MoveBtwnListBox Me, "ListBox3", "ListBox1"
    Me.CB1.Visible = True
    Me.CB2.Visible = False
End Sub

Private Sub CB3_Click()
    MoveBtwnListBox Me, "ListBox1", "ListBox2"
End Sub

Private Sub CB4_Click()
    MoveBtwnListBox Me, "ListBox2", "ListBox1"
End Sub


Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    MoveBtwnListBox Me, "ListBox1", "ListBox2"
End Sub
Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    MoveBtwnListBox Me, "ListBox2", "ListBox1"
End Sub
Private Sub CommandButton5_Click()
    Unload Me
End Sub
Private Sub okbtn_Click()
     
    Dim noll As Integer
    Dim nol  As Integer                                  'nol�� �м�����(ListBox2)�� ���������� ��Ÿ����.
    Dim nocl As Integer                                  'nocl�� ����ڰ� ������ ���������� ��Ÿ����.
    Dim rowlist As String
    'Dim nomaxre As Integer                               'nomaxre�� �ִ�ݺ������� ��Ÿ����.
    'Dim noopcl As Integer                                'noopcl�� ������������ ���ϱ� ���� ���� ���� ��Ÿ����. 1���� noopcl������ ������ ���� ������ �׷����� �����Ѵ�.
    Dim dataRange As Range
    Dim i, j As Integer
    'Dim activePt As Long                                 '��� �м��� ���۵Ǵ� �κ��� �����ֱ� ����
    'Dim rng As Range
    Dim k1(20) As Integer       '�м����� 20�� ������ �������� ����
    Dim Heriarray As String
    Dim cbindstr As String      '������ġ�� �ڵ�
    Dim clmet As String
    
    CleanCharts
'====================================================================================
    '''
    '''��꿡 �ʿ��� ������ ����
    '''
 '   noll = Me.ListBox1.ListCount + Me.ListBox2.ListCount
    nol = Me.ListBox2.ListCount
    If Me.ListBox3.ListCount = 1 Then
    rowlist = Me.ListBox3.List(0)
    End If
    nocl = Me.TextBox1.Value
 '   nomaxre = Me.TextBox2.Value
 '   noopcl = Me.TextBox3.Value
    clmet = Me.ComboBox1.Value
'====================================================================================
    '''
    '''������ �������� �ʾ��� ���
    '''
    If nol = 0 Then
        MsgBox "������ ������ �ֽñ� �ٶ��ϴ�.", vbExclamation, "HIST"
    Exit Sub
    ElseIf nol >= 21 Then
        MsgBox "�м������� 20�� ���Ϸ� �����ؾ� �մϴ�.", vbExclamation, "HIST"
    Exit Sub
    End If

'====================================================================================
    '''
    '''public ���� ���� xlist, DataSheet, RstSheet, m, k1, n
    '''

        DataSheet = ActiveSheet.Name                        'DataSheet : Data�� �ִ� Sheet �̸�
        RstSheet = "_���м����_"                         'RstSheet  : ����� �����ִ� Sheet �̸�

    
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

'====================================================================================
    
    rinterface.StartRServer
'    rinterface.PutDataframe "arraytest", Range(Me.RefEdit1)
'    rinterface.RRun "arraytest1<-kmeans(arraytest,3)"
'    rinterface.RRun "arrayre1<-arraytest1$cluster"
    ReDim xlist(nol - 1)                                            'ListBox2�� �ִ� List(j)��° �������� xlist(j)�� �Ҵ�
        For j = 0 To nol - 1
            xlist(j) = ListBox2.List(j)
        Next j
   
    Set dataRange = ActiveSheet.Cells.CurrentRegion
    m = dataRange.Columns.Count                                     'm  : dataSheet�� �ִ� ���� ����

    tmp = 0
        For j = 0 To nol - 1
            For i = 1 To m
                If xlist(j) = ActiveSheet.Cells(1, i) Then
                    k1(j) = i                                       'k1 : ���õ� ������ ���° ���� �ִ���
                  '  tmp = tmp + 1
                End If
            Next i
            
            n = ActiveSheet.Cells(1, k1(0)).End(xlDown).Row - 1     'n  : ���õ� ������ ����Ÿ ����
        Next j
    
        
        For j = 0 To nol - 1
            
            Heriarray = xlist(j)
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            rinterface.PutArray Heriarray, Range(Cells(2, k1(j)), Cells(n + 1, k1(j)))
            
            If j = 0 Then
                cbindstr = Heriarray
            Else
                cbindstr = cbindstr & "," & Heriarray
            End If
         
        Next j
        
'========================================================================================================================================================================
    If Me.ListBox3.ListCount = 1 Then

    
        tmp2 = 0
        For j = 1 To m
            If rowlist = ActiveSheet.Cells(1, j) Then
                k2 = j  'k1                                 : k1 : ���õ� ������ ���° ���� �ִ���
                'tmp2 = tmp2 + 1
            End If
        Next j
        
            n2 = ActiveSheet.Cells(1, k2).End(xlDown).Row - 1    'n         : ���õ� ������ ����Ÿ ����
            
            rinterface.PutArray "test1", Range(Cells(2, k2), Cells(n2 + 1, k2))
            
            rinterface.RRun "test2<-as.character(test1)"
    
 
    End If
 MsgBox cbindstr
'========================================================================================================================================================================
'========================================================================================================================================================================
 Dim hbind As String, hbind1 As String
 
 hbind = "herivar<-cbind(" & cbindstr & ")"
 
 MsgBox hbind
 rinterface.RRun hbind
  
 rinterface.RRun "rownames(herivar)<-test2"

 hbind1 = "heridata<-as.data.frame(herivar)"
 rinterface.RRun hbind1
 'rinterface.RRun "write.csv(heridata,file=" & Chr(34) & "C:/Users/IME20/Desktop/usexample1.csv" & Chr(34) & ")"
 Dim heex As String
 
 heex = "heex1<-dist(heridata, method = " & Chr(34) & "euclidean" & Chr(34) & ")"
 rinterface.RRun heex
 rinterface.GetArray "heex1", Range("����1!G2")
 
 Dim heex1 As String, heex2 As String, heex3 As String, heex4 As String, heex5 As String
 Dim heex_plot As String
 
 
 If clmet = "�ִܿ����" Then
 
 heex1 = "hc <- hclust(heex1, " & Chr(34) & "single" & Chr(34) & ")"
 rinterface.RRun heex1
' heex_plot = "plot(hc)"
' rinterface.RRun heex_plot
 
'========================================================================================================================================================================
'========================================================================================================================================================================
 Else
 
 If clmet = "���忬���" Then
  
 heex2 = "hc <- hclust(dist(heridata), " & Chr(34) & "complete" & Chr(34) & ")'"
 rinterface.RRun heex2
' heex_plot = "plot(hc)"
' rinterface.RRun heex_plot
'========================================================================================================================================================================
'========================================================================================================================================================================
 Else
 
 If clmet = "��տ����" Then
 
 heex3 = "hc <- hclust(dist(heridata), " & Chr(34) & "average" & Chr(34) & ")'"
 rinterface.RRun heex3
' heex_plot = "plot(hc)"
' rinterface.RRun heex_plot
'========================================================================================================================================================================
'========================================================================================================================================================================
 Else
 
 If clmet = "�߽ɿ����" Then
 
 heex4 = "hc <- hclust(dist(heridata), " & Chr(34) & "centroid" & Chr(34) & ")'"
 rinterface.RRun heex4
' heex_plot = "plot(hc)"
' rinterface.RRun heex_plot
'========================================================================================================================================================================
'========================================================================================================================================================================
 Else
 
 If clmet = "�͵忬���" Then
 
 heex5 = "hc <- hclust(dist(heridata), " & Chr(34) & "wad.D" & Chr(34) & ")'"
 rinterface.RRun heex5
' heex_plot = "plot(hc)"
' rinterface.RRun heex_plot
'========================================================================================================================================================================
'========================================================================================================================================================================
End If
End If
End If
End If
End If
heex_plot = "plot(hc, hang = -0.2, check =TRUE)"
rinterface.RRun heex_plot
rinterface.InsertCurrentRPlot Range("����1!G54"), widthrescale:=2, heightrescale:=1.4, closergraph:=True
 Unload Me
 
End Sub

Private Sub UserForm_Initialize()

    Dim myArray As Variant
        ComboBox1.ColumnCount = 1
        myArray = [{"�ִܿ����";"���忬���";"��տ����";"�߽ɿ����";"�͵忬���"}]
        ComboBox1.List = myArray

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
