VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameDMtree 
   OleObjectBlob   =   "frameDMtree.frx":0000
   Caption         =   "�ǻ�������� "
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8385
   StartUpPosition =   1  '������ ���
   TypeInfoVer     =   27
End
Attribute VB_Name = "frameDMtree"
Attribute VB_Base = "0{855D757B-4D3D-4C37-B1F4-9B3BB885FE61}{42F3D43D-79A5-4B33-897A-0B687AF3603D}"
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

Private Sub CB3_Click()
    MoveBtwnListBox Me, "ListBox1", "ListBox3"
End Sub
Private Sub CB4_Click()
    MoveBtwnListBox Me, "ListBox3", "ListBox1"
    End Sub
    
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    MoveBtwnListBox Me, "ListBox1", "ListBox2"
End Sub
Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    MoveBtwnListBox Me, "ListBox2", "ListBox1"
End Sub
Private Sub ListBox3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    MoveBtwnListBox Me, "ListBox3", "ListBox1"
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

    Dim noll As Integer
    Dim nol  As Integer                                  'nol�� �м�����(ListBox2)�� ���������� ��Ÿ����.
    Dim nocl As Integer                                  'nocl�� ����ڰ� ������ ���������� ��Ÿ����.
    'Dim nomaxre As Integer                               'nomaxre�� �ִ�ݺ������� ��Ÿ����.
    'Dim noopcl As Integer                                'noopcl�� ������������ ���ϱ� ���� ���� ���� ��Ÿ����. 1���� noopcl������ ������ ���� ������ �׷����� �����Ѵ�.
    Dim dataRange As Range
        Dim dataRange2 As Range
    Dim i, j, r As Integer
    Dim activePt As Long                                 '��� �м��� ���۵Ǵ� �κ��� �����ֱ� ����
    Dim rng As Range
    Dim k2(20) As Integer       '�м����� 20�� ������ �������� ����
    'Dim Rformula1 As String 'R���õ� �ڵ� ���ڿ��� (�ϴ� 20����..)
    'Dim Rformula2 As String
    Dim a As String
    Dim b As String
    
    Dim kmeansarray As String
    Dim cbindstr As String      '������ġ�� �ڵ�
    Dim xlistx As String
    Dim xlisty As String
   ' Dim findcluster As String
    
    CleanCharts
'====================================================================================
    '''
    '''��꿡 �ʿ��� ������ ����
    '''
    noll = Me.ListBox1.ListCount + Me.ListBox3.ListCount
    nol = Me.ListBox3.ListCount
    'nocl = Me.TextBox1.Value
    'nomaxre = Me.TextBox2.Value
    'noopcl = Me.TextBox3.Value

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
    
       xlist = Me.ListBox2.List(0)
           'xlist2 = Me.ListBox3.List(0)
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
    
    
    
    
    ReDim xlist2(nol - 1)                                            'ListBox2�� �ִ� List(j)��° �������� xlist(j)�� �Ҵ�
        For j = 0 To nol - 1
            xlist2(j) = ListBox3.List(j)
        Next j
   
    Set dataRange2 = ActiveSheet.Cells.CurrentRegion
    m2 = dataRange.Columns.Count                                     'm  : dataSheet�� �ִ� ���� ����

    tmp2 = 0
        For j = 0 To nol - 1
            For r = 1 To m2
                If xlist2(j) = ActiveSheet.Cells(1, r) Then
                    k2(j) = r                                       'k1 : ���õ� ������ ���° ���� �ִ���
                    tmp2 = tmp2 + 1
                End If
            Next r
            
            n2 = ActiveSheet.Cells(1, k2(0)).End(xlDown).Row - 1     'n  : ���õ� ������ ����Ÿ ����
        Next j

 
        rinterface.RRun "install.packages (" & Chr(34) & "tree" & Chr(34) & ")"
    rinterface.RRun "require (tree)"
rinterface.RRun "install.packages (" & Chr(34) & "party" & Chr(34) & ")"
    rinterface.RRun "require (party)"

    
        For j = 0 To nol - 1
            
            kmeansarray = xlist2(j)
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            rinterface.PutArray kmeansarray, Range(Cells(2, k2(j)), Cells(n + 1, k2(j)))
                If j = 0 Then
                    cbindstr = kmeansarray
                Else
                    cbindstr = cbindstr & "," & kmeansarray
                End If
         
        Next j
 MsgBox cbindstr
 

'========================================================================================================================================================================
    Dim Rf1 As String, Rf2 As String, Rf3 As String, Rf4 As String  'r�ڵ� string������ ����
    Dim Rf5 As String, Rf6 As String, Rf7 As String, Rf8 As String
    Dim Rf9 As String, Rf10 As String, Rf11 As String, Rf12 As String
    Dim Rf13 As String, Rf14 As String, Rf15 As String, Rf16 As String
    Dim ssecl0 As String, ssecl1 As String, ssecl2 As String, ssecl3 As String, ssecl4 As String, ssecl5 As String
    Dim strclbind As String
    
'========================================================================================================================================================================
    Rf1 = " Ndata <- cbind(" & cbindstr & ")" 'cbindstr�� �޾ƿ� �������� �ϳ��� ������ �����Ѵ�(arraytest)

    rinterface.PutArray "arraytest", Range(Cells(2, k1), Cells(n + 1, k1))
    b = "N1<-cbind(Ndata,arraytest)"
   rinterface.RRun b
    Rf2 = "NNdata <- as.data.frame(N1)"                'kmeansvar�� data.frame���� ��ȯ�Ѵ�.
    Rf3 = " NNNdata <- data.frame(Ndata, arraytest)"
    
     
    rinterface.RRun Rf1                                           '���� ��ġ��
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    rinterface.RRun Rf2                                           '���� ��ġ�� �� ������ ���������� ��ȯ
    rinterface.RRun Rf3
    
    a = " ir.tr = tree(arraytest~., NNNdata)"
    MsgBox a

    rinterface.RRun a
    
        rinterface.RRun "plot(ir.tr)"
    rinterface.RRun "text(ir.tr, all=TRUE)"
    
    
    rinterface.RRun "nctree=ctree(arraytest~., data=NNNdata)"
rinterface.RRun "Print (nctree)"
rinterface.RRun "plot(nctree, type= " & Chr(34) & " simple" & Chr(34) & " )"
    
    
   
'========================================================================================================================================================================
'========================================================================================================================================================================
    'If CheckBox11.Value = True Then                               '�׷������


   ' End If
'========================================================================================================================================================================
'========================================================================================================================================================================


    Unload Me
  
 
End Sub
Sub CleanCharts()
    Dim chrt As Picture
    On Error Resume Next
    For Each chrt In ActiveSheet.Pictures
        chrt.Delete
    Next chrt
End Sub





Private Sub UserForm_Click()

End Sub
