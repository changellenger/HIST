VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} framekmeanscl_2 
   OleObjectBlob   =   "framekmeanscl_2.frx":0000
   Caption         =   "K-Means �����м�"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8280
   StartUpPosition =   1  '������ ���
   TypeInfoVer     =   157
End
Attribute VB_Name = "framekmeanscl_2"
Attribute VB_Base = "0{E9B14135-3403-4F56-9726-E8BD5D211304}{D764CFA1-EF7B-4645-96C7-48C1BD465435}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CheckBox1_Click()
    If CheckBox1.Value = True Then CheckBox5.Value = False
End Sub

Private Sub CheckBox11_Click()
    Frame5.Visible = True
    Me.ListBox2.Height = 82.1
    If CheckBox11.Value = False Then
        Frame5.Visible = False
        Me.ListBox2.Height = 181.85
    Else
        Frame5.Visible = True
        Me.ListBox2.Height = 82.1
    End If
       
End Sub


Private Sub CheckBox8_Click()
    Me.TextBox3.Enabled = True
    
    If Me.CheckBox8.Value = False Then
        TextBox3.Enabled = False
    Else
        If Me.CheckBox8.Value = True Then
        TextBox3.Enabled = True
        End If
    End If
    
End Sub

Private Sub CommandButton10_Click()
        Dim i As Integer
    i = 0
    If Me.ListBox5.ListCount = 0 Then
        Do While i <= Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(i) = True Then
               Me.ListBox5.AddItem Me.ListBox1.List(i)
               Me.ListBox1.RemoveItem (i)
               Me.CommandButton10.Visible = False
               Me.CommandButton11.Visible = True
               Exit Sub
            End If
            i = i + 1
        Loop
    End If
End Sub

Private Sub CommandButton11_Click()
        If Me.ListBox5.ListCount <> 0 Then
        Me.ListBox1.AddItem ListBox5.List(0)
        Me.ListBox5.RemoveItem (0)
        Me.CommandButton10.Visible = True
        Me.CommandButton11.Visible = False
    End If
End Sub

Private Sub CommandButton5_Click()
    Unload Me
End Sub
Private Sub CB1_Click()
    MoveBtwnListBox Me, "ListBox1", "ListBox2"
End Sub
Private Sub CB2_Click()
    MoveBtwnListBox Me, "ListBox2", "ListBox1"
End Sub
Private Sub CommandButton8_Click()
   Dim i As Integer
    i = 0
    If Me.ListBox4.ListCount = 0 Then
        Do While i <= Me.ListBox2.ListCount - 1
            If Me.ListBox2.Selected(i) = True Then
               Me.ListBox4.AddItem Me.ListBox2.List(i)
           
               Exit Sub
            End If
            i = i + 1
        Loop
    End If
End Sub
Private Sub CommandButton9_Click()
   Dim i As Integer
    i = 0
    If Me.ListBox3.ListCount = 0 Then
        Do While i <= Me.ListBox2.ListCount - 1
            If Me.ListBox2.Selected(i) = True Then
               Me.ListBox3.AddItem Me.ListBox2.List(i)
           
               Exit Sub
            End If
            i = i + 1
        Loop
    End If
End Sub
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    MoveBtwnListBox Me, "ListBox1", "ListBox2"
End Sub
Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    MoveBtwnListBox Me, "ListBox2", "ListBox1"
End Sub
Private Sub CheckBox5_Click()

    If CheckBox5.Value = True Then CheckBox1.Value = False
    If Me.CheckBox5.Value = True Then Me.TextBox2.Enabled = True
    If Me.CheckBox5.Value = False Then Me.TextBox2.Enabled = False

        
End Sub


Private Sub okbtn_Click()

    Dim noll As Integer
    Dim nol  As Integer                                  'nol�� �м�����(ListBox2)�� ���������� ��Ÿ����.
    Dim nocl As Integer                                  'nocl�� ����ڰ� ������ ���������� ��Ÿ����.
    Dim nomaxre As Integer                               'nomaxre�� �ִ�ݺ������� ��Ÿ����.
    Dim noopcl As Integer                                'noopcl�� ������������ ���ϱ� ���� ���� ���� ��Ÿ����. 1���� noopcl������ ������ ���� ������ �׷����� �����Ѵ�.
    Dim dataRange As Range
    Dim dataRange1 As Range
    Dim i, j As Integer
    Dim activePt As Long                                 '��� �м��� ���۵Ǵ� �κ��� �����ֱ� ����
    Dim rng As Range
    Dim k1(30) As Integer       '�м����� 20�� ������ �������� ����
    Dim a As Integer
    Dim kmeansarray As String
    Dim cbindstr As String      '������ġ�� �ڵ�
    Dim xlistx As String
    Dim xlisty As String
    Dim findcluster As String
    Dim tablestr As String
    
    rinterface.StartRServer     'RExcel ��������
    
    CleanCharts                 '��µ� �׷��� �ߺ� ����
'====================================================================================
    '''
    '''��꿡 �ʿ��� ������ ����
    '''
    noll = Me.ListBox1.ListCount + Me.ListBox2.ListCount + Me.ListBox5.ListCount
    MsgBox noll
    nol = Me.ListBox2.ListCount
    nocl = Me.TextBox1.Value
    nomaxre = Me.TextBox2.Value
    noopcl = Me.TextBox3.Value

'====================================================================================
    '''
    '''������ �������� �ʾ��� ���
    '''
    If nol = 0 Then
        MsgBox "������ ������ �ֽñ� �ٶ��ϴ�.", vbExclamation, "HIST"
    Exit Sub
    'ElseIf nol >= 21 Then
    '    MsgBox "�м������� 20�� ���Ϸ� �����ؾ� �մϴ�.", vbExclamation, "HIST"
    'Exit Sub
    End If
'====================================================================================
    '''
    '''
    '''

    
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
            
            kmeansarray = xlist(j)
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            rinterface.PutArray kmeansarray, Range(Cells(2, k1(j)), Cells(n + 1, k1(j)))
                If j = 0 Then
                    cbindstr = kmeansarray
                Else
                    cbindstr = cbindstr & "," & kmeansarray
                End If
         
        Next j
    
'MsgBox cbindstr




        If ListBox5.ListCount <> 0 Then
        
            Set dataRange1 = ActiveSheet.Cells.CurrentRegion
            m2 = dataRange1.Columns.Count
            
           tablestr = Me.ListBox5.List(0)
           'MsgBox tablestr
           For i = 1 To m2
            If tablestr = ActiveSheet.Cells(1, i) Then
                k2 = i
            End If
           Next i
           
           n2 = ActiveSheet.Cells(1, k2).End(xlDown).Row - 1
           
           rinterface.PutArray "tablestr", Range(Cells(2, k2), Cells(n2 + 1, k2))
        End If

    
'========================================================================================================================================================================
    Dim Rf1 As String, Rf2 As String, Rf3 As String, Rf4 As String  'r�ڵ� string������ ����
    Dim Rf5 As String, Rf5_1 As String, Rf6 As String, Rf7 As String
    Dim Rf8 As String, Rf9 As String, Rf10 As String, Rf11 As String
    Dim Rf12 As String, Rf13 As String, Rf14 As String, Rf15 As String
    Dim Rf16 As String
    Dim ssecl0 As String, ssecl1 As String, ssecl2 As String, ssecl3 As String, ssecl4 As String
    Dim strclbind As String
    
'========================================================================================================================================================================
    Rf1 = "kmeansvar <- cbind(" & cbindstr & ")"                  'cbindstr�� �޾ƿ� �������� �ϳ��� ������ �����Ѵ�(arraytest)
    Rf2 = "kmeansdata <- as.data.frame(kmeansvar)"                'kmeansvar�� data.frame���� ��ȯ�Ѵ�.
    Rf3 = "kmeansresult <- kmeans(kmeansdata," & nocl & ", algorithm = " & Chr(34) & "MacQueen" & Chr(34) & ")"             'kmeansresult�� kmeans�м��� �ǽ��� ����� ����
    Rf4 = "kmeansresult <- kmeans(kmeansdata,nocl,max=" & nomaxre & ")"   'kmeansresult�� �ִ�ݺ������� ������ �м� ����� ����
    Rf5 = "kmeansdata <- scale(kmeansdata)"                       'scale���� ǥ��ȭ��Ų��
    Rf5_1 = "kmeansdata <- as.data.frame(kmeansdata)"
    Rf6 = "kmeansre1 <- kmeansresult$cluster"
    Rf7 = "kmeansre2 <- kmeansresult$center"
    Rf8 = "kmeansre3 <- kmeansresult$totss"
    Rf9 = "kmeansre4 <- kmeansresult$withinss"
    Rf10 = "kmeansre5 <- kmeansresult$tot.withinss"
    Rf11 = "kmeansre6 <- kmeansresult$betweenss"
    Rf12 = "kmeansre7 <- kmeansresult$size"
    Rf13 = "kmeansre8 <- kmeansresult$iter"
    Rf14 = "kmeansre9 <- kmeansresult$ifault"
    Rf15 = "kmeanssumm <-summary(kmeansdata)"
    
    
    ssecl0 = "modelsse <- 0"
    ssecl2 = "findclustno = function(data = kmeansdata, groups = 1 : " & noopcl & "){plot(groups, modelsse, type =" & Chr(34) & "b" & Chr(34) & ", main = " & Chr(34) & "������ ���� ����" & Chr(34) & ", xlab = " & Chr(34) & "������ ��" & Chr(34) & ", ylab = " & Chr(34) & "������ �Ÿ� ������" & Chr(34) & ")}"
    ssecl3 = "findclustno(kmeansdata)"

'========================================================================================================================================================================
'========================================================================================================================================================================
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    rinterface.StartRServer
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    rinterface.RRun Rf1                                           '���� ��ġ��
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    rinterface.RRun Rf2                                           '���� ��ġ�� �� ������ ���������� ��ȯ
        
    If CheckBox6.Value = True Then                                'ǥ��ȭ üũ�ڽ� �Ǿ����� ��
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        rinterface.RRun Rf5                                       'ǥ��ȭ �ǽ�
        rinterface.RRun Rf5_1                                     'ǥ��ȭ �� ������ ���������� ��ȯ
        
    End If

 '======================================================================================================================================================================== ����ϴ��ڵ�
 '========================================================================================================================================================================
    If CheckBox5.Value = True Then                                '�ִ�ݺ����� ������ �ݺ����
        
        TextBox2.Enabled = True
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        rinterface.RRun Rf4
    
    End If
        
    If CheckBox1.Value = True Then                                '�Ϲݰ��(�ݺ��� ������ ����)
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        rinterface.RRun Rf3
    'MsgBox Rf3
    
    End If
'========================================================================================================================================================================
'========================================================================================================================================================================
    If CheckBox8.Value = True Then                                '������

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    rinterface.RRun ssecl0
'====================================================
   For sseno = 1 To noopcl                                         '�� �������� �������� ����Ͽ� �Է��Ѵ�. (sseno�� TextBox3�� ���� �޾ƿ´�)

          ssecl1 = "modelsse[" & sseno & "] <- sum(kmeans(kmeansdata, centers = " & sseno & " )$withinss)"
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          rinterface.RRun ssecl1
          
          ssecl4 = "Cluster" & sseno & "�� <-modelsse[" & sseno & "]"
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          rinterface.RRun ssecl4
                    
         ' MsgBox ssecl1
         ' MsgBox ssecl4
            Next sseno
'=====================================================              '�� ������ ������ ������ ���������� �����ϱ� ���� �غ�
        Dim strcl As String
        
        For i = 1 To sseno - 1
            
            If i = 1 Then
                
                strcl = " Cluster" & i & "�� "
            
            Else
                
                strcl = strcl & ", Cluster" & i & "�� "
            
            End If
            
        Next i
        
            strclbind = "strex1 <- cbind(" & strcl & ")"
        
            rinterface.RRun strclbind
            rinterface.GetArray "strex1", Range(Cells(24, noll + 10), Cells(24, noll + 10))
        
        For i = 1 To sseno - 1
            Range(Cells(24, noll + i + 9), Cells(24, noll + i + 9)).NumberFormat = "0.00"
        Next i
        
        
            'MsgBox strcl
        
        ReDim clustnoarr(sseno - 1)
        
        For i = 1 To sseno - 1
            
            clustnoarr(i) = "Cluster" & i & "��"
        
            Range(Cells(23, noll + i + 9), Cells(23, noll + i + 9)).Value = "" & clustnoarr(i) & ""
            Range(Cells(23, noll + i + 9), Cells(23, noll + i + 9)).Font.Bold = True
            Range(Cells(23, noll + i + 9), Cells(23, noll + i + 9)).Interior.Color = RGB(220, 238, 130)
            Range(Cells(23, noll + i + 9), Cells(23, noll + i + 9)).Cells.ColumnWidth = 15
            
            Range(Cells(24, noll + i + 9), Cells(24, noll + i + 9)).NumberFormat = "0.00"
            Range(Cells(24, noll + i + 9), Cells(24, noll + i + 9)).HorizontalAlignment = xlLeft
                    
            'MsgBox clustnoarr(i)
        Next i
              
        
        rinterface.RRun ssecl2                               '�������� �� �е��׷��� ���(�Լ��� �̿�)
        
        rinterface.RRun ssecl3                               '������ �Լ��� kmeansdata(�Էµ�����) �Է�
        
        rinterface.InsertCurrentRPlot Range(Cells(2, noll + 10), Cells(2, noll + 10)), widthrescale:=0.6, heightrescale:=0.6, closergraph:=True

    End If
'========================================================================================================================================================================
'========================================================================================================================================================================
    If CheckBox7.Value = True Then                                '������跮
        
        rinterface.RRun Rf7                                       'Rf7  = "kmeansre2 <- kmeansresult$center"
        rinterface.RRun Rf12                                      'Rf12 = "kmeansre7 <- kmeansresult$size"
        rinterface.RRun Rf15                                      'Rf15 = "kmeanssumm <-summary(kmeansdata)"




        rinterface.RRun "iris.ta<-table(tablestr,kmeansresult$cluster)"
        
       ' rinterface.RRun "iris.ta1<-as.data.frame(iris.ta)"
        rinterface.GetArray "iris.ta", Range(Cells(37, noll + 3), Cells(37, noll + 3)), , , , True, True
        
        
        Range(Cells(36, noll + 4), Cells(36, noll + nocl + 3)).Merge
        Range(Cells(36, noll + 4), Cells(36, noll + 4)).Value = "����"
        Range(Cells(36, noll + 4), Cells(36, noll + 4)).HorizontalAlignment = xlCenter
        Range(Cells(36, noll + 4), Cells(36, noll + 4)).Font.Bold = True
        For q = 0 To nocl - 1
        
        Range(Cells(37, noll + q + 4), Cells(37, noll + q + 4)).Interior.Color = RGB(220, 238, 130)
        
        Next q

'''                                 '''                                 '''�׷��� ����� ��, ��跮 ��ġ
'''                                 '''                                 '''
'''                                 '''                                 '''
        If CheckBox11.Value = True Then
        
        rinterface.GetArray "kmeansre2", Range(Cells(24, noll + 3), Cells(24, noll + 3))
        
        For q = 0 To nol - 1
            xlist(q) = ListBox2.List(q)
        Range(Cells(23, noll + q + 3), Cells(23, noll + q + 3)).Value = "" & xlist(q) & ""
        Range(Cells(23, noll + q + 3), Cells(23, noll + q + 3)).Font.Bold = True
        Range(Cells(23, noll + q + 3), Cells(23, noll + q + 3)).Interior.Color = RGB(220, 238, 130)
        Range(Cells(23, noll + q + 3), Cells(23, noll + q + 3)).Cells.ColumnWidth = 15
        Range(Cells(24, noll + q + 3), Cells(24 + nocl, noll + q + 3)).NumberFormat = "0.00"
        Range(Cells(24, noll + q + 3), Cells(24 + nocl, noll + q + 3)).HorizontalAlignment = xlLeft
        Next q
     '''''''''''           rinterface.GetArray "strex1", Range(Cells(25, noll + 13), Cells(25, noll + 13))
      '''''''''''          rinterface.InsertCurrentRPlot Range(Cells(2, noll + 3), Cells(2, noll + 3)), widthrescale:=0.6, heightrescale:=0.6, closergraph:=True
        For p = 1 To nocl

        Range(Cells(23 + p, noll + 2), Cells(23 + p, noll + 2)).Value = "" & p & ""
        Range(Cells(23 + p, noll + 2), Cells(23 + p, noll + 2)).Font.Bold = True
        Next p
        
        Range(Cells(23, noll + q + 3), Cells(23, noll + q + 3)).Value = "������ ũ��"
        Range(Cells(23, noll + q + 3), Cells(23, noll + q + 3)).Font.Bold = True
        Range(Cells(23, noll + q + 3), Cells(23, noll + q + 3)).Interior.Color = RGB(220, 238, 130)
        Range(Cells(23, noll + q + 3), Cells(23, noll + q + 3)).Cells.ColumnWidth = 15
        Range(Cells(23, noll + q + 3), Cells(23, noll + q + 3)).HorizontalAlignment = xlLeft
        rinterface.GetArray "kmeansre7", Range(Cells(24, noll + q + 3), Cells(24, noll + q + 3))
        rinterface.GetArray "kmeanssumm", Range(Cells(24 + nocl, noll + 3), Cells(24 + nocl, noll + 3))
'=====================================================
        Else

'''                                 '''                                 '''�׷��� ��� ���� �� ��跮 ��ġ
'''                                 '''                                 '''
'''                                 '''                                 '''
        rinterface.GetArray "kmeansre2", Range(Cells(5, noll + 3), Cells(5, noll + 3))
        
        For q = 0 To nol - 1
            xlist(q) = ListBox2.List(q)
        Range(Cells(4, noll + q + 3), Cells(4, noll + q + 3)).Value = "" & xlist(q) & ""
        Range(Cells(4, noll + q + 3), Cells(4, noll + q + 3)).Font.Bold = True
        Range(Cells(4, noll + q + 3), Cells(4, noll + q + 3)).Interior.Color = RGB(220, 238, 130)
        Range(Cells(4, noll + q + 4), Cells(4, noll + q + 4)).Cells.ColumnWidth = 15
        Range(Cells(5, noll + q + 3), Cells(5 + nocl, noll + q + 3)).NumberFormat = "0.00"
        Range(Cells(5, noll + q + 3), Cells(5 + nocl, noll + q + 3)).HorizontalAlignment = xlLeft
        Next q

        For p = 1 To nocl

        Range(Cells(4 + p, noll + 2), Cells(4 + p, noll + 2)).Value = "" & p & ""
        Range(Cells(4 + p, noll + 2), Cells(4 + p, noll + 2)).Font.Bold = True
        Next p
        
        Range(Cells(4, noll + q + 3), Cells(4, noll + q + 3)).Value = "������ ũ��"
        Range(Cells(4, noll + q + 3), Cells(4, noll + q + 3)).Font.Bold = True
        Range(Cells(4, noll + q + 3), Cells(4, noll + q + 3)).Interior.Color = RGB(220, 238, 130)
        rinterface.GetArray "kmeansre7", Range(Cells(5, noll + q + 3), Cells(5, noll + q + 3))
        rinterface.GetArray "kmeanssumm", Range(Cells(5 + nocl, noll + 3), Cells(5 + nocl, noll + 3))
        End If

    End If
'========================================================================================================================================================================
'========================================================================================================================================================================
    If CheckBox11.Value = True Then                               '�׷������


'''                                 '''                                 '''
'''                                 '''                                 '''
'''                                 '''                                 '''
        If Me.ListBox3.ListCount = 0 Then
            MsgBox "�׷��� X�� ������ ������ �ֽñ� �ٶ��ϴ�.", vbExclamation, "HIST"
            Exit Sub
        End If
        If Me.ListBox4.ListCount = 0 Then
            MsgBox "�׷��� Y�� ������ ������ �ֽñ� �ٶ��ϴ�.", vbExclamation, "HIST"
            Exit Sub
        End If
'''                                 '''                                 '''
'''                                 '''                                 '''
'''                                 '''                                 '''
        

        

        
        xlistx = Me.ListBox3.List(0)                               'ListBox3�� �ִ� �������� xlistx�� �Ҵ�
        xlisty = Me.ListBox4.List(0)                               'ListBox4�� �ִ� �������� xlisty�� �Ҵ�
        Dim Rf17 As String, Rf18 As String, Rf19 As String

    
        Rf17 = "plot(kmeansdata[c(" & Chr(34) & "" & xlistx & "" & Chr(34) & "," & Chr(34) & "" & xlisty & "" & Chr(34) & ")],col=kmeansresult$cluster,main = " & Chr(34) & "�����м� �׷���" & Chr(34) & ")"
        Rf18 = "points(kmeansresult$centers[,c(" & Chr(34) & "" & xlistx & "" & Chr(34) & "," & Chr(34) & "" & xlisty & "" & Chr(34) & ")], col=1:" & nocl & ",pch=8,cex=2)"

        'MsgBox Rf17
        'MsgBox Rf18
        rinterface.RRun Rf17
        rinterface.RRun Rf18
        rinterface.InsertCurrentRPlot Range(Cells(2, noll + 3), Cells(2, noll + 3)), widthrescale:=0.6, heightrescale:=0.6, closergraph:=True
        rinterface.RRun "win.graph()"
        rinterface.RRun Rf17
        rinterface.RRun Rf18
        'MsgBox noll
        

        
    End If
'========================================================================================================================================================================
'========================================================================================================================================================================
    
    If CheckBox12.Value = True Then                               '�Ƿ翧 ��� ����                              '
        Rf19 = "plot(silhouette(kmeansresult$cluster, dist = dist(kmeansdata)), col = 1:" & nocl & ",main = " & Chr(34) & "Silhouette Clustering" & Chr(34) & ")"
        'MsgBox Rf19
        rinterface.RRun "install.packages(" & Chr(34) & "cluster" & Chr(34) & ")"
        rinterface.RRun "require(cluster)"
       ' rinterface.RRun "library(cluster)"
        rinterface.RRun Rf19
        rinterface.InsertCurrentRPlot Range(Cells(36, noll + 10), Cells(36, noll + 10)), widthrescale:=0.6, heightrescale:=0.6, closergraph:=True
        rinterface.RRun "win.graph()"
        rinterface.RRun Rf19
       If CheckBox7.Value = True Then
       
        Range(Cells(55, noll + 10), Cells(55, noll + 10)).Value = "�Ƿ翧 ����� ����� 1�� ����� ���� �ùٸ� �м��� �ǽõǾ����� ���մϴ�."
        Range(Cells(55, noll + 10), Cells(55, noll + 13)).Interior.Color = RGB(220, 238, 130)
       Else
        Range(Cells(55, noll + 10), Cells(55, noll + 10)).Value = "�Ƿ翧 ����� ����� 1�� ����� ���� �ùٸ� �м��� �ǽõǾ����� ���մϴ�."
        Range(Cells(55, noll + 10), Cells(55, noll + 17)).Interior.Color = RGB(220, 238, 130)
       End If
    End If
       ' Range(Cells(36, noll + 10), Cells(36, noll + 12))
'========================================================================================================================================================================
'========================================================================================================================================================================
    If CheckBox9.Value = True Then                                '�Ҽӱ��� ���
    For q = 1 To noll + 1
        If Range(Cells(1, q), Cells(1, q)).Value <> "�Ҽӱ���" Then '������ �ڵ�
            If q = noll + 1 Then
            Range(Cells(1, q), Cells(1, q)).Value = "�Ҽӱ���"
            Range(Cells(1, q), Cells(1, q)).Font.Bold = True
            Range(Cells(1, q), Cells(1, q)).Interior.Color = RGB(220, 238, 130)
            
            rinterface.RRun Rf6
            rinterface.GetArray "kmeansre1", Range(Cells(2, q), Cells(2, q))
             End If
             
                       
        Else    '�����ҽ� �ڵ�
  
            rinterface.RRun Rf6
            rinterface.GetArray "kmeansre1", Range(Cells(2, q), Cells(2, q))
            Exit For
        End If
    Next q
     
    End If



    Range(Cells(34, noll + 9), Cells(34, noll + 23)).Borders(xlEdgeTop).LineStyle = xlContinuous
    Range(Cells(34, noll + 9), Cells(34, noll + 23)).Borders(xlEdgeTop).Color = RGB(34, 116, 34)
    Range(Cells(34, noll + 9), Cells(34, noll + 23)).Borders(xlEdgeTop).Weight = 3
    

    Range(Cells(2, noll + 8), Cells(60, noll + 8)).Borders(xlEdgeRight).LineStyle = xlContinuous
    Range(Cells(2, noll + 8), Cells(60, noll + 8)).Borders(xlEdgeRight).Color = RGB(34, 116, 34)
    Range(Cells(2, noll + 8), Cells(60, noll + 8)).Borders(xlEdgeRight).Weight = 3
    
    
    Unload Me

 
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

Private Sub UserForm_Click()

End Sub
