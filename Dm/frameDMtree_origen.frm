VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameDMtree_origen 
   OleObjectBlob   =   "frameDMtree_origen.frx":0000
   Caption         =   "의사결정나무"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7890
   StartUpPosition =   1  '소유자 가운데
   TypeInfoVer     =   27
End
Attribute VB_Name = "frameDMtree_origen"
Attribute VB_Base = "0{DD8DC8DA-609C-4FD4-963D-6AF62C9DDBD7}{6F5F2DA8-0FDD-4EE6-AB21-85B6AF0F59DA}"
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
    
Private Sub CommandButton5_Click()
Unload Me
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
   If arrName(i) <> "" Then                     '빈칸제거
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
    Dim nol  As Integer                                  'nol은 분석변수(ListBox2)의 변수개수를 나타낸다.
    Dim nocl As Integer                                  'nocl은 사용자가 지정한 군집개수를 나타낸다.
    'Dim nomaxre As Integer                               'nomaxre은 최대반복계산수를 나타낸다.
    'Dim noopcl As Integer                                'noopcl은 최적군집수를 구하기 위한 설정 값을 나타낸다. 1부터 noopcl까지의 군집에 대해 제곱합 그래프를 산출한다.
    Dim dataRange As Range
        Dim dataRange2 As Range
    Dim i, j, r As Integer
    Dim activePt As Long                                 '결과 분석이 시작되는 부분을 보여주기 위함
    Dim rng As Range
    Dim k2(20) As Integer       '분석변수 20개 까지만 가능으로 지정
    'Dim Rformula1 As String 'R관련된 코드 문자열값 (일단 20개만..)
    'Dim Rformula2 As String
    Dim a As String
    Dim b As String
    
    Dim kmeansarray As String
    Dim cbindstr As String      '변수합치는 코드
    Dim xlistx As String
    Dim xlisty As String
   ' Dim findcluster As String
    
    CleanCharts
'====================================================================================
    '''
    '''계산에 필요한 값들을 정의
    '''
    noll = Me.ListBox1.ListCount + Me.ListBox3.ListCount
    nol = Me.ListBox3.ListCount
    'nocl = Me.TextBox1.Value
    'nomaxre = Me.TextBox2.Value
    'noopcl = Me.TextBox3.Value

'====================================================================================
    '''
    '''변수를 선택하지 않았을 경우
    '''
    If nol = 0 Then
        MsgBox "변수를 선택해 주시기 바랍니다.", vbExclamation, "HIST"
    Exit Sub
    ElseIf nol >= 21 Then
        MsgBox "분석변수는 20개 이하로 지정해야 합니다.", vbExclamation, "HIST"
    Exit Sub
    End If
'====================================================================================
    '''
    '''public 변수 선언 xlist, DataSheet, RstSheet, m, k1, n
    '''

        DataSheet = ActiveSheet.Name                        'DataSheet : Data가 있는 Sheet 이름
        RstSheet = "_통계분석결과_"                         'RstSheet  : 결과를 보여주는 Sheet 이름

    
    '맨위에 입력
    'On Error GoTo Err_delete
    Dim val3535 As Long '초기위치 저장할 공간'
    Dim s3535 As Worksheet
            val3535 = 2
        For Each s3535 In ActiveWorkbook.Sheets
            If s3535.Name = RstSheet Then
                val3535 = Sheets(RstSheet).Cells(1, 1).Value
            End If
        Next s3535  '시트가 이미있으면 출력 위치 저장을하고, 없으면 2을 저장한다.

'====================================================================================
    
'    rinterface.StartRServer
'    rinterface.PutDataframe "arraytest", Range(Me.RefEdit1)
'    rinterface.RRun "arraytest1<-kmeans(arraytest,3)"
'    rinterface.RRun "arrayre1<-arraytest1$cluster"
    
       xlist = Me.ListBox2.List(0)
           'xlist2 = Me.ListBox3.List(0)
      Set dataRange = ActiveSheet.Cells.CurrentRegion
    m = dataRange.Columns.Count                         'm         : dataSheet에 있는 변수 개수
    
    tmp = 0
    For i = 1 To m
        If xlist = ActiveSheet.Cells(1, i) Then
            k1 = i  'k1                                 : k1 : 선택된 변수가 몇번째 열에 있는지
            tmp = tmp + 1
        End If
    Next i
    
    n = ActiveSheet.Cells(1, k1).End(xlDown).Row - 1    'n         : 선택된 변수의 데이타 갯수
    
   ' rng = Range(Cells(2, k1), Cells(N + 1, k1))
    
    
    
    
    ReDim xlist2(nol - 1)                                            'ListBox2에 있는 List(j)번째 변수명을 xlist(j)에 할당
        For j = 0 To nol - 1
            xlist2(j) = ListBox3.List(j)
        Next j
   
    Set dataRange2 = ActiveSheet.Cells.CurrentRegion
    m2 = dataRange.Columns.Count                                     'm  : dataSheet에 있는 변수 개수

    tmp2 = 0
        For j = 0 To nol - 1
            For r = 1 To m2
                If xlist2(j) = ActiveSheet.Cells(1, r) Then
                    k2(j) = r                                       'k1 : 선택된 변수가 몇번째 열에 있는지
                    tmp2 = tmp2 + 1
                End If
            Next r
            
            n2 = ActiveSheet.Cells(1, k2(0)).End(xlDown).Row - 1     'n  : 선택된 변수의 데이타 갯수
        Next j


        Dim tms As String
        
    
        For j = 0 To nol - 1
        
            
            kmeansarray = xlist2(j)
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
           ' rinterface.PutArray kmeansarray, Range(Cells(2, k2(j)), Cells(n + 1, k2(j)))
            rinterface.PutDataframe kmeansarray, Range(Cells(1, k2(j)), Cells(n + 1, k2(j)))
            
            'tms = kmeansarray & " <- as.numeric( " & kmeansarray & " )"
            'MsgBox tms
            
            'rinterface.RRun tms
                If j = 0 Then
                    cbindstr = kmeansarray
                Else
                    cbindstr = cbindstr & "," & kmeansarray
                End If
         
        Next j
 MsgBox cbindstr
 
    rinterface.StartRServer
    rinterface.RRun "install.packages (" & Chr(34) & "tree" & Chr(34) & ")"
    rinterface.RRun "require (tree)"
    
'========================================================================================================================================================================
    Dim Rf1 As String, Rf2 As String, Rf3 As String, Rf4 As String  'r코드 string값으로 선언
    Dim Rf5 As String, Rf6 As String, Rf7 As String, Rf8 As String
    Dim Rf9 As String, Rf10 As String, Rf11 As String, Rf12 As String
    Dim Rf13 As String, Rf14 As String, Rf15 As String, Rf16 As String
    Dim ssecl0 As String, ssecl1 As String, ssecl2 As String, ssecl3 As String, ssecl4 As String, ssecl5 As String
    Dim strclbind As String
    Dim strname As String
    
    
'========================================================================================================================================================================
   
   strname = xlist
   
   
   rinterface.PutArray strname, Range(Cells(2, k1), Cells(n + 1, k1))
    Rf1 = " NARdata <- cbind(" & cbindstr & "," & strname & ")" 'cbindstr로 받아온 변수들을 하나의 변수에 지정한다(arraytest)

    
    'b = "N1<-cbind(Ndata3,arraytest3)"
 '  rinterface.RRun b
    Rf2 = "NNARdata <- as.data.frame(NARdata)"                ' Ndata를 data.frame으로 변환한다.
   ' Rf3 = " NNNdata <- data.frame(Ndata, arraytest)"
     
    rinterface.RRun Rf1                                           '변수 합치기
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    rinterface.RRun Rf2                                           '변수 합치기 후 데이터 프레임으로 변환
   ' Rinterface.rrun Rf3
    
    a = "ir.tr = tree(" & strname & "~., NNARdata)"
    'a = "ir.tr = tree(Species ~., iris)"
    
    MsgBox a

    rinterface.RRun a
    
    rinterface.RRun "plot(ir.tr)"
    rinterface.RRun "text(ir.tr, all=TRUE)"
   
'============================================== 가지치기 이전의 나무=============================================
  '  Dim btr As Integer
   ' btr = Me.TextBox1
   ' Rinterface.rrun "fin.tr=prune.misclass(ir.tr, best=" & btr & " )"
   ' Rinterface.rrun "plot(fin.tr)"
   ' Rinterface.rrun "text(fin.tr, all=TRUE)"
'============================================== 가지치기 개수설정=============================================

    'rinterface.rrun "ir.trr1 = snip.tree(ir.tr, Nodes = c(12, 7))"
   ' Rinterface.rrun "plot (ir.trr1)"
   ' Rinterface.rrun "text(ir.trr1, all=TRUE)"
    'rinterface.rrun "par(pty = " & Chr(34) & "s" & Chr(34) & ")"
   ' rinterface.rrun "plot(iris[,3],iris[,4],type=" & Chr(34) & "n" & Chr(34) & ", xlab= " & Chr(34) & "petal length" & Chr(34) & ", ylab = " & Chr(34) & "petal width" & Chr(34) & ")"
    'rinterface.rrun "text(iris[,3], iris[,4],c(" & Chr(34) & "s" & Chr(34) & "," & Chr(34) & "c" & Chr(34) & "," & Chr(34) & "v" & Chr(34) & ")[iris[,5]])"
   ' rinterface.rrun "partition.tree(ir.trr1, add=TRUE,cex=1.5)"
  '  Rinterface.rrun "plot(partition.tree(ir.trr1, add=TRUE,cex=1.5))"
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
