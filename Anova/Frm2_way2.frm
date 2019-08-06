VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm2_way2 
   OleObjectBlob   =   "Frm2_way2.frx":0000
   Caption         =   "�̿���ġ��"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   StartUpPosition =   1  '������ ���
   TypeInfoVer     =   308
End
Attribute VB_Name = "Frm2_way2"
Attribute VB_Base = "0{87666B75-BD2C-4522-988B-CD1EDC0C13EC}{3BE38790-DD30-43B5-8114-BE01DED63B32}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False



Private Sub BtnCan_Click()
   Unload Me
End Sub

Private Sub BtnOK_Click()
    Dim resultsheet As Worksheet
    Dim c As Integer
    Dim r As Integer
    Dim N As Integer
    Dim tst As Double
    Dim tempmean As Double
    Dim cmean() As Double
    Dim rmean() As Double
    Dim cstd(), rstd() As Variant
    Dim xsq As Double
    Dim sst As Double
    Dim SSE As Double
    Dim st1 As Double
    Dim st2 As Double
    Dim st12 As Double
    Dim TempSheet As Worksheet
    Dim temp As Range
    Dim ttemp As Range
    Dim cnt As Integer
    Dim tmp As Double
    Dim i, J, rf As Integer
    Dim tt, tp, ts As Range
    Dim cl As Range
    Dim df12, dfe As Integer
    Dim tmpmean() As Double
    Dim tmpstd() As Double
    Dim rn() As String
    Dim cn() As String
    Set TempSheet = ActiveCell.Worksheet
    Set temp = TempSheet.Cells.CurrentRegion
    c = temp.Columns.count - 1
    r = temp.Rows.count - 1
      '''���� üũ
    Set ttemp = temp.Offset(1, 1).Resize(r, c)
    If FindingRangeError(ttemp) = True Then
          MsgBox "�м������� ���ڳ� ������ �ֽ��ϴ�.", vbExclamation, "HIST"
          Exit Sub
    End If
    
    If Me.TextBox2.Text = "" Or Me.TextBox3.Text = "" Then
          MsgBox "�������� �̸� Ȥ�� ���������̸��� �Է��ؾ� �մϴ�.", vbExclamation, "HIST"
          Exit Sub
    End If
    
    
    cnt = Me.TextBox1.Value
    'If r <= cnt Then
    '      MsgBox "�������� 2�� �̻��̾�� �մϴ�." _
    '             , vbExclamation, "HIST"
    '      Exit Sub
    'End If
    If r Mod cnt <> 0 Then
          MsgBox "�ݺ����� �����ؾ� �մϴ�." _
                 , vbExclamation, "HIST"
          Exit Sub
    End If
    
    Set ts = temp.Columns(1)
    
    If cnt <> 1 Then
    
    rf = r / cnt
    ReDim rn(1 To rf)
    ReDim cn(1 To c)
    N = c * r
    Set tt = temp.Columns(1)
    
    
    For i = 1 To rf
        rn(i) = tt.Cells((i - 1) * cnt + 2)
        If rn(i) = "" Then
            MsgBox "�ݺ����� �߸� �����Ͽ����ϴ�.", vbExclamation, "HIST"
        Exit Sub
        End If
    Next i
    For i = 1 To rf - 1
        For J = i + 1 To rf
            If rn(i) = rn(J) Then
                MsgBox "����Է¹������ �ٽ� �����ؾ� �մϴ�.", vbExclamation, "HIST"
            Exit Sub
            End If
        Next J
    Next i
    
    Set tt = temp.Rows(1)
    For i = 2 To c + 1
        cn(i - 1) = tt.Cells(i).Value
    Next i
    
        
    
    xsq = 0
    st1 = 0
    st2 = 0
    st12 = 0
    SSE = 0
    
    Set temp = temp.Offset(1, 1).Resize(r, c)
    tsum = Application.sum(temp)
    tst = myStdev(temp)
    For Each d In temp
        If d.Value = "" Then
           xsq = xsq
        Else: xsq = xsq + d.Value ^ 2
        End If
    Next d
    sst = xsq - tsum ^ 2 / N
    ReDim cmean(1 To c)
    ReDim rmean(1 To rf)
    ReDim cstd(1 To c)
    ReDim rstd(1 To rf)
    ReDim tmpmean(1 To rf, 1 To c)
    ReDim tmpstd(1 To rf, 1 To c)
    csum = 0
    rcsum = 0
    For i = 1 To c
        Set ttemp = temp.Columns(i)
        cmean(i) = Application.Average(ttemp)
        cstd(i) = myStdev(ttemp)
        csum = csum + (Application.sum(ttemp)) ^ 2 / (rf * cnt)
        rsum = 0
        For J = 1 To rf
            Set ttemp = temp.Range(Cells(cnt * J - cnt + 1, 1), Cells(cnt * J, c))
            rmean(J) = Application.Average(ttemp)
            rstd(J) = myStdev(ttemp)
            rsum = rsum + (Application.sum(ttemp)) ^ 2 / (c * cnt)
            Set ttemp = temp.Columns(i).Range(Cells(cnt * J - cnt + 1, 1), Cells(cnt * J, 1))
            tempmean = Application.Average(ttemp)
            tmpmean(J, i) = tempmean
            tmpstd(J, i) = myStdev(ttemp)
            rcsum = rcsum + (Application.sum(ttemp)) ^ 2 / cnt
        Next J
   Next i
   st1 = rsum - tsum ^ 2 / N
   st2 = csum - tsum ^ 2 / N
   st12 = rcsum - tsum ^ 2 / N - st1 - st2
   SSE = sst - st1 - st2 - st12
   df12 = (rf - 1) * (c - 1)
   dfe = rf * c * (cnt - 1)
   If chkinteract.Value = False Then
      SSE = SSE + st12
      dfe = dfe + df12
      st12 = 0
      df12 = 0
   End If
   Set resultsheet = OpenOutSheet("_���м����_", True)
   
   '''
    '''
    '''
    RstSheet = "_���м����_"
    '����ϴ� �ش� ��⿡ �� ���� ����'
'������ �Է�
On Error GoTo Err_delete
Dim val3535 As Long '�ʱ���ġ ������ ����'
Dim s3535 As Worksheet
val3535 = 2
    For Each s3535 In ActiveWorkbook.Sheets
        If s3535.name = RstSheet Then
val3535 = Sheets(RstSheet).Cells(1, 1).Value
End If
Next s3535  '��Ʈ�� �̹������� ��� ��ġ �������ϰ�, ������ 2�� �����Ѵ�.
   ' Worksheets(RstSheet).Unprotect "prophet"
    activePt = Worksheets(RstSheet).Cells(1, 1).Value
    
    '''Worksheets(RstSheet).Cells(1, 1).Value = "$A$" & Worksheets(RstSheet).Cells(1, 1).Value

    
    ' resultsheet.Unprotect "prophet"
   TwoWay2_Result.dResult tmpmean, tmpstd, cnt, rmean, rstd, cmean, cstd, rn, cn, rf, c, tsum / N, tst, resultsheet, 1
   TwoWay2_Result.a1Result st1, st2, st12, SSE, rf - 1, c - 1, df12, dfe, resultsheet
   If ChkGraph.Value = True Then
       Set twork = opentemp.opentemp()
       Set ttt = twork.Range(twork.Cells(1, 1).Value)
       Set tm1 = maketemp(ttt, tmpmean, rf, c, rn, cn)
       Agraph.makeGraph tm1, resultsheet
    End If
    'resultsheet.Protect Password:="prophet", DrawingObjects:=False, contents:=True, Scenarios:=True
    
    
    
    
   '''Worksheets(RstSheet).Cells(1, 1).Value = Right(Worksheets(RstSheet).Cells(1, 1).Value, Len(Worksheets(RstSheet).Cells(1, 1).Value) - 3)
    
    

    Else
    
    Set tm = temp
    N = c * r
    xsq = 0
    ReDim rname(1 To r)
    ReDim cname(1 To c)
    Set tp = temp.Columns(1)
    For i = 2 To r + 1
        rname(i - 1) = tp.Cells(i, 1)
        If rname(i - 1) = "" Then
        MsgBox "�ݺ����� �߸� �����Ͽ����ϴ�.", vbExclamation, "HIST"
        Exit Sub
        End If
    Next i
    
     For i = 1 To r - 1
        For J = i + 1 To r
            If rname(i) = rname(J) Then
                MsgBox "����Է¹������ �ٽ� �����ؾ� �մϴ�.", vbExclamation, "HIST"
            Exit Sub
            End If
        Next J
    Next i
    
    Set tp = temp.Rows(1)
    For i = 2 To c + 1
        cname(i - 1) = tp.Cells(1, i)
    Next i
    
    
   
    
    
    
    Set temp = temp.Offset(1, 1).Resize(r, c)
    tmean = Application.Average(temp)
    tsum = Application.sum(temp)
    For Each d In temp
        If d.Value = "" Then
           xsq = xsq
        Else: xsq = xsq + d.Value ^ 2
        End If
    Next d
    ReDim cmean(1 To c)
    ReDim rmean(1 To r)
    ReDim rstd(1 To r)
    ReDim cstd(1 To c)
    csum = 0
    For i = 1 To c
        cmean(i) = Application.Average(temp.Columns(i))
        cstd(i) = myStdev(temp.Columns(i))
        csum = csum + (Application.sum(temp.Columns(i))) ^ 2 / r
        rsum = 0
        For J = 1 To r
            rmean(J) = Application.Average(temp.Rows(J))
            rstd(J) = myStdev(temp.Rows(J))
            rsum = rsum + (Application.sum(temp.Rows(J))) ^ 2 / c
        Next J
    Next i
    sst = xsq - tsum ^ 2 / N
    st1 = rsum - tsum ^ 2 / N
    st2 = csum - tsum ^ 2 / N
    SSE = sst - st1 - st2
    Set resultsheet = OpenOutSheet("_���м����_", True)
    
    '''
    '''
    '''
    RstSheet = "_���м����_"
    'Worksheets(RstSheet).Unprotect "prophet"
    activePt = Worksheets(RstSheet).Cells(1, 1).Value

    
    'resultsheet.Unprotect "prophet"
    TwoWay1_Result.dResult rmean, cmean, rstd, cstd, r, c, rname, cname, resultsheet
    TwoWay1_Result.aResult st1, st2, SSE, r - 1, c - 1, (r - 1) * (c - 1), resultsheet
    If ChkGraph.Value = True Then
       Set twork = opentemp.opentemp
       Set ttt = twork.Range(twork.Cells(1, 1).Value)
       Set tm1 = maketemp1(ttt, tm)
       Agraph.makeGraph tm1, resultsheet
    End If
    
    
    
    
    
    
    
    
    End If
    
    Worksheets(RstSheet).Activate

    '���� ���� üũ �� �񱳰� ����
    Dim Cmp_Value As Long
    
    If ModuleControl.ChkVersion(ActiveWorkbook.name) = True Then
        Cmp_Value = 1048000
    Else
        Cmp_Value = 65000
    End If
    
    If Worksheets(RstSheet).Cells(1, 1).Value > Cmp_Value Then
        MsgBox "[_���м����_]��Ʈ�� ���� ��� ����Ͽ����ϴ�." & vbCrLf & "�� ��Ʈ�� �̸��� �ٲٰų� ������ �ּ���", vbExclamation, "HIST"
        Exit Sub
    End If

    Worksheets(RstSheet).Cells(activePt + 10, 1).Select
    Worksheets(RstSheet).Cells(activePt + 10, 1).Activate
                            '��� �м��� ���۵Ǵ� �κ��� �����ָ� ��ģ��.
                            
    Unload Me
    


'�ǵڿ� ���̱�
Exit Sub
Err_delete:

For Each s3535 In ActiveWorkbook.Sheets
        If s3535.name = RstSheet Then
Sheets(RstSheet).Range(Cells(val3535, 1), Cells(5000, 1000)).Select
Selection.Delete
Sheets(RstSheet).Cells(1, 1) = val3535
Sheets(RstSheet).Cells(val3535, 1).Select

If val3535 = 2 Then
Application.DisplayAlerts = False
Sheets(RstSheet).Delete
End If

End If


Next s3535

MsgBox ("���α׷��� ������ �ֽ��ϴ�.")
 'End sub �տ��� ���δ�.

''�ؼ�, ������ ���� Err_delete�� �ͼ� ù�����ķ� �����. ���� ù���� 2�� ��Ʈ�� �����.�׸��� �����޽��� ���
'rSTsheet����⵵ ���� �������� ��쿡�� �ƹ� ���۵� ���� �ʰ�, �����޽����� ����.
End Sub
Function maketemp1(aa, bb)
    For i = 1 To bb.Rows.count
        For J = 1 To bb.Columns.count
            aa.Offset(i - 1, J - 1).Value = bb.Cells(i, J).Value
        Next J
    Next i
    Set maketemp1 = aa.Resize(bb.Rows.count, bb.Columns.count)
    Worksheets("_TempData_").Cells(1, 1).Value = aa.Offset(bb.Rows.count + 1, 0).Address
End Function
Function maketemp(aa, bb, r, c, rname, cname)
    For i = 1 To r
        aa.Offset(i, 0).Value = rname(i)
        For J = 1 To c
            aa.Offset(i, J).Value = bb(i, J)
            aa.Offset(0, J).Value = cname(J)
        Next J
    Next i
    Set maketemp = aa.Resize(r + 1, c + 1)
    Worksheets("_TempData_").Cells(1, 1).Value = aa.Offset(r + 2, 0).Address
End Function

Private Sub CheckBox2_Click()

End Sub

Private Sub CheckBox3_Click()

End Sub

Private Sub CheckBox5_Click()
If CheckBox5.Value = True Then
    Me.CheckBox3.Enabled = True
Else
    Me.CheckBox3.Value = False
    Me.CheckBox3.Enabled = False
End If

End Sub

Private Sub CommandButton11_Click()
    
    Dim i As Integer
    i = 0
    Do While i <= Me.ListBox1.ListCount - 1
        If Me.ListBox1.Selected(i) = True Then
           Me.ListBox2.AddItem Me.ListBox1.list(i)
           Me.ListBox1.RemoveItem (i)
           Me.CommandButton11.Visible = False
           Me.CommandButton14.Visible = True
           Exit Sub
        End If
        i = i + 1
    Loop

End Sub

Private Sub CommandButton12_Click()
    Dim i As Integer
    i = 0
    Do While i <= Me.ListBox1.ListCount - 1
        If Me.ListBox1.Selected(i) = True Then
           Me.ListBox3.AddItem Me.ListBox1.list(i)
           Me.ListBox1.RemoveItem (i)
           Me.CommandButton12.Visible = False
           Me.CommandButton13.Visible = True
           Exit Sub
        End If
        i = i + 1
    Loop
End Sub

Private Sub CommandButton13_Click()
    Me.ListBox1.AddItem Me.ListBox3.list(0)
    Me.ListBox3.RemoveItem (0)
    Me.CommandButton13.Visible = False
    Me.CommandButton12.Visible = True
End Sub

Private Sub CommandButton14_Click()
    Me.ListBox1.AddItem Me.ListBox2.list(0)
    Me.ListBox2.RemoveItem (0)
    Me.CommandButton14.Visible = False
    Me.CommandButton11.Visible = True
End Sub

Private Sub CommandButton15_Click()
    Dim resultsheet As Worksheet
    Dim c, r, N, cnt() As Long
    Dim i, J, df12, dfe, M1, M2, one, uptwo1, uptwo2, loco As Integer
    Dim tst, ave(), st(), cmean(), rmean(), rcmean(), rcnt(), ccnt(), tmp As Double
    Dim cstd(), rstd(), rcstd() As Variant
    Dim xsq, sst, SSE, st1, st2, st12 As Double
    Dim TempSheet As Worksheet
    Dim temp, ttemp, tt, cl As Range
    Dim rn(), cn() As String
    Dim tsum, tmean, rsum, csum, rcsum As Double
    Dim fitted(), resi() As Double
    Dim selvar, fit, X, xx, yy As Range
    Dim posi(0 To 1) As Long
    Dim xnames()
    Dim s6(), SSabab(), SSab(), SSaab(), SSbab(), SSa(), SSb()
    Dim x1(), xtrx1(), inv1(), xtry1(), beta1()
    Dim x2(), xtrx2(), inv2(), xtry2(), beta2()
    Dim x3(), xtrx3(), inv3(), xtry3(), beta3()
    Dim x4(), xtrx4(), inv4(), xtry4(), beta4()
    Dim x5(), xtrx5(), inv5(), xtry5(), beta5()
    Dim x6(), xtrx6(), inv6(), xtry6(), beta6()
    Dim Sabab, Sab, Saab, Sbab, Sa, Sb, ywsh, res As Worksheet
    Dim c1 As Integer    '�ݺ������� Ȯ���ϴ� ����
    
    '''''''''
    Dim Rowname, Colname, valueName As String
    Dim rRn, cRn, vrn, y As Range
        
    If Me.ListBox2.ListCount = 0 Or Me.ListBox3.ListCount = 0 Or Me.ListBox4.ListCount = 0 Then
        MsgBox "������ ������ �ҿ����մϴ�.", vbExclamation, "HIST"
        Exit Sub
    End If
    
    Rowname = ModuleControl.SelectedVariable(Me.ListBox2.list(0), rRn, True)
    Colname = ModuleControl.SelectedVariable(Me.ListBox3.list(0), cRn, True)
    valueName = ModuleControl.SelectedVariable(Me.ListBox4.list(0), vrn, True)
    
    If FindingRangeError(vrn) Then
        MsgBox "�м������� ���ڳ� ������ �ֽ��ϴ�.", vbExclamation, "HIST"
        Exit Sub
    End If
    
    If rRn.count <> cRn.count Or cRn.count <> vrn.count Then
            MsgBox "�з������� �м��������� ������ �߸��Ǿ����ϴ�.", vbExclamation, "HIST"
            Exit Sub
    End If
    
    Set temp = Cells.CurrentRegion
    Application.ScreenUpdating = False
    ModuleControl.PivotMakerforTwoWay temp, Rowname, Colname, valueName, _
        cnt, ave, st, cn, rn
        
    r = UBound(cnt, 1) - 1: c = UBound(cnt, 2) - 1
     
    For i = 1 To r
        For J = 1 To c
            N = N + cnt(i, J)                         '��ü ������ ����
        Next J
    Next i
    
    c1 = 0
    check = False
    
    For i = 1 To r
        For J = 1 To c
            If cnt(i, J) <= 1 Then
                If cnt(i, J) = cnt(1, 1) Then
                    c1 = c1 + 1
                End If
            End If
        Next J
    Next i
    
    Set resultsheet = OpenOutSheet("_���м����_", True)
   'resultsheet.Unprotect "prophet"
    activePt = resultsheet.Cells(1, 1).Value
    
    xsq = 0: st1 = 0: st2 = 0: st12 = 0: SSE = 0
    tsum = Application.sum(vrn)
    tmean = Application.Average(vrn)
    tst = st(r + 1, c + 1)
    For Each d In vrn
        xsq = xsq + d.Value ^ 2
    Next d
    sst = xsq - tsum ^ 2 / N
    ReDim cmean(1 To c): ReDim rmean(1 To r)
    ReDim rstd(1 To r):  ReDim cstd(1 To c)
    
    csum = 0: rsum = 0
    For i = 1 To c
        cmean(i) = ave(r + 1, i)
        cstd(i) = st(r + 1, i)
        csum = csum + (cmean(i) * r * cnt(1, 1)) ^ 2 / (r * cnt(1, 1))
    Next i
    For J = 1 To r
        rmean(J) = ave(J, c + 1)
        rstd(J) = st(J, c + 1)
        rsum = rsum + (rmean(J) * c * cnt(1, 1)) ^ 2 / (c * cnt(1, 1))
    Next J
    
    ReDim rcmean(1 To r, 1 To c): ReDim rcstd(1 To r, 1 To c): rcsum = 0
    For i = 1 To c
        For J = 1 To r
            rcmean(J, i) = ave(J, i)
            rcstd(J, i) = st(J, i)
            rcsum = rcsum + (rcmean(J, i) * cnt(1, 1)) ^ 2 / cnt(1, 1)
        Next J
    Next i
    
    ReDim rcnt(1 To r)
        For i = 1 To r
            rcnt(i) = cnt(i, c + 1)
        Next i
    ReDim ccnt(1 To c)
        For i = 1 To c
            ccnt(i) = cnt(r + 1, i)
        Next i
    
    If c1 <> r * c Then
   
        TwoWay2_Result.dResult rcmean, rcstd, cnt, rmean, rstd, cmean, cstd, rn, cn, r, c, tsum / N, tst, resultsheet, 2
   
    Else
   
        TwoWay1_Result.dResult rmean, cmean, rstd, cstd, r, c, rn, cn, resultsheet
   End If
   '�ڷ� ���� �ľ�
    one = 0: uptwo1 = 0: uptwo2 = 0
    For i = 1 To r: For J = 1 To c
        If cnt(i, J) = 1 Then
            one = one + 1
        End If
        If cnt(i, J) <> 1 And cnt(1, 1) = cnt(i, J) Then
            uptwo1 = uptwo1 + cnt(i, J)
        End If
        If cnt(1, 1) <> cnt(i, J) Then
            uptwo2 = uptwo2 + 1
        Else
            uptwo2 = uptwo2 + 0
        End If
    Next J: Next i
    
   'ss(a,b,ab) ���ϱ� ���� covariate ��Ʈ���� �����
    If uptwo1 = N Or uptwo2 >= 1 Then
        ReDim s6(1 To N, 1 To r * c)
        J = 1: For i = 1 To N: s6(i, J) = 1: Next i     'i=����� row, j=����� column
        For J = 2 To r
            q = J - 1: p = 1: k = 0
            While p <= c
                L = 1
                While L < r
                    For i = k + 1 To k + cnt(L, p)
                        If L = q Then
                            s6(i, J) = 1
                        Else
                            s6(i, J) = 0
                        End If
                    Next i
                    k = k + cnt(L, p): L = L + 1
                Wend
                For i = k + 1 To k + cnt(L, p)
                    s6(i, J) = -1
                Next i
                k = k + cnt(L, p)
                p = p + 1
            Wend
        Next J
        For J = r + 1 To r + c - 1
            q = J - r: p = 1: k = 0
            While p < c
                For i = k + 1 To k + cnt(r + 1, p)
                    If p = q Then
                        s6(i, J) = 1
                    Else
                        s6(i, J) = 0
                    End If
                Next i
                k = k + cnt(r + 1, p)
                p = p + 1
            Wend
            For i = k + 1 To k + cnt(r + 1, p)
                s6(i, J) = -1
            Next i
        Next J
        For i = 1 To N
            J = r + c
            For u = 1 To r - 1
                For k = 1 To c - 1
                    s6(i, J) = s6(i, u + 1) * s6(i, r + k)
                    J = J + 1
                Next k
            Next u
        Next i
        
        '������ �����ؼ� y �� ���
        Set ywsh = Worksheets.Add
        ywsh.Range("A1").Select
        For i = 1 To N
            Selection.Offset(i - 1, 0).Value = rRn(i)
            Selection.Offset(i - 1, 1).Value = cRn(i)
            Selection.Offset(i - 1, 2).Value = vrn(i)
        Next i
        Range("A1" & ":" & "C" & N).Select
        ywsh.Sort.SortFields.Add Key:=Range("B1" & ":" & "B" & N), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ywsh.Sort.SortFields.Add Key:=Range("A1" & ":" & "A" & N), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With ywsh.Sort
            .SetRange Range("A1" & ":" & "C" & N)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        Set y = Range("C1" & ":" & "C" & N)
        ywsh.Visible = xlSheetHidden
                
        '��ũ��Ʈ�� S(a,b,ab) �Ѹ���
        Set Sabab = Worksheets.Add
        Sabab.Range("A1").Select
        For i = 1 To N
            For J = 1 To r * c
                Selection.Offset(i - 1, J - 1).Value = s6(i, J)
            Next J
        Next i
        x1 = ActiveCell.CurrentRegion
        xtrx1 = Application.MMult(Application.Transpose(x1), x1)
        inv1 = Application.MInverse(xtrx1)
        xtry1 = Application.MMult(Application.Transpose(x1), y)
        beta1 = Application.MMult(inv1, xtry1)
        SSabab = Application.MMult(Application.Transpose(beta1), xtry1)
        Application.DisplayAlerts = False
        Sabab.Delete
        Application.DisplayAlerts = True
        
        '��ũ��Ʈ�� S(a,b) �Ѹ���
        Set Sab = Worksheets.Add
        Sab.Range("A1").Select
        For i = 1 To N
            For J = 1 To r + c - 1
                Selection.Offset(i - 1, J - 1).Value = s6(i, J)
            Next J
        Next i
        x2 = ActiveCell.CurrentRegion
        xtrx2 = Application.MMult(Application.Transpose(x2), x2)
        inv2 = Application.MInverse(xtrx2)
        xtry2 = Application.MMult(Application.Transpose(x2), y)
        beta2 = Application.MMult(inv2, xtry2)
        SSab = Application.MMult(Application.Transpose(beta2), xtry2)
        Application.DisplayAlerts = False
        Sab.Delete
        Application.DisplayAlerts = True
            
        '��ũ��Ʈ�� S(a,ab) �Ѹ���
        Set Saab = Worksheets.Add
        Saab.Range("A1").Select
        For i = 1 To N
            For J = 1 To r
                Selection.Offset(i - 1, J - 1).Value = s6(i, J)
            Next J
            For J = r + c To r * c
                Selection.Offset(i - 1, J - c).Value = s6(i, J)
            Next J
        Next i
        x3 = ActiveCell.CurrentRegion
        xtrx3 = Application.MMult(Application.Transpose(x3), x3)
        inv3 = Application.MInverse(xtrx3)
        xtry3 = Application.MMult(Application.Transpose(x3), y)
        beta3 = Application.MMult(inv3, xtry3)
        SSaab = Application.MMult(Application.Transpose(beta3), xtry3)
        Application.DisplayAlerts = False
        Saab.Delete
        Application.DisplayAlerts = True

        '��ũ��Ʈ�� S(b,ab) �Ѹ���
        Set Sbab = Worksheets.Add
        Sbab.Range("A1").Select
        For i = 1 To N
            J = 1
            Selection.Offset(i - 1, J - 1).Value = s6(i, J)
            For J = r + 1 To r * c
                Selection.Offset(i - 1, J - r).Value = s6(i, J)
            Next J
        Next i
        x4 = ActiveCell.CurrentRegion
        xtrx4 = Application.MMult(Application.Transpose(x4), x4)
        inv4 = Application.MInverse(xtrx4)
        xtry4 = Application.MMult(Application.Transpose(x4), y)
        beta4 = Application.MMult(inv4, xtry4)
        SSbab = Application.MMult(Application.Transpose(beta4), xtry4)
        Application.DisplayAlerts = False
        Sbab.Delete
        Application.DisplayAlerts = True

        '��ũ��Ʈ�� S(a) �Ѹ���
        Set Sa = Worksheets.Add
        Sa.Range("A1").Select
        For i = 1 To N
            For J = 1 To r
                Selection.Offset(i - 1, J - 1).Value = s6(i, J)
            Next J
        Next i
        x5 = ActiveCell.CurrentRegion
        xtrx5 = Application.MMult(Application.Transpose(x5), x5)
        inv5 = Application.MInverse(xtrx5)
        xtry5 = Application.MMult(Application.Transpose(x5), y)
        beta5 = Application.MMult(inv5, xtry5)
        SSa = Application.MMult(Application.Transpose(beta5), xtry5)
        Application.DisplayAlerts = False
        Sa.Delete
        Application.DisplayAlerts = True

        '��ũ��Ʈ�� S(b) �Ѹ���
        Set Sb = Worksheets.Add
        Sb.Range("A1").Select
        For i = 1 To N
            J = 1
            Selection.Offset(i - 1, J - 1).Value = s6(i, J)
            For J = r + 1 To r + c - 1
                Selection.Offset(i - 1, J - r).Value = s6(i, J)
            Next J
        Next i
        x6 = ActiveCell.CurrentRegion
        xtrx6 = Application.MMult(Application.Transpose(x6), x6)
        inv6 = Application.MInverse(xtrx6)
        xtry6 = Application.MMult(Application.Transpose(x6), y)
        beta6 = Application.MMult(inv6, xtry6)
        SSb = Application.MMult(Application.Transpose(beta6), xtry6)
        Application.DisplayAlerts = False
        Sb.Delete
        Application.DisplayAlerts = True

        st12 = SSabab(1) - SSab(1)
        SSE = sst + N * (tmean ^ 2) - SSabab(1)
        df12 = (r - 1) * (c - 1)
        dfe = N - r * c
        If Me.CheckBox2 = True Then
           SSE = SSE + SSabab(1) - SSab(1)
           dfe = dfe + df12
           st12 = 0
           df12 = 0
        End If
    End If
    
    '�ݺ��� �ִ� ���� ������
    If uptwo1 = N Then
        '�л�м�ǥ
        st1 = rsum - tsum ^ 2 / N
        st2 = csum - tsum ^ 2 / N
        st12 = rcsum - tsum ^ 2 / N - st1 - st2
        SSE = sst - st1 - st2 - st12
        df12 = (r - 1) * (c - 1)
        dfe = N - r * c
        If Me.CheckBox2 = True Then
           SSE = SSE + st12
           dfe = dfe + df12
           st12 = 0
           df12 = 0
        End If
        TwoWay2_Result.aResult sst, N, st1, st2, st12, SSE, r - 1, c - 1, df12, dfe, resultsheet, ListBox2.list(0), ListBox3.list(0), 0, False
    Else
    If c1 = r * c Then
    st1 = rsum - tsum ^ 2 / N
    st2 = csum - tsum ^ 2 / N
    SSE = sst - st1 - st2
        TwoWay1_Result.aResult st1, st2, SSE, r - 1, c - 1, (r - 1) * (c - 1), resultsheet
    End If
    End If
        
    '�ݺ��� �ִ� �ұ��������� �̿���ġ��
    If uptwo2 >= 1 Then
        '�л�м�ǥ
        stmodel = SSa(1) - N * (tmean ^ 2) + SSab(1) - SSa(1) + st12
        TwoWay2_Result.aResult sst, N, stmodel, 1, 1, SSE, r * c - 1, 1, 1, dfe, resultsheet, ListBox2.list(0), ListBox3.list(0), "1", True
        
        'Type I
        If Frm2_outoption.CheckBox1.Value = True Then
            st1 = SSa(1) - N * (tmean ^ 2)
            st2 = SSab(1) - SSa(1)
            TwoWay2_Result.sResult sst, st1, st2, st12, SSE, r - 1, c - 1, df12, dfe, resultsheet, ListBox2.list(0), ListBox3.list(0), "I"
        End If
        
        'Type II
        If Frm2_outoption.CheckBox2.Value = True Then
            st1 = SSab(1) - SSb(1)
            st2 = SSab(1) - SSa(1)
            TwoWay2_Result.sResult sst, st1, st2, st12, SSE, r - 1, c - 1, df12, dfe, resultsheet, ListBox2.list(0), ListBox3.list(0), "II"
        End If

        'Type III
        If Frm2_outoption.CheckBox3.Value = True Then
            st1 = SSabab(1) - SSbab(1)
            st2 = SSabab(1) - SSaab(1)
            If Me.CheckBox2 = True Then
                st1 = SSab(1) - SSb(1)
                st2 = SSab(1) - SSa(1)
            End If
            TwoWay2_Result.sResult sst, st1, st2, st12, SSE, r - 1, c - 1, df12, dfe, resultsheet, ListBox2.list(0), ListBox3.list(0), "III"
        End If
        
    End If
   
   If c1 <> r * c Then
   '���ߺ�
   If Frm_Multicom.CheckBox1.Value = True Or Frm_Multicom.CheckBox1.Value = True Or Frm_Multicom.CheckBox1.Value = True Then
    TwoWay2_Result.cResult Rowname, rmean, rn, rcnt, SSE, r - 1, dfe, r, _
        Frm_Multicom.Controls("TextBox1").Value, resultsheet, Frm_Multicom.Controls("CheckBox1").Value, _
        Frm_Multicom.Controls("CheckBox2").Value, Frm_Multicom.Controls("CheckBox3").Value
    
    TwoWay2_Result.cResult Colname, cmean, cn, ccnt, SSE, c - 1, dfe, c, _
        Frm_Multicom.Controls("TextBox1").Value, resultsheet, Frm_Multicom.Controls("CheckBox1").Value, _
        Frm_Multicom.Controls("CheckBox2").Value, Frm_Multicom.Controls("CheckBox3").Value
    End If
    Else
    If Frm_Multicom.CheckBox1.Value = True Or Frm_Multicom.CheckBox1.Value = True Or Frm_Multicom.CheckBox1.Value = True Then
   
    dfe = (N - 1) - (r - 1) - (c - 1)
     TwoWay2_Result.cResult Rowname, rmean, rn, rcnt, SSE, r - 1, dfe, r, _
        Frm_Multicom.Controls("TextBox1").Value, resultsheet, Frm_Multicom.Controls("CheckBox1").Value, _
        Frm_Multicom.Controls("CheckBox2").Value, Frm_Multicom.Controls("CheckBox3").Value

    TwoWay2_Result.cResult Colname, cmean, cn, ccnt, SSE, c - 1, dfe, c, _
        Frm_Multicom.Controls("TextBox1").Value, resultsheet, Frm_Multicom.Controls("CheckBox1").Value, _
        Frm_Multicom.Controls("CheckBox2").Value, Frm_Multicom.Controls("CheckBox3").Value
    End If
    End If
        
   If Me.CheckBox4.Value = True Then
  
    M2 = ActiveSheet.Rows(1).Cells(1, 1).End(xlToRight).Column
        For i = 1 To M2
            If ActiveSheet.Rows(1).Cells(1, i).Value = Me.ListBox2.list(0) Then
                k = i
            End If
        Next i
        For i = 1 To M2
            If ActiveSheet.Rows(1).Cells(1, i).Value = Me.ListBox3.list(0) Then
                q = i
            End If
        Next i
        For i = 1 To M2
            If ActiveSheet.Rows(1).Cells(1, i).Value = Me.ListBox4.list(0) Then
                p = i
            End If
        Next i
    
        ActiveSheet.Rows(1).Cells(1, k).Offset(1, 0).Select
        Set X = Range(Selection, Selection.End(xlDown))
        
        ActiveSheet.Rows(1).Cells(1, q).Offset(1, 0).Select
        Set xx = Range(Selection, Selection.End(xlDown))

        ActiveSheet.Rows(1).Cells(1, p).Offset(1, 0).Select
        Set yy = Range(Selection, Selection.End(xlDown))
        
        ChartOutControl posi, True
       ' Worksheets("_���м����_").Unprotect "prophet"
        activePt = Worksheets("_���м����_").Cells(1, 1).Value
        t = 0: loco = 0: cc = 1
        If VarType(X(1)) < 2 Or VarType(X(1)) > 5 Then
        Dim m3 As Integer
        Dim nx
        
        ModuleControl.openTempWorkSheet TempSheet, "_TwoWayTemp1_"
        If TempSheet.Rows(1).Cells(1, 1).End(xlToRight).Column = 16384 Then
            m3 = 1
        Else
            m3 = TempSheet.Rows(1).Cells(1, 1).End(xlToRight).Column
        End If
    
        ReDim xnames(0 To UBound(rn) - 2, 0 To 1)
    
        For i = 0 To (UBound(rn) - 2)
            xnames(i, 0) = i + 1
            xnames(i, 1) = rn(i + 1)
        Next i
        ReDim nx(0 To N)
        TempSheet.Cells(1, m3) = X(0)

        For i = 1 To N
            For J = 1 To (UBound(rn) - 1)
                If X(i) = rn(J) Then
                    TempSheet.Cells(i, 1) = xnames(J - 1, 0)
                    nx(i) = TempSheet.Cells(i, 1)
                    J = (UBound(rn) - 1)
                End If
            Next J
        Next i
        scatterModule.ScatterPlot "_���м����_", posi(0), posi(1), _
        200, 200, X, yy, "", Rowname, Me.ListBox3.list(0), 0
        
        ModuleControl.Trans Rowname, xnames, (UBound(rn) - 1), t, loco, cc, resultsheet
        t = t + 1
        cc = cc + 1
        loco = loco + 1
        Else
        t = 0
        cc = cc + 1
        loco = 0
        scatterModule.ScatterPlot "_���м����_", posi(0), posi(1), _
        200, 200, X, yy, "", Rowname, Me.ListBox3.list(0), 0
        End If
        
        If VarType(xx(1)) < 2 Or VarType(xx(1)) > 5 Then
        Dim nx1
        
        ModuleControl.openTempWorkSheet TempSheet, "_TwoWayTemp2_"
        If TempSheet.Rows(1).Cells(1, 1).End(xlToRight).Column = 16384 Then
            m3 = 1
        Else
            m3 = TempSheet.Rows(1).Cells(1, 1).End(xlToRight).Column
        End If
    
        ReDim xnames(0 To UBound(cn) - 2, 0 To 1)
    
        For i = 0 To (UBound(cn) - 2)
            xnames(i, 0) = i + 1
            xnames(i, 1) = cn(i + 1)
        Next i
        ReDim nx1(0 To N)
        TempSheet.Cells(1, m3) = xx(0)

        For i = 1 To N
            For J = 1 To (UBound(cn) - 1)
                If xx(i) = cn(J) Then
                    TempSheet.Cells(i, 1) = xnames(J - 1, 0)
                    nx1(i) = TempSheet.Cells(i, 1)
                    J = UBound(cn) - 1
                End If
            Next J
        Next i
                
        scatterModule.ScatterPlot "_���м����_", Worksheets("_���м����_").Cells(activePt, 2).Offset(0, 4).Left + t * 115, _
        Worksheets("_���м����_").Cells(activePt, 2).Offset(0, 4).Top, 200, 200, xx, yy, "", Colname, Me.ListBox3.list(0), 0
        
        
        ModuleControl.Trans Colname, xnames, (UBound(cn) - 1), t, loco, cc, resultsheet
        
        Else
        scatterModule.ScatterPlot "_���м����_", Worksheets("_���м����_").Cells(activePt, 2).Offset(0, 4).Left + t * 115, _
        Worksheets("_���м����_").Cells(activePt, 2).Offset(0, 4).Top, 200, 200, xx, yy, "", Colname, Me.ListBox3.list(0), 0
        
    End If

                           

           
    
        ChartOutControl 200, False
        'Worksheets("_���м����_").Unprotect "prophet"
   
       Set twork = opentemp.opentemp()
       Set ttt = twork.Range(twork.Cells(1, 1).Value)
       Set tm1 = maketemp(ttt, rcmean, r, c, rn, cn)
       Agraph.makeGraph tm1, resultsheet
    End If
   
   
   If CheckBox5.Value = True Then
  '�ݺ��� �ִ� �����Ϳ��� ���հ� ���ؼ� �迭�� ������ ����
    ReDim fitted(0 To N - 1)
    J = 1: k = 1
    For i = 1 To N
        Do While cRn(i) <> cn(J)
            J = J + 1
        Loop
        Do While rRn(i) <> rn(k)
            k = k + 1
        Loop
        fitted(i - 1) = ave(k, J)
        fitted(i - 1) = Application.Round(fitted(i - 1), 4)
        J = 1: k = 1
    Next i
    
    
    
    '�ݺ��� �ִ� �����Ϳ��� ���� ���ؼ� �迭�� ������ ����
    ReDim resi(0 To N - 1)
    J = 1: k = 1
    For i = 1 To N
        Do While cRn(i) <> cn(J)
            J = J + 1
        Loop
        Do While rRn(i) <> rn(k)
            k = k + 1
        Loop
        resi(i - 1) = vrn(i) - ave(k, J)
        resi(i - 1) = Application.Round(resi(i - 1), 4)
        J = 1: k = 1
    Next i
    
    Dim count As Integer
    count = 0
        M1 = ActiveSheet.Rows(1).Cells(1, 1).End(xlToRight).Column
        Set ttemp1 = ActiveSheet.Cells(1, M1 + 1)
        For i = 1 To M1
            If Left(ActiveSheet.Cells(1, i).Value, 3) = "���հ�" Then
                count = count + 1
            End If
        Next i
        
        If count = 0 Then
            ttemp1.Value = "���հ�"
        Else
            ttemp1.Value = "���հ�" & count
        End If
        
        For i = 1 To N
            ttemp1.Offset(i, 0) = fitted(i - 1)
        Next i
        
        Set ttemp2 = ActiveSheet.Cells(1, M1 + 2)
        If ttemp1.Value = "���հ�" Then
        ttemp2.Value = "����"
        Else
        ttemp2.Value = "����" & count
        End If
        For i = 1 To N
           ttemp2.Offset(i, 0) = resi(i - 1)
        Next i
   End If
   
   If Me.CheckBox3.Value = True Then
    '���հ�, ���� ��Ʈ
        Set res = Worksheets.Add
        res.Range("A1").Select
        For i = 1 To N
            Selection.Offset(i - 1, 0).Value = fitted(i - 1)
            Selection.Offset(i - 1, 1).Value = resi(i - 1)
        Next i
        Set fit = Range(Selection, Selection.End(xlDown))
        Set selvar = Range(Selection.Offset(0, 1), Selection.Offset(0, 1).End(xlDown))
        res.Visible = xlSheetHidden
    '���� ����Ȯ���� �׸���
        ChartOutControl posi, True
        'Worksheets("_���м����_").Unprotect "prophet"
'        activePt = Worksheets(Rstsheet).Cells(1, 1).Value

        QQmodule.MainNormPlot selvar, posi(0), posi(1), Worksheets("_���м����_"), VarName:="����", NTest:=True
        
'        ChartOutControl 192, False
        'Worksheets("_���м����_").Protect Password:="prophet", DrawingObjects:=False, _
                                    contents:=True, Scenarios:=True
                                    
    '���� ������ �׸���
'        ChartOutControl posi, True
       ' Worksheets("_���м����_").Unprotect "prophet"
        activePt = Worksheets("_���м����_").Cells(1, 1).Value

        scatterModule.OrderScatterPlot "_���м����_", Worksheets("_���м����_").Cells(activePt, 2).Offset(0, 4).Left, _
        Worksheets("_���м����_").Cells(activePt, 2).Offset(0, 4).Top, 200, 200, selvar, "����", 0

'        ChartOutControl 200, False
        'Worksheets("_���м����_").Protect Password:="prophet", DrawingObjects:=False, _
                                            contents:=True, Scenarios:=True

        '���� vs ���հ� ������ �׸���
        'ChartOutControl posi, True
        'Worksheets("_���м����_").Unprotect "prophet"
        activePt = Worksheets("_���м����_").Cells(1, 1).Value

        scatterModule.ScatterPlot "_���м����_", Worksheets("_���м����_").Cells(activePt, 2).Offset(0, 8).Left, _
        Worksheets("_���м����_").Cells(activePt, 2).Offset(0, 8).Top, 200, 200, fit, selvar, "", "���հ�", "����", 0

        ChartOutControl 200, False
       ' Worksheets("_���м����_").Protect Password:="prophet", DrawingObjects:=False, _
                                            contents:=True, Scenarios:=True
        
        Worksheets("_���м����_").Activate
        Worksheets("_���м����_").Cells(activePt + 5, 1).Select
        Worksheets("_���м����_").Cells(activePt + 5, 1).Activate
    End If
    
   
   
    'resultsheet.Protect Password:="prophet", DrawingObjects:=False, contents:=True, Scenarios:=True
   Unload Me
   Worksheets("_���м����_").Activate
   
    '���� ���� üũ �� �񱳰� ����
    Dim Cmp_Value As Long
    
    If ModuleControl.ChkVersion(ActiveWorkbook.name) = True Then
        Cmp_Value = 1048000
    Else
        Cmp_Value = 65000
    End If
    
    If resultsheet.Cells(1, 1).Value > Cmp_Value Then
        MsgBox "[_���м����_]��Ʈ�� ���� ��� ����Ͽ����ϴ�." & vbCrLf & "�� ��Ʈ�� �̸��� �ٲٰų� ������ �ּ���", vbExclamation, "HIST"
        Exit Sub
    End If

    Application.ScreenUpdating = True
    
    resultsheet.Cells(activePt + 10, 1).Select
    resultsheet.Cells(activePt + 10, 1).Activate
                            '��� �м��� ���۵Ǵ� �κ��� �����ָ� ��ģ��.
    Unload Frm_Multicom
    Unload Frm2_outoption
    Unload Me

End Sub

Private Sub CommandButton16_Click()
ShellExecute 0, "open", "hh.exe", ThisWorkbook.Path + "\HIST%202013.chm::/�̿���ġ��.htm", "", 1
End Sub

Private Sub CommandButton17_Click()
    Unload Me
End Sub

Private Sub CommandButton18_Click()
    
    Dim i As Integer
    i = 0
    Do While i <= Me.ListBox1.ListCount - 1
        If Me.ListBox1.Selected(i) = True Then
           Me.ListBox4.AddItem Me.ListBox1.list(i)
           Me.ListBox1.RemoveItem (i)
           Me.CommandButton18.Visible = False
           Me.CommandButton19.Visible = True
           Exit Sub
        End If
        i = i + 1
    Loop

End Sub

Private Sub CommandButton19_Click()
    Me.ListBox1.AddItem Me.ListBox4.list(0)
    Me.ListBox4.RemoveItem (0)
    Me.CommandButton19.Visible = False
    Me.CommandButton18.Visible = True
End Sub

Private Sub CommandButton20_Click()
Frm2_outoption.Show
End Sub

Private Sub CommandButton9_Click()
Frm_Multicom.Show
End Sub



Private Sub ListBox1_Click()

End Sub

Private Sub ListBox2_Click()

End Sub

Private Sub ListBox3_Click()

End Sub

Private Sub ListBox4_Click()

End Sub

Private Sub MultiPage1_Change()
    If Me.MultiPage1.Value = 0 Then
        Me.Height = 180: Me.Width = 214
        Me.MultiPage1.Height = 150: Me.MultiPage1.Width = 198

    Else
        Me.Height = 258: Me.Width = 263.25
        Me.MultiPage1.Height = 228: Me.MultiPage1.Width = 246
        ModuleControl.SetUpforPage2 Me, 2
        Me.CheckBox4.Value = False
        Me.CheckBox3.Value = False
    End If
End Sub

Private Sub MultiPage2_Change()

End Sub

Private Sub SpinButton1_SpinDown()
    If Me.TextBox1.Value >= 2 Then
        Me.TextBox1.Value = Me.TextBox1.Value - 1
    End If
    
    If Me.TextBox1.Value = 1 Then
    Me.chkinteract.Value = False
    Me.chkinteract.Enabled = False
    Me.TextBox2.Enabled = False
    Me.TextBox3.Enabled = False
    Else
    Me.chkinteract.Value = True
    Me.chkinteract.Enabled = True
    Me.TextBox2.Enabled = True
    Me.TextBox3.Enabled = True
    End If
    
End Sub

Private Sub SpinButton1_SpinUp()
    If Me.TextBox1.Value <= 99 Then
        Me.TextBox1.Value = Me.TextBox1.Value + 1
    End If
    
    If Me.TextBox1.Value = 1 Then
    Me.chkinteract.Value = False
    Me.chkinteract.Enabled = False
    Me.TextBox2.Enabled = False
    Me.TextBox3.Enabled = False
    Else
    Me.chkinteract.Value = True
    Me.chkinteract.Enabled = True
    Me.TextBox2.Enabled = True
    Me.TextBox3.Enabled = True
    End If
End Sub




Private Sub UserForm_Initialize()
    Me.Height = 180: Me.Width = 214
    Me.MultiPage1.Height = 150: Me.MultiPage1.Width = 198
End Sub

Private Sub UserForm_Terminate()
     Unload Me
End Sub
