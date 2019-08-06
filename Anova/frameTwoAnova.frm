VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameTwoAnova 
   OleObjectBlob   =   "frameTwoAnova.frx":0000
   Caption         =   "이원배치법"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7470
   StartUpPosition =   1  '소유자 가운데
   TypeInfoVer     =   37
End
Attribute VB_Name = "frameTwoAnova"
Attribute VB_Base = "0{210964CB-4D31-4028-89E1-8F3D1A449F97}{0B27A839-334B-4B7A-B680-BB41F2C8A186}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CheckBox2_Click()

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
    Dim c1 As Integer    '반복유무를 확인하는 변수
    
    '''''''''
    Dim Rowname, Colname, valueName As String
    Dim rRn, cRn, vrn, y As Range
        
    If Me.ListBox2.ListCount = 0 Or Me.ListBox3.ListCount = 0 Or Me.ListBox4.ListCount = 0 Then
        MsgBox "변수의 선택이 불완전합니다.", vbExclamation, "HIST"
        Exit Sub
    End If
    
    Rowname = ModuleControl.SelectedVariable(Me.ListBox2.list(0), rRn, True)
    Colname = ModuleControl.SelectedVariable(Me.ListBox3.list(0), cRn, True)
    valueName = ModuleControl.SelectedVariable(Me.ListBox4.list(0), vrn, True)
    
    If FindingRangeError(vrn) Then
        MsgBox "분석변수에 문자나 공백이 있습니다.", vbExclamation, "HIST"
        Exit Sub
    End If
    
    If rRn.count <> cRn.count Or cRn.count <> vrn.count Then
            MsgBox "분류변수와 분석변수간의 대응이 잘못되었습니다.", vbExclamation, "HIST"
            Exit Sub
    End If
    
    Set temp = Cells.CurrentRegion
    Application.ScreenUpdating = False
    ModuleControl.PivotMakerforTwoWay temp, Rowname, Colname, valueName, _
        cnt, ave, st, cn, rn
        
    r = UBound(cnt, 1) - 1: c = UBound(cnt, 2) - 1
     
    For i = 1 To r
        For J = 1 To c
            N = N + cnt(i, J)                         '전체 데이터 개수
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
    
    Set resultsheet = OpenOutSheet("_통계분석결과_", True)
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
   '자료 유형 파악
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
    
   'ss(a,b,ab) 구하기 위한 covariate 매트릭스 만들기
    If uptwo1 = N Or uptwo2 >= 1 Then
        ReDim s6(1 To N, 1 To r * c)
        J = 1: For i = 1 To N: s6(i, J) = 1: Next i     'i=행렬의 row, j=행렬의 column
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
        
        '데이터 정렬해서 y 값 잡기
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
                
        '워크시트에 S(a,b,ab) 뿌리기
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
        
        '워크시트에 S(a,b) 뿌리기
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
            
        '워크시트에 S(a,ab) 뿌리기
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

        '워크시트에 S(b,ab) 뿌리기
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

        '워크시트에 S(a) 뿌리기
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

        '워크시트에 S(b) 뿌리기
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
    
    '반복이 있는 균형 데이터
    If uptwo1 = N Then
        '분산분석표
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
        
    '반복이 있는 불균형데이터 이원배치법
    If uptwo2 >= 1 Then
        '분산분석표
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
   '다중비교
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

        activePt = Worksheets("_통계분석결과_").Cells(1, 1).Value


    '=============
    

           
    
        ChartOutControl 200, False
        'Worksheets("_통계분석결과_").Unprotect "prophet"
   
       Set twork = opentemp.opentemp()
       Set ttt = twork.Range(twork.Cells(1, 1).Value)
       Set tm1 = maketemp(ttt, rcmean, r, c, rn, cn)
       Agraph.makeGraph tm1, resultsheet
    End If
   
   
   If CheckBox5.Value = True Then
  '반복이 있는 데이터에서 적합값 구해서 배열에 저장해 놓기
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
    
    
    
    '반복이 있는 데이터에서 잔차 구해서 배열에 저장해 놓기
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
            If Left(ActiveSheet.Cells(1, i).Value, 3) = "적합값" Then
                count = count + 1
            End If
        Next i
        
        If count = 0 Then
            ttemp1.Value = "적합값"
        Else
            ttemp1.Value = "적합값" & count
        End If
        
        For i = 1 To N
            ttemp1.Offset(i, 0) = fitted(i - 1)
        Next i
        
        Set ttemp2 = ActiveSheet.Cells(1, M1 + 2)
        If ttemp1.Value = "적합값" Then
        ttemp2.Value = "잔차"
        Else
        ttemp2.Value = "잔차" & count
        End If
        For i = 1 To N
           ttemp2.Offset(i, 0) = resi(i - 1)
        Next i
   End If
   
   If Me.CheckBox3.Value = True Then
    '적합값, 잔차 시트
        Set res = Worksheets.Add
        res.Range("A1").Select
        For i = 1 To N
            Selection.Offset(i - 1, 0).Value = fitted(i - 1)
            Selection.Offset(i - 1, 1).Value = resi(i - 1)
        Next i
        Set fit = Range(Selection, Selection.End(xlDown))
        Set selvar = Range(Selection.Offset(0, 1), Selection.Offset(0, 1).End(xlDown))
        res.Visible = xlSheetHidden
    '잔차 정규확률도 그리기
        ChartOutControl posi, True
        'Worksheets("_통계분석결과_").Unprotect "prophet"
'        activePt = Worksheets(Rstsheet).Cells(1, 1).Value

        QQmodule.MainNormPlot selvar, posi(0), posi(1), Worksheets("_통계분석결과_"), VarName:="잔차", NTest:=True
        
'        ChartOutControl 192, False
        'Worksheets("_통계분석결과_").Protect Password:="prophet", DrawingObjects:=False, _
                                    contents:=True, Scenarios:=True
                                    
    '잔차 산점도 그리기
'        ChartOutControl posi, True
       ' Worksheets("_통계분석결과_").Unprotect "prophet"
        activePt = Worksheets("_통계분석결과_").Cells(1, 1).Value

        scatterModule.OrderScatterPlot "_통계분석결과_", Worksheets("_통계분석결과_").Cells(activePt, 2).Offset(0, 4).Left, _
        Worksheets("_통계분석결과_").Cells(activePt, 2).Offset(0, 4).Top, 200, 200, selvar, "잔차", 0

'        ChartOutControl 200, False
        'Worksheets("_통계분석결과_").Protect Password:="prophet", DrawingObjects:=False, _
                                            contents:=True, Scenarios:=True

        '잔차 vs 적합값 산점도 그리기
        'ChartOutControl posi, True
        'Worksheets("_통계분석결과_").Unprotect "prophet"
        activePt = Worksheets("_통계분석결과_").Cells(1, 1).Value

        scatterModule.ScatterPlot "_통계분석결과_", Worksheets("_통계분석결과_").Cells(activePt, 2).Offset(0, 8).Left, _
        Worksheets("_통계분석결과_").Cells(activePt, 2).Offset(0, 8).Top, 200, 200, fit, selvar, "", "적합값", "잔차", 0

        ChartOutControl 200, False
       ' Worksheets("_통계분석결과_").Protect Password:="prophet", DrawingObjects:=False, _
                                            contents:=True, Scenarios:=True
        
        Worksheets("_통계분석결과_").Activate
        Worksheets("_통계분석결과_").Cells(activePt + 5, 1).Select
        Worksheets("_통계분석결과_").Cells(activePt + 5, 1).Activate
    End If
    
   
   
    'resultsheet.Protect Password:="prophet", DrawingObjects:=False, contents:=True, Scenarios:=True
   Unload Me
   Worksheets("_통계분석결과_").Activate
   
    '파일 버전 체크 후 비교값 정의
    Dim Cmp_Value As Long
    
    If ModuleControl.ChkVersion(ActiveWorkbook.name) = True Then
        Cmp_Value = 1048000
    Else
        Cmp_Value = 65000
    End If
    
    If resultsheet.Cells(1, 1).Value > Cmp_Value Then
        MsgBox "[_통계분석결과_]시트를 거의 모두 사용하였습니다." & vbCrLf & "이 시트의 이름을 바꾸거나 삭제해 주세요", vbExclamation, "HIST"
        Exit Sub
    End If

    Application.ScreenUpdating = True
    
    resultsheet.Cells(activePt + 10, 1).Select
    resultsheet.Cells(activePt + 10, 1).Activate
                            '결과 분석이 시작되는 부분을 보여주며 마친다.
    Unload Frm_Multicom
    Unload Frm2_outoption
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
Private Sub OptionButton1_Click()
   
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

    ReDim myArray(TempSheet.UsedRange.Columns.count - 1)
    a = 0
   For i = 1 To TempSheet.UsedRange.Columns.count
   If arrName(i) <> "" Then                     '빈칸제거
   myArray(a) = arrName(i)
   a = a + 1
   
   Else:
   End If
   Next i
  
   Me.ListBox1.list() = myArray



End Sub

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
