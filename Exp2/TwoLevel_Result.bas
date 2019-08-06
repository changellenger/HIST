Attribute VB_Name = "TwoLevel_Result"
Sub CFDresult(outputsheet, name, factor, run, block, rep, cp, name1)
    Dim ttemp As Range
    Dim addr, qq As Range
    Dim Comment As String
    Dim yp As Double
    
    Set addr = outputsheet.Range("a1")
    addr.value = addr.value + 3
    Set ttemp = outputsheet.Range("a" & addr.value)

    yp = ttemp.Top
    Set title = outputsheet.Shapes.AddShape(msoShapeRectangle, 3.75, yp + 2.25, 250, 24#)
    title.TextFrame.Characters.text = name & "요인설계"
    With title
        .Fill.ForeColor.SchemeColor = 9
        .Line.DashStyle = msoLineSolid
        .Line.Style = msoLineSingle
        .Line.Weight = 1
        .Line.Visible = msoTrue
        .Shadow.Type = msoShadow1
    End With
    With title.TextFrame.Characters.Font
         .Size = 14
         .ColorIndex = 41
    End With
    title.TextFrame.HorizontalAlignment = xlCenter

    Set ttemp = ttemp.Offset(4, 1)
    Set qq = ttemp.Offset(1, 0)
    With ttemp.Resize(, 5).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With qq.Resize(, 5).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    
    ttemp.value = "요인수"
    ttemp.Offset(0, 1) = "실행 횟수"
    ttemp.Offset(0, 2) = "블록수"
    ttemp.Offset(0, 3) = "반복수"
    ttemp.Offset(0, 4) = "중심점"
    
    Set ttemp = ttemp.Offset(1, 0)
    ttemp.value = factor
    ttemp.Offset(0, 1).value = run
    ttemp.Offset(0, 2).value = block
    ttemp.Offset(0, 3).value = rep
    ttemp.Offset(0, 4).value = cp
    
    Set ttemp = ttemp.Offset(1, 0)
    Comment = "위와같이 설계된 설계점들은 하단의 " & name1 & " Sheet에 저장되어 있습니다."
        With ttemp
            .value = Comment
            .Font.Size = 9
            .HorizontalAlignment = xlLeft
        End With
        
    
    
    Set ttemp = ttemp.Offset(4, -1)
    '''addr.Value = ttemp.Address
    addr.value = Right(ttemp.Address, Len(ttemp.Address) - 3)

End Sub




Sub pResult(xlist, ylist, list, t, xm, ym, tt, ave, st, rname, nn, n, ct, Sign, CheckBox1, outputsheet)
    Dim ttemp As Range
    Dim addr As Range
    Dim qq, qqq As Range
    Dim tsum, dfsse, a1, rows As Integer
    Dim yp As Double
    Dim temp, temp1, temp2, M1, rssr, rnc, pvalue As Double
    Dim sst, ssr1, sse As Double
    Dim xmatrix(), txmatrix(), xx(), invxx(), xy(), beta(), fitted(), resi(), checkm(), ckm(), xmatrix1(), pssr(), tnc(), rlist()
    Dim nc1(), nc2(), nc3()
    Dim chmm As Boolean
    Dim Comment1 As String
    'If IsEmpty(outputsheet.Range("a1")) = True Then
    '   Set ttemp = outputsheet.Cells(2, 1)
    '   Set addr = outputsheet.Range("a1")
    'Else: Set addr = outputsheet.Range("a1")
    '      Set ttemp = outputsheet.Range(addr.Value)
    'End If
    
    Set addr = outputsheet.Range("a1")
    Set ttemp = outputsheet.Range("a" & addr.value)
    If tt > 1 Then
    ReDim nc1(0 To Application.WorksheetFunction.Combin(tt, 1) - 1)
    Else
    ReDim nc1(0 To 0)
    End If
    If tt > 2 Then
    ReDim nc2(0 To Application.WorksheetFunction.Combin(tt, 2) - 1)
    Else
    ReDim nc2(0 To 0)
    End If
    
    If tt > 3 Then
    ReDim nc3(0 To Application.WorksheetFunction.Combin(tt, 3) - 1)
    Else
    ReDim nc3(0 To 0)
    End If
    
    h = 0
    For i = 0 To tt - 1
        For j = 0 To t - 1
            If xlist(i) = list(j) Then
               nc1(h) = nn(i + 1) - 2
            j = t - 1
            Else
            nc1(h) = 0
            End If
         Next j
         h = h + 1
    Next i
    
    h = 0
    For i = 0 To tt - 1
        For j = i + 1 To tt - 1
            For k = 0 To t - 1
            If xlist(i) & "*" & xlist(j) = list(k) Then
            nc2(h) = (nn(i + 1) - 2) * (nn(j + 1) - 2)
            k = t - 1
            Else
            nc2(h) = 0
            End If
            Next k
            h = h + 1
        Next j
    Next i
    
    h = 0
    For i = 0 To tt - 1
        For j = i + 1 To tt - 1
            For k = j + 1 To tt - 1
                For L = 0 To t - 1
                    If list(L) = xlist(i) & "*" & xlist(j) & "*" & xlist(k) Then
                        nc3(h) = (nn(i + 1) - 2) * (nn(j + 1) - 2) * (nn(j + 1) - 2)
                        L = t - 1
                    Else
                    nc3(h) = 0
                    End If
                    
                Next L
                h = h + 1
            Next k
        Next j
    Next i
    
    tsum = Application.WorksheetFunction.sum(nc1) _
         + Application.WorksheetFunction.sum(nc2) _
         + Application.WorksheetFunction.sum(nc3)
    
    ReDim xmatrix(0 To n - 1, 0 To tsum)
    ''주효과에 대한 matrix를 만드는 과정
    For i = 0 To Application.WorksheetFunction.Combin(tt, 1)
        If i = 0 Then
            For j = 0 To n - 1
                xmatrix(j, 0) = 1
            Next j
        h = 1
        e = 0
        ElseIf Left(list(i - 1), 2) = "요인" Then
            For j = 0 To n - 1
                xmatrix(j, h) = xm(j, i - 1)
            Next j
        h = h + 1
        
        ElseIf Left(list(i - 1), 2) = "블록" Then
        For k = 0 To nn(i) - 3
            For j = 0 To n - 1
                If rname(i - 1, k) = xm(j, i - 1) Then
                    xmatrix(j, h) = 1
                ElseIf rname(i - 1, nn(i) - 2) = xm(j, i - 1) Then
                    xmatrix(j, h) = -1
                Else
                    xmatrix(j, h) = 0
                End If
            Next j
          h = h + 1
          Next k
     End If
    Next i
    
    ''2차교호효과에 대한 matrix를 만드는 과정
    If Application.WorksheetFunction.sum(nc2) <> 0 Then
    e = 1
    d = e + nn(1) - 2
    j = 1
    w = 1
    For i = 0 To Application.WorksheetFunction.Combin(tt, 2) - 1
    
        If nc2(i) <> 0 Then
            For a = e To e + nn(w) - 3
                For b = d To d + nn(w + j) - 3
                    For k = 0 To n - 1
                        xmatrix(k, h) = xmatrix(k, a) * xmatrix(k, b)
                    Next k
                h = h + 1
                Next b
             Next a
             If b > Application.WorksheetFunction.sum(nn) - 2 * (tt) Then
                    e = e + nn(w) - 2
                    d = e + nn(w + 1) - 2
                    j = 1
                    w = w + 1
             
                    
             Else
             d = d + nn(w + 1) - 2
             j = j + 1
             If i = Application.WorksheetFunction.Combin(tt, 2) - 2 Then
             j = 0
             End If
             
             End If
         ElseIf d > Application.WorksheetFunction.sum(nn) - 2 * (tt) - nn(tt) + 2 Then
                    e = e + nn(w) - 2
                    d = e + nn(w + 1) - 2
                    j = 1
                    w = w + 1
         Else
         d = d + nn(w + 1) - 2
         j = j + 1
         End If
         

    Next i
        
    End If
    
    ''3차교호효과에 대한 matrix를 만드는 과정
    
    If Application.WorksheetFunction.sum(nc3) <> 0 Then
     e = 1
    d = e + nn(1) - 2
    g = d + nn(2) - 2
    j = 1
    w = 1
    z = 1
    For i = 0 To Application.WorksheetFunction.Combin(tt, 3) - 1
    
        If nc3(i) <> 0 Then
            For a = e To e + nn(w) - 3
                For b = d To d + nn(w + j) - 3
                    For f = g To g + nn(w + j + z) - 3
                        For k = 0 To n - 1
                            xmatrix(k, h) = xmatrix(k, a) * xmatrix(k, b) * xmatrix(k, f)
                        Next k
                    h = h + 1
                    Next f
                Next b
             Next a
             If f > Application.WorksheetFunction.sum(nn) - 2 * (tt) And _
                     b <> g Then
                    d = d + nn(w + j) - 2
                    g = d + nn(w + j + z) - 2
                    j = j + 1
                    z = 1
             ElseIf f > Application.WorksheetFunction.sum(nn) - 2 * (tt) And _
                     b = g Then
                    If e + nn(w) + nn(w + j) - 4 = g Then
                    i = Application.WorksheetFunction.Combin(tt, 3) - 1
                    Else
                    e = e + nn(w) - 2
                    d = e + nn(w + j) - 2
                    g = d + nn(w + j + z) - 2
                    w = w + 1
                    j = 1
                    z = 1
                    End If
             Else
             g = g + nn(w + j + z) - 2
             z = z + 1
             End If
          ElseIf g > Application.WorksheetFunction.sum(nn) - 2 * (tt) - nn(tt) + 2 And _
                     d + nn(w + j) - 2 <> g Then
                    d = d + nn(w + j) - 2
                    g = d + nn(w + j + z) - 2
                    j = j + 1
                    z = 1
          ElseIf g > Application.WorksheetFunction.sum(nn) - 2 * (tt) - nn(tt) + 2 And _
                     d + nn(w + j) - 2 = g Then
                     e = e + nn(w) - 2
                     d = e + nn(w + j) - 2
                     g = d + nn(w + j + z) - 2
                     w = w + 1
                     j = 1
                     z = 1
          Else
          g = g + nn(w + j + z) - 2
          z = z + 1
          End If
     Next i
   End If
   
   
   '교락항을 찾는 함수
   checkm = checkmatrix(xmatrix, n, h, list, nn)
   
   
   '교락항을 지우는 과정
    p = 0
    Dim ok1 As Boolean
    ok1 = False
    ReDim chm1(0 To (list(h - 1) - 1) * (list(h) - 1))
    For i = 1 To list(h - 1) - 1
        For j = 0 To list(h) - 1
        If p = 0 Then
        chm1(p) = checkm(i, j)
        p = p + 1
        Else
        For k = 0 To p - 1
             If chm1(k) = checkm(i, j) Then
                ok1 = True
                k = p
                
             End If
        Next k
        If ok1 = False Then
            chm1(p) = checkm(i, j)
            p = p + 1
        End If
        
        End If
        Next j
    Next i
    
    If p <> 0 Then
    ReDim Preserve chm1(0 To p - 1)
    Dim tttemp2 As Integer
    
    For i = 0 To p - 2
        For j = i + 1 To p - 1
            If chm1(i) > chm1(j) Then
                tttemp2 = chm1(j)
                chm1(j) = chm1(i)
                chm1(i) = tttemp2
            End If
        Next j
    Next i
    
   a = 0
   b = 0
   ReDim xmatrix1(0 To n - 1, 0 To tsum)
    For i = 0 To p - 1
        For j = a To chm1(i) - 1
            For k = 0 To n - 1
            xmatrix1(k, b) = xmatrix(k, j)
            Next k
            b = b + 1
        Next j
        
        a = chm1(i) + 1
    Next i
       
   If Sign = True Then
   For j = 0 To n - 1
   ReDim Preserve xmatrix1(0 To n - 1, 0 To b)
   If xmatrix1(j, tt - 1) = 0 Then
   xmatrix1(j, b) = 1
   Else
   xmatrix1(j, b) = 0
   End If
   Next j
   f = b
   
   Else
   ReDim Preserve xmatrix1(0 To n - 1, 0 To b - 1)
   f = b - 1
   End If
   
   
    ''위에서 만든 matrix를 계산하는 과정
    
    'matrix transpose
    txmatrix = Application.WorksheetFunction.Transpose(xmatrix1)
    'matrix 곱
    xx = Application.WorksheetFunction.MMult(txmatrix, xmatrix1)
    'matrix의 역행렬
    invxx = Application.WorksheetFunction.MInverse(xx)
    'matrix 곱
    xy = Application.WorksheetFunction.MMult(txmatrix, ym)
    'matrix 곱
    beta = Application.WorksheetFunction.MMult(invxx, xy)
    'sum of square구하는 과정
    sst = Application.WorksheetFunction.SumProduct(ym, ym) - n * Application.WorksheetFunction.Average(ym) ^ 2
    ssr1 = Application.WorksheetFunction.MMult(Application.WorksheetFunction.Transpose(beta), xy)(1) - n * Application.WorksheetFunction.Average(ym) ^ 2
    sse = sst - ssr1
    Else
        
   If Sign = True Then
   For j = 0 To n - 1
   ReDim Preserve xmatrix(0 To n - 1, 0 To h)
   If xmatrix(j, tt - 1) = 0 Then
   xmatrix(j, h) = 1
   Else
   xmatrix(j, h) = 0
   End If
   Next j
   f = h
   
   Else
   ReDim Preserve xmatrix(0 To n - 1, 0 To h - 1)
   f = h - 1
   End If
   
    ''위에서 만든 matrix를 계산하는 과정
    
    'matrix transpose
    txmatrix = Application.WorksheetFunction.Transpose(xmatrix)
    'matrix 곱
    xx = Application.WorksheetFunction.MMult(txmatrix, xmatrix)
    'matrix의 역행렬
    invxx = Application.WorksheetFunction.MInverse(xx)
    'matrix 곱
    xy = Application.WorksheetFunction.MMult(txmatrix, ym)
    'matrix 곱
    beta = Application.WorksheetFunction.MMult(invxx, xy)
    'sum of square구하는 과정
    sst = Application.WorksheetFunction.SumProduct(ym, ym) - n * Application.WorksheetFunction.Average(ym) ^ 2
    ssr1 = Application.WorksheetFunction.MMult(Application.WorksheetFunction.Transpose(beta), xy)(1) - n * Application.WorksheetFunction.Average(ym) ^ 2
    sse = sst - ssr1
   
    End If
    
    a = 0
    w = 1
    ReDim tnc(0 To f)
    ReDim ssrr(0 To f)
    For i = 1 To f + 1
    If i >= 2 And i <= 1 + nn(1) - 2 Then
    ssrr(a) = ssrr(a) + beta(i, 1) * xy(i, 1)
    If w > 1 + nn(1) - 2 - 2 Then
    tnc(a) = nn(1) - 2
    a = a + 1
    Else
    w = w + 1
    End If
    
    Else
    ssrr(a) = beta(i, 1) * xy(i, 1)
    tnc(a) = 1
    a = a + 1
    End If
    
    Next i
    a1 = a
    If Sign = True Then
    ReDim Preserve ssrr(0 To a)
    ssrr(a) = ssrr(0) + ssrr(a - 1) - n * Application.WorksheetFunction.Average(ym) ^ 2
    Else
    ReDim Preserve ssrr(0 To a - 1)
    End If
    
    ReDim rlist(0 To f)
    For i = 0 To h - 2
    chmm = False
         For j = 0 To p - 1
         If list(i) = list(chm1(j) - 1) Then
            chmm = True
            j = p
         End If
         Next j
         If chmm = False Then
         rlist(w) = list(i)
         w = w + 1
         End If
    Next i
    
    
    
    If p <> 0 Then
    a = a - 1
    f = f - 1
    If Sign = True Then
    dfsse = n - f - 2
    rows = a + 2
    Else
    a = a + 1
    f = f + 1
    dfsse = n - f - 1
    rows = a + 1
    End If
    
    Else
    If Sign = True Then
    a = a - 1
    f = f - 1
    dfsse = n - f - 2
    rows = a + 2
    Else
    dfsse = n - f - 1
    rows = a + 1
    End If
    
    End If
    
    yp = ttemp.Top
    Set title = outputsheet.Shapes.AddShape(msoShapeRectangle, 3.75, yp + 2.25, 250, 24#)
    title.TextFrame.Characters.text = "요인분석 결과"
    With title
        .Fill.ForeColor.SchemeColor = 9
        .Line.DashStyle = msoLineSolid
        .Line.Style = msoLineSingle
        .Line.Weight = 1
        .Line.Visible = msoTrue
        .Shadow.Type = msoShadow1
    End With
    With title.TextFrame.Characters.Font
         .Size = 14
         .ColorIndex = 41
    End With
    title.TextFrame.HorizontalAlignment = xlCenter

    Set ttemp = ttemp.Offset(3, 1)
    yp = ttemp.Top
    Set title = outputsheet.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 80, 20#)
    title.Shadow.Type = msoShadow17
    With title.Fill
      .ForeColor.SchemeColor = 22
      .Visible = msoTrue
      .Solid
    End With
    
    title.TextFrame.Characters.text = "분산분석표"
    With title.TextFrame.Characters.Font
        .Size = 11
        .ColorIndex = xlAutomatic
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
    
    Set ttemp = ttemp.Offset(2, 0)
        Set qq = ttemp.Offset(rows, 0)
        With ttemp.Resize(, 6).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With ttemp.Resize(, 6).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With qq.Resize(, 6).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        
        With qq.Resize(, 6).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        ttemp.value = "요인"
        ttemp.Offset(0, 1) = "제곱합"
        ttemp.Offset(0, 2) = "자유도"
        ttemp.Offset(0, 3) = "평균제곱"
        ttemp.Offset(0, 4) = "F값"
        ttemp.Offset(0, 5) = "유의확률"
        Set ttemp = ttemp.Offset(1, 0)
        
        For i = 1 To a - 1
        ttemp.value = rlist(i)
        ttemp.Offset(0, 1) = Format(ssrr(i), "0.0000")
        ttemp.Offset(0, 2) = Format(tnc(i), "0.0000")
        ttemp.Offset(0, 3) = Format(ssrr(i) / tnc(i), "0.0000")
        If sse = 0 Or dfsse = 0 Then
        ttemp.Offset(0, 4) = "."
        ttemp.Offset(0, 5) = "."
        Else
        ttemp.Offset(0, 4) = Format((ssrr(i) / tnc(i)) / (sse / dfsse), "0.0000")
        ttemp.Offset(0, 5) = Format(Application.FDist((ssrr(i) / tnc(i)) / (sse / dfsse), tnc(i), dfsse), _
        "0.0000")
        End If
        If ttemp.Offset(0, 5).value < 0.0001 Then
        ttemp.Offset(0, 5) = "< 0.0001"
        End If
        Set ttemp = ttemp.Offset(1, 0)
        Next i
        
        If Sign = True Then
        ttemp.value = "곡면성"
        ttemp.Offset(0, 1) = Format(ssrr(a1), "0.0000")
        ttemp.Offset(0, 2) = Format(tnc(a1 - 1), "0.0000")
        ttemp.Offset(0, 3) = Format(ssrr(a1) / tnc(a1 - 1), "0.0000")
        If sse = 0 Or dfsse = 0 Then
        ttemp.Offset(0, 4) = "."
        ttemp.Offset(0, 5) = "."
        Else
        ttemp.Offset(0, 4) = Format((ssrr(a1) / tnc(a1 - 1)) / (sse / dfsse), "0.0000")
        ttemp.Offset(0, 5) = Format(Application.FDist((ssrr(a1) / tnc(a1 - 1)) / (sse / dfsse), tnc(a1 - 1), dfsse), _
        "0.0000")
        End If
        If ttemp.Offset(0, 5).value < 0.0001 Then
        ttemp.Offset(0, 5) = "< 0.0001"
        End If
        Set ttemp = ttemp.Offset(1, 0)
        End If
        
        
        ttemp.value = "잔차"
        ttemp.Offset(0, 1) = Format(sse, "0.0000")
        ttemp.Offset(0, 2) = Format(dfsse, "0.0000")
        If sse = 0 Or dfsse = 0 Then
        ttemp.Offset(0, 3) = "."
        Else
        ttemp.Offset(0, 3) = Format(sse / dfsse, "0.0000")
        End If
        Set ttemp = ttemp.Offset(1, 0)
        
        ttemp.value = "계"
        ttemp.Offset(0, 1) = Format(sst, "0.0000")
        ttemp.Offset(0, 2) = Format(n - 1, "0.0000")
        Set ttemp = ttemp.Offset(1, 0)
        
    Set ttemp = ttemp.Offset(1, 0)
    
    Set ttemp = ttemp.Offset(1, 0)
    yp = ttemp.Top
    Set title = outputsheet.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 80, 20#)
    title.Shadow.Type = msoShadow17
    With title.Fill
      .ForeColor.SchemeColor = 22
      .Visible = msoTrue
      .Solid
    End With
    title.TextFrame.Characters.text = "모수추정값 및 효과"
    With title.TextFrame.Characters.Font
        .Size = 11
        .ColorIndex = xlAutomatic
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
     Set ttemp = ttemp.Offset(2, 0)
        Set qq = ttemp.Offset(rows - 1, 0)
        With ttemp.Resize(, 6).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With ttemp.Resize(, 6).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With qq.Resize(, 6).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        ttemp.value = "요인"
        ttemp.Offset(0, 1) = "효과"
        ttemp.Offset(0, 2) = "계수"
        ttemp.Offset(0, 3) = "SE 계수"
        ttemp.Offset(0, 4) = "T 값"
        ttemp.Offset(0, 5) = "유의확률"
        Set ttemp = ttemp.Offset(1, 0)
        
        ttemp.value = "상수"
        ttemp.Offset(0, 1) = " "
        ttemp.Offset(0, 2) = Format(beta(1, 1), "0.0000")
        If sse = 0 Or dfsse = 0 Then
        ttemp.Offset(0, 3) = "."
        ttemp.Offset(0, 4) = "."
        ttemp.Offset(0, 5) = "."
        Else
        ttemp.Offset(0, 3) = Format((sse / dfsse * invxx(1, 1)) ^ (1 / 2), "0.0000")
        ttemp.Offset(0, 4) = Format(beta(1, 1) / ((sse / dfsse * invxx(1, 1)) ^ (1 / 2)), "0.00")
        ttemp.Offset(0, 5) = Format(Application.TDist(beta(1, 1) / ((sse / dfsse * invxx(1, 1)) ^ (1 / 2)), dfsse, 2), _
        "0.0000")
        End If
        If ttemp.Offset(0, 5).value < 0.0001 Then
        ttemp.Offset(0, 5) = "< 0.0001"
        End If
        Set ttemp = ttemp.Offset(1, 0)
                
        
        For i = 1 To a - 1
        ttemp.value = rlist(i)
        If i = 1 Then
        ttemp.Offset(0, 1) = " "
        Else
        ttemp.Offset(0, 1) = Format(2 * beta(i + 1, 1), "0.0000")
        End If
        ttemp.Offset(0, 2) = Format(beta(i + 1, 1), "0.0000")
        If sse = 0 Or dfsse = 0 Then
        ttemp.Offset(0, 3) = "."
        ttemp.Offset(0, 4) = "."
        ttemp.Offset(0, 5) = "."
        Else
        ttemp.Offset(0, 3) = Format((sse / dfsse * invxx(i + 1, i + 1)) ^ (1 / 2), "0.0000")
        ttemp.Offset(0, 4) = Format(beta(i + 1, 1) / ((sse / dfsse * invxx(i + 1, i + 1)) ^ (1 / 2)), "0.00")
        ttemp.Offset(0, 5) = Format(Application.TDist(Abs(beta(i + 1, 1)) / ((sse / dfsse * invxx(i + 1, i + 1)) ^ (1 / 2)), dfsse, 2), _
        "0.0000")
        End If
        If ttemp.Offset(0, 5).value < 0.0001 Then
        ttemp.Offset(0, 5) = "< 0.0001"
        End If
        Set ttemp = ttemp.Offset(1, 0)
        Next i
        
        If Sign = True Then
        ttemp.value = "중심점"
        ttemp.Offset(0, 1) = " "
        ttemp.Offset(0, 2) = Format(beta(a1, 1), "0.0000")
        If sse = 0 Or dfsse = 0 Then
        ttemp.Offset(0, 3) = "."
        ttemp.Offset(0, 4) = "."
        ttemp.Offset(0, 5) = "."
        Else
        ttemp.Offset(0, 3) = Format((sse / dfsse * invxx(a1, a1)) ^ (1 / 2), "0.0000")
        ttemp.Offset(0, 4) = Format(beta(a1, 1) / ((sse / dfsse * invxx(a1, a1)) ^ (1 / 2)), "0.00")
        ttemp.Offset(0, 5) = Format(Application.TDist(Abs(beta(a1, 1)) / ((sse / dfsse * invxx(a, a)) ^ (1 / 2)), dfsse, 2), _
        "0.0000")
        End If
        If ttemp.Offset(0, 5).value < 0.0001 Then
        ttemp.Offset(0, 5) = "< 0.0001"
        End If
        Else
     
        End If
        
        Set ttemp = ttemp.Offset(1, 0)
    
    
    Set ttemp = ttemp.Offset(3, 0)
    yp = ttemp.Top
    Set title = outputsheet.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 80, 20#)
    title.Shadow.Type = msoShadow17
    With title.Fill
      .ForeColor.SchemeColor = 22
      .Visible = msoTrue
      .Solid
    End With
    title.TextFrame.Characters.text = "교락항 표시"
    With title.TextFrame.Characters.Font
        .Size = 11
        .ColorIndex = xlAutomatic
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
     Set ttemp = ttemp.Offset(2, 0)
        Set qq = ttemp.Offset(list(h) - 1, 0)
        With ttemp.Resize(, 6).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
                        
        With qq.Resize(, 6).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        
        If p <> 0 Then
        For i = 0 To list(h) - 1
            For j = 0 To f
                If list(checkm(0, i) - 1) = list(j) Then
                    Comment1 = list(checkm(0, i) - nn(1) + 2) & "은 "
                    For k = 1 To list(h - 1) - 1
                        If checkm(k, i) <> 0 Then
                            Comment1 = Comment1 & list(checkm(k, i) - nn(1) + 2) & " "
                        End If
                    Next k
                j = f + 1
                Comment1 = Comment1 & "과(와) 교락되어 있습니다"
                With ttemp
            .value = Comment1
            .Font.Size = 9
            .HorizontalAlignment = xlLeft
        End With
        Set ttemp = ttemp.Offset(1, 0)
            End If
            Next j
            
        Next i
        Else
        Comment1 = "선택한 모형에서는 교락되는 항이 없습니다."
        With ttemp
            .value = Comment1
            .Font.Size = 9
            .HorizontalAlignment = xlLeft
        End With
        Set ttemp = ttemp.Offset(1, 0)
        End If
        Set ttemp = ttemp.Offset(3, -1)
    '''addr.Value = ttemp.Address
    addr.value = Right(ttemp.Address, Len(ttemp.Address) - 3)
    
    
    
 If CheckBox1 = True Then
 If p <> 0 Then
    fitted = Application.WorksheetFunction.MMult(xmatrix1, beta)
 Else
    fitted = Application.WorksheetFunction.MMult(xmatrix, beta)
 End If
    ReDim resi(0 To n - 1)
 For i = 0 To n - 1
 resi(i) = ym(i, 0) - Format(fitted(i + 1, 1), "0.0000")
 Next i
 
    M1 = ActiveSheet.rows(1).Cells(1, 1).End(xlToRight).Column
        Set ttemp1 = ActiveSheet.Cells(1, M1 + 1)
        ttemp1.value = "적합값"
        For i = 1 To n
            ttemp1.Offset(i, 0) = Format(fitted(i, 1), "0.0000")
        Next i
        
        Set ttemp2 = ActiveSheet.Cells(1, M1 + 2)
        ttemp2.value = "잔차"
        For i = 1 To n
           ttemp2.Offset(i, 0) = Format(resi(i - 1), "0.0000")
        Next i
                
        ttemp1.Offset(1, 0).Select
        Set fit = Range(Selection, Selection.End(xlDown))
        
        ttemp2.Offset(1, 0).Select
        Set SelVar = Range(Selection, Selection.End(xlDown))
End If

End Sub
Function checkmatrix(xmatrix, n, h, list, nn)
Dim ckeck As Boolean
Dim r, r1, w, maxr As Integer
Dim checklist(), checkloca()
ReDim checkloca(0 To 7, 0 To h - 1)

r1 = 0
maxr = 0
For i = 0 To h - 2
    check = False
    r = 0
    For j = i + 1 To h - 1
    w = 0
        For k = 0 To n - 1
        
        If xmatrix(k, i) = xmatrix(k, j) Then
        w = w + 1
        End If
        
        Next k
        If i > 0 And i <= nn(1) - 2 Then
        
        If w = 0 And r = 0 Then
        
        checkloca(r, r1) = i
        checkloca(r + 1, r1) = j
        r = r + 2
        check = True
        
        ElseIf w = 0 Then
        
        checkloca(r, r1) = j
        r = r + 1
        check = True
        End If
        
        ElseIf w = n And r = 0 Then
        
        checkloca(r, r1) = i
        checkloca(r + 1, r1) = j
        r = r + 2
        check = True
        
        ElseIf w = n Then
        
        checkloca(r, r1) = j
        r = r + 1
        check = True
        End If
    Next j
    If check = True Then
    r1 = r1 + 1
    End If
    If r > maxr Then
    maxr = r
    End If
Next i
checkmatrix = checkloca
ReDim Preserve list(0 To h)
list(h - 1) = maxr
list(h) = r1
End Function
'기술통계량을 만드는 Sub
Sub dresult(list, ave, st, ct, fact, fn, t, outputsheet)
    Dim ttemp As Range
    Dim addr, qq As Range
    Dim yp As Double
    
    Set addr = outputsheet.Range("a1")
    addr.value = addr.value + 3
    Set ttemp = outputsheet.Range("a" & addr.value)

    yp = ttemp.Top
    Set title = outputsheet.Shapes.AddShape(msoShapeRectangle, 3.75, yp + 2.25, 250, 24#)
    title.TextFrame.Characters.text = name & "요인 분석"
    With title
        .Fill.ForeColor.SchemeColor = 9
        .Line.DashStyle = msoLineSolid
        .Line.Style = msoLineSingle
        .Line.Weight = 1
        .Line.Visible = msoTrue
        .Shadow.Type = msoShadow1
    End With
    With title.TextFrame.Characters.Font
         .Size = 14
         .ColorIndex = 41
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
    
    Set ttemp = ttemp.Offset(3, 1)
    yp = ttemp.Top
    Set title = outputsheet.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 80, 20#)
    title.Shadow.Type = msoShadow17
    With title.Fill
         .ForeColor.SchemeColor = 22
         .Visible = msoTrue
         .Solid
    End With
    title.TextFrame.Characters.text = "기술 통계량"
    With title.TextFrame.Characters.Font
        .Size = 11
        .ColorIndex = xlAutomatic
    End With
    

    For j = 0 To t - 1
    title.TextFrame.HorizontalAlignment = xlCenter
    Set ttemp = ttemp.Offset(2, 0)
    Set qq = ttemp.Offset(fn(j + 1), 0)
    With ttemp.Resize(, 4).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With ttemp.Resize(, 4).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With qq.Resize(, 4).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    
    
    ttemp.value = list(j) & " 의 수준"
    ttemp.Offset(0, 1).value = "개수"
    ttemp.Offset(0, 2).value = "평균"
    ttemp.Offset(0, 3).value = "표준편차"
        
    
    For i = 0 To fn(j + 1) - 1
        Set ttemp = ttemp.Offset(1, 0)
        ttemp.value = fact(j, i)
        ttemp.Offset(0, 1).value = ct(j, i)
        ttemp.Offset(0, 2).value = Format(ave(j, i), "0.0000")
        ttemp.Offset(0, 3).value = Format(st(j, i), "0.0000")
           
    Next i
    
    Set ttemp = ttemp.Offset(1, 0)
    addr.value = ttemp.Address
    addr.value = Right(ttemp.Address, Len(ttemp.Address) - 3)
    Next j
    
    Set ttemp = ttemp.Offset(1, 0)
    addr.value = ttemp.Address
    addr.value = Right(ttemp.Address, Len(ttemp.Address) - 3)
End Sub
