Attribute VB_Name = "ModuleResi"
Option Base 1

''' ȸ�������� �������ð� ���������� ���Ŀ� ¥������.
''' �� �迭�� base �� 1�� ���Ǿ���.
''' vector�� (,1) �� ������ �迭�� §��.

''
''ȸ�� ���� ���
''

    '������跮
    'resi(1)    : ���
    'resi(2)    : vs �������� �׷���
    'resi(3)    : vs ����ġ �׷���
    'resi(4)    : ������׷�
    'resi(5)    : ����Ȯ���׸�
    
    'resi(1)~(5) ����
    'resi(6)~(10) ǥ��ȭ ����
    'resi(11)~(15) ǥ��ȭ ���� ����
    
    'resi(16)   :���������
    'resi(17)   :���߰�����
    'resi(18)   :�κ�ȸ�ͻ�����


Sub Diagnosis(resi, intercept, ci, Alpha, simple, tmpx, index)
    Dim k As Integer, check As Integer                'matrix X �� N*K
    Dim exact As Long, sum As Double
    Dim x(), y(), obs()
    Dim pt As Range
    Dim choice As Integer                               '��� ǥ���� �� ���� ���� -ci, e, r,r_
    Dim Flag As Long
    
    Dim id() As Integer
    Dim e(), r(), r_(), H(), diagH()         '����, ǥ��ȭ����, ǥ��ȭ��������, Hat matrix
    Dim DFFITS(), D(), CovR()
    Dim DFBETA(), vif()
    Dim DW, AutoRho
    Dim eval(), evec(), condNum(), stx()
    Dim varPro()
    Dim UCI(), LCI()
    Dim yhat()
    
    Dim s, s_()
    
    Dim mySheet As Worksheet
    Dim xPos As Double, yPos As Double
    Dim obsRn As Range, eRn As Range, rRn As Range, r_Rn As Range
    Dim yRn As Range, yhatRn As Range, lciRn As Range, uciRn As Range, xRn As Range
    Dim hiddenRange As Range
    Dim hiddenM As Long
    Dim tmp_p
    Dim xlist_tmp()
    
    
 
    ''''''�̻� ���� ����
 

 
    ''''''����Ÿ ����ֱ�
 
    
    y = pureY
    
    ReDim x(UBound(tmpx, 1) + 1, UBound(tmpx, 2) + 1)
    For i = 1 To UBound(tmpx, 1) + 1
    For j = 1 To UBound(tmpx, 2) + 1
        x(i, j) = tmpx(i - 1, j - 1)
    Next j
    Next i
    
    tmp_p = UBound(x, 2)
    
    k = tmp_p
    
    If intercept = True Then
        k = tmp_p + 1
        ReDim tmpx(UBound(tmpx, 1) + 1, UBound(tmpx, 2) + 2)
        
        For i = 1 To UBound(tmpx, 1)
        For j = 1 To UBound(tmpx, 2)
            If j = 1 Then tmpx(i, j) = 1
            If j <> 1 Then tmpx(i, j) = x(i, j - 1)
        Next j
        Next i
        
        x = tmpx
    End If
     
 
    ''''''ReDim
 
     
    ReDim e(n, 1), r(n, 1), r_(n, 1), DFFITS(n, 1), D(n, 1), CovR(n, 1)
    ReDim DFBETA(n, k), eval(k), evec(k, k), condNum(k, 1), varPro(k, k), vif(k, 1), stx(n, k)
    ReDim s_(n, 1), yhat(n, 1), diagH(n, 1), obs(n, 1), UCI(n, 1), LCI(n, 1)
    
    
        If k = tmp_p + 1 Then
            ReDim xlist_tmp(k)
            xlist_tmp(1) = "�����"
            j = 2
            For i = 0 To p - 1
            If index(i) = 1 Then
                xlist_tmp(j) = xlist(i)
                j = j + 1
            End If
            Next i
        End If
        
        If k = tmp_p Then
            ReDim xlist_tmp(k)
            j = 1
            For i = 0 To p - 1
            If index(i) = 1 Then
                xlist_tmp(j) = xlist(i)
                j = j + 1
            End If
            Next i
        End If
        
  
    ''''''Hat, yhat, s ���ϰ�
  
    
    H = Hat(x)
    yhat = mm(H, y)
    id = matI(n)
    e = mm(diff(id, H), y)
    s = Application.WorksheetFunction.sum(mm(t(e), e)) / (n - k)        '����Ÿ������.
    s = s ^ 0.5
    matObs obs, n
    s_ = mats_(k, s, e, H)
    r = matr(e, s, H)
    r_ = matr_(e, s_, H)

       
    Set mySheet = Worksheets(DataSheet)
        
        
  '''
    '''''''ci ��� ���ÿ��ο� ���� yhat, ci  ����Ÿ ��Ʈ�� �Ѹ���
  '''
    
    matCI x, UCI, Alpha, s, k
    LCI = diff(yhat, UCI)
    UCI = add(yhat, UCI)

    choice = 0
    If ci = True Then          '''
    
        Set pt = mySheet.Range("a1")
            
        mySheet.Range(pt.Cells(1, m + 2), pt.Cells(1, m + 2)) = "������"
        mySheet.Range(pt.Cells(2, m + 2), pt.Cells(n + 1, m + 2)) = yhat
        
        choice = choice + 1
        
        mySheet.Range(pt.Cells(1, m + 2 + choice), pt.Cells(1, m + 2 + choice)) = Alpha & "% ����"
        mySheet.Range(pt.Cells(2, m + 2 + choice), pt.Cells(n + 1, m + 2 + choice)) = LCI
        mySheet.Activate
        
        choice = choice + 1
        
        mySheet.Range(pt.Cells(1, m + 2 + choice), pt.Cells(1, m + 2 + choice)) = Alpha & "% ����"
        mySheet.Range(pt.Cells(2, m + 2 + choice), pt.Cells(n + 1, m + 2 + choice)) = UCI
        
    End If
    
  ''
    ''''''ȸ�����ܿ��� Title ������� �����ϰ� ����ϱ�
  ''
    check = 0
    For i = 1 To 17
        If resi(i) = True Then check = check + 1
    Next i
                        'check=0 �ε� p>1 �̸� resi(18)�� ������ exit or title, skip
                        'check=0 �ε� p=1 �̸� simple(1,2,3)�� ������ exit or title
    If check = 0 Then
        If tmp_p > 1 And resi(18) = False Then Exit Sub
        If tmp_p > 1 And resi(18) = True Then
            'ModulePrint.Title1 ("ȸ�����ܰ��")
            GoTo SKIP
        End If
        If simple(1) = False And simple(2) = False And simple(3) = False Then Exit Sub
    End If
    
   ' ModulePrint.Title1 ("ȸ�����ܰ��")

 
    ''''''���� ��跮 ���ϱ�
 
    
    
    DFFITS = matdffit(H, r_)
    D = matD(H, k, r)
    CovR = matCovR(H, k, r_)
    DFBETA = matdfbeta(k, x, H, r_)
    
    

  '''''
    ''''''eigenValue, eigenVector, varianceProportion ���ϱ�
  '''''
    'eval(k), evec(k, k), condNum(k, 1), varpro(k, k), vif(k, 1)
    
If k = 1 Then
    eval(1) = 1
    evec(1, 1) = 1
    condNum(1, 1) = 1
    varPro(1, 1) = 1
    vif(1, 1) = 1
    GoTo shortSkip
End If

    
    standardizedX x, stx, k
    exact = 0.00000001
    eval = Eigenvaluesevec(mm(t(stx), stx), exact)
    evec = EigenvectorsEmat(mm(t(stx), stx), exact)
    calCondNum eval, condNum, k
    
    calVarPro eval, evec, varPro, k
    
    
    
    If intercept = True Then
        calVIF x, vif, k - 1
    Else
        calVIF11 x, vif, k - 1
    End If
    
shortSkip:

    DW = matDW(e)
    AutoRho = matAutoRho(e)
    
  '''
    '''''''���� ��� ���ÿ��ο� ���� ����Ÿ ��Ʈ�� �Ѹ���
  '''
    
        Set pt = mySheet.Range("a1")
        
    If resi(1) = True Then
    
        choice = choice + 1
        
        mySheet.Range(pt.Cells(1, m + 2 + choice), pt.Cells(1, m + 2 + choice)) = "����"
        mySheet.Range(pt.Cells(2, m + 2 + choice), pt.Cells(n + 1, m + 2 + choice)) = e
    
    End If
    
    If resi(6) = True Then
    
        choice = choice + 1
        
        mySheet.Range(pt.Cells(1, m + 2 + choice), pt.Cells(1, m + 2 + choice)) = "ǥ��ȭ����"
        mySheet.Range(pt.Cells(2, m + 2 + choice), pt.Cells(n + 1, m + 2 + choice)) = r
    
    End If
    
    If resi(11) = True Then
    
        choice = choice + 1
        
        mySheet.Range(pt.Cells(1, m + 2 + choice), pt.Cells(1, m + 2 + choice)) = "ǥ��ȭ��������"
        mySheet.Range(pt.Cells(2, m + 2 + choice), pt.Cells(n + 1, m + 2 + choice)) = r_
        
    End If
    
    ''' �ӽý�Ʈ�� ������ ���� �����Ѵ�.
    '''��ȣ     1   2   3   4   5   6       7   8   9
    '''��       a   b   c   d   e   f       g   h   i...
    '''����     obs e   r   r_  y   yhat    lci uci x...

    ModulePrint.MakeTmpSheet "_#TmpHIST1#_"
    Set mySheet = Worksheets("_#TmpHIST1#_")
    Set pt = mySheet.Range("a1")
    
    Set hiddenRange = mySheet.Cells.CurrentRegion
    hiddenM = hiddenRange.Cells(1, 1).End(xlToRight).Column       '������ ��ϵǾ� �ִ� ���� ��
    
    '���� ���� üũ �� �񱳰� ����
    Dim Cmp_Value As Integer
    
    If ModuleControl.ChkVersion(ActiveWorkbook.Name) = True Then
        Cmp_Value = 16300
    Else
        Cmp_Value = 250
    End If
    
    If hiddenM > Cmp_Value And hiddenRange.Cells(1, 1).value = "obs" Then  '������ ��ϵǾ� �ִ� ���� ���� �ʹ� ������
        MsgBox "�ʹ� ���� �۾��� ����Ǿ����ϴ�." & vbCrLf & "������ �ڷᰡ �����˴ϴ�.", vbOKOnly, "HIST"
        mySheet.Delete
        ModulePrint.MakeTmpSheet "_#TmpHIST1#_"
        Set mySheet = Worksheets("_#TmpHIST1#_")
        Set pt = Worksheets("_#TmpHIST1#_").Range("a1")
        Set hiddenRange = mySheet.Cells.CurrentRegion
        hiddenM = hiddenRange.Cells(1, 1).End(xlToRight).Column
    End If
    
    If hiddenM > Cmp_Value And hiddenRange.Cells(1, 1).value <> "obs" Then hiddenM = 0
    
                 
                    '''''''''ù�ٿ� ���񾲱�
                 
    pt.Cells(1, hiddenM + 1) = "obs"
    pt.Cells(1, hiddenM + 2) = "e"
    pt.Cells(1, hiddenM + 3) = "r"
    pt.Cells(1, hiddenM + 4) = "r_"
    pt.Cells(1, hiddenM + 5) = "y"
    pt.Cells(1, hiddenM + 6) = "yhat"
    pt.Cells(1, hiddenM + 7) = "lci"
    pt.Cells(1, hiddenM + 8) = "uci"
    
                 
                    '''''''''�ʿ��� ��跮 ����ϱ�
                 
    mySheet.Range(pt.Cells(2, hiddenM + 1), pt.Cells(n + 1, hiddenM + 1)) = obs
    mySheet.Range(pt.Cells(2, hiddenM + 2), pt.Cells(n + 1, hiddenM + 2)) = e
    mySheet.Range(pt.Cells(2, hiddenM + 3), pt.Cells(n + 1, hiddenM + 3)) = r
    mySheet.Range(pt.Cells(2, hiddenM + 4), pt.Cells(n + 1, hiddenM + 4)) = r_
    mySheet.Range(pt.Cells(2, hiddenM + 5), pt.Cells(n + 1, hiddenM + 5)) = y
    mySheet.Range(pt.Cells(2, hiddenM + 6), pt.Cells(n + 1, hiddenM + 6)) = yhat
    mySheet.Range(pt.Cells(2, hiddenM + 7), pt.Cells(n + 1, hiddenM + 7)) = LCI
    mySheet.Range(pt.Cells(2, hiddenM + 8), pt.Cells(n + 1, hiddenM + 8)) = UCI
    
                 ''''''''''''''''''''''''
                    '''''''''�׷��� �׸��� ���� Range�� ����ֱ�
                 ''''''''''''''''''''''''
    Set obsRn = mySheet.Range(pt.Cells(2, hiddenM + 1), pt.Cells(n + 1, hiddenM + 1))
    Set eRn = mySheet.Range(pt.Cells(2, hiddenM + 2), pt.Cells(n + 1, hiddenM + 2))
    Set rRn = mySheet.Range(pt.Cells(2, hiddenM + 3), pt.Cells(n + 1, hiddenM + 3))
    Set r_Rn = mySheet.Range(pt.Cells(2, hiddenM + 4), pt.Cells(n + 1, hiddenM + 4))
    Set yRn = mySheet.Range(pt.Cells(2, hiddenM + 5), pt.Cells(n + 1, hiddenM + 5))
    Set yhatRn = mySheet.Range(pt.Cells(2, hiddenM + 6), pt.Cells(n + 1, hiddenM + 6))
    Set lciRn = mySheet.Range(pt.Cells(2, hiddenM + 7), pt.Cells(n + 1, hiddenM + 7))
    Set uciRn = mySheet.Range(pt.Cells(2, hiddenM + 8), pt.Cells(n + 1, hiddenM + 8))
    

  '''
    '''''''simple �׸���
  '''
    
    tmp_position = 0
    
    If tmp_p = 1 Then
    
        '''Title
        If simple(1) = True Or simple(2) = True Or simple(3) = True Then

            Set mySheet = Worksheets(RstSheet)
        
            Flag = mySheet.Cells(1, 1).value

            ModulePrint.Title3 "�ܼ� ���� ȸ�Ϳ� ���� �׷���"
            mySheet.Cells(1, 1).value = mySheet.Cells(1, 1).value + 1       ' �� ����.
        
        End If
        
        '''��跮 x ����ϱ�
        
        
        
        Set mySheet = Worksheets("_#TmpHIST1#_")
        Set pt = Worksheets("_#TmpHIST1#_").Range("a1")
        pt.Cells(1, hiddenM + 9) = "x"
                
        If k = tmp_p Then mySheet.Range(pt.Cells(2, hiddenM + 9), pt.Cells(n + 1, hiddenM + 9)) = x
        If k = tmp_p + 1 Then
            ReDim tmpx(n, 1)
            For i = 1 To n
                tmpx(i, 1) = x(i, 2)
            Next i
            mySheet.Range(pt.Cells(2, hiddenM + 9), pt.Cells(n + 1, hiddenM + 9)) = tmpx
        End If
        
        Set xRn = mySheet.Range(pt.Cells(2, hiddenM + 9), pt.Cells(n + 1, hiddenM + 9))
        
        
        '''�׷���
        Set mySheet = Worksheets(RstSheet)
        mySheet.Cells(1, 1).value = mySheet.Cells(1, 1).value + 1       ' �� ����.

        Flag = mySheet.Cells(1, 1).value
        Set pt = mySheet.Cells(Flag, 2)

        xPos = pt.Left: yPos = pt.Top - 15
        If simple(1) = True And IsNumeric(yRn(1, 1)) Then '''y vs x, ������
            Application.Run "Grap.xlam!ModuleScatter.ScatterPlot", _
                RstSheet, xPos, yPos, 200, 200, _
                xRn, yRn, "x", "y", False, "������"
            xPos = xPos + 210
        End If
    
        If Alpha <> 0 And simple(2) = True Then
            FittedBand RstSheet, xPos, yPos, 200, 200, yhatRn, _
                lciRn, uciRn, Alpha, xRn, xlist(0)
            xPos = xPos + 210
        End If
    
        'If simple(3) = True Then
        '    regScatterPlot RstSheet, xPos, yPos, 200, 200, yhatRn, _
        '        xRn, xlist(0), yRn, ylist
        '    xPos = xPos + 10: yPos = yPos + 30
        '    gaesoo = gaesoo + 1
        'End If
    
        If simple(3) = True And IsNumeric(xRn(1, 1)) And IsNumeric(rRn(1, 1)) Then '''ǥ��ȭ���� vs x
            Application.Run "Grap.xlam!ModuleScatter.ScatterPlot", _
                RstSheet, xPos, yPos, 200, 200, _
                xRn, rRn, "x", "ǥ��ȭ����", False, "ǥ��ȭ���� vs ��������"
            xPos = xPos + 210
        End If
        




        'If simple(1) = True Or simple(2) = True Or simple(3) = True Then
        'mySheet.Cells(1, 1).value = mySheet.Cells(1, 1).value + 21 + gaesoo
        'tmp_position = tmp_position + 1
        'End If
    
        mySheet.Cells(1, 1).value = mySheet.Cells(1, 1).value + 17
        
    End If
    

    
  '''
    '''''''"���� �׷���" Title ������� �����ϱ�
  '''
    If resi(2) = True Or resi(3) = True Or resi(4) = True Or resi(5) = True Then

        Set mySheet = Worksheets(RstSheet)
        
        Flag = mySheet.Cells(1, 1).value

        ModulePrint.Title3 "�����׷���"
        mySheet.Cells(1, 1).value = mySheet.Cells(1, 1).value + 2       ' �� ����.
        tmp_position = tmp_position + 1
        
    End If
    
  '''
    '''''''���� �׷��� ���ÿ��ο� ���� RstSheet �� ����ϱ�
  '''
    Set mySheet = Worksheets(RstSheet)

    Flag = mySheet.Cells(1, 1).value
    Set pt = mySheet.Cells(Flag, 2)

    xPos = pt.Left: yPos = pt.Top - 15
    gaesoo = 0
    If resi(2) = True And IsNumeric(eRn(1, 1)) Then '''���� vs ��������
        Application.Run "Grap.xlam!ModuleScatter.OrderScatterPlot", _
            RstSheet, xPos, yPos, 200, 200, _
            eRn, "����", 0
        xPos = xPos + 210
        gaesoo = gaesoo + 1
    End If
    
    If resi(3) = True And IsNumeric(yhatRn(1, 1)) And IsNumeric(e(1, 1)) Then '''���� vs ������
        Application.Run "Grap.xlam!ModuleScatter.ScatterPlot", _
            RstSheet, xPos, yPos, 200, 200, _
            yhatRn, eRn, "������", "����", False, "���� vs ������"
        xPos = xPos + 210
        gaesoo = gaesoo + 1
    End If
    
    If resi(4) = True And IsNumeric(e(1, 1)) Then '''���� ����Ȯ���׸�
        Application.Run "Grap.xlam!QQmodule.MainNormPlot", _
            eRn, xPos, yPos, Sheets(RstSheet), "����", True
        xPos = xPos + 210
        gaesoo = gaesoo + 1
    End If

    If resi(5) = True And IsNumeric(e(1, 1)) Then '''���� ������׷�
        Application.Run "Grap.xlam!Histmodule.MainHistogram", _
            eRn, xPos, yPos, Sheets(RstSheet), 0, "����"
        xPos = xPos + 210
        gaesoo = gaesoo + 1
    End If
    
    'If resi(2) = True Or resi(3) = True Or resi(4) = True Or resi(5) = True Then
    'mySheet.Cells(1, 1).value = mySheet.Cells(1, 1).value + 22 + gaesoo
    'End If
    
    mySheet.Cells(1, 1).value = mySheet.Cells(1, 1).value + 17
    
  '''
    '''''''"ǥ��ȭ ���� �׷���" Title ������� �����ϱ�
  '''
    If resi(7) = True Or resi(8) = True Or resi(9) = True Or resi(10) = True Then

        Set mySheet = Worksheets(RstSheet)
        
        Flag = mySheet.Cells(1, 1).value

        ModulePrint.Title3 "ǥ��ȭ����"
        mySheet.Cells(1, 1).value = mySheet.Cells(1, 1).value + 2       ' �� ����.
        tmp_position = tmp_position + 1
        
    End If
    
  '''
    '''''''ǥ��ȭ ���� �׷��� ���ÿ��ο� ���� RstSheet �� ����ϱ�
  '''
    

    Flag = mySheet.Cells(1, 1).value
    Set pt = mySheet.Cells(Flag, 2)

    xPos = pt.Left: yPos = pt.Top - 15
    gaesoo = 0
    If resi(7) = True And IsNumeric(rRn(1, 1)) Then '''ǥ��ȭ���� vs ��������
        Application.Run "Grap.xlam!ModuleScatter.OrderScatterPlot", _
            RstSheet, xPos, yPos, 200, 200, _
            rRn, "ǥ��ȭ����", 0
        xPos = xPos + 210
        gaesoo = gaesoo + 1
    End If

    If resi(8) = True And IsNumeric(yhatRn(1, 1)) And IsNumeric(r(1, 1)) Then '''ǥ��ȭ���� vs ������
        Application.Run "Grap.xlam!ModuleScatter.ScatterPlot", _
            RstSheet, xPos, yPos, 200, 200, _
            yhatRn, rRn, "������", "ǥ��ȭ����", False, "ǥ��ȭ���� vs ������"
        xPos = xPos + 210
        gaesoo = gaesoo + 1
    End If
    
    If resi(9) = True And IsNumeric(r(1, 1)) Then '''ǥ��ȭ���� ����Ȯ���׸�
        Application.Run "Grap.xlam!QQmodule.MainNormPlot", _
            rRn, xPos, yPos, Sheets(RstSheet), "ǥ��ȭ����", True
        xPos = xPos + 210
        gaesoo = gaesoo + 1
    End If

    If resi(10) = True And IsNumeric(r(1, 1)) Then '''ǥ��ȭ���� ������׷�
        Application.Run "Grap.xlam!Histmodule.MainHistogram", _
            rRn, xPos, yPos, Sheets(RstSheet), 0, "ǥ��ȭ����"
        xPos = xPos + 210
        gaesoo = gaesoo + 1
    End If
    
    'If resi(7) = True Or resi(8) = True Or resi(9) = True Or resi(10) = True Then
    'mySheet.Cells(1, 1).value = mySheet.Cells(1, 1).value + 22 + gaesoo
    'tmp_position = tmp_position + 1
    'End If

    mySheet.Cells(1, 1).value = mySheet.Cells(1, 1).value + 17
    
  '''
    '''''''"ǥ��ȭ ���� ���� �׷���" Title ������� �����ϱ�
  '''
    If resi(12) = True Or resi(13) = True Or resi(14) = True Or resi(15) = True Then

        Set mySheet = Worksheets(RstSheet)
        
        Flag = mySheet.Cells(1, 1).value

        ModulePrint.Title3 "ǥ��ȭ��������"
        mySheet.Cells(1, 1).value = mySheet.Cells(1, 1).value + 2       ' �� ����.
        
    End If
    
  '''
    '''''''ǥ��ȭ�������� �׷��� ���ÿ��ο� ���� RstSheet �� ����ϱ�
  '''
    

    Flag = mySheet.Cells(1, 1).value
    Set pt = mySheet.Cells(Flag, 2)

    xPos = pt.Left: yPos = pt.Top - 15
    gaesoo = 0
    If resi(12) = True And IsNumeric(r_Rn(1, 1)) Then '''ǥ��ȭ�������� vs ��������
        Application.Run "Grap.xlam!ModuleScatter.OrderScatterPlot", _
            RstSheet, xPos, yPos, 200, 200, _
            r_Rn, "ǥ��ȭ��������", 0
        xPos = xPos + 210
        gaesoo = gaesoo + 1
    End If

    If resi(13) = True And IsNumeric(yhatRn(1, 1)) And IsNumeric(r_(1, 1)) Then '''ǥ��ȭ�������� vs ������
        Application.Run "Grap.xlam!ModuleScatter.ScatterPlot", _
            RstSheet, xPos, yPos, 200, 200, _
            yhatRn, r_Rn, "������", "ǥ��ȭ��������", False, "ǥ��ȭ�������� vs ������"
        xPos = xPos + 210
        gaesoo = gaesoo + 1
    End If
    
    If resi(14) = True And IsNumeric(r_(1, 1)) Then '''ǥ��ȭ�������� ����Ȯ���׸�
        Application.Run "Grap.xlam!QQmodule.MainNormPlot", _
            r_Rn, xPos, yPos, Sheets(RstSheet), "ǥ��ȭ��������", True
        xPos = xPos + 210
        gaesoo = gaesoo + 1
    End If

    If resi(15) = True And IsNumeric(r_(1, 1)) Then '''ǥ��ȭ�������� ������׷�
        Application.Run "Grap.xlam!Histmodule.MainHistogram", _
            r_Rn, xPos, yPos, Sheets(RstSheet), 0, "ǥ��ȭ��������"
        xPos = xPos + 210
        gaesoo = gaesoo + 1
    End If
    
    'If resi(12) = True Or resi(13) = True Or resi(14) = True Or resi(15) = True Then
    'mySheet.Cells(1, 1).value = mySheet.Cells(1, 1).value + 22 + gaesoo
    'tmp_position = tmp_position + 1
    'End If
    
    mySheet.Cells(1, 1).value = mySheet.Cells(1, 1).value + 17
    
    If tmp_position > 0 Then
        mySheet.Cells(1, 1).value = mySheet.Cells(1, 1).value + 2
        
        Flag = mySheet.Cells(1, 1).value
        Set pt = mySheet.Cells(Flag, 2)
        
        ModulePrint.TABLE 2, 3, 0
        
        pt.Cells(1, 1) = "DW ��跮"
        pt.Cells(1, 2) = "1�� �ڱ������"
        pt.Cells(2, 1) = DW
        pt.Cells(2, 2) = AutoRho
        pt.Cells(2, 1).Resize(1, 2).NumberFormatLocal = "0.0000_ "
        pt.Cells(1, 2).Resize(1, 1).HorizontalAlignment = xlLeft
        
        pt.Cells(4, 1) = "�������� ���� �ڱ����� ������ DW��跮�� 0 �� ����� ���� ����"
        pt.Cells(5, 1) = "���� �ڱ����� ������ 4 �� ����� ���� ���� �ȴ�."
        pt.Cells(4, 1).Resize(2, 1).HorizontalAlignment = xlLeft
        
        mySheet.Cells(1, 1) = mySheet.Cells(1, 1) + 6
    End If

 '''
    '''''''���߰�����
 '''
    If resi(17) = True Then
        
        Set mySheet = Worksheets(RstSheet)

        Flag = mySheet.Cells(1, 1).value
        Set pt = mySheet.Cells(Flag, 2)

    
        matObs obs, k
        
        ModulePrint.Title3 "���� ������"
        
        mySheet.Cells(1, 1).value = mySheet.Cells(1, 1).value + 1       '�� �� ����.
        
        Flag = mySheet.Cells(1, 1).value
        Set pt = mySheet.Cells(Flag, 2)
        
        ModulePrint.TABLE k + 1, 2, 0
        
        pt.Cells(1, 1) = "������"
        pt.Cells(1, 2) = "�л���â" & vbLf & "����"
        
        
        If k = tmp_p + 1 Then
            pt.Cells(2, 1) = "�����"
            For i = 1 To k - 1
            For j = i - 1 To p - 1
            If index(j) = 1 Then
                pt.Cells(2 + i, 1) = xlist(j)
                Exit For
            End If
            Next j
            Next i
        End If
        
        If k = tmp_p Then
            For i = 1 To k - 1
            For j = i - 1 To p - 1
            If index(j) = 1 Then
                pt.Cells(2 + i, 1) = xlist(j)
                Exit For
            End If
            Next j
            Next i
        End If
        
        mySheet.Range(pt.Cells(2, 2), pt.Cells(k + 1, 2)) = vif
        pt.Cells(2, 2).Resize(k + 1, 2).NumberFormatLocal = "0.0000_ "

        pt.Cells(k + 3, 1) = "�л���â���� > 10 �̸� ���߰������� �ɰ��� ������ �ִٰ� �����Ѵ�."
        pt.Cells(k + 3, 1).Resize(1, 1).HorizontalAlignment = xlLeft
        
        mySheet.Cells(1, 1).value = mySheet.Cells(1, 1).value + k + 4
        
        Flag = mySheet.Cells(1, 1).value
        Set pt = mySheet.Cells(Flag, 2)
        
        ModulePrint.TABLE k + 1, k + 3, 0
        
        pt.Cells(1, 1) = "��ȣ"
        pt.Cells(1, 2) = "������"
        pt.Cells(1, 3) = "��������"
        
        If k = tmp_p + 1 Then
            pt.Cells(1, 4) = "�л����" & vbLf & "�����"
            For i = 1 To k - 1
            pt.Cells(1, 4 + i) = vbLf & xlist_tmp(i + 1)
            Next i
        End If
        
        If k = tmp_p Then
            pt.Cells(1, 4) = "�л����" & vbLf & xlist(0)
            For i = 1 To k - 1
            pt.Cells(1, 4 + i) = vbLf & xlist_tmp(i)
            Next i
        End If
        
        
        mySheet.Range(pt.Cells(2, 2), pt.Cells(k + 1, 2)) = t(eval)
        mySheet.Range(pt.Cells(2, 3), pt.Cells(k + 1, 3)) = condNum
        mySheet.Range(pt.Cells(2, 4), pt.Cells(k + 1, 3 + k)) = varPro
        
     ''''''
        ''''''''''Sorting'''''''''''''
     ''''''
        
        mySheet.Activate
        mySheet.Range(pt.Cells(2, 2), pt.Cells(2, 2)).Select
        Application.CutCopyMode = False
        Selection.Sort Key1:=Range(pt.Cells(2, 2), pt.Cells(2, 2)), Order1:=xlDescending, Header:=xlGuess, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
        
        mySheet.Range(pt.Cells(2, 1), pt.Cells(k + 1, 1)) = obs
        
        pt.Cells(k + 3, 1) = "�������� ��� ũ���� 1�� ���� �ɰ��ϰ� ���� ���"
        pt.Cells(k + 4, 1) = "���������� ���� 10���� Ŭ ���"
        pt.Cells(k + 5, 1) = "�л������ 80-90% �̻����� ��Ÿ���� �������� ������ �� �̻��� ���"
        pt.Cells(k + 6, 1) = "���߰������� ������ �ִٰ� �����Ѵ�."
        pt.Cells(k + 3, 1).Resize(4, 1).HorizontalAlignment = xlLeft
        
        pt.Cells(2, 2).Resize(k + 1, k + 3).NumberFormatLocal = "0.0000_ "
        mySheet.Cells(1, 1) = mySheet.Cells(1, 1) + k + 7
        
        
    End If
    
    
 '''
    '''''''���������
 '''


    If resi(16) = True Then
    
        Set mySheet = Worksheets(RstSheet)
    
        Flag = mySheet.Cells(1, 1).value
        Set pt = mySheet.Cells(Flag, 2)

        diagH = matDiagH(H)
        
        ModulePrint.Title3 "���� ������"
        
        mySheet.Cells(1, 1).value = mySheet.Cells(1, 1).value + 1       '�� �� ����.
        
        ModulePrint.TABLE n + 1, k + 5, 0
        
        Flag = mySheet.Cells(1, 1).value
        Set pt = mySheet.Cells(Flag, 2)

        
        pt.Cells(1, 1) = "��ȣ"
        pt.Cells(1, 2) = "Hat" & vbLf & "Diagonal"
        pt.Cells(1, 3) = "Cook's" & vbLf & "Distance"
        pt.Cells(1, 4) = "���л����"
        pt.Cells(1, 5) = "DFFITS"
        
        If k = tmp_p + 1 Then
            pt.Cells(1, 6) = "DFBETAS" & vbLf & "�����"
            For i = 1 To k - 1
            pt.Cells(1, 6 + i) = vbLf & xlist_tmp(i + 1)
            Next i
        End If
        
        If k = tmp_p Then
            pt.Cells(1, 6) = "DFBETAS" & vbLf & xlist(0)
            For i = 1 To k - 1
            pt.Cells(1, 6 + i) = vbLf & xlist_tmp(i)
            Next i
        End If
        
        mySheet.Range(pt.Cells(2, 1), pt.Cells(n + 1, 1)) = obs
        mySheet.Range(pt.Cells(2, 2), pt.Cells(n + 1, 2)) = diagH
        mySheet.Range(pt.Cells(2, 3), pt.Cells(n + 1, 3)) = D
        mySheet.Range(pt.Cells(2, 4), pt.Cells(n + 1, 4)) = CovR
        mySheet.Range(pt.Cells(2, 5), pt.Cells(n + 1, 5)) = DFFITS
        mySheet.Range(pt.Cells(2, 6), pt.Cells(n + 1, 5 + k)) = DFBETA
        
        pt.Cells(2, 2).Resize(n + 1, 5 + k).NumberFormatLocal = "0.0000_ "
        
        mySheet.Cells(1, 1) = mySheet.Cells(1, 1) + n + 2
        
        '''�������� �Ķ������� ��Ÿ����
        
        For i = 1 To n
            If pt.Cells(1 + i, 2) >= 2 * tmp_p / n Then pt.Cells(1 + i, 2).Font.ColorIndex = 41
        Next i
        
        For i = 1 To n
            If pt.Cells(1 + i, 3) > 1 Then pt.Cells(1 + i, 3).Font.ColorIndex = 41
        Next i
        
        For i = 1 To n
            If Abs(pt.Cells(1 + i, 4) - 1) >= 3 * (tmp_p + 1) / n Then pt.Cells(1 + i, 4).Font.ColorIndex = 41
        Next i
        
        For i = 1 To n
            If pt.Cells(1 + i, 5) > 2 * Sqr((tmp_p + 1) / n) Or pt.Cells(1 + i, 5) > 2 Then pt.Cells(1 + i, 5).Font.ColorIndex = 41
        Next i
        
        If n < 10 Then              '���� data
        For j = 1 To k
        For i = 1 To n
            If pt.Cells(1 + i, 4 + j) > 2 Then pt.Cells(1 + i, 4 + j).Font.ColorIndex = 41
        Next i
        Next j
        
        Else                        'ū data
        For j = 1 To k
        For i = 1 To n
            If pt.Cells(1 + i, 4 + j) > 2 / Sqr(n) Then pt.Cells(1 + i, 4 + j).Font.ColorIndex = 41
        Next i
        Next j
        
        End If
        
    End If
    

        
    
SKIP:
    If resi(18) = False Or k = 1 Then Exit Sub
    partialRegPlot k, y, x, index, intercept
    
End Sub

Sub partialRegPlot(k, y, x, index_tmp, intercept)

    
    Dim tmpx(), index(), e_p_y(), e_p_x(), Xk(), H()
    Dim id() As Integer
    Dim i As Integer
    Dim start As Integer
    Dim xPos As Double, yPose As Double
    Dim tmpRnY As Range, tmpRnX As Range
    
    ReDim index(k) ', H(n, n), id(n, n), e_p_y(n, 1), e_p_x(n, 1)
    
    id = matI(n)
    
    ModulePrint.MakeTmpSheet "_#TmpHIST1#_"
    Set mySheet = Worksheets("_#TmpHIST1#_")
    Set pt = Worksheets("_#TmpHIST1#_").Range("a1")
    
    
    
    If pt.value = "" Then
        hiddenM = 0
    Else
        Set hiddenRange = mySheet.Cells.CurrentRegion
        hiddenM = hiddenRange.Cells(1, 1).End(xlToRight).Column       '������ ��ϵǾ� �ִ� ���� ��
    End If
    
    '���� ���� üũ �� �񱳰� ����
    Dim Cmp_Value As Long
    
    If ModuleControl.ChkVersion(ActiveWorkbook.Name) = True Then
        Cmp_Value = 16300
    Else
        Cmp_Value = 250
    End If
    
    If hiddenM > Cmp_Value Then ''And hiddenRange.Cells(1, 1).value = "obs" Then                                       '������ ��ϵǾ� �ִ� ���� ���� �ʹ� ������
        MsgBox "�ʹ� ���� �۾��� ����Ǿ����ϴ�." & vbCrLf & "������ �ڷᰡ �����˴ϴ�.", vbOKOnly, "HIST"
        mySheet.Delete
        ModulePrint.MakeTmpSheet "_#TmpHIST1#_"
        Set mySheet = Worksheets("_#TmpHIST1#_")
        Set pt = Worksheets("_#TmpHIST1#_").Range("a1")
        hiddenM = dataRange.Cells(1, 1).End(xlToRight).Column
    End If
         
    If hiddenM > Cmp_Value Then hiddenM = 0                   ''And hiddenRange.Cells(1, 1).value <> "obs"
                                                        '������ ��ϵǾ� �ִ� ���� ���� �ʹ� ������
  
    For i = 1 To k
        
        index = makeIndex(k, 1)
        index(i) = 0
        tmpx = selectedX(index, x, k)
            
        index = makeIndex(k, 0)
        index(i) = 1
        Xk = selectedX(index, x, k)
            
        H = Hat(x)
        e_p_y = mm(diff(id, H), y)
        e_p_x = mm(diff(id, H), Xk)
            
                ''' ���� ��� ���� , ó�� k��  y�κ�ȸ�Ͱ�, ���� k�� x�κ�ȸ�Ͱ�
                
        mySheet.Range(pt.Cells(1, hiddenM + i), pt.Cells(1, hiddenM + i)) = "pp"
        mySheet.Range(pt.Cells(2, hiddenM + i), pt.Cells(n + 1, hiddenM + i)) = e_p_y
        mySheet.Range(pt.Cells(1, hiddenM + k + i), pt.Cells(1, hiddenM + k + i)) = "pp"
        mySheet.Range(pt.Cells(2, hiddenM + k + i), pt.Cells(n + 1, hiddenM + k + i)) = e_p_x
        
    Next i
        
  '''
    '''''''"�κ�ȸ�ͻ�����" Title ���
  '''
    ModulePrint.Title3 "�κ�ȸ�ͻ�����"
    
  '''
    '''''''�κ�ȸ�ͻ����� RstSheet �� ����ϱ�
  '''
    
    Worksheets(RstSheet).Cells(1, 1).value = Worksheets(RstSheet).Cells(1, 1).value + 1       ' �� ����.

    Flag = Worksheets(RstSheet).Cells(1, 1).value
    Set pt = Worksheets(RstSheet).Cells(Flag, 2)

    xPos = pt.Left: yPos = pt.Top
    gaesoo = 0

    Set mySheet = Worksheets("_#TmpHIST1#_")
    Set pt = Worksheets("_#TmpHIST1#_").Range("a1")

    tmp_p = 0
    For i = 0 To UBound(index_tmp)
        If index_tmp(i) = 1 Then tmp_p = tmp_p + 1
    Next i
    
        If intercept = True Then
            ReDim xlist_tmp(k)
            xlist_tmp(1) = "�����"
            j = 2
            For i = 1 To p
            If index_tmp(i - 1) = 1 Then
                xlist_tmp(j) = xlist(i - 1)
                j = j + 1
            End If
            Next i
        Else
            ReDim xlist_tmp(k)
            j = 1
            For i = 0 To p - 1
            If index_tmp(i) = 1 Then
                xlist_tmp(j) = xlist(i)
                j = j + 1
            End If
            Next i
        End If
        
        
      For i = 1 To k
            Set tmpRnY = mySheet.Range(pt.Cells(2, hiddenM + i), pt.Cells(n + 1, hiddenM + i))
            Set tmpRnX = mySheet.Range(pt.Cells(2, hiddenM + i + k), pt.Cells(n + 1, hiddenM + i + k))
        
            Application.Run "Grap.xlam!ModuleScatter.ScatterPlot", _
                RstSheet, xPos, yPos, 200, 200, _
                tmpRnX, tmpRnY, xlist_tmp(i), ylist, False, "�κ�ȸ�ͻ�����"
            
            xPos = xPos + 210
            gaesoo = gaesoo + 1
        Next i


    'If intercept = True Then
    '
    '        Set tmpRnY = mySheet.Range(pt.Cells(2, hiddenM + 1), pt.Cells(n + 1, hiddenM + 1))
    '        Set tmpRnX = mySheet.Range(pt.Cells(2, hiddenM + 1 + k), pt.Cells(n + 1, hiddenM + 1 + k))
    '
    '        Application.Run "Grap.xlam!ModuleScatter.ScatterPlot", _
    '            RstSheet, xPos, yPos, 200, 200, _
    '            tmpRnX, tmpRnY, "�����", ylist, False, "�κ�ȸ�ͻ�����"
    '        xPos = xPos + 10: yPos = yPos + 30
    '        gaesoo = gaesoo + 1

    '    For i = 2 To k
    '        Set tmpRnY = mySheet.Range(pt.Cells(2, hiddenM + i), pt.Cells(n + 1, hiddenM + i))
    '        Set tmpRnX = mySheet.Range(pt.Cells(2, hiddenM + i + k), pt.Cells(n + 1, hiddenM + i + k))
    '
    '        Application.Run "Grap.xlam!ModuleScatter.ScatterPlot", _
    '            RstSheet, xPos, yPos, 200, 200, _
    '            tmpRnX, tmpRnY, xlist_tmp(i), ylist, False, "�κ�ȸ�ͻ�����"
    '        xPos = xPos + 10: yPos = yPos + 30
    '        gaesoo = gaesoo + 1
    '    Next i
    '
    'End If
    
    'If k = tmp_p Then
        
    '    For i = 1 To k
    '        Set tmpRnY = mySheet.Range(pt.Cells(2, hiddenM + i), pt.Cells(n + 1, hiddenM + i))
    '        Set tmpRnX = mySheet.Range(pt.Cells(2, hiddenM + i + k), pt.Cells(n + 1, hiddenM + i + k))
    '
    '        Application.Run "Grap.xlam!ModuleScatter.ScatterPlot", _
    '            RstSheet, xPos, yPos, 200, 200, _
    '            tmpRnX, tmpRnY, xlist_tmp(i), ylist, False, "�κ�ȸ�ͻ�����"
    '        xPos = xPos + 10: yPos = yPos + 30
    '        gaesoo = gaesoo + 1
    '    Next i
    '
    'End If
        
    Worksheets(RstSheet).Cells(1, 1).value = Worksheets(RstSheet).Cells(1, 1).value + 21
    
End Sub

''index�� 0�� �ƴ� ������������ data���� �迭�� �����ش�.
''�̶� x �� pureX �̴�
Function selectedX(index, x, k)

    Dim p1 As Integer, j As Integer, jj As Integer
    Dim tmpx()
    
    p1 = 0
    For j = 1 To k
        If index(j) <> 0 Then p1 = p1 + 1
    Next j
    
    ReDim tmpx(n, p1)
    j = 1
    For jj = 1 To k
        If index(jj) <> 0 Then
        For i = 1 To n
            tmpx(i, j) = x(i, jj)
        Next i
        j = j + 1
        End If
    Next jj
    
    selectedX = tmpx
    
End Function

Function makeIndex(m, value)

    Dim i As Integer
    Dim tmpIndex()
    
    ReDim tmpIndex(m)
    
    For i = 1 To m
        tmpIndex(i) = value
    Next i
    
    makeIndex = tmpIndex
    
End Function

Sub matObs(obs, num)

    Dim i As Long
    
    For i = 1 To num
        obs(i, 1) = i
    Next i
    
End Sub

Function matDiagH(H)

    Dim i As Long
    Dim tmp()
    
    ReDim tmp(n, 1)
    
    For i = 1 To n
        tmp(i, 1) = H(i, i)
    Next i
    
    matDiagH = tmp
    
End Function


Sub matCI(x, UCI, Alpha, s, k)
    Dim tmp(), tmpx()
    Dim i As Long, j As Long
    Dim tvalue As Double
    
    
    tmp = Inv(mm(t(x), x))
    tvalue = Application.WorksheetFunction.TInv(1 - Alpha / 100, n - k)
    ReDim tmpx(1, k)
    
    For i = 1 To n
            
        For j = 1 To k
            tmpx(1, j) = x(i, j)
        Next j
    
        UCI(i, 1) = Application.WorksheetFunction.sum((mm(mm(tmpx, tmp), t(tmpx))))  'error����
        UCI(i, 1) = s * tvalue * Sqr(UCI(i, 1))                                      ''' 1+ ����
        
    Next i
    
End Sub

Sub calVIF(x, vif, k)
 
    Dim y() As Double
    Dim X_() As Double
    Dim j As Integer, i As Integer, s As Integer
    
    Dim check As Integer
    
    Dim ybar As Double
    Dim yhat() As Variant
    Dim beta() As Variant
    
    Dim sst As Double, ssr As Double, sse As Double
      
      
      
    ReDim y(n)
    ReDim X_(n, k)
    
    vif(1, 1) = 0
        
    For s = 2 To k + 1
    
                For i = 1 To n
                        
                    y(i) = x(i, s)
                    
                Next i
                    
                For i = 1 To n
                    For j = 1 To k
                    
                        If j >= s Then  '�� ĭ �ǳʶڴ�. s ���� ũ��
                            check = 1
                        Else
                            check = 0
                        End If
                        
                        X_(i, j) = x(i, j + check)
                        
                     Next j
                Next i
                
        beta = mm(mm(Inv(mm(t(X_), X_)), t(X_)), t(y))
        yhat = mm(X_, beta)
        
        sst = 0
        ssr = 0
            
        ybar = Application.Average(y)
            
        For i = 1 To n
            sst = sst + (y(i) - ybar) ^ 2
            ssr = ssr + (yhat(i, 1) - ybar) ^ 2
        Next i
                        
        vif(s, 1) = (1 - (ssr / sst)) ^ -1
        
    Next s
    
End Sub

Sub calVIF11(x, vif, k)
 
    Dim y() As Double
    Dim X_() As Double
    Dim j As Integer, i As Integer, s As Integer
    
    Dim check As Integer
    
    Dim ybar As Double
    Dim yhat() As Variant
    Dim beta() As Variant
    
    Dim sst As Double, ssr As Double, sse As Double
      
      
      
    ReDim y(n)
    ReDim X_(n, k)
    
    If intercept = True Then vif(1, 1) = 0
        
    For s = 1 To k + 1
    
                For i = 1 To n
                        
                    y(i) = x(i, s)
                    
                Next i
                    
                For i = 1 To n
                    For j = 1 To k
                    
                        If j >= s Then  '�� ĭ �ǳʶڴ�. s ���� ũ��
                            check = 1
                        Else
                            check = 0
                        End If
                        
                        X_(i, j) = x(i, j + check)
                        
                     Next j
                Next i
                
        beta = mm(mm(Inv(mm(t(X_), X_)), t(X_)), t(y))
        yhat = mm(X_, beta)
        
        sst = 0
        ssr = 0
            
            
        For i = 1 To n
            sst = sst + y(i) ^ 2
            ssr = ssr + yhat(i, 1) ^ 2
        Next i
                        
        vif(s, 1) = (1 - (ssr / sst)) ^ -1
        
    Next s
    
End Sub

Sub calVarPro(eval, evec, varPro, k)

    Dim Cii()
    ReDim Cii(k)
        
    For i = 1 To k
        sum = 0
        For j = 1 To k
            sum = sum + (evec(i, j)) ^ 2 / eval(j)
        Next j
            
        Cii(i) = sum
            
    Next i
        
    For i = 1 To k
        For j = 1 To k
            varPro(j, i) = (evec(i, j)) ^ 2 / eval(j) / Cii(i)
        Next j
    Next i
    
End Sub

Sub calCondNum(eval, condNum, k)
    
    MaxEigenvalue = Application.max(eval)
    
    For i = 1 To k
        condNum(i, 1) = (MaxEigenvalue / eval(i)) ^ 0.5
    Next i

End Sub

Sub standardizedX(x, stx, k)

    Dim i As Long, j As Long
        
    For j = 1 To k
    
        sum = 0
        
        For i = 1 To n
            sum = sum + (x(i, j)) ^ 2
        Next i
        
        Si = sum ^ 0.5
            
        For i = 1 To n
            stx(i, j) = (x(i, j)) / Si
        Next i
        
    Next j

End Sub

Function matAutoRho(e)

    Dim i As Long
    Dim e1()
    
    ReDim e1(n - 1), e2(n - 1)
    
    mu = Application.WorksheetFunction.Average(e)
    For i = 1 To n - 1
        e1(i) = (e(i, 1) - mu) * (e(i + 1, 1) - mu)
    Next i
    
    gamma0 = Application.WorksheetFunction.Var(e) * (n - 1)
    gamma1 = Application.WorksheetFunction.sum(e1)
    
    matAutoRho = gamma1 / gamma0
    
End Function

Function matDW(e)

    Dim i As Long
    Dim num As Long, denum As Long
    
    num = e(1, 1) ^ 2
    demum = 0
    For i = 2 To n
        num = num + e(i, 1) ^ 2
        denum = denum + (e(i, 1) - e(i - 1, 1)) ^ 2
    Next i
            
    If num = 0 Then
        matDW = "infinity"
    Else
        matDW = denum / num
    End If
    
End Function

Function matdfbeta(k, x, H, r_)

    Dim i As Long, j As Long
    Dim tmp()
    Dim r()
    Dim RR()
    Dim tmp1()
    
    ReDim tmp(n, k)
    ReDim RR(k, 1)
    ReDim tmp1(k, k)
    ReDim RR(k, 1)
    
    
    r = mm(Inv(mm(t(x), x)), t(x))
    tmp1 = mm(r, t(r))
        
    If k = 1 Then
        For i = 1 To n
        For j = 1 To k
            tmp(i, j) = r(i) / Sqr(tmp1(j)) * r_(i, 1) / Sqr(1 - H(i, i))
        Next j
        Next i
        
    matdfbeta = tmp
    Exit Function
    
    End If
    
        
    For i = 1 To n
    For j = 1 To k
        tmp(i, j) = r(j, i) / Sqr(tmp1(j, j)) * r_(i, 1) / Sqr(1 - H(i, i))
    Next j
    Next i
    
    matdfbeta = tmp
    
End Function

Function matCovR(H, k, r_)

    Dim i As Long
    Dim tmp()
    
    ReDim tmp(n, 1)
    
    For i = 1 To n
        tmp(i, 1) = (1 + (r_(i, 1) ^ 2 - 1) / (n - k)) ^ k
        tmp(i, 1) = 1 / tmp(i, 1) / (1 - H(i, i))
    Next i
    
    matCovR = tmp
    
End Function

Function matD(H, k, r)

    Dim i As Long
    Dim tmp()
    
    ReDim tmp(n, 1)
    
    For i = 1 To n
        tmp(i, 1) = H(i, i) / k / (1 - H(i, i)) * r(i, 1) ^ 2
    Next i
    
    matD = tmp
    
End Function

Function matdffit(H, r_)

    Dim i As Long
    Dim tmp()
    
    ReDim tmp(n, 1)
    
    For i = 1 To n
        tmp(i, 1) = Sqr(H(i, i) / (1 - H(i, i))) * r_(i, 1)
    Next i
    
    matdffit = tmp
    
End Function

Function matr_(e, s_, H)

    Dim i As Long
    Dim tmp()
    
    ReDim tmp(n, 1)
    
    For i = 1 To n
        tmp(i, 1) = e(i, 1) / s_(i, 1) / Sqr(1 - H(i, i))
    Next i
    
    matr_ = tmp
    
End Function

Function mats_(k, s, e, H)

    Dim i As Long
    Dim tmp()
    
    ReDim tmp(n, 1)
    
    For i = 1 To n
        tmp(i, 1) = ((n - k) * s ^ 2 - e(i, 1) ^ 2 / (1 - H(i, i))) / (n - k - 1)
        tmp(i, 1) = tmp(i, 1) ^ 0.5
    Next i
    
    mats_ = tmp
    
End Function

Function matr(e, s, H)

    Dim i As Long
    Dim tmp()
    
    ReDim tmp(n, 1)
    
    For i = 1 To n
        tmp(i, 1) = e(i, 1) / s / Sqr(1 - H(i, i))
    Next i
    
    matr = tmp
    
End Function

Function diff(a, b)     '�� �迭�� ���� ��ȯ

    Dim i As Long, j As Long, r As Long, c As Long
    Dim mat()
    
    r = UBound(a, 1)
    c = UBound(a, 2)
    ReDim mat(r, c)
    
    For i = 1 To r
    For j = 1 To c
        mat(i, j) = a(i, j) - b(i, j)
    Next j
    Next i
    
    diff = mat
    
End Function

Function add(a, b)     '�� �迭�� ���� ��ȯ

    Dim i As Long, j As Long, r As Long, c As Long
    Dim mat()
    
    r = UBound(a, 1)
    c = UBound(a, 2)
    ReDim mat(r, c)
    
    For i = 1 To r
    For j = 1 To c
        mat(i, j) = a(i, j) + b(i, j)
    Next j
    Next i
    
    add = mat
    
End Function

Function matI(n)

    Dim i As Integer, j As Integer
    Dim mat() As Integer
    ReDim mat(n, n)
    
    For i = 1 To n
        For j = 1 To n
        If i = j Then
            mat(i, j) = 1
        Else
            mat(i, j) = 0
        End If
    Next j
    Next i
    
    matI = mat
    
End Function

Function Hat(x)
    Hat = mm(x, mm(Inv(mm(t(x), x)), t(x)))
End Function

Function t(x)
    t = Application.WorksheetFunction.Transpose(x)
End Function

Function Inv(x)
    Inv = Application.WorksheetFunction.MInverse(x)
End Function

Function mm(x, y)
    mm = Application.WorksheetFunction.MMult(x, y)
End Function

Function pureY() As Variant

    Dim dataRange As Range
    Dim i As Long, j As Integer
    Dim y()
    
    Set dataRange = Worksheets(DataSheet).Cells.CurrentRegion
    
    ReDim y(n, 1)
    
    For j = 1 To m + 1
        If ylist = dataRange.Cells(1, j).value Then
            For i = 1 To n
                y(i, 1) = dataRange.Cells(i + 1, j).value
            Next i
        End If
    Next j
    
    pureY = y
    
End Function

''dataX(1,1)~dataX(N,p)�� �ڷ�
''xlist�� ��޵� ���������� ����Ÿ���� ������ �迭�� ��ȯ

Function pureX() As Variant

    Dim dataRange As Range
    Dim i As Long, j As Integer, k As Integer
    Dim x()
    
    Set dataRange = Worksheets(DataSheet).Cells.CurrentRegion
    ReDim x(n, p)
   
    For j = 1 To p
        For k = 1 To m + 1
            If xlist(j - 1) = dataRange.Cells(1, k) Then
                For i = 1 To n
                    x(i, j) = dataRange.Cells(i + 1, k).value
                Next i
            End If
        Next k
    Next j
    
    pureX = x
    
End Function

Function designX() As Variant

    Dim dataRange As Range
    Dim i As Long, j As Integer, k As Integer
    Dim x()
    
    Set dataRange = Worksheets(DataSheet).Cells.CurrentRegion
    ReDim x(n, p + 1)
    
    
    For i = 1 To n
        x(i, 1) = 1
    Next i
    
    For j = 1 To p
        For k = 1 To m + 1
            If xlist(j - 1) = dataRange.Cells(1, k) Then
                For i = 1 To n
                    x(i, j + 1) = dataRange.Cells(i + 1, k).value
                Next i
            End If
        Next k
    Next j
    
    designX = x
                
End Function



'Option Base 1
'Eigenvaluesevec(mat, 0.001) �� mat ��� ����� ����ġ�� 0.001�� ��Ȯ���� ���Ѵ�.
'EigenvectorsEmat(mat, 0.001) �� mat ��� ����� �������͸� 0.001�� ��Ȯ���� ���Ѵ�.
'exact=0.001 �� �θ� exact �� ����� �۾ƾ� sas��� ������ ����� ��´�.
'��Ī����϶��� �ȴ�. (�߿�!!!)
'�迭�� ��ȯ�Ѵ�.

' ��ó : Advanced Modelling in Finance using Excel and VBA, Mary Jackson and Mike Staunton
'        JOHN WILEY & SONS, LTD   , 2001



Function MatrixIdentity(n As Integer)
'   Returns the (nxn) Identity Matrix
    Dim i As Integer
    Dim Imat() As Double
    ReDim Imat(n, n)
    For i = 1 To n
        Imat(i, i) = 1
    Next i
    MatrixIdentity = Imat
End Function

Function MatrixTrace(Xmat)
'   Returns the trace of a matrix (sum of elements on leading diagonal)
    Dim sum
    Dim i As Integer, n As Integer
    n = Xmat.Columns.count
    sum = 0
    For i = 1 To n
        sum = sum + Xmat(i, i)
    Next i
    MatrixTrace = sum
End Function
    
Function MatrixUTSumSq(Xmat)
'   Returns the Sum of Squares of the Upper Triangle of a Matrix
    Dim sum
    Dim i As Integer, j As Integer, n As Integer
    n = Sqr(Application.count(Xmat))
    sum = 0
    For i = 1 To n
        For j = i + 1 To n
        sum = sum + (Xmat(i, j) ^ 2)
        Next j
    Next i
    MatrixUTSumSq = sum
End Function

Function Jacobirvec(n As Integer, Athis)
'   Returns vector containing mr, mc and jrad
'   These are the row and column vectors and the angle of rotation for the P matrix
    Dim maxval, jrad
    Dim i As Integer, j As Integer, mr As Integer, mc As Integer
    Dim Awork() As Variant
    ReDim Awork(n, n)
    maxval = -1
    mr = -1
    mc = -1
    For i = 1 To n
        For j = i + 1 To n
            Awork(i, j) = Abs(Athis(i, j))
            If Awork(i, j) > maxval Then
                maxval = Awork(i, j)
                mr = i
                mc = j
            End If
        Next j
    Next i
    If Athis(mr, mr) = Athis(mc, mc) Then
        jrad = 0.25 * Application.pi() * Sgn(Athis(mr, mc))
    Else
        jrad = 0.5 * Atn(2 * Athis(mr, mc) / (Athis(mr, mr) - Athis(mc, mc)))
    End If
    Jacobirvec = Array(mr, mc, jrad)
End Function

Function JacobiPmat(n As Integer, rthis)
'   Returns the rotation Pthis matrix
'   Uses MatrixIdentity fn
'   Uses Jacobirvec fn
    Dim Pthis As Variant
    Pthis = MatrixIdentity(n)
    Pthis(rthis(1), rthis(1)) = Cos(rthis(3))
    Pthis(rthis(2), rthis(1)) = sin(rthis(3))
    Pthis(rthis(1), rthis(2)) = -sin(rthis(3))
    Pthis(rthis(2), rthis(2)) = Cos(rthis(3))
    JacobiPmat = Pthis
End Function

Function JacobiAmat(n As Integer, Athis)
'   Returns Anext matrix, updated using the P rotation matrix
'   Uses Jacobirvec fn
'   Uses JacobiPmat fn
    Dim rthis As Variant, Pthis As Variant, Anext As Variant
    rthis = Jacobirvec(n, Athis)
    Pthis = JacobiPmat(n, rthis)
    Anext = Application.MMult(Application.Transpose(Pthis), Application.MMult(Athis, Pthis))
    JacobiAmat = Anext
End Function

Function Eigenvaluesevec(Amat, atol)
'   Uses the Jacobi method to get the eigenvalues for a symmetric matrix
'   Amat is rotated (using the P matrix) until its off-diagonal elements are minimal
'   Uses MatrixUTSumSq fn
'   Uses JacobiAmat fn
    Application.Volatile (False)
    Dim asumsq
    Dim i As Integer, n As Integer, r As Integer
    Dim evec() As Variant
    Dim Anext As Variant
    n = Sqr(Application.count(Amat))
    r = 0
    asumsq = MatrixUTSumSq(Amat)
    Do While asumsq > atol
        Anext = JacobiAmat(n, Amat)
        asumsq = MatrixUTSumSq(Anext)
        Amat = Anext
        r = r + 1
    Loop
    ReDim evec(n)
    For i = 1 To n
        evec(i) = Amat(i, i)
    Next i
    Eigenvaluesevec = evec
End Function

Function JacobiVmat(n As Integer, Athis, Vthis)
'   Returns Vnext matrix
'   Keeps track of the eigenvectors during the rotations
'   Uses Jacobirvec fn
'   Uses JacobiPmat fn
    Dim rthis As Variant, Pthis As Variant, Vnext As Variant
    rthis = Jacobirvec(n, Athis)
    Pthis = JacobiPmat(n, rthis)
    Vnext = Application.MMult(Vthis, Pthis)
    JacobiVmat = Vnext
End Function

Function EigenvectorsEmat(Amat, atol)
'   Uses the Jacobi method to get the eigenvectors for a symmetric matrix
'   Similar to eigenvalue function, but with additional V matrix updated with each rotation
'   Uses MatrixUTSumSq fn
'   Uses JacobiAmat fn
'   Uses JacobiVmat fn
'   Uses MatrixIdentity fn
    Application.Volatile (False)
    Dim asumsq
    Dim n As Integer, r As Integer
    Dim Anext As Variant, Vmat As Variant, Vnext As Variant
    n = Sqr(Application.count(Amat))
    r = 0
    Vmat = MatrixIdentity(n)
    asumsq = MatrixUTSumSq(Amat)
    Do While asumsq > atol
        Anext = JacobiAmat(n, Amat)
        Vnext = JacobiVmat(n, Amat, Vmat)
        asumsq = MatrixUTSumSq(Anext)
        Amat = Anext
        Vmat = Vnext
        r = r + 1
    Loop
    EigenvectorsEmat = Vnext
End Function




Sub FittedBand(OutSheetName, Left, Top, Width, Height, _
               Fitted, LCI, UCI, Alpha, XC, xname)
    
    Dim plot As ChartObject: Dim str1, str2, str As String
    Dim TempX, TempY, TempU, TempL As Range
    Dim TempSheet As Worksheet: Dim temp1, temp2 As Double
    
    '''������ �ձ� ���ؼ��� X������ �켱 ���� ���� �ӽý�Ʈ�� ����
    Set TempSheet = Worksheets.add: TempSheet.Visible = xlSheetHidden
    Set TempX = Range(TempSheet.Cells(1, 1), TempSheet.Cells(Fitted.count, 1))
    Set TempY = Range(TempSheet.Cells(1, 2), TempSheet.Cells(Fitted.count, 2))
    Set TempL = Range(TempSheet.Cells(1, 3), TempSheet.Cells(Fitted.count, 3))
    Set TempU = Range(TempSheet.Cells(1, 4), TempSheet.Cells(Fitted.count, 4))
    XC.Copy: TempSheet.Paste TempSheet.Cells(1, 1)
    Fitted.Copy: TempSheet.Paste TempSheet.Cells(1, 2)
    LCI.Copy: TempSheet.Paste TempSheet.Cells(1, 3)
    UCI.Copy: TempSheet.Paste TempSheet.Cells(1, 4)
    TempSheet.Cells(1, 1).Sort _
        Key1:=TempSheet.Cells(1, 1), _
        Order1:=xlAscending, Header:=xlGuess
    
    str = "������": str2 = Alpha & "%����": str1 = Alpha & "%����"
    
    Set plot = Worksheets(OutSheetName).ChartObjects.add(Left, Top, Width, Height)
    plot.Chart.ChartWizard Source:=TempY, _
        Gallery:=xlXYScatter, Format:=6, CategoryTitle:=xname
    plot.Chart.SeriesCollection(1).XValues = TempX
    plot.Chart.SeriesCollection.NewSeries: plot.Chart.SeriesCollection.NewSeries
    plot.Chart.SeriesCollection(2).Values = TempU
    plot.Chart.SeriesCollection(2).XValues = TempX
    plot.Chart.SeriesCollection(3).Values = TempL
    plot.Chart.SeriesCollection(3).XValues = TempX
    With plot.Chart.SeriesCollection(1)
        .Name = str
        .Border.ColorIndex = 11
        .Border.Weight = xlThin
    End With
    With plot.Chart.SeriesCollection(2)
        .Name = str1
        .Border.ColorIndex = 3
        .Border.Weight = xlThin
    End With
    With plot.Chart.SeriesCollection(3)
        .Name = str2
        .Border.ColorIndex = 3
        .Border.Weight = xlThin
    End With

    With plot.Chart.Axes(xlCategory)
        .TickLabelPosition = xlLow
        .TickLabels.Font.Size = 8
        .AxisTitle.Font.Size = 8
    End With

    With plot.Chart.Axes(xlValue)
        .TickLabels.Font.Size = 8
    End With
    
    With plot.Chart
        .HasTitle = True
        .ChartTitle.Characters.Text = "�ŷڴ� �׷���"
        .ChartTitle.Font.Bold = True
        .ChartTitle.Font.Size = 10
    End With
    plot.Chart.HasLegend = True
    With plot.Chart.Legend
        .position = xlBottom
        .Font.Size = 8
    End With
    temp1 = (Application.max(TempX) - Application.min(TempX)) / 10
    If temp1 <> 0 Then
        With plot.Chart.Axes(xlCategory)
            .MinimumScale = Application.min(TempX) - temp1
            .MaximumScale = Application.max(TempX) + temp1
            .TickLabels.NumberFormat = CStrNumPoint(temp1 * 10, TempX.count)
            .HasMajorGridlines = False
        End With
    End If
    
    temp2 = (Application.max(TempU) - Application.min(TempL)) / 10
    If temp2 <> 0 Then
        With plot.Chart.Axes(xlValue)
            .MinimumScale = Application.min(TempL) - temp2
            .MaximumScale = Application.max(TempU) + temp2
            .TickLabels.NumberFormat = CStrNumPoint(temp2 * 10, TempY.count)
        End With
    End If

    With plot.Chart.PlotArea.Border
        .Weight = xlThin
        .LineStyle = xlAutomatic
        .ColorIndex = 16
    End With
    
End Sub

Sub regScatterPlot(OutSheetName, Left, Top, Width, Height, _
               Fitted, XC, xname, y, YVarName)
    
    Dim plot As ChartObject
    Dim TempX, TempY, TempFitted As Range
    Dim TempSheet As Worksheet
    Dim TempMax, TempMin, temp1, temp2 As Double
    
    '''������ �ձ� ���ؼ��� X������ �켱 ���� ���� �ӽý�Ʈ�� ����
    Set TempSheet = Worksheets.add: TempSheet.Visible = xlSheetHidden
    Set TempX = Range(TempSheet.Cells(1, 1), TempSheet.Cells(Fitted.count, 1))
    Set TempY = Range(TempSheet.Cells(1, 2), TempSheet.Cells(Fitted.count, 2))
    Set TempFitted = Range(TempSheet.Cells(1, 3), TempSheet.Cells(Fitted.count, 3))

    XC.Copy: TempSheet.Paste TempSheet.Cells(1, 1)
    y.Copy: TempSheet.Paste TempSheet.Cells(1, 2)
    Fitted.Copy: TempSheet.Paste TempSheet.Cells(1, 3)
    
    TempSheet.Cells(1, 1).Sort _
        Key1:=TempSheet.Cells(1, 1), _
        Order1:=xlAscending, Header:=xlGuess
    
    Set plot = Worksheets(OutSheetName).ChartObjects.add(Left, Top, Width, Height)
    plot.Chart.ChartWizard Source:=TempY, _
        Gallery:=xlXYScatter, Format:=6, CategoryTitle:=xname, ValueTitle:=YVarName
    
    With plot.Chart.SeriesCollection(1)
        .XValues = TempX
        .Name = "���ڷ�"
        .Border.LineStyle = xlNone
        .MarkerBackgroundColorIndex = 3
        .MarkerForegroundColorIndex = 3
        .MarkerStyle = xlCircle
        .MarkerSize = 3
    End With
    
    plot.Chart.SeriesCollection.NewSeries
    plot.Chart.SeriesCollection(2).Values = TempFitted
    plot.Chart.SeriesCollection(2).XValues = TempX
    
    With plot.Chart.SeriesCollection(2)
        .Name = "ȸ�����ռ�"
        .Border.ColorIndex = 11
    End With

    With plot.Chart.Axes(xlCategory)
        .TickLabelPosition = xlLow
        .TickLabels.Font.Size = 8
        .AxisTitle.Font.Size = 8
    End With

    With plot.Chart.Axes(xlValue)
        .TickLabels.Font.Size = 8
        .AxisTitle.Font.Size = 8
        .AxisTitle.Orientation = xlVertical
    End With
    
    With plot.Chart
        .HasTitle = True
        .ChartTitle.Characters.Text = "������"
        .ChartTitle.Font.FontStyle = "����"
        .ChartTitle.Font.Size = 10
    End With
    plot.Chart.HasLegend = True
    With plot.Chart.Legend
        .position = xlBottom
        .Font.Size = 8
    End With
    temp1 = (Application.max(TempX) - Application.min(TempX)) / 10
    If temp1 <> 0 Then
        With plot.Chart.Axes(xlCategory)
            .MinimumScale = Application.min(TempX) - temp1
            .MaximumScale = Application.max(TempX) + temp1
            .TickLabels.NumberFormat = CStrNumPoint(temp1 * 10, TempX.count)
        End With
    End If
    
    TempMax = Application.max(TempY, TempFitted)
    TempMin = Application.min(TempY, TempFitted)
    temp2 = (TempMax - TempMin) / 10
    If temp2 <> 0 Then
        With plot.Chart.Axes(xlValue)
            .MinimumScale = TempMin - temp2
            .MaximumScale = TempMax + temp2
            .TickLabels.NumberFormat = CStrNumPoint(temp2 * 10, TempY.count)
        End With
    End If

End Sub
Function CStrNumPoint(DataWid, DataCount) As String
    
    Dim i As Integer: Dim LogScale As Double
    Dim temp As String
    
    i = 0: temp = "0."
    LogScale = Application.Power(10, _
             Int(Application.Log10(DataWid / DataCount)))
    If LogScale >= 1 Then
        CStrNumPoint = "0"
    Else
        Do
            temp = temp & "0": i = i + 1
            If LogScale = 10 ^ (-i) Then Exit Do
        Loop While (1)
        CStrNumPoint = CStr(temp)
    End If

End Function




Sub Diagnosis00(resi, intercept, ci, Alpha, simple)
    Dim k As Integer, check As Integer                'matrix X �� N*K
    Dim exact As Long, sum As Double
    Dim x(), y(), obs()
    Dim pt As Range
    Dim choice As Integer                               '��� ǥ���� �� ���� ���� -ci, e, r,r_
    Dim Flag As Long
    
    Dim id() As Integer
    Dim e(), r(), r_(), H(), diagH()         '����, ǥ��ȭ����, ǥ��ȭ��������, Hat matrix
    Dim DFFITS(), D(), CovR()
    Dim DFBETA(), vif()
    Dim DW, AutoRho
    Dim eval(), evec(), condNum(), stx()
    Dim varPro()
    Dim UCI(), LCI()
    Dim yhat()
    
    Dim s, s_()
    
    Dim mySheet As Worksheet
    Dim xPos As Double, yPos As Double
    Dim obsRn As Range, eRn As Range, rRn As Range, r_Rn As Range
    Dim yRn As Range, yhatRn As Range, lciRn As Range, uciRn As Range, xRn As Range
    Dim hiddenRange As Range
    Dim hiddenM As Long
    Dim tmpx()
    
    Dim simple_check As Integer
    
    
 
    ''''''�̻� ���� ����
 

 
    ''''''����Ÿ ����ֱ�
 
    
    y = pureY

    If intercept = True Then
        k = p + 1
        x = designX
    Else:
        k = p
        x = pureX
    End If
     
 
    ''''''ReDim
 
     
    ReDim e(n, 1), r(n, 1), r_(n, 1), DFFITS(n, 1), D(n, 1), CovR(n, 1)
    ReDim DFBETA(n, k), eval(k), evec(k, k), condNum(k, 1), varPro(k, k), vif(k, 1), stx(n, k)
    ReDim s_(n, 1), yhat(n, 1), diagH(n, 1), obs(n, 1), UCI(n, 1), LCI(n, 1)
    
  
    ''''''Hat, yhat, s ���ϰ�
  
    
    H = Hat(x)
    yhat = mm(H, y)
    id = matI(n)
    e = mm(diff(id, H), y)
    s = Application.WorksheetFunction.sum(mm(t(e), e)) / (n - k)        '����Ÿ������.
    s = s ^ 0.5
    matObs obs, n
    s_ = mats_(k, s, e, H)
    r = matr(e, s, H)
    r_ = matr_(e, s_, H)
    
    Set mySheet = Worksheets(DataSheet)
    Set pt = mySheet.Range("a1")

  '''
    '''''''ci ��� ���ÿ��ο� ����  yhat, ci��� ����Ÿ ��Ʈ�� �Ѹ���
  '''
    
       
    matCI x, UCI, Alpha, s, k
    LCI = diff(yhat, UCI)
    UCI = add(yhat, UCI)

    choice = 0
    If ci = True Then          '''
    
        
        mySheet.Range(pt.Cells(1, m + 2), pt.Cells(1, m + 2)) = "������"
        mySheet.Range(pt.Cells(2, m + 2), pt.Cells(n + 1, m + 2)) = yhat
    
        choice = choice + 1
        
        mySheet.Range(pt.Cells(1, m + 2 + choice), pt.Cells(1, m + 2 + choice)) = Alpha & "% ����"
        mySheet.Range(pt.Cells(2, m + 2 + choice), pt.Cells(n + 1, m + 2 + choice)) = LCI
        mySheet.Activate
        
        choice = choice + 1
        
        mySheet.Range(pt.Cells(1, m + 2 + choice), pt.Cells(1, m + 2 + choice)) = Alpha & "% ����"
        mySheet.Range(pt.Cells(2, m + 2 + choice), pt.Cells(n + 1, m + 2 + choice)) = UCI
        
    End If
    
  ''
    ''''''ȸ�����ܿ��� Title ������� �����ϰ� ����ϱ�
  ''
    check = 0
    For i = 1 To 17
        If resi(i) = True Then check = check + 1
    Next i
                        'check=0 �ε� p>1 �̸� resi(18)�� ������ exit or title, skip
                        'check=0 �ε� p=1 �̸� simple(1,2,3)�� ������ exit or title
    If check = 0 Then
        If p > 1 And resi(18) = False Then Exit Sub
        If p > 1 And resi(18) = True Then
           ' ModulePrint.Title1 ("ȸ�����ܰ��")
            GoTo SKIP
        End If
        If simple(1) = False And simple(2) = False And simple(3) = False Then Exit Sub
    End If
    
   ' ModulePrint.Title1 ("ȸ�����ܰ��")

 
    ''''''���� ��跮 ���ϱ�
 
    
    
    DFFITS = matdffit(H, r_)
    D = matD(H, k, r)
    CovR = matCovR(H, k, r_)
    DFBETA = matdfbeta(k, x, H, r_)
    
    

  '''''
    ''''''eigenValue, eigenVector, varianceProportion ���ϱ�
  '''''
    'eval(k), evec(k, k), condNum(k, 1), varpro(k, k), vif(k, 1)
    
If k = 1 Then
    eval(1) = 1
    evec(1, 1) = 1
    condNum(1, 1) = 1
    varPro(1, 1) = 1
    vif(1, 1) = 1
    GoTo shortSkip
End If

    
    standardizedX x, stx, k
    exact = 0.00000001
    eval = Eigenvaluesevec(mm(t(stx), stx), exact)
    evec = EigenvectorsEmat(mm(t(stx), stx), exact)
    calCondNum eval, condNum, k
    
    calVarPro eval, evec, varPro, k
    
    
    
    If intercept = True Then
        calVIF x, vif, k - 1
    Else
        calVIF11 x, vif, k - 1
    End If
    
shortSkip:

    DW = matDW(e)
    AutoRho = matAutoRho(e)
    
  '''
    '''''''���� ��� ���ÿ��ο� ���� ����Ÿ ��Ʈ�� �Ѹ���
  '''
    
    If resi(1) = True Then
    
        choice = choice + 1
        
        mySheet.Range(pt.Cells(1, m + 2 + choice), pt.Cells(1, m + 2 + choice)) = "����"
        mySheet.Range(pt.Cells(2, m + 2 + choice), pt.Cells(n + 1, m + 2 + choice)) = e
    
    End If
    
    If resi(6) = True Then
    
        choice = choice + 1
        
        mySheet.Range(pt.Cells(1, m + 2 + choice), pt.Cells(1, m + 2 + choice)) = "ǥ��ȭ����"
        mySheet.Range(pt.Cells(2, m + 2 + choice), pt.Cells(n + 1, m + 2 + choice)) = r
    
    End If
    
    If resi(11) = True Then
    
        choice = choice + 1
        
        mySheet.Range(pt.Cells(1, m + 2 + choice), pt.Cells(1, m + 2 + choice)) = "ǥ��ȭ��������"
        mySheet.Range(pt.Cells(2, m + 2 + choice), pt.Cells(n + 1, m + 2 + choice)) = r_
        
    End If
    
    ''' �ӽý�Ʈ�� ������ ���� �����Ѵ�.
    '''��ȣ     1   2   3   4   5   6       7   8   9
    '''��       a   b   c   d   e   f       g   h   i...
    '''����     obs e   r   r_  y   yhat    lci uci x...

    ModulePrint.MakeTmpSheet "_#TmpHIST1#_"
    Set mySheet = Worksheets("_#TmpHIST1#_")
    Set pt = mySheet.Range("a1")
    
    Set hiddenRange = mySheet.Cells.CurrentRegion
    hiddenM = hiddenRange.Cells(1, 1).End(xlToRight).Column       '������ ��ϵǾ� �ִ� ���� ��
    
    
    '���� ���� üũ �� �񱳰� ����
    Dim Cmp_Value As Integer
    
    If ModuleControl.ChkVersion(ActiveWorkbook.Name) = True Then
        Cmp_Value = 16300
    Else
        Cmp_Value = 250
    End If
    
    If hiddenM > Cmp_Value And hiddenRange.Cells(1, 1).value = "obs" Then  '������ ��ϵǾ� �ִ� ���� ���� �ʹ� ������
        MsgBox "�ʹ� ���� �۾��� ����Ǿ����ϴ�." & vbCrLf & "������ �ڷᰡ �����˴ϴ�.", vbOKOnly, "HIST"
        mySheet.Delete
        ModulePrint.MakeTmpSheet "_#TmpHIST1#_"
        Set mySheet = Worksheets("_#TmpHIST1#_")
        Set pt = Worksheets("_#TmpHIST1#_").Range("a1")
        Set hiddenRange = mySheet.Cells.CurrentRegion
        hiddenM = hiddenRange.Cells(1, 1).End(xlToRight).Column
    End If
    
    If hiddenM > Cmp_Value And hiddenRange.Cells(1, 1).value <> "obs" Then hiddenM = 0
    
                 
                    '''''''''ù�ٿ� ���񾲱�
                 
    pt.Cells(1, hiddenM + 1) = "obs"
    pt.Cells(1, hiddenM + 2) = "e"
    pt.Cells(1, hiddenM + 3) = "r"
    pt.Cells(1, hiddenM + 4) = "r_"
    pt.Cells(1, hiddenM + 5) = "y"
    pt.Cells(1, hiddenM + 6) = "yhat"
    pt.Cells(1, hiddenM + 7) = "lci"
    pt.Cells(1, hiddenM + 8) = "uci"
    
                 
                    '''''''''�ʿ��� ��跮 ����ϱ�
                 
    mySheet.Range(pt.Cells(2, hiddenM + 1), pt.Cells(n + 1, hiddenM + 1)) = obs
    mySheet.Range(pt.Cells(2, hiddenM + 2), pt.Cells(n + 1, hiddenM + 2)) = e
    mySheet.Range(pt.Cells(2, hiddenM + 3), pt.Cells(n + 1, hiddenM + 3)) = r
    mySheet.Range(pt.Cells(2, hiddenM + 4), pt.Cells(n + 1, hiddenM + 4)) = r_
    mySheet.Range(pt.Cells(2, hiddenM + 5), pt.Cells(n + 1, hiddenM + 5)) = y
    mySheet.Range(pt.Cells(2, hiddenM + 6), pt.Cells(n + 1, hiddenM + 6)) = yhat
    mySheet.Range(pt.Cells(2, hiddenM + 7), pt.Cells(n + 1, hiddenM + 7)) = LCI
    mySheet.Range(pt.Cells(2, hiddenM + 8), pt.Cells(n + 1, hiddenM + 8)) = UCI
    
                 ''''''''''''''''''''''''
                    '''''''''�׷��� �׸��� ���� Range�� ����ֱ�
                 ''''''''''''''''''''''''
    Set obsRn = mySheet.Range(pt.Cells(2, hiddenM + 1), pt.Cells(n + 1, hiddenM + 1))
    Set eRn = mySheet.Range(pt.Cells(2, hiddenM + 2), pt.Cells(n + 1, hiddenM + 2))
    Set rRn = mySheet.Range(pt.Cells(2, hiddenM + 3), pt.Cells(n + 1, hiddenM + 3))
    Set r_Rn = mySheet.Range(pt.Cells(2, hiddenM + 4), pt.Cells(n + 1, hiddenM + 4))
    Set yRn = mySheet.Range(pt.Cells(2, hiddenM + 5), pt.Cells(n + 1, hiddenM + 5))
    Set yhatRn = mySheet.Range(pt.Cells(2, hiddenM + 6), pt.Cells(n + 1, hiddenM + 6))
    Set lciRn = mySheet.Range(pt.Cells(2, hiddenM + 7), pt.Cells(n + 1, hiddenM + 7))
    Set uciRn = mySheet.Range(pt.Cells(2, hiddenM + 8), pt.Cells(n + 1, hiddenM + 8))
    
    tmp_position = 0
    

  '''
    '''''''simple �׸���
  '''
    If p = 1 Then
    
        simple_check = 0
        
        '''Title
        If simple(1) = True Or simple(2) = True Or simple(3) = True Then

            Set mySheet = Worksheets(RstSheet)
        
            Flag = mySheet.Cells(1, 1).value

            ModulePrint.Title3 "�ܼ� ���� ȸ��"
            mySheet.Cells(1, 1).value = mySheet.Cells(1, 1).value + 1       ' �� ����.
            simple_check = 1
        End If
        
        '''��跮 x ����ϱ�
        
        
        
        Set mySheet = Worksheets("_#TmpHIST1#_")
        Set pt = Worksheets("_#TmpHIST1#_").Range("a1")
        pt.Cells(1, hiddenM + 9) = "x"
                
        If k = p Then mySheet.Range(pt.Cells(2, hiddenM + 9), pt.Cells(n + 1, hiddenM + 9)) = x
        If k = p + 1 Then
            ReDim tmpx(n, 1)
            For i = 1 To n
                tmpx(i, 1) = x(i, 2)
            Next i
            mySheet.Range(pt.Cells(2, hiddenM + 9), pt.Cells(n + 1, hiddenM + 9)) = tmpx
        End If
        
        Set xRn = mySheet.Range(pt.Cells(2, hiddenM + 9), pt.Cells(n + 1, hiddenM + 9))
        
        
        '''�׷���
        Set mySheet = Worksheets(RstSheet)
        mySheet.Cells(1, 1).value = mySheet.Cells(1, 1).value + 1       ' �� ����.

        Flag = mySheet.Cells(1, 1).value
        Set pt = mySheet.Cells(Flag, 2)

        xPos = pt.Left: yPos = pt.Top - 15
        If simple(1) = True And IsNumeric(yRn(1, 1)) Then '''y vs x, ������
            Application.Run "Grap.xlam!ModuleScatter.ScatterPlot", _
                RstSheet, xPos, yPos, 200, 200, _
                xRn, yRn, "x", "y", False, "������"
            xPos = xPos + 210
        End If
    


        If Alpha <> 0 And simple(2) = True Then
            FittedBand RstSheet, xPos, yPos, 200, 200, yhatRn, _
                lciRn, uciRn, Alpha, xRn, xlist(0)
            xPos = xPos + 210

        End If
    
        'If simple(3) = True Then
        '    regScatterPlot RstSheet, xPos, yPos, 200, 200, yhatRn, _
        '        xRn, xlist(0), yRn, ylist
        '    xPos = xPos + 10: yPos = yPos + 30
        '    gaesoo = gaesoo + 1
        'End If
    
        If simple(3) = True And IsNumeric(xRn(1, 1)) And IsNumeric(rRn(1, 1)) Then '''ǥ��ȭ���� vs x
            Application.Run "Grap.xlam!ModuleScatter.ScatterPlot", _
                RstSheet, xPos, yPos, 200, 200, _
                xRn, rRn, "x", "ǥ��ȭ����", False, "ǥ��ȭ���� vs ��������"
            xPos = xPos + 210

        End If
        
        'If simple(1) = True Or simple(2) = True Or simple(3) = True Then
        'mySheet.Cells(1, 1).value = mySheet.Cells(1, 1).value + 21 + gaesoo
        'End If
        
        If simple_check = 1 Then mySheet.Cells(1, 1).value = mySheet.Cells(1, 1).value + 17
        
    End If
    
  '''
    '''''''"���� �׷���" Title ������� �����ϱ�
  '''
    If resi(2) = True Or resi(3) = True Or resi(4) = True Or resi(5) = True Then

        Set mySheet = Worksheets(RstSheet)
        
        Flag = mySheet.Cells(1, 1).value

        ModulePrint.Title3 "�����׷���"
        mySheet.Cells(1, 1).value = mySheet.Cells(1, 1).value + 2       ' �� ����.
        
        tmp_position = tmp_position + 1
        
    End If
    
  '''
    '''''''���� �׷��� ���ÿ��ο� ���� RstSheet �� ����ϱ�
  '''
    Set mySheet = Worksheets(RstSheet)

    Flag = mySheet.Cells(1, 1).value
    Set pt = mySheet.Cells(Flag, 2)

    xPos = pt.Left: yPos = pt.Top - 15
    gaesoo = 0
    If resi(2) = True And IsNumeric(eRn(1, 1)) Then '''���� vs ��������
        Application.Run "Grap.xlam!ModuleScatter.OrderScatterPlot", _
            RstSheet, xPos, yPos, 200, 200, _
            eRn, "����", 0
        xPos = xPos + 210
        gaesoo = gaesoo + 1
    End If
    
    If resi(3) = True And IsNumeric(yhatRn(1, 1)) And IsNumeric(e(1, 1)) Then '''���� vs ������
        Application.Run "Grap.xlam!ModuleScatter.ScatterPlot", _
            RstSheet, xPos, yPos, 200, 200, _
            yhatRn, eRn, "������", "����", False, "���� vs ������"
        xPos = xPos + 210
        gaesoo = gaesoo + 1
    End If
    
    If resi(4) = True And IsNumeric(e(1, 1)) Then '''���� ����Ȯ���׸�
        Application.Run "Grap.xlam!QQmodule.MainNormPlot", _
            eRn, xPos, yPos, Sheets(RstSheet), "����", True
        xPos = xPos + 210
        gaesoo = gaesoo + 1
    End If

    If resi(5) = True And IsNumeric(e(1, 1)) Then '''���� ������׷�
        Application.Run "Grap.xlam!Histmodule.MainHistogram", _
            eRn, xPos, yPos, Sheets(RstSheet), 0, "����"
        xPos = xPos + 210
        gaesoo = gaesoo + 1
    End If
    
    'If resi(2) = True Or resi(3) = True Or resi(4) = True Or resi(5) = True Then
    'mySheet.Cells(1, 1).value = mySheet.Cells(1, 1).value + 22 + gaesoo
    'End If
    
    If resi(2) = True Or resi(3) = True Or resi(4) = True Or resi(5) = True Then
        mySheet.Cells(1, 1).value = mySheet.Cells(1, 1).value + 17
    End If
    
    
  '''
    '''''''"ǥ��ȭ ���� �׷���" Title ������� �����ϱ�
  '''
    If resi(7) = True Or resi(8) = True Or resi(9) = True Or resi(10) = True Then

        Set mySheet = Worksheets(RstSheet)
        
        Flag = mySheet.Cells(1, 1).value

        ModulePrint.Title3 "ǥ��ȭ����"
        mySheet.Cells(1, 1).value = mySheet.Cells(1, 1).value + 2       ' �� ����.
        tmp_position = tmp_position + 1
        
    End If
    
  '''
    '''''''ǥ��ȭ ���� �׷��� ���ÿ��ο� ���� RstSheet �� ����ϱ�
  '''
    

    Flag = mySheet.Cells(1, 1).value
    Set pt = mySheet.Cells(Flag, 2)

    xPos = pt.Left: yPos = pt.Top - 15
    gaesoo = 0
    If resi(7) = True And IsNumeric(rRn(1, 1)) Then '''ǥ��ȭ���� vs ��������
        Application.Run "Grap.xlam!ModuleScatter.OrderScatterPlot", _
            RstSheet, xPos, yPos, 200, 200, _
            rRn, "ǥ��ȭ����", 0
        xPos = xPos + 210
        gaesoo = gaesoo + 1
    End If

    If resi(8) = True And IsNumeric(yhatRn(1, 1)) And IsNumeric(r(1, 1)) Then '''ǥ��ȭ���� vs ������
        Application.Run "Grap.xlam!ModuleScatter.ScatterPlot", _
            RstSheet, xPos, yPos, 200, 200, _
            yhatRn, rRn, "������", "ǥ��ȭ����", False, "ǥ��ȭ���� vs ������"
        xPos = xPos + 210
        gaesoo = gaesoo + 1
    End If
    
    If resi(9) = True And IsNumeric(r(1, 1)) Then '''ǥ��ȭ���� ����Ȯ���׸�
        Application.Run "Grap.xlam!QQmodule.MainNormPlot", _
            rRn, xPos, yPos, Sheets(RstSheet), "ǥ��ȭ����", True
        xPos = xPos + 210
        gaesoo = gaesoo + 1
    End If

    If resi(10) = True And IsNumeric(r(1, 1)) Then '''ǥ��ȭ���� ������׷�
        Application.Run "Grap.xlam!Histmodule.MainHistogram", _
            rRn, xPos, yPos, Sheets(RstSheet), 0, "ǥ��ȭ����"
        xPos = xPos + 210
        gaesoo = gaesoo + 1
    End If
    
    'If resi(7) = True Or resi(8) = True Or resi(9) = True Or resi(10) = True Then
    'mySheet.Cells(1, 1).value = mySheet.Cells(1, 1).value + 22 + gaesoo
    'End If
    
    If resi(7) = True Or resi(8) = True Or resi(9) = True Or resi(10) = True Then
        mySheet.Cells(1, 1).value = mySheet.Cells(1, 1).value + 17
    End If
    
  '''
    '''''''"ǥ��ȭ ���� ���� �׷���" Title ������� �����ϱ�
  '''
    If resi(12) = True Or resi(13) = True Or resi(14) = True Or resi(15) = True Then

        Set mySheet = Worksheets(RstSheet)
        
        Flag = mySheet.Cells(1, 1).value

        ModulePrint.Title3 "ǥ��ȭ��������"
        mySheet.Cells(1, 1).value = mySheet.Cells(1, 1).value + 2       ' �� ����.
        
        tmp_position = tmp_position + 1
        
    End If
    
  '''
    '''''''ǥ��ȭ�������� �׷��� ���ÿ��ο� ���� RstSheet �� ����ϱ�
  '''
    

    Flag = mySheet.Cells(1, 1).value
    Set pt = mySheet.Cells(Flag, 2)

    xPos = pt.Left: yPos = pt.Top - 15
    'gaesoo = 0
    If resi(12) = True And IsNumeric(r_Rn(1, 1)) Then '''ǥ��ȭ�������� vs ��������
        Application.Run "Grap.xlam!ModuleScatter.OrderScatterPlot", _
            RstSheet, xPos, yPos, 200, 200, _
            r_Rn, "ǥ��ȭ��������", 0
        xPos = xPos + 210
       ' gaesoo = gaesoo + 1
    End If

    If resi(13) = True And IsNumeric(yhatRn(1, 1)) And IsNumeric(r_(1, 1)) Then '''ǥ��ȭ�������� vs ������
        Application.Run "Grap.xlam!ModuleScatter.ScatterPlot", _
            RstSheet, xPos, yPos, 200, 200, _
            yhatRn, r_Rn, "������", "ǥ��ȭ��������", False, "ǥ��ȭ�������� vs ������"
        xPos = xPos + 210
        'gaesoo = gaesoo + 1
    End If
    
    If resi(14) = True And IsNumeric(r_(1, 1)) Then '''ǥ��ȭ�������� ����Ȯ���׸�
        Application.Run "Grap.xlam!QQmodule.MainNormPlot", _
            r_Rn, xPos, yPos, Sheets(RstSheet), "ǥ��ȭ��������", True
        xPos = xPos + 210
        'gaesoo = gaesoo + 1
    End If

    If resi(15) = True And IsNumeric(r_(1, 1)) Then '''ǥ��ȭ�������� ������׷�
        Application.Run "Grap.xlam!Histmodule.MainHistogram", _
            r_Rn, xPos, yPos, Sheets(RstSheet), 0, "ǥ��ȭ��������"
        xPos = xPos + 210
        'gaesoo = gaesoo + 1
    End If
    
    'If resi(12) = True Or resi(13) = True Or resi(14) = True Or resi(15) = True Then
    'mySheet.Cells(1, 1).value = mySheet.Cells(1, 1).value + 22 + gaesoo
    'tmp_position = tmp_position + 1
    'End If
    If resi(12) = True Or resi(13) = True Or resi(14) = True Or resi(15) = True Then
        mySheet.Cells(1, 1).value = mySheet.Cells(1, 1).value + 17
    End If
    
    
    If tmp_position > 0 Then
        mySheet.Cells(1, 1).value = mySheet.Cells(1, 1).value + 3
        Flag = mySheet.Cells(1, 1).value
        Set pt = mySheet.Cells(Flag, 2)
        
        ModulePrint.TABLE 2, 3, 0
        
        pt.Cells(1, 1) = "DW ��跮"
        pt.Cells(1, 2) = "1�� �ڱ������"
        pt.Cells(2, 1) = DW
        pt.Cells(2, 2) = AutoRho
        pt.Cells(2, 1).Resize(1, 2).NumberFormatLocal = "0.0000_ "
        pt.Cells(1, 2).Resize(1, 1).HorizontalAlignment = xlLeft
        
        pt.Cells(4, 1) = "�������� ���� �ڱ����� ������ DW��跮�� 0 �� ����� ���� ����"
        pt.Cells(5, 1) = "���� �ڱ����� ������ 4 �� ����� ���� ���� �ȴ�."
        pt.Cells(4, 1).Resize(2, 1).HorizontalAlignment = xlLeft
        
        mySheet.Cells(1, 1) = mySheet.Cells(1, 1) + 6
    End If


 '''
    '''''''���߰�����
 '''
    If resi(17) = True Then
        
        Set mySheet = Worksheets(RstSheet)

        Flag = mySheet.Cells(1, 1).value
        Set pt = mySheet.Cells(Flag, 2)

    
        matObs obs, k
        
        ModulePrint.Title3 "���� ������"
        
        mySheet.Cells(1, 1).value = mySheet.Cells(1, 1).value + 1       '�� �� ����.
        
        Flag = mySheet.Cells(1, 1).value
        Set pt = mySheet.Cells(Flag, 2)
        
        ModulePrint.TABLE k + 1, 2, 0
        
        pt.Cells(1, 1) = "������"
        pt.Cells(1, 2) = "�л���â" & vbLf & "����"
        If k = p + 1 Then
            pt.Cells(2, 1) = "�����"
            For i = 1 To k - 1
            pt.Cells(2 + i, 1) = xlist(i - 1)
            Next i
        End If
        
        If k = p Then
            pt.Cells(2, 1) = xlist(0)
            For i = 1 To k - 1
            pt.Cells(2 + i, 1) = xlist(i)
            Next i
        End If
        
        mySheet.Range(pt.Cells(2, 2), pt.Cells(k + 1, 2)) = vif
        pt.Cells(2, 2).Resize(k + 1, 2).NumberFormatLocal = "0.0000_ "

        pt.Cells(k + 3, 1) = "�л���â���� > 10 �̸� ���߰������� �ɰ��� ������ �ִٰ� �����Ѵ�."
        pt.Cells(k + 3, 1).Resize(1, 1).HorizontalAlignment = xlLeft
        
        mySheet.Cells(1, 1).value = mySheet.Cells(1, 1).value + k + 4
        
        Flag = mySheet.Cells(1, 1).value
        Set pt = mySheet.Cells(Flag, 2)
        
        ModulePrint.TABLE k + 1, k + 3, 0
        
        pt.Cells(1, 1) = "��ȣ"
        pt.Cells(1, 2) = "������"
        pt.Cells(1, 3) = "��������"
        
        If k = p + 1 Then
            pt.Cells(1, 4) = "�л����" & vbLf & "�����"
            For i = 1 To k - 1
            pt.Cells(1, 4 + i) = vbLf & xlist(i - 1)
            Next i
        End If
        
        If k = p Then
            pt.Cells(1, 4) = "�л����" & vbLf & xlist(0)
            For i = 1 To k - 1
            pt.Cells(1, 4 + i) = vbLf & xlist(i)
            Next i
        End If
        
        
        mySheet.Range(pt.Cells(2, 2), pt.Cells(k + 1, 2)) = t(eval)
        mySheet.Range(pt.Cells(2, 3), pt.Cells(k + 1, 3)) = condNum
        mySheet.Range(pt.Cells(2, 4), pt.Cells(k + 1, 3 + k)) = varPro
        
     ''''''
        ''''''''''Sorting'''''''''''''
     ''''''
        
        mySheet.Activate
        mySheet.Range(pt.Cells(2, 2), pt.Cells(2, 2)).Select
        Application.CutCopyMode = False
        Selection.Sort Key1:=Range(pt.Cells(2, 2), pt.Cells(2, 2)), Order1:=xlDescending, Header:=xlGuess, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
        
        mySheet.Range(pt.Cells(2, 1), pt.Cells(k + 1, 1)) = obs
        
        pt.Cells(k + 3, 1) = "�������� ��� ũ���� 1�� ���� �ɰ��ϰ� ���� ���"
        pt.Cells(k + 4, 1) = "���������� ���� 10���� Ŭ ���"
        pt.Cells(k + 5, 1) = "�л������ 80-90% �̻����� ��Ÿ���� �������� ������ �� �̻��� ���"
        pt.Cells(k + 6, 1) = "���߰������� ������ �ִٰ� �����Ѵ�."
        pt.Cells(k + 3, 1).Resize(4, 1).HorizontalAlignment = xlLeft
        
        pt.Cells(2, 2).Resize(k + 1, k + 3).NumberFormatLocal = "0.0000_ "
        mySheet.Cells(1, 1) = mySheet.Cells(1, 1) + k + 7
        
        
    End If
    
    

    
 '''
    '''''''���������
 '''


    If resi(16) = True Then
    
        Set mySheet = Worksheets(RstSheet)
    
        Flag = mySheet.Cells(1, 1).value
        Set pt = mySheet.Cells(Flag, 2)

        diagH = matDiagH(H)
        
        ModulePrint.Title3 "���� ������"
        
        mySheet.Cells(1, 1).value = mySheet.Cells(1, 1).value + 1       '�� �� ����.
        
        ModulePrint.TABLE n + 1, k + 5, 0
        
        Flag = mySheet.Cells(1, 1).value
        Set pt = mySheet.Cells(Flag, 2)

        
        pt.Cells(1, 1) = "��ȣ"
        pt.Cells(1, 2) = "Hat" & vbLf & "Diagonal"
        pt.Cells(1, 3) = "Cook's" & vbLf & "Distance"
        pt.Cells(1, 4) = "���л����"
        pt.Cells(1, 5) = "DFFITS"
        
        If k = p + 1 Then
            pt.Cells(1, 6) = "DFBETAS" & vbLf & "�����"
            For i = 1 To k - 1
            pt.Cells(1, 6 + i) = vbLf & xlist(i - 1)
            Next i
        End If
        
        If k = p Then
            pt.Cells(1, 6) = "DFBETAS" & vbLf & xlist(0)
            For i = 1 To k - 1
            pt.Cells(1, 6 + i) = vbLf & xlist(i)
            Next i
        End If
        
        mySheet.Range(pt.Cells(2, 1), pt.Cells(n + 1, 1)) = obs
        mySheet.Range(pt.Cells(2, 2), pt.Cells(n + 1, 2)) = diagH
        mySheet.Range(pt.Cells(2, 3), pt.Cells(n + 1, 3)) = D
        mySheet.Range(pt.Cells(2, 4), pt.Cells(n + 1, 4)) = CovR
        mySheet.Range(pt.Cells(2, 5), pt.Cells(n + 1, 5)) = DFFITS
        mySheet.Range(pt.Cells(2, 6), pt.Cells(n + 1, 5 + k)) = DFBETA
        
        pt.Cells(2, 2).Resize(n + 1, 5 + k).NumberFormatLocal = "0.0000_ "
        
        mySheet.Cells(1, 1) = mySheet.Cells(1, 1) + n + 2
        
        '''�������� �Ķ������� ��Ÿ����
        
        For i = 1 To n
            If pt.Cells(1 + i, 2) > 2 * p / n Then pt.Cells(1 + i, 2).Font.ColorIndex = 41
        Next i
        
        For i = 1 To n
            If pt.Cells(1 + i, 3) > 1 Then pt.Cells(1 + i, 3).Font.ColorIndex = 41
        Next i
        
        For i = 1 To n
            If Abs(pt.Cells(1 + i, 4) - 1) >= 3 * (p + 1) / n Then pt.Cells(1 + i, 4).Font.ColorIndex = 41
        Next i
        
        For i = 1 To n
            If pt.Cells(1 + i, 5) > 2 * Sqr(k / n) Or pt.Cells(1 + i, 5) > 2 Then pt.Cells(1 + i, 5).Font.ColorIndex = 41
        Next i
        
        If n < 10 Then              '���� data
        For j = 1 To k
        For i = 1 To n
            If pt.Cells(1 + i, 4 + j) > 2 Then pt.Cells(1 + i, 4 + j).Font.ColorIndex = 41
        Next i
        Next j
        
        Else                        'ū data
        For j = 1 To k
        For i = 1 To n
            If pt.Cells(1 + i, 4 + j) > 2 / Sqr(n) Then pt.Cells(1 + i, 4 + j).Font.ColorIndex = 41
        Next i
        Next j
        
        End If
        
    End If
    

        
    
SKIP:
    If resi(18) = False Or k = 1 Then Exit Sub
    partialRegPlot00 k, y, x
    
End Sub

Sub partialRegPlot00(k, y, x)
    
    Dim tmpx(), index(), e_p_y(), e_p_x(), Xk(), H()
    Dim id() As Integer
    Dim i As Integer
    Dim start As Integer
    Dim xPos As Double, yPose As Double
    Dim tmpRnY As Range, tmpRnX As Range
    
    ReDim index(k) ', H(n, n), id(n, n), e_p_y(n, 1), e_p_x(n, 1)
    id = matI(n)
    
    ModulePrint.MakeTmpSheet "_#TmpHIST1#_"
    Set mySheet = Worksheets("_#TmpHIST1#_")
    Set pt = Worksheets("_#TmpHIST1#_").Range("a1")
    
    Set hiddenRange = mySheet.Cells.CurrentRegion
    hiddenM = hiddenRange.Cells(1, 1).End(xlToRight).Column       '������ ��ϵǾ� �ִ� ���� ��
    
    '���� ���� üũ �� �񱳰� ����
    Dim Cmp_Value As Integer
    
    If ModuleControl.ChkVersion(ActiveWorkbook.Name) = True Then
        Cmp_Value = 16300
    Else
        Cmp_Value = 250
    End If
    
    If hiddenM > Cmp_Value And hiddenRange.Cells(1, 1).value = "obs" Then                                       '������ ��ϵǾ� �ִ� ���� ���� �ʹ� ������
        MsgBox "�ʹ� ���� �۾��� ����Ǿ����ϴ�." & vbCrLf & "������ �ڷᰡ �����˴ϴ�.", vbOKOnly, "HIST"
        mySheet.Delete
        ModulePrint.MakeTmpSheet "_#TmpHIST1#_"
        Set mySheet = Worksheets("_#TmpHIST1#_")
        Set pt = Worksheets("_#TmpHIST1#_").Range("a1")
        hiddenM = dataRange.Cells(1, 1).End(xlToRight).Column
    End If
         
    If hiddenM > Cmp_Value And hiddenRange.Cells(1, 1).value <> "obs" Then hiddenM = 0                                    '������ ��ϵǾ� �ִ� ���� ���� �ʹ� ������
  
    For i = 1 To k
        
        index = makeIndex(k, 1)
        index(i) = 0
        tmpx = selectedX(index, x, k)
            
        index = makeIndex(k, 0)
        index(i) = 1
        Xk = selectedX(index, x, k)
            
        H = Hat(tmpx)
        e_p_y = mm(diff(id, H), y)
        e_p_x = mm(diff(id, H), Xk)
            
                ''' ���� ��� ���� , ó�� k��  y�κ�ȸ�Ͱ�, ���� k�� x�κ�ȸ�Ͱ�
                
        mySheet.Range(pt.Cells(1, hiddenM + i), pt.Cells(1, hiddenM + i)) = "pp"
        mySheet.Range(pt.Cells(2, hiddenM + i), pt.Cells(n + 1, hiddenM + i)) = e_p_y
        mySheet.Range(pt.Cells(1, hiddenM + k + i), pt.Cells(1, hiddenM + k + i)) = "pp"
        mySheet.Range(pt.Cells(2, hiddenM + k + i), pt.Cells(n + 1, hiddenM + k + i)) = e_p_x
        
    Next i
        
  '''
    '''''''"�κ�ȸ�ͻ�����" Title ���
  '''
    ModulePrint.Title3 "�κ�ȸ�ͻ�����"
    
  '''
    '''''''�κ�ȸ�ͻ����� RstSheet �� ����ϱ�
  '''
    
    Worksheets(RstSheet).Cells(1, 1).value = Worksheets(RstSheet).Cells(1, 1).value + 1       ' �� ����.

    Flag = Worksheets(RstSheet).Cells(1, 1).value
    Set pt = Worksheets(RstSheet).Cells(Flag, 2)

    xPos = pt.Left: yPos = pt.Top
    gaesoo = 0

    Set mySheet = Worksheets("_#TmpHIST1#_")
    Set pt = Worksheets("_#TmpHIST1#_").Range("a1")

    If k = p + 1 Then
        
            Set tmpRnY = mySheet.Range(pt.Cells(2, hiddenM + 1), pt.Cells(n + 1, hiddenM + 1))
            Set tmpRnX = mySheet.Range(pt.Cells(2, hiddenM + 1 + k), pt.Cells(n + 1, hiddenM + 1 + k))
        
            Application.Run "Grap.xlam!ModuleScatter.ScatterPlot", _
                RstSheet, xPos, yPos, 200, 200, _
                tmpRnX, tmpRnY, "�����", ylist, False, "�κ�ȸ�ͻ�����"
            xPos = xPos + 210


        For i = 2 To k
            Set tmpRnY = mySheet.Range(pt.Cells(2, hiddenM + i), pt.Cells(n + 1, hiddenM + i))
            Set tmpRnX = mySheet.Range(pt.Cells(2, hiddenM + i + k), pt.Cells(n + 1, hiddenM + i + k))
        
            Application.Run "Grap.xlam!ModuleScatter.ScatterPlot", _
                RstSheet, xPos, yPos, 200, 200, _
                tmpRnX, tmpRnY, xlist(i - 2), ylist, False, "�κ�ȸ�ͻ�����"
            xPos = xPos + 210
            gaesoo = gaesoo + 1
        Next i
        
    End If
    
    If k = p Then
        
        For i = 1 To k
            Set tmpRnY = mySheet.Range(pt.Cells(2, hiddenM + i), pt.Cells(n + 1, hiddenM + i))
            Set tmpRnX = mySheet.Range(pt.Cells(2, hiddenM + i + k), pt.Cells(n + 1, hiddenM + i + k))
        
            Application.Run "Grap.xlam!ModuleScatter.ScatterPlot", _
                RstSheet, xPos, yPos, 200, 200, _
                tmpRnX, tmpRnY, xlist(i - 1), ylist, False, "�κ�ȸ�ͻ�����"
            xPos = xPos + 210
            gaesoo = gaesoo + 1
        Next i
        
    End If
        
    Worksheets(RstSheet).Cells(1, 1).value = Worksheets(RstSheet).Cells(1, 1).value + 21
    
End Sub
