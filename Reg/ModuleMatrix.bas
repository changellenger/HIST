Attribute VB_Name = "ModuleMatrix"
''N�� data����
''���Ӻ����� ����Ÿ���� �迭�� ��ȯ
''pureY(0)~pureY(N-1)�� �ڷ�

Function pureY() As Variant

    Dim dataRange As Range
    Dim i As Long, j As Integer
    Dim y()
    
    Set dataRange = Worksheets(DataSheet).Cells.CurrentRegion
    
    ReDim y(n - 1, 0)
    For j = 0 To m
        If ylist = dataRange.Cells(1, j + 1).value Then
            For i = 1 To n
                y(i - 1, 0) = dataRange.Cells(i + 1, j + 1).value
            Next i
        End If
    Next j
    
    pureY = y
    
End Function

''dataX(0,0)~dataX(N-1,p-1)�� �ڷ�
''xlist�� ��޵� ���������� ����Ÿ���� ������ �迭�� ��ȯ

Function pureX() As Variant

    Dim dataRange As Range
    Dim i As Long, j As Integer, k As Integer
    Dim x()
    
    Set dataRange = Worksheets(DataSheet).Cells.CurrentRegion
    ReDim x(n - 1, p - 1)
   
    For j = 0 To p - 1
        For k = 0 To m
            If xlist(j) = dataRange.Cells(1, k + 1) Then
                For i = 0 To n - 1
                    x(i, j) = dataRange.Cells(i + 2, k + 1).value
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
    ReDim x(n - 1, p)
    
    
    For i = 0 To n - 1
        x(i, 0) = 1
    Next i
    
    For j = 0 To p - 1
        For k = 0 To m
            If xlist(j) = dataRange.Cells(1, k + 1) Then
                For i = 0 To n - 1
                    x(i, j + 1) = dataRange.Cells(i + 2, k + 1).value
                Next i
            End If
        Next k
    Next j
    
    designX = x
                
End Function

''index�� 0�� �ƴ� ������������ data���� �迭�� �����ش�.
''�̶� x �� pureX �̴�
Function selectedX(index, x)
    Dim p1 As Integer, j As Integer, k As Integer
    Dim tmpx()
    
    p1 = 0
    For j = 0 To p - 1
        If index(j) <> 0 Then p1 = p1 + 1
    Next j
    
    ReDim tmpx(n - 1, p1 - 1)
    j = 0
    For k = 0 To p - 1
        If index(k) <> 0 Then
        For i = 0 To n - 1
            tmpx(i, j) = x(i, k)
        Next i
        j = j + 1
        End If
    Next k
    
    selectedX = tmpx
    
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

Function noIntSSR(y, x)         '������ ���� ����� SSR�� ��ȯ
    Dim H()
    Dim hatY, tmpY
    Dim meanY As Double
    Dim i As Long
    
    H = mm(x, mm(Inv(mm(t(x), x)), t(x)))
    hatY = mm(H, y)
        
    noIntSSR = Application.WorksheetFunction.SumSq(hatY)
    
End Function

Function binStr(number As Long) As String
    If number > 2 ^ 32 Then Exit Function
    If number > 1 Then
        binStr = binStr(number \ 2) & (number Mod 2)
    Else
        binStr = number
    End If
End Function

Function makeIndex(k, value)
    Dim i As Integer
    Dim tmpIndex()
    
    ReDim tmpIndex(k)
    
    For i = 0 To k
        tmpIndex(i) = value
    Next i
    
    makeIndex = tmpIndex
End Function

Function fullModelMSE(y, x, intercept)
    Dim tmpMSE As Double
    
    If intercept <> 0 Then
        tmpMSE = Application.WorksheetFunction.index( _
                Application.WorksheetFunction.LinEst(y, x, 1, 1), 5, 2) / (n - p - 1)
    Else
        tmpMSE = Application.WorksheetFunction.index( _
                Application.WorksheetFunction.LinEst(y, x, 0, 1), 5, 2) / (n - p)
    End If
    
    fullModelMSE = tmpMSE
End Function
