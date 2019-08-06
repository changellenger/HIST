Attribute VB_Name = "ModuleMatrix"
''N은 data개수
''종속변수의 데이타만을 배열로 반환
''pureY(0)~pureY(N-1)은 자료

Function pureY() As Variant

    Dim dataRange As Range
    Dim i As Long, j As Integer
    Dim Y()
    
    Set dataRange = Worksheets(DataSheet).Cells.CurrentRegion
    
    ReDim Y(N - 1, 0)
    For j = 0 To m
        If ylist = dataRange.Cells(1, j + 1).value Then
            For i = 1 To N
                Y(i - 1, 0) = dataRange.Cells(i + 1, j + 1).value
            Next i
        End If
    Next j
    
    pureY = Y
    
End Function

''dataX(0,0)~dataX(N-1,p-1)은 자료
''xlist에 언급된 독립변수의 데이타만을 이차원 배열로 반환

Function pureX() As Variant

    Dim dataRange As Range
    Dim i As Long, j As Integer, K As Integer
    Dim x()
    
    Set dataRange = Worksheets(DataSheet).Cells.CurrentRegion
    ReDim x(N - 1, p - 1)
   
    For j = 0 To p - 1
        For K = 0 To m
            If xlist(j) = dataRange.Cells(1, K + 1) Then
                For i = 0 To N - 1
                    x(i, j) = dataRange.Cells(i + 2, K + 1).value
                Next i
            End If
        Next K
    Next j
    
    pureX = x
    
End Function

Function designX() As Variant

    Dim dataRange As Range
    Dim i As Long, j As Integer, K As Integer
    Dim x()
    
    Set dataRange = Worksheets(DataSheet).Cells.CurrentRegion
    ReDim x(N - 1, p)
    
    
    For i = 0 To N - 1
        x(i, 0) = 1
    Next i
    
    For j = 0 To p - 1
        For K = 0 To m
            If xlist(j) = dataRange.Cells(1, K + 1) Then
                For i = 0 To N - 1
                    x(i, j + 1) = dataRange.Cells(i + 2, K + 1).value
                Next i
            End If
        Next K
    Next j
    
    designX = x
                
End Function

''index가 0이 아닌 독립변수들의 data만의 배열을 돌려준다.
''이때 x 는 pureX 이다
Function selectedX(Index, x)
    Dim p1 As Integer, j As Integer, K As Integer
    Dim tmpx()
    
    p1 = 0
    For j = 0 To p - 1
        If Index(j) <> 0 Then p1 = p1 + 1
    Next j
    
    ReDim tmpx(N - 1, p1 - 1)
    j = 0
    For K = 0 To p - 1
        If Index(K) <> 0 Then
        For i = 0 To N - 1
            tmpx(i, j) = x(i, K)
        Next i
        j = j + 1
        End If
    Next K
    
    selectedX = tmpx
    
End Function

Function T(x)
    T = Application.WorksheetFunction.Transpose(x)
End Function

Function Inv(x)
    Inv = Application.WorksheetFunction.MInverse(x)
End Function

Function mm(x, Y)
    mm = Application.WorksheetFunction.MMult(x, Y)
End Function

Function noIntSSR(Y, x)         '절편이 없는 경우의 SSR을 반환
    Dim H()
    Dim hatY, tmpY
    Dim meanY As Double
    Dim i As Long
    
    H = mm(x, mm(Inv(mm(T(x), x)), T(x)))
    hatY = mm(H, Y)
        
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

Function makeIndex(K, value)
    Dim i As Integer
    Dim tmpIndex()
    
    ReDim tmpIndex(K)
    
    For i = 0 To K
        tmpIndex(i) = value
    Next i
    
    makeIndex = tmpIndex
End Function

Function fullModelMSE(Y, x, intercept)
    Dim tmpMSE As Double
    
    If intercept <> 0 Then
        tmpMSE = Application.WorksheetFunction.Index( _
                Application.WorksheetFunction.LinEst(Y, x, 1, 1), 5, 2) / (N - p - 1)
    Else
        tmpMSE = Application.WorksheetFunction.Index( _
                Application.WorksheetFunction.LinEst(Y, x, 0, 1), 5, 2) / (N - p)
    End If
    
    fullModelMSE = tmpMSE
End Function
