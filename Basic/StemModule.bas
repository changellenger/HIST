Attribute VB_Name = "StemModule"
Option Private Module
Option Explicit

Public StemUnit, X As Double
Public NowStemCount As Integer

'''Sorting 알고리즘 재계산 방지를 위해서
Public pdata(), q1, q3, IQR As Double
Public Obs As Long
Public DiffBug As Boolean

'''
'''Quick Sorting Algorithm
'''
Sub procSort1D(avArray, iLow1 As Long, iHigh1 As Long)

    On Error Resume Next
    
    'Dimension variables
    Dim iLow2 As Long, iHigh2 As Long, i As Long
    Dim vItem1, vItem2 As Variant
    
    'Set new extremes to old extremes
    iLow2 = iLow1
    iHigh2 = iHigh1
    
    'Get value of array item in middle of new extremes
    vItem1 = avArray((iLow1 + iHigh1) \ 2)
    
    'Loop for all the items in the array between the extremes
    While iLow2 < iHigh2
           
        'Find the first item that is greater than the mid-point item
         While avArray(iLow2) < vItem1 And iLow2 < iHigh1
             iLow2 = iLow2 + 1
         Wend
    
         'Find the last item that is less than the mid-point item
         While avArray(iHigh2) > vItem1 And iHigh2 > iLow1
             iHigh2 = iHigh2 - 1
         Wend
    
        'If the two items are in the wrong order, swap the rows
        If iLow2 < iHigh2 Then
            vItem2 = avArray(iLow2)
            avArray(iLow2) = avArray(iHigh2)
            avArray(iHigh2) = vItem2
        End If
    
        'If the pointers are not together, advance to the next item
        If iLow2 <= iHigh2 Then
            iLow2 = iLow2 + 1
            iHigh2 = iHigh2 - 1
        End If
    Wend
    
    'Recurse to sort the lower half of the extremes
    If iHigh2 > iLow1 Then procSort1D avArray, iLow1, iHigh2
    
    'Recurse to sort the upper half of the extremes
    If iLow2 < iHigh1 Then procSort1D avArray, iLow2, iHigh1
    
End Sub

'''
'''sbcflag: 새로운 자료가 들어올 경우 정렬을 다시 하게끔...
'''mystring: 줄기-잎-그림이 문자열로 저장
Sub StemandLeaf(rn, check_Outliers As Boolean, sbcflag As Boolean, _
                ByRef mystring, Optional VarName As String = "")

    Static sbcRange As Double
    Dim c As Range: Dim i As Integer
    Dim prange, clsrng As Double
    Dim loutn, uoutn As Integer
    
    If sbcflag = True Then
        Obs = rn.count
        ReDim pdata(1 To Obs)
            
        For Each c In rn
            i = i + 1
            pdata(i) = c.Value
        Next c
        
        procSort1D pdata, 1, Obs
        q1 = Application.Quartile(pdata, 1): q3 = Application.Quartile(pdata, 3)
        IQR = q3 - q1: sbcRange = pdata(Obs) - pdata(1)
    End If
    prange = sbcRange
    If check_Outliers = False Then
        If StemUnit = 0 Then
            clsrng = get_clsrng(prange)
            StemUnit = clsrng
        Else: clsrng = StemUnit
        End If
        out loutn, uoutn
    Else
        If StemUnit = 0 Then
            clsrng = s_get_clsrng(prange)
            StemUnit = clsrng
        Else: clsrng = StemUnit
        End If
        loutn = 0: uoutn = Obs + 1
    End If
    
    mystring = DrawingStemandLeaf(clsrng, loutn, uoutn, VarName)
    
End Sub

Function get_clsrng(prange) As Double

    Dim pclsrng, qclsrng, lclsrng, temp As Double
    Dim L As Double: Dim N As Integer
    
    DiffBug = False
    N = Obs: L = 2 * Sqr(N)
    If prange = 0 Then
        get_clsrng = 1: X = 1
        DiffBug = True: Exit Function
    End If
    If L > 10 Then L = 10
    temp = prange / L
    
'    ''pclsrng: 자료수의 제곱근으로 구하는 계급 구간
'    ''qclsrng: IQR로 구하는 계급 구간
'    n = Obs: L = 2 * Sqr(n)
'
'    If L > 10 Then L = 10
'    If n >= 30 Then
'        pclsrng = prange / L: qclsrng = 2 * IQR / L
'    Else '''n= 1 - 29 인 경우
'        pclsrng = prange / 6: qclsrng = IQR / 2
'    End If
'
'    If n <= 50 Then
'        If pclsrng >= qclsrng Then
'            temp = qclsrng
'        Else: temp = pclsrng
'        End If
'    Else
'        If pclsrng <= qclsrng Then
'            temp = qclsrng
'        Else: temp = pclsrng
'        End If
'    End If
'    ''자료 크기가 50보다 크면 작은 것을, 작으면 큰 것을 clsrng로 한다.

    
    lclsrng = Application.Log10(temp): X = 10 ^ Int(lclsrng)
    If (temp / X) >= 7.5 Then
        get_clsrng = 10 * X
    ElseIf (temp / X) >= 3.5 Then
        get_clsrng = 5 * X
    ElseIf (temp / X) >= 1.5 Then
        get_clsrng = 2 * X
    Else: get_clsrng = X
    End If
        
End Function

Function s_get_clsrng(prange) As Double

    Dim lclsrng, temp As Double
    Dim L As Integer: Dim N As Integer
    
    DiffBug = False
    N = Obs: L = 2 * Sqr(N)
    If prange = 0 Then
        s_get_clsrng = 1: X = 1
        DiffBug = True: Exit Function
    End If
    If L > 10 Then L = 10
    temp = prange / L
    lclsrng = Application.Log10(temp): X = 10 ^ Int(lclsrng)
    If (temp / X) >= 7.5 Then
        s_get_clsrng = 10 * X
    ElseIf (temp / X) >= 3.5 Then
        s_get_clsrng = 5 * X
    ElseIf (temp / X) >= 1.5 Then
        s_get_clsrng = 2 * X
    Else: s_get_clsrng = X
    End If
        
End Function

Sub out(ByRef loutn, ByRef uoutn)

        Dim lowerout, upperout As Double
        Dim i, J, N As Integer
        
        N = Obs
        lowerout = q1 - 3 * IQR: upperout = q3 + 3 * IQR
        loutn = 0: uoutn = N + 1
        
        For i = 1 To N
            If (pdata(i) < lowerout) Then loutn = i
        Next i
        
        For J = N To 1 Step -1
            If (pdata(J) > upperout) Then uoutn = J
        Next J
        
End Sub

Function DrawingStemandLeaf(clsrng, loutn, uoutn, Optional VarName As String = "") As String

    Dim H, i, J, k, Flag, StemStr1, StemStr2 As Integer
    Dim leaf, Stem, y, startpt, endpt As Double
    Dim TitleStr, StemStr, CountStr, r_CountStr As String
    Dim tmpStemStr, r_StemStr1, OutStr As String
    Dim sFlag As Boolean: Dim sSpace As Integer
    Dim a As Integer: Dim tmpData() As Double
    Dim CumCount, tmpCount, MedianFlag, tmpCount2 As Integer
    Dim MedianPos As Single
    
    H = 0: i = 0: J = 0: k = 0: Flag = 0: r_StemStr1 = "1234567"
    y = clsrng / X: MedianPos = (Obs + 1) / 2: MedianFlag = 0
    
    '''출력을 위한 임시 배열 만듬...메모리 걱정?
    ReDim tmpData(1 To Obs)
    For a = 1 To Obs: tmpData(a) = pdata(a): Next a
    
    startpt = Int(tmpData(loutn + 1) / clsrng) * clsrng
    endpt = Int(tmpData(uoutn - 1) / clsrng) * clsrng
    ''startpt: 줄기를 시작하는 곳, endpt: 줄기를 끝내는 곳
    
    If y = 1 Or y = 10 Then
        y = y / 10
    Else: y = 1
    End If
    
    ''이상점을 출력하기 위한 스트링
    OutStr = "이상점: "
    For a = 1 To loutn
        If a <> loutn Then
            OutStr = OutStr & tmpData(a) & ", "
        Else
            OutStr = OutStr & tmpData(a)
        End If
        sFlag = True
    Next a
    
    If sFlag = True Then
        sSpace = 3
    Else: sSpace = 0
    End If
    
    For a = uoutn To Obs
        If a <> Obs Then
            OutStr = OutStr & Space(sSpace) & tmpData(a) & ", "
        Else
            OutStr = OutStr & Space(sSpace) & tmpData(a)
        End If
    Next a
    If OutStr = "이상점: " Then OutStr = ""
    
    For a = loutn + 1 To uoutn - 1
        tmpData(a) = (Int(tmpData(a) / (X * y)) + 10 ^ (-10)) * (X * y)
    Next a
    If VarName <> "" Then
    TitleStr = "줄기-잎 그림(Stem-and-Leaf Plot)" & Chr(10) _
               & "변수명: " & VarName & Chr(10)
    Else: TitleStr = "줄기-잎 그림(Stem-and-Leaf Plot)"
    End If
    TitleStr = TitleStr & Chr(10) & "Stem Unit: " & clsrng & Space(2) _
               & "Leaf Unit: " & X * y & Chr(10) & Chr(10)
    
    '''잎의 개수를 출력하기 위함.
'    r_CountStr = MakingMyString( _
'                 MaxCountofLeaf(tmpdata, loutn, uoutn, startpt, endpt, clsrng))

    r_CountStr = MakingMyString(Int(Application.Log10(Obs) + 1))
    For i = Int(startpt / clsrng) - 1 To Int(endpt / clsrng) - 1
        Stem = startpt + J * clsrng + 10 ^ (-10)
        If (Stem < 0) Then
            Stem = Stem + clsrng: Flag = 1
            StemStr1 = Int(Stem / (X * y * 10))
        Else: Flag = 0
            StemStr1 = Int(Stem / (X * y * 10))
        End If
        RSet r_StemStr1 = str(StemStr1)
        tmpStemStr = r_StemStr1 & Space(3)
        
        For k = loutn + 1 To uoutn - 1
            If tmpData(k) >= startpt + J * clsrng And tmpData(k) < startpt + (J + 1) * clsrng Then
                leaf = Int(tmpData(k) / (X * y) - Int(tmpData(k) / (X * y * 10)) * 10)
                If Stem < 1 And Flag = 1 Then leaf = Abs(9 - leaf)
                StemStr2 = Int(leaf)
                tmpStemStr = tmpStemStr & StemStr2
                If k >= MedianPos - 0.5 And k <= MedianPos + 0.5 Then
                    MedianFlag = 1
                End If
                H = k
            End If
        Next k
                        
        If MedianFlag = 0 Then
            RSet r_CountStr = str(H)
            CountStr = r_CountStr & Space(1)
        ElseIf MedianFlag = 1 Then
            tmpCount2 = H - tmpCount
            RSet r_CountStr = "(" & str(tmpCount2)
            CountStr = r_CountStr & ")"
            MedianFlag = 2
        Else:
            tmpCount2 = Obs - tmpCount
            RSet r_CountStr = str(tmpCount2)
            CountStr = r_CountStr & Space(1)
        End If
        
        tmpStemStr = CountStr & tmpStemStr & Chr(10)
        StemStr = StemStr & tmpStemStr
        tmpCount = H: J = J + 1
        r_CountStr = MakingMyString(Int(Application.Log10(Obs) + 1))
    Next i
    NowStemCount = J
    DrawingStemandLeaf = TitleStr & StemStr & Chr(10) & OutStr
    
End Function

Function CountofStem(clsrng, loutn, uoutn) As Integer

    Dim startpt, endpt As Double:
    Dim i, J As Integer
            
    startpt = Int(pdata(loutn + 1) / clsrng) * clsrng
    endpt = Int(pdata(uoutn - 1) / clsrng) * clsrng
    ''startpt: 줄기를 시작하는 곳, endpt: 줄기를 끝내는 곳
    
    For i = Int(startpt / clsrng) - 1 To Int(endpt / clsrng) - 1
        J = J + 1
    Next i
    NowStemCount = J: CountofStem = J
        
End Function

Function MaxCountofLeaf(tmpData, loutn, uoutn, startpt, endpt, clsrng) As Integer

    Dim tmpCount, i, J, k, count As Integer
    
    J = 0: tmpCount = 0: count = 0
    
    For i = Int(startpt / clsrng) - 1 To Int(endpt / clsrng) - 1
        For k = loutn + 1 To uoutn - 1
            If tmpData(k) >= startpt + J * clsrng And tmpData(k) < startpt + (J + 1) * clsrng Then
                tmpCount = tmpCount + 1
            End If
        Next k
        count = Application.max(count, tmpCount)
        tmpCount = 0: J = J + 1
    Next i
    
    MaxCountofLeaf = Int(Application.Log10(count) + 1)

End Function

Function MakingMyString(count) As String
    Dim i As Integer: Dim str As String
    For i = 1 To count + 4
        str = str & i
    Next i
    MakingMyString = str
End Function

Sub MainStem(rn, xPos, yPos, outputsheet)

    Dim mystring As String
    Dim co As Shape: Dim i As Integer
    
    StemUnit = 0
    StemModule.StemandLeaf rn, True, True, mystring
    Set co = outputsheet.Shapes.AddTextbox(msoTextOrientationHorizontal, xPos, yPos, 200, 200)
    co.TextFrame.Characters.Text = mystring
    co.Line.ForeColor.SchemeColor = 23
    'For i = 1 To Len(mystring)
    '    co.TextFrame.Characters(i, 1).Text = Mid(mystring, i, 1)
    'Next i
    
    With co.TextFrame
        .AutoSize = True
        .Characters.Font.Name = "굴림체"
        .Characters.Font.Size = 9
    End With
    
End Sub
