Attribute VB_Name = "pgFunction"
Sub CountingFreq(s_Input, class, freq, NofClasses, s_MiniUnit)
    
    Dim WidthofItv, mini, maxi, rng As Double
    Dim n As Long
    Dim PossibleDataN, MaxiNofClass, i, j As Integer
    
    n = UBound(s_Input)
    mini = Application.Min(s_Input)
    maxi = Application.Max(s_Input)
    rng = maxi - mini
    PossibleDataN = Int(rng / s_MiniUnit + 1)
    MaxiNofClass = Application.RoundUp(PossibleDataN / NofClasses, 0)
    WidthofItv = s_MiniUnit * MaxiNofClass
    
    ReDim class(0 To NofClasses + 2)
    class(1) = mini - 0.5 * (NofClasses * MaxiNofClass - PossibleDataN) * s_MiniUnit
    
    If (class(1) / s_MiniUnit) * 10 = Int(class(1) / s_MiniUnit) * 10 Then
        class(1) = class(1) - 0.5 * s_MiniUnit
    End If
    
    class(0) = class(1) - WidthofItv
    For i = 2 To NofClasses + 2
        class(i) = class(i - 1) + WidthofItv
    Next i
    
    ReDim freq(0 To NofClasses + 1)
    For i = 0 To NofClasses + 1
        freq(i) = 0
    Next i
    For j = 1 To UBound(s_Input)
        For i = 1 To NofClasses
            If s_Input(j) >= class(i) And s_Input(j) < class(i + 1) Then
                freq(i) = freq(i) + 1
            End If
        Next i
    Next j
    
End Sub

Function FindingMiniUnit(s_Input)
    
    Dim temp, MiniUnit, zeroindi As Double
    Dim zerocount, j As Integer
    
    MiniUnit = 0
    For j = 1 To UBound(s_Input)
        zerocount = 0
        Do
            temp = s_Input(j) * (10 ^ zerocount)
            zeroindi = temp - Fix(temp)
            zerocount = zerocount + 1
        Loop Until (zeroindi = 0)
        
        MiniUnit = Application.Max(MiniUnit, zerocount - 1)
    Next j
    
    FindingMiniUnit = MiniUnit
        
End Function

Function FindingNofClasses(Obs) As Integer

    If 1 <= Obs < 100 Then
        FindingNofClasses = -Int(-Sqr(Obs))
    ElseIf 100 <= Obs <= 400 Then
        FindingNofClasses = Int(Sqr(Obs))
    ElseIf Obs > 400 Then
        FindingNofClasses = 20
    Else: FindingNofClasses = 0
    End If
    
End Function

Sub FrequencyTable(myArray, s0, s1)

    Dim class(), freq(), WidthofClass As Double
    Dim i, tmpcount As Integer
    Dim TempWorksheet As Worksheet
    
    tmpcount = FindingNofClasses(UBound(myArray))
    CountingFreq myArray, class, freq, tmpcount, 10 ^ (-FindingMiniUnit(myArray))
    
    Set TempWorksheet = Worksheets("tempclt")
    
    WidthofClass = class(1) - class(0)
    For i = 0 To UBound(freq)
        TempWorksheet.Cells(i + 2, 1).Value = (class(i) + class(i + 1)) / 2
        TempWorksheet.Cells(i + 2, 2).Value = freq(i) / UBound(myArray) / WidthofClass
    Next i
    
    Set s1 = Range(TempWorksheet.Cells(2, 2), TempWorksheet.Cells(UBound(freq) + 2, 2))
    Set s0 = Range(TempWorksheet.Cells(2, 1), TempWorksheet.Cells(UBound(freq) + 2, 1))

End Sub

Sub UniformClt(n, iter, smean, Optional Tstat As Boolean = False)
Attribute UniformClt.VB_ProcData.VB_Invoke_Func = " \n14"

    Dim sample() As Double
    Dim i, j As Integer
     
    ReDim sample(1 To n): ReDim smean(1 To iter)
    
    For j = 1 To iter
       For i = 1 To n
           sample(i) = Rnd * 10
       Next i
       If Tstat = False Then
             smean(j) = (Application.Average(sample) - 5) / Sqr(25 / 3 / n)
       Else: smean(j) = (Application.Average(sample) - 5) / Sqr(Application.Var(sample) / n)
       End If
    Next j
    
End Sub

Sub ExponentialClt(n, iter, smean, Optional Tstat As Boolean = False)

    Dim sample() As Double
    Dim i, j As Integer
     
    ReDim sample(1 To n): ReDim smean(1 To iter)
    
    For j = 1 To iter
       For i = 1 To n
           sample(i) = -Application.Ln(Rnd)
       Next i
       If Tstat = False Then
             smean(j) = (Application.Average(sample) - 1) / Sqr(1 / n)
       Else: smean(j) = (Application.Average(sample) - 1) / Sqr(Application.Var(sample) / n)
       End If
    Next j
    
End Sub

Sub NormalClt(n, iter, smean, Optional Tstat As Boolean = False)

    Dim sample() As Double
    Dim i, j As Integer
     
    ReDim sample(1 To n): ReDim smean(1 To iter)
    
    For j = 1 To iter
       For i = 1 To n
           sample(i) = Application.NormInv(Rnd, 0, 1)
       Next i
       If Tstat = False Then
             smean(j) = (Application.Average(sample) - 0) / Sqr(1 / n)
       Else: smean(j) = (Application.Average(sample) - 0) / Sqr(Application.Var(sample) / n)
       End If
    Next j
    
End Sub

Function Tpdf(t, df) As Double
Attribute Tpdf.VB_ProcData.VB_Invoke_Func = " \n14"
   
    Dim temp0, temp As Double
    
    temp0 = (df + 1) / 2
    temp = Application.Power(1 + t * t / df, -temp0)
    temp = 1 / Sqr(3.14159265 * df) * Exp(Application.GammaLn(temp0)) / Exp(Application.GammaLn(temp0 - 0.5)) * temp
    
    Tpdf = temp
    
End Function
      
Function Chipdf(x, df) As Double
Attribute Chipdf.VB_ProcData.VB_Invoke_Func = " \n14"
   
    Dim temp As Double
    
    If x = 0 Then
         Chipdf = 0
    Else: temp = 0.5 * df - 1
         temp = (x / 2) ^ temp * Exp(-x / 2)
         Chipdf = temp / 2 / Exp(Application.GammaLn(df / 2))
    End If
    
End Function
       
Function Fpdf(f, df1, df2) As Double
Attribute Fpdf.VB_ProcData.VB_Invoke_Func = " \n14"
   
    Dim temp0, temp1 As Double
    
    If f = 0 Then
        Fpdf = 0
    Else
        temp0 = Exp(Application.GammaLn((df1 + df2) / 2)) / Exp(Application.GammaLn(df1 / 2)) / Exp(Application.GammaLn(df2 / 2))
        temp1 = Application.Power(df1, df1 / 2) * Application.Power(df2, df2 / 2) * Application.Power(f, df1 / 2 - 1) * Application.Power(df2 + df1 * f, -(df1 + df2) / 2)
        Fpdf = temp0 * temp1
    End If
    
End Function

Function Normal(x, mu, sigma) As Double

    Normal = 1 / Sqr(3.14159265 * 2) / sigma * Exp(-0.5 * ((x - mu) / sigma) ^ 2)
    
End Function

Sub HistandScatter(class, L As Double, U As Double)
    
    Dim Flag As Boolean: Dim wid As Double
    Dim tmpclass() As Double: Dim i As Integer
    
    Flag = False: wid = class(1) - class(0)
    Do
        If class(UBound(class)) + wid / 2 < U Then
            ReDim Preserve class(LBound(class) To UBound(class) + 1)
            class(UBound(class)) = class(UBound(class) - 1) + wid
        ElseIf class(LBound(class)) - wid / 2 > L Then
            ReDim tmpclass(LBound(class) To UBound(class))
            For i = LBound(class) To UBound(class)
                tmpclass(i) = class(i)
            Next i
            ReDim class(LBound(class) - 1 To UBound(class))
            class(LBound(class)) = tmpclass(LBound(class) + 1) - wid
            For i = LBound(class) + 1 To UBound(class)
                class(i) = tmpclass(i)
            Next i
        Else: Flag = True
        End If
    Loop While (Flag = False)
    
End Sub
