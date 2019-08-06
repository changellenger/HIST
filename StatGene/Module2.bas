Attribute VB_Name = "Module2"
Sub FTest2(choice, op)

    Dim rngData1 As Range, rngData2 As Range
    Dim F As Double, Mean1 As Double, Theta As Double, confid As Double
    Dim Mean2 As Double, SD1 As Double, SD2 As Double
    Dim index As Long, nData1 As Long, nData2 As Long
    Dim n1 As Long, n2 As Long
    Dim DF1 As Long, DF2 As Long
    Dim msg As String, H1 As String
    Dim v1 As Double, v2 As Double, Nu As Double, De As Double
    Dim X1bar As Double, X2bar As Double
    Dim s1 As Double, s2 As Double, p As Double
    Dim i As Long, nPage As Long, rngFirst As Range
    Dim strTitle(2) As String, strName As String

    On Error GoTo ErrEnd

If op = 1 Then                                              '표준입력
    
    ReDim dataArray1(n2(1))
    ReDim dataArray2(n2(2))

    With Worksheets(DataSheet)
        For j = 1 To n2(1)
            dataArray1(j) = .Cells(j + 1, k2(1)).Value
        Next j
        For j = 1 To n2(2)
            dataArray2(j) = .Cells(j + 1, k2(2)).Value
        Next j
    End With
    
End If



'   The number of data
    nData1 = rngData1.Count
    nData2 = rngData2.Count

'   Statistics
    With Application.WorksheetFunction
        X1bar = .Average(rngData1)
        X2bar = .Average(rngData2)
        v1 = .Var(rngData1)
        v2 = .Var(rngData2)
        s1 = .StDev(rngData1)
        s2 = .StDev(rngData2)
        If v1 > v2 Then
            Nu = v1
            De = v2
            DF1 = nData1 - 1
            DF2 = nData2 - 1
        Else
            Nu = v2
            De = v1
            DF1 = nData2 - 1
            DF2 = nData1 - 1
        End If
        F = Nu / De
        p = .FDist(F, DF1, DF2) * 2
    End With

'   Insert Worshsheet
    For i = 1 To Sheets.Count
        If Sheets(i).Name = "_stat_" Then
            GoTo 31
        Else
            GoTo 32
        End If
     Next i
    Worksheets.Add Before:=Worksheets(1)
    ActiveSheet.Name = "_stat_"
    ActiveWindow.DisplayGridlines = False
    
    Cells(1, 1) = 1

    Sheets("_stat_").Activate
    Application.ScreenUpdating = False
    
'   Print Location
    nPage = Cells(1, 1)
    Set rngFirst = Cells(Cells(1, 1) + 2, 1)

'   Display Option
    strName = ActiveSheet.Name
    Application.ScreenUpdating = False
    ActiveWindow.DisplayGridlines = False

'   Title
    With rngFirst.Offset(1, 1)
        .Value = "F-Test"
        .Font.Name = "Book Antiqua"
        .Font.Bold = True
        .Font.size = 12
    End With

'   Writing statistics
    With rngFirst
        .Offset(3, 2) = strTitle(1)
        .Offset(3, 3) = strTitle(2)
        
        .Offset(3, 1) = "항목"
        .Offset(4, 1) = "자료수"
        .Offset(4, 2) = nData1
        .Offset(4, 3) = nData2

        .Offset(5, 1) = "평균"
        .Offset(5, 2) = X1bar
        .Offset(5, 3) = X2bar
        .Offset(6, 1) = "표준편차"
        .Offset(6, 2) = s1
        .Offset(6, 3) = s2
        .Offset(7, 1) = "F"
        .Offset(7, 2) = F
        .Offset(7, 2).NumberFormatLocal = "0.00"
        .Offset(8, 1) = "P값"
        .Offset(8, 2) = p
        .Offset(8, 2).NumberFormatLocal = "0.000"
        .Offset(10, 1) = "귀무가설(H0) : 두 모집단의 산포는 같다."
        .Offset(11, 1) = "대립가설(H1) : 두 모집단의 산포는 다르다."
    End With

'   Make a Table
    Call MakeTable(Range(rngFirst.Offset(3, 1), rngFirst.Offset(8, 3)))
    
    For i = 7 To 8
        With Range(rngFirst.Offset(i, 2), rngFirst.Offset(i, 3))
            .Merge
            .HorizontalAlignment = xlCenter
        End With
    Next i

'   Page number reset
    rngFirst = "Created at " & Now()
    Application.Goto rngFirst, Scroll:=True
    Range("A1") = Range("A1") + 15

ErrEnd:

    Application.ScreenUpdating = True
    Unload Me

End Sub
