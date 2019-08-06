Attribute VB_Name = "TModuleGraph"
Option Base 1
Public Sub Ciplot()

        Dim dataArray(), rstArray()
    Dim theta0 As Single, CI As Single, Hyp As Integer
    Dim mySheet As Worksheet
    Dim titleTmp() As String, title As String
    
    
    Worksheets(DataSheet).Activate
    dataArray = Worksheets(DataSheet).Range(Cells(2, k1), Cells(n + 1, k1)).Value
    
    Dim i As Long, nCol As Long, nRow As Long
    Dim nGrp As Integer, nPage As Integer
    Dim CIcht As Chart
    Dim strName As String
    Dim rngData As Range, rngName As Range, rngTitle As Range, arrMean() As Double
    Dim PooledSD As Double, dblSE As Double, dblCL As Double, arrDblSE() As Double
    Dim arrLow() As Double
    Dim rngFirst As Range
    Dim cc1(10) As Double, cc2 As Double
    Dim arrtltB() As String

    On Error GoTo ErrEnd
    Application.ScreenUpdating = False
    
'   Read data
  '  Set rngData = Range(Me.RefEdit1)
    Set rngData = Range(Cells(2, k1), Cells(n + 1, k1))
    
'   Count the # of rows and columns
   ' nGrp = rngData.Columns.Count
    nGrp = 1
    nRow = rngData.Rows.count
    
'   Resize data
    ReDim arrtltB(1)
        arrtltB(1) = frameInterval.ListBox2.List(0)
    
'   Err Check
 '   If nGrp < 2 Then
 '      MsgBox "2개 이상의 그룹을 선택해야 합니다."
 '       Me.RefEdit1 = ""
 '       Me.RefEdit1.SetFocus
 '       Application.ScreenUpdating = True
 '       Exit Sub
 '   End If
    
    ReDim arrMean(nGrp)
    ReDim arrDblSE(nGrp)
    ReDim arrLow(nGrp)
    
'   Standard Error Calculation
    PooledSD = 0
    With WorksheetFunction
    For i = 1 To nGrp
        arrMean(i) = .Average(rngData.Columns(i))
        If frameInterval.OptionButton3 Then
            If .count(rngData.Columns(i)) > 1 Then
                arrDblSE(i) = .StDev(rngData.Columns(i)) / Sqr(.count(rngData.Columns(i)))
            Else
                arrDblSE(i) = 0
            End If
        Else
            PooledSD = PooledSD + .SumSq(rngData.Columns(i)) - .count(rngData.Columns(i)) * arrMean(i) ^ 2
        End If
    Next i
    End With
    
'   Confidence Interval Calcualtion
    With WorksheetFunction
        dblCL = CDbl(frameInterval.TextBox2) * 0.01
        If frameInterval.OptionButton3 Then
            For i = 1 To nGrp
                arrDblSE(i) = .TInv(1 - dblCL, rngData.Rows.count - 1) * arrDblSE(i)
                arrLow(i) = arrMean(i) - arrDblSE(i)
            Next i
        Else
            PooledSD = Sqr(PooledSD / (.count(rngData) - nGrp))
            For i = 1 To nGrp
                arrDblSE(i) = .TInv(1 - dblCL, rngData.Rows.count - 1) * PooledSD / Sqr(rngData.Rows.count)
                arrLow(i) = arrMean(i) - arrDblSE(i)
            Next i
        End If
    End With
        
'   Add a Sheet
    For i = 1 To Sheets.count
        If Sheets(i).Name = "_통계분석결과_" Then
            GoTo 31
        Else
            GoTo 32
        End If
32: Next i
    Worksheets.Add Before:=Worksheets(1)
    ActiveSheet.Name = "_통계분석결과_"
    ActiveWindow.DisplayGridlines = False
    Cells(1, 1) = 1

31: Sheets("_통계분석결과_").Activate
    Application.ScreenUpdating = False

'   Current worksheet's name
    strName = ActiveSheet.Name

'''''위치 수정해야함
'   Chart Location
    Set rngFirst = Cells(Cells(1, 1) - 1, 1)

'   Confidence Interval Chart
    Set CIcht = Charts.Add
    Set CIcht = CIcht.Location(where:=xlLocationAsObject, Name:=strName)

    With CIcht
        .ChartType = xlLineMarkers
        .HasLegend = False
       ' .HasTitle = True
    
      '  With .ChartTitle
      '      .Characters.Text = "Confidence Interval Plot" & vbCrLf & "for mean"
      '      .Font.Size = 12
      '      .Font.Bold = True
      '      .AutoScaleFont = False
      '  End With
       For i = 1 To nRow
        .SeriesCollection.NewSeries
        Next i
        With .SeriesCollection(1)
            .HasDataLabels = True
            .DataLabels.NumberFormat = "##.##"
            .XValues = arrtltB
            .Values = arrMean
            .Border.LineStyle = xlContinuous 'xlNone
            .Border.Weight = xlThin
            .MarkerStyle = xlMarkerStyleStar
            .ErrorBar Direction:=xlY, Include:=xlErrorBarIncludeBoth, Type:=xlErrorBarTypeCustom, _
            Amount:=arrDblSE, MinusValues:=arrDblSE
            .MarkerSize = 7
        End With

        With .Axes(xlValue, xlPrimary)
            .HasTitle = False
            .HasMajorGridlines = False
            .HasMinorGridlines = False
            .MinimumScaleIsAuto = True

'   Ajusting Y axis value
            .MinimumScaleIsAuto = False
            cc1(1) = .MajorUnit
            For i = 1 To 10
                If WorksheetFunction.Min(arrLow) > .MinimumScale + .MajorUnit * 2 Then
                    .MinimumScale = .MinimumScale + .MajorUnit
                End If
            Next i
        End With

        With .Parent
            .Top = rngFirst.Offset(3, 1).Top
            .Left = rngFirst.Offset(3, 1).Left
            .Width = 240
            .Height = 180
        End With

    End With

ErrEnd:

   'Page number reset
      nPage = 18
   ' rngFirst.Offset = "Created at " & Now()
    Application.Goto rngFirst, Scroll:=True
    Range("A1") = Range("A1") + nPage
    
   Application.ScreenUpdating = True

    'Unload Me

End Sub
