Attribute VB_Name = "TModulePrint"
Sub makeOutputSheet(SheetName)

    Dim s As Worksheet
    
    For Each s In ActiveWorkbook.Sheets
        If s.Name = SheetName Then Exit Sub
    Next s
    
    Worksheets.Add.Name = SheetName
    
    With ActiveWindow
        .DisplayGridlines = False
'        .DisplayHeadings = False
    End With
    
    With ActiveWindow.Application.Cells
         .Font.Name = "굴림"
         .Font.size = 9
         .HorizontalAlignment = xlLeft
    End With

    With Worksheets(SheetName).Range("a1")
        .Value = 2
        .Font.ColorIndex = 2
    End With
    Worksheets(SheetName).Rows(1).Hidden = True
    Worksheets(SheetName).Activate
    Cells.Select
    Selection.RowHeight = 13.5
    
    
End Sub

Sub TABLE(row, col)
                                            'Flag에 변화없음. 선만 그려줌
                                            'RstSheet에 (Flag,2)부터 (row,col)만큼
    Dim mySheet As Worksheet
    
    Set mySheet = Worksheets(RstSheet)
    Flag = mySheet.Cells(1, 1).Value
    Set pt = mySheet.Cells(Flag, 2)
    ''
    ''
    With pt.Resize(, col).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    pt.Resize(row, col).HorizontalAlignment = xlLeft
    
    
    With pt.Resize(, col).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    ''
    ''
    With pt.Cells(row, 1).Resize(, col).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    ''
    ''
End Sub



Sub Title1(contents As String)
    Dim Flag As Long
    Dim mySheet As Worksheet
    Dim tmpSign
    
    '''
    tmpSign = 0
    Set mySheet = Worksheets(RstSheet)
    If Left(mySheet.Range("a1"), 1) = "$" Then
        mySheet.Cells(1, 1) = Right(mySheet.Cells(1, 1).Value, Len(mySheet.Cells(1, 1).Value) - 3)
        tmpSign = 1
    End If
    
    Flag = mySheet.Cells(1, 1).Value
    yp = mySheet.Cells(Flag + 2, 1).Top
    
    On Error Resume Next
    
    Set title = mySheet.Shapes.AddShape(msoShapeRectangle, 3.75, yp + 2.5, 400, 25#)
    With title
        .Fill.ForeColor.SchemeColor = 57
        .Line.DashStyle = msoLineSolid
        .Line.Style = msoLineSingle
        .Line.ForeColor.SchemeColor = 8
        .Line.Weight = 1
        .Line.Visible = msoTrue
        '.Shadow.Type = msoShadow1
    End With
   
    With title.TextFrame.Characters
        .Text = contents
        .Font.Name = "돋움"
        .Font.FontStyle = "굵게"
        .Font.size = 14
        .Font.ColorIndex = 2
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
    
    mySheet.Cells(1, 1) = Flag + 4
    
    '''
    If tmpSign = 1 Then
        mySheet.Cells(1, 1) = "$A$" & mySheet.Cells(1, 1).Value
    End If
    
    
End Sub

Sub Title2(contents As String)
    Dim Flag As Long
    Dim mySheet As Worksheet
    
    Set mySheet = Worksheets(RstSheet)
    Flag = mySheet.Cells(1, 1).Value
    yp = mySheet.Cells(Flag + 2, 1).Top
    
    On Error Resume Next
    
    Set title = mySheet.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 250, 25#)
    With title
        .Fill.ForeColor.SchemeColor = 55
        .Solid
        .Line.Visible = msoTrue
        .Shadow.Type = msoShadow17
    End With
   
    With title.TextFrame.Characters
        .Text = contents
        .Font.Name = "굴림"
        .Font.FontStyle = "굵게"
        .Font.size = 11
        .Font.ColorIndex = 2
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
    
    mySheet.Cells(1, 1) = Flag + 4
End Sub

Sub Title3(contents As String)

    Dim Flag As Long
    Dim mySheet As Worksheet
    
    Set mySheet = Worksheets(RstSheet)
    Flag = mySheet.Cells(1, 1).Value
    yp = mySheet.Cells(Flag + 2, 1).Top
    
    On Error Resume Next
    
    Set title = mySheet.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 250, 22#)             '길이늘림
    With title
        .Fill.ForeColor.SchemeColor = 1
        .Solid
          .Line.ForeColor.SchemeColor = 8
           .Line.Weight = 1
      '  .Line.Visible = msoTrue
     '   .Shadow.Type = msoShadow17
    End With
   
    With title.TextFrame.Characters
        .Text = contents
        .Font.Name = "맑은 고딕"
       ' .Font.FontStyle = "굵게"
        .Font.size = 11
        .Font.ColorIndex = xlAutomatic
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
    
    mySheet.Cells(1, 1) = Flag + 4
    
End Sub

Sub printRst(title, rstArray)
    
    Dim mySheet As Worksheet
     
    Set mySheet = Worksheets(RstSheet)
    mySheet.Activate ''''''''''''
    
    mySheet.Cells(1, 1) = mySheet.Cells(1, 1) + 1
    
    If title <> "" Then
        mySheet.Cells(1, 1) = mySheet.Cells(1, 1) + 2
        Flag = mySheet.Cells(1, 1).Value
        mySheet.Cells(Flag - 2, 2) = title
        mySheet.Cells(Flag - 2, 2).Font.size = 10
        mySheet.Cells(Flag - 2, 2).Font.Bold = True
        mySheet.Cells(Flag - 2, 2).HorizontalAlignment = xlGeneral
    End If
    
    Flag = mySheet.Cells(1, 1).Value
    row = UBound(rstArray)
    col = UBound(rstArray, 2)
    
    TABLE row, col
    mySheet.Range(Cells(Flag, 2), Cells(Flag, col + 1)).Select             ''변수명이 1-2 인경우 1월2일이 되지 않게 하기 위하여
    Selection.NumberFormatLocal = "@"
    mySheet.Range(Cells(Flag, 2), Cells(Flag + row - 1, col + 1)).Value = rstArray
    mySheet.Cells(1, 1) = Flag + row
    
End Sub

Sub TitleD(contents As String)                  '타이틀 네개 만들기

    Dim Flag As Long
    Dim mySheet As Worksheet
    
    Set mySheet = Worksheets(RstSheet)
    Flag = mySheet.Cells(1, 1).Value
    yp = mySheet.Cells(Flag + 2, 1).Top
    
    On Error Resume Next
    
    Set title = mySheet.Shapes.AddShape(msoShapeRectangle, 25, yp + 2.5, 280, 22#)           '길이늘림 1
    With title
        .Fill.ForeColor.SchemeColor = 1
        .Solid
          .Line.ForeColor.SchemeColor = 8
           .Line.Weight = 1
      '  .Line.Visible = msoTrue
     '   .Shadow.Type = msoShadow17
    End With
    With title.TextFrame.Characters
        .Text = contents
        .Font.Name = "맑은 고딕"
       ' .Font.FontStyle = "굵게"
        .Font.size = 10
        .Font.ColorIndex = xlAutomatic
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
    
    
    Set titlea = mySheet.Shapes.AddShape(msoShapeRectangle, 350, yp + 2.5, 280, 22#)           '길이늘림 2
    With titlea
        .Fill.ForeColor.SchemeColor = 1
        .Solid
          .Line.ForeColor.SchemeColor = 8
           .Line.Weight = 1
      '  .Line.Visible = msoTrue
     '   .Shadow.Type = msoShadow17
    End With
    With titlea.TextFrame.Characters
        .Text = "결론 및 요약"
        .Font.Name = "맑은 고딕"
       ' .Font.FontStyle = "굵게"
        .Font.size = 10
        .Font.ColorIndex = xlAutomatic
    End With
    titlea.TextFrame.HorizontalAlignment = xlCenter
    
         
    Set titleb = mySheet.Shapes.AddShape(msoShapeRectangle, 25, yp + 180, 280, 22#)           '길이늘림 3
    With titleb
        .Fill.ForeColor.SchemeColor = 1
        .Solid
          .Line.ForeColor.SchemeColor = 8
           .Line.Weight = 1
      '  .Line.Visible = msoTrue
     '   .Shadow.Type = msoShadow17
    End With
    With titleb.TextFrame.Characters
        .Text = "데이터의 분포"
        .Font.Name = "맑은 고딕"
       ' .Font.FontStyle = "굵게"
        .Font.size = 10
        .Font.ColorIndex = xlAutomatic
    End With
    titleb.TextFrame.HorizontalAlignment = xlCenter
    
    '
    '----------------- 오른쪽아래칸
    '
          
    Set titlec = mySheet.Shapes.AddShape(msoShapeRectangle, 350, yp + 180, 280, 22#)           '길이늘림 4
    With titlec
        .Fill.ForeColor.SchemeColor = 1
        .Solid
        .Line.ForeColor.SchemeColor = 8
        .Line.Weight = 1
       ' .Line.Visible = msoTrue
       ' .Shadow.Type = msoShadow17
    End With
    With titlec.TextFrame.Characters
        .Text = "데이터의 평균"
        .Font.Name = "맑은 고딕"
      '  .Font.FontStyle = "굵게"
        .Font.size = 10
        .Font.ColorIndex = xlAutomatic
    End With
    titlec.TextFrame.HorizontalAlignment = xlCenter
    '-------------------------
    
    
    
    mySheet.Cells(1, 1) = Flag + 4
    
End Sub


Sub TitleN(contents As String)          ' 큰 타이틀 한개
    Dim Flag As Long
    Dim mySheet As Worksheet
    Dim tmpSign
    
    '''
    tmpSign = 0
    Set mySheet = Worksheets(RstSheet)
    If Left(mySheet.Range("a1"), 1) = "$" Then
        mySheet.Cells(1, 1) = Right(mySheet.Cells(1, 1).Value, Len(mySheet.Cells(1, 1).Value) - 3)
        tmpSign = 1
    End If
    
    Flag = mySheet.Cells(1, 1).Value
    yp = mySheet.Cells(Flag + 2, 1).Top
    
    On Error Resume Next
    
    Set title = mySheet.Shapes.AddShape(msoShapeRectangle, 3.75, yp + 2.5, 650, 30#)
    With title
        .Fill.ForeColor.SchemeColor = 57
        .Line.DashStyle = msoLineSolid
        .Line.Style = msoLineSingle
        .Line.ForeColor.SchemeColor = 8
        .Line.Weight = 1
        .Line.Visible = msoTrue
        '.Shadow.Type = msoShadow1
    End With
   
    With title.TextFrame.Characters
        .Text = contents
        .Font.Name = "돋움"
        .Font.FontStyle = "굵게"
        .Font.size = 14
        .Font.ColorIndex = 2
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
    
    mySheet.Cells(1, 1) = Flag + 4
    
    '''
    If tmpSign = 1 Then
        mySheet.Cells(1, 1) = "$A$" & mySheet.Cells(1, 1).Value
    End If
    
    
End Sub

Sub printRstNum(title, rstArray, num)
    
    Dim mySheet As Worksheet
     
    Set mySheet = Worksheets(RstSheet)
    mySheet.Activate ''''''''''''
    
    
    If num = 1 Or num = 3 Then
    mySheet.Cells(1, 1) = mySheet.Cells(1, 1) + 1
    End If
    

   
     If title <> "" And num = 1 Then
         mySheet.Cells(1, 1) = mySheet.Cells(1, 1)
         Flag = mySheet.Cells(20, 1).Value
         mySheet.Cells(Flag - 2, 2) = title
         mySheet.Cells(Flag - 2, 2).Font.size = 10
         mySheet.Cells(Flag - 2, 2).Font.Bold = True
         mySheet.Cells(Flag - 2, 2).HorizontalAlignment = xlGeneral
    End If
    
    

    Flag = mySheet.Cells(1, 1).Value
    row = UBound(rstArray)
    col = UBound(rstArray, 2)


   
   If num = 1 Then
    TABLENum row, col, num
    End If
    
    
    If num = 1 Or num = 3 Then
    mySheet.Range(Cells(Flag, 2), Cells(Flag, col + 1)).Select             ''변수명이 1-2 인경우 1월2일이 되지 않게 하기 위하여
    Selection.NumberFormatLocal = "@"
    mySheet.Range(Cells(Flag, 2), Cells(Flag + row - 1, col + 1)).Value = rstArray
    ElseIf num = 2 Or 4 Then
    mySheet.Range(Cells(Flag, 8), Cells(Flag, col + 7)).Select             ''변수명이 1-2 인경우 1월2일이 되지 않게 하기 위하여
    Selection.NumberFormatLocal = "@"
    mySheet.Range(Cells(Flag, 8), Cells(Flag + row - 1, col + 7)).Value = rstArray
    End If
    
    
   If num = 1 Or 2 Or num = 4 Then
   mySheet.Cells(1, 1) = Flag + row
    End If
    
    
End Sub


Sub TABLENum(row, col, num)
                                            'Flag에 변화없음. 선만 그려줌
                                            'RstSheet에 (Flag,2)부터 (row,col)만큼
    Dim mySheet As Worksheet
    
    Set mySheet = Worksheets(RstSheet)
    Flag = mySheet.Cells(1, 1).Value
    
     If num = 1 Then
    Set pt = mySheet.Cells(Flag, 2)
    ElseIf num = 2 Then
    Set pt = mySheet.Cells(Flag, 8)
    End If
    ''
    ''
    With pt.Resize(, col).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    pt.Resize(row, col).HorizontalAlignment = xlLeft
    
    
    With pt.Resize(, col).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    ''
    ''

    ''
    ''
    With pt.Cells(row - 2, 1).Resize(, col).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    ''
    ''
    With pt.Cells(row - 1, 1).Resize(, col).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
 
    ''
    ''

End Sub


Sub printRstOne(title, rstArray)
    
    Dim mySheet As Worksheet
     
    Set mySheet = Worksheets(RstSheet)
    mySheet.Activate ''''''''''''
    
    mySheet.Cells(1, 1) = mySheet.Cells(1, 1) + 1
    
    If title <> "" Then
        mySheet.Cells(1, 1) = mySheet.Cells(1, 1) + 2
        Flag = mySheet.Cells(1, 1).Value
        mySheet.Cells(Flag - 2, 2) = title
        mySheet.Cells(Flag - 2, 2).Font.size = 10
        mySheet.Cells(Flag - 2, 2).Font.Bold = True
        mySheet.Cells(Flag - 2, 2).HorizontalAlignment = xlGeneral
    End If
    
    Flag = mySheet.Cells(1, 1).Value
    row = UBound(rstArray)
    col = UBound(rstArray, 2)
    
    TABLEOne row, col
    mySheet.Range(Cells(Flag, 2), Cells(Flag, col + 1)).Select             ''변수명이 1-2 인경우 1월2일이 되지 않게 하기 위하여
    Selection.NumberFormatLocal = "@"
    mySheet.Range(Cells(Flag, 2), Cells(Flag + row - 1, col + 1)).Value = rstArray
    mySheet.Cells(1, 1) = Flag + row
    
End Sub
Sub printRstResult(title, rstArray)

    Dim mySheet As Worksheet
     
    Set mySheet = Worksheets(RstSheet)
    mySheet.Activate ''''''''''''
    
    mySheet.Cells(1, 1) = mySheet.Cells(1, 1) + 1

    
    Flag = mySheet.Cells(1, 1).Value
    row = UBound(rstArray)
    col = UBound(rstArray, 2)
    
   ' TABLEOne row, col
    mySheet.Range(Cells(Flag, 2), Cells(Flag, col + 1)).Select             ''변수명이 1-2 인경우 1월2일이 되지 않게 하기 위하여
    Selection.NumberFormatLocal = "@"
    mySheet.Range(Cells(Flag, 2), Cells(Flag + row - 1, col + 1)).Value = rstArray
    mySheet.Range(Cells(Flag, 2), Cells(Flag + row - 1, col + 1)).HorizontalAlignment = xlLeft
    mySheet.Cells(1, 1) = Flag + row
    
End Sub


Sub TABLEOne(row, col)
                                            'Flag에 변화없음. 선만 그려줌
                                            'RstSheet에 (Flag,2)부터 (row,col)만큼
    Dim mySheet As Worksheet
    
    Set mySheet = Worksheets(RstSheet)
    Flag = mySheet.Cells(1, 1).Value
    Set pt = mySheet.Cells(Flag, 2)
    ''
    ''
    With pt.Resize(, col).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    pt.Resize(row, col).HorizontalAlignment = xlLeft
    
    
    With pt.Resize(, col).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    ''
    ''
    With pt.Cells(1, 1).Resize(, col).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    ''
    ''
    
End Sub
