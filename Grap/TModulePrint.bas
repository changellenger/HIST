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
         .Font.Name = "±¼¸²"
         .Font.Size = 9
         .HorizontalAlignment = xlLeft            'xlright : HIST
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
                                            'Flag¿¡ º¯È­¾øÀ½. ¼±¸¸ ±×·ÁÁÜ
                                            'RstSheet¿¡ (Flag,2)ºÎÅÍ (row,col)¸¸Å­
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
    
    pt.Resize(row, col).HorizontalAlignment = xlLeft                    'xlright : HIST
    
    
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
        .Line.Weight = 1
        .Line.Visible = msoTrue
       '.Shadow.Type = msoShadow1
    End With
   
    With title.TextFrame.Characters
        .Text = contents
        .Font.Name = "±¼¸²"
        .Font.FontStyle = "±½°Ô"
        .Font.Size = 14
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
    
    Set title = mySheet.Shapes.AddShape(msoShapeRectangle, 60.75, yp, 140, 25#)
    With title
        .Fill.ForeColor.SchemeColor = 9
        .Solid
        .Line.Visible = msoTrue
        .Shadow.Type = msoShadow17
    End With
   
    With title.TextFrame.Characters
        .Text = contents
        .Font.Name = "±¼¸²"
        .Font.FontStyle = "±½°Ô"
        .Font.Size = 11
        .Font.ColorIndex = 41
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
    
    Set title = mySheet.Shapes.AddShape(msoShapeRectangle, 100, yp, 200, 22#)
    With title
        .Fill.ForeColor.SchemeColor = 1
        .Solid
        .Line.Visible = msoTrue
        .Shadow.Type = msoShadow17
    End With
   
    With title.TextFrame.Characters
        .Text = contents
        .Font.Name = "±¼¸²"
        .Font.FontStyle = "±½°Ô"
        .Font.Size = 11
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
        mySheet.Cells(Flag - 2, 2).Font.Size = 10
        mySheet.Cells(Flag - 2, 2).Font.Bold = True
        mySheet.Cells(Flag - 2, 2).HorizontalAlignment = xlGeneral
    End If
    
    Flag = mySheet.Cells(1, 1).Value
    row = UBound(rstArray)
    col = UBound(rstArray, 2)
    
    TABLE row, col
    mySheet.Range(Cells(Flag, 2), Cells(Flag, col + 1)).Select             ''º¯¼ö¸íÀÌ 1-2 ÀÎ°æ¿ì 1¿ù2ÀÏÀÌ µÇÁö ¾Ê°Ô ÇÏ±â À§ÇÏ¿©
    Selection.NumberFormatLocal = "@"
    mySheet.Range(Cells(Flag, 2), Cells(Flag + row - 1, col + 1)).Value = rstArray
    mySheet.Cells(1, 1) = Flag + row
    
End Sub

Sub TitleN(contents As String)
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
        .Font.Name = "µ¸¿ò"
        .Font.FontStyle = "±½°Ô"
        .Font.Size = 14
        .Font.ColorIndex = 2
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
    
    mySheet.Cells(1, 1) = Flag + 4
    
    '''
    If tmpSign = 1 Then
        mySheet.Cells(1, 1) = "$A$" & mySheet.Cells(1, 1).Value
    End If
    
    
End Sub
