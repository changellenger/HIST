Attribute VB_Name = "ModulePrint"
Sub makeOutputSheet(sheetName)

    Dim s As Worksheet
    
    For Each s In ActiveWorkbook.Sheets
        If s.Name = sheetName Then Exit Sub
    Next s
    
    Worksheets.Add.Name = sheetName
    
    With ActiveWindow
        .DisplayGridlines = False
'        .DisplayHeadings = False
    End With
    
    With ActiveWindow.Application.Cells
         .Font.Name = "굴림"
         .Font.Size = 9
         .HorizontalAlignment = xlRight
    End With

    With Worksheets(sheetName).Range("a1")
        .Value = 2
        .Font.ColorIndex = 2
    End With
    Worksheets(sheetName).Rows(1).Hidden = True
    Worksheets(sheetName).Activate
    Cells.Select
    Selection.RowHeight = 13.5
    
End Sub


Sub TABLE(row, col)
                                            'Flag에 변화없음. 선만 그려줌
                                            'RstSheet에 (Flag,2)부터 (row,col)만큼
    Dim mySheet As Worksheet
    
    Set mySheet = Worksheets(rstSheet)
    Flag = mySheet.Cells(1, 1).Value
    Set pt = mySheet.Cells(Flag, 2)
    ''
    ''
    With pt.Resize(, col).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    pt.Resize(row, col).HorizontalAlignment = xlRight
    
    
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


Sub TABLE1(row, col, fore, aft)
                                            'Flag에 변화없음. 선만 그려줌
                                            '(Flag,2)부터 (row,col)만큼 total<>0이면 한줄더 선을 그림
    Dim mySheet As Worksheet
    
    Set mySheet = Worksheets(rstSheet)
    Flag = mySheet.Cells(1, 1).Value
    Set pt = mySheet.Cells(Flag, 2)
    ''
    ''
    With pt.Resize(, col).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    pt.Resize(row + total, col).HorizontalAlignment = xlRight
    
    
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
    If total > 0 Then
    With pt.Cells(row + total, 1).Resize(, col).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    End If
    ''
    ''
End Sub



Sub printRst(data, row, col)
    
    Dim mySheet As Worksheet
    
    Set mySheet = Worksheets(rstSheet)
    Flag = mySheet.Cells(1, 1).Value
    Set pt = mySheet.Cells(Flag, 2)
    
    Range(pt(1, 1), pt(row, col)) = data
    
    mySheet.Cells(1, 1) = Flag + N + 4
    Worksheets(rstSheet).Rows(1).Hidden = True
    
End Sub
Sub Title1(contents As String)
    Dim Flag As Long
    Dim mySheet As Worksheet
    
    Set mySheet = Worksheets(rstSheet)
    Flag = mySheet.Cells(1, 1).Value
    Yp = mySheet.Cells(Flag + 2, 1).Top
    
    On Error Resume Next
    
    Set title = mySheet.Shapes.AddShape(msoShapeRectangle, 3.75, Yp + 2.25, 400, 25#)
    With title
        .Fill.ForeColor.SchemeColor = 57
        .Line.DashStyle = msoLineSolid
        .Line.Style = msoLineSingle
        .Line.ForeColor.SchemeColor = 8
        .Line.Weight = 1
        .Line.Visible = msoTrue
       ' .Shadow.Type = msoShadow1
    End With
   
    With title.TextFrame.Characters
        .Text = contents
        .Font.Name = "굴림"
        .Font.FontStyle = "굵게"
        .Font.Size = 14
        .Font.ColorIndex = 2
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
    
    mySheet.Cells(1, 1) = Flag + 4
End Sub

Sub Title2(contents As String)
    Dim Flag As Long
    Dim mySheet As Worksheet
    
    Set mySheet = Worksheets(rstSheet)
    Flag = mySheet.Cells(1, 1).Value
    Yp = mySheet.Cells(Flag + 2, 1).Top
    
    On Error Resume Next
    
    Set title = mySheet.Shapes.AddShape(msoShapeRectangle, 60.75, Yp, 250, 25#)
    With title
        .Fill.ForeColor.SchemeColor = 1
        .Solid
        .Line.ForeColor.SchemeColor = 8
        .Line.Visible = msoTrue
       ' .Shadow.Type = msoShadow17
    End With
   
    With title.TextFrame.Characters
        .Text = contents
        .Font.Name = "굴림"
        .Font.FontStyle = "굵게"
        .Font.Size = 11
        .Font.ColorIndex = xlAutomatic
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
    
    mySheet.Cells(1, 1) = Flag + 4
End Sub

Sub Title3(contents As String)

    Dim Flag As Long
    Dim mySheet As Worksheet
    
    Set mySheet = Worksheets(rstSheet)
    Flag = mySheet.Cells(1, 1).Value
    Yp = mySheet.Cells(Flag + 2, 1).Top
    
    On Error Resume Next
    
    Set title = mySheet.Shapes.AddShape(msoShapeRectangle, 60.75, Yp, 250, 22#)
    With title
        .Fill.ForeColor.SchemeColor = 1
        .Solid
        .Line.ForeColor.SchemeColor = 8
        .Line.Visible = msoTrue
        .Shadow.Type = msoShadow17
    End With
   
    With title.TextFrame.Characters
        .Text = contents
        .Font.Name = "굴림"
        .Font.FontStyle = "굵게"
        .Font.Size = 11
        .Font.ColorIndex = xlAutomatic
    End With
    title.TextFrame.HorizontalAlignment = xlCenter
    
    mySheet.Cells(1, 1) = Flag + 4
    
End Sub
