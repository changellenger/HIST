Attribute VB_Name = "Module2"
Sub OpenOutSheet(SheetName, Optional IsAddress As Boolean = False)
    
    Dim s, CurS As Worksheet
    
    Application.ScreenUpdating = False
    For Each s In ActiveWorkbook.Sheets
        If s.Name = SheetName Then Exit Sub
    Next s
    Set CurS = ActiveSheet: Set s = Worksheets.Add
    With ActiveWindow
        .DisplayGridlines = False
'        .DisplayHeadings = False
    End With
    
    With Cells
         .Font.Name = "±¼¸²"
         .Font.Size = 9
         .HorizontalAlignment = xlLeft
    End With

    s.Name = SheetName: CurS.Activate
    With Worksheets(SheetName).Range("a1")
        .Value = 2
        '''If IsAddress = True Then .Value = "A2"
        .Font.ColorIndex = 2
    End With
    Worksheets(SheetName).Rows(1).Hidden = True

    's.Protect Password:="prophet", DrawingObjects:=False, contents:=True, Scenarios:=True
    Application.ScreenUpdating = True
    
End Sub
