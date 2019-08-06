Attribute VB_Name = "opentemp"
Function opentemp() As Worksheet
    
    Dim Flag As Boolean: Dim ws As Worksheet
    
    For Each ws In Worksheets
        If ws.name = "_TempData_" Then
            Flag = True
            Set TempWorksheet = ws
            Exit For
        End If
    Next ws
    
    If Flag = False Then
        Set TempWorksheet = Worksheets.Add
        TempWorksheet.name = "_TempData_"
        TempWorksheet.Cells(1, 1).Value = TempWorksheet.Cells(1, 2).Address
    End If
    TempWorksheet.Visible = False
    Set opentemp = TempWorksheet
        
End Function
