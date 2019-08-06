Attribute VB_Name = "openresult"
Function openResultSheet(name) As Worksheet
    
    Dim Flag As Boolean: Dim ws As Worksheet
    Dim TempWindow As Window
    Dim TempWorksheet As Worksheet
    For Each ws In Worksheets
        If ws.name = name & "_Result" Then
            Flag = True
            Set TempWorksheet = ws
            Exit For
        End If
    Next ws
    
    If Flag = False Then
        Set TempWorksheet = Worksheets.Add
        TempWorksheet.name = name & "_Result"
       Set TempWindow = TempWorksheet.Application.ActiveWindow
        With TempWindow
              .DisplayGridlines = False
'              .DisplayHeadings = False
        End With
        With Cells
             .Font.Size = 9
             .HorizontalAlignment = xlRight
        End With
    End If
       
    Set openResultSheet = TempWorksheet
        
End Function

        
