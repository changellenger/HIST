Attribute VB_Name = "DOE_Matrix"
Function xmatrix(xlist, n, m, t)
Dim dataRange As Range
Dim i As Long, j As Integer
Dim X()
Set dataRange = Worksheets(DataSheet).Cells.CurrentRegion

ReDim X(n - 1, t)

For i = 0 To t - 1
    For j = 0 To m
        If xlist(i) = dataRange.Cells(1, j + 1).value Then
            For k = 1 To n
                X(k - 1, i) = dataRange.Cells(k + 1, j + 1).value
            Next k
        End If
    Next j
Next i

xmatrix = X

End Function

Function ymatrix(ylist, n, m) As Variant
Dim dataRange As Range
    Dim i As Long, j As Integer
    Dim Y()
        
    Set dataRange = Worksheets(DataSheet).Cells.CurrentRegion
    
    ReDim Y(n - 1, 0)
    For j = 0 To m
        If ylist = dataRange.Cells(1, j + 1).value Then
            For i = 1 To n
                Y(i - 1, 0) = dataRange.Cells(i + 1, j + 1).value
            Next i
        End If
    Next j
    
    ymatrix = Y
End Function
