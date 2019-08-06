Attribute VB_Name = "pgControl"
'Option Private Module
Public DataSheet As String, RstSheet As String

Public Declare Function ShellExecute _
 Lib "shell32.dll" _
 Alias "ShellExecuteA" ( _
 ByVal hwnd As Long, _
 ByVal lpOperation As String, _
 ByVal lpFile As String, _
 ByVal lpParameters As String, _
 ByVal lpDirectory As String, _
 ByVal nShowCmd As Long) _
 As Long
Function TempWorkbookOpen() As String

    Dim t As Workbook: Dim sin As Long
    
    sin = Application.SheetsInNewWorkbook
    Application.SheetsInNewWorkbook = 6
    Set t = Workbooks.Add
    Application.SheetsInNewWorkbook = sin
    ActiveWindow.Visible = False
    TempWorkbookOpen = t.Name
    
End Function

Sub TempWorkbookClose(workbookname)
    Workbooks(workbookname).Close False
End Sub

Function StorageForStatic(ChartName As String, _
    ChartNum As Integer, Output As Boolean) As String
    
    Static NewChartName(1 To 6) As String
    
    If Output = False Then
        NewChartName(ChartNum) = ChartName
        StorageForStatic = ""
    Else
        StorageForStatic = NewChartName(ChartNum)
    End If
    
End Function

Sub frmSam1Show()
    frmSam1.Show
End Sub

Sub frmSam2Show()
    frmSam2.Show
End Sub

'Sub OpenOutSheet(sheetName)
    
'    Dim s As Worksheet
        
'    For Each s In ActiveWorkbook.Sheets
'        If s.Name = sheetName Then Exit Sub
'    Next s
'    Set s = Worksheets.Add: s.Name = sheetName
'    With Worksheets(sheetName).Range("a1")
'        .Value = 1
'        .Font.ColorIndex = 2
'    End With
'    s.Protect Password:="prophet", DrawingObjects:=False, contents:=True, Scenarios:=True

'End Sub
        
Sub OpenOutSheet(sheetName, Optional IsAddress As Boolean = False)
    
    Dim s, CurS As Worksheet
    
    Application.ScreenUpdating = False
    For Each s In ActiveWorkbook.Sheets
        If s.Name = sheetName Then Exit Sub
    Next s
    Set CurS = ActiveSheet: Set s = Worksheets.Add
    With ActiveWindow
        .DisplayGridlines = False
'        .DisplayHeadings = False
    End With
    
    With Cells
         .Font.Name = "±¼¸²"
         .Font.Size = 9
         .HorizontalAlignment = xlRight
    End With

    s.Name = sheetName: CurS.Activate
    With Worksheets(sheetName).Range("a1")
        .Value = 2
        '''If IsAddress = True Then .Value = "A2"
        .Font.ColorIndex = 2
    End With
    Worksheets(sheetName).Rows(1).Hidden = True
    Worksheets(sheetName).Activate
    Cells.Select
    Selection.RowHeight = 13.5
    

    's.Protect Password:="prophet", DrawingObjects:=False, contents:=True, Scenarios:=True
    Application.ScreenUpdating = True
    
End Sub
