Attribute VB_Name = "Module2"
Public Const ANAL_OUT = "분석결과.xls"
Public Const ANAL_TITLE = "<통계분석결과>"

'새로운 워크북을 만드는 함수
Public Sub Addbook(szName As String, Optional bMusetNew As Boolean = False)  '시트이름을 받는다.
   Dim wbNew As Workbook
   Dim wbCur() As Workbook
   Dim stPath As String
   Dim i As Integer
   Dim Uheight As Double
   Dim Uwidth As Double
   Dim Sfile() As String
   
   Uheight = Application.UsableHeight
   Uwidth = Application.UsableWidth
   
   If Workbooks.Count <> 0 Then
      ReDim wbCur(1 To 1)
      Set wbCur(UBound(wbCur)) = ActiveWorkbook       '현재 시트를 저장해 둔다.
      For i = 1 To Workbooks.Count                    '모든 시트의 이름을 불러서 새로 만들 이름과 같은지 알아본다.
         If Workbooks(i).Name <> szName Then
            ReDim Preserve wbCur(1 To UBound(wbCur) + 1)
            Set wbCur(UBound(wbCur)) = Workbooks(i)
         End If
      Next
   End If
   
   With Workbooks("STEP.xla")
      stPath = .Path & "\"
   End With
      
   If Workbooks.Count <> 0 Then
      For i = 1 To Workbooks.Count                    '모든 시트의 이름을 불러서 새로 만들 이름과 같은지 알아본다.
         If Workbooks(i).Name = szName Then
            Set wbNew = Workbooks(szName)
            GoTo 20
         End If
      Next
   End If
         
   ReDim Sfile(1 To 1)
   Sfile(UBound(Sfile)) = Dir(stPath)
   Do While Sfile(UBound(Sfile)) <> ""
      ReDim Preserve Sfile(1 To UBound(Sfile) + 1)
      Sfile(UBound(Sfile)) = Dir
   Loop
   
   i = 1
   Do While Sfile(i) <> ""
      If Sfile(i) = szName Then
         Workbooks.Open Filename:=stPath & szName
         Set wbNew = Workbooks(szName)
         GoTo 20
      End If
      i = i + 1
   Loop
   Application.SheetsInNewWorkbook = 1
   Set wbNew = Workbooks.Add                       '새로운 시트를 만들고
      wbNew.SaveAs stPath & szName, , , , , False      '정해진 이름을 붙인다.
   
20:
   If Workbooks.Count = 1 Then
      With Windows(wbNew.Name)
         .WindowState = xlMaximized
      End With
      Exit Sub
   End If

   With Windows(wbNew.Name)
      .WindowState = xlNormal
      .Top = 0
      .Left = 0
      .Height = Uheight * 0.7
      .Width = Uwidth
   End With
   
   For i = Workbooks.Count To 1 Step -1
      With Windows(wbCur(i).Name)
         .WindowState = xlNormal
         .Top = Uheight - Uheight * 0.3
         .Left = 0
         .Height = Uheight * 0.3
         .Width = Uwidth
      End With
   Next
   
End Sub

'새로운 시트를 만드는 함수
Public Sub AddSheet(szName As String, Optional bMusetNew As Boolean = False)  '시트이름을 받는다.
   Dim shNew As Worksheet
   Dim i As Integer
   
   For i = 1 To Workbooks(ANAL_OUT).Worksheets.Count    '모든 시트의 이름을 불러서 새로 만들 이름과 같은지 알아본다.
      If Workbooks(ANAL_OUT).Sheets(i).Name = "Sheet1" Then
         Set shNew = Sheets(i)
         GoTo 10
      ElseIf Workbooks(ANAL_OUT).Sheets(i).Name = szName Then
         Exit Sub                                  '만약 같은이름이 있으면 더이상 진행하지 않는다.
      End If
   Next
   Set shNew = Workbooks(ANAL_OUT).Worksheets.Add       '새로운 시트를 만들고

10:
   shNew.Name = szName                             '정해진 이름을 붙인다.
   shNew.Activate                                  '활성화시킨후
   Cells.Font.Name = "굴림"
   Cells.Font.Size = 9
   With ActiveWindow
        .DisplayGridlines = False                  '그리드를 없애고
        .DisplayHeadings = False                   '번호표를 없앤다.
   End With
   
   With shNew.Cells(1, 1)                          '"a1" 셀에 처음프린트될것을 지정한다.
      .Value = "$b$3"
      .Font.ColorIndex = 2
   End With
   shNew.Rows(1).Hidden = True                     '"a1" 셀은 미관상 필요없기때문에 숨긴다.
   
   With shNew.Cells(1, 2)                          '"b1" 셀에 챠트가 처음프린트될것을 지정한다.
      .Value = 1
      .Font.ColorIndex = 2
   End With
   shNew.Rows(1).Hidden = True                     '"b1" 셀은 미관상 필요없기때문에 숨긴다.
      
End Sub


Public Sub AddTempSheet(szSheetName As String)
    Dim shTemp As Worksheet
    
    For Each shTemp In Worksheets
        If shTemp.Name = szSheetName Then Exit Sub
    Next
    
    Set shTemp = Worksheets.Add
    shTemp.Name = szSheetName
    shTemp.Cells(1, 1) = 1
    shTemp.Visible = xlSheetHidden
End Sub

Public Sub DelSheet(szSheetName As String)
   Dim shTemp As Worksheet
    
   For Each shTemp In Worksheets
      If shTemp.Name = szSheetName Then
         Application.DisplayAlerts = False
         shTemp.Delete
         Application.DisplayAlerts = True
         DoEvents
      End If
   Next
End Sub
