Attribute VB_Name = "Module2"
Public Const ANAL_OUT = "�м����.xls"
Public Const ANAL_TITLE = "<���м����>"

'���ο� ��ũ���� ����� �Լ�
Public Sub Addbook(szName As String, Optional bMusetNew As Boolean = False)  '��Ʈ�̸��� �޴´�.
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
      Set wbCur(UBound(wbCur)) = ActiveWorkbook       '���� ��Ʈ�� ������ �д�.
      For i = 1 To Workbooks.Count                    '��� ��Ʈ�� �̸��� �ҷ��� ���� ���� �̸��� ������ �˾ƺ���.
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
      For i = 1 To Workbooks.Count                    '��� ��Ʈ�� �̸��� �ҷ��� ���� ���� �̸��� ������ �˾ƺ���.
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
   Set wbNew = Workbooks.Add                       '���ο� ��Ʈ�� �����
      wbNew.SaveAs stPath & szName, , , , , False      '������ �̸��� ���δ�.
   
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

'���ο� ��Ʈ�� ����� �Լ�
Public Sub AddSheet(szName As String, Optional bMusetNew As Boolean = False)  '��Ʈ�̸��� �޴´�.
   Dim shNew As Worksheet
   Dim i As Integer
   
   For i = 1 To Workbooks(ANAL_OUT).Worksheets.Count    '��� ��Ʈ�� �̸��� �ҷ��� ���� ���� �̸��� ������ �˾ƺ���.
      If Workbooks(ANAL_OUT).Sheets(i).Name = "Sheet1" Then
         Set shNew = Sheets(i)
         GoTo 10
      ElseIf Workbooks(ANAL_OUT).Sheets(i).Name = szName Then
         Exit Sub                                  '���� �����̸��� ������ ���̻� �������� �ʴ´�.
      End If
   Next
   Set shNew = Workbooks(ANAL_OUT).Worksheets.Add       '���ο� ��Ʈ�� �����

10:
   shNew.Name = szName                             '������ �̸��� ���δ�.
   shNew.Activate                                  'Ȱ��ȭ��Ų��
   Cells.Font.Name = "����"
   Cells.Font.Size = 9
   With ActiveWindow
        .DisplayGridlines = False                  '�׸��带 ���ְ�
        .DisplayHeadings = False                   '��ȣǥ�� ���ش�.
   End With
   
   With shNew.Cells(1, 1)                          '"a1" ���� ó������Ʈ�ɰ��� �����Ѵ�.
      .Value = "$b$3"
      .Font.ColorIndex = 2
   End With
   shNew.Rows(1).Hidden = True                     '"a1" ���� �̰��� �ʿ���⶧���� �����.
   
   With shNew.Cells(1, 2)                          '"b1" ���� íƮ�� ó������Ʈ�ɰ��� �����Ѵ�.
      .Value = 1
      .Font.ColorIndex = 2
   End With
   shNew.Rows(1).Hidden = True                     '"b1" ���� �̰��� �ʿ���⶧���� �����.
      
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
