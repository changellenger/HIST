Attribute VB_Name = "Module1"
'frmgen �Ϲݼ�������
'frmgenout ����ɼ�
'frmgenmodel ����(��ȣ�ۿ�)
'frmgengraph �׷���
'frmgencomp ��
'frmgenop ��л�ɼ�


'frmoneway �Ͽ���ġ�л�м�
'frmonewaycomp ��
'frmonewayop ��л�ɼ�
'frmonewaygraph �׷���


Sub Showfrmoneway()
frmoneway.Show

End Sub
Sub Showfrmgen()
frmgen.Show

End Sub

Sub showdoe1()
 Application.run "Exp2.xlam!ModuleControl.FactDE"
End Sub
Sub showdoe2()
   rinterface.RRun "install.packages (" & Chr(34) & "FrF2" & Chr(34) & ")" ' : R ��Ű�� �ʿ����:
    rinterface.RRun "require (FrF2)"
    rinterface.RRun "install.packages (" & Chr(34) & "qualityTools" & Chr(34) & ")" ' : R ��Ű�� �ʿ����:
    rinterface.RRun "require (qualityTools)"
    rinterface.RRun "arrayfrac <- fracChoose()"
    rinterface.RRun "output<-as.data.frame(arrayfrac)" 'ok
    
    rinterface.GetDataframe "output", Range("A1")
    Dim lastColumn As Integer
    lastColumn = ActiveCell.Worksheet.UsedRange.Columns.count
    
   ' lastColumn = Sheet1.Cells(1, Columns.Count).End(xlToLeft).Column
    'MsgBox lastColumn
   ' MsgBox Worksheets("Sheet11").Range("G1").value
  '  MsgBox Worksheets("Sheet11").Cells(1, lastColumn).value
  '  MsgBox ActiveSheet.Cells(1, lastColumn).value
    'Worksheets("Sheet1").Range("A1").Value = 100
    If ActiveSheet.Cells(1, lastColumn).value = "y" Then
    ActiveSheet.Cells(1, lastColumn).value = "Response"
        '  If Worksheets("Sheet11").Cells(1, lastColumn).value = "y" Then
  '  Worksheets("Sheet11").Cells(1, lastColumn).value = "Response"
    End If

End Sub
Sub Showdoe3()
 Dim myRange As Range
   Dim myArray(), SubArray()
   Dim arrName As Variant
   Dim TempSheet As Worksheet
   Set TempSheet = ActiveCell.Worksheet
   
   ReDim arrName(TempSheet.UsedRange.Columns.count)
' Reading Data
   For i = 1 To TempSheet.UsedRange.Columns.count
        arrName(i) = TempSheet.Cells(1, i)
   Next i
   
   doe3.ComboBox1.Clear

   ReDim myArray(TempSheet.UsedRange.Columns.count - 1)
   a = 0
   For i = 1 To TempSheet.UsedRange.Columns.count
   If arrName(i) <> "" Then                     '��ĭ����
   myArray(a) = arrName(i)
   a = a + 1
   
   Else:
   End If
   Next i
    
    doe3.ComboBox1.list() = myArray
    doe3.Show

End Sub
Sub Showdoe4()
 Dim myRange As Range
   Dim myArray(), SubArray()
   Dim arrName As Variant
   Dim TempSheet As Worksheet
   Set TempSheet = ActiveCell.Worksheet
   
   ReDim arrName(TempSheet.UsedRange.Columns.count)
' Reading Data
   For i = 1 To TempSheet.UsedRange.Columns.count
        arrName(i) = TempSheet.Cells(1, i)
   Next i
   
   doe4.ComboBox1.Clear

   ReDim myArray(TempSheet.UsedRange.Columns.count - 1)
   a = 0
   For i = 1 To TempSheet.UsedRange.Columns.count
   If arrName(i) <> "" Then                     '��ĭ����
   myArray(a) = arrName(i)
   a = a + 1
   
   Else:
   End If
   Next i
    
    doe4.ComboBox1.list() = myArray
    doe4.Show

End Sub
Sub Showdoe5()
 Dim myRange As Range
   Dim myArray(), SubArray()
   Dim arrName As Variant
   Dim TempSheet As Worksheet
   Set TempSheet = ActiveCell.Worksheet
    
   ReDim arrName(TempSheet.UsedRange.Columns.count)
' Reading Data
   For i = 1 To TempSheet.UsedRange.Columns.count
        arrName(i) = TempSheet.Cells(1, i)
   Next i
   
   doe5.ComboBox1.Clear

   ReDim myArray(TempSheet.UsedRange.Columns.count - 1)
   a = 0
   For i = 1 To TempSheet.UsedRange.Columns.count
   If arrName(i) <> "" Then                     '��ĭ����
   myArray(a) = arrName(i)
   a = a + 1
   
   Else:
   End If
   Next i
    
    doe5.ComboBox1.list() = myArray
    doe5.Show

End Sub
