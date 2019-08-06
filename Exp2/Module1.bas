Attribute VB_Name = "Module1"
'frmgen 일반선형모형
'frmgenout 결과옵션
'frmgenmodel 모형(교호작용)
'frmgengraph 그래프
'frmgencomp 비교
'frmgenop 등분산옵션


'frmoneway 일원배치분산분석
'frmonewaycomp 비교
'frmonewayop 등분산옵션
'frmonewaygraph 그래프


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
   rinterface.RRun "install.packages (" & Chr(34) & "FrF2" & Chr(34) & ")" ' : R 패키지 필요없음:
    rinterface.RRun "require (FrF2)"
    rinterface.RRun "install.packages (" & Chr(34) & "qualityTools" & Chr(34) & ")" ' : R 패키지 필요없음:
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
   If arrName(i) <> "" Then                     '빈칸제거
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
   If arrName(i) <> "" Then                     '빈칸제거
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
   If arrName(i) <> "" Then                     '빈칸제거
   myArray(a) = arrName(i)
   a = a + 1
   
   Else:
   End If
   Next i
    
    doe5.ComboBox1.list() = myArray
    doe5.Show

End Sub
