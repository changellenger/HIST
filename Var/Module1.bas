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

Sub Showdoe3()
 Dim myRange As Range
   Dim MyArray(), SubArray()
   Dim arrName As Variant
   Dim TempSheet As Worksheet
   Set TempSheet = ActiveCell.Worksheet
   
    ReDim arrName(TempSheet.UsedRange.Columns.count)
' Reading Data
    For I = 1 To TempSheet.UsedRange.Columns.count
        arrName(I) = TempSheet.Cells(1, I)
    Next I
   
   doe3.ComboBox1.Clear

    ReDim MyArray(TempSheet.UsedRange.Columns.count - 1)
    a = 0
   For I = 1 To TempSheet.UsedRange.Columns.count
   If arrName(I) <> "" Then                     '빈칸제거
   MyArray(a) = arrName(I)
   a = a + 1
   
   Else:
   End If
   Next I
   
 '  For q = 0 To Me.ListBox1.count
  ' SubArray(q) = MyArray(0) & "*" & MyArray(q + 1)
  ' Next q
   
 ' Test (MyArray)
  
  
doe3.ComboBox1.List() = MyArray
'doe3.OptionButton1.value = True
'element = Array("가", "나", "다", "라", "마", "바", "사")
doe3.Show

End Sub
Sub test2()
element = Array("k", "g", "b")
Test (element)

End Sub
Sub Test(element)

  Dim InxComb As Integer
  Dim InxResult As Integer
  Dim TestData() As Variant
  Dim Result() As Variant

  TestData = element
  
  Call GenerateCombinations(TestData, Result)

  For InxResult = 0 To UBound(Result)
    Debug.Print Right("  " & InxResult + 1, 3) & " ";   '숫자 출력
    For InxComb = 0 To UBound(Result(InxResult))
      Debug.Print "[" & Result(InxResult)(InxComb) & "] ";  ' 결과값 출력
    Next
    Debug.Print
  Next

End Sub

Sub GenerateCombinations(ByRef AllFields() As Variant, _
                                             ByRef Result() As Variant)

  Dim InxResultCrnt As Integer
  Dim InxField As Integer
  Dim InxResult As Integer
  Dim I As Integer
  Dim NumFields As Integer
  Dim Powers() As Integer
  Dim ResultCrnt() As String

  NumFields = UBound(AllFields) - LBound(AllFields) + 1

  ReDim Result(0 To 2 ^ NumFields - 2)  ' one entry per combination
  ReDim Powers(0 To NumFields - 1)          ' one entry per field name

  ' Generate powers used for extracting bits from InxResult
  For InxField = 0 To NumFields - 1
    Powers(InxField) = 2 ^ InxField
  Next

 For InxResult = 0 To 2 ^ NumFields - 2
    ' Size ResultCrnt to the max number of fields per combination
    ' Build this loop's combination in ResultCrnt
    ReDim ResultCrnt(0 To NumFields - 1)
    InxResultCrnt = -1
    For InxField = 0 To NumFields - 1
      If ((InxResult + 1) And Powers(InxField)) <> 0 Then
        ' This field required in this combination
        InxResultCrnt = InxResultCrnt + 1
        ResultCrnt(InxResultCrnt) = AllFields(InxField)
      End If
    Next
    ' Discard unused trailing entries
    ReDim Preserve ResultCrnt(0 To InxResultCrnt)
    ' Store this loop's combination in return array
    Result(InxResult) = ResultCrnt
  Next

End Sub
