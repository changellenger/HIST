Attribute VB_Name = "combinationModule"


Sub comb(element)

  Dim InxComb As Integer
  Dim InxResult As Integer
  Dim TestData() As Variant
  Dim Result() As Variant
  Dim myArray() As Variant
  Dim temps As String
  Dim n As Integer
  
  n = 0         ' myarray 개수
    temps = ""
  TestData = element
  



  Call GenerateCombinations(TestData, Result)
  
  
     For InxResult = 0 To UBound(Result)
    For InxComb = 0 To UBound(Result(InxResult))
    If InxComb = 0 Then
     Else
   
        End If
        
    Next
       n = n + 1
       
  Next

ReDim myArray(n - 1) As Variant
  

  For InxResult = 0 To UBound(Result)
  '  Debug.Print Right("  " & InxResult + 1, 3) & " ";   '숫자 출력
    For InxComb = 0 To UBound(Result(InxResult))
    If InxComb = 0 Then
    temps = temps & "" & Result(InxResult)(InxComb) & " "
    
    '  Debug.Print "[" & Result(InxResult)(InxComb) & "] ";  ' 결과값 출력
    Else
    temps = temps & "* " & Result(InxResult)(InxComb) & " "
      '  Debug.Print "* [" & Result(InxResult)(InxComb) & "]";  ' 결과값 출력
        End If
        
    Next
   myArray(InxResult) = temps
   temps = ""
'   Debug.Print ' 줄바꿈  즉, 여기에 리스트 추가
   
  Next




doe3.ListBox1.list() = myArray


End Sub


Sub GenerateCombinations(ByRef AllFields() As Variant, _
                                             ByRef Result() As Variant)

  Dim InxResultCrnt As Integer
  Dim InxField As Integer
  Dim InxResult As Integer
  Dim i As Integer
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



Function combstr(element) As String



  Dim InxComb As Integer
  Dim InxResult As Integer
  Dim TestData() As Variant
  Dim Result() As Variant
  Dim myArray() As Variant
  Dim temps As String
  Dim n As Integer
  Dim strp As String
  
  
  n = 0         ' myarray 개수
    temps = ""
  TestData = element
  



  Call GenerateCombinations(TestData, Result)
  
  
     For InxResult = 0 To UBound(Result)
    For InxComb = 0 To UBound(Result(InxResult))
    If InxComb = 0 Then
     Else
   
        End If
        
    Next
       n = n + 1
       
  Next

ReDim myArray(n - 1) As Variant
  

  For InxResult = 0 To UBound(Result)
  '  Debug.Print Right("  " & InxResult + 1, 3) & " ";   '숫자 출력
    For InxComb = 0 To UBound(Result(InxResult))
    If InxComb = 0 Then
    temps = temps & "" & Result(InxResult)(InxComb) & " "
    
    '  Debug.Print "[" & Result(InxResult)(InxComb) & "] ";  ' 결과값 출력
    Else
    temps = temps & "* " & Result(InxResult)(InxComb) & " "
      '  Debug.Print "* [" & Result(InxResult)(InxComb) & "]";  ' 결과값 출력
        End If
        
    Next
    
    If strp = "" Then
    strp = temps
    
    Else
    
    strp = strp & "+" & temps
    
    End If
   temps = ""
'   Debug.Print ' 줄바꿈  즉, 여기에 리스트 추가
   
  Next


Debug.Print strp;


'doe3.ListBox1.List() = myarray

combstr = strp


End Function
