VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameCrfre 
   OleObjectBlob   =   "frameCrfre.frx":0000
   Caption         =   "�����м�"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4140
   StartUpPosition =   1  '������ ���
   TypeInfoVer     =   71
End
Attribute VB_Name = "frameCrfre"
Attribute VB_Base = "0{DD0ADB77-4F06-4BB9-AD52-68C01514899B}{36FABEC2-E16A-490A-9139-89D13AFA0C56}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False





Private Sub BtnOK_Click()                               ''''"_�������ڷ�м����_"
  Dim TempSheet As Worksheet
  Dim resultsheet As Worksheet
  Dim temp As Range
  Dim tm As Range
  Dim rname() As String
  Dim cname() As String
  Dim rtotal() As Long
  Dim ctotal() As Long
  Dim Expect() As Double
  Dim total As Long
  Dim chisq As Double
  Set TempSheet = ActiveCell.Worksheet
  Set temp = TempSheet.Cells.CurrentRegion
  c = temp.Columns.count - 1
  r = temp.Rows.count - 1
      '''���� üũ
    Set tm = temp.Offset(1, 1).Resize(r, c)
    If FindingRangeError2(tm) = True Then
          MsgBox "�м������� ���ڳ� ������ �ֽ��ϴ�.", vbExclamation, "HIST"
          Exit Sub
    End If
  Set tm = temp
  ReDim rname(1 To r)
  ReDim cname(1 To c)
  ReDim rtotal(1 To r)
  ReDim ctotal(1 To c)
  Set tp = temp.Columns(1)
  For i = 2 To r + 1
      rname(i - 1) = tp.Cells(i, 1)
  Next i
  Set tp = temp.Rows(1)
  For i = 2 To c + 1
      cname(i - 1) = tp.Cells(1, i)
  Next i
  Set temp = temp.Offset(1, 1)
  total = Application.sum(temp)
  For i = 1 To r
      rtotal(i) = Application.sum(temp.Rows(i))
  Next i
  For i = 1 To c
      ctotal(i) = Application.sum(temp.Columns(i))
  Next i
  ReDim Expect(1 To r, 1 To c)
  chisq = 0
  For i = 1 To r
      For J = 1 To c
          Expect(i, J) = rtotal(i) * ctotal(J) / total
          chisq = chisq + (temp.Cells(i, J).Value - Expect(i, J)) ^ 2 / Expect(i, J)
      Next J
  Next i
  Set resultsheet = OpenOutSheet2("_���м����_", True)
  
  
    '''
    '''
    '''
    rstSheet = "_���м����_"
    
    '������ �Է�
On Error GoTo Err_delete
Dim val3535 As Long '�ʱ���ġ ������ ����'
Dim s3535 As Worksheet
val3535 = 2
    For Each s3535 In ActiveWorkbook.Sheets
        If s3535.Name = rstSheet Then
val3535 = Sheets(rstSheet).Cells(1, 1).Value
End If
Next s3535  '��Ʈ�� �̹������� ��� ��ġ �������ϰ�, ������ 2�� �����Ѵ�.
    
    'Worksheets(RstSheet).Unprotect "prophet"
    activePt = Worksheets(rstSheet).Cells(1, 1).Value
    

  
  'resultsheet.Unprotect "prophet"
  Conti_Result.cResult r, c, temp, Expect, rtotal, ctotal, chisq, rname, cname, resultsheet
  'resultsheet.Protect Password:="prophet", DrawingObjects:=False, contents:=True, Scenarios:=True
    
    
    
    '''Worksheets(RstSheet).Cells(1, 1).Value = Right(Worksheets(RstSheet).Cells(1, 1).Value, Len(Worksheets(RstSheet).Cells(1, 1).Value) - 3)

    


    Worksheets(rstSheet).Activate

    '���� ���� üũ �� �񱳰� ����
    Dim Cmp_Value As Long
    
    If PublicModule.ChkVersion(ActiveWorkbook.Name) = True Then
        Cmp_Value = 1048000
    Else
        Cmp_Value = 65000
    End If
    
    If Worksheets(rstSheet).Cells(1, 1).Value > Cmp_Value Then
        MsgBox "[_���м����_]��Ʈ�� ���� ��� ����Ͽ����ϴ�." & vbCrLf & "�� ��Ʈ�� �̸��� �ٲٰų� ������ �ּ���", vbExclamation, "HIST"
        Exit Sub
    End If

    Worksheets(rstSheet).Cells(activePt + 10, 1).Select
    Worksheets(rstSheet).Cells(activePt + 10, 1).Activate
                            '��� �м��� ���۵Ǵ� �κ��� �����ָ� ��ģ��.
                            
    Unload Me
    
'�ǵڿ� ���̱�
Exit Sub
Err_delete:

For Each s3535 In ActiveWorkbook.Sheets
        If s3535.Name = rstSheet Then
Sheets(rstSheet).Range(Cells(val3535, 1), Cells(5000, 1000)).Select
Selection.Delete
Sheets(rstSheet).Cells(1, 1) = val3535
Sheets(rstSheet).Cells(val3535, 1).Select

If val3535 = 2 Then
Application.DisplayAlerts = False
Sheets(rstSheet).Delete
End If

End If


Next s3535

MsgBox ("���α׷��� ������ �ֽ��ϴ�.")
 'End sub �տ��� ���δ�.

''�ؼ�, ������ ���� Err_delete�� �ͼ� ù�����ķ� �����. ���� ù���� 2�� ��Ʈ�� �����.�׸��� �����޽��� ���
'rSTsheet����⵵ ���� �������� ��쿡�� �ƹ� ���۵� ���� �ʰ�, �����޽����� ����.
End Sub




Private Sub OptionButton1_Click()
   
   Dim myRange As Range
   Dim myArray()
   Dim arrName As Variant
   Dim TempSheet As Worksheet
   Set TempSheet = ActiveCell.Worksheet
   
    ReDim arrName(TempSheet.UsedRange.Columns.count)
' Reading Data
    For i = 1 To TempSheet.UsedRange.Columns.count
        arrName(i) = TempSheet.Cells(1, i)
    Next i
   
   Me.Listbox1.Clear

    ReDim myArray(TempSheet.UsedRange.Columns.count - 1)
    a = 0
   For i = 1 To TempSheet.UsedRange.Columns.count
   If arrName(i) <> "" Then                     '��ĭ����
   myArray(a) = arrName(i)
   a = a + 1
   
   Else:
   End If
   Next i
  
   Me.Listbox1.list() = myArray



End Sub

Private Sub UserForm_Click()

End Sub
