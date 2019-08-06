VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameF 
   OleObjectBlob   =   "frameF.frx":0000
   Caption         =   "F - 분포"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12720
   StartUpPosition =   1  '소유자 가운데
   TypeInfoVer     =   51
End
Attribute VB_Name = "frameF"
Attribute VB_Base = "0{67256E72-262B-40B5-8F0B-03F87564473B}{57E1BF7B-1BC7-4BFC-8186-04B0E42F569E}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False


Private TempWorkBookName As String
Private tmpSh(1 To 4) As Worksheet
Private temp1 As ChartObject
Sub CommandButton1_Click()
    n1 = Val(TextBox5.Text)
    n2 = Val(TextBox6.Text)
    fvl = Val(TextBox1.Text)
    p = Application.WorksheetFunction.FDist(fvl, n1, n2)
    TextBox2.Text = Format(p, "0.00000")
End Sub

Private Sub CommandButton13_Click()

End Sub

Sub CommandButton2_Click()
'On Error GoTo Err
    n1 = Val(TextBox5.Text)
    n2 = Val(TextBox6.Text)
    fp = Val(TextBox4.Text)
    fvl = Application.WorksheetFunction.FInv(fp, n1, n2)
    TextBox3.Text = Format(fvl, "0.00000")
  End Sub
Private Sub CommandButton24_Click()
    
    Static RowIndex4 As Long
    Dim TmpR As Range: Dim ir1, ir2 As String
    Dim df1, df2, i As Integer: Dim temp As Double
    
    df1 = Val(Me.TextBox5.Value)
    df2 = Val(Me.TextBox6.Value)

    For i = 2 To 52
        temp = (i - 2) * 0.2
        tmpSh(4).Cells(i - 1 + RowIndex4, 2).Formula = Fpdf(temp, df1, df2)
    Next i
    ir1 = tmpSh(4).Cells(1 + RowIndex4, 2).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    ir2 = tmpSh(4).Cells(51 + RowIndex4, 2).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    Set TmpR = tmpSh(4).Range(ir1 & ":" & ir2)
    RowIndex4 = RowIndex4 + 53
    
    For Each temp1 In tmpSh(4).ChartObjects
        temp1.Delete
    Next temp1
    
    Set temp1 = tmpSh(4).ChartObjects.Add(100, 100, 270, 228)
    temp1.Chart.ChartWizard Source:=TmpR, _
        Gallery:=xlLine, Format:=10, HasLegend:=False, _
        Title:="", CategoryTitle:="", ValueTitle:="확률"
    With temp1.Chart.SeriesCollection(1)
        .XValues = tmpSh(4).Range("A1:A51")
        .Border.ColorIndex = 3
        .Border.Weight = xlThin
    End With
    temp1.Chart.ChartArea.Interior.ColorIndex = 2
    
    With temp1.Chart.PlotArea
        .Interior.ColorIndex = 2
        .Border.LineStyle = xlAutomatic
        .Border.ColorIndex = 16
    End With
    With temp1.Chart.Axes(xlValue)
        .TickLabels.NumberFormat = "0.00"
        .MajorTickMark = xlNone
        .MinimumScale = 0
        .TickLabels.Font.Size = 8
        '.TickLabels.Font.Name = "Times New Roman"
        .AxisTitle.Orientation = xlVertical
        .AxisTitle.Font.Size = 8
    End With
    With temp1.Chart.Axes(xlCategory)
        .MajorTickMark = xlNone
        .TickLabels.NumberFormat = "0.0"
        .TickLabels.Font.Size = 8
        '.TickLabels.Font.Name = "Times New Roman"
        .TickLabels.Orientation = xlHorizontal
    End With
    temp1.Chart.Export Filename:="f.tmp", FilterName:="GIF"
    
    StorageForStatic temp1.Name, 6, False
    Me.CommandButton25.Enabled = True
    Me.CommandButton31.Enabled = True
    Me.Image4.Picture = LoadPicture("f.tmp")
    Kill "f.tmp"

End Sub
Private Sub CommandButton25_Click()
    
    Dim temp As String
    
    temp = StorageForStatic("", 6, True)
    tmpSh(4).ChartObjects(temp).Height = 274
    tmpSh(4).ChartObjects(temp).Width = 431
    tmpSh(4).ChartObjects(temp).Chart.Export Filename:="f.tmp", FilterName:="GIF"
    frmChart.Image1.Picture = LoadPicture("f.tmp")
    Kill "f.tmp"
    frmChart.Show

End Sub
Private Sub CommandButton31_Click()
    ChartOut 6, _
    "F-분포(df1=" & TextBox5.Value & _
    ", df2=" & TextBox6.Value & ")"
End Sub

Sub SpinButton1_Change()
    TextBox5.Text = SpinButton1.Value
End Sub
Sub SpinButton2_Change()
    TextBox6.Text = SpinButton2.Value
End Sub

 Sub TextBox1_Change()

End Sub

Sub TextBox2_Change()

End Sub

Sub TextBox4_Change()

End Sub

Sub TextBox5_Change()
    If TextBox5.Text <> "" Then
        SpinButton1.Value = Val(TextBox5.Text)
    End If
End Sub
Sub TextBox6_Change()
    If TextBox6.Text <> "" Then
        SpinButton2.Value = Val(TextBox6.Text)
    End If
End Sub

Private Sub ChartOut(ChartNum As Integer, Comment As String)            ''''"_그래프출력_"
    
    Dim temp As String: Dim tempCO As ChartObject
    Dim position As Range: Dim s, ttemp As Worksheet
    Dim tempheight, tempwidth As Double
    
    On Error GoTo sbcError
    
    OpenOutSheet "_통계분석결과_"
    temp = StorageForStatic("", ChartNum, True)
    Set tempCO = tmpSh(ChartNum - 2).ChartObjects(temp)
    tempheight = tempCO.Height: tempwidth = tempCO.Width
    tempCO.Height = 228: tempCO.Width = 270
    With tempCO.Chart
        .HasTitle = True
        .ChartTitle.Characters.Text = Comment
        .ChartTitle.Font.Size = 10
    End With

    Set s = Worksheets("_통계분석결과_")
    's.Unprotect "prophet"
    
    
    RstSheet = "_통계분석결과_"
    Title1 "그래프출력"

    
    Set position = s.Range("a1")
    tempCO.Chart.ChartArea.Copy
    
    Worksheets("_통계분석결과_").Activate
    Worksheets("_통계분석결과_").Cells(position.Value + 1, 2).Select
    Worksheets("_통계분석결과_").Cells(position.Value + 1, 2).Activate
    Worksheets("_통계분석결과_").Paste
    position.Value = position.Value + Int(245 / s.Range("a2").Height) + 1
    's.Protect Password:="prophet", DrawingObjects:=False, contents:=True, Scenarios:=True
    ActiveWindow.Visible = False
    
    tempCO.Height = tempheight: tempCO.Width = tempwidth
    tempCO.Chart.HasTitle = False
    MsgBox prompt:="그래프 출력이 완료되었습니다.", Title:="그래프 출력"
    Exit Sub

sbcError:
    MsgBox "출력시트를 정할 수 없습니다." & Chr(10) & _
    "[_통계분석결과_]이라는 이름의 시트를 삭제해 주십시오.", vbExclamation, Title:="출력 오류"
    
End Sub
Private Sub UserForm_Initialize()
    
    Dim i As Integer: Dim a, b As Double
    
    Me.TextBox5.Value = Val(Me.SpinButton1.Value) / 100
    
    
    TempWorkBookName = TempWorkbookOpen
    For i = 1 To 4
        Set tmpSh(i) = Workbooks(TempWorkBookName).Sheets(i)
    Next i
        
    For i = 2 To 52
        a = -5 + (i - 2) * 0.2
        b = Normal(-5 + (i - 2) * 0.2, 0, 1)
        tmpSh(2).Cells(i - 1, 1).Value = a
        tmpSh(2).Cells(i - 1, 2).Value = b
        tmpSh(3).Cells(i - 1, 1).Value = (i - 2) * 2
        tmpSh(4).Cells(i - 1, 1).Value = (i - 2) * 0.2
    Next i
End Sub
