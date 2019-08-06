VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameT 
   OleObjectBlob   =   "frameT.frx":0000
   Caption         =   "t - 분포"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12720
   StartUpPosition =   1  '소유자 가운데
   TypeInfoVer     =   42
End
Attribute VB_Name = "frameT"
Attribute VB_Base = "0{EB058C7F-7755-48B7-924D-38020F13C061}{A28D1545-72FF-4D4B-B4E7-FC418FA7FD70}"
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
    n = Val(TextBox5.Text)
    tv = Val(TextBox1.Text)
    
    If TextBox1.Text >= 0 Then
        p = Application.WorksheetFunction.TDist(tv, n, 1)
        TextBox2.Text = Format(p, "0.00000")
    Else
        p = Application.WorksheetFunction.TDist(-tv, n, 1)
        TextBox2.Text = Format(1 - p, "0.00000")
    End If
End Sub
Private Sub CommandButton10_Click()
    
    Static RowIndex2 As Long
    Dim df, i As Integer
    Dim temp As Double
    Dim ir1, ir2 As String: Dim TmpR As Range
    
    df = Val(Me.TextBox5.Value)
    
    For i = 2 To 52
        temp = -5 + 0.2 * (i - 2)
        tmpSh(2).Cells(i - 1 + RowIndex2, 3).Formula = Tpdf(temp, df)
    Next i
    ir1 = tmpSh(2).Cells(1 + RowIndex2, 3).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    ir2 = tmpSh(2).Cells(51 + RowIndex2, 3).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    Set TmpR = tmpSh(2).Range(ir1 & ":" & ir2)
    RowIndex2 = RowIndex2 + 53
    
    For Each temp1 In tmpSh(2).ChartObjects
        temp1.Delete
    Next temp1
    
    Set temp1 = tmpSh(2).ChartObjects.Add(100, 100, 270, 228)
    temp1.Chart.ChartWizard Source:=tmpSh(2).Range("B1:B51"), _
        Gallery:=xlLine, Format:=10, _
        Title:="", CategoryTitle:="", ValueTitle:="확률"
    With temp1.Chart.SeriesCollection(1)
        .XValues = tmpSh(2).Range("A1:A51")
        .Border.ColorIndex = 3
        .Border.Weight = xlThin
    End With
    temp1.Chart.SeriesCollection(1).Name = "N(0,1)"
    temp1.Chart.SeriesCollection.NewSeries
    With temp1.Chart.SeriesCollection(2)
        .Values = TmpR
        .Name = "t-분포"
        .Border.ColorIndex = 5
        .Border.Weight = xlThin
    End With
    With temp1.Chart
        .HasLegend = True
        .Legend.position = xlBottom
        .Legend.Font.Size = 8
        .ChartArea.Interior.ColorIndex = 2
        .PlotArea.Border.ColorIndex = 16
    End With
    With temp1.Chart.PlotArea
        .Interior.ColorIndex = 2
        .Border.LineStyle = xlAutomatic
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
    temp1.Chart.Export Filename:="tdist.tmp", FilterName:="GIF"
    
    StorageForStatic temp1.Name, 4, False
    Me.CommandButton11.Enabled = True
    Me.CommandButton29.Enabled = True
    Me.Image4.Picture = LoadPicture("tdist.tmp")
    Kill "tdist.tmp"

End Sub
Private Sub CommandButton11_Click()
    
    Dim temp As String
    
    temp = StorageForStatic("", 4, True)
    tmpSh(2).ChartObjects(temp).Height = 274
    tmpSh(2).ChartObjects(temp).Width = 431
    tmpSh(2).ChartObjects(temp).Chart.Export Filename:="t.tmp", FilterName:="GIF"
    frmChart.Image1.Picture = LoadPicture("t.tmp")
    Kill "t.tmp"
    frmChart.Show

End Sub

Sub CommandButton2_Click()
    n = Val(TextBox5.Text)
    tp = Val(TextBox4.Text)
    If tp < 0.5 Then
        tv = Application.WorksheetFunction.TInv(tp * 2, n)
    ElseIf tp = 0.5 Then
        tv = 0
    Else
        tv = -Application.WorksheetFunction.TInv((1 - tp) * 2, n)
    End If
    TextBox3.Text = Format(tv, "0.00000")
    

End Sub
Private Sub CommandButton29_Click()
    ChartOut 4, _
    "t-분포(df=" & TextBox5.Value & ")와 표준정규분포"
End Sub

Private Sub Image4_Click()

End Sub

Private Sub SpinButton1_Change()
    TextBox5.Text = SpinButton1.Value
End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub TextBox3_Change()

End Sub

Sub TextBox5_Change()
    If TextBox5.Text <> "" Then
        SpinButton1.Value = Val(TextBox5.Text)
    End If
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
    '''activePt = Worksheets(RstSheet).Cells(1, 1).Value

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
