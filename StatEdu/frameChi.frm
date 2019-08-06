VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameChi 
   OleObjectBlob   =   "frameChi.frx":0000
   Caption         =   "카이제곱분포"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12630
   StartUpPosition =   1  '소유자 가운데
   TypeInfoVer     =   52
End
Attribute VB_Name = "frameChi"
Attribute VB_Base = "0{C6F56424-85F6-4F52-A3F8-0A508EF0751E}{7CBE5903-356A-4E1F-89BD-A2E9785A91EA}"
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
    cv = Val(TextBox1.Text)
    p = Application.WorksheetFunction.ChiDist(cv, n)
    TextBox2.Text = Format(p, "0.00000")
End Sub

Private Sub CommandButton13_Click()

    
    Static RowIndex3 As Long
    Dim TmpR As Range: Dim ir1, ir2 As String
    Dim df, i As Integer: Dim temp As Double
    
    df = Val(Me.TextBox5.Value)

    For i = 2 To 52
        temp = (i - 2) * 2
        tmpSh(3).Cells(i - 1 + RowIndex3, 2).Formula = Chipdf(temp, df)
    Next i
    ir1 = tmpSh(3).Cells(1 + RowIndex3, 2).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    ir2 = tmpSh(3).Cells(51 + RowIndex3, 2).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    Set TmpR = tmpSh(3).Range(ir1 & ":" & ir2)
    RowIndex3 = RowIndex3 + 53
    
    For Each temp1 In tmpSh(3).ChartObjects
        temp1.Delete
    Next temp1

    Set temp1 = tmpSh(3).ChartObjects.Add(100, 100, 270, 228)
    temp1.Chart.ChartWizard Source:=TmpR, _
        Gallery:=xlLine, Format:=10, HasLegend:=False, _
        Title:="", CategoryTitle:="", ValueTitle:="확률"
    With temp1.Chart.SeriesCollection(1)
        .XValues = tmpSh(3).Range("A1:A51")
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
        .TickLabels.Font.Size = 8
        '.TickLabels.Font.Name = "Times New Roman"
        .TickLabels.Orientation = xlHorizontal
    End With
    temp1.Chart.Export Filename:="chi.tmp", FilterName:="GIF"
    
    StorageForStatic temp1.Name, 5, False
    Me.CommandButton14.Enabled = True
    Me.CommandButton30.Enabled = True
    Me.Image5.Picture = LoadPicture("chi.tmp")
    Kill "chi.tmp"
    
End Sub
Private Sub CommandButton14_Click()
    
    Dim temp As String
    
    temp = StorageForStatic("", 5, True)
    tmpSh(3).ChartObjects(temp).Height = 274
    tmpSh(3).ChartObjects(temp).Width = 431
    tmpSh(3).ChartObjects(temp).Chart.Export Filename:="chi.tmp", FilterName:="GIF"
    frmChart.Image1.Picture = LoadPicture("chi.tmp")
    Kill "chi.tmp"
    frmChart.Show

End Sub

Sub CommandButton2_Click()
    n = Val(TextBox5.Text)
    cp = Val(TextBox4.Text)
    cv = Application.WorksheetFunction.ChiInv(cp, n)
    TextBox3.Text = Format(cv, "0.00000")
End Sub
Private Sub CommandButton30_Click()
    ChartOut 5, _
    "카이제곱분포(df=" & TextBox5.Value & ")"
End Sub

Private Sub Image5_Click()

End Sub

Private Sub MultiPage1_Change()

End Sub

Private Sub SpinButton1_Change()
    TextBox5.Text = SpinButton1.Value
End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub TextBox2_Change()

End Sub

Private Sub TextBox3_Change()

End Sub

Private Sub TextBox4_Change()

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
