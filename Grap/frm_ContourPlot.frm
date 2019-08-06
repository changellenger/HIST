VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_ContourPlot 
   OleObjectBlob   =   "frm_ContourPlot.frx":0000
   Caption         =   "등고선도/표면도"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4200
   StartUpPosition =   1  '소유자 가운데
   TypeInfoVer     =   15
End
Attribute VB_Name = "frm_ContourPlot"
Attribute VB_Base = "0{51343001-2273-44AC-A3C4-2C72CB64FAB8}{AD6BB69A-B47C-49D7-B6A9-DC11B97A78C5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Option Base 1


Private Sub cmd_OK_Click()

    Dim rngFirst As Range, rngData As Range, rngPlot As Range
    Dim strName As String
    Dim i As Long
    
    On Error GoTo ErrEnd
    Application.ScreenUpdating = False
    
    Set rngData = Range(Me.RefEdit1)
    
    TModulePrint.Title1 "그래프 출력"
    
    TModulePrint.Title3 "등고선도와 표면도"
'   Insert Worshsheet
    For i = 1 To Sheets.count
        If Sheets(i).Name = "_통계분석결과_" Then
            GoTo 31
        Else
            GoTo 32
        End If
        
    
        
32: Next i


    Worksheets.Add Before:=Worksheets(1)
    ActiveSheet.Name = "_통계분석결과_"
    ActiveWindow.DisplayGridlines = False
    
    Cells(1, 1) = 1

    
    
31: Sheets("_통계분석결과_").Activate
    Application.ScreenUpdating = False

    strName = ActiveSheet.Name
    Set rngFirst = Cells(Cells(1, 1), 1)
     

    If Me.OptionButton3 Then 'All charts
        Set rngPlot = Range(rngFirst.Offset(3, 1), rngFirst.Offset(18, 5))
        Call Contour_Plot(rngData, rngPlot)
        Set rngPlot = Range(rngFirst.Offset(3, 6), rngFirst.Offset(18, 10))
        Call Surface_Plot(rngData, rngPlot)
    End If
    
    If Me.OptionButton1 Then 'Contour Plot
        Set rngPlot = Range(rngFirst.Offset(3, 1), rngFirst.Offset(18, 5))
        Call Contour_Plot(rngData, rngPlot)
    End If
    
    If Me.OptionButton2 Then 'Surface Plot
        Set rngPlot = Range(rngFirst.Offset(3, 1), rngFirst.Offset(18, 5))
        Call Surface_Plot(rngData, rngPlot)
    End If
    
    
ErrEnd:

'   Page number reset
 '   rngFirst = "Created at " & Now()
    Application.ScreenUpdating = True
    Application.Goto rngFirst, Scroll:=True
    Cells(1, 1) = Cells(1, 1) + 30
    
    Unload Me

    
End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub

Private Sub UserForm_Click()

End Sub
