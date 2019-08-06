VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frameContourline 
   OleObjectBlob   =   "frameContourline.frx":0000
   Caption         =   "등고선도"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6780
   StartUpPosition =   1  '소유자 가운데
   TypeInfoVer     =   47
End
Attribute VB_Name = "frameContourline"
Attribute VB_Base = "0{0BA410FA-6962-4546-A746-D06B5FAF1FC1}{0759D976-FE9B-4155-A0EB-24E5F83BBF0E}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Option Base 1


Private Sub CommandButton4_Click()

    Dim rngFirst As Range, rngData As Range, rngPlot As Range
    Dim strName As String
    Dim i As Long
    
    On Error GoTo ErrEnd
    Application.ScreenUpdating = False
    
    Set rngData = Range(Me.RefEdit1)
    
    
    
    Dim Vname(1 To 3) As String
      Vname(1) = PublicModule.SelectedVariable(Me.ListBox2.List(0), x, True)
      Vname(2) = PublicModule.SelectedVariable(Me.ListBox3.List(0), y, True)
      Vname(3) = PublicModule.SelectedVariable(Me.ListBox4.List(0), z, True)
    
'   Insert Worshsheet
    For i = 1 To Sheets.count
        If Sheets(i).Name = "_graph_" Then
            GoTo 31
        Else
            GoTo 32
        End If
32: Next i
    Worksheets.Add Before:=Worksheets(1)
    ActiveSheet.Name = "_graph_"
    ActiveWindow.DisplayGridlines = False
    
    Cells(1, 1) = 1

31: Sheets("_graph_").Activate
    Application.ScreenUpdating = False

    strName = ActiveSheet.Name
    Set rngFirst = Cells(Cells(1, 1) + 2, 1)
     

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
    rngFirst = "Created at " & Now()
    Application.ScreenUpdating = True
    Application.Goto rngFirst, Scroll:=True
    Cells(1, 1) = Cells(1, 1) + 30
    
    Unload Me

    
End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub

Private Sub OptionButton1_Click()

End Sub

Private Sub UserForm_Click()

End Sub
