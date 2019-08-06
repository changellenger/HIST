Attribute VB_Name = "Module2"
Sub ShowfrmDisc() '기술통계
    frmDisc.OptionButton1.Value = True
    frmDisc.Show
   End Sub
Sub ShowfrmFrequency() '빈도분석
    frmFrequency.OptionButton1.Value = True
    frmFrequency.Show
   End Sub
Sub ShowframeNor() '정규성검정
    frameNor.OptionButton1.Value = True
    frameNor.Show
End Sub
Sub ShowframeLe() '등분산성검정
    frameLe.OptionButton1.Value = True
    frameLe.Show
End Sub
Sub ShowframeCrfre() '교차분석
    frameCrfre.OptionButton1.Value = True
    frameCrfre.Show
End Sub
