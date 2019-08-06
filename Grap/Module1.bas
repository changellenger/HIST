Attribute VB_Name = "Module1"
Sub Showscatter()
frameScatterdiagram.ListBox2.Clear
frameScatterdiagram.ListBox3.Clear
frameScatterdiagram.OptionButton1.Value = False
frameScatterdiagram.OptionButton1.Value = True
frameScatterdiagram.CommandButton1.Visible = True
frameScatterdiagram.CommandButton2.Visible = True
frameScatterdiagram.CommandButton3.Visible = False
frameScatterdiagram.CommandButton7.Visible = False
frameScatterdiagram.Image1.Picture = LoadPicture("")
frameScatterdiagram.Show
End Sub
Sub Showhistogram()
frameHistogram.OptionButton1.Value = True
frameHistogram.Show
End Sub
Sub Showbarchart()
framebarchart.OptionButton1.Value = False
framebarchart.OptionButton1.Value = True
    framebarchart.Show
End Sub
Sub ShowLinechart()
frameLinechart.Show
End Sub

Sub ShowCirclechart()
frameCirclechart.Show
End Sub

Sub ShowContourline()
'frameContourline.Show
frm_ContourPlot.Show

End Sub

Sub ShowInterval()
frameInterval.OptionButton1.Value = True
frameInterval.Show
End Sub

Sub ShowBoxchart()
frameBoxchart.OptionButton1.Value = False
frameBoxchart.OptionButton1.Value = True
frameBoxchart.Show
End Sub
Sub ShowParretoChart()
    frameParretochart.OptionButton1.Value = True
    frameParretochart.Show
End Sub
Sub ShowframeReGra()
frameReGra.OptionButton1.Value = True
frameReGra.Show
End Sub
