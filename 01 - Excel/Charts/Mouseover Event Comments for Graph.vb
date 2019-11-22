Private Sub Chart_Activate()
    Application.ShowChartTipNames = False
    Application.ShowChartTipValues = False
End Sub

Private Sub Chart_Deactivate()
    Application.ShowChartTipNames = True
    Application.ShowChartTipValues = True
End Sub

'Uses the mouseover event to select a chart element
'Comments range is one cell on a worksheet
'The comments themselves are concatenated expressions combining values from chart with brief narrative
	'They are offset(1,1) from the 'Comments' named range
Private Sub Chart_MouseMove(ByVal Button As Long, ByVal Shift As Long, _
  ByVal X As Long, ByVal Y As Long)
    Dim ElementId As Long
    Dim arg1 As Long, arg2 As Long
    On Error Resume Next
    ActiveChart.GetChartElement X, Y, ElementId, arg1, arg2
    If ElementId = xlSeries Then
        ActiveChart.Shapes(1).Visible = msoCTrue
        ActiveChart.Shapes(1).TextFrame.Characters.Text = _
          Sheets("Sheet1").Range("Comments").Offset(arg2, arg1)
    Else
        ActiveChart.Shapes(1).Visible = msoFalse
    End If
End Sub