'Changes the formula for the series for all charts from one to another
Sub ChangeSeriesFormulaAllCharts()

Dim oChart As ChartObject
Dim sOldString As String, sNewString As String
Dim chtSrs As Series

sOldString = InputBox("Enter the string to be replaced:", "Enter old string...")

If Len(sOldString) > 0 Then
    sNewString = InputBox("Enter the string to use instead " & """" & sOldString & """:", "Enter new string...")
    For Each oChart In Sheet1.ChartObjects
        For Each mySrs In oChart.Chart.SeriesCollection
            chtSrs.Formula = WorksheetFunction.Substitute(chtSrs.Formula, sOldString, sNewString)
        Next
    Next
Else
    MsgBox "Nothing to be replaced.", vbInformation, "Nothing Entered"
End If

End Sub