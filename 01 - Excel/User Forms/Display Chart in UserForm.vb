'Generates the chart that sits within a user-form
'This example is based upon a group of charts which is a shape object
Sub GenerateChartInForm()

Dim shpSubForm As Shape
Dim chtObjTmp As ChartObject
Dim sFileName As String

Set shpSubForm = cnChart.Shapes("grpChart")
shpSubForm.CopyPicture

Set chtObjTmp = cnChart.ChartObjects.Add(shpSubForm.Left, shpSubForm.Top, _
    shpSubForm.Width, shpSubForm.Height)

sFileName = ThisWorkbook.Path & "\" & shpSubForm.Name & ".jpg"
    
With chtObjTmp
    .Chart.Paste Type:=xlPasteFormats
    .ShapeRange.Line.Visible = msoFalse
    .Chart.Export Filename:=sFileName, FilterName:="JPG"
    .BottomRightCell.Color = RGB(255, 255, 255) 'or vbWhite
End With

cnSelections.Activate

With frmChart
    .imgChart.Picture = stdole.LoadPicture(sFileName)
    .borderstyle = frmBorderStyleNone
    .Caption = cnChart("hcPeerGroup").Value
    .Show
End With

Kill sFileName
chtObjTmp.Delete
Set chtObjTmp = Nothing

End Sub