'Makes an Excel table in a defined format style. Useful for when working with exported data from other systems (MS Access, BusinessObjects, etc)
Sub MakeTable()

Dim sStr As String, sEndPoint As String
Dim loTbl As ListObject

Application.ScreenUpdating = False

With Cells
    .Font.Name = "Calibri"
    .Font.Size = 10
    .Borders.LineStyle = xlNone
    With .Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End With

sStr = ActiveCell.CurrentRegion.Address
sEndPoint = Left$(sStr, InStr(1, sStr, ":", vbTextCompare) - 1)

With ActiveSheet
    .ListObjects.Add(SourceType:=xlSrcRange, Source:=Range(sStr), _
        XlListObjectHasHeaders:=xlYes, TableStyleName:="TableStyleMedium2").Name = "Table1"
End With

Set loTbl = ActiveSheet.ListObjects(1)

With loTbl.HeaderRowRange
    .Interior.Color = RGB(0, 24, 168)
    .WrapText = True
End With

Cells.EntireColumn.AutoFit
Range(sEndPoint).Select

Set loTbl = Nothing
Application.ScreenUpdating = True

End Sub