Sub CreateTableOfContentsPage()

Dim i As Integer
Sheets.Add Before:=Sheets(1)

For i = 2 To Worksheets.Count
  ActiveSheet.Hyperlinks.Add _
	 Anchor:=Cells(i, 1), _
	 Address:="", _
	 'Add the hyperlinks for each individual worksheet name. To include chart sheets, use the 'Sheets' collection
	 SubAddress:="'" & Worksheets(i).Name & "'!A1", _
	 TextToDisplay:=Worksheets(i).Name
 Next i
 
End Sub