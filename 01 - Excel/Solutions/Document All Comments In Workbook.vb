'Documents all of the comments within a workbook
Sub DocumentAllComments()

Dim ws as Worksheet
Dim cmt as Comment
Dim w as Byte

Set ws = Worksheets.Add

With ws
    .Cells(1,1).Value = "Comment"
    .Cells(1,2).Value = "Address"
    .Cells(1,3).Value = "Author"
	r=2
	For w =1 To Worksheets.Count
    	For Each cmt in Worksheets(w).Comments
        	.Cells(r,1).Value = cmt.Text
        	.Cells(r,2).Value = Worksheets(w).Name & "!" & cmt.Parent.Address
        	.Cells(r,3).Value = cmt.Author
        	R = R +1
    	Next cmt
	.Columns.AutoFit
	Next w
End With

End Sub