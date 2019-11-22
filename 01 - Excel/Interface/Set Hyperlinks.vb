Sub SetHyperlinks()

'Adds a hyperlink. The full address for SubAddress can be retrieved using the .vb file for "Get Range Full Address"
cnSheetName.Range("A1").Hyperlinks.Add Anchor:=Range("A1"), Address:="", SubAddress:="'Sheet 2'!$A$1", _
    ScreenTip:="This is the hover text", TextToDisplay:="This is the text visible from within the cell."

'Removes a hyperlink
cnSheetName.Range("A1").Hyperlinks.Delete

End Sub