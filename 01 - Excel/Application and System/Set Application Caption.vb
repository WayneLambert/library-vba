'Sets the caption for the whole Excel application - not just the workbook being used
Sub SetTitleBar()

Application.Caption = "Text to Display..."

'To return it back to it's defaults...
Application.Caption = Empty     'or Application.Caption = vbNullString

End Sub