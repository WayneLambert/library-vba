Private Sub Workbook_Open()
    Application.AutoFormatAsYouTypeReplaceHyperlinks = False
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Application.AutoFormatAsYouTypeReplaceHyperlinks = True
End Sub