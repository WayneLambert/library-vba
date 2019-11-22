Sub ModifyWorksheetCodeName()

With ActiveSheet
    .Name = "Report"
    ThisWorkbook.VBProject.VBComponents(.CodeName).Properties("_CodeName") = "cnReport"
End With

End Sub