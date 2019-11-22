    'Forces the workbook to be saved upon initial entry
    MsgBox "Please save a copy of this workbook to a suitable network folder location. The 'Save As' dialogue box will now open for you." & vbNewLine & vbNewLine & "Please do not rename the file.", , "R2R Form"
    
    Dim ProposedFileName As Variant
    Dim TargetFileName As Variant
    Dim FnameInit As String
    Dim LnameInit As String
    
    FnameInit = Left(Application.UserName, 1)
    LnameInit = Mid(Application.UserName, InStr(Application.UserName, " ") + 1, 1)
    
RetrySaveAs:
    
    ProposedFileName = Application.GetSaveAsFilename("R2R Form - " & Format(Date, "yyyy mm dd") & " " & Format(Time, "hh mm ss") & " - " & FnameInit & LnameInit, "Excel Binary Workbook (*.xlsb), *.xlsb", , "Save R2R Form As...")
    TargetFileName = "R2R Form - " & Format(Date, "yyyy mm dd") & " " & Format(Time, "hh mm ss") & " - " & FnameInit & LnameInit
    
    If ProposedFileName = "False" Then
        MsgBox "CANCELLED: Please save the file to a suitable location on the network drive and keep the filename of the report in tact.", vbOKOnly, "Save As Required"
        GoTo RetrySaveAs
    ElseIf ProposedFileName <> TargetFileName Then
        MsgBox "TRY AGAIN: Please do not attempt to change the filename.", vbOKOnly, "Try Again"
        GoTo RetrySaveAs
    Else
        TargetFileName = ProposedFileName
        ActiveWorkbook.SaveAs _
        Filename:=ProposedFileName, _
        FileFormat:=50
    End If