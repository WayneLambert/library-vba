Option Compare Database
'Purpose:   Count the number of lines of code in your database.
'Author:    Allen Browne (allen@allenbrowne.com)
'Release:   26 November 2007
'Copyright: None. You may use this and modify it for any database you write.
'           All we ask is that you acknowledge the source (leave these comments in your code.)
'Documentation: http://allenbrowne.com/vba-CountLines.html

Private Const micVerboseSummary = 1
Private Const micVerboseListAll = 2

Public Function CountLines(Optional iVerboseLevel As Integer = 3) As Long
On Error GoTo Err_Handler
    'Purpose:   Count the number of lines of code in modules of current database.
    'Requires:  Access 2000 or later.
    'Argument:  This number is a bit field, indicating what should print to the Immediate Window:
    '               0 displays nothing
    '               1 displays a summary for the module type (form, report, stand-alone.)
    '               2 list the lines in each module
    '               3 displays the summary and the list of modules.
    'Notes:     Code will error if dirty (i.e. the project is not compiled and saved.)
    '           Just click Ok if a form/report is assigned to a non-existent printer.
    '           Side effect: all modules behind forms and reports will be closed.
    '           Code window will flash, since modules cannot be opened hidden.
    Dim accObj As AccessObject  'Each module/form/report.
    Dim strDoc As String        'Name of each form/report
    Dim lngObjectCount As Long  'Number of modules/forms/reports
    Dim lngObjectTotal As Long  'Total number of objects.
    Dim lngLineCount As Long    'Number of lines for this object type.
    Dim lngLineTotal As Long    'Total number of lines for all object types.
    Dim bWasOpen As Boolean     'Flag to leave form/report open if it was open.
    
    'Stand-alone modules.
    lngObjectCount = 0&
    lngLineCount = 0&
    For Each accObj In CurrentProject.AllModules
        'OPTIONAL: TO EXCLUDE THE CODE IN THIS MODULE FROM THE COUNT:
        '  a) Uncomment the If ... and End If lines (3 lines later), by removing the single-quote.
        '  b) Replace MODULE_NAME with the name of the module you saved this in (e.g. "Module1")
        '  c) Check that the code compiles after your changes (Compile on Debug menu.)
        'If accObj.Name <> "MODULE_NAME" Then
            lngObjectCount = lngObjectCount + 1&
            lngLineCount = lngLineCount + GetModuleLines(accObj.Name, True, iVerboseLevel)
        'End If

    Next
    lngLineTotal = lngLineTotal + lngLineCount
    lngObjectTotal = lngObjectTotal + lngObjectCount
    If (iVerboseLevel And micVerboseSummary) <> 0 Then
        Debug.Print lngLineCount & " line(s) in " & lngObjectCount & " stand-alone module(s)"
        Debug.Print
    End If
    
    'Modules behind forms.
    lngObjectCount = 0&
    lngLineCount = 0&
    For Each accObj In CurrentProject.AllForms
        strDoc = accObj.Name
        bWasOpen = accObj.IsLoaded
        If Not bWasOpen Then
            DoCmd.OpenForm strDoc, acDesign, WindowMode:=acHidden
        End If
        If Forms(strDoc).HasModule Then
            lngObjectCount = lngObjectCount + 1&
            lngLineCount = lngLineCount + GetModuleLines("Form_" & strDoc, False, iVerboseLevel)
        End If
        If Not bWasOpen Then
            DoCmd.Close acForm, strDoc, acSaveNo
        End If
    Next
    lngLineTotal = lngLineTotal + lngLineCount
    lngObjectTotal = lngObjectTotal + lngObjectCount
    If (iVerboseLevel And micVerboseSummary) <> 0 Then
        Debug.Print lngLineCount & " line(s) in " & lngObjectCount & " module(s) behind forms"
        Debug.Print
    End If
    
    'Modules behind reports.
    lngObjectCount = 0&
    lngLineCount = 0&
    For Each accObj In CurrentProject.AllReports
        strDoc = accObj.Name
        bWasOpen = accObj.IsLoaded
        If Not bWasOpen Then
            'In Access 2000, remove the ", WindowMode:=acHidden" from the next line.
            DoCmd.OpenReport strDoc, acDesign, WindowMode:=acHidden
        End If
        If Reports(strDoc).HasModule Then
            lngObjectCount = lngObjectCount + 1&
            lngLineCount = lngLineCount + GetModuleLines("Report_" & strDoc, False, iVerboseLevel)
        End If
        If Not bWasOpen Then
            DoCmd.Close acReport, strDoc, acSaveNo
        End If
    Next
    lngLineTotal = lngLineTotal + lngLineCount
    lngObjectTotal = lngObjectTotal + lngObjectCount
    If (iVerboseLevel And micVerboseSummary) <> 0 Then
        Debug.Print lngLineCount & " line(s) in " & lngObjectCount & " module(s) behind reports"
        Debug.Print lngLineTotal & " line(s) in " & lngObjectTotal & " module(s)"
    End If
        
    CountLines = lngLineTotal
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
    Case 29068&     'This error actually occurs in GetModuleLines()
        MsgBox "Cannot complete operation." & vbCrLf & "Make sure code is compiled and saved."
    Case Else
        MsgBox "Error " & Err.Number & ": " & Err.Description
    End Select
    Resume Exit_Handler
End Function

Private Function GetModuleLines(strModule As String, bIsStandAlone As Boolean, iVerboseLevel As Integer) As Long
    'Usage:     Called by CountLines().
    'Note:      Do not use error handling: must pass error back to parent routine.
    Dim bWasOpen As Boolean     'Flag applies to standalone modules only.
    
    If bIsStandAlone Then
        bWasOpen = CurrentProject.AllModules(strModule).IsLoaded
    End If
    If Not bWasOpen Then
        DoCmd.OpenModule strModule
    End If
    If (iVerboseLevel And micVerboseListAll) <> 0 Then
        Debug.Print Modules(strModule).CountOfLines, strModule
    End If
    GetModuleLines = Modules(strModule).CountOfLines
    If Not bWasOpen Then
        DoCmd.Close acModule, strModule, acSaveYes
    End If
End Function