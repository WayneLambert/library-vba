'This procedure was created to amend files within a list of folders for each country
'It was created for the Actuals v Proposals (AvP) process at Deutsche Bank
'The IterateFolders procedure can be used to pass the folder object through to another procedure...
    'to perform other actions on the file (e.g. move, copy, rename, open and amend, etc.)
'As an Access query, a DAO recordset can be used to populate the rCountryRange object

Sub IterateFolders()

Dim rCountryRange As Range, r As Range
Dim FSO As FileSystemObject
Dim sAmendFilePath As String, sCountry as String
Dim dStartTime as Double, dEndTIme as Double

dStartTime = Timer

Set FSO = New FileSystemObject
Set rCountryRange = cnRefTables.Range("tblCountries[Host Country]")

'Loop over each country in the defined range ("CountryRange")
For Each r In rCountryRange
    'Optional Select statement to enable some country folders to be omitted
    Select Case r.Value
        Case "Switzerland", "Austria"
            Debug.Print r.Value
        Case Else
            Debug.Print r.Value
            'Sets filepath of object as a string
            sAmendFilePath = Application.ThisWorkbook.Path & "\Country Folders\" & r.Value
            sCountry = r.Value
        'Instantiates object through the GetFolder method
        'Error handling may need to be put here in case sAmendFilePath is not the name of a valid folder path
        Call AmendFiles(FSO.GetFolder(sAmendFilePath), sCountry)
    End Select
Next r

Set FSO = Nothing

End Sub

Sub AmendFiles(ByRef sAmendFilePath As String, ByVal sCountry as String)

Dim wb as Workbook
Dim rConsRange as Range, rng as Range
Dim sWBFilePath as String
Dim CopyRange as Range
Dim tbl as ListObject
Dim TempArr() as Variant, EndArr as Variant
Dim iNoOfCols as Integer

Application.ScreenUpdating = False

'Set complete full name string of file to be amended
sWBFilePath = sAmendFilePath & "\01. Year End Payment Checks Return Files - " & sCountry & ".xlsb"

'Open the file
Set wb = Open(sWBFilePath)

'Ensure the workbook has opened before moving onto the next lines of code
DoEvents

'** START OF BLOCK: EXAMPLES OF HOW I USED THIS TO MAKE EDITS TO EACH OF THE WORKBOOKS
    'Amend the formula for those cells within the specified table/column. This could be any action you would like to take with the workbook!
    Range("tblControlTotalsCash[No Employees]").Formula = "=COUNTA(tblVCPayroll[Employee ID])"

    'Adds in an additional column to sum the number of falses
    Set tbl = wb.Worksheets("Fixed Pay").ListObjects("tblFixedPay")
    Set rng=wb.Worksheets("Fixed Pay").Range("tblFixedPay[#All]").Resize(tbl.Range.Rows.Count,tbl.Range.Columns.Count + 1)
    tbl.Resize rng
    iNoOfCols=tbl.Range.Columns.Count
    'Names the newly added column
    tblListColumns(iNoOfCols).Name = "No of Falses"
    'Adds a formula to the new column
    Range("tblFixedPay[No of Falses").Formula = "=COUNTIFS(tblFixedPay[@[FP CCY Match]:[Flexible Basket Match]],FALSE"
    'Formats the new column to the "General" format
    Range("tblFixedPay[No Of Falses").NumberFormat = "General"

    'Adds a formula to a control totals table in and formats the background colour
    With Range("tblControlTotalsFP[No Sent for Load]")
        .Formula = "=IF([@[Has Element]]=""Y"",COUNTIFS('FP Load'!$C:$C,[@[Element Code]]),"""")"
        .Interior.Color = RGB(242,242,242)
    End With

    Set tbl = Nothing
    Set rng = Nothing
'** END OF BLOCK: EXAMPLES OF HOW I USED THIS TO MAKE EDITS TO EACH OF THE WORKBOOKS

'Close the workbook and save the changes
wb.Close SaveChanges:=True

'Release memory for next iteration
Set wb = Nothing

End Sub