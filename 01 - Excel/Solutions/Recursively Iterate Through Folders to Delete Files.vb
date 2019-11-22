'This procedure was created to delete files within a list of folders for each country
'It was created for the Actuals v Proposals (AvP) process at Deutsche Bank
'The IterateFolders procedure can be used to pass the folder object through to another procedure...
    'to perform other actions on the file (e.g. move, copy, rename, open and amend, etc.)
'As an Access query, a DAO recordset can be used to populate the rCountryRange object

Sub IterateFolders()

Dim rCountryRange As Range, r As Range
Dim FSO As FileSystemObject
Dim sDeleteFilePath As String

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
            sDeleteFilePath = Application.ThisWorkbook.Path & "\Country Folders\" & r.Value
        'Instantiates object through the GetFolder method
        'Error handling may need to be put here in case sDeleteFilePath is not the name of a valid folder path
        Call DeleteFiles(FSO.GetFolder(sDeleteFilePath))
    End Select
Next r

Set FSO = Nothing

End Sub

Sub DeleteFiles(ByRef sDeleteFilePath As String)

Dim File As File

For Each File In sDeleteFilePath.Files
    File.Delete
Next File

End Sub