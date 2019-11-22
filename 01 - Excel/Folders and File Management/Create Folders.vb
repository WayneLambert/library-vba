'Creates a series of folders within the specified filepath based upon each cell (r) in the range
Sub MakeFolders()

Dim CountryRange As Range, r as Range
Dim Path As String

Path = Range("hcFoldersToSavePathway").Value
'Ensure folder path has a slash at the end
If Right(Path, 1) <> "\" Then Path = Path & "\"

MkDir Path & "CountryFolders\"
'Resets Path string variable to include new CountryFolders folder
Path = Path & "CountryFolders\"

'Sets a range object with all of the countries within the Host Country column of the tblCountries table
Set CountryRange = cnTables.Range("tblCountries[Host Country]")

On Error Resume Next
For Each r in CountryRange
    'Create a new folder with the value currently in the cell (i.e. r range object)
    MkDir Path & r.Value
Next r
On Error GoTo o

Set r = Nothing
Set CountryRange = Nothing

End Sub