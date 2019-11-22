Option Compare Text

'This sub routine was used to make a master file for each country within a defined range in a table
'It was used within the Actuals versus Proposals (AvP) process: Return Files
'It requires a reference to the "Microsoft Scripting Library"

Sub MakeCountryMasters()

Dim rCountryRange As Range, r As Range
Dim sMasterPath As String, sMasterFilename As String, sDestPath As String
Dim FSO As FileSystemObject

Application.ScreenUpdating = False

sMasterPath = Application.ThisWorkbook.Path & "\"
sMasterFilename = Application.ThisWorkbook.Name

Set FSO = New FileSystemObject
Set rCountryRange = Application.Range("tblCountries[HostCountry]")

For Each r In rCountryRange
    sDestPath = sMasterPath & "CountryFolders\" & r.Value & "\" & _
        Replace(sMasterFilename, "MASTER", "- " & r.Value)
    FSO.CopyFile Source:=sMasterPath & sMasterFilename, Destination:=sDestPath
Next r

Set FSO = Nothing

Application.ScreenUpdating = True

End Sub