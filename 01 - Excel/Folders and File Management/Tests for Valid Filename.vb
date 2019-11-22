'Determines if a Given Excel File Name Is Valid
Function ValidFileName(FileName As String) As Boolean

Dim wb As Workbook

'Create a Temporary .xlsb file
On Error GoTo InvalidFileName
Set wb = Workbooks.Add
wb.SaveAs Environ("TEMP") & "\" & FileName & ".xlsb", xlExcel12
On Error Resume Next

'Close Temporary Workbook file
 wb.Close (False)

'Delete Temporary File
 Kill Environ("TEMP") & "\" & FileName & ".xlsb"

'File Name is Valid!
 ValidFileName = True

Exit Function

'*** ERROR HANDLERS ***
InvalidFileName:
'Close Temporary Workbook file
wb.Close (False)

'File Name is Invalid
ValidFileName = False

End Function