'Runs a SQL statement from an Excel dataset stored in an Excel table within a worksheet called cnSource
'Exports the output of the SQL statement into another Excel worksheet called cnDestination
Sub RunSelectSQL()

Dim ADOdb As Object
Dim ADOrs As Object
Dim sSQL As String
Dim i As Long

'Connecting to the ADODB Data Source
Set ADOdb = CreateObject("ADODB.Connection")
With ADOdb
    .Provider = "Microsoft.ACE.OLEDB.12.0"
    .ConnectionString = "Data Source=" & ThisWorkbook.Path & "\" & ThisWorkbook.Name & ";" & _
        "Extended Properties=""Excel 12.0 Xml;HDR=YES"";"
    .Open
End With

'Run the SQL SELECT Query
sSQL = GetSQL_String(cnSource.ListObjects(1))
Set ADOrs = ADOdb.Execute(sSQL)

'Export the headers to the destination sheet
For i = 0 To ADOrs.Fields.Count - 1
    cnDestination.Cells(1, i + 1).Value = ADOrs.Fields(i).Name
Next

'Export the contents of the recordset to the destination sheet
cnDestination.Range("A2").CopyFromRecordset ADOrs

'Clean up
ADOrs.Close
ADOdb.Close
Set ADOdb = Nothing
Set ADOrs = Nothing
    
End Sub

'Builds the SQL string - more complex example in other library file
Function GetSQL_String(ByRef loTbl As ListObject) As String
    GetSQL_String = "SELECT [Genres], [Directors] " & _
                    "FROM [" & loTbl.Parent.Name & "$" & Replace(loTbl.Range.Address, "$", "") & "]" & _
                    "GROUP BY [Genres], [Directors];"
End Function