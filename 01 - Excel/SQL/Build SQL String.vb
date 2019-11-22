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

'***************************************************************************************************************************

'Builds the SQL string
Function GetSQL_String(ByRef loTbl As ListObject) As String
    GetSQL_String = "SELECT [Genres], [Directors] " & _
                    "FROM [" & loTbl.Parent.Name & "$" & Replace(loTbl.Range.Address, "$", "") & "]" & _
                    "GROUP BY [Genres], [Directors];"
End Function

'***************************************************************************************************************************

'An example of a more complex SQL string used. Used to build a Target Fixed Pay Ranges Tool
Function BuildComplexSQL_String(ByRef loTbl As ListObject) As String

Dim d As tSelectedDimensions, SQL As tSQL_Strings

d.sLocation = cnSTR.Range("hcLocation").Value
d.sCorpTitle = cnSTR.Range("hcCorpTitle").Value
d.sPRF = cnSTR.Range("hcPRF").Value
d.sUnit = cnSTR.Range("hcUnit").Value

'Build SQL string. The clauses are on separate lines for readability - similar to Access SQL
    SQL.sSelect = "SELECT " & sSortedFP_Field & "/1000 AS FixedPayDividedBy1000" & vbNewLine
    SQL.sFrom = "FROM [" & loTbl.Parent.Name & "$" & Replace$(loTbl.Range.Address, "$", "") & "] " & vbNewLine
    'Build where clause
    Select Case d.sUnit
        Case Is = vbNullString
            SQL.sWhere = "WHERE ((([Dimension 1: Location])=" & """" & d.sLocation & """" & ")" & " AND " & vbNewLine & _
                "(([Dimension 2: Corporate Title])=" & """" & d.sCorpTitle & """" & ")" & " AND " & vbNewLine & _
                "(([Dimension 3: PRF])=" & """" & d.sPRF & """" & ")"
        Case Else
             SQL.sWhere = "WHERE ((([Dimension 1: Location])=" & """" & d.sLocation & """" & ")" & " AND " & vbNewLine & _
                "(([Dimension 2: Corporate Title])=" & """" & d.sCorpTitle & """" & ")" & " AND " & vbNewLine & _
                "(([Dimension 3: PRF])=" & """" & d.sPRF & """" & ")" & " AND " & vbNewLine & _
                "(([Dimension 4: Unit])=" & """" & d.sUnit & """" & ")"
    End Select
            SQL.sOrderBy = "ORDER BY " & sSortedFP_Field & "DESC;"

BuildComplexSQL_String = SQL.sSelect & SQL.sFrom & SQL.sWhere & SQL.sOrderBy

End Function