Sub ImportTextFile()

Const ImportFileName as String = <UNC_PATH_TO_FILE>

With DoCmd
    .SetWarnings = False
    .TransferText TransferType:=acImport, Tablename:="<tablename>", _
        Filename:=ImportFileName, HasFieldNames:=True
    .SetWarnings = False
End With

MsgBox "The text file located at " & ImportFileName & " has been successfully imported",
    vbOKOnly, "Import Complete"

End Sub
