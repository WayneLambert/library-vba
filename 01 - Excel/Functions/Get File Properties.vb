'Use any fully qualified complete filepath for sFile including the file extension to pass through to
    'the GetFileInfo function
Sub RetrieveFileProperties()
    Dim sFile as String
    'MS Excel Example
    sFile = Application.ThisWorkbook.FullName
    'MS Access Example
    sFile = Application.CurrentDb.Name
    Call GetFileProperties(sFile)
End Sub

Public Function GetFileProperties(Optional ByVal sFile As String)

On Error GoTo Error_Handler
Dim fso As FileSystemObject
Dim f As File

Set fso = New FileSystemObject
Set f = fso.GetFile(sFile)

Debug.Print , "Name: " & f.Name
Debug.Print , "Size: " & f.Size
Debug.Print , "Created: " & f.DateCreated
Debug.Print , "Modified: " & f.DateLastModified
Debug.Print , "Accessed: " & f.DateLastAccessed
Debug.Print , "Type: " & f.Type
Debug.Print , "Attributes: " & f.Attributes
Debug.Print , "Drive: " & f.Drive
Debug.Print , "Short Name: " & f.ShortName
Debug.Print , "Parent Folder: " & f.ParentFolder
Debug.Print , "Path: " & f.Path
Debug.Print , "Short Path: " & f.ShortPath
Debug.Print , "Extension: " & Trim(Mid$(f.Path, InStrRev(f.Path, "."), Len(f.Path)))

Error_Handler_Exit:
    On Error Resume Next
    Set f = Nothing
    Set fso = Nothing
    Exit Function
 
Error_Handler:
    MsgBox "The following error has occured." & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: GetFileInfo" & vbCrLf & _
           "Error Description: " & Err.Description, _
           vbCritical, "An Error has Occured!"
    Resume Error_Handler_Exit
End Function