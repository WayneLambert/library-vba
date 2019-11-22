'// TWO METHODS

'METHOD 1 (Simple) - Returns TRUE if the file exists
Public Function FileExists(sFile) As Boolean
	If Len(Dir(sFile))<>0 Then FileExists = TRUE Else FileExists = FALSE
End Function

'// 'FileExists()
'This function returns True if there is a file with the name you pass in, even if it is a hidden or system file.
'Assumes the current directory if you do not include a path.

'Returns False if the file name is a folder, unless you pass True for the second argument.
'Returns False for any error, e.g. invalid file name, permission denied, server not found.

'Does not search subdirectories. To enumerate files in subfolders, see List files recursively.

'Examples
'Look for a file named MyFile.mdb in the Data folder:
'    FileExists("C:\Data\MyFile.mdb")
'Look for a folder named System in the Windows folder on C: drive:
'    FolderExists("C:\Windows\System")
'Look for a file named MyFile.txt on a network server:
'    FileExists("\\MyServer\MyPath\MyFile.txt")
'Check for a file or folder name Wotsit on the server:
'    FileExists("\\MyServer\Wotsit", True)
'Check the folder of the current database for a file named GetThis.xls:
'    FileExists(TrailingSlash(CurrentProject.Path) & "GetThis.xls")
'The code

'METHOD 2 (Comprehensive) - Returns TRUE if the file exists
Function FileExists(ByVal sFile As String, Optional bFindFolders As Boolean) As Boolean

'Purpose:   Return True if the file exists, even if it is hidden.
'Arguments: sFile: File name to look for. Current directory searched if no path included.
'           bFindFolders. If sFile is a folder, FileExists() returns False unless this argument is True.
'Note:      Does not look inside subdirectories for the file.
'Author:    Allen Browne. http://allenbrowne.com June, 2006.
'Source:    http://allenbrowne.com/func-11.html

Dim iAttributes As Long

'Include read-only files, hidden files, system files.
iAttributes = (vbReadOnly Or vbHidden Or vbSystem)

If bFindFolders = True Then
	iAttributes = (iAttributes Or vbDirectory) 'Include folders as well.
Else
	'Strip any trailing slash, so Dir does not look inside the folder.
	Do While Right$(sFile, 1) = "\"
		sFile = Left$(sFile, Len(sFile) - 1)
	Loop
End If

'If Dir() returns something, the file exists.
On Error Resume Next
FileExists = (Len(Dir(sFile, iAttributes)) > 0)

End Function