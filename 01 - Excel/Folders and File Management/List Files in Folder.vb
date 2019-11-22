'Creates a list of specific files within a folder
'SOURCE: https://www.thespreadsheetguru.com/the-code-vault/vba-code-find-replace-words-specific-file-names-in-folder
Private Function PullFileNames(sPath As String, Optional sFileExt As String = vbNullString) As Collection

Dim MyCollection As New Collection
Dim sFileName As String

'Ensure folder sPath has a slash at the end
If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"

'Get first file that meets extension criteria within folder
sFileName = Dir(sPath & sFileExt)

'Loop through all files in folder
Do While Len(sFileName) > 0
    'Store sFileName into a collection (list)
    MyCollection.Add sPath & sFileName
    
    'Next file meeting extension criteria (if any left)
    sFileName = Dir()
Loop
    
'Output files that meet criteria
Set PullsFileNames = MyCollection

End Function