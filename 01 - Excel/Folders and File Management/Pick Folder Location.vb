'Uses the File Dialog box to pick the folder location of the files
Sub GetFilesLocation()

Dim fdFolder As FileDialog
Dim sFilesLocation As String

Set fdFolder = Application.FileDialog(msoFileDialogFolderPicker)

With fdFolder
    .Title = "Please select a folder..."
    .AllowMultiSelect = False
    .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then GoTo NextCode
        sFilesLocation = GetUNC_Path(.SelectedItems(1))
End With

NextCode:
    Sheet1.Range("hcAttachmentPath") = sFilesLocation
    Set fdFolder = Nothing

End Sub

'Gets the universal naming convention (UNC) network path
'There must be network drives available for the this to run
Function GetUNC_Path(sSelectedItem As String) As String

Dim sDrive As String
Dim i As Long

sDrive = UCase$(Left$(sSelectedItem, 2))

With CreateObject("WScript.Network").EnumNetworkDrives
    For i = 0 To .Count - 1 Step 2
        If .Item(i) = sDrive Then
            GetUNC_Path = .Item(i + 1) & Mid$(sSelectedItem, 3)
            Exit For
        End If
    Next i
End With

End Function