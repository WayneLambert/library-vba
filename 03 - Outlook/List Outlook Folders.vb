Option Explicit
'Add Microsoft Outlook object library as an available reference
Public Sub GetListOfFolders()

Dim olSession As Outlook.Namespace
Dim olFolders As Outlook.Folders
Dim olFolder As Outlook.Folder
Dim bRetValue As Boolean
Dim sReport As String
Dim iReply As Long

On Error GoTo On_Error
Set olSession = Outlook.Application.Session

Set olFolders = olSession.Folders
For Each olFolder In olFolders
    Call RecurseFolders(olFolder, vbTab, sReport)
    sReport = sReport & "---------------------------------------------------------------------------" & vbNewLine
Next

bRetValue = CreateReportAsEmail("List of Folders", sReport)
    
Exit_Handler:
    Set olSession = Nothing
    Exit Sub

On_Error:
    MsgBox "Error=" & Err.Number & " " & Err.Description
    Resume Exit_Handler

End Sub

Private Sub RecurseFolders(CurrentFolder As Outlook.Folder, Tabs, sReport As String)

Dim olTable As Outlook.Table
Dim olRow As Outlook.Row
Dim RowValues() As Variant
Dim olSubFolders As Outlook.Folders
Dim olSubFolder As Outlook.Folder

sReport = sReport & Tabs & "Folder Name: " & CurrentFolder.Name & " (Store: " & CurrentFolder.Store.DisplayName & ")" & vbNewLine

Set olSubFolders = CurrentFolder.Folders
For Each olSubFolder In olSubFolders
    Call RecurseFolders(olSubFolder, Tabs & vbTab, sReport)
Next olSubFolder

End Sub

' VBA Function which displays a report inside an email
Public Function CreateReportAsEmail(sTitle As String, sReport As String)

Dim olSession As Outlook.Namespace
Dim olMailItem As MailItem
Dim MyAddress As AddressEntry
Dim olInbox As Outlook.Folder

On Error GoTo On_Error
CreateReportAsEmail = True

Set olSession = Outlook.Application.Session
Set olInbox = olSession.GetDefaultFolder(olFolderInbox)
Set olMailItem = olInbox.Items.Add("IPM.Mail")

Set MyAddress = olSession.CurrentUser.AddressEntry
With olMailItem
    .Recipients.Add (MyAddress.Address)
    .Recipients.ResolveAll
    .Subject = sTitle
    .Body = sReport
    .Save
    .Display
End With
    
Exit_Handler:
    Set olSession = Nothing
    Exit Function

On_Error:
    CreateReportAsEmail = False
    MsgBox "Error=" & Err.Number & " " & Err.Description
    Resume Exit_Handler

End Function