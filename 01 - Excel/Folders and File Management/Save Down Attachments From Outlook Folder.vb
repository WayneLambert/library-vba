Option Compare Text
 
Dim RootFolder As String
Dim OlApp As Outlook.Application
Dim olMAPI As Outlook.Namespace
Dim oParentFolder As Outlook.MAPIFolder
Dim SingleFolderRequired As String
Dim RecurseThroughSingleFolder As Boolean
Dim SingleFolderFound As Boolean
 
Public Sub GetOutlookAttachments()

'''''''''''''''''''''''''''''''''''''''''
'Set reference to Outlook object library'
'''''''''''''''''''''''''''''''''''''''''

'RootFolder: your Outlook root folder (mailbox name)

'SingleFolderRequired: set to blank if you want all mail to be retrieved (always recurses through subfolders); _
    or set to the full path of the folder you want to retrieve the mail from _
        (recurses through subfolders depending on the value of RecurseThroughSingleFolder)


 
    RootFolder = "wayne.a.lambert@gmail.com"	'From mailbox
    SingleFolderRequired = "\\wayne.a.lambert@gmail.com\[Google Mail]\Test Folder" Folder with attachments
    RecurseThroughSingleFolder = False	'True to scan all subfolders; False for only actual folder
    
    Set OlApp = CreateObject("Outlook.Application")
    Set olMAPI = GetObject("", "Outlook.application").GetNamespace("MAPI")
    Set oParentFolder = olMAPI.Folders(RootFolder)
    SingleFolderFound = False
    Call ProcessFolder(oParentFolder)
    Set OlApp = Nothing
End Sub
 
Private Sub ProcessFolder(StartFolder As Outlook.MAPIFolder)

Dim uFolder As Outlook.MAPIFolder
    
    If StartFolder.DefaultItemType = 0 Then
      Call ProcessItems(StartFolder, StartFolder.Items)
      For Each uFolder In StartFolder.Folders
        If SingleFolderFound = False Or RecurseThroughSingleFolder = True Then
          Call ProcessFolder(uFolder)
        End If
      Next uFolder
    End If
    
    Set uFolder = Nothing
    
End Sub
 
Private Sub ProcessItems(CurrentFolder As Outlook.MAPIFolder, Collection As Outlook.Items)
    Dim MailObject As Object
    Dim intAttachment As Integer
    Dim SaveAsFilePath As String
    Dim DateTestFrom As Date, DateTestTo As Date
    
    SaveAsFilePath = "C:\Users\Wayne Lambert\Dropbox\VBA\Code Examples\Save Attachments Project\Attachments\"
    DateTestFrom = Range("hcDateTimeFrom")
    DateTestTo = Range("hcDateTimeTo")
    
    If Len(SingleFolderRequired) > 0 Then
      If Left(CurrentFolder.FolderPath, Len(SingleFolderRequired)) = SingleFolderRequired Then
          SingleFolderFound = True
      Else: Exit Sub
      End If
    End If
    
    For Each MailObject In Collection
      DoEvents
      If TypeOf MailObject Is MailItem Then
        If MailObject.SentOn >= DateTestFrom And MailObject.SentOn <= DateTestTo Then 'could filter emails here
          For intAttachment = 1 To MailObject.Attachments.Count
            On Error Resume Next ' trap unknown attachment types
            MailObject.Attachments(intAttachment).SaveAsFile SaveAsFilePath & MailObject.Attachments(intAttachment).Filename 'change path to suit
            On Error GoTo 0
          Next intAttachment
        End If
      End If
    Next MailObject
    
    Set MailObject = Nothing

End Sub