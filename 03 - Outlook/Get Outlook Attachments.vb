Option Explicit On
Option Compare Text

Dim RootFolder As String
Dim OlApp As outlook.Application, olMAPI As outlook.Namespace, oParentFolder As outlook.MAPIFolder
Dim SingleFolderRequired As String, FolderToCheckName As String
Dim RecurseThroughSingleFolder As Boolean, SingleFolderFound As Boolean

Public Sub GetOutlookAttachments()

    '''''''''''''''''''''''''''''''''''''''''
    'Set reference to Outlook object library'
    '''''''''''''''''''''''''''''''''''''''''

    'RootFolder: your Outlook root folder (mailbox name)

    'SingleFolderRequired: set to blank if you want all mail to be retrieved (always recurses through subfolders);
    'or set to the full path of the folder you want to retrieve the mail from
    '(recurses through subfolders depending on the value of RecurseThroughSingleFolder)

    RootFolder = frmSaveAttachments.tbMailbox                  'From mailbox
    SingleFolderRequired = "\\" & RootFolder & "\[Google Mail]\" & frmSaveAttachments.tbCheckOutlookFolder 'Folder with attachments
    FolderToCheckName = frmSaveAttachments.tbCheckOutlookFolder
    RecurseThroughSingleFolder = False              'True to scan all subfolders; False for only actual folder

Set OlApp = New outlook.Application             'CreateObject("Outlook.Application")
Set olMAPI = OlApp.GetNamespace("MAPI")         'GetObject("", "Outlook.application").GetNamespace ("MAPI")
Set oParentFolder = olMAPI.Folders(RootFolder)
SingleFolderFound = False
    Call ProcessFolder(oParentFolder)
Set OlApp = Nothing

End Sub