'This procedure builds an object collection of all of the emails that are open in the Outlook inspector as at runtime
'Requires a reference to Microsoft Outlook library
Sub BuildCollectionOfOpenEmails

Dim olApp As outlook.Application
Dim olMailItem As outlook.MailItem
Dim olItem As Object, olInspector As Object, olOpenItem As Object
Dim CollOpenItems As Collection
Dim olTestOpenItem As Variant

Set olApp = outlook.Application

'Builds a collection of open items so the processing can be skipped and there is no interference for the user
Set CollOpenItems = New Collection
    For Each olInspector In olApp.Inspectors
        Set olOpenItem = olInspector.CurrentItem
        CollOpenItems.Add olOpenItem
    Next olInspector

'Destroys the object variables to release memory
Set olOpenItem = Nothing
Set olInspector = Nothing
   
'Evaluates whether the item being tested belongs to the open items collection. If so, skip processing. It can run later once the item is closed.
For Each olTestOpenItem In CollOpenItems
    If olTestOpenItem.SentOn = olItem.SentOn Then
        bOpenItemFound = True
        Exit For
    Else: End If
Next olTestOpenItem