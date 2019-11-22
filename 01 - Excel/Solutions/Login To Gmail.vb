Dim HTMLDoc As HTMLDocument
Dim MyBrowser As InternetExplorer

Sub LoginToGmail()

Dim myHTML_Element As IHTMLElement
Dim myURL As String

On Error GoTo Err_Clear
myURL = "https://accounts.google.com/ServiceLogin?service=mail&passive=true&rm=false&continue=https://mail.google.com/mail/&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1#identifier"
Set MyBrowser = New InternetExplorer

With MyBrowser
    .Silent = True
    .navigate myURL
    .Visible = True
End With

Do
Loop Until MyBrowser.readyState = READYSTATE_COMPLETE

'Do While MyBrowser.Busy = True     'Alternative code to readystate... i.e. do nothing until MyBrowser is no longer busy
'    DoEvents
'Loop

Set HTMLDoc = MyBrowser.document

HTMLDoc.all.Email.Value = "wayne.a.lambert@gmail.com"
For Each myHTML_Element In HTMLDoc.getElementsByTagName("input")
    If myHTML_Element.Type = "submit" Then
        myHTML_Element.Click
        Exit For
    End If
Next

Do
Loop Until MyBrowser.readyState = READYSTATE_COMPLETE

Set HTMLDoc = Nothing

Set HTMLDoc = MyBrowser.document
Set myHTML_Element = Nothing

Application.Wait (Now + TimeValue("0:00:01"))   'Makes VBA wait for 1 second until further execution of the code

HTMLDoc.all.Passwd.Value = "" ' << Enter login password between quotes
For Each myHTML_Element In HTMLDoc.getElementsByTagName("input")
    If myHTML_Element.ID = "signIn" Then
        myHTML_Element.Click
        Exit For
    End If
Next

Err_Clear:
If Err <> 0 Then
    Err.Clear
    Resume Next
End If

Set myHTML_Element = Nothing

End Sub