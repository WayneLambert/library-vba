'//Source: https://bettersolutions.com/vba/macros/running.htm

Public Sub FromHyperlink
   Call Msgbox("Fired from a hyperlink")    '// Amend as necessary
End Sub

'Select (Insert > Hyperlink) to display the Insert Hyperlink dialog box
'On the left hand side there is a choice of the type of hyperlink to insert, Select "Place in this document".
'Enter the text to display the top in the "text to display" box
'Click on the ScreenTip button and enter the name of the subroutine into this box. In our example our macro is called "FromHyperlink".
'Enter the cell address of the cell address that contains the hyperlink in the "cell address" box. In our example we are using cell "B3".

'Use the Worksheet_FollowHyperlink event...
Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
Dim objRange As Range
Dim objHyperlink As Hyperlink

If Target.Range.Address = "$B$3" Then
    Set objRange = Target.Range
    For Each objHyperlink In objRange.Hyperlinks
        Application.Run (objHyperlink.ScreenTip)
        Exit Sub
    Next objHyperlink
End If

End Sub