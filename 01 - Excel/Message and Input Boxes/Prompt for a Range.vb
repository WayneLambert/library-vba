Sub GetUserRange()

Dim UserRange As Range

Prompt = "Select a range for the random numbers."
Title = "Select a range"

'Display the Input Box
On Error Resume Next
Set UserRange = Application.InputBox( _
	Prompt:=Prompt, Title:=Title, _
		Default:=ActiveCell.Address, Type:=8) 'Range selection
On Error GoTo 0

'Was the Input Box canceled?

If UserRange Is Nothing Then
	MsgBox "Canceled."
Else
	UserRange.Formula = "=RAND()"
End If

End Sub