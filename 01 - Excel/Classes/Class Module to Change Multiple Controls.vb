 '<Insert into a Class Module>
 Public WithEvents CommandButtonGroup As CommandButton 
 
Private Sub CommandButtonGroup_Click() 
     
MsgBox "You pressed " & CommandButtonGroup.Caption 
     
End Sub 
 
Private Sub CommandButtonGroup_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single) 
     
Dim Ctrl As Control 
 
For Each Ctrl In UserForm1.Controls 
	If TypeName(Ctrl) = "CommandButton" And Ctrl.BackColor = vbRed Then 
		Ctrl.BackColor = UserForm1.BackColor 
	End If 
Next 
CommandButtonGroup.BackColor = vbRed 
UserForm1.Caption = CommandButtonGroup.Caption 
     
End Sub 

 '<Insert into a UserForm>
 
Option Explicit 
 
Dim Buttons() As New CommandButtonClass 
 
Private Sub UserForm_Initialize() 
     
Dim Ctrl As Control 
Dim Count As Integer 
 
For Each Ctrl In UserForm1.Controls 
	If TypeName(Ctrl) = "CommandButton" Then 
		Count = Count + 1 
		ReDim Preserve Buttons(1 To Count) 
		Set Buttons(Count).CommandButtonGroup = Ctrl 
	End If 
Next 
     
End Sub