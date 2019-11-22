#If Win64 Then
    Private Declare PtrSafe Function GetSysColor Lib "user32" _
        (ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetSysColors Lib "user32" _
        (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long
#Else
    Private Declare Function GetSysColor Lib "user32" _
        (ByVal nIndex As Long) As Long
    Private Declare Function SetSysColors Lib "user32" _
        (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long
#End If

Private Const COLOR_WINDOWTEXT As Long = 8
Private Const CHANGE_INDEX As Long = 1

Public Sub MsgBoxColorDemo()
   Dim defaultColour As Long

   'Store the default system color
   defaultColour = GetSysColor(COLOR_WINDOWTEXT)

   'Set system color to red
   SetSysColors CHANGE_INDEX, COLOR_WINDOWTEXT, vbRed
   MsgBox "Incorrect", vbCritical, "Your result is..."

   'Set system color to green
   SetSysColors CHANGE_INDEX, COLOR_WINDOWTEXT, RGB(0, 128, 0)
   MsgBox "Correct", , "Your result is..."
   
   'Restore default value
   SetSysColors CHANGE_INDEX, COLOR_WINDOWTEXT, defaultColour

End Sub