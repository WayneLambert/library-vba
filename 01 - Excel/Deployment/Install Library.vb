Private Sub Workbook_Activate()

     'Macro purpose:  To add a reference to the project using the GUID for the
     'reference library
     
    Dim sGUID As String, theRef As Variant, i As Long
     
     'Update the GUID you need below.
    sGUID = "{00020905-0000-0000-C000-000000000046}"
     
     'Set to continue in case of error
    On Error Resume Next
	     'Remove any missing references
	    For i = ThisWorkbook.VBProject.References.Count To 1 Step -1
	        Set theRef = ThisWorkbook.VBProject.References.Item(i)
	        If theRef.isbroken = True Then
	            ThisWorkbook.VBProject.References.Remove theRef
	        End If
	    Next i
	     
	     'Clear any errors so that error trapping for GUID additions can be evaluated
	    Err.Clear
	     
	     'Add the reference
	    ThisWorkbook.VBProject.References.AddFromGuid _
	    	GUID:=sGUID, Major:=1, Minor:=0
	     
	     'If an error was encountered, inform the user
	    Select Case Err.Number
	    	Case Is = 32813
	         	'Reference already in use.  No action necessary
	    	Case Is = vbNullString
	         	'Reference added without issue
	    	Case Else
	         	'An unknown error was encountered, so alert the user
	        	MsgBox "A problem was encountered trying to" & vbNewLine _
	        		& "add or remove a reference in this file" & vbNewLine & "Please check the " _
	        		& "references in your VBA project!", vbCritical + vbOKOnly, "Error!"
	    End Select
    On Error GoTo 0
End Sub