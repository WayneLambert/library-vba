'Tests each element of TempArr to see if it exists in sString. Where it does, it replaces it with a null string
'The below example specifically can be used to ensure that files are not saved with disallowed characters
Sub ReplaceMultipleStrings()  
 
Dim TempArr as variant, c as variant
Dim sString as String

'The string variable sString would represent a filename that is being cleaned up before it is being saved
sString= "This : is \ a / test ? string. to do * multiple replaces on="
TempArr = Array("/","\",":","=","*",".","?")
For Each c in TempArr
    sString = Replace(sString,c,"")
Next c

End Sub