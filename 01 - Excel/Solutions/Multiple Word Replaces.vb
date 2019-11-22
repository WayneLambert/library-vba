'Tests each element of TempArr to see if it exists in TestString. Where it does, it replaces it will a null string
'The below example specifically can be used to ensure that files are not saved with disallowed characters
Sub MultipleWordReplaces()
 
Dim TempArr as variant, c as variant
Dim TestString as String

TestString= "This : is \ a / test ? string. to do * multiple replaces on="
TempArr = Array("/","\",":","=","*",".","?")
For Each c in TempArr
    TestString = Replace(TestString,c,"")
Next c

End Sub