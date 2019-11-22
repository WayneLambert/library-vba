Function GetNoOfWords(s As String) As Long
'Counts the number of words in a text string

Dim i As Long

'This loop will count the number of spaces within the text string
For i = 1 to Len(s)
    If (Mid$(s,i,1)) = " " Then GetNoOfWords = GetNoOfWords +1
Next i

End Function