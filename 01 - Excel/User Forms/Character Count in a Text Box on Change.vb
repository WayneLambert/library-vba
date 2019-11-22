'Module scope variables...
Dim iChars As Long, iWords As Long
Const iRAT_LEN_REQ As Long = 20
Const iRAT_WORDS_REQ As Long = 5

'Calculates the number of characters typed into the tbRationale textbox each time it is changed
Private Sub tbRationale_Change()

Dim sChars As String

iChars = Len(Me.tbRationale.Value)
sChars = iChars & " of " & iRAT_LEN_REQ & " characters"

Me.lblChars.Caption = sChars
If iChars < iRAT_LEN_REQ Then Me.lblChars.ForeColor = vbRed Else Me.lblChars.ForeColor = vbBlack

End Sub

'**************************************************************************************************

'Calculates the number of words typed into the tbRationale textbox each time it is changed
'Calls upon the function for GetNoOfWords - in my functions library
Private Sub tbRationale_Change()

Dim sWords As String

iWords = GetNoOfWords(Me.tbRationale.Value)
sWords = iWords & " of " & iRAT_WORDS_REQ & " words"

Me.lblWords.Caption = sWords
If iWords < iRAT_WORDS_REQ Then Me.lblWords.ForeColor = vbRed Else Me.lblWords.ForeColor = vbBlack

End Sub