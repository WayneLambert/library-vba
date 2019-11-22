'Get colour reference as Hex code
Function GetHexColour(ByRef rCell As Range) As String

Dim sColor As String

sColor = Right("000000" & Hex(rCell.Interior.Color), 6)
GetHexColour = Right(sColor, 2) & Mid(sColor, 3, 2) & Left(sColor, 2)

End Function

'*****************************************************************************************************

'Retrieves the Hex colour used for VBA userforms using the active cell's fill colour
Sub GetHexColourFromCell
    "&H" & Application.WorksheetFunction.Dec2Hex(ActiveCell.Interior.Color, 6) & "&"
End Sub

'*****************************************************************************************************

Function GetVB_ColourFromRGB(ByVal r As Byte, _
    ByVal g As Byte, ByVal b As Byte) As String

Dim sHex As String

sHex = Left("0" & Hex(r), 2) & _
    Left("0" & Hex(g), 2) & _
    Left("0" & Hex(b), 2)
GetVB_ColourFromRGB = "#" & sHex

End Function

'*****************************************************************************************************

'Get colour reference as RGB value
Function GetRGB_Colour(ByRef rCell As Range) As String

Dim c As Long, r As Long, g As Long, b As Long

c = rCell.Interior.Color
r = c Mod 256
g = c \ 256 Mod 256
b = c \ 65536 Mod 256

GetRGB_Colour = "(" & r & ", " & g & ", " & b & ")"
    
End Function

'*****************************************************************************************************

'Get colour reference as its long number.
'The result from the function can be pasted into the 'BackColor' property of a form or ActiveX Control
Function GetColourAsLong(ByRef rCell As Range, Optional opt As Integer) As Long

Dim iColour As Long, r As Long, g As Long, b As Long

iColour = rCell.Interior.Color
r = iColour Mod 256
g = iColour \ 256 Mod 256
b = iColour \ 65536 Mod 256

If opt = 1 Then
    GetColourAsLong = r
ElseIf opt = 2 Then
    GetColourAsLong = g
ElseIf opt = 3 Then
    GetColourAsLong = b
Else
    GetColourAsLong = iColour
End If
    
End Function

'*****************************************************************************************************
'Example calling procedure to get the long number format of a colour by passing its RGB arguments
Sub GetLongColourCallingProc()
    Dim iColour As Long
    iColour = GetLongColour(100, 100, 100)
End Sub

'Gets colour as its long number from its RGB values
Function GetLongColour(ByVal r As Byte, ByVal g As Byte, ByVal b As Byte) As Long
    GetLongColour = b * 65536 + g * 256 + r
End Function

'*****************************************************************************************************