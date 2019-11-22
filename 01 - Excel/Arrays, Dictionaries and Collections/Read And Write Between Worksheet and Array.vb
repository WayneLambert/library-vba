Sub ReadAndWriteBetweenWorksheetAndArray()

Dim x As Variant
Dim r As Long, c As Integer

'Read the data into the variant
x = Range("Data").Value

'Loop through the variant array
'r outer loop iterates through rows
For r = 1 To UBound(x, 1)
    'c inner loop iterates through columns
    For c = 1 To UBound(x, 2)
'       Multiply by 2
        x(r, c) = x(r, c) * 2
    Next c
Next r

'Write the data within the variant array back to the worksheet range named "Data"
Range("Data") = x

End Sub