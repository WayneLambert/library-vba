'Calling Procedure to to get circle measurements from function
Sub RetrieveCircleMeasurements()
    
Dim CircleMeasurements As Collection
Dim r As Double
Dim s As Single

r = 10  'For an easy maths example

'Calls GetCircleMeasurements function to get all measurements into a collection
Set CircleMeasurements = GetCircleMeasurements(r)
    
'Loop over the collection. Display in the immediate window
For s = 1 To CircleMeasurements.Count
    Debug.Print CircleMeasurements.item(s)
Next s
   
End Sub

Public Function GetCircleMeasurements(Optional ByVal r As Double) As Collection

Const dPI As Double = 3.141592654

Set GetCircleMeasurements = New Collection

With GetCircleMeasurements
    .Add item:=dPI * r ^ 2, Key:="Area"
    .Add item:=dPI * r * 2, Key:="Circumference"
    .Add item:=r * 2, Key:="Diameter"
End With

End Function