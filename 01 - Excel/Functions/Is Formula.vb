'r is the cellthat I want to check
Function IsFormula(ByVal r As Range) As Boolean 
    IsFormula = r.HasFormula 
End Function