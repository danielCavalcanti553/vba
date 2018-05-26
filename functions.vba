
Function AreaX(x As Double, y As Double) As Double

AreaX = x * y

End Function

Function SomaTempate(rg As Range) As Double
  
    Dim tot As Double
    tot = 0
    For Each x In rg.Cells
        tot = tot + x.Value
    Next
    
    SomaTempate = tot

End Function

Function AreaZ(Optional x As Range) As Double

AreaY = 2

End Function
