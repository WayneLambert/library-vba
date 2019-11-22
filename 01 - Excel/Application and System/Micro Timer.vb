Private Declare PtrSafe Function getFrequency Lib "Kernel32" Alias "QueryPerformanceFrequency" _
    (cyFrequency As Currency) As Long
Private Declare PtrSafe Function getTickCount Lib "Kernel32" Alias "QueryPerformanceCounter" _
    (cyTickCount As Currency) As Long

Function MicroTimer() As Double

Dim cyTicks1 As Currency
Static cyFrequency As Currency

MicroTimer = 0
If cyFrequency = 0 Then getFrequency cyFrequency 'Get frequency
getTickCount cyTicks1 'Get ticks
If cyFrequency Then MicroTimer = cyTicks1 / cyFrequency 'Returns Seconds

End Function 