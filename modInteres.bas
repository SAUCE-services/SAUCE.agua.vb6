Attribute VB_Name = "modInteres"
Option Explicit

Public Function interes(total As Currency, tasa As Double, referencia As Date, fechacalculo As Date) As Currency
Dim tasadiaria As Double
Dim factor As Double
Dim calculo_interes As Double
    
    tasadiaria = (1 + tasa) ^ (1 / 30) - 1
    factor = (1 + tasadiaria) ^ (fechacalculo - referencia) - 1
    calculo_interes = total * factor
    
    interes = calculo_interes

End Function

