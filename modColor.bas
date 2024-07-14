Attribute VB_Name = "modColor"
Option Explicit

Private Const cntRojo = &HFF 'Rojo
Private Const cntAmarillo = &H80FFFF 'Amarillo
Private Const cntNaranja = &H80C0FF 'Naranja
Private Const cntVioleta = &HFFC0FF 'Violeta
Private Const cntAzul = &HFF0000 'Azul
Private Const cntVerde = &HFF00 'Verde
Private Const cntBlanco = &HFFFFFF ' Blanco
Private Const cntCeleste = &HFFFF00
Private Const cntMorado = &HC277FF
Private Const cntRosado = &HC0C0FF
Private Const cntVerdeOscuro = &H8000
Private Const cntGris = &H808080

Public Function colores(pOffset) As Long

    Select Case pOffset
        Case 1
            colores = colorAzul
        Case 2
            colores = colorGris
        Case 3
            colores = colorCeleste
        Case 4
            colores = colorMorado
        Case 5
            colores = colorNaranja
        Case 6
            colores = colorRosado
        Case 7
            colores = colorVioleta
        Case 8
            colores = colorVerde
        Case 9
            colores = colorRojo
        Case 10
            colores = colorVerdeOscuro
        Case 11
            colores = colorAmarillo
    End Select
    
End Function

Public Function variant2RGB(pColor As Variant) As Long
Dim objRGB As New clsRGB

    With objRGB
        .color = pColor
        variant2RGB = RGB(objRGB.red, objRGB.green, objRGB.blue)
    End With
    
End Function

Public Function colorRojo() As Long

    colorRojo = variant2RGB(cntRojo)
    
End Function

Public Function colorAmarillo() As Long

    colorAmarillo = variant2RGB(cntAmarillo)
    
End Function

Public Function colorNaranja() As Long

    colorNaranja = variant2RGB(cntNaranja)
    
End Function

Public Function colorVioleta() As Long

    colorVioleta = variant2RGB(cntVioleta)
    
End Function

Public Function colorAzul() As Long

    colorAzul = variant2RGB(cntAzul)
    
End Function

Public Function colorVerde() As Long

    colorVerde = variant2RGB(cntVerde)
    
End Function

Public Function colorBlanco() As Long

    colorBlanco = variant2RGB(cntBlanco)
    
End Function

Public Function colorCeleste() As Long

    colorCeleste = variant2RGB(cntCeleste)
    
End Function

Public Function colorMorado() As Long

    colorMorado = variant2RGB(cntMorado)
    
End Function

Public Function colorRosado() As Long

    colorRosado = variant2RGB(cntRosado)
    
End Function

Public Function colorVerdeOscuro() As Long

    colorVerdeOscuro = variant2RGB(cntVerdeOscuro)
    
End Function

Public Function colorGris() As Long

    colorGris = variant2RGB(cntGris)
    
End Function


