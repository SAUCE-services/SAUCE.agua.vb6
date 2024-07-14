Attribute VB_Name = "modI2of5"
Option Explicit

Public Function generateI2of5(pCUIT As String, pTCoID As Integer, pPuntoVta As Integer, pCAE As String, pCAEVenc As String) As String
Dim strCadena As String

Dim intVerif As Integer

Dim pagofacil_service As New clsCtlPagoFacil

    strCadena = pCUIT & Trim(Format(pTCoID, "00")) & Trim(Format(pPuntoVta, "0000")) & pCAE & pCAEVenc
    
    intVerif = calculateVerif(strCadena)
    
    strCadena = strCadena & Trim(Str(intVerif))
    
    generateI2of5 = pagofacil_service.codigoI2of5(strCadena)
    
End Function

'Public Function i2of5(ByVal pCadena As String) As String
'Dim strCadena As String
'Dim strCaracter As String
'Dim intCiclo As Integer
'Dim intValor As Integer
'
'    If (Len(pCadena) And 1) = 1 Then Exit Function
'    strCadena = Chr(33)
'    For intCiclo = 1 To Len(pCadena) Step 2
'        intValor = Val(Mid(pCadena, intCiclo, 2))
'        Select Case intValor
'            Case 0 To 3
'                strCaracter = Chr(intValor + 35)
'            Case 4
'                strCaracter = "'"
'            Case 5 To 91
'                strCaracter = Chr(intValor + 35)
'            Case 92
'                strCaracter = Chr(196)
'            Case 93
'                strCaracter = Chr(197)
'            Case 94
'                strCaracter = Chr(199)
'            Case 95
'                strCaracter = Chr(201)
'            Case 96
'                strCaracter = Chr(209)
'            Case 97
'                strCaracter = Chr(214)
'            Case 98
'                strCaracter = Chr(220)
'            Case 99
'                strCaracter = Chr(225)
'        End Select
'        strCadena = strCadena & strCaracter
'    Next
'    strCadena = strCadena & Chr(34)
'    i2of5 = strCadena
'
'End Function

Private Function calculateVerif(pCadena As String) As Integer
Dim intPares As Integer
Dim intImpares As Integer
Dim intCiclo As Integer
Dim intResultado As Integer
Dim intMultiplo As Integer
Dim intIzquierda As Integer
Dim intDerecha As Integer
    
    intPares = 0
    intImpares = 0
    
    For intCiclo = 1 To Len(pCadena)
        If intCiclo Mod 2 = 0 Then
            intPares = intPares + CInt(Mid(pCadena, intCiclo, 1))
        Else
            intImpares = intImpares + CInt(Mid(pCadena, intCiclo, 1))
        End If
    Next intCiclo
    
    intImpares = 3 * intImpares
    
    intResultado = intPares + intImpares
    
    intMultiplo = Int(intResultado / 10) * 10
    
    intIzquierda = intResultado - intMultiplo
    
    intMultiplo = intMultiplo + 10
    
    intDerecha = intMultiplo - intResultado
    
    calculateVerif = IIf(intIzquierda < intDerecha, intIzquierda, intDerecha)

End Function

