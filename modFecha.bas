Attribute VB_Name = "modFecha"
Option Explicit

Public Function primerDiaMes(pMes As Integer, pAnho As Integer) As Date

    primerDiaMes = CDate("01/" & pMes & "/" & pAnho)
    
End Function

Public Function primerDiaHorario(ByVal pMes As Integer, ByVal pAnho As Integer, pDiaInicial As Integer) As Date

    If pDiaInicial > 15 Then mesAnterior pMes, pAnho
    
    primerDiaHorario = CDate(pDiaInicial & "/" & pMes & "/" & pAnho & " 00:00:00")
    
End Function

Public Function ultimoDiaMes(ByVal pMes As Integer, ByVal pAnho As Integer)

    mesSiguiente pMes, pAnho
    
    ultimoDiaMes = primerDiaMes(pMes, pAnho) - 1
    
End Function

Public Function ultimoDiaHorario(ByVal pMes As Integer, ByVal pAnho As Integer, pDiaInicial As Integer) As Date

    mesSiguiente pMes, pAnho
    
    ultimoDiaHorario = CDate(primerDiaHorario(pMes, pAnho, pDiaInicial) - 1 & " 23:59:59")
    
End Function

Public Sub mesSiguiente(pMes As Integer, pAnho As Integer)

    pMes = pMes + 1
    
    If pMes = 13 Then
        pMes = 1
        pAnho = pAnho + 1
    End If

End Sub

Public Sub mesAnterior(pMes As Integer, pAnho As Integer)

    pMes = pMes - 1
    
    If pMes = 0 Then
        pMes = 12
        pAnho = pAnho - 1
    End If

End Sub

Public Function diasMes(pMes As Integer, pAnho As Integer)
    
    diasMes = Day(ultimoDiaMes(pMes, pAnho))
    
End Function

Public Sub fillComboMonth(pCombo As ComboBox)
Dim varMeses As Variant

Dim intCiclo As Integer

    varMeses = Array("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre")
    
    pCombo.Clear
    For intCiclo = LBound(varMeses) To UBound(varMeses)
        pCombo.AddItem varMeses(intCiclo)
        pCombo.ItemData(pCombo.NewIndex) = intCiclo + 1
    Next intCiclo
    
    pCombo.ListIndex = Month(Date) - 1
    
End Sub

Public Function horaEntrada(pHoraSalida As Date, pMinutos) As Date

    horaEntrada = minutes2Time(time2Minutes(pHoraSalida) - pMinutos)
    
End Function

Public Function time2Minutes(pHora As Date) As Long

    time2Minutes = Hour(pHora) * 60 + Minute(pHora)
    
End Function

Public Function minutes2Time(pMinutos As Long) As Date
    
    minutes2Time = CDate(Format(pMinutos \ 60, "00") & ":" & Format(pMinutos Mod 60, "00") & ":00")
    
End Function

Public Function minutes2String(pMinutos As Long) As String
    
    minutes2String = Format(pMinutos \ 60, "00") & ":" & Format(pMinutos Mod 60, "00")
    
End Function

Public Function minutesDiff(pDesde As Date, pHasta As Date) As Long

    minutesDiff = CLng((pHasta - pDesde) * 1440)
    
End Function

Public Function timeIntersect(pDesde1 As Date, pHasta1 As Date, pDesde2 As Date, pHasta2 As Date) As Boolean

    timeIntersect = False
    
    If pDesde1 <= pDesde2 And pHasta1 >= pHasta2 Then timeIntersect = True
    
    If pDesde2 <= pDesde1 And pHasta2 >= pHasta1 Then timeIntersect = True
    
    If pDesde1 >= pDesde2 And pDesde1 <= pHasta2 Then timeIntersect = True
    
    If pHasta1 >= pDesde2 And pHasta1 <= pHasta2 Then timeIntersect = True
    
End Function

Public Function YYYYMMDD2DDMMYYYY(pCadena As String) As String

    YYYYMMDD2DDMMYYYY = Mid(pCadena, 7, 2) & Mid(pCadena, 5, 2) & Mid(pCadena, 1, 4)
    
End Function

Private Function diaOtroAnho(pFecha As Date, pOffsetAnhos As Integer) As Date

    diaOtroAnho = CDate(Day(pFecha) & "/" & Month(pFecha) & "/" & Year(pFecha) + pOffsetAnhos)
    
End Function

Public Function diaAnhoSiguiente(pFecha As Date) As Date

    diaAnhoSiguiente = diaOtroAnho(pFecha, 1)
    
End Function

Public Function diaAnhoAnterior(pFecha As Date) As Date

    diaAnhoAnterior = diaOtroAnho(pFecha, -1)
    
End Function
