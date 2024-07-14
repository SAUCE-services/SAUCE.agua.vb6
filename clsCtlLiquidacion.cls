VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCtlLiquidacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Function updateInteres(prefijoId As Integer, facturaId As Long, db As clsDB) As clsMyAFactura
Dim periodo As New clsRESTPeriodo
Dim factura As New clsMyAFactura
Dim liquidacion As New clsMyAFactura

Dim pagoFacilService As New clsCtlPagoFacil
    
    liquidacion.puntoVta = prefijoId
    liquidacion.nroComprob = facturaId
    liquidacion.findByPrimaryKey db
    
    Set updateInteres = liquidacion
    
    ' Lee Periodo para calcular interes
    periodo.periodoId = liquidacion.periodoId + 1
    periodo.findByPrimaryKey
    
    ' Calcula Interes segundo vencimiento
    liquidacion.interes = interes(liquidacion.total, liquidacion.tasa, periodo.fechaPrimero, periodo.fechaSegundo)
    liquidacion.pfcodigo = pagoFacilService.codigopf(oldFactura2newFactura(liquidacion))
    liquidacion.save db
    
    ' Recalcula intereses deuda
    For Each factura In factura.collectionDeudaByClienteId(liquidacion.clienteId, db, liquidacion.periodoId)
        periodo.periodoId = factura.periodoId
        periodo.findByPrimaryKey
        
        factura.interes = interes(factura.total, factura.tasa, periodo.fechaPrimero, liquidacion.fecha)
        factura.pfcodigo = pagoFacilService.codigopf(oldFactura2newFactura(factura))
        factura.save db
    Next

End Function

Public Sub printLiquidacion(hwnd As Long, prefijoId As Integer, facturaId As Long, db As clsDB, pConsumo As PictureBox, pCrystal As CrystalReport, Optional pDefault As Boolean = False, Optional digital As Boolean = False)
Dim periodo As New clsRESTPeriodo

Dim medidor As New clsMyAMedidor
Dim liquidacion As clsMyAFactura

Dim consumoService As New clsCtlConsumo
Dim impresionService As New clsCtlImpresion

Dim datActual As Date
Dim datAnterior As Date

Dim lngActual As Long
Dim lngAnterior As Long
Dim lngConsumo As Long

    Set liquidacion = updateInteres(prefijoId, facturaId, db)
    
    If digital = False Then
        ' Lee periodo siguiente para proximo vencimiento
        periodo.periodoId = liquidacion.periodoId + 1
        periodo.findByPrimaryKey
        
        medidor.clienteId = liquidacion.clienteId
        medidor.findByClienteId db
        
        lngConsumo = consumoService.datosConsumo(liquidacion.clienteId, liquidacion.periodoId, medidor.medidorID, liquidacion.fecha, datActual, lngActual, datAnterior, lngAnterior, db)
        
        consumoService.drawConsumo pConsumo, liquidacion.clienteId, liquidacion.periodoId, medidor.medidorID, db
    
        impresionService.printReport pCrystal, "rptLiquidacion", db.stringConnection, Array("sDeuda"), Array(Array("pID", liquidacion.autoID), Array("pFProximo", toReportDate(periodo.fechaPrimero)), Array("pMedidor", medidor.medidorID), Array("pFActual", toReportDate(datActual)), Array("pEActual", lngActual), Array("pFAnterior", toReportDate(datAnterior)), Array("pEAnterior", lngAnterior), Array("pConsumo", lngConsumo)), , pDefault
    Else
        printDigital hwnd, liquidacion.puntoVta, liquidacion.nroComprob
    End If
    
End Sub

Public Sub printDigital(hwnd As Long, prefijoId As Integer, facturaId As Long)
Dim request As MSXML2.ServerXMLHTTP

Dim file_stream As ADODB.Stream

Dim url As String
Dim filename As String
    
    filename = Environ$("tmp") & "\liquidacion." & prefijoId & "." & facturaId & ".pdf"
    
    Set request = New MSXML2.ServerXMLHTTP
    
    url = modUrls.url_agua & "liquidacion/generatePdf" & modUrls.join(Array(prefijoId, facturaId))

    request.setTimeouts 200000, 200000, 200000, 200000
    request.Open "GET", url, False
    DoEvents
    request.send
    
    If request.Status = 200 Then
        Set file_stream = New ADODB.Stream
        file_stream.Open
        file_stream.Type = adTypeBinary
        
        file_stream.Write request.responseBody
        file_stream.Position = 0
        
        file_stream.SaveToFile filename, adSaveCreateOverWrite
        file_stream.Close
        
        Set file_stream = Nothing
        
        ShellExecute hwnd, "open", filename, vbNullString, vbNullString, 1
        
    End If
    
    Set request = Nothing


End Sub

Public Function sendLiquidacion(prefijoId As Integer, facturaId As Long, db As clsDB) As String
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    updateInteres prefijoId, facturaId, db
    
    Set request = New MSXML2.ServerXMLHTTP
    
    url = modUrls.url_agua & "liquidacion/sendLiquidacion" & modUrls.join(Array(prefijoId, facturaId))

    request.Open "GET", url, False
    request.send
    
    sendLiquidacion = request.responseText
    
    Set request = Nothing


End Function

Public Sub printCuota(clienteId As Long, planID As Integer, cuotaID As Integer, fecha As Date, db As clsDB, crystal As CrystalReport)
Dim mens As String
Dim categoria As String
Dim medidorID As String

Dim np As Integer

Dim total As Currency

Dim deuda As New clsMyADeuda
Dim cliente As clsMODCliente
Dim medidor As New clsMyAMedidor
Dim operador As New clsMyAOperador
Dim cuota As New clsMyACuota

Dim clienteRep As clsREPCliente

Dim impresionService As New clsCtlImpresion

    deuda.clienteId = clienteId
    deuda.planID = planID
    deuda.findByPrimaryKey db
    
    Set clienteRep = New clsREPCliente
    Set cliente = clienteRep.findLastByClienteId(clienteId)
    Set clienteRep = Nothing
    
    mens = ""
    If Len(Trim(cliente.cuit)) > 0 Then mens = Left(cliente.cuit, 2) & "-" & Mid(cliente.cuit, 3, 8) & "-" & Right(cliente.cuit, 1)
    Select Case cliente.situacionIVA
        Case 1
            mens = mens & " R.I."
        Case 2
            mens = mens & " R.N.I."
        Case 3
            mens = "C.Final"
        Case 4
            mens = mens & " IVA Exento"
        Case 5
            mens = mens & " IVA No Resp."
        Case 6
            mens = mens & " Monotributo"
    End Select
    Select Case cliente.categoria
        Case 1
            categoria = "General"
        Case 2
            categoria = "Especial"
    End Select
    medidorID = ""
    medidor.clienteId = clienteId
    medidor.findLastByClienteId db
    If medidor.autoID > 0 Then medidorID = medidor.medidorID
    
    operador.findLast db
    np = 31
    Do While Mid(operador.razonSocial, np, 1) <> " "
        np = np - 1
    Loop
    
    total = 0
    For Each cuota In cuota.collectionByPlanID(clienteId, planID, db)
        total = total + cuota.importe
    Next
        
'Imprime la cuota
    cuota.clienteId = clienteId
    cuota.planID = planID
    cuota.cuotaID = cuotaID
    cuota.findByPrimaryKey db
    
    Select Case operador.situacionIVA
        Case 1:
            mens = "Resp. Inscripto"
        Case 2:
            mens = "Resp. No Inscripto"
        Case 3:
            mens = "Cons. Final"
        Case 4:
            mens = "Exento"
        Case 5:
            mens = "No Responsable"
        Case 6:
            mens = "Resp. Monotributo"
    End Select
    
    impresionService.printReport crystal, "rptCuota", dbapp.stringConnection, , _
        Array(Array("cliente_id", clienteId), Array("plan_id", planID), Array("cuota_id", cuotaID)), , , _
        Array(Array("nomope", Mid(operador.razonSocial, 1, np)), Array("nomop1", Mid(operador.razonSocial, np + 1)), _
        Array("domope", operador.calle & " " & operador.puerta & " " & operador.piso & " " & operador.dpto & " - " & operador.localidad), _
        Array("locope", "C.P. " & operador.codigoPostal & " " & operador.provincia & " - Tel: " & operador.telefono), _
        Array("opegr1", "C.U.I.T. NRO: " & operador.cuit & " ING. BRUTOS: " & operador.ingresosBrutos), _
        Array("opegr2", "I.V.A. " & mens & "   NRO. E.P.A.S. " & operador.numeroEpas), _
        Array("numcuo", deuda.planID & "-" & Right("00000" & cuota.cuotaID, 5) & "/" & deuda.cuotas), _
        Array("fecemi", fecha), Array("nomcli", cliente.apellidonombre), _
        Array("ubiinm", cliente.inmuebleCalle & " " & cliente.inmueblePuerta & " " & cliente.inmueblePiso & " " & cliente.inmuebleDpto & " " & cliente.inmuebleLocalidad & " (" & cliente.inmuebleCodpostal & ")"), _
        Array("nomcat", ""), _
        Array("domcli", cliente.fiscalCalle & " " & cliente.fiscalPuerta & " " & cliente.fiscalPiso & " " & cliente.fiscalDpto & " " & cliente.fiscalLocalidad & " (" & cliente.fiscalCodpostal & ")"), _
        Array("numsoc", cliente.numeroSocio), Array("cuicli", mens), Array("numcli", clienteId), _
        Array("catego", categoria), Array("nummed", medidorID), _
        Array("fecpri", cuota.fechaVencimiento), Array("impprp", "$ " & Format(cuota.importe, "#,###,##0.00")), _
        Array("implet", "SON PESOS: " & num2letras(cuota.importe * 100)), _
        Array("total", "$ " & Format(total, "#,###,##0.00")), _
        Array("pagado", "$ " & Format(total - deuda.deuda, "#,###,##0.00")), _
        Array("saldo", "$ " & Format(deuda.deuda, "#,###,##0.00")), _
        Array("cuopag", deuda.cuotasPagadas & " de " & deuda.cuotas))

End Sub

Public Function oldFactura2newFactura(oldLiquidacion As clsMyAFactura) As clsMODFactura
Dim newLiquidacion As New clsMODFactura

    newLiquidacion.prefijoId = oldLiquidacion.puntoVta
    newLiquidacion.facturaId = oldLiquidacion.nroComprob
    newLiquidacion.fecha = oldLiquidacion.fecha
    newLiquidacion.clienteId = oldLiquidacion.clienteId
    newLiquidacion.periodoId = oldLiquidacion.periodoId
    newLiquidacion.situacionIVA = oldLiquidacion.situacionIVA
    newLiquidacion.tasa = oldLiquidacion.tasa
    newLiquidacion.descuento = oldLiquidacion.descuento
    newLiquidacion.pagada = oldLiquidacion.pagada
    newLiquidacion.fechapago = oldLiquidacion.fechapago
    newLiquidacion.tipoId = oldLiquidacion.tipoId
    newLiquidacion.anulada = oldLiquidacion.anulada
    newLiquidacion.total = oldLiquidacion.total
    newLiquidacion.interes = oldLiquidacion.interes
    newLiquidacion.letras = oldLiquidacion.letras
    newLiquidacion.prefijoIdinteres = oldLiquidacion.puntoVtaInteres
    newLiquidacion.facturaIdinteres = oldLiquidacion.nroComprobInteres
    newLiquidacion.ivacf = oldLiquidacion.ivacf
    newLiquidacion.ivari = oldLiquidacion.ivari
    newLiquidacion.ivarn = oldLiquidacion.ivarn
    newLiquidacion.periodoIdfin = oldLiquidacion.periodoIdfin
    newLiquidacion.cancelada = oldLiquidacion.cancelada
    newLiquidacion.planIdcancela = oldLiquidacion.planIdcancela
    newLiquidacion.pfcodigo = oldLiquidacion.pfcodigo
    newLiquidacion.pfbarras = oldLiquidacion.pfbarras
    newLiquidacion.uid = oldLiquidacion.uid
    newLiquidacion.uniqueId = oldLiquidacion.autoID

    Set oldFactura2newFactura = newLiquidacion
    
End Function
