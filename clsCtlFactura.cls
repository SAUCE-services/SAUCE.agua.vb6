VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCtlFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function detalles2detartic(factura As clsMyAFactura, db As clsDB, fefactura As clsMyAFEFactura) As Collection
Dim detalle As New clsMyADetalle
Dim periodo As New clsRESTPeriodo
Dim rubro As New clsMyARubro
Dim cliente As clsMODCliente
Dim tipocomprobante As New clsMyATipoComprobante
Dim ncredito As New clsMyANCredito
Dim fedetalle As clsMyAFEDetalle
Dim alicuota As New clsMyAAlicuota

Dim clienteRep As New clsREPCliente

Dim IVA As Boolean

Dim interescomprobante As Currency
Dim neto27comprobante As Currency
Dim netocomprobante As Currency
Dim iva27comprobante As Currency
Dim ivacomprobante As Currency
Dim exentocomprobante As Currency
Dim netoncredito As Currency
Dim ivancredito As Currency
Dim exentoncredito As Currency
Dim alicuotacomprobante As Currency

Dim item As Integer

Dim detalles As New Collection

    neto27comprobante = 0
    netocomprobante = 0
    iva27comprobante = 0
    ivacomprobante = 0
    exentocomprobante = 0
    IVA = False

    alicuota.findLast db
    
    Set cliente = clienteRep.findLastByClienteId(factura.clienteId)
    
    Select Case cliente.situacionIVA
        Case 1, 2, 6
            alicuotacomprobante = alicuota.IVA
        Case Else
            alicuotacomprobante = alicuota.ivacf
    End Select
    

    Set fefactura = New clsMyAFEFactura
    fefactura.clienteId = factura.clienteId
    fefactura.tipoId = 1
    If cliente.situacionIVA = 1 Then fefactura.tipoId = 2
    
    tipocomprobante.tipoId = fefactura.tipoId
    tipocomprobante.findByPrimaryKey db
    
    fefactura.prefijo = tipocomprobante.puntoVenta
    fefactura.fecha = Date
    fefactura.tipoCompro = tipocomprobante.tipocomprobante

    periodo.periodoId = factura.periodoId
    periodo.findByPrimaryKey
    
    item = 0

    ' Cálculo adicional por pago posterior al primer vencimiento
    If factura.fechapago > periodo.fechaPrimero Then
        IVA = True
        interescomprobante = modInteres.interes(factura.total, factura.tasa, periodo.fechaPrimero, periodo.fechaSegundo)
        If IVA Then interescomprobante = interescomprobante / (1 + alicuotacomprobante)
        Set fedetalle = New clsMyAFEDetalle
        item = item + 1
        fedetalle.item = item
        fedetalle.rubroID = 0
        fedetalle.concepto = "Recargo 2do Vencimiento"
        fedetalle.fecha = factura.fechapago
        fedetalle.cantidad = 1
        fedetalle.unitarioSinIva = interescomprobante
        fedetalle.unitarioConIva = interescomprobante
        If IVA Then fedetalle.unitarioConIva = fedetalle.unitarioSinIva * (1 + alicuotacomprobante)
        
        If alicuotacomprobante = 0.21 Then
            netocomprobante = netocomprobante + fedetalle.cantidad * fedetalle.unitarioSinIva
            ivacomprobante = ivacomprobante + fedetalle.cantidad * (fedetalle.unitarioConIva - fedetalle.unitarioSinIva)
        ElseIf alicuotacomprobante = 0.27 Then
            neto27comprobante = neto27comprobante + fedetalle.cantidad * fedetalle.unitarioSinIva
            iva27comprobante = iva27comprobante + fedetalle.cantidad * (fedetalle.unitarioConIva - fedetalle.unitarioSinIva)
        Else
            exentocomprobante = exentocomprobante + fedetalle.cantidad * fedetalle.unitarioSinIva
        End If
    
        detalles.add fedetalle, "k." & item
    End If
    
    For Each detalle In detalle.collectionByLiquidacion(factura.puntoVta, factura.nroComprob, db)
        IVA = False
        If detalle.IVA <> 0 Then IVA = True
        If detalle.rubroID = 0 Then
            rubro.concepto = "Intereses por mora de Facturas pagadas fuera de término"
        Else
            rubro.rubroID = detalle.rubroID
            rubro.findByFecha factura.fecha, db
        End If
        Set fedetalle = New clsMyAFEDetalle
        item = item + 1
        fedetalle.item = item
        fedetalle.rubroID = detalle.rubroID
        fedetalle.fecha = factura.fechapago
        fedetalle.cantidad = detalle.cantidad
        fedetalle.unitarioSinIva = detalle.precioUnitario
        fedetalle.unitarioConIva = detalle.precioUnitario
        If IVA Then fedetalle.unitarioConIva = fedetalle.unitarioSinIva * (1 + alicuotacomprobante)
        fedetalle.concepto = rubro.concepto
        
        If IVA Then
            If alicuotacomprobante = 0.21 Then
                netocomprobante = netocomprobante + fedetalle.cantidad * fedetalle.unitarioSinIva
                ivacomprobante = ivacomprobante + fedetalle.cantidad * (fedetalle.unitarioConIva - fedetalle.unitarioSinIva)
            ElseIf alicuotacomprobante = 0.27 Then
                neto27comprobante = neto27comprobante + fedetalle.cantidad * fedetalle.unitarioSinIva
                iva27comprobante = iva27comprobante + fedetalle.cantidad * (fedetalle.unitarioConIva - fedetalle.unitarioSinIva)
            End If
        End If
        If Not IVA Then exentocomprobante = exentocomprobante + fedetalle.cantidad * fedetalle.unitarioSinIva
        
        detalles.add fedetalle, "k." & item
    Next
    
    For Each ncredito In ncredito.collectionByLiquidacion(factura.puntoVta, factura.nroComprob, db)
        netoncredito = 0
        ivancredito = 0
        exentoncredito = 0
        
        ivancredito = -(ncredito.ivacf + ncredito.ivari + ncredito.ivarn)
        netoncredito = ivancredito / alicuotacomprobante
        exentoncredito = -(ncredito.total + ivancredito + netoncredito)
        
        rubro.rubroID = 0
        rubro.concepto = "Descuentos realizados"
        Set fedetalle = New clsMyAFEDetalle
        item = item + 1
        fedetalle.item = item
        fedetalle.rubroID = 0
        fedetalle.fecha = factura.fechapago
        fedetalle.cantidad = 1
        fedetalle.unitarioSinIva = netoncredito
        fedetalle.unitarioConIva = -ncredito.total
        fedetalle.concepto = rubro.concepto
        
        If alicuotacomprobante = 0.27 Then
            neto27comprobante = neto27comprobante + netoncredito
            iva27comprobante = iva27comprobante + ivancredito
        End If
        If alicuotacomprobante = 0.21 Then
            netocomprobante = netocomprobante + netoncredito
            ivacomprobante = ivacomprobante + ivancredito
        End If
        exentocomprobante = exentocomprobante + exentoncredito
        
        detalles.add fedetalle, "k." & item
    Next
    
    fefactura.importe = netocomprobante + neto27comprobante + ivacomprobante + iva27comprobante + exentocomprobante
    fefactura.neto27 = neto27comprobante
    fefactura.neto = netocomprobante
    fefactura.iva27 = iva27comprobante
    fefactura.IVA = ivacomprobante
    fefactura.exento = exentocomprobante
    fefactura.letras = modConv.num2letras(fefactura.importe)
    
    Set detalles2detartic = detalles

End Function

Public Function makeFactura(puntoVta As Integer, nroComprob As Long, db As clsDB) As String
Dim factura As New clsMyAFactura
Dim fedetalle As New clsMyAFEDetalle
Dim fefactura As clsMyAFEFactura
Dim parametro As New clsMyAParametro

Dim detalles As Collection

Dim strBarras As String
Dim strCAEVenc As String
Dim strCAE As String

Dim lngNroComprob As Long

    parametro.findLast db

    factura.puntoVta = puntoVta
    factura.nroComprob = nroComprob
    factura.findByPrimaryKey db
    
    Set detalles = Me.detalles2detartic(factura, db, fefactura)
    
    strCAE = modFEv1.cae(fefactura.tipoId, fefactura.clienteId, fefactura.importe, fefactura.exento, fefactura.neto27, fefactura.neto, 0, fefactura.iva27, fefactura.IVA, 0, parametro.feProduccion, db, lngNroComprob, strCAEVenc, strBarras)
    
    If strCAE <> "" Then
        With fefactura
            .numero = lngNroComprob
            .cae = strCAE
            .caeVencimiento = strCAEVenc
            .caeBarras = strBarras
            .puntoVta = factura.puntoVta
            .nroComprob = factura.nroComprob
            
            .save db
        End With
        
        For Each fedetalle In detalles
            fedetalle.tipoId = fefactura.tipoId
            fedetalle.prefijo = fefactura.prefijo
            fedetalle.numero = fefactura.numero
            fedetalle.save db
        Next
    End If
    
    makeFactura = strCAE
    
End Function

Public Function makeLiquidacion(fechaEmision As Date, cliente As clsMODCliente, alicuota As clsMyAAlicuota, operador As clsMyAOperador, periodo As clsRESTPeriodo, db As clsDB, Optional Status As StatusBar) As Boolean
Dim operadorServicio As Integer
Dim operadorPeriodoFacturado As Integer
Dim prefijoId As Integer

Dim operadorInicio As Date
Dim fechaInicio As Date
Dim fechaFin As Date

Dim facturaId As Long

Dim tasaMayor As Currency
Dim tasaMenor As Currency
Dim factorCobro As Currency
Dim subTotal As Currency
Dim ivct As Currency
Dim ivat As Currency
Dim rnit As Currency
Dim importe As Currency
Dim metrosConsumo As Currency
Dim cantidad As Currency
Dim totalInteresesMora As Currency

Dim tasa As Double

Dim facturar As Boolean
Dim usuarioDesconectado As Boolean
Dim listar As Boolean

Dim fechaActual As Variant
Dim fechaAnterior As Variant

Dim estadoActual As Currency
Dim estadoAnterior As Currency
Dim consumoRegistrado As Long
Dim lrub As Integer
Dim rubroID As Integer
Dim fila As Integer

Dim servicio As String
Dim segmento As String
Dim concepto As String

Dim tsub As Currency
Dim tiva As Currency
Dim tcfl As Currency
Dim trni As Currency
Dim tpri As Currency
Dim tseg As Currency
Dim total As Currency

Dim rfac(1 To 50, 0 To 5) As String

Dim factura As New clsMyAFactura
Dim suspfactura As New clsMyASuspFactura
Dim medidor As New clsMyAMedidor
Dim desconexion As New clsMyADesconexion
Dim lectura As New clsMyALectura
Dim rubro As New clsMyARubro
Dim novedad As New clsMyANovedad
Dim rango As New clsMyARango
Dim periodolocal As New clsRESTPeriodo
Dim detalle As New clsMyADetalle
Dim ncredito As New clsMyANCredito

Dim clienteDato As clsMODClienteDato

Dim clienteDatoRep As New clsREPClienteDato

Dim pagoFacilService As New clsCtlPagoFacil
Dim liquidacionService As New clsCtlLiquidacion

    usuarioDesconectado = False
    facturar = False
    
    operadorServicio = operador.servicio
    operadorPeriodoFacturado = operador.periodoFactura
    operadorInicio = operador.fechaInicio
    prefijoId = operador.puntoVta
    facturaId = operador.nroComprob
    
    factura.puntoVta = prefijoId
    factura.findLast db
    If facturaId < factura.nroComprob + 1 Then facturaId = factura.nroComprob + 1
    
    tasaMayor = 75
    tasaMenor = 25
    If fechaEmision - operadorInicio > 365 Then
        tasaMayor = 50
        tasaMenor = 50
    End If
    If fechaEmision - operadorInicio > 730 Then
        tasaMayor = 25
        tasaMenor = 75
    End If
    If fechaEmision - operadorInicio > 1095 Then
        tasaMayor = 0
        tasaMenor = 100
    End If

    suspfactura.clienteId = cliente.clienteId
    suspfactura.periodoIDInicio = periodo.periodoId
    suspfactura.findLast db
    If suspfactura.autoID > 0 Then If IsNull(suspfactura.periodoIdfin) Or suspfactura.periodoIdfin >= periodo.periodoId Then Exit Function
    
    subTotal = 0
    ivct = 0
    ivat = 0
    rnit = 0
    
    facturar = False
    factura.clienteId = cliente.clienteId
    factura.periodoId = periodo.periodoId
    factura.findByClientePeriodo db
    
    If factura.autoID = 0 Then
        factura.clienteId = cliente.clienteId
        factura.periodoId = periodo.periodoId
        factura.findByClientePeriodoPrev db
        
        If factura.autoID = 0 Then
            facturar = True
        Else
            If factura.periodoIdfin = 0 Or IsNull(factura.periodoIdfin) Or factura.periodoIdfin < periodo.periodoId Then
                facturar = True
            Else
                Exit Function
            End If
        End If
    Else
        If factura.anulada <> 0 Then facturar = True
    End If
    
    If Not facturar Then Exit Function
    
    fechaInicio = periodo.fechaInicio
    fechaFin = periodo.fechaFin
    
    medidor.clienteId = cliente.clienteId
    medidor.findByClienteId db
    If medidor.autoID = 0 Then If cliente.cobro < 3 Then Exit Function
    
    desconexion.clienteId = cliente.clienteId
    desconexion.fechaDesconexion = fechaEmision
    desconexion.findByPrimaryKey db
    
    If desconexion.autoID > 0 Then
        If IsNull(desconexion.fechaReconexion) Then
            facturar = True
            usuarioDesconectado = True
        End If
        If desconexion.fechaReconexion > fechaFin Then
            If desconexion.fechaDesconexion <= fechaInicio Then
                facturar = True
                usuarioDesconectado = True
            End If
        End If
    End If
    
    Status.SimpleText = "Conexión : " & Right("000000" & cliente.clienteId, 6) & "    Factura : " & Right("0000" & prefijoId, 4) & "-" & Right("00000000" & facturaId, 8)
    Status.Refresh
    
    factorCobro = 1
    If cliente.servicio < operadorServicio Then factorCobro = 0.5
    
    If Not usuarioDesconectado Then
        lectura.medidorID = medidor.medidorID
        lectura.periodoId = periodo.periodoId
        lectura.findByPrimaryKey db
        
        If lectura.autoID = 0 Then
            fechaActual = Null
            estadoActual = medidor.estadoInicio
        Else
            fechaActual = lectura.fechaLectura
            estadoActual = medidor.estadoInicio
            If medidor.fechaColocacion <= periodo.fechaInicio Then estadoActual = lectura.estado
        End If
        
        lectura.medidorID = medidor.medidorID
        lectura.periodoId = periodo.periodoId
        lectura.findByMedidorIDPrev db
        If lectura.autoID = 0 Then
            fechaAnterior = Null
            estadoAnterior = 0
            If medidor.autoID > 0 Then estadoAnterior = medidor.estadoInicio
        Else
            periodolocal.periodoId = lectura.periodoId
            periodolocal.findByPrimaryKey
            fechaAnterior = lectura.fechaLectura
            estadoAnterior = medidor.estadoInicio
            If medidor.fechaColocacion <= periodolocal.fechaFin Then estadoAnterior = lectura.estado
        End If
        
        consumoRegistrado = estadoActual - estadoAnterior
    Else
        fechaActual = Null
        fechaAnterior = Null
        estadoActual = 0
        estadoAnterior = 0
        consumoRegistrado = 0
    End If
    
    lrub = 0
    rubroID = 0
    For Each rubro In rubro.collectionSinRepeticion(db)
        If rubroID < rubro.rubroID Then
            rubroID = rubro.rubroID
            rubro.findLast db
            listar = False
            If usuarioDesconectado Then
                If cliente.cobro >= rubro.cobro And rubro.desconectado Then
                    listar = True
                    importe = rubro.precioUnitario * factorCobro
                    servicio = ""
                    segmento = ""
                    
                    novedad.clienteId = cliente.clienteId
                    novedad.periodoId = periodo.periodoId
                    novedad.rubroID = rubro.rubroID
                    novedad.findByPrimaryKey db
                    If novedad.autoID = 0 Then
    'Analiza los rangos aplicados al consumo
                        If rubro.rangoID > 0 Then
                            rango.categoria = cliente.categoria
                            rango.rangoID = rubro.rangoID
                            rango.findLast db
                            If rango.autoID = 0 Then
                                listar = False
                            Else
                                importe = rango.tarifa * factorCobro
                                If consumoRegistrado > rango.limiteSuperior Then metrosConsumo = rango.limiteSuperior - rango.limiteInferior
                                If consumoRegistrado <= rango.limiteSuperior Then metrosConsumo = consumoRegistrado - rango.limiteInferior
                                If consumoRegistrado <= rango.limiteInferior Then
                                    metrosConsumo = 0
                                    listar = False
                                End If
                                If rango.limiteSuperior > 99998 Then
                                    segmento = " ( más de " & rango.limiteInferior & " m³)"
                                Else
                                    segmento = " (" & rango.limiteInferior & "-" & rango.limiteSuperior & " m³)"
                                End If
                                servicio = " Servicio"
                                Select Case cliente.servicio
                                    Case 1
                                        servicio = servicio & " Agua"
                                    Case 2
                                        servicio = servicio & " Cloaca"
                                    Case 3
                                        servicio = servicio & " Agua y Cloaca"
                                End Select
                                servicio = servicio & " Sistema Medido"
                            End If
                        End If
                    Else
                        listar = False
                    End If
                End If
            Else
                If cliente.cobro >= rubro.cobro And (rubro.comun Or (rubro.comunSocio And Val(cliente.numeroSocio) > 0)) Then
                    listar = True
                    importe = rubro.precioUnitario * factorCobro
                    servicio = ""
                    segmento = ""
    
                    novedad.clienteId = cliente.clienteId
                    novedad.periodoId = periodo.periodoId
                    novedad.rubroID = rubro.rubroID
                    novedad.findByPrimaryKey db
                    If novedad.autoID = 0 Then
    'Analiza los rangos aplicados al consumo
                        If rubro.rangoID > 0 Then
                            rango.categoria = cliente.categoria
                            rango.rangoID = rubro.rangoID
                            rango.findLast db
                            If rango.autoID = 0 Then
                                listar = False
                            Else
                                importe = rango.tarifa * factorCobro
                                If consumoRegistrado > rango.limiteSuperior Then metrosConsumo = rango.limiteSuperior - rango.limiteInferior
                                If consumoRegistrado <= rango.limiteSuperior Then metrosConsumo = consumoRegistrado - rango.limiteInferior
                                If consumoRegistrado <= rango.limiteInferior Then
                                    metrosConsumo = 0
                                    listar = False
                                End If
                                If rango.limiteSuperior > 99998 Then
                                    segmento = " ( más de " & rango.limiteInferior & " m³)"
                                Else
                                    segmento = " (" & rango.limiteInferior & "-" & rango.limiteSuperior & " m³)"
                                End If
                                servicio = " Servicio"
                                Select Case cliente.servicio
                                    Case 1
                                        servicio = servicio & " Agua"
                                    Case 2
                                        servicio = servicio & " Cloaca"
                                    Case 3
                                        servicio = servicio & " Agua y Cloaca"
                                End Select
                                servicio = servicio & " Sistema Medido"
                            End If
                        End If
                    Else
                        listar = False
                    End If
                End If
            End If
                
            If cliente.cobro = 3 And rubro.cobro = 1 Then listar = False
        
            If listar Then
                lrub = lrub + 1
                rfac(lrub, 0) = Right("00" & rubro.rubroID, 2)
                concepto = rubro.concepto & servicio & segmento
                If cliente.cobro = 2 And rubro.cobro = 2 Then concepto = concepto & " (" & tasaMayor & "%)"
                If cliente.cobro = 2 And rubro.cobro = 1 Then concepto = concepto & " (" & tasaMenor & "%)"
                rfac(lrub, 1) = concepto
                cantidad = 1
                If metrosConsumo > 0 Then
                    cantidad = metrosConsumo
                    If cliente.cobro = 2 And rubro.cobro = 1 Then cantidad = metrosConsumo * tasaMenor / 100
                End If
                rfac(lrub, 2) = cantidad
    'Controla si el rubro corresponde a un cargo fijo en transicion
                If rubro.cobro = 2 And cliente.cobro < 3 Then importe = importe * tasaMayor / 100
                If rubro.cobro = 1 And cliente.cobro = 2 And rubro.rangoID = 0 Then importe = importe * tasaMenor / 100
                rfac(lrub, 3) = Format(importe, "#,###,##0.00")
                rfac(lrub, 4) = Format(importe * cantidad, "#,###,##0.00")
                If rubro.IVA Then
                    rfac(lrub, 5) = "Si"
                    Select Case cliente.situacionIVA
                        Case 1, 6
                            ivat = ivat + importe * cantidad * alicuota.IVA
                        Case 2
                            ivat = ivat + importe * cantidad * alicuota.IVA
                            rnit = rnit + importe * cantidad * alicuota.rni
                        Case 3 To 5
                            ivct = ivct + importe * cantidad * alicuota.ivacf
                    End Select
                Else
                    rfac(lrub, 5) = "No"
                End If
                subTotal = subTotal + importe * cantidad
                metrosConsumo = 0
            End If
        End If
    Next
    
    If Not usuarioDesconectado Then
'Novedades únicas
        For Each novedad In novedad.collectionUnicasByClienteID(cliente.clienteId, periodo.periodoId, db)
            lrub = lrub + 1
            rfac(lrub, 0) = Right("00" & novedad.rubroID, 2)
            rubro.rubroID = novedad.rubroID
            rubro.findLast db
            rfac(lrub, 1) = rubro.concepto
            cantidad = novedad.cantidad
            If cantidad = 0 Then cantidad = novedad.porcentaje
            rfac(lrub, 2) = cantidad
            rfac(lrub, 3) = Format(rubro.precioUnitario, "#,###,##0.00")
            rfac(lrub, 4) = Format(rubro.precioUnitario * cantidad, "#,###,##0.00")
            If rubro.IVA Then
                rfac(lrub, 5) = "Si"
                Select Case cliente.situacionIVA
                    Case 1, 6
                        ivat = ivat + rubro.precioUnitario * cantidad * alicuota.IVA
                    Case 2
                        ivat = ivat + rubro.precioUnitario * cantidad * alicuota.IVA
                        rnit = rnit + rubro.precioUnitario * cantidad * alicuota.rni
                    Case 3 To 5
                        ivct = ivct + rubro.precioUnitario * cantidad * alicuota.ivacf
                End Select
            Else
                rfac(lrub, 5) = "No"
            End If
            subTotal = subTotal + rubro.precioUnitario * cantidad
        Next
        
'Novedades indefinidas
        For Each novedad In novedad.collectionIndefinidasByClienteID(cliente.clienteId, periodo.periodoId, db)
            lrub = lrub + 1
            rfac(lrub, 0) = Right("00" & novedad.rubroID, 2)
            rubro.rubroID = novedad.rubroID
            rubro.findLast db
            rfac(lrub, 1) = rubro.concepto
            cantidad = novedad.cantidad
            If cantidad = 0 Then cantidad = novedad.porcentaje
            rfac(lrub, 2) = cantidad
            rfac(lrub, 3) = Format(rubro.precioUnitario, "#,###,##0.00")
            rfac(lrub, 4) = Format(rubro.precioUnitario * cantidad, "#,###,##0.00")
            If rubro.IVA Then
                rfac(lrub, 5) = "Si"
                Select Case cliente.situacionIVA
                    Case 1, 6
                        ivat = ivat + rubro.precioUnitario * cantidad * alicuota.IVA
                    Case 2
                        ivat = ivat + rubro.precioUnitario * cantidad * alicuota.IVA
                        rnit = rnit + rubro.precioUnitario * cantidad * alicuota.rni
                    Case 3 To 5
                        ivct = ivct + rubro.precioUnitario * cantidad * alicuota.ivacf
                End Select
            Else
                rfac(lrub, 5) = "No"
            End If
            subTotal = subTotal + rubro.precioUnitario * cantidad
        Next
        
'Novedades en veces
        For Each novedad In novedad.collectionVecesByClienteID(cliente.clienteId, periodo.periodoId, db)
            lrub = lrub + 1
            rfac(lrub, 0) = Right("00" & novedad.rubroID, 2)
            rubro.rubroID = novedad.rubroID
            rubro.findLast db
            rfac(lrub, 1) = rubro.concepto
            cantidad = novedad.cantidad
            If cantidad = 0 Then cantidad = novedad.porcentaje
            rfac(lrub, 2) = cantidad
            rfac(lrub, 3) = Format(novedad.importe / novedad.veces, "#,###,##0.00")
            rfac(lrub, 4) = Format(novedad.importe * cantidad / novedad.veces, "#,###,##0.00")
            If rubro.IVA Then
                rfac(lrub, 5) = "Si"
                Select Case cliente.situacionIVA
                    Case 1, 6
                        ivat = ivat + novedad.importe * cantidad / novedad.veces * alicuota.IVA
                    Case 2
                        ivat = ivat + novedad.importe * cantidad / novedad.veces * alicuota.IVA
                        rnit = rnit + novedad.importe * cantidad / novedad.veces * alicuota.rni
                    Case 3 To 5
                        ivct = ivct + novedad.importe * cantidad / novedad.veces * alicuota.ivacf
                End Select
            Else
                rfac(lrub, 5) = "No"
            End If
            subTotal = subTotal + novedad.importe * cantidad / novedad.veces
        Next
    End If
        
    totalInteresesMora = 0
    For Each factura In factura.collectionParaInteresByClienteID(cliente.clienteId, db)
        total = factura.total
        For Each ncredito In ncredito.collectionByLiquidacion(factura.puntoVta, factura.nroComprob, db)
            total = total - ncredito.total
        Next
        periodolocal.periodoId = factura.periodoId
        periodolocal.findByPrimaryKey
        If periodolocal.fechaSegundo < factura.fechapago Then totalInteresesMora = totalInteresesMora + interes(total, periodolocal.tasa, periodolocal.fechaSegundo, factura.fechapago)
    Next
    If totalInteresesMora > 0 Then
        lrub = lrub + 1
        rfac(lrub, 0) = "00"
        rfac(lrub, 1) = "Intereses por mora de Facturas pagadas fuera de término"
        cantidad = 1
        rfac(lrub, 2) = cantidad
        rfac(lrub, 3) = Format(totalInteresesMora, "#,###,##0.00")
        rfac(lrub, 4) = Format(totalInteresesMora, "#,###,##0.00")
        rfac(lrub, 5) = "Si"
        Select Case cliente.situacionIVA
            Case 1, 6
                ivat = ivat + totalInteresesMora * alicuota.IVA
            Case 2
                ivat = ivat + totalInteresesMora * alicuota.IVA
                rnit = rnit + totalInteresesMora * alicuota.rni
            Case 3 To 5
                ivct = ivct + totalInteresesMora * alicuota.ivacf
        End Select
        subTotal = subTotal + totalInteresesMora
    End If
    tasa = periodo.tasa
    tsub = Format(subTotal, "#,###,##0.00")
    tiva = Format(ivat, "#,###,##0.00")
    tcfl = Format(ivct, "#,###,##0.00")
    trni = Format(rnit, "#,###,##0.00")
    tpri = Format(subTotal + ivct + ivat + rnit, "#,###,##0.00")
    tseg = Format(subTotal + ivct + ivat + rnit + interes(subTotal + ivct + ivat + rnit, tasa, periodo.fechaPrimero, periodo.fechaSegundo), "#,###,##0.00")
    If tsub > 0 Then
'Llena los valores de la factura
        Set factura = New clsMyAFactura
        factura.puntoVta = prefijoId
        factura.nroComprob = facturaId
        factura.fecha = fechaEmision
        factura.clienteId = cliente.clienteId
        factura.periodoId = periodo.periodoId
        factura.tasa = periodo.tasa
        factura.situacionIVA = cliente.situacionIVA
        factura.anulada = 0
        factura.total = tpri
        factura.ivacf = tcfl
        factura.ivari = tiva
        factura.ivarn = trni
        factura.uid = "admin"
        factura.pfcodigo = pagoFacilService.codigopf(liquidacionService.oldFactura2newFactura(factura))
        factura.save db
        
        For fila = 1 To lrub
            Set detalle = New clsMyADetalle
            detalle.puntoVta = prefijoId
            detalle.nroComprob = facturaId
            detalle.rubroID = rfac(fila, 0)
            detalle.concepto = Left(rfac(fila, 1), 80)
            detalle.cantidad = rfac(fila, 2)
            detalle.precioUnitario = rfac(fila, 3)
            If Trim(rfac(fila, 5)) = "Si" Then
                detalle.IVA = 1
            Else
                detalle.IVA = 0
            End If
            detalle.uid = "admin"
            detalle.save db
            
            novedad.clienteId = cliente.clienteId
            novedad.rubroID = rfac(fila, 0)
            novedad.periodoId = periodo.periodoId
            
            novedad.findVeces db
            
            If novedad.autoID > 0 Then
                novedad.vecesCobradas = novedad.vecesCobradas + 1
                novedad.save db
            End If
        Next fila
        
        For Each factura In factura.collectionParaInteresByClienteID(cliente.clienteId, db)
            periodolocal.periodoId = factura.periodoId
            periodolocal.findByPrimaryKey
            If periodolocal.fechaSegundo < factura.fechapago Then
                factura.puntoVtaInteres = prefijoId
                factura.nroComprobInteres = facturaId
                factura.uid = "admin"
                factura.save db
            End If
        Next
    End If
    
    Set clienteDato = clienteDatoRep.findByClienteId(cliente.clienteId)
    If clienteDato.email <> "" Then Status.SimpleText = liquidacionService.sendLiquidacion(prefijoId, facturaId, db)
    
End Function

Public Function deletePeriodo(periodoId As Integer, db As clsDB) As Boolean
Dim factura As New clsMyAFactura

    For Each factura In factura.collectionByPeriodoID(periodoId, db)
        deleteLiquidacion factura.puntoVta, factura.nroComprob, db
    Next
    
End Function

Public Function deleteLiquidacion(puntoVta As Integer, nroComprob As Long, db As clsDB) As Boolean
Dim factura As New clsMyAFactura
Dim facturainteres As New clsMyAFactura
Dim detalle As New clsMyADetalle
Dim novedad As New clsMyANovedad

    ' Carga de Liquidacion para obtener datos
    With factura
        .puntoVta = puntoVta
        .nroComprob = nroComprob
        
        .findByPrimaryKey db
    End With
    ' Eliminación de Detalle
    For Each detalle In detalle.collectionByLiquidacion(puntoVta, nroComprob, db)
        With novedad
            .clienteId = factura.clienteId
            .rubroID = detalle.rubroID
            .periodoId = factura.periodoId
            
            .findVeces db
            
            If .autoID <> 0 Then
                .vecesCobradas = .vecesCobradas - 1
                .save db
            End If
        End With
        
        detalle.delete db
    Next
    ' Desmarca las facturas marcadas por los intereses
    For Each facturainteres In facturainteres.collectionInteresByLiquidacion(puntoVta, nroComprob, db)
        facturainteres.puntoVtaInteres = 0
        facturainteres.nroComprobInteres = 0
        
        facturainteres.save db
    Next
    
    ' Elimina la liquidacion
    factura.delete db

End Function

