VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCtlNotificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function makeCollectionCorte(db As clsDB) As Collection
Dim clienteRep As New clsREPCliente

    Set makeCollectionCorte = clienteRep.collectionDeudoresPlanCorte
    
End Function

Public Function makeCollectionNotificacion15() As Collection
Dim clienteRep As New clsREPCliente

    Set makeCollectionNotificacion15 = clienteRep.sortCollection(clienteRep.collectionDeudoresFactura60Dias)
    
End Function

Public Function makeCollectionNotificacionOC() As Collection
Dim clienteRep As New clsREPCliente

    Set makeCollectionNotificacionOC = clienteRep.sortCollection(clienteRep.collectionDeudoresFactura60Dias)
    
End Function

Public Function makeCollectionNotificacionCorte() As Collection
Dim clienteRep As New clsREPCliente

    Set makeCollectionNotificacionCorte = clienteRep.sortCollection(clienteRep.collectionDeudoresFactura60Dias)
    
End Function

Public Function makeNotificacion15(fecha As Date, notificados As Collection, db As clsDB) As Boolean
Dim clientenotif As New clsMyAClienteNotif
Dim notificacion As clsMODNotificacion
Dim anotador As clsMyAAnotador
Dim periodoactual As New clsRESTPeriodo
Dim periodo As New clsRESTPeriodo
Dim liquidacion As New clsMyAFactura
Dim novedad As clsMyANovedad
Dim cliente As clsMODCliente
Dim tiponotificacion As clsMODTipoNotificacion

Dim clienteRep As New clsREPCliente
Dim notificacionrep As New clsREPNotificacion
Dim tiponotificacionrep As New clsREPTipoNotificacion

Dim periodos As Collection
Dim clientes As Collection

Dim consumo_service As New clsCtlConsumo

Dim periodoId As Integer

    makeNotificacion15 = False
    
    Set periodos = periodoactual.collectionAll
    Set clientes = clienteRep.collectionActivos
    
    periodoactual.findByFecha fecha
    
    If periodoactual.periodoId = 0 Then
        MsgBox "ERROR: Sin PERIODO definido"
        Exit Function
    End If
    
    Set tiponotificacion = tiponotificacionrep.findByTiponotificacionId(1)
    
    For Each clientenotif In notificados
        Set cliente = New clsMODCliente
        If modCollection.collectionExistElement(clientes, "k." & clientenotif.clienteId) Then Set cliente = clientes("k." & clientenotif.clienteId)
        If Not clientenotif.save(db) Then GoTo rollBack
        ' Se registra notificacion
        Set notificacion = notificacionrep.findByUnique(clientenotif.clienteId, clientenotif.ultimaNotificacion15)
        notificacion.clienteId = clientenotif.clienteId
        notificacion.fecha = clientenotif.ultimaNotificacion15
        notificacion.tiponotificacionId = cntNotificacion15
        notificacion.vencimiento = notificacion.fecha + 15
        notificacion.estado = consumo_service.lastLecturaByClienteId(clientenotif.clienteId)
        Set notificacion = notificacionrep.save(notificacion)
    
        db.beginTrans
    
        ' Se agrega novedad en el anotador
        Set anotador = New clsMyAAnotador
        anotador.clienteId = clientenotif.clienteId
        anotador.anotacion = "Notificación 15 días: se envía con Fecha " & notificacion.fecha & " y Vencimiento " & notificacion.vencimiento
        If Not anotador.add(db) Then GoTo rollBack
        ' Busca período para generar la novedad
        periodoId = periodoactual.periodoId - 1
        Set periodo = New clsRESTPeriodo
        If modCollection.collectionExistElement(periodos, "k." & periodoId) Then Set periodo = periodos("k." & periodoId)
        Do
            periodoId = periodoId + 1
            Set periodo = New clsRESTPeriodo
            If modCollection.collectionExistElement(periodos, "k." & periodoId) Then Set periodo = periodos("k." & periodoId)
            If periodo.periodoId = 0 Then
                MsgBox "ERROR: Sin PERIODO definido"
                GoTo rollBack
            End If
            liquidacion.clienteId = clientenotif.clienteId
            liquidacion.periodoId = periodoId
            liquidacion.findByClientePeriodo db
        Loop Until liquidacion.autoID = 0
        ' Agrega novedades
        Set novedad = New clsMyANovedad
        novedad.clienteId = clientenotif.clienteId
        novedad.periodoId = periodoId
        novedad.rubroID = IIf(cliente.numeroSocio > 0, 9, 11)
        novedad.importe = IIf(cliente.numeroSocio > 0, tiponotificacion.valorsocio, tiponotificacion.valornosocio)
        novedad.fecha = notificacion.fecha
        novedad.cantidad = 1
        novedad.veces = 1
        If Not novedad.save(db) Then GoTo rollBack
        
        db.commitTrans
    
    Next
    
    makeNotificacion15 = True
    
    Exit Function

rollBack:
    db.rollBack
    
    MsgBox "ERROR: No Pudo Generar NOTIFICACIONES"
    
End Function

Public Function makeNotificacionOC(fecha As Date, notificados As Collection, db As clsDB) As Boolean
Dim clientenotif As New clsMyAClienteNotif
Dim notificacion As clsMODNotificacion
Dim anotador As clsMyAAnotador
Dim periodoactual As New clsRESTPeriodo
Dim periodo As New clsRESTPeriodo
Dim liquidacion As New clsMyAFactura
Dim novedad As clsMyANovedad
Dim cliente As clsMODCliente
Dim tiponotificacion As clsMODTipoNotificacion

Dim clienteRep As New clsREPCliente
Dim notificacionrep As New clsREPNotificacion
Dim tiponotificacionrep As New clsREPTipoNotificacion

Dim consumo_service As New clsCtlConsumo

Dim periodos As Collection
Dim clientes As Collection

Dim periodoId As Integer

    makeNotificacionOC = False
    
    Set periodos = periodoactual.collectionAll
    Set clientes = clienteRep.collectionActivos
    
    periodoactual.findByFecha fecha
    
    If periodoactual.periodoId = 0 Then
        MsgBox "ERROR: Sin PERIODO definido"
        Exit Function
    End If
    
    Set tiponotificacion = tiponotificacionrep.findByTiponotificacionId(2)
    
    For Each clientenotif In notificados
        Set cliente = New clsMODCliente
        If modCollection.collectionExistElement(clientes, "k." & clientenotif.clienteId) Then Set cliente = clientes("k." & clientenotif.clienteId)
        If Not clientenotif.save(db) Then GoTo rollBack
        ' Se registra notificacion
        Set notificacion = notificacionrep.findByUnique(clientenotif.clienteId, clientenotif.ultimaNotificacion48)
        notificacion.clienteId = clientenotif.clienteId
        notificacion.fecha = clientenotif.ultimaNotificacion48
        notificacion.tiponotificacionId = cntNotificacionOC
        notificacion.vencimiento = notificacion.fecha + 2
        notificacion.estado = consumo_service.lastLecturaByClienteId(clientenotif.clienteId)
        Set notificacion = notificacionrep.save(notificacion)
        
        db.beginTrans
    
        ' Se agrega novedad en el anotador
        Set anotador = New clsMyAAnotador
        anotador.clienteId = clientenotif.clienteId
        anotador.anotacion = "Notificación Orden Corte: se envía con Fecha " & notificacion.fecha & " y Vencimiento " & notificacion.vencimiento & " - Lectura Medidor: " & notificacion.estado
        If Not anotador.add(db) Then GoTo rollBack
        ' Busca período para generar la novedad
        periodoId = periodoactual.periodoId - 1
        
        db.commitTrans
    
    Next
    
    makeNotificacionOC = True
    
    Exit Function

rollBack:
    db.rollBack
    
    MsgBox "ERROR: No Pudo Generar NOTIFICACIONES"
    
End Function

Public Function makeNotificacionCorte(fecha As Date, notificados As Collection, db As clsDB) As Boolean
Dim clientenotif As New clsMyAClienteNotif
Dim notificacion As clsMODNotificacion
Dim anotador As clsMyAAnotador
Dim periodoactual As New clsRESTPeriodo
Dim periodo As New clsRESTPeriodo
Dim liquidacion As New clsMyAFactura
Dim novedad As clsMyANovedad
Dim cliente As clsMODCliente
Dim tiponotificacion As clsMODTipoNotificacion

Dim clienteRep As New clsREPCliente
Dim notificacionrep As New clsREPNotificacion
Dim tiponotificacionrep As New clsREPTipoNotificacion

Dim consumo_service As New clsCtlConsumo

Dim periodos As Collection
Dim clientes As Collection

Dim periodoId As Integer

    makeNotificacionCorte = False
    
    Set periodos = periodoactual.collectionAll
    Set clientes = clienteRep.collectionActivos
    
    periodoactual.findByFecha fecha
    
    If periodoactual.periodoId = 0 Then
        MsgBox "ERROR: Sin PERIODO definido"
        GoTo rollBack
    End If
    
    Set tiponotificacion = tiponotificacionrep.findByTiponotificacionId(3)
    
    For Each clientenotif In notificados
        Set cliente = New clsMODCliente
        If modCollection.collectionExistElement(clientes, "k." & clientenotif.clienteId) Then Set cliente = clientes("k." & clientenotif.clienteId)
        cliente.cortado = 1
        Set cliente = clienteRep.save(cliente)
        ' Se registra notificacion
        Set notificacion = notificacionrep.findByUnique(clientenotif.clienteId, clientenotif.ultimaNotificacionCorte)
        notificacion.clienteId = clientenotif.clienteId
        notificacion.fecha = clientenotif.ultimaNotificacionCorte
        notificacion.tiponotificacionId = cntNotificacionCorte
        notificacion.vencimiento = notificacion.fecha
        notificacion.estado = consumo_service.lastLecturaByClienteId(clientenotif.clienteId)
        Set notificacion = notificacionrep.save(notificacion)
        
        db.beginTrans
    
        If Not clientenotif.save(db) Then GoTo rollBack
        ' Se agrega novedad en el anotador
        Set anotador = New clsMyAAnotador
        anotador.clienteId = clientenotif.clienteId
        anotador.anotacion = "Notificación Corte: se envía con Fecha " & notificacion.fecha & " - Lectura Medidor: " & notificacion.estado
        If Not anotador.add(db) Then GoTo rollBack
        ' Busca período para generar la novedad
        periodoId = periodoactual.periodoId - 1
        Set periodo = New clsRESTPeriodo
        If modCollection.collectionExistElement(periodos, "k." & periodoId) Then Set periodo = periodos("k." & periodoId)
        Do
            periodoId = periodoId + 1
            Set periodo = New clsRESTPeriodo
            If modCollection.collectionExistElement(periodos, "k." & periodoId) Then Set periodo = periodos("k." & periodoId)
            If periodo.periodoId = 0 Then
                MsgBox "ERROR: Sin PERIODO definido"
                GoTo rollBack
            End If
            liquidacion.clienteId = clientenotif.clienteId
            liquidacion.periodoId = periodoId
            liquidacion.findByClientePeriodo db
        Loop Until liquidacion.autoID = 0
        ' Agrega novedades
        Set novedad = New clsMyANovedad
        novedad.clienteId = clientenotif.clienteId
        novedad.periodoId = periodoId
        novedad.rubroID = IIf(Val(cliente.numeroSocio) > 0, 9, 11)
        novedad.importe = IIf(Val(cliente.numeroSocio) > 0, tiponotificacion.valorsocio, tiponotificacion.valornosocio)
        novedad.fecha = notificacion.fecha
        novedad.cantidad = 1
        novedad.veces = 1
        If Not novedad.save(db) Then GoTo rollBack
    
        db.commitTrans
    
    Next
    
    makeNotificacionCorte = True
    
    Exit Function

rollBack:
    db.rollBack
    
    MsgBox "ERROR: No Pudo Generar NOTIFICACIONES"
    
End Function

Public Function updateInteresesByFecha(fecha As Date, db As clsDB) As Boolean
Dim notificacion As New clsMODNotificacion
Dim liquidacion As New clsMyAFactura
Dim periodo As New clsRESTPeriodo

Dim pagofacil_service As New clsCtlPagoFacil
Dim liquidacion_service As New clsCtlLiquidacion

Dim notificacionrep As New clsREPNotificacion

Dim periodos As Collection

    updateInteresesByFecha = False
    
    Set periodos = periodo.collectionAll
    
    db.beginTrans
    
    For Each notificacion In notificacionrep.collectionByFecha(fecha)
        For Each liquidacion In liquidacion.collectionDeudaByClienteId(notificacion.clienteId, db)
            Set periodo = New clsRESTPeriodo
            If modCollection.collectionExistElement(periodos, "k." & liquidacion.periodoId) Then Set periodo = periodos("k." & liquidacion.periodoId)
            liquidacion.interes = interes(liquidacion.total, liquidacion.tasa, periodo.fechaPrimero, notificacion.fecha)
            liquidacion.pfcodigo = pagofacil_service.codigopf(liquidacion_service.oldFactura2newFactura(liquidacion))
            If Not liquidacion.update(db) Then GoTo rollBack
        Next
    Next
    
    db.commitTrans
    
    updateInteresesByFecha = True
    
    Exit Function
    
rollBack:
    db.rollBack
    
    MsgBox "ERROR: Intereses No ACTUALIZADOS"
    
End Function

Public Sub printCorte(clienteId As Long, fecha As Date, report As CrystalReport)
Dim np As Integer
Dim ct As Integer
Dim num As Integer

Dim busq As String
Dim mens As String
Dim tipo As String
Dim cuit As String
Dim numero As String

Dim total As Currency

Dim suspension As New clsMyASuspension
Dim periodo As New clsRESTPeriodo
Dim operador As New clsMyAOperador
Dim cliente As New clsMODCliente
Dim factura As New clsMyAFactura
Dim deuda As New clsMyADeuda

Dim clienteRep As clsREPCliente

Dim consumo_service As New clsCtlConsumo

Dim periodos As Collection

On Error Resume Next
    
    report.ReportFileName = App.path & "\rptCorte.rpt"
    tipo = "C"
    numero = "numcor= "
    
    report.Destination = crptToPrinter
    
    num = 1
    suspension.tipo = tipo
    suspension.findLast dbapp
    num = suspension.numero + 1
    
    suspension.tipo = tipo
    suspension.clienteId = clienteId
    suspension.fecha = fecha
    suspension.findByClienteID dbapp
    If suspension.autoID > 0 Then num = suspension.numero
    
    periodo.findToday
    
    suspension.tipo = tipo
    suspension.numero = num
    suspension.fecha = fecha
    suspension.clienteId = clienteId
    suspension.periodoId = periodo.periodoId + 1
    suspension.uid = "admin"
    suspension.save dbapp
    
    operador.findLast dbapp
    
    Set clienteRep = New clsREPCliente
    Set cliente = clienteRep.findLastByClienteID(clienteId)
    Set clienteRep = Nothing
    
    Set periodos = periodo.collectionAll
    
    cuit = Left(operador.cuit, 2) & "-" & Mid(operador.cuit, 3, 8) & "-" & Right(operador.cuit, 1)
    np = 31
    If np > Len(operador.razonSocial) Then
        np = Len(operador.razonSocial)
    Else
        Do While Mid(operador.razonSocial, np, 1) <> " "
            np = np - 1
        Loop
    End If
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
    End Select
    report.ParameterFields(0) = "cliente_id;" & clienteId & ";TRUE"
    report.ParameterFields(1) = "lectura;" & consumo_service.lastLecturaByClienteId(clienteId) & ";TRUE"
    report.formulas(0) = "nomope= '" & Mid(operador.razonSocial, 1, np) & "'"
    report.formulas(1) = "nomop1= '" & Mid(operador.razonSocial, np + 1) & "'"
    report.formulas(2) = "domope= '" & operador.calle & " " & operador.puerta & " " & operador.piso & " " & operador.dpto & " - " & operador.localidad & "'"
    report.formulas(3) = "locope= 'C.P. " & operador.codigoPostal & " " & operador.provincia & " - Tel: " & operador.telefono & "'"
    report.formulas(4) = "opegr1= 'C.U.I.T. NRO: " & cuit & " ING. BRUTOS: " & operador.ingresosBrutos & "'"
    report.formulas(5) = "opegr2= 'I.V.A. " & mens & "   NRO. E.P.A.S. " & operador.numeroEpas & "'"
    report.formulas(42) = "resol= '" & operador.resolucion & "'"
    report.formulas(43) = "perso= '" & operador.personeria & "'"
    report.formulas(6) = "nomcli= '" & cliente.apellido & ", " & cliente.nombre & "'"
    report.formulas(7) = "ubiinm= '" & cliente.inmuebleCalle & " " & cliente.inmueblePuerta & " " & cliente.inmueblePiso & " " & cliente.inmuebleDpto & "'"
    report.formulas(8) = "nomcat= ''"
    report.formulas(9) = "domcli= '" & cliente.fiscalCalle & " " & cliente.fiscalPuerta & " " & cliente.fiscalPiso & " " & cliente.fiscalDpto & "'"
    report.formulas(10) = "numsoc= '" & cliente.numeroSocio & "'"
    report.formulas(39) = "fecemi= '" & fecha & "'"
    report.formulas(40) = numero & "'" & num & "'"
    mens = ""
    If Len(Trim(cliente.cuit)) > 0 Then mens = Left(cliente.cuit, 2) & "-" & Mid(cliente.cuit, 3, 8) & "-" & Right(cliente.cuit, 1)
    Select Case cliente.situacionIVA
        Case 1
            mens = mens & " R.I."
        Case 2
            mens = mens & " R.N.I."
        Case 3
            mens = "C. Final"
        Case 4
            mens = mens & " IVA Exento"
        Case 5
            mens = mens & " IVA No Resp."
    End Select
    report.formulas(11) = "cuicli= '" & mens & "'"
    report.formulas(12) = "numcli= '" & clienteId & "'"
    Select Case cliente.categoria
        Case 1
            mens = "General"
        Case 2
            mens = "Especial"
    End Select
    report.formulas(13) = "catego= '" & mens & "'"
    For ct = 1 To 5
        report.formulas(13 + ct) = "per(" & ct & ")= ''"
        report.formulas(18 + ct) = "fac(" & ct & ")= ''"
        report.formulas(23 + ct) = "ven(" & ct & ")= ''"
        report.formulas(28 + ct) = "imf(" & ct & ")= ''"
        report.formulas(33 + ct) = "iin(" & ct & ")= ''"
    Next ct
    
    total = 0
    ct = 1
    For Each factura In factura.collectionDeudaByClienteId(clienteId, dbapp)
        If ct < 6 Then
            Set periodo = periodos("k." & factura.periodoId)
            report.formulas(13 + ct) = "per(" & ct & ")= '" & periodo.descripcion & "'"
            report.formulas(18 + ct) = "fac(" & ct & ")= '" & Right("0000" & factura.puntoVta, 4) & "-" & Right("00000000" & factura.nroComprob, 8) & "'"
            report.formulas(23 + ct) = "ven(" & ct & ")= '" & periodo.fechaPrimero & "'"
            report.formulas(28 + ct) = "imf(" & ct & ")= '" & Format(factura.total, "#,###,##0.00") & "'"
            report.formulas(33 + ct) = "iin(" & ct & ")= '" & Format(factura.total + interes(factura.total, factura.tasa, factura.fecha, fecha), "#,###,##0.00") & "'"
            total = total + factura.total + interes(factura.total, factura.tasa, factura.fecha, fecha)
        End If
        ct = ct + 1
    Next
    For Each deuda In deuda.collectionDeudaByClienteId(clienteId, dbapp)
        If ct < 6 Then
            report.formulas(23 + ct) = "ven(" & ct & ")= '" & deuda.cuotas - deuda.cuotasPagadas & " cuota(s)'"
            report.formulas(33 + ct) = "iin(" & ct & ")= '" & Format(deuda.deuda, "#,###,##0.00") & "'"
            total = total + deuda.deuda
        End If
        ct = ct + 1
    Next
    report.formulas(41) = "total= '" & Format(total, "#,###,##0.00") & "'"
    
    frmImpresora.Show vbModal
    frmImpresora.cargar report
    If Not frmImpresora.Cancel Then report.Action = 1

End Sub
