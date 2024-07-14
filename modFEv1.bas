Attribute VB_Name = "modFEv1"
Option Explicit

Public Function cae(tipoId As Integer, clienteId As Long, total As Currency, exento As Currency, neto27 As Currency, neto21 As Currency, neto135 As Currency, iva27 As Currency, iva21 As Currency, iva135 As Currency, produccion As Boolean, pDBmy As clsDB, nroComprob As Long, caeVencimiento As String, barras As String) As String
Dim WSAA As Object
Dim WSFEv1 As Object

Dim strCert As String
Dim strClav As String
Dim strFecha As String
Dim strCAE As String
Dim strExcepcion As String
Dim strCache As String
Dim strProxy As String
Dim strWrapper As String
Dim strCACert As String
Dim strWSDL As String
Dim strDocu As String
Dim strToken As String
Dim strSign As String
Dim strDest As String
Dim strTA_xml As String
Dim strMoneda As String

Dim cliente As clsMODCliente
Dim operador As New clsMyAOperador
Dim clientedato As New clsMyAClienteDato
Dim registrocae As New clsMyARegistroCAE
Dim comprobante As New clsMyATipoComprobante

Dim clienteRep As New clsREPCliente

Dim lngTTL As Long

Dim intFD As Integer
Dim intTipo As Integer

Dim tra
Dim cms
Dim ok
Dim expiration

Dim varUltimo As Variant
Dim v As Variant
    
On Error GoTo ManejoError
    
    cae = ""
    
    operador.findLast pDBmy
    
    With comprobante
        .tipoId = tipoId
        .findByPrimaryKey pDBmy
        
        If .facturaElectronica = 0 Then Exit Function
    End With
    
    strMoneda = "PES"
    
    Set cliente = clienteRep.findLastByClienteID(clienteId)
    
    intTipo = 80
    strDocu = Trim(Replace(cliente.cuit, "-", ""))
    
    If Val(strDocu) = 0 Then
        clientedato.clienteId = clienteId
        clientedato.findByPrimaryKey pDBmy
        
        intTipo = 96
        strDocu = clientedato.documento
    End If
        
    If Val(strDocu) = 0 Then
        intTipo = 99
        strDocu = "0"
    End If
    
    Set WSAA = CreateObject("WSAA")
    
    Debug.Assert WSAA.Version >= "2.04a"
    ' deshabilito errores no manejados (version 2.04 o superior)
    WSAA.LanzarExcepciones = False
    
    ' datos de prueba del certificado (para depuración):
    If produccion Then
        strDest = "C=AR, O=" & operador.razonSocial & ", serialNumber=CUIT " & operador.cuit & ", CN=Facturacion"
    Else
        strDest = "C=AR, O=Daniel Eusebio Quinteros, serialNumber=CUIT 23236938409, CN=Facturacion"
    End If
    
    ' inicializo las variables:
    strToken = ""
    strSign = ""
        
    If Dir("ta.xml") <> "" Then
        ' leo el xml almacenado del archivo
        Open "ta.xml" For Input As #1
        Line Input #1, strTA_xml
        Close #1
        ' analizo el ticket de acceso previo:
        ok = WSAA.AnalizarXml(strTA_xml)
        ' verifico que el destino corresponda (CUIT)
        Debug.Assert WSAA.ObtenerTagXml("destination") = strDest
        ' verificar CUIT
        If Not WSAA.Expirado() Then
            ' puedo reusar el ticket de acceso:
            strToken = WSAA.ObtenerTagXml("token")
            strSign = WSAA.ObtenerTagXml("sign")
        End If
    End If
    
    If strToken = "" Or strSign = "" Then
        ' Generar un Ticket de Requerimiento de Acceso (TRA) para WSFEv1
        lngTTL = 43200 ' tiempo de vida hasta expiración
        tra = WSAA.CreateTRA("wsfe", lngTTL)
        controlarExcepcion WSAA
        Debug.Print tra
    
        ' Certificado: certificado es el firmado por la AFIP
        ' ClavePrivada: la clave privada usada para crear el certificado
        If produccion Then
            strCert = App.path & "\" & cntCertificado & ".crt" ' certificado
            strClav = App.path & "\" & cntCertificado & ".key"  ' clave privada
        Else
            strCert = App.path & "\dqmdz.crt" ' certificado
            strClav = App.path & "\dqmdz.key" ' clave privada
        End If

        cms = WSAA.SignTRA(tra, strCert, strClav)
        controlarExcepcion WSAA
        Debug.Print cms
        
        If cms <> "" Then
            ' Conectarse con el webservice de autenticación:
            strCache = ""
            strProxy = "" '"usuario:clave@localhost:8000"
            strWrapper = "" ' libreria http (httplib2, urllib2, pycurl)
            strCACert = WSAA.InstallDir & "\conf\afip_ca_info.crt" ' certificado de la autoridad de certificante
            
            If produccion Then
                strWSDL = "https://wsaa.afip.gov.ar/ws/services/LoginCms?wsdl" ' Producción
            Else
                strWSDL = "https://wsaahomo.afip.gov.ar/ws/services/LoginCms?wsdl" ' Homologación
            End If
            
            ok = WSAA.Conectar(strCache, strWSDL, strProxy, strWrapper, strCACert)
            controlarExcepcion WSAA
        
            ' Llamar al web service para autenticar:
            strTA_xml = WSAA.LoginCMS(cms)
            controlarExcepcion WSAA

            ' Imprimir el ticket de acceso, ToKen y Sign de autorización
            Debug.Print strTA_xml
            Debug.Print "Token:", WSAA.Token
            Debug.Print "Sign:", WSAA.Sign
            
            If strTA_xml <> "" Then
                ' guardo el ticket de acceso en el archivo
                Open "ta.xml" For Output As #1
                Print #1, strTA_xml
                Close #1
            End If
            
            strToken = WSAA.Token
            strSign = WSAA.Sign
        End If
        ' reviso que no haya errores:
        Debug.Print "excepcion", WSAA.Excepcion
        If WSAA.Excepcion <> "" Then
            Debug.Print WSAA.Traceback
            MsgBox WSAA.Excepcion, vbCritical, "Excepción"
        End If
    End If

    ' Crear objeto interface Web Service de Factura Electrónica de Mercado Interno
    Set WSFEv1 = CreateObject("WSFEv1")
    Debug.Print WSFEv1.Version
    
    ' Setear tocken y sign de autorización (pasos previos)
    WSFEv1.Token = strToken
    WSFEv1.Sign = strSign
    
    ' CUIT del emisor (debe estar registrado en la AFIP)
    If produccion Then
        WSFEv1.cuit = operador.cuit
    Else
        WSFEv1.cuit = "23236938409"
    End If
    
    ' deshabilito errores no manejados
    WSFEv1.LanzarExcepciones = False

    ' Conectar al Servicio Web de Facturación
    strProxy = "" ' "usuario:clave@localhost:8000"
    strCache = "" 'Path
    strWrapper = "" ' libreria http (httplib2, urllib2, pycurl)
    strCACert = WSAA.InstallDir & "\conf\afip_ca_info.crt" ' certificado de la autoridad de certificante (solo pycurl)
    
    If produccion Then
        strWSDL = "https://servicios1.afip.gov.ar/wsfev1/service.asmx?WSDL" ' producción
    Else
        strWSDL = "https://wswhomo.afip.gov.ar/wsfev1/service.asmx?WSDL" ' homologación
    End If
    
    ok = WSFEv1.Conectar(strCache, strWSDL, strProxy, strWrapper, strCACert) ' homologación
    Debug.Print WSFEv1.Version
    controlarExcepcion WSFEv1
    
    ' mostrar bitácora de depuración:
    Debug.Print WSFEv1.DebugLog
    
    ' Llamo a un servicio nulo, para obtener el estado del servidor (opcional)
    WSFEv1.Dummy
    controlarExcepcion WSFEv1
    Debug.Print "appserver status", WSFEv1.AppServerStatus
    Debug.Print "dbserver status", WSFEv1.DbServerStatus
    Debug.Print "authserver status", WSFEv1.AuthServerStatus

    ' Establezco los valores de la factura o lote a autorizar:
    ' Establezco los valores de la factura a autorizar:
    varUltimo = WSFEv1.CompUltimoAutorizado(comprobante.comprobanteId, comprobante.puntoVenta)
    controlarExcepcion WSFEv1
    For Each v In WSFEv1.errores
        Debug.Print v
    Next
    Debug.Print WSFEv1.errmsg
    Debug.Print WSFEv1.errcode
    If varUltimo = "" Then
        varUltimo = 0                ' no hay comprobantes emitidos
    Else
        varUltimo = CLng(varUltimo)   ' convertir a entero largo
    End If
    Debug.Print varUltimo
    
    strFecha = Format(Date, "yyyymmdd")
    
    ok = WSFEv1.crearFactura(1, intTipo, strDocu, comprobante.comprobanteId, comprobante.puntoVenta, _
        varUltimo + 1, varUltimo + 1, Format(Abs(total), "0.00"), Format(Abs(exento), "0.00"), Format(Abs(neto27) + Abs(neto21) + Abs(neto135), "0.00"), Format(Abs(iva27) + Abs(iva21) + Abs(iva135), "0.00"), _
        "0.00", "0.00", strFecha, , , , strMoneda, "1.000")
        
    ' Agrego tasas de iva
    If iva27 > 0 Then ok = WSFEv1.agregariva(6, Format(Abs(neto27), "0.00"), Format(Abs(iva27), "0.00"))
    If iva21 > 0 Then ok = WSFEv1.agregariva(5, Format(Abs(neto21), "0.00"), Format(Abs(iva21), "0.00"))
    If iva135 > 0 Then ok = WSFEv1.agregariva(4, Format(Abs(neto135), "0.00"), Format(Abs(iva135), "0.00"))
    
    ' Habilito reprocesamiento automático (predeterminado):
    WSFEv1.Reprocesar = True

    ' Solicito CAE:
    strCAE = WSFEv1.CAESolicitar()
    controlarExcepcion WSFEv1
    
    Debug.Print "Resultado", WSFEv1.Resultado
    Debug.Print "CAE", WSFEv1.cae

    Debug.Print "Numero de comprobante:", WSFEv1.cbtenro
    
    ' Imprimo pedido y respuesta XML para depuración (errores de formato)
    Debug.Print WSFEv1.XmlRequest
    Debug.Print WSFEv1.XmlResponse
    
    Debug.Print "Reprocesar:", WSFEv1.Reprocesar
    Debug.Print "Reproceso:", WSFEv1.Reproceso
    Debug.Print "CAE:", WSFEv1.cae
    Debug.Print "EmisionTipo:", WSFEv1.EmisionTipo

    'MsgBox "Resultado:" & WSFEv1.Resultado & " CAE: " & WSFEv1.cae & " Venc: " & WSFEv1.Vencimiento & " Obs: " & WSFEv1.obs & " Reproceso: " & WSFEv1.Reproceso, vbInformation + vbOKOnly
    
    nroComprob = WSFEv1.cbtenro
    caeVencimiento = modFecha.YYYYMMDD2DDMMYYYY(WSFEv1.vencimiento)
    barras = generateI2of5(operador.cuit, comprobante.comprobanteId, comprobante.puntoVenta, strCAE, caeVencimiento)
    
    With registrocae
        .tipoId = tipoId
        .prefijo = comprobante.puntoVenta
        .numero = WSFEv1.cbtenro
        .clienteId = clienteId
        .total = total
        .exento = exento
        .neto27 = neto27
        .neto = neto21
        .neto105 = neto135
        .iva27 = iva27
        .IVA = iva21
        .iva105 = iva135
        .cae = strCAE
        .fecha = Format(Date, "ddmmyyyy")
        .caeVencimiento = caeVencimiento
        .barras = barras
    
        .save pDBmy
    End With
    
    cae = strCAE
    
    Exit Function
    
ManejoError:
    ' Si hubo error (tradicional, no controlado):
    
    ' Depuración (grabar a un archivo los detalles del error)
    intFD = FreeFile
    Open App.path & "\error.txt" For Append As intFD
    If Not WSAA Is Nothing Then
        If WSAA.Version >= "1.02a" Then
            Print #intFD, WSAA.Excepcion
            Print #intFD, WSAA.Traceback
            Print #intFD, WSAA.XmlRequest
            Print #intFD, WSAA.XmlResponse
            ' guardo mensaje de error para mostrarlo:
            strExcepcion = WSAA.Excepcion
        End If
    End If
    If Not WSFEv1 Is Nothing Then
        If WSFEv1.Version >= "1.10a" Then
            Print #intFD, WSFEv1.Excepcion
            Print #intFD, WSFEv1.Traceback
            Print #intFD, WSFEv1.XmlRequest
            Print #intFD, WSFEv1.XmlResponse
            Print #intFD, WSFEv1.DebugLog()
            ' guardo mensaje de error para mostrarlo:
            strExcepcion = WSFEv1.Excepcion
        End If
    End If
    Close intFD
    
    Debug.Print Err.description            ' descripción error afip
    Debug.Print Err.Number - vbObjectError ' codigo error afip
    If strExcepcion = "" Then                 ' si no tengo mensaje de excepcion
        strExcepcion = Err.description        ' uso el error de VB
    End If
    
    ' Mostrar el mensaje de error
    Select Case MsgBox(strExcepcion, vbCritical + vbRetryCancel, "Error:" & Err.Number - vbObjectError & " en " & Err.Source)
        Case vbRetry
            Debug.Assert False
            Resume
        Case vbCancel
            Debug.Print Err.description
    End Select
    
End Function

Public Sub consultaComp(pProduccion As Boolean, db As clsDB)
Dim WSAA As Object
Dim WSFEv1 As Object

Dim strCert As String
Dim strClav As String
Dim strFecha As String
Dim strCAE As String
Dim strExcepcion As String
Dim strCache As String
Dim strProxy As String
Dim strWrapper As String
Dim strCACert As String
Dim strWSDL As String
Dim strDocu As String
Dim strToken As String
Dim strSign As String
Dim strDest As String
Dim strTA_xml As String

Dim lngTTL As Long

Dim intFD As Integer
Dim intTipo As Integer

Dim tra
Dim cms
Dim ok

Dim operador As New clsMyAOperador

Dim varUltimo As Variant
Dim v As Variant
    
On Error GoTo ManejoError

    operador.findLast db
    
    Set WSAA = CreateObject("WSAA")
    
    Debug.Assert WSAA.Version >= "2.04a"
    ' deshabilito errores no manejados (version 2.04 o superior)
    WSAA.LanzarExcepciones = False
    
    ' datos de prueba del certificado (para depuración):
    If pProduccion Then
        strDest = "C=AR, O=" & operador.razonSocial & ", serialNumber=CUIT " & operador.cuit & ", CN=Facturacion"
    Else
        strDest = "C=AR, O=Daniel Eusebio Quinteros, serialNumber=CUIT 23236938409, CN=Facturacion"
    End If
    
    ' inicializo las variables:
    strToken = ""
    strSign = ""
        
    If strToken = "" Or strSign = "" Then
        ' Generar un Ticket de Requerimiento de Acceso (TRA) para WSFEv1
        lngTTL = 43200 ' tiempo de vida hasta expiración
        tra = WSAA.CreateTRA("wsfe", lngTTL)
        controlarExcepcion WSAA
        Debug.Print tra
    
        ' Certificado: certificado es el firmado por la AFIP
        ' ClavePrivada: la clave privada usada para crear el certificado
        If pProduccion Then
            strCert = App.path & "\" & cntCertificado & ".crt" ' certificado
            strClav = App.path & "\" & cntCertificado & ".key"  ' clave privada
        Else
            strCert = App.path & "\dqmdz.crt" ' certificado
            strClav = App.path & "\dqmdz.key" ' clave privada
        End If

        cms = WSAA.SignTRA(tra, strCert, strClav)
        controlarExcepcion WSAA
        Debug.Print cms
        
        If cms <> "" Then
            ' Conectarse con el webservice de autenticación:
            strCache = ""
            strProxy = "" '"usuario:clave@localhost:8000"
            strWrapper = "" ' libreria http (httplib2, urllib2, pycurl)
            strCACert = WSAA.InstallDir & "\conf\afip_ca_info.crt" ' certificado de la autoridad de certificante
            
            If pProduccion Then
                strWSDL = "https://wsaa.afip.gov.ar/ws/services/LoginCms?wsdl" ' Producción
            Else
                strWSDL = "https://wsaahomo.afip.gov.ar/ws/services/LoginCms?wsdl" ' Homologación
            End If
            
            ok = WSAA.Conectar(strCache, strWSDL, strProxy, strWrapper, strCACert)
            controlarExcepcion WSAA
        
            ' Llamar al web service para autenticar:
            strTA_xml = WSAA.LoginCMS(cms)
            controlarExcepcion WSAA

            ' Imprimir el ticket de acceso, ToKen y Sign de autorización
            Debug.Print strTA_xml
            Debug.Print "Token:", WSAA.Token
            Debug.Print "Sign:", WSAA.Sign
            
            strToken = WSAA.Token
            strSign = WSAA.Sign
        End If
        ' reviso que no haya errores:
        Debug.Print "excepcion", WSAA.Excepcion
        If WSAA.Excepcion <> "" Then
            Debug.Print WSAA.Traceback
            MsgBox WSAA.Excepcion, vbCritical, "Excepción"
        End If
    End If

    ' Crear objeto interface Web Service de Factura Electrónica de Mercado Interno
    Set WSFEv1 = CreateObject("WSFEv1")
    Debug.Print WSFEv1.Version
    
    ' Setear tocken y sign de autorización (pasos previos)
    WSFEv1.Token = strToken
    WSFEv1.Sign = strSign
    
    ' CUIT del emisor (debe estar registrado en la AFIP)
    If pProduccion Then
        WSFEv1.cuit = operador.cuit
    Else
        WSFEv1.cuit = "23236938409"
    End If
    
    ' deshabilito errores no manejados
    WSFEv1.LanzarExcepciones = False

    ' Conectar al Servicio Web de Facturación
    strProxy = "" ' "usuario:clave@localhost:8000"
    strCache = "" 'Path
    strWrapper = "" ' libreria http (httplib2, urllib2, pycurl)
    strCACert = WSAA.InstallDir & "\conf\afip_ca_info.crt" ' certificado de la autoridad de certificante (solo pycurl)
    
    If pProduccion Then
        strWSDL = "https://servicios1.afip.gov.ar/wsfev1/service.asmx?WSDL" ' producción
    Else
        strWSDL = "https://wswhomo.afip.gov.ar/wsfev1/service.asmx?WSDL" ' homologación
    End If
    
    ok = WSFEv1.Conectar(strCache, strWSDL, strProxy, strWrapper, strCACert) ' homologación
    Debug.Print WSFEv1.Version
    controlarExcepcion WSFEv1
    
    ' mostrar bitácora de depuración:
    Debug.Print WSFEv1.DebugLog
    
    ' Llamo a un servicio nulo, para obtener el estado del servidor (opcional)
    WSFEv1.Dummy
    controlarExcepcion WSFEv1
    Debug.Print "appserver status", WSFEv1.AppServerStatus
    Debug.Print "dbserver status", WSFEv1.DbServerStatus
    Debug.Print "authserver status", WSFEv1.AuthServerStatus

    ' Buscar la factura
    strCAE = WSFEv1.CompConsultar(1, 22, 4052)
    
    Debug.Print "Fecha Comprobante:", WSFEv1.FechaCbte
    Debug.Print "Fecha Vencimiento CAE", WSFEv1.vencimiento
    Debug.Print "Importe Total:", WSFEv1.ImpTotal

    ' Habilito reprocesamiento automático (predeterminado):
    WSFEv1.Reprocesar = True

    Debug.Print "Resultado", WSFEv1.Resultado
    Debug.Print "CAE", WSFEv1.cae

    Debug.Print "Numero de comprobante:", WSFEv1.cbtenro
    
    ' Imprimo pedido y respuesta XML para depuración (errores de formato)
    Debug.Print WSFEv1.XmlRequest
    Debug.Print WSFEv1.XmlResponse
    
    Debug.Print "Reprocesar:", WSFEv1.Reprocesar
    Debug.Print "Reproceso:", WSFEv1.Reproceso
    Debug.Print "CAE:", WSFEv1.cae
    Debug.Print "EmisionTipo:", WSFEv1.EmisionTipo

    MsgBox "Resultado:" & WSFEv1.Resultado & " CAE: " & WSFEv1.cae & " Venc: " & WSFEv1.vencimiento & " Obs: " & WSFEv1.obs & " Reproceso: " & WSFEv1.Reproceso, vbInformation + vbOKOnly
    
    Exit Sub
    
ManejoError:
    ' Si hubo error (tradicional, no controlado):
    
    ' Depuración (grabar a un archivo los detalles del error)
    intFD = FreeFile
    Open App.path & "\error.txt" For Append As intFD
    If Not WSAA Is Nothing Then
        If WSAA.Version >= "1.02a" Then
            Print #intFD, WSAA.Excepcion
            Print #intFD, WSAA.Traceback
            Print #intFD, WSAA.XmlRequest
            Print #intFD, WSAA.XmlResponse
            ' guardo mensaje de error para mostrarlo:
            strExcepcion = WSAA.Excepcion
        End If
    End If
    If Not WSFEv1 Is Nothing Then
        If WSFEv1.Version >= "1.10a" Then
            Print #intFD, WSFEv1.Excepcion
            Print #intFD, WSFEv1.Traceback
            Print #intFD, WSFEv1.XmlRequest
            Print #intFD, WSFEv1.XmlResponse
            Print #intFD, WSFEv1.DebugLog()
            ' guardo mensaje de error para mostrarlo:
            strExcepcion = WSFEv1.Excepcion
        End If
    End If
    Close intFD
    
    Debug.Print Err.description            ' descripción error afip
    Debug.Print Err.Number - vbObjectError ' codigo error afip
    If strExcepcion = "" Then                 ' si no tengo mensaje de excepcion
        strExcepcion = Err.description        ' uso el error de VB
    End If
    
    ' Mostrar el mensaje de error
    Select Case MsgBox(strExcepcion, vbCritical + vbRetryCancel, "Error:" & Err.Number - vbObjectError & " en " & Err.Source)
        Case vbRetry
            Debug.Assert False
            Resume
        Case vbCancel
            Debug.Print Err.description
    End Select
    
End Sub

Private Sub controlarExcepcion(obj As Object)
Dim intFD As Integer

    ' Nueva funcion para verificar que no haya habido errores:
    On Error GoTo 0
    
    If obj.Excepcion <> "" Then
        ' Depuración (grabar a un archivo los detalles del error)
        intFD = FreeFile
        Open App.path & "\excepcion.txt" For Append As intFD
        Print #intFD, obj.Excepcion
        Print #intFD, obj.Traceback
        Print #intFD, obj.XmlRequest
        Print #intFD, obj.XmlResponse
        Close intFD
        MsgBox obj.Excepcion, vbExclamation, "Excepción"
        End
    End If
    
End Sub

