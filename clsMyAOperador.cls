VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyAOperador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vOperadorID As Integer
Private vRazonSocial As String
Private vCalle As String
Private vPuerta As String
Private vPiso As String
Private vDpto As String
Private vCodigoPostal As Integer
Private vLocalidad As String
Private vProvincia As String
Private vTelefono As String
Private vCuit As String
Private vIngresosBrutos As String
Private vSituacionIva As Integer
Private vNumeroEpas As String
Private vFechaInicio As Date
Private vServicio As Integer
Private vPuntoVta As Integer
Private vNroComprob As Long
Private vPeriodoFactura As Integer
Private vResolucion As String
Private vPersoneria As String
Private vReciboSerie As Integer
Private vRecibo As Long
Private vNCreditoSerie As Integer
Private vNCredito As Long
Private vCai As Variant
Private vCaiVencimiento As Variant
Private vPreimpreso As Integer
Private vUid As String
Private vCreated As Date

Private clon As clsMyAOperador

Public ddl As New clsMyDOperador

Private Sub class_Initialize()

    vOperadorID = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vRazonSocial = ""
    vCalle = ""
    vPuerta = ""
    vPiso = ""
    vDpto = ""
    vCodigoPostal = 0
    vLocalidad = ""
    vProvincia = ""
    vTelefono = ""
    vCuit = ""
    vIngresosBrutos = ""
    vSituacionIva = 0
    vNumeroEpas = ""
    vFechaInicio = Date
    vServicio = 0
    vPuntoVta = 0
    vNroComprob = 0
    vPeriodoFactura = 0
    vResolucion = ""
    vPersoneria = ""
    vReciboSerie = 0
    vRecibo = 0
    vNCreditoSerie = 0
    vNCredito = 0
    vCai = Null
    vCaiVencimiento = Null
    vPreimpreso = 0
    vUid = ""
    vCreated = Date

End Sub

Public Sub clean()

    Call class_Initialize

End Sub

Public Property Get operadorID() As Variant

    operadorID = vOperadorID

End Property

Public Property Let operadorID(ByVal vNewValue As Variant)

On Error Resume Next

    vOperadorID = vNewValue

End Property

Public Property Get razonSocial() As Variant

    razonSocial = vRazonSocial

End Property

Public Property Let razonSocial(ByVal vNewValue As Variant)

On Error Resume Next

    vRazonSocial = vNewValue

End Property

Public Property Get calle() As Variant

    calle = vCalle

End Property

Public Property Let calle(ByVal vNewValue As Variant)

On Error Resume Next

    vCalle = vNewValue

End Property

Public Property Get puerta() As Variant

    puerta = vPuerta

End Property

Public Property Let puerta(ByVal vNewValue As Variant)

On Error Resume Next

    vPuerta = vNewValue

End Property

Public Property Get piso() As Variant

    piso = vPiso

End Property

Public Property Let piso(ByVal vNewValue As Variant)

On Error Resume Next

    vPiso = vNewValue

End Property

Public Property Get dpto() As Variant

    dpto = vDpto

End Property

Public Property Let dpto(ByVal vNewValue As Variant)

On Error Resume Next

    vDpto = vNewValue

End Property

Public Property Get codigoPostal() As Variant

    codigoPostal = vCodigoPostal

End Property

Public Property Let codigoPostal(ByVal vNewValue As Variant)

On Error Resume Next

    vCodigoPostal = vNewValue

End Property

Public Property Get localidad() As Variant

    localidad = vLocalidad

End Property

Public Property Let localidad(ByVal vNewValue As Variant)

On Error Resume Next

    vLocalidad = vNewValue

End Property

Public Property Get provincia() As Variant

    provincia = vProvincia

End Property

Public Property Let provincia(ByVal vNewValue As Variant)

On Error Resume Next

    vProvincia = vNewValue

End Property

Public Property Get telefono() As Variant

    telefono = vTelefono

End Property

Public Property Let telefono(ByVal vNewValue As Variant)

On Error Resume Next

    vTelefono = vNewValue

End Property

Public Property Get cuit() As Variant

    cuit = vCuit

End Property

Public Property Let cuit(ByVal vNewValue As Variant)

On Error Resume Next

    vCuit = vNewValue

End Property

Public Property Get ingresosBrutos() As Variant

    ingresosBrutos = vIngresosBrutos

End Property

Public Property Let ingresosBrutos(ByVal vNewValue As Variant)

On Error Resume Next

    vIngresosBrutos = vNewValue

End Property

Public Property Get situacionIVA() As Variant

    situacionIVA = vSituacionIva

End Property

Public Property Let situacionIVA(ByVal vNewValue As Variant)

On Error Resume Next

    vSituacionIva = vNewValue

End Property

Public Property Get numeroEpas() As Variant

    numeroEpas = vNumeroEpas

End Property

Public Property Let numeroEpas(ByVal vNewValue As Variant)

On Error Resume Next

    vNumeroEpas = vNewValue

End Property

Public Property Get fechaInicio() As Variant

    fechaInicio = vFechaInicio

End Property

Public Property Let fechaInicio(ByVal vNewValue As Variant)

On Error Resume Next

    vFechaInicio = vNewValue

End Property

Public Property Get servicio() As Variant

    servicio = vServicio

End Property

Public Property Let servicio(ByVal vNewValue As Variant)

On Error Resume Next

    vServicio = vNewValue

End Property

Public Property Get puntoVta() As Variant

    puntoVta = vPuntoVta

End Property

Public Property Let puntoVta(ByVal vNewValue As Variant)

On Error Resume Next

    vPuntoVta = vNewValue

End Property

Public Property Get nroComprob() As Variant

    nroComprob = vNroComprob

End Property

Public Property Let nroComprob(ByVal vNewValue As Variant)

On Error Resume Next

    vNroComprob = vNewValue

End Property

Public Property Get periodoFactura() As Variant

    periodoFactura = vPeriodoFactura

End Property

Public Property Let periodoFactura(ByVal vNewValue As Variant)

On Error Resume Next

    vPeriodoFactura = vNewValue

End Property

Public Property Get resolucion() As Variant

    resolucion = vResolucion

End Property

Public Property Let resolucion(ByVal vNewValue As Variant)

On Error Resume Next

    vResolucion = vNewValue

End Property

Public Property Get personeria() As Variant

    personeria = vPersoneria

End Property

Public Property Let personeria(ByVal vNewValue As Variant)

On Error Resume Next

    vPersoneria = vNewValue

End Property

Public Property Get reciboSerie() As Variant

    reciboSerie = vReciboSerie

End Property

Public Property Let reciboSerie(ByVal vNewValue As Variant)

On Error Resume Next

    vReciboSerie = vNewValue

End Property

Public Property Get recibo() As Variant

    recibo = vRecibo

End Property

Public Property Let recibo(ByVal vNewValue As Variant)

On Error Resume Next

    vRecibo = vNewValue

End Property

Public Property Get ncreditoSerie() As Variant

    ncreditoSerie = vNCreditoSerie

End Property

Public Property Let ncreditoSerie(ByVal vNewValue As Variant)

On Error Resume Next

    vNCreditoSerie = vNewValue

End Property

Public Property Get ncredito() As Variant

    ncredito = vNCredito

End Property

Public Property Let ncredito(ByVal vNewValue As Variant)

On Error Resume Next

    vNCredito = vNewValue

End Property

Public Property Get cai() As Variant

    cai = vCai

End Property

Public Property Let cai(ByVal vNewValue As Variant)

On Error Resume Next

    vCai = vNewValue

End Property

Public Property Get caiVencimiento() As Variant

    caiVencimiento = vCaiVencimiento

End Property

Public Property Let caiVencimiento(ByVal vNewValue As Variant)

On Error Resume Next

    vCaiVencimiento = vNewValue

End Property

Public Property Get preimpreso() As Variant

    preimpreso = vPreimpreso

End Property

Public Property Let preimpreso(ByVal vNewValue As Variant)

On Error Resume Next

    vPreimpreso = vNewValue

End Property

Public Property Get uid() As Variant

    uid = vUid

End Property

Public Property Let uid(ByVal vNewValue As Variant)

On Error Resume Next

    vUid = vNewValue

End Property

Public Property Get created() As Variant

    created = vCreated

End Property

Public Property Let created(ByVal vNewValue As Variant)

On Error Resume Next

    vCreated = vNewValue

End Property

Public Property Get comboText() As Variant

    comboText = vRazonSocial

End Property

Public Property Get comboKey() As Variant

    comboKey = vOperadorID

End Property

Public Property Get listText() As Variant

    listText = vRazonSocial

End Property

Public Property Get listKey() As Variant

    listKey = vOperadorID

End Property

Public Function clone() As clsMyAOperador
Dim objObjeto As New clsMyAOperador

    With objObjeto
        .operadorID = vOperadorID
        .razonSocial = vRazonSocial
        .calle = vCalle
        .puerta = vPuerta
        .piso = vPiso
        .dpto = vDpto
        .codigoPostal = vCodigoPostal
        .localidad = vLocalidad
        .provincia = vProvincia
        .telefono = vTelefono
        .cuit = vCuit
        .ingresosBrutos = vIngresosBrutos
        .situacionIVA = vSituacionIva
        .numeroEpas = vNumeroEpas
        .fechaInicio = vFechaInicio
        .servicio = vServicio
        .puntoVta = vPuntoVta
        .nroComprob = vNroComprob
        .periodoFactura = vPeriodoFactura
        .resolucion = vResolucion
        .personeria = vPersoneria
        .reciboSerie = vReciboSerie
        .recibo = vRecibo
        .ncreditoSerie = vNCreditoSerie
        .ncredito = vNCredito
        .cai = vCai
        .caiVencimiento = vCaiVencimiento
        .preimpreso = vPreimpreso
        .uid = vUid
        .created = vCreated
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsMyAOperador, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .operadorID = pQuery.Fields(.ddl.operadorID)
        .razonSocial = pQuery.Fields(.ddl.razonSocial)
        .calle = pQuery.Fields(.ddl.calle)
        .puerta = pQuery.Fields(.ddl.puerta)
        .piso = pQuery.Fields(.ddl.piso)
        .dpto = pQuery.Fields(.ddl.dpto)
        .codigoPostal = pQuery.Fields(.ddl.codigoPostal)
        .localidad = pQuery.Fields(.ddl.localidad)
        .provincia = pQuery.Fields(.ddl.provincia)
        .telefono = pQuery.Fields(.ddl.telefono)
        .cuit = pQuery.Fields(.ddl.cuit)
        .ingresosBrutos = pQuery.Fields(.ddl.ingresosBrutos)
        .situacionIVA = pQuery.Fields(.ddl.situacionIVA)
        .numeroEpas = pQuery.Fields(.ddl.numeroEpas)
        .fechaInicio = pQuery.Fields(.ddl.fechaInicio)
        .servicio = pQuery.Fields(.ddl.servicio)
        .puntoVta = pQuery.Fields(.ddl.puntoVta)
        .nroComprob = pQuery.Fields(.ddl.nroComprob)
        .periodoFactura = pQuery.Fields(.ddl.periodoFactura)
        .resolucion = pQuery.Fields(.ddl.resolucion)
        .personeria = pQuery.Fields(.ddl.personeria)
        .reciboSerie = pQuery.Fields(.ddl.reciboSerie)
        .recibo = pQuery.Fields(.ddl.recibo)
        .ncreditoSerie = pQuery.Fields(.ddl.ncreditoSerie)
        .ncredito = pQuery.Fields(.ddl.ncredito)
        .cai = pQuery.Fields(.ddl.cai)
        .caiVencimiento = pQuery.Fields(.ddl.caiVencimiento)
        .preimpreso = pQuery.Fields(.ddl.preimpreso)
        .uid = pQuery.Fields(.ddl.uid)
        .created = pQuery.Fields(.ddl.created)
    End With

End Sub

Private Function makeWherePrimaryKey() As String
Dim sql As String

    sql = ""
    sql = sql & " WHERE " & ddl.operadorID & " = " & vOperadorID

    makeWherePrimaryKey = sql

End Function

Private Sub findSQL(pSQL As String, db As clsDB)
Dim rstQuery As ADODB.Recordset

    Set rstQuery = db.query(pSQL)

    If Not rstQuery.EOF Then
        fillObject Me, rstQuery

        Set clon = Me.clone
    End If

    rstQuery.Close

End Sub

Public Sub findByPrimaryKey(db As clsDB)
Dim sql As String

    initializeOther

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & makeWherePrimaryKey
    sql = sql & ";"

    findSQL sql, db

End Sub

Public Sub findLast(db As clsDB)
Dim sql As String

    initializeOther

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " ORDER BY " & ddl.operadorID & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, db

End Sub

Public Function delete(db As clsDB) As Boolean
Dim sql As String

    sql = "DELETE FROM " & ddl.table
    sql = sql & makeWherePrimaryKey
    sql = sql & ";"

    delete = db.execute(sql)

End Function

Public Function save(db As clsDB) As Boolean

    If exist(db) Then
        save = update(db)
    Else
        save = add(db)
    End If

End Function

Public Function exist(db As clsDB) As Boolean
Dim sql As String

Dim rstQuery As ADODB.Recordset

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & makeWherePrimaryKey
    sql = sql & ";"

    Set rstQuery = db.query(sql)
    exist = Not rstQuery.EOF
    rstQuery.Close

End Function

Public Function add(db As clsDB) As Boolean
Dim sql As String

    sql = "INSERT INTO " & ddl.table
    sql = sql & " SET " & ddl.razonSocial & " = '" & replaceSpecialSymbols(vRazonSocial) & "'"
    sql = sql & ", " & ddl.calle & " = '" & replaceSpecialSymbols(vCalle) & "'"
    sql = sql & ", " & ddl.puerta & " = '" & replaceSpecialSymbols(vPuerta) & "'"
    sql = sql & ", " & ddl.piso & " = '" & replaceSpecialSymbols(vPiso) & "'"
    sql = sql & ", " & ddl.dpto & " = '" & replaceSpecialSymbols(vDpto) & "'"
    sql = sql & ", " & ddl.codigoPostal & " = " & vCodigoPostal
    sql = sql & ", " & ddl.localidad & " = '" & replaceSpecialSymbols(vLocalidad) & "'"
    sql = sql & ", " & ddl.provincia & " = '" & replaceSpecialSymbols(vProvincia) & "'"
    sql = sql & ", " & ddl.telefono & " = '" & replaceSpecialSymbols(vTelefono) & "'"
    sql = sql & ", " & ddl.cuit & " = '" & replaceSpecialSymbols(vCuit) & "'"
    sql = sql & ", " & ddl.ingresosBrutos & " = '" & replaceSpecialSymbols(vIngresosBrutos) & "'"
    sql = sql & ", " & ddl.situacionIVA & " = " & vSituacionIva
    sql = sql & ", " & ddl.numeroEpas & " = '" & replaceSpecialSymbols(vNumeroEpas) & "'"
    sql = sql & ", " & ddl.fechaInicio & " = " & fechaDB(vFechaInicio)
    sql = sql & ", " & ddl.servicio & " = " & vServicio
    sql = sql & ", " & ddl.puntoVta & " = " & vPuntoVta
    sql = sql & ", " & ddl.nroComprob & " = " & vNroComprob
    sql = sql & ", " & ddl.periodoFactura & " = " & vPeriodoFactura
    sql = sql & ", " & ddl.resolucion & " = '" & replaceSpecialSymbols(vResolucion) & "'"
    sql = sql & ", " & ddl.personeria & " = '" & replaceSpecialSymbols(vPersoneria) & "'"
    sql = sql & ", " & ddl.reciboSerie & " = " & vReciboSerie
    sql = sql & ", " & ddl.recibo & " = " & vRecibo
    sql = sql & ", " & ddl.ncreditoSerie & " = " & vNCreditoSerie
    sql = sql & ", " & ddl.ncredito & " = " & vNCredito
    sql = sql & ", " & ddl.cai & " = " & validateNullField(vCai)
    sql = sql & ", " & ddl.caiVencimiento & " = " & fechaDB(vCaiVencimiento)
    sql = sql & ", " & ddl.preimpreso & " = " & vPreimpreso
    sql = sql & ", " & ddl.uid & " = '" & replaceSpecialSymbols(vUid) & "'"
    sql = sql & ", " & ddl.operadorID & " = " & vOperadorID
    sql = sql & ";"

    add = db.execute(sql)

    findByPrimaryKey db

End Function

Public Function update(db As clsDB) As Boolean
Dim sql As String
Dim strSET As String

Dim blnChange As Boolean

    update = True

    strSET = ""
    blnChange = False

    If clon Is Nothing Then
        Set clon = Me.clone
        clon.findByPrimaryKey db
    End If

    If vRazonSocial <> clon.razonSocial Or IsNull(vRazonSocial) <> IsNull(clon.razonSocial) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.razonSocial & " = '" & replaceSpecialSymbols(vRazonSocial) & "'"
        blnChange = True
    End If
    If vCalle <> clon.calle Or IsNull(vCalle) <> IsNull(clon.calle) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.calle & " = '" & replaceSpecialSymbols(vCalle) & "'"
        blnChange = True
    End If
    If vPuerta <> clon.puerta Or IsNull(vPuerta) <> IsNull(clon.puerta) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.puerta & " = '" & replaceSpecialSymbols(vPuerta) & "'"
        blnChange = True
    End If
    If vPiso <> clon.piso Or IsNull(vPiso) <> IsNull(clon.piso) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.piso & " = '" & replaceSpecialSymbols(vPiso) & "'"
        blnChange = True
    End If
    If vDpto <> clon.dpto Or IsNull(vDpto) <> IsNull(clon.dpto) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.dpto & " = '" & replaceSpecialSymbols(vDpto) & "'"
        blnChange = True
    End If
    If vCodigoPostal <> clon.codigoPostal Or IsNull(vCodigoPostal) <> IsNull(clon.codigoPostal) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.codigoPostal & " = " & vCodigoPostal
        blnChange = True
    End If
    If vLocalidad <> clon.localidad Or IsNull(vLocalidad) <> IsNull(clon.localidad) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.localidad & " = '" & replaceSpecialSymbols(vLocalidad) & "'"
        blnChange = True
    End If
    If vProvincia <> clon.provincia Or IsNull(vProvincia) <> IsNull(clon.provincia) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.provincia & " = '" & replaceSpecialSymbols(vProvincia) & "'"
        blnChange = True
    End If
    If vTelefono <> clon.telefono Or IsNull(vTelefono) <> IsNull(clon.telefono) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.telefono & " = '" & replaceSpecialSymbols(vTelefono) & "'"
        blnChange = True
    End If
    If vCuit <> clon.cuit Or IsNull(vCuit) <> IsNull(clon.cuit) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.cuit & " = '" & replaceSpecialSymbols(vCuit) & "'"
        blnChange = True
    End If
    If vIngresosBrutos <> clon.ingresosBrutos Or IsNull(vIngresosBrutos) <> IsNull(clon.ingresosBrutos) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.ingresosBrutos & " = '" & replaceSpecialSymbols(vIngresosBrutos) & "'"
        blnChange = True
    End If
    If vSituacionIva <> clon.situacionIVA Or IsNull(vSituacionIva) <> IsNull(clon.situacionIVA) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.situacionIVA & " = " & vSituacionIva
        blnChange = True
    End If
    If vNumeroEpas <> clon.numeroEpas Or IsNull(vNumeroEpas) <> IsNull(clon.numeroEpas) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.numeroEpas & " = '" & replaceSpecialSymbols(vNumeroEpas) & "'"
        blnChange = True
    End If
    If vFechaInicio <> clon.fechaInicio Or IsNull(vFechaInicio) <> IsNull(clon.fechaInicio) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fechaInicio & " = " & fechaDB(vFechaInicio)
        blnChange = True
    End If
    If vServicio <> clon.servicio Or IsNull(vServicio) <> IsNull(clon.servicio) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.servicio & " = " & vServicio
        blnChange = True
    End If
    If vPuntoVta <> clon.puntoVta Or IsNull(vPuntoVta) <> IsNull(clon.puntoVta) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.puntoVta & " = " & vPuntoVta
        blnChange = True
    End If
    If vNroComprob <> clon.nroComprob Or IsNull(vNroComprob) <> IsNull(clon.nroComprob) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.nroComprob & " = " & vNroComprob
        blnChange = True
    End If
    If vPeriodoFactura <> clon.periodoFactura Or IsNull(vPeriodoFactura) <> IsNull(clon.periodoFactura) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.periodoFactura & " = " & vPeriodoFactura
        blnChange = True
    End If
    If vResolucion <> clon.resolucion Or IsNull(vResolucion) <> IsNull(clon.resolucion) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.resolucion & " = '" & replaceSpecialSymbols(vResolucion) & "'"
        blnChange = True
    End If
    If vPersoneria <> clon.personeria Or IsNull(vPersoneria) <> IsNull(clon.personeria) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.personeria & " = '" & replaceSpecialSymbols(vPersoneria) & "'"
        blnChange = True
    End If
    If vReciboSerie <> clon.reciboSerie Or IsNull(vReciboSerie) <> IsNull(clon.reciboSerie) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.reciboSerie & " = " & vReciboSerie
        blnChange = True
    End If
    If vRecibo <> clon.recibo Or IsNull(vRecibo) <> IsNull(clon.recibo) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.recibo & " = " & vRecibo
        blnChange = True
    End If
    If vNCreditoSerie <> clon.ncreditoSerie Or IsNull(vNCreditoSerie) <> IsNull(clon.ncreditoSerie) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.ncreditoSerie & " = " & vNCreditoSerie
        blnChange = True
    End If
    If vNCredito <> clon.ncredito Or IsNull(vNCredito) <> IsNull(clon.ncredito) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.ncredito & " = " & vNCredito
        blnChange = True
    End If
    If vCai <> clon.cai Or IsNull(vCai) <> IsNull(clon.cai) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.cai & " = " & validateNullField(vCai)
        blnChange = True
    End If
    If vCaiVencimiento <> clon.caiVencimiento Or IsNull(vCaiVencimiento) <> IsNull(clon.caiVencimiento) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.caiVencimiento & " = " & fechaDB(vCaiVencimiento)
        blnChange = True
    End If
    If vPreimpreso <> clon.preimpreso Or IsNull(vPreimpreso) <> IsNull(clon.preimpreso) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.preimpreso & " = " & vPreimpreso
        blnChange = True
    End If
    If vUid <> clon.uid Or IsNull(vUid) <> IsNull(clon.uid) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.uid & " = '" & replaceSpecialSymbols(vUid) & "'"
        blnChange = True
    End If

    sql = "UPDATE " & ddl.table & " SET" & strSET
    sql = sql & makeWherePrimaryKey
    sql = sql & ";"

    If blnChange Then update = db.execute(sql)

    Set clon = Nothing
    Set clon = Me.clone

End Function

Public Sub newID(pFind As Boolean, db As clsDB)

    newID_ pFind, IIf(IsMissing(db), dbapp, db)

End Sub

Private Sub newID_(pFind As Boolean, db As clsDB)
Dim sql As String

Dim rstQuery As ADODB.Recordset

    vOperadorID = 1

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " ORDER BY " & ddl.operadorID & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set rstQuery = db.query(sql)

    If Not rstQuery.EOF Then vOperadorID = 1 + rstQuery.Fields(ddl.operadorID)

    rstQuery.Close

    If pFind Then findByPrimaryKey db

End Sub

Public Function collectionAll(db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & ";"

    Set collectionAll = collectionSQL(sql, db)

End Function

Public Function collectionSQL(pSQL As String, db As clsDB) As Collection

    Set collectionSQL = collection_(pSQL, IIf(IsMissing(db), dbapp, db))

End Function

Private Function collection_(pSQL As String, db As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim elemento As clsMyAOperador

Dim registros As New Collection

    Set rstQuery = db.query(pSQL)

    Do While Not rstQuery.EOF
        Set elemento = New clsMyAOperador
        fillObject elemento, rstQuery

        registros.add elemento

        rstQuery.MoveNext
    Loop

    rstQuery.Close

    Set collection_ = registros

End Function

Public Sub fillCombo(pCombo As ComboBox, db As clsDB)

    fillCombo_ pCombo, collectionAll(db)

End Sub

Private Sub fillCombo_(pCombo As ComboBox, pCollection As Collection)
Dim elemento As clsMyAOperador

    pCombo.Clear

    For Each elemento In pCollection
        pCombo.AddItem elemento.comboText
        pCombo.ItemData(pCombo.NewIndex) = elemento.comboKey
    Next

    If pCombo.ListCount > 0 Then pCombo.ListIndex = 0

End Sub

Public Sub fillList(pList As ListBox, db As clsDB)

    fillList_ pList, collectionAll(db)

End Sub

Private Sub fillList_(pList As ListBox, pCollection As Collection)
Dim elemento As clsMyAOperador

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

