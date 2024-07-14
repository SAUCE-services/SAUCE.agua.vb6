VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyACliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vClienteID As Long
Private vFechaAlta As Date
Private vFechaBaja As Variant
Private vApellido As String
Private vNombre As String
Private vNumeroSocio As Variant
Private vInmuebleCalle As String
Private vInmueblePuerta As String
Private vInmueblePiso As String
Private vInmuebleDpto As String
Private vInmuebleLocalidad As String
Private vInmuebleProvincia As String
Private vInmuebleCodpostal As Integer
Private vFiscalCalle As String
Private vFiscalPuerta As String
Private vFiscalPiso As String
Private vFiscalDpto As String
Private vFiscalLocalidad As String
Private vFiscalProvincia As String
Private vFiscalCodpostal As Integer
Private vCuit As String
Private vSituacionIva As Integer
Private vNomenclaturaCatastral As String
Private vCategoria As Integer
Private vServicio As Integer
Private vCobro As Integer
Private vZona As Integer
Private vRuta As Integer
Private vOrden As Integer
Private vEstadoID As Integer
Private vFechaNacimiento As Variant
Private vCategoriaSocioID As Integer
Private vDestinoID As Integer
Private vAutoID As Long
Private vUid As String
Private vCreated As Date

Private clon As clsMyACliente

Public ddl As New clsMyDCliente

Private Sub class_Initialize()

    vClienteID = 0
    vFechaAlta = Date

    initializeOther

End Sub

Private Sub initializeOther()

    vFechaBaja = Null
    vApellido = ""
    vNombre = ""
    vNumeroSocio = Null
    vInmuebleCalle = ""
    vInmueblePuerta = ""
    vInmueblePiso = ""
    vInmuebleDpto = ""
    vInmuebleLocalidad = ""
    vInmuebleProvincia = ""
    vInmuebleCodpostal = 0
    vFiscalCalle = ""
    vFiscalPuerta = ""
    vFiscalPiso = ""
    vFiscalDpto = ""
    vFiscalLocalidad = ""
    vFiscalProvincia = ""
    vFiscalCodpostal = 0
    vCuit = ""
    vSituacionIva = 0
    vNomenclaturaCatastral = ""
    vCategoria = 0
    vServicio = 0
    vCobro = 0
    vZona = 0
    vRuta = 0
    vOrden = 0
    vEstadoID = 0
    vFechaNacimiento = Null
    vCategoriaSocioID = 0
    vDestinoID = 0
    vAutoID = 0
    vUid = ""
    vCreated = Date

End Sub

Public Sub clean()

    Call class_Initialize

End Sub

Public Property Get clienteID() As Variant

    clienteID = vClienteID

End Property

Public Property Let clienteID(ByVal vNewValue As Variant)

On Error Resume Next

    vClienteID = vNewValue

End Property

Public Property Get keyFound() As Variant
    
    keyFound = vClienteID
    
End Property

Public Property Get key() As String

    key = "k." & vClienteID
    
End Property

Public Property Get fechaAlta() As Variant

    fechaAlta = vFechaAlta

End Property

Public Property Let fechaAlta(ByVal vNewValue As Variant)

On Error Resume Next

    vFechaAlta = vNewValue

End Property

Public Property Get fechaBaja() As Variant

    fechaBaja = vFechaBaja

End Property

Public Property Let fechaBaja(ByVal vNewValue As Variant)

On Error Resume Next

    vFechaBaja = vNewValue

End Property

Public Property Get apellido() As Variant

    apellido = vApellido

End Property

Public Property Let apellido(ByVal vNewValue As Variant)

On Error Resume Next

    vApellido = vNewValue

End Property

Public Property Get nombre() As Variant

    nombre = vNombre

End Property

Public Property Get textFound() As Variant

    textFound = vApellido & ", " & vNombre
    
End Property

Public Property Let nombre(ByVal vNewValue As Variant)

On Error Resume Next

    vNombre = vNewValue

End Property

Public Property Get apellidonombre() As String
    
    apellidonombre = apellido & ", " & nombre
    
End Property

Public Property Get numeroSocio() As Variant

    numeroSocio = vNumeroSocio

End Property

Public Property Let numeroSocio(ByVal vNewValue As Variant)

On Error Resume Next

    vNumeroSocio = vNewValue

End Property

Public Property Get inmuebleCalle() As Variant

    inmuebleCalle = vInmuebleCalle

End Property

Public Property Let inmuebleCalle(ByVal vNewValue As Variant)

On Error Resume Next

    vInmuebleCalle = vNewValue

End Property

Public Property Get inmueblePuerta() As Variant

    inmueblePuerta = vInmueblePuerta

End Property

Public Property Let inmueblePuerta(ByVal vNewValue As Variant)

On Error Resume Next

    vInmueblePuerta = vNewValue

End Property

Public Property Get inmueblePiso() As Variant

    inmueblePiso = vInmueblePiso

End Property

Public Property Let inmueblePiso(ByVal vNewValue As Variant)

On Error Resume Next

    vInmueblePiso = vNewValue

End Property

Public Property Get inmuebleDpto() As Variant

    inmuebleDpto = vInmuebleDpto

End Property

Public Property Let inmuebleDpto(ByVal vNewValue As Variant)

On Error Resume Next

    vInmuebleDpto = vNewValue

End Property

Public Property Get inmuebleLocalidad() As Variant

    inmuebleLocalidad = vInmuebleLocalidad

End Property

Public Property Let inmuebleLocalidad(ByVal vNewValue As Variant)

On Error Resume Next

    vInmuebleLocalidad = vNewValue

End Property

Public Property Get inmuebleProvincia() As Variant

    inmuebleProvincia = vInmuebleProvincia

End Property

Public Property Let inmuebleProvincia(ByVal vNewValue As Variant)

On Error Resume Next

    vInmuebleProvincia = vNewValue

End Property

Public Property Get inmuebleCodpostal() As Variant

    inmuebleCodpostal = vInmuebleCodpostal

End Property

Public Property Let inmuebleCodpostal(ByVal vNewValue As Variant)

On Error Resume Next

    vInmuebleCodpostal = vNewValue

End Property

Public Property Get fiscalCalle() As Variant

    fiscalCalle = vFiscalCalle

End Property

Public Property Let fiscalCalle(ByVal vNewValue As Variant)

On Error Resume Next

    vFiscalCalle = vNewValue

End Property

Public Property Get fiscalPuerta() As Variant

    fiscalPuerta = vFiscalPuerta

End Property

Public Property Let fiscalPuerta(ByVal vNewValue As Variant)

On Error Resume Next

    vFiscalPuerta = vNewValue

End Property

Public Property Get fiscalPiso() As Variant

    fiscalPiso = vFiscalPiso

End Property

Public Property Let fiscalPiso(ByVal vNewValue As Variant)

On Error Resume Next

    vFiscalPiso = vNewValue

End Property

Public Property Get fiscalDpto() As Variant

    fiscalDpto = vFiscalDpto

End Property

Public Property Let fiscalDpto(ByVal vNewValue As Variant)

On Error Resume Next

    vFiscalDpto = vNewValue

End Property

Public Property Get fiscalLocalidad() As Variant

    fiscalLocalidad = vFiscalLocalidad

End Property

Public Property Let fiscalLocalidad(ByVal vNewValue As Variant)

On Error Resume Next

    vFiscalLocalidad = vNewValue

End Property

Public Property Get fiscalProvincia() As Variant

    fiscalProvincia = vFiscalProvincia

End Property

Public Property Let fiscalProvincia(ByVal vNewValue As Variant)

On Error Resume Next

    vFiscalProvincia = vNewValue

End Property

Public Property Get fiscalCodpostal() As Variant

    fiscalCodpostal = vFiscalCodpostal

End Property

Public Property Let fiscalCodpostal(ByVal vNewValue As Variant)

On Error Resume Next

    vFiscalCodpostal = vNewValue

End Property

Public Property Get cuit() As Variant

    cuit = vCuit

End Property

Public Property Let cuit(ByVal vNewValue As Variant)

On Error Resume Next

    vCuit = vNewValue

End Property

Public Property Get situacionIVA() As Variant

    situacionIVA = vSituacionIva

End Property

Public Property Let situacionIVA(ByVal vNewValue As Variant)

On Error Resume Next

    vSituacionIva = vNewValue

End Property

Public Property Get nomenclaturaCatastral() As Variant

    nomenclaturaCatastral = vNomenclaturaCatastral

End Property

Public Property Let nomenclaturaCatastral(ByVal vNewValue As Variant)

On Error Resume Next

    vNomenclaturaCatastral = vNewValue

End Property

Public Property Get categoria() As Variant

    categoria = vCategoria

End Property

Public Property Let categoria(ByVal vNewValue As Variant)

On Error Resume Next

    vCategoria = vNewValue

End Property

Public Property Get servicio() As Variant

    servicio = vServicio

End Property

Public Property Let servicio(ByVal vNewValue As Variant)

On Error Resume Next

    vServicio = vNewValue

End Property

Public Property Get cobro() As Variant

    cobro = vCobro

End Property

Public Property Let cobro(ByVal vNewValue As Variant)

On Error Resume Next

    vCobro = vNewValue

End Property

Public Property Get zona() As Variant

    zona = vZona

End Property

Public Property Let zona(ByVal vNewValue As Variant)

On Error Resume Next

    vZona = vNewValue

End Property

Public Property Get ruta() As Variant

    ruta = vRuta

End Property

Public Property Let ruta(ByVal vNewValue As Variant)

On Error Resume Next

    vRuta = vNewValue

End Property

Public Property Get orden() As Variant

    orden = vOrden

End Property

Public Property Let orden(ByVal vNewValue As Variant)

On Error Resume Next

    vOrden = vNewValue

End Property

Public Property Get estadoID() As Variant

    estadoID = vEstadoID

End Property

Public Property Let estadoID(ByVal vNewValue As Variant)

On Error Resume Next

    vEstadoID = vNewValue

End Property

Public Property Get fechaNacimiento() As Variant

    fechaNacimiento = vFechaNacimiento

End Property

Public Property Let fechaNacimiento(ByVal vNewValue As Variant)

On Error Resume Next

    vFechaNacimiento = vNewValue

End Property

Public Property Get categoriasocioID() As Variant

    categoriasocioID = vCategoriaSocioID

End Property

Public Property Let categoriasocioID(ByVal vNewValue As Variant)

On Error Resume Next

    vCategoriaSocioID = vNewValue

End Property

Public Property Get destinoID() As Variant

    destinoID = vDestinoID

End Property

Public Property Let destinoID(ByVal vNewValue As Variant)

On Error Resume Next

    vDestinoID = vNewValue

End Property

Public Property Get autoID() As Variant

    autoID = vAutoID

End Property

Public Property Let autoID(ByVal vNewValue As Variant)

On Error Resume Next

    vAutoID = vNewValue

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

    comboText = vApellido & ", " & vNombre

End Property

Public Property Get domicilioText() As Variant

    domicilioText = vInmuebleCalle & " " & vInmueblePuerta & " " & vInmueblePiso & " " & vInmuebleDpto & " " & vInmuebleLocalidad

End Property

Public Property Get comboKey() As Variant

    comboKey = vClienteID

End Property

Public Property Get listText() As Variant

    listText = vApellido & ", " & vNombre

End Property

Public Property Get listKey() As Variant

    listKey = vClienteID

End Property

Public Function clone() As clsMyACliente
Dim objObjeto As New clsMyACliente

    With objObjeto
        .clienteID = vClienteID
        .fechaAlta = vFechaAlta
        .fechaBaja = vFechaBaja
        .apellido = vApellido
        .nombre = vNombre
        .numeroSocio = vNumeroSocio
        .inmuebleCalle = vInmuebleCalle
        .inmueblePuerta = vInmueblePuerta
        .inmueblePiso = vInmueblePiso
        .inmuebleDpto = vInmuebleDpto
        .inmuebleLocalidad = vInmuebleLocalidad
        .inmuebleProvincia = vInmuebleProvincia
        .inmuebleCodpostal = vInmuebleCodpostal
        .fiscalCalle = vFiscalCalle
        .fiscalPuerta = vFiscalPuerta
        .fiscalPiso = vFiscalPiso
        .fiscalDpto = vFiscalDpto
        .fiscalLocalidad = vFiscalLocalidad
        .fiscalProvincia = vFiscalProvincia
        .fiscalCodpostal = vFiscalCodpostal
        .cuit = vCuit
        .situacionIVA = vSituacionIva
        .nomenclaturaCatastral = vNomenclaturaCatastral
        .categoria = vCategoria
        .servicio = vServicio
        .cobro = vCobro
        .zona = vZona
        .ruta = vRuta
        .orden = vOrden
        .estadoID = vEstadoID
        .fechaNacimiento = vFechaNacimiento
        .categoriasocioID = vCategoriaSocioID
        .destinoID = vDestinoID
        .autoID = vAutoID
        .uid = vUid
        .created = vCreated
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsMyACliente, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .clienteID = pQuery.Fields(.ddl.clienteID)
        .fechaAlta = pQuery.Fields(.ddl.fechaAlta)
        .fechaBaja = pQuery.Fields(.ddl.fechaBaja)
        .apellido = pQuery.Fields(.ddl.apellido)
        .nombre = pQuery.Fields(.ddl.nombre)
        .numeroSocio = pQuery.Fields(.ddl.numeroSocio)
        .inmuebleCalle = pQuery.Fields(.ddl.inmuebleCalle)
        .inmueblePuerta = pQuery.Fields(.ddl.inmueblePuerta)
        .inmueblePiso = pQuery.Fields(.ddl.inmueblePiso)
        .inmuebleDpto = pQuery.Fields(.ddl.inmuebleDpto)
        .inmuebleLocalidad = pQuery.Fields(.ddl.inmuebleLocalidad)
        .inmuebleProvincia = pQuery.Fields(.ddl.inmuebleProvincia)
        .inmuebleCodpostal = pQuery.Fields(.ddl.inmuebleCodpostal)
        .fiscalCalle = pQuery.Fields(.ddl.fiscalCalle)
        .fiscalPuerta = pQuery.Fields(.ddl.fiscalPuerta)
        .fiscalPiso = pQuery.Fields(.ddl.fiscalPiso)
        .fiscalDpto = pQuery.Fields(.ddl.fiscalDpto)
        .fiscalLocalidad = pQuery.Fields(.ddl.fiscalLocalidad)
        .fiscalProvincia = pQuery.Fields(.ddl.fiscalProvincia)
        .fiscalCodpostal = pQuery.Fields(.ddl.fiscalCodpostal)
        .cuit = pQuery.Fields(.ddl.cuit)
        .situacionIVA = pQuery.Fields(.ddl.situacionIVA)
        .nomenclaturaCatastral = pQuery.Fields(.ddl.nomenclaturaCatastral)
        .categoria = pQuery.Fields(.ddl.categoria)
        .servicio = pQuery.Fields(.ddl.servicio)
        .cobro = pQuery.Fields(.ddl.cobro)
        .zona = pQuery.Fields(.ddl.zona)
        .ruta = pQuery.Fields(.ddl.ruta)
        .orden = pQuery.Fields(.ddl.orden)
        .estadoID = pQuery.Fields(.ddl.estadoID)
        .fechaNacimiento = pQuery.Fields(.ddl.fechaNacimiento)
        .categoriasocioID = pQuery.Fields(.ddl.categoriasocioID)
        .destinoID = pQuery.Fields(.ddl.destinoID)
        .autoID = pQuery.Fields(.ddl.autoID)
        .uid = pQuery.Fields(.ddl.uid)
        .created = pQuery.Fields(.ddl.created)
    End With

End Sub

Private Function makeWherePrimaryKey() As String
Dim sql As String

    sql = ""
    sql = sql & " WHERE " & ddl.clienteID & " = " & vClienteID
    sql = sql & " AND " & ddl.fechaAlta & " = " & fechaDB(vFechaAlta)

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

Public Sub findSearch(clienteID As Long, db As clsDB)

    vClienteID = clienteID
    findLast db
    
End Sub

Public Sub findLast(db As clsDB)
Dim sql As String

    initializeOther

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteID & " = " & vClienteID
    sql = sql & " ORDER BY " & ddl.fechaAlta & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, db

End Sub

Public Sub findLastLast(db As clsDB)
Dim sql As String

    initializeOther

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " ORDER BY " & ddl.clienteID & " DESC, " & ddl.fechaAlta & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, db

End Sub

Public Sub findByAutoID(db As clsDB)
Dim sql As String

Dim lngAutoID As Long

    lngAutoID = vAutoID

    clean

    vAutoID = lngAutoID

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.autoID & " = " & vAutoID
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
    sql = sql & " SET " & ddl.fechaBaja & " = " & fechaDB(vFechaBaja)
    sql = sql & ", " & ddl.apellido & " = '" & replaceSpecialSymbols(vApellido) & "'"
    sql = sql & ", " & ddl.nombre & " = '" & replaceSpecialSymbols(vNombre) & "'"
    sql = sql & ", " & ddl.numeroSocio & " = " & validateNullField(vNumeroSocio)
    sql = sql & ", " & ddl.inmuebleCalle & " = '" & replaceSpecialSymbols(vInmuebleCalle) & "'"
    sql = sql & ", " & ddl.inmueblePuerta & " = '" & replaceSpecialSymbols(vInmueblePuerta) & "'"
    sql = sql & ", " & ddl.inmueblePiso & " = '" & replaceSpecialSymbols(vInmueblePiso) & "'"
    sql = sql & ", " & ddl.inmuebleDpto & " = '" & replaceSpecialSymbols(vInmuebleDpto) & "'"
    sql = sql & ", " & ddl.inmuebleLocalidad & " = '" & replaceSpecialSymbols(vInmuebleLocalidad) & "'"
    sql = sql & ", " & ddl.inmuebleProvincia & " = '" & replaceSpecialSymbols(vInmuebleProvincia) & "'"
    sql = sql & ", " & ddl.inmuebleCodpostal & " = " & vInmuebleCodpostal
    sql = sql & ", " & ddl.fiscalCalle & " = '" & replaceSpecialSymbols(vFiscalCalle) & "'"
    sql = sql & ", " & ddl.fiscalPuerta & " = '" & replaceSpecialSymbols(vFiscalPuerta) & "'"
    sql = sql & ", " & ddl.fiscalPiso & " = '" & replaceSpecialSymbols(vFiscalPiso) & "'"
    sql = sql & ", " & ddl.fiscalDpto & " = '" & replaceSpecialSymbols(vFiscalDpto) & "'"
    sql = sql & ", " & ddl.fiscalLocalidad & " = '" & replaceSpecialSymbols(vFiscalLocalidad) & "'"
    sql = sql & ", " & ddl.fiscalProvincia & " = '" & replaceSpecialSymbols(vFiscalProvincia) & "'"
    sql = sql & ", " & ddl.fiscalCodpostal & " = " & vFiscalCodpostal
    sql = sql & ", " & ddl.cuit & " = '" & replaceSpecialSymbols(vCuit) & "'"
    sql = sql & ", " & ddl.situacionIVA & " = " & vSituacionIva
    sql = sql & ", " & ddl.nomenclaturaCatastral & " = '" & replaceSpecialSymbols(vNomenclaturaCatastral) & "'"
    sql = sql & ", " & ddl.categoria & " = " & vCategoria
    sql = sql & ", " & ddl.servicio & " = " & vServicio
    sql = sql & ", " & ddl.cobro & " = " & vCobro
    sql = sql & ", " & ddl.zona & " = " & vZona
    sql = sql & ", " & ddl.ruta & " = " & vRuta
    sql = sql & ", " & ddl.orden & " = " & vOrden
    sql = sql & ", " & ddl.estadoID & " = " & vEstadoID
    sql = sql & ", " & ddl.fechaNacimiento & " = " & fechaDB(vFechaNacimiento)
    sql = sql & ", " & ddl.categoriasocioID & " = " & vCategoriaSocioID
    sql = sql & ", " & ddl.destinoID & " = " & vDestinoID
    sql = sql & ", " & ddl.uid & " = '" & replaceSpecialSymbols(vUid) & "'"
    sql = sql & ", " & ddl.clienteID & " = " & vClienteID
    sql = sql & ", " & ddl.fechaAlta & " = " & fechaDB(vFechaAlta)
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

    If vFechaBaja <> clon.fechaBaja Or IsNull(vFechaBaja) <> IsNull(clon.fechaBaja) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fechaBaja & " = " & fechaDB(vFechaBaja)
        blnChange = True
    End If
    If vApellido <> clon.apellido Or IsNull(vApellido) <> IsNull(clon.apellido) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.apellido & " = '" & replaceSpecialSymbols(vApellido) & "'"
        blnChange = True
    End If
    If vNombre <> clon.nombre Or IsNull(vNombre) <> IsNull(clon.nombre) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.nombre & " = '" & replaceSpecialSymbols(vNombre) & "'"
        blnChange = True
    End If
    If vNumeroSocio <> clon.numeroSocio Or IsNull(vNumeroSocio) <> IsNull(clon.numeroSocio) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.numeroSocio & " = " & validateNullField(vNumeroSocio)
        blnChange = True
    End If
    If vInmuebleCalle <> clon.inmuebleCalle Or IsNull(vInmuebleCalle) <> IsNull(clon.inmuebleCalle) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.inmuebleCalle & " = '" & replaceSpecialSymbols(vInmuebleCalle) & "'"
        blnChange = True
    End If
    If vInmueblePuerta <> clon.inmueblePuerta Or IsNull(vInmueblePuerta) <> IsNull(clon.inmueblePuerta) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.inmueblePuerta & " = '" & replaceSpecialSymbols(vInmueblePuerta) & "'"
        blnChange = True
    End If
    If vInmueblePiso <> clon.inmueblePiso Or IsNull(vInmueblePiso) <> IsNull(clon.inmueblePiso) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.inmueblePiso & " = '" & replaceSpecialSymbols(vInmueblePiso) & "'"
        blnChange = True
    End If
    If vInmuebleDpto <> clon.inmuebleDpto Or IsNull(vInmuebleDpto) <> IsNull(clon.inmuebleDpto) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.inmuebleDpto & " = '" & replaceSpecialSymbols(vInmuebleDpto) & "'"
        blnChange = True
    End If
    If vInmuebleLocalidad <> clon.inmuebleLocalidad Or IsNull(vInmuebleLocalidad) <> IsNull(clon.inmuebleLocalidad) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.inmuebleLocalidad & " = '" & replaceSpecialSymbols(vInmuebleLocalidad) & "'"
        blnChange = True
    End If
    If vInmuebleProvincia <> clon.inmuebleProvincia Or IsNull(vInmuebleProvincia) <> IsNull(clon.inmuebleProvincia) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.inmuebleProvincia & " = '" & replaceSpecialSymbols(vInmuebleProvincia) & "'"
        blnChange = True
    End If
    If vInmuebleCodpostal <> clon.inmuebleCodpostal Or IsNull(vInmuebleCodpostal) <> IsNull(clon.inmuebleCodpostal) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.inmuebleCodpostal & " = " & vInmuebleCodpostal
        blnChange = True
    End If
    If vFiscalCalle <> clon.fiscalCalle Or IsNull(vFiscalCalle) <> IsNull(clon.fiscalCalle) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fiscalCalle & " = '" & replaceSpecialSymbols(vFiscalCalle) & "'"
        blnChange = True
    End If
    If vFiscalPuerta <> clon.fiscalPuerta Or IsNull(vFiscalPuerta) <> IsNull(clon.fiscalPuerta) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fiscalPuerta & " = '" & replaceSpecialSymbols(vFiscalPuerta) & "'"
        blnChange = True
    End If
    If vFiscalPiso <> clon.fiscalPiso Or IsNull(vFiscalPiso) <> IsNull(clon.fiscalPiso) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fiscalPiso & " = '" & replaceSpecialSymbols(vFiscalPiso) & "'"
        blnChange = True
    End If
    If vFiscalDpto <> clon.fiscalDpto Or IsNull(vFiscalDpto) <> IsNull(clon.fiscalDpto) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fiscalDpto & " = '" & replaceSpecialSymbols(vFiscalDpto) & "'"
        blnChange = True
    End If
    If vFiscalLocalidad <> clon.fiscalLocalidad Or IsNull(vFiscalLocalidad) <> IsNull(clon.fiscalLocalidad) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fiscalLocalidad & " = '" & replaceSpecialSymbols(vFiscalLocalidad) & "'"
        blnChange = True
    End If
    If vFiscalProvincia <> clon.fiscalProvincia Or IsNull(vFiscalProvincia) <> IsNull(clon.fiscalProvincia) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fiscalProvincia & " = '" & replaceSpecialSymbols(vFiscalProvincia) & "'"
        blnChange = True
    End If
    If vFiscalCodpostal <> clon.fiscalCodpostal Or IsNull(vFiscalCodpostal) <> IsNull(clon.fiscalCodpostal) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fiscalCodpostal & " = " & vFiscalCodpostal
        blnChange = True
    End If
    If vCuit <> clon.cuit Or IsNull(vCuit) <> IsNull(clon.cuit) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.cuit & " = '" & replaceSpecialSymbols(vCuit) & "'"
        blnChange = True
    End If
    If vSituacionIva <> clon.situacionIVA Or IsNull(vSituacionIva) <> IsNull(clon.situacionIVA) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.situacionIVA & " = " & vSituacionIva
        blnChange = True
    End If
    If vNomenclaturaCatastral <> clon.nomenclaturaCatastral Or IsNull(vNomenclaturaCatastral) <> IsNull(clon.nomenclaturaCatastral) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.nomenclaturaCatastral & " = '" & replaceSpecialSymbols(vNomenclaturaCatastral) & "'"
        blnChange = True
    End If
    If vCategoria <> clon.categoria Or IsNull(vCategoria) <> IsNull(clon.categoria) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.categoria & " = " & vCategoria
        blnChange = True
    End If
    If vServicio <> clon.servicio Or IsNull(vServicio) <> IsNull(clon.servicio) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.servicio & " = " & vServicio
        blnChange = True
    End If
    If vCobro <> clon.cobro Or IsNull(vCobro) <> IsNull(clon.cobro) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.cobro & " = " & vCobro
        blnChange = True
    End If
    If vZona <> clon.zona Or IsNull(vZona) <> IsNull(clon.zona) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.zona & " = " & vZona
        blnChange = True
    End If
    If vRuta <> clon.ruta Or IsNull(vRuta) <> IsNull(clon.ruta) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.ruta & " = " & vRuta
        blnChange = True
    End If
    If vOrden <> clon.orden Or IsNull(vOrden) <> IsNull(clon.orden) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.orden & " = " & vOrden
        blnChange = True
    End If
    If vEstadoID <> clon.estadoID Or IsNull(vEstadoID) <> IsNull(clon.estadoID) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.estadoID & " = " & vEstadoID
        blnChange = True
    End If
    If vFechaNacimiento <> clon.fechaNacimiento Or IsNull(vFechaNacimiento) <> IsNull(clon.fechaNacimiento) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fechaNacimiento & " = " & fechaDB(vFechaNacimiento)
        blnChange = True
    End If
    If vCategoriaSocioID <> clon.categoriasocioID Or IsNull(vCategoriaSocioID) <> IsNull(clon.categoriasocioID) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.categoriasocioID & " = " & vCategoriaSocioID
        blnChange = True
    End If
    If vDestinoID <> clon.destinoID Or IsNull(vDestinoID) <> IsNull(clon.destinoID) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.destinoID & " = " & vDestinoID
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
Dim sql As String

Dim rstQuery As ADODB.Recordset

    vFechaAlta = 1

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteID & " = " & vClienteID
    sql = sql & " ORDER BY " & ddl.fechaAlta & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set rstQuery = db.query(sql)

    If Not rstQuery.EOF Then vFechaAlta = 1 + rstQuery.Fields(ddl.fechaAlta)

    rstQuery.Close

    If pFind Then findByPrimaryKey db

End Sub

Public Function collectionAll(db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & ";"

    Set collectionAll = collectionSQL(sql, db)

End Function

Public Function collectionAny(db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set collectionAny = collectionSQL(sql, db)

End Function

Public Function collectionRango(desde As Long, hasta As Long, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.fechaBaja & " IS NULL"
    sql = sql & " AND " & ddl.clienteID & " BETWEEN " & desde & " And " & hasta
    sql = sql & " ORDER BY " & ddl.clienteID
    sql = sql & ";"

    Set collectionRango = collectionSQL(sql, db)

End Function

Public Function collectionByClienteID(clienteID As Integer, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteID & " = " & clienteID
    sql = sql & " ORDER BY " & ddl.fechaAlta
    sql = sql & ";"

    Set collectionByClienteID = collectionSQL(sql, db)

End Function

Public Function collectionActivos(db As clsDB, Optional orderByName As Boolean = False) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.fechaBaja & " IS NULL"
    If orderByName Then
        sql = sql & " ORDER BY " & ddl.apellido & ", " & ddl.nombre & ", " & ddl.clienteID
    Else
        sql = sql & " ORDER BY " & ddl.clienteID
    End If
    sql = sql & ";"

    Set collectionActivos = collectionSQL(sql, db)

End Function

Public Function collectionActivosConCuotaFija(db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.fechaBaja & " IS NULL"
    sql = sql & " AND " & ddl.cobro & " = 3"
    sql = sql & " ORDER BY " & ddl.apellido & ", " & ddl.nombre & ", " & ddl.clienteID
    sql = sql & ";"

    Set collectionActivosConCuotaFija = collectionSQL(sql, db)

End Function

Public Function collectionActivosConMedidor(db As clsDB) As Collection
Dim sql As String

Dim m As New clsMyDMedidor

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " JOIN " & m.table
    sql = sql & " ON " & ddl.clienteID_t & " = " & m.clienteID_t
    sql = sql & " WHERE " & ddl.fechaBaja_t & " IS NULL"
    sql = sql & " AND " & m.fechaRetiro_t & " IS NULL"
    sql = sql & " ORDER BY " & ddl.apellido_t & ", " & ddl.nombre_t & ", " & ddl.clienteID_t
    sql = sql & ";"

    Set collectionActivosConMedidor = collectionSQL(sql, db)

End Function

Public Function collectionActivosByZona(db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.fechaBaja & " IS NULL"
    sql = sql & " ORDER BY " & ddl.zona
    sql = sql & ";"

    Set collectionActivosByZona = collectionSQL(sql, db)

End Function

Public Function collectionActivosByRuta(pZona As Integer, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.fechaBaja & " IS NULL"
    sql = sql & " AND " & ddl.zona & " = " & pZona
    sql = sql & " ORDER BY " & ddl.ruta
    sql = sql & ";"

    Set collectionActivosByRuta = collectionSQL(sql, db)

End Function

Public Function collectionActivosByZonaRuta(zona As Integer, ruta As Integer, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.fechaBaja & " IS NULL"
    sql = sql & " AND " & ddl.zona & " = " & zona
    sql = sql & " AND " & ddl.ruta & " = " & ruta
    sql = sql & " ORDER BY " & ddl.orden
    sql = sql & ";"

    Set collectionActivosByZonaRuta = collectionSQL(sql, db)

End Function

Public Function collectionActivosByZonaRutaOtros(zona As Integer, ruta As Integer, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.fechaBaja & " IS NULL"
    sql = sql & " AND NOT (" & ddl.zona & " = " & zona
    sql = sql & " AND " & ddl.ruta & " = " & ruta & ")"
    sql = sql & " ORDER BY " & ddl.zona & ", " & ddl.ruta & ", " & ddl.orden
    sql = sql & ";"

    Set collectionActivosByZonaRutaOtros = collectionSQL(sql, db)

End Function

Public Function collectionActivosMedibles(db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.fechaBaja & " IS NULL"
    sql = sql & " AND " & ddl.cobro & " < 3"
    sql = sql & " ORDER BY " & ddl.clienteID
    sql = sql & ";"

    Set collectionActivosMedibles = collectionSQL(sql, db)

End Function

Public Function collectionSociosActivos(db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.fechaBaja & " IS NULL"
    sql = sql & " AND CAST(" & ddl.numeroSocio & " AS UNSIGNED) > 0"
    sql = sql & " ORDER BY " & ddl.apellido & ", " & ddl.nombre & ", " & ddl.clienteID
    sql = sql & ";"

    Set collectionSociosActivos = collectionSQL(sql, db)

End Function

Public Function collectionSociosActivosConMedidor(db As clsDB) As Collection
Dim sql As String

Dim m As New clsMyDMedidor

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " JOIN " & m.table
    sql = sql & " ON " & ddl.clienteID_t & " = " & m.clienteID_t
    sql = sql & " WHERE " & ddl.fechaBaja_t & " IS NULL"
    sql = sql & " AND " & m.fechaRetiro_t & " IS NULL"
    sql = sql & " AND CAST(" & ddl.numeroSocio_t & " AS UNSIGNED) > 0"
    sql = sql & " ORDER BY " & ddl.apellido_t & ", " & ddl.nombre_t & ", " & ddl.clienteID_t
    sql = sql & ";"

    Set collectionSociosActivosConMedidor = collectionSQL(sql, db)

End Function

Public Function collectionSociosActivosConCuotaFija(db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.fechaBaja & " IS NULL"
    sql = sql & " AND " & ddl.cobro & " = 3"
    sql = sql & " AND CAST(" & ddl.numeroSocio_t & " AS UNSIGNED) > 0"
    sql = sql & " ORDER BY " & ddl.apellido & ", " & ddl.nombre & ", " & ddl.clienteID
    sql = sql & ";"

    Set collectionSociosActivosConCuotaFija = collectionSQL(sql, db)

End Function

Public Function collectionSearch(pBuscar As String, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.apellido & " LIKE '%" & pBuscar & "%'"
    sql = sql & " ORDER BY " & ddl.apellido
    If Len(pBuscar) < 4 Then sql = sql & " LIMIT 50"
    sql = sql & ";"
    
    Set collectionSearch = collectionSQL(sql, db)
    
End Function

Public Function collectionActivos2Lectura(pZonaID As Integer, pRutaID As Integer, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.fechaBaja & " IS NULL"
    sql = sql & " AND " & ddl.zona & " = " & pZonaID
    sql = sql & " AND " & ddl.ruta & " = " & pRutaID
    sql = sql & " ORDER BY " & ddl.orden
    sql = sql & ";"

    Set collectionActivos2Lectura = collectionSQL(sql, db)

End Function

Public Function collectionZona(zona As Integer, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.fechaBaja & " IS NULL"
    sql = sql & " AND " & ddl.zona & " = " & zona
    sql = sql & " ORDER BY " & ddl.zona & ", " & ddl.ruta & ", " & ddl.orden & ", " & ddl.clienteID
    sql = sql & ";"

    Set collectionZona = collectionSQL(sql, db)

End Function

Public Function collectionRuta(zona As Integer, ruta As Integer, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.fechaBaja & " IS NULL"
    sql = sql & " AND " & ddl.zona & " = " & zona
    sql = sql & " AND " & ddl.ruta & " = " & ruta
    sql = sql & " ORDER BY " & ddl.zona & ", " & ddl.ruta & ", " & ddl.orden & ", " & ddl.clienteID
    sql = sql & ";"

    Set collectionRuta = collectionSQL(sql, db)

End Function

Public Function collectionDeudoresPlanCorte(db As clsDB, Optional registrosold As Collection = Nothing) As Collection
Dim sql As String

Dim c As New clsMyDCuota

    sql = "SELECT DISTINCTROW " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " JOIN " & c.table
    sql = sql & " ON " & ddl.clienteID_t & " = " & c.clienteID_t
    sql = sql & " WHERE " & ddl.fechaBaja_t & " IS NULL"
    sql = sql & " AND " & c.fechaPago_t & " IS NULL"
    sql = sql & " AND " & c.fechaVencimiento_t & " < CURDATE()"
    sql = sql & ";"

    Set collectionDeudoresPlanCorte = collectionSQLNoRep(sql, db, registrosold)
    
End Function

Public Function collectionDeudoresFactura60dias(db As clsDB, Optional registrosold As Collection = Nothing) As Collection
Dim sql As String

Dim f As New clsMyDFactura
Dim p As New clsMyDPeriodo
Dim d As New clsMyDDeuda

    sql = "SELECT DISTINCTROW " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " JOIN " & f.table
    sql = sql & " ON " & ddl.clienteID_t & " = " & f.clienteID_t
    sql = sql & " JOIN " & p.table
    sql = sql & " ON " & p.periodoID_t & " = " & f.periodoID_t
    sql = sql & " WHERE " & ddl.fechaBaja_t & " IS NULL"
    sql = sql & " AND " & f.fechaPago_t & " IS NULL"
    sql = sql & " AND " & f.anulada_t & " = 0"
    sql = sql & " AND " & f.cancelada_t & " = 0"
    sql = sql & " AND DATEDIFF(CURDATE(), " & p.fechaSegundo_t & ") > 60"
    sql = sql & " AND NOT " & ddl.clienteID_t & " IN (SELECT " & d.clienteID_t & " FROM " & d.table & " WHERE " & d.pagado_t & " = 0)"
    sql = sql & " GROUP BY " & ddl.clienteID_t
    sql = sql & " HAVING COUNT(*) = 1"
    sql = sql & " ORDER BY " & ddl.apellido_t & ", " & ddl.nombre_t
    sql = sql & ";"

    Set collectionDeudoresFactura60dias = collectionSQLNoRep(sql, db, registrosold)
    
End Function

Public Function collectionDeudoresFacturas2(db As clsDB, Optional registrosold As Collection = Nothing) As Collection
Dim sql As String

Dim f As New clsMyDFactura
Dim p As New clsMyDPeriodo

    sql = "SELECT DISTINCTROW " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " JOIN " & f.table
    sql = sql & " ON " & ddl.clienteID_t & " = " & f.clienteID_t
    sql = sql & " JOIN " & p.table
    sql = sql & " ON " & p.periodoID_t & " = " & f.periodoID_t
    sql = sql & " WHERE " & ddl.fechaBaja_t & " IS NULL"
    sql = sql & " AND " & f.fechaPago_t & " IS NULL"
    sql = sql & " AND " & f.anulada_t & " = 0"
    sql = sql & " AND " & f.cancelada_t & " = 0"
    sql = sql & " AND " & p.fechaSegundo_t & " < CURDATE()"
    sql = sql & " GROUP BY " & ddl.clienteID_t
    sql = sql & " HAVING COUNT(*) = 2"
    sql = sql & " ORDER BY " & ddl.apellido_t & ", " & ddl.nombre_t
    sql = sql & ";"

    Set collectionDeudoresFacturas2 = collectionSQLNoRep(sql, db, registrosold)
    
End Function

Public Function collectionSQL(pSQL As String, db As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim elemento As clsMyACliente

Dim registros As New Collection

    Set rstQuery = db.query(pSQL)

    Do While Not rstQuery.EOF
        Set elemento = New clsMyACliente
        fillObject elemento, rstQuery

        If modCollection.collectionExistElement(registros, elemento.key) Then
            registros.add elemento
        Else
            registros.add elemento, elemento.key
        End If

        rstQuery.MoveNext
    Loop

    rstQuery.Close

    Set collectionSQL = registros

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 29/06/2018
' Description: Hace la consulta SQL sobre la dbapp y arma la colección
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function collectionSQLNoRep(sql As String, db As clsDB, Optional registrosold As Collection = Nothing) As Collection
Dim query As ADODB.Recordset

Dim elemento As clsMyACliente

Dim registros As Collection

    If registrosold Is Nothing Then
        Set registros = New Collection
    Else
        Set registros = registrosold
    End If

    Set query = db.query(sql)

    Do While Not query.EOF
        Set elemento = New clsMyACliente
        fillObject elemento, query

        If Not modCollection.collectionExistElement(registros, elemento.key) Then registros.add elemento, elemento.key

        query.MoveNext
    Loop

    query.Close

    Set collectionSQLNoRep = registros

End Function

Public Sub fillCombo(pCombo As ComboBox, db As clsDB, Optional pComboDomicilio As ComboBox = Nothing)

    fillCombo_ pCombo, collectionActivos(db), pComboDomicilio

End Sub

Public Sub fillComboOtros(pCombo As ComboBox, zona As Integer, ruta As Integer, db As clsDB)

    fillCombo_ pCombo, collectionActivosByZonaRutaOtros(zona, ruta, db)

End Sub

Private Sub fillCombo_(pCombo As ComboBox, pCollection As Collection, Optional pComboDomicilio As ComboBox = Nothing)
Dim elemento As clsMyACliente

    pCombo.Clear

    For Each elemento In pCollection
        pCombo.AddItem elemento.comboText
        pCombo.ItemData(pCombo.NewIndex) = elemento.comboKey
        If Not (pComboDomicilio Is Nothing) Then
            pComboDomicilio.AddItem elemento.domicilioText
            pComboDomicilio.ItemData(pComboDomicilio.NewIndex) = elemento.comboKey
        End If
    Next

    If pCombo.ListCount > 0 Then pCombo.ListIndex = 0

End Sub

Public Sub fillList(pList As ListBox, db As clsDB)

    fillList_ pList, collectionAll(db)

End Sub

Private Sub fillList_(pList As ListBox, pCollection As Collection)
Dim elemento As clsMyACliente

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

Public Sub formSearch(pForm As Form, pObjeto As Variant, pKeyAscii As Integer, pQueBusco As String, db As clsDB)

    Load pForm
    
    With pForm
        .setObjeto pObjeto
        .setDB db
        .txtCadena.Text = Chr(pKeyAscii)
        .Caption = "Búsqueda " & pQueBusco
        .Show vbModal
    End With
    
    Unload pForm
    
End Sub


