VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0   'vbNone
  MTSTransactionMode  = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsVMyAListGral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vMedidorID As String
Private vClienteID As Long
Private vApellido As String
Private vNombre As String
Private vInmuebleCalle As String
Private vInmueblePuerta As String
Private vInmueblePiso As String
Private vInmuebleDpto As String
Private vInmuebleLocalidad As String
Private vInmuebleProvincia As String
Private vFiscalCalle As String
Private vFiscalPuerta As String
Private vFiscalPiso As String
Private vFiscalDpto As String
Private vFiscalLocalidad As String
Private vFiscalProvincia As String
Private vZona As Integer
Private vRuta As Integer
Private vOrden As Integer
Private vUltima As Variant
Private vEstadoInicio As Long
Private vEstado As Variant

Private clon As clsVMyAListGral

Public ddl As New clsDDLVw_listgral

Private Sub class_Initialize()

    initializeOther

End Sub

Private Sub initializeOther()

    vMedidorID = ""
    vClienteID = 0
    vApellido = ""
    vNombre = ""
    vInmuebleCalle = ""
    vInmueblePuerta = ""
    vInmueblePiso = ""
    vInmuebleDpto = ""
    vInmuebleLocalidad = ""
    vInmuebleProvincia = ""
    vFiscalCalle = ""
    vFiscalPuerta = ""
    vFiscalPiso = ""
    vFiscalDpto = ""
    vFiscalLocalidad = ""
    vFiscalProvincia = ""
    vZona = 0
    vRuta = 0
    vOrden = 0
    vUltima = Null
    vEstadoInicio = 0
    vEstado = Null

End Sub

Public Sub clean()

    Call class_Initialize

End Sub

Public Property Get medidorID() As Variant

    medidorID = vMedidorID

End Property

Public Property Let medidorID(ByVal vNewValue As Variant)

On Error Resume Next

    vMedidorID = vNewValue

End Property

Public Property Get clienteID() As Variant

    clienteID = vClienteID

End Property

Public Property Let clienteID(ByVal vNewValue As Variant)

On Error Resume Next

    vClienteID = vNewValue

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

Public Property Let nombre(ByVal vNewValue As Variant)

On Error Resume Next

    vNombre = vNewValue

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

Public Property Get ultima() As Variant

    ultima = vUltima

End Property

Public Property Let ultima(ByVal vNewValue As Variant)

On Error Resume Next

    vUltima = vNewValue

End Property

Public Property Get estadoInicio() As Variant

    estadoInicio = vEstadoInicio

End Property

Public Property Let estadoInicio(ByVal vNewValue As Variant)

On Error Resume Next

    vEstadoInicio = vNewValue

End Property

Public Property Get estado() As Variant

    estado = vEstado

End Property

Public Property Let estado(ByVal vNewValue As Variant)

On Error Resume Next

    vEstado = vNewValue

End Property

Public Function clone() As clsVMyAListGral
Dim objObjeto As New clsVMyAListGral

    With objObjeto
        .medidorID = vMedidorID
        .clienteID = vClienteID
        .apellido = vApellido
        .nombre = vNombre
        .inmuebleCalle = vInmuebleCalle
        .inmueblePuerta = vInmueblePuerta
        .inmueblePiso = vInmueblePiso
        .inmuebleDpto = vInmuebleDpto
        .inmuebleLocalidad = vInmuebleLocalidad
        .inmuebleProvincia = vInmuebleProvincia
        .fiscalCalle = vFiscalCalle
        .fiscalPuerta = vFiscalPuerta
        .fiscalPiso = vFiscalPiso
        .fiscalDpto = vFiscalDpto
        .fiscalLocalidad = vFiscalLocalidad
        .fiscalProvincia = vFiscalProvincia
        .zona = vZona
        .ruta = vRuta
        .orden = vOrden
        .ultima = vUltima
        .estadoInicio = vEstadoInicio
        .estado = vEstado
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsVMyAListGral, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .medidorID = pQuery.fields(.ddl.medidorID)
        .clienteID = pQuery.fields(.ddl.clienteID)
        .apellido = pQuery.fields(.ddl.apellido)
        .nombre = pQuery.fields(.ddl.nombre)
        .inmuebleCalle = pQuery.fields(.ddl.inmuebleCalle)
        .inmueblePuerta = pQuery.fields(.ddl.inmueblePuerta)
        .inmueblePiso = pQuery.fields(.ddl.inmueblePiso)
        .inmuebleDpto = pQuery.fields(.ddl.inmuebleDpto)
        .inmuebleLocalidad = pQuery.fields(.ddl.inmuebleLocalidad)
        .inmuebleProvincia = pQuery.fields(.ddl.inmuebleProvincia)
        .fiscalCalle = pQuery.fields(.ddl.fiscalCalle)
        .fiscalPuerta = pQuery.fields(.ddl.fiscalPuerta)
        .fiscalPiso = pQuery.fields(.ddl.fiscalPiso)
        .fiscalDpto = pQuery.fields(.ddl.fiscalDpto)
        .fiscalLocalidad = pQuery.fields(.ddl.fiscalLocalidad)
        .fiscalProvincia = pQuery.fields(.ddl.fiscalProvincia)
        .zona = pQuery.fields(.ddl.zona)
        .ruta = pQuery.fields(.ddl.ruta)
        .orden = pQuery.fields(.ddl.orden)
        .ultima = pQuery.fields(.ddl.ultima)
        .estadoInicio = pQuery.fields(.ddl.estadoInicio)
        .estado = pQuery.fields(.ddl.estado)
    End With

End Sub

Public Function collectionAll(pDB As clsDB) As Collection
Dim strSQL As String

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & ";"

    Set collectionAll = collectionSQL(strSQL, pDB)

End Function

Public Function collectionAny(pDB As clsDB) As Collection
Dim strSQL As String

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " LIMIT 1"
    strSQL = strSQL & ";"

    Set collectionAny = collectionSQL(strSQL, pDB)

End Function

Public Function collectionSQL(pSQL As String, pDB As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim objElemento As clsVMyAListGral

Dim colRegistros As New Collection

    Set rstQuery = pDB.query(pSQL)

    Do While Not rstQuery.EOF
        Set objElemento = New clsVMyAListGral
        fillObject objElemento, rstQuery

        colRegistros.add objElemento

        rstQuery.MoveNext
    Loop

    rstQuery.Close

    Set collectionSQL = colRegistros

End Function

