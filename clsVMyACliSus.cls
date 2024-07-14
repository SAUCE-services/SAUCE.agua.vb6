VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0   'vbNone
  MTSTransactionMode  = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsVMyACliSus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vClienteID As Long
Private vApellido As String
Private vNombre As String
Private vInmuebleCalle As String
Private vInmueblePuerta As String
Private vInmueblePiso As String
Private vInmuebleDpto As String
Private vCobro As Integer
Private vPeriodoIDInicio As Integer
Private vPeriodoIDFin As Variant

Private clon As clsVMyACliSus

Public ddl As New clsDDLVw_clisus

Private Sub class_Initialize()

    initializeOther

End Sub

Private Sub initializeOther()

    vClienteID = 0
    vApellido = ""
    vNombre = ""
    vInmuebleCalle = ""
    vInmueblePuerta = ""
    vInmueblePiso = ""
    vInmuebleDpto = ""
    vCobro = 0
    vPeriodoIDInicio = 0
    vPeriodoIDFin = Null

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

Public Property Get cobro() As Variant

    cobro = vCobro

End Property

Public Property Let cobro(ByVal vNewValue As Variant)

On Error Resume Next

    vCobro = vNewValue

End Property

Public Property Get periodoIDInicio() As Variant

    periodoIDInicio = vPeriodoIDInicio

End Property

Public Property Let periodoIDInicio(ByVal vNewValue As Variant)

On Error Resume Next

    vPeriodoIDInicio = vNewValue

End Property

Public Property Get periodoIDFin() As Variant

    periodoIDFin = vPeriodoIDFin

End Property

Public Property Let periodoIDFin(ByVal vNewValue As Variant)

On Error Resume Next

    vPeriodoIDFin = vNewValue

End Property

Public Function clone() As clsVMyACliSus
Dim objObjeto As New clsVMyACliSus

    With objObjeto
        .clienteID = vClienteID
        .apellido = vApellido
        .nombre = vNombre
        .inmuebleCalle = vInmuebleCalle
        .inmueblePuerta = vInmueblePuerta
        .inmueblePiso = vInmueblePiso
        .inmuebleDpto = vInmuebleDpto
        .cobro = vCobro
        .periodoIDInicio = vPeriodoIDInicio
        .periodoIDFin = vPeriodoIDFin
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsVMyACliSus, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .clienteID = pQuery.fields(.ddl.clienteID)
        .apellido = pQuery.fields(.ddl.apellido)
        .nombre = pQuery.fields(.ddl.nombre)
        .inmuebleCalle = pQuery.fields(.ddl.inmuebleCalle)
        .inmueblePuerta = pQuery.fields(.ddl.inmueblePuerta)
        .inmueblePiso = pQuery.fields(.ddl.inmueblePiso)
        .inmuebleDpto = pQuery.fields(.ddl.inmuebleDpto)
        .cobro = pQuery.fields(.ddl.cobro)
        .periodoIDInicio = pQuery.fields(.ddl.periodoIDInicio)
        .periodoIDFin = pQuery.fields(.ddl.periodoIDFin)
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

Dim objElemento As clsVMyACliSus

Dim colRegistros As New Collection

    Set rstQuery = pDB.query(pSQL)

    Do While Not rstQuery.EOF
        Set objElemento = New clsVMyACliSus
        fillObject objElemento, rstQuery

        colRegistros.add objElemento

        rstQuery.MoveNext
    Loop

    rstQuery.Close

    Set collectionSQL = colRegistros

End Function

