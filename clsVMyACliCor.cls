VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVMyACliCor"
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

Private clon As clsVMyACliCor

Public ddl As New clsVMyDCliCor

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

Public Function clone() As clsVMyACliCor
Dim objObjeto As New clsVMyACliCor

    With objObjeto
        .clienteID = vClienteID
        .apellido = vApellido
        .nombre = vNombre
        .inmuebleCalle = vInmuebleCalle
        .inmueblePuerta = vInmueblePuerta
        .inmueblePiso = vInmueblePiso
        .inmuebleDpto = vInmuebleDpto
        .cobro = vCobro
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsVMyACliCor, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .clienteID = pQuery.Fields(.ddl.clienteID)
        .apellido = pQuery.Fields(.ddl.apellido)
        .nombre = pQuery.Fields(.ddl.nombre)
        .inmuebleCalle = pQuery.Fields(.ddl.inmuebleCalle)
        .inmueblePuerta = pQuery.Fields(.ddl.inmueblePuerta)
        .inmueblePiso = pQuery.Fields(.ddl.inmueblePiso)
        .inmuebleDpto = pQuery.Fields(.ddl.inmuebleDpto)
        .cobro = pQuery.Fields(.ddl.cobro)
    End With

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

Public Function collectionSQL(pSQL As String, db As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim elemento As clsVMyACliCor

Dim registros As New Collection

    Set rstQuery = db.query(pSQL)

    Do While Not rstQuery.EOF
        Set elemento = New clsVMyACliCor
        fillObject elemento, rstQuery

        registros.add elemento

        rstQuery.MoveNext
    Loop

    rstQuery.Close

    Set collectionSQL = registros

End Function

