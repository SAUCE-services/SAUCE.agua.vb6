VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVMyACliMed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vClienteID As Long
Private vFecha As Variant
Private vMedidorID As String

Private clon As clsVMyACliMed

Public ddl As New clsVMyDCliMed

Private Sub class_Initialize()

    initializeOther

End Sub

Private Sub initializeOther()

    vClienteID = 0
    vFecha = Null
    vMedidorID = ""

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

Public Property Get fecha() As Variant

    fecha = vFecha

End Property

Public Property Let fecha(ByVal vNewValue As Variant)

On Error Resume Next

    vFecha = vNewValue

End Property

Public Property Get medidorID() As Variant

    medidorID = vMedidorID

End Property

Public Property Let medidorID(ByVal vNewValue As Variant)

On Error Resume Next

    vMedidorID = vNewValue

End Property

Public Function clone() As clsVMyACliMed
Dim objObjeto As New clsVMyACliMed

    With objObjeto
        .clienteID = vClienteID
        .fecha = vFecha
        .medidorID = vMedidorID
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsVMyACliMed, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .clienteID = pQuery.Fields(.ddl.clienteID)
        .fecha = pQuery.Fields(.ddl.fecha)
        .medidorID = pQuery.Fields(.ddl.medidorID)
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

Dim elemento As clsVMyACliMed

Dim registros As New Collection

    Set rstQuery = db.query(pSQL)

    Do While Not rstQuery.EOF
        Set elemento = New clsVMyACliMed
        fillObject elemento, rstQuery

        registros.add elemento

        rstQuery.MoveNext
    Loop

    rstQuery.Close

    Set collectionSQL = registros

End Function

