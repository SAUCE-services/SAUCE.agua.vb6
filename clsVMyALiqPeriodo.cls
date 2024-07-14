VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVMyALiqPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vPeriodoID As Integer
Private vLiquidado As Currency

Private clon As clsVMyALiqPeriodo

Public ddl As New clsVMyDLiqPeriodo

Private Sub class_Initialize()

    initializeOther

End Sub

Private Sub initializeOther()

    vPeriodoID = 0
    vLiquidado = 0

End Sub

Public Sub clean()

    Call class_Initialize

End Sub

Public Property Get periodoID() As Variant

    periodoID = vPeriodoID

End Property

Public Property Let periodoID(ByVal vNewValue As Variant)

On Error Resume Next

    vPeriodoID = vNewValue

End Property

Public Property Get liquidado() As Variant

    liquidado = vLiquidado

End Property

Public Property Let liquidado(ByVal vNewValue As Variant)

On Error Resume Next

    vLiquidado = vNewValue

End Property

Public Function clone() As clsVMyALiqPeriodo
Dim objObjeto As New clsVMyALiqPeriodo

    With objObjeto
        .periodoID = vPeriodoID
        .liquidado = vLiquidado
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsVMyALiqPeriodo, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .periodoID = pQuery.Fields(.ddl.periodoID)
        .liquidado = pQuery.Fields(.ddl.liquidado)
    End With

End Sub

Public Function collectionAll(db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & ";"

    Set collectionAll = collectionSQL(sql, db)

End Function

Public Function collectionByPeriodoID(periodoID As Integer, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.periodoID & " = " & periodoID
    sql = sql & ";"

    Set collectionByPeriodoID = collectionSQL(sql, db)

End Function

Public Function collectionSQL(pSQL As String, db As clsDB) As Collection

    Set collectionSQL = collection_(pSQL, IIf(IsMissing(db), dbapp, db))

End Function

Private Function collection_(pSQL As String, db As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim elemento As clsVMyALiqPeriodo

Dim registros As New Collection

    Set rstQuery = db.query(pSQL)

    Do While Not rstQuery.EOF
        Set elemento = New clsVMyALiqPeriodo
        fillObject elemento, rstQuery

        registros.add elemento, "k." & elemento.periodoID

        rstQuery.MoveNext
    Loop

    rstQuery.Close

    Set collection_ = registros

End Function

