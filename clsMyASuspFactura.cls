VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyASuspFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vClienteID As Long
Private vPeriodoIDInicio As Integer
Private vPeriodoIDFin As Variant
Private vMotivo As String
Private vAutoID As Long
Private vUid As String
Private vCreated As Date

Private clon As clsMyASuspFactura

Public ddl As New clsMyDSuspFactura

Private Sub class_Initialize()

    vClienteID = 0
    vPeriodoIDInicio = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vPeriodoIDFin = Null
    vMotivo = ""
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

Public Property Get motivo() As Variant

    motivo = vMotivo

End Property

Public Property Let motivo(ByVal vNewValue As Variant)

On Error Resume Next

    vMotivo = vNewValue

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

    comboText = vPeriodoIDFin

End Property

Public Property Get comboKey() As Variant

    comboKey = vPeriodoIDInicio

End Property

Public Property Get listText() As Variant

    listText = vPeriodoIDFin

End Property

Public Property Get listKey() As Variant

    listKey = vPeriodoIDInicio

End Property

Public Function clone() As clsMyASuspFactura
Dim objObjeto As New clsMyASuspFactura

    With objObjeto
        .clienteID = vClienteID
        .periodoIDInicio = vPeriodoIDInicio
        .periodoIDFin = vPeriodoIDFin
        .motivo = vMotivo
        .autoID = vAutoID
        .uid = vUid
        .created = vCreated
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsMyASuspFactura, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .clienteID = pQuery.Fields(.ddl.clienteID)
        .periodoIDInicio = pQuery.Fields(.ddl.periodoIDInicio)
        .periodoIDFin = pQuery.Fields(.ddl.periodoIDFin)
        .motivo = pQuery.Fields(.ddl.motivo)
        .autoID = pQuery.Fields(.ddl.autoID)
        .uid = pQuery.Fields(.ddl.uid)
        .created = pQuery.Fields(.ddl.created)
    End With

End Sub

Private Function makeWherePrimaryKey() As String
Dim sql As String

    sql = ""
    sql = sql & " WHERE " & ddl.clienteID & " = " & vClienteID
    sql = sql & " AND " & ddl.periodoIDInicio & " = " & vPeriodoIDInicio

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
    sql = sql & " WHERE " & ddl.clienteID & " = " & vClienteID
    sql = sql & " ORDER BY " & ddl.periodoIDInicio & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, db

End Sub

Public Sub findSuspendidaByClienteID(db As clsDB)
Dim sql As String

    initializeOther

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteID & " = " & vClienteID
    sql = sql & " AND " & ddl.periodoIDInicio & " <= " & vPeriodoIDInicio
    sql = sql & " ORDER BY " & ddl.periodoIDInicio & " DESC"
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
    sql = sql & " SET " & ddl.periodoIDFin & " = " & validateNullField(vPeriodoIDFin)
    sql = sql & ", " & ddl.motivo & " = '" & replaceSpecialSymbols(vMotivo) & "'"
    sql = sql & ", " & ddl.uid & " = '" & replaceSpecialSymbols(vUid) & "'"
    sql = sql & ", " & ddl.clienteID & " = " & vClienteID
    sql = sql & ", " & ddl.periodoIDInicio & " = " & vPeriodoIDInicio
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

    If vPeriodoIDFin <> clon.periodoIDFin Or IsNull(vPeriodoIDFin) <> IsNull(clon.periodoIDFin) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.periodoIDFin & " = " & validateNullField(vPeriodoIDFin)
        blnChange = True
    End If
    If vMotivo <> clon.motivo Or IsNull(vMotivo) <> IsNull(clon.motivo) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.motivo & " = '" & replaceSpecialSymbols(vMotivo) & "'"
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

    vPeriodoIDInicio = 1

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteID & " = " & vClienteID
    sql = sql & " ORDER BY " & ddl.periodoIDInicio & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set rstQuery = db.query(sql)

    If Not rstQuery.EOF Then vPeriodoIDInicio = 1 + rstQuery.Fields(ddl.periodoIDInicio)

    rstQuery.Close

    If pFind Then findByPrimaryKey db

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
    sql = sql & " WHERE " & ddl.periodoIDInicio & " <= " & periodoID
    sql = sql & " AND (" & ddl.periodoIDFin & " IS NULL OR " & ddl.periodoIDFin & " >= " & periodoID & ")"
    sql = sql & ";"

    Set collectionByPeriodoID = collectionSQL(sql, db)

End Function

Public Function collectionSQL(pSQL As String, db As clsDB) As Collection

    Set collectionSQL = collection_(pSQL, IIf(IsMissing(db), dbapp, db))

End Function

Private Function collection_(pSQL As String, db As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim elemento As clsMyASuspFactura

Dim registros As New Collection

    Set rstQuery = db.query(pSQL)

    Do While Not rstQuery.EOF
        Set elemento = New clsMyASuspFactura
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
Dim elemento As clsMyASuspFactura

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
Dim elemento As clsMyASuspFactura

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

