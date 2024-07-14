VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyASuspension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vTipo As String
Private vNumero As Long
Private vFecha As Date
Private vClienteID As Long
Private vPeriodoID As Integer
Private vAutoID As Long
Private vUid As String
Private vCreated As Date

Private clon As clsMyASuspension

Public ddl As New clsMyDSuspension

Private Sub class_Initialize()

    vTipo = ""
    vNumero = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vFecha = Date
    vClienteID = 0
    vPeriodoID = 0
    vAutoID = 0
    vUid = ""
    vCreated = Date

End Sub

Public Sub clean()

    Call class_Initialize

End Sub

Public Property Get tipo() As Variant

    tipo = vTipo

End Property

Public Property Let tipo(ByVal vNewValue As Variant)

On Error Resume Next

    vTipo = vNewValue

End Property

Public Property Get numero() As Variant

    numero = vNumero

End Property

Public Property Let numero(ByVal vNewValue As Variant)

On Error Resume Next

    vNumero = vNewValue

End Property

Public Property Get fecha() As Variant

    fecha = vFecha

End Property

Public Property Let fecha(ByVal vNewValue As Variant)

On Error Resume Next

    vFecha = vNewValue

End Property

Public Property Get clienteID() As Variant

    clienteID = vClienteID

End Property

Public Property Let clienteID(ByVal vNewValue As Variant)

On Error Resume Next

    vClienteID = vNewValue

End Property

Public Property Get periodoID() As Variant

    periodoID = vPeriodoID

End Property

Public Property Let periodoID(ByVal vNewValue As Variant)

On Error Resume Next

    vPeriodoID = vNewValue

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

    comboText = vFecha

End Property

Public Property Get comboKey() As Variant

    comboKey = vNumero

End Property

Public Property Get listText() As Variant

    listText = vFecha

End Property

Public Property Get listKey() As Variant

    listKey = vNumero

End Property

Public Function clone() As clsMyASuspension
Dim objObjeto As New clsMyASuspension

    With objObjeto
        .tipo = vTipo
        .numero = vNumero
        .fecha = vFecha
        .clienteID = vClienteID
        .periodoID = vPeriodoID
        .autoID = vAutoID
        .uid = vUid
        .created = vCreated
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsMyASuspension, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .tipo = pQuery.Fields(.ddl.tipo)
        .numero = pQuery.Fields(.ddl.numero)
        .fecha = pQuery.Fields(.ddl.fecha)
        .clienteID = pQuery.Fields(.ddl.clienteID)
        .periodoID = pQuery.Fields(.ddl.periodoID)
        .autoID = pQuery.Fields(.ddl.autoID)
        .uid = pQuery.Fields(.ddl.uid)
        .created = pQuery.Fields(.ddl.created)
    End With

End Sub

Private Function makeWherePrimaryKey() As String
Dim sql As String

    sql = ""
    sql = sql & " WHERE " & ddl.tipo & " = '" & replaceSpecialSymbols(vTipo) & "'"
    sql = sql & " AND " & ddl.numero & " = " & vNumero

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

Public Sub findByClienteID(db As clsDB)
Dim sql As String
Dim tipo As String

Dim clienteID As Long

Dim fecha As Date

    tipo = vTipo
    clienteID = vClienteID
    fecha = vFecha
    
    clean
    
    vTipo = tipo
    vClienteID = clienteID
    vFecha = fecha

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.tipo & " = '" & replaceSpecialSymbols(vTipo) & "'"
    sql = sql & " AND " & ddl.clienteID & " = " & vClienteID
    sql = sql & " AND " & ddl.fecha & " = " & fechaDB(vFecha)
    sql = sql & " ORDER BY " & ddl.numero & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, db

End Sub

Public Sub findLast(db As clsDB)
Dim sql As String

    initializeOther

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.tipo & " = '" & replaceSpecialSymbols(vTipo) & "'"
    sql = sql & " ORDER BY " & ddl.numero & " DESC"
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
    sql = sql & " SET " & ddl.fecha & " = " & fechaDB(vFecha)
    sql = sql & ", " & ddl.clienteID & " = " & vClienteID
    sql = sql & ", " & ddl.periodoID & " = " & vPeriodoID
    sql = sql & ", " & ddl.uid & " = '" & replaceSpecialSymbols(vUid) & "'"
    sql = sql & ", " & ddl.tipo & " = '" & replaceSpecialSymbols(vTipo) & "'"
    sql = sql & ", " & ddl.numero & " = " & vNumero
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

    If vFecha <> clon.fecha Or IsNull(vFecha) <> IsNull(clon.fecha) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fecha & " = " & fechaDB(vFecha)
        blnChange = True
    End If
    If vClienteID <> clon.clienteID Or IsNull(vClienteID) <> IsNull(clon.clienteID) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.clienteID & " = " & vClienteID
        blnChange = True
    End If
    If vPeriodoID <> clon.periodoID Or IsNull(vPeriodoID) <> IsNull(clon.periodoID) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.periodoID & " = " & vPeriodoID
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

    vNumero = 1

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.tipo & " = '" & replaceSpecialSymbols(vTipo) & "'"
    sql = sql & " ORDER BY " & ddl.numero & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set rstQuery = db.query(sql)

    If Not rstQuery.EOF Then vNumero = 1 + rstQuery.Fields(ddl.numero)

    rstQuery.Close

    If pFind Then findByPrimaryKey db

End Sub

Public Function collectionAll(db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & ";"

    Set collectionAll = collectionSQL(sql, db)

End Function

Public Function collectionByPeriodoID(tipo As String, periodoID As Integer, db As clsDB) As Collection
Dim sql As String

Dim c As New clsMyDCliente

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " JOIN " & c.table
    sql = sql & " ON " & c.clienteID_t & " = " & ddl.clienteID_t
    sql = sql & " WHERE " & ddl.tipo_t & " = ' " & tipo & "'"
    sql = sql & " AND " & ddl.periodoID_t & " = " & periodoID
    sql = sql & " AND " & c.fechaBaja_t & " IS NULL"
    sql = sql & " ORDER BY " & ddl.numero_t
    sql = sql & ";"

    Set collectionByPeriodoID = collectionSQL(sql, db)

End Function

Public Function collectionSQL(pSQL As String, db As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim elemento As clsMyASuspension

Dim registros As New Collection

    Set rstQuery = db.query(pSQL)

    Do While Not rstQuery.EOF
        Set elemento = New clsMyASuspension
        fillObject elemento, rstQuery

        registros.add elemento

        rstQuery.MoveNext
    Loop

    rstQuery.Close

    Set collectionSQL = registros

End Function

Public Sub fillCombo(pCombo As ComboBox, db As clsDB)

    fillCombo_ pCombo, collectionAll(db)

End Sub

Private Sub fillCombo_(pCombo As ComboBox, pCollection As Collection)
Dim elemento As clsMyASuspension

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
Dim elemento As clsMyASuspension

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

