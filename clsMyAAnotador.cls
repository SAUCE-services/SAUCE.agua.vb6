VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyAAnotador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vAnotadorID As Long
Private vClienteID As Long
Private vAnotacion As String
Private vCreated As Date

Private clon As clsMyAAnotador

Public ddl As New clsMyDAnotador

Private Sub class_Initialize()

    vAnotadorID = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vClienteID = 0
    vAnotacion = ""
    vCreated = Date

End Sub

Public Sub clean()

    Call class_Initialize

End Sub

Public Property Get anotadorID() As Variant

    anotadorID = vAnotadorID

End Property

Public Property Let anotadorID(ByVal vNewValue As Variant)

On Error Resume Next

    vAnotadorID = vNewValue

End Property

Public Property Get clienteID() As Variant

    clienteID = vClienteID

End Property

Public Property Let clienteID(ByVal vNewValue As Variant)

On Error Resume Next

    vClienteID = vNewValue

End Property

Public Property Get anotacion() As Variant

    anotacion = vAnotacion

End Property

Public Property Let anotacion(ByVal vNewValue As Variant)

On Error Resume Next

    vAnotacion = vNewValue

End Property

Public Property Get created() As Variant

    created = vCreated

End Property

Public Property Let created(ByVal vNewValue As Variant)

On Error Resume Next

    vCreated = vNewValue

End Property

Public Property Get comboText() As Variant

    comboText = vClienteID

End Property

Public Property Get comboKey() As Variant

    comboKey = vAnotadorID

End Property

Public Property Get listText() As Variant

    listText = vClienteID

End Property

Public Property Get listKey() As Variant

    listKey = vAnotadorID

End Property

Public Function clone() As clsMyAAnotador
Dim objObjeto As New clsMyAAnotador

    With objObjeto
        .anotadorID = vAnotadorID
        .clienteID = vClienteID
        .anotacion = vAnotacion
        .created = vCreated
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsMyAAnotador, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .anotadorID = pQuery.Fields(.ddl.anotadorID)
        .clienteID = pQuery.Fields(.ddl.clienteID)
        .anotacion = pQuery.Fields(.ddl.anotacion)
        .created = pQuery.Fields(.ddl.created)
    End With

End Sub

Private Function makeWherePrimaryKey() As String
Dim sql As String

    sql = ""
    sql = sql & " WHERE " & ddl.anotadorID & " = " & vAnotadorID

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

Public Sub findLastID(db As clsDB)
Dim sql As String

Dim rstQuery As ADODB.Recordset

    clean

    sql = "SELECT LAST_INSERT_ID() AS " & ddl.anotadorID
    sql = sql & ";"

    Set rstQuery = db.query(sql)

    If Not rstQuery.EOF Then vAnotadorID = rstQuery.Fields(ddl.anotadorID)

    rstQuery.Close

    findByPrimaryKey db

End Sub

Public Sub findLast(db As clsDB)
Dim sql As String

    initializeOther

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " ORDER BY " & ddl.anotadorID & " DESC"
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
    sql = sql & " SET " & ddl.clienteID & " = " & vClienteID
    sql = sql & ", " & ddl.anotacion & " = '" & replaceSpecialSymbols(vAnotacion) & "'"
    sql = sql & ";"

    add = db.execute(sql)

    findLastID db

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

    If vClienteID <> clon.clienteID Or IsNull(vClienteID) <> IsNull(clon.clienteID) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.clienteID & " = " & vClienteID
        blnChange = True
    End If
    If vAnotacion <> clon.anotacion Or IsNull(vAnotacion) <> IsNull(clon.anotacion) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.anotacion & " = '" & replaceSpecialSymbols(vAnotacion) & "'"
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

    vAnotadorID = 1

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " ORDER BY " & ddl.anotadorID & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set rstQuery = db.query(sql)

    If Not rstQuery.EOF Then vAnotadorID = 1 + rstQuery.Fields(ddl.anotadorID)

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

Public Function collectionByClienteID(clienteID As Long, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteID & " = " & clienteID
    sql = sql & " ORDER BY " & ddl.anotadorID & " DESC"
    sql = sql & ";"

    Set collectionByClienteID = collectionSQL(sql, db)

End Function

Public Function collectionSQL(pSQL As String, db As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim elemento As clsMyAAnotador

Dim registros As New Collection

    Set rstQuery = db.query(pSQL)

    Do While Not rstQuery.EOF
        Set elemento = New clsMyAAnotador
        fillObject elemento, rstQuery

        registros.add elemento

        rstQuery.MoveNext
    Loop

    rstQuery.Close

    Set collectionSQL = registros

End Function

Public Sub fillCombo(pCombo As ComboBox, db As clsDB)

    fillComboCollection pCombo, collectionAll(db)

End Sub

Private Sub fillComboCollection(pCombo As ComboBox, pCollection As Collection)
Dim elemento As clsMyAAnotador

    pCombo.Clear

    For Each elemento In pCollection
        pCombo.AddItem elemento.comboText
        pCombo.ItemData(pCombo.NewIndex) = elemento.comboKey
    Next

    If pCombo.ListCount > 0 Then pCombo.ListIndex = 0

End Sub

Public Sub fillList(pList As ListBox, db As clsDB)

    fillListCollection pList, collectionAll(db)

End Sub

Private Sub fillListCollection(pList As ListBox, pCollection As Collection)
Dim elemento As clsMyAAnotador

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

