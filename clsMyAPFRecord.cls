VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyAPFRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vFileName As String
Private vLine As String
Private vAutoID As Long
Private vUid As String
Private vCreated As Date

Private clon As clsMyAPFRecord

Public ddl As New clsMyDPFRecord

Private Sub class_Initialize()

    vAutoID = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vFileName = ""
    vLine = ""
    vUid = ""
    vCreated = Date

End Sub

Public Sub clean()

    Call class_Initialize

End Sub

Public Property Get fileName() As Variant

    fileName = vFileName

End Property

Public Property Let fileName(ByVal vNewValue As Variant)

On Error Resume Next

    vFileName = vNewValue

End Property

Public Property Get line() As Variant

    line = vLine

End Property

Public Property Let line(ByVal vNewValue As Variant)

On Error Resume Next

    vLine = vNewValue

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

    comboText = vFileName

End Property

Public Property Get comboKey() As Variant

    comboKey = vAutoID

End Property

Public Property Get listText() As Variant

    listText = vFileName

End Property

Public Property Get listKey() As Variant

    listKey = vAutoID

End Property

Public Function clone() As clsMyAPFRecord
Dim objObjeto As New clsMyAPFRecord

    With objObjeto
        .fileName = vFileName
        .line = vLine
        .autoID = vAutoID
        .uid = vUid
        .created = vCreated
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsMyAPFRecord, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .fileName = pQuery.Fields(.ddl.fileName)
        .line = pQuery.Fields(.ddl.line)
        .autoID = pQuery.Fields(.ddl.autoID)
        .uid = pQuery.Fields(.ddl.uid)
        .created = pQuery.Fields(.ddl.created)
    End With

End Sub

Private Function makeWherePrimaryKey() As String
Dim sql As String

    sql = ""
    sql = sql & " WHERE " & ddl.autoID & " = " & vAutoID

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

    findByPrimaryKey_ IIf(IsMissing(db), dbapp, db)

End Sub

Private Sub findByPrimaryKey_(db As clsDB)
Dim sql As String

    initializeOther

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & makeWherePrimaryKey
    sql = sql & ";"

    findSQL sql, db

End Sub

Public Sub findLastID(db As clsDB)

    findLastID_ IIf(IsMissing(db), dbapp, db)

End Sub

Private Sub findLastID_(db As clsDB)
Dim sql As String

Dim rstQuery As ADODB.Recordset

    clean

    sql = "SELECT LAST_INSERT_ID() AS " & ddl.autoID
    sql = sql & ";"

    Set rstQuery = db.query(sql)

    If Not rstQuery.EOF Then vAutoID = rstQuery.Fields(ddl.autoID)

    rstQuery.Close

    findByPrimaryKey_ db

End Sub

Public Sub findLast(db As clsDB)

    findLast_ IIf(IsMissing(db), dbapp, db)

End Sub

Private Sub findLast_(db As clsDB)
Dim sql As String

    initializeOther

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " ORDER BY " & ddl.autoID & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, db

End Sub

Public Function delete(db As clsDB) As Boolean

    delete = delete_(IIf(IsMissing(db), dbapp, db))

End Function

Private Function delete_(db As clsDB) As Boolean
Dim sql As String

    sql = "DELETE FROM " & ddl.table
    sql = sql & makeWherePrimaryKey
    sql = sql & ";"

    delete_ = db.execute(sql)

End Function

Public Function save(db As clsDB) As Boolean

    save = save_(IIf(IsMissing(db), dbapp, db))

End Function

Private Function save_(db As clsDB) As Boolean

    If exist_(db) Then
        save_ = update_(db)
    Else
        save_ = add_(db)
    End If

End Function

Public Function exist(db As clsDB) As Boolean

    exist = exist_(IIf(IsMissing(db), dbapp, db))

End Function

Private Function exist_(db As clsDB) As Boolean
Dim sql As String

Dim rstQuery As ADODB.Recordset

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & makeWherePrimaryKey
    sql = sql & ";"

    Set rstQuery = db.query(sql)
    exist_ = Not rstQuery.EOF
    rstQuery.Close

End Function

Public Function add(db As clsDB) As Boolean

    add = add_(IIf(IsMissing(db), dbapp, db))

End Function

Private Function add_(db As clsDB) As Boolean
Dim sql As String

    sql = "INSERT INTO " & ddl.table
    sql = sql & " SET " & ddl.fileName & " = '" & replaceSpecialSymbols(vFileName) & "'"
    sql = sql & ", " & ddl.line & " = '" & replaceSpecialSymbols(vLine) & "'"
    sql = sql & ", " & ddl.uid & " = '" & replaceSpecialSymbols(vUid) & "'"
    sql = sql & ";"

    add_ = db.execute(sql)

    findLastID_ db

End Function

Public Function update(db As clsDB) As Boolean

    update = update_(IIf(IsMissing(db), dbapp, db))

End Function

Private Function update_(db As clsDB) As Boolean
Dim sql As String
Dim strSET As String

Dim blnChange As Boolean

    update_ = True

    strSET = ""
    blnChange = False

    If clon Is Nothing Then
        Set clon = Me.clone
        clon.findByPrimaryKey db
    End If

    If vFileName <> clon.fileName Or IsNull(vFileName) <> IsNull(clon.fileName) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fileName & " = '" & replaceSpecialSymbols(vFileName) & "'"
        blnChange = True
    End If
    If vLine <> clon.line Or IsNull(vLine) <> IsNull(clon.line) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.line & " = '" & replaceSpecialSymbols(vLine) & "'"
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

    If blnChange Then update_ = db.execute(sql)

    Set clon = Nothing
    Set clon = Me.clone

End Function

Public Sub newID(pFind As Boolean, db As clsDB)

    newID_ pFind, IIf(IsMissing(db), dbapp, db)

End Sub

Private Sub newID_(pFind As Boolean, db As clsDB)
Dim sql As String

Dim rstQuery As ADODB.Recordset

    vAutoID = 1

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " ORDER BY " & ddl.autoID & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set rstQuery = db.query(sql)

    If Not rstQuery.EOF Then vAutoID = 1 + rstQuery.Fields(ddl.autoID)

    rstQuery.Close

    If pFind Then findByPrimaryKey_ db

End Sub

Public Function collectionAll(db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & ";"

    Set collectionAll = collectionSQL(sql, db)

End Function

Public Function collectionByFileName(pFileName As String, db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.fileName & " = '" & pFileName & "'"
    sql = sql & " ORDER BY " & ddl.autoID
    sql = sql & ";"

    Set collectionByFileName = collectionSQL(sql, db)

End Function

Public Function collectionSQL(pSQL As String, db As clsDB) As Collection

    Set collectionSQL = collection_(pSQL, IIf(IsMissing(db), dbapp, db))

End Function

Private Function collection_(pSQL As String, db As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim elemento As clsMyAPFRecord

Dim registros As New Collection

    Set rstQuery = db.query(pSQL)

    Do While Not rstQuery.EOF
        Set elemento = New clsMyAPFRecord
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
Dim elemento As clsMyAPFRecord

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
Dim elemento As clsMyAPFRecord

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

