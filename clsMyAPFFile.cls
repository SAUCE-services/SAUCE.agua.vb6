VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyAPFFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vFileName As String
Private vPath As String
Private vImport As Integer
Private vFechaImport As Variant
Private vAutoID As Long
Private vUid As String
Private vCreated As Date

Private clon As clsMyAPFFile

Public ddl As New clsMyDPFFile

Private Sub class_Initialize()

    vFileName = ""

    initializeOther

End Sub

Private Sub initializeOther()

    vPath = ""
    vImport = 0
    vFechaImport = Null
    vAutoID = 0
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

Public Property Get path() As Variant

    path = vPath

End Property

Public Property Let path(ByVal vNewValue As Variant)

On Error Resume Next

    vPath = vNewValue

End Property

Public Property Get import() As Variant

    import = vImport

End Property

Public Property Let import(ByVal vNewValue As Variant)

On Error Resume Next

    vImport = vNewValue

End Property

Public Property Get fechaImport() As Variant

    fechaImport = vFechaImport

End Property

Public Property Let fechaImport(ByVal vNewValue As Variant)

On Error Resume Next

    vFechaImport = vNewValue

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

    comboText = vPath

End Property

Public Property Get comboKey() As Variant

    comboKey = vFileName

End Property

Public Property Get listText() As Variant

    listText = vPath

End Property

Public Property Get listKey() As Variant

    listKey = vFileName

End Property

Public Function clone() As clsMyAPFFile
Dim objObjeto As New clsMyAPFFile

    With objObjeto
        .fileName = vFileName
        .path = vPath
        .import = vImport
        .fechaImport = vFechaImport
        .autoID = vAutoID
        .uid = vUid
        .created = vCreated
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsMyAPFFile, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .fileName = pQuery.Fields(.ddl.fileName)
        .path = pQuery.Fields(.ddl.path)
        .import = pQuery.Fields(.ddl.import)
        .fechaImport = pQuery.Fields(.ddl.fechaImport)
        .autoID = pQuery.Fields(.ddl.autoID)
        .uid = pQuery.Fields(.ddl.uid)
        .created = pQuery.Fields(.ddl.created)
    End With

End Sub

Private Function makeWherePrimaryKey() As String
Dim sql As String

    sql = ""
    sql = sql & " WHERE " & ddl.fileName & " = '" & replaceSpecialSymbols(vFileName) & "'"

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
    sql = sql & " ORDER BY " & ddl.fileName & " DESC"
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
    sql = sql & " SET " & ddl.path & " = '" & replaceSpecialSymbols(vPath) & "'"
    sql = sql & ", " & ddl.import & " = " & vImport
    sql = sql & ", " & ddl.fechaImport & " = " & fechaDB(vFechaImport)
    sql = sql & ", " & ddl.uid & " = '" & replaceSpecialSymbols(vUid) & "'"
    sql = sql & ", " & ddl.fileName & " = '" & replaceSpecialSymbols(vFileName) & "'"
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

    If vPath <> clon.path Or IsNull(vPath) <> IsNull(clon.path) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.path & " = '" & replaceSpecialSymbols(vPath) & "'"
        blnChange = True
    End If
    If vImport <> clon.import Or IsNull(vImport) <> IsNull(clon.import) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.import & " = " & vImport
        blnChange = True
    End If
    If vFechaImport <> clon.fechaImport Or IsNull(vFechaImport) <> IsNull(clon.fechaImport) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fechaImport & " = " & fechaDB(vFechaImport)
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

    vFileName = 1

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " ORDER BY " & ddl.fileName & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set rstQuery = db.query(sql)

    If Not rstQuery.EOF Then vFileName = 1 + rstQuery.Fields(ddl.fileName)

    rstQuery.Close

    If pFind Then findByPrimaryKey db

End Sub

Public Function collectionAll(db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & ";"

    Set collectionAll = collectionSQL(sql, db)

End Function

Public Function collectionByFileName(pFileName As String, pDB As clsDB) As Collection
Dim strSQL As String

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.fileName & " = '" & pFileName & "'"
    strSQL = strSQL & " ORDER BY " & ddl.autoID
    strSQL = strSQL & ";"

    Set collectionByFileName = collectionSQL(strSQL, pDB)

End Function

Public Function collectionSQL(pSQL As String, db As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim elemento As clsMyAPFFile

Dim registros As New Collection

    Set rstQuery = db.query(pSQL)

    Do While Not rstQuery.EOF
        Set elemento = New clsMyAPFFile
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
Dim elemento As clsMyAPFFile

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
Dim elemento As clsMyAPFFile

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

