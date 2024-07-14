VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyAClienteDato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vClienteID As Long
Private vDocumento As Currency
Private vEmail As String
Private vFijo As String
Private vCelular As String
Private vUid As String
Private vCreated As Date

Private clon As clsMyAClienteDato

Public ddl As New clsMyDClienteDato

Private Sub class_Initialize()

    vClienteID = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vDocumento = 0
    vEmail = ""
    vFijo = ""
    vCelular = ""
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

Public Property Get documento() As Variant

    documento = vDocumento

End Property

Public Property Let documento(ByVal vNewValue As Variant)

On Error Resume Next

    vDocumento = vNewValue

End Property

Public Property Get email() As Variant

    email = vEmail

End Property

Public Property Let email(ByVal vNewValue As Variant)

On Error Resume Next

    vEmail = vNewValue

End Property

Public Property Get fijo() As Variant

    fijo = vFijo

End Property

Public Property Let fijo(ByVal vNewValue As Variant)

On Error Resume Next

    vFijo = vNewValue

End Property

Public Property Get celular() As Variant

    celular = vCelular

End Property

Public Property Let celular(ByVal vNewValue As Variant)

On Error Resume Next

    vCelular = vNewValue

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

    comboText = vDocumento

End Property

Public Property Get comboKey() As Variant

    comboKey = vClienteID

End Property

Public Property Get listText() As Variant

    listText = vDocumento

End Property

Public Property Get listKey() As Variant

    listKey = vClienteID

End Property

Public Function clone() As clsMyAClienteDato
Dim objObjeto As New clsMyAClienteDato

    With objObjeto
        .clienteID = vClienteID
        .documento = vDocumento
        .email = vEmail
        .fijo = vFijo
        .celular = vCelular
        .uid = vUid
        .created = vCreated
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsMyAClienteDato, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .clienteID = pQuery.Fields(.ddl.clienteID)
        .documento = pQuery.Fields(.ddl.documento)
        .email = pQuery.Fields(.ddl.email)
        .fijo = pQuery.Fields(.ddl.fijo)
        .celular = pQuery.Fields(.ddl.celular)
        .uid = pQuery.Fields(.ddl.uid)
        .created = pQuery.Fields(.ddl.created)
    End With

End Sub

Private Function makeWherePrimaryKey() As String
Dim sql As String

    sql = ""
    sql = sql & " WHERE " & ddl.clienteID & " = " & vClienteID

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
    sql = sql & " ORDER BY " & ddl.clienteID & " DESC"
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
    sql = sql & " SET " & ddl.documento & " = " & vDocumento
    sql = sql & ", " & ddl.email & " = '" & replaceSpecialSymbols(vEmail) & "'"
    sql = sql & ", " & ddl.fijo & " = '" & replaceSpecialSymbols(vFijo) & "'"
    sql = sql & ", " & ddl.celular & " = '" & replaceSpecialSymbols(vCelular) & "'"
    sql = sql & ", " & ddl.uid & " = '" & replaceSpecialSymbols(vUid) & "'"
    sql = sql & ", " & ddl.clienteID & " = " & vClienteID
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

    If vDocumento <> clon.documento Or IsNull(vDocumento) <> IsNull(clon.documento) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.documento & " = " & vDocumento
        blnChange = True
    End If
    If vEmail <> clon.email Or IsNull(vEmail) <> IsNull(clon.email) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.email & " = '" & replaceSpecialSymbols(vEmail) & "'"
        blnChange = True
    End If
    If vFijo <> clon.fijo Or IsNull(vFijo) <> IsNull(clon.fijo) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fijo & " = '" & replaceSpecialSymbols(vFijo) & "'"
        blnChange = True
    End If
    If vCelular <> clon.celular Or IsNull(vCelular) <> IsNull(clon.celular) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.celular & " = '" & replaceSpecialSymbols(vCelular) & "'"
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

    vClienteID = 1

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " ORDER BY " & ddl.clienteID & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set rstQuery = db.query(sql)

    If Not rstQuery.EOF Then vClienteID = 1 + rstQuery.Fields(ddl.clienteID)

    rstQuery.Close

    If pFind Then findByPrimaryKey db

End Sub

Public Function collectionAll(db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & ";"

    Set collectionAll = collectionSQL(sql, db)

End Function

Public Function collectionSQL(pSQL As String, db As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim elemento As clsMyAClienteDato

Dim registros As New Collection

    Set rstQuery = db.query(pSQL)

    Do While Not rstQuery.EOF
        Set elemento = New clsMyAClienteDato
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
Dim elemento As clsMyAClienteDato

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
Dim elemento As clsMyAClienteDato

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

