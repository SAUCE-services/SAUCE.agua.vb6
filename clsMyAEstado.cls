VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyAEstado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vEstadoID As Integer
Private vNombre As String
Private vAutoID As Long
Private vCreated As Date

Private clon As clsMyAEstado

Public ddl As New clsMyDEstado

Private Sub class_Initialize()

    vEstadoID = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vNombre = ""
    vAutoID = 0
    vCreated = Date

End Sub

Public Sub clean()

    Call class_Initialize

End Sub

Public Property Get estadoID() As Variant

    estadoID = vEstadoID

End Property

Public Property Let estadoID(ByVal vNewValue As Variant)

On Error Resume Next

    vEstadoID = vNewValue

End Property

Public Property Get key() As String

    key = "k." & estadoID
    
End Property

Public Property Get nombre() As Variant

    nombre = vNombre

End Property

Public Property Let nombre(ByVal vNewValue As Variant)

On Error Resume Next

    vNombre = vNewValue

End Property

Public Property Get autoID() As Variant

    autoID = vAutoID

End Property

Public Property Let autoID(ByVal vNewValue As Variant)

On Error Resume Next

    vAutoID = vNewValue

End Property

Public Property Get created() As Variant

    created = vCreated

End Property

Public Property Let created(ByVal vNewValue As Variant)

On Error Resume Next

    vCreated = vNewValue

End Property

Public Property Get comboText() As Variant

    comboText = vNombre & " (" & vEstadoID & ")"

End Property

Public Property Get comboKey() As Variant

    comboKey = vEstadoID

End Property

Public Property Get listText() As Variant

    listText = vNombre & " (" & vEstadoID & ")"

End Property

Public Property Get listKey() As Variant

    listKey = vEstadoID

End Property

Public Function clone() As clsMyAEstado
Dim objObjeto As New clsMyAEstado

    With objObjeto
        .estadoID = vEstadoID
        .nombre = vNombre
        .autoID = vAutoID
        .created = vCreated
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsMyAEstado, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .estadoID = pQuery.Fields(.ddl.estadoID)
        .nombre = pQuery.Fields(.ddl.nombre)
        .autoID = pQuery.Fields(.ddl.autoID)
        .created = pQuery.Fields(.ddl.created)
    End With

End Sub

Private Function makeWherePrimaryKey() As String
Dim sql As String

    sql = ""
    sql = sql & " WHERE " & ddl.estadoID & " = " & vEstadoID

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
    sql = sql & " ORDER BY " & ddl.estadoID & " DESC"
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
    sql = sql & " SET " & ddl.nombre & " = '" & replaceSpecialSymbols(vNombre) & "'"
    sql = sql & ", " & ddl.estadoID & " = " & vEstadoID
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

    If vNombre <> clon.nombre Or IsNull(vNombre) <> IsNull(clon.nombre) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.nombre & " = '" & replaceSpecialSymbols(vNombre) & "'"
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

    vEstadoID = 1

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " ORDER BY " & ddl.estadoID & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set rstQuery = db.query(sql)

    If Not rstQuery.EOF Then vEstadoID = 1 + rstQuery.Fields(ddl.estadoID)

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

Public Function collectionSQL(pSQL As String, db As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim elemento As clsMyAEstado

Dim registros As New Collection

    Set rstQuery = db.query(pSQL)

    Do While Not rstQuery.EOF
        Set elemento = New clsMyAEstado
        fillObject elemento, rstQuery

        registros.add elemento, elemento.key

        rstQuery.MoveNext
    Loop

    rstQuery.Close

    Set collectionSQL = registros

End Function

Public Sub fillCombo(pCombo As ComboBox, db As clsDB)

    fillComboCollection pCombo, collectionAll(db)

End Sub

Private Sub fillComboCollection(pCombo As ComboBox, pCollection As Collection)
Dim elemento As clsMyAEstado

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
Dim elemento As clsMyAEstado

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub
