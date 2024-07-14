VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyAClienteNotif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vClienteId As Long
Private vUltimaNotificacion15 As Variant
Private vUltimaNotificacion48 As Variant
Private vUltimaNotificacionCorte As Variant
Private vAutoID As Long
Private vCreated As Date
Private vUpdated As Date
Private vUuid As String

Private clon As clsMyAClienteNotif

Public ddl As New clsMyDClienteNotif

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Inicializacion de la Clase
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub class_Initialize()

    vClienteId = 0

    initializeOther

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Inicializacion de los campos que no corresponden a la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub initializeOther()

    vUltimaNotificacion15 = Null
    vUltimaNotificacion48 = Null
    vUltimaNotificacionCorte = Null
    vAutoID = 0
    vCreated = Date
    vUpdated = Date
    vUuid = ""

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Limpia todos los atributos
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub clean()

    Call class_Initialize

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Getter clienteID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get clienteId() As Variant

    clienteId = vClienteId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Setter clienteID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let clienteId(ByVal vNewValue As Variant)

On Error Resume Next

    vClienteId = vNewValue

End Property

Public Property Get key() As String

    key = "k." & clienteId
    
End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Getter ultimaNotificacion15
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get ultimaNotificacion15() As Variant

    ultimaNotificacion15 = vUltimaNotificacion15

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Setter ultimaNotificacion15
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let ultimaNotificacion15(ByVal vNewValue As Variant)

On Error Resume Next

    vUltimaNotificacion15 = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Getter ultimaNotificacion48
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get ultimaNotificacion48() As Variant

    ultimaNotificacion48 = vUltimaNotificacion48

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Setter ultimaNotificacion48
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let ultimaNotificacion48(ByVal vNewValue As Variant)

On Error Resume Next

    vUltimaNotificacion48 = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Getter ultimaNotificacionCorte
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get ultimaNotificacionCorte() As Variant

    ultimaNotificacionCorte = vUltimaNotificacionCorte

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Setter ultimaNotificacionCorte
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let ultimaNotificacionCorte(ByVal vNewValue As Variant)

On Error Resume Next

    vUltimaNotificacionCorte = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Getter autoID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get autoID() As Variant

    autoID = vAutoID

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Setter autoID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let autoID(ByVal vNewValue As Variant)

On Error Resume Next

    vAutoID = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Getter created
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get created() As Variant

    created = vCreated

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Setter created
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let created(ByVal vNewValue As Variant)

On Error Resume Next

    vCreated = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Getter updated
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get updated() As Variant

    updated = vUpdated

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Setter updated
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let updated(ByVal vNewValue As Variant)

On Error Resume Next

    vUpdated = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Getter uuid
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get uuid() As Variant

    uuid = vUuid

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Setter uuid
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let uuid(ByVal vNewValue As Variant)

On Error Resume Next

    vUuid = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Getter comboText
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comboText() As Variant

    comboText = vUltimaNotificacion15

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Getter comboKey
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comboKey() As Variant

    comboKey = vClienteId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Getter listText
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get listText() As Variant

    listText = vUltimaNotificacion15

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Getter listKey
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get listKey() As Variant

    listKey = vClienteId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Clona el objeto this y devuelve uno nuevo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function clone() As clsMyAClienteNotif
Dim objObjeto As New clsMyAClienteNotif

    With objObjeto
        .clienteId = vClienteId
        .ultimaNotificacion15 = vUltimaNotificacion15
        .ultimaNotificacion48 = vUltimaNotificacion48
        .ultimaNotificacionCorte = vUltimaNotificacionCorte
        .autoID = vAutoID
        .created = vCreated
        .updated = vUpdated
        .uuid = vUuid
    End With

    Set clone = objObjeto

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Carga los campos del registro en los atributos de la clase
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub fillObject(pObjeto As clsMyAClienteNotif, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .clienteId = pQuery.Fields(.ddl.clienteId)
        .ultimaNotificacion15 = pQuery.Fields(.ddl.ultimaNotificacion15)
        .ultimaNotificacion48 = pQuery.Fields(.ddl.ultimaNotificacion48)
        .ultimaNotificacionCorte = pQuery.Fields(.ddl.ultimaNotificacionCorte)
        .autoID = pQuery.Fields(.ddl.autoID)
        .created = pQuery.Fields(.ddl.created)
        .updated = pQuery.Fields(.ddl.updated)
        .uuid = pQuery.Fields(.ddl.uuid)
    End With

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Construye el SQL de la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Function makeWherePrimaryKey() As String
Dim sql As String

    sql = ""
    sql = sql & " WHERE " & ddl.clienteId & " = " & vClienteId

    makeWherePrimaryKey = sql

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Consulta la dbapp con un SQL buscando 1 registro
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub findSQL(pSQL As String, db As clsDB)
Dim rstQuery As ADODB.Recordset

    Set rstQuery = db.query(pSQL)

    If Not rstQuery.EOF Then
        fillObject Me, rstQuery

        Set clon = Me.clone
    End If

    rstQuery.Close

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Realiza un find sobre los atributos de la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub findByPrimaryKey(db As clsDB)
Dim sql As String

    initializeOther

    sql = "SELECT DISTINCTROW " & ddl.table & ".* FROM " & ddl.table
    sql = sql & makeWherePrimaryKey
    sql = sql & ";"

    findSQL sql, db

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Consulta la dbapp con un SQL buscando el último registro según la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub findLast(db As clsDB)
Dim sql As String

    initializeOther

    sql = "SELECT DISTINCTROW " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " ORDER BY " & ddl.clienteId & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, db

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Consulta la dbapp con un SQL buscando 1 registro de acuerdo a los atributos cargados
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub findByAutoID(db As clsDB)
Dim sql As String

Dim lngAutoID As Long

    lngAutoID = vAutoID

    clean

    vAutoID = lngAutoID

    sql = "SELECT DISTINCTROW " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.autoID & " = " & vAutoID
    sql = sql & ";"

    findSQL sql, db

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Elimina 1 registro de la dbapp según la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function delete(db As clsDB) As Boolean
Dim sql As String

    sql = "DELETE FROM " & ddl.table
    sql = sql & makeWherePrimaryKey
    sql = sql & ";"

    delete = db.execute(sql)

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Graba un registro verificando si debe agregarlo o actualizarlo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function save(db As clsDB) As Boolean

    If exist(db) Then
        save = update(db)
    Else
        save = add(db)
    End If

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Verifica si el registro asociado a la Primary Key existe
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function exist(db As clsDB) As Boolean
Dim sql As String

Dim rstQuery As ADODB.Recordset

    sql = "SELECT DISTINCTROW " & ddl.table & ".* FROM " & ddl.table
    sql = sql & makeWherePrimaryKey
    sql = sql & ";"

    Set rstQuery = db.query(sql)
    exist = Not rstQuery.EOF
    rstQuery.Close

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Hace 1 INSERT con los valores de los atributos
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function add(db As clsDB) As Boolean
Dim sql As String

    sql = "INSERT INTO " & ddl.table
    sql = sql & " SET " & ddl.ultimaNotificacion15 & " = " & fechaDB(vUltimaNotificacion15)
    sql = sql & ", " & ddl.ultimaNotificacion48 & " = " & fechaDB(vUltimaNotificacion48)
    sql = sql & ", " & ddl.ultimaNotificacionCorte & " = " & fechaDB(vUltimaNotificacionCorte)
    sql = sql & ", " & ddl.created & " = " & "Now()"
    sql = sql & ", " & ddl.uuid & " = REPLACE(UUID(), '-', '')"
    sql = sql & ", " & ddl.clienteId & " = " & vClienteId
    sql = sql & ";"

    add = db.execute(sql)

    findByPrimaryKey db

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Hace 1 UPDATE sólo de los atributos que cambiaron respecto del almacenado en la dbapp
' Log:
' * * * * * * * * * * * * * * * * * * * * *
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

    If vUltimaNotificacion15 <> clon.ultimaNotificacion15 Or IsNull(vUltimaNotificacion15) <> IsNull(clon.ultimaNotificacion15) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.ultimaNotificacion15 & " = " & fechaDB(vUltimaNotificacion15)
        blnChange = True
    End If
    If vUltimaNotificacion48 <> clon.ultimaNotificacion48 Or IsNull(vUltimaNotificacion48) <> IsNull(clon.ultimaNotificacion48) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.ultimaNotificacion48 & " = " & fechaDB(vUltimaNotificacion48)
        blnChange = True
    End If
    If vUltimaNotificacionCorte <> clon.ultimaNotificacionCorte Or IsNull(vUltimaNotificacionCorte) <> IsNull(clon.ultimaNotificacionCorte) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.ultimaNotificacionCorte & " = " & fechaDB(vUltimaNotificacionCorte)
        blnChange = True
    End If
    If vUuid = "" Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.uuid & " = REPLACE(UUID(), '-', '')"
        blnChange = True
    End If

    sql = "UPDATE " & ddl.table & " SET" & strSET
    sql = sql & makeWherePrimaryKey
    sql = sql & ";"

    If blnChange Then update = db.execute(sql)

    Set clon = Nothing
    Set clon = Me.clone

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Busca un nuevo ID según la Primary Key y lo carga al objeto si se le indica como parámetro
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub newID(pFind As Boolean, db As clsDB)
Dim sql As String

Dim rstQuery As ADODB.Recordset

    vClienteId = 1

    sql = "SELECT DISTINCTROW " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " ORDER BY " & ddl.clienteId & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set rstQuery = db.query(sql)

    If Not rstQuery.EOF Then vClienteId = 1 + rstQuery.Fields(ddl.clienteId)

    rstQuery.Close

    If pFind Then findByPrimaryKey db

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Arma una colección con todos los registros
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function collectionAll(db As clsDB, Optional registrosold As Collection = Nothing) As Collection
Dim sql As String

    sql = "SELECT DISTINCTROW " & ddl.table & ".* FROM " & ddl.table
    sql = sql & ";"

    Set collectionAll = collectionSQL(sql, db, registrosold)

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Arma una colección con 1 registro para verificar si algo existe en la tabla
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function collectionAny(db As clsDB) As Collection
Dim sql As String

    sql = "SELECT DISTINCTROW " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set collectionAny = collectionSQL(sql, db)

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Hace la consulta SQL sobre la dbapp y arma la colección
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function collectionSQL(sql As String, db As clsDB, Optional registrosold As Collection = Nothing) As Collection
Dim query As ADODB.Recordset

Dim elemento As clsMyAClienteNotif

Dim registros As Collection

    If registrosold Is Nothing Then
        Set registros = New Collection
    Else
        Set registros = registrosold
    End If

    Set query = db.query(sql)

    Do While Not query.EOF
        Set elemento = New clsMyAClienteNotif
        fillObject elemento, query

        If Not modCollection.collectionExistElement(registros, elemento.key) Then registros.add elemento, elemento.key

        query.MoveNext
    Loop

    query.Close

    Set collectionSQL = registros

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Llena una ComboBox con todos los datos de la tabla
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub fillCombo(pCombo As ComboBox, db As clsDB, Optional registrosold As Collection = Nothing)

    fillComboCollection pCombo, collectionAll(db, registrosold)

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Llena la ComboBox efectivamente
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub fillComboCollection(pCombo As ComboBox, pCollection As Collection)
Dim elemento As clsMyAClienteNotif

    pCombo.Clear

    For Each elemento In pCollection
        pCombo.AddItem elemento.comboText
        pCombo.ItemData(pCombo.NewIndex) = elemento.comboKey
    Next

    If pCombo.ListCount > 0 Then pCombo.ListIndex = 0

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Llena una ListBox con todos los datos de la tabla
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub fillList(pList As ListBox, db As clsDB, Optional registrosold As Collection = Nothing)

    fillListCollection pList, collectionAll(db, registrosold)

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Llena la ListBox efectivamente
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub fillListCollection(pList As ListBox, pCollection As Collection)
Dim elemento As clsMyAClienteNotif

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

