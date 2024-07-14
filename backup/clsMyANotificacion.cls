VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyANotificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vClienteID As Long
Private vFecha As Date
Private vTiponotificacionID As Integer
Private vNotificacionID As Long
Private vCreated As Date
Private vUpdated As Date
Private vUuid As String

Private clon As clsMyANotificacion

Public ddl As New clsMyDNotificacion

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Inicializacion de la Clase
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub class_Initialize()

    vClienteID = 0
    vFecha = Date

    initializeOther

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Inicializacion de los campos que no corresponden a la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub initializeOther()

    vTiponotificacionID = 0
    vNotificacionID = 0
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
Public Property Get clienteID() As Variant

    clienteID = vClienteID

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Setter clienteID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let clienteID(ByVal vNewValue As Variant)

On Error Resume Next

    vClienteID = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Getter fecha
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fecha() As Variant

    fecha = vFecha

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Setter fecha
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let fecha(ByVal vNewValue As Variant)

On Error Resume Next

    vFecha = vNewValue

End Property

Public Property Get key() As String

    key = "k." & clienteID & "." & fecha
    
End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Getter tiponotificacionID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get tiponotificacionID() As Variant

    tiponotificacionID = vTiponotificacionID

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Setter tiponotificacionID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let tiponotificacionID(ByVal vNewValue As Variant)

On Error Resume Next

    vTiponotificacionID = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Getter notificacionID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get notificacionID() As Variant

    notificacionID = vNotificacionID

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Setter notificacionID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let notificacionID(ByVal vNewValue As Variant)

On Error Resume Next

    vNotificacionID = vNewValue

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

    comboText = vTiponotificacionID

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Getter comboKey
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comboKey() As Variant

    comboKey = vFecha

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Getter listText
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get listText() As Variant

    listText = vTiponotificacionID

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Getter listKey
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get listKey() As Variant

    listKey = vFecha

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Clona el objeto this y devuelve uno nuevo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function clone() As clsMyANotificacion
Dim objObjeto As New clsMyANotificacion

    With objObjeto
        .clienteID = vClienteID
        .fecha = vFecha
        .tiponotificacionID = vTiponotificacionID
        .notificacionID = vNotificacionID
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
Private Sub fillObject(pObjeto As clsMyANotificacion, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .clienteID = pQuery.Fields(.ddl.clienteID)
        .fecha = pQuery.Fields(.ddl.fecha)
        .tiponotificacionID = pQuery.Fields(.ddl.tiponotificacionID)
        .notificacionID = pQuery.Fields(.ddl.notificacionID)
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
    sql = sql & " WHERE " & ddl.clienteID & " = " & vClienteID
    sql = sql & " AND " & ddl.fecha & " = " & fechaDB(vFecha)

    makeWherePrimaryKey = sql

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Consulta la DB con un SQL buscando 1 registro
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
' Description: Consulta la DB con un SQL buscando el último registro según la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub findLast(db As clsDB)
Dim sql As String

    initializeOther

    sql = "SELECT DISTINCTROW " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteID & " = " & vClienteID
    sql = sql & " ORDER BY " & ddl.fecha & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, db

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Consulta la DB con un SQL buscando 1 registro de acuerdo a los atributos cargados
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub findByNotificacionID(db As clsDB)
Dim sql As String

Dim lngNotificacionID As Long

    lngNotificacionID = vNotificacionID

    clean

    vNotificacionID = lngNotificacionID

    sql = "SELECT DISTINCTROW " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.notificacionID & " = " & vNotificacionID
    sql = sql & ";"

    findSQL sql, db

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Consulta la DB con un SQL buscando 1 registro de acuerdo a los atributos cargados
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub findByUuid(db As clsDB)
Dim sql As String

Dim uuid As String

    uuid = vUuid

    clean

    vUuid = uuid

    sql = "SELECT DISTINCTROW " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.uuid & " = REPLACE(UUID(), '-', '')"
    sql = sql & ";"

    findSQL sql, db

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Elimina 1 registro de la DB según la Primary Key
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
    sql = sql & " SET " & ddl.tiponotificacionID & " = " & vTiponotificacionID
    sql = sql & ", " & ddl.created & " = " & "Now()"
    sql = sql & ", " & ddl.uuid & " = REPLACE(UUID(), '-', '')"
    sql = sql & ", " & ddl.clienteID & " = " & vClienteID
    sql = sql & ", " & ddl.fecha & " = " & fechaDB(vFecha)
    sql = sql & ";"

    add = db.execute(sql)

    findByPrimaryKey db

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Hace 1 UPDATE sólo de los atributos que cambiaron respecto del almacenado en la DB
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

    If vTiponotificacionID <> clon.tiponotificacionID Or IsNull(vTiponotificacionID) <> IsNull(clon.tiponotificacionID) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.tiponotificacionID & " = " & vTiponotificacionID
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

    vFecha = 1

    sql = "SELECT DISTINCTROW " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteID & " = " & vClienteID
    sql = sql & " ORDER BY " & ddl.fecha & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set rstQuery = db.query(sql)

    If Not rstQuery.EOF Then vFecha = 1 + rstQuery.Fields(ddl.fecha)

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

Public Function collectionByFecha(fecha As Date, db As clsDB, Optional registrosold As Collection = Nothing) As Collection
Dim sql As String

    sql = "SELECT DISTINCTROW " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.fecha & " = " & fechaDB(fecha)
    sql = sql & ";"

    Set collectionByFecha = collectionSQL(sql, db, registrosold)

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Hace la consulta SQL sobre la DB y arma la colección
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function collectionSQL(sql As String, db As clsDB, Optional registrosold As Collection = Nothing) As Collection
Dim query As ADODB.Recordset

Dim elemento As clsMyANotificacion

Dim registros As Collection

    If registrosold Is Nothing Then
        Set registros = New Collection
    Else
        Set registros = registrosold
    End If

    Set query = db.query(sql)

    Do While Not query.EOF
        Set elemento = New clsMyANotificacion
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
Dim elemento As clsMyANotificacion

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
Dim elemento As clsMyANotificacion

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

