VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyATipoNotif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vTiponotificacionId As Integer
Private vNombre As String
Private vValorSocio As Currency
Private vValorNoSocio As Currency
Private vCreated As Date
Private vUpdated As Date
Private vUuid As String

Private clon As clsMyATipoNotif

Public ddl As New clsMyDTipoNotif

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Inicializacion de la Clase
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub class_Initialize()

    vTiponotificacionId = 0

    initializeOther

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Inicializacion de los campos que no corresponden a la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub initializeOther()

    vNombre = ""
    vValorSocio = 0
    vValorNoSocio = 0
    vCreated = Date
    vUpdated = Date
    vUuid = ""

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Limpia todos los atributos
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub clean()

    Call class_Initialize

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Getter tiponotificacionID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get tiponotificacionId() As Variant

    tiponotificacionId = vTiponotificacionId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Setter tiponotificacionID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let tiponotificacionId(ByVal vNewValue As Variant)

On Error Resume Next

    vTiponotificacionId = vNewValue

End Property

Public Property Get key() As String

    key = "k." & tiponotificacionId
    
End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Getter nombre
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get nombre() As Variant

    nombre = vNombre

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Setter nombre
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let nombre(ByVal vNewValue As Variant)

On Error Resume Next

    vNombre = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Getter valorSocio
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get valorSocio() As Variant

    valorSocio = vValorSocio

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Setter valorSocio
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let valorSocio(ByVal vNewValue As Variant)

On Error Resume Next

    vValorSocio = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Getter valorNoSocio
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get valorNoSocio() As Variant

    valorNoSocio = vValorNoSocio

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Setter valorNoSocio
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let valorNoSocio(ByVal vNewValue As Variant)

On Error Resume Next

    vValorNoSocio = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Getter created
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get created() As Variant

    created = vCreated

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Setter created
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let created(ByVal vNewValue As Variant)

On Error Resume Next

    vCreated = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Getter updated
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get updated() As Variant

    updated = vUpdated

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Setter updated
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let updated(ByVal vNewValue As Variant)

On Error Resume Next

    vUpdated = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Getter uuid
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get uuid() As Variant

    uuid = vUuid

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Setter uuid
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let uuid(ByVal vNewValue As Variant)

On Error Resume Next

    vUuid = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Getter comboText
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comboText() As Variant

    comboText = vNombre

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Getter comboKey
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comboKey() As Variant

    comboKey = vTiponotificacionId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Getter listText
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get listText() As Variant

    listText = vNombre

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Getter listKey
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get listKey() As Variant

    listKey = vTiponotificacionId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Clona el objeto this y devuelve uno nuevo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function clone() As clsMyATipoNotif
Dim objObjeto As New clsMyATipoNotif

    With objObjeto
        .tiponotificacionId = vTiponotificacionId
        .nombre = vNombre
        .valorSocio = vValorSocio
        .valorNoSocio = vValorNoSocio
        .created = vCreated
        .updated = vUpdated
        .uuid = vUuid
    End With

    Set clone = objObjeto

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Carga los campos del registro en los atributos de la clase
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub fillObject(pObjeto As clsMyATipoNotif, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .tiponotificacionId = pQuery.Fields(.ddl.tiponotificacionId)
        .nombre = pQuery.Fields(.ddl.nombre)
        .valorSocio = pQuery.Fields(.ddl.valorSocio)
        .valorNoSocio = pQuery.Fields(.ddl.valorNoSocio)
        .created = pQuery.Fields(.ddl.created)
        .updated = pQuery.Fields(.ddl.updated)
        .uuid = pQuery.Fields(.ddl.uuid)
    End With

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Construye el SQL de la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Function makeWherePrimaryKey() As String
Dim sql As String

    sql = ""
    sql = sql & " WHERE " & ddl.tiponotificacionId & " = " & vTiponotificacionId

    makeWherePrimaryKey = sql

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Consulta la DB con un SQL buscando 1 registro
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub findSQL(pSQL As String, pDB As clsDB)
Dim rstQuery As ADODB.Recordset

    Set rstQuery = pDB.query(pSQL)

    If Not rstQuery.EOF Then
        fillObject Me, rstQuery

        Set clon = Me.clone
    End If

    rstQuery.Close

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Realiza un find sobre los atributos de la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub findByPrimaryKey(pDB As clsDB)
Dim sql As String

    initializeOther

    sql = "SELECT DISTINCTROW " & ddl.table & ".* FROM " & ddl.table
    sql = sql & makeWherePrimaryKey
    sql = sql & ";"

    findSQL sql, pDB

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Consulta la DB buscando el último autonumérico agregado
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub findLastID(pDB As clsDB)
Dim sql As String

Dim rstQuery As ADODB.Recordset

    clean

    sql = "SELECT LAST_INSERT_ID() AS " & ddl.tiponotificacionId
    sql = sql & ";"

    Set rstQuery = pDB.query(sql)

    If Not rstQuery.EOF Then vTiponotificacionId = rstQuery.Fields(ddl.tiponotificacionId)

    rstQuery.Close

    findByPrimaryKey pDB

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Consulta la DB con un SQL buscando el último registro según la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub findLast(pDB As clsDB)
Dim sql As String

    initializeOther

    sql = "SELECT DISTINCTROW " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " ORDER BY " & ddl.tiponotificacionId & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, pDB

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Consulta la DB con un SQL buscando 1 registro de acuerdo a los atributos cargados
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub findByUuid(pDB As clsDB)
Dim sql As String

Dim uuid As String

    uuid = vUuid

    clean

    vUuid = uuid

    sql = "SELECT DISTINCTROW " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.uuid & " = REPLACE(UUID(), '-', '')"
    sql = sql & ";"

    findSQL sql, pDB

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Elimina 1 registro de la DB según la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function delete(pDB As clsDB) As Boolean
Dim sql As String

    sql = "DELETE FROM " & ddl.table
    sql = sql & makeWherePrimaryKey
    sql = sql & ";"

    delete = pDB.execute(sql)

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Graba un registro verificando si debe agregarlo o actualizarlo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function save(pDB As clsDB) As Boolean

    If exist(pDB) Then
        save = update(pDB)
    Else
        save = add(pDB)
    End If

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Verifica si el registro asociado a la Primary Key existe
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function exist(pDB As clsDB) As Boolean
Dim sql As String

Dim rstQuery As ADODB.Recordset

    sql = "SELECT DISTINCTROW " & ddl.table & ".* FROM " & ddl.table
    sql = sql & makeWherePrimaryKey
    sql = sql & ";"

    Set rstQuery = pDB.query(sql)
    exist = Not rstQuery.EOF
    rstQuery.Close

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Hace 1 INSERT con los valores de los atributos
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function add(pDB As clsDB) As Boolean
Dim sql As String

    sql = "INSERT INTO " & ddl.table
    sql = sql & " SET " & ddl.nombre & " = '" & replaceSpecialSymbols(vNombre) & "'"
    sql = sql & ", " & ddl.valorSocio & " = '" & Replace(vValorSocio, ",", ".") & "'"
    sql = sql & ", " & ddl.valorNoSocio & " = '" & Replace(vValorNoSocio, ",", ".") & "'"
    sql = sql & ", " & ddl.created & " = " & "Now()"
    sql = sql & ", " & ddl.uuid & " = REPLACE(UUID(), '-', '')"
    sql = sql & ";"

    add = pDB.execute(sql)

    findLastID pDB

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Hace 1 UPDATE sólo de los atributos que cambiaron respecto del almacenado en la DB
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function update(pDB As clsDB) As Boolean
Dim sql As String
Dim strSET As String

Dim blnChange As Boolean

    update = True

    strSET = ""
    blnChange = False

    If clon Is Nothing Then
        Set clon = Me.clone
        clon.findByPrimaryKey pDB
    End If

    If vNombre <> clon.nombre Or IsNull(vNombre) <> IsNull(clon.nombre) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.nombre & " = '" & replaceSpecialSymbols(vNombre) & "'"
        blnChange = True
    End If
    If vValorSocio <> clon.valorSocio Or IsNull(vValorSocio) <> IsNull(clon.valorSocio) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.valorSocio & " = '" & Replace(vValorSocio, ",", ".") & "'"
        blnChange = True
    End If
    If vValorNoSocio <> clon.valorNoSocio Or IsNull(vValorNoSocio) <> IsNull(clon.valorNoSocio) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.valorNoSocio & " = '" & Replace(vValorNoSocio, ",", ".") & "'"
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

    If blnChange Then update = pDB.execute(sql)

    Set clon = Nothing
    Set clon = Me.clone

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Busca un nuevo ID según la Primary Key y lo carga al objeto si se le indica como parámetro
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub newID(pFind As Boolean, pDB As clsDB)
Dim sql As String

Dim rstQuery As ADODB.Recordset

    vTiponotificacionId = 1

    sql = "SELECT DISTINCTROW " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " ORDER BY " & ddl.tiponotificacionId & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set rstQuery = pDB.query(sql)

    If Not rstQuery.EOF Then vTiponotificacionId = 1 + rstQuery.Fields(ddl.tiponotificacionId)

    rstQuery.Close

    If pFind Then findByPrimaryKey pDB

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Arma una colección con todos los registros
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function collectionAll(pDB As clsDB, Optional registrosold As Collection = Nothing) As Collection
Dim sql As String

    sql = "SELECT DISTINCTROW " & ddl.table & ".* FROM " & ddl.table
    sql = sql & ";"

    Set collectionAll = collectionSQL(sql, pDB, registrosold)

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Arma una colección con 1 registro para verificar si algo existe en la tabla
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function collectionAny(pDB As clsDB) As Collection
Dim sql As String

    sql = "SELECT DISTINCTROW " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set collectionAny = collectionSQL(sql, pDB)

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Hace la consulta SQL sobre la DB y arma la colección
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function collectionSQL(sql As String, pDB As clsDB, Optional registrosold As Collection = Nothing) As Collection
Dim query As ADODB.Recordset

Dim elemento As clsMyATipoNotif

Dim registros As Collection

    If registrosold Is Nothing Then
        Set registros = New Collection
    Else
        Set registros = registrosold
    End If

    Set query = pDB.query(sql)

    Do While Not query.EOF
        Set elemento = New clsMyATipoNotif
        fillObject elemento, query

        If Not modCollection.collectionExistElement(registros, elemento.key) Then registros.add elemento, elemento.key

        query.MoveNext
    Loop

    query.Close

    Set collectionSQL = registros

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Llena una ComboBox con todos los datos de la tabla
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub fillCombo(pCombo As ComboBox, pDB As clsDB, Optional registrosold As Collection = Nothing)

    fillComboCollection pCombo, collectionAll(pDB, registrosold)

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Llena la ComboBox efectivamente
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub fillComboCollection(pCombo As ComboBox, pCollection As Collection)
Dim elemento As clsMyATipoNotif

    pCombo.Clear

    For Each elemento In pCollection
        pCombo.AddItem elemento.comboText
        pCombo.ItemData(pCombo.NewIndex) = elemento.comboKey
    Next

    If pCombo.ListCount > 0 Then pCombo.ListIndex = 0

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Llena una ListBox con todos los datos de la tabla
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub fillList(pList As ListBox, pDB As clsDB, Optional registrosold As Collection = Nothing)

    fillListCollection pList, collectionAll(pDB, registrosold)

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Llena la ListBox efectivamente
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub fillListCollection(pList As ListBox, pCollection As Collection)
Dim elemento As clsMyATipoNotif

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

