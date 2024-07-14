VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyAClienteVolumen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vClienteID As Long
Private vPeriodoID As Integer
Private vMedidorIDActual As String
Private vEstadoActual As Long
Private vMedidorIDAnterior As String
Private vEstadoAnterior As Long
Private vConsumido As Long
Private vAutoID As Long
Private vCreated As Date

Private clon As clsMyAClienteVolumen

Public ddl As New clsMyDClienteVolumen

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Inicializacion de la Clase
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub class_Initialize()

    vClienteID = 0
    vPeriodoID = 0

    initializeOther

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Inicializacion de los campos que no corresponden a la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub initializeOther()

    vMedidorIDActual = ""
    vEstadoActual = 0
    vMedidorIDAnterior = ""
    vEstadoAnterior = 0
    vConsumido = 0
    vAutoID = 0
    vCreated = Date

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Limpia todos los atributos
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub clean()

    Call class_Initialize

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Getter clienteID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get clienteID() As Variant

    clienteID = vClienteID

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Setter clienteID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let clienteID(ByVal vNewValue As Variant)

On Error Resume Next

    vClienteID = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Getter periodoID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get periodoID() As Variant

    periodoID = vPeriodoID

End Property

Public Property Get key() As String

    key = "k." & clienteID & "." & periodoID
    
End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Setter periodoID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let periodoID(ByVal vNewValue As Variant)

On Error Resume Next

    vPeriodoID = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Getter medidorIDActual
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get medidorIDActual() As Variant

    medidorIDActual = vMedidorIDActual

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Setter medidorIDActual
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let medidorIDActual(ByVal vNewValue As Variant)

On Error Resume Next

    vMedidorIDActual = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Getter estadoActual
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get estadoActual() As Variant

    estadoActual = vEstadoActual

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Setter estadoActual
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let estadoActual(ByVal vNewValue As Variant)

On Error Resume Next

    vEstadoActual = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Getter medidorIDAnterior
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get medidorIDAnterior() As Variant

    medidorIDAnterior = vMedidorIDAnterior

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Setter medidorIDAnterior
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let medidorIDAnterior(ByVal vNewValue As Variant)

On Error Resume Next

    vMedidorIDAnterior = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Getter estadoAnterior
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get estadoAnterior() As Variant

    estadoAnterior = vEstadoAnterior

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Setter estadoAnterior
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let estadoAnterior(ByVal vNewValue As Variant)

On Error Resume Next

    vEstadoAnterior = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Getter consumido
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get consumido() As Variant

    consumido = vConsumido

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Setter consumido
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let consumido(ByVal vNewValue As Variant)

On Error Resume Next

    vConsumido = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Getter autoID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get autoID() As Variant

    autoID = vAutoID

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Setter autoID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let autoID(ByVal vNewValue As Variant)

On Error Resume Next

    vAutoID = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Getter created
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get created() As Variant

    created = vCreated

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Setter created
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let created(ByVal vNewValue As Variant)

On Error Resume Next

    vCreated = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Getter comboText
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comboText() As Variant

    comboText = vMedidorIDActual

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Getter comboKey
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comboKey() As Variant

    comboKey = vPeriodoID

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Getter listText
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get listText() As Variant

    listText = vMedidorIDActual

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Getter listKey
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get listKey() As Variant

    listKey = vPeriodoID

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Clona el objeto this y devuelve uno nuevo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function clone() As clsMyAClienteVolumen
Dim objObjeto As New clsMyAClienteVolumen

    With objObjeto
        .clienteID = vClienteID
        .periodoID = vPeriodoID
        .medidorIDActual = vMedidorIDActual
        .estadoActual = vEstadoActual
        .medidorIDAnterior = vMedidorIDAnterior
        .estadoAnterior = vEstadoAnterior
        .consumido = vConsumido
        .autoID = vAutoID
        .created = vCreated
    End With

    Set clone = objObjeto

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Carga los campos del registro en los atributos de la clase
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub fillObject(pObjeto As clsMyAClienteVolumen, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .clienteID = pQuery.Fields(.ddl.clienteID)
        .periodoID = pQuery.Fields(.ddl.periodoID)
        .medidorIDActual = pQuery.Fields(.ddl.medidorIDActual)
        .estadoActual = pQuery.Fields(.ddl.estadoActual)
        .medidorIDAnterior = pQuery.Fields(.ddl.medidorIDAnterior)
        .estadoAnterior = pQuery.Fields(.ddl.estadoAnterior)
        .consumido = pQuery.Fields(.ddl.consumido)
        .autoID = pQuery.Fields(.ddl.autoID)
        .created = pQuery.Fields(.ddl.created)
    End With

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Construye el SQL de la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Function makeWherePrimaryKey() As String
Dim sql As String

    sql = ""
    sql = sql & " WHERE " & ddl.clienteID & " = " & vClienteID
    sql = sql & " AND " & ddl.periodoID & " = " & vPeriodoID

    makeWherePrimaryKey = sql

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
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
' Created: 15/04/2018
' Description: Realiza un find sobre los atributos de la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub findByPrimaryKey(db As clsDB)
Dim sql As String

    initializeOther

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & makeWherePrimaryKey
    sql = sql & ";"

    findSQL sql, db

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Consulta la dbapp con un SQL buscando el último registro según la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub findLast(db As clsDB)
Dim sql As String

    initializeOther

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteID & " = " & vClienteID
    sql = sql & " ORDER BY " & ddl.periodoID & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, db

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Consulta la dbapp con un SQL buscando 1 registro de acuerdo a los atributos cargados
' Log:
' * * * * * * * * * * * * * * * * * * * * *
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

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
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
' Created: 15/04/2018
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
' Created: 15/04/2018
' Description: Verifica si el registro asociado a la Primary Key existe
' Log:
' * * * * * * * * * * * * * * * * * * * * *
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

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Hace 1 INSERT con los valores de los atributos
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function add(db As clsDB) As Boolean
Dim sql As String

    sql = "INSERT INTO " & ddl.table
    sql = sql & " SET " & ddl.medidorIDActual & " = '" & replaceSpecialSymbols(vMedidorIDActual) & "'"
    sql = sql & ", " & ddl.estadoActual & " = " & vEstadoActual
    sql = sql & ", " & ddl.medidorIDAnterior & " = '" & replaceSpecialSymbols(vMedidorIDAnterior) & "'"
    sql = sql & ", " & ddl.estadoAnterior & " = " & vEstadoAnterior
    sql = sql & ", " & ddl.consumido & " = " & vConsumido
    sql = sql & ", " & ddl.clienteID & " = " & vClienteID
    sql = sql & ", " & ddl.periodoID & " = " & vPeriodoID
    sql = sql & ";"

    add = db.execute(sql)

    findByPrimaryKey db

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
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

    If vMedidorIDActual <> clon.medidorIDActual Or IsNull(vMedidorIDActual) <> IsNull(clon.medidorIDActual) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.medidorIDActual & " = '" & replaceSpecialSymbols(vMedidorIDActual) & "'"
        blnChange = True
    End If
    If vEstadoActual <> clon.estadoActual Or IsNull(vEstadoActual) <> IsNull(clon.estadoActual) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.estadoActual & " = " & vEstadoActual
        blnChange = True
    End If
    If vMedidorIDAnterior <> clon.medidorIDAnterior Or IsNull(vMedidorIDAnterior) <> IsNull(clon.medidorIDAnterior) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.medidorIDAnterior & " = '" & replaceSpecialSymbols(vMedidorIDAnterior) & "'"
        blnChange = True
    End If
    If vEstadoAnterior <> clon.estadoAnterior Or IsNull(vEstadoAnterior) <> IsNull(clon.estadoAnterior) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.estadoAnterior & " = " & vEstadoAnterior
        blnChange = True
    End If
    If vConsumido <> clon.consumido Or IsNull(vConsumido) <> IsNull(clon.consumido) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.consumido & " = " & vConsumido
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
' Created: 15/04/2018
' Description: Busca un nuevo ID según la Primary Key y lo carga al objeto si se le indica como parámetro
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub newID(pFind As Boolean, db As clsDB)
Dim sql As String

Dim rstQuery As ADODB.Recordset

    vPeriodoID = 1

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.clienteID & " = " & vClienteID
    sql = sql & " ORDER BY " & ddl.periodoID & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set rstQuery = db.query(sql)

    If Not rstQuery.EOF Then vPeriodoID = 1 + rstQuery.Fields(ddl.periodoID)

    rstQuery.Close

    If pFind Then findByPrimaryKey db

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Arma una colección con todos los registros
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function collectionAll(db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & ";"

    Set collectionAll = collectionSQL(sql, db)

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Arma una colección con 1 registro para verificar si algo existe en la tabla
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function collectionAny(db As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set collectionAny = collectionSQL(sql, db)

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Hace la consulta SQL sobre la dbapp y arma la colección
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function collectionSQL(sql As String, db As clsDB, Optional registrosold As Collection = Nothing) As Collection
Dim query As ADODB.Recordset

Dim elemento As clsMyAClienteVolumen

Dim registros As Collection

    If registrosold Is Nothing Then
        Set registros = New Collection
    Else
        Set registros = registrosold
    End If

    Set query = db.query(sql)

    Do While Not query.EOF
        Set elemento = New clsMyAClienteVolumen
        fillObject elemento, query

        If Not modCollection.collectionExistElement(registros, elemento.key) Then registros.add elemento, elemento.key

        query.MoveNext
    Loop

    query.Close

    Set collectionSQL = registros

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Llena una ComboBox con todos los datos de la tabla
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub fillCombo(pCombo As ComboBox, db As clsDB)

    fillComboCollection pCombo, collectionAll(db)

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Llena la ComboBox efectivamente
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub fillComboCollection(pCombo As ComboBox, pCollection As Collection)
Dim elemento As clsMyAClienteVolumen

    pCombo.Clear

    For Each elemento In pCollection
        pCombo.AddItem elemento.comboText
        pCombo.ItemData(pCombo.NewIndex) = elemento.comboKey
    Next

    If pCombo.ListCount > 0 Then pCombo.ListIndex = 0

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Llena una ListBox con todos los datos de la tabla
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub fillList(pList As ListBox, db As clsDB)

    fillListCollection pList, collectionAll(db)

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Llena la ListBox efectivamente
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub fillListCollection(pList As ListBox, pCollection As Collection)
Dim elemento As clsMyAClienteVolumen

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

