VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRESTFacturaPmc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vClienteId As Long
Private vPrefijoId As Integer
Private vFacturaId As Long
Private vTotal As Currency
Private vPeriodoId As Integer
Private vPfcodigo As String
Private vDescripcion As String
Private vFechaPrimero As Date
Private vFechaSegundo As Date
Private vTasa As Currency
Private vUniqueId As Long

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Inicializacion de la Clase
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub class_Initialize()

    vUniqueId = 0

    initializeOther

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Inicializacion de los campos que no corresponden a la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub initializeOther()

    vClienteId = 0
    vPrefijoId = 0
    vFacturaId = 0
    vTotal = 0
    vPeriodoId = 0
    vPfcodigo = ""
    vDescripcion = ""
    vFechaPrimero = Date
    vFechaSegundo = Date
    vTasa = 0

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Limpia todos los atributos
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub clean()

    Call class_Initialize

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Getter clienteId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get clienteId() As Variant

    clienteId = vClienteId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Setter clienteId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let clienteId(ByVal vNewValue As Variant)

On Error Resume Next

    vClienteId = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Getter prefijoId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get prefijoId() As Variant

    prefijoId = vPrefijoId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Setter prefijoId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let prefijoId(ByVal vNewValue As Variant)

On Error Resume Next

    vPrefijoId = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Getter facturaId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get facturaId() As Variant

    facturaId = vFacturaId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Setter facturaId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let facturaId(ByVal vNewValue As Variant)

On Error Resume Next

    vFacturaId = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Getter total
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get total() As Variant

    total = vTotal

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Setter total
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let total(ByVal vNewValue As Variant)

On Error Resume Next

    vTotal = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Getter periodoId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get periodoId() As Variant

    periodoId = vPeriodoId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Setter periodoId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let periodoId(ByVal vNewValue As Variant)

On Error Resume Next

    vPeriodoId = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Getter pfcodigo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get pfcodigo() As Variant

    pfcodigo = vPfcodigo

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Setter pfcodigo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let pfcodigo(ByVal vNewValue As Variant)

On Error Resume Next

    vPfcodigo = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Getter descripcion
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get descripcion() As Variant

    descripcion = vDescripcion

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Setter descripcion
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let descripcion(ByVal vNewValue As Variant)

On Error Resume Next

    vDescripcion = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Getter fechaprimero
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fechaPrimero() As Variant

    fechaPrimero = vFechaPrimero

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Setter fechaprimero
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let fechaPrimero(ByVal vNewValue As Variant)

On Error Resume Next

    vFechaPrimero = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Getter fechasegundo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fechaSegundo() As Variant

    fechaSegundo = vFechaSegundo

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Setter fechasegundo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let fechaSegundo(ByVal vNewValue As Variant)

On Error Resume Next

    vFechaSegundo = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Getter tasa
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get tasa() As Variant

    tasa = vTasa

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Setter tasa
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let tasa(ByVal vNewValue As Variant)

On Error Resume Next

    vTasa = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Getter uniqueId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get uniqueId() As Variant

    uniqueId = vUniqueId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Setter uniqueId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let uniqueId(ByVal vNewValue As Variant)

On Error Resume Next

    vUniqueId = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Atributo key para usar como clave de colecciones
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get key() As String

    key = "k." & uniqueId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Getter comboText
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comboText() As Variant

    comboText = vClienteId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Getter comboKey
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comboKey() As Variant

    comboKey = vUniqueId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Getter listText
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get listText() As Variant

    listText = vClienteId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Getter listKey
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get listKey() As Variant

    listKey = vUniqueId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Clona el objeto this y devuelve uno nuevo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function clone() As clsRESTFacturaPmc
Dim object As New clsRESTFacturaPmc

    object.clienteId = vClienteId
    object.prefijoId = vPrefijoId
    object.facturaId = vFacturaId
    object.total = vTotal
    object.periodoId = vPeriodoId
    object.pfcodigo = vPfcodigo
    object.descripcion = vDescripcion
    object.fechaPrimero = vFechaPrimero
    object.fechaSegundo = vFechaSegundo
    object.tasa = vTasa
    object.uniqueId = vUniqueId

    Set clone = object

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Carga los campos del registro en los atributos de la clase
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub fillObject(object As clsRESTFacturaPmc, json_object As Dictionary)

On Error Resume Next

    object.clienteId = json_object.item("clienteId")
    object.prefijoId = json_object.item("prefijoId")
    object.facturaId = json_object.item("facturaId")
    object.total = json_object.item("total")
    object.periodoId = json_object.item("periodoId")
    object.pfcodigo = json_object.item("pfcodigo")
    object.descripcion = json_object.item("descripcion")
    object.fechaPrimero = modConv.JSON2Date(json_object.item("fechaprimero"))
    object.fechaSegundo = modConv.JSON2Date(json_object.item("fechasegundo"))
    object.tasa = json_object.item("tasa")
    object.uniqueId = json_object.item("uniqueId")

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Consulta la DB con un SQL buscando 1 registro
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub findREST(json_object As Dictionary)

    If json_object.item("status") = "" Then fillObject Me, json_object

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Realiza un find sobre los atributos de la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub findByPrimaryKey()
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    initializeOther

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "facturapmc/"

    request.Open "GET", url
    request.send

    findREST modJsonConverter.ParseJson(request.responseText)

    Set request = Nothing

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Consulta la DB con un SQL buscando el último registro según la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub findLast()
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    initializeOther

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "facturapmc/last"

    request.Open "GET", url
    request.send

    findREST modJsonConverter.ParseJson(request.responseText)

    Set request = Nothing

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Elimina 1 registro de la DB según la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function delete() As Boolean
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    delete = False

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "facturapmc/"

    request.Open "DELETE", url
    request.send

    If request.responseText = "" Then delete = True

    Set request = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Graba un registro verificando si debe agregarlo o actualizarlo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function save() As Boolean

    If exist() Then
        save = update()
    Else
        save = add()
    End If

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Verifica si el registro asociado a la Primary Key existe
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function exist() As Boolean
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

Dim response_object As Dictionary

    exist = False

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "facturapmc/"

    request.Open "GET", url
    request.send

    Set response_object = modJsonConverter.ParseJson(request.responseText)

    If response_object.item("status") = "" Then exist = True

    Set request = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Hace 1 INSERT con los valores de los atributos
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function add() As Boolean
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

Dim response_object As Dictionary

    add = False

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "facturapmc/"

    request.Open "POST", url
    request.setRequestHeader "Content-Type", "application/json"

    request.send modJsonConverter.ConvertToJson(makeParams)

    Set response_object = modJsonConverter.ParseJson(request.responseText)

    If response_object.item("status") = "" Then
        add = True
        fillObject Me, response_object
    End If

    Set request = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Hace 1 UPDATE sólo de los atributos que cambiaron respecto del almacenado en la DB
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function update() As Boolean
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

Dim response_object As Dictionary

    update = False

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "facturapmc/"

    request.Open "PUT", url
    request.setRequestHeader "Content-Type", "application/json"

    request.send modJsonConverter.ConvertToJson(makeParams)

    Set response_object = modJsonConverter.ParseJson(request.responseText)

    If response_object.item("status") = "" Then
        update = True
        fillObject Me, response_object
    End If

    Set request = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Escribe makeParams
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function makeParams(Optional update As Boolean = False) As Dictionary
Dim object As New Dictionary

    object.add "clienteId", clienteId
    object.add "prefijoId", prefijoId
    object.add "facturaId", facturaId
    object.add "total", total
    object.add "periodoId", periodoId
    object.add "pfcodigo", pfcodigo
    object.add "descripcion", descripcion
    object.add "fechaprimero", fechaPrimero
    object.add "fechasegundo", fechaSegundo
    object.add "tasa", tasa
    object.add "uniqueId", uniqueId

    Set makeParams = object

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Arma una colección con todos los registros
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function collectionAll(Optional registrosold As Collection = Nothing) As Collection
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "facturapmc/"

    request.Open "GET", url
    request.send

    Set collectionAll = collectionJSON(request.responseText, registrosold)

    Set request = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Arma una colección con todos los registros
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function collectionDeudaByPeriodo(desde As Date, hasta As Date) As Collection
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "facturapmc/deuda/" & modConv.date2datetimeIso(desde) & "/" & modConv.date2datetimeIso(hasta)

    request.Open "GET", url
    request.send

    Set collectionDeudaByPeriodo = collectionJSON(request.responseText)

    Set request = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Hace la consulta SQL sobre la DB y arma la colección
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function collectionJSON(jsontext As String, Optional registrosold As Collection = Nothing) As Collection
Dim elemento As clsRESTFacturaPmc

Dim registros As Collection

Dim json_object As Dictionary

    If Left(jsontext, 1) <> "[" Or jsontext = "[]" Then
        Set collectionJSON = New Collection
        Exit Function
    End If

    If registrosold Is Nothing Then
        Set registros = New Collection
    Else
        Set registros = registrosold
    End If

    For Each json_object In modJsonConverter.ParseJson(jsontext)
        Set elemento = New clsRESTFacturaPmc
        fillObject elemento, json_object

        If Not modCollection.collectionExistElement(registros, elemento.key) Then registros.add elemento, elemento.key
    Next

    Set collectionJSON = registros

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Llena una ComboBox con todos los datos de la tabla
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub fillCombo(pCombo As ComboBox, Optional registrosold As Collection = Nothing)

    fillComboCollection pCombo, collectionAll(registrosold)

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Llena la ComboBox efectivamente
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub fillComboCollection(pCombo As ComboBox, pCollection As Collection)
Dim elemento As clsRESTFacturaPmc

    pCombo.Clear

    For Each elemento In pCollection
        pCombo.AddItem elemento.comboText
        pCombo.ItemData(pCombo.NewIndex) = elemento.comboKey
    Next

    If pCombo.ListCount > 0 Then pCombo.ListIndex = 0

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Llena una ListBox con todos los datos de la tabla
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub fillList(pList As ListBox, Optional registrosold As Collection = Nothing)

    fillListCollection pList, collectionAll(registrosold)

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 3/5/2020
' Description: Llena la ListBox efectivamente
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub fillListCollection(pList As ListBox, pCollection As Collection)
Dim elemento As clsRESTFacturaPmc

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

