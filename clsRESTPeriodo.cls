VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRESTPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vPeriodoId As Integer
Private vDescripcion As String
Private vFechaInicio As Variant
Private vFechafin As Variant
Private vFechaPrimero As Variant
Private vFechaSegundo As Variant
Private vTasa As Currency
Private vLeyenda As String
Private vLiquidado As Currency
Private vUid As String

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Inicializacion de la Clase
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub class_Initialize()

    vPeriodoId = 0

    initializeOther

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Inicializacion de los campos que no corresponden a la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub initializeOther()

    vDescripcion = ""
    vFechaInicio = Null
    vFechafin = Null
    vFechaPrimero = Null
    vFechaSegundo = Null
    vTasa = 0
    vLeyenda = ""
    vLiquidado = 0
    vUid = ""

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Limpia todos los atributos
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub clean()

    Call class_Initialize

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Getter periodoId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get periodoId() As Variant

    periodoId = vPeriodoId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Setter periodoId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let periodoId(ByVal vNewValue As Variant)

On Error Resume Next

    vPeriodoId = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Getter descripcion
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get descripcion() As Variant

    descripcion = vDescripcion

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Setter descripcion
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let descripcion(ByVal vNewValue As Variant)

On Error Resume Next

    vDescripcion = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Getter fechainicio
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fechaInicio() As Variant

    fechaInicio = vFechaInicio

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Setter fechainicio
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let fechaInicio(ByVal vNewValue As Variant)

On Error Resume Next

    vFechaInicio = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Getter fechafin
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fechaFin() As Variant

    fechaFin = vFechafin

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Setter fechafin
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let fechaFin(ByVal vNewValue As Variant)

On Error Resume Next

    vFechafin = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Getter fechaprimero
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fechaPrimero() As Variant

    fechaPrimero = vFechaPrimero

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Setter fechaprimero
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let fechaPrimero(ByVal vNewValue As Variant)

On Error Resume Next

    vFechaPrimero = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Getter fechasegundo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fechaSegundo() As Variant

    fechaSegundo = vFechaSegundo

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Setter fechasegundo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let fechaSegundo(ByVal vNewValue As Variant)

On Error Resume Next

    vFechaSegundo = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Getter tasa
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get tasa() As Variant

    tasa = vTasa

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Setter tasa
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let tasa(ByVal vNewValue As Variant)

On Error Resume Next

    vTasa = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Getter leyenda
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get leyenda() As Variant

    leyenda = vLeyenda

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Setter leyenda
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let leyenda(ByVal vNewValue As Variant)

On Error Resume Next

    vLeyenda = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Getter liquIdado
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get liquidado() As Variant

    liquidado = vLiquidado

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Setter liquIdado
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let liquidado(ByVal vNewValue As Variant)

On Error Resume Next

    vLiquidado = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Getter uId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get uid() As Variant

    uid = vUid

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Setter uId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let uid(ByVal vNewValue As Variant)

On Error Resume Next

    vUid = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Atributo key para usar como clave de colecciones
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get key() As String

    key = "k." & periodoId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Getter comboText
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comboText() As Variant

    comboText = vDescripcion

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Getter comboKey
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comboKey() As Variant

    comboKey = vPeriodoId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Getter listText
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get listText() As Variant

    listText = vDescripcion

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Getter listKey
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get listKey() As Variant

    listKey = vPeriodoId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Clona el objeto this y devuelve uno nuevo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function clone() As clsRESTPeriodo
Dim object As New clsRESTPeriodo

    object.periodoId = vPeriodoId
    object.descripcion = vDescripcion
    object.fechaInicio = vFechaInicio
    object.fechaFin = vFechafin
    object.fechaPrimero = vFechaPrimero
    object.fechaSegundo = vFechaSegundo
    object.tasa = vTasa
    object.leyenda = vLeyenda
    object.liquidado = vLiquidado
    object.uid = vUid

    Set clone = object

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Carga los campos del registro en los atributos de la clase
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub fillObject(object As clsRESTPeriodo, json_object As Dictionary)

On Error Resume Next

    object.periodoId = json_object.item("periodoId")
    object.descripcion = json_object.item("descripcion")
    object.fechaInicio = modConv.JSON2Date(json_object.item("fechaInicio"))
    object.fechaFin = modConv.JSON2Date(json_object.item("fechaFin"))
    object.fechaPrimero = modConv.JSON2Date(json_object.item("fechaPrimero"))
    object.fechaSegundo = modConv.JSON2Date(json_object.item("fechaSegundo"))
    object.tasa = json_object.item("tasa")
    object.leyenda = json_object.item("leyenda")
    object.liquidado = json_object.item("liquidado")
    object.uid = json_object.item("uid")

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Consulta la DB con un SQL buscando 1 registro
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub findREST(json_object As Dictionary)

    If json_object.item("status") = "" Then fillObject Me, json_object

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Realiza un find sobre los atributos de la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub findByPrimaryKey()
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    initializeOther

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "periodo/" & periodoId

    request.Open "GET", url
    request.send

    findREST modJsonConverter.ParseJson(request.responseText)

    Set request = Nothing

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Consulta la DB con un SQL buscando el último registro según la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub findLast()
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    initializeOther

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "periodo/last"

    request.Open "GET", url
    request.send

    findREST modJsonConverter.ParseJson(request.responseText)

    Set request = Nothing

End Sub

Public Sub findToday()
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    initializeOther

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "periodo/today"

    request.Open "GET", url
    request.send

    findREST modJsonConverter.ParseJson(request.responseText)

    Set request = Nothing

End Sub

Public Sub findByFecha(fecha As Date)
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    clean

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "periodo/byfecha/" & modConv.date2datetimeIso(fecha)

    request.Open "GET", url
    request.send

    findREST modJsonConverter.ParseJson(request.responseText)

    Set request = Nothing

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Elimina 1 registro de la DB según la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function delete() As Boolean
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    delete = False

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "periodo/" & periodoId

    request.Open "DELETE", url
    request.send

    If request.responseText = "" Then delete = True

    Set request = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
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
' Created: 25/5/2020
' Description: Verifica si el registro asociado a la Primary Key existe
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function exist() As Boolean
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

Dim response_object As Dictionary

    exist = False

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "periodo/" & periodoId

    request.Open "GET", url
    request.send

    Set response_object = modJsonConverter.ParseJson(request.responseText)

    If response_object.item("status") = "" Then exist = True

    Set request = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Hace 1 INSERT con los valores de los atributos
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function add() As Boolean
Dim request As MSXML2.ServerXMLHTTP

Dim url As String
Dim json_text As String

Dim response_object As Dictionary

    add = False

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "periodo/"

    request.Open "POST", url
    request.setRequestHeader "Content-Type", "application/json"

    json_text = modJsonConverter.ConvertToJson(makeParams)

    request.send json_text

    Set response_object = modJsonConverter.ParseJson(request.responseText)

    If response_object.item("status") = "" Then
        add = True
        fillObject Me, response_object
    End If

    Set request = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Hace 1 UPDATE sólo de los atributos que cambiaron respecto del almacenado en la DB
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function update() As Boolean
Dim request As MSXML2.ServerXMLHTTP

Dim url As String
Dim json_text As String

Dim response_object As Dictionary

    update = False

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "periodo/" & periodoId

    request.Open "PUT", url
    request.setRequestHeader "Content-Type", "application/json"

    json_text = modJsonConverter.ConvertToJson(makeParams)

    request.send json_text

    Set response_object = modJsonConverter.ParseJson(request.responseText)

    If response_object.item("status") = "" Then
        update = True
        fillObject Me, response_object
    End If

    Set request = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Escribe makeParams
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function makeParams(Optional update As Boolean = False) As Dictionary
Dim object As New Dictionary

    object.add "periodoId", periodoId
    object.add "descripcion", descripcion
    object.add "fechaInicio", modConv.date2JSON(fechaInicio, True)
    object.add "fechaFin", modConv.date2JSON(fechaFin, True)
    object.add "fechaPrimero", modConv.date2JSON(fechaPrimero, True)
    object.add "fechaSegundo", modConv.date2JSON(fechaSegundo, True)
    object.add "tasa", tasa
    object.add "leyenda", leyenda
    object.add "liquidado", liquidado
    object.add "uid", uid

    Set makeParams = object

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Arma una colección con todos los registros
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function collectionAll() As Collection
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "periodo/"

    request.Open "GET", url
    request.send

    Set collectionAll = collectionJSON(request.responseText)

    Set request = Nothing

End Function

Public Function collectionRecaudadoByPeriodo(desde As Date, hasta As Date) As Collection
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "periodo/recaudado/" & modConv.date2datetimeIso(desde) & "/" & modConv.date2datetimeIso(hasta)

    request.Open "GET", url
    request.send

    Set collectionRecaudadoByPeriodo = collectionJSON(request.responseText)

    Set request = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Hace la consulta SQL sobre la DB y arma la colección
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function collectionJSON(jsontext As String, Optional registrosold As Collection = Nothing) As Collection
Dim elemento As clsRESTPeriodo

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
        Set elemento = New clsRESTPeriodo
        fillObject elemento, json_object

        If Not modCollection.collectionExistElement(registros, elemento.key) Then registros.add elemento, elemento.key
    Next

    Set collectionJSON = registros

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Llena una ComboBox con todos los datos de la tabla
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub fillCombo(pCombo As ComboBox)

    fillComboCollection pCombo, collectionAll

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Llena la ComboBox efectivamente
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub fillComboCollection(pCombo As ComboBox, pCollection As Collection)
Dim elemento As clsRESTPeriodo

Dim periodo As New clsRESTPeriodo

    pCombo.Clear

    For Each elemento In pCollection
        pCombo.AddItem elemento.comboText
        pCombo.ItemData(pCombo.NewIndex) = elemento.comboKey
    Next

    If pCombo.ListCount = 0 Then
        pCombo.ListIndex = -1
        Exit Sub
    End If

    periodo.findToday

    If periodo.periodoId = 0 Then
        pCombo.ListIndex = 0
        Exit Sub
    End If
    
    pCombo.Text = periodo.comboText

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Llena una ListBox con todos los datos de la tabla
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub fillList(pList As ListBox, Optional registrosold As Collection = Nothing)

    fillListCollection pList, collectionAll(registrosold)

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 25/5/2020
' Description: Llena la ListBox efectivamente
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub fillListCollection(pList As ListBox, pCollection As Collection)
Dim elemento As clsRESTPeriodo

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

