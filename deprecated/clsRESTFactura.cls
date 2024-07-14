VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRESTFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vPrefijoId As Integer
Private vFacturaId As Long
Private vFecha As Date
Private vClienteId As Long
Private vPeriodoId As Integer
Private vAnulada As Integer

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 20/10/2019
' Description: Inicializacion de la Clase
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub class_Initialize()

    vPrefijoId = 0
    vFacturaId = 0

    initializeOther

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 20/10/2019
' Description: Inicializacion de los campos que no corresponden a la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub initializeOther()

    vFecha = Date
    vClienteId = 0
    vPeriodoId = 0
    vAnulada = 0

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 20/10/2019
' Description: Limpia todos los atributos
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub clean()

    Call class_Initialize

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 20/10/2019
' Description: Getter prefijoId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get prefijoId() As Variant

    prefijoId = vPrefijoId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 20/10/2019
' Description: Setter prefijoId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let prefijoId(ByVal vNewValue As Variant)

On Error Resume Next

    vPrefijoId = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 20/10/2019
' Description: Getter facturaId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get facturaId() As Variant

    facturaId = vFacturaId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 20/10/2019
' Description: Setter facturaId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let facturaId(ByVal vNewValue As Variant)

On Error Resume Next

    vFacturaId = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 20/10/2019
' Description: Getter fecha
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fecha() As Variant

    fecha = vFecha

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 20/10/2019
' Description: Setter fecha
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let fecha(ByVal vNewValue As Variant)

On Error Resume Next

    vFecha = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 20/10/2019
' Description: Getter clienteId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get clienteId() As Variant

    clienteId = vClienteId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 20/10/2019
' Description: Setter clienteId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let clienteId(ByVal vNewValue As Variant)

On Error Resume Next

    vClienteId = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 20/10/2019
' Description: Getter periodoId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get periodoId() As Variant

    periodoId = vPeriodoId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 20/10/2019
' Description: Setter periodoId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let periodoId(ByVal vNewValue As Variant)

On Error Resume Next

    vPeriodoId = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 20/10/2019
' Description: Getter anulada
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get anulada() As Variant

    anulada = vAnulada

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 20/10/2019
' Description: Setter anulada
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let anulada(ByVal vNewValue As Variant)

On Error Resume Next

    vAnulada = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 20/10/2019
' Description: Atributo key para usar como clave de colecciones
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get key() As String

    key = "k." & prefijoId & "." & facturaId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 20/10/2019
' Description: Getter comboText
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comboText() As Variant

    comboText = vFecha

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 20/10/2019
' Description: Getter comboKey
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comboKey() As Variant

    comboKey = vFacturaId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 20/10/2019
' Description: Getter listText
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get listText() As Variant

    listText = vFecha

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 20/10/2019
' Description: Getter listKey
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get listKey() As Variant

    listKey = vFacturaId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 20/10/2019
' Description: Clona el objeto this y devuelve uno nuevo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function clone() As clsRESTFactura
Dim objeto As New clsRESTFactura

    objeto.prefijoId = vPrefijoId
    objeto.facturaId = vFacturaId
    objeto.fecha = vFecha
    objeto.clienteId = vClienteId
    objeto.periodoId = vPeriodoId
    objeto.anulada = vAnulada

    Set clone = objeto

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 20/10/2019
' Description: Carga los campos del registro en los atributos de la clase
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub fillObject(objeto As clsRESTFactura, json As Dictionary)

On Error Resume Next

    objeto.prefijoId = json.item("prefijoId")
    objeto.facturaId = json.item("facturaId")
    objeto.fecha = json.item("fecha")
    objeto.clienteId = json.item("clienteId")
    objeto.periodoId = json.item("periodoId")
    objeto.anulada = json.item("anulada")

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 20/10/2019
' Description: Consulta la DB con un SQL buscando 1 registro
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub findREST(json As Dictionary)
Dim rstQuery As ADODB.Recordset

    If json.item("status") = "" Then fillObject Me, json

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 20/10/2019
' Description: Realiza un find sobre los atributos de la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub findByPrimaryKey()
Dim http As MSXML2.XMLHTTP

Dim url As String

    initializeOther

    Set http = New MSXML2.ServerXMLHTTP

    url = "http://" & properties.item("server") & ":" & properties.item("port") & "/"

    http.Open "GET", url, False
    http.send

    findREST modJsonConverter.ParseJson(http.responseText)

    Set http = Nothing

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 20/10/2019
' Description: Consulta la DB con un SQL buscando el último registro según la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub findLast()

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 20/10/2019
' Description: Elimina 1 registro de la DB según la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function delete() As Boolean
Dim http As MSXML2.XMLHTTP

Dim url As String

    delete = False

    Set http = New MSXML2.ServerXMLHTTP

    url = "http://" & properties.item("server") & ":" & properties.item("port") & "/"

    http.Open "DELETE", url, False
    http.send

    If http.responseText = "" Then delete = True

    Set http = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 20/10/2019
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
' Created: 20/10/2019
' Description: Verifica si el registro asociado a la Primary Key existe
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function exist() As Boolean
Dim http As MSXML2.XMLHTTP

Dim url As String

Dim json As Dictionary

    exist = False

    Set http = New MSXML2.ServerXMLHTTP

    url = "http://" & properties.item("server") & ":" & properties.item("port") & "/"

    http.Open "GET", url, False
    http.send

    Set json = modJsonConverter.ParseJson(http.responseText)

    If json.item("status") = "" Then
        exist = True
        fillObject Me, json
    End If

    Set http = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 20/10/2019
' Description: Hace 1 INSERT con los valores de los atributos
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function add() As Boolean
Dim http As MSXML2.XMLHTTP

Dim url As String

Dim json As Dictionary

    add = False

    Set http = New MSXML2.ServerXMLHTTP

    url = "http://" & properties.item("server") & ":" & properties.item("port") & "/add?" & makeParams

    http.Open "GET", url, False
    http.send

    Set json = modJsonConverter.ParseJson(http.responseText)

    If json.item("status") = "" Then
        add = True
        fillObject Me, json
    End If

    Set http = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 20/10/2019
' Description: Hace 1 UPDATE sólo de los atributos que cambiaron respecto del almacenado en la DB
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function update() As Boolean
Dim http As MSXML2.XMLHTTP

Dim url As String

Dim json As Dictionary

    update = False

    Set http = New MSXML2.ServerXMLHTTP

    url = "http://" & properties.item("server") & ":" & properties.item("port") & "/update?" & makeParams

    http.Open "GET", url, False
    http.send

    Set json = modJsonConverter.ParseJson(http.responseText)

    If json.item("status") = "" Then
        update = True
        fillObject Me, json
    End If

    Set http = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 20/10/2019
' Description: Escribe makeParams
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function makeParams() As String
Dim params As String

    params = params & "&prefijoId=" & prefijoId
    params = params & "&facturaId=" & facturaId
    params = params & "&fecha=" & fecha
    params = params & "&clienteId=" & clienteId
    params = params & "&periodoId=" & periodoId
    params = params & "&anulada=" & anulada

    makeParams = params

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 20/10/2019
' Description: Arma una colección con todos los registros
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function collectionAll(Optional registrosold As Collection = Nothing) As Collection
Dim http As MSXML2.XMLHTTP

Dim url As String

    Set http = New MSXML2.ServerXMLHTTP

    url = "http://" & properties.item("server") & ":" & properties.item("port") & "/"

    http.Open "GET", url, False
    http.send

    Set collectionAll = collectionJSON(http.responseText, registrosold)

    Set http = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 20/10/2019
' Description: Arma una colección con todos los registros
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function collectionByPeriodoIdAndZona(periodoId As Integer, zona As Integer, Optional registrosold As Collection = Nothing) As Collection
Dim http As MSXML2.XMLHTTP

Dim url As String

    Set http = New MSXML2.ServerXMLHTTP

    url = "http://" & properties.item("server") & ":" & properties.item("port") & "/factura/zona/" & periodoId & "/" & zona

    http.Open "GET", url, False
    http.send

    Set collectionByPeriodoIdAndZona = collectionJSON(http.responseText, registrosold)

    Set http = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 20/10/2019
' Description: Hace la consulta SQL sobre la DB y arma la colección
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function collectionJSON(jsonText As String, Optional registrosold As Collection = Nothing) As Collection
Dim elemento As clsRESTFactura

Dim registros As Collection

Dim json As Dictionary

    If Left(jsonText, 1) <> "[" Or jsonText = "[]" Then
        Set collectionJSON = New Collection
        Exit Function
    End If

    If registrosold Is Nothing Then
        Set registros = New Collection
    Else
        Set registros = registrosold
    End If

    For Each json In modJsonConverter.ParseJson(jsonText)
        Set elemento = New clsRESTFactura
        fillObject elemento, json

        If Not modCollection.collectionExistElement(registros, elemento.key) Then registros.add elemento, elemento.key
    Next

    Set collectionJSON = registros

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 20/10/2019
' Description: Llena una ComboBox con todos los datos de la tabla
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub fillCombo(pCombo As ComboBox, Optional registrosold As Collection = Nothing)

    fillComboCollection pCombo, collectionAll(registrosold)

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 20/10/2019
' Description: Llena la ComboBox efectivamente
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub fillComboCollection(pCombo As ComboBox, pCollection As Collection)
Dim elemento As clsRESTFactura

    pCombo.Clear

    For Each elemento In pCollection
        pCombo.AddItem elemento.comboText
        pCombo.ItemData(pCombo.NewIndex) = elemento.comboKey
    Next

    If pCombo.ListCount > 0 Then pCombo.ListIndex = 0

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 20/10/2019
' Description: Llena una ListBox con todos los datos de la tabla
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub fillList(pList As ListBox, Optional registrosold As Collection = Nothing)

    fillListCollection pList, collectionAll(registrosold)

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 20/10/2019
' Description: Llena la ListBox efectivamente
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub fillListCollection(pList As ListBox, pCollection As Collection)
Dim elemento As clsRESTFactura

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

