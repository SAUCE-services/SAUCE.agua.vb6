VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsREPFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Consulta la DB con un SQL buscando 1 registro
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Function findREST(json_object As Dictionary) As clsMODFactura
Dim factura As New clsMODFactura

    If json_object.item("status") <> "" Then
        Set findREST = New clsMODFactura
        Exit Function
    End If

    factura.fillObject json_object

    Set findREST = factura

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Realiza un find sobre los atributos de la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function findByPrimaryKey(prefijoId As Integer, facturaId As Long) As clsMODFactura
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "factura/" & prefijoId & "/" & facturaId

    request.Open "GET", url
    request.send

    Set findByPrimaryKey = findREST(modJsonConverter.ParseJson(request.responseText))

    Set request = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Elimina 1 registro de la DB según la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function delete(facturaId As Long) As Boolean
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    delete = False

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "factura/" & facturaId

    request.Open "DELETE", url
    request.send

    If request.responseText = "" Then delete = True

    Set request = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Graba un registro verificando si debe agregarlo o actualizarlo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function save(factura As clsMODFactura) As Boolean

    If exist(factura.facturaId) Then
        save = update(factura, factura.facturaId)
    Else
        save = add(factura)
    End If

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Verifica si el registro asociado a la Primary Key existe
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function exist(facturaId As Long) As Boolean
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

Dim response_object As Dictionary

    exist = False

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "factura/" & facturaId

    request.Open "GET", url
    request.send

    Set response_object = modJsonConverter.ParseJson(request.responseText)

    If response_object.item("status") = "" Then exist = True

    Set request = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Hace 1 INSERT con los valores de los atributos
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function add(factura As clsMODFactura) As clsMODFactura
Dim request As MSXML2.ServerXMLHTTP

Dim url As String
Dim json_text As String

Dim response_object As Dictionary

Dim object As New clsMODFactura

    Set add = New clsMODFactura

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "factura/"

    request.Open "POST", url
    request.setRequestHeader "Content-Type", "application/json"

    json_text = modJsonConverter.ConvertToJson(factura.makeParams)

    request.send json_text

    Set response_object = modJsonConverter.ParseJson(request.responseText)

    object.fillObject response_object

    If response_object.item("status") = "" Then Set add = object

    Set request = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Hace 1 UPDATE sólo de los atributos que cambiaron respecto del almacenado en la DB
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function update(factura As clsMODFactura, facturaId As Long) As clsMODFactura
Dim request As MSXML2.ServerXMLHTTP

Dim url As String
Dim json_text As String

Dim response_object As Dictionary

Dim object As New clsMODFactura

    Set update = factura

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "factura/" & facturaId

    request.Open "PUT", url
    request.setRequestHeader "Content-Type", "application/json"

    json_text = modJsonConverter.ConvertToJson(factura.makeParams(True))

    request.send json_text

    Set response_object = modJsonConverter.ParseJson(request.responseText)

    object.fillObject response_object

    If response_object.item("status") = "" Then Set update = object

    Set request = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Arma una colección con todos los registros
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function collectionAll(Optional registrosold As Collection = Nothing) As Collection
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "factura/"

    request.Open "GET", url
    request.send

    Set collectionAll = collectionJSON(request.responseText, registrosold)

    Set request = Nothing

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

    url = modUrls.url_agua & "factura/zona/" & periodoId & "/" & zona

    http.Open "GET", url, False
    http.send

    Set collectionByPeriodoIdAndZona = collectionJSON(http.responseText, registrosold)

    Set http = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Hace la consulta SQL sobre la DB y arma la colección
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function collectionJSON(jsontext As String, Optional registrosold As Collection = Nothing) As Collection
Dim elemento As clsMODFactura

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
        Set elemento = New clsMODFactura
        elemento.fillObject json_object

        If Not modCollection.collectionExistElement(registros, elemento.key) Then registros.add elemento, elemento.key
    Next

    Set collectionJSON = registros

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Llena una ComboBox con todos los datos de la tabla
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub fillCombo(pCombo As ComboBox, Optional registrosold As Collection = Nothing)

    fillComboCollection pCombo, collectionAll(registrosold)

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Llena la ComboBox efectivamente
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub fillComboCollection(pCombo As ComboBox, pCollection As Collection)
Dim elemento As clsMODFactura

    pCombo.Clear

    For Each elemento In pCollection
        pCombo.AddItem elemento.comboText
        pCombo.ItemData(pCombo.NewIndex) = elemento.comboKey
    Next

    If pCombo.ListCount > 0 Then pCombo.ListIndex = 0

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Llena una ListBox con todos los datos de la tabla
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub fillList(pList As ListBox, Optional registrosold As Collection = Nothing)

    fillListCollection pList, collectionAll(registrosold)

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Llena la ListBox efectivamente
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub fillListCollection(pList As ListBox, pCollection As Collection)
Dim elemento As clsMODFactura

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

