VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsREPCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Consulta la DB con un SQL buscando 1 registro
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Function findREST(json_object As Dictionary) As clsMODCliente
Dim cliente As New clsMODCliente

    If json_object.item("status") <> "" Then
        Set findREST = New clsMODCliente
        Exit Function
    End If

    cliente.fillObject json_object

    Set findREST = cliente

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Realiza un find sobre los atributos de la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function findLastByClienteId(clienteId As Long) As clsMODCliente
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "cliente/lastbyclienteId/" & clienteId

    request.Open "GET", url
    request.send

    Set findLastByClienteId = findREST(modJsonConverter.ParseJson(request.responseText))

    Set request = Nothing

End Function

Public Function findNextByClienteId(clienteId As Long) As clsMODCliente
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "cliente/nextbyclienteId/" & clienteId

    request.Open "GET", url
    request.send

    Set findNextByClienteId = findREST(modJsonConverter.ParseJson(request.responseText))

    Set request = Nothing

End Function

Public Function findSearch(clienteId As Long) As clsMODCliente

    Set findSearch = findLastByClienteId(clienteId)
    
End Function

Public Function findLastLast() As clsMODCliente
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "cliente/lastcliente"

    request.Open "GET", url
    request.send

    Set findLastLast = findREST(modJsonConverter.ParseJson(request.responseText))

    Set request = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Elimina 1 registro de la DB según la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function delete(clienteId As Long) As Boolean
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    delete = False

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "cliente/" & clienteId

    request.Open "DELETE", url
    request.send

    If request.responseText = "" Then delete = True

    Set request = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Graba un registro verificando si debe agregarlo o actualizarlo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function save(cliente As clsMODCliente) As clsMODCliente

    If IsNull(cliente.uniqueId) Then
        Set save = add(cliente)
    Else
        Set save = update(cliente, cliente.uniqueId)
    End If

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Verifica si el registro asociado a la Primary Key existe
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function exist(uniqueId As Long) As Boolean
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

Dim response_object As Dictionary

    exist = False

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "cliente/" & uniqueId

    request.Open "GET", url
    request.send

    Set response_object = modJsonConverter.ParseJson(request.responseText)

    If response_object.item("status") = "" Then exist = True

    Set request = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Hace 1 INSERT con los valores de los atributos
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function add(cliente As clsMODCliente) As clsMODCliente
Dim request As MSXML2.ServerXMLHTTP

Dim url As String
Dim json_text As String

Dim response_object As Dictionary

Dim object As New clsMODCliente

    Set add = New clsMODCliente

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "cliente/"

    request.Open "POST", url
    request.setRequestHeader "Content-Type", "application/json"

    json_text = modJsonConverter.ConvertToJson(cliente.makeParams)

    request.send json_text

    Set response_object = modJsonConverter.ParseJson(request.responseText)

    object.fillObject response_object

    If response_object.item("status") = "" Then Set add = object

    Set request = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Hace 1 UPDATE sólo de los atributos que cambiaron respecto del almacenado en la DB
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function update(cliente As clsMODCliente, uniqueId As Long) As clsMODCliente
Dim request As MSXML2.ServerXMLHTTP

Dim url As String
Dim json_text As String

Dim response_object As Dictionary

Dim object As New clsMODCliente

    Set update = cliente

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "cliente/" & uniqueId

    request.Open "PUT", url
    request.setRequestHeader "Content-Type", "application/json"

    json_text = modJsonConverter.ConvertToJson(cliente.makeParams(True))

    request.send json_text

    Set response_object = modJsonConverter.ParseJson(request.responseText)

    object.fillObject response_object

    If response_object.item("status") = "" Then Set update = object

    Set request = Nothing

End Function

Public Function collectionByClienteId(clienteId As Integer) As Collection
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "cliente/bycliente/" & clienteId

    request.Open "GET", url
    request.send

    Set collectionByClienteId = collectionJSON(request.responseText)

    Set request = Nothing

End Function

Public Function collectionActivos(Optional orderByName As Boolean = False) As Collection
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "cliente/activos/" & IIf(orderByName, "true", "false")

    request.Open "GET", url
    request.send

    Set collectionActivos = collectionJSON(request.responseText)

    Set request = Nothing

End Function

Public Function collectionRango(clienteIddesde As Long, clienteIdhasta As Long) As Collection
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "cliente/rango/" & clienteIddesde & "/" & clienteIdhasta

    request.Open "GET", url
    request.send

    Set collectionRango = collectionJSON(request.responseText)

    Set request = Nothing

End Function

Public Function collectionActivosMedibles() As Collection
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "cliente/activosmedibles"

    request.Open "GET", url
    request.send

    Set collectionActivosMedibles = collectionJSON(request.responseText)

    Set request = Nothing

End Function

Public Function collectionActivosConCuotaFija() As Collection
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "cliente/activosconcuotafija"

    request.Open "GET", url
    request.send

    Set collectionActivosConCuotaFija = collectionJSON(request.responseText)

    Set request = Nothing

End Function

Public Function collectionActivosConMedidor() As Collection
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "cliente/activosconmedidor"

    request.Open "GET", url
    request.send

    Set collectionActivosConMedidor = collectionJSON(request.responseText)

    Set request = Nothing

End Function

Public Function collectionActivosByZona() As Collection
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "cliente/activosbyzona"

    request.Open "GET", url
    request.send

    Set collectionActivosByZona = collectionJSON(request.responseText)

    Set request = Nothing

End Function

Public Function collectionActivosByRuta(zona As Integer) As Collection
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "cliente/activosbyruta/" & zona

    request.Open "GET", url
    request.send

    Set collectionActivosByRuta = collectionJSON(request.responseText)

    Set request = Nothing

End Function

Public Function collectionActivosByZonaRuta(zona As Integer, ruta As Integer) As Collection
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "cliente/activosbyzonaruta/" & zona & "/" & ruta

    request.Open "GET", url
    request.send

    Set collectionActivosByZonaRuta = collectionJSON(request.responseText)

    Set request = Nothing

End Function

Public Function collectionActivos2Lectura(zona As Integer, ruta As Integer) As Collection
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "cliente/activos2lectura/" & zona & "/" & ruta

    request.Open "GET", url
    request.send

    Set collectionActivos2Lectura = collectionJSON(request.responseText)

    Set request = Nothing

End Function

Public Function collectionZona(zona As Integer) As Collection
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "cliente/zona/" & zona

    request.Open "GET", url
    request.send

    Set collectionZona = collectionJSON(request.responseText)

    Set request = Nothing

End Function

Public Function collectionRuta(zona As Integer, ruta As Integer) As Collection
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "cliente/activosbyzonaruta/" & zona & "/" & ruta

    request.Open "GET", url
    request.send

    Set collectionRuta = collectionJSON(request.responseText)

    Set request = Nothing

End Function

Public Function collectionActivosByZonaRutaOtros(zona As Integer, ruta As Integer) As Collection
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "cliente/activosbyzonarutaotros/" & zona & "/" & ruta

    request.Open "GET", url
    request.send

    Set collectionActivosByZonaRutaOtros = collectionJSON(request.responseText)

    Set request = Nothing

End Function

Public Function collectionSociosActivos() As Collection
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "cliente/sociosactivos"

    request.Open "GET", url
    request.send

    Set collectionSociosActivos = collectionJSON(request.responseText)

    Set request = Nothing

End Function

Public Function collectionSociosActivosConMedidor() As Collection
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "cliente/sociosactivosconmedidor"

    request.Open "GET", url
    request.send

    Set collectionSociosActivosConMedidor = collectionJSON(request.responseText)

    Set request = Nothing

End Function

Public Function collectionSociosActivosConCuotaFija() As Collection
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "cliente/sociosactivosconcuotafija"

    request.Open "GET", url
    request.send

    Set collectionSociosActivosConCuotaFija = collectionJSON(request.responseText)

    Set request = Nothing

End Function

Public Function collectionDeudoresPlanCorte() As Collection
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "cliente/deudoresplancorte"

    request.Open "GET", url
    request.send

    Set collectionDeudoresPlanCorte = collectionJSON(request.responseText)

    Set request = Nothing

End Function

Public Function collectionDeudoresFactura60Dias() As Collection
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "cliente/deudoresfactura60dias"

    request.Open "GET", url
    request.send

    Set collectionDeudoresFactura60Dias = collectionJSON(request.responseText)

    Set request = Nothing

End Function

Public Function collectionSearch(chain As String) As Collection
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_agua & "cliente/search"

    request.Open "POST", url
    request.setRequestHeader "Content-Type", "application/json"
    
    request.send "[""" & Trim(chain) & """]"

    Set collectionSearch = collectionJSON(request.responseText)

    Set request = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Hace la consulta SQL sobre la DB y arma la colección
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function collectionJSON(jsontext As String, Optional registrosold As Collection = Nothing) As Collection
Dim elemento As clsMODCliente

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
        Set elemento = New clsMODCliente
        elemento.fillObject json_object

        If Not modCollection.collectionExistElement(registros, elemento.key) Then registros.add elemento, elemento.key
    Next

    Set collectionJSON = registros

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Llena una ComboBox con todos los datos de la tabla
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub fillCombo(pCombo As ComboBox, Optional pComboDomicilio As ComboBox = Nothing)

    fillComboCollection pCombo, collectionActivos, pComboDomicilio

End Sub

Public Sub fillComboOtros(pCombo As ComboBox, zona As Integer, ruta As Integer)

    fillComboCollection pCombo, collectionActivosByZonaRutaOtros(zona, ruta)

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Llena la ComboBox efectivamente
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub fillComboCollection(pCombo As ComboBox, pCollection As Collection, Optional pComboDomicilio As ComboBox = Nothing)
Dim elemento As clsMODCliente

    pCombo.Clear

    For Each elemento In pCollection
        pCombo.AddItem elemento.comboText
        pCombo.ItemData(pCombo.NewIndex) = elemento.comboKey
        If Not (pComboDomicilio Is Nothing) Then
            pComboDomicilio.AddItem elemento.domicilioText
            pComboDomicilio.ItemData(pComboDomicilio.NewIndex) = elemento.comboKey
        End If
    Next

    If pCombo.ListCount > 0 Then pCombo.ListIndex = 0

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Llena una ListBox con todos los datos de la tabla
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub fillList(pList As ListBox, Optional registrosold As Collection = Nothing)

    fillListCollection pList, collectionActivos

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Llena la ListBox efectivamente
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub fillListCollection(pList As ListBox, pCollection As Collection)
Dim elemento As clsMODCliente

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

Public Function formSearch(pForm As Form, pKeyAscii As Integer, pQueBusco As String) As clsMODCliente

    Load pForm
    
    pForm.repository = Me
    pForm.txtCadena.Text = Chr(pKeyAscii)
    pForm.Caption = "Búsqueda " & pQueBusco
    pForm.Show vbModal
    
    Set formSearch = pForm.model
    
    Unload pForm
    
End Function

Public Function sortCollection(clientes As Collection) As Collection
Dim sorted As Collection

Dim ciclo As Integer
Dim ciclo_secundario As Integer
Dim anterior As Integer

    If clientes Is Nothing Then
        Set sortCollection = Nothing
        Exit Function
    End If
    
    If clientes.Count > 0 Then
        Set sorted = New Collection
        For ciclo = 1 To clientes.Count
            If sorted.Count = 0 Then
                sorted.add clientes(ciclo)
            Else
                anterior = 0
                For ciclo_secundario = sorted.Count To 1 Step -1
                    If clientes(ciclo).clienteId < sorted(ciclo_secundario).clienteId Then
                        anterior = ciclo_secundario
                    Else
                        Exit For
                    End If
                Next
                If anterior = 0 Then
                    sorted.add clientes(ciclo)
                Else
                    sorted.add clientes(ciclo), , anterior
                End If
            End If
        Next
        Set sortCollection = sorted
        Exit Function
    End If
    
    Set sortCollection = clientes
    
End Function
