VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMODNotificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vClienteId As Long
Private vFecha As Date
Private vTiponotificacionId As Variant
Private vVencimiento As Variant
Private vEstado As Long
Private vNotificacionId As Variant

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/9/2020
' Description: Inicializacion de la Clase
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub class_Initialize()

    vNotificacionId = Null
    vClienteId = 0
    vFecha = Date
    vTiponotificacionId = Null
    vVencimiento = Null
    vEstado = 0

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/9/2020
' Description: Limpia todos los atributos
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub clean()

    Call class_Initialize

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/9/2020
' Description: Getter clienteId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get clienteId() As Variant

    clienteId = vClienteId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/9/2020
' Description: Setter clienteId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let clienteId(ByVal vNewValue As Variant)

On Error Resume Next

    vClienteId = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/9/2020
' Description: Getter fecha
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fecha() As Variant

    fecha = vFecha

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/9/2020
' Description: Setter fecha
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let fecha(ByVal vNewValue As Variant)

On Error Resume Next

    vFecha = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/9/2020
' Description: Getter tiponotificacionId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get tiponotificacionId() As Variant

    tiponotificacionId = vTiponotificacionId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/9/2020
' Description: Setter tiponotificacionId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let tiponotificacionId(ByVal vNewValue As Variant)

On Error Resume Next

    vTiponotificacionId = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/9/2020
' Description: Getter vencimiento
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get vencimiento() As Variant

    vencimiento = vVencimiento

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/9/2020
' Description: Setter vencimiento
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let vencimiento(ByVal vNewValue As Variant)

On Error Resume Next

    vVencimiento = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/9/2020
' Description: Getter estado
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get estado() As Variant

    estado = vEstado

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/9/2020
' Description: Setter estado
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let estado(ByVal vNewValue As Variant)

On Error Resume Next

    vEstado = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/9/2020
' Description: Getter notificacionId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get notificacionId() As Variant

    notificacionId = vNotificacionId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/9/2020
' Description: Setter notificacionId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let notificacionId(ByVal vNewValue As Variant)

On Error Resume Next

    vNotificacionId = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/9/2020
' Description: Atributo key para usar como clave de colecciones
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get key() As String

    key = "k." & clienteId & "." & fecha

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/9/2020
' Description: Getter comboText
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comboText() As Variant

    comboText = vClienteId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/9/2020
' Description: Getter comboKey
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comboKey() As Variant

    comboKey = vNotificacionId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/9/2020
' Description: Getter listText
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get listText() As Variant

    listText = vClienteId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/9/2020
' Description: Getter listKey
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get listKey() As Variant

    listKey = vNotificacionId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/9/2020
' Description: Clona el objeto this y devuelve uno nuevo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function clone() As clsMODNotificacion
Dim object As New clsMODNotificacion

    object.clienteId = vClienteId
    object.fecha = vFecha
    object.tiponotificacionId = vTiponotificacionId
    object.vencimiento = vVencimiento
    object.estado = vEstado
    object.notificacionId = vNotificacionId

    Set clone = object

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/9/2020
' Description: Escribe makeParams
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function makeParams(Optional update As Boolean = False) As Dictionary
Dim object As New Dictionary

    object.add "clienteId", clienteId
    object.add "fecha", modConv.date2JSON(fecha, True)
    object.add "tiponotificacionId", tiponotificacionId
    object.add "vencimiento", modConv.date2JSON(vencimiento, True)
    object.add "estado", estado
    If update Then object.add "notificacionId", notificacionId

    Set makeParams = object

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/9/2020
' Description: Carga los campos del registro en los atributos de la clase
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub fillObject(json_object As Dictionary)

On Error Resume Next

    clienteId = json_object.item("clienteId")
    fecha = modConv.JSON2Date(json_object.item("fecha"))
    tiponotificacionId = json_object.item("tiponotificacionId")
    vencimiento = modConv.JSON2Date(json_object.item("vencimiento"))
    estado = json_object.item("estado")
    notificacionId = json_object.item("notificacionId")

End Sub

