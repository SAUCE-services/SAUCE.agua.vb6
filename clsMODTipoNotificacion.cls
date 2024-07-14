VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMODTipoNotificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vTiponotificacionId As Integer
Private vNombre As String
Private vValorsocio As Currency
Private vValornosocio As Currency

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/9/2020
' Description: Inicializacion de la Clase
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub class_Initialize()

    vTiponotificacionId = 0
    vNombre = ""
    vValorsocio = 0
    vValornosocio = 0

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
' Description: Getter nombre
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get nombre() As Variant

    nombre = vNombre

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/9/2020
' Description: Setter nombre
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let nombre(ByVal vNewValue As Variant)

On Error Resume Next

    vNombre = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/9/2020
' Description: Getter valorsocio
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get valorsocio() As Variant

    valorsocio = vValorsocio

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/9/2020
' Description: Setter valorsocio
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let valorsocio(ByVal vNewValue As Variant)

On Error Resume Next

    vValorsocio = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/9/2020
' Description: Getter valornosocio
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get valornosocio() As Variant

    valornosocio = vValornosocio

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/9/2020
' Description: Setter valornosocio
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let valornosocio(ByVal vNewValue As Variant)

On Error Resume Next

    vValornosocio = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/9/2020
' Description: Atributo key para usar como clave de colecciones
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get key() As String

    key = "k." & tiponotificacionId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/9/2020
' Description: Getter comboText
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comboText() As Variant

    comboText = vNombre

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/9/2020
' Description: Getter comboKey
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comboKey() As Variant

    comboKey = vTiponotificacionId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/9/2020
' Description: Getter listText
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get listText() As Variant

    listText = vNombre

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/9/2020
' Description: Getter listKey
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get listKey() As Variant

    listKey = vTiponotificacionId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/9/2020
' Description: Clona el objeto this y devuelve uno nuevo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function clone() As clsMODTipoNotificacion
Dim object As New clsMODTipoNotificacion

    object.tiponotificacionId = vTiponotificacionId
    object.nombre = vNombre
    object.valorsocio = vValorsocio
    object.valornosocio = vValornosocio

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

    object.add "tiponotificacionId", tiponotificacionId
    object.add "nombre", nombre
    object.add "valorsocio", valorsocio
    object.add "valornosocio", valornosocio

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

    tiponotificacionId = json_object.item("tiponotificacionId")
    nombre = json_object.item("nombre")
    valorsocio = json_object.item("valorsocio")
    valornosocio = json_object.item("valornosocio")

End Sub

