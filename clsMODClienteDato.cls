VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMODClienteDato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vClienteId As Variant
Private vDocumento As Currency
Private vEmail As String
Private vFijo As String
Private vCelular As String

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/6/2022
' Description: Inicializacion de la Clase
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub class_Initialize()

    vClienteId = Null
    vDocumento = 0
    vEmail = ""
    vFijo = ""
    vCelular = ""

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/6/2022
' Description: Limpia todos los atributos
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub clean()

    Call class_Initialize

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/6/2022
' Description: Getter clienteId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get clienteId() As Variant

    clienteId = vClienteId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/6/2022
' Description: Setter clienteId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let clienteId(ByVal vNewValue As Variant)

On Error Resume Next

    vClienteId = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/6/2022
' Description: Getter documento
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get documento() As Variant

    documento = vDocumento

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/6/2022
' Description: Setter documento
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let documento(ByVal vNewValue As Variant)

On Error Resume Next

    vDocumento = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/6/2022
' Description: Getter email
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get email() As Variant

    email = vEmail

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/6/2022
' Description: Setter email
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let email(ByVal vNewValue As Variant)

On Error Resume Next

    vEmail = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/6/2022
' Description: Getter fijo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fijo() As Variant

    fijo = vFijo

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/6/2022
' Description: Setter fijo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let fijo(ByVal vNewValue As Variant)

On Error Resume Next

    vFijo = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/6/2022
' Description: Getter celular
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get celular() As Variant

    celular = vCelular

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/6/2022
' Description: Setter celular
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let celular(ByVal vNewValue As Variant)

On Error Resume Next

    vCelular = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/6/2022
' Description: Atributo key para usar como clave de colecciones
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get key() As String

    key = "k." & clienteId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/6/2022
' Description: Getter comboText
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comboText() As Variant

    comboText = vDocumento

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/6/2022
' Description: Getter comboKey
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comboKey() As Variant

    comboKey = vClienteId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/6/2022
' Description: Getter listText
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get listText() As Variant

    listText = vDocumento

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/6/2022
' Description: Getter listKey
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get listKey() As Variant

    listKey = vClienteId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/6/2022
' Description: Clona el objeto this y devuelve uno nuevo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function clone() As clsMODClienteDato
Dim object As New clsMODClienteDato

    object.clienteId = vClienteId
    object.documento = vDocumento
    object.email = vEmail
    object.fijo = vFijo
    object.celular = vCelular

    Set clone = object

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/6/2022
' Description: Escribe makeParams
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function makeParams() As Dictionary
Dim object As New Dictionary

    object.add "clienteId", clienteId
    object.add "documento", documento
    object.add "email", email
    object.add "fijo", fijo
    object.add "celular", celular

    Set makeParams = object

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 21/6/2022
' Description: Carga los campos del registro en los atributos de la clase
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub fillObject(json_object As Dictionary)

On Error Resume Next

    clienteId = json_object.item("clienteId")
    documento = json_object.item("documento")
    email = json_object.item("email")
    fijo = json_object.item("fijo")
    celular = json_object.item("celular")

End Sub

