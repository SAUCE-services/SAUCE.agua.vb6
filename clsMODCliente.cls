VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMODCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vClienteId As Long
Private vFechaAlta As Date
Private vFechaBaja As Variant
Private vApellido As String
Private vNombre As String
Private vNumeroSocio As Variant
Private vInmuebleCalle As String
Private vInmueblePuerta As String
Private vInmueblePiso As String
Private vInmuebleDpto As String
Private vInmuebleLocalidad As String
Private vInmuebleProvincia As String
Private vInmuebleCodpostal As Integer
Private vFiscalCalle As String
Private vFiscalPuerta As String
Private vFiscalPiso As String
Private vFiscalDpto As String
Private vFiscalLocalidad As String
Private vFiscalProvincia As String
Private vFiscalCodpostal As Integer
Private vCuit As String
Private vSituacionIva As Integer
Private vNombreCategoria As String
Private vCategoria As Integer
Private vServicio As Integer
Private vCobro As Integer
Private vZona As Integer
Private vRuta As Integer
Private vOrden As Integer
Private vCortado As Integer
Private vEstadoId As Integer
Private vFechaNacimiento As Variant
Private vCategoriaSocioId As Integer
Private vDestinoId As Integer
Private vUniqueId As Variant
Private vUid As String

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Inicializacion de la Clase
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub class_Initialize()

    vClienteId = 0
    vFechaAlta = Date
    vFechaBaja = Null
    vApellido = ""
    vNombre = ""
    vNumeroSocio = Null
    vInmuebleCalle = ""
    vInmueblePuerta = ""
    vInmueblePiso = ""
    vInmuebleDpto = ""
    vInmuebleLocalidad = ""
    vInmuebleProvincia = ""
    vInmuebleCodpostal = 0
    vFiscalCalle = ""
    vFiscalPuerta = ""
    vFiscalPiso = ""
    vFiscalDpto = ""
    vFiscalLocalidad = ""
    vFiscalProvincia = ""
    vFiscalCodpostal = 0
    vCuit = ""
    vSituacionIva = 0
    vNombreCategoria = ""
    vCategoria = 0
    vServicio = 0
    vCobro = 0
    vZona = 0
    vRuta = 0
    vOrden = 0
    vCortado = 0
    vEstadoId = 0
    vFechaNacimiento = Null
    vCategoriaSocioId = 0
    vDestinoId = 0
    vUid = ""
    vUniqueId = Null

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Limpia todos los atributos
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub clean()

    Call class_Initialize

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Getter clienteId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get clienteId() As Variant

    clienteId = vClienteId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Setter clienteId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let clienteId(ByVal vNewValue As Variant)

On Error Resume Next

    vClienteId = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Getter fechaalta
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fechaAlta() As Variant

    fechaAlta = vFechaAlta

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Setter fechaalta
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let fechaAlta(ByVal vNewValue As Variant)

On Error Resume Next

    vFechaAlta = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Getter fechabaja
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fechaBaja() As Variant

    fechaBaja = vFechaBaja

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Setter fechabaja
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let fechaBaja(ByVal vNewValue As Variant)

On Error Resume Next

    vFechaBaja = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Getter apellido
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get apellido() As Variant

    apellido = vApellido

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Setter apellido
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let apellido(ByVal vNewValue As Variant)

On Error Resume Next

    vApellido = vNewValue

End Property

Public Property Get apellidonombre() As String
    
    apellidonombre = textFound
    
End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Getter nombre
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get nombre() As Variant

    nombre = vNombre

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Setter nombre
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let nombre(ByVal vNewValue As Variant)

On Error Resume Next

    vNombre = vNewValue

End Property

Public Property Get textFound() As Variant

    textFound = IIf(cortado = 0, "", "CORTADO *** ") & apellido & ", " & nombre & " (" & clienteId & ")"
    
End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Getter numerosocio
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get numeroSocio() As Variant

    numeroSocio = vNumeroSocio

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Setter numerosocio
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let numeroSocio(ByVal vNewValue As Variant)

On Error Resume Next

    vNumeroSocio = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Getter inmueblecalle
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get inmuebleCalle() As Variant

    inmuebleCalle = vInmuebleCalle

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Setter inmueblecalle
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let inmuebleCalle(ByVal vNewValue As Variant)

On Error Resume Next

    vInmuebleCalle = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Getter inmueblepuerta
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get inmueblePuerta() As Variant

    inmueblePuerta = vInmueblePuerta

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Setter inmueblepuerta
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let inmueblePuerta(ByVal vNewValue As Variant)

On Error Resume Next

    vInmueblePuerta = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Getter inmueblepiso
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get inmueblePiso() As Variant

    inmueblePiso = vInmueblePiso

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Setter inmueblepiso
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let inmueblePiso(ByVal vNewValue As Variant)

On Error Resume Next

    vInmueblePiso = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Getter inmuebledpto
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get inmuebleDpto() As Variant

    inmuebleDpto = vInmuebleDpto

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Setter inmuebledpto
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let inmuebleDpto(ByVal vNewValue As Variant)

On Error Resume Next

    vInmuebleDpto = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Getter inmueblelocalidad
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get inmuebleLocalidad() As Variant

    inmuebleLocalidad = vInmuebleLocalidad

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Setter inmueblelocalidad
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let inmuebleLocalidad(ByVal vNewValue As Variant)

On Error Resume Next

    vInmuebleLocalidad = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Getter inmuebleprovincia
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get inmuebleProvincia() As Variant

    inmuebleProvincia = vInmuebleProvincia

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Setter inmuebleprovincia
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let inmuebleProvincia(ByVal vNewValue As Variant)

On Error Resume Next

    vInmuebleProvincia = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Getter inmueblecodpostal
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get inmuebleCodpostal() As Variant

    inmuebleCodpostal = vInmuebleCodpostal

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Setter inmueblecodpostal
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let inmuebleCodpostal(ByVal vNewValue As Variant)

On Error Resume Next

    vInmuebleCodpostal = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Getter fiscalcalle
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fiscalCalle() As Variant

    fiscalCalle = vFiscalCalle

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Setter fiscalcalle
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let fiscalCalle(ByVal vNewValue As Variant)

On Error Resume Next

    vFiscalCalle = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Getter fiscalpuerta
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fiscalPuerta() As Variant

    fiscalPuerta = vFiscalPuerta

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Setter fiscalpuerta
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let fiscalPuerta(ByVal vNewValue As Variant)

On Error Resume Next

    vFiscalPuerta = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Getter fiscalpiso
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fiscalPiso() As Variant

    fiscalPiso = vFiscalPiso

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Setter fiscalpiso
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let fiscalPiso(ByVal vNewValue As Variant)

On Error Resume Next

    vFiscalPiso = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Getter fiscaldpto
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fiscalDpto() As Variant

    fiscalDpto = vFiscalDpto

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Setter fiscaldpto
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let fiscalDpto(ByVal vNewValue As Variant)

On Error Resume Next

    vFiscalDpto = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Getter fiscallocalidad
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fiscalLocalidad() As Variant

    fiscalLocalidad = vFiscalLocalidad

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Setter fiscallocalidad
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let fiscalLocalidad(ByVal vNewValue As Variant)

On Error Resume Next

    vFiscalLocalidad = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Getter fiscalprovincia
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fiscalProvincia() As Variant

    fiscalProvincia = vFiscalProvincia

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Setter fiscalprovincia
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let fiscalProvincia(ByVal vNewValue As Variant)

On Error Resume Next

    vFiscalProvincia = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Getter fiscalcodpostal
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fiscalCodpostal() As Variant

    fiscalCodpostal = vFiscalCodpostal

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Setter fiscalcodpostal
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let fiscalCodpostal(ByVal vNewValue As Variant)

On Error Resume Next

    vFiscalCodpostal = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Getter cuit
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get cuit() As Variant

    cuit = vCuit

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Setter cuit
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let cuit(ByVal vNewValue As Variant)

On Error Resume Next

    vCuit = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Getter situacioniva
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get situacionIVA() As Variant

    situacionIVA = vSituacionIva

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Setter situacioniva
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let situacionIVA(ByVal vNewValue As Variant)

On Error Resume Next

    vSituacionIva = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Getter nombrecategoria
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get nombreCategoria() As Variant

    nombreCategoria = vNombreCategoria

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Setter nombrecategoria
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let nombreCategoria(ByVal vNewValue As Variant)

On Error Resume Next

    vNombreCategoria = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Getter categoria
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get categoria() As Variant

    categoria = vCategoria

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Setter categoria
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let categoria(ByVal vNewValue As Variant)

On Error Resume Next

    vCategoria = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Getter servicio
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get servicio() As Variant

    servicio = vServicio

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Setter servicio
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let servicio(ByVal vNewValue As Variant)

On Error Resume Next

    vServicio = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Getter cobro
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get cobro() As Variant

    cobro = vCobro

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Setter cobro
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let cobro(ByVal vNewValue As Variant)

On Error Resume Next

    vCobro = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Getter zona
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get zona() As Variant

    zona = vZona

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Setter zona
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let zona(ByVal vNewValue As Variant)

On Error Resume Next

    vZona = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Getter ruta
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get ruta() As Variant

    ruta = vRuta

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Setter ruta
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let ruta(ByVal vNewValue As Variant)

On Error Resume Next

    vRuta = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Getter orden
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get orden() As Variant

    orden = vOrden

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Setter orden
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let orden(ByVal vNewValue As Variant)

On Error Resume Next

    vOrden = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Getter cortado
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get cortado() As Variant

    cortado = vCortado

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Setter cortado
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let cortado(ByVal vNewValue As Variant)

On Error Resume Next

    vCortado = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Getter estadoId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get estadoID() As Variant

    estadoID = vEstadoId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Setter estadoId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let estadoID(ByVal vNewValue As Variant)

On Error Resume Next

    vEstadoId = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Getter fechanacimiento
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fechaNacimiento() As Variant

    fechaNacimiento = vFechaNacimiento

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Setter fechanacimiento
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let fechaNacimiento(ByVal vNewValue As Variant)

On Error Resume Next

    vFechaNacimiento = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Getter categoriasocioId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get categoriasocioID() As Variant

    categoriasocioID = vCategoriaSocioId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Setter categoriasocioId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let categoriasocioID(ByVal vNewValue As Variant)

On Error Resume Next

    vCategoriaSocioId = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Getter destinoId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get destinoID() As Variant

    destinoID = vDestinoId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Setter destinoId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let destinoID(ByVal vNewValue As Variant)

On Error Resume Next

    vDestinoId = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Getter uid
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get uid() As Variant

    uid = vUid

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Setter uid
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let uid(ByVal vNewValue As Variant)

On Error Resume Next

    vUid = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Getter uniqueId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get uniqueId() As Variant

    uniqueId = vUniqueId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Setter uniqueId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let uniqueId(ByVal vNewValue As Variant)

On Error Resume Next

    vUniqueId = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Atributo key para usar como clave de colecciones
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get key() As String

    key = "k." & clienteId

End Property

Public Property Get keyFound() As Variant
    
    keyFound = clienteId
    
End Property

Public Property Get comboText() As Variant

    comboText = vApellido & ", " & vNombre

End Property

Public Property Get domicilioText() As Variant

    domicilioText = vInmuebleCalle & " " & vInmueblePuerta & " " & vInmueblePiso & " " & vInmuebleDpto & " " & vInmuebleLocalidad

End Property

Public Property Get comboKey() As Variant

    comboKey = vClienteId

End Property

Public Property Get listText() As Variant

    listText = vApellido & ", " & vNombre

End Property

Public Property Get listKey() As Variant

    listKey = vClienteId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Clona el objeto this y devuelve uno nuevo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function clone() As clsMODCliente
Dim object As New clsMODCliente

    object.clienteId = vClienteId
    object.fechaAlta = vFechaAlta
    object.fechaBaja = vFechaBaja
    object.apellido = vApellido
    object.nombre = vNombre
    object.numeroSocio = vNumeroSocio
    object.inmuebleCalle = vInmuebleCalle
    object.inmueblePuerta = vInmueblePuerta
    object.inmueblePiso = vInmueblePiso
    object.inmuebleDpto = vInmuebleDpto
    object.inmuebleLocalidad = vInmuebleLocalidad
    object.inmuebleProvincia = vInmuebleProvincia
    object.inmuebleCodpostal = vInmuebleCodpostal
    object.fiscalCalle = vFiscalCalle
    object.fiscalPuerta = vFiscalPuerta
    object.fiscalPiso = vFiscalPiso
    object.fiscalDpto = vFiscalDpto
    object.fiscalLocalidad = vFiscalLocalidad
    object.fiscalProvincia = vFiscalProvincia
    object.fiscalCodpostal = vFiscalCodpostal
    object.cuit = vCuit
    object.situacionIVA = vSituacionIva
    object.nombreCategoria = vNombreCategoria
    object.categoria = vCategoria
    object.servicio = vServicio
    object.cobro = vCobro
    object.zona = vZona
    object.ruta = vRuta
    object.orden = vOrden
    object.cortado = vCortado
    object.estadoID = vEstadoId
    object.fechaNacimiento = vFechaNacimiento
    object.categoriasocioID = vCategoriaSocioId
    object.destinoID = vDestinoId
    object.uid = vUid
    object.uniqueId = vUniqueId

    Set clone = object

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Escribe makeParams
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function makeParams(Optional update As Boolean = False) As Dictionary
Dim object As New Dictionary

    If update Then object.add "uniqueId", uniqueId
    object.add "clienteId", clienteId
    object.add "fechaAlta", modConv.date2JSON(fechaAlta)
    object.add "fechaBaja", modConv.date2JSON(fechaBaja)
    object.add "apellido", apellido
    object.add "nombre", nombre
    object.add "numeroSocio", numeroSocio
    object.add "inmuebleCalle", inmuebleCalle
    object.add "inmueblePuerta", inmueblePuerta
    object.add "inmueblePiso", inmueblePiso
    object.add "inmuebleDpto", inmuebleDpto
    object.add "inmuebleLocalidad", inmuebleLocalidad
    object.add "inmuebleProvincia", inmuebleProvincia
    object.add "inmuebleCodpostal", inmuebleCodpostal
    object.add "fiscalCalle", fiscalCalle
    object.add "fiscalPuerta", fiscalPuerta
    object.add "fiscalPiso", fiscalPiso
    object.add "fiscalDpto", fiscalDpto
    object.add "fiscalLocalidad", fiscalLocalidad
    object.add "fiscalProvincia", fiscalProvincia
    object.add "fiscalCodpostal", fiscalCodpostal
    object.add "cuit", cuit
    object.add "situacionIva", situacionIVA
    object.add "nombreCategoria", nombreCategoria
    object.add "categoria", categoria
    object.add "servicio", servicio
    object.add "cobro", cobro
    object.add "zona", zona
    object.add "ruta", ruta
    object.add "orden", orden
    object.add "cortado", cortado
    object.add "estadoId", estadoID
    object.add "fechaNacimiento", modConv.date2JSON(fechaNacimiento)
    object.add "categoriasocioId", categoriasocioID
    object.add "destinoId", destinoID
    object.add "uid", uid

    Set makeParams = object

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/7/2020
' Description: Carga los campos del registro en los atributos de la clase
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub fillObject(json_object As Dictionary)

On Error Resume Next

    clienteId = json_object.item("clienteId")
    fechaAlta = modConv.JSON2Date(json_object.item("fechaAlta"))
    fechaBaja = modConv.JSON2Date(json_object.item("fechaBaja"))
    apellido = json_object.item("apellido")
    nombre = json_object.item("nombre")
    numeroSocio = json_object.item("numeroSocio")
    inmuebleCalle = json_object.item("inmuebleCalle")
    inmueblePuerta = json_object.item("inmueblePuerta")
    inmueblePiso = json_object.item("inmueblePiso")
    inmuebleDpto = json_object.item("inmuebleDpto")
    inmuebleLocalidad = json_object.item("inmuebleLocalidad")
    inmuebleProvincia = json_object.item("inmuebleProvincia")
    inmuebleCodpostal = json_object.item("inmuebleCodpostal")
    fiscalCalle = json_object.item("fiscalCalle")
    fiscalPuerta = json_object.item("fiscalPuerta")
    fiscalPiso = json_object.item("fiscalPiso")
    fiscalDpto = json_object.item("fiscalDpto")
    fiscalLocalidad = json_object.item("fiscalLocalidad")
    fiscalProvincia = json_object.item("fiscalProvincia")
    fiscalCodpostal = json_object.item("fiscalCodpostal")
    cuit = json_object.item("cuit")
    situacionIVA = json_object.item("situacionIva")
    nombreCategoria = json_object.item("nombreCategoria")
    categoria = json_object.item("categoria")
    servicio = json_object.item("servicio")
    cobro = json_object.item("cobro")
    zona = json_object.item("zona")
    ruta = json_object.item("ruta")
    orden = json_object.item("orden")
    cortado = json_object.item("cortado")
    estadoID = json_object.item("estadoId")
    fechaNacimiento = modConv.JSON2Date(json_object.item("fechaNacimiento"))
    categoriasocioID = json_object.item("categoriasocioId")
    destinoID = json_object.item("destinoId")
    uniqueId = json_object.item("uniqueId")
    uid = json_object.item("uid")

End Sub

