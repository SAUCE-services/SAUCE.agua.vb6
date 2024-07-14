VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMODFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vPrefijoId As Integer
Private vFacturaId As Long
Private vFecha As Variant
Private vClienteId As Long
Private vPeriodoId As Integer
Private vSituacionIva As Integer
Private vTasa As Currency
Private vDescuento As Currency
Private vPagada As Integer
Private vFechapago As Variant
Private vTipoId As Integer
Private vAnulada As Integer
Private vTotal As Currency
Private vInteres As Currency
Private vLetras As String
Private vPrefijoIdinteres As Integer
Private vFacturaIdinteres As Long
Private vIvacf As Currency
Private vIvari As Currency
Private vIvarn As Currency
Private vPeriodoIdfin As Integer
Private vCancelada As Integer
Private vPlanIdcancela As Variant
Private vPfcodigo As String
Private vPfbarras As String
Private vCajamovimientoId As Variant
Private vUniqueId As Variant
Private vUid As String

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Inicializacion de la Clase
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub class_Initialize()

    vPrefijoId = 0
    vFacturaId = 0
    vFecha = Null
    vClienteId = 0
    vPeriodoId = 0
    vSituacionIva = 0
    vTasa = 0
    vDescuento = 0
    vPagada = 0
    vFechapago = Null
    vTipoId = 0
    vAnulada = 0
    vTotal = 0
    vInteres = 0
    vLetras = ""
    vPrefijoIdinteres = 0
    vFacturaIdinteres = 0
    vIvacf = 0
    vIvari = 0
    vIvarn = 0
    vPeriodoIdfin = 0
    vCancelada = 0
    vPlanIdcancela = Null
    vPfcodigo = ""
    vPfbarras = ""
    vCajamovimientoId = Null
    vUid = ""
    vUniqueId = Null

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Limpia todos los atributos
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub clean()

    Call class_Initialize

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Getter prefijoId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get prefijoId() As Variant

    prefijoId = vPrefijoId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Setter prefijoId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let prefijoId(ByVal vNewValue As Variant)

On Error Resume Next

    vPrefijoId = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Getter facturaId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get facturaId() As Variant

    facturaId = vFacturaId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Setter facturaId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let facturaId(ByVal vNewValue As Variant)

On Error Resume Next

    vFacturaId = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Getter fecha
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fecha() As Variant

    fecha = vFecha

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Setter fecha
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let fecha(ByVal vNewValue As Variant)

On Error Resume Next

    vFecha = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Getter clienteId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get clienteId() As Variant

    clienteId = vClienteId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Setter clienteId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let clienteId(ByVal vNewValue As Variant)

On Error Resume Next

    vClienteId = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Getter periodoId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get periodoId() As Variant

    periodoId = vPeriodoId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Setter periodoId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let periodoId(ByVal vNewValue As Variant)

On Error Resume Next

    vPeriodoId = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Getter situacioniva
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get situacionIVA() As Variant

    situacionIVA = vSituacionIva

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Setter situacioniva
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let situacionIVA(ByVal vNewValue As Variant)

On Error Resume Next

    vSituacionIva = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Getter tasa
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get tasa() As Variant

    tasa = vTasa

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Setter tasa
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let tasa(ByVal vNewValue As Variant)

On Error Resume Next

    vTasa = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Getter descuento
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get descuento() As Variant

    descuento = vDescuento

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Setter descuento
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let descuento(ByVal vNewValue As Variant)

On Error Resume Next

    vDescuento = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Getter pagada
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get pagada() As Variant

    pagada = vPagada

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Setter pagada
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let pagada(ByVal vNewValue As Variant)

On Error Resume Next

    vPagada = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Getter fechapago
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fechapago() As Variant

    fechapago = vFechapago

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Setter fechapago
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let fechapago(ByVal vNewValue As Variant)

On Error Resume Next

    vFechapago = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Getter tipoId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get tipoId() As Variant

    tipoId = vTipoId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Setter tipoId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let tipoId(ByVal vNewValue As Variant)

On Error Resume Next

    vTipoId = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Getter anulada
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get anulada() As Variant

    anulada = vAnulada

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Setter anulada
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let anulada(ByVal vNewValue As Variant)

On Error Resume Next

    vAnulada = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Getter total
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get total() As Variant

    total = vTotal

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Setter total
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let total(ByVal vNewValue As Variant)

On Error Resume Next

    vTotal = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Getter interes
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get interes() As Variant

    interes = vInteres

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Setter interes
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let interes(ByVal vNewValue As Variant)

On Error Resume Next

    vInteres = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Getter letras
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get letras() As Variant

    letras = vLetras

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Setter letras
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let letras(ByVal vNewValue As Variant)

On Error Resume Next

    vLetras = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Getter prefijoIdinteres
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get prefijoIdinteres() As Variant

    prefijoIdinteres = vPrefijoIdinteres

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Setter prefijoIdinteres
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let prefijoIdinteres(ByVal vNewValue As Variant)

On Error Resume Next

    vPrefijoIdinteres = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Getter facturaIdinteres
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get facturaIdinteres() As Variant

    facturaIdinteres = vFacturaIdinteres

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Setter facturaIdinteres
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let facturaIdinteres(ByVal vNewValue As Variant)

On Error Resume Next

    vFacturaIdinteres = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Getter ivacf
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get ivacf() As Variant

    ivacf = vIvacf

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Setter ivacf
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let ivacf(ByVal vNewValue As Variant)

On Error Resume Next

    vIvacf = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Getter ivari
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get ivari() As Variant

    ivari = vIvari

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Setter ivari
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let ivari(ByVal vNewValue As Variant)

On Error Resume Next

    vIvari = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Getter ivarn
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get ivarn() As Variant

    ivarn = vIvarn

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Setter ivarn
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let ivarn(ByVal vNewValue As Variant)

On Error Resume Next

    vIvarn = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Getter periodoIdfin
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get periodoIdfin() As Variant

    periodoIdfin = vPeriodoIdfin

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Setter periodoIdfin
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let periodoIdfin(ByVal vNewValue As Variant)

On Error Resume Next

    vPeriodoIdfin = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Getter cancelada
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get cancelada() As Variant

    cancelada = vCancelada

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Setter cancelada
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let cancelada(ByVal vNewValue As Variant)

On Error Resume Next

    vCancelada = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Getter planIdcancela
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get planIdcancela() As Variant

    planIdcancela = vPlanIdcancela

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Setter planIdcancela
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let planIdcancela(ByVal vNewValue As Variant)

On Error Resume Next

    vPlanIdcancela = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Getter pfcodigo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get pfcodigo() As Variant

    pfcodigo = vPfcodigo

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Setter pfcodigo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let pfcodigo(ByVal vNewValue As Variant)

On Error Resume Next

    vPfcodigo = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Getter pfbarras
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get pfbarras() As Variant

    pfbarras = vPfbarras

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Setter pfbarras
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let pfbarras(ByVal vNewValue As Variant)

On Error Resume Next

    vPfbarras = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Getter cajamovimientoId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get cajamovimientoId() As Variant

    cajamovimientoId = vCajamovimientoId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Setter cajamovimientoId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let cajamovimientoId(ByVal vNewValue As Variant)

On Error Resume Next

    vCajamovimientoId = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Getter uId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get uid() As Variant

    uid = vUid

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Setter uId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let uid(ByVal vNewValue As Variant)

On Error Resume Next

    vUid = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Getter uniqueId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get uniqueId() As Variant

    uniqueId = vUniqueId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Setter uniqueId
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let uniqueId(ByVal vNewValue As Variant)

On Error Resume Next

    vUniqueId = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Atributo key para usar como clave de colecciones
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get key() As String

    key = "k." & prefijoId & "." & facturaId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Getter comboText
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comboText() As Variant

    comboText = vFecha

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Getter comboKey
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comboKey() As Variant

    comboKey = vFacturaId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Getter listText
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get listText() As Variant

    listText = vFecha

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Getter listKey
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get listKey() As Variant

    listKey = vFacturaId

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Clona el objeto this y devuelve uno nuevo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function clone() As clsMODFactura
Dim object As New clsMODFactura

    object.prefijoId = vPrefijoId
    object.facturaId = vFacturaId
    object.fecha = vFecha
    object.clienteId = vClienteId
    object.periodoId = vPeriodoId
    object.situacionIVA = vSituacionIva
    object.tasa = vTasa
    object.descuento = vDescuento
    object.pagada = vPagada
    object.fechapago = vFechapago
    object.tipoId = vTipoId
    object.anulada = vAnulada
    object.total = vTotal
    object.interes = vInteres
    object.letras = vLetras
    object.prefijoIdinteres = vPrefijoIdinteres
    object.facturaIdinteres = vFacturaIdinteres
    object.ivacf = vIvacf
    object.ivari = vIvari
    object.ivarn = vIvarn
    object.periodoIdfin = vPeriodoIdfin
    object.cancelada = vCancelada
    object.planIdcancela = vPlanIdcancela
    object.pfcodigo = vPfcodigo
    object.pfbarras = vPfbarras
    object.cajamovimientoId = vCajamovimientoId
    object.uid = vUid
    object.uniqueId = vUniqueId

    Set clone = object

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Escribe makeParams
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function makeParams(Optional update As Boolean = False) As Dictionary
Dim object As New Dictionary

    object.add "prefijoId", prefijoId
    object.add "facturaId", facturaId
    object.add "fecha", modConv.date2JSON(fecha)
    object.add "clienteId", clienteId
    object.add "periodoId", periodoId
    object.add "situacionIva", situacionIVA
    object.add "tasa", tasa
    object.add "descuento", descuento
    object.add "pagada", pagada
    object.add "fechaPago", modConv.date2JSON(fechapago)
    object.add "tipoId", tipoId
    object.add "anulada", anulada
    object.add "total", total
    object.add "interes", interes
    object.add "letras", letras
    object.add "prefijoIdInteres", prefijoIdinteres
    object.add "facturaIdInteres", facturaIdinteres
    object.add "ivaCf", ivacf
    object.add "ivaRi", ivari
    object.add "ivan", ivarn
    object.add "periodoIdFin", periodoIdfin
    object.add "cancelada", cancelada
    object.add "planIdCancela", planIdcancela
    object.add "pfCodigo", pfcodigo
    object.add "pfBarras", pfbarras
    object.add "cajamovimientoId", cajamovimientoId
    object.add "uniqueId", uniqueId

    Set makeParams = object

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 10/7/2020
' Description: Carga los campos del registro en los atributos de la clase
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub fillObject(json_object As Dictionary)

On Error Resume Next

    prefijoId = json_object.item("prefijoId")
    facturaId = json_object.item("facturaId")
    fecha = modConv.JSON2Date(json_object.item("fecha"))
    clienteId = json_object.item("clienteId")
    periodoId = json_object.item("periodoId")
    situacionIVA = json_object.item("situacionIva")
    tasa = json_object.item("tasa")
    descuento = json_object.item("descuento")
    pagada = json_object.item("pagada")
    fechapago = modConv.JSON2Date(json_object.item("fechaPago"))
    tipoId = json_object.item("tipoId")
    anulada = json_object.item("anulada")
    total = json_object.item("total")
    interes = json_object.item("interes")
    letras = json_object.item("letras")
    prefijoIdinteres = json_object.item("prefijoIdInteres")
    facturaIdinteres = json_object.item("facturaIdInteres")
    ivacf = json_object.item("ivaCf")
    ivari = json_object.item("ivaRi")
    ivarn = json_object.item("ivaRn")
    periodoIdfin = json_object.item("periodoIdFin")
    cancelada = json_object.item("cancelada")
    planIdcancela = json_object.item("planIdCancela")
    pfcodigo = json_object.item("pfCodigo")
    pfbarras = json_object.item("pfBarras")
    cajamovimientoId = json_object.item("cajamovimientoId")
    uniqueId = json_object.item("uniqueId")

End Sub

