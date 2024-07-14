VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPMCCobDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vCodigoRegistro As Integer
Private vNroReferencia As String
Private vIdFactura As String
Private vFechaVencimiento As Date
Private vCodigoMoneda As Integer
Private vFechaAplicacion As Date
Private vImporte As Currency
Private vCodigoMovimiento As Integer
Private vFechaAcreditacion As Date
Private vCanalPago As String
Private vNroControl As String
Private vCodigoProvincia As String
Private vRegistro As String

Private Sub class_Initialize()

    initializeOther

End Sub

Private Sub initializeOther()

    vCodigoRegistro = 0
    vNroReferencia = ""
    vIdFactura = ""
    vFechaVencimiento = Date
    vCodigoMoneda = 0
    vFechaAplicacion = Date
    vImporte = 0
    vCodigoMovimiento = 0
    vFechaAcreditacion = Date
    vCanalPago = ""
    vNroControl = ""
    vCodigoProvincia = ""
    vRegistro = ""

End Sub

Public Sub clean()

    Call class_Initialize

End Sub

Public Property Get codigoRegistro() As Variant

    codigoRegistro = vCodigoRegistro

End Property

Public Property Let codigoRegistro(ByVal vNewValue As Variant)

On Error Resume Next

    vCodigoRegistro = vNewValue

End Property

Public Property Get nroReferencia() As Variant

    nroReferencia = vNroReferencia

End Property

Public Property Let nroReferencia(ByVal vNewValue As Variant)

On Error Resume Next

    vNroReferencia = vNewValue

End Property

Public Property Get idFactura() As Variant

    idFactura = vIdFactura

End Property

Public Property Let idFactura(ByVal vNewValue As Variant)

On Error Resume Next

    vIdFactura = vNewValue

End Property

Public Property Get fechaVencimiento() As Variant

    fechaVencimiento = vFechaVencimiento

End Property

Public Property Let fechaVencimiento(ByVal vNewValue As Variant)

On Error Resume Next

    vFechaVencimiento = vNewValue

End Property

Public Property Get codigoMoneda() As Variant

    codigoMoneda = vCodigoMoneda

End Property

Public Property Let codigoMoneda(ByVal vNewValue As Variant)

On Error Resume Next

    vCodigoMoneda = vNewValue

End Property

Public Property Get fechaAplicacion() As Variant

    fechaAplicacion = vFechaAplicacion

End Property

Public Property Let fechaAplicacion(ByVal vNewValue As Variant)

On Error Resume Next

    vFechaAplicacion = vNewValue

End Property

Public Property Get importe() As Variant

    importe = vImporte

End Property

Public Property Let importe(ByVal vNewValue As Variant)

On Error Resume Next

    vImporte = vNewValue

End Property

Public Property Get codigoMovimiento() As Variant

    codigoMovimiento = vCodigoMovimiento

End Property

Public Property Let codigoMovimiento(ByVal vNewValue As Variant)

On Error Resume Next

    vCodigoMovimiento = vNewValue

End Property

Public Property Get fechaAcreditacion() As Variant

    fechaAcreditacion = vFechaAcreditacion

End Property

Public Property Let fechaAcreditacion(ByVal vNewValue As Variant)

On Error Resume Next

    vFechaAcreditacion = vNewValue

End Property

Public Property Get canalPago() As Variant

    canalPago = vCanalPago

End Property

Public Property Let canalPago(ByVal vNewValue As Variant)

On Error Resume Next

    vCanalPago = vNewValue

End Property

Public Property Get nroControl() As Variant

    nroControl = vNroControl

End Property

Public Property Let nroControl(ByVal vNewValue As Variant)

On Error Resume Next

    vNroControl = vNewValue

End Property

Public Property Get codigoProvincia() As Variant

    codigoProvincia = vCodigoProvincia

End Property

Public Property Let codigoProvincia(ByVal vNewValue As Variant)

On Error Resume Next

    vCodigoProvincia = vNewValue

End Property

Public Function clone()
Dim objObjeto As New clsPMCCobDetalle

    With objObjeto
        .codigoRegistro = vCodigoRegistro
        .nroReferencia = vNroReferencia
        .idFactura = vIdFactura
        .fechaVencimiento = vFechaVencimiento
        .codigoMoneda = vCodigoMoneda
        .fechaAplicacion = vFechaAplicacion
        .importe = vImporte
        .codigoMovimiento = vCodigoMovimiento
        .fechaAcreditacion = vFechaAcreditacion
        .canalPago = vCanalPago
        .nroControl = vNroControl
        .codigoProvincia = vCodigoProvincia
    End With

    Set clone = objObjeto

End Function

Public Property Get registro() As Variant

    registro = vRegistro

End Property

Public Property Let registro(ByVal vNewValue As Variant)

    vRegistro = vNewValue
    
    parseRegistro

End Property

Private Sub parseRegistro()

    With Me
        .codigoRegistro = Mid(.registro, 1, 1)
        .nroReferencia = Mid(.registro, 2, 19)
        .idFactura = Mid(.registro, 21, 20)
        .fechaVencimiento = modConv.stringAAAAMMDD2Date(Mid(.registro, 41, 8))
        .codigoMoneda = Mid(.registro, 49, 1)
        .fechaAplicacion = modConv.stringAAAAMMDD2Date(Mid(.registro, 50, 8))
        .importe = modConv.parseDouble(Mid(.registro, 58, 11), 2)
        .codigoMovimiento = Mid(.registro, 69, 1)
        .fechaAcreditacion = modConv.stringAAAAMMDD2Date(Mid(.registro, 70, 8))
        .canalPago = Mid(.registro, 78, 2)
        .nroControl = Mid(.registro, 80, 4)
        .codigoProvincia = Mid(.registro, 84, 3)
    End With
    
End Sub

