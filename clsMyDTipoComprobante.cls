VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyDTipoComprobante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "tipocomprobante"

End Property

Public Property Get tipoID() As String

    tipoID = "tipo_id"

End Property

Public Property Get tipoID_t() As String

    tipoID_t = "tipocomprobante.tipo_id"

End Property

Public Property Get descripcion() As String

    descripcion = "descripcion"

End Property

Public Property Get descripcion_t() As String

    descripcion_t = "tipocomprobante.descripcion"

End Property

Public Property Get modulo() As String

    modulo = "modulo"

End Property

Public Property Get modulo_t() As String

    modulo_t = "tipocomprobante.modulo"

End Property

Public Property Get aplicaPendiente() As String

    aplicaPendiente = "aplica_pendiente"

End Property

Public Property Get aplicaPendiente_t() As String

    aplicaPendiente_t = "tipocomprobante.aplica_pendiente"

End Property

Public Property Get cuentaCorriente() As String

    cuentaCorriente = "cuenta_corriente"

End Property

Public Property Get cuentaCorriente_t() As String

    cuentaCorriente_t = "tipocomprobante.cuenta_corriente"

End Property

Public Property Get debita() As String

    debita = "debita"

End Property

Public Property Get debita_t() As String

    debita_t = "tipocomprobante.debita"

End Property

Public Property Get iva() As String

    iva = "iva"

End Property

Public Property Get iva_t() As String

    iva_t = "tipocomprobante.iva"

End Property

Public Property Get aplicable() As String

    aplicable = "aplicable"

End Property

Public Property Get aplicable_t() As String

    aplicable_t = "tipocomprobante.aplicable"

End Property

Public Property Get libroIva() As String

    libroIva = "libro_iva"

End Property

Public Property Get libroIva_t() As String

    libroIva_t = "tipocomprobante.libro_iva"

End Property

Public Property Get tipoComprobante() As String

    tipoComprobante = "tipo_comprobante"

End Property

Public Property Get tipoComprobante_t() As String

    tipoComprobante_t = "tipocomprobante.tipo_comprobante"

End Property

Public Property Get recibo() As String

    recibo = "recibo"

End Property

Public Property Get recibo_t() As String

    recibo_t = "tipocomprobante.recibo"

End Property

Public Property Get contado() As String

    contado = "contado"

End Property

Public Property Get contado_t() As String

    contado_t = "tipocomprobante.contado"

End Property

Public Property Get puntoVenta() As String

    puntoVenta = "punto_venta"

End Property

Public Property Get puntoVenta_t() As String

    puntoVenta_t = "tipocomprobante.punto_venta"

End Property

Public Property Get comprobanteId() As String

    comprobanteId = "comprobante_id"

End Property

Public Property Get comprobanteId_t() As String

    comprobanteId_t = "tipocomprobante.comprobante_id"

End Property

Public Property Get facturaElectronica() As String

    facturaElectronica = "factura_electronica"

End Property

Public Property Get facturaElectronica_t() As String

    facturaElectronica_t = "tipocomprobante.factura_electronica"

End Property

Public Property Get uid() As String

    uid = "uid"

End Property

Public Property Get uid_t() As String

    uid_t = "tipocomprobante.uid"

End Property

Public Property Get created() As String

    created = "created"

End Property

Public Property Get created_t() As String

    created_t = "tipocomprobante.created"

End Property

