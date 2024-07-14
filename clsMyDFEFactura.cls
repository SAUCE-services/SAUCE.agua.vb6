VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyDFEFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "fefactura"

End Property

Public Property Get tipoID() As String

    tipoID = "tipo_id"

End Property

Public Property Get tipoID_t() As String

    tipoID_t = "fefactura.tipo_id"

End Property

Public Property Get prefijo() As String

    prefijo = "prefijo"

End Property

Public Property Get prefijo_t() As String

    prefijo_t = "fefactura.prefijo"

End Property

Public Property Get numero() As String

    numero = "numero"

End Property

Public Property Get numero_t() As String

    numero_t = "fefactura.numero"

End Property

Public Property Get clienteID() As String

    clienteID = "cliente_id"

End Property

Public Property Get clienteID_t() As String

    clienteID_t = "fefactura.cliente_id"

End Property

Public Property Get fecha() As String

    fecha = "fecha"

End Property

Public Property Get fecha_t() As String

    fecha_t = "fefactura.fecha"

End Property

Public Property Get importe() As String

    importe = "importe"

End Property

Public Property Get importe_t() As String

    importe_t = "fefactura.importe"

End Property

Public Property Get neto27() As String

    neto27 = "neto27"

End Property

Public Property Get neto27_t() As String

    neto27_t = "fefactura.neto27"

End Property

Public Property Get neto() As String

    neto = "neto"

End Property

Public Property Get neto_t() As String

    neto_t = "fefactura.neto"

End Property

Public Property Get iva27() As String

    iva27 = "iva27"

End Property

Public Property Get iva27_t() As String

    iva27_t = "fefactura.iva27"

End Property

Public Property Get iva() As String

    iva = "iva"

End Property

Public Property Get iva_t() As String

    iva_t = "fefactura.iva"

End Property

Public Property Get exento() As String

    exento = "exento"

End Property

Public Property Get exento_t() As String

    exento_t = "fefactura.exento"

End Property

Public Property Get recibo() As String

    recibo = "recibo"

End Property

Public Property Get recibo_t() As String

    recibo_t = "fefactura.recibo"

End Property

Public Property Get anulada() As String

    anulada = "anulada"

End Property

Public Property Get anulada_t() As String

    anulada_t = "fefactura.anulada"

End Property

Public Property Get tipoCompro() As String

    tipoCompro = "tipo_compro"

End Property

Public Property Get tipoCompro_t() As String

    tipoCompro_t = "fefactura.tipo_compro"

End Property

Public Property Get letras() As String

    letras = "letras"

End Property

Public Property Get letras_t() As String

    letras_t = "fefactura.letras"

End Property

Public Property Get observaciones() As String

    observaciones = "observaciones"

End Property

Public Property Get observaciones_t() As String

    observaciones_t = "fefactura.observaciones"

End Property

Public Property Get cae() As String

    cae = "cae"

End Property

Public Property Get cae_t() As String

    cae_t = "fefactura.cae"

End Property

Public Property Get caeVencimiento() As String

    caeVencimiento = "cae_vencimiento"

End Property

Public Property Get caeVencimiento_t() As String

    caeVencimiento_t = "fefactura.cae_vencimiento"

End Property

Public Property Get caeBarras() As String

    caeBarras = "cae_barras"

End Property

Public Property Get caeBarras_t() As String

    caeBarras_t = "fefactura.cae_barras"

End Property

Public Property Get puntoVta() As String

    puntoVta = "prefijo_id"

End Property

Public Property Get puntoVta_t() As String

    puntoVta_t = "fefactura.prefijo_id"

End Property

Public Property Get nroComprob() As String

    nroComprob = "factura_id"

End Property

Public Property Get nroComprob_t() As String

    nroComprob_t = "fefactura.factura_id"

End Property

Public Property Get autoID() As String

    autoID = "auto_id"

End Property

Public Property Get autoID_t() As String

    autoID_t = "fefactura.auto_id"

End Property

Public Property Get uid() As String

    uid = "uid"

End Property

Public Property Get uid_t() As String

    uid_t = "fefactura.uid"

End Property

Public Property Get created() As String

    created = "created"

End Property

Public Property Get created_t() As String

    created_t = "fefactura.created"

End Property

