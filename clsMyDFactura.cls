VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyDFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "factura"

End Property

Public Property Get puntoVta() As String

    puntoVta = "prefijo_id"

End Property

Public Property Get puntoVta_t() As String

    puntoVta_t = "factura.prefijo_id"

End Property

Public Property Get nroComprob() As String

    nroComprob = "factura_id"

End Property

Public Property Get nroComprob_t() As String

    nroComprob_t = "factura.factura_id"

End Property

Public Property Get fecha() As String

    fecha = "fecha"

End Property

Public Property Get fecha_t() As String

    fecha_t = "factura.fecha"

End Property

Public Property Get clienteID() As String

    clienteID = "cliente_id"

End Property

Public Property Get clienteID_t() As String

    clienteID_t = "factura.cliente_id"

End Property

Public Property Get periodoID() As String

    periodoID = "periodo_id"

End Property

Public Property Get periodoID_t() As String

    periodoID_t = "factura.periodo_id"

End Property

Public Property Get situacionIVA() As String

    situacionIVA = "situacion_iva"

End Property

Public Property Get situacionIva_t() As String

    situacionIva_t = "factura.situacion_iva"

End Property

Public Property Get tasa() As String

    tasa = "tasa"

End Property

Public Property Get tasa_t() As String

    tasa_t = "factura.tasa"

End Property

Public Property Get descuento() As String

    descuento = "descuento"

End Property

Public Property Get descuento_t() As String

    descuento_t = "factura.descuento"

End Property

Public Property Get pagada() As String

    pagada = "pagada"

End Property

Public Property Get pagada_t() As String

    pagada_t = "factura.pagada"

End Property

Public Property Get fechaPago() As String

    fechaPago = "fecha_pago"

End Property

Public Property Get fechaPago_t() As String

    fechaPago_t = "factura.fecha_pago"

End Property

Public Property Get tipoID() As String

    tipoID = "tipo_id"

End Property

Public Property Get tipoID_t() As String

    tipoID_t = "factura.tipo_id"

End Property

Public Property Get anulada() As String

    anulada = "anulada"

End Property

Public Property Get anulada_t() As String

    anulada_t = "factura.anulada"

End Property

Public Property Get total() As String

    total = "total"

End Property

Public Property Get total_t() As String

    total_t = "factura.total"

End Property

Public Property Get interes() As String

    interes = "interes"

End Property

Public Property Get interes_t() As String

    interes_t = "factura.interes"

End Property

Public Property Get letras() As String

    letras = "letras"

End Property

Public Property Get letras_t() As String

    letras_t = "factura.letras"

End Property

Public Property Get puntoVtaInteres() As String

    puntoVtaInteres = "prefijo_id_interes"

End Property

Public Property Get puntoVtaInteres_t() As String

    puntoVtaInteres_t = "factura.prefijo_id_interes"

End Property

Public Property Get nroComprobInteres() As String

    nroComprobInteres = "factura_id_interes"

End Property

Public Property Get nroComprobInteres_t() As String

    nroComprobInteres_t = "factura.factura_id_interes"

End Property

Public Property Get ivaCF() As String

    ivaCF = "iva_cf"

End Property

Public Property Get ivaCF_t() As String

    ivaCF_t = "factura.iva_cf"

End Property

Public Property Get ivaRI() As String

    ivaRI = "iva_ri"

End Property

Public Property Get ivaRI_t() As String

    ivaRI_t = "factura.iva_ri"

End Property

Public Property Get ivaRN() As String

    ivaRN = "iva_rn"

End Property

Public Property Get ivaRN_t() As String

    ivaRN_t = "factura.iva_rn"

End Property

Public Property Get periodoIDFin() As String

    periodoIDFin = "periodo_id_fin"

End Property

Public Property Get periodoIDFin_t() As String

    periodoIDFin_t = "factura.periodo_id_fin"

End Property

Public Property Get cancelada() As String

    cancelada = "cancelada"

End Property

Public Property Get cancelada_t() As String

    cancelada_t = "factura.cancelada"

End Property

Public Property Get planIDCancela() As String

    planIDCancela = "plan_id_cancela"

End Property

Public Property Get planIDCancela_t() As String

    planIDCancela_t = "factura.plan_id_cancela"

End Property

Public Property Get pfcodigo() As String

    pfcodigo = "pf_codigo"

End Property

Public Property Get pfcodigo_t() As String

    pfcodigo_t = "factura.pf_codigo"

End Property

Public Property Get pfbarras() As String

    pfbarras = "pf_barras"

End Property

Public Property Get pfbarras_t() As String

    pfbarras_t = "factura.pf_barras"

End Property

Public Property Get autoID() As String

    autoID = "auto_id"

End Property

Public Property Get autoID_t() As String

    autoID_t = "factura.auto_id"

End Property

Public Property Get uid() As String

    uid = "uid"

End Property

Public Property Get uid_t() As String

    uid_t = "factura.uid"

End Property

Public Property Get created() As String

    created = "created"

End Property

Public Property Get created_t() As String

    created_t = "factura.created"

End Property

