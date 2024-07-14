VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyDNCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "ncredito"

End Property

Public Property Get serieID() As String

    serieID = "serie_id"

End Property

Public Property Get serieID_t() As String

    serieID_t = "ncredito.serie_id"

End Property

Public Property Get numero() As String

    numero = "numero"

End Property

Public Property Get numero_t() As String

    numero_t = "ncredito.numero"

End Property

Public Property Get fecha() As String

    fecha = "fecha"

End Property

Public Property Get fecha_t() As String

    fecha_t = "ncredito.fecha"

End Property

Public Property Get clienteID() As String

    clienteID = "cliente_id"

End Property

Public Property Get clienteID_t() As String

    clienteID_t = "ncredito.cliente_id"

End Property

Public Property Get situacionIVA() As String

    situacionIVA = "situacion_iva"

End Property

Public Property Get situacionIva_t() As String

    situacionIva_t = "ncredito.situacion_iva"

End Property

Public Property Get anulado() As String

    anulado = "anulado"

End Property

Public Property Get anulado_t() As String

    anulado_t = "ncredito.anulado"

End Property

Public Property Get total() As String

    total = "total"

End Property

Public Property Get total_t() As String

    total_t = "ncredito.total"

End Property

Public Property Get puntoVta() As String

    puntoVta = "prefijo_id"

End Property

Public Property Get puntoVta_t() As String

    puntoVta_t = "ncredito.prefijo_id"

End Property

Public Property Get nroComprob() As String

    nroComprob = "factura_id"

End Property

Public Property Get nroComprob_t() As String

    nroComprob_t = "ncredito.factura_id"

End Property

Public Property Get ivaCF() As String

    ivaCF = "iva_cf"

End Property

Public Property Get ivaCF_t() As String

    ivaCF_t = "ncredito.iva_cf"

End Property

Public Property Get ivaRI() As String

    ivaRI = "iva_ri"

End Property

Public Property Get ivaRI_t() As String

    ivaRI_t = "ncredito.iva_ri"

End Property

Public Property Get ivaRN() As String

    ivaRN = "iva_rn"

End Property

Public Property Get ivaRN_t() As String

    ivaRN_t = "ncredito.iva_rn"

End Property

Public Property Get autoID() As String

    autoID = "auto_id"

End Property

Public Property Get autoID_t() As String

    autoID_t = "ncredito.auto_id"

End Property

Public Property Get uid() As String

    uid = "uid"

End Property

Public Property Get uid_t() As String

    uid_t = "ncredito.uid"

End Property

Public Property Get created() As String

    created = "created"

End Property

Public Property Get created_t() As String

    created_t = "ncredito.created"

End Property

