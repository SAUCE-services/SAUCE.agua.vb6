VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyDDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "detalle"

End Property

Public Property Get puntoVta() As String

    puntoVta = "prefijo_id"

End Property

Public Property Get puntoVta_t() As String

    puntoVta_t = "detalle.prefijo_id"

End Property

Public Property Get nroComprob() As String

    nroComprob = "factura_id"

End Property

Public Property Get nroComprob_t() As String

    nroComprob_t = "detalle.factura_id"

End Property

Public Property Get rubroID() As String

    rubroID = "rubro_id"

End Property

Public Property Get rubroID_t() As String

    rubroID_t = "detalle.rubro_id"

End Property

Public Property Get concepto() As String

    concepto = "concepto"

End Property

Public Property Get concepto_t() As String

    concepto_t = "detalle.concepto"

End Property

Public Property Get cantidad() As String

    cantidad = "cantidad"

End Property

Public Property Get cantidad_t() As String

    cantidad_t = "detalle.cantidad"

End Property

Public Property Get precioUnitario() As String

    precioUnitario = "precio_unitario"

End Property

Public Property Get precioUnitario_t() As String

    precioUnitario_t = "detalle.precio_unitario"

End Property

Public Property Get iva() As String

    iva = "iva"

End Property

Public Property Get iva_t() As String

    iva_t = "detalle.iva"

End Property

Public Property Get autoID() As String

    autoID = "auto_id"

End Property

Public Property Get autoID_t() As String

    autoID_t = "detalle.auto_id"

End Property

Public Property Get uid() As String

    uid = "uid"

End Property

Public Property Get uid_t() As String

    uid_t = "detalle.uid"

End Property

Public Property Get created() As String

    created = "created"

End Property

Public Property Get created_t() As String

    created_t = "detalle.created"

End Property

