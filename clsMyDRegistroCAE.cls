VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyDRegistroCAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "registrocae"

End Property

Public Property Get tipoID() As String

    tipoID = "tipo_id"

End Property

Public Property Get tipoID_t() As String

    tipoID_t = "registrocae.tipo_id"

End Property

Public Property Get prefijo() As String

    prefijo = "prefijo"

End Property

Public Property Get prefijo_t() As String

    prefijo_t = "registrocae.prefijo"

End Property

Public Property Get numero() As String

    numero = "numero"

End Property

Public Property Get numero_t() As String

    numero_t = "registrocae.numero"

End Property

Public Property Get clienteID() As String

    clienteID = "cliente_id"

End Property

Public Property Get clienteID_t() As String

    clienteID_t = "registrocae.cliente_id"

End Property

Public Property Get total() As String

    total = "total"

End Property

Public Property Get total_t() As String

    total_t = "registrocae.total"

End Property

Public Property Get exento() As String

    exento = "exento"

End Property

Public Property Get exento_t() As String

    exento_t = "registrocae.exento"

End Property

Public Property Get neto27() As String

    neto27 = "neto27"

End Property

Public Property Get neto27_t() As String

    neto27_t = "registrocae.neto27"

End Property

Public Property Get neto() As String

    neto = "neto"

End Property

Public Property Get neto_t() As String

    neto_t = "registrocae.neto"

End Property

Public Property Get neto105() As String

    neto105 = "neto105"

End Property

Public Property Get neto105_t() As String

    neto105_t = "registrocae.neto105"

End Property

Public Property Get iva27() As String

    iva27 = "iva27"

End Property

Public Property Get iva27_t() As String

    iva27_t = "registrocae.iva27"

End Property

Public Property Get iva() As String

    iva = "iva"

End Property

Public Property Get iva_t() As String

    iva_t = "registrocae.iva"

End Property

Public Property Get iva105() As String

    iva105 = "iva105"

End Property

Public Property Get iva105_t() As String

    iva105_t = "registrocae.iva105"

End Property

Public Property Get cae() As String

    cae = "cae"

End Property

Public Property Get cae_t() As String

    cae_t = "registrocae.cae"

End Property

Public Property Get fecha() As String

    fecha = "fecha"

End Property

Public Property Get fecha_t() As String

    fecha_t = "registrocae.fecha"

End Property

Public Property Get caeVencimiento() As String

    caeVencimiento = "cae_vencimiento"

End Property

Public Property Get caeVencimiento_t() As String

    caeVencimiento_t = "registrocae.cae_vencimiento"

End Property

Public Property Get barras() As String

    barras = "barras"

End Property

Public Property Get barras_t() As String

    barras_t = "registrocae.barras"

End Property

Public Property Get autoID() As String

    autoID = "auto_id"

End Property

Public Property Get autoID_t() As String

    autoID_t = "registrocae.auto_id"

End Property

Public Property Get uid() As String

    uid = "uid"

End Property

Public Property Get uid_t() As String

    uid_t = "registrocae.uid"

End Property

Public Property Get created() As String

    created = "created"

End Property

Public Property Get created_t() As String

    created_t = "registrocae.created"

End Property

