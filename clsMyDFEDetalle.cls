VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyDFEDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "fedetalle"

End Property

Public Property Get tipoID() As String

    tipoID = "tipo_id"

End Property

Public Property Get tipoID_t() As String

    tipoID_t = "fedetalle.tipo_id"

End Property

Public Property Get prefijo() As String

    prefijo = "prefijo"

End Property

Public Property Get prefijo_t() As String

    prefijo_t = "fedetalle.prefijo"

End Property

Public Property Get numero() As String

    numero = "numero"

End Property

Public Property Get numero_t() As String

    numero_t = "fedetalle.numero"

End Property

Public Property Get item() As String

    item = "item"

End Property

Public Property Get item_t() As String

    item_t = "fedetalle.item"

End Property

Public Property Get rubroID() As String

    rubroID = "rubro_id"

End Property

Public Property Get rubroID_t() As String

    rubroID_t = "fedetalle.rubro_id"

End Property

Public Property Get cantidad() As String

    cantidad = "cantidad"

End Property

Public Property Get cantidad_t() As String

    cantidad_t = "fedetalle.cantidad"

End Property

Public Property Get unitarioSinIva() As String

    unitarioSinIva = "unitario_sin_iva"

End Property

Public Property Get unitarioSinIva_t() As String

    unitarioSinIva_t = "fedetalle.unitario_sin_iva"

End Property

Public Property Get unitarioConIva() As String

    unitarioConIva = "unitario_con_iva"

End Property

Public Property Get unitarioConIva_t() As String

    unitarioConIva_t = "fedetalle.unitario_con_iva"

End Property

Public Property Get iva105() As String

    iva105 = "iva105"

End Property

Public Property Get iva105_t() As String

    iva105_t = "fedetalle.iva105"

End Property

Public Property Get exento() As String

    exento = "exento"

End Property

Public Property Get exento_t() As String

    exento_t = "fedetalle.exento"

End Property

Public Property Get fecha() As String

    fecha = "fecha"

End Property

Public Property Get fecha_t() As String

    fecha_t = "fedetalle.fecha"

End Property

Public Property Get concepto() As String

    concepto = "concepto"

End Property

Public Property Get concepto_t() As String

    concepto_t = "fedetalle.concepto"

End Property

Public Property Get autoID() As String

    autoID = "auto_id"

End Property

Public Property Get autoID_t() As String

    autoID_t = "fedetalle.auto_id"

End Property

Public Property Get uid() As String

    uid = "uid"

End Property

Public Property Get uid_t() As String

    uid_t = "fedetalle.uid"

End Property

Public Property Get created() As String

    created = "created"

End Property

Public Property Get created_t() As String

    created_t = "fedetalle.created"

End Property

