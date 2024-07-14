VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyDRecibo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "recibo"

End Property

Public Property Get serieID() As String

    serieID = "serie_id"

End Property

Public Property Get serieID_t() As String

    serieID_t = "recibo.serie_id"

End Property

Public Property Get numero() As String

    numero = "numero"

End Property

Public Property Get numero_t() As String

    numero_t = "recibo.numero"

End Property

Public Property Get fecha() As String

    fecha = "fecha"

End Property

Public Property Get fecha_t() As String

    fecha_t = "recibo.fecha"

End Property

Public Property Get clienteID() As String

    clienteID = "cliente_id"

End Property

Public Property Get clienteID_t() As String

    clienteID_t = "recibo.cliente_id"

End Property

Public Property Get situacionIVA() As String

    situacionIVA = "situacion_iva"

End Property

Public Property Get situacionIva_t() As String

    situacionIva_t = "recibo.situacion_iva"

End Property

Public Property Get anulado() As String

    anulado = "anulado"

End Property

Public Property Get anulado_t() As String

    anulado_t = "recibo.anulado"

End Property

Public Property Get total() As String

    total = "total"

End Property

Public Property Get total_t() As String

    total_t = "recibo.total"

End Property

Public Property Get imputado() As String

    imputado = "imputado"

End Property

Public Property Get imputado_t() As String

    imputado_t = "recibo.imputado"

End Property

Public Property Get autoID() As String

    autoID = "auto_id"

End Property

Public Property Get autoID_t() As String

    autoID_t = "recibo.auto_id"

End Property

Public Property Get uid() As String

    uid = "uid"

End Property

Public Property Get uid_t() As String

    uid_t = "recibo.uid"

End Property

Public Property Get created() As String

    created = "created"

End Property

Public Property Get created_t() As String

    created_t = "recibo.created"

End Property

