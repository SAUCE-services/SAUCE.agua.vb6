VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyDImputado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "imputado"

End Property

Public Property Get serieID() As String

    serieID = "serie_id"

End Property

Public Property Get serieID_t() As String

    serieID_t = "imputado.serie_id"

End Property

Public Property Get numeroID() As String

    numeroID = "numero_id"

End Property

Public Property Get numeroID_t() As String

    numeroID_t = "imputado.numero_id"

End Property

Public Property Get tipoID() As String

    tipoID = "tipo_id"

End Property

Public Property Get tipoID_t() As String

    tipoID_t = "imputado.tipo_id"

End Property

Public Property Get clienteID() As String

    clienteID = "cliente_id"

End Property

Public Property Get clienteID_t() As String

    clienteID_t = "imputado.cliente_id"

End Property

Public Property Get compSerieID() As String

    compSerieID = "comp_serie_id"

End Property

Public Property Get compSerieID_t() As String

    compSerieID_t = "imputado.comp_serie_id"

End Property

Public Property Get compNumeroID() As String

    compNumeroID = "comp_numero_id"

End Property

Public Property Get compNumeroID_t() As String

    compNumeroID_t = "imputado.comp_numero_id"

End Property

Public Property Get fecha() As String

    fecha = "fecha"

End Property

Public Property Get fecha_t() As String

    fecha_t = "imputado.fecha"

End Property

Public Property Get autoID() As String

    autoID = "auto_id"

End Property

Public Property Get autoID_t() As String

    autoID_t = "imputado.auto_id"

End Property

Public Property Get uid() As String

    uid = "uid"

End Property

Public Property Get uid_t() As String

    uid_t = "imputado.uid"

End Property

Public Property Get created() As String

    created = "created"

End Property

Public Property Get created_t() As String

    created_t = "imputado.created"

End Property

