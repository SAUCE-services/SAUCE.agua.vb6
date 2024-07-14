VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyDSuspension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "suspension"

End Property

Public Property Get tipo() As String

    tipo = "tipo"

End Property

Public Property Get tipo_t() As String

    tipo_t = "suspension.tipo"

End Property

Public Property Get numero() As String

    numero = "numero"

End Property

Public Property Get numero_t() As String

    numero_t = "suspension.numero"

End Property

Public Property Get fecha() As String

    fecha = "fecha"

End Property

Public Property Get fecha_t() As String

    fecha_t = "suspension.fecha"

End Property

Public Property Get clienteID() As String

    clienteID = "cliente_id"

End Property

Public Property Get clienteID_t() As String

    clienteID_t = "suspension.cliente_id"

End Property

Public Property Get periodoID() As String

    periodoID = "periodo_id"

End Property

Public Property Get periodoID_t() As String

    periodoID_t = "suspension.periodo_id"

End Property

Public Property Get autoID() As String

    autoID = "auto_id"

End Property

Public Property Get autoID_t() As String

    autoID_t = "suspension.auto_id"

End Property

Public Property Get uid() As String

    uid = "uid"

End Property

Public Property Get uid_t() As String

    uid_t = "suspension.uid"

End Property

Public Property Get created() As String

    created = "created"

End Property

Public Property Get created_t() As String

    created_t = "suspension.created"

End Property

