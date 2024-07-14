VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyDDesconexion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "desconexion"

End Property

Public Property Get clienteID() As String

    clienteID = "cliente_id"

End Property

Public Property Get clienteID_t() As String

    clienteID_t = "desconexion.cliente_id"

End Property

Public Property Get fechaDesconexion() As String

    fechaDesconexion = "fecha_desconexion"

End Property

Public Property Get fechaDesconexion_t() As String

    fechaDesconexion_t = "desconexion.fecha_desconexion"

End Property

Public Property Get fechaReconexion() As String

    fechaReconexion = "fecha_reconexion"

End Property

Public Property Get fechaReconexion_t() As String

    fechaReconexion_t = "desconexion.fecha_reconexion"

End Property

Public Property Get motivo() As String

    motivo = "motivo"

End Property

Public Property Get motivo_t() As String

    motivo_t = "desconexion.motivo"

End Property

Public Property Get autoID() As String

    autoID = "auto_id"

End Property

Public Property Get autoID_t() As String

    autoID_t = "desconexion.auto_id"

End Property

Public Property Get uid() As String

    uid = "uid"

End Property

Public Property Get uid_t() As String

    uid_t = "desconexion.uid"

End Property

Public Property Get created() As String

    created = "created"

End Property

Public Property Get created_t() As String

    created_t = "desconexion.created"

End Property

