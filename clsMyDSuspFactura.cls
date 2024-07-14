VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyDSuspFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "suspfactura"

End Property

Public Property Get clienteID() As String

    clienteID = "cliente_id"

End Property

Public Property Get clienteID_t() As String

    clienteID_t = "suspfactura.cliente_id"

End Property

Public Property Get periodoIDInicio() As String

    periodoIDInicio = "periodo_id_inicio"

End Property

Public Property Get periodoIDInicio_t() As String

    periodoIDInicio_t = "suspfactura.periodo_id_inicio"

End Property

Public Property Get periodoIDFin() As String

    periodoIDFin = "periodo_id_fin"

End Property

Public Property Get periodoIDFin_t() As String

    periodoIDFin_t = "suspfactura.periodo_id_fin"

End Property

Public Property Get motivo() As String

    motivo = "motivo"

End Property

Public Property Get motivo_t() As String

    motivo_t = "suspfactura.motivo"

End Property

Public Property Get autoID() As String

    autoID = "auto_id"

End Property

Public Property Get autoID_t() As String

    autoID_t = "suspfactura.auto_id"

End Property

Public Property Get uid() As String

    uid = "uid"

End Property

Public Property Get uid_t() As String

    uid_t = "suspfactura.uid"

End Property

Public Property Get created() As String

    created = "created"

End Property

Public Property Get created_t() As String

    created_t = "suspfactura.created"

End Property

