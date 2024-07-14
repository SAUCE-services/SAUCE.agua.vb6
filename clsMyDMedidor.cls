VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyDMedidor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "medidor"

End Property

Public Property Get medidorID() As String

    medidorID = "medidor_id"

End Property

Public Property Get medidorID_t() As String

    medidorID_t = "medidor.medidor_id"

End Property

Public Property Get fechaAlta() As String

    fechaAlta = "fecha_alta"

End Property

Public Property Get fechaAlta_t() As String

    fechaAlta_t = "medidor.fecha_alta"

End Property

Public Property Get clienteID() As String

    clienteID = "cliente_id"

End Property

Public Property Get clienteID_t() As String

    clienteID_t = "medidor.cliente_id"

End Property

Public Property Get fechaColocacion() As String

    fechaColocacion = "fecha_colocacion"

End Property

Public Property Get fechaColocacion_t() As String

    fechaColocacion_t = "medidor.fecha_colocacion"

End Property

Public Property Get fechaRetiro() As String

    fechaRetiro = "fecha_retiro"

End Property

Public Property Get fechaRetiro_t() As String

    fechaRetiro_t = "medidor.fecha_retiro"

End Property

Public Property Get motivoRetiro() As String

    motivoRetiro = "motivo_retiro"

End Property

Public Property Get motivoRetiro_t() As String

    motivoRetiro_t = "medidor.motivo_retiro"

End Property

Public Property Get estadoInicio() As String

    estadoInicio = "estado_inicio"

End Property

Public Property Get estadoInicio_t() As String

    estadoInicio_t = "medidor.estado_inicio"

End Property

Public Property Get autoID() As String

    autoID = "auto_id"

End Property

Public Property Get autoID_t() As String

    autoID_t = "medidor.auto_id"

End Property

Public Property Get uid() As String

    uid = "uid"

End Property

Public Property Get uid_t() As String

    uid_t = "medidor.uid"

End Property

Public Property Get created() As String

    created = "created"

End Property

Public Property Get created_t() As String

    created_t = "medidor.created"

End Property

