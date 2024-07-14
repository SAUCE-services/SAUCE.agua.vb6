VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyDMedicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "medicion"

End Property

Public Property Get clienteID() As String

    clienteID = "cliente_id"

End Property

Public Property Get clienteID_t() As String

    clienteID_t = "medicion.cliente_id"

End Property

Public Property Get periodoID() As String

    periodoID = "periodo_id"

End Property

Public Property Get periodoID_t() As String

    periodoID_t = "medicion.periodo_id"

End Property

Public Property Get medidorID() As String

    medidorID = "medidor_id"

End Property

Public Property Get medidorID_t() As String

    medidorID_t = "medicion.medidor_id"

End Property

Public Property Get fechaLectura() As String

    fechaLectura = "fecha_lectura"

End Property

Public Property Get fechaLectura_t() As String

    fechaLectura_t = "medicion.fecha_lectura"

End Property

Public Property Get estado() As String

    estado = "estado"

End Property

Public Property Get estado_t() As String

    estado_t = "medicion.estado"

End Property

Public Property Get autoID() As String

    autoID = "auto_id"

End Property

Public Property Get autoID_t() As String

    autoID_t = "medicion.auto_id"

End Property

Public Property Get uid() As String

    uid = "uid"

End Property

Public Property Get uid_t() As String

    uid_t = "medicion.uid"

End Property

Public Property Get created() As String

    created = "created"

End Property

Public Property Get created_t() As String

    created_t = "medicion.created"

End Property

