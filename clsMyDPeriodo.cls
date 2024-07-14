VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyDPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "periodo"

End Property

Public Property Get periodoID() As String

    periodoID = "periodo_id"

End Property

Public Property Get periodoID_t() As String

    periodoID_t = "periodo.periodo_id"

End Property

Public Property Get descripcion() As String

    descripcion = "descripcion"

End Property

Public Property Get descripcion_t() As String

    descripcion_t = "periodo.descripcion"

End Property

Public Property Get fechaInicio() As String

    fechaInicio = "fecha_inicio"

End Property

Public Property Get fechaInicio_t() As String

    fechaInicio_t = "periodo.fecha_inicio"

End Property

Public Property Get fechaFin() As String

    fechaFin = "fecha_fin"

End Property

Public Property Get fechaFin_t() As String

    fechaFin_t = "periodo.fecha_fin"

End Property

Public Property Get fechaPrimero() As String

    fechaPrimero = "fecha_primero"

End Property

Public Property Get fechaPrimero_t() As String

    fechaPrimero_t = "periodo.fecha_primero"

End Property

Public Property Get fechaSegundo() As String

    fechaSegundo = "fecha_segundo"

End Property

Public Property Get fechaSegundo_t() As String

    fechaSegundo_t = "periodo.fecha_segundo"

End Property

Public Property Get tasa() As String

    tasa = "tasa"

End Property

Public Property Get tasa_t() As String

    tasa_t = "periodo.tasa"

End Property

Public Property Get leyenda() As String

    leyenda = "leyenda"

End Property

Public Property Get leyenda_t() As String

    leyenda_t = "periodo.leyenda"

End Property

Public Property Get liquidado() As String

    liquidado = "liquidado"

End Property

Public Property Get liquidado_t() As String

    liquidado_t = "periodo.liquidado"

End Property

Public Property Get uid() As String

    uid = "uid"

End Property

Public Property Get uid_t() As String

    uid_t = "periodo.uid"

End Property

Public Property Get created() As String

    created = "created"

End Property

Public Property Get created_t() As String

    created_t = "periodo.created"

End Property

