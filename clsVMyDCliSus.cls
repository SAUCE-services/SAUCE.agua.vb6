VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0   'vbNone
  MTSTransactionMode  = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsVMyDCliSus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "vw_clisus"

End Property

Public Property Get clienteID() As String

    clienteID = "cliente_id"

End Property

Public Property Get clienteID_t() As String

    clienteID_t = "vw_clisus.cliente_id"

End Property

Public Property Get apellido() As String

    apellido = "apellido"

End Property

Public Property Get apellido_t() As String

    apellido_t = "vw_clisus.apellido"

End Property

Public Property Get nombre() As String

    nombre = "nombre"

End Property

Public Property Get nombre_t() As String

    nombre_t = "vw_clisus.nombre"

End Property

Public Property Get inmuebleCalle() As String

    inmuebleCalle = "inmueble_calle"

End Property

Public Property Get inmuebleCalle_t() As String

    inmuebleCalle_t = "vw_clisus.inmueble_calle"

End Property

Public Property Get inmueblePuerta() As String

    inmueblePuerta = "inmueble_puerta"

End Property

Public Property Get inmueblePuerta_t() As String

    inmueblePuerta_t = "vw_clisus.inmueble_puerta"

End Property

Public Property Get inmueblePiso() As String

    inmueblePiso = "inmueble_piso"

End Property

Public Property Get inmueblePiso_t() As String

    inmueblePiso_t = "vw_clisus.inmueble_piso"

End Property

Public Property Get inmuebleDpto() As String

    inmuebleDpto = "inmueble_dpto"

End Property

Public Property Get inmuebleDpto_t() As String

    inmuebleDpto_t = "vw_clisus.inmueble_dpto"

End Property

Public Property Get cobro() As String

    cobro = "cobro"

End Property

Public Property Get cobro_t() As String

    cobro_t = "vw_clisus.cobro"

End Property

Public Property Get periodoIDInicio() As String

    periodoIDInicio = "periodo_id_inicio"

End Property

Public Property Get periodoIDInicio_t() As String

    periodoIDInicio_t = "vw_clisus.periodo_id_inicio"

End Property

Public Property Get periodoIDFin() As String

    periodoIDFin = "periodo_id_fin"

End Property

Public Property Get periodoIDFin_t() As String

    periodoIDFin_t = "vw_clisus.periodo_id_fin"

End Property

