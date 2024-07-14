VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0   'vbNone
  MTSTransactionMode  = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsVMyDListGral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "vw_listgral"

End Property

Public Property Get medidorID() As String

    medidorID = "medidor_id"

End Property

Public Property Get medidorID_t() As String

    medidorID_t = "vw_listgral.medidor_id"

End Property

Public Property Get clienteID() As String

    clienteID = "cliente_id"

End Property

Public Property Get clienteID_t() As String

    clienteID_t = "vw_listgral.cliente_id"

End Property

Public Property Get apellido() As String

    apellido = "apellido"

End Property

Public Property Get apellido_t() As String

    apellido_t = "vw_listgral.apellido"

End Property

Public Property Get nombre() As String

    nombre = "nombre"

End Property

Public Property Get nombre_t() As String

    nombre_t = "vw_listgral.nombre"

End Property

Public Property Get inmuebleCalle() As String

    inmuebleCalle = "inmueble_calle"

End Property

Public Property Get inmuebleCalle_t() As String

    inmuebleCalle_t = "vw_listgral.inmueble_calle"

End Property

Public Property Get inmueblePuerta() As String

    inmueblePuerta = "inmueble_puerta"

End Property

Public Property Get inmueblePuerta_t() As String

    inmueblePuerta_t = "vw_listgral.inmueble_puerta"

End Property

Public Property Get inmueblePiso() As String

    inmueblePiso = "inmueble_piso"

End Property

Public Property Get inmueblePiso_t() As String

    inmueblePiso_t = "vw_listgral.inmueble_piso"

End Property

Public Property Get inmuebleDpto() As String

    inmuebleDpto = "inmueble_dpto"

End Property

Public Property Get inmuebleDpto_t() As String

    inmuebleDpto_t = "vw_listgral.inmueble_dpto"

End Property

Public Property Get inmuebleLocalidad() As String

    inmuebleLocalidad = "inmueble_localidad"

End Property

Public Property Get inmuebleLocalidad_t() As String

    inmuebleLocalidad_t = "vw_listgral.inmueble_localidad"

End Property

Public Property Get inmuebleProvincia() As String

    inmuebleProvincia = "inmueble_provincia"

End Property

Public Property Get inmuebleProvincia_t() As String

    inmuebleProvincia_t = "vw_listgral.inmueble_provincia"

End Property

Public Property Get fiscalCalle() As String

    fiscalCalle = "fiscal_calle"

End Property

Public Property Get fiscalCalle_t() As String

    fiscalCalle_t = "vw_listgral.fiscal_calle"

End Property

Public Property Get fiscalPuerta() As String

    fiscalPuerta = "fiscal_puerta"

End Property

Public Property Get fiscalPuerta_t() As String

    fiscalPuerta_t = "vw_listgral.fiscal_puerta"

End Property

Public Property Get fiscalPiso() As String

    fiscalPiso = "fiscal_piso"

End Property

Public Property Get fiscalPiso_t() As String

    fiscalPiso_t = "vw_listgral.fiscal_piso"

End Property

Public Property Get fiscalDpto() As String

    fiscalDpto = "fiscal_dpto"

End Property

Public Property Get fiscalDpto_t() As String

    fiscalDpto_t = "vw_listgral.fiscal_dpto"

End Property

Public Property Get fiscalLocalidad() As String

    fiscalLocalidad = "fiscal_localidad"

End Property

Public Property Get fiscalLocalidad_t() As String

    fiscalLocalidad_t = "vw_listgral.fiscal_localidad"

End Property

Public Property Get fiscalProvincia() As String

    fiscalProvincia = "fiscal_provincia"

End Property

Public Property Get fiscalProvincia_t() As String

    fiscalProvincia_t = "vw_listgral.fiscal_provincia"

End Property

Public Property Get zona() As String

    zona = "zona"

End Property

Public Property Get zona_t() As String

    zona_t = "vw_listgral.zona"

End Property

Public Property Get ruta() As String

    ruta = "ruta"

End Property

Public Property Get ruta_t() As String

    ruta_t = "vw_listgral.ruta"

End Property

Public Property Get orden() As String

    orden = "orden"

End Property

Public Property Get orden_t() As String

    orden_t = "vw_listgral.orden"

End Property

Public Property Get ultima() As String

    ultima = "ultima"

End Property

Public Property Get ultima_t() As String

    ultima_t = "vw_listgral.ultima"

End Property

Public Property Get estadoInicio() As String

    estadoInicio = "estado_inicio"

End Property

Public Property Get estadoInicio_t() As String

    estadoInicio_t = "vw_listgral.estado_inicio"

End Property

Public Property Get estado() As String

    estado = "estado"

End Property

Public Property Get estado_t() As String

    estado_t = "vw_listgral.estado"

End Property

