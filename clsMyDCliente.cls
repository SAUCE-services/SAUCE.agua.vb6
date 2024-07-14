VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyDCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "cliente"

End Property

Public Property Get clienteID() As String

    clienteID = "cliente_id"

End Property

Public Property Get clienteID_t() As String

    clienteID_t = "cliente.cliente_id"

End Property

Public Property Get fechaAlta() As String

    fechaAlta = "fecha_alta"

End Property

Public Property Get fechaAlta_t() As String

    fechaAlta_t = "cliente.fecha_alta"

End Property

Public Property Get fechaBaja() As String

    fechaBaja = "fecha_baja"

End Property

Public Property Get fechaBaja_t() As String

    fechaBaja_t = "cliente.fecha_baja"

End Property

Public Property Get apellido() As String

    apellido = "apellido"

End Property

Public Property Get apellido_t() As String

    apellido_t = "cliente.apellido"

End Property

Public Property Get nombre() As String

    nombre = "nombre"

End Property

Public Property Get nombre_t() As String

    nombre_t = "cliente.nombre"

End Property

Public Property Get numeroSocio() As String

    numeroSocio = "numero_socio"

End Property

Public Property Get numeroSocio_t() As String

    numeroSocio_t = "cliente.numero_socio"

End Property

Public Property Get inmuebleCalle() As String

    inmuebleCalle = "inmueble_calle"

End Property

Public Property Get inmuebleCalle_t() As String

    inmuebleCalle_t = "cliente.inmueble_calle"

End Property

Public Property Get inmueblePuerta() As String

    inmueblePuerta = "inmueble_puerta"

End Property

Public Property Get inmueblePuerta_t() As String

    inmueblePuerta_t = "cliente.inmueble_puerta"

End Property

Public Property Get inmueblePiso() As String

    inmueblePiso = "inmueble_piso"

End Property

Public Property Get inmueblePiso_t() As String

    inmueblePiso_t = "cliente.inmueble_piso"

End Property

Public Property Get inmuebleDpto() As String

    inmuebleDpto = "inmueble_dpto"

End Property

Public Property Get inmuebleDpto_t() As String

    inmuebleDpto_t = "cliente.inmueble_dpto"

End Property

Public Property Get inmuebleLocalidad() As String

    inmuebleLocalidad = "inmueble_localidad"

End Property

Public Property Get inmuebleLocalidad_t() As String

    inmuebleLocalidad_t = "cliente.inmueble_localidad"

End Property

Public Property Get inmuebleProvincia() As String

    inmuebleProvincia = "inmueble_provincia"

End Property

Public Property Get inmuebleProvincia_t() As String

    inmuebleProvincia_t = "cliente.inmueble_provincia"

End Property

Public Property Get inmuebleCodpostal() As String

    inmuebleCodpostal = "inmueble_codpostal"

End Property

Public Property Get inmuebleCodpostal_t() As String

    inmuebleCodpostal_t = "cliente.inmueble_codpostal"

End Property

Public Property Get fiscalCalle() As String

    fiscalCalle = "fiscal_calle"

End Property

Public Property Get fiscalCalle_t() As String

    fiscalCalle_t = "cliente.fiscal_calle"

End Property

Public Property Get fiscalPuerta() As String

    fiscalPuerta = "fiscal_puerta"

End Property

Public Property Get fiscalPuerta_t() As String

    fiscalPuerta_t = "cliente.fiscal_puerta"

End Property

Public Property Get fiscalPiso() As String

    fiscalPiso = "fiscal_piso"

End Property

Public Property Get fiscalPiso_t() As String

    fiscalPiso_t = "cliente.fiscal_piso"

End Property

Public Property Get fiscalDpto() As String

    fiscalDpto = "fiscal_dpto"

End Property

Public Property Get fiscalDpto_t() As String

    fiscalDpto_t = "cliente.fiscal_dpto"

End Property

Public Property Get fiscalLocalidad() As String

    fiscalLocalidad = "fiscal_localidad"

End Property

Public Property Get fiscalLocalidad_t() As String

    fiscalLocalidad_t = "cliente.fiscal_localidad"

End Property

Public Property Get fiscalProvincia() As String

    fiscalProvincia = "fiscal_provincia"

End Property

Public Property Get fiscalProvincia_t() As String

    fiscalProvincia_t = "cliente.fiscal_provincia"

End Property

Public Property Get fiscalCodpostal() As String

    fiscalCodpostal = "fiscal_codpostal"

End Property

Public Property Get fiscalCodpostal_t() As String

    fiscalCodpostal_t = "cliente.fiscal_codpostal"

End Property

Public Property Get cuit() As String

    cuit = "cuit"

End Property

Public Property Get cuit_t() As String

    cuit_t = "cliente.cuit"

End Property

Public Property Get situacionIVA() As String

    situacionIVA = "situacion_iva"

End Property

Public Property Get situacionIva_t() As String

    situacionIva_t = "cliente.situacion_iva"

End Property

Public Property Get nomenclaturaCatastral() As String

    nomenclaturaCatastral = "nombre_categoria"

End Property

Public Property Get nomenclaturaCatastral_t() As String

    nomenclaturaCatastral_t = "cliente.nombre_categoria"

End Property

Public Property Get categoria() As String

    categoria = "categoria"

End Property

Public Property Get categoria_t() As String

    categoria_t = "cliente.categoria"

End Property

Public Property Get servicio() As String

    servicio = "servicio"

End Property

Public Property Get servicio_t() As String

    servicio_t = "cliente.servicio"

End Property

Public Property Get cobro() As String

    cobro = "cobro"

End Property

Public Property Get cobro_t() As String

    cobro_t = "cliente.cobro"

End Property

Public Property Get zona() As String

    zona = "zona"

End Property

Public Property Get zona_t() As String

    zona_t = "cliente.zona"

End Property

Public Property Get ruta() As String

    ruta = "ruta"

End Property

Public Property Get ruta_t() As String

    ruta_t = "cliente.ruta"

End Property

Public Property Get orden() As String

    orden = "orden"

End Property

Public Property Get orden_t() As String

    orden_t = "cliente.orden"

End Property

Public Property Get estadoID() As String

    estadoID = "estado_id"

End Property

Public Property Get estadoID_t() As String

    estadoID_t = "cliente.estado_id"

End Property

Public Property Get fechaNacimiento() As String

    fechaNacimiento = "fecha_nacimiento"

End Property

Public Property Get fechaNacimiento_t() As String

    fechaNacimiento_t = "cliente.fecha_nacimiento"

End Property

Public Property Get categoriasocioID() As String

    categoriasocioID = "categoriasocio_id"

End Property

Public Property Get categoriasocioID_t() As String

    categoriasocioID_t = "cliente.categoriasocio_id"

End Property

Public Property Get destinoID() As String

    destinoID = "destino_id"

End Property

Public Property Get destinoID_t() As String

    destinoID_t = "cliente.destino_id"

End Property

Public Property Get autoID() As String

    autoID = "auto_id"

End Property

Public Property Get autoID_t() As String

    autoID_t = "cliente.auto_id"

End Property

Public Property Get uid() As String

    uid = "uid"

End Property

Public Property Get uid_t() As String

    uid_t = "cliente.uid"

End Property

Public Property Get created() As String

    created = "created"

End Property

Public Property Get created_t() As String

    created_t = "cliente.created"

End Property

