VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyDOperador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "operador"

End Property

Public Property Get operadorID() As String

    operadorID = "operador_id"

End Property

Public Property Get operadorID_t() As String

    operadorID_t = "operador.operador_id"

End Property

Public Property Get razonSocial() As String

    razonSocial = "razon_social"

End Property

Public Property Get razonSocial_t() As String

    razonSocial_t = "operador.razon_social"

End Property

Public Property Get calle() As String

    calle = "calle"

End Property

Public Property Get calle_t() As String

    calle_t = "operador.calle"

End Property

Public Property Get puerta() As String

    puerta = "puerta"

End Property

Public Property Get puerta_t() As String

    puerta_t = "operador.puerta"

End Property

Public Property Get piso() As String

    piso = "piso"

End Property

Public Property Get piso_t() As String

    piso_t = "operador.piso"

End Property

Public Property Get dpto() As String

    dpto = "dpto"

End Property

Public Property Get dpto_t() As String

    dpto_t = "operador.dpto"

End Property

Public Property Get codigoPostal() As String

    codigoPostal = "codigo_postal"

End Property

Public Property Get codigoPostal_t() As String

    codigoPostal_t = "operador.codigo_postal"

End Property

Public Property Get localidad() As String

    localidad = "localidad"

End Property

Public Property Get localidad_t() As String

    localidad_t = "operador.localidad"

End Property

Public Property Get provincia() As String

    provincia = "provincia"

End Property

Public Property Get provincia_t() As String

    provincia_t = "operador.provincia"

End Property

Public Property Get telefono() As String

    telefono = "telefono"

End Property

Public Property Get telefono_t() As String

    telefono_t = "operador.telefono"

End Property

Public Property Get cuit() As String

    cuit = "cuit"

End Property

Public Property Get cuit_t() As String

    cuit_t = "operador.cuit"

End Property

Public Property Get ingresosBrutos() As String

    ingresosBrutos = "ingresos_brutos"

End Property

Public Property Get ingresosBrutos_t() As String

    ingresosBrutos_t = "operador.ingresos_brutos"

End Property

Public Property Get situacionIVA() As String

    situacionIVA = "situacion_iva"

End Property

Public Property Get situacionIva_t() As String

    situacionIva_t = "operador.situacion_iva"

End Property

Public Property Get numeroEpas() As String

    numeroEpas = "numero_epas"

End Property

Public Property Get numeroEpas_t() As String

    numeroEpas_t = "operador.numero_epas"

End Property

Public Property Get fechaInicio() As String

    fechaInicio = "fecha_inicio"

End Property

Public Property Get fechaInicio_t() As String

    fechaInicio_t = "operador.fecha_inicio"

End Property

Public Property Get servicio() As String

    servicio = "servicio"

End Property

Public Property Get servicio_t() As String

    servicio_t = "operador.servicio"

End Property

Public Property Get puntoVta() As String

    puntoVta = "prefijo_id"

End Property

Public Property Get puntoVta_t() As String

    puntoVta_t = "operador.prefijo_id"

End Property

Public Property Get nroComprob() As String

    nroComprob = "factura_id"

End Property

Public Property Get nroComprob_t() As String

    nroComprob_t = "operador.factura_id"

End Property

Public Property Get periodoFactura() As String

    periodoFactura = "periodo_factura"

End Property

Public Property Get periodoFactura_t() As String

    periodoFactura_t = "operador.periodo_factura"

End Property

Public Property Get resolucion() As String

    resolucion = "resolucion"

End Property

Public Property Get resolucion_t() As String

    resolucion_t = "operador.resolucion"

End Property

Public Property Get personeria() As String

    personeria = "personeria"

End Property

Public Property Get personeria_t() As String

    personeria_t = "operador.personeria"

End Property

Public Property Get reciboSerie() As String

    reciboSerie = "recibo_serie"

End Property

Public Property Get reciboSerie_t() As String

    reciboSerie_t = "operador.recibo_serie"

End Property

Public Property Get recibo() As String

    recibo = "recibo"

End Property

Public Property Get recibo_t() As String

    recibo_t = "operador.recibo"

End Property

Public Property Get ncreditoSerie() As String

    ncreditoSerie = "ncredito_serie"

End Property

Public Property Get ncreditoSerie_t() As String

    ncreditoSerie_t = "operador.ncredito_serie"

End Property

Public Property Get ncredito() As String

    ncredito = "ncredito"

End Property

Public Property Get ncredito_t() As String

    ncredito_t = "operador.ncredito"

End Property

Public Property Get cai() As String

    cai = "cai"

End Property

Public Property Get cai_t() As String

    cai_t = "operador.cai"

End Property

Public Property Get caiVencimiento() As String

    caiVencimiento = "cai_vencimiento"

End Property

Public Property Get caiVencimiento_t() As String

    caiVencimiento_t = "operador.cai_vencimiento"

End Property

Public Property Get preimpreso() As String

    preimpreso = "preimpreso"

End Property

Public Property Get preimpreso_t() As String

    preimpreso_t = "operador.preimpreso"

End Property

Public Property Get uid() As String

    uid = "uid"

End Property

Public Property Get uid_t() As String

    uid_t = "operador.uid"

End Property

Public Property Get created() As String

    created = "created"

End Property

Public Property Get created_t() As String

    created_t = "operador.created"

End Property

