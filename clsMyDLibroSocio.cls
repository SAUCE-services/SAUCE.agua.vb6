VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyDLibroSocio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "libro_socio"

End Property

Public Property Get numeroSocio() As String

    numeroSocio = "numero_socio"

End Property

Public Property Get numeroSocio_t() As String

    numeroSocio_t = "libro_socio.numero_socio"

End Property

Public Property Get anho() As String

    anho = "anho"

End Property

Public Property Get anho_t() As String

    anho_t = "libro_socio.anho"

End Property

Public Property Get nombreApellido() As String

    nombreApellido = "nombre_apellido"

End Property

Public Property Get nombreApellido_t() As String

    nombreApellido_t = "libro_socio.nombre_apellido"

End Property

Public Property Get domicilio() As String

    domicilio = "domicilio"

End Property

Public Property Get domicilio_t() As String

    domicilio_t = "libro_socio.domicilio"

End Property

Public Property Get documento() As String

    documento = "documento"

End Property

Public Property Get documento_t() As String

    documento_t = "libro_socio.documento"

End Property

Public Property Get estado() As String

    estado = "estado"

End Property

Public Property Get estado_t() As String

    estado_t = "libro_socio.estado"

End Property

Public Property Get edad() As String

    edad = "edad"

End Property

Public Property Get edad_t() As String

    edad_t = "libro_socio.edad"

End Property

Public Property Get categoria() As String

    categoria = "categoria"

End Property

Public Property Get categoria_t() As String

    categoria_t = "libro_socio.categoria"

End Property

Public Property Get ingreso() As String

    ingreso = "ingreso"

End Property

Public Property Get ingreso_t() As String

    ingreso_t = "libro_socio.ingreso"

End Property

Public Property Get enero() As String

    enero = "enero"

End Property

Public Property Get enero_t() As String

    enero_t = "libro_socio.enero"

End Property

Public Property Get febrero() As String

    febrero = "febrero"

End Property

Public Property Get febrero_t() As String

    febrero_t = "libro_socio.febrero"

End Property

Public Property Get marzo() As String

    marzo = "marzo"

End Property

Public Property Get marzo_t() As String

    marzo_t = "libro_socio.marzo"

End Property

Public Property Get abril() As String

    abril = "abril"

End Property

Public Property Get abril_t() As String

    abril_t = "libro_socio.abril"

End Property

Public Property Get mayo() As String

    mayo = "mayo"

End Property

Public Property Get mayo_t() As String

    mayo_t = "libro_socio.mayo"

End Property

Public Property Get junio() As String

    junio = "junio"

End Property

Public Property Get junio_t() As String

    junio_t = "libro_socio.junio"

End Property

Public Property Get julio() As String

    julio = "julio"

End Property

Public Property Get julio_t() As String

    julio_t = "libro_socio.julio"

End Property

Public Property Get agosto() As String

    agosto = "agosto"

End Property

Public Property Get agosto_t() As String

    agosto_t = "libro_socio.agosto"

End Property

Public Property Get setiembre() As String

    setiembre = "setiembre"

End Property

Public Property Get setiembre_t() As String

    setiembre_t = "libro_socio.setiembre"

End Property

Public Property Get octubre() As String

    octubre = "octubre"

End Property

Public Property Get octubre_t() As String

    octubre_t = "libro_socio.octubre"

End Property

Public Property Get noviembre() As String

    noviembre = "noviembre"

End Property

Public Property Get noviembre_t() As String

    noviembre_t = "libro_socio.noviembre"

End Property

Public Property Get diciembre() As String

    diciembre = "diciembre"

End Property

Public Property Get diciembre_t() As String

    diciembre_t = "libro_socio.diciembre"

End Property

Public Property Get autoID() As String

    autoID = "auto_id"

End Property

Public Property Get autoID_t() As String

    autoID_t = "libro_socio.auto_id"

End Property

Public Property Get created() As String

    created = "created"

End Property

Public Property Get created_t() As String

    created_t = "libro_socio.created"

End Property

