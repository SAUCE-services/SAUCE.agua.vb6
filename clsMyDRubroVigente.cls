VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0   'vbNone
  MTSTransactionMode  = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsMyDRubroVigente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "rubrovigente"

End Property

Public Property Get rubroID() As String

    rubroID = "rubro_id"

End Property

Public Property Get rubroID_t() As String

    rubroID_t = "rubrovigente.rubro_id"

End Property

Public Property Get fecha() As String

    fecha = "fecha"

End Property

Public Property Get fecha_t() As String

    fecha_t = "rubrovigente.fecha"

End Property

Public Property Get rangoID() As String

    rangoID = "rango_id"

End Property

Public Property Get rangoID_t() As String

    rangoID_t = "rubrovigente.rango_id"

End Property

Public Property Get concepto() As String

    concepto = "concepto"

End Property

Public Property Get concepto_t() As String

    concepto_t = "rubrovigente.concepto"

End Property

Public Property Get precioUnitario() As String

    precioUnitario = "precio_unitario"

End Property

Public Property Get precioUnitario_t() As String

    precioUnitario_t = "rubrovigente.precio_unitario"

End Property

Public Property Get iva() As String

    iva = "iva"

End Property

Public Property Get iva_t() As String

    iva_t = "rubrovigente.iva"

End Property

Public Property Get comun() As String

    comun = "comun"

End Property

Public Property Get comun_t() As String

    comun_t = "rubrovigente.comun"

End Property

Public Property Get comunSocio() As String

    comunSocio = "comun_socio"

End Property

Public Property Get comunSocio_t() As String

    comunSocio_t = "rubrovigente.comun_socio"

End Property

Public Property Get cobro() As String

    cobro = "cobro"

End Property

Public Property Get cobro_t() As String

    cobro_t = "rubrovigente.cobro"

End Property

Public Property Get desconectado() As String

    desconectado = "desconectado"

End Property

Public Property Get desconectado_t() As String

    desconectado_t = "rubrovigente.desconectado"

End Property

Public Property Get autoID() As String

    autoID = "auto_id"

End Property

Public Property Get autoID_t() As String

    autoID_t = "rubrovigente.auto_id"

End Property

Public Property Get uid() As String

    uid = "uid"

End Property

Public Property Get uid_t() As String

    uid_t = "rubrovigente.uid"

End Property

Public Property Get created() As String

    created = "created"

End Property

Public Property Get created_t() As String

    created_t = "rubrovigente.created"

End Property

