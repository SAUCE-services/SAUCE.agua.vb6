VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyDNovedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "novedad"

End Property

Public Property Get clienteID() As String

    clienteID = "cliente_id"

End Property

Public Property Get clienteID_t() As String

    clienteID_t = "novedad.cliente_id"

End Property

Public Property Get periodoID() As String

    periodoID = "periodo_id"

End Property

Public Property Get periodoID_t() As String

    periodoID_t = "novedad.periodo_id"

End Property

Public Property Get rubroID() As String

    rubroID = "rubro_id"

End Property

Public Property Get rubroID_t() As String

    rubroID_t = "novedad.rubro_id"

End Property

Public Property Get fecha() As String

    fecha = "fecha"

End Property

Public Property Get fecha_t() As String

    fecha_t = "novedad.fecha"

End Property

Public Property Get porcentaje() As String

    porcentaje = "porcentaje"

End Property

Public Property Get porcentaje_t() As String

    porcentaje_t = "novedad.porcentaje"

End Property

Public Property Get cantidad() As String

    cantidad = "cantidad"

End Property

Public Property Get cantidad_t() As String

    cantidad_t = "novedad.cantidad"

End Property

Public Property Get importe() As String

    importe = "importe"

End Property

Public Property Get importe_t() As String

    importe_t = "novedad.importe"

End Property

Public Property Get veces() As String

    veces = "veces"

End Property

Public Property Get veces_t() As String

    veces_t = "novedad.veces"

End Property

Public Property Get vecesCobradas() As String

    vecesCobradas = "veces_cobradas"

End Property

Public Property Get vecesCobradas_t() As String

    vecesCobradas_t = "novedad.veces_cobradas"

End Property

Public Property Get indefinida() As String

    indefinida = "indefinida"

End Property

Public Property Get indefinida_t() As String

    indefinida_t = "novedad.indefinida"

End Property

Public Property Get periodoIdSuspension() As String

    periodoIdSuspension = "periodo_id_suspension"

End Property

Public Property Get periodoIdSuspension_t() As String

    periodoIdSuspension_t = "novedad.periodo_id_suspension"

End Property

Public Property Get autoID() As String

    autoID = "auto_id"

End Property

Public Property Get autoID_t() As String

    autoID_t = "novedad.auto_id"

End Property

Public Property Get uid() As String

    uid = "uid"

End Property

Public Property Get uid_t() As String

    uid_t = "novedad.uid"

End Property

Public Property Get created() As String

    created = "created"

End Property

Public Property Get created_t() As String

    created_t = "novedad.created"

End Property

