VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyDRango"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "rango"

End Property

Public Property Get categoria() As String

    categoria = "categoria"

End Property

Public Property Get categoria_t() As String

    categoria_t = "rango.categoria"

End Property

Public Property Get rangoID() As String

    rangoID = "rango_id"

End Property

Public Property Get rangoID_t() As String

    rangoID_t = "rango.rango_id"

End Property

Public Property Get fecha() As String

    fecha = "fecha"

End Property

Public Property Get fecha_t() As String

    fecha_t = "rango.fecha"

End Property

Public Property Get limiteInferior() As String

    limiteInferior = "limite_inferior"

End Property

Public Property Get limiteInferior_t() As String

    limiteInferior_t = "rango.limite_inferior"

End Property

Public Property Get limiteSuperior() As String

    limiteSuperior = "limite_superior"

End Property

Public Property Get limiteSuperior_t() As String

    limiteSuperior_t = "rango.limite_superior"

End Property

Public Property Get tarifa() As String

    tarifa = "tarifa"

End Property

Public Property Get tarifa_t() As String

    tarifa_t = "rango.tarifa"

End Property

Public Property Get autoID() As String

    autoID = "auto_id"

End Property

Public Property Get autoID_t() As String

    autoID_t = "rango.auto_id"

End Property

Public Property Get uid() As String

    uid = "uid"

End Property

Public Property Get uid_t() As String

    uid_t = "rango.uid"

End Property

Public Property Get created() As String

    created = "created"

End Property

Public Property Get created_t() As String

    created_t = "rango.created"

End Property

