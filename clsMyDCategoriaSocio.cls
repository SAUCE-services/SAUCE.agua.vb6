VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0   'vbNone
  MTSTransactionMode  = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsMyDCategoriaSocio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "categoria_socio"

End Property

Public Property Get categoriasocioID() As String

    categoriasocioID = "categoriasocio_id"

End Property

Public Property Get categoriasocioID_t() As String

    categoriasocioID_t = "categoria_socio.categoriasocio_id"

End Property

Public Property Get nombre() As String

    nombre = "nombre"

End Property

Public Property Get nombre_t() As String

    nombre_t = "categoria_socio.nombre"

End Property

Public Property Get autoID() As String

    autoID = "auto_id"

End Property

Public Property Get autoID_t() As String

    autoID_t = "categoria_socio.auto_id"

End Property

Public Property Get created() As String

    created = "created"

End Property

Public Property Get created_t() As String

    created_t = "categoria_socio.created"

End Property

