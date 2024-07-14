VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyDAnotador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "anotador"

End Property

Public Property Get anotadorID() As String

    anotadorID = "anotador_id"

End Property

Public Property Get anotadorID_t() As String

    anotadorID_t = "anotador.anotador_id"

End Property

Public Property Get clienteID() As String

    clienteID = "cliente_id"

End Property

Public Property Get clienteID_t() As String

    clienteID_t = "anotador.cliente_id"

End Property

Public Property Get anotacion() As String

    anotacion = "anotacion"

End Property

Public Property Get anotacion_t() As String

    anotacion_t = "anotador.anotacion"

End Property

Public Property Get created() As String

    created = "created"

End Property

Public Property Get created_t() As String

    created_t = "anotador.created"

End Property

