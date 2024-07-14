VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0   'vbNone
  MTSTransactionMode  = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsMyDEstado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "estado"

End Property

Public Property Get estadoID() As String

    estadoID = "estado_id"

End Property

Public Property Get estadoID_t() As String

    estadoID_t = "estado.estado_id"

End Property

Public Property Get nombre() As String

    nombre = "nombre"

End Property

Public Property Get nombre_t() As String

    nombre_t = "estado.nombre"

End Property

Public Property Get autoID() As String

    autoID = "auto_id"

End Property

Public Property Get autoID_t() As String

    autoID_t = "estado.auto_id"

End Property

Public Property Get created() As String

    created = "created"

End Property

Public Property Get created_t() As String

    created_t = "estado.created"

End Property

