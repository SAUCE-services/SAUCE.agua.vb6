VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0   'vbNone
  MTSTransactionMode  = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsMyDParametro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "parametro"

End Property

Public Property Get parametroID() As String

    parametroID = "parametro_id"

End Property

Public Property Get parametroID_t() As String

    parametroID_t = "parametro.parametro_id"

End Property

Public Property Get feProduccion() As String

    feProduccion = "fe_produccion"

End Property

Public Property Get feProduccion_t() As String

    feProduccion_t = "parametro.fe_produccion"

End Property

Public Property Get uid() As String

    uid = "uid"

End Property

Public Property Get uid_t() As String

    uid_t = "parametro.uid"

End Property

Public Property Get created() As String

    created = "created"

End Property

Public Property Get created_t() As String

    created_t = "parametro.created"

End Property

