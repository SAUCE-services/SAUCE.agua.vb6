VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyDAlicuota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "alicuota"

End Property

Public Property Get ivaCF() As String

    ivaCF = "iva_cf"

End Property

Public Property Get ivaCF_t() As String

    ivaCF_t = "alicuota.iva_cf"

End Property

Public Property Get iva() As String

    iva = "iva"

End Property

Public Property Get iva_t() As String

    iva_t = "alicuota.iva"

End Property

Public Property Get rni() As String

    rni = "rni"

End Property

Public Property Get rni_t() As String

    rni_t = "alicuota.rni"

End Property

Public Property Get fecha() As String

    fecha = "fecha"

End Property

Public Property Get fecha_t() As String

    fecha_t = "alicuota.fecha"

End Property

Public Property Get autoID() As String

    autoID = "auto_id"

End Property

Public Property Get autoID_t() As String

    autoID_t = "alicuota.auto_id"

End Property

Public Property Get uid() As String

    uid = "uid"

End Property

Public Property Get uid_t() As String

    uid_t = "alicuota.uid"

End Property

Public Property Get created() As String

    created = "created"

End Property

Public Property Get created_t() As String

    created_t = "alicuota.created"

End Property

