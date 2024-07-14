VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyDCuota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "cuota"

End Property

Public Property Get clienteID() As String

    clienteID = "cliente_id"

End Property

Public Property Get clienteID_t() As String

    clienteID_t = "cuota.cliente_id"

End Property

Public Property Get planID() As String

    planID = "plan_id"

End Property

Public Property Get planID_t() As String

    planID_t = "cuota.plan_id"

End Property

Public Property Get cuotaID() As String

    cuotaID = "cuota_id"

End Property

Public Property Get cuotaID_t() As String

    cuotaID_t = "cuota.cuota_id"

End Property

Public Property Get fechaVencimiento() As String

    fechaVencimiento = "fecha_vencimiento"

End Property

Public Property Get fechaVencimiento_t() As String

    fechaVencimiento_t = "cuota.fecha_vencimiento"

End Property

Public Property Get fechaPago() As String

    fechaPago = "fecha_pago"

End Property

Public Property Get fechaPago_t() As String

    fechaPago_t = "cuota.fecha_pago"

End Property

Public Property Get importe() As String

    importe = "importe"

End Property

Public Property Get importe_t() As String

    importe_t = "cuota.importe"

End Property

Public Property Get cancelada() As String

    cancelada = "cancelada"

End Property

Public Property Get cancelada_t() As String

    cancelada_t = "cuota.cancelada"

End Property

Public Property Get planIDCancela() As String

    planIDCancela = "plan_id_cancela"

End Property

Public Property Get planIDCancela_t() As String

    planIDCancela_t = "cuota.plan_id_cancela"

End Property

Public Property Get autoID() As String

    autoID = "auto_id"

End Property

Public Property Get autoID_t() As String

    autoID_t = "cuota.auto_id"

End Property

Public Property Get uid() As String

    uid = "uid"

End Property

Public Property Get uid_t() As String

    uid_t = "cuota.uid"

End Property

Public Property Get created() As String

    created = "created"

End Property

Public Property Get created_t() As String

    created_t = "cuota.created"

End Property

