VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyDDeuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "deuda"

End Property

Public Property Get clienteID() As String

    clienteID = "cliente_id"

End Property

Public Property Get clienteID_t() As String

    clienteID_t = "deuda.cliente_id"

End Property

Public Property Get planID() As String

    planID = "plan_id"

End Property

Public Property Get planID_t() As String

    planID_t = "deuda.plan_id"

End Property

Public Property Get deuda() As String

    deuda = "deuda"

End Property

Public Property Get deuda_t() As String

    deuda_t = "deuda.deuda"

End Property

Public Property Get cuotas() As String

    cuotas = "cuotas"

End Property

Public Property Get cuotas_t() As String

    cuotas_t = "deuda.cuotas"

End Property

Public Property Get cuotasPagadas() As String

    cuotasPagadas = "cuotas_pagadas"

End Property

Public Property Get cuotasPagadas_t() As String

    cuotasPagadas_t = "deuda.cuotas_pagadas"

End Property

Public Property Get tasa() As String

    tasa = "tasa"

End Property

Public Property Get tasa_t() As String

    tasa_t = "deuda.tasa"

End Property

Public Property Get pagado() As String

    pagado = "pagado"

End Property

Public Property Get pagado_t() As String

    pagado_t = "deuda.pagado"

End Property

Public Property Get periodo() As String

    periodo = "periodo"

End Property

Public Property Get periodo_t() As String

    periodo_t = "deuda.periodo"

End Property

Public Property Get cancelada() As String

    cancelada = "cancelada"

End Property

Public Property Get cancelada_t() As String

    cancelada_t = "deuda.cancelada"

End Property

Public Property Get planIDCancela() As String

    planIDCancela = "plan_id_cancela"

End Property

Public Property Get planIDCancela_t() As String

    planIDCancela_t = "deuda.plan_id_cancela"

End Property

Public Property Get autoID() As String

    autoID = "auto_id"

End Property

Public Property Get autoID_t() As String

    autoID_t = "deuda.auto_id"

End Property

Public Property Get uid() As String

    uid = "uid"

End Property

Public Property Get uid_t() As String

    uid_t = "deuda.uid"

End Property

Public Property Get created() As String

    created = "created"

End Property

Public Property Get created_t() As String

    created_t = "deuda.created"

End Property

