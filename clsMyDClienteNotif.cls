VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0   'vbNone
  MTSTransactionMode  = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsMyDClienteNotif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Asocia TABLA
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get table() As String

    table = "cliente_notificacion"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Asocia clienteID
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get clienteID() As String

    clienteID = "cliente_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Asocia clienteID completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get clienteID_t() As String

    clienteID_t = table & ".cliente_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Asocia ultimaNotificacion15
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get ultimaNotificacion15() As String

    ultimaNotificacion15 = "ultima_notificacion_15"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Asocia ultimaNotificacion15 completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get ultimaNotificacion15_t() As String

    ultimaNotificacion15_t = table & ".ultima_notificacion_15"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Asocia ultimaNotificacion48
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get ultimaNotificacion48() As String

    ultimaNotificacion48 = "ultima_notificacion_48"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Asocia ultimaNotificacion48 completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get ultimaNotificacion48_t() As String

    ultimaNotificacion48_t = table & ".ultima_notificacion_48"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Asocia ultimaNotificacionCorte
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get ultimaNotificacionCorte() As String

    ultimaNotificacionCorte = "ultima_notificacion_corte"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Asocia ultimaNotificacionCorte completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get ultimaNotificacionCorte_t() As String

    ultimaNotificacionCorte_t = table & ".ultima_notificacion_corte"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Asocia autoID
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get autoID() As String

    autoID = "auto_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Asocia autoID completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get autoID_t() As String

    autoID_t = table & ".auto_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Asocia created
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get created() As String

    created = "created"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Asocia created completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get created_t() As String

    created_t = table & ".created"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Asocia updated
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get updated() As String

    updated = "updated"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Asocia updated completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get updated_t() As String

    updated_t = table & ".updated"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Asocia uuid
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get uuid() As String

    uuid = "uuid"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 30/06/2018
' Description: Asocia uuid completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get uuid_t() As String

    uuid_t = table & ".uuid"

End Property

