VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0   'vbNone
  MTSTransactionMode  = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsMyDNotificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Asocia TABLA
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get table() As String

    table = "notificacion"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Asocia clienteID
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get clienteID() As String

    clienteID = "cliente_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Asocia clienteID completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get clienteID_t() As String

    clienteID_t = table & ".cliente_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Asocia fecha
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fecha() As String

    fecha = "fecha"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Asocia fecha completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fecha_t() As String

    fecha_t = table & ".fecha"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Asocia tiponotificacionID
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get tiponotificacionID() As String

    tiponotificacionID = "tiponotificacion_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Asocia tiponotificacionID completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get tiponotificacionID_t() As String

    tiponotificacionID_t = table & ".tiponotificacion_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Asocia vencimiento
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get vencimiento() As String

    vencimiento = "vencimiento"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Asocia vencimiento completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get vencimiento_t() As String

    vencimiento_t = table & ".vencimiento"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Asocia notificacionID
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get notificacionID() As String

    notificacionID = "notificacion_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Asocia notificacionID completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get notificacionID_t() As String

    notificacionID_t = table & ".notificacion_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Asocia created
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get created() As String

    created = "created"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Asocia created completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get created_t() As String

    created_t = table & ".created"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Asocia updated
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get updated() As String

    updated = "updated"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Asocia updated completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get updated_t() As String

    updated_t = table & ".updated"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Asocia uuid
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get uuid() As String

    uuid = "uuid"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Asocia uuid completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get uuid_t() As String

    uuid_t = table & ".uuid"

End Property

