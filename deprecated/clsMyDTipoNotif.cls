VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0   'vbNone
  MTSTransactionMode  = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsMyDTipoNotif"
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

    table = "tipo_notificacion"

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
' Description: Asocia nombre
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get nombre() As String

    nombre = "nombre"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Asocia nombre completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get nombre_t() As String

    nombre_t = table & ".nombre"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Asocia valorSocio
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get valorSocio() As String

    valorSocio = "valor_socio"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Asocia valorSocio completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get valorSocio_t() As String

    valorSocio_t = table & ".valor_socio"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Asocia valorNoSocio
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get valorNoSocio() As String

    valorNoSocio = "valor_no_socio"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/07/2018
' Description: Asocia valorNoSocio completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get valorNoSocio_t() As String

    valorNoSocio_t = table & ".valor_no_socio"

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

