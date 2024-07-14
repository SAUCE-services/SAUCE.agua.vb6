VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0   'vbNone
  MTSTransactionMode  = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsMyDClienteVolumen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Asocia TABLA
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get table() As String

    table = "cliente_volumen"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Asocia clienteID
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get clienteID() As String

    clienteID = "cliente_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Asocia clienteID completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get clienteID_t() As String

    clienteID_t = table & ".cliente_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Asocia periodoID
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get periodoID() As String

    periodoID = "periodo_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Asocia periodoID completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get periodoID_t() As String

    periodoID_t = table & ".periodo_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Asocia medidorIDActual
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get medidorIDActual() As String

    medidorIDActual = "medidor_id_actual"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Asocia medidorIDActual completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get medidorIDActual_t() As String

    medidorIDActual_t = table & ".medidor_id_actual"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Asocia estadoActual
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get estadoActual() As String

    estadoActual = "estado_actual"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Asocia estadoActual completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get estadoActual_t() As String

    estadoActual_t = table & ".estado_actual"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Asocia medidorIDAnterior
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get medidorIDAnterior() As String

    medidorIDAnterior = "medidor_id_anterior"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Asocia medidorIDAnterior completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get medidorIDAnterior_t() As String

    medidorIDAnterior_t = table & ".medidor_id_anterior"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Asocia estadoAnterior
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get estadoAnterior() As String

    estadoAnterior = "estado_anterior"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Asocia estadoAnterior completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get estadoAnterior_t() As String

    estadoAnterior_t = table & ".estado_anterior"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Asocia consumido
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get consumido() As String

    consumido = "consumido"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Asocia consumido completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get consumido_t() As String

    consumido_t = table & ".consumido"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Asocia autoID
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get autoID() As String

    autoID = "auto_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Asocia autoID completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get autoID_t() As String

    autoID_t = table & ".auto_id"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Asocia created
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get created() As String

    created = "created"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 15/04/2018
' Description: Asocia created completo
' Log: 
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get created_t() As String

    created_t = table & ".created"

End Property

