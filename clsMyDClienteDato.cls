VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0   'vbNone
  MTSTransactionMode  = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsMyDClienteDato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "clientedato"

End Property

Public Property Get clienteID() As String

    clienteID = "cliente_id"

End Property

Public Property Get clienteID_t() As String

    clienteID_t = "clientedato.cliente_id"

End Property

Public Property Get documento() As String

    documento = "documento"

End Property

Public Property Get documento_t() As String

    documento_t = "clientedato.documento"

End Property

Public Property Get email() As String

    email = "email"

End Property

Public Property Get email_t() As String

    email_t = "clientedato.email"

End Property

Public Property Get fijo() As String

    fijo = "fijo"

End Property

Public Property Get fijo_t() As String

    fijo_t = "clientedato.fijo"

End Property

Public Property Get celular() As String

    celular = "celular"

End Property

Public Property Get celular_t() As String

    celular_t = "clientedato.celular"

End Property

Public Property Get uid() As String

    uid = "uid"

End Property

Public Property Get uid_t() As String

    uid_t = "clientedato.uid"

End Property

Public Property Get created() As String

    created = "created"

End Property

Public Property Get created_t() As String

    created_t = "clientedato.created"

End Property

