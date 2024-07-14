VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyDPFFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "pffile"

End Property

Public Property Get fileName() As String

    fileName = "file_name"

End Property

Public Property Get fileName_t() As String

    fileName_t = "pffile.file_name"

End Property

Public Property Get path() As String

    path = "path"

End Property

Public Property Get path_t() As String

    path_t = "pffile.path"

End Property

Public Property Get import() As String

    import = "import"

End Property

Public Property Get import_t() As String

    import_t = "pffile.import"

End Property

Public Property Get fechaImport() As String

    fechaImport = "fecha_import"

End Property

Public Property Get fechaImport_t() As String

    fechaImport_t = "pffile.fecha_import"

End Property

Public Property Get autoID() As String

    autoID = "auto_id"

End Property

Public Property Get autoID_t() As String

    autoID_t = "pffile.auto_id"

End Property

Public Property Get uid() As String

    uid = "uid"

End Property

Public Property Get uid_t() As String

    uid_t = "pffile.uid"

End Property

Public Property Get created() As String

    created = "created"

End Property

Public Property Get created_t() As String

    created_t = "pffile.created"

End Property

