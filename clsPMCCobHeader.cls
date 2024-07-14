VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPMCCobHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vCodigoRegistro As Integer
Private vCodigoBanelco As Integer
Private vCodigoEmpresa As Integer
Private vFechaArchivo As Date
Private vRegistro As String

Private Sub class_Initialize()

    initializeOther

End Sub

Private Sub initializeOther()

    vCodigoRegistro = 0
    vCodigoBanelco = 0
    vCodigoEmpresa = 0
    vFechaArchivo = Date
    vRegistro = ""

End Sub

Public Sub clean()

    Call class_Initialize

End Sub

Public Property Get codigoRegistro() As Variant

    codigoRegistro = vCodigoRegistro

End Property

Public Property Let codigoRegistro(ByVal vNewValue As Variant)

On Error Resume Next

    vCodigoRegistro = vNewValue

End Property

Public Property Get codigoBanelco() As Variant

    codigoBanelco = vCodigoBanelco

End Property

Public Property Let codigoBanelco(ByVal vNewValue As Variant)

On Error Resume Next

    vCodigoBanelco = vNewValue

End Property

Public Property Get codigoEmpresa() As Variant

    codigoEmpresa = vCodigoEmpresa

End Property

Public Property Let codigoEmpresa(ByVal vNewValue As Variant)

On Error Resume Next

    vCodigoEmpresa = vNewValue

End Property

Public Property Get fechaArchivo() As Variant

    fechaArchivo = vFechaArchivo

End Property

Public Property Let fechaArchivo(ByVal vNewValue As Variant)

On Error Resume Next

    vFechaArchivo = vNewValue

End Property

Public Function clone()
Dim objObjeto As New clsPMCCobHeader

    With objObjeto
        .codigoRegistro = vCodigoRegistro
        .codigoBanelco = vCodigoBanelco
        .codigoEmpresa = vCodigoEmpresa
        .fechaArchivo = vFechaArchivo
    End With

    Set clone = objObjeto

End Function

Public Property Get registro() As Variant

    registro = vRegistro

End Property

Public Property Let registro(ByVal vNewValue As Variant)

    vRegistro = vNewValue
    
    parseRegistro

End Property

Private Sub parseRegistro()

    With Me
        .codigoRegistro = Mid(.registro, 1, 1)
        .codigoBanelco = Mid(.registro, 2, 3)
        .codigoEmpresa = Mid(.registro, 5, 4)
        .fechaArchivo = modConv.stringAAAAMMDD2Date(Mid(.registro, 9, 8))
    End With
    
End Sub


