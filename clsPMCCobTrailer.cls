VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPMCCobTrailer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vCodigoRegistro As Integer
Private vCodigoBanelco As Integer
Private vCodigoEmpresa As Integer
Private vFechaArchivo As Date
Private vCantidadRegistrosPesos As Double
Private vCantidadRegistrosDolares As Double
Private vTotalImportePesos As Double
Private vTotalImportesDolares As Double
Private vRegistro As String

Private Sub class_Initialize()

    initializeOther

End Sub

Private Sub initializeOther()

    vCodigoRegistro = 0
    vCodigoBanelco = 0
    vCodigoEmpresa = 0
    vFechaArchivo = Date
    vCantidadRegistrosPesos = 0
    vCantidadRegistrosDolares = 0
    vTotalImportePesos = 0
    vTotalImportesDolares = 0
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

Public Property Get cantidadRegistrosPesos() As Variant

    cantidadRegistrosPesos = vCantidadRegistrosPesos

End Property

Public Property Let cantidadRegistrosPesos(ByVal vNewValue As Variant)

On Error Resume Next

    vCantidadRegistrosPesos = vNewValue

End Property

Public Property Get cantidadRegistrosDolares() As Variant

    cantidadRegistrosDolares = vCantidadRegistrosDolares

End Property

Public Property Let cantidadRegistrosDolares(ByVal vNewValue As Variant)

On Error Resume Next

    vCantidadRegistrosDolares = vNewValue

End Property

Public Property Get totalImportePesos() As Variant

    totalImportePesos = vTotalImportePesos

End Property

Public Property Let totalImportePesos(ByVal vNewValue As Variant)

On Error Resume Next

    vTotalImportePesos = vNewValue

End Property

Public Property Get totalImportesDolares() As Variant

    totalImportesDolares = vTotalImportesDolares

End Property

Public Property Let totalImportesDolares(ByVal vNewValue As Variant)

On Error Resume Next

    vTotalImportesDolares = vNewValue

End Property

Public Function clone()
Dim objObjeto As New clsPMCCobTrailer

    With objObjeto
        .codigoRegistro = vCodigoRegistro
        .codigoBanelco = vCodigoBanelco
        .codigoEmpresa = vCodigoEmpresa
        .fechaArchivo = vFechaArchivo
        .cantidadRegistrosPesos = vCantidadRegistrosPesos
        .cantidadRegistrosDolares = vCantidadRegistrosDolares
        .totalImportePesos = vTotalImportePesos
        .totalImportesDolares = vTotalImportesDolares
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
        .cantidadRegistrosPesos = Mid(.registro, 17, 7)
        .cantidadRegistrosDolares = Mid(.registro, 24, 7)
        .totalImportePesos = modConv.parseDouble(Mid(.registro, 31, 11), 2)
        .totalImportesDolares = modConv.parseDouble(Mid(.registro, 42, 11), 2)
    End With
    
End Sub


