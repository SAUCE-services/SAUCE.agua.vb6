VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0   'vbNone
  MTSTransactionMode  = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsPFFooterFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vRecordCode As Integer
Private vCreateDate As Date
Private vTotalBatches As Long
Private vFilePaymentCount As Long
Private vFilePaymentAmount As Currency
Private vFileCount As Long

Private Sub class_Initialize()

    initializeOther

End Sub

Private Sub initializeOther()

    vRecordCode = 0
    vCreateDate = Date
    vTotalBatches = 0
    vFilePaymentCount = 0
    vFilePaymentAmount = 0
    vFileCount = 0

End Sub

Public Sub clean()

    Call class_Initialize

End Sub

Public Property Get recordCode() As Variant

    recordCode = vRecordCode

End Property

Public Property Let recordCode(ByVal vNewValue As Variant)

On Error Resume Next

    vRecordCode = vNewValue

End Property

Public Property Get createDate() As Variant

    createDate = vCreateDate

End Property

Public Property Let createDate(ByVal vNewValue As Variant)

On Error Resume Next

    vCreateDate = vNewValue

End Property

Public Property Get totalBatches() As Variant

    totalBatches = vTotalBatches

End Property

Public Property Let totalBatches(ByVal vNewValue As Variant)

On Error Resume Next

    vTotalBatches = vNewValue

End Property

Public Property Get filePaymentCount() As Variant

    filePaymentCount = vFilePaymentCount

End Property

Public Property Let filePaymentCount(ByVal vNewValue As Variant)

On Error Resume Next

    vFilePaymentCount = vNewValue

End Property

Public Property Get filePaymentAmount() As Variant

    filePaymentAmount = vFilePaymentAmount

End Property

Public Property Let filePaymentAmount(ByVal vNewValue As Variant)

On Error Resume Next

    vFilePaymentAmount = vNewValue

End Property

Public Property Get fileCount() As Variant

    fileCount = vFileCount

End Property

Public Property Let fileCount(ByVal vNewValue As Variant)

On Error Resume Next

    vFileCount = vNewValue

End Property

Public Function clone() As clsPFFooterFile
Dim objObjeto As New clsPFFooterFile

    With objObjeto
        .recordCode = vRecordCode
        .createDate = vCreateDate
        .totalBatches = vTotalBatches
        .filePaymentCount = vFilePaymentCount
        .filePaymentAmount = vFilePaymentAmount
        .fileCount = vFileCount
    End With

    Set clone = objObjeto

End Function

