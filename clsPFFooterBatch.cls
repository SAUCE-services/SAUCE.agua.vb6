VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0   'vbNone
  MTSTransactionMode  = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsPFFooterBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vRecordCode As Integer
Private vCreateDate As Date
Private vBatchNumber As Long
Private vBatchPaymentCount As Long
Private vBatchPaymentAmount As Currency
Private vBatchCount As Long

Private Sub class_Initialize()

    initializeOther

End Sub

Private Sub initializeOther()

    vRecordCode = 0
    vCreateDate = Date
    vBatchNumber = 0
    vBatchPaymentCount = 0
    vBatchPaymentAmount = 0
    vBatchCount = 0

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

Public Property Get batchNumber() As Variant

    batchNumber = vBatchNumber

End Property

Public Property Let batchNumber(ByVal vNewValue As Variant)

On Error Resume Next

    vBatchNumber = vNewValue

End Property

Public Property Get batchPaymentCount() As Variant

    batchPaymentCount = vBatchPaymentCount

End Property

Public Property Let batchPaymentCount(ByVal vNewValue As Variant)

On Error Resume Next

    vBatchPaymentCount = vNewValue

End Property

Public Property Get batchPaymentAmount() As Variant

    batchPaymentAmount = vBatchPaymentAmount

End Property

Public Property Let batchPaymentAmount(ByVal vNewValue As Variant)

On Error Resume Next

    vBatchPaymentAmount = vNewValue

End Property

Public Property Get batchCount() As Variant

    batchCount = vBatchCount

End Property

Public Property Let batchCount(ByVal vNewValue As Variant)

On Error Resume Next

    vBatchCount = vNewValue

End Property

Public Function clone() As clsPFFooterBatch
Dim objObjeto As New clsPFFooterBatch

    With objObjeto
        .recordCode = vRecordCode
        .createDate = vCreateDate
        .batchNumber = vBatchNumber
        .batchPaymentCount = vBatchPaymentCount
        .batchPaymentAmount = vBatchPaymentAmount
        .batchCount = vBatchCount
    End With

    Set clone = objObjeto

End Function

