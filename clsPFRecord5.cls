VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPFRecord5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vRecordCode As Integer
Private vRecordSequence As Long
Private vTransactionCode As Integer
Private vWorkDate As Date
Private vTransferDate As Date
Private vAccountNumber As String
Private vCurrencyCode As String
Private vAmount As Currency
Private vTerminalID As String
Private vPaymentDate As Date
Private vPaymentTime As Date
Private vTermSeqNumber As Integer

Private Sub class_Initialize()

    initializeOther

End Sub

Private Sub initializeOther()

    vRecordCode = 0
    vRecordSequence = 0
    vTransactionCode = 0
    vWorkDate = Date
    vTransferDate = Date
    vAccountNumber = ""
    vCurrencyCode = ""
    vAmount = 0
    vTerminalID = ""
    vPaymentDate = Date
    vPaymentTime = Date
    vTermSeqNumber = 0

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

Public Property Get recordSequence() As Variant

    recordSequence = vRecordSequence

End Property

Public Property Let recordSequence(ByVal vNewValue As Variant)

On Error Resume Next

    vRecordSequence = vNewValue

End Property

Public Property Get transactionCode() As Variant

    transactionCode = vTransactionCode

End Property

Public Property Let transactionCode(ByVal vNewValue As Variant)

On Error Resume Next

    vTransactionCode = vNewValue

End Property

Public Property Get workDate() As Variant

    workDate = vWorkDate

End Property

Public Property Let workDate(ByVal vNewValue As Variant)

On Error Resume Next

    vWorkDate = vNewValue

End Property

Public Property Get transferDate() As Variant

    transferDate = vTransferDate

End Property

Public Property Let transferDate(ByVal vNewValue As Variant)

On Error Resume Next

    vTransferDate = vNewValue

End Property

Public Property Get accountNumber() As Variant

    accountNumber = vAccountNumber

End Property

Public Property Let accountNumber(ByVal vNewValue As Variant)

On Error Resume Next

    vAccountNumber = vNewValue

End Property

Public Property Get currencyCode() As Variant

    currencyCode = vCurrencyCode

End Property

Public Property Let currencyCode(ByVal vNewValue As Variant)

On Error Resume Next

    vCurrencyCode = vNewValue

End Property

Public Property Get amount() As Variant

    amount = vAmount

End Property

Public Property Let amount(ByVal vNewValue As Variant)

On Error Resume Next

    vAmount = vNewValue

End Property

Public Property Get terminalID() As Variant

    terminalID = vTerminalID

End Property

Public Property Let terminalID(ByVal vNewValue As Variant)

On Error Resume Next

    vTerminalID = vNewValue

End Property

Public Property Get paymentDate() As Variant

    paymentDate = vPaymentDate

End Property

Public Property Let paymentDate(ByVal vNewValue As Variant)

On Error Resume Next

    vPaymentDate = vNewValue

End Property

Public Property Get paymentTime() As Variant

    paymentTime = vPaymentTime

End Property

Public Property Let paymentTime(ByVal vNewValue As Variant)

On Error Resume Next

    vPaymentTime = vNewValue

End Property

Public Property Get termSeqNumber() As Variant

    termSeqNumber = vTermSeqNumber

End Property

Public Property Let termSeqNumber(ByVal vNewValue As Variant)

On Error Resume Next

    vTermSeqNumber = vNewValue

End Property

Public Function clone() As clsPFRecord5
Dim objObjeto As New clsPFRecord5

    With objObjeto
        .recordCode = vRecordCode
        .recordSequence = vRecordSequence
        .transactionCode = vTransactionCode
        .workDate = vWorkDate
        .transferDate = vTransferDate
        .accountNumber = vAccountNumber
        .currencyCode = vCurrencyCode
        .amount = vAmount
        .terminalID = vTerminalID
        .paymentDate = vPaymentDate
        .paymentTime = vPaymentTime
        .termSeqNumber = vTermSeqNumber
    End With

    Set clone = objObjeto

End Function

Public Sub fillByLine(pLine As String)

    With Me
        .recordCode = Mid(pLine, 1, 1)
        .recordSequence = Mid(pLine, 2, 5)
        .transactionCode = Mid(pLine, 7, 2)
        .workDate = modConv.stringAAAAMMDD2Date(Mid(pLine, 9, 8))
        .transferDate = modConv.stringAAAAMMDD2Date(Mid(pLine, 17, 8))
        .accountNumber = Trim(Mid(pLine, 25, 21))
        .currencyCode = Trim(Mid(pLine, 46, 3))
        .amount = modConv.parseDouble(Mid(pLine, 49, 10), 2)
        .terminalID = Trim(Mid(pLine, 59, 6))
        .paymentDate = modConv.stringAAAAMMDD2Date(Mid(pLine, 65, 8))
        .paymentTime = modConv.stringHHMM2Time(Mid(pLine, 73, 4))
        .termSeqNumber = Mid(pLine, 77, 4)
    End With
    
End Sub


