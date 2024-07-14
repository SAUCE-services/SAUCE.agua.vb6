VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPFRecord7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vRecordCode As Integer
Private vCurrencyCode As String
Private vPayInstrument As String
Private vCodeBarPayInst As String
Private vAmount As Currency

Private Sub class_Initialize()

    initializeOther

End Sub

Private Sub initializeOther()

    vRecordCode = 0
    vCurrencyCode = ""
    vPayInstrument = ""
    vCodeBarPayInst = ""
    vAmount = 0

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

Public Property Get currencyCode() As Variant

    currencyCode = vCurrencyCode

End Property

Public Property Let currencyCode(ByVal vNewValue As Variant)

On Error Resume Next

    vCurrencyCode = vNewValue

End Property

Public Property Get payInstrument() As Variant

    payInstrument = vPayInstrument

End Property

Public Property Let payInstrument(ByVal vNewValue As Variant)

On Error Resume Next

    vPayInstrument = vNewValue

End Property

Public Property Get codeBarPayInst() As Variant

    codeBarPayInst = vCodeBarPayInst

End Property

Public Property Let codeBarPayInst(ByVal vNewValue As Variant)

On Error Resume Next

    vCodeBarPayInst = vNewValue

End Property

Public Property Get amount() As Variant

    amount = vAmount

End Property

Public Property Let amount(ByVal vNewValue As Variant)

On Error Resume Next

    vAmount = vNewValue

End Property

Public Function clone() As clsPFRecord7
Dim objObjeto As New clsPFRecord7

    With objObjeto
        .recordCode = vRecordCode
        .currencyCode = vCurrencyCode
        .payInstrument = vPayInstrument
        .codeBarPayInst = vCodeBarPayInst
        .amount = vAmount
    End With

    Set clone = objObjeto

End Function

Public Sub fillByLine(pLine As String)

    With Me
        .recordCode = Mid(pLine, 1, 1)
        .currencyCode = Trim(Mid(pLine, 2, 3))
        .payInstrument = Trim(Mid(pLine, 5, 1))
        .codeBarPayInst = Trim(Mid(pLine, 6, 80))
        .amount = modConv.parseCurrency(Mid(pLine, 86, 15), 2)
    End With
    
End Sub



