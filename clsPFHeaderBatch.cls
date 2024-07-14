VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPFHeaderBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vRecordCode As Integer
Private vCreateDate As Date
Private vBatchNumber As Long
Private vDescription As String

Private Sub class_Initialize()

    initializeOther

End Sub

Private Sub initializeOther()

    vRecordCode = 0
    vCreateDate = Date
    vBatchNumber = 0
    vDescription = ""

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

Public Property Get description() As Variant

    description = vDescription

End Property

Public Property Let description(ByVal vNewValue As Variant)

On Error Resume Next

    vDescription = vNewValue

End Property

Public Function clone() As clsPFHeaderBatch
Dim objObjeto As New clsPFHeaderBatch

    With objObjeto
        .recordCode = vRecordCode
        .createDate = vCreateDate
        .batchNumber = vBatchNumber
        .description = vDescription
    End With

    Set clone = objObjeto

End Function

Public Sub fillByLine(pLine As String)

    With Me
        .recordCode = Mid(pLine, 1, 1)
        .createDate = modConv.stringAAAAMMDD2Date(Mid(pLine, 2, 8))
        .batchNumber = Mid(pLine, 10, 6)
        .description = Trim(Mid(pLine, 16, 35))
    End With
    
End Sub

