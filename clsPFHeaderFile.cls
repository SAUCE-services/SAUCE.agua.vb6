VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPFHeaderFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vRecordCode As Integer
Private vCreateDate As Date
Private vOriginName As String
Private vClientNumber As Currency
Private vClientName As String

Private Sub class_Initialize()

    initializeOther

End Sub

Private Sub initializeOther()

    vRecordCode = 0
    vCreateDate = Date
    vOriginName = ""
    vClientNumber = 0
    vClientName = ""

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

Public Property Get originName() As Variant

    originName = vOriginName

End Property

Public Property Let originName(ByVal vNewValue As Variant)

On Error Resume Next

    vOriginName = vNewValue

End Property

Public Property Get clientNumber() As Variant

    clientNumber = vClientNumber

End Property

Public Property Let clientNumber(ByVal vNewValue As Variant)

On Error Resume Next

    vClientNumber = vNewValue

End Property

Public Property Get clientName() As Variant

    clientName = vClientName

End Property

Public Property Let clientName(ByVal vNewValue As Variant)

On Error Resume Next

    vClientName = vNewValue

End Property

Public Function clone() As clsPFHeaderFile
Dim objObjeto As New clsPFHeaderFile

    With objObjeto
        .recordCode = vRecordCode
        .createDate = vCreateDate
        .originName = vOriginName
        .clientNumber = vClientNumber
        .clientName = vClientName
    End With

    Set clone = objObjeto

End Function

Public Sub fillByLine(pLine As String)

    With Me
        .recordCode = Mid(pLine, 1, 1)
        .createDate = modConv.stringAAAAMMDD2Date(Mid(pLine, 2, 8))
        .originName = Trim(Mid(pLine, 10, 25))
        .clientNumber = Mid(pLine, 35, 9)
        .clientName = Trim(Mid(pLine, 44, 35))
    End With
    
End Sub
