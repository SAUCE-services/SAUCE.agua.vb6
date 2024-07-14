VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPFRecord6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vRecordCode As Integer
Private vBarCode As String
Private vTypeCode As String

Private Sub class_Initialize()

    initializeOther

End Sub

Private Sub initializeOther()

    vRecordCode = 0
    vBarCode = ""
    vTypeCode = ""

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

Public Property Get barCode() As Variant

    barCode = vBarCode

End Property

Public Property Let barCode(ByVal vNewValue As Variant)

On Error Resume Next

    vBarCode = vNewValue

End Property

Public Property Get typeCode() As Variant

    typeCode = vTypeCode

End Property

Public Property Let typeCode(ByVal vNewValue As Variant)

On Error Resume Next

    vTypeCode = vNewValue

End Property

Public Function clone() As clsPFRecord6
Dim objObjeto As New clsPFRecord6

    With objObjeto
        .recordCode = vRecordCode
        .barCode = vBarCode
        .typeCode = vTypeCode
    End With

    Set clone = objObjeto

End Function

Public Sub fillByLine(pLine As String)

    With Me
        .recordCode = Mid(pLine, 1, 1)
        .barCode = Trim(Mid(pLine, 2, 80))
        .typeCode = Trim(Mid(pLine, 82, 1))
    End With
    
End Sub


