VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRGB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vRed As Integer
Private vGreen As Integer
Private vBlue As Integer
Private vColor As Variant

Private Sub class_Initialize()

    initializeOther

End Sub

Private Sub initializeOther()

    vRed = 0
    vGreen = 0
    vBlue = 0
    vColor = 0

End Sub

Public Sub clean()

    Call class_Initialize

End Sub

Public Property Get red() As Variant

    red = vRed

End Property

Public Property Get green() As Variant

    green = vGreen

End Property

Public Property Get blue() As Variant

    blue = vBlue

End Property

Public Property Let color(ByVal vNewValue As Variant)

On Error Resume Next

    vColor = vNewValue
    
    convert

End Property

Public Function clone()
Dim objObjeto As New clsRGB

    With objObjeto
        .red = vRed
        .green = vGreen
        .blue = vBlue
    End With

    Set clone = objObjeto

End Function

Private Sub convert()
Dim strHex As String

    strHex = Right(String(6, "0") & Hex(vColor), 6)
    
    vRed = Val("&H" & Mid(strHex, 5, 2))
    vGreen = Val("&H" & Mid(strHex, 3, 2))
    vBlue = Val("&H" & Mid(strHex, 1, 2))
    
End Sub
