VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPFRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public PFRec5 As clsPFRecord5
Public PFRec6 As clsPFRecord6
Public PFRec7 As clsPFRecord7

Private Sub class_Initialize()

    initializeOther

End Sub

Private Sub initializeOther()

    Set PFRec5 = Nothing
    Set PFRec6 = Nothing
    Set PFRec7 = Nothing
    
    Set PFRec5 = New clsPFRecord5
    Set PFRec6 = New clsPFRecord6
    Set PFRec7 = New clsPFRecord7

End Sub

Public Sub clean()

    Call class_Initialize

End Sub

