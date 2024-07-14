VERSION 5.00
Begin VB.Form frmBuscarRest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Búsqueda"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstEncontrados 
      Height          =   2010
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   5535
   End
   Begin VB.TextBox txtCadena 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   2520
      Width           =   1695
   End
End
Attribute VB_Name = "frmBuscarRest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private vModel As Variant
Private vRepository As Variant

Private vColeccion As Collection

Private Sub cmdSalir_Click()

    Me.Hide
    
End Sub

Public Property Get model() As Variant

    If IsEmpty(vModel) Then Set vModel = Nothing

    Set model = vModel
    
End Property

Public Property Let repository(vNewValue As Variant)

    Set vRepository = vNewValue
    
End Property

Private Sub Form_Activate()

    Me.txtCadena.SetFocus
    Me.txtCadena.SelStart = Len(Me.txtCadena.Text)
    
End Sub

Private Sub lstEncontrados_DblClick()

    If Me.lstEncontrados.ListIndex < 0 Then Exit Sub
    
    Set vModel = vRepository.findSearch(Me.lstEncontrados.ItemData(Me.lstEncontrados.ListIndex))
    
    cmdSalir_Click

End Sub

Private Sub lstEncontrados_KeyPress(KeyAscii As Integer)

    If Me.lstEncontrados.ListIndex < 0 Then Exit Sub
    If KeyAscii <> 13 Then Exit Sub
    
    Set vModel = vRepository.findSearch(Me.lstEncontrados.ItemData(Me.lstEncontrados.ListIndex))
    
    cmdSalir_Click

End Sub

Private Sub txtCadena_Change()
Dim local_ As Variant

    Set vColeccion = vRepository.collectionSearch(Me.txtCadena.Text)
    
    Me.lstEncontrados.Clear
    
    For Each local_ In vColeccion
        Me.lstEncontrados.AddItem local_.textFound
        Me.lstEncontrados.ItemData(Me.lstEncontrados.NewIndex) = local_.keyFound
    Next
    
End Sub
