VERSION 5.00
Begin VB.Form frmDestino 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Destino del Servicio"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9855
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   9855
   Begin VB.TextBox txtCodigo 
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "Limpiar"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   5535
   End
   Begin VB.ListBox lstDestinos 
      Height          =   1425
      Left            =   6000
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label lblCodigo 
      AutoSize        =   -1  'True
      Caption         =   "Código"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Destino"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   540
   End
   Begin VB.Label lblHotel 
      AutoSize        =   -1  'True
      Caption         =   "Destinos del Servicio"
      Height          =   195
      Left            =   6000
      TabIndex        =   6
      Top             =   120
      Width           =   1485
   End
End
Attribute VB_Name = "frmDestino"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private destino As New clsMyADestinoServ

Private Sub fillForm()

    With destino
        Me.txtCodigo.Text = .destinoID
        Me.txtDescripcion.Text = .nombre
    End With
    
End Sub

Private Sub cmdGrabar_Click()
 
    If Me.txtCodigo.Text = "0" Then Exit Sub
    If Trim(Me.txtDescripcion.Text) = "" Then Exit Sub
    
    With destino
        .nombre = Me.txtDescripcion.Text
        .save dbapp
        
        .fillList Me.lstDestinos, dbapp
    End With
    
    cmdLimpiar_Click

End Sub

Private Sub cmdLimpiar_Click()
    
    destino.newID True, dbapp
    
    fillForm

End Sub

Private Sub cmdSalir_Click()
    
    Unload Me

End Sub

Private Sub Form_Load()

    destino.fillList Me.lstDestinos, dbapp
    
    cmdLimpiar_Click
    
End Sub

Private Sub lstDestinos_Click()

    If Me.lstDestinos.ListIndex < 0 Then Exit Sub
    
    destino.destinoID = Me.lstDestinos.ItemData(Me.lstDestinos.ListIndex)
    destino.findByPrimaryKey dbapp
    
    fillForm
    
End Sub

Private Sub txtDescripcion_GotFocus()

    marcarseleccion Me.txtDescripcion
    
End Sub


