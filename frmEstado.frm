VERSION 5.00
Begin VB.Form frmEstado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estado"
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
   Begin VB.ListBox lstEstados 
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
      Caption         =   "Estado"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblHotel 
      AutoSize        =   -1  'True
      Caption         =   "Estados"
      Height          =   195
      Left            =   6000
      TabIndex        =   6
      Top             =   120
      Width           =   570
   End
End
Attribute VB_Name = "frmEstado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private estado As New clsMyAEstado

Private Sub fillForm()

    With estado
        Me.txtCodigo.Text = .estadoID
        Me.txtDescripcion.Text = .nombre
    End With
    
End Sub

Private Sub cmdGrabar_Click()
 
    If Me.txtCodigo.Text = "0" Then Exit Sub
    If Trim(Me.txtDescripcion.Text) = "" Then Exit Sub
    
    With estado
        .nombre = Me.txtDescripcion.Text
        .save dbapp
        
        .fillList Me.lstEstados, dbapp
    End With
    
    cmdLimpiar_Click

End Sub

Private Sub cmdLimpiar_Click()
    
    estado.newID True, dbapp
    
    fillForm

End Sub

Private Sub cmdSalir_Click()
    
    Unload Me

End Sub

Private Sub Form_Load()

    estado.fillList Me.lstEstados, dbapp
    
    cmdLimpiar_Click
    
End Sub

Private Sub lstestados_Click()

    If Me.lstEstados.ListIndex < 0 Then Exit Sub
    
    estado.estadoID = Me.lstEstados.ItemData(Me.lstEstados.ListIndex)
    estado.findByPrimaryKey dbapp
    
    fillForm
    
End Sub

Private Sub txtDescripcion_GotFocus()

    marcarseleccion Me.txtDescripcion
    
End Sub


