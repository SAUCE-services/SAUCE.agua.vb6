VERSION 5.00
Begin VB.Form frmCategoria 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Categoría de Socio"
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
   Begin VB.ListBox lstCategorias 
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
      Caption         =   "Categoría"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   705
   End
   Begin VB.Label lblHotel 
      AutoSize        =   -1  'True
      Caption         =   "Categorías del Socio"
      Height          =   195
      Left            =   6000
      TabIndex        =   6
      Top             =   120
      Width           =   1485
   End
End
Attribute VB_Name = "frmCategoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private categoria As New clsMyACategoriaSocio

Private Sub fillForm()

    With categoria
        Me.txtCodigo.Text = .categoriasocioID
        Me.txtDescripcion.Text = .nombre
    End With
    
End Sub

Private Sub cmdGrabar_Click()
 
    If Me.txtCodigo.Text = "0" Then Exit Sub
    If Trim(Me.txtDescripcion.Text) = "" Then Exit Sub
    
    With categoria
        .nombre = Me.txtDescripcion.Text
        .save dbapp
        
        .fillList Me.lstCategorias, dbapp
    End With
    
    cmdLimpiar_Click

End Sub

Private Sub cmdLimpiar_Click()
    
    categoria.newID True, dbapp
    
    fillForm

End Sub

Private Sub cmdSalir_Click()
    
    Unload Me

End Sub

Private Sub Form_Load()

    categoria.fillList Me.lstCategorias, dbapp
    
    cmdLimpiar_Click
    
End Sub

Private Sub lstCategorias_Click()

    If Me.lstCategorias.ListIndex < 0 Then Exit Sub
    
    categoria.categoriasocioID = Me.lstCategorias.ItemData(Me.lstCategorias.ListIndex)
    categoria.findByPrimaryKey dbapp
    
    fillForm
    
End Sub

Private Sub txtDescripcion_GotFocus()

    marcarseleccion Me.txtDescripcion
    
End Sub


