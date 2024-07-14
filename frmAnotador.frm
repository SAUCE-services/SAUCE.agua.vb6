VERSION 5.00
Begin VB.Form frmAnotador 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anotador"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10110
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   10110
   Begin VB.TextBox txtCliente 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   9615
   End
   Begin VB.TextBox txtAnotacion 
      Height          =   975
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   720
      Width           =   7695
   End
   Begin VB.TextBox txtAnotador 
      Height          =   4695
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1800
      Width           =   9615
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Agregar"
      Height          =   375
      Left            =   8160
      TabIndex        =   2
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8160
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Cliente"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmAnotador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cliente As New clsMODCliente

Private Sub fillDetalle()
Dim historia As String

Dim anotador As New clsMyAAnotador

On Error Resume Next

    Me.txtAnotador.Text = ""
    Me.txtAnotacion.Text = ""
    
    If cliente.clienteId = 0 Then Exit Sub
    
    Me.MousePointer = 11
    
    historia = ""
    
    For Each anotador In anotador.collectionByClienteID(cliente.clienteId, dbapp)
        historia = historia & anotador.created & " -> " & anotador.anotacion & vbCrLf
        historia = historia & vbCrLf
    Next
    
    Me.txtAnotador.Text = historia
    
    Me.MousePointer = 0

End Sub

Private Sub cmdAgregar_Click()
Dim anotador As New clsMyAAnotador

    If cliente.clienteId = 0 Then Exit Sub
    If Trim(Me.txtAnotacion.Text) = "" Then Exit Sub
    
    anotador.clienteId = cliente.clienteId
    anotador.anotacion = Me.txtAnotacion.Text
    anotador.add dbapp
    
    fillDetalle
    
End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub txtCliente_GotFocus()

    marcarseleccion Me.txtCliente
    
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
Dim clienteRep As New clsREPCliente

    Set cliente = clienteRep.formSearch(frmBuscarRest, KeyAscii, "Clientes")
    
    Me.txtCliente.Text = cliente.textFound
    KeyAscii = 0
    
    fillDetalle

End Sub

