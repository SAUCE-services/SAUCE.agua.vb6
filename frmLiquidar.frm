VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLiquidar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Liquidación General"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9855
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   9855
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   6000
      TabIndex        =   18
      ToolTipText     =   "Elimina TODAS las Liquidaciones del PERIODO Seleccionado"
      Top             =   2040
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid grdClientes 
      Height          =   5295
      Left            =   240
      TabIndex        =   17
      Top             =   2520
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   9340
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar stbEstado 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   8115
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdExtremos 
      Caption         =   "Extremos"
      Height          =   375
      Left            =   6000
      TabIndex        =   15
      ToolTipText     =   "Factura e Imprime todas las CONEXIONES pendientes de facturar"
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox txtClienteD 
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1080
      Width           =   7455
   End
   Begin VB.TextBox txtClienteH 
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1680
      Width           =   7455
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   375
      Left            =   4080
      TabIndex        =   12
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   100270081
      CurrentDate     =   42705
   End
   Begin VB.TextBox txtConexionD 
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtConexionH 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton cmdFacturar 
      Caption         =   "&Facturar"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      ToolTipText     =   "Factura e Imprime todas las CONEXIONES pendientes de facturar"
      Top             =   2040
      Width           =   1695
   End
   Begin VB.ComboBox cboPeriodo 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Seleccione el Período a Facturar"
      Top             =   360
      Width           =   3615
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7920
      TabIndex        =   1
      ToolTipText     =   "Fin de la TAREA"
      Top             =   2040
      Width           =   1695
   End
   Begin MSComctlLib.ProgressBar prbProgreso 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7920
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Conexión"
      Height          =   195
      Index           =   11
      Left            =   240
      TabIndex        =   11
      Top             =   1440
      Width           =   660
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Desde"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   2160
      TabIndex        =   10
      Top             =   840
      Width           =   555
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Hasta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   2160
      TabIndex        =   9
      Top             =   1440
      Width           =   510
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de Emisión"
      Height          =   195
      Index           =   5
      Left            =   4080
      TabIndex        =   8
      Top             =   120
      Width           =   1260
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Período"
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   570
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Conexión"
      Height          =   195
      Index           =   8
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   660
   End
End
Attribute VB_Name = "frmLiquidar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub fillGrid()
Dim cliente As clsMODCliente

Dim clienteRep As New clsREPCliente

    Me.grdClientes.Rows = 1
    Me.grdClientes.Redraw = False
    For Each cliente In clienteRep.collectionRango(Val(Me.txtConexionD.Text), Val(Me.txtConexionH.Text))
        With cliente
            Me.grdClientes.AddItem modGrid.array2itemGrid(Array(Me.grdClientes.Rows, .clienteId, .apellido & ", " & .nombre, .zona, .ruta, .orden))
        End With
    Next
    Me.grdClientes.Redraw = True
    
End Sub

Private Sub cmdEliminar_Click()
Dim ctlFac As New clsCtlFactura

    If Me.cboPeriodo.ListIndex < 0 Then Exit Sub
    
    If MsgBox("Está SEGURO?", vbYesNo) = vbNo Then Exit Sub
    If MsgBox("Está Realmente SEGURO?", vbYesNo) = vbNo Then Exit Sub
    If MsgBox("Está Segurísimo?", vbYesNo) = vbNo Then Exit Sub
    
    Me.cmdEliminar.Enabled = False
    Me.MousePointer = 11
    
    Me.stbEstado.SimpleText = "Eliminando . . ."
    
    ctlFac.deletePeriodo Me.cboPeriodo.ItemData(Me.cboPeriodo.ListIndex), dbapp
    
    Me.stbEstado.SimpleText = " . . . TERMINADO"
    
    Me.MousePointer = 0
    Me.cmdEliminar.Enabled = True
    
End Sub

Private Sub cmdExtremos_Click()
Dim cliente As clsMODCliente

Dim clienteRep As New clsREPCliente

Dim clientes As Collection

    Set clientes = clienteRep.collectionActivos
    
    Set cliente = clientes.item(1)
    
    Me.txtConexionD.Text = cliente.clienteId
    Me.txtClienteD.Text = cliente.apellidonombre
    
    Set cliente = clientes.item(clientes.Count)
    
    Me.txtConexionH.Text = cliente.clienteId
    Me.txtClienteH.Text = cliente.apellidonombre
    
    fillGrid
    
End Sub

Private Sub cmdFacturar_Click()
Dim alicuota As New clsMyAAlicuota
Dim cliente As clsMODCliente
Dim operador As New clsMyAOperador
Dim periodo As New clsRESTPeriodo
Dim factura As New clsMyAFactura

Dim clienteRep As New clsREPCliente

Dim clientes As Collection

Dim ctlFac As New clsCtlFactura

    alicuota.findLast dbapp
    operador.findLast dbapp
    
    periodo.periodoId = Me.cboPeriodo.ItemData(Me.cboPeriodo.ListIndex)
    periodo.findByPrimaryKey

    Set clientes = clienteRep.collectionRango(Val(Me.txtConexionD.Text), Val(Me.txtConexionH.Text))
    
    Me.prbProgreso.Min = 1
    Me.prbProgreso.Max = clientes.Count + 1
    Me.prbProgreso.value = 1
    
    For Each cliente In clientes
        DoEvents
        
        Me.prbProgreso.value = Me.prbProgreso.value + 1
        
        ctlFac.makeLiquidacion Me.dtpFecha.value, cliente, alicuota, operador, periodo, dbapp, Me.stbEstado
    Next
    
    MsgBox "Facturación del Período Terminada . . ."

End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()
Dim periodo As New clsRESTPeriodo

    modGrid.makeGrid Me.grdClientes, Array(Array("#", 500), Array("Conexion", 1000), Array("operador", 5000), Array("Zona", 800), Array("Ruta", 800), Array("Orden", 800)), 0, 1, flexSelectionByRow

    cmdExtremos_Click
    
    periodo.fillCombo Me.cboPeriodo
    
    Me.dtpFecha.value = Date
    
End Sub

Private Sub txtConexionD_GotFocus()

    marcarseleccion Me.txtConexionD
    
End Sub

Private Sub txtConexionD_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtConexionD_LostFocus
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtConexionD_LostFocus()
Dim cliente As clsMODCliente

Dim clienteRep As New clsREPCliente

    Set cliente = clienteRep.findLastByClienteID(Val(Me.txtConexionD.Text))
    
    Me.txtConexionD.Text = cliente.clienteId
    Me.txtClienteD.Text = cliente.apellidonombre
    
End Sub

Private Sub txtConexionH_GotFocus()

    marcarseleccion Me.txtConexionH
    
End Sub

Private Sub txtConexionH_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtConexionH_LostFocus
        KeyAscii = 0
    End If

End Sub

Private Sub txtConexionH_LostFocus()
Dim cliente As clsMODCliente

Dim clienteRep As New clsREPCliente

    Set cliente = clienteRep.findLastByClienteID(Val(Me.txtConexionH.Text))
    
    Me.txtConexionH.Text = cliente.clienteId
    Me.txtClienteH.Text = cliente.apellidonombre

End Sub
