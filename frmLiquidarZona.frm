VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLiquidarZona 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Liquidación por Zona"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9885
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   9885
   Begin VB.ComboBox cboZona 
      Height          =   315
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   9
      ToolTipText     =   "Seleccione el Período a Facturar"
      Top             =   360
      Width           =   3615
   End
   Begin MSFlexGridLib.MSFlexGrid grdClientes 
      Height          =   5295
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   9340
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar stbEstado 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   6930
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   109707265
      CurrentDate     =   42705
   End
   Begin VB.CommandButton cmdFacturar 
      Caption         =   "&Facturar"
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      ToolTipText     =   "Factura e Imprime todas las CONEXIONES pendientes de facturar"
      Top             =   840
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
      Top             =   840
      Width           =   1695
   End
   Begin MSComctlLib.ProgressBar prbProgreso 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6720
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Zona"
      Height          =   195
      Index           =   0
      Left            =   6000
      TabIndex        =   10
      Top             =   120
      Width           =   375
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de Emisión"
      Height          =   195
      Index           =   5
      Left            =   4080
      TabIndex        =   5
      Top             =   120
      Width           =   1260
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Período"
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   570
   End
End
Attribute VB_Name = "frmLiquidarZona"
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
    For Each cliente In clienteRep.collectionZona(Val(Me.cboZona.Text))
        Me.grdClientes.AddItem modGrid.array2itemGrid(Array(Me.grdClientes.Rows, cliente.clienteId, cliente.apellido & ", " & cliente.nombre, cliente.zona, cliente.ruta, cliente.orden, ""))
        Me.grdClientes.RowData(Me.grdClientes.Rows - 1) = cliente.clienteId
    Next
    Me.grdClientes.Redraw = True
    
End Sub

Private Sub fillZonas()
Dim cliente As clsMODCliente

Dim clienteRep As New clsREPCliente

Dim zona As Integer

    zona = -1
    Me.cboZona.Clear
    For Each cliente In clienteRep.collectionActivosByZona
        If zona <> cliente.zona Then
            Me.cboZona.AddItem cliente.zona
            zona = cliente.zona
        End If
    Next
    If Me.cboZona.ListCount > 0 Then Me.cboZona.ListIndex = 0
    
End Sub

Private Sub cboZona_Click()

    If Me.cboZona.ListIndex < 0 Then Exit Sub
    
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

    Set clientes = clienteRep.collectionZona(Val(Me.cboZona.Text))
    
    Me.prbProgreso.Min = 1
    Me.prbProgreso.Max = clientes.Count + 1
    Me.prbProgreso.value = 1
    
    For Each cliente In clientes
        DoEvents
        
        Me.prbProgreso.value = Me.prbProgreso.value + 1
        
        ctlFac.makeLiquidacion Me.dtpFecha.value, cliente, alicuota, operador, periodo, dbapp, Me.stbEstado
    Next
    
    MsgBox "Facturación de Zona Terminada . . ."

End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()
Dim periodo As New clsRESTPeriodo

    modGrid.makeGrid Me.grdClientes, Array(Array("#", 500), Array("Conexion", 1000), Array("Cliente", 5000), Array("Zona", 800), Array("Ruta", 800), Array("Orden", 800)), 0, 1, flexSelectionByRow

    periodo.fillCombo Me.cboPeriodo
    
    Me.dtpFecha.value = Date
    
    fillZonas
    
End Sub

