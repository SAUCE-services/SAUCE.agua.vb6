VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmImpRuta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión de Liquidación por Ruta"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9840
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   9840
   Begin VB.ComboBox cboRuta 
      Height          =   315
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   10
      ToolTipText     =   "Seleccione el Período a Facturar"
      Top             =   360
      Width           =   1695
   End
   Begin VB.PictureBox picConsumo 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000E&
      Height          =   5895
      Left            =   9960
      ScaleHeight     =   5835
      ScaleWidth      =   5835
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.ComboBox cboZona 
      Height          =   315
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   7
      ToolTipText     =   "Seleccione el Período a Facturar"
      Top             =   360
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid grdClientes 
      Height          =   5295
      Left            =   240
      TabIndex        =   6
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
      TabIndex        =   5
      Top             =   6960
      Width           =   9840
      _ExtentX        =   17357
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   7920
      TabIndex        =   3
      ToolTipText     =   "Factura e Imprime todas las CONEXIONES pendientes de facturar"
      Top             =   360
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
   Begin Crystal.CrystalReport crpLiquidacion 
      Left            =   8400
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Ruta"
      Height          =   195
      Index           =   1
      Left            =   6000
      TabIndex        =   11
      Top             =   120
      Width           =   345
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Zona"
      Height          =   195
      Index           =   0
      Left            =   4080
      TabIndex        =   8
      Top             =   120
      Width           =   375
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
Attribute VB_Name = "frmImpRuta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub fillGrid()
Dim cliente As clsMODCliente
Dim factura As New clsMyAFactura

Dim clienteRep As New clsREPCliente

    Me.grdClientes.Rows = 1
    Me.grdClientes.Redraw = False
    For Each cliente In clienteRep.collectionRuta(Val(Me.cboZona.Text), Val(Me.cboRuta.Text))
        With cliente
            Me.grdClientes.AddItem modGrid.array2itemGrid(Array(Me.grdClientes.Rows, .clienteId, .apellido & ", " & .nombre, .zona, .ruta, .orden))
            Me.grdClientes.RowData(Me.grdClientes.Rows - 1) = .clienteId
        End With
        With factura
            .clienteId = cliente.clienteId
            .periodoId = Me.cboPeriodo.ItemData(Me.cboPeriodo.ListIndex)

            .findByClientePeriodo dbapp

            modGrid.letCheckCell Me.grdClientes, Me.grdClientes.Rows - 1, 6, (.autoID > 0)
        End With
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

Private Sub fillRutas()
Dim cliente As clsMODCliente

Dim clienteRep As New clsREPCliente

Dim ruta As Integer

    ruta = -1
    Me.cboRuta.Clear
    For Each cliente In clienteRep.collectionActivosByRuta(Val(Me.cboZona.Text))
        If ruta <> cliente.ruta Then
            Me.cboRuta.AddItem cliente.ruta
            ruta = cliente.ruta
        End If
    Next
    If Me.cboRuta.ListCount > 0 Then Me.cboRuta.ListIndex = 0

End Sub

Private Sub cboPeriodo_Click()

    If Me.cboPeriodo.ListIndex < 0 Then Exit Sub
    
    fillGrid
    
End Sub

Private Sub cboRuta_Click()

    If Me.cboRuta.ListIndex < 0 Then Exit Sub
    
    fillGrid
    
End Sub

Private Sub cboZona_Click()

    If Me.cboZona.ListIndex < 0 Then Exit Sub

    fillRutas
    
End Sub

Private Sub cmdImprimir_Click()
Dim intCiclo As Integer

Dim factura As New clsMyAFactura

Dim liquidacion_service As New clsCtlLiquidacion

    Me.cmdImprimir.Enabled = False
    Me.MousePointer = 11
    
    Me.prbProgreso.Min = 0
    Me.prbProgreso.Max = Me.grdClientes.Rows - 1
    Me.prbProgreso.value = Me.prbProgreso.Min
    For intCiclo = 1 To Me.grdClientes.Rows - 1
        Me.prbProgreso.value = Me.prbProgreso.value + 1
        Me.prbProgreso.Refresh
        
        DoEvents
        
        With factura
            .clienteId = Me.grdClientes.RowData(intCiclo)
            .periodoId = Me.cboPeriodo.ItemData(Me.cboPeriodo.ListIndex)
            
            .findByClientePeriodo dbapp
            
            If .autoID > 0 Then
                liquidacion_service.printLiquidacion Me.hWnd, .puntoVta, .nroComprob, dbapp, Me.picConsumo, Me.crpLiquidacion, True

                Me.stbEstado.SimpleText = "Imprimiendo -> Conexión: " & .clienteId & " - Liquidación: " & .puntoVta & "/" & .nroComprob & " . . ."
            End If
        End With
    Next intCiclo
    
    Me.stbEstado.SimpleText = ""
    
    Me.MousePointer = 0
    Me.cmdImprimir.Enabled = True
    
End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()
Dim periodo As New clsRESTPeriodo

    modGrid.makeGrid Me.grdClientes, Array(Array("#", 500), Array("Conexion", 1000), Array("Cliente", 5000), Array("Zona", 700), Array("Ruta", 700), Array("Orden", 800), Array("Fac", 300)), 0, 1, flexSelectionByRow

    periodo.fillCombo Me.cboPeriodo
    
    fillZonas
    
End Sub

