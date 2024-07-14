VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmOrden 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Zonas y Rutas"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10065
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   10065
   Begin VB.ComboBox LCli 
      Height          =   315
      Left            =   240
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   840
      Width           =   3375
   End
   Begin MSFlexGridLib.MSFlexGrid GCli 
      Height          =   1935
      Left            =   4680
      TabIndex        =   5
      Top             =   840
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   3413
      _Version        =   393216
      Cols            =   3
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin VB.CommandButton elimi 
      Caption         =   "<<"
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      ToolTipText     =   "Elimina la fila seleccionada"
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton inser 
      Caption         =   ">>"
      Height          =   255
      Left            =   3720
      TabIndex        =   3
      ToolTipText     =   "Inserta el Cliente debajo de la fila elegida"
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox ruta 
      Height          =   285
      Left            =   6480
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox zona 
      Height          =   285
      Left            =   4680
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton fin 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8280
      TabIndex        =   6
      ToolTipText     =   "Fin de la TAREA"
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label ncon 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Conexión"
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   16
      Top             =   2280
      Width           =   660
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Orden Actual"
      Height          =   195
      Index           =   5
      Left            =   2040
      TabIndex        =   15
      Top             =   2280
      Width           =   930
   End
   Begin VB.Label aorden 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2040
      TabIndex        =   14
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Ruta Actual"
      Height          =   195
      Index           =   4
      Left            =   2040
      TabIndex        =   13
      Top             =   1680
      Width           =   840
   End
   Begin VB.Label aruta 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2040
      TabIndex        =   12
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Zona Actual"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   1680
      Width           =   870
   End
   Begin VB.Label azona 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Clientes"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   600
      Width           =   555
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Ruta"
      Height          =   195
      Index           =   1
      Left            =   6480
      TabIndex        =   8
      Top             =   120
      Width           =   345
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Zona"
      Height          =   195
      Index           =   0
      Left            =   4680
      TabIndex        =   7
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmOrden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private vl As Integer

Private Sub llenar()
Dim cliente As clsMODCliente

Dim clienteRep As New clsREPCliente

On Error Resume Next
    
    If Val(ruta.Text) = 0 Then ruta.Text = 1
    If Val(zona.Text) = 0 Then zona.Text = 1
    
    GCli.Rows = 1
    For Each cliente In clienteRep.collectionActivosByZonaRuta(Val(zona.Text), Val(ruta.Text))
        GCli.AddItem modGrid.array2itemGrid(Array(GCli.Rows - 1, cliente.clienteId, cliente.apellidonombre))
    Next
    
    clienteRep.fillComboOtros Me.LCli, Val(zona.Text), Val(ruta.Text)

End Sub

Private Sub elimi_Click()
Dim ecli As Long
Dim cx As Long
Dim crit, nm, enm As String
Dim ps, nf, ds As Integer

Dim cliente As clsMODCliente

Dim clienteRep As New clsREPCliente

On Error Resume Next
    
    If GCli.row < 1 Then Exit Sub
    GCli.col = 1
    ecli = GCli.Text
    GCli.col = 2
    enm = GCli.Text
    ps = GCli.row
    nf = GCli.Rows
    For ds = ps + 1 To GCli.Rows - 1
        GCli.row = ds
        GCli.col = 1
        cx = GCli.Text
        Set cliente = clienteRep.findLastByClienteID(cx)
        cliente.zona = zona.Text
        cliente.ruta = ruta.Text
        cliente.orden = ds - 1
        Set cliente = clienteRep.update(cliente, cliente.uniqueId)
        GCli.col = 2
        nm = GCli.Text
        GCli.row = ds - 1
        GCli.col = 1
        GCli.Text = cx
        GCli.col = 2
        GCli.Text = nm
    Next ds
    
    Set cliente = clienteRep.findLastByClienteID(ecli)
    cliente.zona = 0
    cliente.ruta = 0
    cliente.orden = 0
    Set cliente = clienteRep.update(cliente, cliente.uniqueId)
    
    LCli.AddItem enm
    LCli.ItemData(LCli.NewIndex) = ecli
    LCli.ListIndex = LCli.NewIndex
    nf = nf - 1
    GCli.Rows = nf

End Sub

Private Sub fin_Click()
    
    Unload Me
    
End Sub

Private Sub Form_Activate()
    
    If Val(zona.Text) = 0 Then
        zona.Text = 1
        llenar
    End If

End Sub

Private Sub Form_Load()

On Error Resume Next
    
    GCli.Rows = 1
    GCli.Cols = 3
    GCli.fixedRows = 1
    GCli.fixedCols = 0
    GCli.row = 1
    GCli.col = 0
    GCli.Text = "Orden"
    GCli.ColWidth(0) = 600
    GCli.col = 1
    GCli.Text = "Conexión"
    GCli.ColWidth(1) = 800
    GCli.col = 2
    GCli.Text = "Cliente"
    GCli.ColWidth(2) = 3300

End Sub

Private Sub inser_Click()
Dim ps, nf, ds As Integer
Dim cx As Long, nm, crit As String

Dim cliente As clsMODCliente

Dim clienteRep As New clsREPCliente

On Error Resume Next
    
    If LCli.ListIndex < 0 Then Exit Sub
    ps = GCli.row + 1
    nf = GCli.Rows
    nf = nf + 1
    GCli.Rows = nf
    GCli.row = nf - 1
    GCli.col = 0
    GCli.Text = nf - 1
    For ds = GCli.Rows - 1 To ps + 1 Step -1
        GCli.row = ds - 1
        GCli.col = 1
        cx = GCli.Text
        Set cliente = clienteRep.findLastByClienteID(cx)
        cliente.zona = zona.Text
        cliente.ruta = ruta.Text
        cliente.orden = ds
        Set cliente = clienteRep.update(cliente, cliente.uniqueId)
        GCli.col = 2
        nm = GCli.Text
        GCli.row = ds
        GCli.col = 1
        GCli.Text = cx
        GCli.col = 2
        GCli.Text = nm
    Next ds
    GCli.row = ps
    GCli.col = 1
    GCli.Text = LCli.ItemData(LCli.ListIndex)
    Set cliente = clienteRep.findLastByClienteID(Me.LCli.ItemData(Me.LCli.ListIndex))
    cliente.zona = zona.Text
    cliente.ruta = ruta.Text
    cliente.orden = ps
    Set cliente = clienteRep.update(cliente, cliente.uniqueId)
    GCli.col = 2
    GCli.Text = LCli.Text
    LCli.RemoveItem LCli.ListIndex
    LCli.ListIndex = 0

End Sub

Private Sub lcli_Click()
Dim cliente As clsMODCliente

Dim clienteRep As New clsREPCliente

On Error Resume Next
    
    Set cliente = clienteRep.findLastByClienteID(Me.LCli.ItemData(Me.LCli.ListIndex))
    azona.Caption = ""
    aruta.Caption = ""
    aorden.Caption = ""
    If cliente.uniqueId > 0 Then
        azona.Caption = cliente.zona
        aruta.Caption = cliente.ruta
        aorden.Caption = cliente.orden
        ncon.Caption = cliente.clienteId
    End If

End Sub

Private Sub ruta_GotFocus()
    
    vl = Val(ruta.Text)
    ruta.SelStart = 0
    ruta.SelLength = Len(ruta.Text)

End Sub

Private Sub ruta_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then LCli.SetFocus

End Sub

Private Sub ruta_LostFocus()
    
    If Val(ruta.Text) = 0 Then ruta.Text = 1
    If vl = 0 Then vl = ruta.Text
    If vl <> Val(ruta.Text) Then llenar

End Sub

Private Sub zona_GotFocus()
    
    vl = Val(zona.Text)
    zona.SelStart = 0
    zona.SelLength = Len(zona.Text)

End Sub

Private Sub zona_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then ruta.SetFocus

End Sub

Private Sub zona_LostFocus()
    
    If Val(zona.Text) = 0 Then zona.Text = 1
    If vl = 0 Then vl = zona.Text
    If vl <> Val(zona.Text) Then llenar

End Sub
