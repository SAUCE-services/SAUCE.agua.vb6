VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRepLecturas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lecturas por Zona y Ruta"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7935
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   7935
   Begin VB.CommandButton cmdImprimir 
      Cancel          =   -1  'True
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      ToolTipText     =   "Fin de la TAREA"
      Top             =   840
      Width           =   1695
   End
   Begin VB.ComboBox cboRuta 
      Height          =   315
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   5
      ToolTipText     =   "Seleccione el Período a Facturar"
      Top             =   360
      Width           =   1695
   End
   Begin VB.ComboBox cboZona 
      Height          =   315
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   3
      ToolTipText     =   "Seleccione el Período a Facturar"
      Top             =   360
      Width           =   1695
   End
   Begin VB.ComboBox cboPeriodo 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Seleccione el Período a Facturar"
      Top             =   360
      Width           =   3615
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6000
      TabIndex        =   0
      ToolTipText     =   "Fin de la TAREA"
      Top             =   840
      Width           =   1695
   End
   Begin Crystal.CrystalReport crpReporte 
      Left            =   240
      Top             =   840
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
      TabIndex        =   6
      Top             =   120
      Width           =   345
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Zona"
      Height          =   195
      Index           =   0
      Left            =   4080
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Período"
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   570
   End
End
Attribute VB_Name = "frmRepLecturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub fillZonas()
Dim cliente As clsMODCliente

Dim clienteRep As New clsREPCliente

Dim intZona As Integer

    intZona = -1
    Me.cboZona.Clear
    For Each cliente In clienteRep.collectionActivosByZona
        If intZona <> cliente.zona Then
            Me.cboZona.AddItem cliente.zona
            intZona = cliente.zona
        End If
    Next
    If Me.cboZona.ListCount > 0 Then Me.cboZona.ListIndex = 0
    
End Sub

Private Sub cboZona_Click()
Dim cliente As clsMODCliente

Dim clienteRep As New clsREPCliente

Dim ruta As Integer

    If Me.cboZona.ListIndex < 0 Then Exit Sub
    
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

Private Sub cmdImprimir_Click()
Dim impresionService As New clsCtlImpresion
Dim consumoService As New clsCtlConsumo

Dim cliente As clsMODCliente
Dim medidor As clsMyAMedidor
Dim operador As New clsMyAOperador

Dim clienteRep As New clsREPCliente

Dim fechaActual As Date
Dim fechaAnterior As Date

Dim estadoActual As Long
Dim estadoAnterior As Long

Dim cuit As String
Dim mens As String

    Me.MousePointer = 11
    
    Me.cmdImprimir.Enabled = False
    
    operador.findLast dbapp
    
    cuit = Left(operador.cuit, 2) & "-" & Mid(operador.cuit, 3, 8) & "-" & Right(operador.cuit, 1)
    Select Case operador.situacionIVA
        Case 1:
            mens = "Resp. Inscripto"
        Case 2:
            mens = "Resp. No Inscripto"
        Case 3:
            mens = "Cons. Final"
        Case 4:
            mens = "Exento"
        Case 5:
            mens = "No Responsable"
    End Select
    
    For Each cliente In clienteRep.collectionActivos2Lectura(Val(Me.cboZona.Text), Val(Me.cboRuta.Text))
        Set medidor = New clsMyAMedidor
        With medidor
            .clienteId = cliente.clienteId
            
            .findByClienteId dbapp
        End With
        
        consumoService.datosConsumo cliente.clienteId, Me.cboPeriodo.ItemData(Me.cboPeriodo.ListIndex), medidor.medidorID, Date, fechaActual, estadoActual, fechaAnterior, estadoAnterior, dbapp
    Next
    
    impresionService.printReport Me.crpReporte, "rptLecturas", dbapp.stringConnection, , Array(Array("zona", Val(Me.cboZona.Text)), Array("ruta", Val(Me.cboRuta.Text)), Array("periodo", Me.cboPeriodo.Text), Array("periodoID", Me.cboPeriodo.ItemData(Me.cboPeriodo.ListIndex))), , , _
        Array(Array("nomope", operador.razonSocial), _
        Array("domope", operador.calle & " " & operador.puerta & " " & operador.piso & " " & operador.dpto & " - " & operador.localidad), _
        Array("locope", "C.P. " & operador.codigoPostal & " " & operador.provincia & " - Tel: " & operador.telefono), _
        Array("opegr1", "C.U.I.T. NRO: " & cuit & " ING. BRUTOS: " & operador.ingresosBrutos), _
        Array("opegr2", "I.V.A. " & mens & "   NRO. E.P.A.S. " & operador.numeroEpas), _
        Array("titulo", "Lecturas por Zona y Ruta"))
    
    Me.cmdImprimir.Enabled = True
    
    Me.MousePointer = 0
    
End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()
Dim periodo As New clsRESTPeriodo

    periodo.fillCombo Me.cboPeriodo
    
    fillZonas
    
End Sub

