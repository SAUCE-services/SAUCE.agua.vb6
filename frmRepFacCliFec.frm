VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRepFacCliFec 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturas por Cliente entre Fechas"
   ClientHeight    =   1965
   ClientLeft      =   1410
   ClientTop       =   2805
   ClientWidth     =   6030
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   6030
   Begin VB.TextBox txtCliente 
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   5535
   End
   Begin VB.CommandButton cmdDuplicar 
      Caption         =   ">"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   960
      Width           =   255
   End
   Begin Crystal.CrystalReport crpFacturas 
      Left            =   3360
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   1
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdImprimir 
      Cancel          =   -1  'True
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      ToolTipText     =   "Imprime el Detalle de la DEUDA"
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      ToolTipText     =   "Fin de la TAREA"
      Top             =   1440
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker dtpDesde 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   109248513
      CurrentDate     =   42930
   End
   Begin MSComCtl2.DTPicker dtpHasta 
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   109248513
      CurrentDate     =   42930
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Cliente"
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   480
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Desde"
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   465
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Hasta"
      Height          =   195
      Index           =   0
      Left            =   2160
      TabIndex        =   6
      Top             =   720
      Width           =   420
   End
End
Attribute VB_Name = "frmRepFacCliFec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cliente As New clsMODCliente

Private Sub cmdDuplicar_Click()

    Me.dtpHasta.value = Me.dtpDesde.value
    
End Sub

Private Sub cmdImprimir_Click()
Dim operador As New clsMyAOperador

Dim cuit As String
Dim mens As String

Dim impresionService As New clsCtlImpresion

    Me.MousePointer = 11
    
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
    
    impresionService.printReport Me.crpFacturas, "rptDMovCli", dbapp.stringConnection, , _
        Array(Array("cliente_id", cliente.clienteId), Array("desde", toReportDate(Me.dtpDesde.value)), _
        Array("hasta", toReportDate(Me.dtpHasta.value))), , , _
        Array(Array("nomope", operador.razonSocial), _
        Array("domope", operador.calle & " " & operador.puerta & " " & operador.piso & " " & operador.dpto & " - " & operador.localidad), _
        Array("locope", "C.P. " & operador.codigoPostal & " " & operador.provincia & " - Tel: " & operador.telefono), _
        Array("opegr1", "C.U.I.T. NRO: " & cuit & " ING. BRUTOS: " & operador.ingresosBrutos), _
        Array("opegr2", "I.V.A. " & mens & "   NRO. E.P.A.S. " & operador.numeroEpas), _
        Array("titulo", "Facturas por Cliente entre Fechas"))
        
    Me.MousePointer = 0

End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()

    Me.dtpDesde.value = Date
    Me.dtpHasta.value = Date
    
End Sub

Private Sub txtCliente_GotFocus()

    marcarseleccion Me.txtCliente
    
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
Dim clienteRep As New clsREPCliente

    Set cliente = clienteRep.formSearch(frmBuscarRest, KeyAscii, "Clientes")
    
    Me.txtCliente.Text = cliente.textFound
    
    KeyAscii = 0
    
End Sub


