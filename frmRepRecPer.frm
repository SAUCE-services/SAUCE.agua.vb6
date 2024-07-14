VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRepRecPer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recaudación por Rango de Fechas"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   5985
   Begin VB.CommandButton cmdDuplicar 
      Caption         =   ">"
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   360
      Width           =   255
   End
   Begin MSComCtl2.DTPicker dtpDesde 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   109248513
      CurrentDate     =   42930
   End
   Begin VB.PictureBox picConsumo 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000E&
      Height          =   5895
      Left            =   9960
      ScaleHeight     =   5835
      ScaleWidth      =   5835
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      ToolTipText     =   "Factura e Imprime todas las CONEXIONES pendientes de facturar"
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      ToolTipText     =   "Fin de la TAREA"
      Top             =   840
      Width           =   1695
   End
   Begin Crystal.CrystalReport crpReporte 
      Left            =   5160
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker dtpHasta 
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   109248513
      CurrentDate     =   42930
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Hasta"
      Height          =   195
      Index           =   0
      Left            =   2160
      TabIndex        =   5
      Top             =   120
      Width           =   420
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Desde"
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   465
   End
End
Attribute VB_Name = "frmRepRecPer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDuplicar_Click()

    Me.dtpHasta.value = Me.dtpDesde.value
    
End Sub

Private Sub cmdImprimir_Click()
Dim periodo As New clsRESTPeriodo
Dim liqperiodo As New clsVMyALiqPeriodo
Dim operador As New clsMyAOperador

Dim impresionService As New clsCtlImpresion

Dim cuit As String
Dim mens As String

    Me.MousePointer = 11
    Me.cmdImprimir.Enabled = False
    
    For Each periodo In periodo.collectionRecaudadoByPeriodo(Me.dtpDesde.value, Me.dtpHasta.value)
        periodo.liquidado = liqperiodo.collectionByPeriodoID(periodo.periodoId, dbapp)("k." & periodo.periodoId).liquidado
        periodo.save
    Next
    
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
    
    impresionService.printReport Me.crpReporte, "rptRecPeriodo", dbapp.stringConnection, , Array(Array("pDesde", toReportDate(Me.dtpDesde.value)), Array("pHasta", toReportDate(Me.dtpHasta.value))), , , _
        Array(Array("nomope", operador.razonSocial), _
        Array("domope", operador.calle & " " & operador.puerta & " " & operador.piso & " " & operador.dpto & " - " & operador.localidad), _
        Array("locope", "C.P. " & operador.codigoPostal & " " & operador.provincia & " - Tel: " & operador.telefono), _
        Array("opegr1", "C.U.I.T. NRO: " & cuit & " ING. BRUTOS: " & operador.ingresosBrutos), _
        Array("opegr2", "I.V.A. " & mens & "   NRO. E.P.A.S. " & operador.numeroEpas), _
        Array("titulo", "Recaudación Diaria"), _
        Array("info1", "Período : " & Me.dtpDesde.value & " - " & Me.dtpHasta.value))
    
    Me.cmdImprimir.Enabled = True
    Me.MousePointer = 0
    
End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()
    
    Me.dtpDesde.value = Date
    Me.dtpHasta.value = Date
    
End Sub

