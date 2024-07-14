VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRepRubFacPer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rubros Facturados por Período"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   4095
   Begin VB.CommandButton cmdImprimir 
      Cancel          =   -1  'True
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      ToolTipText     =   "Fin de la TAREA"
      Top             =   840
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
      Left            =   2160
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
      Caption         =   "Período"
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   570
   End
End
Attribute VB_Name = "frmRepRubFacPer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdImprimir_Click()
Dim impresionService As New clsCtlImpresion
Dim ctlRub As New clsCtlRubro

Dim operador As New clsMyAOperador

Dim cuit As String
Dim mens As String

    Me.MousePointer = 11
    
    Me.cmdImprimir.Enabled = False
    
    ctlRub.updateVigente dbapp
    
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
    
    impresionService.printReport Me.crpReporte, "rptRubroFact", dbapp.stringConnection, Array("impuestos"), Array(Array("periodoID", Me.cboPeriodo.ItemData(Me.cboPeriodo.ListIndex))), , , _
        Array(Array("nomope", operador.razonSocial), _
        Array("domope", operador.calle & " " & operador.puerta & " " & operador.piso & " " & operador.dpto & " - " & operador.localidad), _
        Array("locope", "C.P. " & operador.codigoPostal & " " & operador.provincia & " - Tel: " & operador.telefono), _
        Array("opegr1", "C.U.I.T. NRO: " & cuit & " ING. BRUTOS: " & operador.ingresosBrutos), _
        Array("opegr2", "I.V.A. " & mens & "   NRO. E.P.A.S. " & operador.numeroEpas), _
        Array("titulo", "Rubros Facturados por Período"))
    
    Me.cmdImprimir.Enabled = True
    
    Me.MousePointer = 0
    
End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()
Dim periodo As New clsRESTPeriodo

    periodo.fillCombo Me.cboPeriodo
    
End Sub

