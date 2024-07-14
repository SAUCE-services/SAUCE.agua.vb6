VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportRP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importar RapiPago"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11790
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   11790
   Begin MSComctlLib.StatusBar stbEstado 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   6720
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdImportar 
      Cancel          =   -1  'True
      Caption         =   "Importar"
      Height          =   375
      Left            =   9840
      TabIndex        =   6
      ToolTipText     =   "Fin de la TAREA"
      Top             =   360
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid grdPagos 
      Height          =   5175
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   9128
      _Version        =   393216
   End
   Begin VB.CommandButton cmdRevisar 
      Caption         =   "Revisar"
      Height          =   375
      Left            =   7920
      TabIndex        =   4
      ToolTipText     =   "Fin de la TAREA"
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdArchivo 
      Caption         =   ". . ."
      Height          =   255
      Left            =   6960
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin MSComDlg.CommonDialog cdlArchivo 
      Left            =   6000
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtArchivo 
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   7455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   9840
      TabIndex        =   0
      ToolTipText     =   "Fin de la TAREA"
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Pagos"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   450
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Archivo"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   540
   End
End
Attribute VB_Name = "frmImportRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private vPFR As Collection

Private Sub cmdArchivo_Click()
Dim pagofacil_service As New clsCtlPagoFacil

Dim pffile As New clsMyAPFFile

    Me.grdPagos.Rows = 1

    Me.cdlArchivo.Filter = "RapiPago | *.4768"
    Me.cdlArchivo.ShowOpen
    
    Me.txtArchivo.Text = Me.cdlArchivo.filename
    
    Me.MousePointer = 11
    Me.stbEstado.SimpleText = "Cargando ARCHIVO . . ."
    
    Set vPFR = pagofacil_service.loadFile(Me.txtArchivo.Text, dbapp)
    
    pffile.filename = modConv.parseFilename(Me.txtArchivo.Text)
    pffile.findByPrimaryKey dbapp
    
    Me.cmdImportar.Enabled = True
    If pffile.import <> 0 Then Me.cmdImportar.Enabled = False
    
    Me.stbEstado.SimpleText = ""
    
    Me.MousePointer = 0
    
End Sub

Private Sub cmdImportar_Click()
Dim pffile As New clsMyAPFFile
Dim pfrecord As clsPFRecord

Dim factura As New clsMyAFactura
Dim cliente As clsMODCliente
Dim periodo As New clsRESTPeriodo

Dim clienteRep As New clsREPCliente

Dim transactionId As String

Dim importado As Boolean

    Me.stbEstado.SimpleText = "Importando . . ."
    Me.MousePointer = 11

    pffile.filename = modConv.parseFilename(Me.txtArchivo.Text)
    pffile.findByPrimaryKey dbapp
    
    importado = False
    
    ' Marcando Pagos de PF
    For Each pfrecord In vPFR
        transactionId = Mid(pfrecord.PFRec6.barCode, 18, 14)
        If Left(transactionId, 1) = "1" Then
            factura.pfcodigo = pfrecord.PFRec6.barCode
            factura.findByPFCodigo dbapp
            
            If factura.autoID = 0 Then
                MsgBox "ERROR: LIQUIDACION NO Encontrada"
            Else
                ' Si la factura está anulada avisa
                If factura.anulada = 1 Then
                    Set cliente = clienteRep.findLastByClienteID(factura.clienteId)
                    periodo.periodoId = factura.periodoId
                    periodo.findByPrimaryKey
                    MsgBox "ADVERTENCIA: Está PAGANDO una Liquidación ANULADA - Cliente " & cliente.clienteId & " " & cliente.apellidonombre & " - " & periodo.descripcion
                End If
                ' Asienta Pago Sistema Nuevo
                factura.pagada = 1
                factura.fechapago = pfrecord.PFRec5.paymentDate
                factura.tipoId = cntPagoRapiPago
                factura.save dbapp
                
                importado = True
            End If
        End If
    Next
    
    ' Marca archivo importado
    If importado Then
        pffile.import = 1
        pffile.fechaImport = Date
        pffile.save dbapp
    End If
    
    MsgBox "Importación TERMINADA"
    
    Me.cmdImportar.Enabled = False
    Me.stbEstado.SimpleText = ""
    Me.MousePointer = 0
    
End Sub

Private Sub cmdRevisar_Click()
Dim pfrecord As clsPFRecord

Dim factura As New clsMyAFactura
Dim cliente As clsMODCliente
Dim periodo As New clsRESTPeriodo

Dim clienteRep As New clsREPCliente

Dim transactionId As String

    Me.grdPagos.Rows = 1
    
    If vPFR.Count = 0 Then Exit Sub
    
    Me.cmdRevisar.Enabled = False
    Me.MousePointer = 11
    
    Me.grdPagos.Redraw = False
    For Each pfrecord In vPFR
        transactionId = Mid(pfrecord.PFRec6.barCode, 18, 14)
        If Left(transactionId, 1) = "1" Then
            factura.pfcodigo = pfrecord.PFRec6.barCode
            factura.findByPFCodigo dbapp
            
            Set cliente = clienteRep.findLastByClienteID(factura.clienteId)
            
            periodo.periodoId = factura.periodoId
            periodo.findByPrimaryKey
            
            Me.grdPagos.AddItem modGrid.array2itemGrid(Array(factura.clienteId, cliente.apellidonombre, periodo.descripcion, Format(factura.total, "0.00"), Format(pfrecord.PFRec5.amount, "0.00"), pfrecord.PFRec5.paymentDate))
        End If
    Next
    Me.grdPagos.Redraw = True
    
    Me.cmdRevisar.Enabled = True
    Me.MousePointer = 0

End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()

    modGrid.makeGrid Me.grdPagos, Array(Array("Conexion", 800), Array("Cliente", 3900), Array("Periodo", 2000), Array("Total Liquidación", 1500), Array("Importe cobrado", 1500), Array("Fecha Pago", 1200)), 0, 1, flexSelectionByRow
    
End Sub
