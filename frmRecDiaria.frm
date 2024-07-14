VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRecDiaria 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recaudación Diaria"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   7920
   Begin Crystal.CrystalReport crpReporte 
      Left            =   5280
      Top             =   6840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   4080
      TabIndex        =   8
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   7080
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid grdComprobantes 
      Height          =   5655
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   9975
      _Version        =   393216
   End
   Begin VB.CommandButton cmdRevisar 
      Caption         =   "Revisar"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   6000
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   77398017
      CurrentDate     =   42575
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Total Recaudación"
      Height          =   195
      Index           =   7
      Left            =   6000
      TabIndex        =   7
      Top             =   6840
      Width           =   1365
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Comprobantes"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   1020
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Pago"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   870
   End
End
Attribute VB_Name = "frmRecDiaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdImprimir_Click()
Dim ctlImp As New clsCtlImpresion

    Me.cmdImprimir.Enabled = False
    Me.MousePointer = 11
    
    ctlImp.printReport Me.crpReporte, "rptRecDiaria", dbmy.stringConnection, Array("sCuota"), Array(Array("pFechaPago", toReportDate(Me.dtpFecha.Value)), Array("pTotal", Val(Me.txtTotal.Text)))
    
    Me.MousePointer = 0
    Me.cmdImprimir.Enabled = True
    
End Sub

Private Sub cmdRevisar_Click()
Dim curInteres As Currency
Dim curTotalDia As Currency

Dim lngCliID As Long

Dim intPerID As Integer

Dim objMFac As New clsMyAFactura
Dim objMCli As New clsMyACliente
Dim objMPer As New clsMyAPeriodo
Dim objMCuo As New clsMyACuota
Dim objMNCr As New clsMyANCredito
Dim objMRec As New clsMyARecibo

    lngCliID = 0
    intPerID = 0
    curTotalDia = 0

    Me.grdComprobantes.Rows = 1
    Me.grdComprobantes.Redraw = False
    For Each objMFac In objMFac.collectionByPago(Me.dtpFecha.Value, dbmy)
        If lngCliID <> objMFac.clienteID Then
            objMCli.clienteID = objMFac.clienteID
            objMCli.findLast dbmy
        End If
        If intPerID <> objMFac.periodoID Then
            objMPer.periodoID = objMFac.periodoID
            objMPer.findByPrimaryKey dbmy
        End If
        
        curInteres = 0
        If objMFac.fechaPago > objMPer.fechaPrimero Then curInteres = modInteres.interes(objMFac.total, objMFac.tasa, objMPer.fechaPrimero, objMPer.fechaSegundo)
        objMFac.interes = curInteres
        objMFac.save dbmy
        
        Me.grdComprobantes.AddItem modGrid.array2itemGrid(Array("Liquidación", objMCli.apellido & ", " & objMCli.nombre, objMFac.puntoVta & "/" & objMFac.nroComprob, Format(objMFac.total + curInteres, "0.00")))
        curTotalDia = curTotalDia + objMFac.total + curInteres
    Next
    
    For Each objMCuo In objMCuo.collectionByPago(Me.dtpFecha.Value, dbmy)
        If lngCliID <> objMCuo.clienteID Then
            objMCli.clienteID = objMCuo.clienteID
            objMCli.findLast dbmy
        End If
        Me.grdComprobantes.AddItem modGrid.array2itemGrid(Array("Cuota", objMCli.apellido & ", " & objMCli.nombre, objMCuo.planID & "/" & objMCuo.cuotaID, Format(objMCuo.importe, "0.00")))
        curTotalDia = curTotalDia + objMCuo.importe
    Next
    
    For Each objMNCr In objMNCr.collectionByPago(Me.dtpFecha.Value, dbmy)
        If lngCliID <> objMNCr.clienteID Then
            objMCli.clienteID = objMNCr.clienteID
            objMCli.findLast dbmy
        End If
        Me.grdComprobantes.AddItem modGrid.array2itemGrid(Array("N.Crédito", objMCli.apellido & ", " & objMCli.nombre, objMNCr.serieID & "/" & objMNCr.numero, Format(-objMNCr.total, "0.00")))
        curTotalDia = curTotalDia - objMNCr.total
    Next
    
    For Each objMRec In objMRec.collectionByPago(Me.dtpFecha.Value, dbmy)
        If lngCliID <> objMRec.clienteID Then
            objMCli.clienteID = objMRec.clienteID
            objMCli.findLast dbmy
        End If
        Me.grdComprobantes.AddItem modGrid.array2itemGrid(Array("Recibo", objMCli.apellido & ", " & objMCli.nombre, objMRec.serieID & "/" & objMRec.numero, Format(objMRec.total, "0.00")))
        curTotalDia = curTotalDia + objMRec.total
    Next
    Me.grdComprobantes.Redraw = True
    
    Me.txtTotal.Text = Format(curTotalDia, "0.00")
    
    Me.cmdImprimir.Enabled = True
    
End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub dtpFecha_Change()

    Me.grdComprobantes.Rows = 1
    Me.txtTotal.Text = ""
    Me.cmdImprimir.Enabled = False
    
End Sub

Private Sub Form_Load()

    Me.dtpFecha.Value = Date
    
    modGrid.makeGrid2 Me.grdComprobantes, Array(Array("Tipo", 1000), Array("Cliente", 3500), Array("Número", 1200), Array("Total", 1200)), 0, 1, flexSelectionByRow
    
    Me.cmdImprimir.Enabled = False
    
End Sub
