VERSION 5.00
Begin VB.Form frmPago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carga de Pagos"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7860
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   7860
   Begin VB.CommandButton fin 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6000
      TabIndex        =   12
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton anul 
      Caption         =   "&Anular"
      Height          =   375
      Left            =   6000
      TabIndex        =   6
      ToolTipText     =   "Elimina la fecha de PAGO cargada"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton conf 
      Caption         =   "&Confirmar"
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      ToolTipText     =   "Graba los DATOS del Pago de la FACTURA Activa"
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox fpag 
      Height          =   285
      Left            =   4080
      TabIndex        =   4
      Top             =   1560
      Width           =   1575
   End
   Begin VB.ComboBox ncli 
      Height          =   315
      Left            =   240
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   960
      Width           =   5415
   End
   Begin VB.ComboBox pdef 
      Height          =   315
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox facn 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox ncon 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label impor 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Importe 1er Vto."
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   13
      Top             =   1320
      Width           =   1125
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de PAGO"
      Height          =   195
      Index           =   3
      Left            =   4080
      TabIndex        =   11
      Top             =   1320
      Width           =   1170
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Nombre del CLIENTE"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   720
      Width           =   1530
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Período"
      Height          =   195
      Index           =   4
      Left            =   4080
      TabIndex        =   9
      Top             =   120
      Width           =   570
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Número de FACTURA"
      Height          =   195
      Index           =   0
      Left            =   2160
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Número de CONEXION"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   1665
   End
End
Attribute VB_Name = "frmPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sfac As Integer
Dim nfac As Long

Public Sub llenar()
Dim factura As New clsMyAFactura
Dim ncredito As New clsMyANCredito

Dim total As Currency

On Error Resume Next
    
    facn.Text = ""
    total = 0
    conf.Enabled = False
    anul.Enabled = False
    If factura.collectionAny(dbapp).Count = 0 Then
        MsgBox "No hay FACTURAS emitidas"
        Unload Me
        Exit Sub
    End If
    factura.clienteId = ncon.Text
    factura.periodoId = pdef.ItemData(pdef.ListIndex)
    factura.findByClientePeriodo dbapp
    If factura.autoID = 0 Then
        factura.clienteId = ncon.Text
        factura.periodoId = pdef.ItemData(pdef.ListIndex)
        factura.findByClientePeriodoPrev dbapp, True
        If factura.autoID = 0 Then Exit Sub
    End If
    sfac = factura.puntoVta
    nfac = factura.nroComprob
    facn.Text = Right("0000" & sfac, 4) & "-" & Right("00000000" & nfac, 8)
    
    total = factura.total
    For Each ncredito In ncredito.collectionByLiquidacion(factura.puntoVta, factura.nroComprob, dbapp)
        total = total - ncredito.total
    Next
    
    impor.Caption = Format(total, "#,###,##0.00")
    If factura.pagada Then
        fpag.Text = factura.fechapago
        anul.Enabled = True
        Exit Sub
    End If
    If factura.cancelada Then
        MsgBox "Esta FACTURA fue cancelada por Plan de Pago . . ."
        Exit Sub
    End If
    conf.Enabled = True
    
End Sub

Private Sub anul_Click()
Dim periodo As New clsRESTPeriodo
Dim factura As New clsMyAFactura

On Error Resume Next
    
    periodo.periodoId = pdef.ItemData(pdef.ListIndex)
    periodo.findByPrimaryKey
    
    factura.puntoVta = sfac
    factura.nroComprob = nfac
    factura.findByPrimaryKey dbapp
    factura.pagada = 0
    factura.fechapago = Null
    factura.puntoVtaInteres = Null
    factura.nroComprobInteres = Null
    factura.uid = "admin"
    factura.update dbapp
    
    anul.Enabled = False
    conf.Enabled = True
    fpag.Text = Date

End Sub

Private Sub conf_Click()
Dim periodo As New clsRESTPeriodo
Dim factura As New clsMyAFactura

On Error Resume Next
    
    If Not IsDate(fpag.Text) Then
        MsgBox "La Fecha de PAGO no es válida"
        Exit Sub
    End If
    
    periodo.periodoId = pdef.ItemData(pdef.ListIndex)
    periodo.findByPrimaryKey
    
    factura.puntoVta = sfac
    factura.nroComprob = nfac
    factura.findByPrimaryKey dbapp
    factura.pagada = 1
    factura.fechapago = CDate(fpag.Text)
    factura.puntoVtaInteres = Null
    factura.nroComprobInteres = Null
    If CDate(fpag.Text) <= periodo.fechaSegundo Then
        factura.puntoVtaInteres = 0
        factura.nroComprobInteres = 0
    End If
    factura.uid = "admin"
    factura.update dbapp
    
    conf.Enabled = False
    anul.Enabled = True

End Sub

Private Sub facn_GotFocus()

    Me.facn.SelStart = 0
    Me.facn.SelLength = Len(Me.facn.Text)
    
End Sub

Private Sub facn_KeyPress(KeyAscii As Integer)
Dim factura As New clsMyAFactura
Dim periodo As New clsRESTPeriodo

On Error Resume Next
    
    If KeyAscii = 13 Then
        factura.nroComprob = facn.Text
        factura.findByNroComprob dbapp
        If factura.autoID = 0 Then
            facn.Text = Right("0000" & sfac, 4) & "-" & Right("00000000" & nfac, 8)
            Exit Sub
        End If
        sfac = factura.puntoVta
        nfac = factura.nroComprob
        facn.Text = Right("0000" & sfac, 4) & "-" & Right("00000000" & nfac, 8)
        ncon.Text = factura.clienteId
        
        periodo.periodoId = factura.periodoId
        periodo.findByPrimaryKey
        
        pdef.Text = periodo.comboText
        
        llenar
    End If

End Sub

Private Sub fin_Click()
    
    Unload Me

End Sub

Private Sub Form_Activate()
Dim periodo As New clsRESTPeriodo
Dim clienteRep As New clsREPCliente

On Error Resume Next
    
    periodo.fillCombo Me.pdef
    clienteRep.fillCombo Me.ncli
    If Val(ncon.Text) Then llenar
    fpag.Text = Date

End Sub

Private Sub fpag_GotFocus()
    
    fpag.SelStart = 0
    fpag.SelLength = Len(fpag.Text)

End Sub

Private Sub fpag_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then conf.SetFocus

End Sub

Private Sub fpag_LostFocus()
    
    If Not IsDate(fpag.Text) Then MsgBox "La Fecha de PAGO no es válida"

End Sub

Private Sub ncli_Click()

On Error Resume Next
    
    ncon.Text = ncli.ItemData(ncli.ListIndex)
    
    llenar

End Sub

Private Sub ncon_GotFocus()
    
    ncon.SelStart = 0
    ncon.SelLength = Len(ncon.Text)

End Sub

Private Sub ncon_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then pdef.SetFocus

End Sub

Private Sub ncon_LostFocus()
Dim clienteId As Long

Dim cliente As clsMODCliente

Dim clienteRep As New clsREPCliente

On Error Resume Next
    
    Set cliente = clienteRep.findLastByClienteID(Val(Me.ncon.Text))
    If IsNull(cliente.uniqueId) Then
        ncli.ListIndex = 0
        ncon.Text = ncli.ItemData(ncli.ListIndex)
        Exit Sub
    End If
    If Not IsNull(cliente.fechaBaja) Then
        MsgBox "Cliente Dado de Baja . . ."
        ncli.ListIndex = 0
        ncon.Text = ncli.ItemData(ncli.ListIndex)
        Exit Sub
    End If
    clienteId = Val(ncon.Text)
    ncli.Text = cliente.comboText
    Do While ncli.ItemData(ncli.ListIndex) <> clienteId
        ncli.ListIndex = ncli.ListIndex + 1
    Loop
    
    llenar

End Sub

Private Sub pdef_Click()
    
    llenar

End Sub
