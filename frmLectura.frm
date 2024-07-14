VERSION 5.00
Begin VB.Form frmLectura 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lecturas de MEDIDORES"
   ClientHeight    =   2640
   ClientLeft      =   3435
   ClientTop       =   3930
   ClientWidth     =   7815
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   7815
   Begin VB.TextBox txtCliente 
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   5535
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      ToolTipText     =   "Fin de la TAREA"
      Top             =   2040
      Width           =   1575
   End
   Begin VB.ComboBox pdef 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1560
      Width           =   3495
   End
   Begin VB.CommandButton elim 
      Caption         =   "&Eliminar última"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6000
      TabIndex        =   6
      ToolTipText     =   "Elimina la última LECTURA del Cliente ACTIVO"
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox emed 
      Height          =   285
      Left            =   3960
      TabIndex        =   3
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox flec 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox nmed 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton conf 
      Caption         =   "&Confirmar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      ToolTipText     =   "Graba los DATOS de la LECTURA"
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Período"
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   14
      Top             =   1320
      Width           =   570
   End
   Begin VB.Label eant 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3960
      TabIndex        =   8
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Estado ANTERIOR"
      Height          =   195
      Index           =   5
      Left            =   3960
      TabIndex        =   13
      Top             =   1920
      Width           =   1380
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Estado del MEDIDOR"
      Height          =   195
      Index           =   4
      Left            =   3960
      TabIndex        =   12
      Top             =   1320
      Width           =   1560
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de Lectura"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   1920
      Width           =   1260
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Cliente"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   480
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Número de MEDIDOR"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   720
      Width           =   1590
   End
End
Attribute VB_Name = "frmLectura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cliente As New clsMODCliente

Private Sub fillCliente()
Dim medidor As New clsMyAMedidor
Dim lectura As New clsMyALectura

Dim estado As Currency
    
    Me.txtCliente.Text = cliente.textFound
    
    medidor.clienteId = cliente.clienteId
    medidor.findColocadoByClienteID dbapp
    If medidor.autoID = 0 Then
        MsgBox "Cliente sin medidor"
        nmed.Text = ""
        emed.Text = ""
        eant.Caption = ""
        conf.Enabled = False
        elim.Enabled = False
        Exit Sub
    End If
    
    nmed.Text = medidor.medidorID
    estado = 0
    lectura.medidorID = Trim(nmed.Text)
    lectura.findLast dbapp
    If lectura.autoID > 0 Then
        estado = lectura.estado
    Else
        medidor.medidorID = Trim(nmed.Text)
        medidor.findByMedidorID dbapp
        If medidor.autoID > 0 Then estado = medidor.estadoInicio
    End If
    eant.Caption = estado
    emed.Text = estado

End Sub

Private Sub fillForm()
Dim clienteId As Long
Dim estado As Currency

Dim lectura As New clsMyALectura
Dim medidor As New clsMyAMedidor

Dim clienteRep As New clsREPCliente

On Error Resume Next

    estado = 0
    lectura.medidorID = Trim(nmed.Text)
    lectura.findLast dbapp
    If lectura.autoID > 0 Then
        estado = lectura.estado
    Else
        medidor.medidorID = Trim(nmed.Text)
        medidor.findByMedidorID dbapp
        If medidor.autoID > 0 Then estado = medidor.estadoInicio
    End If
    eant.Caption = estado
    emed.Text = estado

End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub conf_Click()
Dim medidor As New clsMyAMedidor
Dim factura As New clsMyAFactura
Dim lectura As New clsMyALectura

Dim clienteRep As New clsREPCliente

On Error Resume Next
    
    If Not IsDate(flec.Text) Then
        MsgBox "Fecha de LECTURA no válida"
        Exit Sub
    End If
    If Val(eant.Caption) > Val(emed.Text) Then
        MsgBox "Estado ACTUAL incorrecto"
        Exit Sub
    End If
    medidor.medidorID = Trim(nmed.Text)
    medidor.findByMedidorID dbapp
    If medidor.fechaColocacion > CDate(flec.Text) Then
        MsgBox "Fecha de LECTURA anterior a la Fecha de COLOCACIÓN"
        Exit Sub
    End If
    factura.clienteId = cliente.clienteId
    factura.periodoId = Me.pdef.ItemData(Me.pdef.ListIndex)
    factura.findByClientePeriodo dbapp
    If factura.autoID > 0 Then
        MsgBox "No puede modificar esta lectura, ya fue FACTURADA . . ."
        Exit Sub
    End If
    
    lectura.medidorID = Trim(Me.nmed.Text)
    lectura.periodoId = Me.pdef.ItemData(Me.pdef.ListIndex)
    lectura.fechaLectura = CDate(flec.Text)
    lectura.estado = Val(Me.emed.Text)
    lectura.uid = "admin"
    lectura.save dbapp
    
    eant.Caption = emed.Text
    
    Set cliente = clienteRep.findNextByClienteId(cliente.clienteId)
    fillCliente
    
    emed.SetFocus
    
End Sub

Private Sub elim_Click()
Dim factura As New clsMyAFactura
Dim lectura As New clsMyALectura

On Error Resume Next
    
    If Len(Trim(nmed.Text)) = 0 Then Exit Sub
    
    lectura.medidorID = Trim(Me.nmed.Text)
    lectura.periodoId = Me.pdef.ItemData(Me.pdef.ListIndex)
    lectura.findByPrimaryKey dbapp
    
    If lectura.autoID = 0 Then Exit Sub
    
    factura.clienteId = cliente.clienteId
    factura.periodoId = Me.pdef.ItemData(Me.pdef.ListIndex)
    factura.findByClientePeriodo dbapp
    
    If factura.autoID = 0 Then
        lectura.delete dbapp
    Else
        If factura.anulada <> 0 Then
            lectura.delete dbapp
        Else
            MsgBox "No puede eliminar : Período Facturado"
            Exit Sub
        End If
    End If
    fillForm
    pdef.SetFocus

End Sub

Private Sub emed_GotFocus()
    
    emed.SelStart = 0
    emed.SelLength = Len(emed.Text)

End Sub

Private Sub emed_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then flec.SetFocus

End Sub

Private Sub emed_LostFocus()
    
    conf.Enabled = True
    If Val(emed.Text) < Val(eant.Caption) Then conf.Enabled = False

End Sub

Private Sub flec_GotFocus()
    
    flec.SelStart = 0
    flec.SelLength = Len(flec.Text)

End Sub

Private Sub flec_KeyPress(KeyAscii As Integer)

On Error Resume Next
    
    If KeyAscii = 13 Then
        conf.SetFocus
        If Err.Number = 5 Then
            MsgBox "Revise la Lectura Cargada, no puedo confirmar . . ."
            emed.SetFocus
        End If
    End If

End Sub

Private Sub flec_LostFocus()
    
    If Not IsDate(flec.Text) Then MsgBox "La Fecha de LECTURA no es válida"

End Sub

Private Sub Form_Activate()
Dim periodo As New clsRESTPeriodo
Dim clienteRep As New clsREPCliente

On Error Resume Next
    
    periodo.fillCombo Me.pdef
    flec.Text = Date

End Sub

Private Sub nmed_GotFocus()
    
    nmed.SelStart = 0
    nmed.SelLength = Len(nmed.Text)

End Sub

Private Sub nmed_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then pdef.SetFocus

End Sub

Private Sub nmed_LostFocus()
Dim medidor As New clsMyAMedidor
Dim clienteRep As New clsREPCliente

On Error Resume Next
    
    medidor.medidorID = Trim(nmed.Text)
    medidor.findByMedidorID dbapp
    If medidor.autoID = 0 Then
        MsgBox "Medidor Inexistente"
        nmed.Text = ""
        emed.Text = ""
        eant.Caption = ""
        Exit Sub
    End If
    If Not IsNull(medidor.fechaRetiro) Then
        MsgBox "Medidor Retirado"
        nmed.Text = ""
        emed.Text = ""
        eant.Caption = ""
        Exit Sub
    End If
    If medidor.clienteId = 0 Then
        MsgBox "Medidor no asignado"
        nmed.Text = ""
        emed.Text = ""
        eant.Caption = ""
        Exit Sub
    End If
    Set cliente = clienteRep.findLastByClienteID(medidor.clienteId)
    Me.txtCliente.Text = cliente.textFound
    
    fillForm
    conf.Enabled = True
    elim.Enabled = True

End Sub

Private Sub txtCliente_GotFocus()

    marcarseleccion Me.txtCliente
    
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
Dim clienteRep As New clsREPCliente

    Set cliente = clienteRep.formSearch(frmBuscarRest, KeyAscii, "Clientes")
    KeyAscii = 0
    
    fillCliente
    
End Sub


