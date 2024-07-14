VERSION 5.00
Begin VB.Form frmSuspFac 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suspensión y Reanudación de Facturación"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   6015
   Begin VB.TextBox txtCliente 
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   360
      Width           =   5535
   End
   Begin VB.CommandButton fin 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      ToolTipText     =   "Fin de la TAREA"
      Top             =   2040
      Width           =   1695
   End
   Begin VB.ComboBox prea 
      Height          =   315
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.ComboBox psus 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox motsus 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   3615
   End
   Begin VB.CommandButton anul 
      Caption         =   "&Anular Suspensión"
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      ToolTipText     =   "Anula la última suspensión"
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton rean 
      Caption         =   "&Reanudar"
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      ToolTipText     =   "Reanudar la Facturación de un Usuario"
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton susp 
      Caption         =   "&Suspender"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      ToolTipText     =   "Confirma la SUSPENSION del Usuario"
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Motivo Suspensión"
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   10
      Top             =   1560
      Width           =   1350
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Cliente"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   480
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Fin Suspensión"
      Height          =   195
      Index           =   3
      Left            =   2160
      TabIndex        =   8
      Top             =   840
      Width           =   1080
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Inicio Suspensión"
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   1245
   End
End
Attribute VB_Name = "frmSuspFac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cliente As New clsMODCliente

Private Sub fillForm()
Dim suspfactura As New clsMyASuspFactura

    With suspfactura
        .clienteId = cliente.clienteId
        .findLast dbapp
        
        If .autoID = 0 Then
            susp.Enabled = True
            rean.Enabled = False
            psus.Enabled = True
            psus.ListIndex = -1
            prea.ListIndex = -1
            prea.Enabled = False
            motsus.Enabled = True
            motsus.Text = ""
        Else
            If IsNull(.periodoIdfin) Then
                susp.Enabled = False
                If psus.ListCount > 0 Then psus.ListIndex = .periodoIDInicio - 1
                psus.Enabled = False
                motsus.Text = .motivo
                motsus.Enabled = False
                rean.Enabled = True
                prea.Enabled = True
                prea.ListIndex = -1
            Else
                susp.Enabled = True
                psus.Enabled = True
                psus.ListIndex = -1
                prea.ListIndex = -1
                rean.Enabled = False
                prea.Enabled = False
                motsus.Text = ""
                motsus.Enabled = True
            End If
        End If
    End With
    
    If psus.Enabled Then
        psus.SetFocus
    Else
        prea.SetFocus
    End If

End Sub

Private Sub anul_Click()
Dim suspfactura As New clsMyASuspFactura

On Error Resume Next
    
    With suspfactura
        .clienteId = cliente.clienteId
        .findLast dbapp
        
        .delete dbapp
    End With
    
    lcli_Click

End Sub

Private Sub lcli_Click()
Dim suspfactura As New clsMyASuspFactura

On Error Resume Next
    
    With suspfactura
        .clienteId = cliente.clienteId
        .findLast dbapp
        
        If .autoID = 0 Then
            susp.Enabled = True
            rean.Enabled = False
            psus.Enabled = True
            psus.ListIndex = -1
            prea.ListIndex = -1
            prea.Enabled = False
            motsus.Enabled = True
            motsus.Text = ""
        Else
            If IsNull(.periodoIdfin) Then
                susp.Enabled = False
                If psus.ListCount > 0 Then psus.ListIndex = .periodoIDInicio - 1
                psus.Enabled = False
                motsus.Text = .motivo
                motsus.Enabled = False
                rean.Enabled = True
                prea.Enabled = True
                prea.ListIndex = -1
            Else
                susp.Enabled = True
                psus.Enabled = True
                psus.ListIndex = -1
                prea.ListIndex = -1
                rean.Enabled = False
                prea.Enabled = False
                motsus.Text = ""
                motsus.Enabled = True
            End If
        End If
    End With
    
    If psus.Enabled Then
        psus.SetFocus
    Else
        prea.SetFocus
    End If
    
End Sub
Private Sub susp_Click()
Dim suspfactura As New clsMyASuspFactura

On Error Resume Next
    
    If Len(Trim(motsus.Text)) < 1 Then
        MsgBox "Debe poner un MOTIVO de Suspensión"
        motsus.SetFocus
        Exit Sub
    End If
    
    With suspfactura
        .clienteId = cliente.clienteId
        .periodoIDInicio = psus.ListIndex + 1
        .motivo = motsus.Text
        .uid = "admin"
        
        .add dbapp
    End With
    
    lcli_Click
    
    susp.Enabled = False
    rean.Enabled = True

End Sub

Private Sub fin_Click()
    
    Unload Me

End Sub

Private Sub Form_Activate()
Dim objMPer As New clsRESTPeriodo

    susp.Enabled = False
    rean.Enabled = False
    
    objMPer.fillCombo Me.psus
    objMPer.fillCombo Me.prea
    
End Sub

Private Sub motsus_GotFocus()
    
    motsus.SelStart = 0
    motsus.SelLength = Len(motsus.Text)

End Sub

Private Sub motsus_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then susp.SetFocus

End Sub

Private Sub rean_Click()
Dim suspfactura As New clsMyASuspFactura

On Error Resume Next
    
    With suspfactura
        .clienteId = cliente.clienteId
        .findLast dbapp
        If .autoID = 0 Then Exit Sub
        If prea.ItemData(prea.ListIndex) < psus.ItemData(psus.ListIndex) Then
            MsgBox "El Período de Finalización no puede ser menor al de Inicio"
            prea.SetFocus
            Exit Sub
        End If
        
        .periodoIdfin = prea.ListIndex + 1
        .uid = "admin"
        .update dbapp
    End With
    
    rean.Enabled = False
    susp.Enabled = True
    
    lcli_Click

End Sub

Private Sub txtCliente_GotFocus()

    marcarseleccion Me.txtCliente
    
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
Dim clienteRep As New clsREPCliente

    Set cliente = clienteRep.formSearch(frmBuscarRest, KeyAscii, "Clientes")
    
    Me.txtCliente.Text = cliente.textFound
    KeyAscii = 0
    
    fillForm
    
End Sub


