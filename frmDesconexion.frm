VERSION 5.00
Begin VB.Form frmDesconexion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Desconexión y Reconexión de Medidores"
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
      TabIndex        =   0
      Top             =   360
      Width           =   5535
   End
   Begin VB.TextBox motdes 
      Height          =   285
      Left            =   240
      TabIndex        =   11
      Top             =   2160
      Width           =   3615
   End
   Begin VB.TextBox nmed 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox fecrec 
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox fecdes 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   1695
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
   Begin VB.CommandButton reco 
      Caption         =   "&Reconectar"
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      ToolTipText     =   "Confirma la RECONEXION del Cliente"
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton desc 
      Caption         =   "&Desconectar"
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      ToolTipText     =   "Confirma la DESCONEXION del Cliente"
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Motivo Desconexión"
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   12
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Número de MEDIDOR"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   720
      Width           =   1590
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
      Caption         =   "Fecha Reconexión"
      Height          =   195
      Index           =   3
      Left            =   2160
      TabIndex        =   8
      Top             =   1320
      Width           =   1350
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Desconexión"
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   1425
   End
End
Attribute VB_Name = "frmDesconexion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cliente As New clsMODCliente

Private Sub desc_Click()
Dim desconexion As New clsMyADesconexion

On Error Resume Next
    
    If Len(Trim(motdes.Text)) < 1 Then
        MsgBox "Debe poner un MOTIVO de Desconexión"
        motdes.SetFocus
        Exit Sub
    End If
    If Not IsDate(fecdes.Text) Then
        MsgBox "La fecha de DESCONEXION no es válida"
        Exit Sub
    End If
        
    With desconexion
        .clienteId = cliente.clienteId
        .fechaDesconexion = fecdes.Text
        .motivo = motdes.Text
        .uid = "admin"
        
        .add dbapp
    End With
    
    desc.Enabled = False
    reco.Enabled = True

End Sub

Private Sub fecdes_GotFocus()
    
    fecdes.SelStart = 0
    fecdes.SelLength = Len(fecdes.Text)

End Sub

Private Sub fecdes_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then motdes.SetFocus

End Sub

Private Sub fecdes_LostFocus()
    
    If Not IsDate(fecdes.Text) Then MsgBox "La Fecha de DESCONEXION no es válida"

End Sub

Private Sub fecrec_GotFocus()
    
    fecrec.SelStart = 0
    fecrec.SelLength = Len(fecrec.Text)

End Sub

Private Sub fecrec_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then reco.SetFocus

End Sub

Private Sub fecrec_LostFocus()
    
    If Not IsDate(fecrec.Text) Then MsgBox "La Fecha de RECONEXION no es válida"

End Sub

Private Sub fin_Click()
    
    Unload Me

End Sub

Private Sub Form_Activate()
Dim medidor As New clsMyAMedidor

Dim clienteRep As New clsREPCliente

On Error Resume Next
    
    desc.Enabled = False
    reco.Enabled = False
    
    If medidor.collectionAny(dbapp).Count = 0 Then MsgBox "No hay información de Medidores"
    
End Sub

Private Sub motdes_GotFocus()
    
    motdes.SelStart = 0
    motdes.SelLength = Len(motdes.Text)

End Sub

Private Sub motdes_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then desc.SetFocus

End Sub

Private Sub nmed_GotFocus()
    
    nmed.SelStart = 0
    nmed.SelLength = Len(nmed.Text)

End Sub

Private Sub nmed_KeyPress(KeyAscii As Integer)

On Error Resume Next
    
    If KeyAscii = 13 Then
        If fecdes.Enabled Then
            fecdes.SetFocus
        Else
            fecrec.SetFocus
        End If
    End If

End Sub

Private Sub nmed_LostFocus()
Dim medidor As New clsMyAMedidor
Dim desconexion As New clsMyADesconexion

Dim clienteRep As New clsREPCliente

On Error Resume Next
    
    With medidor
        If .collectionAny(dbapp).Count = 0 Then
            MsgBox "No hay información sobre Medidores"
            Exit Sub
        End If
        .medidorID = Trim(nmed.Text)
        .findLast dbapp
        If .autoID = 0 Then
            MsgBox "No hay información sobre este Medidor . . ."
            Exit Sub
        End If
        If Not IsNull(.fechaRetiro) Then
            MsgBox "Este medidor fue Retirado"
            Exit Sub
        End If
        If IsNull(.fechaColocacion) Then
            MsgBox "Este medidor no está Asignado"
            Exit Sub
        End If
        If IsNull(.clienteId) Then
            MsgBox "Este medidor no está Asignado"
            Exit Sub
        End If
    End With
        
    Set cliente = clienteRep.findLastByClienteID(medidor.clienteId)
    Me.txtCliente.Text = cliente.textFound
    
    With desconexion
        .clienteId = cliente.clienteId
        .findLast dbapp
        If .autoID = 0 Then
            desc.Enabled = True
            reco.Enabled = False
            fecdes.Enabled = True
            fecdes.Text = Date
            fecrec.Enabled = False
        Else
            If IsNull(.fechaReconexion) Then
                desc.Enabled = False
                fecdes.Text = .fechaDesconexion
                fecdes.Enabled = False
                reco.Enabled = True
                fecrec.Enabled = True
                fecrec.Text = Date
            Else
                desc.Enabled = True
                fecdes.Enabled = True
                fecdes.Text = Date
                reco.Enabled = False
                fecrec.Enabled = False
            End If
        End If
    End With

End Sub

Private Sub reco_Click()
Dim desconexion As New clsMyADesconexion

On Error Resume Next
    
    With desconexion
        If .collectionAny(dbapp).Count = 0 Then Exit Sub
        If Not IsDate(fecrec.Text) Then
            MsgBox "La Fecha de RECONEXION no es válida"
            Exit Sub
        End If
        .clienteId = cliente.clienteId
        .findLast dbapp
        
        If .autoID = 0 Then Exit Sub
        
        .fechaReconexion = fecrec.Text
        .uid = "admin"
        .update dbapp
    End With
    
    reco.Enabled = False
    desc.Enabled = True

End Sub

Private Sub txtCliente_GotFocus()

    marcarseleccion Me.txtCliente
    
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
Dim clienteRep As New clsREPCliente
Dim medidor As New clsMyAMedidor
Dim desconexion As New clsMyADesconexion

    Set cliente = clienteRep.formSearch(frmBuscarRest, KeyAscii, "Clientes")
    
    Me.txtCliente.Text = cliente.textFound
    KeyAscii = 0

    Set cliente = clienteRep.formSearch(frmBuscarRest, KeyAscii, "Clientes")
    
    Me.txtCliente.Text = cliente.textFound
    KeyAscii = 0
    
    fecdes.Text = ""
    fecrec.Text = ""
    motdes.Text = ""
    With medidor
        If medidor.collectionAny(dbapp).Count > 0 Then
            .clienteId = cliente.clienteId
            .findLastByClienteID dbapp
            If .autoID > 0 Then
                If Not IsNull(.fechaRetiro) Then
                    MsgBox "Este medidor fue Retirado"
                    Exit Sub
                End If
                If IsNull(.fechaColocacion) Then
                    MsgBox "Este medidor no está Asignado"
                    Exit Sub
                End If
                nmed.Text = .medidorID
            End If
        End If
    End With
    
    With desconexion
        .clienteId = cliente.clienteId
        .findLast dbapp
        If .autoID = 0 Then
            desc.Enabled = True
            reco.Enabled = False
            fecdes.Enabled = True
            fecdes.Text = Date
            fecrec.Enabled = False
        Else
            If IsNull(.fechaReconexion) Then
                desc.Enabled = False
                fecdes.Text = .fechaDesconexion
                fecdes.Enabled = False
                reco.Enabled = True
                fecrec.Enabled = True
                fecrec.Text = Date
            Else
                desc.Enabled = True
                fecdes.Enabled = True
                fecdes.Text = Date
                reco.Enabled = False
                fecrec.Enabled = False
            End If
        End If
    End With

End Sub


