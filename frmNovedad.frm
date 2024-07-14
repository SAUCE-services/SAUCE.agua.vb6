VERSION 5.00
Begin VB.Form frmNovedad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carga de Novedades"
   ClientHeight    =   6420
   ClientLeft      =   2010
   ClientTop       =   2520
   ClientWidth     =   7935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   7935
   Begin VB.TextBox txtCliente 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5535
   End
   Begin VB.CommandButton fin 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      ToolTipText     =   "Fin de la TAREA"
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton rubc 
      Caption         =   "M&od. Rubro Común"
      Height          =   375
      Left            =   6000
      TabIndex        =   6
      ToolTipText     =   "Modifica el Importe de un RUBRO COMUN"
      Top             =   2160
      Width           =   1695
   End
   Begin VB.ComboBox pdef 
      Height          =   315
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton enov 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      ToolTipText     =   "Elimina la NOVEDAD Seleccionada"
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton mnov 
      Caption         =   "&Modificar Novedad"
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      ToolTipText     =   "Permite modificar algunos de los DATOS de la NOVEDAD"
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton anov 
      Caption         =   "&Agregar Novedad"
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      ToolTipText     =   "Permite agregar un NOVEDAD de Facturación para el Período elegido"
      Top             =   1080
      Width           =   1695
   End
   Begin VB.ListBox ndef 
      Columns         =   2
      Height          =   1815
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   5535
   End
   Begin VB.Frame nmod 
      Caption         =   "Modificar NOVEDADES"
      Height          =   3135
      Left            =   480
      TabIndex        =   25
      Top             =   3120
      Width           =   6975
      Begin VB.Frame mgcobr 
         Caption         =   "Cobro"
         Height          =   1455
         Left            =   360
         TabIndex        =   68
         Top             =   960
         Width           =   1575
         Begin VB.OptionButton muvez 
            Caption         =   "Una vez"
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton mnvec 
            Caption         =   "'N' veces"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton mindef 
            Caption         =   "Indefinido"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   1080
            Width           =   1095
         End
      End
      Begin VB.CommandButton mcon 
         Caption         =   "&Confirmar"
         Height          =   375
         Left            =   5040
         TabIndex        =   39
         ToolTipText     =   "Graba las modificaciones de la NOVEDAD"
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Frame mgind 
         Caption         =   "Indefinido"
         Height          =   1455
         Left            =   2160
         TabIndex        =   69
         Top             =   960
         Width           =   4455
         Begin VB.CommandButton mlim 
            Caption         =   ">"
            Height          =   255
            Left            =   4080
            TabIndex        =   38
            ToolTipText     =   "Elimina Fin de Cobro"
            Top             =   960
            Width           =   255
         End
         Begin VB.ComboBox midef 
            Height          =   315
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   960
            Width           =   1575
         End
         Begin VB.TextBox mican 
            Height          =   285
            Left            =   2520
            TabIndex        =   36
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox mipor 
            Height          =   285
            Left            =   360
            TabIndex        =   35
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad"
            Height          =   195
            Index           =   19
            Left            =   2520
            TabIndex        =   72
            Top             =   240
            Width           =   630
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Porcentaje (%)"
            Height          =   195
            Index           =   8
            Left            =   360
            TabIndex        =   71
            Top             =   240
            Width           =   1020
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Fin de Cobro"
            Height          =   195
            Index           =   7
            Left            =   1320
            TabIndex        =   70
            Top             =   960
            Width           =   900
         End
      End
      Begin VB.Frame mguna 
         Caption         =   "Una vez"
         Height          =   1455
         Left            =   2160
         TabIndex        =   78
         Top             =   960
         Width           =   4455
         Begin VB.TextBox mpor 
            Height          =   285
            Left            =   360
            TabIndex        =   31
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox mcan 
            Height          =   285
            Left            =   2520
            TabIndex        =   32
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Porcentaje (%)"
            Height          =   195
            Index           =   24
            Left            =   360
            TabIndex        =   80
            Top             =   360
            Width           =   1020
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad"
            Height          =   195
            Index           =   23
            Left            =   2520
            TabIndex        =   79
            Top             =   360
            Width           =   630
         End
      End
      Begin VB.Frame mgnve 
         Caption         =   "'N' veces"
         Height          =   1455
         Left            =   2160
         TabIndex        =   73
         Top             =   960
         Width           =   4455
         Begin VB.TextBox mveces 
            Height          =   285
            Left            =   360
            TabIndex        =   33
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox mtotal 
            Height          =   285
            Left            =   2520
            TabIndex        =   34
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Veces"
            Height          =   195
            Index           =   22
            Left            =   360
            TabIndex        =   77
            Top             =   240
            Width           =   450
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Importe Total"
            Height          =   195
            Index           =   21
            Left            =   2520
            TabIndex        =   76
            Top             =   240
            Width           =   930
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Pagadas"
            Height          =   195
            Index           =   20
            Left            =   1560
            TabIndex        =   75
            Top             =   960
            Width           =   630
         End
         Begin VB.Label mpaga 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2520
            TabIndex        =   74
            Top             =   960
            Width           =   1575
         End
      End
      Begin VB.Label mrub 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   480
         Width           =   6255
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Rubro"
         Height          =   195
         Index           =   9
         Left            =   360
         TabIndex        =   26
         Top             =   240
         Width           =   435
      End
   End
   Begin VB.Frame nagr 
      Caption         =   "Agregar NOVEDADES"
      Height          =   3135
      Left            =   480
      TabIndex        =   22
      Top             =   3120
      Width           =   6975
      Begin VB.Frame gcobr 
         Caption         =   "Cobro"
         Height          =   1455
         Left            =   360
         TabIndex        =   55
         Top             =   960
         Width           =   1575
         Begin VB.OptionButton aindef 
            Caption         =   "Indefinido"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   1080
            Width           =   1095
         End
         Begin VB.OptionButton anvec 
            Caption         =   "'N' veces"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton auvez 
            Caption         =   "Una vez"
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.CommandButton acon 
         Caption         =   "&Confirmar"
         Height          =   375
         Left            =   5040
         TabIndex        =   20
         ToolTipText     =   "Graba la NOVEDAD"
         Top             =   2640
         Width           =   1575
      End
      Begin VB.ComboBox arub 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   480
         Width           =   6255
      End
      Begin VB.Frame aguna 
         Caption         =   "Una vez"
         Height          =   1455
         Left            =   2160
         TabIndex        =   56
         Top             =   960
         Width           =   4455
         Begin VB.TextBox acan 
            Height          =   285
            Left            =   2520
            TabIndex        =   13
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox apor 
            Height          =   285
            Left            =   360
            TabIndex        =   12
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad"
            Height          =   195
            Index           =   6
            Left            =   2520
            TabIndex        =   58
            Top             =   360
            Width           =   630
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Porcentaje (%)"
            Height          =   195
            Index           =   5
            Left            =   360
            TabIndex        =   57
            Top             =   360
            Width           =   1020
         End
      End
      Begin VB.Frame agind 
         Caption         =   "Indefinido"
         Height          =   1455
         Left            =   2160
         TabIndex        =   64
         Top             =   960
         Width           =   4455
         Begin VB.TextBox aipor 
            Height          =   285
            Left            =   360
            TabIndex        =   16
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox aican 
            Height          =   285
            Left            =   2520
            TabIndex        =   17
            Top             =   480
            Width           =   1575
         End
         Begin VB.ComboBox aidef 
            Height          =   315
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   960
            Width           =   1575
         End
         Begin VB.CommandButton alim 
            Caption         =   ">"
            Height          =   255
            Left            =   4080
            TabIndex        =   19
            ToolTipText     =   "Elimina Fin de Cobro"
            Top             =   960
            Width           =   255
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Fin de Cobro"
            Height          =   195
            Index           =   18
            Left            =   1320
            TabIndex        =   67
            Top             =   960
            Width           =   900
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Porcentaje (%)"
            Height          =   195
            Index           =   17
            Left            =   360
            TabIndex        =   66
            Top             =   240
            Width           =   1020
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad"
            Height          =   195
            Index           =   16
            Left            =   2520
            TabIndex        =   65
            Top             =   240
            Width           =   630
         End
      End
      Begin VB.Frame agnve 
         Caption         =   "'N' veces"
         Height          =   1455
         Left            =   2160
         TabIndex        =   59
         Top             =   960
         Width           =   4455
         Begin VB.TextBox atotal 
            Height          =   285
            Left            =   2520
            TabIndex        =   15
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox aveces 
            Height          =   285
            Left            =   360
            TabIndex        =   14
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label apaga 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2520
            TabIndex        =   63
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Pagadas"
            Height          =   195
            Index           =   15
            Left            =   1560
            TabIndex        =   62
            Top             =   960
            Width           =   630
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Importe Total"
            Height          =   195
            Index           =   14
            Left            =   2520
            TabIndex        =   61
            Top             =   240
            Width           =   930
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Veces"
            Height          =   195
            Index           =   13
            Left            =   360
            TabIndex        =   60
            Top             =   240
            Width           =   450
         End
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Rubro"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   23
         Top             =   240
         Width           =   435
      End
   End
   Begin VB.Frame ncom 
      Caption         =   "Modificar RUBROS COMUNES"
      Height          =   3135
      Left            =   480
      TabIndex        =   53
      Top             =   3120
      Width           =   6975
      Begin VB.Frame rgcob 
         Caption         =   "Cobro"
         Height          =   1455
         Left            =   360
         TabIndex        =   81
         Top             =   960
         Width           =   1575
         Begin VB.OptionButton ruvez 
            Caption         =   "Una vez"
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton rnvec 
            Caption         =   "'N' veces"
            Height          =   255
            Left            =   240
            TabIndex        =   42
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton rindef 
            Caption         =   "Indefinido"
            Height          =   255
            Left            =   240
            TabIndex        =   43
            Top             =   1080
            Width           =   1095
         End
      End
      Begin VB.ComboBox rcom 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   480
         Width           =   6255
      End
      Begin VB.CommandButton rcon 
         Caption         =   "&Confirmar"
         Height          =   375
         Left            =   5040
         TabIndex        =   52
         ToolTipText     =   "Graba la NOVEDAD"
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Frame rgind 
         Caption         =   "Indefinido"
         Height          =   1455
         Left            =   2160
         TabIndex        =   82
         Top             =   960
         Width           =   4455
         Begin VB.CommandButton rlim 
            Caption         =   ">"
            Height          =   255
            Left            =   4080
            TabIndex        =   51
            ToolTipText     =   "Elimina Fin de Cobro"
            Top             =   960
            Width           =   255
         End
         Begin VB.ComboBox ridef 
            Height          =   315
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   960
            Width           =   1575
         End
         Begin VB.TextBox rican 
            Height          =   285
            Left            =   2520
            TabIndex        =   49
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox ripor 
            Height          =   285
            Left            =   360
            TabIndex        =   48
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad"
            Height          =   195
            Index           =   25
            Left            =   2520
            TabIndex        =   85
            Top             =   240
            Width           =   630
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Porcentaje (%)"
            Height          =   195
            Index           =   11
            Left            =   360
            TabIndex        =   84
            Top             =   240
            Width           =   1020
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Fin de Cobro"
            Height          =   195
            Index           =   10
            Left            =   1320
            TabIndex        =   83
            Top             =   960
            Width           =   900
         End
      End
      Begin VB.Frame rguna 
         Caption         =   "Una vez"
         Height          =   1455
         Left            =   2160
         TabIndex        =   91
         Top             =   960
         Width           =   4455
         Begin VB.TextBox rpor 
            Height          =   285
            Left            =   360
            TabIndex        =   44
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox rcan 
            Height          =   285
            Left            =   2520
            TabIndex        =   45
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Porcentaje (%)"
            Height          =   195
            Index           =   30
            Left            =   360
            TabIndex        =   93
            Top             =   360
            Width           =   1020
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad"
            Height          =   195
            Index           =   29
            Left            =   2520
            TabIndex        =   92
            Top             =   360
            Width           =   630
         End
      End
      Begin VB.Frame rgnve 
         Caption         =   "'N' veces"
         Height          =   1455
         Left            =   2160
         TabIndex        =   86
         Top             =   960
         Width           =   4455
         Begin VB.TextBox rveces 
            Height          =   285
            Left            =   360
            TabIndex        =   46
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox rtotal 
            Height          =   285
            Left            =   2520
            TabIndex        =   47
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Veces"
            Height          =   195
            Index           =   28
            Left            =   360
            TabIndex        =   90
            Top             =   240
            Width           =   450
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Importe Total"
            Height          =   195
            Index           =   27
            Left            =   2520
            TabIndex        =   89
            Top             =   240
            Width           =   930
         End
         Begin VB.Label etiq 
            AutoSize        =   -1  'True
            Caption         =   "Pagadas"
            Height          =   195
            Index           =   26
            Left            =   1560
            TabIndex        =   88
            Top             =   960
            Width           =   630
         End
         Begin VB.Label rpaga 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2520
            TabIndex        =   87
            Top             =   960
            Width           =   1575
         End
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Rubro"
         Height          =   195
         Index           =   12
         Left            =   360
         TabIndex        =   54
         Top             =   240
         Width           =   435
      End
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Cliente"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   94
      Top             =   120
      Width           =   480
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Período"
      Height          =   195
      Index           =   4
      Left            =   6000
      TabIndex        =   24
      Top             =   120
      Width           =   570
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Novedades CARGADAS"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   21
      Top             =   840
      Width           =   1755
   End
End
Attribute VB_Name = "frmNovedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsin As Integer

Private cliente As New clsMODCliente

Private Sub acan_GotFocus()
    
    marcarseleccion Me.acan

End Sub

Private Sub acan_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then acon.SetFocus

End Sub

Private Sub acon_Click()
Dim crit As String

Dim novedad As New clsMyANovedad

On Error Resume Next
    
    If arub.ListIndex < 0 Then Exit Sub
    If anvec.value Then
        If Val(aveces.Text) < 2 Then
            MsgBox "La cantidad de veces no puede ser menor que 2 . . ."
            aveces.SetFocus
            Exit Sub
        End If
        If Val(atotal.Text) = 0 Then
            MsgBox "Debe especificar un importe"
            atotal.SetFocus
            Exit Sub
        End If
    End If
    If aindef.value Then
        If Me.aidef.ItemData(Me.aidef.ListIndex) > -1 And Me.aidef.ItemData(Me.aidef.ListIndex) <= Me.pdef.ItemData(Me.pdef.ListIndex) Then
            MsgBox "El Período de suspensión debe ser posterior"
            aidef.SetFocus
            Exit Sub
        End If
    End If
    novedad.clienteId = cliente.clienteId
    novedad.periodoId = Me.pdef.ItemData(Me.pdef.ListIndex)
    novedad.rubroID = Val(Left(arub.Text, 2))
    novedad.findByPrimaryKey dbapp
    If novedad.autoID > 0 Then
        MsgBox "Repetido . . ."
        Exit Sub
    End If
    Set novedad = New clsMyANovedad
    novedad.clienteId = cliente.clienteId
    novedad.periodoId = Me.pdef.ItemData(Me.pdef.ListIndex)
    novedad.rubroID = Val(Left(arub.Text, 2))
    novedad.fecha = Date
    If auvez.value Then
        If apor.Text > 0 Then acan.Text = 0
        novedad.porcentaje = CDbl(apor.Text) / 100
        novedad.cantidad = acan.Text
        novedad.veces = 1
    End If
    If anvec.value Then
        novedad.veces = Val(aveces.Text)
        novedad.importe = CDbl(atotal.Text)
        novedad.vecesCobradas = 0
        novedad.cantidad = 1
    End If
    If aindef.value Then
        novedad.indefinida = 1
        If Me.aidef.ItemData(Me.aidef.ListIndex) > -1 Then novedad.periodoIdSuspension = Me.aidef.ItemData(Me.aidef.ListIndex)
        If aipor.Text > 0 Then aican.Text = 0
        novedad.porcentaje = CDbl(aipor.Text) / 100
        novedad.cantidad = aican.Text
    End If
    novedad.uid = "admin"
    
    novedad.add dbapp
    
    llenar
    
    frmNovedad.Height = rsin
    nagr.Visible = False

End Sub

Private Sub aindef_Click()
    
    aguna.Visible = False
    agind.Visible = True
    agnve.Visible = False

End Sub

Private Sub aipor_GotFocus()
    
    aipor.SelStart = 0
    aipor.SelLength = Len(aipor.Text)

End Sub

Private Sub aipor_LostFocus()
    
    If Not IsNumeric(aipor.Text) Then aipor.Text = "0"
    aipor.Text = Format(CDbl(aipor.Text), "#,###,##0.00")

End Sub

Private Sub alim_Click()
    
    aidef.ListIndex = -1

End Sub

Private Sub anov_Click()
Dim rubro As New clsMyARubro

Dim rubros As Collection

On Error Resume Next

    Set rubros = rubro.collectionSinRepeticion(dbapp)
    
    If rubros.Count = 0 Then
        Set rubros = Nothing
        Exit Sub
    End If

    nagr.Visible = True
    nmod.Visible = False
    ncom.Visible = False
    frmNovedad.Height = rsin + 3400
    arub.Clear
    
    For Each rubro In rubros
        If Not rubro.comunSocio And Not rubro.comun And Not rubro.desconectado Then arub.AddItem Right("00" & rubro.rubroID, 2) & " - " & rubro.concepto
    Next
    apor.Text = Format(0, "#,###,##0.00")
    acan.Text = 0
    aipor.Text = Format(0, "#,###,##0.00")
    aican.Text = 0
    aidef.ListIndex = -1
    aveces.Text = 0
    atotal.Text = Format(0, "#,###,##0.00")
    apaga.Caption = 0
    agind.Visible = False
    agnve.Visible = False
    aguna.Visible = True
    auvez.value = True
    arub.SetFocus

End Sub

Private Sub anvec_Click()
    
    aguna.Visible = False
    agind.Visible = False
    agnve.Visible = True

End Sub

Private Sub apor_GotFocus()
    
    apor.SelStart = 0
    apor.SelLength = Len(apor.Text)

End Sub

Private Sub apor_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then acan.SetFocus

End Sub

Private Sub apor_LostFocus()
    
    If Not IsNumeric(apor.Text) Then apor.Text = "0"
    apor.Text = Format(CDbl(apor.Text), "#,###,##0.00")

End Sub

Private Sub atotal_GotFocus()
    
    atotal.SelStart = 0
    atotal.SelLength = Len(atotal.Text)

End Sub

Private Sub atotal_LostFocus()
    
    If Not IsNumeric(atotal.Text) Then atotal.Text = "0"
    atotal.Text = Format(CDbl(atotal.Text), "#,###,##0.00")

End Sub

Private Sub auvez_Click()
    
    aguna.Visible = True
    agind.Visible = False
    agnve.Visible = False

End Sub

Private Sub aveces_GotFocus()
    aveces.SelStart = 0
    aveces.SelLength = Len(aveces.Text)

End Sub

Private Sub enov_Click()
Dim novedad As New clsMyANovedad

Dim ubic As Integer

On Error Resume Next
    
    If ndef.ListIndex < 0 Then Exit Sub
    
    ubic = InStr(ndef.List(ndef.ListIndex), " - ") + 3
    
    novedad.clienteId = cliente.clienteId
    novedad.periodoId = Me.pdef.ItemData(Me.pdef.ListIndex)
    novedad.rubroID = Val(Mid(ndef.List(ndef.ListIndex), ubic, 2))
    novedad.findByPrimaryKey dbapp
    
    If novedad.autoID > 0 Then novedad.delete dbapp
    
    llenar
    
    nagr.Visible = False
    nmod.Visible = False
    ncom.Visible = False
    frmNovedad.Height = rsin

End Sub

Private Sub fin_Click()
    
    Unload Me

End Sub

Private Sub Form_Activate()
Dim periodo As New clsRESTPeriodo

On Error Resume Next
    
    frmNovedad.Height = rsin
    nagr.Visible = False
    
    periodo.fillCombo Me.pdef
    periodo.fillCombo Me.aidef
    periodo.fillCombo Me.midef
    periodo.fillCombo Me.ridef
    
    llenar

End Sub

Private Sub Form_Load()
    
    rsin = 3475

End Sub

Public Sub llenar()
Dim rubro As New clsMyARubro
Dim novedad As New clsMyANovedad

On Error Resume Next

    ndef.Clear
    For Each novedad In novedad.collectionByClienteID(cliente.clienteId, dbapp, Me.pdef.ItemData(Me.pdef.ListIndex))
        rubro.rubroID = novedad.rubroID
        rubro.findLast dbapp
        ndef.AddItem novedad.fecha & " - " & Right("00" & novedad.rubroID, 2) & " - " & rubro.concepto
    Next

End Sub

Private Sub mcan_GotFocus()
    
    mcan.SelStart = 0
    mcan.SelLength = Len(mcan.Text)

End Sub

Private Sub mcan_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then mcon.SetFocus

End Sub

Private Sub mcon_Click()
Dim novedad As New clsMyANovedad

On Error Resume Next

    If mnvec.value Then
        If Val(mveces.Text) < 2 Then
            MsgBox "La cantidad de veces no puede ser menor que 2 . . ."
            mveces.SetFocus
            Exit Sub
        End If
        If Val(mveces.Text) < Val(mpaga.Caption) Then
            MsgBox "La cantidad de veces no puede ser menor que " & Val(mpaga.Caption) & ". . ."
            mveces.SetFocus
            Exit Sub
        End If
        If Val(mtotal.Text) = 0 Then
            MsgBox "Debe especificar un importe"
            mtotal.SetFocus
            Exit Sub
        End If
    End If
    If mindef.value Then
        If Me.midef.ItemData(Me.midef.ListIndex) > -1 And Me.midef.ItemData(Me.midef.ListIndex) <= Me.pdef.ItemData(Me.pdef.ListIndex) Then
            MsgBox "El Período de suspensión debe ser posterior"
            midef.SetFocus
            Exit Sub
        End If
    End If
    novedad.clienteId = cliente.clienteId
    novedad.periodoId = Me.pdef.ItemData(Me.pdef.ListIndex)
    novedad.rubroID = Val(Left(mrub.Caption, 2))
    novedad.findByPrimaryKey dbapp

    novedad.clienteId = cliente.clienteId
    novedad.periodoId = Me.pdef.ItemData(Me.pdef.ListIndex)
    novedad.rubroID = Val(Left(mrub.Caption, 2))
    novedad.fecha = Date

    If muvez.value Then
        If mpor.Text > 0 Then mcan.Text = 0
        novedad.porcentaje = CDbl(mpor.Text) / 100
        novedad.cantidad = mcan.Text
        novedad.veces = 1
        novedad.vecesCobradas = 0
        novedad.indefinida = False
        novedad.periodoIdSuspension = Null
        novedad.importe = 0
    End If
    If mnvec.value Then
        novedad.veces = Val(mveces.Text)
        novedad.importe = CDbl(mtotal.Text)
        novedad.porcentaje = 0
        novedad.cantidad = 1
        novedad.indefinida = False
        novedad.periodoIdSuspension = Null
    End If
    If mindef.value Then
        novedad.indefinida = True
        If Me.midef.ItemData(Me.midef.ListIndex) > -1 Then novedad.periodoIdSuspension = Me.midef.ItemData(Me.midef.ListIndex)
        If mipor.Text > 0 Then mican.Text = 0
        novedad.porcentaje = CDbl(mipor.Text) / 100
        novedad.cantidad = mican.Text
        novedad.veces = 1
        novedad.importe = 0
    End If
    novedad.uid = "admin"
    novedad.save dbapp
    
    llenar
    
    frmNovedad.Height = rsin
    nmod.Visible = False

End Sub

Private Sub mindef_Click()
    
    mguna.Visible = False
    mgnve.Visible = False
    mgind.Visible = True

End Sub

Private Sub mipor_GotFocus()
    
    mipor.SelStart = 0
    mipor.SelLength = Len(mipor.Text)

End Sub

Private Sub mipor_LostFocus()
    
    If Not IsNumeric(mipor.Text) Then mipor.Text = "0"
    mipor.Text = Format(CDbl(mipor.Text), "#,###,##0.00")

End Sub

Private Sub mlim_Click()
    
    midef.ListIndex = -1

End Sub

Private Sub mnov_Click()
Dim rube As Integer
Dim ubic As Long
Dim regi As Variant
Dim enc As Boolean

Dim rubro As New clsMyARubro
Dim novedad As New clsMyANovedad
Dim periodo As New clsRESTPeriodo

On Error Resume Next
    
    If ndef.ListIndex < 0 Then Exit Sub
    
    ubic = InStr(ndef.List(ndef.ListIndex), " - ") + 3
    rube = Val(Mid(ndef.List(ndef.ListIndex), ubic, 2))
    nmod.Visible = True
    nagr.Visible = False
    ncom.Visible = False
    frmNovedad.Height = rsin + 3400
    enc = False
    For Each rubro In rubro.collectionSinRepeticion(dbapp)
        If Not rubro.comun And Not rubro.comunSocio And Not rubro.desconectado Then
            If rubro.rubroID = rube Then
                enc = True
                mrub.Caption = Right("00" & rubro.rubroID, 2) & " - " & rubro.concepto
            End If
        End If
    Next
    If Not enc Then
        MsgBox "No corresponde a una novedad . . ."
        frmNovedad.Height = rsin
        nmod.Visible = False
        Exit Sub
    End If
    mpor.Text = Format(0, "#,###,##0.00")
    mcan.Text = 0
    mveces.Text = 0
    mtotal.Text = Format(0, "#,###,##0.00")
    mipor.Text = Format(0, "#,###,##0.00")
    mican.Text = 0
    midef.ListIndex = -1
    
    novedad.clienteId = cliente.clienteId
    novedad.periodoId = Me.pdef.ItemData(Me.pdef.ListIndex)
    novedad.rubroID = Val(Left(mrub.Caption, 2))
    novedad.findByPrimaryKey dbapp

    mguna.Visible = False
    mgnve.Visible = False
    mgind.Visible = False
    If novedad.indefinida Then
        mindef.value = True
        mgind.Visible = True
        mipor.Text = Format(novedad.porcentaje * 100, "#,###,##0.00")
        mican.Text = novedad.cantidad
        If Not IsNull(novedad.periodoIdSuspension) Then
            periodo.periodoId = novedad.periodoIdSuspension
            periodo.findByPrimaryKey
            midef.Text = periodo.comboText
        End If
        mipor.SetFocus
    End If
    If novedad.veces > 1 Then
        mnvec.value = True
        mgnve.Visible = True
        mveces.Text = novedad.veces
        mtotal.Text = Format(novedad.importe, "#,###,##0.00")
        mpaga.Caption = novedad.vecesCobradas
    End If
    If Not novedad.indefinida And novedad.veces < 2 Then
        muvez.value = True
        mguna.Visible = True
        mpor.Text = Format(novedad.porcentaje * 100, "#,###,##0.00")
        mcan.Text = novedad.cantidad
        mpor.SetFocus
    End If

End Sub

Private Sub mnvec_Click()
    
    mguna.Visible = False
    mgnve.Visible = True
    mgind.Visible = False

End Sub

Private Sub mpor_GotFocus()
    
    mpor.SelStart = 0
    mpor.SelLength = Len(mpor.Text)

End Sub

Private Sub mpor_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then mcan.SetFocus

End Sub

Private Sub mpor_LostFocus()
    
    If Not IsNumeric(mpor.Text) Then mpor.Text = 0
    mpor.Text = Format(CDbl(mpor.Text), "#,###,##0.00")

End Sub

Private Sub mtotal_GotFocus()
    
    mtotal.SelStart = 0
    mtotal.SelLength = Len(mtotal.Text)

End Sub

Private Sub mtotal_LostFocus()
    
    If Not IsNumeric(mtotal.Text) Then mtotal.Text = 0
    mtotal.Text = Format(CDbl(mtotal.Text), "#,###,##0.00")

End Sub

Private Sub muvez_Click()
    
    mguna.Visible = True
    mgnve.Visible = False
    mgind.Visible = False

End Sub

Private Sub mveces_GotFocus()
    
    mveces.SelStart = 0
    mveces.SelLength = Len(mveces.Text)

End Sub

Private Sub ndef_Click()
    
    frmNovedad.Height = rsin
    nagr.Visible = False
    ncom.Visible = False

End Sub

Private Sub pdef_Click()
    
    llenar

End Sub

Private Sub rcon_Click()
Dim crit As String
Dim novedad As New clsMyANovedad

On Error Resume Next
    
    If rcom.ListIndex < 0 Then Exit Sub
    If rnvec.value Then
        If Val(rveces.Text) < 2 Then
            MsgBox "La cantidad de veces no puede ser menor que 2 . . ."
            rveces.SetFocus
            Exit Sub
        End If
        If Val(rveces.Text) < Val(rpaga.Caption) Then
            MsgBox "La cantidad de veces no puede ser menor que " & Val(rpaga.Caption) & ". . ."
            mveces.SetFocus
            Exit Sub
        End If
        If Val(rtotal.Text) = 0 Then
            MsgBox "Debe especificar un importe"
            rtotal.SetFocus
            Exit Sub
        End If
    End If
    If rindef.value Then
        If Me.ridef.ItemData(Me.ridef.ListIndex) > -1 And Me.ridef.ItemData(Me.ridef.ListIndex) <= Me.pdef.ItemData(Me.pdef.ListIndex) Then
            MsgBox "El Período de suspensión debe ser posterior"
            ridef.SetFocus
            Exit Sub
        End If
    End If
    
    novedad.clienteId = cliente.clienteId
    novedad.periodoId = Me.pdef.ItemData(Me.pdef.ListIndex)
    novedad.rubroID = Left(rcom.Text, 2)
    novedad.fecha = Date
    If ruvez.value Then
        If rpor.Text > 0 Then rcan.Text = 0
        novedad.porcentaje = CDbl(rpor.Text) / 100
        novedad.cantidad = rcan.Text
        novedad.veces = 1
        novedad.vecesCobradas = 0
        novedad.indefinida = 0
        novedad.periodoIdSuspension = Null
        novedad.importe = 0
    End If
    If rnvec.value Then
        novedad.veces = Val(rveces.Text)
        novedad.importe = CDbl(rtotal.Text)
        novedad.porcentaje = 0
        novedad.cantidad = 1
        novedad.indefinida = 0
        novedad.periodoIdSuspension = Null
    End If
    If rindef.value Then
        novedad.indefinida = 1
        If Me.ridef.ItemData(Me.ridef.ListIndex) > -1 Then novedad.periodoIdSuspension = Me.ridef.ItemData(Me.ridef.ListIndex)
        If ripor.Text > 0 Then rican.Text = 0
        novedad.porcentaje = CDbl(ripor.Text) / 100
        novedad.cantidad = rican.Text
        novedad.veces = 1
        novedad.importe = 0
    End If
    novedad.uid = "admin"
    novedad.save dbapp
    
    llenar
    
    frmNovedad.Height = rsin
    ncom.Visible = False

End Sub

Private Sub rindef_Click()
    
    rguna.Visible = False
    rgnve.Visible = False
    rgind.Visible = True

End Sub

Private Sub ripor_GotFocus()
    
    ripor.SelStart = 0
    ripor.SelLength = Len(ripor.Text)

End Sub

Private Sub ripor_LostFocus()
    
    If Not IsNumeric(ripor.Text) Then ripor.Text = 0
    ripor.Text = Format(CDbl(ripor.Text), "#,###,##0.00")

End Sub

Private Sub rlim_Click()
    
    ridef.ListIndex = -1

End Sub

Private Sub rnvec_Click()
    
    rguna.Visible = False
    rgnve.Visible = True
    rgind.Visible = False

End Sub

Private Sub rpor_GotFocus()
    
    rpor.SelStart = 0
    rpor.SelLength = Len(rpor.Text)

End Sub

Private Sub rpor_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then rcan.SetFocus

End Sub

Private Sub rpor_LostFocus()
    
    If Not IsNumeric(rpor.Text) Then rpor.Text = 0
    rpor.Text = Format(CDbl(rpor.Text), "#,###,##0.00")

End Sub

Private Sub rcan_GotFocus()
    
    rcan.SelStart = 0
    rcan.SelLength = Len(rcan.Text)

End Sub

Private Sub rcan_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then rcon.SetFocus

End Sub

Private Sub rtotal_GotFocus()
    
    rtotal.SelStart = 0
    rtotal.SelLength = Len(rtotal.Text)

End Sub

Private Sub rtotal_LostFocus()
    
    If Not IsNumeric(rtotal.Text) Then rtotal.Text = 0
    rtotal.Text = Format(CDbl(rtotal.Text), "#,###,##0.00")

End Sub

Private Sub rubc_Click()
Dim cadx As String
Dim rube As Integer
Dim ubic As Integer
Dim enc As Boolean

Dim rubro As New clsMyARubro
Dim novedad As New clsMyANovedad
Dim periodo As New clsRESTPeriodo

Dim rubros As Collection

On Error Resume Next
    
    Set rubros = rubro.collectionSinRepeticion(dbapp)
    
    If rubros.Count = 0 Then
        Set rubros = Nothing
        Exit Sub
    End If
    
    If ndef.ListIndex > -1 Then
        ubic = InStr(ndef.List(ndef.ListIndex), " - ") + 3
        rube = Val(Mid(ndef.List(ndef.ListIndex), ubic, 2))
    Else
        rube = -1
    End If
    nagr.Visible = False
    nmod.Visible = False
    ncom.Visible = True
    frmNovedad.Height = rsin + 3400
    rcom.Clear
    enc = False
    For Each rubro In rubros
        If (rubro.comunSocio = 1 Or rubro.comun = True Or rubro.desconectado = 1) And rubro.rangoID = 0 Then
            rcom.AddItem Right("00" & rubro.rubroID, 2) & " - " & rubro.concepto
            If rubro.rubroID = rube Then
                enc = True
                cadx = Right("00" & rubro.rubroID, 2) & " - " & rubro.concepto
            End If
        End If
    Next
    If Not enc And rube > -1 Then
        MsgBox "No corresponde a un rubro común . . ."
        frmNovedad.Height = rsin
        ncom.Visible = False
        Exit Sub
    End If
    rpor.Text = Format(0, "#,###,##0.00")
    rcan.Text = 0
    rveces.Text = 0
    rtotal.Text = Format(0, "#,###,##0.00")
    ripor.Text = Format(0, "#,###,##0.00")
    rican.Text = 0
    ridef.ListIndex = -1
    rguna.Visible = True
    rgnve.Visible = False
    rgind.Visible = False
    ruvez.value = True
    rnvec.value = False
    rindef.value = False
    If rube < 0 Then
        rcom.SetFocus
        Exit Sub
    End If
    rcom.Text = cadx
    novedad.clienteId = cliente.clienteId
    novedad.periodoId = Me.pdef.ItemData(Me.pdef.ListIndex)
    novedad.rubroID = Val(Left(cadx, 2))
    novedad.findByPrimaryKey dbapp
    
    rguna.Visible = False
    rgnve.Visible = False
    rgind.Visible = False
    If novedad.indefinida Then
        rindef.value = True
        rgind.Visible = True
        ripor.Text = Format(novedad.porcentaje * 100, "#,###,##0.00")
        rican.Text = novedad.cantidad
        If Not IsNull(novedad.periodoIdSuspension) Then
            periodo.periodoId = novedad.periodoIdSuspension
            periodo.findByPrimaryKey
            ridef.Text = periodo.comboText
        End If
    End If
    If novedad.veces > 1 Then
        rnvec.value = True
        rgnve.Visible = True
        rveces.Text = novedad.veces
        rtotal.Text = Format(novedad.importe, "#,###,##0.00")
        rpaga.Caption = novedad.vecesCobradas
    End If
    If Not novedad.indefinida And novedad.veces < 2 Then
        ruvez.value = True
        rguna.Visible = True
        rpor.Text = Format(novedad.porcentaje * 100, "#,###,##0.00")
        rcan.Text = novedad.cantidad
    End If
    
    rcom.SetFocus

End Sub

Private Sub ruvez_Click()
    
    rguna.Visible = True
    rgnve.Visible = False
    rgind.Visible = False

End Sub

Private Sub rveces_GotFocus()
    
    rveces.SelStart = 0
    rveces.SelLength = Len(rveces.Text)

End Sub

Private Sub txtCliente_GotFocus()

    marcarseleccion Me.txtCliente
    
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
Dim clienteRep As New clsREPCliente

    Set cliente = clienteRep.formSearch(frmBuscarRest, KeyAscii, "Clientes")
    
    Me.txtCliente.Text = cliente.textFound
    KeyAscii = 0
    
    frmNovedad.Height = rsin
    
    llenar
    
End Sub


