VERSION 5.00
Begin VB.Form frmOperador 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos del Operador"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   7440
   Begin VB.TextBox opevto 
      Height          =   285
      Left            =   2040
      TabIndex        =   14
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox opecai 
      Height          =   285
      Left            =   240
      TabIndex        =   13
      Top             =   3960
      Width           =   1575
   End
   Begin VB.ComboBox sitdgi 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox nroepas 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox opeing 
      Height          =   285
      Left            =   3840
      TabIndex        =   12
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox opecui 
      Height          =   285
      Left            =   2040
      TabIndex        =   11
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox opepro 
      Height          =   285
      Left            =   3840
      TabIndex        =   9
      Text            =   "Mendoza"
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox opeloc 
      Height          =   285
      Left            =   2040
      TabIndex        =   8
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox opetel 
      Height          =   285
      Left            =   3840
      TabIndex        =   6
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox opecod 
      Height          =   285
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox opedep 
      Height          =   285
      Left            =   2040
      TabIndex        =   5
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox opepis 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox openum 
      Height          =   285
      Left            =   3840
      TabIndex        =   3
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox opecal 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   3375
   End
   Begin VB.TextBox razsoc 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   5175
   End
   Begin VB.CommandButton fin 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5640
      TabIndex        =   16
      ToolTipText     =   "Fin de la TAREA"
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton actua 
      Caption         =   "&Actualizar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5640
      TabIndex        =   15
      ToolTipText     =   "Graba los DATOS cargados"
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Vto. CAI"
      Height          =   195
      Index           =   14
      Left            =   2040
      TabIndex        =   31
      Top             =   3720
      Width           =   585
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "CAI"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   30
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Situación DGI"
      Height          =   195
      Index           =   13
      Left            =   240
      TabIndex        =   29
      Top             =   3120
      Width           =   990
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Número EPAS"
      Height          =   195
      Index           =   12
      Left            =   240
      TabIndex        =   28
      Top             =   120
      Width           =   1020
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Número Ing. Brutos"
      Height          =   195
      Index           =   11
      Left            =   3840
      TabIndex        =   27
      Top             =   3120
      Width           =   1365
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Número CUIT"
      Height          =   195
      Index           =   10
      Left            =   2040
      TabIndex        =   26
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Provincia"
      Height          =   195
      Index           =   9
      Left            =   3840
      TabIndex        =   25
      Top             =   2520
      Width           =   660
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Localidad"
      Height          =   195
      Index           =   8
      Left            =   2040
      TabIndex        =   24
      Top             =   2520
      Width           =   690
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Teléfono"
      Height          =   195
      Index           =   7
      Left            =   3840
      TabIndex        =   23
      Top             =   1920
      Width           =   630
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Código Postal"
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   22
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Departamento"
      Height          =   195
      Index           =   5
      Left            =   2040
      TabIndex        =   21
      Top             =   1920
      Width           =   1005
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Piso"
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   20
      Top             =   1920
      Width           =   300
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Número"
      Height          =   195
      Index           =   3
      Left            =   3840
      TabIndex        =   19
      Top             =   1320
      Width           =   555
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Calle"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   18
      Top             =   1320
      Width           =   345
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Razón Social"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   17
      Top             =   720
      Width           =   945
   End
End
Attribute VB_Name = "frmOperador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private vari As Boolean

Private Sub actua_Click()
Dim cuit As String
Dim i As Integer

Dim operador As New clsMyAOperador

On Error Resume Next
    
    If sitdgi.ListIndex <> 2 And Len(opecui.Text) = 0 Then
        MsgBox "Debe ingresar el número de CUIT"
        opecui.SetFocus
        Exit Sub
    End If
    
    With operador
        .findLast dbapp
        .numeroEpas = Left(nroepas.Text, 10)
        .razonSocial = Left(razsoc.Text & " ", 100)
        .calle = Left(opecal.Text, 25)
        .puerta = Left(openum.Text, 5)
        .piso = Left(opepis.Text, 3)
        .dpto = Left(opedep.Text, 4)
        .codigoPostal = Val(opecod.Text)
        .telefono = Left(opetel.Text, 25)
        .localidad = Left(opeloc.Text, 30)
        .provincia = Left(opepro.Text, 15)
        .situacionIVA = sitdgi.ListIndex + 1
        cuit = ""
        For i = 1 To Len(opecui.Text)
            If Mid(opecui.Text, i, 1) <> "-" Then
                cuit = cuit & Mid(opecui.Text, i, 1)
            End If
        Next i
        i = 11
        If Len(cuit) < i Then i = Len(cuit)
        .cuit = Left(cuit, i)
        .ingresosBrutos = opeing.Text
        .cai = Left(opecai.Text, 25)
        If Not IsDate(opevto.Text) Then
            .caiVencimiento = Null
        Else
            .caiVencimiento = CDate(opevto.Text)
        End If
        .uid = "admin"
        
        .save dbapp
    End With
    
    vari = False
    actua.Enabled = False

End Sub

Private Sub fin_Click()

    Unload Me

End Sub

Private Sub Form_Activate()
Dim operador As New clsMyAOperador

On Error Resume Next
    
    With operador
        If operador.collectionAll(dbapp).Count = 0 Then
            sitdgi.ListIndex = 0
            Exit Sub
        End If
        .findLast dbapp
        If IsNull(.numeroEpas) Then
            nroepas.Text = ""
        Else
            nroepas.Text = .numeroEpas
        End If
        If IsNull(.razonSocial) Then
            razsoc.Text = ""
        Else
            razsoc.Text = .razonSocial
        End If
        If IsNull(.calle) Then
            opecal.Text = ""
        Else
            opecal.Text = .calle
        End If
        If IsNull(.puerta) Then
            openum.Text = ""
        Else
            openum.Text = .puerta
        End If
        If IsNull(.piso) Then
            opepis.Text = ""
        Else
            opepis.Text = .piso
        End If
        If IsNull(.dpto) Then
            opedep.Text = ""
        Else
            opedep.Text = .dpto
        End If
        If IsNull(.codigoPostal) Then
            opecod.Text = ""
        Else
            opecod.Text = .codigoPostal
        End If
        If IsNull(.telefono) Then
            opetel.Text = ""
        Else
            opetel.Text = .telefono
        End If
        If IsNull(.localidad) Then
            opeloc.Text = ""
        Else
            opeloc.Text = .localidad
        End If
        If IsNull(.provincia) Then
            opepro.Text = ""
        Else
            opepro.Text = .provincia
        End If
        sitdgi.ListIndex = .situacionIVA - 1
        If IsNull(.cuit) Then
            opecui.Text = ""
        Else
            If Len(.cuit) = 11 Then opecui.Text = Left(.cuit, 2) & "-" & Mid(.cuit, 3, 8) & "-" & Right(.cuit, 1)
        End If
        If IsNull(.ingresosBrutos) Then
            opeing.Text = ""
        Else
            opeing.Text = .ingresosBrutos
        End If
        If IsNull(.cai) Then
            opecai.Text = ""
        Else
            opecai.Text = .cai
        End If
        If IsNull(.caiVencimiento) Then
            opevto.Text = ""
        Else
            opevto.Text = .caiVencimiento
        End If
    End With
    actua.Enabled = False

End Sub

Private Sub Form_Load()

    sitdgi.Clear
    sitdgi.AddItem "Responsable Inscripto"
    sitdgi.AddItem "Responsable No Inscripto"
    sitdgi.AddItem "Consumidor Final"
    sitdgi.AddItem "iva Exento"
    sitdgi.AddItem "iva No Responsable"

End Sub

Private Sub nroepas_Change()
    
    vari = True
    actua.Enabled = True

End Sub

Private Sub nroepas_GotFocus()
    
    nroepas.SelStart = 0
    nroepas.SelLength = Len(nroepas.Text)

End Sub

Private Sub nroepas_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then razsoc.SetFocus

End Sub

Private Sub opecai_Change()
    
    If Len(opecai.Text) > 25 Then
        MsgBox "No debe ingresar más de 25 caracteres"
        opecai.Text = Left(opecai.Text, 25)
    End If
    vari = True
    actua.Enabled = True

End Sub

Private Sub opecai_GotFocus()
    
    opecai.SelStart = 0
    opecai.SelLength = Len(opecai.Text)

End Sub

Private Sub opecai_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then opecal.SetFocus

End Sub

Private Sub opecal_Change()
    
    If Len(opecal.Text) > 25 Then
        MsgBox "No debe ingresar más de 25 caracteres"
        opecal.Text = Left(opecal.Text, 25)
    End If
    vari = True
    actua.Enabled = True

End Sub

Private Sub opecal_GotFocus()
    
    opecal.SelStart = 0
    opecal.SelLength = Len(opecal.Text)

End Sub

Private Sub opecal_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then openum.SetFocus

End Sub

Private Sub opecod_Change()
    
    vari = True
    actua.Enabled = True

End Sub

Private Sub opecod_GotFocus()
    
    opecod.SelStart = 0
    opecod.SelLength = Len(opecod.Text)

End Sub

Private Sub opecod_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then opeloc.SetFocus

End Sub

Private Sub opecui_Change()
    
    vari = True
    actua.Enabled = True

End Sub

Private Sub opecui_GotFocus()
    
    opecui.SelStart = 0
    opecui.SelLength = Len(opecui.Text)

End Sub

Private Sub opecui_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then opeing.SetFocus

End Sub

Private Sub opecui_LostFocus()
    
    If Len(Trim(opecui.Text)) = 0 Then MsgBox "Debe ingresar el número de CUIT"

End Sub

Private Sub opedep_Change()
    
    vari = True
    actua.Enabled = True

End Sub

Private Sub opedep_GotFocus()
    
    opedep.SelStart = 0
    opedep.SelLength = Len(opedep.Text)

End Sub

Private Sub opedep_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then opetel.SetFocus

End Sub

Private Sub opeing_Change()
    
    vari = True
    actua.Enabled = True

End Sub

Private Sub opeing_GotFocus()
    
    opeing.SelStart = 0
    opeing.SelLength = Len(opeing.Text)

End Sub

Private Sub opeing_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        opecai.SetFocus
    End If

End Sub

Private Sub opeloc_Change()
    
    vari = True
    actua.Enabled = True

End Sub

Private Sub opeloc_GotFocus()
    
    opeloc.SelStart = 0
    opeloc.SelLength = Len(opeloc.Text)

End Sub

Private Sub opeloc_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then opepro.SetFocus

End Sub

Private Sub openum_Change()
    
    vari = True
    actua.Enabled = True

End Sub

Private Sub openum_GotFocus()
    
    openum.SelStart = 0
    openum.SelLength = Len(openum.Text)

End Sub

Private Sub openum_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then opepis.SetFocus

End Sub

Private Sub opepis_Change()
    
    vari = True
    actua.Enabled = True

End Sub

Private Sub opepis_GotFocus()
    
    opepis.SelStart = 0
    opepis.SelLength = Len(opepis.Text)

End Sub

Private Sub opepis_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then opedep.SetFocus

End Sub

Private Sub opepro_Change()
    
    vari = True
    actua.Enabled = True

End Sub

Private Sub opepro_GotFocus()
    
    opepro.SelStart = 0
    opepro.SelLength = Len(opepro.Text)

End Sub

Private Sub opepro_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then sitdgi.SetFocus

End Sub

Private Sub opetel_Change()
    
    vari = True
    actua.Enabled = True

End Sub

Private Sub opetel_GotFocus()
    
    opetel.SelStart = 0
    opetel.SelLength = Len(opetel.Text)

End Sub

Private Sub opetel_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then opecod.SetFocus

End Sub

Private Sub opevto_Change()
    
    vari = True
    actua.Enabled = True

End Sub

Private Sub opevto_GotFocus()
    
    opevto.SelStart = 0
    opevto.SelLength = Len(opevto.Text)

End Sub

Private Sub opevto_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then actua.SetFocus

End Sub

Private Sub razsoc_Change()
    
    vari = True
    actua.Enabled = True

End Sub

Private Sub razsoc_GotFocus()
    
    razsoc.SelStart = 0
    razsoc.SelLength = Len(razsoc.Text)

End Sub

Private Sub razsoc_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then opecal.SetFocus

End Sub

Private Sub sitdgi_Change()
    
    vari = True
    actua.Enabled = True

End Sub

Private Sub sitdgi_Click()
    
    If sitdgi.ListIndex = 2 Then
        opecui.Visible = False
    Else
        opecui.Visible = True
    End If

End Sub
