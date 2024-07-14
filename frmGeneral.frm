VERSION 5.00
Begin VB.Form frmGeneral 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Información General"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   6270
   Begin VB.CheckBox preimp 
      Caption         =   "Factura Preimpresa"
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox alicf 
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin VB.Frame susp 
      Caption         =   "Corte y Restricción"
      Height          =   975
      Left            =   240
      TabIndex        =   25
      Top             =   4440
      Width           =   3975
      Begin VB.TextBox perso 
         Height          =   285
         Left            =   2160
         TabIndex        =   14
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox resol 
         Height          =   285
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "N. Personería"
         Height          =   195
         Index           =   8
         Left            =   2160
         TabIndex        =   27
         Top             =   240
         Width           =   990
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Resolución"
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.ComboBox perfa 
      Height          =   315
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1080
      Width           =   1575
   End
   Begin VB.ComboBox servi 
      Height          =   315
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Frame gral 
      Caption         =   "Facturación"
      Height          =   2775
      Left            =   240
      TabIndex        =   19
      Top             =   1560
      Width           =   3975
      Begin VB.TextBox ncrini 
         Height          =   285
         Left            =   2160
         TabIndex        =   12
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox ncrneg 
         Height          =   285
         Left            =   240
         TabIndex        =   11
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox recneg 
         Height          =   285
         Left            =   240
         TabIndex        =   9
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox recini 
         Height          =   285
         Left            =   2160
         TabIndex        =   10
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox facini 
         Height          =   285
         Left            =   2160
         TabIndex        =   8
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox unineg 
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox fecini 
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "N. Crédito Inicial"
         Height          =   195
         Index           =   13
         Left            =   2160
         TabIndex        =   32
         Top             =   2040
         Width           =   1155
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Unidad de Negocio"
         Height          =   195
         Index           =   12
         Left            =   240
         TabIndex        =   31
         Top             =   2040
         Width           =   1380
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Unidad de Negocio"
         Height          =   195
         Index           =   11
         Left            =   240
         TabIndex        =   30
         Top             =   1440
         Width           =   1380
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Recibo Inicial"
         Height          =   195
         Index           =   10
         Left            =   2160
         TabIndex        =   29
         Top             =   1440
         Width           =   960
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Factura Inicial"
         Height          =   195
         Index           =   4
         Left            =   2160
         TabIndex        =   22
         Top             =   840
         Width           =   990
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Unidad de Negocio"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   21
         Top             =   840
         Width           =   1380
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Inicio de Transición"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   1380
      End
   End
   Begin VB.CommandButton fin 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4440
      TabIndex        =   16
      ToolTipText     =   "Fin de la TAREA"
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton actua 
      Caption         =   "&Actualizar"
      Height          =   375
      Left            =   4440
      TabIndex        =   15
      ToolTipText     =   "Graba los DATOS cargados"
      Top             =   4680
      Width           =   1575
   End
   Begin VB.TextBox alirni 
      Height          =   285
      Left            =   4320
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox aliiva 
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "IVA (%) Cons. Final"
      Height          =   195
      Index           =   9
      Left            =   480
      TabIndex        =   28
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Período Facturación"
      Height          =   195
      Index           =   6
      Left            =   2400
      TabIndex        =   24
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Servicio Prestado"
      Height          =   195
      Index           =   5
      Left            =   480
      TabIndex        =   23
      Top             =   840
      Width           =   1245
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "IVA (%) No Inscripto"
      Height          =   195
      Index           =   1
      Left            =   4320
      TabIndex        =   18
      Top             =   120
      Width           =   1410
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "IVA (%) Resp. Inscripto"
      Height          =   195
      Index           =   0
      Left            =   2400
      TabIndex        =   17
      Top             =   120
      Width           =   1620
   End
End
Attribute VB_Name = "frmGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub actua_Click()
Dim alicuota As New clsMyAAlicuota
Dim operador As New clsMyAOperador

On Error Resume Next
    
    If Not IsDate(fecini.Text) Then
        MsgBox "La Fecha de INICIO de ACTIVIDADES no es válida"
        Exit Sub
    End If
    If Len(Trim(aliiva.Text)) = 0 Then aliiva.Text = Format(0, "#,###,##0.00")
    If Len(Trim(alirni.Text)) = 0 Then alirni.Text = Format(0, "#,###,##0.00")
    If Len(Trim(alicf.Text)) = 0 Then alicf.Text = Format(0, "#,###,##0.00")
    
    With alicuota
        .ivaCF = CDbl(alicf.Text) / 100
        .IVA = CDbl(aliiva.Text) / 100
        .rni = CDbl(alirni.Text) / 100
        .fecha = Date
        .uid = "admin"
        
        .save dbapp
    End With
    
    If operador.collectionAll(dbapp).Count = 0 Then
        MsgBox "No hay información del operador"
        Exit Sub
    End If
    
    With operador
        .findLast dbapp
        .fechaInicio = CDate(fecini.Text)
        .puntoVta = Val(unineg.Text)
        .nroComprob = Val(facini.Text)
        .reciboSerie = Val(recneg.Text)
        .recibo = Val(recini.Text)
        .ncreditoSerie = Val(ncrneg.Text)
        .ncredito = Val(ncrini.Text)
        .servicio = servi.ListIndex + 1
        .periodoFactura = perfa.ListIndex + 1
        .resolucion = Left(resol.Text, 10)
        .personeria = Left(perso.Text, 10)
        .uid = "admin"
        .preimpreso = False
        If preimp.Value = 1 Then .preimpreso = True
        .save dbapp
    End With

End Sub

Private Sub alicf_GotFocus()
    
    alicf.SelStart = 0
    alicf.SelLength = Len(alicf.Text)

End Sub

Private Sub alicf_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then aliiva.SetFocus

End Sub

Private Sub alicf_LostFocus()
    
    If Not IsNumeric(alicf.Text) Then alicf.Text = "0"
    alicf.Text = Format(CDbl(alicf.Text), "#,###,##0.00")

End Sub

Private Sub aliiva_GotFocus()
    
    aliiva.SelStart = 0
    aliiva.SelLength = Len(aliiva.Text)

End Sub

Private Sub aliiva_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then alirni.SetFocus

End Sub

Private Sub aliiva_LostFocus()
    
    If Not IsNumeric(aliiva.Text) Then aliiva.Text = "0"
    aliiva.Text = Format(CDbl(aliiva.Text), "#,###,##0.00")

End Sub

Private Sub alirni_GotFocus()
    
    alirni.SelStart = 0
    alirni.SelLength = Len(alirni.Text)

End Sub

Private Sub alirni_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then servi.SetFocus

End Sub

Private Sub alirni_LostFocus()
    
    If Not IsNumeric(alirni.Text) Then alirni.Text = "0"
    alirni.Text = Format(CDbl(alirni.Text), "#,###,##0.00")

End Sub

Private Sub facini_GotFocus()
    
    facini.SelStart = 0
    facini.SelLength = Len(facini.Text)

End Sub

Private Sub facini_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then resol.SetFocus

End Sub

Private Sub fecini_GotFocus()
    
    fecini.SelStart = 0
    fecini.SelLength = Len(fecini.Text)

End Sub

Private Sub fecini_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then unineg.SetFocus

End Sub

Private Sub fecini_LostFocus()
    
    If Not IsDate(fecini.Text) Then MsgBox "La Fecha de INICIO de ACTIVIDADES no es válida"

End Sub

Private Sub fin_Click()
    
    Unload Me

End Sub

Private Sub Form_Activate()
Dim alicuota As New clsMyAAlicuota
Dim operador As New clsMyAOperador

On Error Resume Next
    
    With alicuota
        If .collectionAny(dbapp).Count > 0 Then
            .findLast dbapp
            If IsNull(.ivaCF) Then
                alicf.Text = Format(0, "#,###,##0.00")
            Else
                alicf.Text = Format(.ivaCF * 100, "#,###,##0.00")
            End If
            aliiva.Text = Format(.IVA * 100, "#,###,##0.00")
            alirni.Text = Format(.rni * 100, "#,###,##0.00")
        End If
    End With
    With operador
        If .collectionAll(dbapp).Count > 0 Then
            .findLast dbapp
            servi.ListIndex = .servicio - 1
            perfa.ListIndex = .periodoFactura - 1
            fecini.Text = Date
            unineg.Text = ""
            facini.Text = ""
            recneg.Text = ""
            recini.Text = ""
            ncrneg.Text = ""
            ncrini.Text = ""
            resol.Text = ""
            perso.Text = ""
            preimp.Value = 0
            If Not IsNull(.fechaInicio) Then fecini.Text = .fechaInicio
            If Not IsNull(.puntoVta) Then unineg.Text = .puntoVta
            If Not IsNull(.nroComprob) Then facini.Text = .nroComprob
            If Not IsNull(.reciboSerie) Then recneg.Text = .reciboSerie
            If Not IsNull(.recibo) Then recini.Text = .recibo
            If Not IsNull(.ncreditoSerie) Then ncrneg.Text = .ncreditoSerie
            If Not IsNull(.ncredito) Then ncrini.Text = .ncredito
            If Not IsNull(.resolucion) Then resol.Text = .resolucion
            If Not IsNull(.personeria) Then perso.Text = .personeria
            If .preimpreso Then preimp.Value = 1
        Else
            servi.ListIndex = 2
            perfa.ListIndex = 0
            fecini.Text = Date
        End If
    End With
    alicf.SetFocus

End Sub

Private Sub Form_Load()
Dim operador As New clsMyAOperador
Dim alicuota As New clsMyAAlicuota

    servi.Clear
    servi.AddItem "Agua"
    servi.AddItem "Cloaca"
    servi.AddItem "Agua y Cloaca"
    perfa.Clear
    perfa.AddItem "Bimestral"
    perfa.AddItem "Mensual"

End Sub

Private Sub ncrini_GotFocus()
    
    ncrini.SelStart = 0
    ncrini.SelLength = Len(ncrini.Text)

End Sub

Private Sub ncrini_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then resol.SetFocus

End Sub

Private Sub perso_GotFocus()
    
    perso.SelStart = 0
    perso.SelLength = Len(perso.Text)

End Sub

Private Sub perso_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then actua.SetFocus

End Sub

Private Sub recini_GotFocus()
    
    recini.SelStart = 0
    recini.SelLength = Len(recini.Text)

End Sub

Private Sub recini_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then ncrneg.SetFocus

End Sub

Private Sub recneg_GotFocus()
    
    recneg.SelStart = 0
    recneg.SelLength = Len(recneg.Text)

End Sub

Private Sub recneg_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then recini.SetFocus

End Sub

Private Sub resol_gotfocus()
    
    resol.SelStart = 0
    resol.SelLength = Len(resol.Text)

End Sub

Private Sub resol_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then perso.SetFocus

End Sub

Private Sub unineg_GotFocus()
    
    unineg.SelStart = 0
    unineg.SelLength = Len(unineg.Text)

End Sub

Private Sub unineg_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then facini.SetFocus

End Sub
