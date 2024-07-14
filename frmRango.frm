VERSION 5.00
Begin VB.Form frmRango 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rangos de Consumo"
   ClientHeight    =   5355
   ClientLeft      =   2670
   ClientTop       =   2745
   ClientWidth     =   5640
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   5640
   Begin VB.CommandButton fin 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      ToolTipText     =   "Fin de la TAREA"
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton elim 
      Caption         =   "&Eliminar Rango"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      ToolTipText     =   "Permite la Eliminación de un RANGO"
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton agre 
      Caption         =   "&Agregar Rango"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      ToolTipText     =   "Permite agregar un RANGO a la lista"
      Top             =   1680
      Width           =   1575
   End
   Begin VB.ComboBox rcat 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton modi 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      ToolTipText     =   "Permite modificar los DATOS del RANGO seleccionado"
      Top             =   1320
      Width           =   1575
   End
   Begin VB.ListBox rdef 
      Height          =   1815
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton hist 
      Caption         =   "&Ver Histórico"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      ToolTipText     =   "Muestra la Evolución del RANGO seleccionado"
      Top             =   960
      Width           =   1575
   End
   Begin VB.Frame vmod 
      Caption         =   "Modificación"
      Height          =   1695
      Left            =   240
      TabIndex        =   18
      Top             =   3120
      Width           =   5055
      Begin VB.CommandButton conf 
         Caption         =   "&Confirmar"
         Height          =   375
         Left            =   2640
         TabIndex        =   14
         ToolTipText     =   "Graba las modificaciones del RANGO"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox tari 
         Height          =   285
         Left            =   720
         TabIndex        =   13
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox lsup 
         Height          =   285
         Left            =   2640
         TabIndex        =   12
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox linf 
         Height          =   285
         Left            =   720
         TabIndex        =   11
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Tarifa ($)"
         Height          =   195
         Index           =   3
         Left            =   720
         TabIndex        =   21
         Top             =   960
         Width           =   630
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Límite Superior"
         Height          =   195
         Index           =   2
         Left            =   2640
         TabIndex        =   20
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Límite Inferior"
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   19
         Top             =   240
         Width           =   960
      End
   End
   Begin VB.Frame vhis 
      Caption         =   "Histórico"
      Height          =   2055
      Left            =   240
      TabIndex        =   17
      Top             =   3120
      Width           =   5055
      Begin VB.ListBox rhis 
         Height          =   1425
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   4575
      End
   End
   Begin VB.Frame vagr 
      Caption         =   "Agregar Rango"
      Height          =   1695
      Left            =   240
      TabIndex        =   22
      Top             =   3120
      Width           =   5055
      Begin VB.TextBox ninf 
         Height          =   285
         Left            =   720
         TabIndex        =   7
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox nsup 
         Height          =   285
         Left            =   2640
         TabIndex        =   8
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox ntari 
         Height          =   285
         Left            =   720
         TabIndex        =   9
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton acon 
         Caption         =   "&Confirmar"
         Height          =   375
         Left            =   2640
         TabIndex        =   10
         ToolTipText     =   "Graba los DATOS del nuevo RANGO definido"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Límite Inferior"
         Height          =   195
         Index           =   6
         Left            =   720
         TabIndex        =   25
         Top             =   240
         Width           =   960
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Límite Superior"
         Height          =   195
         Index           =   5
         Left            =   2640
         TabIndex        =   24
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Tarifa ($)"
         Height          =   195
         Index           =   4
         Left            =   720
         TabIndex        =   23
         Top             =   960
         Width           =   630
      End
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Categoría"
      Height          =   195
      Index           =   7
      Left            =   960
      TabIndex        =   26
      Top             =   120
      Width           =   705
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Rangos DEFINIDOS"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   16
      Top             =   720
      Width           =   1470
   End
End
Attribute VB_Name = "frmRango"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsin As Integer

Private Sub llenar()
Dim rangoID As Integer

Dim rango As New clsMyARango

On Error Resume Next

    rangoID = 1
    rdef.Clear
    Do
        With rango
            .categoria = rcat.ListIndex + 1
            .rangoID = rangoID
            .findLast dbapp
        
            If .autoID > 0 Then rdef.AddItem Right("00" & .rangoID, 2) & " - Rango : " & Right("000" & .limiteInferior, 3) & " - " & Right("000" & .limiteSuperior, 3) & " m3  Tarifa : $ " & Format(.tarifa, "#,###,##0.00")
        End With
        rangoID = rangoID + 1
    Loop While rango.autoID > 0

End Sub

Private Sub acon_Click()
Dim rango As New clsMyARango

On Error Resume Next
    
    With rango
        .categoria = rcat.ListIndex + 1
        .rangoID = rdef.ListCount + 1
        .fecha = Date
        .limiteInferior = ninf.Text
        .limiteSuperior = nsup.Text
        .tarifa = CDbl(ntari.Text)
        .uid = "admin"
        
        .save dbapp
    End With
    
    llenar
    
    vagr.Visible = False
    
    frmRango.Height = rsin

End Sub

Private Sub agre_Click()
Dim rangoID As Integer

Dim rango As New clsMyARango

On Error Resume Next
    
    rangoID = rdef.ListCount
    vhis.Visible = False
    vmod.Visible = False
    vagr.Visible = True
    frmRango.Height = rsin + 2000
    
    With rango
        .categoria = rcat.ListIndex + 1
        .rangoID = rangoID
        .findLast dbapp
        
        If .autoID = 0 Then
            ninf.Text = 0
            nsup.Text = 0
            ntari.Text = Format(0, "#,###,##0.00")
        Else
            ninf.Text = .limiteSuperior
            nsup.Text = .limiteSuperior + (.limiteSuperior - .limiteInferior)
            ntari.Text = Format(.tarifa, "#,###,##0.00")
        End If
    End With
    ninf.SetFocus
    
End Sub

Private Sub conf_Click()
Dim rango As New clsMyARango

On Error Resume Next
    
    With rango
        .categoria = rcat.ListIndex + 1
        .rangoID = rdef.ListIndex + 1
        .fecha = Date
        .limiteInferior = linf.Text
        .limiteSuperior = lsup.Text
        .tarifa = tari.Text
        .uid = "admin"
        
        .save dbapp
    End With
    
    llenar
    
    vmod.Visible = False
    frmRango.Height = rsin

End Sub

Private Sub elim_Click()
Dim rangoID As Integer

Dim rango As New clsMyARango
Dim rubro As New clsMyARubro

On Error Resume Next
    
    If rdef.ListIndex < 0 Then Exit Sub
    
    rangoID = rdef.ListIndex + 1
    If rubro.collectionByRangoID(rangoID, dbapp).Count > 0 Then
        MsgBox "Este RANGO esta utilizado por un Rubro"
        Exit Sub
    End If
    
    If rango.collectionByRangoID(rcat.ListIndex + 1, rangoID, dbapp).Count = 1 And rdef.ListIndex < rdef.ListCount - 1 Then Exit Sub
    
    With rango
        .categoria = rcat.ListIndex + 1
        .rangoID = rangoID
        .findLast dbapp
        
        .delete dbapp
    End With
    
    llenar
    
    frmRango.Height = rsin

End Sub

Private Sub fin_Click()
    
    frmRango.Height = rsin
    vhis.Visible = False
    vmod.Visible = False
    vagr.Visible = False
    
    Unload Me

End Sub

Private Sub Form_Activate()
    
    frmRango.Height = rsin
    vhis.Visible = False
    vmod.Visible = False
    vagr.Visible = False
    rcat.ListIndex = 0

End Sub

Private Sub Form_Load()
    
    rcat.Clear
    rcat.AddItem "General"
    rcat.AddItem "Especial"
    rsin = 3350
    llenar

End Sub

Private Sub hist_Click()
Dim rangoID As Integer

Dim rango As New clsMyARango

On Error Resume Next
    
    If rdef.ListIndex < 0 Then Exit Sub
    
    rangoID = rdef.ListIndex + 1
    
    frmRango.Height = rsin + 2400
    
    vhis.Visible = True
    vmod.Visible = False
    vagr.Visible = False
    rhis.Clear
    
    For Each rango In rango.collectionByRangoID(rcat.ListIndex + 1, rangoID, dbapp)
        With rango
            rhis.AddItem "Fecha " & .fecha & "  Rango : " & Right("000" & .limiteInferior, 3) & " - " & Right("000" & .limiteSuperior, 3) & " m3  Tarifa : $ " & Format(.tarifa, "#,###,##0.00")
        End With
    Next

End Sub

Private Sub linf_GotFocus()
    
    linf.SelStart = 0
    linf.SelLength = Len(linf.Text)

End Sub

Private Sub linf_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then lsup.SetFocus

End Sub

Private Sub lsup_GotFocus()
    
    lsup.SelStart = 0
    lsup.SelLength = Len(lsup.Text)

End Sub

Private Sub lsup_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then tari.SetFocus

End Sub

Private Sub modi_Click()
Dim rangoID As Integer

Dim rango As New clsMyARango

On Error Resume Next
    
    If rdef.ListIndex < 0 Then Exit Sub
    
    rangoID = rdef.ListIndex + 1
    
    vhis.Visible = False
    vagr.Visible = False
    vmod.Visible = True
    frmRango.Height = rsin + 2000
    
    With rango
        .categoria = rcat.ListIndex + 1
        .rangoID = rangoID
        .findLast dbapp
        linf.Text = .limiteInferior
        lsup.Text = .limiteSuperior
        tari.Text = Format(.tarifa, "#,###,##0.00")
    End With
    
    linf.SetFocus

End Sub

Private Sub ninf_GotFocus()
    
    ninf.SelStart = 0
    ninf.SelLength = Len(ninf.Text)

End Sub

Private Sub ninf_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then nsup.SetFocus

End Sub

Private Sub nsup_GotFocus()
    
    nsup.SelStart = 0
    nsup.SelLength = Len(nsup.Text)

End Sub

Private Sub nsup_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then ntari.SetFocus

End Sub

Private Sub ntari_GotFocus()
    
    ntari.SelStart = 0
    ntari.SelLength = Len(ntari.Text)

End Sub

Private Sub ntari_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then acon.SetFocus

End Sub

Private Sub ntari_LostFocus()
    
    If Not IsNumeric(ntari.Text) Then ntari.Text = 0
    ntari.Text = Format(CDbl(ntari.Text), "#,###,##0.00")

End Sub

Private Sub rcat_Click()
    
    llenar
    frmRango.Height = rsin

End Sub

Private Sub rdef_Click()
    
    vhis.Visible = False
    vmod.Visible = False
    vagr.Visible = False
    frmRango.Height = rsin

End Sub

Private Sub tari_GotFocus()
    
    tari.SelStart = 0
    tari.SelLength = Len(tari.Text)

End Sub

Private Sub tari_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then conf.SetFocus

End Sub

Private Sub tari_LostFocus()
    
    If Not IsNumeric(tari.Text) Then tari.Text = 0
    tari.Text = Format(CDbl(tari.Text), "#,###,##0.00")

End Sub
