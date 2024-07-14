VERSION 5.00
Begin VB.Form frmImpresora 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Imprimir"
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optImpresora 
      Caption         =   "I&mpresora"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   720
      Width           =   1335
   End
   Begin VB.OptionButton optPantalla 
      Caption         =   "Pan&talla"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   240
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5160
      TabIndex        =   18
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   5160
      TabIndex        =   17
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame fraCopias 
      Caption         =   "Copias"
      Height          =   1215
      Left            =   3720
      TabIndex        =   14
      Top             =   3000
      Width           =   2535
      Begin VB.CheckBox chkIntercalar 
         Caption         =   "&Intercalar"
         Height          =   255
         Left            =   1320
         TabIndex        =   19
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtCopias 
         Height          =   315
         Left            =   1680
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Número de &copias:"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame fraIntervalo 
      Caption         =   "Intervalo de ìmpresión"
      Height          =   1215
      Left            =   240
      TabIndex        =   7
      Top             =   3000
      Width           =   3375
      Begin VB.TextBox txtHasta 
         Height          =   315
         Left            =   2520
         TabIndex        =   13
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtDesde 
         Height          =   315
         Left            =   1440
         TabIndex        =   11
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton optPaginas 
         Caption         =   "&Páginas"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton optTodo 
         Caption         =   "&Todo"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "&a"
         Height          =   195
         Index           =   4
         Left            =   2280
         TabIndex        =   12
         Top             =   720
         Width           =   90
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "&de"
         Height          =   195
         Index           =   1
         Left            =   1200
         TabIndex        =   10
         Top             =   720
         Width           =   180
      End
   End
   Begin VB.Frame fraImpresora 
      Caption         =   "Impresora"
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   6015
      Begin VB.ComboBox cboImpresoras 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label lblTipo 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         Height          =   195
         Left            =   1080
         TabIndex        =   6
         Top             =   840
         Width           =   315
      End
      Begin VB.Label lblUbicacion 
         AutoSize        =   -1  'True
         Caption         =   "Ubicacion"
         Height          =   195
         Left            =   1080
         TabIndex        =   5
         Top             =   1200
         Width           =   720
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Ubicación:"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   765
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   360
      End
      Begin VB.Label etiq 
         AutoSize        =   -1  'True
         Caption         =   "&Nombre:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmImpresora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private vCopies As Integer
Private vMin As Integer
Private vMax As Integer
Private vFromPage As Integer
Private vToPage As Integer
Private vCollate As Boolean
Private vDefaultPrinter As Integer
Private vCancel As Boolean
Private vImpresora As Boolean

Private Sub cmdAceptar_Click()
    
    vCancel = False
    Me.Hide

End Sub

Private Sub cmdCancelar_Click()
    
    vCancel = True
    Me.Hide

End Sub

Private Sub txtCopias_GotFocus()
    
    marcarseleccion Me.txtCopias

End Sub

Private Sub txtCopias_LostFocus()
    
    If Not IsNumeric(Me.txtCopias.Text) Then Me.txtCopias.Text = vCopies
    vCopies = Me.txtCopias.Text

End Sub

Private Sub txtDesde_GotFocus()
    
    Me.optPaginas.Value = True
    marcarseleccion Me.txtDesde

End Sub

Private Sub txtDesde_LostFocus()
    
    If Not IsNumeric(Me.txtDesde.Text) Then Me.txtDesde.Text = vFromPage
    If Me.txtDesde.Text < vMin Then Me.txtDesde.Text = vMin
    vFromPage = Me.txtDesde.Text

End Sub

Private Sub txtHasta_GotFocus()
    
    Me.optPaginas.Value = True
    marcarseleccion Me.txtHasta

End Sub

Private Sub txtHasta_LostFocus()
    
    If Not IsNumeric(Me.txtHasta.Text) Then Me.txtHasta.Text = vToPage
    If Me.txtHasta.Text > vMax Then Me.txtHasta.Text = vMax
    vToPage = Me.txtHasta.Text

End Sub

Private Sub optImpresora_Click()
    
    vImpresora = True
    Me.Height = 4740

End Sub

Private Sub chkIntercalar_Click()
    
    vCollate = Not vCollate

End Sub

Private Sub optPaginas_Click()
    
    vFromPage = Me.txtDesde.Text
    vToPage = Me.txtHasta.Text

End Sub

Private Sub optPantalla_Click()
    
    vImpresora = False
    Me.Height = 1560

End Sub

Private Sub optTodo_Click()
    
    vFromPage = vMin
    vToPage = vMax

End Sub

Private Sub Form_Activate()
    
    Me.txtCopias.Text = vCopies
    Me.txtDesde.Text = vMin
    Me.txtHasta.Text = vMax
    Me.optTodo.Value = True
    Me.chkIntercalar.Value = 0

End Sub

Private Sub Form_Load()
Dim prtImpresora As Printer
    
    vCopies = 1
    vMin = 1
    vMax = 32767
    vFromPage = vMin
    vToPage = vMax
    vCollate = False
    vCancel = False
    vImpresora = False
    For Each prtImpresora In Printers
        Me.cboImpresoras.AddItem prtImpresora.DeviceName
        If prtImpresora.DeviceName = Printer.DeviceName Then vDefaultPrinter = Me.cboImpresoras.NewIndex
    Next
    If Me.cboImpresoras.ListCount > 0 Then Me.cboImpresoras.ListIndex = vDefaultPrinter
    Me.optPantalla.Value = True

End Sub

Private Sub cboImpresoras_Click()
    
    Me.lblTipo.Caption = Printers(Me.cboImpresoras.ListIndex).DeviceName
    Me.lblUbicacion.Caption = Printers(Me.cboImpresoras.ListIndex).Port
    vDefaultPrinter = Me.cboImpresoras.ListIndex

End Sub

Property Let Copies(ByVal p_Copies As Integer)
    
    vCopies = p_Copies

End Property

Property Get Copies() As Integer
    
    Copies = vCopies

End Property

Property Let Min(ByVal p_Min As Integer)
    
    vMin = p_Min

End Property

Property Get Min() As Integer
    
    Min = vMin

End Property

Property Let Max(ByVal p_Max As Integer)
    
    vMax = p_Max

End Property

Property Get Max() As Integer
    
    Max = vMax

End Property

Property Let FromPage(ByVal p_FromPage As Integer)
    
    vFromPage = p_FromPage

End Property

Property Get FromPage() As Integer
    
    FromPage = vFromPage

End Property

Property Let ToPage(ByVal p_ToPage As Integer)
    
    vToPage = p_ToPage

End Property

Property Get ToPage() As Integer
    
    ToPage = vToPage

End Property

Property Let Collate(ByVal p_Collate As Boolean)
    
    vCollate = p_Collate

End Property

Property Get Collate() As Boolean
    
    Collate = vCollate

End Property

Property Let DefaultPrinter(ByVal p_DefaultPrinter As Integer)
    
    vDefaultPrinter = p_DefaultPrinter

End Property

Property Get DefaultPrinter() As Integer
    
    DefaultPrinter = vDefaultPrinter

End Property

Property Get Cancel() As Boolean
    
    Cancel = vCancel

End Property

Property Get prtImpresora() As Boolean
    
    prtImpresora = vImpresora

End Property

Public Sub cargar(pReport As CrystalReport)
    
    pReport.Destination = IIf(vImpresora, crptToPrinter, crptToWindow)
    pReport.WindowShowPrintBtn = True
    pReport.WindowShowPrintSetupBtn = True
    pReport.PrinterCollation = IIf(vCollate, crptCollated, crptUncollated)
    pReport.PrinterCopies = vCopies
    If Me.cboImpresoras.ListCount = 0 Then Exit Sub
    pReport.PrinterDriver = Printers(vDefaultPrinter).DriverName
    pReport.PrinterName = Printers(vDefaultPrinter).DeviceName
    pReport.PrinterPort = Printers(vDefaultPrinter).Port
    pReport.PrinterStartPage = vFromPage
    pReport.PrinterStopPage = vToPage

End Sub
