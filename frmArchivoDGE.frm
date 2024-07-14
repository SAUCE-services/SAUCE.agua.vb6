VERSION 5.00
Begin VB.Form frmArchivoDGE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generar Archivo DGE"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6045
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   6045
   Begin VB.PictureBox picConsumo 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000E&
      Height          =   5895
      Left            =   9960
      ScaleHeight     =   5835
      ScaleWidth      =   5835
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "Generar"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      ToolTipText     =   "Factura e Imprime todas las CONEXIONES pendientes de facturar"
      Top             =   360
      Width           =   1695
   End
   Begin VB.ComboBox cboPeriodo 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Seleccione el Período a Facturar"
      Top             =   360
      Width           =   3615
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      ToolTipText     =   "Fin de la TAREA"
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Período"
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   570
   End
End
Attribute VB_Name = "frmArchivoDGE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdGenerar_Click()
Dim request As MSXML2.ServerXMLHTTP

Dim file_stream As ADODB.Stream

Dim periodo As New clsRESTPeriodo

Dim url As String
Dim filename As String

    If Me.cboPeriodo.ListIndex < 0 Then Exit Sub
    
    periodo.periodoId = Me.cboPeriodo.ItemData(Me.cboPeriodo.ListIndex)
    periodo.findByPrimaryKey

    filename = App.path & "\UVSPES." & Year(periodo.fechaInicio) & Format(Month(periodo.fechaInicio), "00") & ".txt"

    Me.MousePointer = 11
    Me.cmdGenerar.Enabled = False
    
    Set request = New MSXML2.ServerXMLHTTP
    
    url = modUrls.url_agua & "dgefile/generate/" & Me.cboPeriodo.ItemData(Me.cboPeriodo.ListIndex)

    request.setTimeouts 200000, 200000, 200000, 200000
    request.Open "GET", url, False
    DoEvents
    request.send
    
    If request.Status = 200 Then
        Set file_stream = New ADODB.Stream
        file_stream.Open
        file_stream.Type = adTypeBinary
        
        file_stream.Write request.responseBody
        file_stream.Position = 0
        
        file_stream.SaveToFile filename, adSaveCreateOverWrite
        file_stream.Close
        
        Set file_stream = Nothing
        
        ShellExecute Me.hwnd, "open", filename, vbNullString, vbNullString, 1
        
    End If
    
    Set request = Nothing
    
    Me.cmdGenerar.Enabled = True
    Me.MousePointer = 0

End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()
Dim periodo As New clsRESTPeriodo

    periodo.fillCombo Me.cboPeriodo
    
End Sub

