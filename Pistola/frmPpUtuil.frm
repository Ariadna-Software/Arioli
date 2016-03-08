VERSION 5.00
Begin VB.Form frmPpUtuil 
   Caption         =   "Utilidades auxiliares"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3480
   Icon            =   "frmPpUtuil.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   3480
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "S A L I R"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Verificar artículo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "(c) Ariadna Software"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   3255
   End
End
Attribute VB_Name = "frmPpUtuil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdMoval_Click()
    frmStock.Show vbModal
End Sub

Private Sub cmdMovi_Click()
    frmComprobarMovClie.Show vbModal
End Sub

Private Sub cmdRevInvne_Click()
    frmRevisonInven.Show vbModal
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    frmLotaje.Show vbModal
End Sub

Private Sub Command3_Click()
    frmFichaTecnica.Show vbModal
End Sub

Private Sub Command4_Click()
    Me.Visible = False
    frmPruebasPedido.Show vbModal
    Me.Visible = True
End Sub

Private Sub Form_Load()
Dim C As String

    
    'Le asigno al codusu el usuario
    C = DevuelveDesdeBD(1, "codtraba", "straba", "login", vUsu.Login, "T")
    If C = "" Then C = "0"
    vUsu.CadenaConexion = C
    
    FijarContadorActualizaciones
    
    
End Sub


Private Sub FijarContadorActualizaciones()
Dim C As String
    On Error GoTo EF
    '   "'INV" & NombrePistola & "'
    ContadorActualizaciones = 0
    C = "fechamov= '" & Format(Now, FormatoFecha) & "' AND document='INV" & NombrePistola & "' AND detamovi"
    C = DevuelveDesdeBD("max(numlinea)", "smoval", C, "DFI", "T")
    If C <> "" Then ContadorActualizaciones = Val(C)
    
    
    
    Exit Sub
EF:
    Err.Clear
End Sub
