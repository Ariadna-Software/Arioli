VERSION 5.00
Begin VB.Form frmFacTraerOferta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Traer Lineas de Oferta"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   ClipControls    =   0   'False
   Icon            =   "frmFacTraerOferta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CheckBox Check1 
      Caption         =   "Datos carta"
      Height          =   195
      Index           =   1
      Left            =   2880
      TabIndex        =   6
      Top             =   960
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Copiar observaciones"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   1680
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   1680
      Width           =   1035
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   1320
      ToolTipText     =   "Buscar artículo"
      Top             =   240
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Nº Oferta"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "Cargando datos ........."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   8220
      Visible         =   0   'False
      Width           =   3495
   End
End
Attribute VB_Name = "frmFacTraerOferta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim NombreTabla As String
Dim Ordenacion As String


Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1

Private Modo As Byte
'Solo utilizamos el Modo=4 -> Modificar

Dim kCampo As Integer

Public Event CargarOferta(NumOfert As String)

Private HaDevueltoDatos As Boolean


Private Sub cmdAceptar_Click()
Dim NumOfe As String
On Error GoTo Error1

    Screen.MousePointer = vbHourglass
    NumOfe = Text1.Text & "|" & Check1(0).Value & "|" & Check1(1).Value & "|"
    Unload Me
    RaiseEvent CargarOferta(NumOfe)

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub cmdCancelar_Click()
On Error GoTo ECancelar
    Unload Me
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    Me.imgBuscar(0).Picture = frmPpal.imgListComun.ListImages(19).Picture
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Text1.Text = RecuperaValor(CadenaDevuelta, 1)
    Text1_LostFocus
End Sub

Private Sub imgBuscar_Click(Index As Integer)

    If Index = 0 Then
        Set frmB = New frmBuscaGrid
        frmB.vCampos = "Nº Oferta|scapre|numofert|N|0000000|13·Fecha Ofer.|scapre|fecofert|F|dd/mm/yyyy|15·Cliente|scapre|codclien|N|000000|12·" & _
            "Nombre Cliente|scapre|nomclien|T||45·Importe||sum(importel)|T|#,##0.00|13·"
        frmB.vTabla = "scapre,slipre"
        frmB.vSQL = "scapre.numofert=slipre.numofert group by numofert"
        frmB.vDevuelve = "0|"
        frmB.vTitulo = "Ofertas"
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri
        frmB.Show vbModal
    End If
End Sub

Private Sub Text1_GotFocus()
    ConseguirFoco Text1, Modo
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text1_LostFocus()
Dim Devuelve As String

    With Text1
        If .Text = "" Then Exit Sub
        .Text = Format(.Text, "0000000")
        'Comprobar que la oferta existe
        Devuelve = DevuelveDesdeBDNew(conAri, "scapre", "numofert", "numofert", .Text, "N")
        If Devuelve = "" Then
            MsgBox "No existe la Oferta: " & .Text, vbInformation
            Text1.Text = ""
            PonerFoco Text1
        End If
    End With
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerModo(Kmodo As Byte)
       
    Modo = Kmodo
End Sub

