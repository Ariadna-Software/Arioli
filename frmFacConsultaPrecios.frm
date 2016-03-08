VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacConsultaPrecios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta precios"
   ClientHeight    =   8280
   ClientLeft      =   345
   ClientTop       =   2430
   ClientWidth     =   12480
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   12480
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FramePedirDatos 
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   8655
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   0
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Text5"
         Top             =   480
         Width           =   3735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2820
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "Text5"
         Top             =   960
         Width           =   3975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   1
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   7080
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   0
         Left            =   7080
         TabIndex        =   6
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   66
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   585
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   0
         Left            =   960
         Picture         =   "frmFacConsultaPrecios.frx":0000
         ToolTipText     =   "Buscar cliente"
         Top             =   480
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   1
         Left            =   1035
         Picture         =   "frmFacConsultaPrecios.frx":0102
         ToolTipText     =   "Buscar artículo"
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Articulo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   54
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   660
      End
   End
   Begin VB.Frame FrameMostrarDatos 
      Height          =   8175
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   12255
      Begin VB.TextBox txtResultado 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   17
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   47
         Text            =   "Text1"
         Top             =   7665
         Width           =   1215
      End
      Begin VB.TextBox txtResultado 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   16
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   46
         Text            =   "Text1"
         Top             =   7665
         Width           =   735
      End
      Begin VB.TextBox txtResultado 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   15
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   45
         Text            =   "Text1"
         Top             =   7665
         Width           =   735
      End
      Begin VB.TextBox txtResultado 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   14
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   44
         Text            =   "Text1"
         Top             =   7665
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmFacConsultaPrecios.frx":0204
         Left            =   9960
         List            =   "frmFacConsultaPrecios.frx":0214
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   1560
         Width           =   2055
      End
      Begin MSComctlLib.ListView ListStock 
         Height          =   1815
         Left            =   240
         TabIndex        =   40
         Top             =   5640
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   3201
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cod."
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Almacen"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Stock"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txtResultado 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   13
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   38
         Text            =   "Text1"
         Top             =   3360
         Width           =   1335
      End
      Begin MSComctlLib.ListView listTarifa 
         Height          =   1575
         Left            =   240
         TabIndex        =   37
         Top             =   3840
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   2778
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tarifa"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Des. tarifa"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Precio"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txtResultado 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   12
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox txtResultado 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   11
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   33
         Text            =   "Text1"
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox txtResultado 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox txtResultado 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   9
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   2400
         Width           =   4215
      End
      Begin VB.TextBox txtResultado 
         Height          =   285
         Index           =   8
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox txtResultado 
         Height          =   285
         Index           =   7
         Left            =   8280
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   465
         Width           =   615
      End
      Begin VB.TextBox txtResultado 
         Height          =   285
         Index           =   6
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   465
         Width           =   735
      End
      Begin VB.TextBox txtResultado 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   10560
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   465
         Width           =   975
      End
      Begin VB.TextBox txtResultado 
         Height          =   285
         Index           =   4
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   1080
         Width           =   3975
      End
      Begin VB.TextBox txtResultado 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   3
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox txtResultado 
         Height          =   285
         Index           =   2
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtResultado 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox txtResultado 
         Height          =   285
         Index           =   0
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Volver"
         Height          =   330
         Index           =   1
         Left            =   10920
         TabIndex        =   5
         Top             =   7680
         Width           =   1095
      End
      Begin MSComctlLib.ListView listDatos 
         Height          =   5535
         Left            =   5640
         TabIndex        =   41
         Top             =   1920
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   9763
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "T"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Cant."
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Precio"
            Object.Width           =   2152
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Dto1"
            Object.Width           =   1023
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Dto2"
            Object.Width           =   1023
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Importe"
            Object.Width           =   1765
         EndProperty
      End
      Begin VB.Label lblSituacion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   53
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Importe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   7320
         TabIndex        =   52
         Top             =   7710
         Width           =   2535
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Importe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   5160
         TabIndex        =   51
         Top             =   7710
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Dto2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   3840
         TabIndex        =   50
         Top             =   7710
         Width           =   465
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Dto1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   2520
         TabIndex        =   49
         Top             =   7710
         Width           =   495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Precio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   48
         Top             =   7710
         Width           =   975
      End
      Begin VB.Label lblIndicador 
         Alignment       =   1  'Right Justify
         Caption         =   "Label2"
         Height          =   255
         Left            =   10200
         TabIndex        =   42
         Top             =   7440
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Stock"
         Height          =   255
         Index           =   14
         Left            =   3000
         TabIndex        =   39
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "P.V.P."
         Height          =   255
         Index           =   13
         Left            =   3000
         TabIndex        =   36
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "P.M.P"
         Height          =   255
         Index           =   12
         Left            =   360
         TabIndex        =   34
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "P.U.C"
         Height          =   255
         Index           =   11
         Left            =   360
         TabIndex        =   32
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Descripcion"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   30
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Artículo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   27
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Pendiente(€)"
         Height          =   255
         Index           =   8
         Left            =   9600
         TabIndex        =   22
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Dto Pp."
         Height          =   255
         Index           =   7
         Left            =   7680
         TabIndex        =   21
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Dto gral"
         Height          =   255
         Index           =   6
         Left            =   5640
         TabIndex        =   20
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "F. pago"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   19
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Situación"
         Height          =   255
         Index           =   4
         Left            =   5640
         TabIndex        =   18
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Ofertas,pedidos, albaranes,facturas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   5640
         TabIndex        =   12
         Top             =   1560
         Width           =   3855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   2
         X1              =   7800
         X2              =   9840
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00008000&
         BorderWidth     =   2
         Index           =   1
         X1              =   1080
         X2              =   5400
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Index           =   0
         X1              =   1080
         X2              =   12000
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   1215
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   240
         Top             =   7560
         Width           =   9855
      End
   End
End
Attribute VB_Name = "frmFacConsultaPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private WithEvents frmA As frmAlmArticulos
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmC As frmFacClientes
Attribute frmC.VB_VarHelpID = -1

Dim Cad As String
Dim It As ListItem

Dim Valor As Currency
'

'
'
Private Sub cmdBuscar_Click()

    If txtCodigo(0).Text = "" Or txtCodigo(1).Text = "" Then
        MsgBox "Debe poner cliente /articulo", vbExclamation
        Exit Sub
    End If
    Combo1.ListIndex = 0
    LimpiarResultados
    PonerVisiblePedir False
    DoEvents
    Screen.MousePointer = vbHourglass
    Set miRsAux = New ADODB.Recordset
    CargarDatos
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
    lblIndicador.Caption = ""
    
End Sub


Private Sub LimpiarResultados()
Dim T As TextBox
    lblIndicador.Caption = ""
    For Each T In Me.txtResultado
        T.Text = ""
    Next
    Me.listTarifa.ListItems.Clear
    Me.ListStock.ListItems.Clear
    Me.listDatos.ListItems.Clear
    Label2(4).Caption = ""
    lblSituacion.Caption = ""
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    If Index = 1 Then
        PonerVisiblePedir True
    Else
        Unload Me
    End If
End Sub



Private Sub Combo1_Click()
    'para que al inicio NO haga nada
    If FramePedirDatos.visible Then Exit Sub
    
    '------------------------------------------
    Set miRsAux = New ADODB.Recordset
    CargarDatosFacturacion
    lblIndicador.Caption = ""
    Set miRsAux = Nothing
End Sub

Private Sub Form_Load()
    Me.Icon = frmppal.Icon
    limpiar Me
    PonerVisiblePedir True
    'ASigno estos iconos
    
    Me.listDatos.SmallIcons = frmppal.ImgListPpal
End Sub


Private Sub PonerVisiblePedir(b As Boolean)
    Me.FrameMostrarDatos.visible = Not b
    Me.FramePedirDatos.visible = b
    If b Then
        Me.cmdCancelar(0).Cancel = True
        Me.Width = FramePedirDatos.Width
        Me.Height = FramePedirDatos.Height
    Else
        Me.cmdCancelar(1).Cancel = True
        Me.Width = FrameMostrarDatos.Width
        Me.Height = FrameMostrarDatos.Height
    End If
    Me.Width = Me.Width + 360
    Me.Height = Me.Height + 620
End Sub




Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(1).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(1).Text = RecuperaValor(CadenaSeleccion, 2)
    Cad = "O"
End Sub

Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(0).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(0).Text = RecuperaValor(CadenaSeleccion, 2)
    Cad = "O"
End Sub

Private Sub imgBuscarG_Click(Index As Integer)
    Cad = "" 'Para ver si devuelve datos
    If Index = 0 Then
        'Cliente
        
        Set frmC = New frmFacClientes
        frmC.DatosADevolverBusqueda = "1|"
        frmC.Show vbModal
        Set frmC = Nothing
        If Cad <> "" Then PonerFoco txtCodigo(1)
    Else
        'Articulo
        Set frmA = New frmAlmArticulos
        frmA.DeConsulta = True
        frmA.DatosADevolverBusqueda2 = "@1@"
        frmA.Show vbModal
        Set frmA = Nothing
        If Cad <> "" Then PonerFocoBtn Me.cmdBuscar
    End If
    Cad = ""
        
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub



Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
    txtCodigo(Index).Text = Trim(txtCodigo(Index))
    
    Cad = ""
    If Index = 0 Then
        
        'Cliente
        If txtCodigo(Index).Text <> "" Then
            If Not IsNumeric(txtCodigo(Index).Text) Then
                MsgBox "Campo codigo cliente debe ser numérico", vbExclamation
                PonerFoco txtCodigo(Index)
            Else
                Cad = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", txtCodigo(Index).Text, "N")
                If Cad = "" Then
                    MsgBox "No existe el cliente : " & txtCodigo(Index).Text, vbExclamation
                End If
            End If
        End If
        
        
    Else
        'articulo
        If txtCodigo(Index).Text <> "" Then
            Cad = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", txtCodigo(Index).Text, "T")
            If Cad = "" Then
                MsgBox "No existe el articulo: " & txtCodigo(Index).Text, vbExclamation
                PonerFoco txtCodigo(Index)
            End If
        End If
    End If
    Me.txtNombre(Index).Text = Cad
    If Cad = "" Then txtCodigo(Index).Text = ""
    
End Sub


Private Sub CargaStock()
Dim i As Currency
    Valor = 0
    txtResultado(13).Text = ""
    Cad = "select salmac.codalmac,nomalmac,canstock   from salmac,salmpr where salmac.codalmac="
    Cad = Cad & "salmpr.codalmac AND  codartic=" & DBSet(txtCodigo(1).Text, "T") & " ORDER BY salmac.codalmac"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set It = ListStock.ListItems.Add()
        It.Text = miRsAux!Codalmac
        It.SubItems(1) = miRsAux!nomalmac
        i = DBLet(miRsAux!CanStock, "N")
        It.SubItems(2) = Format(i, FormatoCantidad)
        Valor = Valor + 1
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    'Stock
    txtResultado(13).Text = Format(Valor, FormatoCantidad)
    
End Sub


Private Sub CargaTarifas()

    listTarifa.ListItems.Clear
    Cad = "select slista.codlista,nomlista,precioac from slista,starif where slista.codlista="
    Cad = Cad & "starif.codlista and codartic = " & DBSet(txtCodigo(1).Text, "T") & " ORDER BY slista.codlista"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set It = listTarifa.ListItems.Add()
        It.Text = miRsAux!codlista
        It.SubItems(1) = miRsAux!nomlista
        It.SubItems(2) = Format(DBLet(miRsAux!precioac, "N"), FormatoPrecio)
        If miRsAux!codlista = listTarifa.Tag Then
            'Tarifa del cliente
            It.Bold = True
            It.ForeColor = vbBlue
            It.ListSubItems(1).Bold = True
            It.ListSubItems(1).ForeColor = vbBlue
    
            It.ListSubItems(2).Bold = True
            It.ListSubItems(2).ForeColor = vbBlue
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
End Sub




Private Sub CargarDatos()

    On Error GoTo EC
    
    Cad = "OK"

    
    lblIndicador.Caption = "Datos cliente"
    lblIndicador.Refresh
    
    Cad = "select codclien ,nomclien ,dtoppago ,dtognral  ,codsitua ,codmacta,codforpa,codtarif  from sclien where codclien =" & Me.txtCodigo(0).Text
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    'Ponemos los campos
    '--------------------------------------------------------
    Me.txtResultado(0).Text = miRsAux!CodClien
    Me.txtResultado(1).Text = miRsAux!nomclien
    Me.txtResultado(2).Text = miRsAux!codforpa
    Me.txtResultado(3).Text = DevuelveDesdeBD(conAri, "nomforpa", "sforpa", "codforpa", miRsAux!codforpa, "N")
    Me.txtResultado(4).Text = DevuelveDesdeBD(conAri, "nomsitua", "ssitua", "codsitua", miRsAux!codsitua, "N")
    Me.txtResultado(6).Text = Format(miRsAux!DtoGnral, FormatoDescuento)
    Me.txtResultado(7).Text = Format(miRsAux!DtoPPago, FormatoDescuento)
    
    'Cargo la cta contable
    Cad = DBLet(miRsAux!Codmacta, "T")
    
    'Cargo la tarifa
    Me.listTarifa.Tag = miRsAux!codTarif
    
    'Cerramos el RS
    miRsAux.Close
    
    lblIndicador.Caption = "Cobros pendientes"
    lblIndicador.Refresh
    
    PonerCobrosPendientes Cad
    
    txtResultado(5).Text = Format(Valor, FormatoImporte)
    
    DoEvents
    
    'Datos articulo
    lblIndicador.Caption = "Articulo"
    lblIndicador.Refresh
    
    Cad = "select codartic,nomartic,preciouc,preciomp,preciove,unicajas,codartic,nomartic,preciouc,preciomp,preciove,unicajas,codstatu    from sartic where codartic =" & DBSet(Me.txtCodigo(1).Text, "T")
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Me.txtResultado(8).Text = miRsAux!codArtic
    Me.txtResultado(9).Text = miRsAux!NomArtic
    Me.txtResultado(12).Text = Format(DBLet(miRsAux!precioUC, "N"), FormatoPrecio)
    Me.txtResultado(12).Text = Format(DBLet(miRsAux!preciomp, "N"), FormatoPrecio)
    Me.txtResultado(12).Text = Format(DBLet(miRsAux!preciove, "N"), FormatoPrecio)
    Me.txtResultado(8).Tag = miRsAux!unicajas
    Select Case Val(miRsAux!codstatu)
    Case 1
        lblSituacion.Caption = "Bloqueado"
    Case 2
        lblSituacion.Caption = "Caducado"
    Case Else
        lblSituacion.Caption = ""
    End Select
    
    miRsAux.Close
    
    lblIndicador.Caption = "Stock"
    lblIndicador.Refresh
    CargaStock
    
    
    lblIndicador.Caption = "Tarifas"
    lblIndicador.Refresh
    CargaTarifas
    
    DoEvents
    
    'Datos albaranes......
    CargarDatosFacturacion
    
    
    'Ponemos el precio fianl
    CalcularPrecioFinal
    
    'Insertamos log de consulta
    lblIndicador.Caption = "Ins. log"
    lblIndicador.Refresh
    Cad = "insert into `sconsulta` (`DiaHora`,`Usuario`,`codclien`,`nomclien`,"
    '----------                                       cogera la fecha del mysql
    Cad = Cad & "`codartic`,`nomartic`) values (" & "concat(curdate(),' ',curtime())" & ","
    Cad = Cad & DBSet(vUsu.Nombre, "T") & "," & txtCodigo(0).Text & "," & DBSet(txtNombre(0), "T")
    Cad = Cad & "," & DBSet(txtCodigo(1), "T") & "," & DBSet(txtNombre(1), "T") & ")"
    Conn.Execute Cad
    
    Exit Sub
EC:
    MuestraError Err.Number, Err.Description
End Sub





Private Sub PonerCobrosPendientes(ByVal Codmacta As String)
    Valor = 0
    If Codmacta = "" Then Exit Sub
    'Obtener a partir de la cuenta del cliente si hay cobros pendientes en Contabilidad
    Cad = " WHERE scobro.codmacta = '" & Codmacta & "'"
    Cad = Cad & " AND fecvenci <= ' " & Format(Now, FormatoFecha) & "' "
    Cad = Cad & " AND (sforpa.tipforpa between 0 and 3)"
    Cad = "SELECT sum(impvenci - if(isnull(impcobro),0,impcobro)) FROM scobro INNER JOIN sforpa ON scobro.codforpa=sforpa.codforpa " & Cad
    
    'Lee de la Base de Datos de CONTABILIDAD
    miRsAux.Open Cad, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not miRsAux.EOF Then Valor = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
End Sub



'Cargara los datos de las lineas
'de OFERTAS,PEDIDOS,ALBARANES,FACTURA
Private Sub CargarDatosFacturacion()

    Me.listDatos.ListItems.Clear
    
        
    
    CargaDatosTablas CByte(Combo1.ListIndex)
    
    
End Sub



Private Sub CargaDatosTablas(KTabla As Byte)
Dim Aux As String
Dim Ico As Integer
    Select Case KTabla
    Case 3
        Ico = 5
        Cad = "slipre,scapre WHERE slipre.numofert=scapre.numofert"
        Aux = " '' as Primero,slipre.numofert as elnumero,fecofert as fecha"
        Me.lblIndicador.Caption = "Ofertas"
    Case 2
        Ico = 6
        Cad = "sliped,scaped where sliped.numpedcl=scaped.numpedcl"
        Aux = " '' as Primero,sliped.numpedcl as elnumero,fecpedcl  as fecha"
        Me.lblIndicador.Caption = "Pedidos"
    Case 1
        Ico = 7
        Cad = "slialb,scaalb where slialb.numalbar=scaalb.numalbar and slialb.codtipom=scaalb.codtipom"
        Aux = " slialb.codtipom as Primero,slialb.numalbar as elnumero,fechaalb as fecha"
        Me.lblIndicador.Caption = "Albaranes"
    Case 0
        Ico = 8
        Cad = " slifac,scafac where slifac.numfactu=scafac.numfactu and slifac.codtipom=scafac.codtipom and slifac.fecfactu=scafac.fecfactu"
        Aux = "slifac.codtipom as primero,slifac.numfactu as elnumero,slifac.fecfactu as fecha"
        Me.lblIndicador.Caption = "Facturas"
    End Select
    Me.lblIndicador.Refresh
    
    Aux = "Select " & Aux & ",Cantidad, precioar, dtoline1, dtoline2, ImporteL FROM " & Cad
    Cad = Aux & " AND codartic = " & DBSet(Me.txtCodigo(1).Text, "T")
    Cad = Cad & " AND codclien = " & txtCodigo(0).Text
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set It = Me.listDatos.ListItems.Add()
        Cad = miRsAux!primero & Format(miRsAux!elnumero, "000000")
        It.Text = ""
        
        It.SubItems(1) = Format(miRsAux!Fecha, "dd/mm/yyyy")
        It.SubItems(2) = Format(miRsAux!Cantidad, FormatoCantidad)
        It.SubItems(3) = Format(miRsAux!precioar, FormatoPrecio)
        It.SubItems(4) = Format(miRsAux!dtoline1, FormatoDescuento)
        It.SubItems(5) = Format(miRsAux!dtoline2, FormatoDescuento)
        It.SubItems(6) = Format(miRsAux!ImporteL, FormatoImporte)
        It.SmallIcon = Ico
        It.ToolTipText = Me.lblIndicador.Caption & " " & Cad
        miRsAux.MoveNext
    Wend
    miRsAux.Close
End Sub





'Calculamos el precio que se le va a quedar
Private Sub CalcularPrecioFinal()
Dim CPrecioFact As CPreciosFact
Dim Precio As Currency
Dim Cantidad As Currency
Dim PorCaja As Boolean
Dim NumCajas As Integer
                Set CPrecioFact = New CPreciosFact
                'Si se puede vender por cajas(devuelve>1) poner numero de cajas en una linea con el
                'precio de caja, y otra linea con el resto unidades un precio unidad
                'Cantidad = txtAux(Index).Text
                Cantidad = 1
                NumCajas = CPrecioFact.ObtenerNumCajas(CStr(Cantidad), CStr(txtResultado(8).Tag))
                'RestoUnid = CInt(ComprobarCero(Cantidad)) - NumCajas * CInt(devuelve)
                'Obtenemos la Tarifa del Cliente
                CPrecioFact.CodigoLista = Me.listTarifa.Tag  'la tarifa del cliente
                    
                CPrecioFact.CodigoArtic = Me.txtCodigo(1).Text
                CPrecioFact.CodigoClien = Me.txtCodigo(0).Text
                PorCaja = (NumCajas > 0)
                Precio = CPrecioFact.ObtenerPrecio(PorCaja, CStr(Now), Cad)
                    
                'En cad TENGO el origen del precio
                Select Case Cad
                    Case "P": Label2(4).Caption = "Promoción"
                    Case "E": Label2(4).Caption = "Precio Especial"
                    Case "T": Label2(4).Caption = "Tarifa Artículo"
                    Case "A": Label2(4).Caption = "Precio Artículo"
                    Case "M": Label2(4).Caption = "Manual"
                End Select
                
                    txtResultado(14).Text = Precio
                    PonerFormatoDecimal txtResultado(14), 2
                    txtResultado(15).Text = CPrecioFact.Descuento1
                    PonerFormatoDecimal txtResultado(15), 4
                    txtResultado(16).Text = CPrecioFact.Descuento2
                    PonerFormatoDecimal txtResultado(16), 4
                    txtResultado(17).Text = CalcularImporte(CStr(Cantidad), txtResultado(14).Text, txtResultado(15).Text, txtResultado(16).Text, vParamAplic.TipoDtos)
                    PonerFormatoDecimal txtResultado(17), 1
                Set CPrecioFact = Nothing
End Sub
