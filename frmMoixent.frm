VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMoixent 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Moixent"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   7200
      TabIndex        =   15
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   5760
      TabIndex        =   14
      Top             =   4800
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Albar�n venta"
      TabPicture(0)   =   "frmMoixent.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtEAN(0)"
      Tab(0).Control(1)=   "txtImporte(0)"
      Tab(0).Control(2)=   "txtDescClie(0)"
      Tab(0).Control(3)=   "txtCliente(0)"
      Tab(0).Control(4)=   "txtLote(0)"
      Tab(0).Control(5)=   "txtArticulo(0)"
      Tab(0).Control(6)=   "txtDescArticulo(0)"
      Tab(0).Control(7)=   "imgLote(0)"
      Tab(0).Control(8)=   "Label1(2)"
      Tab(0).Control(9)=   "Label1(1)"
      Tab(0).Control(10)=   "Label1(0)"
      Tab(0).Control(11)=   "Label4(7)"
      Tab(0).Control(12)=   "lblDpto(33)"
      Tab(0).Control(13)=   "imgCliente(0)"
      Tab(0).Control(14)=   "Label4(0)"
      Tab(0).Control(15)=   "Label4(2)"
      Tab(0).Control(16)=   "imgArticulo(0)"
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "Trasiegos"
      TabPicture(1)   =   "frmMoixent.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label4(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label4(3)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(3)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "imgArticulo(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label1(4)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label1(5)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label4(4)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label4(5)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label4(6)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label4(8)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label4(9)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label4(10)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label4(11)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label4(12)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "imgLote(1)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "imgLote(2)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "txtEAN(1)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "txtArticulo(1)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "txtDescArticulo(1)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "txtLote(1)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "txtImporte(1)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "txtLote(2)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "txtImporte(2)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "txtObslin(0)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "txtObslin(1)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "txtObslin(2)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).ControlCount=   26
      Begin VB.TextBox txtObslin 
         Height          =   285
         Index           =   2
         Left            =   5280
         MaxLength       =   50
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   2640
         Width           =   2895
      End
      Begin VB.TextBox txtObslin 
         Height          =   285
         Index           =   1
         Left            =   5280
         MaxLength       =   50
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1920
         Width           =   2895
      End
      Begin VB.TextBox txtObslin 
         Height          =   615
         Index           =   0
         Left            =   1440
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   13
         Text            =   "frmMoixent.frx":0038
         Top             =   3360
         Width           =   5535
      End
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   3120
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox txtLote 
         Height          =   285
         Index           =   2
         Left            =   960
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   3120
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtLote 
         Height          =   285
         Index           =   1
         Left            =   960
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txtDescArticulo 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "Text5"
         Top             =   1080
         Width           =   4215
      End
      Begin VB.TextBox txtArticulo 
         Height          =   285
         Index           =   1
         Left            =   2280
         MaxLength       =   16
         TabIndex        =   6
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtEAN 
         Height          =   285
         Index           =   1
         Left            =   600
         MaxLength       =   16
         TabIndex        =   5
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtEAN 
         Height          =   285
         Index           =   0
         Left            =   -74640
         MaxLength       =   16
         TabIndex        =   0
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   -71760
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   -73440
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   3360
         Width           =   5175
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   0
         Left            =   -74640
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   3360
         Width           =   1095
      End
      Begin VB.TextBox txtLote 
         Height          =   285
         Index           =   0
         Left            =   -74640
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox txtArticulo 
         Height          =   285
         Index           =   0
         Left            =   -72720
         MaxLength       =   16
         TabIndex        =   1
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtDescArticulo 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   -71160
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "Text5"
         Top             =   1200
         Width           =   4215
      End
      Begin VB.Image imgLote 
         Height          =   240
         Index           =   2
         Left            =   720
         Picture         =   "frmMoixent.frx":003E
         Top             =   2640
         Width           =   240
      End
      Begin VB.Image imgLote 
         Height          =   240
         Index           =   1
         Left            =   720
         Picture         =   "frmMoixent.frx":0140
         Top             =   1920
         Width           =   240
      End
      Begin VB.Image imgLote 
         Height          =   240
         Index           =   0
         Left            =   -74160
         Picture         =   "frmMoixent.frx":0242
         Top             =   1920
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Articulo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   12
         Left            =   120
         TabIndex        =   39
         Top             =   480
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Observa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   11
         Left            =   4560
         TabIndex        =   38
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Observa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   10
         Left            =   4560
         TabIndex        =   37
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   240
         TabIndex        =   36
         Top             =   3360
         Width           =   1065
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   2400
         TabIndex        =   35
         Top             =   2640
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Lote"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   360
         TabIndex        =   34
         Top             =   2640
         Width           =   315
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   2400
         TabIndex        =   33
         Top             =   1920
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Lote"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   360
         TabIndex        =   32
         Top             =   1920
         Width           =   315
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo"
         Height          =   255
         Index           =   5
         Left            =   2280
         TabIndex        =   30
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Descripcion"
         Height          =   255
         Index           =   4
         Left            =   3840
         TabIndex        =   29
         Top             =   840
         Width           =   1095
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   1
         Left            =   960
         Picture         =   "frmMoixent.frx":0344
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "EAN"
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   28
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "ENTRADA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   27
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "SALIDA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   1560
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "Descripcion"
         Height          =   255
         Index           =   2
         Left            =   -71160
         TabIndex        =   25
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "EAN"
         Height          =   255
         Index           =   1
         Left            =   -74640
         TabIndex        =   24
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo"
         Height          =   255
         Index           =   0
         Left            =   -72720
         TabIndex        =   23
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad"
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
         Index           =   7
         Left            =   -71760
         TabIndex        =   22
         Top             =   1920
         Width           =   750
      End
      Begin VB.Label lblDpto 
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
         Index           =   33
         Left            =   -74640
         TabIndex        =   21
         Top             =   3000
         Width           =   585
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   0
         Left            =   -73920
         Picture         =   "frmMoixent.frx":0446
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Lote"
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
         Index           =   0
         Left            =   -74640
         TabIndex        =   19
         Top             =   1920
         Width           =   375
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
         Index           =   2
         Left            =   -74640
         TabIndex        =   18
         Top             =   720
         Width           =   660
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   0
         Left            =   -73920
         Picture         =   "frmMoixent.frx":0548
         Top             =   720
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmMoixent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmMtoArticulos As frmAlmArticulos
Attribute frmMtoArticulos.VB_VarHelpID = -1
Private WithEvents frmCli As frmFacClientes
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmL As frmAlmPartidas
Attribute frmL.VB_VarHelpID = -1


Dim Codigo As String
Dim PriVez As Boolean


Private Sub Cancelar_Click()
    Unload Me
End Sub

Private Sub cmdAceptar_Click()


    If Me.SSTab1.Tab = 0 Then
        'Albaranes
        BotonAlbaran
        
    Else
        BotonTrasiego
    End If


End Sub


Private Sub BotonAlbaran()
Dim Cli As CCliente
    'Comprobaciones
    Codigo = ""
    If Me.txtArticulo(0).Text = "" Then Codigo = Codigo & "    -Articulo" & vbCrLf
    If Me.txtLote(0).Text = "" Then Codigo = Codigo & "    -Lote" & vbCrLf
    If Me.txtImporte(0).Text = "" Then Codigo = Codigo & "    -Cantidad" & vbCrLf
    If Me.txtCliente(0).Text = "" Then Codigo = Codigo & "    -Cliente" & vbCrLf
    
    If Codigo <> "" Then
        Codigo = "Campos vacios: " & vbCrLf & vbCrLf & Codigo
        MsgBox Codigo, vbExclamation
        Exit Sub
    End If
    
    'Pregunta
    If MsgBox("Desea realizar el albar�n?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    
    'Vamos p'alla
    NumRegElim = 0 'Para guardarno el numalbar
    Set Cli = New CCliente
    If Cli.LeerDatos(Me.txtCliente(0).Text) Then
        If GeneraAlbaran(Cli) Then
            ImprimeAlbaran
            limpiar Me
            PonerFoco Me.txtArticulo(0)
        End If
    End If
    Set Cli = Nothing
    
End Sub



Private Sub BotonTrasiego()
Dim cP As cPartidas
    'Comprobaciones
    Codigo = PonerTrabajadorConectado("") 'Codigo trabajador concetado
    If Codigo = "" Then
        MsgBox "Error trabajador conectado", vbExclamation
        Exit Sub
    End If
    
    Codigo = ""
    If Me.txtArticulo(1).Text = "" Then Codigo = Codigo & "    -Articulo" & vbCrLf
    For NumRegElim = 1 To 2

        If Me.txtLote(NumRegElim).Text = "" Then Codigo = Codigo & "    -Lote(" & NumRegElim & ")" & vbCrLf
        If Me.txtImporte(NumRegElim).Text = "" Then Codigo = Codigo & "    -Cantidad" & NumRegElim & ")" & vbCrLf
    Next
    
    If Codigo <> "" Then
        Codigo = "Campos vacios: " & vbCrLf & vbCrLf & Codigo
        MsgBox Codigo, vbExclamation
        Exit Sub
    End If
        
    If Me.txtLote(2).Text = Me.txtLote(1).Text Then
        MsgBox "Mismo lote entrada-salida", vbExclamation
        Exit Sub
    End If
        
    'Todos los datos puestos
    'Comprobaciones
    'Los lotes son validos
    Set cP = New cPartidas
    
    Codigo = ""
    If Not cP.LeerDesdeArticulo(txtArticulo(1).Text, 1, Me.txtLote(1).Text) Then
        Codigo = "No se encuentra el lote: " & Me.txtLote(1).Text & " del articulo " & txtArticulo(1).Text & vbCrLf
        NumRegElim = 1
    End If
    
    If Not cP.LeerDesdeArticulo(txtArticulo(1).Text, 1, Me.txtLote(2).Text) Then
        Codigo = Codigo & "No se encuentra el lote: " & Me.txtLote(2).Text & " del articulo " & txtArticulo(1).Text & vbCrLf
        NumRegElim = 2
    End If
    If Codigo <> "" Then
        MsgBox Codigo, vbInformation
        PonerFoco txtLote(NumRegElim)
        Exit Sub
    End If
    
    'Si las cantidades no son iguales
    If Me.txtImporte(1).Text <> txtImporte(2).Text Then
        If MsgBox("Cantidades distintas. �Continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    Else
        'Pregunta
        If MsgBox("Desea realizar el trasiego?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    'Vamos p'alla
    Screen.MousePointer = vbHourglass
    HacerTrasiego
    Screen.MousePointer = vbDefault
    
End Sub






Private Sub Form_Activate()
    If PriVez Then
        PriVez = False
        PonerFoco Me.txtEAN(1)
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = frmppal.Icon
    limpiar Me
    SSTab1.Tab = 1
    PriVez = True
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
     Codigo = CadenaSeleccion
End Sub

Private Sub frmL_DatoSeleccionado(CadenaSeleccion As String)
    Codigo = CadenaSeleccion
End Sub

Private Sub frmMtoArticulos_DatoSeleccionado(CadenaSeleccion As String)
    Codigo = CadenaSeleccion
End Sub

Private Sub imgArticulo_Click(Index As Integer)
  
    Codigo = ""
    Set frmMtoArticulos = New frmAlmArticulos
    frmMtoArticulos.DatosADevolverBusqueda2 = "@1@" 'Abrimos en Modo Busqueda
    frmMtoArticulos.Show vbModal
    Set frmMtoArticulos = Nothing
    If Codigo <> "" Then
        Me.txtArticulo(Index).Text = RecuperaValor(Codigo, 1)
        txtArticulo_LostFocus Index
    End If
End Sub



Private Sub imgCliente_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    Codigo = ""
    Set frmCli = New frmFacClientes
    frmCli.DatosADevolverBusqueda = "0|1|"
    frmCli.Show vbModal
    Set frmCli = Nothing
    If Codigo <> "" Then
        Me.txtCliente(Index).Text = RecuperaValor(Codigo, 1)
        Me.txtDescClie(Index).Text = RecuperaValor(Codigo, 2)
        Codigo = ""
    End If
End Sub

Private Sub imgLote_Click(Index As Integer)
Dim C As Currency
Dim Ind As Integer
    Ind = Index
    If Index = 2 Then Ind = 1
        
    If txtArticulo(Ind).Text = "" Then
        MsgBox "Seleccione el articulo", vbExclamation
        PonerFoco txtEAN(Ind)
        Exit Sub
    End If
    
    
    Set frmL = New frmAlmPartidas
    frmL.DatosADevolverBusqueda = txtArticulo(Ind).Text
    Codigo = ""
    frmL.Show vbModal
    Set frmL = Nothing
    If Codigo <> "" Then
        C = CCur(RecuperaValor(Codigo, 2))
        If C < 0 Then
            MsgBox "Cantidad negativa.", vbExclamation
        Else
         
             txtLote(Index).Text = RecuperaValor(Codigo, 1)
             PonerFoco Me.txtImporte(Index)
        End If

    End If
    
End Sub

Private Sub txtArticulo_GotFocus(Index As Integer)
    ConseguirFoco txtArticulo(Index), 3
End Sub

Private Sub txtArticulo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub


Private Sub txtArticulo_LostFocus(Index As Integer)
Dim T As String
    
    txtArticulo(Index).Text = Trim(txtArticulo(Index).Text)
    If txtArticulo(Index).Text = "" Then
        'EN blanco
        txtDescArticulo(Index).Text = ""
        txtEAN(Index).Text = ""
        Exit Sub
    End If
    
    
    T = "codigoea"
    Codigo = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", txtArticulo(Index).Text, "T", T)
    If Codigo = "" Then
        MsgBox "No existe el art�culo : " & txtArticulo(Index).Text, vbExclamation
        T = ""
        txtArticulo(Index).Text = ""
        PonerFoco txtArticulo(Index)
        txtEAN(Index).Text = ""
    Else
        'txtArticulo(Index).Text = T
    End If
    Me.txtDescArticulo(Index).Text = Codigo
    txtEAN(Index).Text = T
    Codigo = ""
    
End Sub

Private Sub txtEAN_GotFocus(Index As Integer)
    ConseguirFoco txtImporte(Index), 3
End Sub

Private Sub txtEAN_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtEAN_LostFocus(Index As Integer)
Dim T As String
    
    txtEAN(Index).Text = Trim(txtEAN(Index).Text)
    If txtEAN(Index).Text = "" Then Exit Sub
    
    T = "codartic"
    Codigo = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codigoea", txtEAN(Index).Text, "T", T)
    If Codigo = "" Then
        MsgBox "No existe el art�culo con EAN: " & txtEAN(Index).Text, vbExclamation
        T = ""
        txtArticulo(Index).Text = ""
        PonerFoco txtEAN(Index)
        txtEAN(Index).Text = ""
    Else
       
            'Ok. Nos ponemos en lote
            PonerFoco txtLote(Index)
 
        
    End If
    Me.txtDescArticulo(Index).Text = Codigo
    txtArticulo(Index).Text = T
    Codigo = ""
    
    
    
End Sub

Private Sub txtImporte_GotFocus(Index As Integer)
    ConseguirFoco txtImporte(Index), 3
End Sub

Private Sub txtImporte_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtImporte_LostFocus(Index As Integer)
    txtImporte(Index).Text = Trim(txtImporte(Index).Text)
    If txtImporte(Index).Text = "" Then Exit Sub
    If Index = 0 Then
        If Not PonerFormatoDecimal(txtImporte(Index), 3) Then   '10,2  en formato decimal
            txtImporte(Index).Text = ""
        End If
    End If
End Sub



Private Sub txtCliente_GotFocus(Index As Integer)
    ConseguirFoco txtCliente(Index), 3
End Sub

Private Sub txtCliente_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtCliente_LostFocus(Index As Integer)

    Codigo = ""
    txtCliente(Index).Text = Trim(txtCliente(Index).Text)
    If txtCliente(Index).Text <> "" Then
        If Not IsNumeric(txtCliente(Index).Text) Then
            MsgBox "Campo codigo cliente debe ser num�rico", vbExclamation
            txtCliente(Index).Text = ""
            PonerFoco txtCliente(Index)
        Else
            Codigo = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", txtCliente(Index).Text, "N")
            If Codigo = "" Then
                MsgBox "No existe el cliente : " & txtCliente(Index).Text, vbExclamation
                
                txtCliente(Index).Text = ""
                PonerFoco txtCliente(Index)
            End If
        End If
    End If
    Me.txtDescClie(Index).Text = Codigo
    
    
    
End Sub


Private Sub txtLote_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub


'**************************************************************************
'genera albaran

Private Function GeneraAlbaran(vCli As CCliente) As Boolean
Dim vTipomov As CTiposMov
Dim vCStock As CStock
Dim CPrecioFact As CPreciosFact
Dim cArt As CArticulo
Dim cP As cPartidas
Dim cLot As cLotaje

Dim bol As Boolean
Dim vSQL As String
Dim MenError As String
Dim PorCaja As Boolean
Dim NumCajas As Long
Dim RestoUnid As Long
Dim Precio As String
Dim OrigP As String
    
    On Error GoTo EInsertarOferta
    GeneraAlbaran = False
    'Comprobar si mientras tanto se incremento el contador de Pedidos
    'para ello vemos si existe una oferta con ese contador y si existe la incrementamos
    
    Set vTipomov = New CTiposMov
    Set vCStock = New CStock
    Set cArt = New CArticulo
    Set cP = New cPartidas
    MenError = "leer articulo"
    If Not cArt.LeerDatos(Me.txtArticulo(0).Text) Then Err.Raise 513
    
    
    MenError = "ID trazabilidad"     'FALTA###  codalmac=1
    'vCodartic As String, vCodAlmac As Integer, cLote As String
    If Not cP.LeerDesdeArticulo(cArt.Codigo, 1, Me.txtLote(0).Text) Then
        MsgBox "No se encuentra el lote: " & Me.txtLote(0).Text & " del articulo " & cArt.Nombre, vbExclamation
        Set cP = Nothing
        Exit Function
    End If
    
    
    'Aqui empieza la transaccion. El contador lo dejamos a parte
    Conn.BeginTrans
    
    MenError = "leer movimiento"
    If Not vTipomov.Leer("ALV") Then Err.Raise 513   'Obtenenmos el contador
    
    
    MenError = "Error al insertar en la tabla Cabecera de Albaranes ."
    vSQL = PonerTrabajadorConectado("") 'Codigo trabajador concetado
    If vSQL = "" Then Err.Raise 513
    vCStock.Trabajador = Val(vSQL)
    
    
    
    MenError = "Error al insertar en la tabla Cabecera de Albaranes ."
    vSQL = " insert into `scaalb` (`codtipom`,`numalbar`,`fechaalb`,`factursn`,`codclien`,`nomclien`,"
    vSQL = vSQL & " `domclien`,`codpobla`,`pobclien`,`proclien`,`nifclien`,`telclien`,`coddirec`,"
    vSQL = vSQL & " `nomdirec`,`referenc`,`facturkm`,`cantidkm`,`codtraba`,`codtrab1`,`codtrab2`,`codagent`,"
    vSQL = vSQL & " `codforpa`,`codenvio`,`dtoppago`,`dtognral`,`tipofact`,"
    vSQL = vSQL & " `observa01`,`observa02`,`observa03`,`observa04`,`observa05`,"
    vSQL = vSQL & " `numofert`,`fecofert`,`numpedcl`,`fecpedcl`,`fecentre`,`sementre`,`codtipmf`,"
    vSQL = vSQL & " `numfactu`,`fecfactu`,`esticket`,`numtermi`,`numventa`,`aportacion`,`observa6`,"
    vSQL = vSQL & " `refproduccion`) values ( "
    
    '`codtipom`,`numalbar`,`fechaalb`,`factursn`,`codclien`,`nomclien`,
    vTipomov.contador = vTipomov.ConseguirContador(vTipomov.TipoMovimiento)
    vSQL = vSQL & "'" & vTipomov.TipoMovimiento & "'," & vTipomov.contador & "," & DBSet(Now, "F") & ",1," & vCli.Codigo & "," & DBSet(vCli.Nombre, "T") & ","

    ' `domclien`,`codpobla`,`pobclien`,`proclien`,`nifclien`,`telclien`,`coddirec`,"
    vSQL = vSQL & DBSet(vCli.Domicilio, "T") & "," & DBSet(vCli.CPostal, "T") & "," & DBSet(vCli.Poblacion, "T") & ","
    vSQL = vSQL & DBSet(vCli.Provincia, "T") & "," & DBSet(vCli.NIF, "T") & "," & DBSet(vCli.TfnoClien, "T") & ",NULL,"
    ' `nomdirec`,`referenc`,`facturkm`,`cantidkm`,`codtraba`,`codtrab1`,`codtrab2`,`codagent`,"
    vSQL = vSQL & "NULL,NULL,0,0," & vCStock.Trabajador & "," & vCStock.Trabajador & "," & vCStock.Trabajador & "," & vCli.Agente
    ' `codforpa`,`codenvio`,`dtoppago`,`dtognral`,`tipofact`,"
    vSQL = vSQL & "," & vCli.ForPago & "," & vCli.FEnvio & ",0,0," & vCli.TipoFactu
    ' `observa01`,`observa02`,`observa03`,`observa04`,`observa05`,"
    vSQL = vSQL & ",NULL,NULL,NULL,NULL,NULL"
    ' `numofert`,`fecofert`,`numpedcl`,`fecpedcl`,`fecentre`,`sementre`,`codtipmf`,"
    vSQL = vSQL & ",NULL,NULL,NULL,NULL,NULL,NULL,NULL"
    ' `numfactu`,`fecfactu`,`esticket`,`numtermi`,`numventa`,`aportacion`,`observa6`,refproduccion"
     vSQL = vSQL & ",NULL,NULL,0,NULL,NULL,NULL,NULL,NULL)"
    
    Conn.Execute vSQL, , adCmdText
    
 
    'Actualizar el campo fechamov (ult. movimiento) de la tabla de clientes (sclien)
    MenError = "Actualizando Fecha Movimiento del Cliente."
    vCli.ActualizaUltFecMovim Now
    
    'Linea del albaran
    '----------------------------------------------------------------------------------------
    vCStock.tipoMov = "S"
    vCStock.DetaMov = vTipomov.TipoMovimiento
    vCStock.Documento = vTipomov.contador
    vCStock.Fechamov = Now
    vCStock.codArtic = cArt.Codigo
    vCStock.codalmac = 1 'FALTA###
    vCStock.Cantidad = ImporteFormateado(Me.txtImporte(0).Text)
    vCStock.LineaDocu = 1
    vCStock.HoraMov = Now
    
    'Comprobar si se mueve mover stock (hay stock, o si no hay pero no control de stock)
    bol = True
    If vCStock.MueveStock Then
        bol = vCStock.MoverStock(False)
    End If
    
    
    
    If bol Then
                Set CPrecioFact = New CPreciosFact
                'Si se puede vender por cajas(devuelve>1) poner numero de cajas en una linea con el
                'precio de caja, y otra linea con el resto unidades un precio unidad
                
                NumCajas = CPrecioFact.ObtenerNumCajas(vCStock.Cantidad, CStr(cArt.UnidCaja))
                RestoUnid = CLng(ComprobarCero(vCStock.Cantidad)) - NumCajas * CInt(cArt.UnidCaja)
                CPrecioFact.CodigoLista = vCli.Tarifa
                CPrecioFact.CodigoArtic = vCStock.codArtic
                CPrecioFact.CodigoClien = vCli.Codigo
                PorCaja = (NumCajas > 0)
                Precio = CPrecioFact.ObtenerPrecio(PorCaja, Now, OrigP)
                'Si PorCaja vuelve de ObtenerPrecio a false se calcula con precio unidad aunque NumCajas>0
                'Ya que a regresado con pvp del Articulo
                If PorCaja And NumCajas > 0 And RestoUnid > 0 Then
                    vSQL = "El Art�culo puede venderse por Cajas " & vbCrLf
                    vSQL = vSQL & vbCrLf & "Inserte dos Lineas:   "
                   ' vSQL = vSQL & vbCrLf & "   Linea 1:  " & NumCajas * CInt(devuelve) & " uds a Precio Caja"
                   ' vSQL = vSQL & vbCrLf & "   Linea 2:  " & CInt(Cantidad) - NumCajas * CInt(devuelve) & " uds a Precio Unidad"
                    MsgBox vSQL, vbInformation
                    bol = False
                End If
                    
                    
                    
                vCStock.Importe = CCur(Precio)
                vCStock.Importe = vCStock.Importe * vCStock.Cantidad   'precio * cantidad
    
    End If
    
    'Inserta en tabla "slialb"
    If bol Then
        vSQL = "INSERT INTO slialb"
        vSQL = vSQL & "(codtipom, numalbar,numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, precioar, "
        vSQL = vSQL & "dtoline1, dtoline2, importel, origpre,codprovex,cajas,PrecioLitro,palets,hectogrado) "
        vSQL = vSQL & "VALUES ('" & vTipomov.TipoMovimiento & "', " & vTipomov.contador & ",1 , " & vCStock.codalmac & ","
        vSQL = vSQL & DBSet(vCStock.codArtic, "T") & ", " & DBSet(cArt.Nombre, "T") & ", NULL , "
        vSQL = vSQL & DBSet(vCStock.Cantidad, "N") & ", "
        vSQL = vSQL & DBSet(Precio, "N") & ", " & DBSet(CPrecioFact.Descuento1, "N") & ", "
        vSQL = vSQL & DBSet(CPrecioFact.Descuento2, "N") & ","
        'Importe
        vSQL = vSQL & DBSet(vCStock.Importe, "N") & ", "
        vSQL = vSQL & DBSet(OrigP, "T", "N") & ","
        vSQL = vSQL & DBSet(cArt.codProve, "N", "N") & ","
        vSQL = vSQL & DBSet(NumCajas, "N", "N") & ","
        If cArt.LitrosxUd = 1 Then
            OrigP = Precio
        Else
            OrigP = Round2(CCur(Precio) / cArt.LitrosxUd, 4)
        End If
        vSQL = vSQL & DBSet(OrigP, "N", "N") & "," 'precio litro
        vSQL = vSQL & "0," 'palets
        vSQL = vSQL & "1)" 'hectogrado es un UNO
    
        MenError = "Insertar linea."
        Conn.Execute vSQL
    
    
    
        'LOTE y Partida
        MenError = "Error movimientos lotes."
        Set cLot = New cLotaje
        cLot.codArtic = vCStock.codArtic
        cLot.codalmac = vCStock.codalmac
        cLot.DetaMov = vCStock.DetaMov
        cLot.Documento = vCStock.Documento
        cLot.ProvCliTra = vCStock.Trabajador
        cLot.Numlote = cP.Numlote
        cLot.Cantidad = vCStock.Cantidad
        cLot.LineaDocu = 1
        cLot.SubLinea = 1 'La sublinea del lote 'Normalmente 1 o 2
        cLot.Fechamov = vCStock.Fechamov
        cLot.HoraMov = vCStock.HoraMov
        cLot.tipoMov = 0
        cLot.InsertarLote
        
        cP.IncrementarCantidad -vCStock.Cantidad


        MenError = "Lotes albaran."
        vSQL = "insert into `slialblotes` (`codtipom`,`numalbar`,`numlinea`,`linea`,`numlote`,cantidad) "
        vSQL = vSQL & "VALUES ('" & vTipomov.TipoMovimiento & "', " & vTipomov.contador & "," & vCStock.LineaDocu & " , " & cLot.SubLinea
   
        'Ahora
        vSQL = vSQL & ",'" & DevNombreSQL(cP.Numlote) & "'," & DBSet(vCStock.Cantidad, "N") & ")"
        Conn.Execute vSQL
         
        
    

    

    
    
    
    
        MenError = "Error al actualizar el contador."
        NumRegElim = vTipomov.contador  'para la impresion
        vTipomov.contador = vTipomov.contador - 1
        vTipomov.IncrementarContador vTipomov.TipoMovimiento

        GeneraAlbaran = True
        
    End If
    
EInsertarOferta:
    If Err.Number <> 0 Then
        MuestraError Err.Number, MenError, Err.Description
        Conn.RollbackTrans
    Else
        Conn.CommitTrans
    End If
    Set vTipomov = Nothing
    Set vCStock = Nothing
    Set cArt = Nothing
    Set CPrecioFact = Nothing
    Set cP = Nothing
    Set cLot = Nothing
End Function




Private Sub ImprimeAlbaran()
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
'Dim cadSelect As String 'select para insertar en tabla temporal
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim ImpresionDirecta As Boolean

    '------------------------ Esta copiado del boton imprimir en albaran

                
    If Not PonerParamRPT(10, cadParam, numParam, nomDocu, ImpresionDirecta) Then Exit Sub
   
    'A�adir el codigo de usuario como parametro para link con tabla Temporal (tmptiposiva) en el Report
    'tabla temporal para el calculo del bruto total para cada tipo de IVA
    cadParam = cadParam & "pCodUsu=" & vUsu.Codigo & "|"
    numParam = numParam + 1
    
    
    
    'PUNTO VERDE
    cadParam = cadParam & "PuntoVerde= """ & vParamAplic.ArtReciclado & """|"
    numParam = numParam + 1
    
    
    'Si se imprimen importes y/o
    Codigo = "tipoiva"  'tpo IVA cliente
    devuelve = DevuelveDesdeBD(conAri, "albarcon", "sclien", "codclien", Me.txtCliente(0).Text, "N", Codigo)
    If devuelve = "" Then devuelve = "0"
    ' 0 "Todo"
    ' 1 "Cantidad y Precio"
    ' 2 "Cantidad"
    cadParam = cadParam & "Albarcon=" & devuelve & "|"
    numParam = numParam + 1

    frmImprimir.NombreRPT = nomDocu
        
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion N� de Albaran
    '---------------------------------------------------
    
    'Cod Tipo Movimiento
    devuelve = "{scaalb.codtipom}='ALV'"
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    'N� Albaran
    devuelve = "{scaalb.numalbar}=" & NumRegElim
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    
    
        
    '=========================================================================
    'Aqui sabemos que valor tiene CodClien y a�adimos a los parametros el tipo de IVA
    'que se aplica a ese cliente
    'devuelve = DevuelveDesdeBDNew(conAri, "sclien", "tipoiva", "codclien", Text1(4).Text, "N")
    'esta en cadselect
    
    If Codigo = "" Then Codigo = "0"
    cadParam = cadParam & "pTipoIVA=" & Codigo & "|"
    numParam = numParam + 1


'
'    If ImpresionDirecta Then
'        'Imrpimie directamente. Tipo 4tonda
'        If MsgBox("�Imprimir el albar�n?", vbQuestion + vbYesNo) = vbYes Then ImprimirDirectoAlb cadSelect
'    Else
        With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .opcion = 45 'Es el que esta en frmfacentalb
            .Titulo = "Albaran de Cliente"
            .ConSubInforme = True
            .Show vbModal
        End With
'    End If

End Sub


Private Function HacerTrasiego()
Dim Ok As Boolean
    'Insertamos en scatra y en slitra
    'despues abriremos el frmTraspaso y que actualice directamente
    Conn.BeginTrans
    
    Ok = InsertarTablasTraspaso
    If Ok Then
        Conn.CommitTrans
        Espera 0.3
        Screen.MousePointer = vbHourglass
            With frmAlmMovimientos2
                .TrasiegoMoixent = True  'UNICO SITIO A TRUE
                .EsHistorico = False
                .hcoCodMovim = NumRegElim
                .hcoFechaMovim = Now
                .Show vbModal
            End With
        
        If CadenaDesdeOtroForm = "" Then
            'Sginifica que ha ido bien
            limpiar Me
            PonerFoco Me.txtArticulo(1)
            MsgBox "Proceso finalizado con exito", vbInformation
        Else
            MsgBox "Avise soporte tecnico", vbCritical
        End If
    Else
        Conn.RollbackTrans
    End If
End Function


Private Function InsertarTablasTraspaso() As Boolean
Dim vTipomov As CTiposMov
Dim B As Boolean

    InsertarTablasTraspaso = False
    
    Set vTipomov = New CTiposMov
    NumRegElim = -1
    B = False
    If vTipomov.Leer("REG") Then
  
    
        'cabecera
        'scamov codmovim,codalmac,fecmovim,codtraba,observa1,situacio
        Codigo = PonerTrabajadorConectado("")  'NO SERA "", ya lo he comprobado arriba
        Codigo = " VALUES (" & vTipomov.contador & ",1,now()," & Codigo & ","
        Codigo = Codigo & DBSet(Me.txtObslin(0), "T", "S") & ",1)"  '1: YA esta impreso
        Codigo = "INSERT INTO scamov(codmovim,codalmac,fecmovim,codtraba,observa1,situacio)" & Codigo
        B = EjecutaSQL(conAri, Codigo, True)
        
        If B Then
            'lineas
            'slimov codmovim ,numlinea,codartic,cantidad,tipomovi,motimovi,numlote,
            Codigo = ""
            For NumRegElim = 1 To 2
                Codigo = Codigo & ", (" & vTipomov.contador & "," & NumRegElim & "," & DBSet(Me.txtArticulo(1).Text, "T") & "," & DBSet(Me.txtImporte(NumRegElim), "N")
                Codigo = Codigo & "," & Abs(NumRegElim = 2) & "," & DBSet(txtObslin(NumRegElim), "T") & "," & DBSet(txtLote(NumRegElim).Text, "T") & ")"
            Next
            
            Codigo = Mid(Codigo, 2)
            Codigo = "INSERT INTO slimov(codmovim ,numlinea,codartic,cantidad,tipomovi,motimovi,numlote) VALUES " & Codigo
            B = EjecutaSQL(conAri, Codigo, True)
                    
        End If
        
        If B Then
            NumRegElim = vTipomov.contador
            vTipomov.IncrementarContador ("REG")
        Else
            NumRegElim = -1
        End If
    End If
    Set vTipomov = Nothing
    InsertarTablasTraspaso = B
End Function


Private Sub txtObslin_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub
