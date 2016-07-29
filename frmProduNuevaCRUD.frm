VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProduNuevaCRUD2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Linea producción"
   ClientHeight    =   9270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   16170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FramePalets 
      Height          =   2175
      Left            =   4335
      TabIndex        =   66
      Top             =   4320
      Width           =   4080
      Begin MSComctlLib.ListView lwPalet 
         Height          =   1695
         Left            =   120
         TabIndex        =   67
         Top             =   240
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   2990
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Id"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Caj. pal."
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Cajas"
            Object.Width           =   1323
         EndProperty
      End
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   5400
      TabIndex        =   46
      Top             =   8760
      Width           =   1455
   End
   Begin VB.Frame FrameCambioLote 
      Caption         =   "Cambio LOTE mataria prima/auxiliar"
      Height          =   2175
      Left            =   360
      TabIndex        =   36
      Top             =   4320
      Visible         =   0   'False
      Width           =   3975
      Begin VB.CheckBox chkFin 
         Caption         =   "Fin dep."
         Height          =   255
         Left            =   2880
         TabIndex        =   71
         Top             =   1560
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   8
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   43
         Text            =   "Text1"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton cmdModLote 
         Height          =   375
         Index           =   1
         Left            =   2400
         Picture         =   "frmProduNuevaCRUD.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   375
         Index           =   7
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   40
         Text            =   "Text1"
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000018&
         Height          =   375
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   38
         Text            =   "Text1"
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Height          =   375
         Index           =   6
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   37
         Text            =   "Text1"
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "NUEVO lote"
         Height          =   195
         Index           =   9
         Left            =   1440
         TabIndex        =   44
         Top             =   1200
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Lote anterior"
         Height          =   195
         Index           =   8
         Left            =   240
         TabIndex        =   41
         Top             =   1200
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Materia prima  / auxiliar"
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   39
         Top             =   360
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmdAceptarCantidad 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   8760
      Width           =   1455
   End
   Begin VB.CommandButton cmdIniciarProduccio 
      Caption         =   "Iniciar producción"
      Height          =   375
      Left            =   5040
      TabIndex        =   22
      Top             =   8760
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   7080
      TabIndex        =   21
      Top             =   8760
      Width           =   1335
   End
   Begin VB.CommandButton cmdAsignarProd 
      Height          =   375
      Left            =   840
      Picture         =   "frmProduNuevaCRUD.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Asignar nueva produccion"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox txtNomartic 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   120
      Width           =   4935
   End
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   8175
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   15
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   48
         TabIndex        =   73
         Text            =   "Text1"
         Top             =   3120
         Width           =   4575
      End
      Begin VB.ComboBox cboTipoImpresion 
         Height          =   315
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Tag             =   "Tipo|N|N|0||prodpalets|TipoImpresion|||"
         Top             =   690
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   14
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   48
         TabIndex        =   60
         Text            =   "Text1"
         Top             =   2760
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   13
         Left            =   840
         TabIndex        =   59
         Text            =   "Text1"
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   315
         Index           =   12
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   56
         Text            =   "Text1"
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         ForeColor       =   &H00808080&
         Height          =   315
         Index           =   11
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   51
         Text            =   "Text1"
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   10
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   48
         Text            =   "Text1"
         Top             =   1740
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   9
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   47
         Text            =   "Text1"
         Top             =   1260
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   1740
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   1260
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   690
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Linea extra 2"
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   74
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Impr. palets"
         Height          =   195
         Index           =   17
         Left            =   2640
         TabIndex        =   69
         Top             =   750
         Width           =   885
      End
      Begin VB.Line Line2 
         X1              =   910
         X2              =   910
         Y1              =   1200
         Y2              =   2160
      End
      Begin VB.Label lblManual 
         Caption         =   "MANUAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   495
         Left            =   4680
         TabIndex        =   65
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Linea extra 1"
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   62
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha cad."
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   61
         Top             =   2310
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Meses para la fecha caducidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   1320
         TabIndex        =   58
         Top             =   2310
         Width           =   2415
      End
      Begin VB.Label Label2 
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   57
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   120
         X2              =   5520
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   120
         X2              =   5040
         Y1              =   1150
         Y2              =   1150
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Caj/Pal"
         Height          =   195
         Index           =   12
         Left            =   150
         TabIndex        =   52
         Top             =   1320
         Width           =   645
      End
      Begin VB.Label Label2 
         Caption         =   "Cajas esti."
         Height          =   195
         Index           =   11
         Left            =   1080
         TabIndex        =   50
         Top             =   1320
         Width           =   720
      End
      Begin VB.Label Label2 
         Caption         =   "Cajas prod."
         Height          =   195
         Index           =   10
         Left            =   1080
         TabIndex        =   49
         Top             =   1800
         Width           =   795
      End
      Begin VB.Label Label2 
         Caption         =   "Cant. prod"
         Height          =   255
         Index           =   6
         Left            =   2880
         TabIndex        =   20
         Top             =   1770
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Cant. estimada"
         Height          =   255
         Index           =   5
         Left            =   2880
         TabIndex        =   18
         Top             =   1290
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Linea"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   5640
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblLINEA 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   96
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2655
         Left            =   6120
         TabIndex        =   15
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Id trazabilidad"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Id sublinea"
         Height          =   195
         Index           =   1
         Left            =   2640
         TabIndex        =   9
         Top             =   300
         Width           =   765
      End
      Begin VB.Label Label2 
         Caption         =   "Id produccion"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame FrameIntroduccionCantidad 
      Height          =   1935
      Left            =   360
      TabIndex        =   28
      Top             =   6480
      Visible         =   0   'False
      Width           =   8055
      Begin VB.CheckBox chkFindepositoEnCierreLinea 
         Caption         =   "Fin dep."
         Height          =   255
         Left            =   6720
         TabIndex        =   72
         Top             =   240
         Width           =   975
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Width           =   7695
         Begin VB.TextBox TxtUD 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Enabled         =   0   'False
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
            Left            =   3000
            TabIndex        =   30
            Text            =   "Text2"
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
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
            Text            =   "Text2"
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
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
            Left            =   4440
            TabIndex        =   2
            Text            =   "Text2"
            Top             =   360
            Width           =   3135
         End
         Begin VB.Label Label4 
            Caption         =   "="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   4
            Left            =   4080
            TabIndex        =   34
            Top             =   480
            Width           =   315
         End
         Begin VB.Label Label4 
            Caption         =   "x"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   3
            Left            =   2280
            TabIndex        =   33
            Top             =   480
            Width           =   315
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Unidades"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   2
            Left            =   5760
            TabIndex        =   32
            Top             =   0
            Width           =   1755
         End
         Begin VB.Label Label4 
            Caption         =   "Cajas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   0
            Left            =   240
            TabIndex        =   31
            Top             =   0
            Width           =   915
         End
      End
      Begin VB.Label Label4 
         Caption         =   "Introduzca la cantidad producida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   120
         TabIndex        =   35
         Top             =   120
         Width           =   5775
      End
   End
   Begin VB.Frame FrameLine 
      Height          =   4335
      Left            =   240
      TabIndex        =   23
      Top             =   4320
      Width           =   8175
      Begin VB.Frame FramePesosLotes 
         Height          =   1695
         Left            =   120
         TabIndex        =   106
         Top             =   2520
         Width           =   7815
         Begin MSComctlLib.ListView ListView3 
            Height          =   1335
            Left            =   120
            TabIndex        =   107
            Top             =   240
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   2355
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Tipo"
               Object.Width           =   2558
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Articulo"
               Object.Width           =   5998
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Peso"
               Object.Width           =   2363
            EndProperty
         End
      End
      Begin VB.CommandButton cmdCambiarTipoImpresionPalet 
         Height          =   375
         Left            =   7500
         Picture         =   "frmProduNuevaCRUD.frx":61E4
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   "Cambiar tipo impresion palet"
         Top             =   1560
         Width           =   375
      End
      Begin VB.CommandButton cmdLinExtra 
         Height          =   375
         Left            =   7500
         Picture         =   "frmProduNuevaCRUD.frx":7C56
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Cambiar datos linea extra"
         Top             =   2030
         Width           =   375
      End
      Begin VB.CommandButton cmdFecCad 
         Height          =   375
         Left            =   7500
         Picture         =   "frmProduNuevaCRUD.frx":8658
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "Cambiar fecha caducidad"
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton cmdControlProduccion 
         Height          =   375
         Left            =   7500
         Picture         =   "frmProduNuevaCRUD.frx":8BE2
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Impr. control produccion"
         Top             =   640
         Width           =   375
      End
      Begin VB.CommandButton cmdVerLote 
         Height          =   375
         Left            =   1560
         Picture         =   "frmProduNuevaCRUD.frx":F434
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Lote"
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdImpr 
         Height          =   375
         Left            =   7500
         Picture         =   "frmProduNuevaCRUD.frx":FE36
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Imprimir etiquetas caja"
         Top             =   240
         Width           =   375
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   1575
         Left            =   120
         TabIndex        =   45
         Top             =   2520
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   2778
         _Version        =   393217
         Indentation     =   471
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton cmdModLote 
         Height          =   375
         Index           =   0
         Left            =   1560
         Picture         =   "frmProduNuevaCRUD.frx":103C0
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   120
         Width           =   375
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1455
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   2566
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   2558
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Articulo"
            Object.Width           =   5998
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   1834
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "LOTE"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "LOTEANTERIOR"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Aceite(Si no"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Depostio"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Shape Shape1 
         Height          =   2315
         Left            =   7320
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "X"
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
         Index           =   1
         Left            =   120
         TabIndex        =   27
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Componentes"
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
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame FramePesajes 
      Height          =   8415
      Left            =   8640
      TabIndex        =   75
      Top             =   120
      Width           =   7335
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   645
         Index           =   12
         Left            =   4200
         TabIndex        =   105
         Text            =   "Text5"
         Top             =   7680
         Width           =   2895
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   11
         Left            =   2640
         TabIndex        =   103
         Text            =   "Text5"
         Top             =   7920
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   1200
         TabIndex        =   101
         Text            =   "Text5"
         Top             =   7920
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   5400
         TabIndex        =   99
         Text            =   "Text5"
         Top             =   7200
         Width           =   855
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   4200
         TabIndex        =   97
         Text            =   "Text5"
         Top             =   7200
         Width           =   855
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   2760
         TabIndex        =   95
         Text            =   "Text5"
         Top             =   7200
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   1440
         TabIndex        =   93
         Text            =   "Text5"
         Top             =   7200
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   480
         TabIndex        =   91
         Text            =   "Text5"
         Top             =   7200
         Width           =   855
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   6120
         TabIndex        =   87
         Text            =   "Text5"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   4920
         TabIndex        =   85
         Text            =   "Text5"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   3840
         TabIndex        =   83
         Text            =   "Text5"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2640
         TabIndex        =   81
         Text            =   "Text5"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   79
         Text            =   "Text5"
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   0
         TabIndex        =   78
         Top             =   960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox cboSerie 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   77
         Top             =   480
         Width           =   855
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   5415
         Left            =   240
         TabIndex        =   89
         Top             =   960
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   9551
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
            Text            =   "Id"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Llena(gr)"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Gr. aceite"
            Object.Width           =   1676
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Volumen"
            Object.Width           =   2029
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Emp"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Err1"
            Object.Width           =   1146
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Err2"
            Object.Width           =   1147
         EndProperty
      End
      Begin VB.Label Label5 
         Caption         =   "Desv.tip"
         Height          =   375
         Index           =   13
         Left            =   2640
         TabIndex        =   104
         Top             =   7680
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Promedio"
         Height          =   375
         Index           =   12
         Left            =   1200
         TabIndex        =   102
         Top             =   7680
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "No Emp2"
         Height          =   375
         Index           =   11
         Left            =   5400
         TabIndex        =   100
         Top             =   6960
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "NO emp"
         Height          =   375
         Index           =   10
         Left            =   4200
         TabIndex        =   98
         Top             =   6960
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Mínimo"
         Height          =   375
         Index           =   9
         Left            =   2880
         TabIndex        =   96
         Top             =   6960
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Máximo"
         Height          =   375
         Index           =   8
         Left            =   1560
         TabIndex        =   94
         Top             =   6960
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Nº pesos"
         Height          =   375
         Index           =   7
         Left            =   480
         TabIndex        =   92
         Top             =   6960
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Control media"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   90
         Top             =   6600
         Width           =   1335
      End
      Begin VB.Line Line3 
         X1              =   1320
         X2              =   6720
         Y1              =   6720
         Y2              =   6720
      End
      Begin VB.Label Label5 
         Caption         =   "Otros"
         Height          =   375
         Index           =   5
         Left            =   6120
         TabIndex        =   88
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Etiqueta"
         Height          =   375
         Index           =   4
         Left            =   4920
         TabIndex        =   86
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Tapon"
         Height          =   375
         Index           =   3
         Left            =   3960
         TabIndex        =   84
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Botella"
         Height          =   375
         Index           =   2
         Left            =   2640
         TabIndex        =   82
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Nominal"
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   80
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Serie:"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   76
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Artículo"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   13
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   8760
      Width           =   3495
   End
End
Attribute VB_Name = "frmProduNuevaCRUD2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Modo As Byte   '1  Ver,   2 Modificar   3 Cerrar
Public cLP As cLineaProduccion   'Si es nuevo es NOTHING y en modo tendremos la LINEA

Public SubLinea As Byte   'Para el cambio de lote

Private WithEvents frmL As frmAlmPartidas
Attribute frmL.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmB2 As frmBuscaGrid
Attribute frmB2.VB_VarHelpID = -1

Dim SePuedeSalir As Boolean
Dim SQL As String
Dim PrimVez As Boolean

'Para los nuevos
Dim idProd As Long
Dim LinProd As Integer


Dim NumDeposito As Byte


Private Sub cboSerie_Click()
    CargarPesos
End Sub

Private Sub cmdAceptarCantidad_Click()
Dim Can As Currency
Dim Cajas As Integer
Dim C As Long
Dim cL As cLineaProCompo
Dim Cp As cPartidas
Dim CajasDistintas As String


Dim FinDepositoLote As Boolean
Dim NUevoDeposito As Integer
    CadenaDesdeOtroForm = ""
    If Modo = 2 Then
        If Text1(8).Text = "" Then
            MsgBox "Falta NUEVO lote", vbExclamation
            Exit Sub
        Else
            If Text1(8).Text = Text1(7).Text Then
                MsgBox "Mismo lote", vbExclamation
                Exit Sub
            End If
        End If
        
        
    End If

    If Text2.Text = "" Then
        Can = 0
    Else
        Can = ImporteFormateado(Text2.Text)
    End If
    If Can <= 0 Then
        
            If MsgBox("Desea eliminar de la linea?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        
    End If
    
    
    
    
    If Text3.Text = "" Then
        Cajas = 0
    Else
        Cajas = Val(Text3.Text)
    End If
    
    
  
    
    'Cajas teoricas
    '-------------
    C = Can \ CInt(Me.TxtUD.Text)
    If (Can Mod CInt(Me.TxtUD.Text)) > 0 Then C = C + 1
    
    C = C * CInt(Me.TxtUD.Text)  'Cantidad producida si llenaramos las cajas
    
    SQL = String(40, "-") & vbCrLf
    SQL = SQL & vbCrLf & "UNIDADES: " & Format(Val(Can), "#,###,##0") & vbCrLf & "Cajas:        " & Cajas & vbCrLf
    If Val(Can) <> C Then
        
        'Cantidad de cajas a producir  distinto
        C = C - CInt(Can)
        C = CInt(Me.TxtUD.Text) - C
        SQL = SQL & vbCrLf & "Cajas incompletas " & vbCrLf & "Una cajas con  " & C & " uds" & vbCrLf
    End If
    SQL = vbCrLf & SQL & String(40, "-")
    
    
    CajasDistintas = ""
    If vParamAplic.ProduccionNueva Then
        'aqui aqui aqui
        C = cLP.CajasLeidasLector
        If Cajas <> C Then
            
            CajasDistintas = String(17, "*")
            CajasDistintas = CajasDistintas & CajasDistintas
            CajasDistintas = CajasDistintas & vbCrLf & CajasDistintas
            
            SQL = SQL & vbCrLf & vbCrLf & CajasDistintas & vbCrLf & vbCrLf
            SQL = SQL & " N O      D E B E R I A      C O N T I N U AR " & vbCrLf & vbCrLf
            SQL = SQL & CajasDistintas & vbCrLf
            
            'Guardo un log
            CajasDistintas = "CIERRE PROD. Cantidad: " & C & "   Cantidad indicada para el cierre: " & Cajas
            CajasDistintas = DBSet(CajasDistintas, "T")
            CajasDistintas = "INSERT INTO proderrorcierrepalet(Fechora,idpalet,observaciones) VALUES (now(),0," & SQL & ")"
       
        End If
    End If
        
    SQL = SQL & vbCrLf & "¿CONTINUAR?"
    
    
    If Modo = 2 Then
        SQL = "Va a proceder con el cambio de lote habiendo producido: " & SQL
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Else
        'Cierre produccion
        SQL = "Finalizar la produccion: " & SQL
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        
    End If
    
    
    
    'Ha intentado cerrar con cajas <> e las leidas
    If CajasDistintas <> "" Then EjecutaSQL conAri, CajasDistintas, False
    
    CadenaDesdeOtroForm = "OK"
    If Modo = 2 Then

        'JUNIO 2014
        'FALTA###
        'De momento esta solo para el aceite. Tambien podremos regularizar las
        'partidas cuando sean final de lote
        FinDepositoLote = False
        NUevoDeposito = 0
        If Me.chkFin.visible Then
            If Me.chkFin.Value = 1 Then FinDepositoLote = True
            
            'Vere, el nuevo LOTE, ande esta, en que depos
            SQL = DevuelveDesdeBD(conAri, "numdeposito", "proddepositos", "numlote", Text1(8).Text, "T")
            If SQL = "" Then
                SQL = "Error obteniendo deposito." & vbCrLf & " Avise soporte técnico" & vbCrLf & vbCrLf & "El programa continuará"
                MsgBox SQL, vbExclamation
            Else
                NUevoDeposito = Val(SQL)
            End If
        End If
    
        If cLP.CerrarParaCambioLote(Can, Cajas, CInt(Me.SubLinea), Text1(8).Text, FinDepositoLote, NUevoDeposito) Then
            SePuedeSalir = True
            Cajas = 1 'Todo mal
        
            'Deberiamos buscar parar marcar la etiqueta nueva
            If cLP.DevuelveComponenteLinea(CInt(Me.SubLinea), cL) Then
                '----------------------------------------------------
                Set Cp = New cPartidas             'FALTA EL ALMACEN
                If Cp.LeerDesdeArticulo(cL.codarticCompo, 1, cL.LoteMateria) Then
                    'Veremos las etiquetas
                    '-----------------------------------
                    'Si hay una etiqueta libre
                    SQL = "fechaulizada is null and id "
                    SQL = DevuelveDesdeBD(conAri, "min(bulto)", "spartidaslin", SQL, Cp.idPartida)
                    If SQL = "" Then
                        'No hay ninguna libre
                        'Veo a ver si hay
                        SQL = DevuelveDesdeBD(conAri, "bulto", "spartidaslin", "id", Cp.idPartida)
                        If SQL = "" Then
                            SQL = "ERROR leyendo etiquetas. No hay ninguna etiqueta para " & Cp.codartic
                        
                        Else
                            MsgBox "No existe etiqueta libre", vbExclamation
                            SQL = " WHERE id = " & Cp.idPartida & " AND bulto = " & SQL
                            SQL = "UPDATE spartidaslin Set fechaulizada = " & DBSet(Now, "FH") & SQL
                            EjecutaSQL conAri, SQL, True
                            Cajas = 0
                        End If
                    Else
                        'Si que hay libre
                        SQL = " WHERE id = " & Cp.idPartida & " AND bulto = " & SQL
                        SQL = "UPDATE spartidaslin Set fechaulizada = " & DBSet(Now, "FH") & SQL
                        EjecutaSQL conAri, SQL, True
                        Cajas = 0
                    End If
                    
                Else
                    SQL = "Error leyendo partida: " & cL.codarticCompo & "   Lote  " & cL.LoteMateria
                End If
                If Cajas = 1 Then
                    MsgBox SQL, vbExclamation
                Else
                    'OK . LAnzamos impresion etiquetas
                    
                    'Ahora YA NO SE IMPRIMEN DESDE AQUI
                    'LanzaImpresionEtiquetas
                End If
            Else
                
            End If
            Unload Me
        End If
    Else
        'MsgBox "FALTA###"
        'Habria que ver si es final de existencias. Primero de deposito y cuando lo pida "Ramon" de materia auxiliar
        If cLP.CerrarProduccion(Can, Cajas, chkFindepositoEnCierreLinea.Value = 1) Then
            'Ya no se imprimien etiquetas
            'LanzaImpresionEtiquetas
            'IMPRIMI
            SePuedeSalir = True
            Unload Me
        End If
    End If
        
    
End Sub

Private Sub cmdAsignarProd_Click()
    
    If Me.Text1(4).Text <> "" Then
        If MsgBox("Ya existen datos. ¿Borrarlos?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
        'Ha querido volver a pedir datos
        limpiar Me
        Me.ListView1.ListItems.Clear
        InsertarTablaProduccion False
        
    End If
    
    
    
    
    frmProdSeleccionarLineaProd.Show vbModal
    If CadenaDesdeOtroForm <> "" Then
    
        InsertarTablaProduccion True
           
    
        Set cLP = New cLineaProduccion
        If cLP.LeerDeOrdenProduccion(idProd, CLng(LinProd)) Then
            PonerCampos
        Else
            Set cLP = Nothing
        End If
        CadenaDesdeOtroForm = ""
    End If
End Sub

Private Sub InsertarTablaProduccion(Insertar As Boolean)
    If Not Insertar Then
        SQL = "DELETE FROM prodlin where codigo=" & idProd & " AND idlin =" & LinProd
        conn.Execute SQL
        If LinProd = 1 Then
            'Hay que borrar la cabecera tb
            SQL = "DELETE FROM prodcab where codigo=" & idProd
            conn.Execute SQL
        End If
    Else
        If LinProd = 1 Then
        
            SQL = "insert into `prodcab` (`codigo`,`descripcion`,`feccreacion`,`fecproduccion`,`almacen2`,`producido`) values ("
            SQL = SQL & idProd & "," & DBSet(Format(Now, "dd/mm/yyyy"), "T") & "," & DBSet(Now, "F") & ",NULL,1,0)"
            conn.Execute SQL
        End If

        SQL = "INSERT INTO prodlin (codigo ,idlin ,codartic ,cantesti )"
        SQL = SQL & " VALUES (" & idProd & "," & LinProd & ",'" & RecuperaValor(CadenaDesdeOtroForm, 1)
        SQL = SQL & "'," & RecuperaValor(CadenaDesdeOtroForm, 3) & ")"
        conn.Execute SQL
        Espera 0.5
    End If
End Sub





Private Sub cmdCambiarTipoImpresionPalet_Click()
    
    If Me.cboTipoImpresion.ListIndex = 0 Then
        SQL = Me.cboTipoImpresion.List(1)
        SubLinea = 1
    Else
        SQL = Me.cboTipoImpresion.List(0)
        SubLinea = 0
    End If
    
    SQL = "Desea cambiar el tipo de impresion de etiquetas de palet a : " & SQL
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        SQL = "UPDATE prodlin set TipoImpresionPalet=" & SubLinea & " WHERE   codigo = " & cLP.CodProduccion & " AND idlin =" & cLP.idLiProd
        If EjecutaSQL(conAri, SQL, True) Then
            Me.cboTipoImpresion.ListIndex = SubLinea
            cLP.TipoImpresionPalet = SubLinea
            'Cogere el palet que esta en marcha y le cambio el tipo deimpresion d palet
            SQL = cLP.LeerLineaDondeEstaPaletizando
            If SQL <> "" Then
                If InStr(1, SQL, "-") > 0 Then
               
                    SQL = Trim(Mid(SQL, InStr(1, SQL, "-") + 1))
                    SQL = "UPDATE prodpalets set TipoImpresion = " & SubLinea & " WHERE idpalet =" & SQL
                    
                    EjecutaSQL conAri, SQL, False
                End If
            End If
            
        End If
    End If
    SubLinea = 0 'reestablezco
End Sub

Private Sub cmdCancelar_Click()
    SePuedeSalir = True
    If Me.cmdAsignarProd.visible Then InsertarTablaProduccion False

    CadenaDesdeOtroForm = ""
    Unload Me
End Sub


Private Function DatosOk() As Boolean
Dim I As Integer
Dim CambioLote As Byte  'Si ha cambiado algun lote

    
    DatosOk = False
    
    
    If Text1(4).Text = "" Or Me.txtNomartic.Text = "" Then
        MsgBox "Falta articulo", vbExclamation
        Exit Function
    End If
    
    If Modo = 0 Then
        Text1(13).Text = Trim(Text1(13).Text)
        If Text1(13).Text = "" Then
            MsgBox "Faltan meses caducidad", vbExclamation
            Exit Function
        End If
    End If
    
    'Comprobaremos el numero de LOTE de los componentes
    CambioLote = 0
    SQL = ""
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).SubItems(3) = "" Then
            SQL = SQL & "  -" & ListView1.ListItems(I).SubItems(1) & "  --> FALTA LOTE" & vbCrLf
        Else
            If Modo = 2 Then
                If ListView1.ListItems(I).SubItems(3) <> ListView1.ListItems(I).SubItems(4) Then CambioLote = CambioLote + 1
            End If
        End If
    Next
    If SQL <> "" Then
        SQL = SQL & vbCrLf & vbCrLf & "¿Continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
    End If
    If Modo = 2 Then
        'Esta cambiando lote. Si no cambia nada.....
        
        
    End If
    DatosOk = True
End Function


Private Sub cmdControlProduccion_Click()
            
    SQL = "{prodcab.codigo}=" & Text1(0).Text & " AND {prodlin.idlin} = " & Text1(1).Text
    SQL = SQL & " AND {prodtrazcompo.lotetraza}=" & Text1(2).Text
    LlamaImprimirGral SQL, "", 0, "produccionControl.rpt", "Control produccion: " & Text1(0).Text & " - " & Text1(1).Text
    SQL = ""
    
End Sub

Private Sub cmdFecCad_Click()
    Set frmC = New frmCal
    
    If Text1(12).Text <> "" Then
        frmC.Fecha = CDate(Text1(12).Text)
    Else
        'NO DEBERIA HABER PASADO
        frmC.Fecha = DateAdd("yyyy", 22, cLP.FH_Incio)
    End If
    SQL = ""
    frmC.Show vbModal
    Set frmC = Nothing
    If SQL <> "" Then
        If MsgBox("Desea establecer la fecha de caducidad a: " & SQL, vbQuestion + vbYesNo) = vbYes Then
            If MsgBox("Fecha: " & SQL & "         ¿Continuar?      ", vbQuestion + vbYesNo) = vbYes Then
                Text1(12).Text = SQL
                cLP.FechaCaducidad = SQL
                SQL = "UPDATE prodlin set feccaduca =" & DBSet(SQL, "F")
                SQL = SQL & " WHERE codigo = " & cLP.CodProduccion & " AND idlin = " & cLP.idLiProd
                EjecutaSQL conAri, SQL, True
            End If
        End If
    End If
End Sub

Private Sub cmdImpr_Click()
    ImprimeEtiquetas False

End Sub

Private Sub ImprimeEtiquetas(Nuevo As Boolean)
Dim L As Long
Dim I As Integer

    L = 0
    I = cLP.UnidadesCaja
    If I = 0 Then I = 1
    
    SQL = ""
    If Not Nuevo Then
        If cLP.EtiquetasImpresas > 0 Then SQL = "NO"
    End If
    
    If SQL = "" Then
        If MsgBox("Imprimir etiquetas?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        L = cLP.CantidadEstimada \ I
        L = Round(L * 1.1, 0)   'Un 10 % mas
        If L = 0 Then L = 1
    Else
    
        SQL = "0"
        SQL = InputBox("Numero de cajas a imprimir:  (ult: " & cLP.EtiquetasImpresas & ")", "Cajas", SQL)
        
        If SQL <> "" Then
            If IsNumeric(SQL) Then L = CLng(SQL)
        End If
            
    End If
    If L = 0 Then Exit Sub
    
    If ImprimeEtiquetasCajas(cLP.linea, cLP.LoteTrazabilidad, cLP.codartic, cLP.EtiquetasImpresas, L, cLP.LineaExtraEtiquetas, cLP.LineaExtraEtiqueta2) Then
        'Si ha ido bien tengo que updaear las
        L = cLP.EtiquetasImpresas + L
        cLP.EstablecerEtiquetasImpresas L
    End If
    
End Sub
    
    
    
'Private Sub ImpresionCajasAntiguo()
'    SQL = InputBox("Nºincio", "Cajas")
'    i = 0
'    If SQL <> "" Then
'        If IsNumeric(SQL) Then i = CInt(SQL)
'    End If
'    If i = 0 Then Exit Sub
'    Inicio = i
'
'    SQL = InputBox("NºFin", "Cajas")
'    i = 0
'    If SQL <> "" Then
'        If IsNumeric(SQL) Then i = CInt(SQL)
'    End If
'    If i = 0 Then Exit Sub
'    Fin = i
'
'
'    If Inicio > Fin Then Exit Sub
'
'
'
'    SQL = ""
'    'prodcajas lotetraza,idcaja,idpalet,fcreacion
'    For i = Inicio To Fin
'        SQL = SQL & ", (" & Text1(2).Text & "," & i & ",NULL,NULL)"
'    Next i
'    SQL = Mid(SQL, 2)
'    SQL = "INSERT INTO prodcajas(lotetraza,idcaja,idpalet,fcreacion) VALUES " & SQL
'    If EjecutaSQL(conAri, SQL, True) Then
'
'        Espera 0.5
'
'        SQL = " AND {prodcajas.idcaja} >= " & Inicio & " AND {prodcajas.idcaja} <= " & Fin
'
'        LanzaImpresionEtiquetas2 SQL
'
'        SQL = "DELETE FROM prodcajas where prodcajas.lotetraza = " & Text1(2).Text
'        SQL = SQL & " AND prodcajas.idcaja >= " & Inicio & " AND prodcajas.idcaja <= " & Fin
'        Conn.Execute SQL
'    End If
'
'End Sub

Private Sub cmdImprimir_Click()
    
    SQL = "{prodcab.codigo}=" & Text1(0).Text & " AND {prodlin.idlin} = " & Text1(1).Text
    LlamaImprimirGral SQL, "", 0, "produccionNueva.rpt", "Produccion: " & Text1(0).Text & " - " & Text1(1).Text
    SQL = ""
End Sub

Private Sub cmdIniciarProduccio_Click()
Dim I As Byte
Dim VaBien As Boolean
Dim F As Date
Dim TodasMateriasPrimasAsignadas As Byte
Dim ElDeposito As Integer
Dim Aux As Integer

    If Not DatosOk Then Exit Sub
    
    'Ahora asigno los nuevos lotes de produccion
    'Dentro de la funcion hay transacciones...
    
    'Asigno los lotes de MP
    CadenaDesdeOtroForm = ""
    VaBien = True
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).SubItems(3) <> ListView1.ListItems(I).SubItems(4) Then
            SQL = ListView1.ListItems(I).SubItems(3)
            If Not cLP.AsignarLoteLinea(CInt(I), SQL, False) Then
                VaBien = False
                Exit For
            End If
            TodasMateriasPrimasAsignadas = TodasMateriasPrimasAsignadas + 1
            
            If ElDeposito = 0 Then
                'Aun nO ha asignado el deposito
                SQL = DevuelveDesdeBD(conAri, "numdeposito", "proddepositos", "numlote", ListView1.ListItems(I).SubItems(3), "T")
                If SQL <> "" Then ElDeposito = Val(SQL)

            End If
        End If
    Next I
    
    
    
    
    If VaBien Then
        cLP.linea = lblLINEA.Caption
        F = DateAdd("m", Val(Me.Text1(13).Text), cLP.FH_Incio)
        cLP.FechaCaducidad = Format(F, "dd/mm/yyyy")
        cLP.LineaExtraEtiquetas = Trim(Text1(14).Text)
        cLP.LineaExtraEtiqueta2 = Trim(Text1(15).Text)
        cLP.TipoImpresionPalet = Me.cboTipoImpresion.ListIndex
        
        
        'Vemos que deposito esta cogiendo
        
        If cLP.AsignarA_LineaProduccion(ElDeposito) Then
        
                    
                    
                    
                    
            CadenaDesdeOtroForm = "OK"
            'Lanzar impresion etiquetas
            If TodasMateriasPrimasAsignadas = Me.ListView1.ListItems.Count Then
            
                'Ha asignado todas las materias primas con numero de lote
                If Me.lblLINEA.Caption <> "10" Then   'la linea 9 NO se imprime.. de momento
                    'Juli 0212
                    'No se imprme directamente NINGUNA
                    'If MsgBox("Lanzar impresion etiquetas?", vbQuestion + vbYesNo) = vbYes Then ImprimeEtiquetas True
                End If
            End If
        End If
    End If
    SePuedeSalir = True
    Unload Me
    
End Sub

Private Sub cmdLinExtra_Click()



    CadenaDesdeOtroForm = Text1(14).Text & "|" & Text1(15).Text & "|"
    frmListado2.opcion = 31
    frmListado2.Show vbModal
    
    If CadenaDesdeOtroForm <> "" Then
    
        Text1(14).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
        Text1(15).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
        
        SQL = "UPDATE prodlin SET LineaExtraEtiqueta = " & DBSet(Text1(14).Text, "T", "S")
        SQL = SQL & ", LineaExtraEtiqueta2 = " & DBSet(Text1(15).Text, "T", "S")
        SQL = SQL & " WHERE codigo = " & cLP.CodProduccion & " AND idlin = " & cLP.idLiProd
        If EjecutaSQL(conAri, SQL, True) Then
            cLP.LineaExtraEtiquetas = Text1(14).Text
            cLP.LineaExtraEtiqueta2 = Text1(15).Text
        End If
    End If
End Sub

Private Sub cmdModLote_Click(Index As Integer)
Dim C As Currency
Dim Aux As String

    
    'Nuevo y cambio lote materia prima
    If Modo <> 0 And Modo <> 2 Then Exit Sub
    
    
    If Modo = 0 Then
        'Son Nuevos
        SQL = ""
        If Me.ListView1.ListItems.Count = 0 Then
            SQL = "Vacio"
        Else
            If ListView1.SelectedItem Is Nothing Then SQL = "Ninguno seleccionado"
        End If
        If SQL <> "" Then
            MsgBox SQL, vbExclamation
            Exit Sub
        End If
        Aux = ListView1.SelectedItem.Text
    Else

        Aux = Text1(6).Text
        
    End If
    SQL = ""
    
    
    If ListView1.SelectedItem.SubItems(5) = "1" Then
        'Si es del depostio, al final obtendremos el lote
        NumDeposito = 200 'No acepta -1
        ObtenerLoteAceiteDeposito ListView1.SelectedItem.Text
    Else
    
        Set frmL = New frmAlmPartidas
        frmL.DatosADevolverBusqueda = Aux
        frmL.Show vbModal
        Set frmL = Nothing
    End If
    If SQL <> "" Then
        C = CCur(RecuperaValor(SQL, 2))
        If C < 0 Then
            MsgBox "Cantidad negativa.", vbExclamation
        Else
            If Modo = 0 Then
                If C < CCur(ListView1.SelectedItem.SubItems(2)) Then MsgBox "No tiene cantidad suficiente", vbExclamation
            End If
           'YA tengo el LOTE.
           If Modo = 0 Then
                ListView1.SelectedItem.SubItems(3) = RecuperaValor(SQL, 1)
                LeerPesos
           Else
                Text1(8).Text = RecuperaValor(SQL, 1)
                PonerFoco Text3
           End If
        End If
    End If
End Sub

Private Sub cmdVerLote_Click()
    'Ver lote
    If ListView1.ListItems.Count = 0 Then Exit Sub
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    
    frmAlmPartidas.DatosADevolverBusqueda = ListView1.SelectedItem.Text
    frmAlmPartidas.ParaMostrarDesdeNuevaProduccion = ListView1.SelectedItem.SubItems(3)
    frmAlmPartidas.Show vbModal
End Sub




Private Sub Command1_Click()

Dim I As Integer
Dim C As String
Dim Peso As Currency
Dim PesoLlena As Currency
Dim Signo As Integer
Dim Emp As Currency

    Peso = Rnd()
    Peso = 400 + (Peso)
    
    Emp = 500 * 0.03   'Formato * constante dependiente formato

    For I = 1 To 50
        C = "INSERT INTO  prodlinpesos(codigo,idlin,serie,secuencial,fechahora,"
        C = C & "pesoLleno,pesoBotella,pesoTapon,pesoEtiqueta,pesoOtro,volumenLlenado"
        C = C & ",EMP,CumpleEMP,Cumple2EMP) VALUES ("
        
        C = C & cLP.CodProduccion & "," & cLP.idLiProd & "," & Me.cboSerie.ListCount + 1
        C = C & "," & I & "," & DBSet(Now, "FH") & ","

                
                
        Signo = Int((4 * Rnd) + 1)
        If Signo > 3 Then
            Signo = -1
            PesoLlena = (1.1 * Rnd)
        Else
            Signo = 1
            PesoLlena = (3.5 * Rnd)
        End If
        
        
        PesoLlena = 857 + (Signo * PesoLlena)
                
      
        
        ' "pesoLleno,pesoBotella,pesoTapon,pesoEtiqueta,pesoOtro,volumenLlenado"
        C = C & DBSet(PesoLlena, "N") & "," & DBSet(Peso, "N") & ",0.15,0.1,0.08,"
        
        
        
        PesoLlena = PesoLlena - (Peso + 0.15 + 0.1 + 0.08)  'Botella + tapon +
        PesoLlena = PesoLlena / 0.916     'Volumen
        C = C & DBSet(PesoLlena, "N") & "," & DBSet(Emp, "N") & ","
        
        If 500 - PesoLlena > Emp Then
            'NO cumple EMP
            C = C & "0,"
        Else
            C = C & "1,"
        End If
        
        'doble emp
        If 500 - PesoLlena > (2 * Emp) Then
            'NO cumple EMP
            C = C & "0)"
        Else
            C = C & "1)"
        End If
        conn.Execute C
    Next I
    CargaSeries
End Sub

Private Sub Form_Activate()
    If PrimVez Then
        PrimVez = False
        If Modo < 2 Then
        
            If Modo = 1 Then
                CargaSeries
            End If
        
            PonerFocoBtn Me.cmdCancelar
        Else
            PonerFoco Text3
        End If
    End If
End Sub

Private Sub Form_Load()
Dim Cajas As Long

    Me.Icon = frmppal.Icon
    PrimVez = True
    limpiar Me
    SePuedeSalir = False
    
    CargaComboTipoImpresionPalet Me.cboTipoImpresion
    FrameLine.Left = 240
    
    If cLP Is Nothing Then
        'Es nuevo
        lblLINEA.Caption = Modo 'EN modo, cuando clp es nothing llevamos la linea donde trabajamos
        Modo = 0
        Me.Label1.ForeColor = &H808080
        Me.Label1.Caption = "Nueva prod."
        
        
        cmdIniciarProduccio.visible = True
        
        FijarValoresParaInsertProducion
        
        'Caducidad
        Label2(13).Caption = "MESES"
        Me.Text1(12).visible = False
        Me.Text1(13).visible = True
        Text1(13).Text = "18"
        Label2(14).visible = True
        FramePalets.visible = False
        cboTipoImpresion.Enabled = True
        Me.cboTipoImpresion.ListIndex = 0 'Por defecto tipo impr
        
        Label3(1).Caption = "Pesos"
        FramePesosLotes.visible = True
    Else
        Label3(1).Caption = "Historico trazabilidad"
        FramePesosLotes.visible = False
        'Caducidad
        Label2(13).Caption = "F.caducidad"
        Me.Text1(12).visible = True
        Me.Text1(13).visible = False
        Label2(14).visible = False
        
    
    
        cmdIniciarProduccio.visible = False
        lblLINEA.Caption = cLP.linea
        
        cboTipoImpresion.Enabled = False
        PonerCampos
        
            
        
            
            FrameIntroduccionCantidad.visible = Modo > 1
            FrameIntroduccionCantidad.Enabled = True
            'If (vUsu.Codigo Mod 1000) = 0 Then FrameIntroduccionCantidad.Enabled = True
            FrameLine.visible = Modo < 2
            FrameCambioLote.visible = Modo = 2
                
                
            FramePalets.visible = False
            If Modo > 1 Then
                FramePalets.visible = True
                CargaPalets
                If Modo = 2 Then
                    FramePalets.Left = 4215
                    FramePalets.Width = 4080
                Else
                    FramePalets.Left = 0
                    FramePalets.Width = 8500
                End If
            End If
                
                
            Me.TxtUD.Text = cLP.UnidadesCaja
                
                '
                
                'Ahora. Abril 2011
                Cajas = cLP.CajasLeidasLector
                
                Text3.Text = Cajas
                PonerDatosUdsCajas False
                
                Text1(5).Text = Text2.Text
                Text1(10).Text = Cajas
                'End If
            
            'VA A MODIFICAR ALGO
            If Modo = 1 Then
                Label1.Caption = "Ver linea"
                SePuedeSalir = True
            ElseIf Modo = 2 Then
                Label1.Caption = "Cambio lote"
            
            Else
                'modo=3
                Label1.Caption = "Fin produccion"
            End If
        End If
    
    cmdModLote(0).visible = Modo = 0 'or modo =3
    cmdModLote(1).visible = Modo = 2
    cmdImpr.visible = Modo = 1
    Me.cmdControlProduccion.visible = Modo = 1
    cmdVerLote.visible = Modo = 1
    cmdAsignarProd.visible = Modo = 0
    cmdAceptarCantidad.visible = Modo > 1
    cmdImprimir.visible = Modo = 1
    cmdLinExtra.visible = Modo = 1
    Me.cmdCambiarTipoImpresionPalet.visible = Modo = 1
    lblManual.visible = False
    
    'El tamaño SI importa
    lblLINEA.Font.SIZE = 86
    lblLINEA.Top = 120
    If Val(Me.lblLINEA.Caption) >= 8 Then
        lblManual.visible = True
        If Me.lblLINEA.Caption = "10" Then
            lblLINEA.Font.SIZE = 62
            lblLINEA.Top = 270
            lblManual.Caption = "MUESTRAS"
        Else
            lblManual.Caption = "MANAUAL"
        End If
    End If
    
    
    Me.Label2(15).visible = Modo > 0
    Text1(14).Locked = Modo > 0 'Solo se puede escribir creando
    Text1(15).Locked = Modo > 0 'Solo se puede escribir creando
    cmdFecCad.visible = Modo = 1 And vUsu.Nivel = 0
    
    
    'Junio 16
    'Sistemas de pesajes
    If Modo = 1 Then
        Me.Width = 16260
    Else
        Me.Width = 8625
    End If
    
    FramePesajes.visible = Modo = 1
    
    
    
    
    
    
End Sub


Private Sub PonerCampos()
Dim It As ListItem
Dim L As cLineaProCompo
Dim I As Long



    Text1(4).Text = cLP.codartic
    txtNomartic.Text = cLP.NomArtic
    Text1(0).Text = cLP.CodProduccion
    Text1(1).Text = cLP.idLiProd
    
    Text1(3).Text = Format(cLP.CantidadEstimada, FormatoCantidad)
    I = cLP.UnidadesCaja
    If I = 0 Then I = 1
    I = CLng(cLP.CantidadEstimada \ I)
    Text1(9).Text = I
    
    If Modo = 0 Then
        Text1(2).Text = ""  'lot trazabilidad
        Text1(5).Text = ""  'cant producida
        Text1(10).Text = ""
    Else
        Text1(2).Text = cLP.LoteTrazabilidad
        Text1(5).Text = ""  'sera la suma. la pongo cargando el tree
    End If
        
    Text1(12).Text = cLP.FechaCaducidad
    Text1(14).Text = cLP.LineaExtraEtiquetas
    Text1(15).Text = cLP.LineaExtraEtiqueta2
        
        
    Me.cboTipoImpresion.ListIndex = cLP.TipoImpresionPalet
        
    SQL = DevuelveDesdeBD(conAri, "pal_udbas * pal_udalt", "sarti4", "codartic", cLP.codartic, "T")
    Text1(11).Text = SQL
        
    Me.ListView1.ListItems.Clear
    Me.chkFin.Value = 0
    Me.chkFin.visible = False
    chkFindepositoEnCierreLinea.visible = Modo = 3
    chkFindepositoEnCierreLinea.Value = 0
    For I = 1 To cLP.CuantasMP
        If cLP.DevuelveComponenteLinea(CInt(I), L) Then
            Set It = ListView1.ListItems.Add()
            It.Text = L.codarticCompo
            It.SubItems(1) = L.NomArticCompo
            It.SubItems(2) = Format(L.CantidadEstimada, FormatoCantidad)
            It.SubItems(3) = L.LoteMateria
            If Modo > 0 Then It.SubItems(4) = L.LoteMateria
            
            It.SubItems(5) = Abs(L.EsMateriaPrima)
            
            If Modo = 2 Then
                'Esta modificando
                If I = SubLinea Then
                    'Esta es la linea que vamos a modificar
                    Text1(6).Text = L.codarticCompo
                    Text4.Text = L.NomArticCompo
                    Text1(7).Text = L.LoteMateria
                    Set ListView1.SelectedItem = It
                    If L.EsMateriaPrima Then
                        Me.chkFin.visible = True
                        chkFin.Value = 1
                    End If
                    
                End If
            End If
        End If
    Next
    
    
    'Junio 2016
    'Pesos para SGS
    If Modo = 0 Then
        LeerPesos
    End If
    
    
    'Cargamos los datos . Si modo =1 se mostrara el treee, si no, na de na.
    CargarHco

End Sub



Private Sub LeerPesos()
Dim RT As ADODB.Recordset
Dim SQL As String
Dim I As Integer
Dim J As Integer
Dim K As Integer
Dim cLinPro As cLineaProCompo
    
    On Error GoTo eLeerPesos
    
    SQL = "SELECT sarti1.codarti1,sartic.nomartic, sarti1.Cantidad,tipartic,desctipfamia From  sarti1 , sartic,stipfamia Where  "
    SQL = SQL & " sarti1.codarti1 = sartic.codartic And stipfamia.tipfamia = sartic.tipartic"
    SQL = SQL & " AND sarti1.codartic=" & DBSet(cLP.codartic, "T")
    'SQL = SQL & " AND tipartic in (2,3,4,5,8) ORDER BY sarti1.numlinea"
    SQL = SQL & " AND tipartic in (2,3,8) ORDER BY sarti1.numlinea"


    '2,3,4,5,8
    '  Envase,tapon,etiq fron,etiq dor,,retractil
    

    '000300010702  quiar
    ListView3.ListItems.Clear
    Set cLinPro = New cLineaProCompo
    Set RT = New ADODB.Recordset
    J = 0
    RT.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RT.EOF
        J = J + 1
        ListView3.ListItems.Add
        ListView3.ListItems(J).Text = RT!desctipfamia
        ListView3.ListItems(J).SubItems(1) = RT!NomArtic
        If RT!tipartic = 4 Or RT!tipartic = 5 Then
            'El peso viene de parametros
            ListView3.ListItems(J).SubItems(2) = " "
        Else
            If RT!tipartic = 8 Then
                'retractil, viene de la ficha
                SQL = DevuelveDesdeBD(conAri, "pesoneto", "sarti4", "codartic", RT!codarti1, "T")
                If SQL = "" Then
                    PonerElTextoPeso J, SQL, 1
                Else
                    PonerElTextoPeso J, SQL, 3
                End If
            Else
                'Emvase o tapon
                For K = 1 To ListView1.ListItems.Count
                    If ListView1.ListItems(K).Text = RT!codarti1 Then
                        'Veremos si tiene lote asignado
                        If ListView1.ListItems(K).SubItems(3) = "" Then
                            PonerElTextoPeso J, SQL, 0
                            
                            
                        Else
                            SQL = "numlote = " & DBSet(ListView1.ListItems(K).SubItems(3), "T") & " AND codartic"
                            SQL = DevuelveDesdeBD(conAri, "id", "spartidas", SQL, RT!codarti1, "T")
                            If SQL = "" Then
                                'Error GRAVE EN el lote. NO deberia pasar. PONGO un msgbox
                                MsgBox "Error GRAVE. Consulte soporte técnico", vbCritical
                            Else
                                'Veremos si esta pesado el lote
                                SQL = DevuelveDesdeBD(conAri, "avg(peso)", "spartidaspesos", "idpartida", SQL, "T")
                                If SQL = "" Then
                                    PonerElTextoPeso J, SQL, 1
                                Else
                                    PonerElTextoPeso J, SQL, 3
                                End If
                            End If
                        End If
                        Exit For
                    End If
                Next
            End If
        End If
        
        
        RT.MoveNext
    Wend
    RT.Close
    Set RT = Nothing




eLeerPesos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Leyendo pesos"
    Set RT = Nothing
    Set cLinPro = Nothing
End Sub

'EsOK:  0.- Falta Lote    1.- Falta peso     3.- OK
Private Sub PonerElTextoPeso(linea As Integer, Valor As String, EsOK As Byte)

    Select Case EsOK
    Case 0
        
        ListView3.ListItems(linea).SubItems(2) = "Falta lote"
        ListView3.ListItems(linea).ListSubItems(2).ForeColor = vbRed

    Case 1
        ListView3.ListItems(linea).SubItems(2) = "SIN PESO"
        ListView3.ListItems(linea).ListSubItems(2).ForeColor = vbRed
        ListView3.ListItems(linea).ListSubItems(2).Bold = True
    
    
    Case 2
    
    
    Case 3
        ListView3.ListItems(linea).SubItems(2) = Format(Valor, FormatoPrecio)
    End Select
End Sub




Private Sub Form_Unload(Cancel As Integer)
    If Not SePuedeSalir Then Cancel = 1
End Sub

Private Sub frmB2_Selecionado(CadenaDevuelta As String)
    SQL = CadenaDevuelta
End Sub

Private Sub frmC_Selec(vFecha As Date)
    SQL = vFecha
End Sub

Private Sub frmL_DatoSeleccionado(CadenaSeleccion As String)
    SQL = CadenaSeleccion
End Sub

Private Sub ListView1_DblClick()
    If Modo = 1 Then
        cmdVerLote_Click
    End If
End Sub

Private Sub ListView2_KeyPress(KeyAscii As Integer)
Dim Peso As Currency
Dim VolumenProd As Currency

    If ListView2.ListItems.Count = 0 Then Exit Sub
    If ListView2.SelectedItem Is Nothing Then Exit Sub
    
    If vUsu.Nivel > 1 Then Exit Sub
    
    If Not (KeyAscii = 112 Or KeyAscii = 80) Then Exit Sub
    
    
    SQL = DevuelveDesdeBD(conAri, "litrosunidad", "sartic", "codartic", Text1(4).Text, "T")
    If SQL = "" Then Err.Raise 513, , "Error LITROS - UNIDAD "
    VolumenProd = 1000 * CCur(SQL)
    
    Set miRsAux = New ADODB.Recordset
    SQL = "select  * from prodlinpesos where codigo =" & cLP.CodProduccion & " and idlin=" & cLP.idLiProd
    SQL = SQL & " AND serie =" & cboSerie.Text & " and secuencial = " & ListView2.SelectedItem.Text
    SQL = SQL & " AND lotetraza = " & cLP.LoteTrazabilidad
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If miRsAux.EOF Then Err.Raise 513, , "No existe dato (EOF)"
    
    SQL = InputBox("Peso en gr:", , CStr(miRsAux!pesolleno))
    If SQL <> "" Then
        
        Peso = CCur(SQL)
        SQL = "UPDATE prodlinpesos set pesolleno=" & DBSet(Peso, "N")
        
        
            'Volumen llenado
            Peso = Peso - (miRsAux!PesoBotella + miRsAux!pesotapon + miRsAux!pesoetiqueta + miRsAux!pesootro)
            Peso = Peso / 0.916
            SQL = SQL & ", volumenllenado=" & DBSet(Peso, "N")
            
            If VolumenProd - Peso > miRsAux!Emp Then
                SQL = SQL & ", cumpleemp=0"
                If VolumenProd - Peso > 2 * miRsAux!Emp Then
                   ' lw1.SelectedItem.ForeColor = vbRed
                    SQL = SQL & ", cumple2emp=0"
                Else
                   ' lw1.SelectedItem.ForeColor = vbGreen
                    SQL = SQL & ", cumple2emp=1"
                End If
            Else
                SQL = SQL & ", cumpleemp=1, cumple2emp=1"
            End If
            miRsAux.Close
            Set miRsAux = Nothing
            
            SQL = SQL & " Where Codigo = " & cLP.CodProduccion & " And idlin = " & cLP.idLiProd
            SQL = SQL & " AND serie =" & cboSerie.Text & " and secuencial = " & ListView2.SelectedItem.Text
            SQL = SQL & " AND lotetraza = " & cLP.LoteTrazabilidad
            conn.Execute SQL
            CargarPesos
    End If
     Set miRsAux = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    If Index = 13 Then ConseguirFoco Text1(13), 3
        
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Index = 13 Then PonerFocoBtn Me.cmdIniciarProduccio
    End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    If Index = 13 Then
        If Text1(13).Text <> "" Then
            If Not IsNumeric(Text1(13).Text) Then
                MsgBox "MESES debe ser numerico", vbExclamation
                Text1(13).Text = "18"
                PonerFoco Text1(13)
            End If
        End If
    Else
        If Index = 14 Or Index = 15 Then Text1(Index).Text = Replace(Text1(Index).Text, "|", "-")
        
    End If
End Sub

Private Sub Text2_GotFocus()
    ConseguirFoco Text2, 3
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then PonerFocoBtn Me.cmdAceptarCantidad
End Sub

Private Sub Text2_LostFocus()
    If Not PonerFormatoDecimal(Text2, 3) Then
        Text2.Text = ""
        Text3.Text = ""
    Else
        PonerDatosUdsCajas True
    End If
End Sub

Private Sub PonerDatosUdsCajas(DesdeUds As Boolean)
Dim L As Long
    If DesdeUds Then
        L = ImporteFormateado(Text2.Text) \ CInt(Me.TxtUD.Text)
        If (ImporteFormateado(Text2.Text) Mod CInt(Me.TxtUD.Text)) > 0 Then L = L + 1
        Text3.Text = Format(L, "0")
    Else
        L = Val(Text3.Text) * CInt(Me.TxtUD.Text)
        Text2.Text = Format(L, FormatoCantidad)
    End If
        
End Sub

Private Sub Text3_GotFocus()
    ConseguirFoco Text3, 3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then PonerFoco Text2
End Sub

Private Sub Text3_LostFocus()
    If Not PonerFormatoEntero(Text3) Then
        Text3.Text = ""
        Text2.Text = ""
    Else
        PonerDatosUdsCajas False
    End If
End Sub




Private Sub CargarHco()
Dim N As Node
Dim idTraza As Long
Dim Cantidad As Currency
Dim L As Byte
Dim C2 As Currency

    If Modo = 1 Then TreeView1.Nodes.Clear
    idTraza = -1
    SQL = "select prodtrazcompo.*,nomartic,cantprodu,depositoL,factorconversion from prodtrazlin,prodtrazcompo,sartic where"
    SQL = SQL & " prodtrazcompo.codigo = prodtrazlin.codigo and prodtrazcompo.idlin  = prodtrazlin.idlin  and"
    SQL = SQL & " prodtrazcompo.lineaprod   = prodtrazlin.lineaprod  and    prodtrazcompo.lotetraza = prodtrazlin.lotetraza and"
    SQL = SQL & " prodtrazcompo.codartic = sartic.codartic and prodtrazlin.codigo=" & cLP.CodProduccion & " and prodtrazlin.idlin= " & cLP.idLiProd
    SQL = SQL & " and prodtrazlin.lotetraza <>" & cLP.LoteTrazabilidad & "  order by lotetraza,factorconversion"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    idTraza = -1
    Cantidad = 0
    While Not miRsAux.EOF
        If idTraza <> miRsAux!lotetraza Then
            idTraza = miRsAux!lotetraza
            Cantidad = Cantidad + DBLet(miRsAux!cantprodu, "N")
            
            If Modo = 1 Then
                Set N = TreeView1.Nodes.Add(, , "C" & idTraza)
                N.Text = "LOTE " & Format(idTraza, "00000") & "  (" & Format(miRsAux!cantprodu, FormatoCantidad) & ")"
            End If
            
        End If
        If Modo = 1 Then
            Set N = TreeView1.Nodes.Add("C" & idTraza, tvwChild)
            
            SQL = miRsAux!codartic & " " & miRsAux!NomArtic
            L = Len(SQL)
            If L > 45 Then
                SQL = Mid(SQL, 1, 45)
                L = 1
            Else
                L = 46 - L
            End If
            
            SQL = SQL & Space(CLng(L))
            SQL = SQL & "Lot:" & miRsAux!NUmlote & " / "
            C2 = DBLet(miRsAux!cantutili, "N")
            If Int(C2) = C2 Then
                SQL = SQL & Format(C2, "#,##0")
            Else
                
                SQL = SQL & Format(C2, FormatoCantidad)
            End If
            
            If miRsAux!FactorConversion < 1 Then
                'DEPOSITO
                If DBLet(miRsAux!depositol, "N") > 0 Then SQL = SQL & "[" & miRsAux!depositol & "]"
            End If
            
            N.Text = SQL
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If Cantidad > 0 Then
        Text1(5).Text = Format(Cantidad, FormatoCantidad)
        idTraza = Me.cLP.UnidadesCaja
        If idTraza = 0 Then idTraza = 1
        idTraza = Cantidad \ idTraza
        Text1(10).Text = idTraza
        
    Else
        Text1(5).Text = ""
        Text1(10).Text = ""
    End If
    Set miRsAux = Nothing
End Sub


Private Sub LanzaImpresionEtiquetas2(RestoSql As String)
Dim C As String
    
    
    'Si hay que salir... se sale
    'Imprime el total de etiquetas

    'C = "{prodcajas.lotetraza} = " & Me.cLP.LoteTrazabilidad  b
    'Ponemos el txt pq si es cambio de lote, en la clase ya tenemos NUEVO lote de trazabilidad
    C = "{prodcajas.lotetraza} = " & Text1(2).Text & RestoSql
    LlamaImprimirGral C, "", 0, "EtiCaja.rpt", "Etiquetas de caja"
    
End Sub



Private Sub FijarValoresParaInsertProducion()


    SQL = DevuelveDesdeBD(conAri, "codigo", "prodcab", "feccreacion", Format(Now, FormatoFecha), "F")
    If SQL = "" Then
        'No hay nada hoy
        SQL = DevuelveDesdeBD(conAri, "max(codigo)", "prodcab", "1", "1")
        If SQL = "" Then SQL = "0"
        idProd = Val(SQL) + 1
        LinProd = 1
        
    Else
        idProd = Val(SQL)
        SQL = DevuelveDesdeBD(conAri, "max(idlin)", "prodLIN", "codigo", CStr(idProd))
        If SQL = "" Then SQL = "0"
        LinProd = Val(SQL) + 1
    End If
End Sub



Private Sub CargaPalets()
Dim I As Integer
Dim C As Integer
Dim It
    Set miRsAux = New ADODB.Recordset
    SQL = "select * from prodpalets where idpalet in (select distinct(idpalet)"
    SQL = SQL & " from prodcajas where lotetraza=" & cLP.LoteTrazabilidad & ") ORDER BY idpalet "
    lwPalet.ListItems.Clear
    I = 0
    C = 0
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    While Not miRsAux.EOF
            I = I + 1
            Set It = lwPalet.ListItems.Add()
            It.Text = Format(miRsAux!IdPalet, "0000")
            
            If Format(miRsAux!fhinicio, "dd/mm/yyyy") = SQL Then
                It.SubItems(1) = " "
            Else
                It.SubItems(1) = Format(miRsAux!fhinicio, "dd/mm/yyyy")
                SQL = It.SubItems(1)
            End If
            It.SubItems(2) = Format(miRsAux!Cajasprod, "#,##0")
            C = C + DBLet(miRsAux!Cajasprod, "N")
            It.SubItems(3) = " "
            miRsAux.MoveNext
    Wend
    miRsAux.Close
    If I > 0 Then
            Set It = lwPalet.ListItems.Add()
            It.Text = "TOTAL"
            It.SubItems(1) = I
            It.SubItems(2) = Format(C, "#,##0")
            It.Bold = True
            C = 0
            For I = 1 To lwPalet.ListItems.Count - 1
                SQL = "Select count(*) from prodcajas where lotetraza=" & cLP.LoteTrazabilidad
                SQL = SQL & " AND idpalet =" & lwPalet.ListItems(I).Text
                miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not miRsAux.EOF Then
                    If Not IsNull(miRsAux.Fields(0)) Then
                        lwPalet.ListItems(I).SubItems(3) = Format(miRsAux.Fields(0), "#,##0")
                        C = C + miRsAux.Fields(0)
                    End If
                End If
                miRsAux.Close
            Next
            If C > 0 Then lwPalet.ListItems(lwPalet.ListItems.Count).SubItems(3) = Format(C, "#,##0")
            
                
    End If
End Sub







Private Sub ObtenerLoteAceiteDeposito(articulonecesario As String)
Dim cad As String
Dim Depo As Integer
Dim B As Boolean
Dim I As Integer

        
        
        Screen.MousePointer = vbHourglass
        Set frmB2 = New frmBuscaGrid
        'CAMPOS
        'numdeposito,nomartic,spartidas.codartic,spartidas.numlote,litros
        cad = "Deposito|proddepositos|numdeposito|N||5·"
        cad = cad & "Cod. art|spartidas|codartic|T||20·"
        cad = cad & "Articulo|sartic|nomartic|T||45·"
        cad = cad & "Lote|spartidas|numlote|T||12·"
        'Si quiero litros lo pondria aqui
        cad = cad & "Litros||(kilos * factorconversion)|N|" & FormatoPrecio & "|16·"
        'Cad = Cad & "kilos||kilos|N|" & FormatoPrecio & "|16·"
        
        frmB2.vCampos = cad
        'TABLA
        cad = " proddepositos left join spartidas on spartidas.numlote=proddepositos.numlote"
        cad = cad & " inner join sartic on spartidas.codartic=sartic.codartic AND sartic.factorconversion<1"
        
        
        cad = cad & " and sartic.codartic = '" & articulonecesario & "'"
        '
        
        frmB2.vTabla = cad
        'WHERE
        frmB2.vSQL = "not spartidas.numlote is null"

        frmB2.vDevuelve = "0|3|4|"
        frmB2.vTitulo = "Depositos"
        frmB2.vselElem = 0
        frmB2.vConexionGrid = conAri 'Conexión a BD: Ariges
        SQL = ""
        frmB2.Show vbModal
        Set frmB2 = Nothing
        If SQL <> "" Then

            
            I = InStr(1, SQL, "|")
            NumDeposito = CByte(Mid(SQL, 1, I - 1))
            SQL = Mid(SQL, I + 1)
            
                        
            
        End If

End Sub


'***************************************************************************************
'***************************************************************************************
'***************************************************************************************
'
'   Sistema pesaje
'
'***************************************************************************************
'***************************************************************************************
'***************************************************************************************
Private Sub CargaSeries()
    cboSerie.Clear
    Set miRsAux = New ADODB.Recordset
    SQL = "select distinct serie from prodlinpesos where codigo =" & cLP.CodProduccion & " and idlin=" & cLP.idLiProd
    SQL = SQL & " AND lotetraza = " & cLP.LoteTrazabilidad & " order by 1"
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        cboSerie.AddItem miRsAux.Fields(0)
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    If cboSerie.ListCount > 0 Then cboSerie.ListIndex = cboSerie.ListCount - 1
    
End Sub

Private Sub CargarPesos()
Dim It As ListItem
Dim numero As Currency
Dim C As Integer

    Me.ListView2.ListItems.Clear
    For C = 0 To Text5.Count - 1
        Text5(C).Text = ""
    Next
    
    If Me.cboSerie.ListIndex < 0 Then Exit Sub
    
    Set miRsAux = New ADODB.Recordset
    SQL = "select  secuencial,pesoBotella,pesoTapon,pesoEtiqueta,pesoOtro,pesoLleno,volumenLlenado,EMP,CumpleEMP,Cumple2EMP "
    SQL = SQL & " from prodlinpesos where codigo =" & cLP.CodProduccion & " and idlin=" & cLP.idLiProd
    SQL = SQL & " AND lotetraza = " & cLP.LoteTrazabilidad
    SQL = SQL & " AND serie =" & cboSerie.Text & " order by secuencial"
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        
        
        'SQL = "K" & Format(miRsAux!secuencial, "00")
        Set It = ListView2.ListItems.Add()
        It.Text = Format(miRsAux!secuencial, "00")
        
        It.SubItems(1) = Format(miRsAux!pesolleno, FormatoPrecio)
       
        
        numero = 0.916 * miRsAux!volumenLlenado
        It.SubItems(2) = Format(numero, FormatoPrecio)
        It.SubItems(3) = Format(miRsAux!volumenLlenado, FormatoPrecio)
        It.SubItems(4) = Format(miRsAux!Emp, FormatoPorcen)
        
        If miRsAux!cumpleemp = 1 Then
            'OK
            SQL = " "
        Else
            SQL = "NO"
            It.ForeColor = vbRed
        End If
        
        It.SubItems(5) = SQL
        If miRsAux!cumple2emp = 1 Then
            'OK
            SQL = " "
        Else
            SQL = "NO"
            It.ForeColor = vbRed
            For C = 1 To It.ListSubItems.Count
                It.ListSubItems(C).ForeColor = vbRed
            Next
            
        End If
        It.SubItems(6) = SQL
        If SQL = "NO" Then It.ListSubItems(6).ForeColor = vbRed
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    'Resumen de la muestra
    ValoresEstadisticosMuestra2
    
    
    
    Set miRsAux = Nothing
            
    
End Sub



'De cada SERIE
Private Sub ValoresEstadisticosMuestra2()
Dim I As Integer
Dim V As Currency
Dim MedidaCorrecta As Boolean
Dim LitrosUnidadMiles As Integer


    For I = 0 To 12
        Text5(I).Text = ""
    Next
    
    
    SQL = "select avg(round(volumenLlenado,2)) media,count(*) cuantos,std(volumenLlenado) desviacion,"
    SQL = SQL & " sum(if(cumpleemp=0,1,0)) NoEmp1,sum(if(Cumple2EMP=0,1,0)) NoEmp2, max(volumenLlenado) maximo,min(volumenLlenado) minimo"
    SQL = SQL & " ,pesobotella,pesotapon,pesoEtiqueta,pesootro"
    SQL = SQL & " from prodlinpesos where codigo =" & cLP.CodProduccion & " and idlin=" & cLP.idLiProd
    SQL = SQL & " AND lotetraza = " & cLP.LoteTrazabilidad
    SQL = SQL & " AND serie =" & cboSerie.Text
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    If Not miRsAux.EOF Then I = DBLet(miRsAux!Cuantos, "N")
    MedidaCorrecta = False
    If I > 0 Then
        'Hay datos. Vamos a mostrarlos
        MedidaCorrecta = True
        Text5(1).Text = Format(miRsAux!PesoBotella, FormatoPrecio)
        Text5(2).Text = Format(miRsAux!pesotapon, FormatoPrecio)
        Text5(3).Text = Format(miRsAux!pesoetiqueta, FormatoPrecio)
        Text5(4).Text = Format(miRsAux!pesootro, FormatoPrecio)
        
        Text5(5).Text = Format(miRsAux!Cuantos, "0000")
        Text5(6).Text = Format(miRsAux!Maximo, FormatoPrecio)
        Text5(7).Text = Format(miRsAux!Minimo, FormatoPrecio)
        
        Text5(8).Text = Format(miRsAux!noemp1, "0000")
        Text5(9).Text = Format(miRsAux!noemp2, "0000")
        
        Text5(10).Text = Format(miRsAux!media, FormatoPrecio)
        Text5(11).Text = Format(miRsAux!Desviacion, FormatoPrecio)
        
        
        'OK emp1
        If miRsAux!noemp1 > 2 Then
            If miRsAux!Cuantos > 50 Then
                'Segunda tanda pesadas
                If miRsAux!noemp1 >= 7 Then
                    MedidaCorrecta = False
                    PonerColorText 1, 8
                Else
                    'Caution. Ha sido correcto
                    PonerColorText 2, 8
                End If
            Else
                If miRsAux!noemp1 >= 5 Then
                    MedidaCorrecta = False
                    PonerColorText 1, 8
                Else
                    'Caution. Ha sido correcto
                    PonerColorText 2, 8
                End If
            End If
        Else
            PonerColorText 0, 8
        End If
        
        If miRsAux!noemp2 > 0 Then
            MedidaCorrecta = False
            PonerColorText 1, 9
        Else
            PonerColorText 0, 9
        End If
        SQL = DevuelveDesdeBD(conAri, "litrosunidad", "sartic", "codartic", Text1(4).Text, "T")
        If SQL = "" Then MsgBox "Error LITROS - UNIDAD ", vbCritical
        LitrosUnidadMiles = 1000 * CCur(SQL)
        
        'B13>=B2-0,379*B14
        V = miRsAux!Desviacion * 0.397
        V = LitrosUnidadMiles - V
        
        
        If miRsAux!media > V Then
            'ok
        Else
            MedidaCorrecta = False
        End If
        
        If MedidaCorrecta Then
            'COnforme
            Text5(12).Text = "CONFORME"
            PonerColorText 0, 12
        Else
            'NO conforme
            Text5(12).Text = "NO"
            PonerColorText 1, 12
        End If
        
    End If
    miRsAux.Close
    
    
    'Nominal.
    'En sartic
    SQL = "Select LitrosUnidad from sartic where codartic=" & DBSet(cLP.codartic, "T")
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Text5(0).Text = ""
    If Not miRsAux.EOF Then
        Text5(0).Text = DBLet(miRsAux!LitrosUnidad, "N") * 1000
    End If
    miRsAux.Close
    
    
    
End Sub

'Ok:  0: OK     1: Mal      2: Cuation
Private Sub PonerColorText(OK As Byte, Indice As Integer)
Dim ColoresFondo As String
Dim ColoresFore As String

    ColoresFondo = "&H4000|&H80|&HFFFF|"
    ColoresFore = "&HC0FFC0|&HFFFFFF|&HFF0000|"
    If OK = 0 Then
        Text5(Indice).BackColor = RecuperaValor(ColoresFondo, OK + 1)
        Text5(Indice).ForeColor = RecuperaValor(ColoresFore, OK + 1)
        Text5(Indice).FontBold = False
    Else
        Text5(Indice).BackColor = RecuperaValor(ColoresFondo, OK + 1)
        Text5(Indice).ForeColor = RecuperaValor(ColoresFore, OK + 1)
        Text5(Indice).FontBold = True
        Text5(Indice).ForeColor = vbWhite
    End If
End Sub
