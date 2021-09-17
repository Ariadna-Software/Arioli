VERSION 5.00
Begin VB.Form frmProduVarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multi form para muchas cosas de produccion"
   ClientHeight    =   9360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   12585
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameVaciado 
      Height          =   2175
      Left            =   2880
      TabIndex        =   26
      Top             =   1200
      Width           =   7695
      Begin VB.Frame FrameFechaVall 
         Caption         =   "Frame1"
         Height          =   615
         Left            =   240
         TabIndex        =   89
         Top             =   1320
         Width           =   4815
         Begin VB.TextBox txtFecha 
            Height          =   285
            Index           =   5
            Left            =   960
            TabIndex        =   28
            Text            =   "Text1"
            Top             =   120
            Width           =   1215
         End
         Begin VB.TextBox txtHora 
            Height          =   285
            Index           =   5
            Left            =   2280
            TabIndex        =   29
            Text            =   "Text1"
            Top             =   135
            Width           =   1215
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   5
            Left            =   720
            Picture         =   "frmProduVarios.frx":0000
            Top             =   135
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   90
            Top             =   135
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdVaciadoDeposito 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   5160
         TabIndex        =   30
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   3
         Left            =   6360
         TabIndex        =   31
         Top             =   1440
         Width           =   975
      End
      Begin VB.ComboBox cboDeposito 
         Height          =   315
         Index           =   2
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   960
         Width           =   7095
      End
      Begin VB.Label lbFec 
         AutoSize        =   -1  'True
         Caption         =   "Forzar vaciado depósito"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   6
         Left            =   1920
         TabIndex        =   33
         Top             =   240
         Width           =   3420
      End
      Begin VB.Label lbFec 
         AutoSize        =   -1  'True
         Caption         =   "Depósito"
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
         Index           =   5
         Left            =   240
         TabIndex        =   32
         Top             =   720
         Width           =   750
      End
   End
   Begin VB.Frame FrameLaVallTraFiltrado 
      Height          =   6255
      Left            =   4440
      TabIndex        =   54
      Top             =   3600
      Visible         =   0   'False
      Width           =   8535
      Begin VB.CommandButton cmdLavall 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   6120
         TabIndex        =   84
         Top             =   5640
         Width           =   975
      End
      Begin VB.OptionButton optLaVall 
         Caption         =   "Trasiego"
         Height          =   195
         Index           =   1
         Left            =   5880
         TabIndex        =   73
         Top             =   1245
         Width           =   1095
      End
      Begin VB.OptionButton optLaVall 
         Caption         =   "Filtrado"
         Height          =   195
         Index           =   0
         Left            =   4200
         TabIndex        =   72
         Top             =   1245
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox txtNumeroDec 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   6720
         TabIndex        =   76
         Text            =   "Text1"
         Top             =   2760
         Width           =   1455
      End
      Begin VB.ComboBox cboDeposito 
         Height          =   315
         Index           =   6
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   75
         Top             =   2760
         Width           =   6255
      End
      Begin VB.ComboBox cboDeposito 
         Height          =   315
         Index           =   5
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   74
         Top             =   1920
         Width           =   6255
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   4
         Left            =   1200
         TabIndex        =   70
         Text            =   "Text1"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtHora 
         Height          =   285
         Index           =   4
         Left            =   2520
         TabIndex        =   71
         Text            =   "Text1"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Height          =   495
         Left            =   7560
         Picture         =   "frmProduVarios.frx":008B
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   6
         Left            =   7320
         TabIndex        =   86
         Top             =   5640
         Width           =   975
      End
      Begin VB.Frame FrameFilltroLaVall2 
         Height          =   2175
         Left            =   240
         TabIndex        =   55
         Top             =   3240
         Width           =   8055
         Begin VB.TextBox txtNumeroDec 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   6720
            TabIndex        =   79
            Text            =   "Text1"
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtNumeroDec 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   6720
            TabIndex        =   81
            Text            =   "Text1"
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox txtNumeroDec 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   3
            Left            =   6720
            TabIndex        =   83
            Text            =   "Text1"
            Top             =   1560
            Width           =   975
         End
         Begin VB.TextBox txtArtFiltrado 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   4
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   66
            Text            =   "Text1"
            Top             =   1080
            Width           =   3255
         End
         Begin VB.TextBox txtArtFiltrado 
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   77
            Text            =   "Text1"
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox txtArtFiltrado 
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   65
            Text            =   "Text1"
            Top             =   1080
            Width           =   1575
         End
         Begin VB.TextBox txtArtFiltrado 
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   64
            Text            =   "Text1"
            Top             =   1560
            Width           =   1575
         End
         Begin VB.TextBox txtArtFiltrado 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   3
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   63
            Text            =   "Text1"
            Top             =   600
            Width           =   3255
         End
         Begin VB.TextBox txtArtFiltrado 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   5
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   62
            Text            =   "Text1"
            Top             =   1560
            Width           =   3255
         End
         Begin VB.TextBox txtLote 
            Height          =   285
            Index           =   0
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   61
            Text            =   "Text1"
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtLote 
            Height          =   285
            Index           =   1
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   60
            Text            =   "Text1"
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtLote 
            Height          =   285
            Index           =   2
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   59
            Text            =   "Text1"
            Top             =   1560
            Width           =   1215
         End
         Begin VB.CommandButton cmdLote 
            Caption         =   "+"
            Height          =   255
            Index           =   0
            Left            =   5160
            TabIndex        =   58
            Top             =   600
            Width           =   135
         End
         Begin VB.CommandButton cmdLote 
            Caption         =   "+"
            Height          =   255
            Index           =   1
            Left            =   5160
            TabIndex        =   57
            Top             =   1080
            Width           =   135
         End
         Begin VB.CommandButton cmdLote 
            Caption         =   "+"
            Height          =   255
            Index           =   2
            Left            =   5160
            TabIndex        =   56
            Top             =   1560
            Width           =   135
         End
         Begin VB.Label Label2 
            Caption         =   "Producto"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   69
            Top             =   240
            Width           =   1125
         End
         Begin VB.Label Label2 
            Caption         =   "Descripcion"
            Height          =   195
            Index           =   1
            Left            =   2520
            TabIndex        =   68
            Top             =   240
            Width           =   1485
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Kilos"
            Height          =   195
            Index           =   2
            Left            =   7080
            TabIndex        =   67
            Top             =   240
            Width           =   330
         End
      End
      Begin VB.Label lbFec 
         AutoSize        =   -1  'True
         Caption         =   "Kilos"
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
         Index           =   19
         Left            =   6720
         TabIndex        =   88
         Top             =   2520
         Width           =   390
      End
      Begin VB.Label lbFec 
         AutoSize        =   -1  'True
         Caption         =   "Destino"
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
         Index           =   18
         Left            =   240
         TabIndex        =   87
         Top             =   2520
         Width           =   645
      End
      Begin VB.Label lbFec 
         AutoSize        =   -1  'True
         Caption         =   "Origen"
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
         Index           =   17
         Left            =   240
         TabIndex        =   85
         Top             =   1680
         Width           =   555
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   4
         Left            =   840
         Picture         =   "frmProduVarios.frx":0A8D
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label lbFec 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Index           =   16
         Left            =   240
         TabIndex        =   82
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lbFec 
         AutoSize        =   -1  'True
         Caption         =   "Movimientos depósitos  cooperativa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   15
         Left            =   240
         TabIndex        =   80
         Top             =   240
         Width           =   5190
      End
   End
   Begin VB.Frame FrCierreOrdenProduccion 
      Height          =   2175
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5295
      Begin VB.TextBox txtHora 
         Height          =   285
         Index           =   0
         Left            =   2520
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtMeses 
         Height          =   285
         Left            =   1920
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton cmdCierreOrdProd 
         Caption         =   "Cerrar orden"
         Height          =   375
         Left            =   2880
         TabIndex        =   3
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   4080
         TabIndex        =   4
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   705
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Meses caducidad"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lbFec 
         AutoSize        =   -1  'True
         Caption         =   "Cierre orden de producción"
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
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   2280
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   960
         Picture         =   "frmProduVarios.frx":0B18
         Top             =   720
         Width           =   240
      End
   End
   Begin VB.Frame FrameTrasiego 
      Height          =   3015
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   8775
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   3
         Left            =   1920
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtHora 
         Height          =   285
         Index           =   3
         Left            =   3240
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   840
         Width           =   1215
      End
      Begin VB.Frame FrameTrasiegoLaVAll 
         Height          =   735
         Left            =   240
         TabIndex        =   45
         Top             =   1920
         Width           =   3975
         Begin VB.TextBox txtNumeroDec 
            Alignment       =   1  'Right Justify
            Height          =   375
            Index           =   0
            Left            =   840
            TabIndex        =   20
            Text            =   "Text1"
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lbFec 
            AutoSize        =   -1  'True
            Caption         =   "Kilos"
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
            Index           =   12
            Left            =   240
            TabIndex        =   46
            Top             =   360
            Width           =   390
         End
      End
      Begin VB.ComboBox cboDeposito 
         Height          =   315
         Index           =   1
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1560
         Width           =   1695
      End
      Begin VB.ComboBox cboDeposito 
         Height          =   315
         Index           =   0
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1560
         Width           =   6495
      End
      Begin VB.CommandButton cmdtrasiego 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   6480
         TabIndex        =   21
         Top             =   2280
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   2
         Left            =   7560
         TabIndex        =   22
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label lbFec 
         AutoSize        =   -1  'True
         Caption         =   "Fecha/hora"
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
         Index           =   13
         Left            =   240
         TabIndex        =   47
         Top             =   840
         Width           =   975
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   1560
         Picture         =   "frmProduVarios.frx":0BA3
         Top             =   840
         Width           =   240
      End
      Begin VB.Label lbFec 
         AutoSize        =   -1  'True
         Caption         =   "Destino"
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
         Index           =   4
         Left            =   6840
         TabIndex        =   25
         Top             =   1320
         Width           =   645
      End
      Begin VB.Label lbFec 
         AutoSize        =   -1  'True
         Caption         =   "Origen"
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
         Index           =   3
         Left            =   240
         TabIndex        =   24
         Top             =   1320
         Width           =   555
      End
      Begin VB.Label lbFec 
         AutoSize        =   -1  'True
         Caption         =   "Trasiego"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   2
         Left            =   3480
         TabIndex        =   23
         Top             =   240
         Width           =   1230
      End
   End
   Begin VB.Frame FrCoupage 
      Height          =   1695
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6495
      Begin VB.TextBox txtHora 
         Height          =   285
         Index           =   1
         Left            =   2160
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   1
         Left            =   840
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   5160
         TabIndex        =   11
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdCoupage 
         Caption         =   "Hacer"
         Height          =   375
         Left            =   3840
         TabIndex        =   10
         Top             =   720
         Width           =   1095
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   480
         Picture         =   "frmProduVarios.frx":0C2E
         Top             =   720
         Width           =   240
      End
      Begin VB.Label lbFec 
         AutoSize        =   -1  'True
         Caption         =   "Hacer coupage"
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
         Index           =   1
         Left            =   360
         TabIndex        =   12
         Top             =   240
         Width           =   4620
      End
   End
   Begin VB.Frame FrameFiltrado 
      Height          =   3735
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Visible         =   0   'False
      Width           =   8775
      Begin VB.CommandButton cmdImpreFiltrado 
         Height          =   495
         Left            =   7920
         Picture         =   "frmProduVarios.frx":0CB9
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtHora 
         Height          =   285
         Index           =   2
         Left            =   2520
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Frame FramefiltroMorales 
         Height          =   735
         Left            =   240
         TabIndex        =   48
         Top             =   2640
         Width           =   4455
         Begin VB.CheckBox chkFiltrado 
            Caption         =   "Depósito 8"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   50
            Top             =   360
            Width           =   1815
         End
         Begin VB.CheckBox chkFiltrado 
            Caption         =   "Depósito 9"
            Height          =   255
            Index           =   1
            Left            =   2520
            TabIndex        =   49
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label lbFec 
            AutoSize        =   -1  'True
            Caption         =   "Depósitos auxiliares"
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
            Index           =   10
            Left            =   0
            TabIndex        =   51
            Top             =   0
            Width           =   1710
         End
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   2
         Left            =   1200
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   4
         Left            =   7560
         TabIndex        =   40
         Top             =   2880
         Width           =   975
      End
      Begin VB.CommandButton cmdFiltrar 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   6480
         TabIndex        =   39
         Top             =   2880
         Width           =   975
      End
      Begin VB.ComboBox cboDeposito 
         Height          =   315
         Index           =   4
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   1800
         Width           =   6495
      End
      Begin VB.ComboBox cboDeposito 
         Height          =   315
         Index           =   3
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label lbFec 
         AutoSize        =   -1  'True
         Caption         =   "Impresión parte filtrado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   14
         Left            =   5640
         TabIndex        =   52
         Top             =   600
         Width           =   2040
      End
      Begin VB.Label lbFec 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Index           =   11
         Left            =   240
         TabIndex        =   44
         Top             =   1080
         Width           =   495
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   840
         Picture         =   "frmProduVarios.frx":16BB
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label lbFec 
         AutoSize        =   -1  'True
         Caption         =   "Proceso filtrado aceite"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   9
         Left            =   360
         TabIndex        =   43
         Top             =   360
         Width           =   4170
      End
      Begin VB.Label lbFec 
         AutoSize        =   -1  'True
         Caption         =   "Origen"
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
         Index           =   8
         Left            =   240
         TabIndex        =   42
         Top             =   1560
         Width           =   555
      End
      Begin VB.Label lbFec 
         AutoSize        =   -1  'True
         Caption         =   "Destino"
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
         Left            =   6840
         TabIndex        =   41
         Top             =   1560
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmProduVarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte
    '0  .-Cierrer de una orden de produccion
    '1  .-Hacer coupage
        
    '2  .- trasiego
    '3  .- Vaciado
    '4  .- Filtrado
    
    '5  .- Hacer el coupage autmatico. Lo llama desde proceso almazara
    
    
    '6.- Trasiego/Filtrado desde LAVALLL
    
    
Public Intercambio As String
    '0 : codiog|fecha creacion
    '1:  codigo|fecha|almacen
    
    
'Para evitar hacer una select cad vez que lle alguna linea para el stock
Private TrabajadorConectado_ As Integer
    
    
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmL As frmAlmPartidas
Attribute frmL.VB_VarHelpID = -1

Dim cad As String  'multi proposito
Dim I As Integer

Dim PrimeraVez As Boolean

Private Sub cboDeposito_Click(Index As Integer)
    If vParamAplic.QUE_EMPRESA = 4 Then
        'En el camopo kilos pongo la cantidad total
        If Index = 0 Then
            I = InStr(1, cboDeposito(Index).Text, "(")
            If I > 0 Then
                cad = Mid(cboDeposito(Index).Text, I + 1)
                I = InStr(1, cad, ")")
                If I > 0 Then Me.txtNumeroDec(0).Text = Mid(cad, 1, I - 1)
            End If
        ElseIf Index = 5 Then
            I = InStr(1, cboDeposito(Index).Text, "(")
            If I > 0 Then
                cad = Mid(cboDeposito(Index).Text, I + 1)
                I = InStr(1, cad, ")")
                If I > 0 Then Me.txtNumeroDec(4).Text = Mid(cad, 1, I - 1)
            End If
        End If
    End If
End Sub

Private Sub cboDeposito_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub chkFiltrado_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdCierreOrdProd_Click()
    If txtFecha(0).Text = "" Then Exit Sub
    If txtMeses.Text = "" Then
        MsgBox "Indique los meses para la fecha de caducidad", vbExclamation
        PonerFoco txtMeses
        Exit Sub
    End If
    
    
    
    'Fecha activa.
    'Puesta por  para la VALL. Al resto sera 01/01/1900
    If CDate(txtFecha(0).Text) < vParamAplic.FechaActiva Then
        MsgBox "Periodo de produccion cerrado", vbExclamation
        Exit Sub
    End If
    
    
    
    
    
    If vParamAplic.QUE_EMPRESA = 4 Then
        If CDate(txtFecha(0).Text) > vParamAplic.FechaActivaMasUno Then
            MsgBox "Periodo de produccion no abierto", vbExclamation
            Exit Sub
        End If
    
        'Tienen que indicar el campo HORA
        If txtHora(0).Text = "" Then
            MsgBox "Indique la hora del cierre de produccion", vbExclamation
            PonerFoco txtHora(0)
            Exit Sub
        End If
    End If
    
    
    cad = RecuperaValor(Intercambio, 2)
    If CDate(cad) > CDate(txtFecha(0).Text) Then
        cad = "Fecha anterior a la creacion del parte de produccion." & vbCrLf & vbCrLf & "Creacion: " & cad
        cad = String(60, "*") & vbCrLf & cad & vbCrLf & vbCrLf & String(60, "*") & vbCrLf
        If vParamAplic.QUE_EMPRESA = 4 Then
            MsgBox cad, vbExclamation
            Exit Sub
        Else
            cad = cad & vbCrLf & "Cierre: " & txtFecha(0).Text
            cad = cad & vbCrLf & "Caducidad. Meses: " & txtMeses.Text & "    "
            cad = cad & "EXP: " & Format(DateAdd("m", Val(txtMeses.Text), CDate(txtFecha(0).Text)), "mm/yyyy") & vbCrLf & vbCrLf & "¿Continuar?"
            If MsgBox(cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
    Else
        cad = "Va a cerrar la orden de producción " & RecuperaValor(Intercambio, 1) & " - " & RecuperaValor(Intercambio, 2)
        cad = cad & vbCrLf & " Fecha prod. : " & txtFecha(0).Text
        If vParamAplic.QUE_EMPRESA = 4 Then cad = cad & "   Hora: " & txtHora(0).Text
        cad = cad & vbCrLf & vbCrLf & "¿Continuar?"
        If MsgBox(cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    If CerrarOrdenProduccion(True) Then
        If CerrarOrdenProduccion(False) Then Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdCoupage_Click()



    If txtFecha(1).Text = "" Then Exit Sub
    
    If vParamAplic.QUE_EMPRESA = 4 Then
        If CDate(txtFecha(1).Text) < vParamAplic.FechaActiva Then
            MsgBox "Periodo de produccion cerrado", vbExclamation
            Exit Sub
        End If
        
        If CDate(txtFecha(1).Text) > vParamAplic.FechaActivaMasUno Then
            MsgBox "Periodo de produccion sin abrir", vbExclamation
            Exit Sub
        End If
        
        If Me.txtHora(1).Text = "" Then
            MsgBox "Indique hora del proceso", vbExclamation
            PonerFoco txtHora(1)
            Exit Sub
        End If
    End If
    
    If Opcion = 5 Then
        'No hacemos pregunta a que lanzamos autmaticamente
        '-----
    Else
        cad = "¿Seguro que desea hacer el coupage " & RecuperaValor(Intercambio, 1) & " - " & RecuperaValor(Intercambio, 2)
        If MsgBox(cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    
    
     
    If vParamAplic.QUE_EMPRESA = 4 Then
        
        If RealizarCoupageVALL(True) Then
            If RealizarCoupageVALL(False) Then
                'Si ha ido bien, y el articulo es UNO de los que se tiene que actualizar el upc
                ActualizarPrecio
                '---------
                Unload Me
            End If
        End If
        
        
    Else
        If RealizarCoupage(True) Then
            If RealizarCoupage(False) Then
                'Si ha ido bien, y el articulo es UNO de los que se tiene que actualizar el upc
                ActualizarPrecio
                '---------
                Unload Me
            End If
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub cmdFiltrar_Click()
Dim C1 As cDeposito
Dim C2 As cDeposito
Dim CC As CTiposMov
Dim FechaHora2 As Date

    cad = ""
    If txtFecha(2).Text = "" Then cad = "-Fecha"
    If vParamAplic.QUE_EMPRESA = 4 Then If Me.txtHora(2).Text = "" Then cad = "   -Hora"
    If cboDeposito(3).ListIndex < 0 Or cboDeposito(4).ListIndex < 0 Then cad = cad & "  -Deposito"
    For I = 0 To 2
        
        If Me.txtArtFiltrado(I).Text <> "" And Me.txtNumeroDec(I + 1).Text <> "" And Me.txtLote(I).Text = "" Then cad = cad & vbCrLf & "  -Lote para " & txtArtFiltrado(I).Text
    Next
    If cad <> "" Then
        cad = "Campos requeridos: " & vbCrLf & cad
        MsgBox cad, vbExclamation
        Exit Sub
    End If
        
        
        
        
    If CDate(txtFecha(2).Text) < vParamAplic.FechaActiva Then
        MsgBox "Menor que fecha activa", vbExclamation
        Exit Sub
    End If
    
    For I = 0 To 1
        cad = ""
        If Me.chkFiltrado(I).Value = 1 Then
            'El deposito 8 no puede ser destino ni estar lleno
            NumRegElim = cboDeposito(3).ItemData(cboDeposito(3).ListIndex)
            If NumRegElim = 8 + I Then cad = "Deposito " & NumRegElim & " no puede ser destino ya que se utiliza como intermedio"
        End If
        
        If cad = "" Then
            If Me.chkFiltrado(I).Value = 1 Then
                Set C1 = New cDeposito
                If C1.LeerDatos(8 + I, False) Then
                    If C1.numLote <> "" Then cad = "Deposito intermedio  no esta vacio"
                End If
                Set C1 = Nothing
            End If
        End If
        If cad <> "" Then
            MsgBox cad, vbExclamation
            Exit Sub
        End If
    Next
    
    TrabajadorConectado_ = Val(PonerTrabajadorConectado(vUsu.Login))
    cad = "Va a realizar el filtrado: " & vbCrLf & "Origen: " & cboDeposito(4).Text
    cad = cad & vbCrLf & "Destino: " & cboDeposito(3).Text & vbCrLf & vbCrLf
    
    'Si hay gasto de productos en filtrado
    For I = 1 To 3
        If Me.txtNumeroDec(I).Text <> "" Then cad = cad & "      - " & Me.txtArtFiltrado(2 + I).Text & ": " & txtNumeroDec(I).Text & "  Kilos" & vbCrLf
    Next I

    If vParamAplic.QUE_EMPRESA = 1 Then
        If Me.chkFiltrado(0).Value = 1 Then cad = cad & vbCrLf & "Deposito auxiliar 8"
        If Me.chkFiltrado(1).Value = 1 Then cad = cad & vbCrLf & "Deposito auxiliar 9"
    End If
    If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    
    'Fecha hora la indica
    FechaHora2 = CDate(txtFecha(2).Text & " " & Me.txtHora(2).Text)
    
    If vParamAplic.QUE_EMPRESA = 4 Then
        Me.chkFiltrado(0).Value = 0
        Me.chkFiltrado(1).Value = 0
    End If
    
            'ANTES para morales. Para obtener la hora
''                cad = "select horamovi from proddepositoshco  where horamovi>=" & DBSet(txtFecha(2).Text, "F")
''                'menor que el dia siguiente
''                cad = cad & " AND horamovi<" & DBSet(DateAdd("d", 1, CDate(txtFecha(2).Text)), "F")
''                cad = cad & " AND tipoaccion=8 order by horamovi desc"
''                Set miRsAux = New ADODB.Recordset
''                miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
''                FechaHora = CDate(txtFecha(2).Text & " " & "07:00:00")
''                If Not miRsAux.EOF Then
''                    If Not IsNull(miRsAux!horamovi) Then
''                        FechaHora = miRsAux!horamovi
''                        FechaHora = DateAdd("s", 1, FechaHora)
''                    End If
''                End If
''                miRsAux.Close
''                Set miRsAux = Nothing
''            End If
    
    'Hacemos el trasiego.
    cad = ""
    Set CC = New CTiposMov
    If CC.ConseguirContador("TRO") Then
        
        Set C1 = New cDeposito
        Set C2 = New cDeposito
        
        
        
        If C1.LeerDatos(cboDeposito(4).ItemData(cboDeposito(4).ListIndex), False) Then
            If C2.LeerDatos(cboDeposito(3).ItemData(cboDeposito(3).ListIndex), False) Then
                C1.HacerFiltrado C2, Me.chkFiltrado(0).Value = 1, Me.chkFiltrado(1).Value = 1, CC.contador + 1, FechaHora2, 0
                
                CC.IncrementarContador CC.TipoMovimiento
                
                
                'sfiltradoaceite(idFiltrado,Trabajador,FechaHora,DepositoInicial,DepositoFinal,Kilos,usaAux8,usaAux9,idPartida)
                cad = "(" & CC.contador & "," & TrabajadorConectado_ & "," & DBSet(FechaHora2, "FH") & "," & C1.NumDeposito & ","
                cad = cad & C2.NumDeposito & "," & DBSet(C2.Kilos, "N") & "," & Abs(Me.chkFiltrado(0).Value)
                cad = cad & "," & Abs(Me.chkFiltrado(0).Value) & "," & C2.idPartida & ")"
                cad = "insert into sfiltradoaceite(idFiltrado,Trabajador,FechaHora,DepositoInicial,DepositoFinal,Kilos,usaAux8,usaAux9,idPartida) values " & cad
                
                If Not EjecutaSQL(conAri, cad, False) Then MsgBox "El programa continuará. Llame a soporte tecnico" & vbCrLf & cad, vbExclamation
                    
                
                
                
                
                cad = "OK"
            End If
        End If
    
            
        'productos filtrado
        If cad = "OK" Then
            '                           el mas uno ya esta hecho
            HacerStockProductosFiltrados CC.contador, FechaHora2
            
        End If
        
    End If
    Set CC = Nothing
    
    Set C1 = Nothing
    Set C2 = Nothing


    If cad <> "" Then Unload Me
End Sub



Private Sub HacerStockProductosFiltrados(idFil As Long, Fecha As Date)
Dim vCStock As cStock
Dim cPar As cPartidas
Dim cLot As cLotaje
Dim Aux As String
Dim Cantidad As Currency

    
    
    Set cLot = New cLotaje
   Set vCStock = New cStock
    vCStock.tipoMov = "S"
    vCStock.DetaMov = "TRO"
    vCStock.Trabajador = TrabajadorConectado_
    vCStock.Documento = Format(idFil, "00000")
    vCStock.Fechamov = Format(Fecha, "dd/mm/yyyy")
    vCStock.HoraMov = Fecha
    vCStock.codAlmac = 1
    
    
    cLot.codAlmac = vCStock.codAlmac
    cLot.DetaMov = vCStock.DetaMov
    cLot.Fechamov = vCStock.Fechamov
    cLot.HoraMov = vCStock.HoraMov
    cLot.Documento = vCStock.Documento
    cLot.ProvCliTra = vCStock.Trabajador
    
    
    For I = 1 To 3
        If Me.txtNumeroDec(I).Text <> "" Then
            'OK este lleva
            vCStock.LineaDocu = I
            vCStock.Cantidad = ImporteFormateado(txtNumeroDec(I).Text)
            
            cad = DevuelveDesdeBD(conAri, "concat(sartic.codartic,'|',coalesce(preciouc,0),'|')", "vallparam, sartic", IIf(I = 1, "diatomeasRojas", IIf(I = 2, "diatomeasVerdas", "celulosa")) & " = sartic.codartic AND 1", "1")
            If cad = "" Then
                MsgBox "Error obteniendo articulo filtrado:" & IIf(I = 1, "diatomeasRojas", IIf(I = 2, "diatomeasVerdas", "celulosa")), vbExclamation
            Else
                
                vCStock.codArtic = RecuperaValor(cad, 1)
                cad = RecuperaValor(cad, 2)
                vCStock.Importe = TransformaPuntosComas(cad)
                vCStock.Importe = vCStock.Importe * vCStock.Cantidad
                vCStock.ActualizarStock False
                
                cLot.codArtic = vCStock.codArtic
                
                Aux = "spartidas.codartic=sartic.codartic"
                Aux = Aux & " AND spartidas.codartic= " & DBSet(vCStock.codArtic, "T") & " AND numlote"
                cad = DevuelveDesdeBD(conAri, "id", "spartidas,sartic", Aux, Me.txtLote(I - 1).Text, "T")
                If cad = "" Then
                    MsgBox "No se encuentra la partida" & Aux, vbExclamation
                Else
                    Set cPar = New cPartidas
                    cPar.Leer CLng(cad)
                                    
                                    
                    cLot.numLote = cPar.numLote
        
                    cLot.Cantidad = vCStock.Cantidad
                    cLot.LineaDocu = I
                   
                    
                    cLot.InsertarLote
                
                    cPar.IncrementarCantidad -vCStock.Cantidad
                    Set cPar = Nothing
                End If
            End If
            
        End If
    Next
    Set vCStock = New cStock
    
End Sub











Private Sub cmdImpreFiltrado_Click()



        
        'select descripcion,date(horamovi) lafecha from proddepositoshco where tipoaccion In (8,9) group by 1 order by 2 desc
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        cad = "Código||idFiltrado|N|0000|10·Fecha||date(FechaHora)|T||15·Articulo||nomartic|T||40·"
        cad = cad & "Origen||DepositoInicial|T||9·Destino||DepositoFinal|T||9·Kilos||kilos|T||15·"

        
        
        frmB.vCampos = cad
        frmB.vTabla = "sfiltradoaceite ,spartidas inner join sartic on spartidas.codartic=sartic.codartic"
        frmB.vSQL = "idpartida=id"
        frmB.vDevuelve = "0|"
        frmB.vTitulo = "Filtrado"
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri  'Conexión a BD: Ariges
        cad = ""
        frmB.Show vbModal
        Set frmB = Nothing
        Screen.MousePointer = vbDefault
        If cad <> "" Then
            I = CInt(RecuperaValor(cad, 1))
            
            cad = DevuelveDesdeBDNew(conAri, "scryst", "documrpt", "codcryst", "41", "N")
 
            LlamaImprimirGral "{sfiltradoaceite.idFiltrado}=" & I, "", 0, cad, "Filtrado"
            
            
           
        End If
        
End Sub

Private Sub cmdLavall_Click()
Dim b As Boolean
Dim cOrig As cDeposito
Dim cDest As cDeposito
Dim Kilos As Currency
Dim CC As CTiposMov
   ' aqui aqui aqui
    b = False
   
    
    If cboDeposito(5).ListIndex = -1 Then Exit Sub
    If cboDeposito(6).ListIndex = -1 Then Exit Sub
    
    
    Set cOrig = New cDeposito
    Set cDest = New cDeposito
    
    If cOrig.LeerDatos(cboDeposito(5).ItemData(cboDeposito(5).ListIndex), False) Then
        If cDest.LeerDatos(cboDeposito(6).ItemData(cboDeposito(6).ListIndex), False) Then b = True
    End If
    
    
    
    
    If b Then
        
        If TodoCorrectoParaFiltrar(cOrig, cDest) Then
                Kilos = ImporteFormateado(txtNumeroDec(4).Text)
                conn.BeginTrans
                If Me.optLaVall(0).Value Then
                                    
                    
                    'FILTRADO
                    Set CC = New CTiposMov
                    CC.ConseguirContador ("TRO")
    
                    b = cOrig.HacerFiltrado(cDest, False, False, CC.contador + 1, CDate(Me.txtFecha(4).Text & " " & Me.txtHora(4).Text), Kilos)
                    
                    If b Then
                        CC.IncrementarContador CC.TipoMovimiento
                    
                        'hco filtrados
                        'sfiltradoaceite(idFiltrado,Trabajador,FechaHora,DepositoInicial,DepositoFinal,Kilos,usaAux8,usaAux9,idPartida)
                        cad = "(" & CC.contador & "," & TrabajadorConectado_ & "," & DBSet(CDate(Me.txtFecha(4).Text & " " & Me.txtHora(4).Text), "FH") & "," & cOrig.NumDeposito & ","
                        cad = cad & cDest.NumDeposito & "," & DBSet(Kilos, "N") & ",false,false," & cDest.idPartida & ")"
                        cad = "insert into sfiltradoaceite(idFiltrado,Trabajador,FechaHora,DepositoInicial,DepositoFinal,Kilos,usaAux8,usaAux9,idPartida) values " & cad
                        
                        If Not EjecutaSQL(conAri, cad, False) Then MsgBox "El programa continuará. Llame a soporte tecnico" & vbCrLf & cad, vbExclamation
                            
                    
                        HacerStockProductosFiltrados CC.contador, CDate(Me.txtFecha(4).Text & " " & Me.txtHora(4).Text)
            
                
                    End If
                              
                Else
                    b = cOrig.HacerTrasiego(cDest, Kilos = cOrig.Kilos, Kilos, CDate(Me.txtFecha(4).Text & " " & Me.txtHora(4).Text))
                
                End If
                If b Then
                    conn.CommitTrans
                Else
                    conn.RollbackTrans
                End If

        
                If b Then Unload Me
        
        
        End If


    End If
    Set cOrig = Nothing
    Set cDest = Nothing
End Sub

Private Sub cmdLote_Click(Index As Integer)
    If Me.txtArtFiltrado(Index) = "" Then Exit Sub
    
    
    cad = ""
    Set frmL = New frmAlmPartidas
    frmL.DatosADevolverBusqueda = txtArtFiltrado(Index).Text
    frmL.Show vbModal
    Set frmL = Nothing
    If cad <> "" Then
        'Comprobamos cantidad
        
        Me.txtLote(Index).Text = RecuperaValor(cad, 1)
        
   End If
    
End Sub

Private Sub cmdtrasiego_Click()
Dim C1 As cDeposito
Dim C2 As cDeposito
Dim Kilos As Currency
Dim Litros As Currency
Dim b As Boolean
'Si mueve el deposito entero, no genera NUEVO numero de lote
Dim MueveDepositoEntero As Boolean

    If cboDeposito(0).ListIndex < 0 Or cboDeposito(1).ListIndex < 0 Then Exit Sub
    If Me.txtFecha(3).Text = "" Or Me.txtHora(3).Text = "" Then
        MsgBox "Indique fecha hora", vbExclamation
        Exit Sub
    End If
    If CDate(txtFecha(3).Text) < vParamAplic.FechaActiva Then
        MsgBox "Menor que fecha activa", vbExclamation
        Exit Sub
    End If
    
    cad = "Va a realizar el trasiego: " & vbCrLf & "Origen: " & cboDeposito(0).Text
    cad = cad & vbCrLf & "Destino: " & cboDeposito(1).Text
    cad = cad & vbCrLf & "Fecha: " & Me.txtFecha(3).Text & " " & Me.txtHora(3).Text
    
    If vParamAplic.QUE_EMPRESA = 4 Then
        If Me.txtNumeroDec(0).Text = "" Then
            MsgBox "Debe indicar los kilos", vbExclamation
            Exit Sub
        End If
        
        cad = cad & vbCrLf & "Kilos : " & Me.txtNumeroDec(0).Text
        
        
    End If
    
    
    If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    
    
    'Hacemos el trasiego.
    Set C1 = New cDeposito
    Set C2 = New cDeposito
    
    If C1.LeerDatos(cboDeposito(0).ItemData(cboDeposito(0).ListIndex), False) Then
        If C2.LeerDatos(cboDeposito(1).ItemData(cboDeposito(1).ListIndex), False) Then
            
            b = True
            If vParamAplic.QUE_EMPRESA = 4 Then
                'Trasiego especifico la VALL
                'Factor conversion
                cad = "spartidas.id =partida and spartidas.codartic=sartic.codartic AND numdeposito "
                cad = DevuelveDesdeBD(conAri, "factorconversion", "spartidas,proddepositos,sartic", cad, C1.NumDeposito)
                If cad = "" Then
                    MsgBox "Error obteniendo articulo"
                    b = False
                Else
                    Litros = CCur(cad) 'factor conversion
                    
                    'Veremos si los kilos a traspasar son mas de los que hay o no
                    Kilos = ImporteFormateado(Me.txtNumeroDec(0).Text)
                                        
                    Litros = Round(Kilos / Litros, 2) '/factorconversion. Me da litros
                    cad = ""
                    If Litros > C2.Capacidad Then
                        cad = "Excede de la capacidad del deposito destino"
                    Else
                        If Kilos > C1.Kilos Then cad = "No tiene suficiente cantidad en el deposito origen"
                    End If
                    If cad <> "" Then
                        MsgBox cad, vbExclamation
                        b = False
                    Else
                        'Si la cantidad es igual a la del deposito, entonces MUEVE el deposito entero
                        MueveDepositoEntero = Kilos = C1.Kilos
                        
                    End If
                End If
            Else
            
                If Val(C2.idPartida) > 0 Then
                    MsgBox "Deposito destino contiene datos. partida " & C2.idPartida, vbExclamation
                    b = False
                End If
                MueveDepositoEntero = True
                Kilos = C1.Kilos
            End If
            If b Then
                'El que estaba
                conn.BeginTrans
                If C1.HacerTrasiego(C2, MueveDepositoEntero, Kilos, CDate(Me.txtFecha(3).Text & " " & Me.txtHora(3).Text)) Then
                    conn.CommitTrans
                    cad = ""
                Else
                    cad = "NO"
                    conn.RollbackTrans
                End If
                
            End If
        End If
    End If
    
    Set C1 = Nothing
    Set C2 = Nothing
    If cad = "" Then Unload Me
    
End Sub




Private Sub cmdTrasLavall_Click()
    If Me.FrameTrasiegoLaVAll.visible Then
         CargaComobosTrasiegos 0, 1
    Else
        cboDeposito(1).Clear
        cboDeposito(0).Clear
        
        Set miRsAux = New ADODB.Recordset
        cad = "Select partida,numdeposito from proddepositos where numdeposito=18"
        miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If miRsAux.EOF Then
            
        Else
            cad = "Deposito " & miRsAux!NumDeposito
            cboDeposito(1).AddItem cad
            cboDeposito(1).ItemData(cboDeposito(1).NewIndex) = miRsAux!NumDeposito
            cad = DBLet(miRsAux!partida, "T")
            miRsAux.Close
            
            If cad <> "" Then
                cad = "Select * from proddepositos where numdeposito<>18 AND partida=" & cad
            Else
                'Esta vacio. Cualquiera puede ser traspasado
                cad = "Select * from proddepositos where numdeposito<>18 AND partida>0 "
            End If
            miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                cad = Format(miRsAux!NumDeposito, "00") & "  "
                cad = cad & Mid(miRsAux!numLote & "       ", 1, 9) & " " & " Kilos: " & Format(miRsAux!Kilos, FormatoCantidad)
                cboDeposito(0).AddItem cad
                cboDeposito(0).ItemData(cboDeposito(0).NewIndex) = miRsAux!NumDeposito
                miRsAux.MoveNext
            Wend
            
        End If
        miRsAux.Close
            
        Set miRsAux = Nothing
    End If
    FrameTrasiegoLaVAll.visible = Not FrameTrasiegoLaVAll.visible
End Sub



Private Sub cmdVaciadoDeposito_Click()
Dim C1 As cDeposito
Dim F As Date
    
    
    If vParamAplic.QUE_EMPRESA = 4 Then
    
        If txtFecha(5).Text = "" Or txtHora(5).Text = "" Then
            MsgBox "Fecha / Hora obligatorio", vbExclamation
            Exit Sub
        End If
        'En VALL pueden haber el mismo lote en varios deposi    tos. Comprobaremos que no es asi
        Set C1 = New cDeposito
        cad = ""
        If C1.LeerDatos(cboDeposito(2).ItemData(cboDeposito(2).ListIndex), True) Then
            cad = DevuelveDesdeBD(conAri, "numdeposito", "proddepositos", "numdeposito<>" & C1.NumDeposito & " AND partida ", C1.idPartida)
            If cad <> "" Then
                cad = "LOTE en varios depositos"
            End If
               
            If C1.AvisarMovimientoHcoPosterior(CDate(txtFecha(5).Text & " " & txtHora(5).Text)) Then cad = cad & vbCrLf & " -Movimientos posteriores"
               
               
        Else
            cad = "Error leyendo deposito"
        End If
        Set C1 = Nothing
        
        
        
        
        If cad <> "" Then
            MsgBox cad, vbExclamation
            Exit Sub
        End If
    
    
    
    
    End If
    
    
    
    TrabajadorConectado_ = Val(PonerTrabajadorConectado(cad))
    If MsgBox("Continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    
    
    cad = "N"
    Set C1 = New cDeposito
    F = Now
    If vParamAplic.QUE_EMPRESA = 4 Then F = CDate(txtFecha(5).Text & " " & txtHora(5).Text)
    If C1.LeerDatos(cboDeposito(2).ItemData(cboDeposito(2).ListIndex), False) Then
        If C1.Kilos > 0 Then
            'DEBERIAOS REGULARIZAR
            
            RegularizarFinLote_Partida_DEPOS C1, F
        End If
        
        C1.QuitarAsignacionDeposito2 2, F, -C1.Kilos
        
        cad = ""
    End If
    Set C1 = Nothing
    If cad = "" Then Unload Me
End Sub

Private Sub Command1_Click()
    cmdImpreFiltrado_Click
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        
        If Opcion = 5 Then
            Screen.MousePointer = vbHourglass
            cmdCoupage.visible = False
            cmdCancelar(1).visible = False
            DoEvents
            Me.Refresh
            
            cmdCoupage_Click
            Screen.MousePointer = vbDefault
        End If
    End If
End Sub

Private Sub Form_Load()
    PrimeraVez = True
    Me.Icon = frmppal.Icon
    FrCierreOrdenProduccion.visible = False
    FrCoupage.visible = False
    FrameTrasiego.visible = False
    FrameVaciado.visible = False
    FrameFiltrado.visible = False
    limpiar Me
    TrabajadorConectado_ = Val(PonerTrabajadorConectado(cad))
    Select Case Opcion
    Case 0
        PonerFrameVisible FrCierreOrdenProduccion
        Me.Caption = "Cierre orden producción"
        lbFec(0).Caption = lbFec(0).Caption & ": " & RecuperaValor(Intercambio, 1) & " " & RecuperaValor(Intercambio, 2)
        Me.txtMeses.Text = "18"
        
        
    Case 1, 5
        PonerFrameVisible Me.FrCoupage
        Me.Caption = "Hacer coupage"
        lbFec(1).Caption = lbFec(1).Caption & ": " & RecuperaValor(Intercambio, 1) & " " & RecuperaValor(Intercambio, 2)
        If Opcion = 5 Then
            lbFec(1).Caption = lbFec(1).Caption & "   ---> Desde molturacion"
            cad = DevuelveDesdeBD(conAri, "fecha", "olicoupage", "codigo", RecuperaValor(Intercambio, 1))
            txtFecha(1).Text = Format(cad, "dd/mm/yyyy")
            cad = Format(cad, "hh:mm:ss")
            txtHora(1).Text = DateAdd("s", 1, cad)
                       
        End If
                       
            
        
    Case 2
        PonerFrameVisible FrameTrasiego
        Me.Caption = "Realizar trasiego"
        FrameTrasiegoLaVAll.visible = vParamAplic.QUE_EMPRESA = 4
        CargaComobosTrasiegos 0, 1
        
    Case 3
        FrameFechaVall.visible = vParamAplic.QUE_EMPRESA = 4
        FrameFechaVall.BorderStyle = 0
    
        PonerFrameVisible FrameVaciado
        Me.Caption = "Vaciar deposito"
        CargaComobosTrasiegos 2, 2
    Case 4
        PonerFrameVisible FrameFiltrado
        Me.Caption = "Filtrado"
        CargaComobosTrasiegos 3, 4
        FramefiltroMorales.visible = vParamAplic.QUE_EMPRESA = 1
        'Cargamos los articulos de vallparam (servira tanto para morales como para La VALL
        'diatomeasRojas diatomeasVerdas  celulosa
        cad = DevuelveDesdeBD(conAri, "concat(diatomeasRojas ,'|',diatomeasVerdas,'|',celulosa,'|')", "vallparam", "1", "1")
                
        For I = 0 To 2
            BloquearTxt Me.txtArtFiltrado(I), True
            BloquearTxt Me.txtArtFiltrado(I + 3), True
            txtArtFiltrado(I).Text = RecuperaValor(cad, I + 1)
            If txtArtFiltrado(I).Text <> "" Then txtArtFiltrado(I + 3).Text = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", txtArtFiltrado(I).Text, "T")
        Next I
    Case 6
        PonerFrameVisible FrameLaVallTraFiltrado
        Me.Caption = vEmpresa.nomempre
        CargaComobosLaVall
    
    
             'Cargamos los articulos de vallparam (servira tanto para morales como para La VALL
        'diatomeasRojas diatomeasVerdas  celulosa
        cad = DevuelveDesdeBD(conAri, "concat(diatomeasRojas ,'|',diatomeasVerdas,'|',celulosa,'|')", "vallparam", "1", "1")
                
        For I = 0 To 2
            BloquearTxt Me.txtArtFiltrado(I), True
            BloquearTxt Me.txtArtFiltrado(I + 3), True
            txtArtFiltrado(I).Text = RecuperaValor(cad, I + 1)
            If txtArtFiltrado(I).Text <> "" Then txtArtFiltrado(I + 3).Text = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", txtArtFiltrado(I).Text, "T")
        Next I
    
    
    
    
    End Select
    If Opcion = 5 Then
        cmdCancelar(1).Cancel = True
    Else
        cmdCancelar(Opcion).Cancel = True
    End If
End Sub



Private Sub PonerFrameVisible(ByRef Fr As Frame)

    Fr.visible = True
    Fr.Top = 30
    Fr.Left = 30
    Me.Width = Fr.Width + 180
    Me.Height = Fr.Height + 520
    
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
    cad = CadenaDevuelta
End Sub

Private Sub frmC_Selec(vFecha As Date)
    txtFecha(I).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmL_DatoSeleccionado(CadenaSeleccion As String)
    cad = CadenaSeleccion
End Sub

Private Sub imgFecha_Click(Index As Integer)
    'El index tiene que ser el mismo que el del txtfecha al que acompaña
    Set frmC = New frmCal
    frmC.Fecha = Now
    I = Index
    If txtFecha(Index).Text <> "" Then frmC.Fecha = CDate(txtFecha(Index).Text)
    frmC.Show vbModal
    Set frmC = Nothing
    
End Sub

Private Sub optLaVall_Click(Index As Integer)
    FrameFilltroLaVall2.visible = Index = 0
End Sub

Private Sub optLaVall_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtArtFiltrado_LostFocus(Index As Integer)
Dim C As String
    If Index >= 0 And Index <= 2 Then
        cad = ""
        txtArtFiltrado(Index).Text = Trim(txtArtFiltrado(Index).Text)
        If txtArtFiltrado(Index).Text <> "" Then
            C = "codartic"
            cad = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", txtArtFiltrado(Index).Text, "T", C)
            If cad = "" Then
                MsgBox "No existe articulo: " & txtArtFiltrado(Index).Text
            Else
                txtArtFiltrado(Index).Text = C
            End If
        End If
        txtArtFiltrado(Index + 3).Text = cad
    End If
End Sub

Private Sub txtFecha_GotFocus(Index As Integer)
    ConseguirFoco txtFecha(Index), 3
End Sub

Private Sub txtFecha_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtFecha_LostFocus(Index As Integer)
Dim T As String
    txtFecha(Index).Text = Trim(txtFecha(Index).Text)
    If txtFecha(Index).Text <> "" Then
        T = txtFecha(Index).Text
        If EsFechaOK(T) Then
            txtFecha(Index).Text = T
        Else
            MsgBox "Fecha con formato incorrecto: " & txtFecha(Index).Text, vbExclamation
            txtFecha(Index).Text = ""
            PonerFoco txtFecha(Index)
        End If
    End If
    
End Sub

Private Function CerrarOrdenProduccion(SoloComprobar As Boolean) As Boolean
Dim vCStock As cStock
Dim b As Boolean

    'ACciones a realizar
    'Comprobar stock sublineas, ya que es la que van a disminuir la cantidad
    'Damos de alta en stock (y smoval) las lienas ppales
    'Damos de baja   "        "        las sublineas
    CerrarOrdenProduccion = False
    Set miRsAux = New ADODB.Recordset
    Set vCStock = New cStock
    
    'Veamos las sub lineas  si tienen stock. Antes comprobabamos cantidad x sarti1.cntidad
    'Cad = "select codarti1,codalmac,sarti1.cantidad multiplicador,sum(sliordpr.cantidad) cantilinea from sliordpr,sarti1 where "
    'Cad = Cad & " sliordpr.codartic=sarti1.codartic and  codigo=" & RecuperaValor(Me.Intercambio, 1) & " group by 1,2,3"
    'AHora hay una tabla para los componentes
'    Cad = "select codarti2,sliordpr.codalmac,sliordpr2.cantidad cantilinea from sliordpr,sliordpr2 where"
'    Cad = Cad & " sliordpr.codartic=sliordpr2.codartic and sliordpr.codalmac=sliordpr2.codalmac and"
'    Cad = Cad & " sliordpr.codigo=1 group by 1,2"
'
    cad = "select sliordpr2.*,sartic.factorconversion from sliordpr2,sartic where sliordpr2.codarti2=sartic.codartic and codigo=" & RecuperaValor(Me.Intercambio, 1)
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    b = False
    
    If Not SoloComprobar Then conn.BeginTrans

    
    While Not miRsAux.EOF

        b = False
        If InicializarCStock(vCStock, "S", True) Then
            
            If vCStock.MueveStock Then
                If SoloComprobar Then
                    b = vCStock.MoverStock(False)
                Else
                    'Estamos ejecutando la actualizacion
                    '---------------------------------------------
                    'si hay control de stock para el articulo actualizar en salmac e insertar en smoval
                    'en actualizar stock comprobamos si el articulo tiene control de stock
                    b = vCStock.ActualizarStock(False)
                End If
            Else
                b = True
            End If
        End If
                             
        
        If Not b Then
            While Not miRsAux.EOF
                miRsAux.MoveNext  'para que no siga
            Wend
        Else
            'Al siguiente
            miRsAux.MoveNext
        End If
    Wend
    miRsAux.Close
  
    
    If Not b Then
        Set miRsAux = Nothing
        Set vCStock = Nothing
        If Not SoloComprobar Then conn.RollbackTrans
        Exit Function 'Si no puede inicializar los stocks, de las sublineas salimos
    End If
    
    
    
    
    'SSi ha ido bien comprobamos los LOTES
    If Not RealizarProduccionLOTES(SoloComprobar) Then
    
            Set miRsAux = Nothing
            Set vCStock = Nothing
            If Not SoloComprobar Then conn.RollbackTrans
            Exit Function 'Si no puede inicializar los stocks, de las sublineas salimos
    End If
    Set miRsAux = Nothing
    Set miRsAux = New ADODB.Recordset
    
    'AHora comprobamos los stcosk de las entraddas , de las lineas            factor=1
    cad = "select codartic codarti2,codalmac,sum(sliordpr.cantidad) cantidad,1 factorconversion,numlote from sliordpr where "
    cad = cad & " codigo=" & RecuperaValor(Me.Intercambio, 1) & " group by 1,2"
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    b = False
    While Not miRsAux.EOF
        b = False
        If InicializarCStock(vCStock, "E", False) Then   'Las lineas son de netrada
            If vCStock.MueveStock Then
                If SoloComprobar Then
                   ' B = vCStock.MoverStock(False, True)
                   b = True
                Else
                    b = vCStock.ActualizarStock(False)
                End If
            Else
                b = True
            End If
        End If
        
        If Not b Then
            While Not miRsAux.EOF
                miRsAux.MoveNext  'para que no siga
            Wend
        Else
            If Not SoloComprobar Then
                '-------------------------- LOTES
                
                'Si ha puesto numero de lote
                
            End If
            'Al siguiente
            miRsAux.MoveNext
        End If
    Wend
    miRsAux.Close
    
    
    If Not b Then
        Set miRsAux = Nothing
        Set vCStock = Nothing
        If Not SoloComprobar Then conn.RollbackTrans
        Exit Function 'Si no puede inicializar los stocks, de las sublineas salimos
    End If
    
    
    
    
    'Acutailizaremos algnas cosas como la fecha de baja
    If Not SoloComprobar Then
        conn.CommitTrans
        cad = "UPDATE sordprod set fecproduccion = " & DBSet(txtFecha(0).Text, "F")
        'Marzo 2012. Caducidad
        cad = cad & ",feccaduca  = " & DBSet(DateAdd("m", Val(txtMeses.Text), CDate(txtFecha(0).Text)), "F")
        cad = cad & " WHERE  codigo=" & RecuperaValor(Me.Intercambio, 1)
        conn.Execute cad
        
        
        'Para LA VALL, si el articulo producido esta en algun albaran en SCAALB que avise
        cad = "select distinct numalbar from slialb where codartic in (select codartic from sliordpr where codigo=" & RecuperaValor(Me.Intercambio, 1) & ")"
        miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        cad = ""
        While Not miRsAux.EOF
            If cad <> "" Then cad = cad & " - "
            cad = cad & miRsAux!NumAlbar
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        If cad <> "" Then MsgBox "Existen albaranes con esta referencia: " & vbCrLf & vbCrLf & cad, vbInformation
        
        
    End If
    
    CerrarOrdenProduccion = True
    
    Set miRsAux = Nothing
    Set vCStock = Nothing
    
    
End Function






'-------------------------------------------------------------------------
' Realizar el coupage
'
Private Function RealizarCoupage(SoloComprobar As Boolean) As Boolean
Dim vCStock As cStock
Dim b As Boolean
Dim CantidadTotalAProducir As Currency 'Cuatro decimales


    'ACciones a realizar
    'Comprobar stock sublineas, ya que es la que van a disminuir la cantidad
    'Damos de alta en stock (y smoval) las lienas ppales
    'Damos de baja   "        "        las sublineas
    RealizarCoupage = False
    Set miRsAux = New ADODB.Recordset
    Set vCStock = New cStock
    
    
    
    
    
    If Not SoloComprobar Then conn.BeginTrans

        
        
    
    'Los mezclantes
    'Como no lleva factor conversion. Necesito los precios para los calculos de importes
    cad = "select olicoupagelin.*,preciouc, preciomp from olicoupagelin,sartic where olicoupagelin.codartic=sartic.codartic and "
    cad = cad & "  codigo = " & RecuperaValor(Me.Intercambio, 1)
    'cad = "select * from olicoupagelin where codigo = " & RecuperaValor(Me.Intercambio, 1)
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    b = False
    
    CantidadTotalAProducir = 0
    While Not miRsAux.EOF
        b = False
        If InicializarCStockCoupage(vCStock, "S", False) Then
            
            If vCStock.MueveStock Then
                If SoloComprobar Then
                    b = vCStock.MoverStock(False)
                Else
                    'Estamos ejecutando la actualizacion
                    '---------------------------------------------
                    'si hay control de stock para el articulo actualizar en salmac e insertar en smoval
                    'en actualizar stock comprobamos si el articulo tiene control de stock
                    b = vCStock.ActualizarStock(False)
                End If
            Else
                b = True
            End If
            CantidadTotalAProducir = CantidadTotalAProducir + miRsAux!Kilos
        End If
        
        If Not b Then
            While Not miRsAux.EOF
                miRsAux.MoveNext  'para que no siga
            Wend
        Else
            'Al siguiente
            miRsAux.MoveNext
        End If
    Wend
    miRsAux.Close
    
    
    If Not b Then
        Set miRsAux = Nothing
        Set vCStock = Nothing
        If Not SoloComprobar Then conn.RollbackTrans
        Exit Function 'Si no puede inicializar los stocks, de las sublineas salimos
    End If
    
    
    'SSi ha ido bien comprobamos los LOTES
    If Not RealizarCoupageLOTES(SoloComprobar, CantidadTotalAProducir) Then
    
            Set miRsAux = Nothing
            Set vCStock = Nothing
            If Not SoloComprobar Then conn.RollbackTrans
            Exit Function 'Si no puede inicializar los stocks, de las sublineas salimos
    End If
    Set miRsAux = Nothing
    Set miRsAux = New ADODB.Recordset
    
    
    
    
    
    
    'AHora comprobamos los stcosk de las entraddas , de las lineas
    cad = TransformaComasPuntos(CStr(CantidadTotalAProducir))
    
    cad = "select olicoupage.codartic," & cad & " kilos,preciouc from olicoupage,sartic where"
    cad = cad & " olicoupage.codartic=sartic.codartic and codigo=" & RecuperaValor(Me.Intercambio, 1)
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    b = False
    
    While Not miRsAux.EOF
        b = False
        If InicializarCStockCoupage(vCStock, "E", False) Then   'Las lineas son de netrada
        
            If vCStock.MueveStock Then
                If SoloComprobar Then
                    'B = vCStock.MoverStock(False)
                    b = True
                    
                    
                    
                    
                Else
                    b = vCStock.ActualizarStock(False)
                End If
            Else
                b = True
            End If
        End If
        
        If Not b Then
            While Not miRsAux.EOF
                miRsAux.MoveNext  'para que no siga
            Wend
        Else
            'Al siguiente
            miRsAux.MoveNext
        End If
    Wend
    miRsAux.Close
    
    
    If Not b Then
        Set miRsAux = Nothing
        Set vCStock = Nothing
        If Not SoloComprobar Then conn.RollbackTrans
        Exit Function 'Si no puede inicializar los stocks, de las sublineas salimos
    End If
    
    
    
    
    'Acutailizaremos algnas cosas como la fecha de baja
    If Not SoloComprobar Then
        conn.CommitTrans
        cad = "UPDATE olicoupage set YaCreado = 1"
        cad = cad & " WHERE  codigo=" & RecuperaValor(Me.Intercambio, 1)
        conn.Execute cad
    End If
    
    
        
    
    
    RealizarCoupage = True
    
    Set miRsAux = Nothing
    Set vCStock = Nothing
    
    
End Function






'No le paso el recodset pq es mirsaux que es comun
Private Function InicializarCStock(ByRef vCStock As cStock, TipoM As String, Sublineas As Boolean) As Boolean
Dim CantidadNecesaria As Single
Dim MateriaPrima As Boolean
    
    On Error Resume Next

    vCStock.tipoMov = TipoM
    vCStock.DetaMov = "PRO"
    vCStock.Trabajador = TrabajadorConectado_
    vCStock.Documento = RecuperaValor(Intercambio, 1)
    vCStock.Fechamov = txtFecha(0).Text '
    vCStock.codAlmac = CInt(miRsAux!codAlmac)
        
    If vParamAplic.QUE_EMPRESA = 4 Then vCStock.HoraMov = vCStock.Fechamov & " " & Format(txtHora(0).Text, "hh:mm:ss")
   
    
    
    CantidadNecesaria = miRsAux!FactorConversion
    MateriaPrima = False
    If CantidadNecesaria <> 1 Then MateriaPrima = True
    
    'mAYO 2010.   eL FACTOR CONVERSION VIENE ya grabado en sliorpr2
    '           quiero decir que no hay que volver a multiplcarlo
    'If CantidadNecesaria <> 1 Then
    CantidadNecesaria = 1  'YA hemos grabado la sliordpr
    
    If Sublineas Then
        If vCStock.codAlmac = 2 And Not MateriaPrima Then
            'Es el del B
            'Solo el aceite vendra de las garrafas de B. Lo demas todo del limpio
             vCStock.codAlmac = 1
        End If
    End If
    vCStock.codArtic = miRsAux!codarti2
    
   
    If CantidadNecesaria = 0 Then CantidadNecesaria = 1 'PARA QUE NO DE ERROR
    CantidadNecesaria = Round2(miRsAux!Cantidad * CantidadNecesaria, 5)
    vCStock.Cantidad = CantidadNecesaria
    vCStock.Importe = 0
    vCStock.LineaDocu = 0

    If Err.Number <> 0 Then
        MsgBox "No se han podido inicializar la clase para actualizar Stock:" & Err.Description, vbExclamation
        InicializarCStock = False
    Else
        InicializarCStock = True
    End If
End Function

Private Function InicializarCStockCoupage(ByRef vCStock As cStock, TipoM As String, ParaLotes As Boolean) As Boolean
'Dim CantidadNecesaria As Single   'No lleva factor conversion, ya que esta en KILOS que es el stcok
Dim Impor As Currency
    On Error Resume Next

    vCStock.tipoMov = TipoM
    vCStock.DetaMov = "CUP"  'Coupages
   ' vCStock.Trabajador = PonerTrabajadorConectado(cad)
    vCStock.Trabajador = TrabajadorConectado_
    vCStock.Documento = RecuperaValor(Intercambio, 1)
    vCStock.Fechamov = txtFecha(1).Text '
    If vParamAplic.QUE_EMPRESA = 4 Then
        vCStock.HoraMov = vCStock.Fechamov & " " & Format(txtHora(1).Text, "hh:mm:ss")
    Else
        If txtHora(1).Text <> "" Then
            vCStock.HoraMov = vCStock.Fechamov & " " & Format(txtHora(1).Text, "hh:mm:ss")
        Else
            vCStock.HoraMov = vCStock.Fechamov & " " & Format(Now, "hh:mm:ss")
        End If
    End If
    vCStock.codArtic = miRsAux!codArtic
    vCStock.codAlmac = RecuperaValor(Intercambio, 3)
'    CantidadNecesaria = miRsAux!FactorConversion
'    If CantidadNecesaria = 0 Then CantidadNecesaria = 1 'PARA QUE NO DE ERROR
'    CantidadNecesaria = Round2(miRsAux!kilos / CantidadNecesaria, 5)
'    vCStock.Cantidad = CantidadNecesaria
    vCStock.Cantidad = miRsAux!Kilos
    If Not ParaLotes Then
        Impor = DBLet(miRsAux!PrecioUC, "N")
        Impor = Round2(Impor * vCStock.Cantidad, 4)
        vCStock.Importe = Impor
    Else
        vCStock.Importe = 0
    End If
    vCStock.LineaDocu = 0

    If Err.Number <> 0 Then
        MsgBox "No se han podido inicializar la clase para actualizar Stock", vbExclamation
        InicializarCStockCoupage = False
    Else
        InicializarCStockCoupage = True
    End If
End Function





Private Function RealizarProduccionLOTES(SoloComprobar As Boolean) As Boolean
Dim ErroresEnPartidas As String
Dim LotesNecesartios As Collection
Dim CantidadNecesaria As Currency
Dim AuxPartida As String
Dim Err_x_Articulo As String
Dim MiNumeroLote As String
Dim cP As cPartidas   'Para los numeros de lote
Dim Rc As Byte
Dim vvCstock As cStock
Dim b As Boolean
Dim RL As ADODB.Recordset
Dim CantidadQueLLevo As Currency
Dim Aux As String
Dim TieneLotesMP As Boolean
Dim II As Integer
Dim Cant2 As Currency
Dim cL As cLotaje
Dim LoteReal As String  'Con fecha
Dim cDe As cDeposito
Dim ParaDeposito As String
Dim Secuencia As Integer
Dim leerDeLotesAsignadosEnProduccion As Boolean
Dim OtroTexto As String
Dim Cadena2 As String


    On Error GoTo ERealizarProduccionLOTES

    RealizarProduccionLOTES = False
    ErroresEnPartidas = ""
    AuxPartida = ""
    Set cP = New cPartidas

    If Not SoloComprobar Then

        Set cL = New cLotaje
        cL.DetaMov = "PRO"
        cL.Documento = RecuperaValor(Intercambio, 1)
        cL.Fechamov = CDate(Me.txtFecha(0).Text)
        If vParamAplic.QUE_EMPRESA = 4 Then
            cL.HoraMov = CDate(Me.txtFecha(0).Text & " " & txtHora(0).Text)
        Else
            cL.HoraMov = CDate(Me.txtFecha(0).Text & " " & Format(Now, "hh:nn:ss"))
        End If
        cL.ProvCliTra = TrabajadorConectado_
        cL.LineaDocu = 0
        cL.SubLinea = 0
    End If
        


    cad = "select sliordpr2.*,sartic.factorconversion,trazabilidad,nomartic from sliordpr2,sartic where "
    cad = cad & " sliordpr2.codarti2=sartic.codartic and codigo=" & RecuperaValor(Me.Intercambio, 1)
    cad = cad & " AND trazabilidad = 1" 'Solo miraremos los que lleven trazabilidad
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    b = False
    


    AuxPartida = ""
    Secuencia = 0
    Set vvCstock = New cStock
    While Not miRsAux.EOF
        Secuencia = Secuencia + 1 'Para la hora de insercion en le deposito
        
        If Err_x_Articulo <> miRsAux!codArtic Then
            'Han habido errores en el articulo anterior.
            If AuxPartida <> "" Then
                cad = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", DevNombreSQL(Err_x_Articulo), "T")
                AuxPartida = "-  " & Err_x_Articulo & "  " & cad & AuxPartida & vbCrLf
                ErroresEnPartidas = ErroresEnPartidas & AuxPartida & vbCrLf
            End If
            Err_x_Articulo = miRsAux!codArtic
            AuxPartida = ""
        End If

        b = False
        If InicializarCStock(vvCstock, "E", False) Then   'Las lineas son de netrada
  
            CantidadNecesaria = vvCstock.Cantidad  'Para tenerla despues en los lotes

            '// NUmeros de LOTE
            ' Las materias primas (en ppio solo ellas) pueden forzarle los lotes
            'en el mantenimiento de produccion. Con lo cual, si se lo han asignado comprobare
            'que de lo que le asignan tengo disponible. Si no se lo asigno YO
            Set LotesNecesartios = New Collection
            Set RL = New ADODB.Recordset
            'De momento solo para las MATERIAS PRIMAS
            ' factorconversion<>1
            If miRsAux!FactorConversion = 1 Then
                'Vere si tiene algun lote asiganad a mano
                Aux = "Select * from sliordpr2lotes WHERE  codigo = " & RecuperaValor(Intercambio, 1)
                Aux = Aux & " AND codalmac =" & vvCstock.codAlmac & " AND codArtic = " & DBSet(miRsAux!codArtic, "T")
                Aux = Aux & " AND codArti2 = " & DBSet(vvCstock.codArtic, "T")
                RL.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If RL.EOF Then
                    TieneLotesMP = False
                    leerDeLotesAsignadosEnProduccion = False
                Else
                    leerDeLotesAsignadosEnProduccion = True
                    TieneLotesMP = True
                End If
                RL.Close
            Else
                leerDeLotesAsignadosEnProduccion = True
                TieneLotesMP = True
            End If
            If leerDeLotesAsignadosEnProduccion Then
                Aux = "Select * from sliordpr2lotes WHERE  codigo = " & RecuperaValor(Intercambio, 1)
                Aux = Aux & " AND codalmac =" & vvCstock.codAlmac & " AND codArtic = " & DBSet(miRsAux!codArtic, "T")
                Aux = Aux & " AND codArti2 = " & DBSet(vvCstock.codArtic, "T")
                Set RL = New ADODB.Recordset
                RL.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If RL.EOF Then
                    TieneLotesMP = False
                Else
                    TieneLotesMP = True
                    CantidadQueLLevo = CantidadNecesaria
                    'Para cada lote especficiado veremos SI existe el lote o no en partidas
                    While Not RL.EOF
                        'ANTES MAYO 2010
                        'Cant2 = Round2(miRsAux!FactorConversion * RL!cantlote, 5)
                        'AHORA. Mayo 2010.  YA he grabado la sliord2 con el factor conversion multimplicado NO debo volver a miultiplicarlo
                        Cant2 = RL!cantlote
                        
                        
                        CantidadQueLLevo = CantidadQueLLevo - Cant2
                    
                        Aux = RL!numLote & "|"
                        RL.MoveNext
                        If RL.EOF Then
                            'Es la utlima. Ajusto los decimales
                            If CantidadQueLLevo > 0 Then Cant2 = Cant2 + CantidadQueLLevo
                        End If
                        Aux = Aux & Cant2 & "|"
                        LotesNecesartios.Add Aux
                        
                        
                        
                    Wend
                    RL.Close
                    CantidadQueLLevo = 0
                End If
            End If
            
            If TieneLotesMP Then
                
                'Los busco en partidas
                For II = 1 To LotesNecesartios.Count
                    Aux = LotesNecesartios(II)
                    Cant2 = CCur(RecuperaValor(Aux, 2))
                    Aux = RecuperaValor(Aux, 1)
                    Aux = "  AND numlote = '" & DevNombreSQL(Aux) & "'"
                    Aux = " AND codalmac =" & vvCstock.codAlmac & Aux
                    Aux = " where codartic = " & DBSet(vvCstock.codArtic, "T") & Aux
                    Aux = "Select id,cantotal from spartidas " & Aux
                    RL.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    
                    If RL.EOF Then
                        'NO existe el registro en partidas para ese LOTE - articulo
                        cad = "NO existe LOTE: " & RecuperaValor(LotesNecesartios(II), 1)
                        If Not SoloComprobar Then
                            cP.Cantidad = -1 * CantidadNecesaria
                            cP.codAlmac = vvCstock.codAlmac
                            cP.codArtic = vvCstock.codArtic
                            cP.codProve = 0
                            cP.Fecha = vvCstock.Fechamov
                            
                            cP.NumAlbar = "PR" & miRsAux!Codigo
                            cP.numLote = cP.NumAlbar
                            cP.Insertar


                            InsertarMovientosLotesProduccion cL, cP, cP.Cantidad, miRsAux!codArtic

                            'Si es aceite..
                            
                        End If
                        
                    Else
                        'SI que existe el LOTE veamos si tiene suficiente
                        If RL!cantotal < Cant2 Then
                            'No tengo suficiente
                            'FALTA
                            cad = "No tengo suficiente. (" & LotesNecesartios(II) & ")"

    
                        Else
                            'Todo OK
                            cad = ""
                            
                        End If
                        
                        If cad = "" Then
                            If miRsAux!FactorConversion < 1 Then
                                ParaDeposito = LotesNecesartios(II)
                                ParaDeposito = RecuperaValor(ParaDeposito, 1)
                                
                                
                                'OK. Es el aceite virgen. Vamos a buscar su deposito
                                Set cDe = New cDeposito
                                ParaDeposito = DevuelveDesdeBD(conAri, "numdeposito", "proddepositos", "numlote", ParaDeposito, "T")
                                If vParamAplic.QUE_EMPRESA = 4 Then
                                    
                                    'Para LAVall nop puede envasar fuera del 18.  Noviembre 2017. Permitimos envasar desde otro
                                    If ParaDeposito <> "18" Then
                                            
                                        'SOOOOLO envasamos del 18
                                        Cadena2 = DevuelveDesdeBD(conAri, "numlote", "proddepositos", "numdeposito", "18", "T")
                                        If Cadena2 <> RecuperaValor(LotesNecesartios(II), 1) Then
                                            If SoloComprobar Then
                                                cad = vbCrLf & "-No es depósito 18. Deposito selecc: " & ParaDeposito
                                            Else
                                                'Err.Raise 513, , "No es deposito 18"
                                                Dim OtroAux As String
                                                
                                                OtroAux = RecuperaValor(LotesNecesartios(II), 1)
                                                OtroAux = DevuelveDesdeBD(conAri, "numdeposito", "proddepositos", "numdeposito<>" & ParaDeposito & " AND numlote", OtroAux, "T")
                                                If OtroAux <> "" Then
                                                    'Esta en mas de un deposito
                                                    OtroAux = "Esta en mas de un deposito. (" & ParaDeposito & " : " & OtroAux & "..)"
                                                    OtroAux = OtroAux & vbCrLf & vbCrLf & "Indique numero deposito"
                                                    Cadena2 = InputBox(OtroAux, "Nº deposito", ParaDeposito)
                                                    If Cadena2 = "" Then
                                                        'Cancelamos
                                                        Err.Raise 513, , "Cancelacion en seleccion deposito"
                                                    Else
                                                        If Not IsNumeric(Cadena2) Then
                                                            Err.Raise 513, , "Deposito debe ser campo numerico"
                                                        Else
                                                            If Val(Cadena2) < 1 Or Val(Cadena2) > MaxNumDepositos_ Then
                                                                Err.Raise 513, , "Deposito entre 1 y " & MaxNumDepositos_
                                                            
                                                            Else
                                                                'El deposito contiene el lote
                                                                OtroAux = RecuperaValor(LotesNecesartios(II), 1)
                                                                OtroAux = DevuelveDesdeBD(conAri, "numdeposito", "proddepositos", "numdeposito=" & Cadena2 & " AND numlote", OtroAux, "T")
                                                                If OtroAux = "" Then
                                                                    Err.Raise 513, , "Deposito no contiene la muestra procesada"
                                                                Else
                                                                    'OK. CONFIRMACION
                                                                    If MsgBox("Deposito seleccionado: " & Cadena2 & vbCrLf & vbCrLf & "¿Continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then Err.Raise 513, , "Cancelacion en seleccion deposito (II)"
                                                                    ParaDeposito = Cadena2
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                            Cadena2 = ""
                                            
                                        Else
                                            ParaDeposito = "18"
                                        End If
                                    End If
                                    
                                  
                                End If
                                If cad = "" Then
                                    If vParamAplic.QUE_EMPRESA = 3 Then
                                        'OK. Moixent NO lleva depositos
                                    Else
                                        If Not cDe.LeerDatos(CInt(ParaDeposito), True) Then
                                            cad = "Error leyendo datos deposito " & ParaDeposito 'NO DEBERIA PASAR NUNCA
                                        Else
                                        
                                            If SoloComprobar Then
                                                vvCstock.HoraMov = vvCstock.Fechamov & " " & Format(Now, "hh:mm:ss")
                                                If cDe.AvisarMovimientoHcoPosterior(vvCstock.HoraMov) Then
                                                    If MsgBox("Movimientos posteriores en el deposito. ¿Continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then cad = vbCrLf & "- Moviemientos posteriores en deposito"
                                                End If
                                            End If
                                            If Not SoloComprobar Then
                                                cDe.VariacionKilosDeposito -Cant2
                                                cDe.InsertarEnHco 5, DateAdd("s", Secuencia, cL.HoraMov), "Prod: " & Format(RecuperaValor(Intercambio, 1), "00") & " - " & RecuperaValor(Intercambio, 2), -Cant2
                                                Espera 0.95 'porque si no puede dar entrada duplicada
                                                
                                                
                                                'Updatemaso las lineas de produccion para indicar deposito Paradeposito
                                                cad = "UPDATE sliordpr2lotes SET numdeposito = " & Val(ParaDeposito)
                                                cad = cad & " WHERE  codigo = " & RecuperaValor(Intercambio, 1)
                                                cad = cad & " AND codalmac =" & vvCstock.codAlmac & " AND codArtic = " & DBSet(miRsAux!codArtic, "T")
                                                cad = cad & " AND codArti2 = " & DBSet(vvCstock.codArtic, "T")
                                                conn.Execute cad
                                                
                                            End If
                                        End If
                                    End If
                                End If
                                Set cDe = Nothing
                              End If
                         
                           
                        End If
                        
                        
                        'Si estamos ya realizando la produccion actualizamos tablas
                        If Not SoloComprobar Then
                            
                            cP.Leer Val(RL!ID)
                            cP.IncrementarCantidad -1 * Cant2
                            
                            InsertarMovientosLotesProduccion cL, cP, -1 * Cant2, miRsAux!codArtic
                        End If
                    End If
                    RL.Close
                    If SoloComprobar Then
                        If cad <> "" Then
                                cad = Space(19) & "-- " & vvCstock.codArtic & "  " & Mid(miRsAux!NomArtic & Space(45), 1, 45) & cad
                                AuxPartida = AuxPartida & vbCrLf & cad
                        End If
                    
                    End If
                Next   'LotesNecesartios.Count
            
            Else
                
                'Asi es como estaba antes
                Rc = cP.RecuperarLotes(vvCstock.codArtic, vvCstock.codAlmac, CantidadNecesaria, LotesNecesartios)
            
                If Rc = 2 Then
                    'No tengo el articulo dado de alta
                    cad = "NO hay ningun lote "
                    
                    'Si estoyNO es solo comprobar, entonces NO dejo que continue en este caso
                    If Not SoloComprobar Then
                        'Realmente deberia salir
                      
                        
                        'FALTA####
                        'Deberian existir. Como No existe lo damos de alta
                        
                        cP.Cantidad = -1 * CantidadNecesaria
                        cP.codAlmac = vvCstock.codAlmac
                        cP.codArtic = vvCstock.codArtic
                        cP.codProve = 0
                        cP.Fecha = vvCstock.Fechamov
                        
                        cP.NumAlbar = "PR" & miRsAux!Codigo
                        cP.numLote = cP.NumAlbar
                        If cP.Insertar Then
                            b = True
                            Insertar_sliordpr2lotes cP, 1, CantidadNecesaria
                        End If
                        InsertarMovientosLotesProduccion cL, cP, cP.Cantidad, miRsAux!codArtic
                        
                        
                    End If
                ElseIf Rc = 1 Then
                
                    cad = "NO hay suficiente cantidad"
                    
                    If Not SoloComprobar Then
                        
                        cP.IncrementarCantidad -1 * CantidadNecesaria
                        Insertar_sliordpr2lotes cP, 1, CantidadNecesaria
                        InsertarMovientosLotesProduccion cL, cP, -1 * CantidadNecesaria, miRsAux!codArtic
                    End If
                Else
                    'Ahora si
                    cad = ""
                    b = True
                    
                End If
                If SoloComprobar Then
                        If cad <> "" Then
                            cad = Space(19) & "-- " & vvCstock.codArtic & "  " & Mid(miRsAux!NomArtic & Space(45), 1, 45) & cad
                            AuxPartida = AuxPartida & vbCrLf & cad
                        End If
                
                Else
                    'Estamos ejecutando
                    If b Then
                      For I = 1 To LotesNecesartios.Count
                            cad = LotesNecesartios(I)
                            
                            'ACciones a realizar. Disminnuir cantidad en LOTES
                            NumRegElim = RecuperaValor(cad, 1)
                            CantidadNecesaria = CCur(RecuperaValor(cad, 2))
                            
                            If Not cP.Leer(NumRegElim) Then
                                'MAAAAAAl
                                MsgBox "Error grave partidas/lotes: " & NumRegElim, vbExclamation
                            Else
                                CantidadNecesaria = -1 * CantidadNecesaria
                                cP.IncrementarCantidad CantidadNecesaria
                            
                            
                                'ACtualizar la fila con el numero de lote asignado
                                Insertar_sliordpr2lotes cP, I, Abs(CantidadNecesaria)
                                
                                InsertarMovientosLotesProduccion cL, cP, CantidadNecesaria, miRsAux!codArtic
                                
                                
                            End If  'de cp.leer
                        Next
                    End If  'De B
                End If 'Solo comprobar
            End If  'Tiene lotes MP
            


            
            Set LotesNecesartios = Nothing
        End If 'DE incializa stock
        miRsAux.MoveNext
    Wend
    miRsAux.Close
        
    If AuxPartida <> "" Then
        cad = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", DevNombreSQL(Err_x_Articulo), "T")
        AuxPartida = "-  " & Err_x_Articulo & "   " & cad & AuxPartida
        ErroresEnPartidas = ErroresEnPartidas & AuxPartida
    End If

    If ErroresEnPartidas <> "" Then ErroresEnPartidas = "Error en numeros de lote. " & vbCrLf & String(75, "=") & vbCrLf & ErroresEnPartidas


    AuxPartida = ""
    
        
    cad = "select codartic codarti2,codalmac,sum(sliordpr.cantidad) cantidad,1 factorconversion,numlote from sliordpr where "
    cad = cad & " codigo=" & RecuperaValor(Me.Intercambio, 1) & " group by 1,2"
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not miRsAux.EOF
        b = False
            If InicializarCStock(vvCstock, "E", False) Then   'Las lineas son de netrada
                
                    'AHora veremos los numeros de lote
                    'EL nUMERO DE LOTE NO puede ser NULO
                    CantidadNecesaria = vvCstock.Cantidad  'Para tenerla despues en los lotes
                    cad = "select codalmac,codartic,numlote,cantlote from sliordprlotes where "
                    cad = cad & " codigo=" & RecuperaValor(Me.Intercambio, 1)
                    cad = cad & " AND codartic= '" & miRsAux!codarti2 & "'"
                    
                    CantidadQueLLevo = 0
                    Set RL = New ADODB.Recordset
                    RL.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    While Not RL.EOF
                        CantidadQueLLevo = CantidadQueLLevo + RL!cantlote
                        If Not SoloComprobar Then
                                Set cP = New cPartidas
                                'Vemos si ya existe
                                LoteReal = RL!numLote & " " & Format(txtFecha(0).Text, "yyyy/mm/dd")
                                If cP.LeerDesdeArticulo(miRsAux!codarti2, miRsAux!codAlmac, LoteReal) Then
                                    'Ya existia(por algun motivo)
                                    cP.IncrementarCantidad RL!cantlote
                                    
                                Else
                                    cP.Cantidad = RL!cantlote
                                    cP.codAlmac = vvCstock.codAlmac
                                    cP.codArtic = vvCstock.codArtic
                                    cP.codProve = 0
                                    cP.Fecha = CDate(txtFecha(0).Text)
                                    cP.NumAlbar = "PR" & RecuperaValor(Me.Intercambio, 1)
                                    cP.numLote = LoteReal
                                    If Not cP.Insertar Then
                                        cad = "Error insertando partidas/lotes: " & cP.codArtic
                                        MsgBox cad, vbExclamation
                                    End If
                                    
                                End If
                                
                                'En movimientos lote
                                cL.tipoMov = 1
                                cL.Cantidad = cP.Cantidad
                                cL.codAlmac = cP.codAlmac
                                cL.codArtic = cP.codArtic
                                cL.codarti2 = ""
                                cL.numLote = cP.numLote
                                If Not cL.InsertarLote Then Err.Raise vbObjectError + 513, , "Error insertando en mov lotes: " & cP.codArtic
                                Set cP = Nothing
                                
                                
                                'MAYO 2010
                                'UPDATEO el LOTE que antes era de 4 digitos
                                'a otro que sera los 4 mas la fecha
                                cad = "UPDATE sliordprlotes set numlote=" & DBSet(LoteReal, "T")
                                cad = cad & " where codigo=" & RecuperaValor(Me.Intercambio, 1)
                                cad = cad & " AND codartic= '" & miRsAux!codarti2 & "'"
                                cad = cad & " AND numlote= '" & RL!numLote & "'"
                                conn.Execute cad
                        End If
                        RL.MoveNext
                   Wend
                   RL.Close
                   If CantidadQueLLevo <> CantidadNecesaria Then
                        If Not SoloComprobar Then AuxPartida = AuxPartida & vvCstock.codArtic & ":   necesaria/lotes: " & Format(CantidadNecesaria, FormatoCantidad) & " / " & Format(CantidadQueLLevo, FormatoCantidad) & vbCrLf
                   End If
            End If 'Ini stock
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        If AuxPartida <> "" Then   'Si han habido errores en comprobar cantidades lotes los añado
            AuxPartida = vbCrLf & vbCrLf & "Articulos producidos: " & vbCrLf & AuxPartida
            ErroresEnPartidas = ErroresEnPartidas & AuxPartida
        End If
        b = True
        
        If SoloComprobar Then
            If ErroresEnPartidas <> "" Then
                If InStr(1, ErroresEnPartidas, "depósito 18") > 0 Then
                    ErroresEnPartidas = ErroresEnPartidas & vbCrLf & vbCrLf
                    
                    If UCase(InputBox(ErroresEnPartidas, "Escriba  contraseña para continuar")) <> "ARIADNA" Then b = False
                Else
                    ErroresEnPartidas = ErroresEnPartidas & vbCrLf & vbCrLf & "¿Continuar?"
                    If MsgBox(ErroresEnPartidas, vbQuestion + vbYesNo) = vbNo Then b = False
                End If
            End If
        End If
    
        RealizarProduccionLOTES = b


    
ERealizarProduccionLOTES:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set RL = Nothing
    Set miRsAux = Nothing
    Set vvCstock = Nothing
 
End Function


Private Sub InsertarMovientosLotesProduccion(ByRef cLot As cLotaje, cPar As cPartidas, Cantidad As Currency, ArticuloProduccion As String)

    
    
    cLot.tipoMov = 0  'Salida
    cLot.Cantidad = Abs(Cantidad)
    cLot.codAlmac = cPar.codAlmac
    cLot.codArtic = cPar.codArtic
    cLot.codarti2 = ArticuloProduccion
    cLot.numLote = cPar.numLote

    If Not cLot.InsertarLote Then Err.Raise vbObjectError + 513, , "Error insertando en mov lotes: " & cPar.codArtic
    
End Sub


Private Sub Insertar_sliordpr2lotes(ByRef Par As cPartidas, LineaLote As Integer, Cantidad As Currency)
Dim SQL As String

    
    SQL = "insert into sliordpr2lotes (`codigo`,`codalmac`,`codartic`,`codarti2`,"
    SQL = SQL & "`linea`,`numlote`,`cantlote`) values ( "

    SQL = SQL & RecuperaValor(Intercambio, 1) & ","
    'En misraux tengo los datos que necesito
    SQL = SQL & miRsAux!codAlmac & ",'" & miRsAux!codArtic & "','" & miRsAux!codarti2 & "',"
    SQL = SQL & LineaLote & ",'" & DevNombreSQL(Par.numLote) & "'," & TransformaComasPuntos(CStr(Cantidad)) & ")"
    EjecutaSQL conAri, SQL, True
    
End Sub






'------------------------  LOTES COUPAGE
Private Function RealizarCoupageLOTES(SoloComprobar As Boolean, CantidadMezcla As Currency) As Boolean
Dim ErroresEnPartidas As String
Dim CantidadNecesaria As Currency
Dim AuxPartida As String
Dim Err_x_Articulo As String
Dim MiNumeroLote As String
Dim cP As cPartidas   'Para los numeros de lote
Dim Rc As Byte
Dim vvCstock As cStock
Dim b As Boolean
'Si lleva marca de fin depoisto
Dim RegularizacionDeposito As Currency
Dim cDEP As cDeposito

Dim T1 As Single
Dim FLin As Date
Dim CantidadQueLLevo As Currency
Dim cL As cLotaje
Dim MoviMIentosPosteriores As String

    On Error GoTo ERealizarCUPLOTES

    RealizarCoupageLOTES = False
    

    If Not SoloComprobar Then

        Set cL = New cLotaje
        cL.DetaMov = "CUP"
        cL.Documento = RecuperaValor(Intercambio, 1)
        cL.Fechamov = CDate(Me.txtFecha(1).Text)
        cL.HoraMov = CDate(Me.txtFecha(1).Text & " " & Format(Now, "hh:nn:ss"))
        cL.ProvCliTra = TrabajadorConectado_
        cL.LineaDocu = 0
        cL.SubLinea = 0
    End If
    'Por si acaso no ha puesto numero de lotes. DEBERIA HABERLOS PUESTO
    cad = "select olicoupagelin.codartic,kilos,olicoupagelinlotes.codartic artlote,numlote,cantlote"
    'Juni 2014
    cad = cad & " ,fincuba,deposito"
    cad = cad & " FROM olicoupagelin left join olicoupagelinlotes on"
    cad = cad & " olicoupagelin.codArtic = olicoupagelinlotes.codArtic"
    cad = cad & " and olicoupagelin.codigo= olicoupagelinlotes.codigo WHERE  olicoupagelin.codigo ="
    cad = cad & RecuperaValor(Me.Intercambio, 1) & " ORDER BY codartic"
    miRsAux.Open cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
    b = False
    cad = ""
    ErroresEnPartidas = ""
              
    'Comprobaremos que todos traen el numero de lote puesto y que los
  
    While Not miRsAux.EOF
        If IsNull(miRsAux!artlote) Then
            AuxPartida = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", miRsAux!codArtic, "T")
            cad = cad & miRsAux!codArtic & "   " & AuxPartida
        Else
            If MiNumeroLote <> miRsAux!codArtic Then
                If MiNumeroLote <> "" Then
                    If CantidadQueLLevo <> CantidadNecesaria Then
                        AuxPartida = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", MiNumeroLote, "T")
                        ErroresEnPartidas = ErroresEnPartidas & MiNumeroLote & "   " & AuxPartida & vbCrLf
                    End If
                End If
                MiNumeroLote = miRsAux!codArtic
                CantidadNecesaria = miRsAux!Kilos
                CantidadQueLLevo = miRsAux!cantlote
            Else
                'Dos lineas del mismo articulo
                CantidadQueLLevo = CantidadQueLLevo + miRsAux!cantlote
            End If
            
        End If
        
        
        
        
        miRsAux.MoveNext
        
        
        
        
    Wend
    
    
    'La utlima linea
    If MiNumeroLote <> "" Then
        If CantidadQueLLevo <> CantidadNecesaria Then
            AuxPartida = "S"
            If vParamAplic.QUE_EMPRESA = 4 Then
                  If Opcion = 5 Then AuxPartida = ""  'no compruebo. Viene de molturacion
            End If
            
            If AuxPartida <> "" Then
                AuxPartida = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", MiNumeroLote, "T")
            
            
                ErroresEnPartidas = ErroresEnPartidas & MiNumeroLote & "   " & AuxPartida & vbCrLf
            End If
        End If
    End If
    
    If cad <> "" Or ErroresEnPartidas <> "" Then
        If cad <> "" Then cad = "Lineas articulo sin indicar numero de lote: " & vbCrLf & String(60, "-") & vbCrLf & cad
        If ErroresEnPartidas <> "" Then cad = cad & vbCrLf & vbCrLf & "Articulos lineas sin coincidir cantidades lotes: " & vbCrLf & String(70, "-") & vbCrLf & ErroresEnPartidas
        miRsAux.Close
        MsgBox cad, vbExclamation
        Exit Function
    End If
        
    miRsAux.MoveFirst
    MiNumeroLote = ""
    AuxPartida = ""
    ErroresEnPartidas = ""
    MoviMIentosPosteriores = ""
    Set cP = New cPartidas
    Set vvCstock = New cStock
    Set cDEP = New cDeposito

    While Not miRsAux.EOF
        If Err_x_Articulo <> miRsAux!codArtic Then
            'Han habido errores en el articulo anterior.
            If AuxPartida <> "" Then
                AuxPartida = "-  " & Err_x_Articulo & vbCrLf & AuxPartida & vbCrLf
                ErroresEnPartidas = ErroresEnPartidas & AuxPartida & vbCrLf
            End If
            Err_x_Articulo = miRsAux!codArtic
            AuxPartida = ""
        End If

        RegularizacionDeposito = 0
        b = False
        If InicializarCStockCoupage(vvCstock, "E", True) Then    'Las lineas son de netrada
            
            CantidadNecesaria = CCur(miRsAux!cantlote)
            b = True
            '// NUmeros de LOTE
            cad = ""
            If cP.LeerDesdeArticulo(vvCstock.codArtic, vvCstock.codAlmac, miRsAux!numLote) Then
            
                If cP.Cantidad >= CantidadNecesaria Then
                    'PERFECTO. NO HAgo nada
                    If Val(miRsAux!fincuba) = 1 Then
                        'Regulzarizaremos el deposito
                        If vParamAplic.QUE_EMPRESA <> 4 Then RegularizacionDeposito = cP.Cantidad - CantidadNecesaria
                    End If
                Else
                    If Val(miRsAux!fincuba) = 0 Then
                        'No es fin deposito, no puede seguir
                        cad = "NO hay suficiente cantidad"
                    Else
                        'OK, es fin deposito y habria que "REGULARIZARLO"
                        ' es decir meter una linea para dejar la cantidad del deposito a cero,
                        ' LA PARTIDA a cero
                        ' y una vez acabado el proceso dejar el deposito preparado para llenarlo de nuevo
                        RegularizacionDeposito = cP.Cantidad - CantidadNecesaria
                    End If
                     
                End If
            Else
                'NO existe lote. De momento dejo continuar
                b = False
                cad = "NO hay ningun lote "
                
            End If
    
        
            If SoloComprobar Then
                If cad <> "" Then
                    cad = cad & " (" & miRsAux!numLote & ")"
                    cad = Space(15) & "-- " & vvCstock.codArtic & "  " & cad
                    AuxPartida = AuxPartida & vbCrLf & cad
                End If
            
                If vParamAplic.QUE_EMPRESA = 4 Then
                    cDEP.LeerDatos miRsAux!Deposito, False
                    If cDEP.AvisarMovimientoHcoPosterior(vvCstock.HoraMov) Then MoviMIentosPosteriores = MoviMIentosPosteriores & vbCrLf & "Deposito: " & cDEP.NumDeposito & "  - Movimientos posteriores"
                    Set cDEP = Nothing
                    Set cDEP = New cDeposito
                End If
            Else
                'Por si acaso es FIN deposito
                RegularizacionDeposito = cP.Cantidad - CantidadNecesaria
                
                
                'En la VALL no regularizamos partida
                If vParamAplic.QUE_EMPRESA = 4 Then RegularizacionDeposito = 0
                
                CantidadNecesaria = -1 * CantidadNecesaria  'En negativo
                
                'Incrementamos los kilos
                cDEP.LeerDatos miRsAux!Deposito, False
                
                If vParamAplic.QUE_EMPRESA = 0 Then
                    If vvCstock.HoraMov = "" Then vvCstock.HoraMov = txtFecha(1).Text & " " & Format(Now, "hh:mm:ss")
                End If
                    
                cDEP.InsertarEnHco 2, vvCstock.HoraMov, "CUP" & RecuperaValor(Me.Intercambio, 1), CantidadNecesaria
                cDEP.VariacionKilosDeposito CantidadNecesaria
                
                
                If Not b Then
                    'NO existe. Lo creo
                    cP.Cantidad = CantidadNecesaria
                    cP.codAlmac = vvCstock.codAlmac
                    cP.codArtic = vvCstock.codArtic
                    cP.codProve = 0
                    cP.Fecha = CDate(txtFecha(1).Text)
                    cP.NumAlbar = "CUP" & RecuperaValor(Me.Intercambio, 1)
                    cP.numLote = DBLet(miRsAux!numLote, "T")
                    If cP.numLote Then cP.numLote = cP.NumAlbar
                    
                    If Not cP.Insertar Then
                        cad = "Error insertando partidas/lotes: " & cP.codArtic
                        Err.Raise vbObjectError + 513, , cad
                    End If
        
                Else
                    'Si existe. Lo decremento
                    cP.IncrementarCantidad CantidadNecesaria
                                    
                End If
                'Insertamos en la linea de smoval
                cL.tipoMov = 0
                cL.Cantidad = Abs(CantidadNecesaria)
                cL.codAlmac = vvCstock.codAlmac
                cL.codArtic = vvCstock.codArtic
                cL.numLote = cP.numLote
                cL.InsertarLote
                
                'JUNIO 2014
                'Regulzarizacion FIN DEPOSITO
                If Val(miRsAux!fincuba) = 1 Then
                    
                    If RegularizacionDeposito <> 0 Then
                        Espera 1.25 'PAra que el apunte lo haga un poco despues en la smoval
                        'Regulzarizaremos el deposito
                        
                        
                        Debug.Assert False
                        'Un linea mas en smoval
                        vvCstock.DetaMov = "DEP"
                        
                        
    
                        cL.DetaMov = "DEP"  'FIN DEPOSITO
                        cL.HoraMov = CDate(Me.txtFecha(1).Text & " " & Format(Now, "hh:nn:ss"))
                        cL.tipoMov = 1  '0 entrada 1 salida
                        vvCstock.tipoMov = "E"
                        If RegularizacionDeposito > 0 Then
                            cL.tipoMov = 0
                            vvCstock.tipoMov = "S"
                        End If
                        cL.LineaDocu = cDEP.NumDeposito
                        vvCstock.LineaDocu = cL.LineaDocu
                        cL.Cantidad = Abs(RegularizacionDeposito)
                        cL.InsertarLote
                                                                                           
                        If vParamAplic.QUE_EMPRESA = 4 Then
                            cad = DevuelveDesdeBD(conAri, "count(*)", "proddepositos", "numlote", cP.numLote, "T")
                            If Val(cad) > 1 Then
                                'Tiene mas de un datos en los depositos
                                cad = "MAS"
                            Else
                                cad = ""
                            End If
                        Else
                            cad = ""
                        End If
                                                                                           
                        If cad = "" Then
                            cP.FinPartida   'POndra a cero la cantidad
                       
                        Else
                            cP.IncrementarCantidad CantidadNecesaria
                        End If
                        'Cantidad
                        
                        If vvCstock.Cantidad > 0 Then vvCstock.Importe = (vvCstock.Importe / vvCstock.Cantidad) * cL.Cantidad
                        vvCstock.Cantidad = cL.Cantidad
                        
                        'Septiembre 2021
                        ' Email Ramon. No movemos stock. Si que ajustamos deposito
                        ' vvCstock.ActualizarStock False
                        
                        
                        'Dejamos donde estaba el tipo de movimiento
                        cL.DetaMov = "CUP"
                        vvCstock.DetaMov = "CUP"
                    End If
                    'Ponemos vacios los campos del deposito
                    'Fuera numero de lote y fuera kilos
                    If vvCstock.HoraMov = "" Then vvCstock.HoraMov = Format(Now, "hh:mm")
                    
                    cDEP.QuitarAsignacionDeposito2 0, vvCstock.HoraMov, -vvCstock.Cantidad
                    Espera 0.75
                End If
            End If
        End If 'DE incializa stock
        miRsAux.MoveNext
    Wend
    miRsAux.Close
        
    Set cDEP = Nothing

    If SoloComprobar Then
        RealizarCoupageLOTES = True
        If AuxPartida <> "" Then
            AuxPartida = "-  " & Err_x_Articulo & AuxPartida & vbCrLf
            ErroresEnPartidas = ErroresEnPartidas & AuxPartida
        End If
        If ErroresEnPartidas <> "" Then
            ErroresEnPartidas = ErroresEnPartidas & "¿Continuar?"
            If MsgBox(ErroresEnPartidas, vbExclamation + vbYesNo) = vbNo Then RealizarCoupageLOTES = False
        End If
        If Opcion = 5 Then MoviMIentosPosteriores = ""
        If MoviMIentosPosteriores <> "" Then
            MoviMIentosPosteriores = MoviMIentosPosteriores & vbCrLf & vbCrLf & "¿Continuar de igual modo?"
            If MsgBox(MoviMIentosPosteriores, vbQuestion + vbYesNoCancel) <> vbYes Then RealizarCoupageLOTES = False
        End If
        
        GoTo ERealizarCUPLOTES 'para k haga los =nothing
    End If

        

    AuxPartida = ""
    
        

    'AHora comprobamos los stcosk de las entraddas , de las lineas
    cad = TransformaComasPuntos(CStr(CantidadMezcla))
    cad = "select codartic," & cad & " kilos,numlote,codalmac,deposito from olicoupage where codigo=" & RecuperaValor(Me.Intercambio, 1)
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'SOLO HAY una linea
    If Not miRsAux.EOF Then
        b = False
        If InicializarCStockCoupage(vvCstock, "E", True) Then    'Las lineas son de netrada
                
                                
                'AHora veremos los numeros de lote
                'EL nUMERO DE LOTE NO puede ser NULO
                CantidadNecesaria = vvCstock.Cantidad  'Para tenerla despues en los lotes
                
                
                                                        'Vemos si ya existe
                If cP.LeerDesdeArticulo(miRsAux!codArtic, miRsAux!codAlmac, miRsAux!numLote) Then
                    'Ya existia(por algun motivo)
                    cP.IncrementarCantidad CantidadNecesaria
                    
                Else
                    cP.Cantidad = CantidadNecesaria
                    cP.codAlmac = miRsAux!codAlmac
                    cP.codArtic = vvCstock.codArtic
                    cP.codProve = 0
                    cP.Fecha = CDate(txtFecha(1).Text)
                    cP.NumAlbar = "CUP" & RecuperaValor(Me.Intercambio, 1)
                    cP.numLote = miRsAux!numLote
                    If Not cP.Insertar Then Err.Raise vbObjectError + 513, , cad
                    
                End If
                
                'Insertamos en la linea de smoval
                cL.tipoMov = 1
                cL.Cantidad = Abs(CantidadNecesaria)
                cL.codAlmac = vvCstock.codAlmac
                cL.codArtic = vvCstock.codArtic
                cL.numLote = cP.numLote
                cL.InsertarLote
                
                b = True
                
                Set cDEP = New cDeposito
                'Para que no de error insertando en hco
                T1 = Timer
                If Not cDEP.LeerDatos(miRsAux!Deposito, False) Then b = False
                
                AuxPartida = DevuelveDesdeBD(conAri, "factorconversion", "sartic", "codartic", miRsAux!codArtic, "T")
                CantidadNecesaria = CCur(AuxPartida)
                
                
                cDEP.Kilos = cL.Cantidad
                cDEP.numLote = cP.numLote
                cDEP.idPartida = cP.idPartida
                Espera 0.5
                
                FLin = DateAdd("s", 1, vvCstock.HoraMov)
                cDEP.InsertarEnDeposito2 1, FLin, "CUP" & RecuperaValor(Me.Intercambio, 1)
                
                
                Espera 0.8
        End If
    End If
        
    RealizarCoupageLOTES = b


    
ERealizarCUPLOTES:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set cP = Nothing
    Set miRsAux = Nothing
    Set vvCstock = Nothing
    Set cDEP = Nothing
End Function






Private Function ActualizarPrecio() As Boolean
Dim b As Boolean
Dim CantidadTotalAProducir As Currency 'Cuatro decimales
Dim PrecioTotal As Currency
Dim C As Currency
Dim Articulo As String
    
    
    Set miRsAux = New ADODB.Recordset
    
    'Vemos si la referencia es de esas
    cad = "select olicoupage.codartic from olicoupage where codigo=" & RecuperaValor(Me.Intercambio, 1)
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Articulo = miRsAux!codArtic
    miRsAux.Close
    
    'Estos articulos me los indico ramoon en un Email
    '003500411513  003500421513 003900431513
    b = (Articulo = "003500411513") Or (Articulo = "003500421513") Or (Articulo = "003900431513")
    If Not b Then
        Set miRsAux = Nothing
        Exit Function
    End If
    
    
    'OK.Calculo el precio
    
    
    
    
    'Los mezclantes
    
    cad = "select olicoupagelin.*,preciouc, preciomp from olicoupagelin,sartic where olicoupagelin.codartic=sartic.codartic and "
    cad = cad & "  codigo = " & RecuperaValor(Me.Intercambio, 1)
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    b = False
    
    CantidadTotalAProducir = 0
    PrecioTotal = 0
    While Not miRsAux.EOF
        C = DBLet(miRsAux!PrecioUC, "N")
        C = miRsAux!Kilos * C
        PrecioTotal = PrecioTotal + C
        CantidadTotalAProducir = CantidadTotalAProducir + miRsAux!Kilos
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    'Si no produce nada nos piramos
    If CantidadTotalAProducir = 0 Then Exit Function
    
    PrecioTotal = Round(PrecioTotal / CantidadTotalAProducir, 4)
    
    cad = "select preciouc,ultfecco from sartic where codartic='" & Articulo & "'"
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    b = False 'Tiene que actualizar
    If IsNull(miRsAux!ultfecco) Then
        b = True
    Else
        If CDate(miRsAux!ultfecco) < CDate(txtFecha(1).Text) Then
            'OK
            'Veremos los importes
            C = DBLet(miRsAux!PrecioUC)
                                            'Ha cambiado
            If C <> PrecioTotal Then b = True
        End If
    End If
    miRsAux.Close
    
    
  
    If b Then
        'OK. Hay que actualizar los importes
        lbFec(1).Caption = "Act. precio"
        lbFec(1).Refresh
        Espera 0.3
        ActualizarPrecioCosteArticulo PrecioTotal, Articulo
    End If
    Set miRsAux = Nothing
End Function




Private Sub ActualizarPrecioCosteArticulo(ByRef Pre As Currency, ByRef codart As String)


On Error GoTo EActualizarPrecioCosteArt


    cad = "UPDATE sartic set PrecioUC = " & TransformaComasPuntos(CStr(Pre))
    cad = cad & ", ultfecco = '" & Format(txtFecha(1).Text, FormatoFecha) & "'"
    cad = cad & " WHERE codartic = '" & codart & "'"
    
    'Ejecutar
    conn.Execute cad
    Espera 0.2
    
    
    
    
    'Para que se actualice bien
    CommitConexion
    Espera 0.1
    
    'AHora va el meollo. Si el articulo es materia prima, todos los artiuclos
    'de venta en los que el entra como materia prima deben sera actualizados
    cad = "select sartic.codartic,nomartic,codunida from sarti1,sartic where sarti1.codartic = sartic.codartic"
    cad = cad & " AND codarti1 = '" & codart & "'"
    miRsAux.Open cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        NumRegElim = 0
        While Not miRsAux.EOF
            NumRegElim = NumRegElim + 1
            miRsAux.MoveNext
        Wend
        miRsAux.MoveFirst
        Pre = 1
        While Not miRsAux.EOF
            lbFec(1).Caption = "UPC " & CInt(Pre) & " de " & NumRegElim
            lbFec(1).Refresh
            ActualizaUPCArticuloCabecera miRsAux!codArtic, CInt(miRsAux!CodUnida)
            Pre = CInt(Pre) + 1
            miRsAux.MoveNext
            If (CInt(Pre) Mod 15) = 0 Then
                Me.Refresh
                DoEvents
            End If
        Wend
     
    End If
    miRsAux.Close
    
EActualizarPrecioCosteArt:
    
    If Err.Number <> 0 Then MuestraError Err.Number, "Actualizando precio coste"
    Screen.MousePointer = vbDefault
    Set miRsAux = Nothing

End Sub




Private Sub ActualizaUPCArticuloCabecera(ByRef C As String, CodUnida As Integer)
Dim Aux As String
Dim RS As ADODB.Recordset
Dim Im0 As Currency
Dim Im1 As Currency

    On Error GoTo eActualizaUPCArticuloCabecera
    Set RS = New ADODB.Recordset
    Aux = "SELECT sarti1.codartic, numlinea, sarti1.codarti1,sartic.nomartic, sarti1.Cantidad ,"
    Aux = Aux & "sartic.preciove , sartic.precioUC, FactorConversion"
    Aux = Aux & " FROM   sarti1 INNER JOIN sartic ON sarti1.codarti1 = sartic.codArtic where sarti1.codartic='"
    Aux = Aux & C & "' ORDER BY sarti1.numlinea"
    RS.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Im1 = 0
    Aux = ""
    While Not RS.EOF
        Aux = RS!NomArtic
        Im0 = DBLet(RS!FactorConversion, "N")  'del articulo de la linea

        'COSTE
        Im0 = DBLet(RS!Cantidad, "N") * Im0
        Im0 = Im0 * DBLet(RS!PrecioUC, "N")
        Im1 = Im1 + Im0
        
        RS.MoveNext
    Wend

    RS.Close
    
    'El formato
    Aux = "Select sum(importe) from sunilin where codunida=" & CodUnida
    RS.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Im0 = 0
    If Not RS.EOF Then Im0 = DBLet(RS.Fields(0), "N")
    RS.Close

    'Redondeamos (al igual que en el mantenimiento de articulos) a 3 antes de sumar el formato
    Im1 = Round(Im1, 3)

    Im1 = Im1 + Im0
    Im1 = Round2(Im1, 3)
    
    'UPDATEAMOS
    Aux = "UPDATE sartic set PrecioUC = " & TransformaComasPuntos(CStr(Im1))
    Aux = Aux & ", ultfecco = '" & Format(txtFecha(1).Text, FormatoFecha) & "'"
    Aux = Aux & " WHERE codartic = '" & C & "'"
    conn.Execute Aux
    
eActualizaUPCArticuloCabecera:
    If Err.Number <> 0 Then MuestraError Err.Number, Aux
    Set RS = Nothing
End Sub




Private Sub txtHora_GotFocus(Index As Integer)
     ConseguirFoco txtHora(Index), 3
End Sub

Private Sub txtHora_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtHora_LostFocus(Index As Integer)
Dim cad As String
    txtHora(Index).Text = Trim(txtHora(Index).Text)
    If txtHora(Index).Text = "" Then Exit Sub
    cad = Replace(txtHora(Index).Text, ".", ":")
    If Not EsHoraOK(cad) Then
        MsgBox "Error en campo hora", vbExclamation
        txtHora(Index).Text = ""
        PonerFoco txtHora(Index)
    Else
        txtHora(Index).Text = cad
    End If
End Sub

Private Sub txtMeses_GotFocus()
    ConseguirFoco txtMeses, 3
End Sub

Private Sub txtMeses_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub


Private Sub txtMeses_LostFocus()
    txtMeses.Text = Trim(txtMeses.Text)
    If txtMeses.Text = "" Then Exit Sub
    
    If Not IsNumeric(txtMeses.Text) Then
        MsgBox "Campo numerico", vbExclamation
        txtMeses.Text = "18"
        PonerFoco txtMeses
    End If
    
    txtMeses.Text = Abs(Val(txtMeses.Text))
    
        
        
End Sub

Private Sub txtNumeroDec_GotFocus(Index As Integer)
    ConseguirFoco txtNumeroDec(Index), 3
End Sub


Private Sub txtNumeroDec_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtNumeroDec_LostFocus(Index As Integer)
    txtNumeroDec(Index).Text = Trim(txtNumeroDec(Index).Text)
    If txtNumeroDec(Index).Text = "" Then Exit Sub
    
    
    If Not PonerFormatoDecimal(txtNumeroDec(Index), 3) Then txtNumeroDec(Index).Text = ""
    

End Sub



Private Sub CargaComobosTrasiegos(Inicio As Byte, Fin As Byte)

    Set miRsAux = New ADODB.Recordset
    For I = Inicio To Fin
        cboDeposito(I).Clear
        
        If I = 0 Or I = 2 Or I = 4 Then
            cad = "SELECT proddepositos.numdeposito, spartidas.codartic, sartic.nomartic, spartidas.numlote, kilos vlitros"
            '(kilos * factorconversion) vlitros"
            cad = cad & " FROM  proddepositos left join spartidas on spartidas.numlote=proddepositos.numlote"
            cad = cad & " inner join sartic on spartidas.codartic=sartic.codartic AND sartic.factorconversion<1"
            cad = cad & " Where Not spartidas.numLote Is Null"
            
            cad = cad & " ORDER BY numdeposito"
    
        Else

            cad = "select * from proddepositos where numlote is null"
            If vParamAplic.QUE_EMPRESA = 0 Then
                If I <> 2 Then cad = cad & " AND  numdeposito < 107"   '108 y 109 vinagreta
            End If
        End If
        
        miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            If I = 0 Or I = 2 Or I = 4 Then
                cad = Format(miRsAux!NumDeposito, "00") & "  "
                If vParamAplic.QUE_EMPRESA <> 4 Then cad = cad & "L"
                cad = cad & Mid(miRsAux!numLote & "       ", 1, 9) & " " & miRsAux!NomArtic & " (" & Format(miRsAux!vlitros, FormatoCantidad) & ")"
            Else
                cad = "Deposito " & miRsAux!NumDeposito
            End If
            cboDeposito(I).AddItem cad
            cboDeposito(I).ItemData(cboDeposito(I).NewIndex) = miRsAux!NumDeposito
            
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    Next I
    Set miRsAux = Nothing
End Sub

'LA VALL
'  Pondremos los destinos TODOS, y los destinos los que esten llenos.
'  Por programa comprobaremos
Private Sub CargaComobosLaVall()

    Set miRsAux = New ADODB.Recordset
      For I = 5 To 6
        cboDeposito(I).Clear
        
        
            cad = "SELECT proddepositos.numdeposito, spartidas.codartic, sartic.nomartic, spartidas.numlote, kilos vlitros"
            '(kilos * factorconversion) vlitros"
            cad = cad & " FROM  proddepositos left join spartidas on spartidas.numlote=proddepositos.numlote"
            cad = cad & " left join sartic on spartidas.codartic=sartic.codartic AND sartic.factorconversion<1"
            
            If I = 5 Then cad = cad & " Where Not spartidas.numLote Is Null"
            
            cad = cad & " ORDER BY numdeposito"
    
        
        
            miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                
                cad = Format(miRsAux!NumDeposito, "00") & "  "
                cad = "Deposito " & cad
                If Not IsNull(miRsAux!numLote) Then
                
                    cad = cad & Mid(miRsAux!numLote & "       ", 1, 9) & " " & miRsAux!NomArtic & " (" & Format(miRsAux!vlitros, FormatoCantidad) & ")"
                End If
                
                
                cboDeposito(I).AddItem cad
                cboDeposito(I).ItemData(cboDeposito(I).NewIndex) = miRsAux!NumDeposito
                
                miRsAux.MoveNext
            Wend
            miRsAux.Close
    Next I
    Set miRsAux = Nothing
End Sub








'Este trozo esta copiado de proceso produccion
'De momento solo entra aqui para materia prima
Private Sub RegularizarFinLote_Partida_DEPOS(ByRef cDEP As cDeposito, FechaHoraMov As Date)
Dim cPar As cPartidas

Dim cLot As cLotaje
Dim vvCstock2 As cStock
Dim Aux As String
Dim Donde As String
Dim Cantidad As Currency

    On Error GoTo eRegularizarFinLote_Partida

    
    
    Set cPar = New cPartidas
    Set cLot = New cLotaje
  '  Set vvCstock = New cStock
    
    Donde = "Leyendo clases"
    
    'select * from spartidas,sartic where spartidas.codartic=sartic.codartic and sartic.factorconversion<1 and numlote in (select numlote from proddepositos)
    Aux = "spartidas.codartic=sartic.codartic and sartic.factorconversion<1 and numlote"
    Aux = DevuelveDesdeBD(conAri, "id", "spartidas,sartic", Aux, cDEP.numLote, "T")
    If Aux = "" Then Err.Raise 513, , "No se encuentra la partida"
    cPar.Leer CLng(Aux)
    
    
    
        
    Set cLot = New cLotaje
 '   Set vvCstock = New cStock
        
   
    
    
    'Un linea mas en smoval
  
 '   vvCstock.DetaMov = "DEP"
    '0=Salida, 1=Entrada
    If cPar.Cantidad >= 0 Then
  '      vvCstock.tipoMov = "S"
        cLot.tipoMov = 0
    Else
  '      vvCstock.tipoMov = "E"
        cLot.tipoMov = 1
    End If
  '  vvCstock.Cantidad = Abs(cPar.Cantidad)
  '  vvCstock.Trabajador = TrabajadorConectado_
  '
  '  vvCstock.Fechamov = Format(FechaHoraMov, "dd/mm/yyyy")
  '  vvCstock.HoraMov = FechaHoraMov
  '  vvCstock.codAlmac = cPar.codAlmac
  '  vvCstock.codArtic = cPar.codArtic
  '  vvCstock.Importe = 0
  '  vvCstock.Documento = "VACIA " & Format(cPar.idPartida, "0000000")
    
    cLot.codAlmac = cPar.codAlmac 'vvCstock.codAlmac
    cLot.codArtic = cPar.codArtic 'vvCstock.codArtic
    cLot.DetaMov = "DEP"   'vvCstock.DetaMov
    cLot.Fechamov = Format(FechaHoraMov, "dd/mm/yyyy")   'vvCstock.Fechamov
    cLot.HoraMov = FechaHoraMov 'vvCstock.HoraMov
    cLot.numLote = cPar.numLote
    
    cLot.Cantidad = Abs(cPar.Cantidad)
    cLot.LineaDocu = cDEP.NumDeposito
    cLot.Documento = "VACIA " & Format(cPar.idPartida, "0000000") 'vvCstock.Documento
    
    cLot.InsertarLote


    'Septiembre 2021
    ' Mensaje de  Ramon.
    ' El fallo de stock que tenia es pq realmente, aunque el deposito lo vacia, a nivel de stock NO lo mueve
    ' por lo tanto, tema de partidas (trazbilidad) SI que ajusta , a nivel de smoval no
    ' vvCstock.ActualizarStock False
    
    
    
    cPar.AjustarFinPartida
    
    
                        
    
eRegularizarFinLote_Partida:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set cPar = Nothing
    Set cLot = Nothing
    'Set vvCstock = Nothing
    
End Sub




'----------------------------------------------------------------------------------
Private Function TodoCorrectoParaFiltrar(C1 As cDeposito, C2 As cDeposito) As Boolean
Dim CC As CTiposMov
Dim FechaHora2 As Date
Dim TextoFiltrado As String
Dim Cantidad As Currency
Dim Cadena2 As String

    TodoCorrectoParaFiltrar = False
    cad = ""
    If txtFecha(4).Text = "" Then cad = "-Fecha"
    If Me.txtHora(4).Text = "" Then cad = cad & vbCrLf & "  -Hora"
    If Me.txtNumeroDec(4).Text = "" Then cad = cad & vbCrLf & "   -Cantidad a mover"
    
    For I = 0 To 2
        If Me.txtArtFiltrado(I).Text <> "" And Me.txtNumeroDec(I + 1).Text <> "" And Me.txtLote(I).Text = "" Then
            cad = cad & vbCrLf & "  -Lote para " & txtArtFiltrado(I).Text
        Else
            If Me.txtNumeroDec(I + 1).Text <> "" Then TextoFiltrado = "S"
        End If
    Next
    
    If TextoFiltrado = "" Then
        If Me.optLaVall(0).Value Then cad = cad & vbCrLf & "   -Selecciona filtrado y no indica filtros a utilizar"
    End If
    
    'No puede haber TRASIEGO sobre mismo deposito
    If Me.optLaVall(1).Value Then
        If C1.NumDeposito = C2.NumDeposito Then cad = cad & vbCrLf & "   -No puede realizar trasiego sobre el mismo deposito"
    End If
            
    If C1.NumDeposito <> C2.NumDeposito Then
        'Distintos depositos
        'Solo podriamos pasar si el destino esta vacio, o el destino tiene el mismo numero de lote.
        'Si no, a coupar
        If C2.idPartida > 0 Then
            'Tiene algo
            If C1.numLote <> C2.numLote Then
                'deber realizar coupage
                cad = cad & vbCrLf & "-Lotes distintos"
            End If
        End If
    End If
    
    
    If cad <> "" Then
        cad = "Campos requeridos: " & vbCrLf & cad
        MsgBox cad, vbExclamation
        Exit Function
    End If
        
    
    
    If CDate(txtFecha(4).Text) < vParamAplic.FechaActiva Then
        MsgBox "Menor que fecha activa", vbExclamation
        Exit Function
    End If
    
    
    cad = ""
    If C1.AvisarMovimientoHcoPosterior(CDate(Me.txtFecha(4).Text & " " & Me.txtHora(4).Text)) Then cad = cad & vbCrLf & " -Origen"
    If C2.AvisarMovimientoHcoPosterior(CDate(Me.txtFecha(4).Text & " " & Me.txtHora(4).Text)) Then cad = cad & vbCrLf & " -Destino"
    If cad <> "" Then
        cad = "Hay movimientos posteriores en el depósito: " & cad & vbCrLf & "¿Continuar?"
        If MsgBox(cad, vbExclamation + vbYesNoCancel) <> vbYes Then Exit Function
    End If
    
    
    'cantidad a mover
    cad = ""
    Cantidad = ImporteFormateado(txtNumeroDec(4).Text)
    If Cantidad > C1.Kilos Then cad = cad & "    -No puede traspasar mas cantidad de la disponible    "
    If C1.NumDeposito <> C2.NumDeposito Then
        If Cantidad + C2.Kilos > C2.Capacidad Then cad = cad & vbCrLf & "    -Aceite en destino mayor que capacidad "
    End If
    
    'Fitrlado sobre el mismo deposito. Dbe ser el total
    If Me.optLaVall(0).Value Then
        If C1.NumDeposito = C2.NumDeposito Then
            If Cantidad <> C1.Kilos Then
                cad = cad & vbCrLf & "   -Sobre un mismo deposito debe filtrar todos los kilos."
            End If
        End If
    End If
            
    
    If cad <> "" Then
        cad = "Error en KILOS" & vbCrLf & vbCrLf & cad
        MsgBox cad, vbExclamation
        Exit Function
    End If
    
    TrabajadorConectado_ = Val(PonerTrabajadorConectado(vUsu.Login))
    cad = "Va a realizar el " & IIf(Me.optLaVall(0).Value, "filtrado", "trasiego") & " : " & vbCrLf & "Origen: " & cboDeposito(5).Text
    cad = cad & vbCrLf & "Destino: " & cboDeposito(6).Text & vbCrLf & vbCrLf
    cad = cad & Space(20) & "KILOS: " & Me.txtNumeroDec(4).Text & vbCrLf & vbCrLf
    'Si hay gasto de productos en filtrado
    For I = 1 To 3
        If Me.txtNumeroDec(I).Text <> "" Then cad = cad & "      - " & Me.txtArtFiltrado(2 + I).Text & ": " & txtNumeroDec(I).Text & "  Kilos" & vbCrLf
    Next I

   
    If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
    
    TodoCorrectoParaFiltrar = True
End Function

















'*****************************************************************************************+



'-------------------------------------------------------------------------
' Realizar el coupage
'
Private Function RealizarCoupageVALL(SoloComprobar As Boolean) As Boolean
Dim vCStock As cStock
Dim b As Boolean
Dim CantidadTotalAProducir As Currency 'Cuatro decimales

Dim DepositoDestinoCopuage As Integer

    'ACciones a realizar
    'Comprobar stock sublineas, ya que es la que van a disminuir la cantidad
    'Damos de alta en stock (y smoval) las lienas ppales
    'Damos de baja   "        "        las sublineas
    RealizarCoupageVALL = False
    Set miRsAux = New ADODB.Recordset
    Set vCStock = New cStock
    
    
    
    
    
    If Not SoloComprobar Then conn.BeginTrans

        
        
    cad = "Select deposito from olicoupage where codigo=" & RecuperaValor(Me.Intercambio, 1)
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'NO PUEDE SER EOF
    DepositoDestinoCopuage = miRsAux!Deposito
    miRsAux.Close
    
    'Los mezclantes
    'Como no lleva factor conversion. Necesito los precios para los calculos de importes
    cad = "select olicoupagelin.*,preciouc, preciomp from olicoupagelin,sartic where olicoupagelin.codartic=sartic.codartic and "
    cad = cad & "  codigo = " & RecuperaValor(Me.Intercambio, 1)
    'cad = "select * from olicoupagelin where codigo = " & RecuperaValor(Me.Intercambio, 1)
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    b = False
    
    CantidadTotalAProducir = 0
    While Not miRsAux.EOF
        b = False
        If InicializarCStockCoupage(vCStock, "S", False) Then
            
            If vCStock.MueveStock Then
                If SoloComprobar Then
                    b = vCStock.MoverStock(False)
                Else
                    'Estamos ejecutando la actualizacion
                    '---------------------------------------------
                    'si hay control de stock para el articulo actualizar en salmac e insertar en smoval
                    'en actualizar stock comprobamos si el articulo tiene control de stock
                    b = vCStock.ActualizarStock(False)
                End If
            Else
                b = True
            End If
            CantidadTotalAProducir = CantidadTotalAProducir + miRsAux!Kilos
        End If
        
        If Not b Then
            While Not miRsAux.EOF
                miRsAux.MoveNext  'para que no siga
            Wend
        Else
            'Al siguiente
            miRsAux.MoveNext
        End If
    Wend
    miRsAux.Close
    
    
    If Not b Then
        Set miRsAux = Nothing
        Set vCStock = Nothing
        If Not SoloComprobar Then conn.RollbackTrans
        Exit Function 'Si no puede inicializar los stocks, de las sublineas salimos
    End If
    
    
    'SSi ha ido bien comprobamos los LOTES
    If Not RealizarCoupageLOTESVALL(SoloComprobar, CantidadTotalAProducir, DepositoDestinoCopuage) Then
    
            Set miRsAux = Nothing
            Set vCStock = Nothing
            If Not SoloComprobar Then conn.RollbackTrans
            Exit Function 'Si no puede inicializar los stocks, de las sublineas salimos
    End If
    Set miRsAux = Nothing
    Set miRsAux = New ADODB.Recordset
    
    
    
    
    
    
    'AHora comprobamos los stcosk de las entraddas , de las lineas
    cad = TransformaComasPuntos(CStr(CantidadTotalAProducir))
    
    cad = "select olicoupage.codartic," & cad & " kilos,preciouc from olicoupage,sartic where"
    cad = cad & " olicoupage.codartic=sartic.codartic and codigo=" & RecuperaValor(Me.Intercambio, 1)
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    b = False
    
    While Not miRsAux.EOF
        b = False
        If InicializarCStockCoupage(vCStock, "E", False) Then   'Las lineas son de netrada
        
            If vCStock.MueveStock Then
                If SoloComprobar Then
                    'B = vCStock.MoverStock(False)
                    b = True
                    
                    
                    
                    
                Else
                    b = vCStock.ActualizarStock(False)
                End If
            Else
                b = True
            End If
        End If
        
        If Not b Then
            While Not miRsAux.EOF
                miRsAux.MoveNext  'para que no siga
            Wend
        Else
            'Al siguiente
            miRsAux.MoveNext
        End If
    Wend
    miRsAux.Close
    
    
    If Not b Then
        Set miRsAux = Nothing
        Set vCStock = Nothing
        If Not SoloComprobar Then conn.RollbackTrans
        Exit Function 'Si no puede inicializar los stocks, de las sublineas salimos
    End If
    
    
    
    
    'Acutailizaremos algnas cosas como la fecha de baja
    If Not SoloComprobar Then
        conn.CommitTrans
        cad = "UPDATE olicoupage set YaCreado = 1"
        cad = cad & " WHERE  codigo=" & RecuperaValor(Me.Intercambio, 1)
        conn.Execute cad
    End If
    
    
        
    
    
    RealizarCoupageVALL = True
    
    Set miRsAux = Nothing
    Set vCStock = Nothing
    
    
End Function






'------------------------  LOTES COUPAGE
Private Function RealizarCoupageLOTESVALL(SoloComprobar As Boolean, CantidadMezcla As Currency, DepositoDestino As Integer) As Boolean
Dim ErroresEnPartidas As String
Dim CantidadNecesaria As Currency
Dim AuxPartida As String
Dim Err_x_Articulo As String
Dim MiNumeroLote As String
Dim cP As cPartidas   'Para los numeros de lote
Dim Rc As Byte
Dim vvCstock As cStock
Dim b As Boolean
Dim cDe As cDeposito

Dim T1 As Single
Dim FLin As Date
Dim CantidadQueLLevo As Currency
Dim cL As cLotaje
Dim MoviMIentosPosteriores As String
Dim TextoParaHcoDepositosLaVall As String
Dim CuantoHabiaDepositoDestino As Currency
Dim cDestino As cDeposito

Dim Segundos As Integer

    On Error GoTo ERealizarCUPLOTES

    RealizarCoupageLOTESVALL = False
    

    If Not SoloComprobar Then

        Set cL = New cLotaje
        cL.DetaMov = "CUP"
        cL.Documento = RecuperaValor(Intercambio, 1)
        cL.Fechamov = CDate(Me.txtFecha(1).Text)
        cL.HoraMov = CDate(Me.txtFecha(1).Text & " " & Format(Now, "hh:nn:ss"))
        cL.ProvCliTra = TrabajadorConectado_
        cL.LineaDocu = 0
        cL.SubLinea = 0
    End If
    'Por si acaso no ha puesto numero de lotes. DEBERIA HABERLOS PUESTO
    cad = "select olicoupagelin.codartic,kilos,olicoupagelinlotes.codartic artlote,numlote,cantlote"
    cad = cad & " ,fincuba,deposito"
    'Para que ordene primeropor los depositos que no es el destino. El order by 8 es pq el campo es el octavo
    cad = cad & " , if(deposito=" & DepositoDestino & ",1,0)"
    cad = cad & " FROM olicoupagelin left join olicoupagelinlotes on"
    cad = cad & " olicoupagelin.codArtic = olicoupagelinlotes.codArtic"
    cad = cad & " and olicoupagelin.codigo= olicoupagelinlotes.codigo WHERE  olicoupagelin.codigo ="
    cad = cad & RecuperaValor(Me.Intercambio, 1) & " ORDER BY 8,codartic"
    miRsAux.Open cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
    b = False
    cad = ""
    ErroresEnPartidas = ""
    TextoParaHcoDepositosLaVall = ""
    CuantoHabiaDepositoDestino = 0
    Set cDestino = New cDeposito
    'Comprobaremos que todos traen el numero de lote puesto y que los
  
    While Not miRsAux.EOF
        If IsNull(miRsAux!artlote) Then
            AuxPartida = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", miRsAux!codArtic, "T")
            cad = cad & miRsAux!codArtic & "   " & AuxPartida
        Else
            If MiNumeroLote <> miRsAux!codArtic Then
                If MiNumeroLote <> "" Then
                    If CantidadQueLLevo <> CantidadNecesaria Then
                        AuxPartida = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", MiNumeroLote, "T")
                        ErroresEnPartidas = ErroresEnPartidas & MiNumeroLote & "   " & AuxPartida & vbCrLf
                    End If
                End If
                MiNumeroLote = miRsAux!codArtic
                CantidadNecesaria = miRsAux!Kilos
                CantidadQueLLevo = miRsAux!cantlote
            Else
                'Dos lineas del mismo articulo
                CantidadQueLLevo = CantidadQueLLevo + miRsAux!cantlote
            End If
            
            If miRsAux!Deposito <> DepositoDestino Then
                TextoParaHcoDepositosLaVall = miRsAux!numLote
            End If
        End If
        
        If miRsAux!Deposito = DepositoDestino Then
            If Not SoloComprobar Then
                CuantoHabiaDepositoDestino = miRsAux!cantlote
                cDestino.LeerDatos DepositoDestino, False
            End If
        End If
        miRsAux.MoveNext
        
        
        
        
    Wend
    If cDestino.NumDeposito = 0 Then
        
        cDestino.LeerDatos DepositoDestino, False
    End If
    
    'La utlima linea
    If MiNumeroLote <> "" Then
        If CantidadQueLLevo <> CantidadNecesaria Then
            AuxPartida = "S"
            If vParamAplic.QUE_EMPRESA = 4 Then
                  If Opcion = 5 Then AuxPartida = ""  'no compruebo. Viene de molturacion
            End If
            
            If AuxPartida <> "" Then
                AuxPartida = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", MiNumeroLote, "T")
            
            
                ErroresEnPartidas = ErroresEnPartidas & MiNumeroLote & "   " & AuxPartida & vbCrLf
            End If
        End If
    End If
    
    If cad <> "" Or ErroresEnPartidas <> "" Then
        If cad <> "" Then cad = "Lineas articulo sin indicar numero de lote: " & vbCrLf & String(60, "-") & vbCrLf & cad
        If ErroresEnPartidas <> "" Then cad = cad & vbCrLf & vbCrLf & "Articulos lineas sin coincidir cantidades lotes: " & vbCrLf & String(70, "-") & vbCrLf & ErroresEnPartidas
        miRsAux.Close
        MsgBox cad, vbExclamation
        Exit Function
    End If
        
    miRsAux.MoveFirst
    MiNumeroLote = ""
    AuxPartida = ""
    ErroresEnPartidas = ""
    MoviMIentosPosteriores = ""
    Set cP = New cPartidas
    Set vvCstock = New cStock
    Set cDe = New cDeposito
    Segundos = 0
    While Not miRsAux.EOF
        If Err_x_Articulo <> miRsAux!codArtic Then
            'Han habido errores en el articulo anterior.
            If AuxPartida <> "" Then
                AuxPartida = "-  " & Err_x_Articulo & vbCrLf & AuxPartida & vbCrLf
                ErroresEnPartidas = ErroresEnPartidas & AuxPartida & vbCrLf
            End If
            Err_x_Articulo = miRsAux!codArtic
            AuxPartida = ""
        End If

        b = False
        Segundos = Segundos + 1
        If InicializarCStockCoupage(vvCstock, "E", True) Then    'Las lineas son de netrada
    
            CantidadNecesaria = CCur(miRsAux!cantlote)
            b = True
            '// NUmeros de LOTE
            cad = ""
            If cP.LeerDesdeArticulo(vvCstock.codArtic, vvCstock.codAlmac, miRsAux!numLote) Then
            
            Else
                'NO existe lote. De momento dejo continuar
                b = False
                cad = "NO hay ningun lote "
                
            End If
    
        
            If SoloComprobar Then
                If cad <> "" Then
                    cad = cad & " (" & miRsAux!numLote & ")"
                    cad = Space(15) & "-- " & vvCstock.codArtic & "  " & cad
                    AuxPartida = AuxPartida & vbCrLf & cad
                End If
            
                If vParamAplic.QUE_EMPRESA = 4 Then
                    cDe.LeerDatos miRsAux!Deposito, False
                    If cDe.AvisarMovimientoHcoPosterior(vvCstock.HoraMov) Then MoviMIentosPosteriores = MoviMIentosPosteriores & vbCrLf & "Deposito: " & cDe.NumDeposito & "  - Movimientos posteriores"
                    
                    
                    'Si hay mas kilos que los que tiene el deposito
                    If CantidadNecesaria > cDe.Kilos Then
                        cad = "DEP: " & cDe.NumDeposito
                        cad = Space(15) & "-- Mas kilos que los que quedan en deposito   " & cad
                        AuxPartida = AuxPartida & vbCrLf & cad
                    End If
                    If Abs(CantidadNecesaria - cDe.Kilos) < 5 And Val(miRsAux!fincuba) = 0 Then
                        cad = "DEP: " & cDe.NumDeposito
                        cad = Space(15) & "-- parece FIN deposito y no esta marcado como tal  " & cad
                        AuxPartida = AuxPartida & vbCrLf & cad
                    End If
                                    
                    If Val(miRsAux!fincuba) = 1 And Abs(cDe.Kilos - CantidadNecesaria) > 5 Then
                        cad = "DEP: " & cDe.NumDeposito
                        cad = Space(15) & "-- NO es FIN deposito y si que esta marcado como tal  " & cad
                        AuxPartida = AuxPartida & vbCrLf & cad
                    End If
                    Set cDe = Nothing
                    Set cDe = New cDeposito
                End If
            Else
                CantidadNecesaria = -1 * CantidadNecesaria  'En negativo
                
                'Incrementamos los kilos
                cDe.LeerDatos miRsAux!Deposito, False
                
                If cDestino.NumDeposito = 0 Then
                     If Not cDestino.LeerDatos(DepositoDestino, False) Then Err.Raise 513, , "Error leyendo deposito destino: " & DepositoDestino
                End If
                
                
                cDe.InsertarEnHco 2, vvCstock.HoraMov, "CUP" & RecuperaValor(Me.Intercambio, 1), CantidadNecesaria
                cDe.VariacionKilosDeposito CantidadNecesaria
               
                If Not b Then
                    'NO existe. Lo creo
                    cP.Cantidad = CantidadNecesaria
                    cP.codAlmac = vvCstock.codAlmac
                    cP.codArtic = vvCstock.codArtic
                    cP.codProve = 0
                    cP.Fecha = CDate(txtFecha(1).Text)
                    cP.NumAlbar = "CUP" & RecuperaValor(Me.Intercambio, 1)
                    cP.numLote = DBLet(miRsAux!numLote, "T")
                    If cP.numLote Then cP.numLote = cP.NumAlbar
                    
                    If Not cP.Insertar Then
                        cad = "Error insertando partidas/lotes: " & cP.codArtic
                        Err.Raise vbObjectError + 513, , cad
                    End If
        
                Else
                    'Si existe. Lo decremento
                    cP.IncrementarCantidad CantidadNecesaria
                                    
                End If
                'Insertamos en la linea de smoval
                cL.tipoMov = 0
                cL.Cantidad = Abs(CantidadNecesaria)
                cL.codAlmac = vvCstock.codAlmac
                cL.codArtic = vvCstock.codArtic
                cL.numLote = cP.numLote
                cL.InsertarLote
                
                'Mohco
                If cDe.NumDeposito <> DepositoDestino Then
                  cDestino.numLote = "*" & cDe.numLote
                  cDestino.idPartida = cDe.idPartida
                  cDestino.InsertarEnHco 1, DateAdd("s", Segundos, vvCstock.HoraMov), cP.NumAlbar, Abs(CantidadNecesaria)
                End If
                
                If Val(miRsAux!fincuba) = 1 Then
                    
                    'Ponemos vacios los campos del deposito
                    'Fuera numero de lote y fuera kilos
                    If vvCstock.HoraMov = "" Then vvCstock.HoraMov = Format(Now, "hh:mm")
                    
                    If cDe.NumDeposito <> DepositoDestino Then
                        cDe.QuitarAsignacionDepositoVALL vvCstock.HoraMov, -cL.Cantidad, "CUP" & RecuperaValor(Me.Intercambio, 1)
                    Else
                        cDe.QuitarAsignacionDepositoVALL vvCstock.HoraMov, vvCstock.Cantidad - cL.Cantidad, TextoParaHcoDepositosLaVall
                    End If
                    Espera 0.75
                   
                Else
                    
                End If
            End If
        End If 'DE incializa stock
        miRsAux.MoveNext
    Wend
    miRsAux.Close
        
    Set cDe = Nothing

    If SoloComprobar Then
        RealizarCoupageLOTESVALL = True
        If AuxPartida <> "" Then
            AuxPartida = "-  " & Err_x_Articulo & AuxPartida & vbCrLf
            ErroresEnPartidas = ErroresEnPartidas & AuxPartida
        End If
        If ErroresEnPartidas <> "" Then
            ErroresEnPartidas = ErroresEnPartidas & "¿Continuar?"
            If MsgBox(ErroresEnPartidas, vbExclamation + vbYesNo) = vbNo Then RealizarCoupageLOTESVALL = False
        End If
        If Opcion = 5 Then MoviMIentosPosteriores = ""
        If MoviMIentosPosteriores <> "" Then
            MoviMIentosPosteriores = MoviMIentosPosteriores & vbCrLf & vbCrLf & "¿Continuar de igual modo?"
            If MsgBox(MoviMIentosPosteriores, vbQuestion + vbYesNoCancel) <> vbYes Then RealizarCoupageLOTESVALL = False
        End If
        
        GoTo ERealizarCUPLOTES 'para k haga los =nothing
    End If

        

    AuxPartida = ""
    
        

    'AHora comprobamos los stcosk de las entraddas , de las lineas
    cad = TransformaComasPuntos(CStr(CantidadMezcla))
    cad = "select codartic," & cad & " kilos,numlote,codalmac,deposito from olicoupage where codigo=" & RecuperaValor(Me.Intercambio, 1)
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'SOLO HAY una linea
    If Not miRsAux.EOF Then
        b = False
        If InicializarCStockCoupage(vvCstock, "E", True) Then    'Las lineas son de netrada
                
                                
                'AHora veremos los numeros de lote
                'EL nUMERO DE LOTE NO puede ser NULO
                CantidadNecesaria = vvCstock.Cantidad  'Para tenerla despues en los lotes
                
                
                                                        'Vemos si ya existe
                If cP.LeerDesdeArticulo(miRsAux!codArtic, miRsAux!codAlmac, miRsAux!numLote) Then
                    'Ya existia(por algun motivo)
                    cP.IncrementarCantidad CantidadNecesaria
                    
                Else
                    cP.Cantidad = CantidadNecesaria
                    cP.codAlmac = miRsAux!codAlmac
                    cP.codArtic = vvCstock.codArtic
                    cP.codProve = 0
                    cP.Fecha = CDate(txtFecha(1).Text)
                    cP.NumAlbar = "CUP" & RecuperaValor(Me.Intercambio, 1)
                    cP.numLote = miRsAux!numLote
                    If Not cP.Insertar Then Err.Raise vbObjectError + 513, , cad
                    
                End If
                
                'Insertamos en la linea de smoval
                cL.tipoMov = 1
                cL.Cantidad = Abs(CantidadNecesaria)
                cL.codAlmac = vvCstock.codAlmac
                cL.codArtic = vvCstock.codArtic
                cL.numLote = cP.numLote
                cL.InsertarLote
                
                b = True
                
                Set cDe = New cDeposito
                'Para que no de error insertando en hco
                T1 = Timer
                If Not cDe.LeerDatos(miRsAux!Deposito, False) Then b = False
                
                AuxPartida = DevuelveDesdeBD(conAri, "factorconversion", "sartic", "codartic", miRsAux!codArtic, "T")
                CantidadNecesaria = CCur(AuxPartida)
                
                
               
                cDe.numLote = cP.numLote
                cDe.idPartida = cP.idPartida
                Espera 0.5
                
                FLin = DateAdd("s", 1, vvCstock.HoraMov)
                'Metemos en el deposito los lotes pero no LA CANTIDAD, asi no genera el movimiento del deposito
                cDe.InsertarEnDeposito2 1, FLin, "CUP" & RecuperaValor(Me.Intercambio, 1)
                If CuantoHabiaDepositoDestino > 0 Then
                    
                    
                    cDe.InsertarEnHco 1, DateAdd("s", 5, FLin), "CUP" & RecuperaValor(Me.Intercambio, 1), CuantoHabiaDepositoDestino
                Else
                    'Inserto LINEA 0
                    cDe.InsertarEnHco 1, DateAdd("s", 5, FLin), "CUP" & RecuperaValor(Me.Intercambio, 1), 0
                End If
                cDe.VariacionKilosDeposito cL.Cantidad
                
                
                
                
                Espera 0.8
        End If
    End If
        
    RealizarCoupageLOTESVALL = b


    
ERealizarCUPLOTES:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set cP = Nothing
    Set miRsAux = Nothing
    Set vvCstock = Nothing
    Set cDe = Nothing
    Set cDestino = Nothing
End Function


