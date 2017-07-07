VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPist1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Produccion"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9420
   Icon            =   "frmPist1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   9420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameDesapareceCaja 
      Height          =   4215
      Left            =   0
      TabIndex        =   159
      Top             =   0
      Visible         =   0   'False
      Width           =   3615
      Begin VB.TextBox Text17 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   165
         Top             =   480
         Width           =   3255
      End
      Begin VB.TextBox Text16 
         ForeColor       =   &H00808080&
         Height          =   300
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   164
         Top             =   1080
         Width           =   3255
      End
      Begin VB.CommandButton cmdQuitarCajaSistema 
         Height          =   375
         Index           =   2
         Left            =   120
         Picture         =   "frmPist1.frx":1082
         Style           =   1  'Graphical
         TabIndex        =   163
         ToolTipText     =   "Realizar proceso"
         Top             =   3720
         Width           =   375
      End
      Begin VB.CommandButton cmdQuitarCajaSistema 
         Height          =   375
         Index           =   0
         Left            =   1200
         Picture         =   "frmPist1.frx":1A84
         Style           =   1  'Graphical
         TabIndex        =   162
         ToolTipText     =   "Quitar una"
         Top             =   3720
         Width           =   375
      End
      Begin VB.CommandButton cmdQuitarCajaSistema 
         Height          =   375
         Index           =   1
         Left            =   1800
         Picture         =   "frmPist1.frx":2486
         Style           =   1  'Graphical
         TabIndex        =   161
         ToolTipText     =   "Quitar TODOS"
         Top             =   3720
         Width           =   375
      End
      Begin VB.CommandButton cmdSalir 
         Height          =   375
         Index           =   12
         Left            =   2520
         Picture         =   "frmPist1.frx":7910
         Style           =   1  'Graphical
         TabIndex        =   160
         Top             =   3720
         Width           =   855
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2055
         Index           =   3
         Left            =   120
         TabIndex        =   166
         Top             =   1440
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   3625
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Caja"
            Object.Width           =   2364
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Palet"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label Label5 
         Caption         =   "Caja"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   167
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame FrameDevolucionOrdenCarga 
      Height          =   4215
      Left            =   0
      TabIndex        =   150
      Top             =   0
      Visible         =   0   'False
      Width           =   3615
      Begin VB.TextBox Text15 
         ForeColor       =   &H00808080&
         Height          =   300
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   158
         Top             =   840
         Width           =   3255
      End
      Begin VB.CommandButton cmdDevolucion 
         Height          =   375
         Index           =   2
         Left            =   1680
         Picture         =   "frmPist1.frx":7E9A
         Style           =   1  'Graphical
         TabIndex        =   157
         ToolTipText     =   "Devolver"
         Top             =   3720
         Width           =   375
      End
      Begin VB.CommandButton cmdDevolucion 
         Height          =   375
         Index           =   1
         Left            =   720
         Picture         =   "frmPist1.frx":889C
         Style           =   1  'Graphical
         TabIndex        =   156
         ToolTipText     =   "Quitar TODOS"
         Top             =   3720
         Width           =   375
      End
      Begin VB.CommandButton cmdDevolucion 
         Height          =   375
         Index           =   0
         Left            =   120
         Picture         =   "frmPist1.frx":DD26
         Style           =   1  'Graphical
         TabIndex        =   155
         ToolTipText     =   "Quitar una"
         Top             =   3720
         Width           =   375
      End
      Begin VB.TextBox Text14 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   153
         Top             =   240
         Width           =   2535
      End
      Begin VB.CommandButton cmdSalir 
         Height          =   375
         Index           =   11
         Left            =   2520
         Picture         =   "frmPist1.frx":E728
         Style           =   1  'Graphical
         TabIndex        =   151
         Top             =   3720
         Width           =   855
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2415
         Index           =   2
         Left            =   120
         TabIndex        =   154
         Top             =   1200
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   4260
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Caja"
            Object.Width           =   2364
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Orden"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Descrip"
            Object.Width           =   5644
         EndProperty
      End
      Begin VB.Label Label5 
         Caption         =   "Caja"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   152
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   3720
      Top             =   3000
   End
   Begin VB.Frame FrameCierrePalet 
      Height          =   4215
      Left            =   0
      TabIndex        =   79
      Top             =   0
      Visible         =   0   'False
      Width           =   3615
      Begin VB.CommandButton cmdCierrePalet 
         Caption         =   "6"
         Height          =   735
         Index           =   5
         Left            =   3030
         Picture         =   "frmPist1.frx":ECB2
         Style           =   1  'Graphical
         TabIndex        =   131
         ToolTipText     =   "Linea 6"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdCierrePalet 
         Caption         =   "5"
         Height          =   735
         Index           =   4
         Left            =   2448
         Picture         =   "frmPist1.frx":15504
         Style           =   1  'Graphical
         TabIndex        =   130
         ToolTipText     =   "Linea 5"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdCierrePalet 
         Caption         =   "4"
         Height          =   735
         Index           =   3
         Left            =   1866
         Picture         =   "frmPist1.frx":1BD56
         Style           =   1  'Graphical
         TabIndex        =   129
         ToolTipText     =   "Linea 4"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdCierrePalet 
         Caption         =   "3"
         Height          =   735
         Index           =   2
         Left            =   1284
         Picture         =   "frmPist1.frx":225A8
         Style           =   1  'Graphical
         TabIndex        =   128
         ToolTipText     =   "Linea 3"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdCierrePalet 
         Caption         =   "2"
         Height          =   735
         Index           =   1
         Left            =   702
         Picture         =   "frmPist1.frx":28DFA
         Style           =   1  'Graphical
         TabIndex        =   127
         ToolTipText     =   "Linea 2"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdCierrePalet 
         Caption         =   "1"
         Height          =   735
         Index           =   0
         Left            =   120
         Picture         =   "frmPist1.frx":2F64C
         Style           =   1  'Graphical
         TabIndex        =   126
         ToolTipText     =   "Linea 1"
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox chkNoContinuar 
         Caption         =   "Fin palets"
         Height          =   255
         Left            =   120
         TabIndex        =   120
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox txtCajaCierre 
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
         Index           =   0
         Left            =   840
         TabIndex        =   81
         Text            =   "Text2"
         Top             =   2760
         Width           =   2535
      End
      Begin VB.CommandButton cmdCerrarElPalet 
         Caption         =   "Cerrar palet"
         Height          =   495
         Left            =   1440
         TabIndex        =   82
         Top             =   3600
         Width           =   1335
      End
      Begin VB.TextBox txtCierrPalet 
         Alignment       =   1  'Right Justify
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
         Index           =   2
         Left            =   2640
         TabIndex        =   80
         Text            =   "Text11"
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtCierrPalet 
         Alignment       =   1  'Right Justify
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
         Index           =   1
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   88
         Text            =   "Text11"
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtCierrPalet 
         Alignment       =   1  'Right Justify
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
         Index           =   0
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   87
         Text            =   "Text11"
         Top             =   2040
         Width           =   735
      End
      Begin VB.CommandButton cmdSalir 
         Height          =   495
         Index           =   8
         Left            =   2880
         Picture         =   "frmPist1.frx":35E9E
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   3600
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Linea:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   84
         Top             =   1800
         Width           =   435
      End
      Begin VB.Label Label7 
         Caption         =   "Ult. caja"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   93
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Cerrar con"
         Height          =   255
         Index           =   3
         Left            =   2640
         TabIndex        =   91
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Leidas"
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   90
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Cajas/Pal"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   89
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label9 
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   86
         Top             =   1200
         Width           =   3375
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Index           =   0
         Left            =   240
         TabIndex        =   85
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   92
         Top             =   3120
         Width           =   3135
      End
   End
   Begin VB.Frame FramelecturaPostePalet 
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   0
      TabIndex        =   45
      Top             =   0
      Width           =   3615
      Begin VB.CommandButton cmdLectura 
         Caption         =   "OK"
         Height          =   495
         Left            =   2280
         TabIndex        =   50
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox TextErrPoste 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   1215
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   49
         Text            =   "frmPist1.frx":36428
         Top             =   2400
         Width           =   3375
      End
      Begin VB.TextBox txtlecturaPelt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   47
         Text            =   "Text2"
         Top             =   1200
         Width           =   3375
      End
      Begin VB.CommandButton cmdSalir 
         Height          =   495
         Index           =   4
         Left            =   720
         Picture         =   "frmPist1.frx":3642E
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Lectura poste paletizado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Index           =   6
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   3615
      Begin VB.Frame FrPrimeraProduccion 
         Height          =   3375
         Left            =   120
         TabIndex        =   95
         Top             =   120
         Width           =   3375
         Begin VB.TextBox Text11 
            Height          =   1095
            Index           =   1
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   99
            Text            =   "frmPist1.frx":369B8
            Top             =   2040
            Width           =   3015
         End
         Begin VB.TextBox Text11 
            Height          =   375
            Index           =   0
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   97
            Text            =   "Text11"
            Top             =   960
            Width           =   3015
         End
         Begin VB.Label Label2 
            Caption         =   "Falta tambien:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   240
            Index           =   12
            Left            =   120
            TabIndex        =   100
            Top             =   1800
            Width           =   2790
         End
         Begin VB.Label Label2 
            Caption         =   "Referencia a actualizar:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   240
            Index           =   11
            Left            =   120
            TabIndex        =   98
            Top             =   600
            Width           =   2790
         End
         Begin VB.Label Label2 
            Caption         =   "Asignar lotes linea produc."
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
            Height          =   240
            Index           =   10
            Left            =   120
            TabIndex        =   96
            Top             =   240
            Width           =   2790
         End
      End
      Begin VB.CommandButton cmdSalir 
         Height          =   495
         Index           =   1
         Left            =   120
         Picture         =   "frmPist1.frx":369BF
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         TabIndex        =   14
         Text            =   "Text2"
         Top             =   2400
         Width           =   2055
      End
      Begin VB.CommandButton cmdAceptarCambioLote 
         Caption         =   "Aceptar"
         Height          =   495
         Left            =   1680
         TabIndex        =   15
         Top             =   3600
         Width           =   1695
      End
      Begin VB.TextBox TxtUD 
         Alignment       =   1  'Right Justify
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
         Height          =   495
         Left            =   2040
         TabIndex        =   24
         Text            =   "Text2"
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         TabIndex        =   13
         Text            =   "Text2"
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Indique produccion realizada"
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
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   26
         Top             =   120
         Width           =   3375
      End
      Begin VB.Label Label2 
         Caption         =   "Uds"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   25
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Uds x caja"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   23
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "CAJAS"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   22
         Top             =   960
         Width           =   615
      End
   End
   Begin VB.Frame FrameNPalet 
      Height          =   4215
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   3615
      Begin VB.CheckBox chkNoImprPaletManual 
         Caption         =   "  NO IMPR"
         Height          =   735
         Left            =   240
         TabIndex        =   143
         Top             =   3240
         Width           =   735
      End
      Begin VB.TextBox Text12 
         Height          =   495
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   39
         Text            =   "frmPist1.frx":36F49
         Top             =   1680
         Width           =   2295
      End
      Begin VB.CommandButton cmdQuiCajaPalMan 
         Height          =   375
         Index           =   1
         Left            =   600
         Picture         =   "frmPist1.frx":36F4F
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Limpiar"
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox txtNumCajas 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   118
         Text            =   "0"
         Top             =   1080
         Width           =   735
      End
      Begin VB.CommandButton cmdQuiCajaPalMan 
         Height          =   375
         Index           =   0
         Left            =   120
         Picture         =   "frmPist1.frx":374D9
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Quitar UNA caja"
         Top             =   2640
         Width           =   375
      End
      Begin VB.CommandButton cmdAceptarNPalet 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   1080
         TabIndex        =   40
         Top             =   3720
         Width           =   1455
      End
      Begin VB.TextBox Text7 
         Height          =   495
         Left            =   120
         TabIndex        =   38
         Text            =   "Text7"
         Top             =   240
         Width           =   3255
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   41
         Text            =   "frmPist1.frx":37EDB
         Top             =   1080
         Width           =   2415
      End
      Begin VB.CommandButton cmdSalir 
         Height          =   375
         Index           =   3
         Left            =   2880
         Picture         =   "frmPist1.frx":37EE1
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   3720
         Width           =   495
      End
      Begin MSComctlLib.ListView lwNPalet 
         Height          =   1335
         Left            =   1080
         TabIndex        =   78
         Top             =   2280
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   2355
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "codartic"
            Object.Width           =   3422
         EndProperty
      End
      Begin VB.Label Label6 
         Caption         =   "IdPalet"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   123
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Total"
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   117
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Mensaje"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   77
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Cajas"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   76
         Top             =   2280
         Width           =   495
      End
   End
   Begin VB.Frame FrAjusteCajas 
      Height          =   4215
      Left            =   0
      TabIndex        =   101
      Top             =   0
      Visible         =   0   'False
      Width           =   3615
      Begin VB.CommandButton cmdCierrePalet 
         Caption         =   "6"
         Height          =   615
         Index           =   15
         Left            =   3000
         Picture         =   "frmPist1.frx":3846B
         Style           =   1  'Graphical
         TabIndex        =   125
         ToolTipText     =   "Linea 6"
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton cmdCierrePalet 
         Caption         =   "5"
         Height          =   615
         Index           =   14
         Left            =   2424
         Picture         =   "frmPist1.frx":3ECBD
         Style           =   1  'Graphical
         TabIndex        =   124
         ToolTipText     =   "Linea 5"
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton cmdCierrePalet 
         Caption         =   "4"
         Height          =   615
         Index           =   13
         Left            =   1848
         Picture         =   "frmPist1.frx":4550F
         Style           =   1  'Graphical
         TabIndex        =   119
         ToolTipText     =   "Linea 4"
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton cmdSalir 
         Height          =   495
         Index           =   9
         Left            =   2760
         Picture         =   "frmPist1.frx":4BD61
         Style           =   1  'Graphical
         TabIndex        =   108
         Top             =   3600
         Width           =   615
      End
      Begin VB.CommandButton cmdAjustar 
         Caption         =   "Ajustar"
         Height          =   495
         Left            =   120
         TabIndex        =   107
         Top             =   3600
         Width           =   2055
      End
      Begin VB.TextBox txtCierrPalet 
         Alignment       =   1  'Right Justify
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
         Left            =   2520
         TabIndex        =   106
         Text            =   "Text11"
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox txtCierrPalet 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   112
         Text            =   "Text11"
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtCajaCierre 
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
         Index           =   1
         Left            =   120
         TabIndex        =   105
         Text            =   "Text2"
         Top             =   2880
         Width           =   2295
      End
      Begin VB.CommandButton cmdCierrePalet 
         Caption         =   "1"
         Height          =   615
         Index           =   10
         Left            =   120
         Picture         =   "frmPist1.frx":4C2EB
         Style           =   1  'Graphical
         TabIndex        =   102
         ToolTipText     =   "Linea 1"
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton cmdCierrePalet 
         Caption         =   "2"
         Height          =   615
         Index           =   11
         Left            =   696
         Picture         =   "frmPist1.frx":52B3D
         Style           =   1  'Graphical
         TabIndex        =   103
         ToolTipText     =   "Linea 2"
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton cmdCierrePalet 
         Caption         =   "3"
         Height          =   615
         Index           =   12
         Left            =   1272
         Picture         =   "frmPist1.frx":5938F
         Style           =   1  'Graphical
         TabIndex        =   104
         ToolTipText     =   "Linea 3"
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   555
         Index           =   1
         Left            =   720
         TabIndex        =   109
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Ajustar numero cajas"
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
         Height          =   240
         Index           =   13
         Left            =   600
         TabIndex        =   116
         Top             =   240
         Width           =   2790
      End
      Begin VB.Label Label7 
         Caption         =   "Leidas"
         Height          =   255
         Index           =   9
         Left            =   1800
         TabIndex        =   115
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Real"
         Height          =   255
         Index           =   8
         Left            =   2520
         TabIndex        =   114
         Top             =   2640
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Ult. caja"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   113
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label9 
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   111
         Top             =   1440
         Width           =   3255
      End
      Begin VB.Label Label7 
         Caption         =   "Linea:"
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   110
         Top             =   1920
         Width           =   615
      End
   End
   Begin VB.Frame FrameSelect 
      Height          =   4215
      Left            =   120
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   3495
      Begin VB.Frame FrameSegundasAcciones 
         Height          =   3495
         Left            =   0
         TabIndex        =   144
         Top             =   0
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton cmdSelect 
            Caption         =   "Perdida o baja"
            Height          =   615
            Index           =   11
            Left            =   120
            TabIndex        =   149
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CommandButton cmdSelect 
            Caption         =   "Devolucion"
            Height          =   615
            Index           =   10
            Left            =   1800
            TabIndex        =   148
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton cmdSelect 
            Caption         =   "mover cajas"
            Height          =   615
            Index           =   5
            Left            =   120
            TabIndex        =   146
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Esta dentro del frame ppal. Es para tener mas acciones"
            Height          =   495
            Left            =   120
            TabIndex        =   145
            Top             =   2880
            Visible         =   0   'False
            Width           =   3255
         End
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Palet ""manual"""
         Height          =   495
         Index           =   8
         Left            =   1920
         TabIndex        =   5
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Ajuste cajas"
         Height          =   495
         Index           =   7
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Cerrar palet"
         Height          =   495
         Index           =   6
         Left            =   1920
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdSalir 
         Height          =   495
         Index           =   7
         Left            =   1920
         Picture         =   "frmPist1.frx":5FBE1
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3600
         Width           =   1455
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Expedicion"
         Height          =   615
         Index           =   4
         Left            =   120
         TabIndex        =   6
         Top             =   2760
         Width           =   1455
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Poste paletizado"
         Height          =   495
         Index           =   3
         Left            =   1920
         TabIndex        =   3
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Nuevo palet"
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Paletización"
         Height          =   615
         Index           =   1
         Left            =   3360
         TabIndex        =   8
         Top             =   1320
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Producción"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Mas ------>"
         Height          =   615
         Index           =   9
         Left            =   1920
         TabIndex        =   147
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Image imgRepetidos 
         Height          =   240
         Left            =   960
         Picture         =   "frmPist1.frx":6016B
         ToolTipText     =   "Poste sin leer"
         Top             =   3660
         Width           =   240
      End
      Begin VB.Image imgPoste 
         Height          =   240
         Left            =   240
         Picture         =   "frmPist1.frx":669BD
         ToolTipText     =   "Poste sin leer"
         Top             =   3660
         Width           =   240
      End
      Begin VB.Line Line4 
         X1              =   840
         X2              =   840
         Y1              =   3600
         Y2              =   4120
      End
      Begin VB.Shape Shape1 
         Height          =   560
         Left            =   120
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Line Line3 
         Index           =   1
         X1              =   120
         X2              =   3360
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line3 
         Index           =   0
         X1              =   120
         X2              =   3360
         Y1              =   1680
         Y2              =   1680
      End
   End
   Begin VB.Frame FrameExpedicion 
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   0
      TabIndex        =   51
      Top             =   0
      Visible         =   0   'False
      Width           =   3615
      Begin VB.CommandButton cmdPaletVariosAlbaranes 
         Caption         =   "Palet"
         Height          =   495
         Left            =   1560
         TabIndex        =   136
         Top             =   3600
         Width           =   615
      End
      Begin VB.CommandButton cmdVerComoVaExpedicion 
         Caption         =   "Ver"
         Height          =   495
         Left            =   840
         TabIndex        =   57
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox Text9 
         Height          =   2055
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   56
         Text            =   "frmPist1.frx":6D20F
         Top             =   1440
         Width           =   3255
      End
      Begin VB.CommandButton cmdLimpExpedicion 
         Caption         =   "Limpiar"
         Height          =   495
         Left            =   120
         TabIndex        =   54
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox Text10 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   53
         Text            =   "Text2"
         Top             =   720
         Width           =   3255
      End
      Begin VB.CommandButton cmdSalir 
         Height          =   495
         Index           =   5
         Left            =   2880
         Picture         =   "frmPist1.frx":6D215
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   3600
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Orden carga"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         TabIndex        =   63
         Top             =   120
         Width           =   3495
      End
      Begin VB.Label Label5 
         Caption         =   "Jugamos con la Len de las etiq"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   55
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Frame FrameQuitaCaja 
      Height          =   4215
      Left            =   0
      TabIndex        =   65
      Top             =   0
      Visible         =   0   'False
      Width           =   3615
      Begin VB.TextBox txtMoverUltimasNcajas 
         Height          =   375
         Left            =   1080
         TabIndex        =   68
         Text            =   "Text13"
         Top             =   1200
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdImprEtiqPal 
         Height          =   375
         Left            =   120
         Picture         =   "frmPist1.frx":6D79F
         Style           =   1  'Graphical
         TabIndex        =   122
         ToolTipText     =   "Imprimie etiqueta palet destino"
         Top             =   3120
         Width           =   375
      End
      Begin VB.CommandButton cmdCreaPalet 
         Height          =   375
         Left            =   120
         Picture         =   "frmPist1.frx":6E1A1
         Style           =   1  'Graphical
         TabIndex        =   121
         ToolTipText     =   "Crear nuevo  palet"
         Top             =   2640
         Width           =   375
      End
      Begin VB.CommandButton cmdSalir 
         Height          =   375
         Index           =   6
         Left            =   2880
         Picture         =   "frmPist1.frx":6EBA3
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   3720
         Width           =   495
      End
      Begin VB.TextBox txtObsCambioPalet 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   720
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   73
         Text            =   "frmPist1.frx":6F12D
         Top             =   2040
         Width           =   2655
      End
      Begin VB.TextBox txtCCaja 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   69
         Text            =   "Text2"
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txtCPalet 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1080
         TabIndex        =   67
         Text            =   "Text11"
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtCPalet 
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
         Index           =   0
         Left            =   1080
         TabIndex        =   66
         Text            =   "Text11"
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Cuantas"
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   133
         Top             =   1260
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Palet destino"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   75
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Palet origen"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   74
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Caja:"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   72
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Quita las cajas de un palet.       Puede ponerlas en otros (o no)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   71
         Top             =   3720
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4215
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   3615
      Begin VB.CommandButton cmdLimpiar 
         Height          =   495
         Index           =   0
         Left            =   2160
         Picture         =   "frmPist1.frx":6F133
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   120
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Height          =   2055
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   20
         Text            =   "frmPist1.frx":745BD
         Top             =   1440
         Width           =   3255
      End
      Begin VB.CommandButton cmdValidar 
         Caption         =   "Aceptar"
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   3600
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   10
         Text            =   "Text2"
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   840
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdSalir 
         Height          =   495
         Index           =   0
         Left            =   2760
         Picture         =   "frmPist1.frx":745C3
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   150
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Lectura:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "LINEA"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   330
         Width           =   975
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   3360
         Y1              =   690
         Y2              =   690
      End
   End
   Begin VB.Frame FrameVerSituacionCarga 
      Height          =   4215
      Left            =   0
      TabIndex        =   58
      Top             =   0
      Visible         =   0   'False
      Width           =   3495
      Begin VB.CommandButton cmdEliminarCargaExpedicion 
         Caption         =   "-"
         Height          =   255
         Left            =   2880
         TabIndex        =   132
         ToolTipText     =   "Eliminar carga referencia"
         Top             =   3840
         Width           =   375
      End
      Begin VB.CommandButton cmdVolverExp 
         Height          =   375
         Left            =   2880
         Picture         =   "frmPist1.frx":74B4D
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "VOLVER"
         Top             =   150
         Width           =   495
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   975
         Index           =   0
         Left            =   120
         TabIndex        =   60
         Top             =   600
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   1720
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Albaran"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fin"
            Object.Width           =   882
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1935
         Index           =   1
         Left            =   120
         TabIndex        =   61
         Top             =   1800
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   3413
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "codartic"
            Object.Width           =   2364
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cajas"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Carga"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Nombre"
            Object.Width           =   5644
         EndProperty
      End
      Begin VB.Label Label5 
         Caption         =   "Orden carga"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   4
         Left            =   1320
         TabIndex        =   64
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Orden carga"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   59
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame FramePalet 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4215
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   3615
      Begin VB.TextBox txtIdPal 
         Height          =   375
         Left            =   1320
         TabIndex        =   35
         Text            =   "Text5"
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox txtObservaCajas 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   31
         Text            =   "frmPist1.frx":750D7
         Top             =   2160
         Width           =   3255
      End
      Begin VB.CommandButton cmdSalir 
         Height          =   495
         Index           =   2
         Left            =   2760
         Picture         =   "frmPist1.frx":750DD
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   120
         Width           =   615
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1560
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   120
         Width           =   855
      End
      Begin VB.TextBox txtCajas 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   30
         Text            =   "Text2"
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Label Label3 
         Caption         =   "ID Palet"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   36
         Top             =   960
         Width           =   735
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   3360
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label3 
         Caption         =   "LINEA PALET"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   34
         Top             =   330
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Caja:"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   33
         Top             =   1560
         Width           =   615
      End
   End
   Begin VB.Frame FrameExpedPalet 
      Height          =   4215
      Left            =   120
      TabIndex        =   134
      Top             =   0
      Visible         =   0   'False
      Width           =   3375
      Begin VB.CommandButton cmdAceptarPaletVariosAlb 
         Caption         =   "Aceptar"
         Height          =   495
         Left            =   120
         TabIndex        =   142
         Top             =   3600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text13 
         Height          =   2895
         Index           =   2
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   141
         Text            =   "frmPist1.frx":75667
         Top             =   600
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.TextBox Text13 
         Height          =   2175
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   140
         Text            =   "frmPist1.frx":7566D
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox Text13 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   720
         TabIndex        =   138
         Text            =   "Text2"
         Top             =   600
         Width           =   2535
      End
      Begin VB.CommandButton cmdSalir 
         Height          =   495
         Index           =   10
         Left            =   2520
         Picture         =   "frmPist1.frx":75673
         Style           =   1  'Graphical
         TabIndex        =   135
         Top             =   3600
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Palet"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   139
         Top             =   600
         Width           =   450
      End
      Begin VB.Label Label5 
         Caption         =   "Orden carga"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   137
         Top             =   240
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmPist1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private PrimVez As Boolean

Dim cLinPr As cLineaProduccion  'Para saber lo que estamos produciendo en la linea
Dim SubLinea As cLineaProCompo
Dim IndexSublinea As Integer  'identificara dentro del vector que sublinea es



Dim cPal As CPalet

Dim SQL As String
Dim RS As ADODB.Recordset
Dim Cp As cPartidas   'la lee en el lostfocus

Dim Tiempo1 As Single  'Para que cada xTiempo vuelva a leer las lineas


'Para acelerar las busquedas
Dim idTrazaAntiguo As Long
Dim ValoresLeidos As String


Private Sub cmdAceptar_Click()

End Sub

Private Sub cmdAccionePPal_Click()
    
End Sub

Private Sub cmdAceptarCambioLote_Click()
Dim Can As Currency
Dim C As Long
Dim Cajas As Integer
Dim Indice As Integer



    If Me.FrPrimeraProduccion.Visible Then
        'NO es cierre de PALET, es asignar numero de lote a la produccion
        If AsignacionPrimerLoteProduccion Then
            Text2.Text = "": Text3.Text = ""
            cmdSalir_Click 1
            PonerFoco Text2
            Exit Sub
        End If
        
    End If
    
    If Text4.Text = "" Then
        Cajas = 0
    Else
        Cajas = ImporteFormateado(Text4.Text)
    End If

    If Text6.Text = "" Then
        Can = 0
    Else
        Can = ImporteFormateado(Text6.Text)
    End If
    If Can <= 0 Then
        MsgBox "Indique la cantidad producida. Debe ser mayor que cero", vbExclamation
        Exit Sub
    End If

    C = Can \ CInt(Me.TxtUD.Text)
    If (Can Mod CInt(Me.TxtUD.Text)) > 0 Then C = C + 1
    
    C = C * CInt(Me.TxtUD.Text)  'Cantidad producida si llenaramos las cajas
    
    SQL = String(30, "-") & vbCrLf
    SQL = SQL & vbCrLf & "UNIDADES: " & Format(CInt(Can), "#,###,##0") & vbCrLf & "Cajas:        " & Cajas & vbCrLf
    If CInt(Can) <> C Then
        
        'Cantidad de cajas a producir  distinto
        C = C - CInt(Can)
        C = CInt(Me.TxtUD.Text) - C
        SQL = SQL & vbCrLf & "Cajas incompletas " & vbCrLf & "Una cajas con  " & C & " uds" & vbCrLf
    End If
    SQL = vbCrLf & SQL & String(30, "-") & vbCrLf & vbCrLf & "¿CONTINUAR?"
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        
    
    'JUNIO 2014
    'FALTA###
    'De momento esta solo para el aceite. Tambien podremos regularizar las
    'partidas cuando sean final de lote
    ' Como NO cierra el aceite desde aqui, numdeposito=0
    If cLinPr.CerrarParaCambioLote(Can, Cajas, IndexSublinea, Cp.NUmlote, False, 0) Then
        'Y marco la etiqueta como "empuieza en produccion"
        SQL = "UPDATE spartidaslin set fechaulizada = " & DBSet(Now, "FH")
        SQL = SQL & " WHERE  bulto = " & Right(Text2.Text, 3) & " AND id = " & Cp.idPartida
        EjecutaSQL conAri, SQL, True
        Unload Me
    End If
    Conn.Execute "commit"   'tipo flush
End Sub

Private Sub cmdAceptarNPalet_Click()
Dim OK As Boolean
Dim Aux As String
Dim i As Integer
Dim NP As Long
Dim Imprime As Boolean
    If Me.lwNPalet.ListItems.Count = 0 Then Exit Sub
    
    
    
    Set cPal = Nothing
    'Comprobaremos que las cajas asignadas son mas de una trazabilidad
    SQL = ""
    If InStr(1, Me.Caption, "*") = 0 Then
        'Para los palets de produccion MANUAL
        SQL = Mid(Me.lwNPalet.ListItems(1).Text, 1, 8) 'Primera traza
        For i = 1 To Me.lwNPalet.ListItems.Count
            If Mid(Me.lwNPalet.ListItems(i).Text, 1, 8) = SQL Then
                'ok
            Else
                SQL = ""
                Exit For
            End If
        Next
                
        If SQL = "" Then
            SQL = String(30, "*") & vbCrLf & vbCrLf
            SQL = SQL & " Distintos lotes de trazabilidad " & vbCrLf & vbCrLf & SQL
        Else
            SQL = ""
        End If
      
    
        'PALET MANUAL. VERemos si es nuevo o abre uno
        If Text12.Text <> "" Then
            'VA A poner las cajas sobre otro palet
            Set cPal = New CPalet
            
            If cPal.Leer(CLng(Mid(Text12.Text, 2, 8))) Then
                i = 1
            Else
                i = 0
            End If
            If i = 0 Then
                MsgBox "No existe el palet", vbExclamation
                Set cPal = Nothing
                Exit Sub
            End If
            SQL = "Va a abrir el palet " & cPal.ID & " y añadirle las cajas?"
        Else
            SQL = SQL & vbCrLf & "¿Crear palet?"
        End If
    Else
        SQL = SQL & vbCrLf & "¿Crear palet?"
    End If
    
    
        
    
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    
    Conn.BeginTrans
    
    If cPal Is Nothing Then
    
        SQL = DevuelveDesdeBD(conAri, "max(idpalet)", "prodpalets", "1", "1")
        If SQL = "" Then SQL = "0"
        NP = CStr(Val(SQL) + 1)
    
        'FALTA###  mvarTipoImpresion
         '`idpalet`,`LineaPeletiza`,`fhinicio`,`fhFin`,`CajasProd`,`L0`,`L1`,`L2`,`L3`,`L4`,`L5`,`L6`,`L7`
        SQL = NP & ",0,NOW(),NOW()," & Me.lwNPalet.ListItems.Count & ",'0','0','0','0','0','0','0','0',1)"   'UNO en manual=linea 8
        SQL = "insert into `prodpalets` (`idpalet`,`LineaPeletiza`,`fhinicio`,`fhFin`,`CajasProd`,`L0`,`L1`,`L2`,`L3`,`L4`,`L5`,`L6`,`L7`,`L8`) values (" & SQL
    
        OK = EjecutaSQL(conAri, SQL, True)   'insertamos el palet
    
    
    Else
        NP = cPal.ID
        OK = True
    End If
    
    
   
    
    
    'Insertamos en palettraza para el MANUAL
    If OK Then
        If InStr(1, Me.Caption, "*") = 0 Then
            ValoresLeidos = "|"
            SQL = ""
            For i = 1 To Me.lwNPalet.ListItems.Count
                'prodpaletstraza idpalet lotetraza fh
                idTrazaAntiguo = Val(Mid(Me.lwNPalet.ListItems(i).Text, 1, 8))
                
                If InStr(1, ValoresLeidos, "|" & idTrazaAntiguo & "|") = 0 Then
                
                    'Vere si no esa ya insertado
                    Aux = "idpalet=" & NP & " AND lotetraza "
                    Aux = DevuelveDesdeBD(conAri, "lotetraza", "prodpaletstraza", Aux, CStr(idTrazaAntiguo))
                    'No existia
                    If Aux = "" Then SQL = SQL & ", (" & NP & "," & Mid(Me.lwNPalet.ListItems(i).Text, 1, 8) & ",NOW())"
                    
                    ValoresLeidos = ValoresLeidos & idTrazaAntiguo & "|"
                End If
                
            Next
            If SQL <> "" Then
                SQL = Mid(SQL, 2) 'quito la primera coma
                SQL = "INSERT INTO prodpaletstraza (idpalet, lotetraza, fh) VALUES " & SQL
                          
                OK = EjecutaSQL(conAri, SQL, True)   'insertamos el palet
            Else
                OK = True
            End If
            idTrazaAntiguo = 0
            ValoresLeidos = ""
        End If
    End If
    
    'las cajas las llevamos a ese palet
    If OK Then
        For i = 1 To Me.lwNPalet.ListItems.Count
        
            If InStr(1, Me.Caption, "*") = 0 Then
                'Nuevo palet de produccion manual
                'insert into `prodcajas` (`lotetraza`,`idcaja`,`idpalet`,`fcreacion`)
                SQL = "insert into `prodcajas` (`lotetraza`,`idcaja`,`idpalet`,`fcreacion`) VALUES ("
                SQL = SQL & Mid(Me.lwNPalet.ListItems(i).Text, 1, 8) & "," & Mid(Me.lwNPalet.ListItems(i).Text, 9) & ","
                SQL = SQL & NP & ",NOW())"
            Else
                'Reasignacion de cajas que ya existen
                SQL = "UPDATE prodcajas SET idpalet=" & NP
                SQL = SQL & " WHERE lotetraza =" & Mid(Me.lwNPalet.ListItems(i).Text, 1, 8) & " AND idcaja = " & Mid(Me.lwNPalet.ListItems(i).Text, 9)
            End If
            OK = EjecutaSQL(conAri, SQL, True)   'insertamos el palet
            If Not OK Then i = Me.lwNPalet.ListItems.Count + 1 'Se salga
        Next i
        
        If InStr(1, Me.Caption, "*") = 0 Then
            If Not cPal Is Nothing Then
                Espera 0.2
                SQL = DevuelveDesdeBD(conAri, "count(*)", "prodcajas", "idpalet", cPal.ID)
                If SQL = "" Then SQL = "0"
                SQL = "UPDATE prodpalets set fhFin=" & DBSet(Now, "FH") & " ,CajasProd= " & SQL & " WHERE idpalet =" & cPal.ID
                EjecutaSQL conAri, SQL, True
            End If
        End If
    End If
    
    If OK Then
        Conn.CommitTrans
        If cPal Is Nothing Then
            MsgBox "Palet creado: " & NP, vbExclamation
        Else
            MsgBox "Cajas añadidas", vbInformation
        End If
        Set cPal = Nothing
        Text7.Text = ""
        Text5.Text = ""
        Text12.Text = ""
        Me.txtNumCajas = ""
        lwNPalet.ListItems.Clear
        
        
        
        
        'MAYO. Imprimimos
        Imprime = True
        If InStr(1, Me.Caption, "*") = 0 Then
            'Manual
            If Me.chkNoImprPaletManual.Value = 1 Then Imprime = False
        End If
        
        If Imprime Then
            Text5.Text = "Imprimiendo"
            Text5.Refresh
            Dim C As Collection
            Set cPal = New CPalet
            If cPal.Leer(NP) Then
                    cPal.CargaDatosPalet C, True, i, False
                    ImprimirPalet cPal.ID, cPal.TipoImpresion
            End If
            Set cPal = Nothing
        
         End If
        
        Text5.Text = ""
        PonerFoco Text7
    Else
        Conn.RollbackTrans
    End If
End Sub

Private Sub cmdAjustar_Click()
Dim InsertarCajasNoLeidasPost As Integer

    If Me.txtCajaCierre(1).Text = "" Then Exit Sub
    If Me.txtCierrPalet(4).Text = "" Then Me.txtCierrPalet(4).Text = "0"


    SQL = "idcaja<=" & Val(Mid(Me.txtCajaCierre(1).Text, 9)) & " AND idpalet"
    SQL = DevuelveDesdeBD(conAri, "count(*)", "prodcajas", SQL, CStr(cPal.ID))
    If SQL = "" Then SQL = "0"
    InsertarCajasNoLeidasPost = CInt(Me.txtCierrPalet(4).Text) - Val(SQL)
    
    If InsertarCajasNoLeidasPost < 0 Then
        'Hay mas cajas de las que deberian
        MsgBox "Hay mas cajas de las que deberian: " & CInt(Me.txtCierrPalet(3).Text) & " / " & Val(SQL), vbInformation
        Exit Sub
    End If
    
    If InsertarCajasNoLeidasPost = 0 Then
        Me.txtCierrPalet(4).Text = Me.txtCierrPalet(3).Text
        MsgBox "Ok. ", vbInformation
        LimpiarCierrePalet
        Exit Sub
    End If
    
    
    If MsgBox("Va a insertar las cajas que faltan: " & InsertarCajasNoLeidasPost & " , continuar?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub

    'FALTA

    If Not InsertarCajasNoLeidasPosteSub(InsertarCajasNoLeidasPost, False) Then Exit Sub
    LimpiarCierrePalet

    
          
End Sub

Private Sub cmdCerrarElPalet_Click()
Dim Situar As Byte

    SQL = ""
    Situar = 1
    If Me.txtCierrPalet(1).Text = "" Then
        SQL = "Palet sin cajas"
    Else
        If Val(Me.txtCierrPalet(2).Text) <= 0 Then
            SQL = "Cantidad real: " & txtCierrPalet(2).Text & "?"
        Else
            If Me.txtCierrPalet(2).Text = "" Then
                SQL = "Escriba cantidad cierre"
            Else
                If Val(Me.txtCierrPalet(2).Text) <= 0 Then SQL = "Cantidad debe ser positiva"
            End If
              
        End If
      
    End If
    If txtCajaCierre(0).Text = "" Then
        SQL = SQL & vbCrLf & "Indique caja cierre"
        Situar = 2
    End If
        
    If SQL <> "" Then
        MsgBox SQL, vbExclamation
        
        If Situar = 2 Then
            PonerFoco txtCajaCierre(0)
        Else
            PonerFoco txtCierrPalet(2)
        End If
        Exit Sub
    End If

    
    Screen.MousePointer = vbHourglass
    Label10.Caption = "Cerrando"
    Label10.Refresh
    If CerrarPalet_ Then cmdSalir_Click 6
    Screen.MousePointer = vbDefault
    Label10.Caption = ""
End Sub

Private Sub cmdCierrePalet_Click(Index As Integer)
Dim linea As Byte

    If Index < 10 Then
        linea = Index + 1
    Else
        linea = Index - 9
    End If
    
    
    'Cierre palet
    If LeerLineaPalet(linea) Then
        
        PonerDatosLineaPalet2 linea, Index < 10
    Else
        LimpiarCierrePalet
    End If

End Sub

Private Sub cmdCreaPalet_Click()
    idTrazaAntiguo = 0
    If cmdCreaPalet.Tag > 0 Then
        'Creo un palet nuevo
        If MsgBox("Palet ya creado, desea crear OTRO?", vbQuestion + vbYesNo) = vbNo Then idTrazaAntiguo = cmdCreaPalet.Tag
        
        
    End If
    If idTrazaAntiguo = 0 Then
    
        
    
        Set cPal = New CPalet
        cPal.FechaInicio = Now
        cPal.LineaPeletizacion = 0
        If cPal.CrearPalet Then
            cmdCreaPalet.Tag = cPal.ID
            idTrazaAntiguo = 1 'para que pinte el palet
        Else
            cmdCreaPalet.Tag = 0
        End If
        
        cPal.CerrarPalet 0
        
        Set cPal = Nothing
        
        
    End If
    If idTrazaAntiguo <> 0 Then
        txtCPalet(1).Text = "1" & Format(cmdCreaPalet.Tag, "00000000") & "1"
        PonerFoco txtCPalet(1)
    End If
End Sub

Private Sub cmdDevolucion_Click(Index As Integer)
    
    If Me.ListView1(2).ListItems.Count = 0 Then Exit Sub
    
    
    If Index = 1 Then If Me.ListView1(2).SelectedItem Is Nothing Then Exit Sub
    
    
    Select Case Index
    Case 0
        SQL = "Va a quitar la caja " & ListView1(2).SelectedItem.Text & vbCrLf & "¿Continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        
        ListView1(2).ListItems.Remove ListView1(2).SelectedItem.Index
        
    Case 1
        SQL = "Desea limpiar todas las cajas?"
        If MsgBox(SQL, vbExclamation + vbYesNo) = vbNo Then Exit Sub
        
        If MsgBox("Seguro?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        
        ListView1(2).ListItems.Clear
        PonerFoco Text14
        
    Case 2
        ' Adelante con el proceso
        SQL = "Finalizar el proceso de devolucion." & vbCrLf & "¿Continuar"
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        
        
        
        RealizarProcesoDevolucion
    
    End Select
    
End Sub

Private Sub cmdEliminarCargaExpedicion_Click()
    If Me.ListView1(1).ListItems.Count = 0 Then Exit Sub
    If Me.ListView1(1).SelectedItem Is Nothing Then Exit Sub
    
    SQL = "Eliminar expedicion : " & vbCrLf & ListView1(1).SelectedItem.SubItems(3)
    SQL = SQL & vbCrLf & " ALBARAN " & Me.ListView1(0).SelectedItem.Text & "?"
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    idTrazaAntiguo = 0
    
    SQL = "DELETE FROM srepartolotcaj WHERE " & ListView1(1).SelectedItem.Tag
    EjecutaSQL conAri, SQL, True
     
    
    SQL = "update `srepartolot` set llevamos=0"
    SQL = SQL & " where " & ListView1(1).SelectedItem.Tag
    EjecutaSQL conAri, SQL, True
     
    If Not Me.ListView1(0).SelectedItem Is Nothing Then CargarLineasAlbaranExpedicion Me.ListView1(0).SelectedItem, True
  
    
End Sub

Private Sub cmdImprEtiqPal_Click()
    If txtCPalet(1).Text = "" Then
        MsgBox "palet destino", vbExclamation
        Exit Sub
    End If
    
    SQL = Mid(txtCPalet(1).Text, 2, 8) 'los ocho centrales
    If SQL = "" Then Exit Sub
    If Not IsNumeric(SQL) Then Exit Sub
    
    idTrazaAntiguo = Val(SQL)
    
        Dim C As Collection
        Set cPal = New CPalet
        If cPal.Leer(idTrazaAntiguo) Then
                cPal.CargaDatosPalet C, True, CInt(idTrazaAntiguo), False
                ImprimirPalet cPal.ID, cPal.TipoImpresion
                MsgBox "Fichero impresion OK", vbInformation
        End If
        Set cPal = Nothing
    
    idTrazaAntiguo = 0
    SQL = ""
End Sub

Private Sub cmdLectura_Click()
    ProcesarCodigoCaja
End Sub

Private Sub cmdLimpExpedicion_Click()
    If Trim(CStr(Label5(0).Tag)) <> "" Then
        If Label5(0).Tag > 0 Then
            If MsgBox("Limpiar pantalla orden carga?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
    End If
    Label5(0).Caption = "Orden carga "
    Label5(0).Tag = 0
    
    Label5(1).Caption = "Albaran"
    Label5(1).Tag = ""
    Text9.Text = ""
    Text10.Text = ""
    PonerFoco Text10
End Sub



Private Sub cmdLimpiar_Click(Index As Integer)
    If Index = 0 Then
        'limpiar datos produccion
        Text2.Text = ""
        Text3.Text = ""
    End If
End Sub



Private Sub cmdPaletVariosAlbaranes_Click()
    If Val(Label5(0).Tag) = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    Label5(3).Caption = "Orden carga: " & Label5(0).Tag
    Me.Text13(0).Text = "": Me.Text13(1).Text = "": Me.Text13(2).Text = ""
    Me.FrameExpedPalet.Visible = True
    FrameExpedicion.Visible = False
    cmdSalir(10).Cancel = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdQuiCajaPalMan_Click(Index As Integer)
    If Me.lwNPalet.ListItems.Count = 0 Then Exit Sub
    
    
    If Index = 0 Then
        If Me.lwNPalet.SelectedItem Is Nothing Then Exit Sub
        If MsgBox("Quitar la caja: " & lwNPalet.SelectedItem.Text & "?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        Me.lwNPalet.ListItems.Remove Me.lwNPalet.SelectedItem.Index
        
    Else
        If MsgBox("Limpiar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        lwNPalet.ListItems.Clear
    End If
    txtNumCajas.Text = Me.lwNPalet.ListItems.Count
End Sub

Private Sub cmdQuitarCajaSistema_Click(Index As Integer)
  
    If Me.ListView1(3).ListItems.Count = 0 Then Exit Sub
    
    
    If Index = 1 Then If Me.ListView1(3).SelectedItem Is Nothing Then Exit Sub
    
    
    Select Case Index
    Case 0
        SQL = "Va a quitar la caja " & ListView1(3).SelectedItem.Text & vbCrLf & "¿Continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        
        ListView1(3).ListItems.Remove ListView1(3).SelectedItem.Index
        
    Case 1
        SQL = "Desea limpiar todas las cajas?"
        If MsgBox(SQL, vbExclamation + vbYesNo) = vbNo Then Exit Sub
        
        If MsgBox("Seguro?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        
        ListView1(3).ListItems.Clear
        PonerFoco Text17
        
    Case 2
        ' Adelante con el proceso
        SQL = "Finalizar el proceso de baja/perdida." & vbCrLf & "¿Continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        
        
        'falta###
        Screen.MousePointer = vbHourglass
        ProcesoDeBaja
        Screen.MousePointer = vbDefault
    End Select
End Sub

Private Sub cmdSalir_Click(Index As Integer)
    If Not Me.FrameSelect.Visible Then
     
         If Index = 3 Then
            'Esta creando NUEVO palet
            'Si tienen cajas ya metidas en el campo txtCajasEnNuevoPalet
            'Preguntamos si quiere salir o no
            If lwNPalet.ListItems.Count > 0 Then
                If MsgBox("Perderá los datos hasta el momento" & vbCrLf & "¿Continuar?", vbYesNo + vbDefaultButton2) = vbNo Then
                    Exit Sub
                Else
                    lwNPalet.ListItems.Clear
                End If
            End If
         End If
    
    
         If Index = 1 Then
              If Frame2.Visible Then
                Frame1.Visible = True
                Frame2.Visible = False
                Exit Sub
              End If
         End If
    
         If Index = 9 Then LimpiarCierrePalet
    
         If Index = 5 Then
            'EXPEDICION
            If Val(Label5(0).Tag) > 0 Then
                If MsgBox("Esta en proceso de expedición. Salir?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
            End If
        End If
         
         If Index = 10 Then
            If Text13(2).Visible Then
                'Quito el visible y si vuelve a opinchar salimos
                Text13(0).Text = ""
                Text13(2).Visible = False
                PonerFoco Text13(0)
                Me.cmdAceptarPaletVariosAlb.Visible = False
                
            Else
                Me.FrameExpedPalet.Visible = False
                FrameExpedicion.Visible = True
                 
            End If
            Exit Sub
         End If
         If Index = 11 Then Text17.Text = "": Text16.Text = "": ListView1(3).ListItems.Clear
            
         
         
         Conn.Execute "commit"   'tipo flush
         PonerFramePpal
         FramesNoVisibles
    Else
        If Me.FrameSegundasAcciones.Visible Then
            FrameSegundasAcciones.Visible = False
            FrameSelect.Visible = True
        Else
    
            Unload Me
        End If
    End If
    
End Sub

Private Sub PonerFramePpal()
    FrameSegundasAcciones.Visible = False
    FrameSelect.Visible = True
    ComprobarLuces
    Timer1.Enabled = True
End Sub

Private Sub cmdSelect_Click(Index As Integer)

    

    Me.FrameSelect.Visible = False
    Timer1.Enabled = False
    Select Case Index
    Case 0
        Me.Caption = "Linea produccion"
        Text1.Text = "": Text2.Text = "": Text3.Text = ""
        Frame1.Visible = True
        PonerFoco Text1
        cmdSalir(0).Cancel = True
    Case 1
        Me.Caption = "Paletización"
        FramePalet.Visible = True
        PonerFoco txtlecturaPelt
    Case 2, 8
        If Index = 2 Then
            Me.Caption = "Nuevo palet *"    'meto el * para saber que la caja TIENE QUE EXISTOR
        Else
            Me.Caption = "Palet prod. manual "  'las cajas no existen. Las dare de alta en este proceso
        End If
        Me.Text12.Visible = Index <> 2
        Label6(3).Visible = Index <> 2
        chkNoImprPaletManual.Visible = Index <> 2
        chkNoImprPaletManual.Value = 0
        
        FrameNPalet.Visible = True
        limpiar Me
        PonerFoco Text7
        cmdSalir(3).Cancel = True
    Case 3
        Me.Caption = "Poste"
        FramelecturaPostePalet.Visible = True
        PonerFoco txtlecturaPelt
        idTrazaAntiguo = -1
        Tiempo1 = Timer
        cmdSalir(4).Cancel = True
     Case 4
        Me.Caption = "Expedicion"
        FrameExpedicion.Visible = True
        cmdLimpExpedicion_Click
        Set RS = New ADODB.Recordset
        PonerFoco Text10
        cmdSalir(5).Cancel = True
    Case 5
        'Quitar cajas de un palet
        Me.Caption = "Ajuste cajas"
        FrameQuitaCaja.Visible = True
        cmdCreaPalet.Tag = 0 'Crear palet automatico
        txtCPalet(0).Text = "": txtCPalet(1).Text = "": txtCCaja.Text = "": txtObsCambioPalet.Text = ""
        Me.txtMoverUltimasNcajas.Text = "1"
        cmdSalir(6).Cancel = True
    Case 6
        Me.Caption = "Cierre palet"
        ponerOpcionesCierrePalet
        cmdSalir(8).Cancel = True
    Case 7
        Me.Caption = "Ajustar cajas"
        ponerOpcionesAjustePalet
        cmdSalir(9).Cancel = True
        
    Case 9
        FrameSelect.Visible = True
        Me.FrameSegundasAcciones.Visible = True
    Case 10
        Me.Caption = "Devolucion"
        Label5(6).Tag = 0
        
        FrameDevolucionOrdenCarga.Visible = True
        cmdSalir(11).Cancel = True
        PonerFoco Text14
        
    Case 11
        FrameDesapareceCaja.Visible = True
        Me.Caption = "Perdida o baja"
        cmdSalir(12).Cancel = True
        PonerFoco Text17
    End Select
End Sub

Private Sub cmdValidar_Click()
Dim SL As cLineaProCompo

    SQL = ""
    If Me.Text1.Text = "" Then
        SQL = "Lea etiqueta /linea "
    Else
        If Text2.Text = "" Or Text3.Text = "" Then
            SQL = "Lea etiqueta materia prima"
        Else
            If Cp Is Nothing Then
                SQL = "No se ha cargado los datos de produccion actual de la linea "
            Else
                If SubLinea Is Nothing Then SQL = "No se ha identificado la materia prima/auxiliar"
            End If
        End If
    End If
    If SQL <> "" Then
        MsgBox SQL, vbExclamation
        Exit Sub
    End If
    FrPrimeraProduccion.Visible = False
    If SubLinea.LoteMateria = "" Then
        'Primera asignacion LOTEs materia auxiliar
        FrPrimeraProduccion.Visible = True
        Text11(0).Text = SubLinea.NomArticCompo
        
        
        'Vere cuantas materias primas me falta tb
        Text11(1).Text = ""
        For idTrazaAntiguo = 1 To cLinPr.CuantasMP
            If cLinPr.DevuelveComponenteLinea(CInt(idTrazaAntiguo), SL) Then
                If SubLinea.codarticCompo = SL.codarticCompo Then
                    'Es el mismo. No hago nada
                    
                Else
                    If SL.LoteMateria = "" Then
                        'Tab falta este por asignar
                        Text11(1).Text = Text11(1).Text & ".-" & SL.NomArticCompo & vbCrLf
                        
                    End If
                End If
            End If
        Next
        Set SL = Nothing
    End If
    If Cp.NUmlote = SubLinea.LoteMateria Then
        'Mismo lote. Ha cambiado el bulto de materia auxiliar para embasar
        SQL = Right(Text2.Text, 3)
        SQL = "bulto = " & SQL & " AND id "
        SQL = DevuelveDesdeBD(conAri, "fechaulizada", "spartidaslin", SQL, Cp.idPartida, "N")
        If SQL <> "" Then
            SQL = "El bulto esta marcado como producido (" & SQL & ")"
            SQL = SQL & vbCrLf & "¿Continuar?"
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
    
        'Marcamos como en produccion y salimos
        SQL = "UPDATE spartidaslin set fechaulizada = " & DBSet(Now, "FH")
        SQL = SQL & " WHERE  bulto = " & Right(Text2.Text, 3) & " AND id = " & Cp.idPartida
        EjecutaSQL conAri, SQL, True
        
        'Nos salimos
        Frame1.Visible = True
        Frame2.Visible = False
        Text2.Text = "": Text3.Text = "": Text1.Text = ""
        PonerFoco Text1
    Else
        If IndexSublinea = 0 Then
            MsgBox "Aqui no deberia haber llegado", vbExclamation
            Exit Sub
        End If
        ' Lotes distintos
    
            TxtUD.Text = cLinPr.UnidadesCaja
            Text4.Text = cLinPr.CajasLeidasLector
            PonerDatosUdsCajas False
            Frame1.Visible = False
            Frame2.Visible = True
            BloquearTxt Text4, True
            BloquearTxt Text6, True
    End If
    

        

End Sub



Private Sub cmdVerComoVaExpedicion_Click()
    If Val(Label5(0).Tag) = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    Label5(4).Caption = Label5(0).Tag
    CargaExpedicion
    Me.FrameVerSituacionCarga.Visible = True
    FrameExpedicion.Visible = False
    cmdVolverExp.Cancel = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdVolverExp_Click()
    Me.FrameVerSituacionCarga.Visible = False
    FrameExpedicion.Visible = True
    PonerFoco Text10
    cmdSalir(5).Cancel = True
End Sub



Private Sub Form_Load()
   ' Me.Icon = frmppal.Icon
    PrimVez = True
    Me.Height = 320 * Screen.TwipsPerPixelX
    Me.Width = 240 * Screen.TwipsPerPixelX
    limpiar Me
    FramesNoVisibles
    PonerFramePpal
End Sub

Private Sub FramesNoVisibles()
    Me.Frame1.Visible = False
    Me.Frame2.Visible = False
    FramePalet.Visible = False
    FrameNPalet.Visible = False
    FramelecturaPostePalet.Visible = False
    FrameQuitaCaja.Visible = False
    FrameExpedicion.Visible = False
    Me.FrameCierrePalet.Visible = False
    Me.FrAjusteCajas.Visible = False
    Me.FrameDevolucionOrdenCarga.Visible = False
    FrameDesapareceCaja.Visible = False
    Me.Caption = "Aceites Morales (" & App.Major & "." & App.Minor & "." & App.Revision & ")"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'No cierra si no esta el de seleccionar
    If Not Me.FrameSelect.Visible Then
        FramesNoVisibles
        PonerFramePpal
        Cancel = 1
    Else
        Set Cp = Nothing
        Set cLinPr = Nothing
        Set SubLinea = Nothing
        Set RS = Nothing
    End If
End Sub

Private Sub imgPoste_Click()
    MsgBox "Poste sin leer cajas", vbExclamation
    
End Sub

Private Sub imgRepetidos_Click()
    MsgBox "Cajas repetidas", vbExclamation
End Sub

Private Sub ListView1_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    If Index = 1 Then Exit Sub
    If Not Me.FrameVerSituacionCarga.Visible Then Exit Sub  'Hasta que no sea visible nO va este
    CargarLineasAlbaranExpedicion Me.ListView1(0).SelectedItem, True
End Sub

Private Sub Text1_GotFocus()
    ConseguirFoco Text1, 3
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then PonerFoco Text2
End Sub

Private Sub Text1_LostFocus()
    SQL = ""
    If Text1.Text <> "" Then
        If Not IsNumeric(Text1.Text) Then
            SQL = "Campo numerico"
        Else
            If Val(Text1.Text) < 0 Or Val(Text1.Text) > 10 Then
                SQL = "Error leyendo linea(0-8)"
            Else
                If Not LeerLinea(CByte(Text1.Text)) Then
                    SQL = "Error leyendo produccion linea: " & Text1.Text
                    Set cLinPr = Nothing
                End If
            End If
        End If
    End If
    If SQL <> "" Then
        MsgBox SQL, vbExclamation
        Text1.Text = ""
        PonerFoco Text1
    Else
        Text2.Text = ""
        Text3.Text = ""
    End If
    If Text1.Text = "" Then
        LimpiarLineas
    End If
End Sub

Private Sub LimpiarLineas()
    Text2.Text = ""
    Text3.Text = ""
    IndexSublinea = 0
    Set Cp = Nothing
    Set SubLinea = Nothing
End Sub

Private Sub Text10_GotFocus()
   ConseguirFoco Text10, 3
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
    
    KEYpressGnral KeyAscii, 2, False
End Sub

Private Sub Text10_LostFocus()
Dim cadErr As String
Dim L As Integer
'Dim CadAux As String

 
    Text10.Text = Trim(Text10.Text)
    If Text10.Text = "" Then Exit Sub
    
    
    cadErr = ""
    '
    'Abril 2013
    'SSCC   ->  Ver [SSCC] en impresion de palets  EAN128
    ' LA etiquetas se leen de palet de oliveline se leen C1003841259400000592423300000387
    L = InStr(1, Text10.Text, "C10038412594")   'este es el barrado
    
    If L = 0 Then
        L = InStr(Text10.Text, "38412594")
    End If
    
    If L > 0 Then
        'En la expedicion dejare que lea los palets etiquetados
        'en formato SSCC.  Con lo cual
        If L = 4 Then
            'Lo que estaba antes
            SQL = Mid(Text10.Text, L + 4)
        Else
            SQL = Mid(Text10.Text, L)
        End If
        'El fin del sscc es 10 caracteres despues
        If Len(SQL) > 18 Then SQL = Mid(SQL, 1, 18)
      
      
      
        
        'Comprobacion
        
        If Len(SQL) <> 18 Then
            cadErr = "Longitud incorrecta"
        Else
            If Not IsNumeric(SQL) Then
                cadErr = "Campo numerico"
            Else
                If Mid(SQL, 1, 8) <> "38412594" Then   'los de morales empiezan asi
                    cadErr = "no pertenece a Aceites Morales"
                Else
                    '(00)38412594xxxxxxxxxC  valen las ultimas 7 x
                    SQL = Right(SQL, 8) 'los utilmos 8
                    SQL = Left(SQL, 7)  'los primeros 7
                    'AHORA tenemos el palet
                    L = Val(SQL)
                    
                    'Ahora lo formateo para nosotros, con len 10
                    SQL = "0" & Format(L, "00000000") & "0"
                                        
                    TextoLecturasExpedicion "SSCC " & Mid(Text10.Text, 1, 12) & "... >  " & L
        
                    Text10.Text = SQL
        
                End If
            End If
            
        End If
        If cadErr <> "" Then
            cadErr = "SSCC " & cadErr
            TextoLecturasExpedicion Mid(Text10.Text, 14) & "  Error " & cadErr
            Text10.Text = ""
            PonerFoco Text10
            Exit Sub
        End If
        SQL = ""
    End If
    
    
    
    
    If Not IsNumeric(Text10.Text) Then
        cadErr = "Campo numerico"
        
    Else
    
        'Si la longitud no es 11,12 o 13
        L = Len(Text10.Text)
        If L < 10 Or L > 13 Then
            cadErr = "Longitud incorrecta"
        Else
            '
            Select Case L
            Case 10
                'Lectura PALET
                cadErr = LeerPaletExpedicion
            Case 11
                'ID orden carga
                cadErr = lecturaOrdenExpedicion
            Case 12
                'ID albaran
                'Le toca introducir albaran
                cadErr = lecturaAlbaranExpedicion
                
            Case 13
                'CAJA
                cadErr = LeerCajaExpedicion
                
            End Select
        End If
    End If
    
    If cadErr <> "" Then
        'MsgBox CadErr, vbExclamation
        TextoLecturasExpedicion Text10.Text & " Error " & cadErr
    Else
        
        If L = 13 Then
            SQL = "Caja OK"
        Else
            If L = 12 Then
                SQL = "  Albaran OK"
            ElseIf L = 11 Then
                SQL = "  Orden OK"
            Else
                SQL = " Palet OK"
            End If
        End If
        TextoLecturasExpedicion "#" & Text10.Text & " " & SQL
    End If
        '
    Text10.Text = ""
    PonerFoco Text10
    
    
End Sub



Private Sub Text12_GotFocus()
    ConseguirFoco Text12, 3
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
     KEYpressGnral KeyAscii, 2, False
End Sub

Private Sub Text12_LostFocus()
Dim L As Integer
Dim cadErr As String
Dim SQL As String

    If Text12.Text = "" Then Exit Sub




        '--------------------------------------------------------- ZC1003841259400000592423300000387
        '
    'Junio 2014
    'SSCC   ->  Ver [SSCC] en impresion de palets  EAN128
    ' LA etiquetas se leen de palet de oliveline se leen C1003841259400000592423300000387
    L = InStr(1, Text12.Text, "C10038412594")   'este es el barrado
    If L > 0 Then
        'En la expedicion dejare que lea los palets etiquetados
        'en formato SSCC.  Con lo cual
        SQL = Mid(Text12.Text, L + 4)
        
        'El fin del sscc es 10 caracteres despues
        If Len(SQL) > 18 Then SQL = Mid(SQL, 1, 18)
      
        
        'Comprobacion
        cadErr = ""
        If Len(SQL) <> 18 Then
            cadErr = "Longitud incorrecta"
        Else
            If Not IsNumeric(SQL) Then
                cadErr = "Campo numerico"
            Else
                If Mid(SQL, 1, 8) <> "38412594" Then   'los de morales empiezan asi
                    cadErr = "no pertenece a Aceites Morales"
                Else
                    '(00)38412594xxxxxxxxxC  valen las ultimas 7 x
                    SQL = Right(SQL, 8) 'los utilmos 8
                    SQL = Left(SQL, 7)  'los primeros 7
                    'AHORA tenemos el palet
                    L = Val(SQL)
                    
                    'Ahora lo formateo para nosotros, con len 10 y 1 al incio y al final
                    SQL = "1" & Format(L, "00000000") & "1"
                                        
                    'TextoLecturasExpedicion "SSCC " & Mid(Text12.Text = "", 1, 12) & "... >  " & L
        
                    Text12.Text = SQL
                End If
            End If
            
        End If
        If cadErr <> "" Then
            cadErr = "SSCC " & cadErr
            'MsgBox SQL, vbExclamation
            Text12.Text = ""
            PonerFoco Text12
            
            Exit Sub
        End If
        SQL = ""
    End If

    
    



    SQL = EtiquetaPaletCorrecta2(Text12.Text)
    If SQL = "" Then
        
            
            Set cPal = New CPalet
            SQL = Mid(Text12.Text, 2, 8) 'los ocho centrales
            If Not cPal.Leer(CLng(SQL)) Then
                SQL = "No existe "
               
            Else
                SQL = ""
            End If
            
            
    End If
    If SQL <> "" Then
        MsgBox SQL, vbExclamation
        Text12.Text = ""
        PonerFoco Text12
    End If
     Set cPal = Nothing
End Sub

Private Sub Text13_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, False
End Sub

Private Sub Text13_LostFocus(Index As Integer)
Dim C As String
Dim Aux2 As String
Dim L As Long
Dim txSCC As String
     If Index <> 0 Then Exit Sub 'Solo el index=0
    
    
     Text13(0).Text = Trim(Text13(0).Text)
     If Text13(0).Text = "" Then Exit Sub
     
     
    'Abril 2017
    'Si es lectura PALET SSCC EAN128, lo transforammos a nuestra lectura
    'SSCC   ->  Ver [SSCC] en impresion de palets  EAN128
    ' LA etiquetas se leen de palet de oliveline se leen C1003841259400000592423300000387
    L = InStr(1, Text13(0).Text, "C10038412594")   'este es el barrado
    If L > 0 Then
        txSCC = Text13(0).Text
        'En la expedicion dejare que lea los palets etiquetados
        'en formato SSCC.  Con lo cual
        Aux2 = Mid(Text13(0).Text, L + 4)
        
        'El fin del sscc es 10 caracteres despues
        If Len(Aux2) > 18 Then Aux2 = Mid(Aux2, 1, 18)
      
         
     
     
        'Comprobacion
        If Len(Aux2) <> 18 Then
            Text13(1).Text = "Longitud incorrecta"
        Else
            If Not IsNumeric(Aux2) Then
                Text13(1).Text = "Campo numerico"
            Else
                If Mid(Aux2, 1, 8) <> "38412594" Then   'los de morales empiezan asi
                    Text13(1).Text = "no pertenece a Aceites Morales"
                Else
                    '(00)38412594xxxxxxxxxC  valen las ultimas 7 x
                    Aux2 = Right(Aux2, 8) 'los utilmos 8
                    Aux2 = Left(Aux2, 7)  'los primeros 7
                    'AHORA tenemos el palet
                    L = Val(Aux2)
                    
                    'Ahora lo formateo para nosotros, con len 10 y 1 al incio y al final
                    Aux2 = "1" & Format(L, "00000000") & "1"
                                        
                    Text13(1).Text = "SSCC " & Mid(Text13(0).Text, 1, 12) & "... >  " & L
        
                    Text13(0).Text = Aux2
        
                End If
            End If
            
        End If
      End If
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     If Not IsNumeric(Text13(0).Text) Then
        Text13(1).Text = "Campo no numerico:" & Text13(0).Text & vbCrLf & txSCC
        
     Else
    
        'Si la longitud no es 11,12 o 13
        
        If Len(Text13(0).Text) <> 10 Then
            Text13(1).Text = "Longitud incorrecta"
        Else
            Text13(1).Text = ""
            Text13(1).Text = LeerPaletExpedicionParaVariosAlbaranes(True)
            If Text13(1).Text = "" Then
                Text13(1).Text = LeerPaletExpedicionParaVariosAlbaranes(False)
            End If
        End If
    End If
    If Text13(1).Text <> "" Then
        Text13(0).Text = ""
        PonerFoco Text13(0)
    End If
End Sub



'-------------------------------
' Devolucion mercacnia expedida
Private Sub Text14_GotFocus()
    ConseguirFoco Text14, 3
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
     KEYpressGnral KeyAscii, 2, False
End Sub



Private Sub Text14_LostFocus()
Dim cadErr As String
'Dim L As Integer
'Dim CadAux As String

 
    Text14.Text = Trim(Text14.Text)
    If Text14.Text = "" Then
        Text15.Text = ""
        Exit Sub
    End If
    
    
    cadErr = ""
    If Not IsNumeric(Text14.Text) Then
        cadErr = "Campo numerico"
        
    Else
    
        'Si la longitud no es 11,12 o 13
        'L = Len(Text14.Text)
        If Len(Text14.Text) <> 13 Then
            cadErr = "Longitud incorrecta"
        Else
            '
'            Select Case L
'            Case 10
'                'Lectura PALET
'                cadErr = "No se pueden leer palets"
'            Case 11
'                'ID orden carga
'                cadErr = lecturaOrdenCargaDevolucion
'            Case 12
'                'ID albaran
'                'Le toca introducir albaran
'                cadErr = lecturaAlbaranExpedicion
'
'            Case 13
'                'CAJA


                Set RS = New ADODB.Recordset
                
                cadErr = LeerCajaDevolucion
                
                Set RS = Nothing
                
'
'            End Select
        End If
    End If
    
    If cadErr <> "" Then
        Text15.Text = "Err " & cadErr
        
    Else
        
'        If L = 13 Then
'            SQL = "Caja OK"
'        Else
'            If L = 12 Then
'                SQL = "  Albaran OK"
'            ElseIf L = 11 Then
'                SQL = "  Orden OK"
'            Else
'                SQL = " Palet OK"
'            End If
'        End If
        Text15.Text = "# Caja OK"
    End If
        '
    Text14.Text = ""
    PonerFoco Text14
    
    idTrazaAntiguo = 0  'esta variable la es gnral y la utilizo en la funcion
End Sub






Private Sub Text2_GotFocus()
    ConseguirFoco Text2, 3
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, False
End Sub

Private Sub Text2_LostFocus()
    Text2.Text = Trim(Text2.Text)
    If Text2.Text = "" Then Exit Sub
    If Text1.Text = "" Then
        MsgBox "FALTA LA LINEA", vbExclamation
        Text2.Text = ""
        Text3.Text = ""
        PonerFoco Text1
    End If
    If Not IsNumeric(Text2.Text) Then
        MsgBox "Error en lectura etiqueta", vbExclamation
        Text2.Text = ""
        Text3.Text = ""
        PonerFoco Text2
        Exit Sub
    End If
    
    'AHOR
    If Not PonerDatosDesdelecturaEtiqueta Then
        LimpiarLineas
        PonerFoco Text2
    End If
    
End Sub

Private Function LeerLinea(KLinea As Byte) As Boolean
Dim i As Byte


    LeerLinea = False
    Set RS = New ADODB.Recordset
    SQL = "select prodlin.codigo,prodlin.idlin ,lotetraza from prodlin,prodtrazlin  where prodlin.codigo= prodtrazlin.codigo AND prodlin.idlin = prodtrazlin.idlin"
    SQL = SQL & " and lineaprod = " & KLinea & " and estado >0 and estado<10 ORDER BY lotetraza DESC"  'Pq puede que haya varios cambios de trazabilidad para la misma linea
    RS.Open SQL, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    i = 0
    If Not RS.EOF Then
        'UNO tiene seguro
        
        Set cLinPr = New cLineaProduccion
        cLinPr.LeerDesdeTrazabilidad RS!Codigo, RS!idlin, KLinea, RS!lotetraza
        i = 1
        
        'veremos que SOLO hay una linea en marcha
        Do
            RS.MoveNext
        
            If Not RS.EOF Then
                'Si el codigo y prodilin es el mismo, es que solo hay una produccion
                If RS!Codigo <> cLinPr.CodProduccion Or cLinPr.idLiProd <> RS!idlin Then
                    'MAAAAAAAl
                    'Hay mas de una produccion en la linea
                    Set cLinPr = Nothing
                    i = 2
                End If
            End If
        Loop Until RS.EOF
    End If
    RS.Close
    SQL = ""
    If i = 1 Then LeerLinea = True
End Function

Private Function PonerDatosDesdelecturaEtiqueta() As Boolean
Dim N As Long
Dim Aux As String
Dim Seguir As Boolean
Dim NOEXISTE As Boolean

Dim RegularizarStockLotes As String  'FALTA### Esto hay que "hacerlo"


    On Error GoTo EPonerDatosDesdelecturaEtiqueta
    PonerDatosDesdelecturaEtiqueta = False
    
    If Len(Text2.Text) <> 9 Then
        MsgBox "Codigo etiqueta incorrecto", vbExclamation
        Exit Function
    End If


    'Los tres utlimos digitos hacen referencia al secuencial de etiqueta
    'Si es el cero estan leyendo la etiqueta del albaran
    If Right(Text2.Text, 3) = "000" Then
        MsgBox "Etiqueta incorrecta(Albaran)", vbExclamation
        Exit Function
    End If
    'Los tres utlimos es el identificador de la etiqueta. El resto es el idPartida/lote
    SQL = Mid(Text2.Text, 1, Len(Text2.Text) - 3)
    N = CLng(SQL) 'Si da error de desboramienot ya hablaremos
    SQL = ""
    Set Cp = New cPartidas
    If Not Cp.Leer(N) Then
        'La etiqueta no pertenece a ninunga partida
        MsgBox "No pertence a ninguna partida", vbExclamation
        
    Else
        'Correcto. Pertenece a una partida
        'QUe  hacemos ahora.....
        'Facil. Comprobamos si el articulo es un subcomponente
        For N = 1 To cLinPr.CuantasMP
            If cLinPr.DevuelveComponenteLinea(CInt(N), SubLinea) Then
                If SubLinea.codarticCompo = Cp.codartic Then
                    IndexSublinea = N 'para cuando mandemos el cierre de lote
                    Exit For
                Else
                    Set SubLinea = Nothing
                End If
            End If
        Next
        
        
        If N > cLinPr.CuantasMP Then
            SQL = "El articulo no pertence a lo que se esta produciendo en esta linea"
            MsgBox SQL, vbExclamation
            Exit Function
        End If
        
        'Veremos si para esta etiqueta esta marcado la fecha de produccion. Sginificaria que
        'ya ha sido utilizada
        SQL = "fechaulizada"
        N = Val(DevuelveDesdeBD(conAri, "id", "spartidaslin", "bulto = " & Right(Text2.Text, 3) & " AND id", Cp.idPartida, "N", SQL))
        
        
        
        If N > 0 Then
            Seguir = True
            
            If SQL <> "" Then
                SQL = "El bulto ya ha sido utilizado: " & Format(SQL, "dd/mm/yyyy hh:nn:ss")
                SQL = SQL & vbCrLf & vbCrLf & "¿Continuar?"
                If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Seguir = False
            End If
            NOEXISTE = False
        Else
            NOEXISTE = True
            
            
            Seguir = True
        End If
            'MAYO 2012
            If Seguir Then
            
                Set RS = New ADODB.Recordset
                
                RegularizarStockLotes = ""
                If SubLinea.LoteMateria <> Cp.NUmlote Then
                    'Cambia de materia auxiliar. Veremos si ya no quedan de ese
                    SQL = "Select * from spartidas where  codartic = " & DBSet(Cp.codartic, "T") & " AND numlote = " & DBSet(SubLinea.LoteMateria, "T")
                    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    If Not RS.EOF Then
                        If RS!cantotal <> 0 Then
                            RegularizarStockLotes = "Deberia regularizar existencia " & vbCrLf & "Lote: " & SubLinea.LoteMateria & "  Cantidad: " & RS!cantotal
                            MsgBox RegularizarStockLotes, vbInformation
                        End If
                    End If
                    RS.Close
                End If
            
                'Vemos si es un cambio de lote, si existe uno mas antiguo que es el que deberia coger
                If SubLinea.LoteMateria <> Cp.NUmlote Then
                    'HA CAMBIADO EL NUMERO DE LOTE
                    'select concat(numlote,"|",fecha,"|",cantotal,"|") from spartidas where codartic='002400180306'
                    'and numlote <> '429-A' and cantotal>0 order by fecha asc
                    
                    SQL = "select * from spartidas left join sprove on spartidas.codprove=sprove.codprove where codartic = " & DBSet(Cp.codartic, "T") & " AND numlote <> " & DBSet(SubLinea.LoteMateria, "T")
                    SQL = SQL & "  and cantotal>0 order by fecha asc "
                    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    If Not RS.EOF Then
                        'Si el lote NO es el que ha seleccionado, significa que hay uno mas antiguo
                        
                        If RS!NUmlote <> Cp.NUmlote Then
                            SQL = "Lote mas antiguo: " & vbCrLf & "Lote: " & RS!NUmlote & "   Uds:" & RS!cantotal & vbCrLf
                            SQL = SQL & "Fecha: " & RS!Fecha & " - Alb:" & DBLet(RS!NumAlbar, "T") & vbCrLf
                            SQL = SQL & "Prov: " & RS!codProve & " " & DBLet(RS!nomprove, "T")
                            SQL = SQL & vbCrLf & "¿Continuar?"
                            If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Seguir = False
                        End If
                    End If
                    RS.Close
                    
                    
                    
                    
                    If Seguir Then
                        'Ha cambiado el lote. Comprobar existencias
                        'Veremos si quedaban entodavia
                        
                        
                    End If
                
                End If
                Set RS = Nothing
            End If
            
            If Seguir Then
                'OKKKKKK
                'Perfecto, el bulto NO ha sido utilizado
                Text3 = "Bulto: " & Right(Text2.Text, 3)
                If NOEXISTE Then Text3.Text = Text3.Text & "      NO existe en BD"
                Text3.Text = Text3.Text & vbCrLf
                'articulo del bulto
                SQL = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", Cp.codartic)
                Text3.Text = Text3.Text & "Art: " & SQL & vbCrLf
                Text3.Text = Text3.Text & "Albaran: " & Cp.NumAlbar & "   Lote: " & Cp.NUmlote & vbCrLf
                If Cp.codProve > 0 Then
                    SQL = DevuelveDesdeBD(conAri, "nomprove", "sprove", "codprove", Cp.codProve)
                    Text3.Text = Text3.Text & "Prov: " & Cp.codProve & "   " & SQL & vbCrLf
                End If
                
                'Articulo en produccion
                SQL = String(15, "-")
                Text3 = Text3 & vbCrLf & SQL & "Produciendo" & SQL & vbCrLf
                Text3.Text = Text3.Text & "Articulo: " & cLinPr.codartic & vbCrLf
                Text3.Text = Text3.Text & " " & cLinPr.NomArtic
                
                
                
                PonerDatosDesdelecturaEtiqueta = True
                
            End If
        
            
        
        
        
    End If
EPonerDatosDesdelecturaEtiqueta:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set RS = Nothing
    SQL = ""
End Function



Private Sub Text4_GotFocus()
    ConseguirFoco Text4, 3
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then PonerFoco Text6
End Sub

Private Sub Text4_LostFocus()
    If Not PonerFormatoEntero(Text4) Then
        Text4.Text = ""
        Text6.Text = ""
    Else
        PonerDatosUdsCajas False
    End If
End Sub


Private Sub Text6_GotFocus()
    ConseguirFoco Text6, 3
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then PonerFocoBtn Me.cmdAceptarCambioLote
End Sub

Private Sub Text6_LostFocus()
    If Not PonerFormatoDecimal(Text6, 3) Then
        Text4.Text = ""
        Text6.Text = ""
    Else
        PonerDatosUdsCajas True
    End If
End Sub

Private Sub PonerDatosUdsCajas(DesdeUds As Boolean)
Dim L As Long
    If DesdeUds Then
        L = ImporteFormateado(Text6.Text) \ CInt(Me.TxtUD.Text)
        If (ImporteFormateado(Text6.Text) Mod CInt(Me.TxtUD.Text)) > 0 Then L = L + 1
        Text4.Text = Format(L, "##,##0")
    Else
        L = Val(Text4.Text) * CInt(Me.TxtUD.Text)
        Text6.Text = Format(L, FormatoCantidad)
    End If
        
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub Text7_LostFocus()
Dim C As String
Dim Aux As String
Dim OK As Boolean
Dim EsUnaProduccionManual As Boolean
    
    Text7.Text = Trim(Text7.Text)
    If Text7.Text = "" Then Exit Sub
    
    'cAJA PARA AÑADIR A UN PALET MANUAL
    If Len(Text7.Text) < 13 Then
        'MsgBox "Longitudad etiqueta incorrecta", vbExclamation
        Text5.Text = "Longitudad etiqueta incorrecta " & Text7.Text
        Text7.Text = ""
        PonerFoco Text7
        Exit Sub
    End If
    If Not IsNumeric(Text7.Text) Then
        'MsgBox "Campo numérico", vbExclamation
        Text5.Text = "Campo numerico " & Text7.Text
        Text7.Text = ""
        PonerFoco Text7
        Exit Sub
    End If
    
        
    'EsUnaProduccionManual :  LAS cajas NO estan dadas de alta. Sera en este proceso cuando lo hagamos
    EsUnaProduccionManual = InStr(1, Me.Caption, "*") = 0  'el caption no lleva el *
    
        
        'Dividimos la etiqueta leida en 2
        'los cinco ultimos son el IDCAJa
        'el resto ID trza
        'select * from prodcajas where lotetraza=27 and idcaja=5
        C = Mid(Text7, 1, Len(Text7.Text) - 5)
        C = "lotetraza = " & C & " AND idcaja"
        Aux = "lotetraza"
        C = DevuelveDesdeBD(conAri, "idpalet", "prodcajas", C, Right(Text7.Text, 5), "N", Aux)
        OK = False
        If Aux = "lotetraza" Then
        
            If EsUnaProduccionManual Then
                'Veremos si esta en marcha el lote
                Aux = "lineaprod"  'SOLO EN LA 8 o 9
                C = DevuelveDesdeBD(conAri, "codigo", "prodtrazlin", "lotetraza", Left(Text7.Text, 8), "N", Aux)
                If C = "" Then
                    'No existe el lote
                    Aux = "No existe numero trazabilidad"
                Else
                    If Aux = "8" Or Aux = "9" Then
                        Aux = InsertarEnNuevosPalets(True)
                    Else
                        'OK intentamos insertaer
                        Aux = "No es la linea prod. manual"
                        
                    End If
                End If
                
            Else
                Aux = Text7.Text & "   caja no encon."
            End If
        Else
                
            'Ok esta en el sistema
            If EsUnaProduccionManual Then
                'En las producciones manuales(de la linea de produccion 8(manual)
                'las cajas NO deben existir. las crearemos desde aqui
                
                Aux = "Ya existe la caja: " & Text7.Text
                
            Else
                If C <> "" Then
                    'CAJA YA asignada
                    Aux = txtCajas.Text & " EN PALET " & C  'IGUAL DEBERIAMOS DEJARLA Y DESPUES DESASIGNARLA
                Else
                    
                            
                    'OK la asigno
                    Aux = InsertarEnNuevosPalets(False)
                    
                End If
            End If
        End If
   
        Text5.Text = Aux
        Text7.Text = ""
        PonerFoco Text7
    


End Sub

'Cuando el palet es manual me preguntar las cajas
'Insertaremos a mano las cajas que falten
Private Function InsertarEnNuevosPalets(EsPaletManual As Boolean) As String
Dim IT As ListItem
Dim i As Integer


    On Error Resume Next
    InsertarEnNuevosPalets = ""
    'Recorrp para que no exista
    For i = 1 To lwNPalet.ListItems.Count
        If Text7.Text = lwNPalet.ListItems(i).Text Then
            InsertarEnNuevosPalets = "ya introducida"
            Exit Function
        End If
    Next
    
    If EsPaletManual Then
            i = 0
            Do
                SQL = Val(Right(Text7, 5))
                SQL = InputBox("Numero de cajas a insertar", , SQL)
                If SQL <> "" Then
                    If IsNumeric(SQL) Then
                        If Val(SQL) > Val(Right(Text7, 5)) Then
                            MsgBox "Maximo " & Right(Text7, 5), vbExclamation
                        Else
                            i = 1
                        End If
                    Else
                        MsgBox "Numerico", vbExclamation
                    End If
                Else
                    i = 1
                End If
                
                
                
            Loop Until i = 1
            
            If SQL <> "" Then
                'ok insetare cad cajas
                i = Val(SQL)
                idTrazaAntiguo = Val(Right(Text7, 5))
      
                While i <> 0
                    SQL = Mid(Text7.Text, 1, 8) & Format(idTrazaAntiguo, "00000")
                    Set IT = Me.lwNPalet.ListItems.Add(, "C" & SQL, SQL)
                    idTrazaAntiguo = idTrazaAntiguo - 1
                    i = i - 1
                Wend
                
            End If
            
            
        
    Else
        Set IT = Me.lwNPalet.ListItems.Add(, "C" & Text7.Text, Text7.Text)
        If Err.Number <> 0 Then
            InsertarEnNuevosPalets = Err.Description
            Err.Clear
        End If
    End If
    txtNumCajas.Text = Me.lwNPalet.ListItems.Count
End Function

Private Sub Text8_GotFocus()
    ConseguirFoco Text8, 3
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then PonerFoco Me.txtCajas
End Sub

Private Sub Text8_LostFocus()
    SQL = ""
    If Text8.Text <> "" Then
        If Not IsNumeric(Text8.Text) Then
            SQL = "Campo numerico"
        Else
            If Val(Text8.Text) < 1 Or Val(Text8.Text) > 4 Then
                SQL = "Error leyendo linea paletizacion(1-4)"
            Else
                'SQL = "select idpalet from prodpalets where  LineaPeletiza = " & KLinea & " and fhFin is null "  '
                SQL = DevuelveDesdeBD(conAri, "idpalet", "prodpalets", "fhFin is null AND LineaPeletiza", Text8.Text)
                If SQL = "" Then
                    SQL = "No hay nada en la linea de paletizacion"
                Else
                    Me.txtIdPal.Text = SQL
                    SQL = ""
                    Set cPal = New CPalet
                    If Not cPal.Leer(Me.txtIdPal.Text) Then SQL = "Error leyendo PALET"
                    
                    
                End If
                Text8.Text = Val(Text8.Text)
            End If
        End If
    End If
    If SQL <> "" Then
        MsgBox SQL, vbExclamation
        Text8.Text = ""
        Me.txtIdPal.Text = ""
        PonerFoco Text8
    Else
        
        txtCajas.Text = ""
        Me.txtObservaCajas.Text = ""
    End If
    If Text8.Text = "" Then
        Me.txtIdPal.Text = ""
        txtCajas.Text = ""
        Me.txtObservaCajas.Text = ""
    End If
    If Me.txtIdPal.Text = "" Then Set cPal = Nothing
End Sub



Private Sub Timer1_Timer()
    ComprobarLuces
End Sub

Private Sub txtCajaCierre_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.FrAjusteCajas.Visible Then
            
            PonerFoco Me.txtCierrPalet(4)
        Else
            PonerFocoBtn cmdCerrarElPalet
        End If
    End If
End Sub

Private Sub txtCajaCierre_LostFocus(Index As Integer)
    txtCajaCierre(Index).Text = Trim(txtCajaCierre(Index).Text)
    txtCajaCierre(Index).Tag = 0
    SQL = ""
    If txtCajaCierre(Index).Text <> "" Then
        If Len(txtCajaCierre(Index).Text) <> 13 Then
                SQL = "Longitud etiqueta incorrecta"
            
            Else
                If Not IsNumeric(txtCajaCierre(Index).Text) Then
                    SQL = "Campo numerico"
                Else
                    '*****************************************
                    If Index = 0 Then
                    
                    
                        'Vere si existe la traza es la que estamos paletizando
                        
                        SQL = Mid(txtCajaCierre(Index).Text, 1, 8)
                        SQL = Val(SQL) & "|"
                        If InStr(1, cPal.TrazabilidadPaletizando, SQL) = 0 Then
                            SQL = "No pertenece a este palet"
                            
                        Else
                            
                            'SQL = DevuelveDesdeBD(conAri, "idpalet", "prodcajas", SQL, Mid(txtCajaCierre.Text, 9), "N")
                            SQL = "lotetraza = " & Mid(txtCajaCierre(Index).Text, 1, 8) & " AND idcaja  "
                            SQL = DevuelveDesdeBD(conAri, "idpalet", "prodcajas", SQL, Mid(txtCajaCierre(Index).Text, 9), "N")
                            
                            'Veremos si la caja esta en el palet que toca o no
                            If SQL = "" Then
                                
                                txtCajaCierre(Index).Tag = 1 'HAy QUE DAR DE ALTA ESTA CAJA...seguro
                            Else
                                If Val(SQL) <> Val(cPal.ID) Then
                                    SQL = "Caja no pertenece al palet origen"
                                Else
                                    'SIP OK. vamos a quitarla del palet.
                                    SQL = ""
                                End If
                            End If
                        End If
                        
                    Else
                        'INDEX=1  Cierre lote traza
                        
                    End If
                 End If 'de numerico
            End If  '<>13
        
    End If 'de <>""
    If SQL <> "" Then
        MsgBox SQL, vbExclamation
        Me.txtCajaCierre(Index).Text = ""
        PonerFoco txtCajaCierre(Index)
    End If
    
End Sub

Private Sub txtCajas_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub txtCajas_LostFocus()
Dim C As String
Dim Aux As String
Dim OK As Boolean
Dim LineaPal As String

    If txtCajas.Text = "" Then Exit Sub
    If Text8.Text = "" Then
        txtCajas.Text = ""
        MsgBox "Falta palet"
        PonerFoco Text8
        Exit Sub
    End If
    If Len(txtCajas.Text) < 13 Then
        'MsgBox "Longitudad etiqueta incorrecta", vbExclamation
        txtObservaCajas.Text = "Longitudad etiqueta incorrecta"
        txtCajas.Text = ""
        PonerFoco txtCajas
        Exit Sub
    End If
    If Not IsNumeric(txtCajas.Text) Then
        'MsgBox "Campo numérico", vbExclamation
        txtObservaCajas.Text = "Campo numerico"
        txtCajas.Text = ""
        PonerFoco txtCajas
    Else
        'Dividimos la etiqueta leida en 2
        'los cinco ultimos son el IDCAJa
        'el resto ID trza
        'select * from prodcajas where lotetraza=27 and idcaja=5
        C = Mid(txtCajas, 1, Len(txtCajas.Text) - 5)
        C = "lotetraza = " & C & " AND idcaja"
        Aux = "lotetraza"
        C = DevuelveDesdeBD(conAri, "idpalet", "prodcajas", C, Right(txtCajas.Text, 5), "N", Aux)
        OK = False
        If Aux = "lotetraza" Then
            Aux = txtCajas.Text & "   caja no encon."
        Else
            
            'Ok esta en el sistema
            If C <> "" Then
                'CAJA YA asignada
                Aux = txtCajas.Text & " ya asig" & C
            Else
                
                'Ya tengo el lote de traza. Vere si esta produciendo asi o no
                'Es decir. Vere si en la linea que me han indicado se esta paletizando todo esto
                
                'select * from prodtrazlin where lotetraza=1 lineaprod
                LineaPal = DevuelveDesdeBD(conAri, "lineaprod", "prodtrazlin", "lotetraza", Aux)
                If LineaPal = "" Then
                    Aux = "Error obteniedo linea produccion del articulo"
                Else
                    If Not cPal.LineasProd(CInt(LineaPal - 1)) Then
                        Aux = "NO en la linea  de produccion asociada" & LineaPal
                    Else
                        
                        'OK la asigno
                        C = "UPDATE prodcajas set idpalet = " & txtIdPal.Text & " WHERE idcaja=" & Right(txtCajas.Text, 5) & " AND lotetraza=" & Aux
                        If EjecutaSQL(conAri, C) Then
                            Aux = "     OK. "
                            OK = True
                        Else
                            Aux = "ERROR SQL"
                        End If
                    End If
                End If
            End If
        End If
        'If Not Ok Then
        '    txtObservaCajas.Text = Aux & vbCrLf & txtObservaCajas.Text
        'Else
            txtObservaCajas.Text = Aux
        'End If
        txtCajas.Text = ""
        PonerFoco txtCajas
    End If
End Sub

Private Sub txtCCaja_KeyPress(KeyAscii As Integer)
    
     KEYpressGnral KeyAscii, 2, False
End Sub

Private Sub txtCCaja_LostFocus()
    ProcesaCajaCambioPalet
End Sub

Private Sub txtCierrPalet_GotFocus(Index As Integer)
    If Index = 2 Or Index = 4 Then ConseguirFoco txtCierrPalet(Index), 3
    
End Sub

Private Sub txtCierrPalet_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, False
End Sub

Private Sub txtCierrPalet_LostFocus(Index As Integer)
    
    txtCierrPalet(Index).Text = Trim(txtCierrPalet(Index).Text)
    If txtCierrPalet(Index).Text = "" Then Exit Sub
    
    
    'Uds para cerrar el palet
    If Index = 2 Or Index = 4 Then
        If Not PonerFormatoEntero(txtCierrPalet(Index)) Then
            txtCierrPalet(Index).Text = ""
        Else
            If Val(txtCierrPalet(Index).Text) < 0 Then
                MsgBox "Cantidades postivas", vbExclamation
                txtCierrPalet(Index).Text = ""
            End If
        End If
    End If
End Sub

Private Sub txtCPalet_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = 13 Then ProcesaPaletCambio Index
End Sub
    
Private Sub ProcesaPaletCambio(Index As Integer)
Dim L As Integer
Dim cadErr As String


    SQL = ""
    txtCPalet(Index).Text = Trim(txtCPalet(Index).Text)
    If txtCPalet(Index).Text = "" Then
        
        'El campo 1 puede estar vacio
        If Index = 1 Then PonerFoco Me.txtCCaja
        Exit Sub
    End If
    
    
    
    '---------------------------------------------------------
        '
    'Junio 2014
    'SSCC   ->  Ver [SSCC] en impresion de palets  EAN128
    ' LA etiquetas se leen de palet de oliveline se leen C1003841259400000592423300000387
    L = InStr(1, txtCPalet(Index).Text, "C10038412594")   'este es el barrado
    If L > 0 Then
        'En la expedicion dejare que lea los palets etiquetados
        'en formato SSCC.  Con lo cual
        SQL = Mid(txtCPalet(Index).Text, L + 4)
        
        'El fin del sscc es 10 caracteres despues
        If Len(SQL) > 18 Then SQL = Mid(SQL, 1, 18)
      
        
        'Comprobacion
        
        If Len(SQL) <> 18 Then
            cadErr = "Longitud incorrecta"
        Else
            If Not IsNumeric(SQL) Then
                cadErr = "Campo numerico"
            Else
                If Mid(SQL, 1, 8) <> "38412594" Then   'los de morales empiezan asi
                    cadErr = "no pertenece a Aceites Morales"
                Else
                    '(00)38412594xxxxxxxxxC  valen las ultimas 7 x
                    SQL = Right(SQL, 8) 'los utilmos 8
                    SQL = Left(SQL, 7)  'los primeros 7
                    'AHORA tenemos el palet
                    L = Val(SQL)
                    
                    'Ahora lo formateo para nosotros, con len 10 y 1 al incio y al final
                    SQL = "1" & Format(L, "00000000") & "1"
                                        
                    TextoLecturasExpedicion "SSCC " & Mid(txtCPalet(Index).Text, 1, 12) & "... >  " & L
        
                    txtCPalet(Index).Text = SQL
        
                End If
            End If
            
        End If
        If cadErr <> "" Then
            cadErr = "SSCC " & cadErr
            TextoLecturasExpedicion Mid(txtCPalet(Index), 14) & "  Error " & cadErr
            
            AñadeACuadroMsgCboPalet SQL, False
            AñadeACuadroMsgCboPalet txtCPalet(Index).Text, True
            
            
            
            txtCPalet(Index).Text = ""
            PonerFoco txtCPalet(Index)
            Exit Sub
        End If
        SQL = ""
    End If

    
    
    
    
    
    
    
    
    
    
    'Este es el palet origen
    Set cPal = Nothing
    
    SQL = EtiquetaPaletCorrecta2(txtCPalet(Index).Text)
    If SQL = "" Then
        
            
            Set cPal = New CPalet
            SQL = Mid(txtCPalet(Index).Text, 2, 8) 'los ocho centrales
            If Not cPal.Leer(CLng(SQL)) Then
                SQL = "No existe " & Label2(8 + Index).Caption
                Set cPal = Nothing
            Else
                'Palet origen y destino NO puede ser el mismo
                If txtCPalet(0).Text = txtCPalet(1).Text Then
                    SQL = "Palet origen y destino NO pueden ser el mismo"
                    Set cPal = Nothing
                Else
                    SQL = "OK.  Palet: " & txtCPalet(Index).Text & vbCrLf & vbCrLf
                End If
            End If
            
            
    End If
    SQL = "(" & Index & ") " & SQL
    
    AñadeACuadroMsgCboPalet SQL, False
    AñadeACuadroMsgCboPalet txtCPalet(Index).Text, True
    
    
    If cPal Is Nothing Then
        txtCCaja.Text = ""
        txtCPalet(Index).Text = ""
        
        PonerFoco txtCPalet(Index)
    Else
        
        If Index = 0 Then
            PonerFoco txtCPalet(1)
        Else
            PonerFoco txtCCaja
        End If
        Set cPal = Nothing
    End If
        
End Sub


Private Sub txtlecturaPelt_KeyPress(KeyAscii As Integer)
    'Pulsa enter
    If KeyAscii = 13 Then ProcesarCodigoCaja
    
End Sub



Private Sub AñadeACuadroMsg(Texto As String, EsLectura As Boolean)
    TextErrPoste.Text = Texto & vbCrLf & TextErrPoste.Text
    If EsLectura Then TextErrPoste.Text = "#-" & TextErrPoste.Text
    If Len(TextErrPoste.Text) > 600 Then TextErrPoste.Text = Mid(TextErrPoste.Text, 1, 400)

End Sub

Private Sub AñadeACuadroMsgCboPalet(Texto As String, EsLectura As Boolean)
    txtObsCambioPalet.Text = Texto & vbCrLf & txtObsCambioPalet.Text
    If EsLectura Then txtObsCambioPalet.Text = "#" & txtObsCambioPalet.Text
    If Len(txtObsCambioPalet.Text) > 600 Then txtObsCambioPalet.Text = Mid(txtObsCambioPalet.Text, 1, 400)

End Sub


Private Sub ProcesarCodigoCaja()
Dim C As String
Dim Aux As String
Dim OK As Boolean
Dim LineaPal As String
Dim CodigoCaja As String

    CodigoCaja = Trim(txtlecturaPelt.Text)

    If CodigoCaja = "" Then Exit Sub
    'haga lo que haga
    txtlecturaPelt.Text = ""
    PonerFoco txtlecturaPelt
    AñadeACuadroMsg CodigoCaja, True
    
    If Len(CodigoCaja) <> 13 Then
        'MsgBox "Longitudad etiqueta incorrecta", vbExclamation
        
        AñadeACuadroMsg "Longitudad etiqueta incorrecta" & Len(CodigoCaja), False
        Exit Sub
    End If
    If Not IsNumeric(CodigoCaja) Then
        'MsgBox "Campo numérico", vbExclamation
        AñadeACuadroMsg "Campo numerico", False
        
    Else
        'Dividimos la etiqueta leida en 2
        'los cinco ultimos son el IDCAJa
        'el resto ID trza
        'select * from prodcajas where lotetraza=27 and idcaja=5
        C = Mid(CodigoCaja, 1, Len(CodigoCaja) - 5)
        
        
     '   If Timer - Tiempo1 > 60 Then
     '       'Cada minuto vuelvo a leer las lineas
            idTrazaAntiguo = -1
     '       Tiempo1 = Timer
     '   End If
        
        
        'Ya tengo el codigo de trazabilidad
        If idTrazaAntiguo <> Val(C) Then
            'Vamos a ver en que palet se esta produciendo esto
            idTrazaAntiguo = Val(C)
            'Pongo ffin de la propalettrza no la del palet. Puede que vayamos a paletizar otr
            C = DevuelveDesdeBD(conAri, "prodpalets.idpalet", "prodpalets,prodpaletstraza", "prodpalets.idpalet= prodpaletstraza.idpalet and fhfin is null and lotetraza ", C)
            
            'Este no se porque esta. Lo comento
            'If C <> "" Then C = DevuelveDesdeBD(conAri, "prodpalets.idpalet", "prodpalets,prodpaletstraza", "prodpalets.idpalet= prodpaletstraza.idpalet and fhfin is null and lotetraza ", C)
            ValoresLeidos = ""
            If C = "" Then
                '******************************************************************
                'ERROR. No encuentra el palet
                '******************************************************************
                idTrazaAntiguo = -1
                AñadeACuadroMsg "No existe palet asignado", False
            Else
                ValoresLeidos = C
            End If
        
        End If
        
        If idTrazaAntiguo > 0 Then
            'Si es correcta metemos en la linea de produccion
            '
            C = "insert into prodcajas(lotetraza,idcaja,idpalet,fcreacion) VALUES (" & idTrazaAntiguo & "," & Right(CodigoCaja, 5) & "," & ValoresLeidos & ",'" & Format(Now, FormatoFechaHora) & "')"
            If EjecutaSQL(conAri, C, False) Then
                'TODO BIEN
                
            Else
                AñadeACuadroMsg "YA existe la caja", False
            End If
        End If
    End If
End Sub

Private Function lecturaAlbaranExpedicion() As String

    'Si esta la orden de carga leida
    SQL = ""
    If Me.Label5(0).Tag = 0 Then
        SQL = "Falta orden carga"
    Else
        'Veremos si la etiqueta es correcta
        If Left(Text10.Text, 1) <> Right(Text10.Text, 1) Then SQL = "Etiqueta incorrecta(2)"
    End If
    
    If SQL <> "" Then
        lecturaAlbaranExpedicion = SQL
        Exit Function
    End If
    
    
    
        'Ahora veremos si pertence a la orden de carga
        SQL = Mid(Text10.Text, 3, 6)
        SQL = "numalbar= " & SQL & " AND id = " & Label5(0).Tag & " AND codtipom="
        'If Mid(Text10.Text, 2, 1) <> 1 Then msgbox
        SQL = SQL & " 'ALV'"
        
        SQL = "Select * from srepartol WHERE " & SQL
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        If RS.EOF Then
            SQL = "NO pertence a la orden de expedición"
        Else
            ' el albaran YA ha sido cargado
            If Val(RS!albexpedido) = 1 Then SQL = "Albarán ya ha sido cargado"
        End If
        RS.Close
        
        If SQL <> "" Then
            lecturaAlbaranExpedicion = SQL
            Exit Function
        End If
        
        'Veremos si quiere cerrar el albaran
        If Mid(Text10, 1, 1) = "2" Then
            
            'Vamos a cerrar el albaran, tendremos que ver cual es. Puede QUE nosea el que esta entrando ahora
            '
            lecturaAlbaranExpedicion = CerrarAlbaranExpedicion(Val(Mid(Text10.Text, 3, 6)))
            If lecturaAlbaranExpedicion = "" Then
                Label5(1).Tag = 0
                Label5(1).Caption = "Albaran: "
            End If
            
        Else
            Label5(1).Tag = Val(Mid(Text10.Text, 3, 6))
            Label5(1).Caption = "Albaran: " & Mid(Text10.Text, 3, 6)
        End If
        
        
End Function




Private Function lecturaOrdenExpedicion() As String
Dim MismaOrden As Boolean

    'Si esta la orden de carga leida
    SQL = ""
    If Left(Text10.Text, 1) <> Right(Text10.Text, 1) Then SQL = "Etiqueta incorrecta(1)"
    
    If SQL <> "" Then
        lecturaOrdenExpedicion = SQL
        Exit Function
    End If
    
    
    
        'Ahora veremos si existe y esta en situacion a la orden de carga
        SQL = Mid(Text10.Text, 2, 7)
        SQL = "Select * from srepartoc WHERE id = " & SQL
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        If RS.EOF Then
            SQL = "NO existe la orden "
        Else
            If Val(RS!Situacion) = 3 Then SQL = "Orden cerrada"
            If Val(RS!Situacion) = 2 Then SQL = "Ya realizada la carga"
        End If
        RS.Close
        
        If SQL <> "" Then
            lecturaOrdenExpedicion = SQL
            Exit Function
        End If
        
        'Veremos si quiere cerrar la orden
        If Mid(Text10, 1, 1) = "9" Then
            
            'If PuedeCerrarAlbaran Then
            lecturaOrdenExpedicion = CerrarOrdenExpedicion
            
        Else
            'Sigue leyenod introduccion
            MismaOrden = False
            SQL = Mid(Text10.Text, 2, 7)
            If Label5(0).Tag > 0 Then
                MismaOrden = Label5(0).Tag = Val(SQL)
            Else
                'Primera vez
                MismaOrden = True
                Label5(0).Tag = Val(SQL)
                Label5(0).Caption = "Orden carga: " & SQL
            End If
            If Not MismaOrden Then
                If MsgBox("¿Cambiar orden de expedicion?", vbQuestion + vbYesNoCancel) = vbYes Then
                    Label5(0).Tag = Val(SQL)
                    Label5(0).Caption = "Orden carga: " & SQL
                    Label5(1).Tag = 0
                    Label5(1).Caption = "Albaran: "
                Else
                    lecturaOrdenExpedicion = "Proceso nueva exp. cancelado"
                End If
            End If
        End If
        
        
End Function

Private Function LeerPaletExpedicion() As String
Dim Co As Collection
Dim i As Integer
Dim Aux As Integer
Dim CajasPorCargar As Integer
Dim J As Integer
Dim RN As ADODB.Recordset
Dim CargaPalet As Boolean
Dim CadeLot As String

    SQL = ""
    If Me.Label5(0).Tag = 0 Then
        SQL = "Falta orden carga"
    Else
        If Me.Label5(1).Tag = 0 Then
            'Albaran si leer
            SQL = "Falta leer albaran"

        End If
    End If
    
    If SQL <> "" Then
        LeerPaletExpedicion = SQL
        Exit Function
    End If

    'FALTA
    'Deberiamos comprobar que el palet no esta asignado a ningun albaran
    


    'Ya tengo el ID del palet.
    'Veamos que existe y que los datos que esta metiendo son articulos del la orden de expedicion que esta metiendo
    'Ahora veremos si existe y esta en situacion a la orden de carga
    SQL = "select codartic,count(*) from prodlin,prodtrazlin,prodcajas  where  prodlin.Codigo = prodtrazlin.Codigo "
    SQL = SQL & " AND prodlin.idlin = prodtrazlin.idlin And prodtrazlin.lotetraza = prodcajas.lotetraza"
    SQL = SQL & " AND idpalet=" & Mid(Text10.Text, 2, 8) & " group by 1"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        LeerPaletExpedicion = "Error leyendo datos trazab. palet: " & Mid(Text10.Text, 2, 8)
        RS.Close
        Exit Function
    
    Else
        'Existe el palet. Veamos que lo que lleva el palet esta en el albaran que estamos cargando
        '1.- Que todo lo que lleva el palet es del albaran ese
        'MEtere el codartic y las cajas
        Set Co = New Collection
        While Not RS.EOF
            SQL = RS!codartic & "|" & RS.Fields(1) & "|"
            Co.Add SQL
            RS.MoveNext
        Wend
    
      
    End If
    RS.Close
    
      
    For i = 1 To Co.Count
         SQL = "select codartic,sum(cajas) vcajas, sum(llevamos) vllevamos from srepartolot "
         SQL = SQL & " WHERE idreparto=" & Label5(0).Tag & " and numalbar=" & Label5(1).Tag & " and codartic=" & DBSet(RecuperaValor(Co.Item(i), 1), "T")
         SQL = SQL & " group by 1"
         RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
         SQL = ""
         If RS.EOF Then
            'Error grave. No se encuentra id/alb/art "
            SQL = "Error grave. No se encuentra alb/art " & Label5(1).Tag & " /" & RecuperaValor(Co.Item(i), 1)
         Else
             CajasPorCargar = RS!vcajas - RS!vllevamos
             Aux = Val(RecuperaValor(Co.Item(i), 2))
             If CajasPorCargar < Aux Then
                SQL = RS!codartic & vbCrLf & "  Falta: " & CajasPorCargar & " Palet: " & Aux
             End If
         End If
         RS.Close
         If SQL <> "" Then
            Set Co = Nothing
            LeerPaletExpedicion = SQL
            Exit Function
         End If
    Next
        
        
    'Otra comprobacion.
    'Veremos si los lotes tienen algo mas antiguo
    ValoresLeidos = ""
    For i = 1 To Co.Count
         SQL = "select distinct(lotetraza) from prodcajas where idpalet=" & Mid(Text10.Text, 2, 8)
         RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
         CadeLot = "|"
         While Not RS.EOF
            CadeLot = CadeLot & RS.Fields(0) & "|"
            RS.MoveNext
         Wend
         RS.Close
             
             
         SQL = "select * from spartidas where  cantotal>0 and codartic=" & DBSet(RecuperaValor(Co.Item(i), 1), "T")
         SQL = SQL & " order by fecha asc"
         RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
         SQL = ""
         If Not RS.EOF Then
            SQL = RS!NUmlote
            If InStr(1, SQL, " ") > 0 Then
                'ANtiguo
                SQL = Trim(Mid(SQL, 1, InStr(1, SQL, " ")))
            Else
                'No hacemos nada, esta bien
                SQL = Val(SQL)
            End If
            
            SQL = "|" & SQL & "|"
            If InStr(1, CadeLot, SQL) = 0 Then
                'Significa que hay un lote antorior
                SQL = Mid(SQL, 2)
                SQL = Mid(SQL, 1, Len(SQL) - 1)
                ValoresLeidos = ValoresLeidos & vbCrLf & RecuperaValor(Co.Item(i), 1) & ": " & SQL
            End If
            
         End If
         CadeLot = ""
            
         
         RS.Close
         
    Next
    If ValoresLeidos <> "" Then
        ValoresLeidos = "Lotes anteriores" & ValoresLeidos & "   ¿Continuar?"
        If MsgBox(ValoresLeidos, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Set Co = Nothing
            LeerPaletExpedicion = ValoresLeidos
            Exit Function
        End If
    End If
        
    
        
    'Hacemos el insert
    'Iremos metiendo en el CO los inserts
    Set RN = New ADODB.Recordset
    For i = 1 To Co.Count
        'Cada articulo vere en el albaran cuantas lineas voy a meter
        'SOLO deberia haber un codartic.
        SQL = "select codartic,cajas,llevamos,codtipom,numalbar,numlinea  from srepartolot "
        SQL = SQL & " WHERE idreparto=" & Label5(0).Tag & " and numalbar=" & Label5(1).Tag & " and codartic=" & DBSet(RecuperaValor(Co.Item(i), 1), "T")
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Aux = Val(RecuperaValor(Co.Item(i), 2))
        'NO DEBE SER EOF
        While Not RS.EOF
             CajasPorCargar = RS!Cajas - RS!llevamos
             If CajasPorCargar >= Aux Then
                  'De todas las del palet cojere las que son de este articulo
                  SQL = "select distinct(prodcajas.lotetraza) from prodlin,prodtrazlin,prodcajas  where "
                  SQL = SQL & " prodlin.Codigo = prodtrazlin.Codigo  AND prodlin.idlin = prodtrazlin.idlin And"
                  SQL = SQL & " prodtrazlin.lotetraza = prodcajas.lotetraza AND idpalet=" & Mid(Text10.Text, 2, 8)
                  SQL = SQL & " AND codartic=" & DBSet(RS!codartic, "T")
                  J = 0
                  RN.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                  CadeLot = ""
                  While Not RN.EOF
                        J = J + 1
                        CadeLot = CadeLot & ", " & RN.Fields(0)
                        RN.MoveNext
                  Wend
                  RN.Close
                  
                  If J = 0 Then
                    'DEBERIA HABER ENCONTRADO un lote por lo menos
                     Set Co = Nothing
                     RS.Close
 
                
                     LeerPaletExpedicion = "Palet. NO encuentra lote trazabilidad"
                     Exit Function
            
                  End If
                  CadeLot = Mid(CadeLot, 2)
                  
                  If J = 1 Then
                        CadeLot = " AND lotetraza = " & CadeLot
                  Else
                        CadeLot = " AND lotetraza IN (" & CadeLot & ")"
                  End If
                  'Voy a insertar cajas
                  SQL = "select * from prodcajas where idpalet=" & Mid(Text10.Text, 2, 8) & CadeLot & " and not (lotetraza,idcaja)  in "
                  SQL = SQL & " (select lotetraza,idcaja FROM srepartolotcaj "
                  SQL = SQL & " WHERE idreparto=" & Label5(0).Tag & " and numalbar=" & Label5(1).Tag & " and codtipom=1 ) ORDER BY idCaja"
                  RN.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                  LeerPaletExpedicion = ""
                  J = 0
                  SQL = ""
                  Do
                    If RN.EOF Then
                        'YA NO inserto mas
                        'J = CajasPorCargar
                        CajasPorCargar = J    'No cambio J que despues se utiliza bajo
                    Else
                        J = J + 1
                        'insert
                        'INSERT INTO  srepartolotcaj(idreparto,codtipom,numalbar,numlinea,codartic,fecha,lotetraza,idcaja,idPalet) VALUES ("
                        SQL = SQL & ", (" & Label5(0).Tag & ",1," & Label5(1).Tag & "," & RS!numlinea & "," & DBSet(RS!codartic, "T") & ",now(),"
                        SQL = SQL & RN!lotetraza & "," & RN!idcaja & "," & RN!IdPalet & ")"
                        RN.MoveNext
                        
                    End If
                  Loop Until J >= CajasPorCargar
                  RN.Close
                  
                  'Actualiz lllevamps
                  If J > 0 Then
                        'Insertamos cajas
                        SQL = Mid(SQL, 2) 'quito la coma
                        SQL = "INSERT INTO  srepartolotcaj(idreparto,codtipom,numalbar,numlinea,codartic,fecha,lotetraza,idcaja,idPalet) VALUES " & SQL
                        If Not EjecutaSQL(conAri, SQL, False) Then
                            'ERRROR
                            Set Co = Nothing
                            RS.Close

                            
                            LeerPaletExpedicion = "Error insertando cajas repartol"
                            Exit Function
                        End If
                        'Actualizmos cajas
                        SQL = " where idreparto=" & Label5(0).Tag & " and numalbar=" & Label5(1).Tag & " and codtipom=1 AND numlinea = " & RS!numlinea
                        SQL = "UPDATE srepartolot set llevamos=llevamos + " & J & SQL
                        EjecutaSQL conAri, SQL, True  'NO DEBERIA DAR ERROR
                        Espera 0.2
                  End If
                  
            End If
            RS.MoveNext
        Wend
        RS.Close
    Next
             
             
    Set RN = Nothing
End Function








Private Function LeerCajaExpedicion() As String
Dim EnEsta As Boolean
Dim IdPalet As Long
    'Si esta la orden de carga y el albaran leidos
    SQL = ""
    If Me.Label5(0).Tag = 0 Then
        SQL = "Falta orden carga"
    Else
        If Me.Label5(1).Tag = 0 Then
            'Albaran si leer
            SQL = "Falta leer albaran"

        End If
    End If
    
    If SQL <> "" Then
        LeerCajaExpedicion = SQL
        Exit Function
    End If
    


        'Ahora veremos si pertence a una orden de carga
        SQL = Mid(Text10.Text, 1, 8) & " AND idcaja = " & Val(Mid(Text10, 9))
        SQL = "Select * from prodcajas WHERE lotetraza = " & SQL
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        IdPalet = 0
        If RS.EOF Then
            LeerCajaExpedicion = "NO existe la caja en el sistema"
        Else
            'Vemos si lo que lee es lo que escribe ;
            'Si realmente la caja es de lo que me ha dicho en el albarna
            IdPalet = DBLet(RS!IdPalet, "N")
            SQL = "select codartic from prodlin,prodtrazlin  where prodlin.codigo= prodtrazlin.codigo AND prodlin.idlin = prodtrazlin.idlin and lotetraza=" & Mid(Text10.Text, 1, 8)
        End If
        RS.Close
        
        If SQL = "" Then Exit Function
        
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        If RS.EOF Then
            LeerCajaExpedicion = "Error en lote de trazabilidad"
        Else
            'El codartic esta en nla primera linea del label, desde la posicion 4 hasta el ·
            'SQL = Mid(Label4(1).Caption, 1, InStr(Label4(1).Caption, "·") - 1)
            SQL = DBSet(RS!codartic, "T")
            SQL = "select * from srepartolot where idreparto=" & Label5(0).Tag & " and numalbar=" & Label5(1).Tag & " and codartic=" & SQL
        End If
        RS.Close
        
        If SQL = "" Then Exit Function
        
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        If RS.EOF Then
            LeerCajaExpedicion = "No datos para articulo"
        Else
            EnEsta = False
            While Not EnEsta
                If RS!Cajas - RS!llevamos = 0 Then
                    'NO CABEN
                    
                Else
                    
                    'OK la caja pertenece al articulo
                    SQL = "INSERT INTO  srepartolotcaj(idreparto,codtipom,numalbar,numlinea,codartic,fecha,lotetraza,idcaja,idPalet) VALUES ("
                    SQL = SQL & Label5(0).Tag & ",1," & Label5(1).Tag & "," & RS!numlinea & "," & DBSet(RS!codartic, "T") & ",now(),"
                    'caja y traza y palet que leyendo las cajas a mano ---> Lo tengo leido en IdPalet
                    SQL = SQL & Val(Mid(Text10.Text, 1, 8)) & " ," & Val(Mid(Text10, 9)) & "," & IdPalet & ")"
                    
                    If Not EjecutaSQL(conAri, SQL, False) Then
                        LeerCajaExpedicion = "Caja ya asignada"
                        EnEsta = True 'No hace falta que busque mas
                    Else
                        LeerCajaExpedicion = ""
                        'Updateamos las cajas que llevamos
                        'idreparto codtipom numalbar numlinea
                        SQL = "where idreparto=" & Label5(0).Tag & " and numalbar=" & Label5(1).Tag & " and codtipom=1 AND numlinea = " & RS!numlinea
                        
                        SQL = "UPDATE srepartolot set llevamos=llevamos +1 " & SQL
                        EjecutaSQL conAri, SQL, True  'NO DEBERIA DAR ERROR
                        Espera 0.2
                        
                        EnEsta = True
                    End If
                End If
                If Not EnEsta Then
                    RS.MoveNext
                    If RS.EOF Then EnEsta = True
                End If
            Wend
            If SQL = "" Then LeerCajaExpedicion = "Referencia completa"
        End If
        RS.Close
        
    




End Function
'Cierre de una orden de produccion
Private Function CerrarOrdenExpedicion() As String
Dim Co As Collection
Dim cad As String
Dim J As Integer
Dim QueAlbaranFalta As Long

    SQL = "Select * FROM srepartol where id = " & Me.Label5(0).Tag
    cad = ""
    Set Co = New Collection
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    J = 0
    While Not RS.EOF
        If RS!albexpedido = 0 Then
            J = J + 1
            QueAlbaranFalta = RS!NumAlbar
        End If
        SQL = RS!NumAlbar
        Co.Add SQL
        
        RS.MoveNext
    Wend
    RS.Close
    
    If J > 0 Then   'De momento NO dejo seguir si falta. En futuro cerraremos aqui TOOOODOs
        'Existe MAS de un albaran por cerrar la orden de expedicion
        cad = "Albaranes sin finalizar expedición: " & QueAlbaranFalta
    End If
    
    If cad <> "" Then
        CerrarOrdenExpedicion = cad
        Exit Function
    End If
    
    If Co.Count = 0 Then
        CerrarOrdenExpedicion = "Error leyendo tabla lineas reparto(4)"
        Exit Function
    End If
    
    
    
    
    'Compruebo que el albaran NO tiene lotes asignados
    cad = "select count(*) from slialblotes where (codtipom,numalbar) in (select codtipom,numalbar from srepartol where id=" & Me.Label5(0).Tag & ")"
    RS.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cad = ""
    If Not RS.EOF Then
        If DBLet(RS.Fields(0), "N") > 0 Then cad = "Algunos albaranes ya tienen lote asignado"
    End If
    RS.Close
    If cad <> "" Then
        CerrarOrdenExpedicion = cad
        Exit Function
    End If
    
    

    For J = 1 To Co.Count
        'Comprobaremos si estan todos los albaranes cerrados
        cad = PuedeCerrarAlbaran(CLng(Me.Label5(0).Tag), CLng(Me.Label5(1).Tag), 1)  'alv
        If cad <> "" Then
            CerrarOrdenExpedicion = "Error leyendo tabla lineas reparto(4)"
            Set Co = Nothing
            Exit Function
        End If
    Next J
    
    'CERRAMOS LA ORDEN DE CARGA y...... actualizamos slotes
    cad = ActualizarLotajeAlbaran
    If cad <> "" Then
        CerrarOrdenExpedicion = cad
        Exit Function
    End If
    
    
    
    'Sept 2012
    '-  Dara de baja las cajas y ajustara los palets
    'Ajustar palets
    AjusteCajasPaletsPorCierreOrdenExpedicion
    
    
    SQL = "UPDATE srepartoc SET situacion=3 WHERE id = " & CLng(Me.Label5(0).Tag)
    If Not EjecutaSQL(conAri, SQL, True) Then
        CerrarOrdenExpedicion = "Error actualizando sreparto(5). Ver situacion"
    Else
        cmdLimpExpedicion_Click
    End If
    Set Co = Nothing
End Function






'Cierre de un albaran en expedicion
Private Function CerrarAlbaranExpedicion(K_Albaran As Long) As String
Dim cad As String
    
    cad = PuedeCerrarAlbaran(CLng(Me.Label5(0).Tag), K_Albaran, 1)  'alv
    If cad <> "" Then
        'VEAMOS.   Vamos a dejar que el programa modifique las lineas del albaran (y stocks y .....)
        'para poder cuadrar despues las cantidadades
       ' If MsgBox(Cad & vbCrLf & "Desea modificar el albaran?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
       '     If MsgBox("Guardara LOG de acciones del usuario " & vUsu.Nombre & " . ¿Continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
       '         ModificarAlbaranConCantidadesCargadas CLng(Me.Label5(0).Tag), K_Albaran, 1
       '         Cad = PuedeCerrarAlbaran(CLng(Me.Label5(0).Tag), K_Albaran, 1)  'alv
       '         If Cad = "" Then TextoLecturasExpedicion "Lineas de albaran ajustada cantidad"
       '     End If
       ' End If
    End If
    If cad = "" Then

            SQL = "UPDATE srepartol SET albexpedido=1 WHERE numalbar= " & K_Albaran & " AND id = " & CLng(Me.Label5(0).Tag) & " AND codtipom='ALV'"
            If Not EjecutaSQL(conAri, SQL, True) Then cad = "Error albexpedido update table"

    End If
    CerrarAlbaranExpedicion = cad
End Function



Private Function PuedeCerrarAlbaran(idCarga As Long, idAlbaran As Long, idTipAlbaran As Byte) As String
Dim TodoBien As Boolean
Dim RN As ADODB.Recordset
Dim ColFalta As Collection
Dim CuentaCajas As Integer
Dim Aux As String
On Error GoTo eCierreAlb
    
    
    
    
    'Veamos si puede cerrar el albaran
    
    SQL = "numalbar= " & idAlbaran & " AND idreparto = " & idCarga & " AND codtipom=" & idTipAlbaran
    
    SQL = "select numlinea,codartic,cajas,llevamos from srepartolot where " & SQL & " ORDER BY numlinea"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    Set RN = New ADODB.Recordset
    SQL = "numalbar= " & idAlbaran & " AND idreparto = " & idCarga & " AND codtipom=" & idTipAlbaran
    SQL = "select numlinea,codartic,count(*) cuantashay from srepartolotcaj where  " & SQL
    SQL = SQL & " group by 1,2  ORDER BY numlinea"
    RN.Open SQL, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    Set ColFalta = New Collection
    'NO PUEDE SER EOF
    While Not RS.EOF
        'Veremos la diferencia para cada linea
        SQL = "numlinea = " & RS!numlinea
        
        RN.Find SQL, , adSearchForward, 1
        
        CuentaCajas = 0
        If Not RN.EOF Then CuentaCajas = DBLet(RN!cuantashay, "N")
        
        If RS!Cajas <> CuentaCajas Then
            'NO cuentan las mismas
            'Aqui añadiriamos ams cosas para el error
            
            SQL = "Linea: " & RS!numlinea & "  -  " & RS!codartic
            ColFalta.Add SQL
            
            
        End If
        
        If RS!llevamos <> CuentaCajas Then
            SQL = "numalbar= " & idAlbaran & " AND idreparto = " & idCarga & " AND codtipom=" & idTipAlbaran
            SQL = "UPDATE srepartolot SET llevamos = " & CuentaCajas & " WHERE " & SQL
            SQL = SQL & " AND numlinea = " & RS!numlinea
            EjecutaSQL conAri, SQL, True
        End If
        
        RS.MoveNext
    Wend
    RS.Close
    RN.Close
    
    
    If ColFalta.Count > 0 Then
        'Abril 2017. No dejo continuar
        'If MsgBox("Faltan lineas por expedir", vbQuestion + vbYesNo) = vbNo Then
            PuedeCerrarAlbaran = "Faltan lineas por expedir"
            Exit Function
        'End If
    End If
        
    
        'ULTIMA COMPROBACION.
        'NO deberian haber cambiado el numero de cajas(uds) del albaran.
        'PEeeeeeero, por si acaaaaaso, compruebo que es lo que hay y lo que pone el alb
        SQL = "select codartic,sum(cajas) cajasp from srepartolot where idreparto=" & Label5(0).Tag & " and numalbar=" & Label5(1).Tag & " AND codtipom=" & idTipAlbaran & " GROUP BY codartic"
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        Set ColFalta = Nothing
        Set ColFalta = New Collection
        While Not RS.EOF
            SQL = "select slialb.codartic,unicajas,sum(cajas) vcajas,sum(cantidad) vcantidad from slialb,sartic where slialb.codartic=Sartic.codartic"
            SQL = SQL & " AND numalbar = " & Label5(1).Tag
            If idTipAlbaran = 1 Then SQL = SQL & " AND codtipom = 'ALV'"
            SQL = SQL & " AND slialb.codartic = " & DBSet(RS!codartic, "T")
            SQL = SQL & " group by 1"
            
            RN.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If RN.EOF Then
                SQL = "NO se encuentra articulo albaran " & RS!codartic & "/" & Label5(1).Tag
            Else
                'Si que esta
                'veremos si el albaran esta bien las cajas uds
                If (RN!UniCajas * RN!vcajas) <> RN!vCantidad Then
                    SQL = "Alb: cajas * unicajas <> cantidad"
                Else
                    'Veremos si las cajas de aqui son las del albaran
                    If RS!cajasp <> RN!vcajas Then
                        'Distinto numero de cajas tabla  con las del albaran
                        SQL = "Reparto (lot) <> nº cajas"
                    Else
                        'OK
                        SQL = ""
                    End If
                End If
            End If
            RN.Close
            
            If SQL <> "" Then ColFalta.Add SQL
                
            RS.MoveNext
        Wend
        RS.Close
        
        'Veremos si hay errores en las contra el albaran
        SQL = ""
        For CuentaCajas = 1 To ColFalta.Count
            SQL = SQL & ColFalta.Item(CuentaCajas) & vbCrLf
        Next
        PuedeCerrarAlbaran = SQL

    
    Set RN = Nothing
    Set ColFalta = Nothing
    Exit Function
eCierreAlb:
    PuedeCerrarAlbaran = "ERROR " & Err.Description
    Err.Clear
    Set RS = Nothing
    Set RS = New ADODB.Recordset
    Set ColFalta = Nothing



End Function

Private Sub TextoLecturasExpedicion(TextoAAñadir As String)
    Text9.Text = TextoAAñadir & vbCrLf & vbCrLf & Text9.Text
    If Len(Text9.Text) > 1000 Then Text9.Text = Mid(Text9.Text, 1, 750)
End Sub





'Cerrar orden de expedicion. Actualiza los lotes del albaran, y, obviamnete, las partidas
Private Function ActualizarLotajeAlbaran() As String
Dim cPar As cPartidas
Dim Bien As Boolean
Dim linea As String
Dim vLin As Integer

    SQL = "select numalbar,numlinea,codartic,lotetraza ,count(*)  as cantidad  from srepartolotcaj where idreparto="
    SQL = SQL & Label5(0).Tag & " group by 1,2,3,4 order by 1,2,3,4"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        
        If MsgBox("Ninguna caja expedida.¿Continuar?", vbQuestion + vbYesNo) = vbNo Then
            RS.Close
            ActualizarLotajeAlbaran = "Error agrupando cajas orden expedición"
            Exit Function
        End If
    End If
    ActualizarLotajeAlbaran = ""
    Conn.BeginTrans
    Bien = True
    While Not RS.EOF
        Set cPar = New cPartidas
        SQL = RS!NumAlbar & "|" & RS!numlinea & "|" & RS!codartic
        If linea <> SQL Then
             vLin = 1
             linea = SQL
        Else
            vLin = vLin + 1
        End If
        If Bien Then  'para que no haga las demas a partir del fallo
            If cPar.LeerDesdeExpedicion(RS!codartic, 1, Format(RS!lotetraza, "0000000000")) Then
                If Not InsertarModificarLoteLinea(cPar, vLin) Then
                    ActualizarLotajeAlbaran = "Ins/Mod lote: " & RS!codartic & " " & cPar.NUmlote
                    Bien = False
                End If
            Else
                Bien = False
                ActualizarLotajeAlbaran = "Lote traza: " & RS!codartic & " " & " " & cPar.NUmlote
            End If
        End If
        RS.MoveNext
        Set cPar = Nothing
    Wend
    RS.Close
    If Bien Then
        Conn.CommitTrans
    Else
        Conn.RollbackTrans
    End If

End Function

'Com RS es public , NO paso nada
Private Function InsertarModificarLoteLinea(ByRef cPa As cPartidas, vLineaLot As Integer) As Boolean
Dim Leido As Boolean
Dim Can As Currency
Dim cLot As cLotaje


    On Error GoTo EInsertarModificar
    InsertarModificarLoteLinea = False
    '---------------------
    Set cLot = New cLotaje

    cLot.codartic = cPa.codartic
    cLot.codalmac = cPa.codalmac
    cLot.DetaMov = "ALV"
    cLot.LineaDocu = RS!numlinea
    cLot.Documento = RS!NumAlbar
    cLot.tipoMov = 0

    
    SQL = DevuelveDesdeBD(conAri, "unicajas", "sartic", "codartic", cPa.codartic, "T")
    Can = Val(SQL)
    If Can = 0 Then Can = 1
    
    
    cLot.NUmlote = cPa.NUmlote
    cLot.Cantidad = Can * ImporteFormateado(RS!Cantidad)
    cLot.SubLinea = vLineaLot 'La sublinea del lote 'Normalmente 1 o 2
  

    SQL = "insert into `slialblotes` (`codtipom`,`numalbar`,`numlinea`,`linea`,`numlote`,cantidad) values ('"
    SQL = SQL & cLot.DetaMov & "'," & cLot.Documento & "," & RS!numlinea & ","
    'SQL = SQL & txtAux(0).Text & ",'" & DevNombreSQL(txtAux(1).Text) & "'," & DBSet(txtAux(2).Text, "N") & ")"
    'Ahora
    SQL = SQL & vLineaLot & ",'" & cPa.NUmlote & "'," & DBSet(cLot.Cantidad, "N") & ")"



    Conn.Execute SQL


    'Hay k rellenar el resto de valores
    cLot.Fechamov = Now
    cLot.HoraMov = Now
    cLot.InsertarLote
    
   
    InsertarModificarLoteLinea = True  'Ya ponemos que esta bien
                               'aunque de errores bajo


     

    cPa.IncrementarCantidad -cLot.Cantidad

     Set cLot = Nothing






    
    Exit Function
EInsertarModificar:
    MuestraError Err.Number, Err.Description
End Function


'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
'EXPEDICION
'
Private Sub CargaExpedicion()
Dim IT As ListItem
        ListView1(0).ListItems.Clear


        SQL = "Select * from srepartol WHERE id = " & Label5(0).Tag & " ORDER BY numalbar"
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
            Set IT = Me.ListView1(0).ListItems.Add(, , Format(RS!NumAlbar, "000000"))
            IT.SubItems(1) = Format(RS!FechaAlb, "dd/mm/yyyy")
            IT.SubItems(2) = ""
            RS.MoveNext
        Wend
        
        RS.Close
        For idTrazaAntiguo = 1 To Me.ListView1(0).ListItems.Count
            CargarLineasAlbaranExpedicion ListView1(0).ListItems(idTrazaAntiguo), idTrazaAntiguo = 1
        Next
        idTrazaAntiguo = 0
        
        
        
End Sub


Private Sub CargarLineasAlbaranExpedicion(ByRef IT As ListItem, Cargar As Boolean)
Dim ITx As ListItem
        If Cargar Then ListView1(1).ListItems.Clear
        SQL = "select srepartolot.*,nomartic from srepartolot,sartic where srepartolot.codartic=sartic.codartic and"
        SQL = SQL & " idreparto=" & Label5(0).Tag & " and numalbar=" & IT.Text & " order by numlinea"
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        While Not RS.EOF
            
            If Cargar Then
                Set ITx = Me.ListView1(1).ListItems.Add(, , RS!codartic)
                ITx.SubItems(1) = RS!Cajas
                ITx.SubItems(2) = RS!llevamos
                ITx.SubItems(3) = RS!NomArtic
                ITx.Tag = " idreparto=" & Label5(0).Tag & " and numalbar=" & IT.Text & " AND " & _
                            "codtipom=" & RS!Codtipom & " AND numlinea=" & RS!numlinea & " AND codartic = '" & RS!codartic & "'"
                            
            End If
            If RS!Cajas - RS!llevamos <> 0 Then SQL = "NO"
            RS.MoveNext
        Wend
        RS.Close
        IT.SubItems(2) = SQL
End Sub




'''''''''Private Sub ModificarAlbaranConCantidadesCargadas(idCarga As Long, idAlbaran As Long, idTipAlbaran As Byte)
'''''''''Dim RT As ADODB.Recordset
'''''''''Dim vCS As CStock
'''''''''Dim Uds As Long
'''''''''
'''''''''        '  LOG de acciones
'''''''''        Set LOG = New cLOG
'''''''''
'''''''''
'''''''''    'Veremos cuantas cajas/unidades ha cargado
'''''''''    'y a partir de ahi rearemos el albaran
'''''''''
'''''''''    Set vCS = New CStock
'''''''''    SQL = "select * from scaalb where numalbar = " & idAlbaran
'''''''''    If idTipAlbaran = 1 Then SQL = SQL & " AND codtipom = 'ALV'"
'''''''''    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'''''''''
'''''''''    vCS.codalmac = 1
'''''''''    vCS.DetaMov = RS!Codtipom
'''''''''    vCS.Fechamov = RS!FechaAlb
'''''''''    vCS.HoraMov = RS!FechaAlb & " " & Format(Now, "hh:mm:ss")
'''''''''    vCS.Trabajador = RS!CodClien
'''''''''    vCS.Documento = Format(RS!NumAlbar, "0000000")
'''''''''    RS.Close
'''''''''
'''''''''
'''''''''    Set RT = New ADODB.Recordset
'''''''''    SQL = " srepartolot.codartic=sartic.codartic"
'''''''''    SQL = SQL & " AND numalbar= " & idAlbaran & " AND idreparto = " & idCarga & " AND codtipom=" & idTipAlbaran & " AND cajas - llevamos>0"
'''''''''    SQL = "select numlinea,unicajas,llevamos,srepartolot.codartic,nomartic from srepartolot,sartic where " & SQL & " ORDER BY numlinea"
'''''''''    RT.Open SQL, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
'''''''''    'NO puede ser eof, si no el paso anterior nu hubiera dicho que no puede cerrar albaran
'''''''''    While Not RT.EOF
'''''''''
'''''''''
'''''''''        SQL = "select * from slialb where numalbar = " & idAlbaran & " AND numlinea = " & RT!numlinea
'''''''''        If idTipAlbaran = 1 Then SQL = SQL & " AND codtipom = 'ALV'"
'''''''''        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'''''''''        If RS.EOF Then
'''''''''            '******   ERROR GRAVE
'''''''''            SQL = "*************************"
'''''''''            SQL = SQL & vbCrLf & SQL & vbCrLf
'''''''''            TextoLecturasExpedicion SQL & " NO EXISTE LINEA ALBARAN" & vbCrLf & SQL
'''''''''
'''''''''        Else
'''''''''            Conn.BeginTrans
'''''''''
'''''''''            'Modificamos
'''''''''            Uds = RT!llevamos * RT!UniCajas
'''''''''            If ModificarLinea(vCS, CInt(RT!llevamos), Uds) Then
'''''''''
'''''''''                'Pongo en "cajas" lo mismo que llevamos
'''''''''                SQL = "UPDATE srepartolot SET cajas=llevamos WHERE idreparto= " & idCarga & " AND codtipom=" & idTipAlbaran
'''''''''                SQL = SQL & " AND numalbar= " & idAlbaran & " AND numlinea = " & RT!numlinea
'''''''''                If EjecutaSQL(conAri, SQL, False) Then
'''''''''                    'Anteriores, ahora LOG
'''''''''                    SQL = "Alb: " & idAlbaran & " / " & RT!numlinea & "     IdExp: " & idCarga
'''''''''                    SQL = SQL & vbCrLf & RT!codArtic & " - " & RT!NomArtic
'''''''''                    SQL = SQL & vbCrLf & "Cajas (antes/ahora): " & RS!Cajas & " / " & RS!llevamos
'''''''''
'''''''''                    LOG.Insertar 10, vUsu, SQL
'''''''''                    Conn.CommitTrans
'''''''''                    Espera 0.5  'para garantizarnos que tarda mas de un seg
'''''''''                Else
'''''''''                    TextoLecturasExpedicion " Error updateando srepartol" & vbCrLf & SQL
'''''''''                    Conn.RollbackTrans
'''''''''                End If
'''''''''            Else
'''''''''                Conn.RollbackTrans
'''''''''            End If
'''''''''        End If
'''''''''        RS.Close
'''''''''        RT.MoveNext
'''''''''    Wend
'''''''''    RT.Close
'''''''''    Set RT = Nothing
'''''''''
'''''''''    Set LOG = Nothing
'''''''''
'''''''''
'''''''''End Sub
'''''''''
'''''''''
'''''''''
'''''''''
'''''''''
'''''''''
'''''''''Private Function ModificarLinea(ByRef cStk As CStock, CajasAlAlbaran As Integer, UdsAlAlbaran As Long) As Boolean
''''''''''Modifica un registro en la tabla de lineas de Albaran: slialb
'''''''''Dim SQL As String
'''''''''Dim B As Boolean
'''''''''Dim Importe As Currency
'''''''''
'''''''''    On Error GoTo EModificarLinea
'''''''''
'''''''''    ModificarLinea = False
'''''''''    SQL = ""
'''''''''    cStk.LineaDocu = RS!numlinea
'''''''''    cStk.codArtic = RS!codArtic
'''''''''
'''''''''
'''''''''
'''''''''    'Ponemos la cantodad de entrada QUE TENIA ANTES
'''''''''    'b = InicializarCStock(vCStock, "E")
'''''''''    cStk.tipoMov = "E"
'''''''''    'Cantidad del albaran
'''''''''    cStk.Cantidad = RS!Cantidad
'''''''''    cStk.Importe = RS!ImporteL
'''''''''    Importe = RS!precioar
'''''''''
'''''''''
'''''''''
'''''''''    'eliminamos de smoval y devolvemos stock valores anteriores
'''''''''    If Not cStk.DevolverStock2 Then Exit Function
'''''''''
'''''''''    'CajasAlAlbaran  UdsAlAlbaran
'''''''''    'b = InicializarCStock(vCStock, "S")
'''''''''    cStk.tipoMov = "S"
'''''''''    'Cantidad real que voy a poder "expedir"
'''''''''    cStk.Cantidad = UdsAlAlbaran
'''''''''    SQL = CalcularImporte(CStr(cStk.Cantidad), CStr(Importe), CStr(RS!dtoline1), CStr(RS!dtoline2), vParamAplic.TipoDtos)
'''''''''    cStk.Importe = CCur(SQL)
'''''''''
'''''''''
'''''''''
'''''''''    'insertamos en smoval y actualizamos stock a los valores nuevos
'''''''''    If Not cStk.ActualizarStock Then Exit Function
'''''''''
'''''''''    'actualizar la linea de Albaran
'''''''''
'''''''''
'''''''''    SQL = "UPDATE slialb SET "
'''''''''    SQL = SQL & "cantidad= " & DBSet(UdsAlAlbaran, "N") & ", "
'''''''''   ' SQL = SQL & "precioar= " & DBSet(txtAux(4).Text, "N") & ", " 'precio
'''''''''   ' SQL = SQL & "dtoline1= " & DBSet(txtAux(6).Text, "N") & ", dtoline2= " & DBSet(txtAux(7).Text, "N") & ", "
'''''''''    SQL = SQL & "importel= " & DBSet(cStk.Importe, "N") & ", " 'Importe
'''''''''    'Abril 2009
'''''''''    SQL = SQL & "cajas=" & DBSet(CajasAlAlbaran, "N", "N")
'''''''''
'''''''''
'''''''''    SQL = SQL & " WHERE numalbar =" & RS!NumAlbar & " AND codtipom='" & RS!Codtipom & "' AND numlinea =" & RS!numlinea
'''''''''    Conn.Execute SQL
'''''''''
'''''''''
'''''''''    'Hay que updatear
'''''''''
'''''''''    ModificarLinea = True
'''''''''
'''''''''    Exit Function
'''''''''EModificarLinea:
'''''''''   MuestraError Err.Number, "Modificar Lineas Albaran" & vbCrLf & Err.Description
'''''''''
'''''''''
'''''''''End Function

Private Sub ProcesaCajaCambioPalet()
Dim InsertaCaja As Boolean

    txtCCaja.Text = Trim(txtCCaja.Text)
    If txtCCaja.Text = "" Then Exit Sub
    
    
    
    SQL = "NO"
    If Me.txtCPalet(0).Text = "" Or txtCPalet(1).Text = "" Then
        SQL = "Debe poner el palet origen y/o destino"
        If txtCPalet(0).Text = "" Then
            PonerFoco txtCPalet(0)
        Else
            PonerFoco txtCPalet(1)
        End If
    Else
    
        
        If txtMoverUltimasNcajas.Text <> "" Then
            If Val(txtMoverUltimasNcajas.Text) > 1 Then
                'OBLCIADO PALET ORIGNE DESTINO
                If Me.txtCPalet(0).Text = "" Or txtCPalet(1).Text = "" Then
                    SQL = "Debe poner el palet origen y/o destino"
                    If txtCPalet(0).Text = "" Then
                        PonerFoco txtCPalet(0)
                    Else
                        PonerFoco txtCPalet(1)
                    End If
                Else
                    SQL = ""
                End If
            Else
                SQL = ""
            End If
        End If
              
        If SQL = "NO" Then
            'No hago nad
        
        Else
            'Longituda etc
            If Len(txtCCaja.Text) <> 13 Then
                SQL = "Longitud etiqueta incorrecta"
            
            Else
                If Not IsNumeric(txtCCaja.Text) Then
                    SQL = "Campo numerico"
                Else
                   InsertaCaja = False
                   
                    SQL = "lotetraza = " & Mid(txtCCaja.Text, 1, 8) & " AND idcaja  "
                    SQL = DevuelveDesdeBD(conAri, "idpalet", "prodcajas", SQL, Mid(txtCCaja.Text, 9), "N")
                    
                    'Veremos si la caja esta en el palet que toca o no
                    If SQL = "" Then
                        'Si no tieene palet asignado puede ser que no este creado
                        If Me.txtCPalet(0).Text <> "" Then
                            SQL = "No existe/No asignada"
                        Else
                            'Veremos si existe una trazabilidad
                            SQL = "lotetraza = " & Mid(txtCCaja.Text, 1, 8) & " AND 1"
                            SQL = DevuelveDesdeBD(conAri, "codigo", "prodtrazlin", SQL, "1", "N")
                            If SQL = "" Then
                                SQL = "No existe lote traza"
                            Else
                                SQL = "lotetraza = " & Mid(txtCCaja.Text, 1, 8) & " AND idcaja  "
                                SQL = DevuelveDesdeBD(conAri, "idcaja", "prodcajas", SQL, Mid(txtCCaja.Text, 9), "N")
                                If SQL = "" Then InsertaCaja = True
                            End If
                        End If
                    Else
                        If txtCPalet(0).Text <> "" Then
                            If Val(SQL) <> Val(Mid(Me.txtCPalet(0).Text, 2, 8)) Then
                                SQL = "Caja no pertenece al palet origen"
                            Else
                                'SIP OK. vamos a quitarla del palet.
                                SQL = ""
                                
                            End If
                        Else
                            'La caja se puede mover al palet destino. Faltar ver SI no es del palet destino
                            If txtCPalet(1).Text = "" Then
                                SQL = ""    'ok
                            Else
                                If Val(SQL) = Val(Mid(Me.txtCPalet(1).Text, 2, 8)) Then
                                    SQL = "YA esta en este palet"
                                Else
                                    SQL = ""
                                End If
                            End If
                        End If
                    End If
                End If
                
                
            End If
        End If  '"NO"
    End If
    If SQL <> "" Then AñadeACuadroMsgCboPalet SQL, False
    AñadeACuadroMsgCboPalet txtCCaja.Text, True
        

    If SQL = "" Then
        If InsertaCaja Then
            'Hay que insertar caja
            SQL = "INSERT INTO prodcajas (lotetraza ,idcaja ,idpalet ,fcreacion) VALUES ("
            SQL = SQL & Mid(txtCCaja.Text, 1, 8) & "," & Mid(txtCCaja.Text, 9) & ","
            If txtCPalet(1).Text <> "" Then
                SQL = SQL & Mid(txtCPalet(1).Text, 2, 8)
            Else
                SQL = SQL & "NULL"
            End If
            SQL = SQL & ",NOW())"
            If Not EjecutaSQL(conAri, SQL, False) Then
                AñadeACuadroMsgCboPalet "INSERTANDO CAJA", False
                Exit Sub
            End If
        End If
    
        'OK Procedemos con el UPDATE
        'Palet destino
        SQL = "NULL"
        If txtCPalet(1).Text <> "" Then SQL = Mid(txtCPalet(1).Text, 2, 8)
        SQL = "UPDATE prodcajas set idpalet = " & SQL
        SQL = SQL & " WHERE lotetraza = " & Mid(txtCCaja.Text, 1, 8) & " AND idcaja  =" & Mid(txtCCaja.Text, 9)
        If Not EjecutaSQL(conAri, SQL, False) Then
            AñadeACuadroMsgCboPalet "Error UPDATAE tabla", False
            
        Else
            'OK. Ahora sumamos una caja en uno y restamos en otro
            If txtCPalet(0).Text <> "" Then
                SQL = "UPDATE prodpalets set CajasProd = CajasProd-1 where idpalet =" & Mid(txtCPalet(0).Text, 2, 8)
                If Not EjecutaSQL(conAri, SQL, False) Then AñadeACuadroMsgCboPalet "Actualizar palet ORIGEN", False
            End If
            If txtCPalet(1).Text <> "" Then
                SQL = "UPDATE prodpalets set CajasProd = CajasProd+1 where idpalet =" & Mid(txtCPalet(1).Text, 2, 8)
                If Not EjecutaSQL(conAri, SQL, False) Then AñadeACuadroMsgCboPalet "Actualizar palet ORIGEN", False
            End If
            
            AñadeACuadroMsgCboPalet "OK" & vbCrLf & "" & vbCrLf, False
            
        End If
    End If
    txtCCaja.Text = ""
    PonerFoco txtCCaja
End Sub


'Siempre debe venir algo en el txt
'
Private Function EtiquetaPaletCorrecta2(etiquetaPalet As String) As String
    
    If Len(etiquetaPalet) <> 10 Then
        EtiquetaPaletCorrecta2 = "Longitud incorrecta"
    Else
        If Not IsNumeric(etiquetaPalet) Then
            EtiquetaPaletCorrecta2 = "Campo numerico"
        Else
            If Mid(etiquetaPalet, 1, 1) <> "1" Then
                EtiquetaPaletCorrecta2 = "Cod. control incorrecto(I)"
            Else
                If Mid(etiquetaPalet, 10, 1) <> "1" Then
                    EtiquetaPaletCorrecta2 = "Cod. control incorrecto(II)"
                Else
                    'Ok la etiqueta es correcta. Faltara ver si es de algun palet, pero eso es otra historia
                    EtiquetaPaletCorrecta2 = ""
                End If
            End If
        End If
    End If
 

End Function

Private Sub DesablarBotonCierrePalet(Cierre As Boolean)
Dim i As Byte
Dim J As Byte
    If Cierre Then
        J = 0
    Else
        J = 10
    End If
    
    For i = 0 To 5
        cmdCierrePalet(J + i).Enabled = False
        cmdCierrePalet(J + i).FontBold = False
    Next
End Sub

Private Sub ponerOpcionesCierrePalet()
Dim Indice As Byte
    'Aqui
    LimpiarCierrePalet
    DesablarBotonCierrePalet True
    
    
    
    SQL = "select LineaPeletiza from prodpalets where  fhFin is null group by 1"  '
    Set MiRsAux = New ADODB.Recordset
    MiRsAux.Open SQL, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    While Not MiRsAux.EOF
        If MiRsAux.Fields(0) > 0 Then
            Indice = MiRsAux.Fields(0) - 1
            Me.cmdCierrePalet(Indice).Enabled = True
            Me.cmdCierrePalet(Indice).FontBold = True
        End If
        MiRsAux.MoveNext
    Wend
    
    MiRsAux.Close
    Set MiRsAux = Nothing
    Me.FrameCierrePalet.Visible = True
    PonerFoco txtCierrPalet(Indice)
End Sub
    
    
 Private Sub ponerOpcionesAjustePalet()
 Dim Indice As Byte
 
    'Aqui
    LimpiarCierrePalet
    DesablarBotonCierrePalet False
    
    SQL = "select LineaPeletiza from prodpalets where  fhFin is null group by 1"  '
    Set MiRsAux = New ADODB.Recordset
    MiRsAux.Open SQL, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    While Not MiRsAux.EOF
        If MiRsAux.Fields(0) > 0 Then
            Indice = (MiRsAux.Fields(0) - 1) + 10
            'Me.cmdCierrePalet((miRsAux.Fields(0) - 1) + 10).Enabled = True
            Me.cmdCierrePalet(Indice).Enabled = True
            Me.cmdCierrePalet(Indice).FontBold = True
        End If
        MiRsAux.MoveNext
    Wend
    
    MiRsAux.Close
    Set MiRsAux = Nothing
    Me.FrAjusteCajas.Visible = True
    PonerFoco txtCierrPalet(Indice)
End Sub
    
    
Private Sub LimpiarCierrePalet()
    Label8(0).Caption = "": Label9(0).Caption = ""
    Me.txtCierrPalet(0).Text = "": Me.txtCierrPalet(1).Text = "": Me.txtCierrPalet(2).Text = ""
    txtCajaCierre(0).Text = ""
    
    
    Label8(1).Caption = "": Label9(1).Caption = ""
    Me.txtCierrPalet(3).Text = "": Me.txtCierrPalet(4).Text = ""
    txtCajaCierre(1).Text = ""
    
End Sub



Private Function LeerLineaPalet(KLinea As Byte) As Boolean
Dim i As Byte
Dim SQL As String

    LeerLineaPalet = False
    Set MiRsAux = New ADODB.Recordset
    SQL = "select idpalet from prodpalets where  LineaPeletiza = " & KLinea & " and fhFin is null "  '
    MiRsAux.Open SQL, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    i = 0
    If Not MiRsAux.EOF Then
        'UNO tiene seguro
        Set cPal = Nothing
        Set cPal = New CPalet
        cPal.Leer MiRsAux!IdPalet
        
        'Añadimos el hco de trazabilidads del palet
        cPal.TodasLasTrazabilidades
        i = 1
        
        'veremos que SOLO hay una linea en marcha
        Do
            MiRsAux.MoveNext
        
            If Not MiRsAux.EOF Then
                
                
                    'Hay mas de una produccion en la linea
                    Set cPal = Nothing
                    i = 2
                
            End If
        Loop Until MiRsAux.EOF
    
    End If
    MiRsAux.Close
    Set MiRsAux = Nothing
    If i = 1 Then LeerLineaPalet = True
End Function



Private Sub PonerDatosLineaPalet2(linea As Byte, CierrePalet As Boolean)
Dim CajasPalet As Integer
Dim SQL As String
Dim i As Byte

   'Si esta en cierre palet sera unos textos y si no otros
    i = 0
    If Not CierrePalet Then i = 1

    ValoresLeidos = ""
    SQL = cPal.TrazabilidadPaletizando
    Do
        idTrazaAntiguo = InStr(1, SQL, "|")
        If idTrazaAntiguo = 0 Then
            SQL = ""
        Else
            ValoresLeidos = ValoresLeidos & ", " & Mid(SQL, 1, idTrazaAntiguo - 1)
            SQL = Mid(SQL, idTrazaAntiguo + 1)
        End If
    Loop Until SQL = ""
    ValoresLeidos = Mid(ValoresLeidos, 2) 'quito la primera coma
    If ValoresLeidos = "" Then
        'ERROR
        MsgBox "Error leyendo trazabilidad", vbExclamation
        ValoresLeidos = "-1"
    End If
            
    Label8(i).Caption = linea
    txtCajaCierre(i) = ""
    SQL = "select prodlin.codartic,nomartic from prodlin,prodtrazlin,sartic  where prodlin.codigo= prodtrazlin.codigo"
    SQL = SQL & " AND prodlin.idlin = prodtrazlin.idlin and sartic.codartic=prodlin.codartic and lotetraza in "
    SQL = SQL & "(" & ValoresLeidos & ")"
    
    Set MiRsAux = New ADODB.Recordset
    
    
  
    MiRsAux.Open SQL, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    CajasPalet = 0
    If Not MiRsAux.EOF Then
        SQL = DevuelveDesdeBD(conAri, "pal_udbas*pal_udalt", "sarti4", "codartic", MiRsAux!codartic, "T")
        If SQL = "" Then SQL = "0"
        CajasPalet = Val(SQL)
        Label9(i).Caption = MiRsAux!NomArtic
        MiRsAux.MoveNext
        While Not MiRsAux.EOF
        
            'Palet combinado
            If Label9(i).Caption <> MiRsAux!NomArtic Then
                SQL = "0"
                CajasPalet = 0
                Label9(i).Caption = "Multi"
            End If
            MiRsAux.MoveNext
        Wend
        
        
    End If
    MiRsAux.Close
    Me.txtCierrPalet(0).Text = CajasPalet
    
    
    'Asignacion cajas NULL
    '-----------------------------------------
    SQL = "Select count(*) from prodcajas where idpalet is null AND lotetraza IN "
    SQL = SQL & "(" & ValoresLeidos & ")"
    MiRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    idTrazaAntiguo = 0
    If Not MiRsAux.EOF Then
        idTrazaAntiguo = DBLet(MiRsAux.Fields(0), "N")
    End If
    MiRsAux.Close
    
    If idTrazaAntiguo > 0 Then
        MsgBox "Se asignarán las cajas NULL", vbExclamation
        SQL = "UPDATE prodcajas set idpalet = " & cPal.ID & " where idpalet is null AND lotetraza IN "
        SQL = SQL & "(" & ValoresLeidos & ")"
        Conn.Execute SQL
        Conn.Execute "commit"   'tipo flush
        Espera 0.5
    End If
    
    'cajas leidas
    ValoresLeidos = "select count(*) from prodcajas where idpalet = " & cPal.ID
    MiRsAux.Open ValoresLeidos, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    idTrazaAntiguo = 0
    If Not MiRsAux.EOF Then
        idTrazaAntiguo = DBLet(MiRsAux.Fields(0), "N")
    End If
    MiRsAux.Close
    Set MiRsAux = Nothing
    If CierrePalet Then
        Me.txtCierrPalet(1).Text = idTrazaAntiguo
        If CajasPalet < idTrazaAntiguo Then
            Me.txtCierrPalet(2).Text = CajasPalet
        Else
            Me.txtCierrPalet(2).Text = idTrazaAntiguo
        End If
    Else
        Me.txtCierrPalet(3).Text = idTrazaAntiguo
        Me.txtCierrPalet(4).Text = idTrazaAntiguo
    End If
    
    If CierrePalet Then
        If Val(SQL) > 0 Then
            'Debia haber sql cajas
            If Val(SQL) < idTrazaAntiguo Then Me.txtCierrPalet(2).Text = SQL
        End If
            
        PonerFoco txtCierrPalet(2)
    
    Else
         PonerFoco txtCajaCierre(1)
    End If
End Sub



Private Function CerrarPalet_() As Boolean
Dim CPN As CPalet
Dim HayQueAbrirOtroPalet2 As Boolean
Dim Cajas As Integer
Dim InsertarCajasNoLeidasPost  As Integer

Dim ColTrazas As Collection   'Tendremos que trazas


    CerrarPalet_ = False
    
    
    
    CargaColTrazasPalet ColTrazas, Cajas
    
    HayQueAbrirOtroPalet2 = True
    'Mayo 2012
    'Siempre iniciara un nuevo palet EXCEPTO cuando se marque el chec Me.chkNoContinuar
    'Si no esta marcado y hubiera que abrir otro palet lo indicaremos
   
    SQL = "Va a cerrar el palet." & vbCrLf
    
    If Cajas - CInt(Me.txtCierrPalet(2).Text) > 0 Then
        
        SQL = SQL & "Ultimas: " & Cajas - CInt(Me.txtCierrPalet(2).Text) & " cajas " & vbCrLf & "     a un nuevo palet" & vbCrLf
        HayQueAbrirOtroPalet2 = True
    End If
    SQL = SQL & "   ¿CONTINUAR?"
    
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
    
            
            'Veremos si hay que meter nuevas cajas
            'Monteremos el SQL de abajo
            'idpalet=43 and (lotetraza<82 or (lotetraza=82 and idcaja<=48))
            SQL = Val(Mid(Me.txtCajaCierre(0).Text, 1, 8))  'lotetraza
            SQL = "(lotetraza<" & SQL & " or (lotetraza=" & SQL & " AND idcaja<="
            SQL = SQL & Val(Mid(Me.txtCajaCierre(0).Text, 9)) & ")) AND idpalet"
            SQL = DevuelveDesdeBD(conAri, "count(*)", "prodcajas", SQL, CStr(cPal.ID))
            If SQL = "" Then SQL = "0"
            InsertarCajasNoLeidasPost = CInt(Me.txtCierrPalet(2).Text) - Val(SQL)
            
            If InsertarCajasNoLeidasPost < 0 Then
                'Hay mas cajas de las que deberian
                MsgBox "Hay mas cajas de las que deberian: " & CInt(Me.txtCierrPalet(2).Text) & " / " & Val(SQL), vbInformation
                Exit Function
            End If
            
            If InsertarCajasNoLeidasPost > 0 Then
                If Not InsertarCajasNoLeidasPosteSub(InsertarCajasNoLeidasPost, True) Then Exit Function
                
                
                
                
                'Despues de insertar cajas, volvemos a ver si HAY qu insertar otro palet
                Espera 0.3
                SQL = DevuelveDesdeBD(conAri, "count(*)", "prodcajas", "idpalet", CStr(cPal.ID))
                If Val(SQL) - CInt(Me.txtCierrPalet(2).Text) > 0 Then
                    'HAY QUE ABRIR OTRO PALET
                    If Me.chkNoContinuar.Value = 1 Then MsgBox "Es obligado abrir otro palet", vbExclamation
                    
                    HayQueAbrirOtroPalet2 = True
                End If
            Else
                'No hay que insertar mas cajas. Si tiene la marca de que no abrir otro, no lo abriremos
                If Me.chkNoContinuar.Value = 1 Then HayQueAbrirOtroPalet2 = False
            End If
            
            
            If HayQueAbrirOtroPalet2 Then
                Set CPN = New CPalet
                CopiarPalet CPN
            End If
            
            If cPal.CerrarPalet(CInt(Me.txtCierrPalet(2).Text)) Then
                If HayQueAbrirOtroPalet2 Then cPal.PasarUltimasCajasA_OtroPalet CInt(Mid(Me.txtCajaCierre(0).Text, 9)), CPN.ID, cPal.ID, CLng(Mid(Me.txtCajaCierre(0).Text, 1, 8))
                
                
                
            
                Label10.Caption = "Imprimiendo"
                Label10.Refresh
                Dim C As Collection
                
                Conn.Execute "DELETE FROM tmppartidas WHERE codusu = " & vUsu.Codigo
                
                cPal.CargaDatosPalet C, True, CInt(Me.txtCierrPalet(2).Text), False
                ImprimirPalet cPal.ID, cPal.TipoImpresion
        
                 
            
                    
                    
                
                
                'ImprimeEtiquetaPalet cPal.ID
                Set cPal = Nothing
                CerrarPalet_ = True
            Else
                'error cerrando palet
                If Not CPN Is Nothing Then
                    'Ha ido mal arriba, deshago el nuevo pal
                      Conn.Execute "DELETE FROM prodpalets WHERE idpalet = " & CPN.ID
                      EjecutaSQL conAri, "DELETE FROM prodpaletstraza WHERE idpalet = " & CPN.ID
            
                End If
            End If
    End If
    
    Set CPN = Nothing
    Conn.Execute "commit"   'tipo flush
End Function


Private Sub CargaColTrazasPalet(ByRef C As Collection, ByRef Cajas_ As Integer)
    Set C = New Collection
    
    
    'Tengo para cada trazabilidad del palet cuantas cajas hay 0000008200048
    SQL = "select lotetraza,count(*)  from prodcajas where  idpalet=" & cPal.ID & " group by 1 ORDER BY 1"  '
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    Cajas_ = 0
    'NO PUEDE SER EOV
    If Not RS.EOF Then
        Do
            Cajas_ = Cajas_ + DBLet(RS.Fields(1), "N")
            SQL = RS.Fields(0) & "|" & DBLet(RS.Fields(1), "N") & "|"
            C.Add SQL
            RS.MoveNext
        Loop Until RS.EOF
    Else
        SQL = cPal.ID & "|0|"
    End If
    RS.Close
    
    Set RS = Nothing
    
End Sub


Private Sub CopiarPalet(ByRef CPN As CPalet)
Dim i As Integer
        
    Set CPN = New CPalet
    For i = 0 To 7
        CPN.LineasProd(i) = cPal.LineasProd(i)
    Next i


    CPN.FechaInicio = Now
    CPN.LineaPeletizacion = cPal.LineaPeletizacion
    ' Hay que ver QUE esta paletizanod AHORA, sin nada mas
    SQL = " idPalet "
    SQL = DevuelveDesdeBD(conAri, "lotetraza", "prodpaletstraza", SQL, cPal.ID & " ORDER BY 1 DESC")
    'Si no encotramos NINGUNA ponermos la que habia
    If SQL = "" Then
        SQL = cPal.TrazabilidadPaletizando
    Else
        SQL = SQL & "|"
    End If
    While SQL <> ""
        i = InStr(1, SQL, "|")
        If i = 0 Then
            SQL = ""
        Else
            CPN.AñadirIdTraza CLng(Mid(SQL, 1, i - 1))
            SQL = Mid(SQL, i + 1)
        End If
    Wend
    
    
    CPN.TipoImpresion = cPal.TipoImpresion
    CPN.CrearPalet
  
    
End Sub



'Proceso complicado
'Cuando cierra un palet y no estan todas las cajas VEREMOS


Private Function InsertarCajasNoLeidasPosteSub(CuantasInserto As Integer, CierrePalet As Boolean) As Boolean
Dim Fin As Boolean
Dim LaCaja As Integer
Dim CadenaInsertCajas As String
Dim Cual As Byte
Dim K As Byte

'quita la K
Dim MaxLoteAnterior As Long
Dim LoteTr As Long

    InsertarCajasNoLeidasPosteSub = False
   
    If CierrePalet Then
        Cual = 0
    Else
        Cual = 1
    End If
   
   
    CadenaInsertCajas = ""
    'Primera, si el txtCajaCierre.tag =1 entonces tengo que insertar esta caja SEGURO
    LaCaja = Val(Mid(txtCajaCierre(Cual).Text, 9))
    
    
    SQL = "lotetraza= " & Mid(txtCajaCierre(Cual).Text, 1, 8) & " AND idcaja "
    SQL = DevuelveDesdeBD(conAri, "idcaja", "prodcajas", SQL, CStr(LaCaja))
    If SQL = "" Then
        'Esta caja hay que insertarla
        SQL = Mid(txtCajaCierre(Cual).Text, 1, 8)
        CadenaInsertCajas = CadenaInsertCajas & ", (" & SQL & "," & LaCaja & "," & cPal.ID & "," & DBSet(Now, "FH") & ")"
        CuantasInserto = CuantasInserto - 1
        LaCaja = LaCaja - 1
        Espera 0.1
    End If
    
    'Busco hueco
    
    MaxLoteAnterior = -1   'Sabre si hay dos lotes en el palet
    If CuantasInserto > 0 Then
    
        
            'Cuando k=lotpal.count significa que es el ultimo lote procesado en el palet
            Set RS = New ADODB.Recordset
            SQL = Mid(txtCajaCierre(Cual).Text, 1, 8)
            LoteTr = Val(SQL)
            SQL = " and (lotetraza<" & SQL & " or (lotetraza=" & SQL & " and idcaja<=" & Mid(Me.txtCajaCierre(Cual).Text, 9) & "))"
            SQL = "select * from prodcajas where idpalet=" & cPal.ID & SQL
            SQL = SQL & " order by lotetraza desc,idcaja desc"
            'Salgo
            RS.Open SQL, Conn, adOpenKeyset, adLockPessimistic, adCmdText
            
            If RS.EOF Then
                'Las creamos todas?
            Else
                
                While Not Fin
                        
                        If RS!lotetraza = LoteTr Then
                            
                            If LaCaja = 0 Then
                                'Stop   'NO PUEDE INSERTARSE CAJA 0, -1,-2...
                                MsgBox "No se ha encontrado hueco. Caja=0. ", vbExclamation
                                RS.Close
                                Exit Function
                            End If
                        Else
                            'Teneia mas de un lote el palet
                            
                            'NO COMPROBAREMOS EL SIGUIENTE LOTE ya que deberia estar cerrado correctamente
                            If LaCaja > 0 Then
                                RS.MovePrevious
                            Else
                                MsgBox "No se ha encotrado hueco", vbExclamation
                                RS.Close
                                Exit Function
                            End If
                        End If
                        
                        If RS!idcaja = LaCaja Then
                            LaCaja = LaCaja - 1
                            RS.MoveNext
                            
                            
                            'If LaCaja = 3 Then Stop
                            
                        Else
    
                            ' prodcajas lotetraza idcaja idpalet fcreacion
                            'Compruebo que la caja NO esta en otro palet
                            SQL = "lotetraza = " & LoteTr & " AND idcaja"
                            SQL = DevuelveDesdeBD(conAri, "idcaja", "prodcajas", SQL, CStr(LaCaja))
                            'Si SQL<>"" siginifica que la caja existe ya
                            If SQL = "" Then
                                SQL = LoteTr
                                CadenaInsertCajas = CadenaInsertCajas & ", (" & SQL & "," & LaCaja & "," & cPal.ID & "," & DBSet(Now, "FH") & ")"
                                LaCaja = LaCaja - 1
                                CuantasInserto = CuantasInserto - 1
                                If CuantasInserto = 0 Then Fin = True
                            Else
                                RS.MoveNext 'La caja esta en otro palet
                                LaCaja = LaCaja - 1
                            End If
                        End If
                
                        If RS.EOF Then Fin = True
                Wend
            End If
            
            RS.Close
            Set RS = Nothing
            If CuantasInserto > 0 Then 'HAy que insertarla
                If LaCaja > 0 Then
                    SQL = LoteTr
                    While LaCaja > 0
                        CadenaInsertCajas = CadenaInsertCajas & ", (" & SQL & "," & LaCaja & "," & cPal.ID & "," & DBSet(Now, "FH") & ")"
                        LaCaja = LaCaja - 1
                        CuantasInserto = CuantasInserto - 1
                        If CuantasInserto <= 0 Then LaCaja = 0
                    Wend
                
                Else
                    'LACAJA=0 'no ha encontrado el hueco
                    'Si tenia mas de un lote podemos ir hacia arriba a partir de la ultima caja leida
                    If MaxLoteAnterior > 0 Then
                        
                        SQL = LoteTr
                        While CuantasInserto > 0
                            MaxLoteAnterior = MaxLoteAnterior + 1
                            CadenaInsertCajas = CadenaInsertCajas & ", (" & SQL & "," & MaxLoteAnterior & "," & cPal.ID & "," & DBSet(Now, "FH") & ")"
                            LaCaja = LaCaja - 1
                            CuantasInserto = CuantasInserto - 1
                            If CuantasInserto <= 0 Then LaCaja = 0
                        Wend
                        
                        
                    End If
                End If
            End If
            If LaCaja = 0 Then
                If CuantasInserto > 0 Then
                    MsgBox "Falta insertar:" & CuantasInserto & " caja(s)", vbExclamation
                    CadenaInsertCajas = ""
                End If
            End If
    
    
    
        
    End If  'de cuantas>0
    
    
    If CadenaInsertCajas <> "" Then
        'Insertamos
        CadenaInsertCajas = Mid(CadenaInsertCajas, 2)
        SQL = "INSERT INTO prodcajas (lotetraza ,idcaja ,idpalet ,fcreacion) VALUES " & CadenaInsertCajas
        If Not EjecutaSQL(conAri, SQL, False) Then
        
            'Abril 2012,25
            'Si da error haremos lo siguiente, preguntaremos si quiere continuar
            If MsgBox("Error ajustando cajas. ¿Continuar?", vbQuestion + vbYesNo) = vbNo Then
                InsertarCajasNoLeidasPosteSub = False
            Else
                SQL = DBSet(SQL, "T")
                SQL = "insert into proderrorcierrepalet(Fechora,idpalet,observaciones) values (now()," & cPal.ID & "," & SQL & ")"
                EjecutaSQL conAri, SQL, False
                InsertarCajasNoLeidasPosteSub = True
            End If
        Else
            InsertarCajasNoLeidasPosteSub = True
        End If
    End If
   
End Function



Private Function AsignacionPrimerLoteProduccion() As Boolean

    On Error GoTo EAsignacionPrimerLoteProduccion
    'Si hay una etiqueta libre
    AsignacionPrimerLoteProduccion = False
    
    If Not cLinPr.AsignarLoteLinea(IndexSublinea, Cp.NUmlote, True) Then Exit Function
     
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

        End If
    Else
        'Si que hay libre
        SQL = " WHERE id = " & Cp.idPartida & " AND bulto = " & SQL
        SQL = "UPDATE spartidaslin Set fechaulizada = " & DBSet(Now, "FH") & SQL
        EjecutaSQL conAri, SQL, True
        
    End If
    AsignacionPrimerLoteProduccion = True
    Exit Function
EAsignacionPrimerLoteProduccion:
    MuestraError Err.Number, Err.Description
End Function


Private Sub txtMoverUltimasNcajas_LostFocus()
    txtMoverUltimasNcajas.Text = Trim(txtMoverUltimasNcajas.Text)
    
    If txtMoverUltimasNcajas.Text = "" Then txtMoverUltimasNcajas.Text = "1"
    If Not IsNumeric(txtMoverUltimasNcajas.Text) Then txtMoverUltimasNcajas.Text = "1"
    
    
    
End Sub



Private Function LeerPaletExpedicionParaVariosAlbaranes(Comprobar As Boolean) As String
Dim Co As Collection
Dim i As Integer
Dim Aux As Integer
Dim CajasPorCargar As Integer
Dim J As Integer
Dim RN As ADODB.Recordset
Dim CargaPalet As Boolean
Dim CadeLot As String

    

    
    'Primera comprobacion. Palet con solo un tipo de cosas
    SQL = "Select cajasprod from prodpalets where idpalet=" & Mid(Text13(0).Text, 2, 8)
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 0
    SQL = "No se encuentra el palet"
    If Not RS.EOF Then
        i = DBLet(RS!Cajasprod, "N")
        If i = 0 Then
            SQL = "No tiene cajas el palet"
        Else
            SQL = ""
        End If
    End If
    RS.Close
    
    If SQL <> "" Then
        LeerPaletExpedicionParaVariosAlbaranes = SQL
        Exit Function
    End If
        

    'Veremos si las cajas dispobibles son las que tienen el palet y si solo hay una referencia
    SQL = "select codartic,count(*) from prodlin,prodtrazlin,prodcajas  where  prodlin.Codigo = prodtrazlin.Codigo "
    SQL = SQL & " AND prodlin.idlin = prodtrazlin.idlin And prodtrazlin.lotetraza = prodcajas.lotetraza"
    SQL = SQL & " AND idpalet=" & Mid(Text13(0).Text, 2, 8) & " group by 1"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        LeerPaletExpedicionParaVariosAlbaranes = "Error leyendo datos trazab. palet: " & Mid(Text13(0).Text, 2, 8)
        RS.Close
        Exit Function
    
    End If

    Set Co = New Collection
    While Not RS.EOF
          SQL = RS!codartic & "|" & RS.Fields(1) & "|"
          Co.Add SQL
          RS.MoveNext
    Wend
    
    RS.Close
            
            
            
     If Comprobar Then
            SQL = ""
            If Co.Count > 1 Then
                SQL = "mas de una referencia en el palet"
            Else
                SQL = RecuperaValor(Co.Item(1), 2)
                If Val(SQL) <> i Then
                    SQL = "Cajas palet distinta de disponible. " & SQL & " - Mto : " & i
                Else
                    SQL = ""
                End If
            End If
            
            If SQL <> "" Then
                LeerPaletExpedicionParaVariosAlbaranes = SQL
                Exit Function
            End If
            
            
            
            SQL = "Select * from prodcajas where idpalet=" & Mid(Text13(0).Text, 2, 8) & " AND (lotetraza,idcaja) IN ("
            SQL = SQL & " SELECT lotetraza,idcaja from srepartolotcaj)"
            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            SQL = ""
            If Not RS.EOF Then
                
                While Not RS.EOF
                    SQL = SQL & Format(RS!lotetraza, "00000000") & Format(RS!idcaja, "00000") & "    "
                    RS.MoveNext
                Wend
            
                
            End If
            RS.Close
            If SQL <> "" Then
                SQL = "Cajas YA asignadas" & vbCrLf & SQL
                LeerPaletExpedicionParaVariosAlbaranes = SQL
                Exit Function
            End If
            
            
            
            For i = 1 To Co.Count
                 SQL = "select codartic,sum(cajas) vcajas, sum(llevamos) vllevamos from srepartolot "
                 SQL = SQL & " WHERE idreparto=" & Label5(0).Tag & "  and codartic=" & DBSet(RecuperaValor(Co.Item(i), 1), "T")
                 SQL = SQL & " group by 1"
                 RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                 SQL = ""
                 If RS.EOF Then
                    'Error grave. No se encuentra id/alb/art "
                    SQL = "Error grave. No se encuentra articulo de orden de carga" & Label5(1).Tag & " /" & RecuperaValor(Co.Item(i), 1)
                 Else
                     CajasPorCargar = RS!vcajas - RS!vllevamos
                     Aux = Val(RecuperaValor(Co.Item(i), 2))
                     If CajasPorCargar < Aux Then SQL = RS!codartic & vbCrLf & "  Falta: " & CajasPorCargar & " Palet: " & Aux
                     
                 End If
                 RS.Close
                 If SQL <> "" Then
                    Set Co = Nothing
                    LeerPaletExpedicionParaVariosAlbaranes = SQL
                    Exit Function
                 End If
            Next
                
                
            'Otra comprobacion.
            'Veremos si los lotes tienen algo mas antiguo
            ValoresLeidos = ""
            For i = 1 To Co.Count
                 SQL = "select distinct(lotetraza) from prodcajas where idpalet=" & Mid(Text13(0).Text, 2, 8)
                 RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                 CadeLot = "|"
                 While Not RS.EOF
                    CadeLot = CadeLot & RS.Fields(0) & "|"
                    RS.MoveNext
                 Wend
                 RS.Close
                     
                     
                 SQL = "select * from spartidas where  cantotal>0 and codartic=" & DBSet(RecuperaValor(Co.Item(i), 1), "T")
                 SQL = SQL & " order by fecha asc"
                 RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                 SQL = ""
                 If Not RS.EOF Then
                    SQL = RS!NUmlote
                    If InStr(1, SQL, " ") > 0 Then
                        'ANtiguo
                        SQL = Trim(Mid(SQL, 1, InStr(1, SQL, " ")))
                    Else
                        'No hacemos nada, esta bien
                        SQL = Val(SQL)
                    End If
                    
                    SQL = "|" & SQL & "|"
                    If InStr(1, CadeLot, SQL) = 0 Then
                        'Significa que hay un lote antorior
                        SQL = Mid(SQL, 2)
                        SQL = Mid(SQL, 1, Len(SQL) - 1)
                        ValoresLeidos = ValoresLeidos & vbCrLf & RecuperaValor(Co.Item(i), 1) & ": " & SQL
                    End If
                    
                 End If
                 CadeLot = ""
                    
                 
                 RS.Close
                 
            Next
            If ValoresLeidos <> "" Then
                ValoresLeidos = "Lotes anteriores" & ValoresLeidos & "   ¿Continuar?"
                If MsgBox(ValoresLeidos, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                    Set Co = Nothing
                    LeerPaletExpedicionParaVariosAlbaranes = ValoresLeidos
                    Exit Function
                End If
            End If
                
                
                
            'Cargaremos en un txt didicendo que albaranres vamos a ponerles las cajas
            CadeLot = "Palet: " & Mid(Text13(0).Text, 2, 8) & vbCrLf & vbCrLf
            CadeLot = CadeLot & "Albaran      Cargadas        Palet" & vbCrLf
            CadeLot = CadeLot & String(25, "=") & vbCrLf
            SQL = "select * from srepartolot where idreparto=" & Label5(0).Tag & " and codartic=" & DBSet(RecuperaValor(Co.Item(1), 1), "T") & " ORDER By numalbar"
            CajasPorCargar = Val(RecuperaValor(Co.Item(1), 2))
            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            SQL = ""
            idTrazaAntiguo = -1
            
            While Not RS.EOF
                i = RS!Cajas - RS!llevamos
                
                'Aux: albaran expedido
                'idTrazaAntiguo: albaran
                If RS!NumAlbar <> idTrazaAntiguo Then
                    idTrazaAntiguo = RS!NumAlbar
                    SQL = "id=" & Label5(0).Tag & " and numalbar"
                    SQL = DevuelveDesdeBD(conAri, "albexpedido", "srepartol", SQL, CStr(idTrazaAntiguo))
                    Aux = 0
                    If SQL = "1" Then Aux = 1
                    
                End If
                
                SQL = Format(RS!NumAlbar, "000000") & "        " & Format(RS!llevamos, "0000") & "           "
                
                If Aux = 1 Then
                    'ALBARAN EXPEDIDO
                    SQL = SQL & "    *"
                Else
                    If i > 0 And CajasPorCargar > 0 Then
                        
                    
                        CajasPorCargar = CajasPorCargar - i
                        SQL = SQL & Format(i, "0000")
                    Else
                        SQL = SQL & "    -"
                    End If
                
                End If
                CadeLot = CadeLot & SQL & vbCrLf
                RS.MoveNext
            Wend
            RS.Close
            Text13(2).Text = CadeLot
            Text13(2).Visible = True
            PonerFoco Text13(2)
            Me.cmdAceptarPaletVariosAlb.Visible = True
            
            
    Else
        'Asignar las cajas a los palets
            SQL = "select * from srepartolot where idreparto=" & Label5(0).Tag & " and codartic=" & DBSet(RecuperaValor(Co.Item(1), 1), "T") & " ORDER By numalbar"
            CajasPorCargar = Val(RecuperaValor(Co.Item(1), 2))
            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            SQL = ""
            
            SQL = "select * from prodcajas where idpalet=" & Mid(Text13(0).Text, 2, 8) & " and not (lotetraza,idcaja)  in "
            SQL = SQL & " (select lotetraza,idcaja FROM srepartolotcaj "
            SQL = SQL & " WHERE idreparto=" & Label5(0).Tag & ") ORDER BY idCaja"
            Set RN = New ADODB.Recordset
            RN.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            'NO PUEDE SER EOF
            
            
            While Not RS.EOF
                
                i = RS!Cajas - RS!llevamos
                'SQL = Format(RS!NumAlbar, "000000") & "        " & Format(RS!llevamos, "0000") & "           "
                
                SQL = ""
                If i > 0 And CajasPorCargar > 0 Then
                    For J = 1 To i
                        If RN.EOF Then
                            LeerPaletExpedicionParaVariosAlbaranes = "Error. Todas las cajas asignadas. No se puede continuar"
                        Else
                        
                        'INSERT INTO  srepartolotcaj(idreparto,codtipom,numalbar,numlinea,codartic,fecha,lotetraza,idcaja,idPalet) VALUES "
                            SQL = SQL & ", (" & Label5(0).Tag & ",1," & RS!NumAlbar & "," & RS!numlinea & "," & DBSet(RS!codartic, "T") & ",now(),"
                            SQL = SQL & RN!lotetraza & "," & RN!idcaja & "," & RN!IdPalet & ")"
                            RN.MoveNext
                        End If
                    Next
                    If SQL <> "" Then
                        SQL = Mid(SQL, 2)
                        SQL = "INSERT INTO  srepartolotcaj(idreparto,codtipom,numalbar,numlinea,codartic,fecha,lotetraza,idcaja,idPalet) VALUES " & SQL
                        If Not EjecutaSQL(conAri, SQL) Then
                            
                            LeerPaletExpedicionParaVariosAlbaranes = "Error asignado cajas del palet"
                            RS.Close
                            RN.Close
                            Exit Function
                        End If
            
                        
                        SQL = "UPDATE srepartolot set llevamos=llevamos + " & i
                        'idreparto codtipom numalbar numlinea codartic
                        SQL = SQL & " WHERE idreparto =" & RS!idreparto & " and codartic=" & DBSet(RS!codartic, "T")
                        SQL = SQL & " AND codtipom =" & RS!Codtipom & " and numalbar=" & RS!NumAlbar
                        SQL = SQL & " AND numlinea =" & RS!numlinea
                        Conn.Execute SQL
                    End If
                End If
                
                RS.MoveNext
            Wend
            RS.Close
            Text13(1).Text = "Cajas asignadas"
            Text13(2).Visible = False
            PonerFoco Text13(0)
            Me.cmdAceptarPaletVariosAlb.Visible = False
          
    End If
        
        
End Function
    
'Private Sub InsertarCajasAlbaranesDesdeUnPalet()
        
'    'Hacemos el insert
'    'Iremos metiendo en el CO los inserts
'    Set RN = New ADODB.Recordset
'    For i = 1 To Co.Count
'        'Cada articulo vere en el albaran cuantas lineas voy a meter
'        'SOLO deberia haber un codartic.
'        SQL = "select codartic,cajas,llevamos,codtipom,numalbar,numlinea  from srepartolot "
'        SQL = SQL & " WHERE idreparto=" & Label5(0).Tag & " and numalbar=" & Label5(1).Tag & " and codartic=" & DBSet(RecuperaValor(Co.Item(i), 1), "T")
'        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        Aux = Val(RecuperaValor(Co.Item(i), 2))
'        'NO DEBE SER EOF
'        While Not RS.EOF
'             CajasPorCargar = RS!Cajas - RS!llevamos
'             If CajasPorCargar >= Aux Then
'                  'De todas las del palet cojere las que son de este articulo
'                  SQL = "select distinct(prodcajas.lotetraza) from prodlin,prodtrazlin,prodcajas  where "
'                  SQL = SQL & " prodlin.Codigo = prodtrazlin.Codigo  AND prodlin.idlin = prodtrazlin.idlin And"
'                  SQL = SQL & " prodtrazlin.lotetraza = prodcajas.lotetraza AND idpalet=" & Mid(Text10.Text, 2, 8)
'                  SQL = SQL & " AND codartic=" & DBSet(RS!codArtic, "T")
'                  J = 0
'                  RN.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'                  CadeLot = ""
'                  While Not RN.EOF
'                        J = J + 1
'                        CadeLot = CadeLot & ", " & RN.Fields(0)
'                        RN.MoveNext
'                  Wend
'                  RN.Close
'
'                  If J = 0 Then
'                    'DEBERIA HABER ENCONTRADO un lote por lo menos
'                     Set Co = Nothing
'                     RS.Close
'
'
'                     LeerPaletExpedicionParaVariosAlbaranes = "Palet. NO encuentra lote trazabilidad"
'                     Exit Sub
'
'                  End If
'                  CadeLot = Mid(CadeLot, 2)
'
'                  If J = 1 Then
'                        CadeLot = " AND lotetraza = " & CadeLot
'                  Else
'                        CadeLot = " AND lotetraza IN (" & CadeLot & ")"
'                  End If
'                  'Voy a insertar cajas
'                  SQL = "select * from prodcajas where idpalet=" & Mid(Text10.Text, 2, 8) & CadeLot & " and not (lotetraza,idcaja)  in "
'                  SQL = SQL & " (select lotetraza,idcaja FROM srepartolotcaj "
'                  SQL = SQL & " WHERE idreparto=" & Label5(0).Tag & " and numalbar=" & Label5(1).Tag & " and codtipom=1 ) ORDER BY idCaja"
'                  RN.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'                  LeerPaletExpedicionParaVariosAlbaranes = ""
'                  J = 0
'                  SQL = ""
'                  Do
'                    If RN.EOF Then
'                        'YA NO inserto mas
'                        'J = CajasPorCargar
'                        CajasPorCargar = J    'No cambio J que despues se utiliza bajo
'                    Else
'                        J = J + 1
'                        'insert
'                        'INSERT INTO  srepartolotcaj(idreparto,codtipom,numalbar,numlinea,codartic,fecha,lotetraza,idcaja,idPalet) VALUES ("
'                        SQL = SQL & ", (" & Label5(0).Tag & ",1," & Label5(1).Tag & "," & RS!numlinea & "," & DBSet(RS!codArtic, "T") & ",now(),"
'                        SQL = SQL & RN!lotetraza & "," & RN!idcaja & "," & RN!IdPalet & ")"
'                        RN.MoveNext
'
'                    End If
'                  Loop Until J >= CajasPorCargar
'                  RN.Close
'
'                  'Actualiz lllevamps
'                  If J > 0 Then
'                        'Insertamos cajas
'                        SQL = Mid(SQL, 2) 'quito la coma
'                        SQL = "INSERT INTO  srepartolotcaj(idreparto,codtipom,numalbar,numlinea,codartic,fecha,lotetraza,idcaja,idPalet) VALUES " & SQL
'                        If Not EjecutaSQL(conAri, SQL, False) Then
'                            'ERRROR
'                            Set Co = Nothing
'                            RS.Close
'
'
'                            LeerPaletExpedicionParaVariosAlbaranes = "Error insertando cajas repartol"
'                            Exit Sub
'                        End If
'                        'Actualizmos cajas
'                        SQL = " where idreparto=" & Label5(0).Tag & " and numalbar=" & Label5(1).Tag & " and codtipom=1 AND numlinea = " & RS!numlinea
'                        SQL = "UPDATE srepartolot set llevamos=llevamos + " & J & SQL
'                        EjecutaSQL conAri, SQL, True  'NO DEBERIA DAR ERROR
'                        Espera 0.2
'                  End If
'
'            End If
'            RS.MoveNext
'        Wend
'        RS.Close
'    Next
'
'
'    Set RN = Nothing
'End Sub





Private Sub AjusteCajasPaletsPorCierreOrdenExpedicion()
Dim RN As ADODB.Recordset
Dim i As Integer
On Error GoTo eAjusteCajasPaletsPorCierreOrdenExpedicion

    
    idTrazaAntiguo = CLng(Me.Label5(0).Tag) 'id reparto
    
    'Metemos en el HCO de ordenes de expedicion
    'las cajasxpalet que vamos a llevar
    SQL = "insert into srepartohco"
    SQL = SQL & " select " & idTrazaAntiguo & ",idpalet,count(*),now() from prodcajas where (lotetraza,idcaja)"
    SQL = SQL & " in (select lotetraza,idcaja from srepartolotcaj where idreparto=" & idTrazaAntiguo & " ) group by idpalet"
    Conn.Execute SQL

    Set RN = New ADODB.Recordset
    'Actualizamos los palets con las nuevas cantidas
    SQL = " select idpalet,count(*) ncaj from prodcajas where (lotetraza,idcaja)"
    SQL = SQL & " in (select lotetraza,idcaja from srepartolotcaj where idreparto=" & idTrazaAntiguo & " ) group by idpalet"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        SQL = "Select * from  prodpalets where idpalet=" & RS!IdPalet
        i = DBLet(RS!ncaj, "N")
        RN.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RN.EOF Then
            i = RN!Cajasprod - i
            If i < 0 Then
                i = 0
                'Habria que meter un LOG
            End If
            SQL = "UPDATE prodpalets set cajasprod=" & i & " WHERE idpalet=" & RS!IdPalet
            Conn.Execute SQL
        End If
        RN.Close
        RS.MoveNext
    Wend
    RS.Close
    
    
    
    'Borramos las cajas que esten en la orden de carga
    SQL = "delete from prodcajas where (lotetraza,idcaja)"
    SQL = SQL & " in (select lotetraza,idcaja from srepartolotcaj where idreparto=" & idTrazaAntiguo & " )"
    Conn.Execute SQL
    
eAjusteCajasPaletsPorCierreOrdenExpedicion:
    If Err.Number <> 0 Then MuestraError Err.Number, "Ajuste cajas-palets. Consulte soporte tecnico"
    idTrazaAntiguo = 0
    Set RN = Nothing
End Sub



Private Sub ComprobarLuces()
Dim Aux As String
Dim N As Long
   Aux = DevuelveDesdeBD(conAri, "count(*)", "prodcajasduplicadas", "1", "1")
   If Aux = "" Then Aux = "0"
   Me.imgRepetidos.Visible = Val(Aux) > 0
   
   SQL = "now()"
   Aux = DevuelveDesdeBD(conAri, "max(fechahora)", "prodlecturaposte", "1", "1", "N", SQL)
   If Aux <> "" Then
        N = DateDiff("n", Aux, SQL)
        If N < 4 Then Aux = ""
   End If
   Me.imgPoste.Visible = Aux <> ""
        
End Sub



'-------------------------------------------------------
' Devolucion mercancia
Private Function LeerCajaDevolucion() As String
Dim EnEsta As Boolean
Dim Articu As String
Dim IT As ListItem


    On Error GoTo eLeerCajaDevolucion

        'Ahora veremos si pertence a una orden de carga
        SQL = Mid(Text14.Text, 1, 8) & " AND idcaja = " & Val(Mid(Text14, 9))
        SQL = "Select * from srepartolotcaj WHERE lotetraza = " & SQL
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        idTrazaAntiguo = 0
        If RS.EOF Then
            LeerCajaDevolucion = "caja NO esta expedida"
        Else
            'Vemos si lo que lee es lo que escribe ;
            'Si realmente la caja es de lo que me ha dicho en el albarna
            idTrazaAntiguo = DBLet(RS!idreparto, "N")
            Articu = RS!codartic
        End If
        RS.Close
        
        
        
        If idTrazaAntiguo = 0 Then Exit Function
        
        'Veremos si la orden de carga esta expedida
        'Veremos si ya hay una entrada de esa orden de carga(para no tener que hacer un select cada vez)
        SQL = "N" 'hay que hacer select
        For IndexSublinea = 1 To Me.ListView1(2).ListItems.Count
            If Val(Me.ListView1(2).ListItems(IndexSublinea).SubItems(1)) = idTrazaAntiguo Then
                SQL = ""
                Exit For
            End If
        Next
        
        If SQL <> "" Then
            'Hay que comprobar si la orden de carga esta EXPEDIDA
            SQL = "SELECT * from srepartoc WHERE id = " & CStr(idTrazaAntiguo)
            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            SQL = ""
            If RS.EOF Then
                 SQL = "leyendo expedicion " & idTrazaAntiguo
            Else
                If DBLet(RS!Situacion, "N") < 3 Then SQL = "situacion expedicion"
            End If
            RS.Close
            If SQL <> "" Then
                LeerCajaDevolucion = SQL
                Exit Function
            End If
        End If
        
        
        'TODO correcto. Añadiremos la caja al lw de devolucion
        SQL = "SELECT codartic,nomartic from sartic WHERE codartic = " & DBSet(Articu, "T")
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Articu = RS!NomArtic  'NO puede ser EOF
        SQL = RS!codartic
        RS.Close
        
        Set IT = Me.ListView1(2).ListItems.Add(, "T" & Text14.Text)
        IT.Text = Text14.Text
        IT.SubItems(1) = idTrazaAntiguo
        IT.SubItems(2) = Articu
        IT.Tag = SQL
        
eLeerCajaDevolucion:
    If Err.Number <> 0 Then
        LeerCajaDevolucion = Err.Description
        Err.Clear
    End If
        
End Function



Private Sub RealizarProcesoDevolucion()
Dim IdDev As Integer


    Set RS = New ADODB.Recordset
    'Veremos que id le asignamos
    SQL = "Select max(IdDevol) from sreparto_dev"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    IdDev = 0
    If Not RS.EOF Then IdDev = DBLet(RS.Fields(0), "N")
    IdDev = IdDev + 1
    RS.Close
    
    
    
    'Ahora meteremos en una tabla las cajas que devolvemos tal como estaban
    ValoresLeidos = ""
    For idTrazaAntiguo = 1 To Me.ListView1(2).ListItems.Count
        ValoresLeidos = ValoresLeidos & ", (" & Mid(ListView1(2).ListItems(idTrazaAntiguo).Text, 1, 8)
        ValoresLeidos = ValoresLeidos & "," & Mid(ListView1(2).ListItems(idTrazaAntiguo).Text, 9) & ")"
    Next
    
    'NO puede ser ""
    ValoresLeidos = Mid(ValoresLeidos, 2)
    
    SQL = "INSERT INTO sreparto_dev(IdDevol,codtraba,fechadev,idreparto,numalbar,"
    SQL = SQL & "numlinea,codartic,lotetraza,idcaja,idpalet) "
    SQL = SQL & "SELECT " & IdDev & "," & vUsu.CodigoTrabajador & " codtraba, now() fechadev,"
    SQL = SQL & " idreparto,numalbar,numlinea,codartic,lotetraza,idcaja,idpalet "
    SQL = SQL & " from srepartolotcaj  WHERE (lotetraza,idcaja) IN (" & ValoresLeidos & ")"
    
    Conn.Execute SQL
    
    'Borramos de serpartolotcag
    SQL = "DELETE from srepartolotcaj  WHERE (lotetraza,idcaja) IN (" & ValoresLeidos & ")"
    Conn.Execute SQL
    
    'AHora comprobamos palets
    
    
    SQL = DevuelveDesdeBD(conAri, "max(idpalet)", "prodpalets", "1", "1")
    If SQL = "" Then SQL = "0"
    idTrazaAntiguo = CStr(Val(SQL) + 1)
    
     '`idpalet`,`LineaPeletiza`,`fhinicio`,`fhFin`,`CajasProd`,`L0`,`L1`,`L2`,`L3`,`L4`,`L5`,`L6`,`L7`
    SQL = idTrazaAntiguo & ",0,NOW(),NOW()," & Me.lwNPalet.ListItems.Count & ",'0','0','0','0','0','0','0','0',0)"   'UNO en manual=linea 8
    SQL = "insert into `prodpalets` (`idpalet`,`LineaPeletiza`,`fhinicio`,`fhFin`,`CajasProd`,`L0`,`L1`,`L2`,`L3`,`L4`,`L5`,`L6`,`L7`,`L8`) values (" & SQL
    
    ValoresLeidos = ""
    If EjecutaSQL(conAri, SQL, True) Then
        Espera 0.4
        'Ya tengo el palet
        'Ahora meto las cajas en prodcajas con el id palet
        SQL = "INSERT INTO prodcajas(lotetraza,idcaja,idpalet,fcreacion) "
        SQL = SQL & "Select lotetraza,idcaja," & idTrazaAntiguo & ",fechadev FROM sreparto_dev WHERE IdDevol = " & IdDev
        If Not EjecutaSQL(conAri, SQL, True) Then ValoresLeidos = "Llame soporte técnico"
    Else
        ValoresLeidos = "Llame soporte técnico"
    End If
    
    
    If ValoresLeidos <> "" Then
        MsgBox ValoresLeidos, vbCritical
    Else
        ValoresLeidos = String(30, "*") & vbCrLf
        ValoresLeidos = ValoresLeidos & ValoresLeidos
        SQL = "Devolucion generada en palet: " & idTrazaAntiguo
        SQL = SQL & vbCrLf & vbCrLf & "Etiqueta palet:   " & "1" & Format(idTrazaAntiguo, "00000000") & "1" & vbCrLf & vbCrLf
        SQL = ValoresLeidos & SQL & ValoresLeidos
        SQL = SQL & "Copiado en portapapeles"
        MsgBox SQL, vbExclamation
        SQL = "1" & Format(idTrazaAntiguo, "00000000") & "1"
        CopiarPortaPapeles SQL
        
        
        Me.ListView1(2).ListItems.Clear
        
    End If
    ValoresLeidos = ""
    idTrazaAntiguo = 0
End Sub

Private Sub CopiarPortaPapeles(Cadena As String)

    On Error Resume Next
    
    Clipboard.Clear
    Clipboard.SetText Cadena
    
    If Err.Number <> 0 Then Err.Clear
End Sub





'------------------------------------------------------------------
'
'
'   Perdida o bja de mercancia
Private Sub text17_GotFocus()
    ConseguirFoco Text17, 3
End Sub

Private Sub text17_KeyPress(KeyAscii As Integer)
     KEYpressGnral KeyAscii, 2, False
End Sub



Private Sub text17_LostFocus()
Dim cadErr As String

 
    Text17.Text = Trim(Text17.Text)
    If Text17.Text = "" Then
        Text16.Text = ""
        Exit Sub
    End If
    
    
    cadErr = ""
    If Not IsNumeric(Text17.Text) Then
        cadErr = "Campo numerico"
        
    Else
    
        'Si la longitud no es 11,12 o 13
        'L =Len(text17.Text)
        If Len(Text17.Text) <> 13 Then
            cadErr = "Longitud incorrecta"
        Else
            '
            cadErr = "lotetraza = " & Mid(Text17, 1, 8) & " AND idcaja"
            
            
            SQL = DevuelveDesdeBD(conAri, "idpalet", "prodcajas", cadErr, Mid(Text17, 9))
            If SQL = "" Then
                cadErr = "No existe caja en sistema"
            Else
                
                'Metemos la caja
                cadErr = MeteCajaEnBajasPerdidas
            End If
        End If
    End If
    
    If cadErr <> "" Then
        Text16.Text = "Err " & cadErr
        
    Else
        
'        If L = 13 Then
'            SQL = "Caja OK"
'        Else
'            If L = 12 Then
'                SQL = "  Albaran OK"
'            ElseIf L = 11 Then
'                SQL = "  Orden OK"
'            Else
'                SQL = " Palet OK"
'            End If
'        End If
        Text16.Text = "# Caja OK"
    End If
        '
    Text17.Text = ""
    PonerFoco Text17
    
    idTrazaAntiguo = 0  'esta variable la es gnral y la utilizo en la funcion
End Sub


Private Function MeteCajaEnBajasPerdidas() As String
Dim IT
    On Error Resume Next
    Set IT = Me.ListView1(3).ListItems.Add(, "C" & Text17.Text, Text17.Text)
    If Err.Number <> 0 Then
        
        MeteCajaEnBajasPerdidas = Err.Description
        Err.Clear
    Else
        IT.SubItems(1) = Format(SQL, "0000") 'PALET
        MeteCajaEnBajasPerdidas = ""
    End If
End Function


Private Sub ProcesoDeBaja()
Dim OK As Boolean
Dim C As Collection
Dim J As Integer
    'Las meteremos en la tabla srepartolotcaj
    'para que luego sepamos que fue baja
    
    'El idreparto sera -1 para que no este vinculado a ninguna orden de carga
    'En nuimalbaran tendremos un secuencial
    
    SQL = DevuelveDesdeBD(conAri, "max(numalbar)", "srepartolotcaj", "idreparto", "-1")
    If SQL = "" Then SQL = "0"
    idTrazaAntiguo = Val(SQL) + 1
    
    ValoresLeidos = DBSet(Now, "FH")
    '(idreparto,codtipom,numalbar,numlinea,codartic,fecha,lotetraza,idcaja,idpalet)
    SQL = ""
    For J = 1 To Me.ListView1(3).ListItems.Count
        SQL = SQL & ", (-1,1," & idTrazaAntiguo & "," & J & ",'PERD_BAJA'," & ValoresLeidos & ","
        SQL = SQL & Mid(ListView1(3).ListItems(J), 1, 8) & "," & Mid(ListView1(3).ListItems(J), 9) & "," & ListView1(3).ListItems(J).SubItems(1) & ")"
        
    Next
    
    SQL = Mid(SQL, 2) 'quitamos primera coma
    ValoresLeidos = "INSERT INTO srepartolotcaj(idreparto,codtipom,numalbar,numlinea,codartic,fecha,lotetraza,idcaja,idpalet) VALUES " & SQL
    
    
    'Hacemos el SQL dentro de commit
    '
    Conn.BeginTrans
    OK = True
        SQL = ""
        For J = 1 To Me.ListView1(3).ListItems.Count
            SQL = SQL & ", (" & Mid(ListView1(3).ListItems(J), 1, 8) & "," & Mid(ListView1(3).ListItems(J), 9) & ")"
        Next
        SQL = Mid(SQL, 2)
        SQL = "DELETE FROM prodcajas WHERE (lotetraza,idcaja) IN (" & SQL & ")"
            
        OK = EjecutaSQL(conAri, SQL, True)
        
        If OK Then
            OK = EjecutaSQL(conAri, ValoresLeidos, True)
            If OK Then
                Espera 0.3
                Set RS = New ADODB.Recordset
                ValoresLeidos = ""
                SQL = "SELECT idpalet,count(*) from srepartolotcaj WHERE idreparto=-1 and numalbar=" & idTrazaAntiguo & "  group by  1"
                RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                idTrazaAntiguo = 0
                While Not RS.EOF
                    idTrazaAntiguo = idTrazaAntiguo + RS.Fields(1)
                    ValoresLeidos = ValoresLeidos & ", " & RS.Fields(0)
                    SQL = "UPDATE prodpalets SET CajasProd=cajasprod - " & RS.Fields(1) & " WHERE idpalet =" & RS!IdPalet
                    Conn.Execute SQL
                    RS.MoveNext
                Wend
                RS.Close
                J = -1
                If idTrazaAntiguo <> ListView1(3).ListItems.Count Then
                    SQL = "Error cajas baja:" & vbCrLf & "Por pistola: " & Me.ListView1(3).ListItems.Count & vbCrLf
                    SQL = SQL & "En BD: " & idTrazaAntiguo & vbCrLf & "Consulte soporte tecnico"
                    MsgBox SQL, vbInformation
                    
                Else
                        
                    SQL = "Desea imprimir las etiqueta(s) de palet "
                    SQL = SQL & "por la impresora SATO?"
                    
                    
                    
                    SQL = MsgBox(SQL, vbQuestion + vbYesNoCancel)
                    If CByte(SQL) = vbCancel Then
                        J = 0
                    ElseIf CByte(SQL) = vbYes Then
                        J = 1  'por la sato
                    Else
                        J = 0  'impresion normal NO HAY impresion normal todavia
                    End If
                End If
               
            End If
    End If
    
    If OK Then
        Conn.CommitTrans
        
        If J > 0 Then
            'Imprime etiquetas
            ValoresLeidos = Mid(ValoresLeidos, 2)
            SQL = "Select * from prodpalets where idpalet in (" & ValoresLeidos & ")"
            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            Conn.Execute "DELETE FROM tmppartidas WHERE codusu = " & vUsu.Codigo
            
            Set cPal = New CPalet
            idTrazaAntiguo = IndexSublinea 'copio lo que hay
            While Not RS.EOF
               
                If cPal.Leer(Val(RS!IdPalet)) Then
                    cPal.CargaDatosPalet C, True, IndexSublinea, True
            
                    If J = 1 Then
                        ImprimirPalet cPal.ID, cPal.TipoImpresion
                        Espera 0.85
                    End If
                    
                End If
                RS.MoveNext
            Wend
            RS.Close
            IndexSublinea = CInt(idTrazaAntiguo)  'dejo lo que habia
            If J = 2 Then
                SQL = "{tmppartidas.codusu}=" & vUsu.Codigo
                'LlamaImprimirGral SQL, "", 0, "EtiqPalet.rpt", "Etiquetas palets "
            End If
            
            
            cmdSalir_Click 11
            
        End If
    Else
        Conn.RollbackTrans
    End If
    Set RS = Nothing
End Sub
    
    
    


    
        
            
        

