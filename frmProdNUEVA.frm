VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmProdNUEVA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Producción / Envasado"
   ClientHeight    =   8715
   ClientLeft      =   -840
   ClientTop       =   -105
   ClientWidth     =   14115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   14115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   13800
      Top             =   7680
   End
   Begin VB.Frame FramePB 
      Height          =   735
      Left            =   120
      TabIndex        =   26
      Top             =   7920
      Width           =   13935
      Begin MSComctlLib.ProgressBar pb1 
         Height          =   375
         Left            =   2880
         TabIndex        =   27
         Top             =   240
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Max             =   8
      End
      Begin VB.Label Label3 
         Caption         =   "Leyendo datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   2655
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   13361
      _Version        =   393216
      Tabs            =   7
      Tab             =   5
      TabsPerRow      =   7
      TabHeight       =   520
      TabCaption(0)   =   "Lineas 0-2"
      TabPicture(0)   =   "frmProdNUEVA.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FrameLinea1(2)"
      Tab(0).Control(1)=   "FrameLinea1(1)"
      Tab(0).Control(2)=   "FrameLinea1(0)"
      Tab(0).Control(3)=   "LineaIndicadora(2)"
      Tab(0).Control(4)=   "LineaIndicadora(1)"
      Tab(0).Control(5)=   "LineaIndicadora(0)"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Lineas 3-5"
      TabPicture(1)   =   "frmProdNUEVA.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameLinea1(5)"
      Tab(1).Control(1)=   "FrameLinea1(4)"
      Tab(1).Control(2)=   "FrameLinea1(3)"
      Tab(1).Control(3)=   "LineaIndicadora(5)"
      Tab(1).Control(4)=   "LineaIndicadora(4)"
      Tab(1).Control(5)=   "LineaIndicadora(3)"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Lineas 6-7 "
      TabPicture(2)   =   "frmProdNUEVA.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameLinea1(7)"
      Tab(2).Control(1)=   "FrameLinea1(6)"
      Tab(2).Control(2)=   "LineaIndicadora(7)"
      Tab(2).Control(3)=   "LineaIndicadora(6)"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Manual / Muestras"
      TabPicture(3)   =   "frmProdNUEVA.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "LineaIndicadora(8)"
      Tab(3).Control(1)=   "LineaIndicadora(9)"
      Tab(3).Control(2)=   "LineaIndicadora(10)"
      Tab(3).Control(3)=   "FrameLinea1(8)"
      Tab(3).Control(4)=   "FrameLinea1(9)"
      Tab(3).Control(5)=   "FrameLinea1(10)"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "Paletización"
      TabPicture(4)   =   "frmProdNUEVA.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "FramePalet(5)"
      Tab(4).Control(1)=   "FramePalet(4)"
      Tab(4).Control(2)=   "FramePalet(3)"
      Tab(4).Control(3)=   "FramePalet(2)"
      Tab(4).Control(4)=   "FramePalet(1)"
      Tab(4).Control(5)=   "FramePalet(0)"
      Tab(4).Control(6)=   "LineaIndicadoraP(5)"
      Tab(4).Control(7)=   "LineaIndicadoraP(4)"
      Tab(4).Control(8)=   "LineaIndicadoraP(3)"
      Tab(4).Control(9)=   "LineaIndicadoraP(2)"
      Tab(4).Control(10)=   "LineaIndicadoraP(1)"
      Tab(4).Control(11)=   "LineaIndicadoraP(0)"
      Tab(4).ControlCount=   12
      TabCaption(5)   =   "Planning"
      TabPicture(5)   =   "frmProdNUEVA.frx":008C
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "lblPlann(2)"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Line1(0)"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "lblPlann(0)"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "Line1(1)"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "lblPlann(3)"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "Line1(2)"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).Control(6)=   "lwp2(0)"
      Tab(5).Control(6).Enabled=   0   'False
      Tab(5).Control(7)=   "lwp"
      Tab(5).Control(7).Enabled=   0   'False
      Tab(5).Control(8)=   "cmdBuscarRef(1)"
      Tab(5).Control(8).Enabled=   0   'False
      Tab(5).Control(9)=   "cmdBuscarRef(0)"
      Tab(5).Control(9).Enabled=   0   'False
      Tab(5).ControlCount=   10
      TabCaption(6)   =   "Materia auxiliar"
      TabPicture(6)   =   "frmProdNUEVA.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "lblPlann(4)"
      Tab(6).Control(1)=   "Line1(3)"
      Tab(6).Control(2)=   "lwp2(1)"
      Tab(6).ControlCount=   3
      Begin VB.Frame FramePalet 
         Caption         =   "  Paletizado 6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6735
         Index           =   5
         Left            =   -63480
         TabIndex        =   325
         Top             =   600
         Width           =   2295
         Begin VB.TextBox txtQueProdLin 
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
            Height          =   1335
            Index           =   5
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   346
            Text            =   "frmProdNUEVA.frx":00C4
            Top             =   5280
            Width           =   2055
         End
         Begin VB.TextBox txtCajas 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   5
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   332
            Text            =   "Text1"
            Top             =   1680
            Width           =   855
         End
         Begin VB.ListBox lstCajas 
            Height          =   2595
            Index           =   5
            Left            =   120
            TabIndex        =   331
            Top             =   2400
            Width           =   2055
         End
         Begin VB.CommandButton cmdCerrarPalet 
            Height          =   375
            Index           =   5
            Left            =   1800
            Picture         =   "frmProdNUEVA.frx":00CA
            Style           =   1  'Graphical
            TabIndex        =   330
            ToolTipText     =   "Finalizar produccion"
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtInicioP 
            Height          =   315
            Index           =   5
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   329
            Text            =   "Text1"
            Top             =   1080
            Width           =   2055
         End
         Begin VB.TextBox txtDuracionP 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000013&
            Height          =   315
            Index           =   5
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   328
            Text            =   "Text1"
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox txtIdP 
            Height          =   315
            Index           =   5
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   327
            Text            =   "Text1"
            Top             =   480
            Width           =   975
         End
         Begin VB.CommandButton cmdInciarPalet 
            Height          =   375
            Index           =   5
            Left            =   1800
            Picture         =   "frmProdNUEVA.frx":0ACC
            Style           =   1  'Graphical
            TabIndex        =   326
            ToolTipText     =   "Asignar nueva produccion"
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "Cajas"
            Height          =   255
            Index           =   90
            Left            =   1320
            TabIndex        =   338
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Cajas"
            Height          =   255
            Index           =   89
            Left            =   120
            TabIndex        =   337
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Duracion"
            Height          =   255
            Index           =   88
            Left            =   120
            TabIndex        =   336
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Id Palet"
            Height          =   255
            Index           =   87
            Left            =   120
            TabIndex        =   335
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "F. incio"
            Height          =   255
            Index           =   86
            Left            =   120
            TabIndex        =   334
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Lineas"
            Height          =   255
            Index           =   85
            Left            =   120
            TabIndex        =   333
            Top             =   5040
            Width           =   855
         End
      End
      Begin VB.Frame FramePalet 
         Caption         =   "  Paletizado 5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6735
         Index           =   4
         Left            =   -65760
         TabIndex        =   311
         Top             =   600
         Width           =   2295
         Begin VB.TextBox txtQueProdLin 
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
            Height          =   1335
            Index           =   4
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   345
            Text            =   "frmProdNUEVA.frx":14CE
            Top             =   5280
            Width           =   2055
         End
         Begin VB.TextBox txtCajas 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   4
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   318
            Text            =   "Text1"
            Top             =   1680
            Width           =   855
         End
         Begin VB.ListBox lstCajas 
            Height          =   2595
            Index           =   4
            Left            =   120
            TabIndex        =   317
            Top             =   2400
            Width           =   2055
         End
         Begin VB.CommandButton cmdCerrarPalet 
            Height          =   375
            Index           =   4
            Left            =   1800
            Picture         =   "frmProdNUEVA.frx":14D4
            Style           =   1  'Graphical
            TabIndex        =   316
            ToolTipText     =   "Finalizar produccion"
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtInicioP 
            Height          =   315
            Index           =   4
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   315
            Text            =   "Text1"
            Top             =   1080
            Width           =   2055
         End
         Begin VB.TextBox txtDuracionP 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000013&
            Height          =   315
            Index           =   4
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   314
            Text            =   "Text1"
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox txtIdP 
            Height          =   315
            Index           =   4
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   313
            Text            =   "Text1"
            Top             =   480
            Width           =   975
         End
         Begin VB.CommandButton cmdInciarPalet 
            Height          =   375
            Index           =   4
            Left            =   1800
            Picture         =   "frmProdNUEVA.frx":1ED6
            Style           =   1  'Graphical
            TabIndex        =   312
            ToolTipText     =   "Asignar nueva produccion"
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "Cajas"
            Height          =   255
            Index           =   84
            Left            =   1320
            TabIndex        =   324
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Cajas"
            Height          =   255
            Index           =   83
            Left            =   120
            TabIndex        =   323
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Duracion"
            Height          =   255
            Index           =   82
            Left            =   120
            TabIndex        =   322
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Id Palet"
            Height          =   255
            Index           =   81
            Left            =   120
            TabIndex        =   321
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "F. incio"
            Height          =   255
            Index           =   80
            Left            =   120
            TabIndex        =   320
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Lineas"
            Height          =   255
            Index           =   79
            Left            =   120
            TabIndex        =   319
            Top             =   5040
            Width           =   855
         End
      End
      Begin VB.Frame FramePalet 
         Caption         =   "  Paletizado 4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6735
         Index           =   3
         Left            =   -68040
         TabIndex        =   297
         Top             =   600
         Width           =   2295
         Begin VB.TextBox txtQueProdLin 
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
            Height          =   1335
            Index           =   3
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   344
            Text            =   "frmProdNUEVA.frx":28D8
            Top             =   5280
            Width           =   2055
         End
         Begin VB.TextBox txtCajas 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   3
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   304
            Text            =   "Text1"
            Top             =   1680
            Width           =   855
         End
         Begin VB.ListBox lstCajas 
            Height          =   2595
            Index           =   3
            Left            =   120
            TabIndex        =   303
            Top             =   2400
            Width           =   2055
         End
         Begin VB.CommandButton cmdCerrarPalet 
            Height          =   375
            Index           =   3
            Left            =   1800
            Picture         =   "frmProdNUEVA.frx":28DE
            Style           =   1  'Graphical
            TabIndex        =   302
            ToolTipText     =   "Finalizar produccion"
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtInicioP 
            Height          =   315
            Index           =   3
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   301
            Text            =   "Text1"
            Top             =   1080
            Width           =   2055
         End
         Begin VB.TextBox txtDuracionP 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000013&
            Height          =   315
            Index           =   3
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   300
            Text            =   "Text1"
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox txtIdP 
            Height          =   315
            Index           =   3
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   299
            Text            =   "Text1"
            Top             =   480
            Width           =   975
         End
         Begin VB.CommandButton cmdInciarPalet 
            Height          =   375
            Index           =   3
            Left            =   1800
            Picture         =   "frmProdNUEVA.frx":32E0
            Style           =   1  'Graphical
            TabIndex        =   298
            ToolTipText     =   "Asignar nueva produccion"
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "Cajas"
            Height          =   255
            Index           =   73
            Left            =   1320
            TabIndex        =   310
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Cajas"
            Height          =   255
            Index           =   72
            Left            =   120
            TabIndex        =   309
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Duracion"
            Height          =   255
            Index           =   71
            Left            =   120
            TabIndex        =   308
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Id Palet"
            Height          =   255
            Index           =   70
            Left            =   120
            TabIndex        =   307
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "F. incio"
            Height          =   255
            Index           =   69
            Left            =   120
            TabIndex        =   306
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Lineas"
            Height          =   255
            Index           =   68
            Left            =   120
            TabIndex        =   305
            Top             =   5040
            Width           =   855
         End
      End
      Begin VB.Frame FramePalet 
         Caption         =   "  Paletizado 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6735
         Index           =   2
         Left            =   -70320
         TabIndex        =   283
         Top             =   600
         Width           =   2295
         Begin VB.TextBox txtQueProdLin 
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
            Height          =   1335
            Index           =   2
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   343
            Text            =   "frmProdNUEVA.frx":3CE2
            Top             =   5280
            Width           =   2055
         End
         Begin VB.TextBox txtCajas 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   2
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   290
            Text            =   "Text1"
            Top             =   1680
            Width           =   855
         End
         Begin VB.ListBox lstCajas 
            Height          =   2595
            Index           =   2
            Left            =   120
            TabIndex        =   289
            Top             =   2400
            Width           =   2055
         End
         Begin VB.CommandButton cmdCerrarPalet 
            Height          =   375
            Index           =   2
            Left            =   1800
            Picture         =   "frmProdNUEVA.frx":3CE8
            Style           =   1  'Graphical
            TabIndex        =   288
            ToolTipText     =   "Finalizar produccion"
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtInicioP 
            Height          =   315
            Index           =   2
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   287
            Text            =   "Text1"
            Top             =   1080
            Width           =   2055
         End
         Begin VB.TextBox txtDuracionP 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000013&
            Height          =   315
            Index           =   2
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   286
            Text            =   "Text1"
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox txtIdP 
            Height          =   315
            Index           =   2
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   285
            Text            =   "Text1"
            Top             =   480
            Width           =   975
         End
         Begin VB.CommandButton cmdInciarPalet 
            Height          =   375
            Index           =   2
            Left            =   1800
            Picture         =   "frmProdNUEVA.frx":46EA
            Style           =   1  'Graphical
            TabIndex        =   284
            ToolTipText     =   "Asignar nueva produccion"
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "Cajas"
            Height          =   255
            Index           =   67
            Left            =   1320
            TabIndex        =   296
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Cajas"
            Height          =   255
            Index           =   66
            Left            =   120
            TabIndex        =   295
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Duracion"
            Height          =   255
            Index           =   65
            Left            =   120
            TabIndex        =   294
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Id Palet"
            Height          =   255
            Index           =   64
            Left            =   120
            TabIndex        =   293
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "F. incio"
            Height          =   255
            Index           =   54
            Left            =   120
            TabIndex        =   292
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Lineas"
            Height          =   255
            Index           =   53
            Left            =   120
            TabIndex        =   291
            Top             =   5040
            Width           =   855
         End
      End
      Begin VB.Frame FramePalet 
         Caption         =   "  Paletizado 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6735
         Index           =   1
         Left            =   -72600
         TabIndex        =   269
         Top             =   600
         Width           =   2295
         Begin VB.TextBox txtQueProdLin 
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
            Height          =   1335
            Index           =   1
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   342
            Text            =   "frmProdNUEVA.frx":50EC
            Top             =   5280
            Width           =   2055
         End
         Begin VB.TextBox txtCajas 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   1
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   276
            Text            =   "Text1"
            Top             =   1680
            Width           =   855
         End
         Begin VB.ListBox lstCajas 
            Height          =   2595
            Index           =   1
            Left            =   120
            TabIndex        =   275
            Top             =   2400
            Width           =   2055
         End
         Begin VB.CommandButton cmdCerrarPalet 
            Height          =   375
            Index           =   1
            Left            =   1800
            Picture         =   "frmProdNUEVA.frx":50F2
            Style           =   1  'Graphical
            TabIndex        =   274
            ToolTipText     =   "Finalizar produccion"
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtInicioP 
            Height          =   315
            Index           =   1
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   273
            Text            =   "Text1"
            Top             =   1080
            Width           =   2055
         End
         Begin VB.TextBox txtDuracionP 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000013&
            Height          =   315
            Index           =   1
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   272
            Text            =   "Text1"
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox txtIdP 
            Height          =   315
            Index           =   1
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   271
            Text            =   "Text1"
            Top             =   480
            Width           =   975
         End
         Begin VB.CommandButton cmdInciarPalet 
            Height          =   375
            Index           =   1
            Left            =   1800
            Picture         =   "frmProdNUEVA.frx":5AF4
            Style           =   1  'Graphical
            TabIndex        =   270
            ToolTipText     =   "Asignar nueva produccion"
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "Cajas"
            Height          =   255
            Index           =   51
            Left            =   1320
            TabIndex        =   282
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Cajas"
            Height          =   255
            Index           =   50
            Left            =   120
            TabIndex        =   281
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Duracion"
            Height          =   255
            Index           =   49
            Left            =   120
            TabIndex        =   280
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Id Palet"
            Height          =   255
            Index           =   47
            Left            =   120
            TabIndex        =   279
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "F. incio"
            Height          =   255
            Index           =   46
            Left            =   120
            TabIndex        =   278
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Lineas"
            Height          =   255
            Index           =   45
            Left            =   120
            TabIndex        =   277
            Top             =   5040
            Width           =   855
         End
      End
      Begin VB.Frame FrameLinea1 
         Caption         =   "Linea 10                Producción MUESTRAS"
         ClipControls    =   0   'False
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
         Height          =   2175
         Index           =   10
         Left            =   -74640
         TabIndex        =   247
         Top             =   5160
         Width           =   13335
         Begin VB.TextBox txtLineProdPalet 
            Enabled         =   0   'False
            Height          =   375
            Index           =   10
            Left            =   1080
            TabIndex        =   259
            Text            =   " "
            Top             =   1680
            Width           =   1575
         End
         Begin VB.TextBox txtTraza 
            Height          =   375
            Index           =   10
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   258
            Top             =   1170
            Width           =   1335
         End
         Begin VB.CommandButton cmdVerProd 
            Height          =   375
            Index           =   10
            Left            =   12720
            Picture         =   "frmProdNUEVA.frx":64F6
            Style           =   1  'Graphical
            TabIndex        =   257
            ToolTipText     =   "ver produccion"
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtDuracion 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000013&
            Height          =   375
            Index           =   10
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   256
            Top             =   1200
            Width           =   855
         End
         Begin VB.TextBox txtInicio 
            Height          =   375
            Index           =   10
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   255
            Top             =   1170
            Width           =   1935
         End
         Begin VB.CommandButton cmdAsignarProd 
            Height          =   375
            Index           =   10
            Left            =   120
            Picture         =   "frmProdNUEVA.frx":6EF8
            Style           =   1  'Graphical
            TabIndex        =   254
            ToolTipText     =   "Asignar nueva produccion"
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtCantidad 
            Alignment       =   1  'Right Justify
            Height          =   375
            Index           =   10
            Left            =   4680
            Locked          =   -1  'True
            TabIndex        =   253
            Top             =   1680
            Width           =   975
         End
         Begin VB.CommandButton cmdCerrarProd 
            Height          =   375
            Index           =   10
            Left            =   12720
            Picture         =   "frmProdNUEVA.frx":78FA
            Style           =   1  'Graphical
            TabIndex        =   252
            ToolTipText     =   "Finalizar produccion"
            Top             =   1680
            Width           =   375
         End
         Begin VB.CommandButton cmdCambioLotLin 
            Height          =   375
            Index           =   10
            Left            =   12720
            Picture         =   "frmProdNUEVA.frx":82FC
            Style           =   1  'Graphical
            TabIndex        =   251
            ToolTipText     =   "Cambiar LOTE componente"
            Top             =   1080
            Width           =   375
         End
         Begin VB.TextBox txtNomartic 
            BackColor       =   &H80000018&
            Height          =   375
            Index           =   10
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   250
            Top             =   600
            Width           =   4095
         End
         Begin VB.TextBox txtCodartPPal 
            Height          =   375
            Index           =   10
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   249
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtCajasEstimadas 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   10
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   248
            Top             =   1680
            Width           =   855
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1575
            Index           =   10
            Left            =   6960
            TabIndex        =   260
            Top             =   480
            Width           =   5535
            _ExtentX        =   9763
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Articulo"
               Object.Width           =   5997
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Lote"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label lblLinPalet 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lin. palet."
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   266
            Top             =   1800
            Width           =   690
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Trazabilidad"
            Height          =   195
            Index           =   78
            Left            =   120
            TabIndex        =   265
            Top             =   1260
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Duración"
            Height          =   195
            Index           =   77
            Left            =   5040
            TabIndex        =   264
            Top             =   1260
            Width           =   645
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inicio"
            Height          =   195
            Index           =   76
            Left            =   2520
            TabIndex        =   263
            Top             =   1260
            Width           =   375
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Cant. estimada"
            Height          =   195
            Index           =   75
            Left            =   3360
            TabIndex        =   262
            Top             =   1800
            Width           =   1305
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Articulo"
            Height          =   255
            Index           =   74
            Left            =   720
            TabIndex        =   261
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame FrameLinea1 
         Caption         =   "Linea 9                Producción manual 2"
         ClipControls    =   0   'False
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
         Height          =   2175
         Index           =   9
         Left            =   -74640
         TabIndex        =   224
         Top             =   2880
         Width           =   13335
         Begin VB.TextBox txtCajasEstimadas 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   9
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   236
            Top             =   1680
            Width           =   855
         End
         Begin VB.TextBox txtCodartPPal 
            Height          =   375
            Index           =   9
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   235
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtNomartic 
            BackColor       =   &H80000018&
            Height          =   375
            Index           =   9
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   234
            Top             =   600
            Width           =   4095
         End
         Begin VB.CommandButton cmdCambioLotLin 
            Height          =   375
            Index           =   9
            Left            =   12720
            Picture         =   "frmProdNUEVA.frx":8CFE
            Style           =   1  'Graphical
            TabIndex        =   233
            ToolTipText     =   "Cambiar LOTE componente"
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton cmdCerrarProd 
            Height          =   375
            Index           =   9
            Left            =   12720
            Picture         =   "frmProdNUEVA.frx":9700
            Style           =   1  'Graphical
            TabIndex        =   232
            ToolTipText     =   "Finalizar produccion"
            Top             =   1680
            Width           =   375
         End
         Begin VB.TextBox txtCantidad 
            Alignment       =   1  'Right Justify
            Height          =   375
            Index           =   9
            Left            =   4680
            Locked          =   -1  'True
            TabIndex        =   231
            Top             =   1680
            Width           =   975
         End
         Begin VB.CommandButton cmdAsignarProd 
            Height          =   375
            Index           =   9
            Left            =   120
            Picture         =   "frmProdNUEVA.frx":A102
            Style           =   1  'Graphical
            TabIndex        =   230
            ToolTipText     =   "Asignar nueva produccion"
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtInicio 
            Height          =   375
            Index           =   9
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   229
            Top             =   1170
            Width           =   1935
         End
         Begin VB.TextBox txtDuracion 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000013&
            Height          =   375
            Index           =   9
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   228
            Top             =   1200
            Width           =   855
         End
         Begin VB.CommandButton cmdVerProd 
            Height          =   375
            Index           =   9
            Left            =   12720
            Picture         =   "frmProdNUEVA.frx":AB04
            Style           =   1  'Graphical
            TabIndex        =   227
            ToolTipText     =   "ver produccion"
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtTraza 
            Height          =   375
            Index           =   9
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   226
            Top             =   1170
            Width           =   1335
         End
         Begin VB.TextBox txtLineProdPalet 
            Enabled         =   0   'False
            Height          =   375
            Index           =   9
            Left            =   1080
            TabIndex        =   225
            Text            =   " "
            Top             =   1680
            Width           =   1575
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1575
            Index           =   9
            Left            =   6960
            TabIndex        =   237
            Top             =   480
            Width           =   5535
            _ExtentX        =   9763
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Articulo"
               Object.Width           =   5997
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Lote"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Articulo"
            Height          =   255
            Index           =   63
            Left            =   720
            TabIndex        =   243
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Cant. estimada"
            Height          =   195
            Index           =   62
            Left            =   3360
            TabIndex        =   242
            Top             =   1800
            Width           =   1305
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inicio"
            Height          =   195
            Index           =   61
            Left            =   2520
            TabIndex        =   241
            Top             =   1260
            Width           =   375
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Duración"
            Height          =   195
            Index           =   60
            Left            =   5040
            TabIndex        =   240
            Top             =   1260
            Width           =   645
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Trazabilidad"
            Height          =   195
            Index           =   59
            Left            =   120
            TabIndex        =   239
            Top             =   1260
            Width           =   1095
         End
         Begin VB.Label lblLinPalet 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lin. palet."
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   238
            Top             =   1800
            Width           =   690
         End
      End
      Begin VB.Frame FrameLinea1 
         Caption         =   "Linea 8                Producción manual "
         ClipControls    =   0   'False
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
         Height          =   2175
         Index           =   8
         Left            =   -74640
         TabIndex        =   204
         Top             =   480
         Width           =   13335
         Begin VB.TextBox txtLineProdPalet 
            Enabled         =   0   'False
            Height          =   375
            Index           =   8
            Left            =   1080
            TabIndex        =   216
            Text            =   " "
            Top             =   1680
            Width           =   1575
         End
         Begin VB.TextBox txtTraza 
            Height          =   375
            Index           =   8
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   215
            Top             =   1170
            Width           =   1335
         End
         Begin VB.CommandButton cmdVerProd 
            Height          =   375
            Index           =   8
            Left            =   12720
            Picture         =   "frmProdNUEVA.frx":B506
            Style           =   1  'Graphical
            TabIndex        =   214
            ToolTipText     =   "ver produccion"
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtDuracion 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000013&
            Height          =   375
            Index           =   8
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   213
            Top             =   1200
            Width           =   855
         End
         Begin VB.TextBox txtInicio 
            Height          =   375
            Index           =   8
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   212
            Top             =   1170
            Width           =   1935
         End
         Begin VB.CommandButton cmdAsignarProd 
            Height          =   375
            Index           =   8
            Left            =   120
            Picture         =   "frmProdNUEVA.frx":BF08
            Style           =   1  'Graphical
            TabIndex        =   211
            ToolTipText     =   "Asignar nueva produccion"
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtCantidad 
            Alignment       =   1  'Right Justify
            Height          =   375
            Index           =   8
            Left            =   4680
            Locked          =   -1  'True
            TabIndex        =   210
            Top             =   1680
            Width           =   975
         End
         Begin VB.CommandButton cmdCerrarProd 
            Height          =   375
            Index           =   8
            Left            =   12720
            Picture         =   "frmProdNUEVA.frx":C90A
            Style           =   1  'Graphical
            TabIndex        =   209
            ToolTipText     =   "Finalizar produccion"
            Top             =   1680
            Width           =   375
         End
         Begin VB.CommandButton cmdCambioLotLin 
            Height          =   375
            Index           =   8
            Left            =   12720
            Picture         =   "frmProdNUEVA.frx":D30C
            Style           =   1  'Graphical
            TabIndex        =   208
            ToolTipText     =   "Cambiar LOTE componente"
            Top             =   1080
            Width           =   375
         End
         Begin VB.TextBox txtNomartic 
            BackColor       =   &H80000018&
            Height          =   375
            Index           =   8
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   207
            Top             =   600
            Width           =   4095
         End
         Begin VB.TextBox txtCodartPPal 
            Height          =   375
            Index           =   8
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   206
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtCajasEstimadas 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   8
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   205
            Top             =   1680
            Width           =   855
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1575
            Index           =   8
            Left            =   6960
            TabIndex        =   217
            Top             =   480
            Width           =   5535
            _ExtentX        =   9763
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Articulo"
               Object.Width           =   5997
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Lote"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label lblLinPalet 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lin. palet."
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   223
            Top             =   1800
            Width           =   690
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Trazabilidad"
            Height          =   195
            Index           =   58
            Left            =   120
            TabIndex        =   222
            Top             =   1260
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Duración"
            Height          =   195
            Index           =   57
            Left            =   5040
            TabIndex        =   221
            Top             =   1260
            Width           =   645
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inicio"
            Height          =   195
            Index           =   56
            Left            =   2520
            TabIndex        =   220
            Top             =   1260
            Width           =   375
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Cant. estimada"
            Height          =   195
            Index           =   55
            Left            =   3360
            TabIndex        =   219
            Top             =   1800
            Width           =   1305
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Articulo"
            Height          =   255
            Index           =   48
            Left            =   720
            TabIndex        =   218
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame FramePalet 
         Caption         =   "  Paletizado 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6735
         Index           =   0
         Left            =   -74880
         TabIndex        =   191
         Top             =   600
         Width           =   2295
         Begin VB.TextBox txtQueProdLin 
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
            Height          =   1335
            Index           =   0
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   268
            Text            =   "frmProdNUEVA.frx":DD0E
            Top             =   5280
            Width           =   2055
         End
         Begin VB.CommandButton cmdInciarPalet 
            Height          =   375
            Index           =   0
            Left            =   1800
            Picture         =   "frmProdNUEVA.frx":DD14
            Style           =   1  'Graphical
            TabIndex        =   198
            ToolTipText     =   "Asignar nueva produccion"
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtIdP 
            Height          =   315
            Index           =   0
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   197
            Text            =   "Text1"
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txtDuracionP 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000013&
            Height          =   315
            Index           =   0
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   196
            Text            =   "Text1"
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox txtInicioP 
            Height          =   315
            Index           =   0
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   195
            Text            =   "Text1"
            Top             =   1080
            Width           =   2055
         End
         Begin VB.CommandButton cmdCerrarPalet 
            Height          =   375
            Index           =   0
            Left            =   1800
            Picture         =   "frmProdNUEVA.frx":E716
            Style           =   1  'Graphical
            TabIndex        =   194
            ToolTipText     =   "Finalizar produccion"
            Top             =   480
            Width           =   375
         End
         Begin VB.ListBox lstCajas 
            Height          =   2595
            Index           =   0
            Left            =   120
            TabIndex        =   193
            Top             =   2400
            Width           =   2055
         End
         Begin VB.TextBox txtCajas 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   0
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   192
            Text            =   "Text1"
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Lineas"
            Height          =   255
            Index           =   44
            Left            =   120
            TabIndex        =   245
            Top             =   5040
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "F. incio"
            Height          =   255
            Index           =   40
            Left            =   120
            TabIndex        =   203
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Id Palet"
            Height          =   255
            Index           =   41
            Left            =   120
            TabIndex        =   202
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Duracion"
            Height          =   255
            Index           =   42
            Left            =   120
            TabIndex        =   201
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Cajas"
            Height          =   255
            Index           =   43
            Left            =   120
            TabIndex        =   200
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Cajas"
            Height          =   255
            Index           =   52
            Left            =   1320
            TabIndex        =   199
            Top             =   1440
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdBuscarRef 
         Height          =   375
         Index           =   0
         Left            =   1920
         Picture         =   "frmProdNUEVA.frx":F118
         Style           =   1  'Graphical
         TabIndex        =   188
         ToolTipText     =   "Buscar referencia anterior"
         Top             =   5160
         Width           =   375
      End
      Begin VB.CommandButton cmdBuscarRef 
         Height          =   375
         Index           =   1
         Left            =   1920
         Picture         =   "frmProdNUEVA.frx":F6A2
         Style           =   1  'Graphical
         TabIndex        =   187
         ToolTipText     =   "Buscar referencai posterior"
         Top             =   5640
         Width           =   375
      End
      Begin VB.Frame FrameLinea1 
         Caption         =   "Linea 7"
         ClipControls    =   0   'False
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
         Height          =   2175
         Index           =   7
         Left            =   -74640
         TabIndex        =   152
         Top             =   4200
         Width           =   13335
         Begin VB.TextBox txtCajasEstimadas 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   7
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   178
            Top             =   1680
            Width           =   855
         End
         Begin VB.TextBox txtCodartPPal 
            Height          =   375
            Index           =   7
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   164
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtNomartic 
            BackColor       =   &H80000018&
            Height          =   375
            Index           =   7
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   163
            Top             =   600
            Width           =   4095
         End
         Begin VB.CommandButton cmdCambioLotLin 
            Height          =   375
            Index           =   7
            Left            =   12720
            Picture         =   "frmProdNUEVA.frx":FC2C
            Style           =   1  'Graphical
            TabIndex        =   161
            ToolTipText     =   "Cambiar LOTE componente"
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton cmdCerrarProd 
            Height          =   375
            Index           =   7
            Left            =   12720
            Picture         =   "frmProdNUEVA.frx":1062E
            Style           =   1  'Graphical
            TabIndex        =   160
            ToolTipText     =   "Finalizar produccion"
            Top             =   1680
            Width           =   375
         End
         Begin VB.TextBox txtCantidad 
            Alignment       =   1  'Right Justify
            Height          =   375
            Index           =   7
            Left            =   4680
            Locked          =   -1  'True
            TabIndex        =   159
            Top             =   1680
            Width           =   975
         End
         Begin VB.CommandButton cmdAsignarProd 
            Height          =   375
            Index           =   7
            Left            =   120
            Picture         =   "frmProdNUEVA.frx":11030
            Style           =   1  'Graphical
            TabIndex        =   158
            ToolTipText     =   "Asignar nueva produccion"
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtInicio 
            Height          =   375
            Index           =   7
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   157
            Top             =   1170
            Width           =   1935
         End
         Begin VB.TextBox txtDuracion 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000013&
            Height          =   375
            Index           =   7
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   156
            Top             =   1200
            Width           =   855
         End
         Begin VB.CommandButton cmdVerProd 
            Height          =   375
            Index           =   7
            Left            =   12720
            Picture         =   "frmProdNUEVA.frx":11A32
            Style           =   1  'Graphical
            TabIndex        =   155
            ToolTipText     =   "ver produccion"
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtTraza 
            Height          =   375
            Index           =   7
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   154
            Top             =   1170
            Width           =   1335
         End
         Begin VB.TextBox txtLineProdPalet 
            Enabled         =   0   'False
            Height          =   375
            Index           =   7
            Left            =   1080
            TabIndex        =   153
            Text            =   " "
            Top             =   1680
            Width           =   1575
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1575
            Index           =   7
            Left            =   6960
            TabIndex        =   162
            Top             =   480
            Width           =   5535
            _ExtentX        =   9763
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Articulo"
               Object.Width           =   5997
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Lote"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Articulo"
            Height          =   255
            Index           =   39
            Left            =   720
            TabIndex        =   170
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Cant. estimada"
            Height          =   195
            Index           =   38
            Left            =   3360
            TabIndex        =   169
            Top             =   1800
            Width           =   1305
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inicio"
            Height          =   195
            Index           =   37
            Left            =   2520
            TabIndex        =   168
            Top             =   1260
            Width           =   375
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Duración"
            Height          =   195
            Index           =   36
            Left            =   5040
            TabIndex        =   167
            Top             =   1260
            Width           =   645
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Trazabilidad"
            Height          =   195
            Index           =   35
            Left            =   120
            TabIndex        =   166
            Top             =   1260
            Width           =   1095
         End
         Begin VB.Label lblLinPalet 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lin. palet."
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   165
            Top             =   1800
            Width           =   690
         End
      End
      Begin VB.Frame FrameLinea1 
         Caption         =   "Linea 6"
         ClipControls    =   0   'False
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
         Height          =   2175
         Index           =   6
         Left            =   -74640
         TabIndex        =   133
         Top             =   960
         Width           =   13335
         Begin VB.TextBox txtCajasEstimadas 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   6
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   177
            Top             =   1680
            Width           =   855
         End
         Begin VB.TextBox txtCodartPPal 
            Height          =   375
            Index           =   6
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   145
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtNomartic 
            BackColor       =   &H80000018&
            Height          =   375
            Index           =   6
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   144
            Top             =   600
            Width           =   4095
         End
         Begin VB.CommandButton cmdCambioLotLin 
            Height          =   375
            Index           =   6
            Left            =   12720
            Picture         =   "frmProdNUEVA.frx":12434
            Style           =   1  'Graphical
            TabIndex        =   142
            ToolTipText     =   "Cambiar LOTE componente"
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton cmdCerrarProd 
            Height          =   375
            Index           =   6
            Left            =   12720
            Picture         =   "frmProdNUEVA.frx":12E36
            Style           =   1  'Graphical
            TabIndex        =   141
            ToolTipText     =   "Finalizar produccion"
            Top             =   1680
            Width           =   375
         End
         Begin VB.TextBox txtCantidad 
            Alignment       =   1  'Right Justify
            Height          =   375
            Index           =   6
            Left            =   4680
            Locked          =   -1  'True
            TabIndex        =   140
            Top             =   1680
            Width           =   975
         End
         Begin VB.CommandButton cmdAsignarProd 
            Height          =   375
            Index           =   6
            Left            =   120
            Picture         =   "frmProdNUEVA.frx":13838
            Style           =   1  'Graphical
            TabIndex        =   139
            ToolTipText     =   "Asignar nueva produccion"
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtInicio 
            Height          =   375
            Index           =   6
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   138
            Top             =   1170
            Width           =   1935
         End
         Begin VB.TextBox txtDuracion 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000013&
            Height          =   375
            Index           =   6
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   137
            Top             =   1200
            Width           =   855
         End
         Begin VB.CommandButton cmdVerProd 
            Height          =   375
            Index           =   6
            Left            =   12720
            Picture         =   "frmProdNUEVA.frx":1423A
            Style           =   1  'Graphical
            TabIndex        =   136
            ToolTipText     =   "ver produccion"
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtTraza 
            Height          =   375
            Index           =   6
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   135
            Top             =   1170
            Width           =   1335
         End
         Begin VB.TextBox txtLineProdPalet 
            Enabled         =   0   'False
            Height          =   375
            Index           =   6
            Left            =   1080
            TabIndex        =   134
            Text            =   " "
            Top             =   1680
            Width           =   1575
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1575
            Index           =   6
            Left            =   6960
            TabIndex        =   143
            Top             =   480
            Width           =   5535
            _ExtentX        =   9763
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Articulo"
               Object.Width           =   5997
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Lote"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Articulo"
            Height          =   255
            Index           =   34
            Left            =   720
            TabIndex        =   151
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Cant. estimada"
            Height          =   195
            Index           =   33
            Left            =   3360
            TabIndex        =   150
            Top             =   1800
            Width           =   1305
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inicio"
            Height          =   195
            Index           =   32
            Left            =   2520
            TabIndex        =   149
            Top             =   1260
            Width           =   375
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Duración"
            Height          =   195
            Index           =   31
            Left            =   5040
            TabIndex        =   148
            Top             =   1260
            Width           =   645
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Trazabilidad"
            Height          =   195
            Index           =   30
            Left            =   120
            TabIndex        =   147
            Top             =   1260
            Width           =   1095
         End
         Begin VB.Label lblLinPalet 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lin. palet."
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   146
            Top             =   1800
            Width           =   690
         End
      End
      Begin VB.Frame FrameLinea1 
         Caption         =   "Linea 5"
         ClipControls    =   0   'False
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
         Height          =   2175
         Index           =   5
         Left            =   -74640
         TabIndex        =   114
         Top             =   5160
         Width           =   13335
         Begin VB.TextBox txtCajasEstimadas 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   5
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   176
            Top             =   1680
            Width           =   855
         End
         Begin VB.TextBox txtCodartPPal 
            Height          =   375
            Index           =   5
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   126
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtNomartic 
            BackColor       =   &H80000018&
            Height          =   375
            Index           =   5
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   125
            Top             =   600
            Width           =   4095
         End
         Begin VB.CommandButton cmdCambioLotLin 
            Height          =   375
            Index           =   5
            Left            =   12720
            Picture         =   "frmProdNUEVA.frx":14C3C
            Style           =   1  'Graphical
            TabIndex        =   123
            ToolTipText     =   "Cambiar LOTE componente"
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton cmdCerrarProd 
            Height          =   375
            Index           =   5
            Left            =   12720
            Picture         =   "frmProdNUEVA.frx":1563E
            Style           =   1  'Graphical
            TabIndex        =   122
            ToolTipText     =   "Finalizar produccion"
            Top             =   1680
            Width           =   375
         End
         Begin VB.TextBox txtCantidad 
            Alignment       =   1  'Right Justify
            Height          =   375
            Index           =   5
            Left            =   4680
            Locked          =   -1  'True
            TabIndex        =   121
            Top             =   1680
            Width           =   975
         End
         Begin VB.CommandButton cmdAsignarProd 
            Height          =   375
            Index           =   5
            Left            =   120
            Picture         =   "frmProdNUEVA.frx":16040
            Style           =   1  'Graphical
            TabIndex        =   120
            ToolTipText     =   "Asignar nueva produccion"
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtInicio 
            Height          =   375
            Index           =   5
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   119
            Top             =   1170
            Width           =   1935
         End
         Begin VB.TextBox txtDuracion 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000013&
            Height          =   375
            Index           =   5
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   118
            Top             =   1200
            Width           =   855
         End
         Begin VB.CommandButton cmdVerProd 
            Height          =   375
            Index           =   5
            Left            =   12720
            Picture         =   "frmProdNUEVA.frx":16A42
            Style           =   1  'Graphical
            TabIndex        =   117
            ToolTipText     =   "ver produccion"
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtTraza 
            Height          =   375
            Index           =   5
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   116
            Top             =   1170
            Width           =   1335
         End
         Begin VB.TextBox txtLineProdPalet 
            Enabled         =   0   'False
            Height          =   375
            Index           =   5
            Left            =   1080
            TabIndex        =   115
            Text            =   " "
            Top             =   1680
            Width           =   1575
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1575
            Index           =   5
            Left            =   6960
            TabIndex        =   124
            Top             =   480
            Width           =   5535
            _ExtentX        =   9763
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Articulo"
               Object.Width           =   5997
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Lote"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Articulo"
            Height          =   255
            Index           =   29
            Left            =   720
            TabIndex        =   132
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Cant. estimada"
            Height          =   195
            Index           =   28
            Left            =   3240
            TabIndex        =   131
            Top             =   1800
            Width           =   1305
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inicio"
            Height          =   195
            Index           =   27
            Left            =   2520
            TabIndex        =   130
            Top             =   1260
            Width           =   375
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Duración"
            Height          =   195
            Index           =   26
            Left            =   5040
            TabIndex        =   129
            Top             =   1260
            Width           =   645
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Trazabilidad"
            Height          =   195
            Index           =   25
            Left            =   120
            TabIndex        =   128
            Top             =   1260
            Width           =   1095
         End
         Begin VB.Label lblLinPalet 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lin. palet."
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   127
            Top             =   1800
            Width           =   690
         End
      End
      Begin VB.Frame FrameLinea1 
         Caption         =   "Linea 4"
         ClipControls    =   0   'False
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
         Height          =   2175
         Index           =   4
         Left            =   -74640
         TabIndex        =   95
         Top             =   2880
         Width           =   13335
         Begin VB.TextBox txtCajasEstimadas 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   4
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   175
            Top             =   1680
            Width           =   855
         End
         Begin VB.TextBox txtCodartPPal 
            Height          =   375
            Index           =   4
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   107
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtNomartic 
            BackColor       =   &H80000018&
            Height          =   375
            Index           =   4
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   106
            Top             =   600
            Width           =   4095
         End
         Begin VB.CommandButton cmdCambioLotLin 
            Height          =   375
            Index           =   4
            Left            =   12720
            Picture         =   "frmProdNUEVA.frx":17444
            Style           =   1  'Graphical
            TabIndex        =   104
            ToolTipText     =   "Cambiar LOTE componente"
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton cmdCerrarProd 
            Height          =   375
            Index           =   4
            Left            =   12720
            Picture         =   "frmProdNUEVA.frx":17E46
            Style           =   1  'Graphical
            TabIndex        =   103
            ToolTipText     =   "Finalizar produccion"
            Top             =   1680
            Width           =   375
         End
         Begin VB.TextBox txtCantidad 
            Alignment       =   1  'Right Justify
            Height          =   375
            Index           =   4
            Left            =   4680
            Locked          =   -1  'True
            TabIndex        =   102
            Top             =   1680
            Width           =   975
         End
         Begin VB.CommandButton cmdAsignarProd 
            Height          =   375
            Index           =   4
            Left            =   120
            Picture         =   "frmProdNUEVA.frx":18848
            Style           =   1  'Graphical
            TabIndex        =   101
            ToolTipText     =   "Asignar nueva produccion"
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtInicio 
            Height          =   375
            Index           =   4
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   100
            Top             =   1170
            Width           =   1935
         End
         Begin VB.TextBox txtDuracion 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000013&
            Height          =   375
            Index           =   4
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   99
            Top             =   1200
            Width           =   855
         End
         Begin VB.CommandButton cmdVerProd 
            Height          =   375
            Index           =   4
            Left            =   12720
            Picture         =   "frmProdNUEVA.frx":1924A
            Style           =   1  'Graphical
            TabIndex        =   98
            ToolTipText     =   "ver produccion"
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtTraza 
            Height          =   375
            Index           =   4
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   97
            Top             =   1170
            Width           =   1335
         End
         Begin VB.TextBox txtLineProdPalet 
            Enabled         =   0   'False
            Height          =   375
            Index           =   4
            Left            =   1080
            TabIndex        =   96
            Text            =   " "
            Top             =   1680
            Width           =   1575
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1575
            Index           =   4
            Left            =   6960
            TabIndex        =   105
            Top             =   480
            Width           =   5535
            _ExtentX        =   9763
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Articulo"
               Object.Width           =   5997
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Lote"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Articulo"
            Height          =   255
            Index           =   24
            Left            =   720
            TabIndex        =   113
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Cant. estimada"
            Height          =   195
            Index           =   23
            Left            =   3240
            TabIndex        =   112
            Top             =   1800
            Width           =   1305
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inicio"
            Height          =   195
            Index           =   22
            Left            =   2520
            TabIndex        =   111
            Top             =   1260
            Width           =   375
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Duración"
            Height          =   195
            Index           =   21
            Left            =   5040
            TabIndex        =   110
            Top             =   1260
            Width           =   645
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Trazabilidad"
            Height          =   195
            Index           =   20
            Left            =   120
            TabIndex        =   109
            Top             =   1260
            Width           =   1095
         End
         Begin VB.Label lblLinPalet 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lin. palet."
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   108
            Top             =   1800
            Width           =   690
         End
      End
      Begin VB.Frame FrameLinea1 
         Caption         =   "Linea 3"
         ClipControls    =   0   'False
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
         Height          =   2175
         Index           =   3
         Left            =   -74640
         TabIndex        =   76
         Top             =   480
         Width           =   13335
         Begin VB.TextBox txtCajasEstimadas 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   3
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   174
            Top             =   1680
            Width           =   855
         End
         Begin VB.TextBox txtCodartPPal 
            Height          =   375
            Index           =   3
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   88
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtNomartic 
            BackColor       =   &H80000018&
            Height          =   375
            Index           =   3
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   87
            Top             =   600
            Width           =   4095
         End
         Begin VB.CommandButton cmdCambioLotLin 
            Height          =   375
            Index           =   3
            Left            =   12720
            Picture         =   "frmProdNUEVA.frx":19C4C
            Style           =   1  'Graphical
            TabIndex        =   85
            ToolTipText     =   "Cambiar LOTE componente"
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton cmdCerrarProd 
            Height          =   375
            Index           =   3
            Left            =   12720
            Picture         =   "frmProdNUEVA.frx":1A64E
            Style           =   1  'Graphical
            TabIndex        =   84
            ToolTipText     =   "Finalizar produccion"
            Top             =   1680
            Width           =   375
         End
         Begin VB.TextBox txtCantidad 
            Alignment       =   1  'Right Justify
            Height          =   375
            Index           =   3
            Left            =   4680
            Locked          =   -1  'True
            TabIndex        =   83
            Top             =   1680
            Width           =   975
         End
         Begin VB.CommandButton cmdAsignarProd 
            Height          =   375
            Index           =   3
            Left            =   120
            Picture         =   "frmProdNUEVA.frx":1B050
            Style           =   1  'Graphical
            TabIndex        =   82
            ToolTipText     =   "Asignar nueva produccion"
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtInicio 
            Height          =   375
            Index           =   3
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   81
            Top             =   1170
            Width           =   1935
         End
         Begin VB.TextBox txtDuracion 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000013&
            Height          =   375
            Index           =   3
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   80
            Top             =   1200
            Width           =   855
         End
         Begin VB.CommandButton cmdVerProd 
            Height          =   375
            Index           =   3
            Left            =   12720
            Picture         =   "frmProdNUEVA.frx":1BA52
            Style           =   1  'Graphical
            TabIndex        =   79
            ToolTipText     =   "ver produccion"
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtTraza 
            Height          =   375
            Index           =   3
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   78
            Top             =   1170
            Width           =   1335
         End
         Begin VB.TextBox txtLineProdPalet 
            Enabled         =   0   'False
            Height          =   375
            Index           =   3
            Left            =   1080
            TabIndex        =   77
            Text            =   " "
            Top             =   1680
            Width           =   1575
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1575
            Index           =   3
            Left            =   6960
            TabIndex        =   86
            Top             =   480
            Width           =   5535
            _ExtentX        =   9763
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Articulo"
               Object.Width           =   5997
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Lote"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Articulo"
            Height          =   255
            Index           =   19
            Left            =   720
            TabIndex        =   94
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Cant. estimada"
            Height          =   195
            Index           =   18
            Left            =   3240
            TabIndex        =   93
            Top             =   1800
            Width           =   1305
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inicio"
            Height          =   195
            Index           =   17
            Left            =   2520
            TabIndex        =   92
            Top             =   1260
            Width           =   375
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Duración"
            Height          =   195
            Index           =   16
            Left            =   5040
            TabIndex        =   91
            Top             =   1260
            Width           =   645
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Trazabilidad"
            Height          =   195
            Index           =   15
            Left            =   120
            TabIndex        =   90
            Top             =   1260
            Width           =   1095
         End
         Begin VB.Label lblLinPalet 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lin. palet."
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   89
            Top             =   1800
            Width           =   690
         End
      End
      Begin VB.Frame FrameLinea1 
         Caption         =   "Linea 2"
         ClipControls    =   0   'False
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
         Height          =   2175
         Index           =   2
         Left            =   -74640
         TabIndex        =   57
         Top             =   5160
         Width           =   13335
         Begin VB.TextBox txtCajasEstimadas 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   2
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   173
            Top             =   1680
            Width           =   855
         End
         Begin VB.TextBox txtCodartPPal 
            Height          =   375
            Index           =   2
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   69
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtNomartic 
            BackColor       =   &H80000018&
            Height          =   375
            Index           =   2
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   68
            Top             =   600
            Width           =   4095
         End
         Begin VB.CommandButton cmdCambioLotLin 
            Height          =   375
            Index           =   2
            Left            =   12720
            Picture         =   "frmProdNUEVA.frx":1C454
            Style           =   1  'Graphical
            TabIndex        =   66
            ToolTipText     =   "Cambiar LOTE componente"
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton cmdCerrarProd 
            Height          =   375
            Index           =   2
            Left            =   12720
            Picture         =   "frmProdNUEVA.frx":1CE56
            Style           =   1  'Graphical
            TabIndex        =   65
            ToolTipText     =   "Finalizar produccion"
            Top             =   1680
            Width           =   375
         End
         Begin VB.TextBox txtCantidad 
            Alignment       =   1  'Right Justify
            Height          =   375
            Index           =   2
            Left            =   4680
            Locked          =   -1  'True
            TabIndex        =   64
            Top             =   1680
            Width           =   975
         End
         Begin VB.CommandButton cmdAsignarProd 
            Height          =   375
            Index           =   2
            Left            =   120
            Picture         =   "frmProdNUEVA.frx":1D858
            Style           =   1  'Graphical
            TabIndex        =   63
            ToolTipText     =   "Asignar nueva produccion"
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtInicio 
            Height          =   375
            Index           =   2
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   62
            Top             =   1170
            Width           =   1935
         End
         Begin VB.TextBox txtDuracion 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000013&
            Height          =   375
            Index           =   2
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   61
            Top             =   1200
            Width           =   855
         End
         Begin VB.CommandButton cmdVerProd 
            Height          =   375
            Index           =   2
            Left            =   12720
            Picture         =   "frmProdNUEVA.frx":1E25A
            Style           =   1  'Graphical
            TabIndex        =   60
            ToolTipText     =   "ver produccion"
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtTraza 
            Height          =   375
            Index           =   2
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   59
            Top             =   1170
            Width           =   1335
         End
         Begin VB.TextBox txtLineProdPalet 
            Enabled         =   0   'False
            Height          =   375
            Index           =   2
            Left            =   1080
            TabIndex        =   58
            Text            =   " "
            Top             =   1680
            Width           =   1575
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1575
            Index           =   2
            Left            =   6960
            TabIndex        =   67
            Top             =   480
            Width           =   5535
            _ExtentX        =   9763
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Articulo"
               Object.Width           =   5997
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Lote"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Articulo"
            Height          =   255
            Index           =   14
            Left            =   720
            TabIndex        =   75
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Cant. estimada"
            Height          =   195
            Index           =   13
            Left            =   3240
            TabIndex        =   74
            Top             =   1800
            Width           =   1305
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inicio"
            Height          =   195
            Index           =   12
            Left            =   2520
            TabIndex        =   73
            Top             =   1260
            Width           =   375
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Duración"
            Height          =   195
            Index           =   11
            Left            =   5040
            TabIndex        =   72
            Top             =   1260
            Width           =   645
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Trazabilidad"
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   71
            Top             =   1260
            Width           =   1095
         End
         Begin VB.Label lblLinPalet 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lin. palet."
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   70
            Top             =   1800
            Width           =   690
         End
      End
      Begin VB.Frame FrameLinea1 
         Caption         =   "Linea 1"
         ClipControls    =   0   'False
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
         Height          =   2175
         Index           =   1
         Left            =   -74640
         TabIndex        =   38
         Top             =   2880
         Width           =   13335
         Begin VB.TextBox txtCajasEstimadas 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   1
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   172
            Top             =   1680
            Width           =   855
         End
         Begin VB.TextBox txtCodartPPal 
            Height          =   375
            Index           =   1
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   50
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtNomartic 
            BackColor       =   &H80000018&
            Height          =   375
            Index           =   1
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   49
            Top             =   600
            Width           =   4095
         End
         Begin VB.CommandButton cmdCambioLotLin 
            Height          =   375
            Index           =   1
            Left            =   12720
            Picture         =   "frmProdNUEVA.frx":1EC5C
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Cambiar LOTE componente"
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton cmdCerrarProd 
            Height          =   375
            Index           =   1
            Left            =   12720
            Picture         =   "frmProdNUEVA.frx":1F65E
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Finalizar produccion"
            Top             =   1680
            Width           =   375
         End
         Begin VB.TextBox txtCantidad 
            Alignment       =   1  'Right Justify
            Height          =   375
            Index           =   1
            Left            =   4680
            Locked          =   -1  'True
            TabIndex        =   45
            Top             =   1680
            Width           =   975
         End
         Begin VB.CommandButton cmdAsignarProd 
            Height          =   375
            Index           =   1
            Left            =   120
            Picture         =   "frmProdNUEVA.frx":20060
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Asignar nueva produccion"
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtInicio 
            Height          =   375
            Index           =   1
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   43
            Top             =   1170
            Width           =   1935
         End
         Begin VB.TextBox txtDuracion 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000013&
            Height          =   375
            Index           =   1
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   42
            Top             =   1200
            Width           =   855
         End
         Begin VB.CommandButton cmdVerProd 
            Height          =   375
            Index           =   1
            Left            =   12720
            Picture         =   "frmProdNUEVA.frx":20A62
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "ver produccion"
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtTraza 
            Height          =   375
            Index           =   1
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   1170
            Width           =   1335
         End
         Begin VB.TextBox txtLineProdPalet 
            Enabled         =   0   'False
            Height          =   375
            Index           =   1
            Left            =   1080
            TabIndex        =   39
            Text            =   " "
            Top             =   1680
            Width           =   1575
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1575
            Index           =   1
            Left            =   6960
            TabIndex        =   48
            Top             =   480
            Width           =   5535
            _ExtentX        =   9763
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Articulo"
               Object.Width           =   5997
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Lote"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Articulo"
            Height          =   255
            Index           =   9
            Left            =   720
            TabIndex        =   56
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Cant. estimada"
            Height          =   195
            Index           =   8
            Left            =   3240
            TabIndex        =   55
            Top             =   1800
            Width           =   1305
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inicio"
            Height          =   195
            Index           =   7
            Left            =   2520
            TabIndex        =   54
            Top             =   1260
            Width           =   375
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Duración"
            Height          =   195
            Index           =   6
            Left            =   5040
            TabIndex        =   53
            Top             =   1260
            Width           =   645
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Trazabilidad"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   52
            Top             =   1260
            Width           =   1095
         End
         Begin VB.Label lblLinPalet 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lin. palet."
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   51
            Top             =   1800
            Width           =   690
         End
      End
      Begin VB.Frame FrameLinea1 
         Caption         =   "Linea 0"
         ClipControls    =   0   'False
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
         Height          =   2175
         Index           =   0
         Left            =   -74640
         TabIndex        =   2
         Top             =   480
         Width           =   13335
         Begin VB.TextBox txtCajasEstimadas 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   0
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   171
            Top             =   1680
            Width           =   855
         End
         Begin VB.TextBox txtLineProdPalet 
            Enabled         =   0   'False
            Height          =   375
            Index           =   0
            Left            =   1080
            TabIndex        =   36
            Text            =   " "
            Top             =   1680
            Width           =   1575
         End
         Begin VB.TextBox txtTraza 
            Height          =   375
            Index           =   0
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   35
            Top             =   1170
            Width           =   1335
         End
         Begin VB.CommandButton cmdVerProd 
            Height          =   375
            Index           =   0
            Left            =   12720
            Picture         =   "frmProdNUEVA.frx":21464
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "ver produccion"
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtDuracion 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000013&
            Height          =   375
            Index           =   0
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   1200
            Width           =   855
         End
         Begin VB.TextBox txtInicio 
            Height          =   375
            Index           =   0
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   1170
            Width           =   1935
         End
         Begin VB.CommandButton cmdAsignarProd 
            Height          =   375
            Index           =   0
            Left            =   120
            Picture         =   "frmProdNUEVA.frx":21E66
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Asignar nueva produccion"
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtCantidad 
            Alignment       =   1  'Right Justify
            Height          =   375
            Index           =   0
            Left            =   4680
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   1680
            Width           =   975
         End
         Begin VB.CommandButton cmdCerrarProd 
            Height          =   375
            Index           =   0
            Left            =   12720
            Picture         =   "frmProdNUEVA.frx":22868
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Finalizar produccion"
            Top             =   1680
            Width           =   375
         End
         Begin VB.CommandButton cmdCambioLotLin 
            Height          =   375
            Index           =   0
            Left            =   12720
            Picture         =   "frmProdNUEVA.frx":2326A
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Cambiar LOTE componente"
            Top             =   1080
            Width           =   375
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1575
            Index           =   0
            Left            =   6960
            TabIndex        =   13
            Top             =   480
            Width           =   5535
            _ExtentX        =   9763
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Articulo"
               Object.Width           =   5997
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Lote"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.TextBox txtNomartic 
            BackColor       =   &H80000018&
            Height          =   375
            Index           =   0
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   600
            Width           =   4095
         End
         Begin VB.TextBox txtCodartPPal 
            Height          =   375
            Index           =   0
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label lblLinPalet 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lin. palet."
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   37
            Top             =   1800
            Width           =   690
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Trazabilidad"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   24
            Top             =   1260
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Duración"
            Height          =   195
            Index           =   3
            Left            =   5040
            TabIndex        =   23
            Top             =   1260
            Width           =   645
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inicio"
            Height          =   195
            Index           =   2
            Left            =   2520
            TabIndex        =   21
            Top             =   1260
            Width           =   375
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Cant. estimada"
            Height          =   195
            Index           =   1
            Left            =   3240
            TabIndex        =   18
            Top             =   1800
            Width           =   1305
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Articulo"
            Height          =   255
            Index           =   0
            Left            =   720
            TabIndex        =   14
            Top             =   360
            Width           =   855
         End
      End
      Begin MSComctlLib.ListView lwp2 
         Height          =   6735
         Index           =   1
         Left            =   -74880
         TabIndex        =   182
         Top             =   720
         Width           =   13575
         _ExtentX        =   23945
         _ExtentY        =   11880
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Referencia"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripion"
            Object.Width           =   6773
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Stock"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "stock en currency-oculta"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Pedidos"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lwp 
         Height          =   4095
         Left            =   240
         TabIndex        =   184
         Top             =   720
         Width           =   13455
         _ExtentX        =   23733
         _ExtentY        =   7223
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Referencia"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripion"
            Object.Width           =   6773
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "cajaspal"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "unicajas"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "SUMA"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "STOCK"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "PDTE"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView lwp2 
         Height          =   2175
         Index           =   0
         Left            =   2400
         TabIndex        =   189
         Top             =   5160
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   3836
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Referencia"
            Object.Width           =   2717
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripion"
            Object.Width           =   7479
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Pedido"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Sem."
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Pal. pedidos"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Stock"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "PDTE"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Pendiente oculta"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "udsporpalet"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Line LineaIndicadoraP 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   5
         X1              =   -63360
         X2              =   -61200
         Y1              =   7440
         Y2              =   7440
      End
      Begin VB.Line LineaIndicadoraP 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   4
         X1              =   -65760
         X2              =   -63600
         Y1              =   7440
         Y2              =   7440
      End
      Begin VB.Line LineaIndicadoraP 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   3
         X1              =   -68040
         X2              =   -65880
         Y1              =   7440
         Y2              =   7440
      End
      Begin VB.Line LineaIndicadoraP 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   2
         X1              =   -70320
         X2              =   -68160
         Y1              =   7440
         Y2              =   7440
      End
      Begin VB.Line LineaIndicadoraP 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   1
         X1              =   -72600
         X2              =   -70440
         Y1              =   7440
         Y2              =   7440
      End
      Begin VB.Line LineaIndicadoraP 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   0
         X1              =   -74880
         X2              =   -72720
         Y1              =   7440
         Y2              =   7440
      End
      Begin VB.Line LineaIndicadora 
         BorderWidth     =   2
         Index           =   10
         Visible         =   0   'False
         X1              =   -74760
         X2              =   -74760
         Y1              =   5280
         Y2              =   7320
      End
      Begin VB.Line LineaIndicadora 
         BorderWidth     =   2
         Index           =   9
         Visible         =   0   'False
         X1              =   -74760
         X2              =   -74760
         Y1              =   3000
         Y2              =   5040
      End
      Begin VB.Line LineaIndicadora 
         BorderWidth     =   2
         Index           =   8
         Visible         =   0   'False
         X1              =   -74760
         X2              =   -74760
         Y1              =   600
         Y2              =   2520
      End
      Begin VB.Line LineaIndicadora 
         BorderWidth     =   2
         Index           =   7
         Visible         =   0   'False
         X1              =   -74760
         X2              =   -74760
         Y1              =   4320
         Y2              =   6360
      End
      Begin VB.Line LineaIndicadora 
         BorderWidth     =   2
         Index           =   6
         Visible         =   0   'False
         X1              =   -74760
         X2              =   -74760
         Y1              =   1080
         Y2              =   3000
      End
      Begin VB.Line LineaIndicadora 
         BorderWidth     =   2
         Index           =   5
         Visible         =   0   'False
         X1              =   -74760
         X2              =   -74760
         Y1              =   5280
         Y2              =   7320
      End
      Begin VB.Line LineaIndicadora 
         BorderWidth     =   2
         Index           =   4
         Visible         =   0   'False
         X1              =   -74760
         X2              =   -74760
         Y1              =   3000
         Y2              =   5040
      End
      Begin VB.Line LineaIndicadora 
         BorderWidth     =   2
         Index           =   3
         Visible         =   0   'False
         X1              =   -74760
         X2              =   -74760
         Y1              =   600
         Y2              =   2520
      End
      Begin VB.Line LineaIndicadora 
         BorderWidth     =   2
         Index           =   2
         Visible         =   0   'False
         X1              =   -74760
         X2              =   -74760
         Y1              =   5280
         Y2              =   7320
      End
      Begin VB.Line LineaIndicadora 
         BorderWidth     =   2
         Index           =   1
         Visible         =   0   'False
         X1              =   -74760
         X2              =   -74760
         Y1              =   3000
         Y2              =   5040
      End
      Begin VB.Line LineaIndicadora 
         BorderWidth     =   2
         Index           =   0
         Visible         =   0   'False
         X1              =   -74760
         X2              =   -74760
         Y1              =   720
         Y2              =   2640
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00400000&
         BorderWidth     =   3
         Index           =   2
         X1              =   240
         X2              =   13680
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Label lblPlann 
         Caption         =   "Semana entrega"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   190
         Top             =   5280
         Width           =   1800
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004000&
         BorderWidth     =   3
         Index           =   1
         X1              =   7680
         X2              =   13680
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblPlann 
         Caption         =   "Producto acabado"
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
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   186
         Top             =   480
         Width           =   1800
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004000&
         BorderWidth     =   3
         Index           =   0
         X1              =   2280
         X2              =   5520
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblPlann 
         Caption         =   "Totales(Palets)"
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
         Height          =   195
         Index           =   2
         Left            =   6000
         TabIndex        =   185
         Top             =   480
         Width           =   1650
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004040&
         BorderWidth     =   3
         Index           =   3
         X1              =   -73560
         X2              =   -61320
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblPlann 
         Caption         =   "Materia auxliar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   195
         Index           =   4
         Left            =   -74880
         TabIndex        =   183
         Top             =   480
         Width           =   1800
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Lineas producción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   7920
      Width           =   7485
      Begin VB.Image imgLinea 
         Height          =   360
         Index           =   9
         Left            =   6360
         ToolTipText     =   "Produccion manual (II)"
         Top             =   300
         Width           =   345
      End
      Begin VB.Image imgLinea 
         Height          =   360
         Index           =   7
         Left            =   4680
         Top             =   300
         Width           =   360
      End
      Begin VB.Image imgLinea 
         Height          =   360
         Index           =   6
         Left            =   4200
         Top             =   300
         Width           =   360
      End
      Begin VB.Image imgLinea 
         Height          =   360
         Index           =   5
         Left            =   3360
         Top             =   300
         Width           =   360
      End
      Begin VB.Image imgLinea 
         Height          =   360
         Index           =   2
         Left            =   1330
         Top             =   300
         Width           =   360
      End
      Begin VB.Image imgLinea 
         Height          =   360
         Index           =   10
         Left            =   7080
         ToolTipText     =   "Produccion MUESTRAS"
         Top             =   300
         Width           =   345
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Mu"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   14
         Left            =   6720
         TabIndex        =   267
         Top             =   300
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "M2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   12
         Left            =   6000
         TabIndex        =   244
         Top             =   300
         Width           =   375
      End
      Begin VB.Image imgLinea 
         Height          =   360
         Index           =   8
         Left            =   5640
         ToolTipText     =   "Produccion manual"
         Top             =   300
         Width           =   345
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "M1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   11
         Left            =   5280
         TabIndex        =   181
         Top             =   300
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   7
         Left            =   4560
         TabIndex        =   10
         Top             =   300
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   3960
         TabIndex        =   9
         Top             =   300
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   3240
         TabIndex        =   8
         Top             =   300
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   1920
         TabIndex        =   6
         Top             =   300
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   1200
         TabIndex        =   5
         Top             =   300
         Width           =   165
      End
      Begin VB.Image imgLinea 
         Height          =   360
         Index           =   4
         Left            =   2760
         Top             =   300
         Width           =   360
      End
      Begin VB.Image imgLinea 
         Height          =   360
         Index           =   3
         Left            =   2160
         Top             =   300
         Width           =   360
      End
      Begin VB.Image imgLinea 
         Height          =   360
         Index           =   1
         Left            =   790
         Top             =   300
         Width           =   360
      End
      Begin VB.Image imgLinea 
         Height          =   360
         Index           =   0
         Left            =   270
         Top             =   300
         Width           =   360
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   2
         Left            =   0
         Picture         =   "frmProdNUEVA.frx":23C6C
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   1080
         Picture         =   "frmProdNUEVA.frx":256DE
         Top             =   1800
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   0
         Left            =   480
         Picture         =   "frmProdNUEVA.frx":27150
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   660
         TabIndex        =   4
         Top             =   300
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   2640
         TabIndex        =   7
         Top             =   300
         Width           =   165
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Paletización"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7800
      TabIndex        =   32
      Top             =   7920
      Width           =   3495
      Begin VB.Image imgLineaPalet 
         Height          =   360
         Index           =   5
         Left            =   3060
         Top             =   300
         Width           =   345
      End
      Begin VB.Image imgLineaPalet 
         Height          =   360
         Index           =   4
         Left            =   2520
         Top             =   300
         Width           =   345
      End
      Begin VB.Image imgLineaPalet 
         Height          =   360
         Index           =   3
         Left            =   1920
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label1 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   1740
         TabIndex        =   246
         Top             =   300
         Width           =   255
      End
      Begin VB.Image imgLineaPalet 
         Height          =   360
         Index           =   1
         Left            =   840
         Top             =   300
         Width           =   315
      End
      Begin VB.Image imgLineaPalet 
         Height          =   360
         Index           =   0
         Left            =   270
         Top             =   300
         Width           =   345
      End
      Begin VB.Label Label1 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   720
         TabIndex        =   34
         Top             =   300
         Width           =   255
      End
      Begin VB.Image imgLineaPalet 
         Height          =   360
         Index           =   2
         Left            =   1320
         Top             =   300
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   1200
         TabIndex        =   33
         Top             =   300
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   2370
         TabIndex        =   340
         Top             =   323
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   2880
         TabIndex        =   341
         Top             =   310
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   339
         Top             =   300
         Width           =   255
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   12120
      TabIndex        =   29
      Top             =   7920
      Width           =   1935
      Begin VB.CommandButton cmdPlanning 
         Height          =   465
         Left            =   480
         Picture         =   "frmProdNUEVA.frx":28BC2
         Style           =   1  'Graphical
         TabIndex        =   180
         ToolTipText     =   "Refrescar datos"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Height          =   465
         Left            =   1440
         Picture         =   "frmProdNUEVA.frx":295C4
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Salir"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton cmdPistola 
         Height          =   465
         Left            =   960
         Picture         =   "frmProdNUEVA.frx":29B4E
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Pistola"
         Top             =   180
         Width           =   375
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Poste"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11400
      TabIndex        =   179
      Top             =   7920
      Width           =   710
      Begin VB.Image imgPoste 
         Height          =   360
         Index           =   0
         Left            =   120
         Top             =   300
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmProdNUEVA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Si ve uno o/y otro
Public PermisoProduccion As Boolean
Public PermisoPlanning As Boolean

Private Const SegundosDeRefresco = 15  ' Cada cuantos segundos se va ir a mirar a la BD a ver si ha cambiado algo
Private Const RefrescarTodosLosDatos = 20 'ej:cada 20 bloques de segundosrefrescos refrescara todos
                                          'Es decir 20* 15=300 segundos->5 minutos

Dim BloquesSegundosDeRefrescos As Byte

Dim LineasDeProduccion(10) As cLineaProduccion  'abril 2012
Dim LineasDePaletizacion(5) As CPalet  'Abril 2012: 4 lineas    Junio 2012: 6 lineas
Dim seg As Byte


Dim PrimeraVez As Boolean

Dim Aux As String
Dim Orden_lwp As Byte
Dim Asc_lwp As Boolean

Dim ParalecturaPoste As Byte

Dim AuxLP As cLineaProduccion
Dim MinAnyo As Integer
Private Const NumeroLineasProduccion = 10
'Abril 2012:   8 lineas, la ultima la MANUAL
'Abril 2012:   9 lineas, la ultima la muestras
'Abril 2012 :  10 OOtra linea manual  ---> Seran 8,9 Manual 10: muestras

'Octubre 2012
'-------------------
'Se definen dos permisos para poder ver este formulario
'   -Se ven lineas(y palets)
'   -Se ven planning
'Cuando entra un usuario veremos si tiene permisos




'Modificacion  14/12/2011
'
'ANTES:  LINEA -1
'   AHORA las lineas empiezan en el cero, hasta el 8
Private Sub PonerDatosLinea(ByRef LinPr As cLineaProduccion)
Dim linea As Byte
Dim L As cLineaProCompo
Dim J As Integer
Dim It As ListItem

    With LinPr
        
        linea = .linea
        Me.txtCodartPPal(linea).Text = .codartic
        Me.txtNomartic(linea).Text = .NomArtic
        txtCantidad(linea).Text = Format(.CantidadEstimada, FormatoCantidad)
        Me.txtTraza(linea).Text = .LoteTrazabilidad
        Me.txtInicio(linea).Text = .FH_Incio
        
        ListView1(linea).ListItems.Clear
        For J = 1 To .CuantasMP
            If .DevuelveComponenteLinea(J, L) Then
                'OK, ha ledio la sublinea de componentes
                Set It = ListView1(linea).ListItems.Add
                It.Text = L.NomArticCompo
                It.SubItems(1) = L.LoteMateria
            End If
        Next
        
        If .UnidadesCaja > 0 Then
            J = .CantidadEstimada \ .UnidadesCaja
            Me.txtCajasEstimadas(linea).Text = Format(J, "#,##0")
        End If
        
        
        
        PonerDuracion linea
        
        'Estado. Para pintar luz verde o roja
        If .Estado = 1 Then
            J = 1
        Else
            J = 2
        End If
        
        Me.txtLineProdPalet(linea).Text = .LeerLineaDondeEstaPaletizando
        
    End With
    Set L = Nothing
    'El icono de abajo
    
    Me.imgLinea(linea).Picture = Me.Image1(J).Picture '0 azul   1 verde    2 rojo
    Me.FrameLinea1(linea).ForeColor = &H8000&
    cmdAsignarProd(linea).visible = False
    Me.cmdCambioLotLin(linea).visible = True
    Me.cmdCerrarProd(linea).visible = True
    cmdVerProd(linea).visible = True
End Sub



Private Sub PonerDatosLineaPalet(ByRef cP As CPalet, DesdePonerTodosLosDatos As Boolean)
Dim linea As Byte
Dim Total As Integer
Dim J As Integer
Dim Cadena As String

    With cP
        
        linea = .LineaPeletizacion - 1
        Me.txtIdP(linea).Text = .ID
        Me.txtInicioP(linea).Text = .FechaInicio
        
       
        
    
        'La linea 8-9 (MANUAL/muestras) NO SE PALETIZA, por eso pone: to 7
        txtQueProdLin(.LineaPeletizacion - 1).Text = ""
        For J = 0 To 7
            Total = 0
            If txtNomartic(J).Text <> "" Then
                If .LineasProd(J) Then
'                     Set IT = Me.lwLinPalet(.LineaPeletizacion - 1).ListItems.Add()
'
'
'                     IT.Text = J
'                     IT.Bold = True
'                     IT.ForeColor = vbBlue
'                     IT.SubItems(1) = Mid(txtNomartic(J).Text, 1, 25)
'                     IT.ToolTipText = txtNomartic(J).Text
                     Cadena = J & ".- " & txtNomartic(J).Text & " (" & txtTraza(J).Text & ")"
                    If Total = 0 Then
                        'Es el primero veremos la longitud
                        If Len(Cadena) > 12 Then
                        
                        Else
                        
                        End If
                    End If
                    If txtQueProdLin(.LineaPeletizacion - 1).Text <> "" Then txtQueProdLin(.LineaPeletizacion - 1).Text = txtQueProdLin(.LineaPeletizacion - 1).Text & vbCrLf & vbCrLf
                    txtQueProdLin(.LineaPeletizacion - 1).Text = txtQueProdLin(.LineaPeletizacion - 1).Text & Cadena
                End If
            End If
            
        Next J
        
        
'        'Para cada linea de produccion vere que paletiza
        If Not DesdePonerTodosLosDatos Then
            'Stop
            For J = 0 To NumeroLineasProduccion
                If Not (LineasDeProduccion(J) Is Nothing) Then Me.txtLineProdPalet(J).Text = LineasDeProduccion(J).LeerLineaDondeEstaPaletizando
            Next J
        End If

        
        txtDuracionP(linea).Text = Now
        PonerDuracionPalet linea
       
       
        .CargaCajasPaletList Total, Me.lstCajas(linea)
        txtCajas(linea).Text = Total
        Me.imgLineaPalet(linea).Picture = Me.Image1(1).Picture '0 azul   1 verde    2 rojo
       
        Me.FramePalet(linea).ForeColor = &H8000&
        cmdInciarPalet(linea).visible = False
        Me.cmdCerrarPalet(linea).visible = True

    

    End With
End Sub



Private Sub PonerDuracion(Index As Byte)
Dim F1 As Date
    
    On Error GoTo ED
    If txtInicio(Index).Text = "" Then
        txtDuracion(Index).Text = ""
    Else
        F1 = CDate(Me.txtInicio(Index).Text)
        If F1 > Now Then
            txtDuracion(Index).Text = "Mayor"
        Else
            F1 = Now - F1
            txtDuracion(Index).Text = Format(F1, "hh:mm:ss")
        End If
    End If
ED:
    If Err.Number <> 0 Then
        Err.Clear
        Me.txtDuracion(Index).Text = "Error"
    End If
End Sub

Private Sub PonerDuracionPalet(Index As Byte)
Dim F1 As Date
    
    On Error GoTo ED
    If txtDuracionP(Index).Text = "" Then
        txtDuracionP(Index).Text = ""
    Else
        F1 = CDate(Me.txtInicioP(Index).Text)
        If F1 > Now Then
            txtDuracionP(Index).Text = "Mayor"
        Else
            F1 = Now - F1
            txtDuracionP(Index).Text = Format(F1, "hh:mm:ss")
        End If
    End If
ED:
    If Err.Number <> 0 Then
        Err.Clear
        Me.txtDuracionP(Index).Text = "Error"
    End If
End Sub



'LINEA -1-->Ahora sin -1
Private Sub LimpiarLinea(ByVal linea As Byte)
    'linea = linea - 1
    Me.txtCodartPPal(linea).Text = ""
    Me.txtNomartic(linea).Text = ""
    Me.ListView1(linea).ListItems.Clear
    txtCantidad(linea).Text = ""
    Me.txtCajasEstimadas(linea).Text = ""
    Me.txtLineProdPalet(linea).Text = ""
    'El icono de abajo
    Me.imgLinea(linea).Picture = Me.Image1(0).Picture '0azul   1 verde    2 rojo  3-naranja
    Me.FrameLinea1(linea).ForeColor = vbBlack
    cmdAsignarProd(linea).visible = True
    Me.cmdCambioLotLin(linea).visible = False
    Me.cmdCerrarProd(linea).visible = False
    cmdVerProd(linea).visible = False
    
    'Tiempos
    txtDuracion(linea).Text = ""
    txtInicio(linea).Text = ""
    txtTraza(linea).Text = ""
End Sub


Private Sub LimpiarLineaPalet(ByVal linea As Byte)
    linea = linea - 1
    
    Me.imgLineaPalet(linea).Picture = Me.Image1(0).Picture '0azul   1 verde    2 rojo
    Me.FramePalet(linea).ForeColor = vbBlack
    txtQueProdLin(linea).Text = ""

    txtIdP(linea).Text = ""
    txtDuracionP(linea).Text = ""
    txtInicioP(linea).Text = ""
    Me.lstCajas(linea).Clear
    txtCajas(linea).Text = ""
    cmdInciarPalet(linea).visible = True
    Me.cmdCerrarPalet(linea).visible = False

    
End Sub






Private Sub cmdAsignarProd_Click(Index As Integer)
Dim Salir As Boolean
Dim LineasPaletizado As String

    Timer1.Enabled = False
    Salir = False
    If Me.lwp.ListItems.Count = 0 Then
        MsgBox "Cargue el planning de producción", vbExclamation
        Salir = True
    End If
    

    
    
    If Not ComprobarLineaAHORA(CByte(Index), True) Then Salir = True
    Timer1.Enabled = False 'de nuevo
    
    
    If Not Salir Then
        'Comprobaremos que las lineas de paletizado estan libres
        LineasPaletizado = ""
        If Index < 8 Then
            For NumRegElim = 0 To 5
                If Me.txtIdP(NumRegElim).Text = "" Then
                    LineasPaletizado = LineasPaletizado & "0"
                Else
                    LineasPaletizado = LineasPaletizado & "1"
                End If
            Next
            
            If InStr(1, LineasPaletizado, "0") = 0 Then
                MsgBox "Ninguna linea de paletizacion disponible", vbExclamation
                Salir = True
            End If
        End If
    End If
    
    If Salir Then
        Timer1.Enabled = True
        Exit Sub
    End If
   



    


    If BloqueoManual("PonerEnProd", "1") Then
        OcultarLineaIndicadora
        OcultarLineaIndicadoraPalet
        Me.LineaIndicadora(Index).visible = True
        
        Set frmProduNuevaCRUD2.cLP = Nothing
        frmProduNuevaCRUD2.Modo = Index 'LLEVARA La linea de produccion donde vamos a producir el producto producido
        frmProduNuevaCRUD2.Show vbModal
        If CadenaDesdeOtroForm <> "" Then
            'ALgo ha cambiado. Tendre que leer otra vez la produccion
            Set LineasDeProduccion(Index) = Nothing
            Set miRsAux = New ADODB.Recordset
            If LeerLinea(CByte(Index)) Then
                PonerDatosLinea LineasDeProduccion(Index)
                
                
                   'Ahora mandamos directamente a PALETIZAR si linea <8
                    If Index < 8 Then
                        frmPaletProduccion.DesdeNuevaProd = LineasPaletizado
                        
                        If Index = 0 Then
                            CadenaDesdeOtroForm = "1" & String(8, "0")
                            LineasPaletizado = txtTraza(Index).Text & String(8, "|")
                            
                        Else
                            CadenaDesdeOtroForm = String(8, "0")
                            CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, 1, Index) & "1" & Mid(CadenaDesdeOtroForm, Index + 2)
                  
                            LineasPaletizado = String(Index, "|")
                            LineasPaletizado = LineasPaletizado & Me.txtTraza(Index).Text & String(8 - Index, "|")
                        End If
                        frmPaletProduccion.Varios = "0|" & CadenaDesdeOtroForm & "|"
                        
                        frmPaletProduccion.TrazaEnLineas = LineasPaletizado
                        Set frmPaletProduccion.cPal = Nothing
                        CadenaDesdeOtroForm = ""
                        frmPaletProduccion.Show vbModal
                        If CadenaDesdeOtroForm <> "" Then
                            'leemos las lineas
                            Set miRsAux = New ADODB.Recordset
                            'Antes
                            
                           
                            
                            
                            For NumRegElim = 0 To 5
                                If Me.txtIdP(NumRegElim).Text = "" Then
                                    If LeerLineaPalet(CByte(NumRegElim) + 1) Then PonerDatosLineaPalet LineasDePaletizacion(NumRegElim), False
                                End If
                            Next
                        End If
                End If '<8
        
                
                
                
            Else
                'PonerDatosLinea LineasDeProduccion(0)
                LimpiarLinea CByte(Index)
            End If
            Set miRsAux = Nothing
            
            
         
            
        End If
        
        DesBloqueoManual "PonerEnProd"
        
        
        
        
        
        
    End If
    
    Timer1.Enabled = True
End Sub

Private Sub cmdBuscarRef_Click(Index As Integer)
Dim I As Integer

    If lwp2(0).ListItems.Count = 0 Then Exit Sub
    If Me.lwp2(0).SelectedItem Is Nothing Then Exit Sub
    
    NumRegElim = -1
    If Index = 0 Then
        For I = Me.lwp2(0).SelectedItem.Index - 1 To 1 Step -1
            If Me.lwp2(0).ListItems(I).Text = lwp2(0).SelectedItem.Text Then
                NumRegElim = I
                Exit For
            End If
        Next
    Else
        For I = Me.lwp2(0).SelectedItem.Index + 1 To lwp2(0).ListItems.Count
            If Me.lwp2(0).ListItems(I).Text = lwp2(0).SelectedItem.Text Then
                NumRegElim = I
                Exit For
            End If
        Next
    End If
    If NumRegElim > 0 Then
        Set lwp2(0).SelectedItem = lwp2(0).ListItems(NumRegElim)
        lwp2(0).SelectedItem.EnsureVisible
    End If
End Sub

Private Sub cmdCambioLotLin_Click(Index As Integer)
    'Vamos acambiar el lote.
    'Preguntara cuanto se ha producido en esta tanda, y nuevo lote
    If Me.ListView1(Index).ListItems.Count = 0 Then Exit Sub
    If Me.ListView1(Index).SelectedItem Is Nothing Then Exit Sub
    
    If Not ComprobarLineaAHORA(CByte(Index), False) Then Exit Sub
    
    
    
    Timer1.Enabled = False
    PoneLaLineaIndicadora Index
    
    
    If LineasDeProduccion(Index).Estado = 2 Then
        MsgBox "Faltan lotes por asignar", vbExclamation
    
    Else
        Set frmProduNuevaCRUD2.cLP = LineasDeProduccion(Index)
        frmProduNuevaCRUD2.SubLinea = ListView1(Index).SelectedItem.Index
        frmProduNuevaCRUD2.Modo = 2
        frmProduNuevaCRUD2.Show vbModal
        
        'Refrescamos los datos de la linea
        PonerDatosLinea LineasDeProduccion(Index)
        
    End If
    HabilitaTimer
End Sub

Private Sub cmdCerrarPalet_Click(Index As Integer)
Dim J As Integer
Dim cP As CPalet
    'CERRAR PALET
    Timer1.Enabled = False
    
    
    'Puede cerrar el palet desde la pistola, por lo tanto, COMPROBAREMOS que esta aquin
    NumRegElim = 0
    If Me.txtIdP(Index).Text <> "" Then
        NumRegElim = Me.txtIdP(Index).Text
        Set cP = New CPalet
        cP.Leer NumRegElim
        If cP.ID <> LineasDePaletizacion(Index).ID Then
            MsgBox "Id palet distintos", vbExclamation
            NumRegElim = 0
        End If
        Set cP = Nothing
    Else
        MsgBox "Linea vacia", vbExclamation
    End If
    

    
    If NumRegElim > 0 Then
        LineasDePaletizacion(Index).Leer NumRegElim
    
        If LineasDePaletizacion(Index).FechaFin > CDate("2000-01-01") Then
            MsgBox "Error leyendo palet. Palet cerrado", vbExclamation
            PonerValoresLineasProduccion_
        Else
            frmPaletProduccion.Varios = ""
            frmPaletProduccion.TrazaEnLineas = ""
            Set frmPaletProduccion.cPal = LineasDePaletizacion(Index)
            CadenaDesdeOtroForm = ""
            frmPaletProduccion.Show vbModal
            If CadenaDesdeOtroForm <> "" Then
                'leemos las lineas
                    If LeerLineaPalet(Index + 1) Then
                        PonerDatosLineaPalet LineasDePaletizacion(Index), False
                    Else
                        LimpiarLineaPalet Index + 1
                        
                        'Quito de donde esta paletizando
                        For J = 0 To NumeroLineasProduccion
                            If Not (LineasDeProduccion(J) Is Nothing) Then
                                Me.txtLineProdPalet(J).Text = LineasDeProduccion(J).LeerLineaDondeEstaPaletizando
                            Else
                                Me.txtLineProdPalet(J).Text = ""
                            End If
                                
                        Next J
                        
                    End If
            End If
        End If
    End If 'numregelim
    HabilitaTimer
End Sub

Private Sub cmdCerrarProd_Click(Index As Integer)
Dim CerrarPaletLineaManual As Boolean
Dim cP As CPalet
Dim Tot As Integer
Dim C As Collection

    'Vamos acambiar el lote.
    'Preguntara cuanto se ha producido en esta tanda, y cerraremos todos los valores
    If Me.ListView1(Index).ListItems.Count = 0 Then Exit Sub
    If Me.ListView1(Index).SelectedItem Is Nothing Then Exit Sub
    
    If Not ComprobarLineaAHORA(CByte(Index), False) Then Exit Sub
    
    PoneLaLineaIndicadora Index
    
    'Veamos si no se ha cerrado el palet
    Aux = LineasDeProduccion(Index).LeerLineaDondeEstaPaletizando
    If Aux <> "" Then
        CerrarPaletLineaManual = False
        If Mid(Aux, 1, 2) = "L0" Then
            'Es la linea manual
            CerrarPaletLineaManual = True 'Lo cierro luego
            Aux = Mid(Aux, InStr(1, Aux, "-") + 1)
            Set cP = New CPalet
            If Not cP.Leer(CLng(Aux)) Then
                Set cP = Nothing
                Exit Sub
            End If
        Else
            'Aun se esta paletizando. El palet NO ha sido cerrado
            Aux = "Todavia esta en paletizado: " & Aux & vbCrLf & vbCrLf & "¿Continuar?"
            If MsgBox(Aux, vbQuestion + vbYesNo) = vbNo Then Exit Sub
            Aux = "No deberia seguir.  No podemos(todavia) asignar otra produccion al palet en curso"
            MsgBox Aux, vbExclamation
            Exit Sub
        End If
    End If
    
    
    
    Timer1.Enabled = False
    

    
            Set frmProduNuevaCRUD2.cLP = LineasDeProduccion(Index)
            frmProduNuevaCRUD2.SubLinea = ListView1(Index).SelectedItem.Index
            frmProduNuevaCRUD2.Modo = 3
            frmProduNuevaCRUD2.Show vbModal
            
            'Refrescamos los datos de la linea
            'Volveremos a leer
            If CadenaDesdeOtroForm <> "" Then
                Set miRsAux = New ADODB.Recordset
                
                If CerrarPaletLineaManual Then
                    cP.CargaDatosPalet C, False, Tot, False
                    cP.CerrarPalet Tot
                    Set cP = Nothing
                End If
                
                
                If LeerLinea(CByte(Index)) Then
                    PonerDatosLinea LineasDeProduccion(Index)
                Else
                    Set LineasDeProduccion(Index) = Nothing
                    LimpiarLinea Index
                End If
                Set miRsAux = Nothing
            End If
   
    
    HabilitaTimer
End Sub

Private Sub cmdInciarPalet_Click(Index As Integer)
Dim TrazaEnLineas As String


    'Nuevo palet
    CadenaDesdeOtroForm = ""
    TrazaEnLineas = ""
    For NumRegElim = 0 To NumeroLineasProduccion
        If Me.txtTraza(NumRegElim).Text <> "" Then
            If LineasDeProduccion(NumRegElim).Estado = 1 Then
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & "1"
                
                If Me.txtLineProdPalet(NumRegElim).Text <> "" Then
                    'YA esta en linea de produccion
                    TrazaEnLineas = TrazaEnLineas & "#" & txtTraza(NumRegElim).Text   'marcamos que la linea esta en paletizandose
                Else
                    'Ok. Esta sin asignar a paletizar
                    TrazaEnLineas = TrazaEnLineas & txtTraza(NumRegElim).Text
                End If
            Else
                'La linea NO esta disponible todavia
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & "0"
            End If
        Else
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & "0"
        End If
        TrazaEnLineas = TrazaEnLineas & "|"
    Next
    CadenaDesdeOtroForm = Index + 1 & "|" & CadenaDesdeOtroForm & "|"
    Timer1.Enabled = False
    frmPaletProduccion.DesdeNuevaProd = ""
    frmPaletProduccion.Varios = CadenaDesdeOtroForm
    frmPaletProduccion.TrazaEnLineas = TrazaEnLineas
    Set frmPaletProduccion.cPal = Nothing
    CadenaDesdeOtroForm = ""
    frmPaletProduccion.Show vbModal
    If CadenaDesdeOtroForm <> "" Then
        'leemos las lineas
        Set miRsAux = New ADODB.Recordset
        'Antes
        If LeerLineaPalet(CByte(Index) + 1) Then PonerDatosLineaPalet LineasDePaletizacion(Index), False
       
        Set miRsAux = Nothing
    End If
    HabilitaTimer
End Sub

Private Sub cmdPistola_Click()
    Timer1.Enabled = False
    frmPist1.Show vbModal
    'PonerValoresLineasProduccion_
    Me.pb1.Value = 0
    Me.FramePB.visible = True
    Label3.Caption = "Leyendo datos"
    Label3.Refresh
    seg = SegundosDeRefresco 'Asi refresca los valores YA
    BloquesSegundosDeRefrescos = RefrescarTodosLosDatos
    ParalecturaPoste = 3
    HabilitaTimer
    
End Sub

Private Sub cmdPlanning_Click()
    Timer1.Enabled = False
    Screen.MousePointer = vbHourglass
    PonerValoresLineasProduccion_
    Screen.MousePointer = vbDefault
    Timer1.Enabled = True
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdVerProd_Click(Index As Integer)
    Timer1.Enabled = False
    PoneLaLineaIndicadora Index
    Set frmProduNuevaCRUD2.cLP = LineasDeProduccion(Index)
    frmProduNuevaCRUD2.Modo = 1
    frmProduNuevaCRUD2.Show vbModal
    HabilitaTimer
End Sub



Private Sub Form_Activate()


    If PrimeraVez Then
        PrimeraVez = False
        pb1.Value = 0
        DoEvents
        Espera 0.2
        PreparaListviewPlanning
        
        
        
        PonerValoresLineasProduccion_
        
        ComprobarPoste True
        Timer1.Enabled = True
    End If
End Sub

Private Function LeerLinea(KLinea As Byte) As Boolean
Dim I As Byte
Dim SQL As String

    LeerLinea = False
    
    SQL = "select prodlin.codigo,prodlin.idlin ,lotetraza from prodlin,prodtrazlin  where prodlin.codigo= prodtrazlin.codigo AND prodlin.idlin = prodtrazlin.idlin"
    SQL = SQL & " and lineaprod = " & KLinea & " and estado >0 and estado<10 ORDER BY lotetraza DESC"  'Pq puede que haya varios cambios de trazabilidad para la misma linea
    miRsAux.Open SQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    I = 0
    If Not miRsAux.EOF Then
        'UNO tiene seguro
        If Not LineasDeProduccion(KLinea) Is Nothing Then Set LineasDeProduccion(KLinea) = Nothing
        Set LineasDeProduccion(KLinea) = New cLineaProduccion
        LineasDeProduccion(KLinea).LeerDesdeTrazabilidad miRsAux!Codigo, miRsAux!idlin, KLinea, miRsAux!lotetraza
        I = 1
        
        'veremos que SOLO hay una linea en marcha
        Do
            miRsAux.MoveNext
        
            If Not miRsAux.EOF Then
                'Si el codigo y prodilin es el mismo, es que solo hay una produccion
                If miRsAux!Codigo <> LineasDeProduccion(KLinea).CodProduccion Or LineasDeProduccion(KLinea).idLiProd <> miRsAux!idlin Then
                    'MAAAAAAAl
                    'Hay mas de una produccion en la linea
                    Set LineasDeProduccion(KLinea) = Nothing
                    I = 2
                End If
            End If
        Loop Until miRsAux.EOF
    End If
    miRsAux.Close
    
    
    
    
    If I = 1 Then LeerLinea = True
End Function



Private Function LeerLineaPalet(KLinea As Byte) As Boolean
Dim I As Byte
Dim SQL As String

    LeerLineaPalet = False
    
    SQL = "select idpalet from prodpalets where  LineaPeletiza = " & KLinea & " and fhFin is null "  '
    miRsAux.Open SQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    I = 0
    If Not miRsAux.EOF Then
        'UNO tiene seguro
        If Not LineasDePaletizacion(KLinea - 1) Is Nothing Then Set LineasDePaletizacion(KLinea - 1) = Nothing
        Set LineasDePaletizacion(KLinea - 1) = New CPalet
        LineasDePaletizacion(KLinea - 1).Leer miRsAux!IdPalet
        I = 1
        
        'veremos que SOLO hay una linea en marcha
        Do
            miRsAux.MoveNext
        
            If Not miRsAux.EOF Then
                
                
                    'Hay mas de una produccion en la linea
                    Set LineasDePaletizacion(KLinea) = Nothing
                    I = 2
                
            End If
        Loop Until miRsAux.EOF
    Else
        If Not LineasDePaletizacion(KLinea - 1) Is Nothing Then Set LineasDePaletizacion(KLinea - 1) = Nothing
    End If
    miRsAux.Close
    
    If I = 1 Then LeerLineaPalet = True
End Function








Private Sub Form_Load()
    PrimeraVez = True
    Me.Icon = frmppal.Icon
    ParalecturaPoste = 0
    
    
    Me.cmdPistola.visible = PermisoProduccion
    'Ocutbre 2012
    'Se podra ver produccion y-o planning en funcion de permisos
    For NumRegElim = 0 To 4
        Me.SSTab1.TabVisible(CInt(NumRegElim)) = PermisoProduccion
    Next
    For NumRegElim = 5 To 6
        Me.SSTab1.TabVisible(CInt(NumRegElim)) = PermisoPlanning
    Next
End Sub


Private Sub HabilitaTimer()
    PonTiempos
    Timer1.Enabled = True
End Sub

Private Sub imgSemana_Click(Index As Integer)

    If Me.lwp.ListItems.Count = 0 Then Exit Sub

    If Index = 0 Then
        If Me.lblPlann(1).Tag = 0 Then Exit Sub 'mas a la izda NO
        Me.lblPlann(1).Tag = Me.lblPlann(1).Tag - 1
    Else
    
        If Trim(Me.lwp.ColumnHeaders(10).Text) = "" Then
            If Me.lwp.ColumnHeaders(5).Text <> "" Then Exit Sub
        End If
    
        Me.lblPlann(1).Tag = Me.lblPlann(1).Tag + 1
    End If
    CargarPlanningSemana
End Sub



Private Sub Frame5_DblClick()
    VerDuplicidades
End Sub

Private Sub imgLinea_Click(Index As Integer)
      OcultarLineaIndicadora
        If Index <= 2 Then
            Me.SSTab1.Tab = 0
        ElseIf Index <= 5 Then
            Me.SSTab1.Tab = 1
        ElseIf Index <= 7 Then
            Me.SSTab1.Tab = 2
        Else
            Me.SSTab1.Tab = 3
        End If
        Me.LineaIndicadora(Index).visible = True
End Sub

Private Sub imgLineaPalet_Click(Index As Integer)
    OcultarLineaIndicadora
    OcultarLineaIndicadoraPalet
    Me.SSTab1.Tab = 4
    Me.LineaIndicadoraP(Index).visible = True
End Sub

Private Sub imgPoste_DblClick(Index As Integer)
    VerDuplicidades
End Sub

Private Sub lwp_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    
    If ColumnHeader.Index > 2 And ColumnHeader.Index <= 7 Then
        MsgBox "Ordenar solo por codigo/descripcion", vbExclamation
        Exit Sub
    End If
    
    If ColumnHeader.Index < 8 Then
        
        If Orden_lwp = ColumnHeader.Index Then
    
            Asc_lwp = Not Asc_lwp
        Else
            Orden_lwp = ColumnHeader.Index
            Asc_lwp = True
        End If
        
        PonerPlanningListview
        CargarPlanningSemana
        
    Else
        'Ha pinchado en una semmana. Buscaremos donde tiene el valor y ensurevis
        Dim J As Integer
        
        NumRegElim = -1
        Aux = ""
        For J = 2 To lwp.ListItems.Count  'la primera fila estan los pedidos
            If Trim(Me.lwp.ListItems(J).SubItems(ColumnHeader.Index - 1)) <> "" Then
                'aqui esta el valor
                If NumRegElim < 0 Then NumRegElim = J
                Aux = Aux & lwp.ListItems(J).Text & " " & lwp.ListItems(J).SubItems(1)
                Aux = Aux & ": " & Me.lwp.ListItems(J).SubItems(ColumnHeader.Index - 1) & vbCrLf
                
            End If
        Next
        If NumRegElim > 0 Then
            Timer1.Enabled = False
                Me.lwp.ListItems(NumRegElim).EnsureVisible
                Me.lwp.SelectedItem = Me.lwp.ListItems(NumRegElim)
                lwp_DblClick
                Aux = "Pedido: " & Me.lwp.ListItems(1).SubItems(ColumnHeader.Index - 1) & vbCrLf & vbCrLf & Aux
                Aux = Aux & vbCrLf & "¿Ver pedido?"
                If MsgBox(Aux, vbQuestion + vbYesNoCancel + vbDefaultButton3) = vbYes Then ImprimirPedido ColumnHeader.Index - 1
                        
             Timer1.Enabled = True
        End If
        
    End If
End Sub

Private Sub lwp_DblClick()
Dim J As Integer
        If Me.lwp.ListItems.Count = 0 Then Exit Sub
        If Me.lwp.SelectedItem Is Nothing Then Exit Sub
        
        
        For J = 1 To lwp2(0).ListItems.Count
            If Me.lwp2(0).ListItems(J).Text = Me.lwp.SelectedItem.Text Then
                'Esta es la referencia
                Set lwp2(0).SelectedItem = lwp2(0).ListItems(J)
                lwp2(0).ListItems(J).Selected = True
                lwp2(0).ListItems(J).EnsureVisible
                Exit For
            End If
        Next
End Sub

Private Sub lwp2_DblClick(Index As Integer)
Dim J As Integer
    If Index = 0 Then
        'en semana
        If Me.lwp2(0).ListItems.Count = 0 Then Exit Sub
        If Me.lwp2(0).SelectedItem Is Nothing Then Exit Sub
        
        
        For J = 1 To lwp.ListItems.Count
            If Me.lwp.ListItems(J).Text = Me.lwp2(0).SelectedItem.Text Then
                'Esta es la referencia
                Set lwp.SelectedItem = lwp.ListItems(J)
                lwp.ListItems(J).Selected = True
                lwp.ListItems(J).EnsureVisible
                Exit For
            End If
        Next
    Else
        'Es materia auxiliar
        If Me.lwp2(1).ListItems.Count = 0 Then Exit Sub
        If Me.lwp2(1).SelectedItem Is Nothing Then Exit Sub
        
'        'Ahora veremos para esa semana
'        Aux = "select sartic.codartic,sartic.nomartic,sementre,scaped.numpedcl,sliped.cantidad from sarti1,sartic,scaped,sliped where scaped.numpedcl=sliped.numpedcl AND"
'        Aux = Aux & " sliped.codartic=sartic.codartic and sarti1.codArtic = sartic.codArtic"
'        Aux = Aux & " and codarti1='" & lwp2(1).SelectedItem.Text & "' order by codartic,numpedcl"
'        Set miRsAux = New ADODB.Recordset
'        miRsAux.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        Aux = ""
'        NumRegElim = 0
'        While Not miRsAux.EOF
'            If InStr(1, Aux, "    - " & miRsAux!NomArtic) = 0 Then Aux = Aux & "    - " & miRsAux!NomArtic & vbCrLf
'            Aux = Aux & Space(10) & " P: " & miRsAux!numpedcl & " --S: " & miRsAux!sementre & "  C:" & miRsAux!Cantidad & vbCrLf
'            NumRegElim = NumRegElim + Val(miRsAux!Cantidad)
'            miRsAux.MoveNext
'        Wend
'        miRsAux.Close
'        Set miRsAux = Nothing
'
'        Aux = lwp2(1).SelectedItem.Text & " - " & lwp2(1).SelectedItem.SubItems(1) & vbCrLf & String(30, "=") & vbCrLf & vbCrLf & Aux
'        Aux = Aux & vbCrLf & " APROX: " & NumRegElim
        
        Timer1.Enabled = False
                'MsgBox Aux, vbInformation
        CadenaDesdeOtroForm = lwp2(1).SelectedItem.Text & "|" & lwp2(1).SelectedItem.SubItems(1) & "|"
        frmVarios.Opcion = 9
        frmVarios.Show vbModal
                
        Timer1.Enabled = True
    End If
End Sub

Private Sub lwp2_GotFocus(Index As Integer)
    lwp2(Index).ToolTipText = ""
    If lwp2(Index).ListItems.Count > 0 Then
        If Not lwp2(Index).SelectedItem Is Nothing Then lwp2(Index).ToolTipText = lwp2(Index).SelectedItem.SubItems(1)
    End If
    
End Sub

Private Sub lwp2_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    lwp2(Index).ToolTipText = Item.SubItems(1)
End Sub

Private Sub lwp2_LostFocus(Index As Integer)
    lwp2(Index).ToolTipText = ""
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    OcultarLineaIndicadora
    OcultarLineaIndicadoraPalet
End Sub

Private Sub Timer1_Timer()
    seg = seg + 1
    PonTiempos
    If seg > SegundosDeRefresco Then
        Screen.MousePointer = vbHourglass
        
        BloquesSegundosDeRefrescos = BloquesSegundosDeRefrescos + 1
        
        
        If BloquesSegundosDeRefrescos > RefrescarTodosLosDatos Then PonerValoresLineasProduccion_
        
        'Haremos consulta en BD
        'Poner cajas
        PonerCajasLeidas
        
        ComprobarPoste False
        
        'Ponemos a 0
        Screen.MousePointer = vbDefault
        seg = 0
    End If
    
End Sub



Private Sub PonTiempos()
Dim J As Byte
    For J = 0 To NumeroLineasProduccion
        If J < 6 Then PonerDuracionPalet J
        PonerDuracion J
    Next J
End Sub



Private Sub PonerValoresLineasProduccion_()
Dim I As Byte
Dim Rc
    
    Rc = Screen.MousePointer
    Screen.MousePointer = vbHourglass

    FramePB.visible = True
    Me.Frame1.visible = False
    Me.Frame3.visible = False
    Me.Frame5.visible = False
    pb1.Max = 21 '10 de produccion + 6 de paletizacion +  5 planning
    pb1.Value = 0
    DoEvents
       
    OcultarLineaIndicadora
    
        Set miRsAux = New ADODB.Recordset
        
        If PermisoProduccion Then
        
            For I = 0 To NumeroLineasProduccion
                Label3.Caption = "Linea: " & I
                pb1.Value = pb1.Value + 1
                If LeerLinea(I) Then
                
                    PonerDatosLinea LineasDeProduccion(I)
                
                Else
                    'PonerDatosLinea LineasDeProduccion(0)
                    LimpiarLinea I
                End If
            Next I
            
            '************************
            ' paletizacion
            For I = 1 To 6
                Label3.Caption = "Lin. palet : " & I
                pb1.Value = pb1.Value + 1
                If LeerLineaPalet(I) Then
                    PonerDatosLineaPalet LineasDePaletizacion(I - 1), True
                Else
                    LimpiarLineaPalet I
                End If
            Next I
            
        End If
        
        'Planning
        
        If PermisoPlanning Then
            Espera 0.5
            CargarPlanning
            
        End If
        
        Set miRsAux = Nothing
        FramePB.visible = False
        Me.Frame1.visible = PermisoProduccion
        Me.Frame3.visible = PermisoProduccion
        Me.Frame5.visible = True
        BloquesSegundosDeRefrescos = 0
        Screen.MousePointer = Rc
End Sub


Private Sub PonerCajasLeidas()
Dim SQL As String
Dim I As Integer
Dim Total As Integer
      
    SQL = ""
    For I = 1 To 6
        If Not LineasDePaletizacion(I - 1) Is Nothing Then
            SQL = "OK"
            Exit For
        End If
    Next I
    If SQL = "" Then Exit Sub 'Ninguna linea paletizando
    
    
    'FramePB.visible = True
    'Me.Frame1.visible = False
    'pb1.Max = 3 'paletizado 3 de paletizacion
    'pb1.Value = 0
    
    For I = 0 To 5
        'pb1.Value = pb1.Value + 1
        If Not LineasDePaletizacion(I) Is Nothing Then
            LineasDePaletizacion(I).CargaCajasPaletList Total, Me.lstCajas(I)
            txtCajas(I).Text = Total
        
        End If
    Next I

    'FramePB.visible = False
    'Me.Frame1.visible = True

End Sub



Private Sub CargarPlanning()
    'Habra que leer en una tabla si hay que recargar pedidos ...
    If True Then
    
        Aux = DevuelveDesdeBD(conAri, "min(fecentre)", "scaped", "1", "1")
        If Aux = "" Then Aux = vEmpresa.FechaIni
        MinAnyo = Year(CDate(Aux))
    
    
        Label3.Caption = "Planning"
        Label3.Refresh
        
        CargaTablaPlanning
        
        Label3.Caption = "Pl. semanal"
        Label3.Refresh
        PonerPlanningListview
        
        Label3.Caption = "Semana"
        Label3.Refresh
        CargarPlanningSemana
        CargaPorSemanaEntrega
        
        Label3.Caption = "Materia auxiliar"
        Label3.Refresh
        CargarPlanningMateriaAuxiliar
    End If
End Sub




Private Sub PreparaListviewPlanning()
    For NumRegElim = 3 To Me.lwp.ColumnHeaders.Count
        Me.lwp.ColumnHeaders(NumRegElim).Alignment = lvwColumnRight
       ' If NumRegElim >= 4 And NumRegElim <= 10 Then Me.lwp.ColumnHeaders(NumRegElim).Text = ""
    Next
    Orden_lwp = 1  'codartic
    Asc_lwp = True 'acenente
    
End Sub


Private Sub CargaTablaPlanning()


    Aux = "DELETE FROM tmpPlanning "
    conn.Execute Aux
    'Cargamos los datos del pedido
    Aux = "INSERT INTO tmpPlanning"
    Aux = Aux & " select sliped.codartic codartic ,sliped.nomartic nomartic"
    Aux = Aux & ",if(pal_udbas is null,0,pal_udbas)*if(pal_udalt is null,0,pal_udalt) cajaspal"
    Aux = Aux & ",unicajas,0 enproduccion,sum(cantidad) ud_pedidos,0.00 stock"
    Aux = Aux & " From scaped,sliped,sartic left join sarti4 on sartic.codartic=sarti4.codartic"
    Aux = Aux & " where scaped.numpedcl=sliped.numpedcl and"
    Aux = Aux & " sliped.codartic=sartic.codartic  and sliped.codartic <> '" & vParamAplic.ArtReciclado & "'"
    Aux = Aux & " and conjunto=1 group by 1,2,3,4"
    conn.Execute Aux
    
    'Actualizamos el stock actual
    Set miRsAux = New ADODB.Recordset
    Aux = "Select canstock,codartic FROM salmac where codalmac =1 AND codartic in (Select codartic from tmpplanning)"
    miRsAux.Open Aux, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not miRsAux.EOF
        Aux = "UPDATE tmpPlanning set stock=" & DBSet(miRsAux!CanStock, "N")
        Aux = Aux & " WHERE codartic = " & DBSet(miRsAux!codartic, "T")
        miRsAux.MoveNext
        conn.Execute Aux
    Wend
    miRsAux.Close
    
    

    
    
    Set miRsAux = Nothing
End Sub


Private Sub PonerPlanningListview()
Dim It As ListItem
Dim Pal_ped As Currency
Dim Pal_stock As Currency




    Set miRsAux = New ADODB.Recordset
    Aux = "Select * from tmpplanning order by "
    Select Case Orden_lwp
    Case 1
        Aux = Aux & " codartic"
    
    Case 2
        Aux = Aux & " nomartic"
    End Select
    If Not Asc_lwp Then Aux = Aux & " DESC"
    
    
    Me.lwp.ListItems.Clear
    
    miRsAux.Open Aux, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not miRsAux.EOF
        If It Is Nothing Then
            'primera linea
            Set It = lwp.ListItems.Add()
            It.Text = " "
            For NumRegElim = 2 To Me.lwp.ColumnHeaders.Count - 1
                It.SubItems(CInt(NumRegElim)) = " "
            Next
            It.SubItems(6) = "Pedidos"
        End If
        
        Set It = lwp.ListItems.Add
        It.Text = miRsAux!codartic
        It.SubItems(1) = miRsAux!NomArtic
        
        'cajaspal,unicajas,enproduccion,
        'estas dos no se ven
        It.SubItems(2) = miRsAux!CajasPal
        If miRsAux!CajasPal = 0 Then It.ForeColor = vbRed
        
        It.SubItems(3) = miRsAux!Unicajas
        
        
        'Totales
        'ud_pedidos ,stock
        If It.ForeColor <> vbRed Then
            Pal_ped = miRsAux!ud_pedidos / miRsAux!Unicajas 'cuantas cajas
            Pal_ped = Round(Pal_ped / miRsAux!CajasPal, 2)
            It.SubItems(4) = Format(Pal_ped, FormatoImporte)
            
            Pal_stock = miRsAux!stock / miRsAux!Unicajas 'cuantas cajas
            Pal_stock = Round(Pal_stock / miRsAux!CajasPal, 2)
            It.SubItems(5) = Format(Pal_stock, FormatoImporte)
            
            'diferencia
            Pal_ped = Pal_ped - Pal_stock
            If Pal_ped > 0 Then
                It.SubItems(6) = Format(Pal_ped, FormatoImporte)
            Else
                It.SubItems(6) = " "
            End If
        Else
            It.SubItems(4) = " "
            It.SubItems(5) = It.SubItems(4)
            It.SubItems(6) = It.SubItems(4)
        End If
            
        miRsAux.MoveNext
    Wend
    miRsAux.Close
End Sub



Private Sub CargarPlanningSemana()
Dim I As Integer
Dim J As Integer
Dim Cantidad As Currency
Dim Colu As ColumnHeader
Dim numPed As Long
Dim Pedido As String 'hay que vincularlos con avab y buscar el campo referncia



    
    
    If Me.lwp.ColumnHeaders.Count > 8 Then
        For J = lwp.ColumnHeaders.Count To 8 Step -1
            lwp.ColumnHeaders.Remove J
        Next J
    End If
            
    
    Set miRsAux = New ADODB.Recordset
    
    'Vamos p'alla
    
    Aux = "select ((year(fecentre)-" & MinAnyo & ") *100) + sementre  as sementre,scaped.numpedcl,sliped.codartic,cantidad,codclien,referenc,refproduccion from  scaped,sliped ,sartic"
    Aux = Aux & " Where scaped.numpedcl = sliped.numpedcl And sliped.codArtic = sartic.codArtic"
    Aux = Aux & " And Conjunto = 1  and sliped.codartic <> '" & vParamAplic.ArtReciclado & "' order by sementre,numpedcl,numlinea "
    miRsAux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    numPed = 0
    While Not miRsAux.EOF
        
        
        
        'el primer
        If numPed <> miRsAux!numpedcl Then
            'Si el cliente NO es avab
            If miRsAux!CodClien <> 1 Then
                Pedido = miRsAux!numpedcl
            Else
                'AVAB, tenemos que ver si tiene refprov
                Aux = DBLet(miRsAux!refproduccion, "N")
                If Val(Aux) = 0 Then
                    'No estaba vinculado
                    Pedido = miRsAux!numpedcl & "(1)"
                Else
                    Aux = DevuelveDesdeBD(conAri, "referenc", "ariges" & EmprAVAB & ".scaped", "numpedcl", Aux)
                    If Aux = "" Then
                        Pedido = miRsAux!numpedcl & "(A)"
                    Else
                        Pedido = Aux
                    End If
                End If
            End If
            numPed = miRsAux!numpedcl
            
            Set Colu = lwp.ColumnHeaders.Add()
        
            'encabezado con la semna
            Aux = ""
            J = miRsAux!sementre
            If J > 100 Then
                Aux = CStr(J)
                Aux = Mid(Aux, 1, Len(Aux) - 2)
                
                Aux = (Val(Aux) + MinAnyo) Mod 100 & "/"
                
                J = J Mod 100
               
            End If
            Aux = Aux & CStr(J)
            Colu.Text = Aux
            Colu.Width = 900
            Colu.Tag = miRsAux!numpedcl
            Me.lwp.ListItems(1).SubItems(Colu.Index - 1) = Pedido
        End If
        
        
        
        
        
        'BUsco la referencia
        For J = 1 To lwp.ListItems.Count
            If Me.lwp.ListItems(J).Text = miRsAux!codartic Then
                'Este es
                Exit For
            End If
        Next J
        
        If J > lwp.ListItems.Count Then
            'NO LO HA ENCINTRADO
            Stop
        Else
            'Ya lo tengo
            If lwp.ListItems(J).ForeColor = vbRed Then
                'NADA. Mal las caja
                
            Else
                Cantidad = lwp.ListItems(J).SubItems(3) 'unicajas
                If lwp.ListItems(J).SubItems(2) = "" Then
                    Cantidad = 0
                Else
                    Cantidad = Cantidad * lwp.ListItems(J).SubItems(2) 'uds palet
                    Cantidad = Round(miRsAux!Cantidad / Cantidad, 2)
                End If
                
                'Puede haber dos veces en un pedido un mismo articulo
                If Trim(lwp.ListItems(J).SubItems(Colu.Index - 1)) = "" Then
                    lwp.ListItems(J).SubItems(Colu.Index - 1) = Cantidad
                Else
                    'YA HABIA UN VALOR
                     lwp.ListItems(J).SubItems(Colu.Index - 1) = CCur(lwp.ListItems(J).SubItems(Colu.Index - 1)) + Cantidad
                End If
            End If
            
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
End Sub




Private Sub CargaPorSemanaEntrega()
Dim It As ListItem
Dim Cantidad As Currency
Dim stoc As Currency
Dim Ped As Currency
Dim J As Integer

    Me.lwp2(0).ListItems.Clear
    Aux = "select ((year(fecentre)-" & MinAnyo & ") *100) + sementre  as sementre,scaped.numpedcl,sliped.codartic,cantidad,cajaspal,"
    Aux = Aux & "tmpplanning.* "
    Aux = Aux & " from  scaped,sliped ,tmpplanning"
    Aux = Aux & " Where scaped.numpedcl = sliped.numpedcl And sliped.codArtic = tmpplanning.codArtic"
    Aux = Aux & "  order by sementre,numpedcl,numlinea  "

    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
    
      
    
        'Metemos el ITEM
        Set It = Me.lwp2(0).ListItems.Add
        It.Text = miRsAux!codartic
        It.SubItems(1) = miRsAux!NomArtic

        It.SubItems(2) = miRsAux!numpedcl
        'Antes Dic2018
        'It.SubItems(3) = miRsAux!sementre
        Aux = ""
        J = miRsAux!sementre
        If J > 100 Then
            Aux = CStr(J)
            Aux = Mid(Aux, 1, Len(Aux) - 2)
            
            Aux = (Val(Aux) + MinAnyo) Mod 100 & "/"
            
            J = J Mod 100
           
        End If
        Aux = Aux & CStr(J)
        It.SubItems(3) = Aux


        Cantidad = miRsAux!CajasPal
        Cantidad = miRsAux!Unicajas * Cantidad

        If Cantidad = 0 Then
            'ERROR
            It.SubItems(4) = "N/D"
            It.SubItems(5) = "N/D"
            It.SubItems(6) = " "
            It.SubItems(7) = " "
            It.SubItems(8) = 0 'uds por palet
            
        Else
            'Pedido
            It.SubItems(8) = Cantidad
            Ped = Round2(miRsAux!Cantidad / Cantidad, 2) 'cjas palet * uds
            
            It.SubItems(4) = Format(Ped, FormatoImporte)
            'stock
            If Not VerStockArrastrado(It, stoc) Then
                stoc = miRsAux!stock
                stoc = Round2(stoc / Cantidad, 2) 'palets en stock
            End If
            It.SubItems(5) = Format(stoc, FormatoImporte)
            
            stoc = stoc - Ped
            It.SubItems(7) = stoc 'columna oculta
            It.SubItems(6) = " "
            If stoc < 0 Then
                It.SubItems(6) = Format(Abs(stoc), FormatoImporte)
               
            End If
        End If
        
        miRsAux.MoveNext
        
    Wend
    miRsAux.Close
    Set miRsAux = Nothing

End Sub

'Para cada articulo veremos si el stcok ha sido utlizado anteriormente
' para ello veremos desde la linea que acabamos de insertar hacia arriba si encontramos
'la cantidad disponible
Private Function VerStockArrastrado(ByRef It As ListItem, ByRef KCantidas As Currency) As Boolean
Dim J As Integer

    VerStockArrastrado = False
    For J = Me.lwp2(0).ListItems.Count - 1 To 1 Step -1
        If Me.lwp2(0).ListItems(J).Text = It.Text Then
            'Vemos si tiene "disponible"
            
            If Trim(Me.lwp2(0).ListItems(J).SubItems(5)) = "" Then
                'Significa que NO nos queda stock.
                KCantidas = 0
            Else
                'Si queda.)
                KCantidas = CCur(lwp2(0).ListItems(J).SubItems(7))  'columna oculta con lo que queda
            End If
            VerStockArrastrado = True
            Exit For
        End If
    Next J
    
End Function

''''''''''''''''''''''''''''''''''''''''''''''''
''''
''''        ANTIGUO
''''
''''Private Sub CargarPlanningMateriaAuxiliar()
''''Dim I As Integer
''''Dim J As Integer
''''Dim Cantidad As Currency
''''Dim C1 As Currency
''''Dim It As ListItem
''''Dim Colu As ColumnHeader
''''Dim RN As ADODB.Recordset
''''Dim ErrMateriasPrimas As String
''''Dim CantidadSegunPlanning As Currency
''''Dim Semana As Byte
''''
''''    Me.lwp2(1).ListItems.Clear
''''
''''    J = lwp2(1).ColumnHeaders.Count
''''    While J > 4
''''        lwp2(1).ColumnHeaders.Remove lwp2(1).ColumnHeaders.Count
''''        J = lwp2(1).ColumnHeaders.Count
''''    Wend
''''    Set miRsAux = New ADODB.Recordset
''''
''''    'Vamos p'alla
''''    'Cargamos las materias primas, lo que hay de stock
''''    Aux = "select sartic.codartic,nomartic,canstock from sarti1,sartic,salmac where"
''''    Aux = Aux & " sarti1.codarti1=sartic.codartic and sartic.codartic=salmac.codartic"
''''    Aux = Aux & " and codalmac=1 and factorconversion=1"
''''    Aux = Aux & " and sarti1.codartic in (select codartic from tmpplanning) group by 1,2 order by 2"
''''    miRsAux.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
''''    While Not miRsAux.EOF
''''        Set It = Me.lwp2(1).ListItems.Add
''''        It.Text = miRsAux!codArtic
''''        It.SubItems(1) = miRsAux!NomArtic
''''        Cantidad = miRsAux!CanStock
''''        If Cantidad < 0 Then
''''            Cantidad = 0
''''            It.SubItems(2) = " "
''''            It.ForeColor = vbRed
''''
''''        Else
''''            It.SubItems(2) = Format(Cantidad, FormatoCantidad)
''''        End If
''''        It.SubItems(3) = Cantidad 'stock en numero, columna oculta
''''
''''
''''        miRsAux.MoveNext
''''    Wend
''''    miRsAux.Close
''''
''''
''''
''''
''''    'Ya tengo cargadas todas la materias primas que intervienen en los pedidos
''''    'AHora , por semana ire creando columnas
''''    'Empezare en la columna 4
''''
''''    Aux = "select sementre from scaped group by 1 order by 1"
''''    miRsAux.Open Aux, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
''''    NumRegElim = 4 'empieza en la 4
''''    While Not miRsAux.EOF
''''        NumRegElim = NumRegElim + 1
''''        Set Colu = Me.lwp2(1).ColumnHeaders.Add
''''        Colu.Alignment = lvwColumnRight
''''        Colu.Text = "Sem" & miRsAux!sementre
''''        Colu.Width = 1000
''''
''''        For J = 1 To Me.lwp2(1).ListItems.Count
''''            lwp2(1).ListItems(J).SubItems(NumRegElim - 1) = " "
''''        Next
''''        miRsAux.MoveNext
''''    Wend
''''    miRsAux.Close
''''
''''    Set RN = New ADODB.Recordset
''''    ErrMateriasPrimas = ""
''''    I = 0 'para que meta las referncias con errores de 5 en 5
''''    NumRegElim = 4
''''    'While NumRegElim <= Me.lwp2(1).ColumnHeaders.Count - 1
''''    Do   'el primero se hace
''''        NumRegElim = NumRegElim + 1
''''        Aux = Mid(lwp2(1).ColumnHeaders(NumRegElim).Text, 4)  'SEM39-->39
''''        Semana = CByte(Aux)
''''        Aux = " and sementre=" & Aux & " and conjunto=1  and sliped.codartic<>'000000000000'  group by 1"
''''        Aux = " scaped.numpedcl = sliped.numpedcl And sliped.codArtic = sartic.codArtic " & Aux
''''        Aux = "select sliped.codartic,sum(cantidad) totalpedido from scaped,sliped,sartic where" & Aux
''''        miRsAux.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
''''
''''
''''
''''        While Not miRsAux.EOF
''''            'Cada producto, con la cantidad a producir
''''            If DBLet(miRsAux!totalpedido, "N") <> 0 Then
''''                Aux = "select codarti1,cantidad from sarti1,sartic where sarti1.codarti1=sartic.codartic AND "
''''                Aux = Aux & " sarti1.codartic=" & DBSet(miRsAux!codArtic, "T") & " AND factorconversion=1"
''''
''''                RN.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
''''                While Not RN.EOF
''''
''''                    For J = 1 To Me.lwp2(1).ListItems.Count
''''                               'es este
''''                        'If RN!codarti1 = "003300290602" Then Stop
''''                        If Me.lwp2(1).ListItems(J).Text = RN!codArti1 Then Exit For
''''                    Next
''''                    If J <= Me.lwp2(1).ListItems.Count Then
''''                        'OK, el NODO j es el articulo que buscabamos
''''                        'If RN!codArti1 = "000300160120" Then Stop
''''                        Cantidad = DBLet(RN!Cantidad, "N") 'cantidad en el escandallo
''''                        'Cantidad = Cantidad * miRsAux!totalpedido  'suma en pedidos por semana /artc
''''
''''                        CantidadSegunPlanning = CantidadAProducirSabiendoStock(Semana, miRsAux!codArtic)
''''                        'If CantidadSegunPlanning <= 0 Then Stop
''''                        Cantidad = Cantidad * CantidadSegunPlanning
''''
''''                        'ahora tengo en cantidad cuanto necesito para esta semana
''''                        Cantidad = lwp2(1).ListItems(J).SubItems(3) - Cantidad 'cuanto me queda
''''                        lwp2(1).ListItems(J).SubItems(3) = Cantidad
''''                        lwp2(1).ListItems(J).SubItems(NumRegElim - 1) = Format(Cantidad, FormatoCantidad)
''''                        If Cantidad < 0 Then lwp2(1).ListItems(J).ListSubItems(NumRegElim - 1).ForeColor = vbRed
''''                    Else
''''                        If InStr(1, ErrMateriasPrimas, RN!codArti1) = 0 Then
''''
''''                            I = I + 1
''''                            ErrMateriasPrimas = ErrMateriasPrimas & RN!codArti1 & "  "
''''                            If I = 5 Then
''''                                ErrMateriasPrimas = ErrMateriasPrimas & vbCrLf
''''                                I = 0
''''                            End If
''''                        End If
''''                    End If
''''                    RN.MoveNext
''''                Wend
''''                RN.Close
''''            End If
''''            miRsAux.MoveNext
''''        Wend
''''        miRsAux.Close
''''
''''    Loop While NumRegElim < Me.lwp2(1).ColumnHeaders.Count
''''
''''
''''    If ErrMateriasPrimas <> "" Then
''''        ErrMateriasPrimas = "Error materias primas" & vbCrLf & ErrMateriasPrimas
''''        MsgBox ErrMateriasPrimas, vbExclamation
''''    End If
''''    Set RN = Nothing
''''    Set miRsAux = Nothing
''''End Sub
''''


Private Sub CargarPlanningMateriaAuxiliar()
Dim I As Integer
Dim J As Integer
Dim Cantidad As Currency
Dim C1 As Currency
Dim It As ListItem
Dim Colu As ColumnHeader
Dim RN As ADODB.Recordset
Dim ErrMateriasPrimas As String
Dim CantidadSegunPlanning As Currency
Dim Semana2 As Integer
Dim AuxD As String
Dim K As Integer
Dim NumeroAnyoPedido As Integer 'Para los pedidos con fecha de entrega posteriores



    Me.lwp2(1).ListItems.Clear
    
    J = lwp2(1).ColumnHeaders.Count
    While J > 5
        lwp2(1).ColumnHeaders.Remove lwp2(1).ColumnHeaders.Count
        J = lwp2(1).ColumnHeaders.Count
    Wend
    Set miRsAux = New ADODB.Recordset
    
    'Vamos p'alla
    'Cargamos las materias primas, lo que hay de stock
    Aux = "select sartic.codartic,nomartic,canstock from sarti1,sartic,salmac where"
    Aux = Aux & " sarti1.codarti1=sartic.codartic and sartic.codartic=salmac.codartic"
    Aux = Aux & " and codalmac=1 and factorconversion=1"
    Aux = Aux & " and sarti1.codartic in (select codartic from tmpplanning) group by 1,2 order by 2"
    miRsAux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set It = Me.lwp2(1).ListItems.Add
        It.Text = miRsAux!codartic
        It.SubItems(1) = miRsAux!NomArtic
        Cantidad = miRsAux!CanStock
        If Cantidad < 0 Then
            Cantidad = 0
            It.SubItems(2) = " "
            It.ForeColor = vbRed
            
        Else
            It.SubItems(2) = Format(Cantidad, FormatoCantidad)
        End If
        It.SubItems(3) = Cantidad 'stock en numero, columna oculta
        
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    
    'Ya tengo cargadas todas la materias primas que intervienen en los pedidos
    'AHora , por semana ire creando columnas
    'Empezare en la columna 4
    
    Aux = "select sementre from scaped group by 1 order by 1"
    'Dic2018
    Aux = "select ((year(fecentre)-" & MinAnyo & ") *100) + sementre  as sementre from scaped group by 1 order by 1"
    
    miRsAux.Open Aux, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    NumRegElim = 5 'empieza en la 5
    While Not miRsAux.EOF
        NumRegElim = NumRegElim + 1
        Set Colu = Me.lwp2(1).ColumnHeaders.Add
        Colu.Alignment = lvwColumnRight
        
        
        'encabezado con la semna
        AuxD = ""
        K = miRsAux!sementre
        If K > 100 Then
            AuxD = CStr(K)
            AuxD = Mid(AuxD, 1, Len(AuxD) - 2)
            
            AuxD = (Val(AuxD) + MinAnyo) Mod 100 & "/"
            
            K = K Mod 100
           
        End If
        AuxD = AuxD & CStr(K)
        
        
        Colu.Text = "Sem" & AuxD
        Colu.Width = 1000
        
        For J = 1 To Me.lwp2(1).ListItems.Count
            lwp2(1).ListItems(J).SubItems(NumRegElim - 1) = " "
        Next
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    Set RN = New ADODB.Recordset
    ErrMateriasPrimas = ""
    I = 0 'para que meta las referncias con errores de 5 en 5
    NumRegElim = 5
    
    Do   'el primero se hace
        NumRegElim = NumRegElim + 1
        NumeroAnyoPedido = 0
        Aux = Mid(lwp2(1).ColumnHeaders(NumRegElim).Text, 4)  'SEM39-->39
        If InStr(1, Aux, "/") > 0 Then
            NumeroAnyoPedido = Mid(Aux, 1, InStr(1, Aux, "/") - 1)
            Aux = Mid(Aux, InStr(1, Aux, "/") + 1)
        End If
        Semana2 = CByte(Aux)
        Aux = " and sementre=" & Aux & " and conjunto=1  and sliped.codartic<>'000000000000'"
        Aux = " scaped.numpedcl = sliped.numpedcl And sliped.codArtic = sartic.codArtic " & Aux
         Aux = Aux & " AND year(fecentre)="
        If NumeroAnyoPedido > 0 Then
            Aux = Aux & CStr(2000 + NumeroAnyoPedido)
        Else
            Aux = Aux & CStr(MinAnyo)
        End If
        'No miro la cantidad. VOY a ver la referencia y buscare
        ' el lwiew de la semana,y producto que valor hay que producir
        Aux = "select sliped.codartic from scaped,sliped,sartic where" & Aux & "   group by 1"
        miRsAux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        
        
        While Not miRsAux.EOF
            'Cada producto, con la cantidad a producir
            
                'If Right(miRsAux!codartic, 3) = "103" Then Stop
                
                
                CantidadSegunPlanning = CantidadAProducirSabiendoStock(Semana2, miRsAux!codartic, NumeroAnyoPedido)
                
                Aux = "select codarti1,cantidad from sarti1,sartic where sarti1.codarti1=sartic.codartic AND "
                Aux = Aux & " sarti1.codartic=" & DBSet(miRsAux!codartic, "T") & " AND factorconversion=1"
                RN.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not RN.EOF
                    
                    For J = 1 To Me.lwp2(1).ListItems.Count
                               'es este
                        'If RN!codarti1 = "003300290602" Then Stop
                        If Me.lwp2(1).ListItems(J).Text = RN!codarti1 Then Exit For
                    Next
                  
                    If J <= Me.lwp2(1).ListItems.Count Then
                    
                         'OK, el NODO j es el articulo que buscabamos
                        'If RN!codArti1 = "000300160120" Then Stop
                        Cantidad = DBLet(RN!Cantidad, "N") 'cantidad en el escandallo
                      
                        
                        'If CantidadSegunPlanning <= 0 Then Stop
                        Cantidad = Cantidad * CantidadSegunPlanning
                        
                        'ahora tengo en cantidad cuanto necesito para esta semana
                        Cantidad = lwp2(1).ListItems(J).SubItems(3) - Cantidad 'cuanto me queda
                        lwp2(1).ListItems(J).SubItems(3) = Cantidad
                        lwp2(1).ListItems(J).SubItems(NumRegElim - 1) = Format(Cantidad, FormatoCantidad)
                        If Cantidad < 0 Then lwp2(1).ListItems(J).ListSubItems(NumRegElim - 1).ForeColor = vbRed
                    Else
                        If InStr(1, ErrMateriasPrimas, RN!codarti1) = 0 Then
                    
                            I = I + 1
                            ErrMateriasPrimas = ErrMateriasPrimas & RN!codarti1 & "  "
                            If I = 5 Then
                                ErrMateriasPrimas = ErrMateriasPrimas & vbCrLf
                                I = 0
                            End If
                        End If
                    End If
                    RN.MoveNext
                Wend
                RN.Close
       
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
    Loop While NumRegElim < Me.lwp2(1).ColumnHeaders.Count
    
    
    
    'Ahora cargaremos los pedidos a proveedor
    Aux = ""
    I = 0 'para quelos selects de 10 en 10 artic
    NumRegElim = 0
    J = 0
    Do
        NumRegElim = NumRegElim + 1
        I = I + 1
        
        Aux = Aux & ", '" & Me.lwp2(1).ListItems(NumRegElim).Text & "'"
        
        If NumRegElim = Me.lwp2(1).ListItems.Count Then
            J = 1 'Que haga el sql
        Else
            If I > 15 Then J = 1
        End If
        
        
        If J = 1 Then
            'OK vamos a cargar los pedidos de estos articulos
            Aux = Mid(Aux, 2)
            
            Aux = " WHERE codartic in (" & Aux & ")"
            Aux = "select slippr.codartic,sum(cantidad) from slippr  " & Aux
            Aux = Aux & " group by 1"
            miRsAux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
        
                For J = 1 To Me.lwp2(1).ListItems.Count
                    If lwp2(1).ListItems(J).Text = miRsAux!codartic Then
                        
                        lwp2(1).ListItems(J).SubItems(4) = Format(miRsAux.Fields(1), FormatoCantidad)
                        'lwp2(1).ListItems(J).ListSubItems(4).Bold = True
                        lwp2(1).ListItems(J).ListSubItems(4).ForeColor = vbBlue
                    End If
                Next
        
        
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            'Para la proxima
            J = 0
            I = 0
            Aux = ""
        End If
        
    Loop Until NumRegElim = Me.lwp2(1).ListItems.Count
    
    
    
    
    If ErrMateriasPrimas <> "" Then
        ErrMateriasPrimas = "Error materias primas" & vbCrLf & ErrMateriasPrimas
        MsgBox ErrMateriasPrimas, vbExclamation
    End If
    Set RN = Nothing
    Set miRsAux = Nothing
End Sub











Private Sub ImprimirPedido(columna As Integer)

    With frmImprimir
        'indRPT = 7 '7: Pedidos de Clientes
        Aux = DevuelveDesdeBD(conAri, "documrpt", "scryst", "codcryst", "7")
        .NombreRPT = Aux
        .Titulo = "Pedido de Ventas desde produccion"
        .FormulaSeleccion = "{scaped.numpedcl}=" & Me.lwp.ColumnHeaders(columna + 1).Tag
        .OtrosParametros = "|pCodigoISO=""1""|pCodigoRev=""02""|pLinea1=""Desde produccion""|SinValorar= 1|pTipoIVA=0|"
        .NumeroParametros = 4
        .Show vbModal
    End With
End Sub


Private Function CantidadAProducirSabiendoStock(Semana2 As Integer, QueArticulo As String, AnyoSiguiente As Integer) As Currency
Dim I As Integer
Dim Esta As Boolean
Dim SemanaCalculada As Integer
Dim SemanaListview As Integer


        CantidadAProducirSabiendoStock = 0
        Esta = False
        
        
        SemanaCalculada = Semana2
        If AnyoSiguiente > 0 Then SemanaCalculada = SemanaCalculada + 53
    
        For I = 1 To Me.lwp2(0).ListItems.Count
        
          
            If Not IsNumeric(Me.lwp2(0).ListItems(I).SubItems(3)) Then
                
                SemanaListview = 53 + Mid(Me.lwp2(0).ListItems(I).SubItems(3), 4)
            Else
                SemanaListview = Val(Me.lwp2(0).ListItems(I).SubItems(3))
            End If
           
                
        
            'Misma semana
            If SemanaListview > SemanaCalculada Then
                'Ya salimos
                Exit For
            Else
                If SemanaListview = SemanaCalculada Then
                    If Me.lwp2(0).ListItems(I).Text = QueArticulo Then
                        'Este es el articulo
                        Esta = True
                        If Trim(lwp2(0).ListItems(I).SubItems(6)) = "" Then
                            'Significa que tengo bastante
                            
                        Else
                                                            
                            'Voy a producir, pero cuanto....
                            'Si el stock es negativo significa que
                            'Ya teneia algo pendiente de producir, con lo cual, para esta semana
                            'la produccuion es  la de pendiente de producir de la SEMANA, no la PDTE

                            If ImporteFormateado(lwp2(0).ListItems(I).SubItems(5)) <= 0 Then
                                CantidadAProducirSabiendoStock = CantidadAProducirSabiendoStock + Abs((lwp2(0).ListItems(I).SubItems(4) * lwp2(0).ListItems(I).SubItems(8)))
                            Else
                                
                                CantidadAProducirSabiendoStock = Abs((lwp2(0).ListItems(I).SubItems(7) * lwp2(0).ListItems(I).SubItems(8)))
                            End If
                        End If
                    End If
                End If
            End If
        Next I
    
        
        If Not Esta Then
            MsgBox "No encontrado. Semana " & Semana2 & " Art. " & QueArticulo, vbExclamation
        Else
            CantidadAProducirSabiendoStock = Round(CantidadAProducirSabiendoStock, 0)
        End If
        
End Function

Private Sub ComprobarPoste(Inicio As Boolean)
Dim Actu As String

    If Not Inicio Then
        ParalecturaPoste = ParalecturaPoste + 1
        If ParalecturaPoste < 3 Then Exit Sub
    End If
    ParalecturaPoste = 0
    Actu = txtIdP(0).Text & Me.txtIdP(1).Text & Me.txtIdP(2).Text & Me.txtIdP(3).Text & Me.txtIdP(4).Text & Me.txtIdP(5).Text
    If Actu <> "" Then
        Actu = "now()"
        Aux = DevuelveDesdeBD(conAri, "max(fechahora)", "prodlecturaposte", "1", "1", "N", Actu)
        If Aux <> "" Then
            NumRegElim = DateDiff("n", Aux, Actu)
            If NumRegElim > 5 Then
                'rojo
                Me.imgPoste(0).Picture = Me.Image1(2).Picture
                
            ElseIf NumRegElim > 2 Then
                'azul
                Me.imgPoste(0).Picture = Me.Image1(2).Picture
            Else
                'verde
                Me.imgPoste(0).Picture = Me.Image1(1).Picture
            End If
        Else
            Me.imgPoste(0).Picture = Me.Image1(2).Picture
        End If
    Else
        Me.imgPoste(0).Picture = LoadPicture()
    End If
    
    ComprobarEntradaDuplicadas
    
End Sub



Private Function ComprobarLineaAHORA(Lin As Byte, Nuevo As Boolean) As Boolean

    Timer1.Enabled = False
    ComprobarLineaAHORA = False
    Set AuxLP = Nothing
    If LeerLineaAux(Lin) Then
        If Nuevo Then
            If Not AuxLP Is Nothing Then
                MsgBox "Error leyendo datos", vbExclamation
            Else
                ComprobarLineaAHORA = True
            End If
        Else
            'Esta modificando algo
            If AuxLP Is Nothing Then
               MsgBox "Error. Nada en produccion", vbExclamation
            Else
                If AuxLP.LoteTrazabilidad <> AuxLP.LoteTrazabilidad Then
                    MsgBox "Error. Ha cambiado lote linea", vbExclamation
                Else
                    ComprobarLineaAHORA = True
                End If
            End If
        End If
    
    End If
    
    If Not ComprobarLineaAHORA Then
        cmdPlanning_Click
    Else
        Timer1.Enabled = True
    End If
End Function


'Antes de cualquier cambio veremos si la linea HA cambiado
'es decir, que si no es la misma que refresque
Private Function LeerLineaAux(KLinea As Byte) As Boolean
Dim I As Byte
Dim SQL As String

    LeerLineaAux = False
    Set miRsAux = New ADODB.Recordset
    SQL = "select prodlin.codigo,prodlin.idlin ,lotetraza from prodlin,prodtrazlin  where prodlin.codigo= prodtrazlin.codigo AND prodlin.idlin = prodtrazlin.idlin"
    SQL = SQL & " and lineaprod = " & KLinea & " and estado >0 and estado<10 ORDER BY lotetraza DESC"  'Pq puede que haya varios cambios de trazabilidad para la misma linea
    miRsAux.Open SQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    I = 0
    If Not miRsAux.EOF Then
        'UNO tiene seguro
        
        Set AuxLP = New cLineaProduccion
        AuxLP.LeerDesdeTrazabilidad miRsAux!Codigo, miRsAux!idlin, KLinea, miRsAux!lotetraza
        I = 1
        
        'veremos que SOLO hay una linea en marcha
        Do
            miRsAux.MoveNext
        
            If Not miRsAux.EOF Then
                'Si el codigo y prodilin es el mismo, es que solo hay una produccion
                If miRsAux!Codigo <> AuxLP.CodProduccion Or AuxLP.idLiProd <> miRsAux!idlin Then
                    
                    
                    I = 2
                End If
            End If
        Loop Until miRsAux.EOF
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    If I < 2 Then LeerLineaAux = True
End Function


Private Sub PoneLaLineaIndicadora(Cual As Integer)
    OcultarLineaIndicadora
    LineaIndicadora(Cual).visible = True
    
End Sub

Private Sub OcultarLineaIndicadora()
Dim I As Byte
    For I = 0 To LineaIndicadora.Count - 1
        LineaIndicadora(I).visible = False
    Next I
    
End Sub
Private Sub OcultarLineaIndicadoraPalet()
Dim I As Byte
 
    For I = 0 To LineaIndicadoraP.Count - 1
        LineaIndicadoraP(I).visible = False
    Next I
    
End Sub



Private Sub ComprobarEntradaDuplicadas()
    Aux = DevuelveDesdeBD(conAri, "count(*)", "prodcajasduplicadas", "1", "1")
    If Aux = "" Then Aux = "0"
    
    If Val(Aux) >= 1 Then
        'ENTRADAS DUPLICADAS
        Aux = "*******   *****"
        Aux = Aux & "  D U P L I C A D O S   " & Aux
        Me.Caption = Aux
        'Frame del poste
        Me.Frame5.Caption = ""
        Me.Frame5.BackColor = vbRed
        Frame5.BorderStyle = 0
        
        VerLineasEntradasDuplicadas
        
        
    Else
        'RESTAURAMOS
        If Me.Frame5.Caption = "" Then
            'Significa que estaba mal
            Me.Caption = "Producción / Envasado"
            Me.Frame5.Caption = "Poste"
            Me.Frame5.BackColor = -2147483633
            Frame5.BorderStyle = 1
        
            
            Me.Refresh
        End If
    End If
    
End Sub


Private Sub VerLineasEntradasDuplicadas()

    Aux = "Select lotetraza  from prodcajasduplicadas GROUP BY 1"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        For NumRegElim = 0 To NumeroLineasProduccion - 1
            If Val(txtTraza(NumRegElim).Text) = Val(miRsAux!lotetraza) Then
                'ESTA ESTA MAL
                imgLinea(NumRegElim).Picture = Me.Image1(2).Picture
                Me.FrameLinea1(NumRegElim).BackColor = vbRed
            End If
        Next
        miRsAux.MoveNext
    Wend
End Sub



Private Sub VerDuplicidades()
    If Me.Frame5.BackColor <> vbRed Then Exit Sub

    'Octubre 2012
    If Not PermisoProduccion Then Exit Sub

    Timer1.Enabled = False
    
    CadenaDesdeOtroForm = ""
    frmProdDuplicados.Show vbModal
    
    If CadenaDesdeOtroForm <> "" Then
        'Reestablecer
        For NumRegElim = 0 To NumeroLineasProduccion - 1
            Me.FrameLinea1(NumRegElim).BackColor = -2147483633
            
        Next
        'Cargamos todos
        seg = RefrescarTodosLosDatos + 1
        BloquesSegundosDeRefrescos = RefrescarTodosLosDatos + 1
        ParalecturaPoste = 3
    End If
    Timer1.Enabled = True
End Sub
