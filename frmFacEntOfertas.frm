VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmFacEntOfertas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ofertas Clientes"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   4335
   ClientWidth     =   13185
   Icon            =   "frmFacEntOfertas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   13185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   920
      Left            =   120
      TabIndex        =   81
      Top             =   390
      Width           =   12855
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   3
         Left            =   7965
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   86
         Text            =   "Text2"
         Top             =   160
         Width           =   3345
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   7200
         MaxLength       =   4
         TabIndex        =   4
         Tag             =   "Realizada Por|N|N|0|9999|scapre|codtraba|0000|N|"
         Text            =   "Text1"
         Top             =   160
         Width           =   720
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   5
         Left            =   7965
         MaxLength       =   40
         TabIndex        =   12
         Tag             =   "Nombre Cliente|T|N|||scapre|nomclien||N|"
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwww aq"
         Top             =   520
         Width           =   3345
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   7200
         MaxLength       =   6
         TabIndex        =   5
         Tag             =   "Cod. Cliente|N|N|0|999999|scapre|codclien|000000|N|"
         Text            =   "Text1"
         Top             =   520
         Width           =   720
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   1220
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Oferta|F|N|||scapre|fecofert|dd/mm/yyyy|N|"
         Top             =   430
         Width           =   1065
      End
      Begin VB.CheckBox chkAceptado 
         Caption         =   "Aceptada"
         Height          =   255
         Index           =   0
         Left            =   4440
         TabIndex        =   3
         Tag             =   "Aceptada|N|N|||scapre|aceptado||N|"
         Top             =   405
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   0
         Left            =   200
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "N� Oferta|N|S|0||scapre|numofert|0000000|S|"
         Text            =   "Text1 7"
         Top             =   430
         Width           =   885
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha Entrega|F|N|||scapre|fecentre|dd/mm/yyyy|N|"
         Top             =   430
         Width           =   1065
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   6900
         Picture         =   "frmFacEntOfertas.frx":000C
         ToolTipText     =   "Buscar trabajador"
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Realiz. por"
         Height          =   255
         Index           =   21
         Left            =   6105
         TabIndex        =   87
         Top             =   165
         Width           =   855
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   6900
         Picture         =   "frmFacEntOfertas.frx":010E
         ToolTipText     =   "Buscar cliente"
         Top             =   570
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         Height          =   255
         Index           =   0
         Left            =   6105
         TabIndex        =   85
         Top             =   525
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "F. Oferta"
         Height          =   255
         Index           =   14
         Left            =   1235
         TabIndex        =   84
         Top             =   240
         Width           =   615
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   2060
         Picture         =   "frmFacEntOfertas.frx":0210
         ToolTipText     =   "Buscar fecha"
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "N� Oferta"
         Height          =   255
         Index           =   50
         Left            =   200
         TabIndex        =   83
         Top             =   240
         Width           =   735
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   3275
         Picture         =   "frmFacEntOfertas.frx":029B
         ToolTipText     =   "Buscar fecha"
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "F. Entrega"
         Height          =   255
         Index           =   51
         Left            =   2450
         TabIndex        =   82
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   16
      Left            =   2400
      Locked          =   -1  'True
      MaxLength       =   60
      TabIndex        =   57
      Text            =   "Text2 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwqa"
      Top             =   6780
      Visible         =   0   'False
      Width           =   6885
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   120
      TabIndex        =   39
      Top             =   6615
      Width           =   2175
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   240
         TabIndex        =   40
         Top             =   180
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   11850
      TabIndex        =   37
      Top             =   6720
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   10680
      TabIndex        =   36
      Top             =   6720
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   360
      Top             =   6360
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   1800
      Top             =   6480
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5280
      Left            =   120
      TabIndex        =   43
      Tag             =   "Fecha Oferta|F|N|||scapre|fecentre|dd/mm/yyyy|N|"
      Top             =   1320
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   9313
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos b�sicos"
      TabPicture(0)   =   "frmFacEntOfertas.frx":0326
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "DataGrid1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtAux(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtAux(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtAux(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtAux(4)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtAux(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtAux(7)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtAux(8)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtAux(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdAux(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdAux(1)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "FrameCliente"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtAux(5)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtAux(9)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtAux(10)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "Textos de la Carta"
      TabPicture(1)   =   "frmFacEntOfertas.frx":0342
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(3)"
      Tab(1).Control(1)=   "Label1(5)"
      Tab(1).Control(2)=   "Label1(45)"
      Tab(1).Control(3)=   "Label1(2)"
      Tab(1).Control(4)=   "Text1(21)"
      Tab(1).Control(5)=   "Text1(22)"
      Tab(1).Control(6)=   "Text1(23)"
      Tab(1).Control(7)=   "Text1(24)"
      Tab(1).Control(8)=   "Text1(25)"
      Tab(1).Control(9)=   "Text1(26)"
      Tab(1).Control(10)=   "Text1(27)"
      Tab(1).Control(11)=   "Text1(28)"
      Tab(1).Control(12)=   "Text1(29)"
      Tab(1).Control(13)=   "Text1(30)"
      Tab(1).Control(14)=   "Text1(18)"
      Tab(1).Control(15)=   "Text1(20)"
      Tab(1).Control(16)=   "Text1(19)"
      Tab(1).Control(17)=   "Text1(33)"
      Tab(1).ControlCount=   18
      TabCaption(2)   =   "Concepto y Gesti�n Oferta"
      TabPicture(2)   =   "frmFacEntOfertas.frx":035E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1(38)"
      Tab(2).Control(1)=   "Label1(37)"
      Tab(2).Control(2)=   "Text1(31)"
      Tab(2).Control(3)=   "Text1(32)"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Totales"
      TabPicture(3)   =   "frmFacEntOfertas.frx":037A
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "FrameFactura"
      Tab(3).ControlCount=   1
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   10
         Left            =   6960
         MaxLength       =   12
         TabIndex        =   52
         Tag             =   "cajas"
         Text            =   "cajas"
         Top             =   4080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   9
         Left            =   6000
         MaxLength       =   12
         TabIndex        =   51
         Tag             =   "palets"
         Text            =   "palets"
         Top             =   4080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   645
         Index           =   33
         Left            =   -73200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   33
         Tag             =   "O6|T|S|||scapre|observa6||N|"
         Top             =   4440
         Width           =   7845
      End
      Begin VB.Frame FrameFactura 
         Height          =   3300
         Left            =   -74520
         TabIndex        =   88
         Top             =   1320
         Width           =   10575
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   43
            Left            =   3240
            MaxLength       =   15
            TabIndex        =   123
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   37
            Left            =   2520
            MaxLength       =   4
            TabIndex        =   122
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   40
            Left            =   4920
            MaxLength       =   5
            TabIndex        =   121
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   46
            Left            =   5520
            MaxLength       =   15
            TabIndex        =   120
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   44
            Left            =   3240
            MaxLength       =   15
            TabIndex        =   119
            Text            =   "Text1 7"
            Top             =   1875
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   38
            Left            =   2520
            MaxLength       =   4
            TabIndex        =   118
            Text            =   "Text1 7"
            Top             =   1875
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   41
            Left            =   4920
            MaxLength       =   5
            TabIndex        =   117
            Text            =   "Text1 7"
            Top             =   1875
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   47
            Left            =   5520
            MaxLength       =   15
            TabIndex        =   116
            Text            =   "Text1 7"
            Top             =   1875
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   45
            Left            =   3240
            MaxLength       =   15
            TabIndex        =   115
            Text            =   "Text1 7"
            Top             =   2205
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   39
            Left            =   2520
            MaxLength       =   4
            TabIndex        =   114
            Text            =   "Text1 7"
            Top             =   2205
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   42
            Left            =   4920
            MaxLength       =   5
            TabIndex        =   113
            Text            =   "Text1 7"
            Top             =   2205
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   48
            Left            =   5520
            MaxLength       =   15
            TabIndex        =   112
            Text            =   "Text1 7"
            Top             =   2205
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   54
            Left            =   7800
            MaxLength       =   15
            TabIndex        =   111
            Text            =   "Text1 7"
            Top             =   2205
            Width           =   1245
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   51
            Left            =   7200
            MaxLength       =   5
            TabIndex        =   110
            Text            =   "Text1 7"
            Top             =   2205
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   53
            Left            =   7800
            MaxLength       =   15
            TabIndex        =   109
            Text            =   "Text1 7"
            Top             =   1875
            Width           =   1245
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   50
            Left            =   7200
            MaxLength       =   5
            TabIndex        =   108
            Text            =   "Text1 7"
            Top             =   1875
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   52
            Left            =   7800
            MaxLength       =   15
            TabIndex        =   107
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   1245
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   49
            Left            =   7200
            MaxLength       =   5
            TabIndex        =   106
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   33
            Left            =   240
            MaxLength       =   15
            TabIndex        =   93
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   34
            Left            =   2160
            MaxLength       =   15
            TabIndex        =   92
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1365
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   35
            Left            =   3960
            MaxLength       =   15
            TabIndex        =   91
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1365
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   36
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   90
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   55
            Left            =   7200
            MaxLength       =   15
            TabIndex        =   89
            Text            =   "Text1 7"
            Top             =   2880
            Width           =   1845
         End
         Begin VB.Label Label1 
            Caption         =   "Base Imponible"
            Height          =   255
            Index           =   9
            Left            =   3360
            TabIndex        =   129
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. IVA"
            Height          =   255
            Index           =   33
            Left            =   5520
            TabIndex        =   128
            Top             =   1350
            Width           =   1335
         End
         Begin VB.Line Line1 
            X1              =   2400
            X2              =   9120
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Label Label1 
            Caption         =   "% IVA"
            Height          =   255
            Index           =   41
            Left            =   4920
            TabIndex        =   127
            Top             =   1350
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Cod. IVA"
            Height          =   255
            Index           =   42
            Left            =   2520
            TabIndex        =   126
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "% RE"
            Height          =   255
            Index           =   48
            Left            =   7200
            TabIndex        =   125
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. RE"
            Height          =   255
            Index           =   22
            Left            =   8040
            TabIndex        =   124
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Bruto"
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   102
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Dto PP"
            Height          =   255
            Index           =   11
            Left            =   2160
            TabIndex        =   101
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Dto Gn"
            Height          =   255
            Index           =   12
            Left            =   3960
            TabIndex        =   100
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Base Imponible"
            Height          =   255
            Index           =   8
            Left            =   5760
            TabIndex        =   99
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   30
            Left            =   1920
            TabIndex        =   98
            Top             =   480
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   31
            Left            =   3720
            TabIndex        =   97
            Top             =   480
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   32
            Left            =   5520
            TabIndex        =   96
            Top             =   480
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   36
            Left            =   11880
            TabIndex        =   95
            Top             =   2160
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "TOTAL OFERTA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   255
            Index           =   39
            Left            =   5640
            TabIndex        =   94
            Top             =   2880
            Width           =   1530
         End
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   5
         Left            =   10080
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   58
         Tag             =   "Descuento 1"
         Text            =   "OF"
         Top             =   4080
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Frame FrameCliente 
         Caption         =   "Datos Cliente"
         ForeColor       =   &H00972E0B&
         Height          =   2475
         Left            =   240
         TabIndex        =   65
         Top             =   370
         Width           =   10935
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   12
            Left            =   7430
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   78
            Tag             =   "Direccion/Dpto.|T|S|||scapre|nomdirec||N|"
            Text            =   "Text2"
            Top             =   210
            Width           =   3405
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   12
            Left            =   6885
            MaxLength       =   3
            TabIndex        =   7
            Tag             =   "Direccion/Dpto.|N|S|0|999|scapre|coddirec|000|N|"
            Text            =   "Text1"
            Top             =   210
            Width           =   540
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   11
            Left            =   1200
            MaxLength       =   30
            TabIndex        =   17
            Tag             =   "Provincia|T|N|||scapre|proclien||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text22"
            Top             =   1620
            Width           =   2565
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   9
            Left            =   1200
            MaxLength       =   6
            TabIndex        =   15
            Tag             =   "CPostal|T|N|||scapre|codpobla||N|"
            Text            =   "Text15"
            Top             =   1275
            Width           =   630
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   10
            Left            =   1875
            MaxLength       =   30
            TabIndex        =   16
            Tag             =   "Poblaci�n|T|N|||scapre|pobclien||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
            Top             =   1275
            Width           =   3405
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   7
            Left            =   1200
            MaxLength       =   20
            TabIndex        =   18
            Tag             =   "tel�fono Cliente|T|S|||scapre|telclien||N|"
            Text            =   "12345678911234567899"
            Top             =   2040
            Width           =   1965
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   6
            Left            =   1200
            MaxLength       =   15
            TabIndex        =   13
            Tag             =   "NIF Cliente|T|N|||scapre|nifclien||N|"
            Text            =   "123456789"
            Top             =   210
            Width           =   990
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   13
            Left            =   4320
            MaxLength       =   20
            TabIndex        =   6
            Tag             =   "Referencia Cliente|T|S|||scapre|referenc||N|"
            Text            =   "Text1 Text1 Text1 Te"
            Top             =   2040
            Width           =   1605
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   17
            Left            =   6885
            MaxLength       =   4
            TabIndex        =   8
            Tag             =   "Cod. Agente|N|N|0|9999|scapre|codagent|0000|N|"
            Text            =   "Text1"
            Top             =   560
            Width           =   540
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   17
            Left            =   7430
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   72
            Text            =   "Text2"
            Top             =   562
            Width           =   3405
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   14
            Left            =   6885
            MaxLength       =   3
            TabIndex        =   9
            Tag             =   "Forma de Pago|N|N|0|999|scapre|codforpa|000|N|"
            Text            =   "Text1"
            Top             =   910
            Width           =   510
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   14
            Left            =   7440
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   67
            Text            =   "Text2"
            Top             =   914
            Width           =   3390
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   15
            Left            =   6885
            MaxLength       =   5
            TabIndex        =   10
            Tag             =   "Descuento P.Pago|N|N|0|99.90|scapre|dtoppago|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1260
            Width           =   510
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   16
            Left            =   6885
            MaxLength       =   5
            TabIndex        =   19
            Tag             =   "Descuento General|N|N|0|99.90|scapre|dtognral|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1610
            Width           =   510
         End
         Begin VB.ComboBox cboFacturacion 
            Height          =   315
            Left            =   7845
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Tag             =   "Tipo Facturaci�n|N|N|||scapre|tipofact||N|"
            Top             =   1610
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Height          =   660
            Index           =   8
            Left            =   1200
            MaxLength       =   35
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Tag             =   "Domicilio|T|N|||scapre|domclien||N|"
            Text            =   "frmFacEntOfertas.frx":0396
            Top             =   562
            Width           =   4070
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   6
            Left            =   900
            Picture         =   "frmFacEntOfertas.frx":03BA
            ToolTipText     =   "Buscar poblaci�n"
            Top             =   1275
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Direc."
            Height          =   255
            Index           =   1
            Left            =   5700
            TabIndex        =   80
            Top             =   210
            Width           =   495
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   2
            Left            =   6600
            Picture         =   "frmFacEntOfertas.frx":04BC
            ToolTipText     =   "Buscar direc./dpto"
            Top             =   210
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Provincia"
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   79
            Top             =   1620
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Poblaci�n"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   77
            Top             =   1275
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Tfno:"
            Height          =   195
            Index           =   19
            Left            =   120
            TabIndex        =   76
            Top             =   2040
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "NIF"
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   75
            Top             =   210
            Width           =   615
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   900
            Picture         =   "frmFacEntOfertas.frx":05BE
            ToolTipText     =   "Buscar cliente varios"
            Top             =   220
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Ref. Cliente"
            Height          =   255
            Index           =   13
            Left            =   3360
            TabIndex        =   74
            Top             =   2040
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Agente"
            Height          =   255
            Index           =   34
            Left            =   5700
            TabIndex        =   73
            Top             =   562
            Width           =   615
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   5
            Left            =   6600
            Picture         =   "frmFacEntOfertas.frx":06C0
            ToolTipText     =   "Buscar agente"
            Top             =   562
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Forma Pago"
            Height          =   255
            Index           =   15
            Left            =   5700
            TabIndex        =   71
            Top             =   914
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. P.P"
            Height          =   255
            Index           =   25
            Left            =   5700
            TabIndex        =   70
            Top             =   1266
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. Gral"
            Height          =   255
            Index           =   26
            Left            =   5700
            TabIndex        =   69
            Top             =   1610
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Facturac."
            Height          =   255
            Index           =   4
            Left            =   7845
            TabIndex        =   68
            Top             =   1380
            Width           =   1095
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   6600
            Picture         =   "frmFacEntOfertas.frx":07C2
            ToolTipText     =   "Buscar forma de pago"
            Top             =   914
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Domicilio"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   66
            Top             =   562
            Width           =   735
         End
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   19
         Left            =   -73200
         MaxLength       =   80
         TabIndex        =   21
         Tag             =   "Plazo Entrega 2|T|S|||scapre|plazos02||N|"
         Top             =   740
         Width           =   7845
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   20
         Left            =   -73200
         MaxLength       =   80
         TabIndex        =   22
         Tag             =   "Validez de la oferta|T|S|||scapre|plazos03||N|"
         Top             =   1100
         Width           =   7845
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   18
         Left            =   -73200
         MaxLength       =   80
         TabIndex        =   20
         Tag             =   "Plazo Entrega 1|T|S|||scapre|plazos01||N|"
         Top             =   450
         Width           =   7845
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Index           =   1
         Left            =   2640
         TabIndex        =   64
         ToolTipText     =   "Buscar art�culo"
         Top             =   3960
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Index           =   0
         Left            =   960
         TabIndex        =   63
         ToolTipText     =   "Buscar almacen"
         Top             =   3960
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   2
         Left            =   2880
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   50
         Tag             =   "Nombre Art�culo"
         Text            =   "nomArtic"
         Top             =   4080
         Visible         =   0   'False
         Width           =   3285
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   8
         Left            =   11760
         MaxLength       =   12
         TabIndex        =   59
         Tag             =   "Importe"
         Text            =   "Importe"
         Top             =   4080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   7
         Left            =   11160
         MaxLength       =   30
         TabIndex        =   56
         Tag             =   "Descuento 2"
         Text            =   "Dto2"
         Top             =   4080
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   6
         Left            =   10560
         MaxLength       =   5
         TabIndex        =   55
         Tag             =   "Descuento 1"
         Text            =   "Dto1"
         Top             =   4080
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   4
         Left            =   9240
         MaxLength       =   12
         TabIndex        =   54
         Tag             =   "Precio"
         Text            =   "123,456.7879"
         Top             =   4080
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   1275
         Index           =   32
         Left            =   -74520
         MultiLine       =   -1  'True
         TabIndex        =   35
         Tag             =   "Gesti�n Oferta|T|S|||scapre|seguiofe||N|"
         Text            =   "frmFacEntOfertas.frx":08C4
         Top             =   2640
         Width           =   9285
      End
      Begin VB.TextBox Text1 
         Height          =   1275
         Index           =   31
         Left            =   -74520
         MultiLine       =   -1  'True
         TabIndex        =   34
         Tag             =   "Concepto Oferta|T|S|||scapre|concepto||N|"
         Text            =   "frmFacEntOfertas.frx":08CC
         Top             =   840
         Width           =   9285
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   3
         Left            =   8040
         MaxLength       =   16
         TabIndex        =   53
         Tag             =   "Cantidad"
         Text            =   "1,234,567,891.25"
         Top             =   4080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   1200
         MaxLength       =   18
         TabIndex        =   49
         Tag             =   "C�digo Art�culo"
         Text            =   "Artic Artic Artic5"
         Top             =   3900
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   360
         MaxLength       =   15
         TabIndex        =   48
         Tag             =   "C�digo Almacen"
         Text            =   "codalmac"
         Top             =   3900
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   30
         Left            =   -73200
         MaxLength       =   80
         TabIndex        =   32
         Tag             =   "Observaci�n 5|T|S|||scapre|observa05||N|"
         Top             =   4100
         Width           =   7845
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   29
         Left            =   -73200
         MaxLength       =   80
         TabIndex        =   31
         Tag             =   "Observaci�n 4|T|S|||scapre|observa04||N|"
         Top             =   3830
         Width           =   7845
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   28
         Left            =   -73200
         MaxLength       =   80
         TabIndex        =   30
         Tag             =   "Observaci�n 3|T|S|||scapre|observa03||N|"
         Top             =   3560
         Width           =   7845
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   27
         Left            =   -73200
         MaxLength       =   80
         TabIndex        =   29
         Tag             =   "Observaci�n 2|T|S|||scapre|observa02||N|"
         Top             =   3290
         Width           =   7845
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   26
         Left            =   -73200
         MaxLength       =   80
         TabIndex        =   28
         Tag             =   "Observaci�n 1|T|S|||scapre|observa01||N|"
         Top             =   3020
         Width           =   7845
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   25
         Left            =   -73200
         MaxLength       =   80
         TabIndex        =   27
         Tag             =   "Asunto Carta 5|T|S|||scapre|asunto05||N|"
         Top             =   2600
         Width           =   7845
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   24
         Left            =   -73200
         MaxLength       =   80
         TabIndex        =   26
         Tag             =   "Asunto Carta 4|T|S|||scapre|asunto04||N|"
         Top             =   2330
         Width           =   7845
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   23
         Left            =   -73200
         MaxLength       =   80
         TabIndex        =   25
         Tag             =   "Asunto Carta 3|T|S|||scapre|asunto03||N|"
         Top             =   2060
         Width           =   7845
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   22
         Left            =   -73200
         MaxLength       =   80
         TabIndex        =   24
         Tag             =   "Asunto Carta|T|S|||scapre|asunto02||N|"
         Top             =   1790
         Width           =   7845
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   21
         Left            =   -73200
         MaxLength       =   80
         TabIndex        =   23
         Tag             =   "Asunto Carta 1|T|S|||scapre|asunto01||N|"
         Top             =   1520
         Width           =   7845
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmFacEntOfertas.frx":08D4
         Height          =   2145
         Left            =   240
         TabIndex        =   60
         Top             =   2880
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   3784
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   $"frmFacEntOfertas.frx":08E9
         Height          =   2055
         Left            =   11520
         TabIndex        =   130
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Validez Oferta"
         Height          =   255
         Index           =   2
         Left            =   -74400
         TabIndex        =   105
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Concepto Oferta"
         Height          =   255
         Index           =   37
         Left            =   -74490
         TabIndex        =   62
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Gesti�n Oferta"
         Height          =   255
         Index           =   38
         Left            =   -74490
         TabIndex        =   61
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones"
         Height          =   255
         Index           =   45
         Left            =   -74400
         TabIndex        =   47
         Top             =   3020
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Asunto Carta"
         Height          =   255
         Index           =   5
         Left            =   -74400
         TabIndex        =   45
         Top             =   1520
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Plazo Entrega"
         Height          =   255
         Index           =   3
         Left            =   -74400
         TabIndex        =   44
         Top             =   450
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   11850
      TabIndex        =   38
      Top             =   6720
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Width           =   13185
      _ExtentX        =   23257
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   28
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Lineas Oferta"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Generar Pedido"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Cargar Plantilla"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Traer de Oferta"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Recordatorio"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Valoraci�n"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Oferta"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Fact. Pro forma"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "�ltimo"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   285
         Index           =   0
         Left            =   8400
         MaxLength       =   15
         TabIndex        =   104
         Text            =   "TOTAL"
         Top             =   100
         Width           =   1490
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   56
         Left            =   9960
         MaxLength       =   15
         TabIndex        =   103
         Text            =   "Text1 7"
         Top             =   80
         Width           =   1530
      End
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   6960
         TabIndex        =   42
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Ampliaci�n L�nea"
      Height          =   255
      Index           =   35
      Left            =   2400
      TabIndex        =   46
      Top             =   6600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver Todos"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         HelpContextID   =   2
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   2
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         HelpContextID   =   2
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnLineas 
         Caption         =   "&Lineas"
         HelpContextID   =   2
         Shortcut        =   ^L
      End
      Begin VB.Menu mnGenPedido 
         Caption         =   "&Generar Pedido"
         HelpContextID   =   2
         Shortcut        =   ^G
      End
      Begin VB.Menu mnPlantillas 
         Caption         =   "&Plantillas"
         HelpContextID   =   2
         Shortcut        =   ^P
         Visible         =   0   'False
      End
      Begin VB.Menu mnOferta 
         Caption         =   "Traer &Oferta"
         HelpContextID   =   2
         Shortcut        =   ^O
         Visible         =   0   'False
      End
      Begin VB.Menu mnBarra4 
         Caption         =   "-"
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Begin VB.Menu mnImpOferta 
            Caption         =   "&Oferta"
            Shortcut        =   ^I
         End
         Begin VB.Menu mnImpRecordatorio 
            Caption         =   "&Recordatorio"
            Shortcut        =   ^R
         End
         Begin VB.Menu mnImpValoracion 
            Caption         =   "&Valoraci�n"
            HelpContextID   =   2
            Shortcut        =   ^V
         End
         Begin VB.Menu mnImpFactProF 
            Caption         =   "&Factura Pro Forma"
            Shortcut        =   ^T
         End
      End
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmFacEntOfertas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
'Public Event DatoSeleccionado(CadenaSeleccion As String)

Public EsHistorico As Boolean 'Si es true abrir el formulario con la tabla de
                              'de historico schpre, y solo en modo de consulta


Public DatosOferta As String   'Para situarla

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1

Private WithEvents frmCP As frmCPostal 'Codigos Postales
Attribute frmCP.VB_VarHelpID = -1
Private WithEvents frmC As frmFacClientes 'Form Mto Clientes
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCV As frmFacClientesV  'Form Mto Clientes Varios
Attribute frmCV.VB_VarHelpID = -1
Private WithEvents frmFP As frmFacFormasPago 'Form Mto Formas de Pago
Attribute frmFP.VB_VarHelpID = -1
Private WithEvents frmT As frmAdmTrabajadores  'Form Mto Trabajadores
Attribute frmT.VB_VarHelpID = -1
Private WithEvents frmA As frmFacAgentesCom   'Form Mto Agentes
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmAlm As frmAlmAlPropios   'Form Almacenes Propios
Attribute frmAlm.VB_VarHelpID = -1
Private WithEvents frmArt As frmAlmArticulos   'Form Articulos
Attribute frmArt.VB_VarHelpID = -1

Private WithEvents frmList As frmListadoOfer 'Listados para Ofertas
Attribute frmList.VB_VarHelpID = -1

'Carga de Plantillas en la linea de la Oferta
Private WithEvents frmPlant As frmFacCargaPlantilla  'Form para cargar plantillas
Attribute frmPlant.VB_VarHelpID = -1
'Carga las lineas de otra Oferta
Private WithEvents frmTOferta As frmFacTraerOferta
Attribute frmTOferta.VB_VarHelpID = -1


'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'   5.-  Mantenimiento Lineas
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
Private Modo As Byte

Dim ModificaLineas As Byte
'1.- A�adir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

Dim PrimeraVez As Boolean

Dim EsCabecera As Boolean
'Para saber en MandaBusquedaPrevia si busca en la tabla scapla o en la tabla sdirec

Dim CodTipoMov As String
'Codigo tipo de movimiento en funci�n del valor en la tabla de par�metros: stipom

Dim EsDeVarios As Boolean
'Si el cliente mostrado es de Varios o No


Private CadenaConsulta As String 'SQL de la tabla principal del formulario
Private CadenaSQL As String 'Para crear consulta de Generar Pedido a partir de la Oferta

Private Ordenacion As String   'ORDER BY de la cadena consulta
Private NombreTabla As String  'Nombre de la tabla o de la
Private NomTablaLineas As String 'Nombre de la Tabla de lineas
Private Kcampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Dim btnAnyadir As Byte
'Variable que indica el n�mero del Boton  Anyadir en la Toolbar1
Dim btnPrimero As Byte
'Variable que indica el n�mero del Boton  PrimerRegistro en la Toolbar1

Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos
Private HaCambiadoCP As Boolean
'Para saber si tras haber vuelto de prismaticos ha cambiado el valor del CPostal


Dim PorCaja As Boolean
'Para Saber si se ha salido con precio caja y hay que calcular el importe de la
'linea aplicando el precio de la caja. Si PorCaja=false se aplicaca el precio de unidad

Dim Precio As String 'Precio de la linea de Articulo

Dim txtAnterior As String

'=====================================================================================


Private Sub cboFacturacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        Me.SSTab1.Tab = 1
        PonerFoco Text1(18)
    End If
End Sub

Private Sub chkAceptado_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub chkAceptado_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim PrimeraLin As Boolean

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                InsertarCabecera
                Me.SSTab1.Tab = 0
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me, 1) Then
                    'Actualizar los datos del cliente si es de varios
                    EsDeVarios = EsClienteVarios(Text1(4).Text)
                    If EsDeVarios Then ActualizarClienteVarios Text1(4).Text, Text1(6).Text
                    TerminaBloquear
                    PosicionarData
                End If
            End If
            
         Case 5 'INSERTAR MODIFICAR LINEA
            'Actualizar el registro en la tabla de lineas 'slima1' (Revisiones)
            If ModificaLineas = 1 Then 'INSERTAR lineas Ofertas
                PrimeraLin = False
                If Data2.Recordset.EOF = True Then PrimeraLin = True
                If InsertarLinea Then
                    If PrimeraLin Then
                        CargaGrid DataGrid1, Data2, True
                    Else
                        CargaGrid2 DataGrid1, Data2
                    End If
                    BotonAnyadirLinea
                End If
            ElseIf ModificaLineas = 2 Then 'MODIFICAR lineas
                If ModificarLinea Then
                    TerminaBloquear
                    CargaTxtAux False, False
                    CargaGrid2 DataGrid1, Data2
                    ModificaLineas = 0
                    PonerBotonCabecera True
                    BloquearTxt Text2(16), True
                End If
                Me.DataGrid1.Enabled = True
            End If
            CalcularDatosFactura
    End Select
    Screen.MousePointer = vbDefault

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdAux_Click(Index As Integer)
    Select Case Index
        Case 0 'Busqueda de Cod. Almacen
            Set frmAlm = New frmAlmAlPropios
            frmAlm.DatosADevolverBusqueda = "0"
            frmAlm.Show vbModal
            Set frmAlm = Nothing
        Case 1 'Busqueda de Cod. Artic
            Set frmArt = New frmAlmArticulos
            frmArt.DatosADevolverBusqueda2 = "@1@" 'Poner en Modo Busqueda
            frmArt.ParaVenta = True
            frmArt.DeConsulta = True
            frmArt.Show vbModal
            Set frmArt = Nothing
    End Select
    PonerFoco txtAux(Index)
End Sub

Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 1, 3 'Busqueda, Insertar
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            LimpiarDataGrids
            PonerModo 0
            PonerFoco Text1(0)
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
        Case 5 'Lineas Detalle
            TerminaBloquear
            CargaTxtAux False, False
            BloquearTxt Text2(16), True
            If ModificaLineas = 1 Then 'INSERTAR
                DataGrid1.AllowAddNew = False
                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
            End If
            ModificaLineas = 0
            PonerBotonCabecera True
            Me.DataGrid1.Enabled = True
    End Select
End Sub


Private Sub BotonAnyadir()
'A�adir registro en tabla de trabajadores: straba (Cabecera)
Dim NomTraba As String

    LimpiarCampos 'Vac�a los TextBox
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    BloquearTxt Text1(0), True, True
    
    'Poner el nombre del trabajador que esta conectado
    Text1(3).Text = PonerTrabajadorConectado(NomTraba)
    Text2(3).Text = NomTraba

    Text1(1).Text = Format(Now, "dd/mm/yyyy") 'Fecha Oferta
    PonerFoco Text1(1)
End Sub


Private Sub BotonAnyadirLinea()
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
       
    ModificaLineas = 1 'Ponemos Modo A�adir Linea
    Precio = ""
    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerBotonCabecera False
    Me.SSTab1.Tab = 0
    lblIndicador.Caption = "INSERTAR"
    
    AnyadirLinea DataGrid1, Data2
    CargaTxtAux True, True
    
    'Poner el Almacen por defecto del Trabajador
    txtAux(0).Text = DevuelveDesdeBDNew(1, "straba", "codalmac", "codtraba", Text1(3).Text, "N")
    If txtAux(0).Text <> "" Then txtAux(0).Text = Format(txtAux(0).Text, "000")
    
    'Campo Ampliacion Linea
    Text2(16).Text = ""
    BloquearTxt Text2(16), False
    
    PonerFoco txtAux(1)
    Me.DataGrid1.Enabled = False
End Sub


Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
        PonerModo 1
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Text1(0).BackColor = vbYellow
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(Kcampo).Text = ""
            Text1(Kcampo).BackColor = vbYellow
            PonerFoco Text1(Kcampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
'Ver todos
    
    If chkVistaPrevia.Value = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia ""
    Else
        LimpiarCampos
        LimpiarDataGrids
        CadenaConsulta = "Select * from " & NombreTabla & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index
    PonerCampos
End Sub


Private Sub BotonModificar()
Dim DeVarios As Boolean

    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    PonerFoco Text1(1)
        
    'Si es Cliente de Varios no se pueden modificar sus datos
    DeVarios = EsClienteVarios(Text1(4).Text)
    BloquearDatosCliente (DeVarios)
End Sub


Private Sub BotonModificarLinea()
'Modificar una linea
Dim vWhere As String
    
    On Error GoTo EModificarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub '1= Insertar
    Me.SSTab1.Tab = 0
    
    If Data2.Recordset.EOF Then Exit Sub
    vWhere = ObtenerWhereCP & " and numlinea=" & Data2.Recordset!numlinea
    If Not BloqueaRegistro(NomTablaLineas, vWhere) Then Exit Sub
    CargaTxtAux True, False
    ModificaLineas = 2 'Modificar
    
    'A�adiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR"
    PonerBotonCabecera False
    BloquearTxt Text2(16), False 'Campo Ampliacion Linea
    BloquearTxt txtAux(2), True 'campo nombre articulo
    PonerFoco txtAux(0)
    Me.DataGrid1.Enabled = False
    
EModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Mantenimientos (scaman)
' y los registros correspondientes de las tablas de lineas (sliman y slima1)
Dim Cad As String
Dim vTipoMov As CTiposMov
Dim NumOferElim As Long

    
    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    Cad = "Cabecera de Ofertas." & vbCrLf
    Cad = Cad & "-----------------------------" & vbCrLf & vbCrLf
    Cad = Cad & "Va a eliminar la Oferta:            "
    Cad = Cad & vbCrLf & "N�:  " & Format(Text1(0).Text, "0000000")
    Cad = Cad & vbCrLf & "Cliente:  " & Format(Text1(4).Text, "000000") & " - " & Text1(5).Text
    Cad = Cad & vbCrLf & vbCrLf & " �Desea Eliminarla? "
    
    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        
        NumRegElim = Data1.Recordset.AbsolutePosition
        NumOferElim = Data1.Recordset.Fields(0).Value
        
        If Not Eliminar Then
            Screen.MousePointer = vbDefault
            Exit Sub
        ElseIf SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            LimpiarDataGrids
            PonerModo 0
        End If
        
        'Devolvemos contador, si no estamos actualizando
        Set vTipoMov = New CTiposMov
        vTipoMov.DevolverContador CodTipoMov, NumOferElim
        Set vTipoMov = Nothing
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Oferta", Err.Description
End Sub


Private Sub BotonEliminarLinea()
'Eliminar una linea De Mantenimiento. Tabla: slima1
Dim SQL As String
    
    On Error GoTo EEliminarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar

    If Data2.Recordset.EOF Then Exit Sub
        
    Me.SSTab1.Tab = 0
    ModificaLineas = 3 'Eliminar
    
    SQL = "�Seguro que desea eliminar la l�nea de Oferta?     "
    SQL = SQL & vbCrLf & "NumLinea:  " & Data2.Recordset!numlinea & vbCrLf
    SQL = SQL & "Almacen:  " & Format(Data2.Recordset!codAlmac, "000")
    SQL = SQL & vbCrLf & "Art�culo:  " & Data2.Recordset!codartic & " - " & Data2.Recordset!NomArtic
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Data2.Recordset.AbsolutePosition
        SQL = "Delete from " & NomTablaLineas & " WHERE " & ObtenerWhereCP
        SQL = SQL & " and numlinea=" & Data2.Recordset!numlinea
        conn.Execute SQL
        
        ModificaLineas = 0
        CargaGrid2 DataGrid1, Data2
        SituarDataTrasEliminar Data2, NumRegElim
        CalcularDatosFactura
'        CancelaADODC
    End If
    PonerFocoBtn Me.cmdRegresar

EEliminarLinea:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Lineas Mantenimientos", Err.Description
End Sub


Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim Cad As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
        PonerModo 2
        'BloquearTabs False
        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        If DataGrid1.Row >= 0 Then
            DeseleccionaGrid DataGrid1
            DataGrid1.Bookmark = 1
        End If
    Else 'Se llama desde alg�n Prismatico de otro Form al Mantenimiento de Trabajadores
'        If Data1.Recordset.EOF Then
'            MsgBox "Ning�n registro devuelto.", vbExclamation
'            Exit Sub
'        End If
'        Cad = Data1.Recordset.Fields(0) & "|"
'        Cad = Cad & Data1.Recordset.Fields(1) & "|"
'        RaiseEvent DatoSeleccionado(Cad)
        Unload Me
    End If
End Sub


Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    On Error GoTo EKeyPress
    
    If KeyAscii = 27 Then 'ESC
        If Modo = 5 Then 'Modo Lineas
            cmdRegresar_Click
        ElseIf Modo = 0 Or Modo = 2 Then 'Estamos en Cabecera
            Unload Me
        End If
    End If
    
EKeyPress:
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub DataGrid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Ayuda de Etiqueta de precio de salida de la Funci�n de Precios
    If (Modo <> 2 And Modo <> 5) Then Exit Sub
    If Data2.Recordset.EOF Then Exit Sub
    If (Modo = 2) Or (Modo = 5 And ModificaLineas = 0) Then
        If X > 7750 And X < 8000 Then
            Select Case DataGrid1.Columns(8).Value
                Case "P": Me.DataGrid1.ToolTipText = "P: Promoci�n"
                Case "E": Me.DataGrid1.ToolTipText = "E: Precio Especial"
                Case "T": Me.DataGrid1.ToolTipText = "T: Tarifa Art�culo"
                Case "A": Me.DataGrid1.ToolTipText = "A: Precio Art�culo"
                Case "M": Me.DataGrid1.ToolTipText = "M: Manual"
                Case Else
                    Me.DataGrid1.ToolTipText = ""
            End Select
        Else
            Me.DataGrid1.ToolTipText = ""
        End If
    End If
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo Error1
    
    If Not Data2.Recordset.EOF And ModificaLineas <> 1 Then '1: Insertar
        'Poner descripcion de ampliacion lineas
        Text2(16).Text = DevuelveDesdeBDNew(1, NomTablaLineas, "ampliaci", "numofert", Text1(0).Text, "N", , "numlinea", Data2.Recordset!numlinea, "N")
    Else
        Text2(16).Text = ""
    End If

Error1:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        If DatosOferta <> "" Then
            PonerModo 1
            Text1(0).Text = DatosOferta
            HacerBusqueda
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmppal.Icon
    PrimeraVez = True
    ' ICONITOS DE LA BARRA
    btnAnyadir = 5
    btnPrimero = 25
    With Me.Toolbar1
        .ImageList = frmppal.imgListComun
        .Buttons(1).Image = 1   'Bot�n Buscar
        .Buttons(2).Image = 2   'Bot�n Todos
        .Buttons(5).Image = 3   'Insertar Nuevo
        .Buttons(6).Image = 4   'Modificar
        .Buttons(7).Image = 5   'Borrar
        
        .Buttons(10).Image = 10 'Mto Lineas Ofertas
        .Buttons(11).Image = 26 'Generar Pedido
        .Buttons(12).Image = 32 'Cargar Plantilla
        .Buttons(13).Image = 24 'Traer Lineas de Otra Oferta
        
        .Buttons(16).Image = 30 'Recordatorio
        .Buttons(17).Image = 31 'Valoracion
        .Buttons(18).Image = 16 'Imprimir
        .Buttons(19).Image = 40 'Imprimir factura pro forma
        
        .Buttons(22).Image = 15  'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 '�ltimo
    End With
    Me.SSTab1.Tab = 0
      
    LimpiarCampos   'Limpia los campos TextBox
'    CargarComboTipoPago
    CargarComboFacturacion
    CodTipoMov = "OFE"
    VieneDeBuscar = False 'Para el CPostal
   
    'Comprobar si es Departamento o Direccion
    If vParamAplic.Departamento Then
        Me.Label1(1).Caption = "Dpto."
    Else
        Me.Label1(1).Caption = "Direc."
    End If
        
    '## A mano
    If Not EsHistorico Then
        NombreTabla = "scapre"
        NomTablaLineas = "slipre" 'Tabla lineas de Ofertas
        Me.Caption = "Ofertas Clientes"
    Else
        NombreTabla = "schpre"
        NomTablaLineas = "slhpre"
        CargarTagsHco Me, "scapre", NombreTabla
        Me.Caption = "Hist�rico Ofertas Clientes"
    End If
    Ordenacion = " ORDER BY numofert "



  'Si es AVAB longitud NIF y domicilio cmabian
    If vParamAplic.EsAVAB Then
        'AVAB
        '.....................
        'NIF
        Text1(6).MaxLength = 50
        Text1(6).Width = 3990
        'Domicilio
        Text1(8).MaxLength = 100
        Text1(8).Height = 675
  
        Text1(33).visible = True
  
    Else
        'MORALES
        Text1(6).MaxLength = 15
        Text1(6).Width = 1590
        Text1(8).MaxLength = 35
        Text1(8).Height = Text1(6).Height


        Text1(33).visible = False   'Para morales NO dejo ver el observa6 ... de momento

    End If
  





    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    Data1.RecordSource = "Select * from " & NombreTabla & " where numofert=-1"
    Data1.Refresh
'    If DatosADevolverBusqueda = "" Then
    If Me.DatosOferta = "" Then
        PonerModo 0
    Else
        PonerModo 1
        Text1(0).BackColor = vbYellow
    End If
    
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True

    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    
    PrimeraVez = True
End Sub


Private Sub LimpiarCampos()
On Error Resume Next

    limpiar Me   'Metodo general: Limpia los controles TextBox
    
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.chkAceptado(0).Value = False
    Me.cboFacturacion.ListIndex = -1
    
    Text3(0).Text = "BASE IMP."
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Trabajadores
Dim Indice As Byte
    Indice = 17
    Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000") 'Cod Trabajador
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Trabajador
End Sub

Private Sub frmAlm_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Almacenes Propios
    txtAux(0).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Almacen
End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Articulos
    txtAux(1).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Artic
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim cadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        If EsCabecera Then 'Llama desde VerTodos del Form
            cadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
            cadB = Aux
            If EsHistorico Then
                Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
                cadB = cadB & " and " & Aux
            End If
            
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
            Text1(0).Text = Format(RecuperaValor(CadenaDevuelta, 1), "0000000")
        Else 'Llama desde Prismatico Direcciones/Departamentos
            Text1(12).Text = Format(RecuperaValor(CadenaDevuelta, 1), "000")
            Text2(12).Text = RecuperaValor(CadenaDevuelta, 2)
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Clientes
    Text1(4).Text = RecuperaValor(CadenaSeleccion, 1)  'Cod Clien
    HaDevueltoDatos = True
End Sub


Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento C. Postales
Dim Indice As Byte
Dim Devuelve As String

    Indice = 9
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
    Text1(Indice + 1).Text = ObtenerPoblacion(Text1(Indice).Text, Devuelve)  'Poblacion
    'provincia
    Text1(Indice + 2).Text = Devuelve
End Sub


Private Sub frmCV_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Clientes Varios
Dim Indice As Byte

    Indice = 6
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'NIF
    Text1(Indice - 1).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Clien
    PonerDatosClienteVario (Text1(Indice).Text)
End Sub

Private Sub frmF_Selec(vFecha As Date) 'Calendario Fechas
Dim Indice As Byte
    Indice = CByte(Me.imgFecha(0).Tag) + 1
    Text1(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Formas de Pago
Dim Indice As Byte
    Indice = 14
    Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Forma Pago
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Forma Pago
End Sub


Private Sub frmList_DatoSeleccionado(CadenaSeleccion As String)
'Aqui devuelve los valores que se introducen el Listado de Oferta para generar el Pedido
Dim vSQL As String

    'Construimos parte de la SQL para insertar en Pedidos
    vSQL = ""
    vSQL = " '" & Format(RecuperaValor(CadenaSeleccion, 2), FormatoFecha) & "' as fecpedcl, '" 'Fecha Pedido
    vSQL = vSQL & Format(RecuperaValor(CadenaSeleccion, 3), FormatoFecha) & "' as fecentre, " 'Fecha entrega
    vSQL = vSQL & RecuperaValor(CadenaSeleccion, 4) & " as sementre, " 'Sem entrega
    vSQL = vSQL & "0 as visadore, " & "codclien, nomclien, domclien, codpobla, pobclien, proclien, nifclien, "
    vSQL = vSQL & "telclien, coddirec, nomdirec, referenc, "
    vSQL = vSQL & RecuperaValor(CadenaSeleccion, 1) & " as codtraba, " 'Operador de Pedido
    vSQL = vSQL & "codagent, codforpa, dtoppago, dtognral, tipofact, observa01, observa02, observa03, "
    vSQL = vSQL & "observa04, observa05, 0 as servcomp,0 as restoped, " & Text1(0).Text & " as numofert, '" 'N� Oferta
    vSQL = vSQL & Format(Text1(1).Text, FormatoFecha) & "' as fecofert " 'Fecha Oferta
    'NUEVO. Observaciones 6
    vSQL = vSQL & ",observa6 "
    
    
    CadenaSQL = vSQL
End Sub


Private Sub frmPlant_CargarPlantillas()
Dim RS As ADODB.Recordset
Dim RSLineas As ADODB.Recordset
Dim SQL As String, Devuelve As String
Dim codAlmac As String, codTarif As String
Dim Cantidad As Integer
Dim NumCajas As Integer, RestoUnid As Integer
Dim Precio As String, Dto1 As String, Dto2 As String
Dim OrigP As String 'De donde viene el precio: promocion, precio especial,...
Dim CPrecioFact As CPreciosFact

    Screen.MousePointer = vbHourglass
    
    'Si se ha seleccionado alguna plantilla para a�adir sus lineas a la Oferta
    '(cantidad de alguna linea de tmpscapla > 0), entonces a�adimos todas las
    'lineas de esa oferta poniendo en cantidad de slipre de lineas de oferta
    'el resultado de multiplicar la cantidad de tmpscapla * cantidad de slipla
    SQL = "SELECT * FROM tmpscapla WHERE codusu=" & vUsu.Codigo & " AND cantidad>0"
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    'Obtener el almacen por defecto del trabajador
    'Poner el Almacen por defecto del Trabajador
    codAlmac = DevuelveDesdeBDNew(conAri, "straba", "codalmac", "codtraba", Text1(3).Text, "N")
    'Obtener la tarifa del cliente
    codTarif = DevuelveDesdeBDNew(conAri, "sclien", "codtarif", "codclien", Text1(4).Text, "N")

    While Not RS.EOF  'Para cada plantilla
        'A�adimos todas las lineas de esa plantilla en la cantidad correcta en las
        'lineas de la Oferta
        SQL = "SELECT * FROM slipla WHERE codplant=" & RS!codPlant
        Set RSLineas = New ADODB.Recordset
        RSLineas.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not RSLineas.EOF
            'Comprobar si el articulo se vende por cajas antes de entrar a la funci�n
            Devuelve = DevuelveDesdeBDNew(conAri, "sartic", "unicajas", "codartic", RSLineas!codartic, "T")
            If Devuelve <> "" Then
            'Si se puede vender por cajas(devuelve>1) poner numero de unidades/caja en una linea con el
            'precio de caja, y otra linea con el resto unidades un precio unidad
                Cantidad = (RS!Cantidad * RSLineas!Cantidad)
                NumCajas = ObtenerNumCajas(CStr(Cantidad), Devuelve)
                RestoUnid = CInt(Cantidad) - NumCajas * CInt(Devuelve)
                'Obtener el precio a aplicar
                Set CPrecioFact = New CPreciosFact
                CPrecioFact.CodigoLista = codTarif
                CPrecioFact.CodigoArtic = RSLineas!codartic
                CPrecioFact.CodigoClien = Text1(4).Text
                PorCaja = (NumCajas > 0)
                Precio = CPrecioFact.ObtenerPrecio(PorCaja, Text1(1).Text, OrigP)
                
                'Si PorCaja vuelve de ObtenerPrecio a False se aplica precio
                'de Unidad aunque se venda por cajas, ya que ha regresado con pvp de articulo
                Dto1 = CPrecioFact.Descuento1
                Dto2 = CPrecioFact.Descuento2
                Set CPrecioFact = Nothing
                    
                If PorCaja And NumCajas > 0 Then 'El Articulo se Vende Por Cajas y Cantidad supera la cant en 1 caja
                    'Obtener el precio y los descuentos adecuados
                    'Insertar 2 lineas: 1 linea con la cantidad que se puede vender en cajas y al precio de caja
                    InsertarLineaDePlantilla RSLineas!codartic, codAlmac, NumCajas * CInt(Devuelve), Precio, Dto1, Dto2, OrigP
                    '2 linea con el resto de la cantidad que no llega a una caja a precio de unidad
                    If RestoUnid > 0 Then InsertarLineaDePlantilla RSLineas!codartic, codAlmac, RestoUnid, Precio, Dto1, Dto2, OrigP
'                    Else
'                        InsertarLineaDePlantilla rsLineas!codArtic, codAlmac, codTarif, Cantidad, 0
'                    End If
                Else 'No llega a una caja
                    InsertarLineaDePlantilla RSLineas!codartic, codAlmac, Cantidad, Precio, Dto1, Dto2, OrigP
                End If
            End If
            RSLineas.MoveNext
        Wend
        RSLineas.Close
        Set RSLineas = Nothing
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing

    'Borrar de la Tabla Temporal (tmpscapla) los registros insertados tras a�adir
    'las lineas de las plantillas seleccionadas
    DescargarDatosTMP
    'Actualizar el Grid con las lineas A�adidas
    PonerCamposLineas
    DataGrid1.Enabled = True
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Trabajadores
Dim Indice As Byte
    Indice = 3
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Trabajador
    FormateaCampo Text1(Indice)
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Trabajador
End Sub


Private Sub frmTOferta_CargarOferta(NumOfe As String)
Dim RS As ADODB.Recordset
Dim SQL As String
Dim numlinea As String, vWhere As String
Dim I  As Integer
    On Error GoTo ECargarOferta
    
    'Si se ha seleccionado alguna oferta para a�adir sus lineas a la Oferta
    If NumOfe = "" Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    SQL = "Select * from " & NomTablaLineas & " where numofert=" & RecuperaValor(NumOfe, 1)
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    While Not RS.EOF  'Para cada linea de oferta
        'Obtener el siguiente numero de linea
        vWhere = ObtenerWhereCP
        numlinea = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", vWhere)
        
        SQL = "INSERT INTO " & NomTablaLineas & " (numofert, numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel,origpre) "
        SQL = SQL & " VALUES(" & Text1(0).Text & ", " & numlinea & ", " & RS!codAlmac & ", " & DBSet(RS!codartic, "T") & ", " & DBSet(RS!NomArtic, "T") & ", "
        SQL = SQL & DBSet(RS!ampliaci, "T", "S")
        SQL = SQL & ", " & DBSet(RS!Cantidad, "N") & ", " & DBSet(RS!precioar, "N") & ", " & DBSet(RS!dtoline1, "N") & ", " & DBSet(RS!dtoline2, "N") & ", " & DBSet(RS!ImporteL, "N") & ", "
        SQL = SQL & DBSet(CStr(RS!origpre), "T", "S") & ")"
        
        conn.Execute SQL
        RS.MoveNext
    Wend
    RS.Close
    
    
    SQL = RecuperaValor(NumOfe, 2)  'Copio observaciones
    vWhere = RecuperaValor(NumOfe, 3)  'Copio datos carta
    I = Val(SQL) + Val(vWhere)
    If I > 0 Then
        'Cargo en RS la oferta
        SQL = "Select * from " & NombreTabla & " where numofert=" & RecuperaValor(NumOfe, 1)
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then
            'UPDATEAMOS los campos de la oferta de observaciones
            SQL = ""
            If RecuperaValor(NumOfe, 2) = "1" Then 'Copio observaciones
                
                
                For I = 1 To 5
                    vWhere = "observa0" & I
                    numlinea = ", " & vWhere & " = " & DBSet(DBLet(RS.Fields(vWhere), "T"), "T", "S")
                    SQL = SQL & numlinea
                Next I
                
                
            End If
            
            If RecuperaValor(NumOfe, 2) = "1" Then 'Copio carta
                For I = 1 To 5
                    vWhere = "asunto0" & I
                    numlinea = ", " & vWhere & " = " & DBSet(DBLet(RS.Fields(vWhere), "T"), "T", "S")
                    SQL = SQL & numlinea
                Next I
                
            End If
            SQL = Mid(SQL, 2)  'quito la primera coma
            SQL = SQL & " WHERE numofert = " & Text1(0).Text
            SQL = "UPDATE " & NombreTabla & " SET " & SQL
            RS.Close
        conn.Execute SQL
        PosicionarData  'vuelvo a cargar los datos
        PonerCampos
        Else
            MsgBox "Error buscando oferta destino: " & Text1(0).Text & ".  EOF", vbExclamation
        End If
    End If
    
    
    Set RS = Nothing

    'Actualizar el Grid con las lineas A�adidas
    If I = 0 Then CalcularDatosFactura   'Si no mete obser y carta que carge los totales
    PonerCamposLineas
    DataGrid1.Enabled = True
    Screen.MousePointer = vbDefault
    
ECargarOferta:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Traer de otra Oferta.", Err.Description
End Sub


Private Sub imgBuscar_Click(Index As Integer)
Dim Indice As Byte

    If Modo = 2 Or Modo = 0 Then Exit Sub
    
    'Trabajador albaran
    If Index = 3 Then
        If Text1(3).Text <> "" Then Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    TerminaBloquear

    Select Case Index
        Case 0 'Cod. Cliente
            HaDevueltoDatos = False
            PonerFoco Text1(4)
            Set frmC = New frmFacClientes
            frmC.DatosADevolverBusqueda = "0"
            frmC.Show vbModal
            Set frmC = Nothing
            Indice = 5
            If HaDevueltoDatos Then
                txtAnterior = ""
                Text1_LostFocus 4
                txtAnterior = Text1(4).Text
            End If
        Case 1 'NIF para cliente de Varios
            Set frmCV = New frmFacClientesV
            frmCV.DatosADevolverBusqueda = "0"
            frmCV.Show vbModal
            Set frmCV = Nothing
            Indice = 6
            
        Case 2 'Cod. Direc.
             'Mostrar las Direc. o Dptos del cliente seleccionado
             If Trim(Text1(4).Text) = "" Then
                MsgBox "Debe seleccionar un cliente.", vbInformation
                Screen.MousePointer = vbDefault
                Exit Sub
             Else
                EsCabecera = False
                MandaBusquedaPrevia " codclien= " & Val(Text1(4).Text)
                Indice = 12
             End If
             
        Case 3 'Realizada Por Trabajador
            Indice = 3
            Set frmT = New frmAdmTrabajadores
            frmT.DatosADevolverBusqueda = "0"
            frmT.Show vbModal
            Set frmT = Nothing
            
        Case 4 'Forma de Pago
            Indice = 14
            PonerFoco Text1(Indice)
            Set frmFP = New frmFacFormasPago
            frmFP.DatosADevolverBusqueda = "0"
            frmFP.Show vbModal
            Set frmFP = Nothing
            
        Case 5 'Agente
            Indice = 17
            PonerFoco Text1(Indice)
            Set frmA = New frmFacAgentesCom
            frmA.DatosADevolverBusqueda = "0"
            frmA.Show vbModal
            Set frmA = Nothing
            
        Case 6 'Cod. Postal
            Set frmCP = New frmCPostal
            frmCP.DatosADevolverBusqueda = "0"
            frmCP.Show vbModal
            Set frmCP = Nothing
            Indice = 9
            VieneDeBuscar = True
    End Select
    PonerFoco Text1(Indice)
    Screen.MousePointer = vbDefault
    
    If Modo = 4 Then
        If Not BLOQUEADesdeFormulario(Me) Then cmdCancelar_Click
    End If
End Sub


Private Sub imgFecha_Click(Index As Integer) 'Abre calendario Fechas
Dim Indice As Byte

   If Modo = 2 Or Modo = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass
   Set frmF = New frmCal
   frmF.Fecha = Now
   Indice = Index + 1
   Me.imgFecha(0).Tag = Index
   
   PonerFormatoFecha Text1(Indice)
   If Text1(Indice).Text <> "" Then frmF.Fecha = CDate(Text1(Indice).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco Text1(Indice)
End Sub


Private Sub mnBuscar_Click()
    Me.SSTab1.Tab = 0
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar lineas de trabajadores
         BotonEliminarLinea
    Else   'Eliminar Trabajador
         BotonEliminar
    End If
End Sub


Private Sub mnGenPedido_Click()
'Pasar una Oferta a Pedido
Dim Devuelve As String

    'Comprobar que hay una Oferta seleccionada
    If Text1(0).Text = "" Then Exit Sub
    
    'Comprobar que la Oferta seleccionada esta aceptada
    Devuelve = DevuelveDesdeBDNew(conAri, NombreTabla, "aceptado", "numofert", Text1(0).Text, "N")
    If Devuelve = "0" Then
        MsgBox "La Oferta debe estar Aceptada para pasar a Pedido."
        Exit Sub
    End If
    If Devuelve = "1" Then
        'Pedir: Operador de Pedido, fecha pedido, y fecha entrega (calcular semana)
'        AbrirListadoOfer (37) '37: Pedir datos para Pedido (NO IMPRIME LISTADO)
        Set frmList = New frmListadoOfer
        frmList.OpcionListado = 37
        frmList.CodClien = Text1(4).Text
        frmList.FecEntre = Text1(2).Text
        frmList.Show vbModal
        Set frmList = Nothing
        
        'Tenemos en CadenaSQL parte de la SELECT para insertar el Pedido
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        PasarOfertaAPedido (CadenaSQL)

        If SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            LimpiarDataGrids
        End If
        Screen.MousePointer = vbDefault
    End If
End Sub


Private Sub mnImpFactProF_Click()
'Imprime factura pro forma
    BotonImprimirProForma (59) '59: Informe Factura ProForma
End Sub

Private Sub mnImpOferta_Click()
'Imprime una Oferta
       frmListadoOfer.NumCod = Text1(0).Text   'N� de Oferta
       frmListadoOfer.FecEntre = Text1(1).Text 'Fecha de Oferta
       If EsHistorico Then
            AbrirListadoOfer (35) '35: Informe Historico de Ofertas
       Else
            AbrirListadoOfer (31) '31: Informe de Ofertas
       End If
End Sub

Private Sub mnImpRecordatorio_Click()
    frmListadoOfer.NumCod = Text1(0).Text
    frmListadoOfer.CodClien = Text1(4).Text
    AbrirListadoOfer (32) '32: Recordatorio de Ofertas
End Sub

Private Sub mnImpValoracion_Click()
    frmListadoOfer.CodClien = Text1(4).Text
    frmListadoOfer.NumCod = Text1(0).Text 'N� de Oferta
    AbrirListadoOfer (33) '33: Valoracion de Ofertas
End Sub

Private Sub mnLineas_Click()
    BotonMtoLineas 0, "Ofertas"
End Sub


Private Sub mnModificar_Click()
    If Modo = 5 Then 'Modificar lineas
         BotonModificarLinea
    Else   'Modificar Trabajador
         If BLOQUEADesdeFormulario(Me) Then BotonModificar
    End If
End Sub


Private Sub mnNuevo_Click()
    If Modo = 5 Then 'A�adir lineas
         BotonAnyadirLinea
    Else 'A�adir Cabecera de Ofertas
         Me.SSTab1.Tab = 0
         BotonAnyadir
    End If
End Sub


Private Sub mnOferta_Click()
'A�adir las lineas de otra oferta a la Oferta
    Set frmTOferta = New frmFacTraerOferta
    frmTOferta.Show vbModal
    Set frmTOferta = Nothing
End Sub

Private Sub mnPlantillas_Click()
'A�adir Plantilla de Oferta
    Set frmPlant = New frmFacCargaPlantilla
    frmPlant.Show vbModal
    Set frmPlant = Nothing
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    If (Modo = 5) Then 'Modo 5: Mto Lineas
        '1:Insertar linea, 2: Modificar
        If ModificaLineas = 1 Or ModificaLineas = 2 Then cmdCancelar_Click
        cmdRegresar_Click
        Exit Sub
    End If
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub


Private Sub Text1_Change(Index As Integer)
    If Index = 9 Then HaCambiadoCP = True 'Cod. Postal
End Sub

'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
Dim cadkey As Integer
    If Index = 33 Then Exit Sub
    txtAnterior = Text1(Index).Text
    cadkey = ObtenerCadKey(Kcampo, Index)
    Kcampo = Index
    If Index = 9 Then HaCambiadoCP = False
    
    If Modo > 2 And Index = 3 Then
        If Text1(3).Text <> "" Then PonerFoco Text1(4)
    Else
        If (Index <> 31 And Index <> 32) Then ConseguirFoco Text1(Index), Modo, cadkey
    End If
    
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Index = 33 Then Exit Sub
    If Index = 30 And KeyCode = 40 Then
        Me.SSTab1.Tab = 2
        PonerFoco Text1(31)
    Else
        If Not Text1(Index).MultiLine Then KEYdown KeyCode
    End If
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not Text1(Index).MultiLine Then
        If KeyAscii = 13 And (Index = 30 Or Index = 32) Then 'ENTER
            If Index = 32 Then
    '            PonerFocoBtn Me.cmdAceptar
            ElseIf Index = 30 Then
                Me.SSTab1.Tab = 2
                PonerFoco Text1(31)
            End If
        Else
            KEYpress KeyAscii
        End If
    End If
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
' Cunado el campo de texto pierde el enfoque
' Es especifico de cada formulario y en el podremos controlar
' lo que queramos, desde formatear un campo si asi lo deseamos
' hasta pedir que nos devuelva los datos de la empresa
'----------------------------------------------------------------
'----------------------------------------------------------------
Private Sub Text1_LostFocus(Index As Integer)
Dim Devuelve As String
       
     If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
 
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    'Por si no ha cambiado nada
    If txtAnterior = Text1(Index).Text Then Exit Sub
    
    'Si queremos hacer algo ..
    Select Case Index
        Case 1, 2 'Fecha Oferta, Fecha Entrega
            If Text1(Index).Text = "" Then Exit Sub
            PonerFormatoFecha Text1(Index)
            If Index = 2 And Text1(Index).Text <> "" Then 'Fecha Entrega
                'Comprobar que es posterior a la del pedido
                If Not EsFechaIgualPosterior(Text1(1).Text, Text1(2).Text, True, "La Fecha de Entrega debe ser posterior a la Fecha del Pedido.") Then
                    Text1(Index).Text = ""
                    PonerFoco Text1(Index)
                    Exit Sub
                End If
            End If
            
        Case 3 'Cod Vendedor
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "straba", "nomtraba")
            Else
                Text2(Index).Text = ""
            End If
            
        Case 4 'Cod. Cliente
            If PonerFormatoEntero(Text1(Index)) Then
                If Modo = 1 Then 'Modo=1 Busqueda
                    Text1(5).Text = PonerNombreDeCod(Text1(Index), conAri, "sclien", "nomclien")
                Else
                    PonerDatosCliente (Text1(Index).Text)
                End If
            Else
                LimpiarDatosCliente
            End If
            
        Case 6 'NIF
'            If Not EsDeVarios Then Exit Sub
'            If Modo = 4 Then 'Modificar
'                'si no se ha modificado el nif del cliente no hacer nada
'                If Text1(6).Text = Data1.Recordset!nifClien Then
'                    Exit Sub
'                End If
'            End If
'            PonerDatosClienteVario (Text1(Index).Text)
        

        Case 9 'Cod. Postal
             If Text1(Index).Locked Then Exit Sub
             If Text1(Index).Text = "" Then
                Text1(Index + 1).Text = ""
                Text1(Index + 2).Text = ""
                Exit Sub
             End If
             If (Not VieneDeBuscar) Or (VieneDeBuscar And HaCambiadoCP) Then
                Text1(Index + 1).Text = ObtenerPoblacion(Text1(Index).Text, Devuelve)
                Text1(Index + 2).Text = Devuelve
            End If
            VieneDeBuscar = False
            
        Case 12 'Cod. Direc
            If Text1(Index).Text = "" Then
                Text2(Index).Text = ""
                Exit Sub
            End If
            Text1(Index).Text = Format(Text1(Index).Text, "000")
            
            'Comprobar que el cliente seleccionada tiene esa direccion
           If PonerDptoEnCliente Then
                'Comprobar que el cliente tiene mantenimientos en esa dired/dpto
                Devuelve = DevuelveDesdeBDNew(conAri, "scaman", "nummante", "codclien", Text1(4).Text, "N", , "coddirec", Text1(12).Text, "N")
                If Devuelve <> "" Then MsgBox "El cliente tiene Mantenimientos.", vbInformation
            Else
                PonerFoco Text1(Index)
            End If
            
        Case 13 'Referencia Obligatoria
            If Trim(Text1(4).Text) <> "" Then
                ComprobarRefObligatoria
            End If
            
        Case 14 'Forma de Pago
            If Me.SSTab1.Tab = 0 Then
                If PonerFormatoEntero(Text1(Index)) Then
                    Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sforpa", "nomforpa")
                Else
                    Text2(Index).Text = ""
                End If
            End If
            
        Case 15, 16 'Descuentos
            If PonerFormatoDecimal(Text1(Index), 4) Then  'Tipo 4: Decimal(4,2)
                If Modo = 4 Then CalcularDatosFactura
            End If
            
        Case 17 'Cod. Agente
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sagent", "nomagent")
            Else
                Text2(Index).Text = ""
            End If
    End Select
End Sub


Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False)
    If chkVistaPrevia = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
Dim Tabla As String
Dim Titulo As String
Dim Desc As String, Devuelve As String
    'Llamamos a al form
    '##A mano
    Cad = ""
    If EsCabecera Then
            Cad = Cad & ParaGrid(Text1(0), 15, "N� Oferta")
            Cad = Cad & ParaGrid(Text1(1), 20, "Fecha Ofer.")
            Cad = Cad & ParaGrid(Text1(4), 15, "Cliente")
            Cad = Cad & ParaGrid(Text1(5), 50, "Nombre Cliente")
            Tabla = NombreTabla
            If EsHistorico Then
                Titulo = "Hist�rico de Ofertas"
                Devuelve = "0|1|"
            Else
                Titulo = "Ofertas"
                Devuelve = "0|"
            End If
'            devuelve = "0|"
    Else 'Llama desde lineas, para cargar solo los depart/direc. del cliente seleccionado
            If vParamAplic.Departamento Then
                Titulo = "Dptos Cliente: "
                Desc = "Dpto."
            Else
                Titulo = "Direc. Cliente: "
                Desc = "Direc."
            End If
            Titulo = Titulo & Text1(4).Text & " - " & Text1(5).Text
            Cad = Cad & "Cod. " & Desc & "|sdirec|coddirec|N||15�"
            Cad = Cad & "Desc. " & Desc & "|sdirec|nomdirec|T||35�"
            Tabla = "sdirec"
            Devuelve = "0|1|"
    End If
           
    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = Tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
'        frmB.vDevuelve = "0|1|"
        frmB.vDevuelve = Devuelve
        frmB.vTitulo = Titulo
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri
        If Not EsCabecera Then frmB.Label1.FontSize = 11
'        frmB.vBuscaPrevia = chkVistaPrevia
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
'        If HaDevueltoDatos Then
'''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'''                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            PonerFoco Text1(kCampo)
        'End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    
    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ning�n registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        If Modo = 1 Then
            Me.cboFacturacion.ListIndex = -1
            PonerFoco Text1(Kcampo)
        End If
        Exit Sub
    Else
        Data1.Recordset.MoveFirst
        PonerModo 2
        PonerCampos
    End If

    Screen.MousePointer = vbDefault
    Exit Sub
    
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCamposLineas()
'Carga las Pesta�as con las tablas de lineas del Trabajador seleccionado para mostrar

    On Error GoTo EPonerLineas

    Screen.MousePointer = vbHourglass

    'Datos de la tabla slipre
    CargaGrid DataGrid1, Data2, True

    Screen.MousePointer = vbDefault
    Exit Sub
    
EPonerLineas:
    MuestraError Err.Number, "PonerCamposLineas"
    PonerModo 2
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
On Error Resume Next

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    Text2(3).Text = PonerNombreDeCod(Text1(3), conAri, "straba", "nomtraba")
    Text2(12).Text = DevuelveDesdeBDNew(conAri, "sdirec", "nomdirec", "codclien", Text1(4).Text, "N", , "coddirec", Text1(12).Text, "N")
    Text2(17).Text = PonerNombreDeCod(Text1(17), conAri, "sagent", "nomagent")
    Text2(14).Text = PonerNombreDeCod(Text1(14), conAri, "sforpa", "nomforpa")
    
    PonerCamposLineas 'Pone los datos de las tablas de lineas de Ofertas
    
    CalcularDatosFactura
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    If Err.Number <> 0 Then Err.Clear
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim I As Byte, NumReg As Byte
Dim b As Boolean

    On Error GoTo EPonerModo

    'Actualiza Iconos Insertar,Modificar,Eliminar
    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    

    
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    cmdRegresar.visible = Modo = 5 And ModificaLineas = 0
    
    
        
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    'Poner Flechas de desplazamiento visibles
    b = (Modo = 2)
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
        
        
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar adem�s limpia los campos Text1
    BloquearText1 Me, Modo
    
    'Poner siempre el campo numOferta (contador) bloqueado, excepto cuando
    'estamos en modo de Busqueda
    b = (Modo <> 1)
    BloquearTxt Text1(0), b, True
    
    '-----  Datos Totales de Factura siempre bloqueado
    For I = 33 To 56
        BloquearTxt Text3(I), True
    Next I
    'Campo B.Imp y Imp. IVA siempre en azul
    Text3(36).BackColor = &HFFFFC0
    For I = 46 To 48
        Text3(I).BackColor = &HFFFFC0
        Text3(I + 6).BackColor = &HFFFFC0
    Next I
    'Campos total Factura en verde
    Text3(55).BackColor = &HC0FFC0
    Text3(56).BackColor = &HC0FFC0    'Tatal factura
    '---------------------------------------------------
    
    b = (Modo = 3) Or (Modo = 4) Or (Modo = 1)
    Me.cboFacturacion.Enabled = b
    Me.chkAceptado(0).Enabled = b
    
    'Si no es modo lineas Boquear los TxtAux
    For I = 0 To txtAux.Count - 1
        BloquearTxt txtAux(I), (Modo <> 5)
    Next I
    BloquearTxt Text2(16), (Modo <> 5)
    
    
    '---------------------------------------------
    b = Modo <> 0 And Modo <> 2 And Modo <> 5
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    For I = 0 To Me.imgFecha.Count - 1
        Me.imgFecha(I).Enabled = b
    Next I
    
    For I = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(I).Enabled = b
    Next I
    Me.imgBuscar(1).visible = False
    
       
    'Modo Linea de Ofertas
    Me.Label1(35).visible = (Modo = 5)
    Me.Text2(16).visible = (Modo = 5)
    BloquearTxt Text2(16), True
       
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
    
    'Poner el tama�o de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
       
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu seg�n modo
    PonerOpcionesMenu 'Activar opciones de menu seg�n nivel de permisos del usuario
    
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de b�squeda o no
'para los campos que permitan introducir criterios m�s largos del tama�o del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Devuelve As String

    On Error GoTo EDatosOK

    DatosOk = False
    b = CompForm(Me, 1) 'Comprobar formato datos ok
    If Not b Then Exit Function
    
    'Comprobar que la Fecha Entrega es posterior a la del pedido
    If Not EsFechaIgualPosterior(Text1(1).Text, Text1(2).Text, True, "La Fecha de Entrega debe ser posterior a la Fecha del Pedido.") Then Exit Function
   
    
    'Comprobar si la referencia del cliente es obligatoria que tenga valor
     If Trim(Text1(4).Text) <> "" Then
        Devuelve = DevuelveDesdeBDNew(1, "sclien", "referobl", "codclien", Text1(4).Text, "N")
        If Devuelve = "1" And Text1(13).Text = "" Then 'Referencia Obligatoria
            MsgBox "La Referencia del Cliente es Obligatoria.", vbInformation
            PonerFoco Text1(13)
            b = False
        End If
    End If
    If Not b Then Exit Function
          
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLinea() As Boolean
Dim b As Boolean
Dim I As Byte
Dim vArtic As CArticulo

    On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    b = True

    


    For I = 0 To txtAux.Count - 1
        If I < 9 Then 'el 9 y el 10
            If txtAux(I).Text = "" Then
                MsgBox "El campo " & txtAux(I).Tag & " no puede ser nulo", vbExclamation
                b = False
                PonerFoco txtAux(I)
                Exit Function
            End If
        End If
    Next I
        
    'Comprobar que existe de el articulo en el almacen seleccionado
    Set vArtic = New CArticulo
    vArtic.Codigo = txtAux(1).Text
    If Not vArtic.ExisteEnAlmacen(txtAux(0).Text) Then
        b = False
        PonerFoco txtAux(1)
    End If
    Set vArtic = Nothing
    
    DatosOkLinea = b
    
EDatosOkLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function



Private Sub Text2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 16 And KeyCode = 40 Then
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYdown KeyCode
    End If
End Sub


Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 16 And KeyAscii = 13 Then 'ENTER
        PonerFocoBtn Me.cmdAceptar
    End If
End Sub

Private Sub Text2_LostFocus(Index As Integer)
    If Index = 16 And (Text2(Index).Locked = False) Then Text2(Index).Text = UCase(Text2(Index).Text)
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Buscar
            mnBuscar_Click
        Case 2  'Todos
            BotonVerTodos
        Case 5  'Nuevo
            mnNuevo_Click
        Case 6  'Modificar
            mnModificar_Click
        Case 7  'Borrar
            mnEliminar_Click
            
        Case 10  'Lineas
            mnLineas_Click
        Case 11 'Generar Pedido
            mnGenPedido_Click
        Case 12 ' Plantillas. Solo visible en Mantenimiento Lineas.
            mnPlantillas_Click
        Case 13 'Traer Lineas de Otra Oferta
            mnOferta_Click
            
        Case 16 'Recordatorio
            mnImpRecordatorio_Click
        Case 17 'Valoracion
            mnImpValoracion_Click
        Case 18 'Imprimir
            mnImpOferta_Click
        Case 19 'Imprimir factura por forma
             mnImpFactProF_Click
        
        Case 22    'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub


Private Sub PonerOpcionesMenu()
Dim J As Byte

    On Error Resume Next
    
    PonerOpcionesMenuGeneral Me
        
    J = Val(Me.mnGenPedido.HelpContextID)
    If J < vUsu.Nivel Then Me.mnGenPedido.Enabled = False
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean
    
    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub
    
    
Private Function InsertarLinea() As Boolean
'Inserta un registro en la tabla de lineas de Ofertas: slipre
Dim SQL As String
Dim numlinea As String
    
    On Error GoTo EInsertarLinea

    InsertarLinea = False
    SQL = ""
    If DatosOkLinea Then 'Lineas de Ofertas
        'Conseguir el siguiente numero de linea
        numlinea = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", ObtenerWhereCP)
        SQL = "INSERT INTO " & NomTablaLineas
        SQL = SQL & "(numofert,numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel, origpre,palets,cajas) "
        SQL = SQL & "VALUES (" & Val(Text1(0).Text) & ", " & numlinea & ", " & Val(txtAux(0).Text) & ","
        SQL = SQL & DBSet(txtAux(1).Text, "T") & ", " & DBSet(txtAux(2).Text, "T") & ", " & DBSet(Text2(16).Text, "T") & ", "
        SQL = SQL & DBSet(txtAux(3).Text, "N") & ", "
        SQL = SQL & DBSet(txtAux(4).Text, "N") & ", " & DBSet(txtAux(6).Text, "N") & ", "
        SQL = SQL & DBSet(txtAux(7).Text, "N") & ", " 'Dto 2
        SQL = SQL & DBSet(txtAux(8).Text, "N") & ", " 'Importe
        SQL = SQL & DBSet(txtAux(5).Text, "T") & ","
        'palets cajas
        SQL = SQL & DBSet(txtAux(9).Text, "N", "S") & "," & DBSet(txtAux(10).Text, "N", "S") & ")"
     End If
    
    If SQL <> "" Then
        conn.Execute SQL
        InsertarLinea = True
    End If
    Exit Function

EInsertarLinea:
    MuestraError Err.Number, "Insertar Lineas Oferta" & vbCrLf & Err.Description
End Function


Private Function InsertarLineaDePlantilla(codartic As String, codAlmac As String, Cantidad As Integer, Precio As String, Dto1 As String, Dto2 As String, OrigP) As Boolean
'Inserta un registro en la tabla de lineas de Ofertas: slipre
Dim SQL As String
Dim numlinea As String
Dim NomArtic As String
Dim Importe As String

    On Error GoTo EInsertarLinea

    InsertarLineaDePlantilla = False
    SQL = ""
    
    'Conseguir el siguiente numero de linea
    numlinea = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", ObtenerWhereCP)
    
    SQL = "INSERT INTO " & NomTablaLineas
    SQL = SQL & " (numofert,numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel, origpre) "
    SQL = SQL & "VALUES (" & Val(Text1(0).Text) & ", " & numlinea & ", "
    SQL = SQL & codAlmac & ", " & DBSet(codartic, "T") & ", "
    NomArtic = DevuelveDesdeBDNew(1, "sartic", "nomartic", "codartic", codartic, "T")
    SQL = SQL & DBSet(NomArtic, "T") & ", " & ValorNulo & ", " & DBSet(Cantidad, "N") & ", "
                   
    Importe = CalcularImporte(CStr(Cantidad), Precio, Dto1, Dto2, vParamAplic.TipoDtos)
    SQL = SQL & DBSet(Precio, "N") & ", "
    SQL = SQL & DBSet(Dto1, "N") & ", "
    SQL = SQL & DBSet(Dto2, "N") & ", "
    SQL = SQL & DBSet(Importe, "N") & ", '"
    SQL = SQL & OrigP & "')"
     
    If SQL <> "" Then
        conn.Execute SQL
        InsertarLineaDePlantilla = True
    End If
    Exit Function
    
EInsertarLinea:
    MuestraError Err.Number, "Insertar Lineas Oferta." & vbCrLf & Err.Description
End Function



Private Function ModificarLinea() As Boolean
'Modifica un registro en la tabla de Revisiones: slima1
Dim SQL As String
    
    On Error GoTo EModificarLinea

    ModificarLinea = False
    SQL = ""
    If DatosOkLinea Then
        SQL = "UPDATE " & NomTablaLineas & " Set codalmac = " & txtAux(0).Text & ", codartic=" & DBSet(txtAux(1).Text, "T") & ", "
        SQL = SQL & "nomartic=" & DBSet(txtAux(2).Text, "T") & ", ampliaci=" & DBSet(Text2(16).Text, "T") & ", "
        SQL = SQL & " cantidad = " & DBSet(txtAux(3).Text, "N", "N") & ", precioar = " & DBSet(txtAux(4).Text, "N", "N") & ", "
        SQL = SQL & "dtoline1= " & DBSet(txtAux(6).Text, "N", "N") & ", dtoline2= " & DBSet(txtAux(7).Text, "N", "N") & ", "
        SQL = SQL & " importel = " & DBSet(txtAux(8).Text, "N") & ", origpre=" & DBSet(txtAux(5).Text, "T")
        SQL = SQL & " , palets = " & DBSet(txtAux(9).Text, "N", "S") & ", cajas=" & DBSet(txtAux(10).Text, "N", "S")
        SQL = SQL & " WHERE " & ObtenerWhereCP & " AND numlinea=" & Data2.Recordset!numlinea
    End If

    If SQL <> "" Then
        conn.Execute SQL
        ModificarLinea = True
    End If
    Exit Function
    
EModificarLinea:
    MuestraError Err.Number, "Modificar Lineas Oferta" & vbCrLf & Err.Description
End Function



Private Sub PonerBotonCabecera(b As Boolean)
'Pone el boton de Regresar a la Cabecera si pasamos a MAntenimiento de Lineas
'o Pone los botones de Aceptar y cancelar en Insert,update o delete lineas
On Error Resume Next
    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    Me.cmdRegresar.visible = b
    Me.cmdRegresar.Caption = "Cabecera"
    If b Then
        Me.lblIndicador.Caption = "L�neas " & TituloLinea
        PonerFocoBtn Me.cmdRegresar
    End If
    'Habilitar las opciones correctas del menu
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
Dim b As Boolean
Dim SQL As String

    On Error GoTo ECargaGrid

    b = DataGrid1.Enabled
    
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral vDataGrid, vData, SQL, PrimeraVez

    
    CargaGrid2 vDataGrid, vData
    vDataGrid.ScrollBars = dbgAutomatic
        
    b = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2)
    vDataGrid.Enabled = Not b

    PrimeraVez = False
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim I As Integer

    On Error GoTo ECargaGrid

    vData.Refresh

    vDataGrid.Columns(0).visible = False
    vDataGrid.Columns(1).visible = False

    Select Case vDataGrid.Name
        Case "DataGrid1" 'Cod. Almacen
                vDataGrid.Columns(2).Caption = "Alm."
                vDataGrid.Columns(2).Width = 500
                vDataGrid.Columns(2).NumberFormat = "000"
                
                vDataGrid.Columns(3).Caption = "Articulo"
                vDataGrid.Columns(3).Width = 1700
                
                vDataGrid.Columns(4).Caption = "Desc. Art�culo"
                vDataGrid.Columns(4).Width = 3300
                
'                vDataGrid.Columns(5).Caption = "Ampl. L�nea"
'                vDataGrid.Columns(5).Width = 7980
                vDataGrid.Columns(5).visible = False
                
                
                vDataGrid.Columns(6).Caption = "Palets"
                vDataGrid.Columns(6).Width = 850
                vDataGrid.Columns(6).Alignment = dbgRight
                vDataGrid.Columns(6).NumberFormat = "0"
                vDataGrid.Columns(7).Caption = "Cajas"
                vDataGrid.Columns(7).Width = 850
                vDataGrid.Columns(7).Alignment = dbgRight
                vDataGrid.Columns(7).NumberFormat = "0"
                
                
                
                vDataGrid.Columns(8).Caption = "Cantidad"
                vDataGrid.Columns(8).Width = 850
                vDataGrid.Columns(8).Alignment = dbgRight
                vDataGrid.Columns(8).NumberFormat = FormatoImporte
                
                vDataGrid.Columns(9).Caption = "Precio"
                vDataGrid.Columns(9).Width = 1000
                vDataGrid.Columns(9).Alignment = dbgRight
                vDataGrid.Columns(9).NumberFormat = FormatoPrecio
                
                vDataGrid.Columns(10).Caption = "OP"
                vDataGrid.Columns(10).Width = 350
                vDataGrid.Columns(10).Alignment = dbgCenter
                
                
                vDataGrid.Columns(11).Caption = "Dto. 1"
                vDataGrid.Columns(11).Width = 600
                vDataGrid.Columns(11).Alignment = dbgRight
                vDataGrid.Columns(11).NumberFormat = FormatoDescuento
                
                vDataGrid.Columns(12).Caption = "Dto. 2"
                vDataGrid.Columns(12).Width = 600
                vDataGrid.Columns(12).Alignment = dbgRight
                vDataGrid.Columns(12).NumberFormat = FormatoDescuento
                
                vDataGrid.Columns(13).Caption = "Importe L�nea"
                vDataGrid.Columns(13).Width = 1400
                vDataGrid.Columns(13).Alignment = dbgRight
                vDataGrid.Columns(13).NumberFormat = FormatoImporte
    End Select

    For I = 0 To vDataGrid.Columns.Count - 1
        vDataGrid.Columns(I).Locked = True
        vDataGrid.Columns(I).AllowSizing = False
    Next I
    vDataGrid.HoldFields
    Exit Sub
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


'Esta funcion sustituye a LlamaLineas
Private Sub CargaTxtAux(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posici�n adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single
Dim I As Byte

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For I = 0 To txtAux.Count - 1 'TextBox
            txtAux(I).Top = 290
            txtAux(I).visible = visible
        Next I
'        txtAux2.visible = visible
        cmdAux(0).visible = visible
        cmdAux(1).visible = visible
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid1
            For I = 0 To txtAux.Count - 1
                txtAux(I).Text = ""
                BloquearTxt txtAux(I), False
            Next I
        Else 'Vamos a modificar
            For I = 0 To txtAux.Count - 1
                
                If I < 3 Then
                    'Los primeros
                    txtAux(I).Text = DataGrid1.Columns(I + 2).Text
                ElseIf I < 9 Then
                    
                    txtAux(I).Text = DataGrid1.Columns(I + 5).Text
                    
                Else
                    txtAux(I).Text = DataGrid1.Columns(I - 3).Text
                End If
                txtAux(I).Locked = False
            Next I
        End If
        
        'El Campo de Origen del precio se actualiza por programa al modificar el precio
        BloquearTxt txtAux(5), True
        'El campo Importe es calculado y lo bloqueamos.
        BloquearTxt txtAux(8), True
    

        'Fijamos altura(Height) y posici�n Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid1, 10)
        
        For I = 0 To txtAux.Count - 1
            txtAux(I).Top = alto
            txtAux(I).Height = DataGrid1.RowHeight
        Next I
        cmdAux(0).Top = alto
        cmdAux(1).Top = alto
        cmdAux(0).Height = DataGrid1.RowHeight
        cmdAux(1).Height = DataGrid1.RowHeight
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'Cod. Almac
        txtAux(0).Left = DataGrid1.Left + 330
        txtAux(0).Width = DataGrid1.Columns(2).Width - 160
        cmdAux(0).Left = txtAux(0).Left + txtAux(0).Width - 40
        'Cod Artic
        txtAux(1).Left = cmdAux(0).Left + cmdAux(0).Width + 20
        txtAux(1).Width = DataGrid1.Columns(3).Width - 160
        cmdAux(1).Left = txtAux(1).Left + txtAux(1).Width - 50
        'Nom Artic
        txtAux(2).Left = cmdAux(1).Left + cmdAux(1).Width
        txtAux(2).Width = DataGrid1.Columns(4).Width - 10
        'Cantidad
        'txtAux(3).Left = txtAux(2).Left + txtAux(2).Width + 10
        'txtAux(3).Width = DataGrid1.Columns(6).Width - 10
        
        'Palets
        txtAux(9).Left = txtAux(2).Left + txtAux(2).Width + 10
        txtAux(9).Width = DataGrid1.Columns(6).Width - 10
        txtAux(10).Left = txtAux(9).Left + txtAux(9).Width + 10
        txtAux(10).Width = DataGrid1.Columns(7).Width - 10
        
        'Cantidad
        txtAux(3).Left = txtAux(10).Left + txtAux(10).Width + 10
        txtAux(3).Width = DataGrid1.Columns(8).Width - 10
        
        'Precio, Dto1, Dto2, Precio
        For I = 4 To 8   'txtAux.Count - 1
            txtAux(I).Left = txtAux(I - 1).Left + txtAux(I - 1).Width + 10
            txtAux(I).Width = DataGrid1.Columns(I + 5).Width - 10
        Next I
        
        'Los ponemos Visibles o No
        '--------------------------
        For I = 0 To txtAux.Count - 1
            txtAux(I).visible = visible
        Next I
        cmdAux(0).visible = visible
        cmdAux(1).visible = visible
    End If
End Sub


Private Sub TxtAux_Change(Index As Integer)
    'Precio y Modo Borrar Lineas
    If Index = 4 And ModificaLineas = 2 Then txtAux(5).Text = "M"
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
Dim cadkey As Integer

    cadkey = ObtenerCadKey(Kcampo, Index)
    Kcampo = Index
'    ConseguirFoco txtAux(Index), Modo, cadkey
    ConseguirFocoLin txtAux(Index), cadkey
End Sub

Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Not (Index = 0 And KeyCode = 38) Then KEYdown KeyCode
End Sub


Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim Devuelve As String, cadMen As String
Dim codTarif As String
Dim CPrecioFact As CPreciosFact
Dim vCStock As cStock
Dim NumCajas As Long, RestoUnid As Long
Dim OrigP As String 'De donde viene el precio
Dim Cantidad As String
Dim b As Boolean

    'Quitar espacios en blanco
    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub
    'txtAux(Index).Text = Trim(txtAux(Index))
    If txtAux(Index).Text = "" And (Index <> 1 And Index <> 4) Then Exit Sub
    'If Screen.ActiveForm.ActiveControl.Name = "cmdCancelar" Then Exit Sub
    
    
    Select Case Index
        Case 0 'Cod Almacen
            'Comprobar que existe el almacen
            Devuelve = PonerAlmacen(txtAux(Index).Text)
            txtAux(Index).Text = Devuelve
            If Devuelve = "" Then PonerFoco txtAux(Index)
            
        Case 1 'Cod. Articulo
            If txtAux(1).Text = "" Then
                txtAux(2).Text = ""
                Exit Sub
            End If
            
            If txtAux(0).Text = "" Then
                MsgBox "Debe seleccionar un almacen.", vbInformation
                PonerFoco txtAux(0)
                Exit Sub
            End If
            
            Devuelve = ""
            If ModificaLineas = 2 Then
                If Not Data2.Recordset.EOF Then Devuelve = Data2.Recordset!codartic
            End If
            
            If PonerArticulo(txtAux(1), txtAux(2), txtAux(0).Text, CodTipoMov, ModificaLineas, Devuelve) Then
                b = (Me.ActiveControl.Name = "txtAux")
                If b Then b = (Me.ActiveControl.Index = 0)
                
                If Not b Then
                    If txtAux(2).Locked Then
'                        PonerFoco txtAux(3)
                    Else
                        PonerFoco txtAux(2)
                    End If
                Else
                    PonerFoco txtAux(0)
                End If
            Else
                PonerFoco txtAux(Index)
            End If
'
'             'Si es articulo de varios podemos modificar la descripci�n del articulo, sino bloqueamos.
'            If Not EsArticuloVarios(txtAux(Index).Text) Then
'                BloquearTxt txtAux(2), True
'                PonerFoco txtAux(3)
'            Else
'                BloquearTxt txtAux(2), False
'                PonerFoco txtAux(2)
'            End If
        
        Case 2 'desc articulo
            If txtAux(Index).Locked = False Then txtAux(Index).Text = UCase(txtAux(Index).Text)
        
        Case 3 'CANTIDAD
            If PonerFormatoDecimal(txtAux(Index), 1) Then  'Tipo 1: Decimal(12,2)
                'Comprobar si hay suficiente stock
                b = True
                Set vCStock = New cStock
                If Not InicializarCStock(vCStock, "S") Then b = False
                If vCStock.MueveStock Then
                    If Not vCStock.MoverStock(False) Then b = False
                End If
                If Not b Then
                    Set vCStock = Nothing
                    Exit Sub
                End If
                
                b = False
                If Modo = 5 Then 'Modo lineas
                    If ModificaLineas = 1 Then 'insertar linea
                        b = True
                    ElseIf ModificaLineas = 2 Then 'modificar linea
                        If Data2.Recordset!codartic <> txtAux(1).Text Then b = True
                    End If
                End If
                
                If b Then 'Modo Insertar en Mto Lineas
                    'Obtener el precio correspondiente y los descuentos
                    'Comprobar si el articulo se vende por cajas antes de entrar a la funci�n
                    Devuelve = DevuelveDesdeBDNew(conAri, "sartic", "unicajas", "codartic", txtAux(1).Text, "T")
                    If Devuelve <> "" Then
                        'Si se puede vender por cajas(devuelve>1) poner numero de cajas en una linea con el
                        'precio de caja, y otra linea con el resto unidades un precio unidad
                        Cantidad = txtAux(Index).Text
                        NumCajas = ObtenerNumCajas(Cantidad, Devuelve)
                        RestoUnid = CLng(Cantidad) - NumCajas * CLng(Devuelve)
            
                        codTarif = DevuelveDesdeBDNew(conAri, "sclien", "codtarif", "codclien", Text1(4).Text, "N")
                        Set CPrecioFact = New CPreciosFact
                        CPrecioFact.CodigoLista = codTarif
                        CPrecioFact.CodigoArtic = txtAux(1).Text
                        CPrecioFact.CodigoClien = Text1(4).Text
                        PorCaja = (NumCajas > 0)
                        Precio = CPrecioFact.ObtenerPrecio(PorCaja, Text1(1).Text, OrigP)
                        
                        'Si PorCaja vuelve de ObtenerPrecio a false se calcula con precio unidad aunque NumCajas>0
                        'Ya que a regresado con pvp del Articulo
                        If PorCaja And NumCajas > 0 And RestoUnid > 0 Then
                            cadMen = "El Art�culo puede venderse por Cajas (" & Devuelve & "uds. por Caja)." & vbCrLf
                            cadMen = cadMen & vbCrLf & "Inserte dos Lineas:   "
                            cadMen = cadMen & vbCrLf & "   Linea 1:  " & NumCajas * CInt(Devuelve) & " uds a Precio Caja"
                            cadMen = cadMen & vbCrLf & "   Linea 2:  " & CInt(Cantidad) - NumCajas * CInt(Devuelve) & " uds a Precio Unidad"
                            MsgBox cadMen, vbInformation
    '                        TxtAux(3).Text = NumCajas * CInt(devuelve)
                            PonerFoco txtAux(Index)
                        Else
                            If (txtAux(4).Text = "") Or (txtAux(4).Text <> "" And ModificaLineas = 2 And b) Then
                                txtAux(4).Text = Precio
                                txtAux(5).Text = OrigP 'De donde viene el precio
                            End If
                            PonerFormatoDecimal txtAux(4), 2
                            If txtAux(6).Text = "" Then txtAux(6).Text = CPrecioFact.Descuento1
                            PonerFormatoDecimal txtAux(6), 4
                            If txtAux(7).Text = "" Then txtAux(7).Text = CPrecioFact.Descuento2
                            PonerFormatoDecimal txtAux(7), 4
                            PonerFoco txtAux(4)
                            ConseguirFocoLin txtAux(4)
    '                            PonerFoco Text2(16)
                        End If
                        Set CPrecioFact = Nothing
                    End If
                End If 'modo 5
                Set vCStock = Nothing
            End If 'formato decimal
            
        Case 4 'Precio
            If txtAux(Index).Text <> "" Then
                PonerFormatoDecimal txtAux(Index), 2 'Tipo 2: Decimal(10,4)
                If ModificaLineas = 1 Then
                    If CSng(txtAux(Index).Text) <> CSng(ComprobarCero(Precio)) Then txtAux(5).Text = "M"
                End If
            End If
            
        Case 6, 7 'Descuentos
            PonerFormatoDecimal txtAux(Index), 4 'Tipo 4: Decimal(4,2)
            If Index = 7 Then PonerFocoBtn Me.cmdAceptar
        Case 8 'Importe Linea
            PonerFormatoDecimal txtAux(Index), 1 'Tipo 3: Decimal(12,2)
            
        Case 9, 10
            If Not PonerFormatoEntero(txtAux(Index)) Then
                txtAux(Index).Text = ""
                PonerFoco txtAux(Index)
                Exit Sub
            End If
            
            If txtAux(1).Text = "" Then Exit Sub
            
            'Si que es numero entero
            If Index = 9 Then
                'Ya hay puestas cajas
                If txtAux(10).Text <> "" Then Exit Sub
            Else
                'Ya estan puestas las uds
                If txtAux(3).Text <> "" Then Exit Sub
            End If
            
            If Index = 9 Then
                Devuelve = DevuelveDesdeBD(conAri, "pal_udbas * pal_udalt", "sarti4", "codartic", txtAux(1).Text)
                NumCajas = 10 'sera el txtaux
            Else
                Devuelve = DevuelveDesdeBD(conAri, "unicajas", "sartic", "codartic", txtAux(1).Text)
                NumCajas = 3 'sera el txtaux
            End If
            If Devuelve = "" Then Devuelve = "0"
            RestoUnid = Val(Devuelve) * CInt(txtAux(Index).Text)
            txtAux(NumCajas).Text = RestoUnid
    End Select
    
    If (Index = 3 Or Index = 4 Or Index = 6 Or Index = 7) Then
        If txtAux(1).Text = "" Then Exit Sub
        txtAux(8).Text = CalcularImporte(txtAux(3).Text, txtAux(4).Text, txtAux(6).Text, txtAux(7).Text, vParamAplic.TipoDtos)
        PonerFormatoDecimal txtAux(8), 1
        
    End If
End Sub


Private Sub BotonMtoLineas(numTab As Integer, Cad As String)
        Me.SSTab1.Tab = numTab
        TituloLinea = Cad
        ModificaLineas = 0
        PonerModo 5
        PonerBotonCabecera True
End Sub


Private Function Eliminar() As Boolean
Dim SQL As String

    On Error GoTo FinEliminar

    conn.BeginTrans
    SQL = " WHERE  numofert=" & Data1.Recordset!NumOfert

    'Lineas de Ofertas
    conn.Execute "Delete from " & NomTablaLineas & SQL
    
    'Cabecera
    conn.Execute "Delete from " & NombreTabla & SQL

FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
        conn.RollbackTrans
        Eliminar = False
    Else
        conn.CommitTrans
        Eliminar = True
    End If
End Function


Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ning�n registro
On Error Resume Next
    CargaGrid DataGrid1, Data2, False
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = "(" & ObtenerWhereCP & ")"
         If SituarData(Data1, vWhere, Indicador) Then
             PonerModo 2
             lblIndicador.Caption = Indicador
        Else
             LimpiarCampos
             'Poner los grid sin apuntar a nada
             LimpiarDataGrids
             PonerModo 0
         End If
    Else
        'El Data esta vacio, desde el modo de inicio se pulsa Insertar
        CadenaConsulta = "Select * from " & NombreTabla & " WHERE " & ObtenerWhereCP & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Function ObtenerWhereCP() As String
On Error Resume Next

    ObtenerWhereCP = " numofert= " & Text1(0).Text
End Function


Private Function MontaSQLCarga(enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Bas�ndose en la informaci�n proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim SQL As String
    
    SQL = "SELECT numofert, numlinea, codalmac, codartic, nomartic, ampliaci,palets,cajas, cantidad, precioar, origpre, dtoline1, dtoline2,importel "
    SQL = SQL & " FROM " & NomTablaLineas
    If enlaza Then
        SQL = SQL & " WHERE " & ObtenerWhereCP
        If EsHistorico Then SQL = SQL & " and fecofert='" & Format(Text1(1).Text, FormatoFecha) & "'"
    Else
        SQL = SQL & " WHERE numofert = -1"
    End If
    SQL = SQL & " Order by numofert, numlinea"
    MontaSQLCarga = SQL
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar seg�n el modo en que estemos
Dim b As Boolean, bol As Boolean
Dim I As Byte

    'Si visualizamos el historico no mostrar botones de Mantenimiento, solo es consulta
    For I = 5 To 17
        Toolbar1.Buttons(I).visible = Not EsHistorico
    Next I
    Me.mnNuevo.visible = Not EsHistorico
    Me.mnModificar.visible = Not EsHistorico
    Me.mnEliminar.visible = Not EsHistorico
    Me.mnLineas.visible = Not EsHistorico
    Me.mnGenPedido.visible = Not EsHistorico
    Me.mnPlantillas.visible = Not EsHistorico
    Me.mnOferta.visible = Not EsHistorico 'Traer de Oferta
    Me.mnImpRecordatorio.visible = Not EsHistorico
    Me.mnImpValoracion.visible = Not EsHistorico
    Me.mnBarra2.visible = Not EsHistorico
    Me.mnBarra3.visible = Not EsHistorico
    Me.mnBarra4.visible = Not EsHistorico
    
    Me.Toolbar1.Buttons(19).Enabled = Not EsHistorico
    Me.mnImpFactProF.Enabled = Not EsHistorico
    
    If Not EsHistorico Then
        b = (Modo = 2) Or (Modo = 5 And ModificaLineas = 0)
        'Insertar
        Toolbar1.Buttons(5).Enabled = (b Or Modo = 0)
        Me.mnNuevo.Enabled = (b Or Modo = 0)
        'Modificar
        Toolbar1.Buttons(6).Enabled = b
        Me.mnModificar.Enabled = b
        'eliminar
        Toolbar1.Buttons(7).Enabled = b
        Me.mnEliminar.Enabled = b
            
        b = (Modo = 2)
        'Mantenimiento lineas
        Toolbar1.Buttons(10).Enabled = b
        Me.mnLineas.Enabled = b
        'Generar Pedido
        Toolbar1.Buttons(11).Enabled = b
        Me.mnGenPedido.Enabled = b
        
        
        b = (Modo = 5) And (ModificaLineas = 0)
        'Plantillas
        Toolbar1.Buttons(12).visible = b
        Toolbar1.Buttons(12).Enabled = b
        Me.mnPlantillas.visible = b
        Me.mnPlantillas.Enabled = b
        'Traer Lineas de Otra Oferta
        Toolbar1.Buttons(13).visible = b
        Toolbar1.Buttons(13).Enabled = b
        Me.mnOferta.visible = b
        Me.mnOferta.Enabled = b
        
        'Recordatorio
        b = (Modo = 2)
        bol = (Modo <> 5)
        Toolbar1.Buttons(16).visible = bol
        Toolbar1.Buttons(16).Enabled = b
        Me.mnImpRecordatorio.visible = bol
        Me.mnImpRecordatorio.Enabled = b
        'Valoraci�n
        Toolbar1.Buttons(17).visible = bol
        Toolbar1.Buttons(17).Enabled = b
        Me.mnImpValoracion.visible = bol
        Me.mnImpValoracion.Enabled = b
        
        b = (Modo >= 3) Or Modo = 1
        'Buscar
        Toolbar1.Buttons(1).Enabled = Not b
        Me.mnBuscar.Enabled = Not b
        'Ver Todos
        Toolbar1.Buttons(2).Enabled = Not b
        Me.mnVerTodos.Enabled = Not b
    End If
End Sub



Private Sub CargarComboFacturacion()
'### Combo Tipo Facturaci�n
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Factura Colectiva, 1-Factura x Albaran

    cboFacturacion.Clear
    cboFacturacion.AddItem "Factura Colectiva"
    cboFacturacion.ItemData(cboFacturacion.NewIndex) = 0

    cboFacturacion.AddItem "Factura x Albaran"
    cboFacturacion.ItemData(cboFacturacion.NewIndex) = 1

End Sub


Private Function InsertarOferta(vSQL As String, vTipoMov As CTiposMov) As Boolean
Dim MenError As String
Dim bol As Boolean, Existe As Boolean
Dim Devuelve As String
Dim cambiaSQL As Boolean

    On Error GoTo EInsertarOferta
    
    bol = True
    
    cambiaSQL = False
    'Comprobar si mientras tanto se incremento el contador de Ofertas
    'para ello vemos si existe una oferta con ese contador y si existe la incrementamos
    Do
        Devuelve = DevuelveDesdeBDNew(conAri, NombreTabla, "numofert", "numofert", Text1(0).Text, "N")
        If Devuelve <> "" Then
            'Ya existe el contador incrementarlo
            Existe = True
            vTipoMov.IncrementarContador (CodTipoMov)
            Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
            cambiaSQL = True
        Else
            Existe = False
        End If
    Loop Until Not Existe
    If cambiaSQL Then vSQL = CadenaInsertarDesdeForm(Me)
    
    
    'Aqui empieza transaccion
    conn.BeginTrans
    MenError = "Error al insertar en la tabla Cabecera de Ofertas (scapre)."
    conn.Execute vSQL, , adCmdText
    
    'Actualizar los datos del cliente si es de varios
    EsDeVarios = EsClienteVarios(Text1(4).Text)
    If EsDeVarios Then
'        MenError = "Error al actualizar el Cliente de Varios (sclvar)."
        MenError = "Modificando datos cliente varios"
        bol = ActualizarClienteVarios(Text1(4).Text, Text1(6).Text)
    End If
    
    MenError = "Error al actualizar el contador de la Oferta."
'    bol = vTipoMov.IncrementarContador("REG")
    vTipoMov.IncrementarContador (CodTipoMov)

EInsertarOferta:
        If Err.Number <> 0 Then
            MenError = "Insertando Oferta." & vbCrLf & "----------------------------" & vbCrLf & MenError
            MuestraError Err.Number, MenError, Err.Description
            bol = False
        End If
        If bol Then
            conn.CommitTrans
            InsertarOferta = True
        Else
            conn.RollbackTrans
            InsertarOferta = False
        End If
End Function


Private Sub LimpiarDatosCliente()
Dim I As Byte

    For I = 4 To 13
        Text1(I).Text = ""
    Next I
    If Modo = 3 Then
        For I = 14 To 17
            Text1(I).Text = ""
        Next I
        Text2(12).Text = ""
        Text2(14).Text = ""
        Text2(17).Text = ""
        Me.cboFacturacion.ListIndex = -1
    End If
End Sub
    

Private Function ObtenerNumCajas(TUnidades As String, UniCaja As String) As Integer
Dim NumCajas As Integer
Dim Cantidad As Integer, UniPorCaja As Integer

    On Error Resume Next

    Cantidad = CInt(TUnidades)
    UniPorCaja = CInt(UniCaja)
    If UniPorCaja > 1 Then 'Se vende en cajas
        NumCajas = Int(Cantidad / UniPorCaja)
    Else 'No se vende por cajas
        NumCajas = 0
    End If
    ObtenerNumCajas = NumCajas
End Function


Private Function DescargarDatosTMP()
'Al salir de la aplicacion se borran los datos de la tabla temporal
Dim SQL As String

    On Error GoTo EDescargaDatos

    '------------- AHORA
    SQL = "DELETE from tmpscapla" & " where codusu= " & vUsu.Codigo
    conn.Execute SQL
    Exit Function
    
EDescargaDatos:
        MuestraError Err.Number, "Descargar Tabla Temporal", Err.Description
End Function



Private Function InsertarPedido(cadSQL As String, MenError As String, numPed As String) As Boolean
'Devuelve el mensane de error si se produce
'OUT -> numPed: N� Pedido que inserta
'Dim cadError As String
Dim bol As Boolean, Existe As Boolean
Dim Devuelve As String
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim Codtipom As String
Dim vSQL As String

    On Error GoTo EInsertarPedido
    
    bol = False
    InsertarPedido = bol
    
    'Obtener el Contador de PEDIDO
    Codtipom = "PEV"
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(Codtipom) Then
        'Comprobar si mientras tanto se incremento el contador de Pedidos
        'para ello vemos si existe una oferta con ese contador y si existe la incrementamos
        Do
            numPed = vTipoMov.ConseguirContador(Codtipom)
            Devuelve = DevuelveDesdeBDNew(1, "scaped", "numpedcl", "numpedcl", numPed, "N")
            If Devuelve <> "" Then
                'Ya existe el contador incrementarlo
                Existe = True
                vTipoMov.IncrementarContador (Codtipom)
                numPed = vTipoMov.ConseguirContador(Codtipom)
            Else
                Existe = False
            End If
        Loop Until Not Existe
            
    Else 'No existe el tipo de Movimiento
        Set vTipoMov = Nothing
        Exit Function
    End If
    

    'Acabar la sql con el contador seleccionado
    vSQL = "INSERT INTO scaped (numpedcl,fecpedcl,fecentre,sementre,visadore,codclien,nomclien,domclien,codpobla,pobclien,proclien,nifclien,telclien,coddirec, nomdirec, referenc,codtraba,codagent, codforpa, dtoppago, dtognral, tipofact,"
    vSQL = vSQL & "observa01, observa02, observa03, observa04, observa05,servcomp,restoped,numofert,fecofert,observa6)"
    vSQL = vSQL & " SELECT " & numPed & " as numpedcl, " & cadSQL
    vSQL = vSQL & " FROM " & NombreTabla & " WHERE numofert=" & Text1(0).Text

    'Insertar Cabecera
    MenError = "Error al insertar en la tabla Cabecera de Pedidos (scaped )."
    conn.Execute vSQL, , adCmdText
    
    'Insertar Lineas Pedido
    MenError = "Error al insertar en la tabla Lineas de Pedido (sliped)."
    If Not InsertarLineasPedido(numPed) Then Exit Function
    
    MenError = "Error al actualizar el contador del Pedido."
'    bol = vTipoMov.IncrementarContador("REG")
    bol = vTipoMov.IncrementarContador(Codtipom)
    Set vTipoMov = Nothing
    'bol = True
    
EInsertarPedido:
        If Err.Number <> 0 Then bol = False
        InsertarPedido = bol
End Function


Private Sub PasarOfertaAPedido(vSQL As String)
Dim bol As Boolean
Dim MenError As String
Dim numPed As String

    On Error GoTo EGenPedido

    bol = False
        
    'Aqui empieza transaccion
    conn.BeginTrans
    'Insertar en tablas de Pedido la Oferta
    bol = InsertarPedido(vSQL, MenError, numPed)
    If bol Then 'Si se inserta Pedido
       'Pasar la Oferta al Historico de Oferta y Borrarla de Ofertas
       vSQL = " scapre.numofert= " & Text1(0).Text
       bol = ActualizarElTraspaso(MenError, vSQL, "OFE")
    End If
    
EGenPedido:
    If Err.Number <> 0 Then
        MenError = "Pasando Oferta a Pedido." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
        bol = False
    End If
    If bol Then
        conn.CommitTrans
        Screen.MousePointer = vbDefault
        MsgBox "La Oferta de Venta N�: " & Text1(0).Text & vbCrLf & vbCrLf & "ha generado el Pedido N�: " & Format(numPed, "0000000")
    Else
        MsgBox MenError, vbExclamation
        conn.RollbackTrans
    End If
End Sub


Private Function InsertarLineasPedido(NumPedido As String) As Boolean
Dim SQL As String

    On Error Resume Next

    'Insertar en la tabla de Pedido, los registros seleccionados de la tabla de Ofertas
    SQL = ""
    SQL = "SELECT " & NumPedido & " as numpedcl, numlinea, codalmac,"
    SQL = SQL & NomTablaLineas & ".codartic, " & NomTablaLineas & ".nomartic, ampliaci, "
    SQL = SQL & "cantidad, " & "0 as servidas, precioar, dtoline1, dtoline2, importel, origpre "
    'Cajas y preciolitro.  YA esta el campo cajas
    '''SQL = SQL & ",IF(unicajas=0,1,cantidad div unicajas)  ,if(LitrosUnidad=0,precioar,round( (precioar/LitrosUnidad),2))"
    SQL = SQL & ",cajas  ,if(LitrosUnidad=0,precioar,round( (precioar/LitrosUnidad),2)),palets"
    SQL = SQL & " FROM " & NomTablaLineas & ",sartic WHERE " & NomTablaLineas & ".codartic=sartic.codartic "
    SQL = SQL & " AND numofert=" & Text1(0).Text
    SQL = "insert into `sliped` (`numpedcl`,`numlinea`,`codalmac`,`codartic`,`nomartic`,`ampliaci`,`cantidad`," & _
        "`servidas`,`precioar`,`dtoline1`,`dtoline2`,`importel`,`origpre`,`cajas`,`PrecioLitro`,`cajserv`) " & SQL

    
    conn.Execute SQL
    
        
    If Err.Number <> 0 Then
         'Hay error , almacenamos y salimos
        InsertarLineasPedido = False
    Else
        InsertarLineasPedido = True
    End If
End Function


Private Function InicializarCStock(ByRef vCStock As cStock, TipoM As String, Optional numlinea As String) As Boolean
'On Error Resume Next
On Error Resume Next

    vCStock.tipoMov = TipoM
    vCStock.DetaMov = CodTipoMov
    vCStock.Trabajador = CLng(Text1(4).Text) 'guardamos el cliente de la oferta
    vCStock.Documento = Text1(0).Text 'N� de oferta
    vCStock.Fechamov = Text1(1).Text 'Fecha oferta
    If ModificaLineas = 1 Or ModificaLineas = 2 Then '1=Insertar, 2=Modificar
        vCStock.codartic = txtAux(1).Text
        vCStock.codAlmac = CInt(txtAux(0).Text)
        vCStock.Cantidad = CSng(ComprobarCero(txtAux(3).Text))
        vCStock.Importe = CCur(ComprobarCero(txtAux(8).Text))
    Else
        vCStock.codartic = Data2.Recordset!codartic
        vCStock.codAlmac = CInt(Data2.Recordset!codAlmac)
        vCStock.Cantidad = CSng(Data2.Recordset!Cantidad)
        vCStock.Importe = CCur(Data2.Recordset!ImporteL)
    End If
    If ModificaLineas = 1 Then '1=Insertar Linea
         vCStock.LineaDocu = CInt(ComprobarCero(numlinea))
    Else
        vCStock.LineaDocu = CInt(Data2.Recordset!numlinea)
    End If
    
    If Err.Number <> 0 Then
        MsgBox "No se han podido inicializar la clase para actualizar Stock", vbExclamation
        InicializarCStock = False
    Else
        InicializarCStock = True
    End If
End Function



Private Sub PonerDatosCliente(CodClien As String, Optional nifClien As String)
Dim vCliente As CCliente
Dim Observaciones As String
    
    On Error GoTo EPonerDatos
    
    If CodClien = "" Then
        LimpiarDatosCliente
        Exit Sub
    End If

    Set vCliente = New CCliente
    'si se ha modificado el cliente volver a cargar los datos
    If vCliente.Existe(CodClien) Then
        If vCliente.LeerDatos(CodClien) Then
            'si el cliente esta bloqueado salimos
            If vCliente.ClienteBloqueado Then
                LimpiarDatosCliente
                Set vCliente = Nothing
                Exit Sub
            End If
            
'            EsDeVarios = vCliente.EsClienteVarios(Text1(4).Text)
            EsDeVarios = vCliente.DeVarios
            BloquearDatosCliente (EsDeVarios)
        
            If Modo = 4 And EsDeVarios Then 'Modificar
                'si no se ha modificado el cliente no hacer nada
                If CLng(Text1(4).Text) = CLng(Data1.Recordset!CodClien) Then
                    Set vCliente = Nothing
                    Exit Sub
                End If
            End If
        
'            If Actualizar = False And EsDeVarios = False Then Exit Sub
            
'            If (Not EsDeVarios) Or (EsDeVarios And modo = 3) Then
            Text1(4).Text = Format(vCliente.Codigo, "000000")
            If (Modo = 3) Or (Modo = 4) Then
                Text1(5).Text = vCliente.Nombre  'Nom clien
                Text1(8).Text = vCliente.Domicilio
                Text1(9).Text = vCliente.CPostal
                Text1(10).Text = vCliente.Poblacion
                Text1(11).Text = vCliente.Provincia
                Text1(6).Text = vCliente.NIF
                Text1(7).Text = DBLet(vCliente.TfnoClien, "T")
            End If
            
            If Modo = 3 Then 'insertar
                Text1(14).Text = vCliente.ForPago
                Text2(14).Text = PonerNombreDeCod(Text1(14), conAri, "sforpa", "nomforpa")
                Text1(15).Text = Format(vCliente.DtoPPago, FormatoDescuento)
                Text1(16).Text = Format(vCliente.DtoGnral, FormatoDescuento)
                Text1(17).Text = vCliente.Agente
                Text2(17).Text = PonerNombreDeCod(Text1(17), conAri, "sagent", "nomagent")
                Me.cboFacturacion.ListIndex = vCliente.TipoFactu
            End If

            Observaciones = DBLet(vCliente.Observaciones)
            If Observaciones <> "" Then
                MsgBox Observaciones, vbInformation, "Observaciones del cliente"
            End If
            
            'Comprobar si el cliente tiene cobros pendientes
            ComprobarCobrosCliente CodClien, Text1(1).Text
        End If
    Else
        LimpiarDatosCliente
    End If
    Set vCliente = Nothing

EPonerDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poner Datos Cliente", Err.Description
End Sub



Private Sub PonerDatosClienteVario(nifClien As String)
Dim vCliente As CCliente
Dim b As Boolean
   
    If nifClien = "" Then Exit Sub
   
    Set vCliente = New CCliente
    b = vCliente.LeerDatosCliVario(nifClien)
    Text1(5).Text = vCliente.Nombre  'Nom clien
    Text1(8).Text = vCliente.Domicilio
    Text1(9).Text = vCliente.CPostal
    Text1(10).Text = vCliente.Poblacion
    Text1(11).Text = vCliente.Provincia
'    Text1(6).Text = vCliente.NIF
    Text1(7).Text = DBLet(vCliente.TfnoClien, "T")
            
    If Not b Then PonerFoco Text1(6)
    Set vCliente = Nothing
End Sub


Private Sub BloquearDatosCliente(bol As Boolean)
Dim I As Byte

    'bloquear/desbloquear campos de datos segun sea de varios o no
    If Modo <> 5 Then
        Me.imgBuscar(1).visible = bol
        Me.imgBuscar(1).Enabled = bol
        Me.imgBuscar(6).Enabled = bol
        
        For I = 5 To 11 'si no es de varios no se pueden modificar los datos
            BloquearTxt Text1(I), Not bol
        Next I
    End If
End Sub


Private Function ActualizarClienteVarios(clien As String, NIF As String) As Boolean
Dim vCliente As CCliente

    On Error GoTo EActualizarCV

    ActualizarClienteVarios = False
    
    Set vCliente = New CCliente
    If EsClienteVarios(clien) Then
        vCliente.NIF = NIF
        vCliente.Nombre = Text1(5).Text
        vCliente.Domicilio = Text1(8).Text
        vCliente.CPostal = Text1(9).Text
        vCliente.Poblacion = Text1(10).Text
        vCliente.Provincia = Text1(11).Text
        vCliente.TfnoClien = Text1(7).Text
        vCliente.ActualizarClienteV (NIF)
    End If
    Set vCliente = Nothing
    
    ActualizarClienteVarios = True
    
EActualizarCV:
    If Err.Number <> 0 Then
        ActualizarClienteVarios = False
    Else
        ActualizarClienteVarios = True
    End If
End Function


Private Sub BotonImprimirProForma(OpcionListado As Byte)
Dim cadFormula As String
Dim Cadparam As String
Dim NumParam As Byte
Dim Cadselect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim Devuelve As String
Dim vTipom As CTiposMov

    If Text1(0).Text = "" Then
        MsgBox "Debe seleccionar una Oferta para Imprimir.", vbInformation
        Exit Sub
    End If
    
    cadFormula = ""
    Cadparam = ""
    Cadselect = ""
    NumParam = 0
    
    '===================================================
    '============ PARAMETROS ===========================
    indRPT = 17 'Facturas Proforma Clientes
    If Not PonerParamRPT(indRPT, Cadparam, NumParam, nomDocu) Then
        Exit Sub
    End If
    
    'Pasar la letra serie de la factura como par�metro
    Set vTipom = New CTiposMov
    If vTipom.Leer("FAV") Then
        
    End If
    Set vTipom = Nothing
    
      
    'Nombre fichero .rpt a Imprimir
    frmImprimir.NombreRPT = nomDocu
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion N� de Oferta
    '---------------------------------------------------
    If Text1(0).Text <> "" Then
        'N� Oferta
        Devuelve = "{" & NombreTabla & ".numofert}=" & Val(Text1(0).Text)
        If Not AnyadirAFormula(cadFormula, Devuelve) Then Exit Sub
        Cadselect = cadFormula
    End If
   
    If Not HayRegParaInforme(NombreTabla, Cadselect) Then Exit Sub
     
     
    

        Devuelve = DevuelveDesdeBDNew(conAri, "sclien", "tipoiva", "codclien", Text1(4).Text, "N")
        
 
        
        If Devuelve <> "" Then
            Cadparam = Cadparam & "pTipoIVA=" & Devuelve & "|"
            NumParam = NumParam + 1
        End If
   
    
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = Cadparam
        .NumeroParametros = NumParam
        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = OpcionListado
        .Titulo = "Factura ProForma"
        .ConSubInforme = True
        .Show vbModal
    End With
End Sub


Private Sub CalcularDatosFactura()
Dim T
Dim cadWhere As String
Dim SQL As String
Dim vFactu As CFactura

    'Limpiar en el form los datos calculados de la factura
    'y volvemos a recalcular
    For Each T In Text3
        T.Text = ""
    Next
    
    'Comprobar que hay lineas de albaran para calcular totales
    cadWhere = ObtenerWhereCP
    SQL = "Select count(*) from " & NomTablaLineas & " Where " & Replace(cadWhere, NombreTabla, NomTablaLineas)
    If RegistrosAListar(SQL) = 0 Then Exit Sub
    
    Set vFactu = New CFactura
    vFactu.DtoPPago = CCur(ComprobarCero(Text1(15).Text))
    vFactu.DtoGnral = CCur(ComprobarCero(Text1(16).Text))
    vFactu.Cliente = Text1(4).Text
    
    If vFactu.CalcularDatosFactura(cadWhere, NombreTabla, NomTablaLineas, False) Then
        Text3(33).Text = vFactu.BrutoFac
        Text3(34).Text = vFactu.ImpPPago
        Text3(35).Text = vFactu.ImpGnral
        Text3(36).Text = vFactu.BaseImp
        Text3(37).Text = QuitarCero(vFactu.TipoIVA1)
        Text3(38).Text = QuitarCero(vFactu.TipoIVA2)
        Text3(39).Text = QuitarCero(vFactu.TipoIVA3)
        Text3(40).Text = vFactu.PorceIVA1
        Text3(41).Text = vFactu.PorceIVA2
        Text3(42).Text = vFactu.PorceIVA3
        Text3(43).Text = vFactu.BaseIVA1
        Text3(44).Text = vFactu.BaseIVA2
        Text3(45).Text = vFactu.BaseIVA3
        Text3(46).Text = vFactu.ImpIVA1
        Text3(47).Text = vFactu.ImpIVA2
        Text3(48).Text = vFactu.ImpIVA3
        Text3(55).Text = vFactu.TotalFac
        Text3(56).Text = vFactu.BaseImp
        
        
        'Recargos de equivalencia
        Text3(49).Text = vFactu.PorceIVA1RE
        Text3(50).Text = vFactu.PorceIVA2RE
        Text3(51).Text = vFactu.PorceIVA3RE
        Text3(52).Text = vFactu.ImpIVA1RE
        Text3(53).Text = vFactu.ImpIVA2RE
        Text3(54).Text = vFactu.ImpIVA3RE
        
        
        FormatoDatosTotales
    Else
        MuestraError Err.Number, "Calculando Totales", Err.Description
    End If
    Set vFactu = Nothing
End Sub



Private Function FormatoDatosTotales()
Dim I As Byte

    For I = 33 To 36
        Text3(I).Text = QuitarCero(Text3(I).Text)
        Text3(I).Text = Format(Text3(I).Text, FormatoImporte)
    Next I
    For I = 49 To 54
        Text3(I).Text = QuitarCero(Text3(I).Text)
        Text3(I).Text = Format(Text3(I).Text, FormatoImporte)
    Next I
    'Desglose B.Imponible por IVA
    For I = 43 To 45
        If Text3(I).Text <> "" Then
             If CSng(Text3(I).Text) = 0 Then
                Text3(I).Text = QuitarCero(Text3(I).Text)
                Text3(I - 3).Text = QuitarCero(Text3(I - 3).Text)
                Text3(I - 6).Text = QuitarCero(Text3(I - 6).Text)
                Text3(I + 3).Text = QuitarCero(Text3(I).Text)
            Else
                Text3(I).Text = Format(Text3(I).Text, FormatoImporte)
                Text3(I - 3) = Format(Text3(I - 3).Text, FormatoDescuento)
                Text3(I + 3).Text = Format(Text3(I + 3).Text, FormatoImporte)
            End If
        End If
    Next I
     'TOTALES
    Text3(55).Text = Format(Text3(55).Text, FormatoImporte)
    Text3(56).Text = Format(Text3(56).Text, FormatoImporte)
End Function




Private Function PonerDptoEnCliente() As Boolean
Dim vClien As CCliente
Dim NomDpto As String

    Set vClien = New CCliente
    vClien.Codigo = Text1(4).Text
    'si existe el departamento para el cliente
    If vClien.DptoCliente(Text1(12).Text, NomDpto) Then
        Text2(12).Text = NomDpto
        PonerDptoEnCliente = True
    Else
        PonerDptoEnCliente = False
    End If
    Set vClien = Nothing
End Function


Private Sub ComprobarRefObligatoria()
Dim vClien As CCliente

    Set vClien = New CCliente
    vClien.Codigo = Text1(4).Text
    If vClien.TieneRefObligatoria(Text1(13).Text) Then
        If Text1(13).Text = "" Then PonerFoco Text1(13)
    End If
    Set vClien = Nothing
End Sub


Private Sub InsertarCabecera()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim SQL As String

    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
        Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
        SQL = CadenaInsertarDesdeForm(Me)
        If SQL <> "" Then
            If InsertarOferta(SQL, vTipoMov) Then
                'El Data esta vacio, desde el modo de inicio se pulsa Insertar
                CadenaConsulta = "Select * from " & NombreTabla & " WHERE " & ObtenerWhereCP & Ordenacion
                PonerCadenaBusqueda
                PonerModo 2
                'Ponerse en Modo Insertar Lineas
                BotonMtoLineas 1, "Oferta"
                BotonAnyadirLinea
            End If
        End If
        FormateaCampo Text1(0)
    End If
    Set vTipoMov = Nothing
End Sub
