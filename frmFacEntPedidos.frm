VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmFacEntPedidos 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   4335
   ClientWidth     =   14040
   Icon            =   "frmFacEntPedidos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   14040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   76
      Top             =   410
      Width           =   13815
      Begin VB.CheckBox chkServirCom 
         Caption         =   "Servir completo"
         Enabled         =   0   'False
         Height          =   240
         Left            =   4680
         TabIndex        =   4
         Tag             =   "Servir completo|N|N|||scaped|servcomp||N|"
         Top             =   120
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   5
         Left            =   9360
         MaxLength       =   40
         TabIndex        =   9
         Tag             =   "Nombre Cliente|T|N|||scaped|nomclien||N|"
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwww aq"
         Top             =   480
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   8445
         MaxLength       =   30
         TabIndex        =   8
         Tag             =   "Cod. Cliente|N|N|||scaped|codclien|000000|N|"
         Text            =   "Text1"
         Top             =   480
         Width           =   780
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   8445
         MaxLength       =   30
         TabIndex        =   7
         Tag             =   "Realizada Por|N|N|0|9999|scaped|codtraba|0000|N|"
         Text            =   "Text1"
         Top             =   130
         Width           =   780
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   3
         Left            =   9360
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   82
         Text            =   "Text2"
         Top             =   130
         Width           =   3260
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   1230
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Pedido|F|N|||scaped|fecpedcl|dd/mm/yyyy|N|"
         Top             =   360
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   0
         Left            =   240
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "Nº Pedido|N|S|0||scaped|numpedcl|0000000|S|"
         Text            =   "Text1 7"
         Top             =   360
         Width           =   885
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha Entrega|F|N|||scaped|fecentre|dd/mm/yyyy|N|"
         Top             =   360
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   18
         Left            =   3570
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Semana Entrega|N|N|||scaped|sementre||N|"
         Top             =   360
         Width           =   465
      End
      Begin VB.CheckBox chkVisadoRes 
         Caption         =   "Visado Responsable"
         Height          =   240
         Left            =   4680
         TabIndex        =   6
         Tag             =   "Visado Responsable|N|N|||scaped|visadore||N|"
         Top             =   590
         Width           =   1815
      End
      Begin VB.CheckBox chkRestoPed 
         Caption         =   "Resto de Pedido"
         Enabled         =   0   'False
         Height          =   240
         Left            =   4680
         TabIndex        =   5
         Tag             =   "Resto de Pedido|N|N|||scaped|restoped||N|"
         Top             =   355
         Width           =   1815
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   8145
         Picture         =   "frmFacEntPedidos.frx":000C
         ToolTipText     =   "Buscar cliente"
         Top             =   510
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         Height          =   255
         Index           =   0
         Left            =   7200
         TabIndex        =   83
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Realiz. Por"
         Height          =   255
         Index           =   21
         Left            =   7200
         TabIndex        =   81
         Top             =   135
         Width           =   780
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   8145
         Picture         =   "frmFacEntPedidos.frx":010E
         ToolTipText     =   "Buscar trabajador"
         Top             =   165
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "F. Pedido"
         Height          =   255
         Index           =   14
         Left            =   1230
         TabIndex        =   80
         Top             =   165
         Width           =   735
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   2055
         Picture         =   "frmFacEntPedidos.frx":0210
         ToolTipText     =   "Buscar fecha"
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Pedido"
         Height          =   255
         Index           =   50
         Left            =   240
         TabIndex        =   79
         Top             =   165
         Width           =   735
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   3225
         Picture         =   "frmFacEntPedidos.frx":029B
         ToolTipText     =   "Buscar fecha"
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "F. Entrega"
         Height          =   255
         Index           =   51
         Left            =   2400
         TabIndex        =   78
         Top             =   165
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Semana"
         Height          =   255
         Index           =   8
         Left            =   3570
         TabIndex        =   77
         Top             =   165
         Width           =   615
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   16
      Left            =   2400
      Locked          =   -1  'True
      MaxLength       =   60
      TabIndex        =   51
      Text            =   "Text2 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwqa"
      Top             =   6660
      Visible         =   0   'False
      Width           =   6885
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   120
      TabIndex        =   34
      Top             =   6495
      Width           =   2175
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   240
         TabIndex        =   35
         Top             =   180
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   12690
      TabIndex        =   32
      Top             =   6600
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   11520
      TabIndex        =   31
      Top             =   6600
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   12360
      Top             =   2280
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   14040
      _ExtentX        =   24765
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   24
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
            Object.ToolTipText     =   "Lineas Pedido"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Generar Albaran"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Facturar"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Pedido"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Vincular pedido a produccion"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Orden Instal."
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   56
         Left            =   10080
         MaxLength       =   15
         TabIndex        =   116
         Text            =   "Text1 7"
         Top             =   80
         Width           =   1530
      End
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
         Left            =   8520
         MaxLength       =   15
         TabIndex        =   115
         Text            =   "TOTAL"
         Top             =   100
         Width           =   1490
      End
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   6720
         TabIndex        =   37
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   12480
      Top             =   2640
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
      Height          =   5100
      Left            =   120
      TabIndex        =   38
      Tag             =   "Fecha Oferta|F|N|||scapre|fecentre|dd/mm/yyyy|N|"
      Top             =   1275
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   8996
      _Version        =   393216
      Style           =   1
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
      TabCaption(0)   =   "Datos básicos"
      TabPicture(0)   =   "frmFacEntPedidos.frx":0326
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DataGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtAux(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtAux(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtAux(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtAux(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtAux(6)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtAux(7)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtAux(8)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtAux(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdAux(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdAux(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "FrameCliente"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtAux(5)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtAux(10)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtAux(9)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtAux(11)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "Otros Datos"
      TabPicture(1)   =   "frmFacEntPedidos.frx":0342
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(45)"
      Tab(1).Control(1)=   "Label1(3)"
      Tab(1).Control(2)=   "Label1(5)"
      Tab(1).Control(3)=   "Label1(18)"
      Tab(1).Control(4)=   "imgBuscar(9)"
      Tab(1).Control(5)=   "Text1(19)"
      Tab(1).Control(6)=   "Text1(20)"
      Tab(1).Control(7)=   "Text1(21)"
      Tab(1).Control(8)=   "Text1(22)"
      Tab(1).Control(9)=   "Text1(23)"
      Tab(1).Control(10)=   "Text1(24)"
      Tab(1).Control(11)=   "Text1(25)"
      Tab(1).Control(12)=   "FrameHco"
      Tab(1).Control(13)=   "Text1(30)"
      Tab(1).Control(14)=   "Text1(29)"
      Tab(1).Control(15)=   "Text1(31)"
      Tab(1).ControlCount=   16
      TabCaption(2)   =   "Totales"
      TabPicture(2)   =   "frmFacEntPedidos.frx":035E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameFactura"
      Tab(2).ControlCount=   1
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   11
         Left            =   6000
         MaxLength       =   12
         TabIndex        =   44
         Tag             =   "Palets"
         Text            =   "Palets"
         Top             =   4080
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   1485
         Index           =   31
         Left            =   -73680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Tag             =   "06|T|S|||scaped|observa6||N|"
         Top             =   3360
         Width           =   8805
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   9
         Left            =   10920
         MaxLength       =   12
         TabIndex        =   45
         Tag             =   "Cajas"
         Text            =   "Cajas"
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
         Index           =   10
         Left            =   11760
         MaxLength       =   12
         TabIndex        =   136
         Tag             =   "PrecioLit"
         Text            =   "PrecioLitro"
         Top             =   4080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   29
         Left            =   -71040
         MaxLength       =   80
         TabIndex        =   23
         Tag             =   "Observación pedido 1|T|S|||scaped|observap1||N|"
         Top             =   840
         Width           =   8805
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   30
         Left            =   -71040
         MaxLength       =   80
         TabIndex        =   24
         Tag             =   "Observación pedido 2|T|S|||scaped|observap2||N|"
         Top             =   1140
         Width           =   8805
      End
      Begin VB.Frame FrameHco 
         Height          =   1275
         Left            =   -64680
         TabIndex        =   117
         Top             =   1800
         Width           =   2775
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   26
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   122
            Top             =   200
            Width           =   1185
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   27
            Left            =   1560
            MaxLength       =   30
            TabIndex        =   121
            Text            =   "Text1"
            Top             =   570
            Width           =   660
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   27
            Left            =   2235
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   120
            Text            =   "Text2"
            Top             =   570
            Width           =   3285
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   28
            Left            =   1560
            MaxLength       =   30
            TabIndex        =   119
            Text            =   "Text1"
            Top             =   940
            Width           =   660
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   28
            Left            =   2235
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   118
            Text            =   "Text2"
            Top             =   940
            Width           =   3285
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Eliminación"
            Height          =   255
            Index           =   37
            Left            =   120
            TabIndex        =   125
            Top             =   200
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Trabajador"
            Height          =   255
            Index           =   38
            Left            =   120
            TabIndex        =   124
            Top             =   570
            Width           =   825
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   7
            Left            =   1080
            Picture         =   "frmFacEntPedidos.frx":037A
            ToolTipText     =   "Buscar trabajador"
            Top             =   570
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Incidencia"
            Height          =   255
            Index           =   40
            Left            =   120
            TabIndex        =   123
            Top             =   940
            Width           =   825
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   8
            Left            =   1080
            Picture         =   "frmFacEntPedidos.frx":047C
            ToolTipText     =   "Buscar incidencia"
            Top             =   940
            Width           =   240
         End
      End
      Begin VB.Frame FrameFactura 
         Height          =   3300
         Left            =   -73200
         TabIndex        =   84
         Top             =   1200
         Width           =   10575
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   49
            Left            =   6600
            MaxLength       =   5
            TabIndex        =   133
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   52
            Left            =   7200
            MaxLength       =   15
            TabIndex        =   132
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   1245
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   50
            Left            =   6600
            MaxLength       =   5
            TabIndex        =   131
            Text            =   "Text1 7"
            Top             =   1755
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   53
            Left            =   7200
            MaxLength       =   15
            TabIndex        =   130
            Text            =   "Text1 7"
            Top             =   1755
            Width           =   1245
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   51
            Left            =   6600
            MaxLength       =   5
            TabIndex        =   129
            Text            =   "Text1 7"
            Top             =   2085
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   54
            Left            =   7200
            MaxLength       =   15
            TabIndex        =   128
            Text            =   "Text1 7"
            Top             =   2085
            Width           =   1245
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
            Left            =   6600
            MaxLength       =   15
            TabIndex        =   101
            Text            =   "Text1 7"
            Top             =   2760
            Width           =   1845
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   48
            Left            =   4920
            MaxLength       =   15
            TabIndex        =   100
            Text            =   "Text1 7"
            Top             =   2085
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   42
            Left            =   4320
            MaxLength       =   5
            TabIndex        =   99
            Text            =   "Text1 7"
            Top             =   2085
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   39
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   98
            Text            =   "Text1 7"
            Top             =   2085
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   45
            Left            =   2640
            MaxLength       =   15
            TabIndex        =   97
            Text            =   "Text1 7"
            Top             =   2085
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   47
            Left            =   4920
            MaxLength       =   15
            TabIndex        =   96
            Text            =   "Text1 7"
            Top             =   1755
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   41
            Left            =   4320
            MaxLength       =   5
            TabIndex        =   95
            Text            =   "Text1 7"
            Top             =   1755
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   38
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   94
            Text            =   "Text1 7"
            Top             =   1755
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   44
            Left            =   2640
            MaxLength       =   15
            TabIndex        =   93
            Text            =   "Text1 7"
            Top             =   1755
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   46
            Left            =   4920
            MaxLength       =   15
            TabIndex        =   92
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   40
            Left            =   4320
            MaxLength       =   5
            TabIndex        =   91
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   37
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   90
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   43
            Left            =   2640
            MaxLength       =   15
            TabIndex        =   89
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   36
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   88
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   35
            Left            =   3960
            MaxLength       =   15
            TabIndex        =   87
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1365
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   34
            Left            =   2160
            MaxLength       =   15
            TabIndex        =   86
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1365
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   33
            Left            =   240
            MaxLength       =   15
            TabIndex        =   85
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1485
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. RE"
            Height          =   255
            Index           =   22
            Left            =   7440
            TabIndex        =   135
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "% RE"
            Height          =   255
            Index           =   48
            Left            =   6600
            TabIndex        =   134
            Top             =   1200
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Cod. IVA"
            Height          =   255
            Index           =   42
            Left            =   1920
            TabIndex        =   114
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "% IVA"
            Height          =   255
            Index           =   41
            Left            =   4320
            TabIndex        =   113
            Top             =   1230
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "TOTAL PEDIDO"
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
            Left            =   4800
            TabIndex        =   112
            Top             =   2760
            Width           =   1530
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
            TabIndex        =   111
            Top             =   2160
            Width           =   135
         End
         Begin VB.Line Line1 
            X1              =   1800
            X2              =   8520
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. IVA"
            Height          =   255
            Index           =   33
            Left            =   4920
            TabIndex        =   110
            Top             =   1230
            Width           =   1335
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
            TabIndex        =   109
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
            TabIndex        =   108
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
            Index           =   30
            Left            =   1920
            TabIndex        =   107
            Top             =   480
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "Base Imponible"
            Height          =   255
            Index           =   2
            Left            =   5760
            TabIndex        =   106
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Dto Gn"
            Height          =   255
            Index           =   12
            Left            =   3960
            TabIndex        =   105
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Dto PP"
            Height          =   255
            Index           =   11
            Left            =   2160
            TabIndex        =   104
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Bruto"
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   103
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Base Imponible"
            Height          =   255
            Index           =   9
            Left            =   2760
            TabIndex        =   102
            Top             =   1200
            Width           =   1215
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   25
         Left            =   -72600
         MaxLength       =   10
         TabIndex        =   73
         Tag             =   "Fecha Oferta|F|S|||scaped|fecofert|dd/mm/yyyy|N|"
         Top             =   795
         Width           =   1185
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   24
         Left            =   -73680
         MaxLength       =   7
         TabIndex        =   72
         Tag             =   "Nº Oferta|N|S|||scaped|numofert|0000000|N|"
         Text            =   "Text1 7"
         Top             =   795
         Width           =   885
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   5
         Left            =   8280
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   48
         Tag             =   "Descuento 1"
         Text            =   "OF"
         Top             =   4080
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Frame FrameCliente 
         Height          =   2310
         Left            =   240
         TabIndex        =   56
         Top             =   310
         Width           =   13215
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   32
            Left            =   10440
            MaxLength       =   30
            TabIndex        =   138
            Tag             =   "Ref. Produccion|N|S|0||scaped|refproduccion|0|N|"
            Text            =   "Text1"
            Top             =   1920
            Width           =   900
         End
         Begin VB.CheckBox chkRecogeClien 
            Caption         =   "Recoge cliente"
            Enabled         =   0   'False
            Height          =   240
            Left            =   8160
            TabIndex        =   127
            Tag             =   "Recoge cliente|N|N|||scaped|recogecl||N|"
            Top             =   1560
            Width           =   1455
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   12
            Left            =   7755
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   69
            Tag             =   "Direccion/Dpto.|T|S|||scaped|nomdirec||N|"
            Text            =   "Text2"
            Top             =   165
            Width           =   3645
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   12
            Left            =   7170
            MaxLength       =   30
            TabIndex        =   17
            Tag             =   "Direccion/Dpto.|N|S|0|999|scaped|coddirec|000|N|"
            Text            =   "Text1"
            Top             =   165
            Width           =   540
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   11
            Left            =   1170
            MaxLength       =   30
            TabIndex        =   15
            Tag             =   "Provincia|T|N|||scaped|proclien||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text22"
            Top             =   1575
            Width           =   2445
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   9
            Left            =   1170
            MaxLength       =   6
            TabIndex        =   13
            Tag             =   "CPostal|T|N|||scaped|codpobla||N|"
            Text            =   "Text15"
            Top             =   1230
            Width           =   630
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   10
            Left            =   1820
            MaxLength       =   30
            TabIndex        =   14
            Tag             =   "Población|T|N|||scaped|pobclien||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
            Top             =   1230
            Width           =   3405
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   7
            Left            =   3360
            MaxLength       =   20
            TabIndex        =   11
            Tag             =   "teléfono Cliente|T|S|||scaped|telclien||N|"
            Text            =   "12345678911234567899"
            Top             =   1920
            Width           =   1845
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   6
            Left            =   1170
            MaxLength       =   15
            TabIndex        =   10
            Tag             =   "NIF Cliente|T|N|||scaped|nifclien||N|"
            Text            =   "123456789"
            Top             =   165
            Width           =   990
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   13
            Left            =   1170
            MaxLength       =   20
            TabIndex        =   16
            Tag             =   "Referencia Cliente|T|S|||scaped|referenc||N|"
            Text            =   "Text1 Text1 Text1 Te"
            Top             =   1920
            Width           =   1605
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   17
            Left            =   7170
            MaxLength       =   30
            TabIndex        =   18
            Tag             =   "Cod. Agente|N|N|0|9999|scaped|codagent|0000|N|"
            Text            =   "Text1"
            Top             =   516
            Width           =   540
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   17
            Left            =   7755
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   63
            Text            =   "Text2"
            Top             =   516
            Width           =   3645
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   14
            Left            =   7170
            MaxLength       =   30
            TabIndex        =   19
            Tag             =   "Forma de Pago|N|N|0|999|scaped|codforpa|000|N|"
            Text            =   "Text1"
            Top             =   867
            Width           =   540
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   14
            Left            =   7755
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   58
            Text            =   "Text2"
            Top             =   867
            Width           =   3630
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   15
            Left            =   7170
            MaxLength       =   7
            TabIndex        =   20
            Tag             =   "Descuento P.Pago|N|N|0|99.90|scaped|dtoppago|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1218
            Width           =   540
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   16
            Left            =   7170
            MaxLength       =   7
            TabIndex        =   21
            Tag             =   "Descuento General|N|N|0|99.90|scaped|dtognral|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   540
         End
         Begin VB.ComboBox cboFacturacion 
            Height          =   315
            Left            =   7200
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Tag             =   "Tipo Facturación|N|N|||scaped|tipofact||N|"
            Top             =   1920
            Width           =   1800
         End
         Begin VB.TextBox Text1 
            Height          =   675
            Index           =   8
            Left            =   1170
            MaxLength       =   35
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Tag             =   "Domicilio|T|N|||scaped|domclien||N|"
            Text            =   "frmFacEntPedidos.frx":057E
            Top             =   516
            Width           =   4050
         End
         Begin VB.Label Label1 
            Caption         =   "Ref. producción"
            Height          =   255
            Index           =   6
            Left            =   9240
            TabIndex        =   137
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   6
            Left            =   900
            Picture         =   "frmFacEntPedidos.frx":05A2
            ToolTipText     =   "Buscar población"
            Top             =   1245
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Direc."
            Height          =   255
            Index           =   1
            Left            =   5940
            TabIndex        =   71
            Top             =   165
            Width           =   495
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   2
            Left            =   6885
            Picture         =   "frmFacEntPedidos.frx":06A4
            ToolTipText     =   "Buscar direc./dpto"
            Top             =   180
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Provincia"
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   70
            Top             =   1575
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Población"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   68
            Top             =   1230
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Tfno:"
            Height          =   195
            Index           =   19
            Left            =   2925
            TabIndex        =   67
            Top             =   1920
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "NIF"
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   66
            Top             =   165
            Width           =   615
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   900
            Picture         =   "frmFacEntPedidos.frx":07A6
            ToolTipText     =   "Buscar cliente varios"
            Top             =   180
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Ref. Cliente"
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   65
            Top             =   1920
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Agente"
            Height          =   255
            Index           =   34
            Left            =   5940
            TabIndex        =   64
            Top             =   510
            Width           =   615
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   5
            Left            =   6885
            Picture         =   "frmFacEntPedidos.frx":08A8
            ToolTipText     =   "Buscar agente"
            Top             =   525
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Forma Pago"
            Height          =   255
            Index           =   15
            Left            =   5940
            TabIndex        =   62
            Top             =   870
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. P. Pago"
            Height          =   255
            Index           =   25
            Left            =   5940
            TabIndex        =   61
            Top             =   1215
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. Gral"
            Height          =   255
            Index           =   26
            Left            =   5940
            TabIndex        =   60
            Top             =   1575
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Facturac."
            Height          =   255
            Index           =   4
            Left            =   5880
            TabIndex        =   59
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   6885
            Picture         =   "frmFacEntPedidos.frx":09AA
            ToolTipText     =   "Buscar forma de pago"
            Top             =   900
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Domicilio"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   57
            Top             =   516
            Width           =   735
         End
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Index           =   1
         Left            =   2640
         TabIndex        =   55
         ToolTipText     =   "Buscar artículo"
         Top             =   4080
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Index           =   0
         Left            =   960
         TabIndex        =   54
         ToolTipText     =   "Buscar almacen"
         Top             =   4080
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
         TabIndex        =   43
         Tag             =   "Nombre Artículo"
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
         Left            =   10080
         MaxLength       =   12
         TabIndex        =   52
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
         Left            =   9480
         MaxLength       =   30
         TabIndex        =   50
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
         Left            =   8880
         MaxLength       =   5
         TabIndex        =   49
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
         Left            =   7440
         MaxLength       =   12
         TabIndex        =   47
         Tag             =   "Precio"
         Text            =   "123,456.7879"
         Top             =   4080
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   3
         Left            =   6240
         MaxLength       =   16
         TabIndex        =   46
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
         TabIndex        =   42
         Tag             =   "Código Artículo"
         Text            =   "Artic Artic Artic5"
         Top             =   4020
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
         TabIndex        =   41
         Tag             =   "Código Almacen"
         Text            =   "codalmac"
         Top             =   4080
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   23
         Left            =   -73680
         MaxLength       =   80
         TabIndex        =   29
         Tag             =   "Observación 5|T|S|||scaped|observa05||N|"
         Top             =   3000
         Width           =   8805
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   22
         Left            =   -73680
         MaxLength       =   80
         TabIndex        =   28
         Tag             =   "Observación 4|T|S|||scaped|observa04||N|"
         Top             =   2715
         Width           =   8805
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   21
         Left            =   -73680
         MaxLength       =   80
         TabIndex        =   27
         Tag             =   "Observación 3|T|S|||scaped|observa03||N|"
         Top             =   2430
         Width           =   8805
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   20
         Left            =   -73680
         MaxLength       =   80
         TabIndex        =   26
         Tag             =   "Observación 2|T|S|||scaped|observa02||N|"
         Top             =   2145
         Width           =   8805
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   19
         Left            =   -73680
         MaxLength       =   80
         TabIndex        =   25
         Tag             =   "Observación 1|T|S|||scaped|observa01||N|"
         Top             =   1860
         Width           =   8805
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmFacEntPedidos.frx":0AAC
         Height          =   2400
         Left            =   240
         TabIndex        =   53
         Top             =   2640
         Width           =   13455
         _ExtentX        =   23733
         _ExtentY        =   4233
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
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   9
         Left            =   -72480
         Picture         =   "frmFacEntPedidos.frx":0AC1
         ToolTipText     =   "Buscar cliente varios"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones del Pedido"
         Height          =   255
         Index           =   18
         Left            =   -71040
         TabIndex        =   126
         Top             =   645
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Oferta"
         Height          =   255
         Index           =   5
         Left            =   -72600
         TabIndex        =   75
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Oferta"
         Height          =   255
         Index           =   3
         Left            =   -73680
         TabIndex        =   74
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones"
         Height          =   255
         Index           =   45
         Left            =   -73680
         TabIndex        =   40
         Top             =   1635
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   12690
      TabIndex        =   33
      Top             =   6600
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "Ampliación Línea"
      Height          =   255
      Index           =   35
      Left            =   2400
      TabIndex        =   39
      Top             =   6480
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
      Begin VB.Menu mnBarra6 
         Caption         =   "-"
      End
      Begin VB.Menu mnLineas 
         Caption         =   "&Lineas"
         HelpContextID   =   2
         Shortcut        =   ^L
      End
      Begin VB.Menu mnGenAlbaran 
         Caption         =   "&Generar Albaran"
         HelpContextID   =   2
         Shortcut        =   ^G
      End
      Begin VB.Menu mnGeneraFactura 
         Caption         =   "Generar factura"
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Begin VB.Menu mnImpPedido 
            Caption         =   "&Pedido"
            Shortcut        =   ^P
         End
         Begin VB.Menu mnImpOrde 
            Caption         =   "&Orden Instalación"
            Shortcut        =   ^O
         End
      End
      Begin VB.Menu mnbarra5 
         Caption         =   "-"
      End
      Begin VB.Menu mnTodosLosAlmacenes 
         Caption         =   "Todos los almacenes"
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
Attribute VB_Name = "frmFacEntPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda2 As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado2(CadenaSeleccion As String)

Public EsHistorico As Boolean 'Si es true abrir el formulario con la tabla de
                              'de historico schped, y solo en modo de consulta


Private WithEvents frmB As frmBuscaGrid  'Form para busquedas
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

Private WithEvents frmList As frmListadoPed 'Listados para Pedidos (pasar pedido a albaran)
Attribute frmList.VB_VarHelpID = -1
Private WithEvents frmList2 As frmListadoOfer  'Listados para pedir datos para grabar en historico
Attribute frmList2.VB_VarHelpID = -1
Private WithEvents frmNSerie As frmRepCargarNSerie  'Form Cargar nº Series
Attribute frmNSerie.VB_VarHelpID = -1
Private WithEvents frmMen As frmMensajes  'Form Mensajes para mostrar los Nº serie y elegir
Attribute frmMen.VB_VarHelpID = -1

Private WithEvents frmO As frmFacCopiarObservaciones2
Attribute frmO.VB_VarHelpID = -1

Private Modo As Byte
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'   5.-  Mantenimiento Lineas
'   6.- Cargar cantidad servidas al Generar Albaran no completo (Pedido --> Albaran)
'-------------------------------------------------------------------------


Private ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

Dim PrimeraVez As Boolean

Dim EsCabecera As Boolean
'Para saber en MandaBusquedaPrevia si busca en la tabla scapla o en la tabla sdirec

Dim CodTipoMov As String
'Codigo tipo de movimiento en función del valor en la tabla de parámetros: stipom

Dim EsDeVarios As Boolean
'Si el cliente mostrado es de Varios o No

'SQL de la tabla principal del formulario
Private CadenaConsulta As String
Private CadenaSQL As String 'Para crear consulta de Generar Albaran a partir del Pedido

Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla de Cabecera
Private NomTablaLineas As String 'Nombre de la Tabla de lineas
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Dim btnAnyadir As Byte
'Variable que indica el número del Boton  Anyadir en la Toolbar1
Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1


Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos

Private HaCambiadoCP As Boolean
'Para saber si tras haber vuelto de prismaticos ha cambiado el valor del CPostal

Dim PorCaja As Boolean
'Para Saber si se ha salido con precio caja y hay que calcular el importe de la
'linea aplicando el precio de la caja. Si PorCaja=false se aplicaca el precio de unidad

Dim Precio As String 'Precio de la linea de Articulo

Dim ImprimeAlb As Boolean 'Para saber cuando vuelve de Generar ALbaran si se ha solicitado Imprimir Albaran o no
Dim FechaAlb As String 'Para cuando vuelve de pedir datos para Generar Albaran, saber la fecha que se introdujo

Dim gridCargado As Boolean 'Saber si el grid esta cargado cuando se ejecuta DataGrid1_RowColChange

Dim AlbCompleto As Boolean 'Si se va a servir el Pedido Completo (slialb.cantidad=sliped.cantidad)
                            'o se va a servir una parte (slialb.cantidad=sliped.servidas)

Dim CtaBancoPropi As String 'Cuando facturamos el pedido directamente, para saber la caja

Dim txtAnterior As String   'Para que no realice las acciones en el lost_focus si NO ha cambiado nada
Dim ElArticulo As CArticulo

Dim ClienteConTasaReciclado As Boolean  'Cuando pasamos a las lineas pondremos esta variab


'Para servidas. Para el txtaux.left de servidas.
'Dim PosicionTxtauxServidas As Integer

'================================================================================

Private Sub cboFacturacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        Me.SSTab1.Tab = 1
        PonerFoco Text1(19)
    End If
End Sub


Private Sub chkServirCom_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub chkVisadoRes_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim SQL As String
Dim PrimeraLin As Boolean 'Si se inserta la primera linea no esta creado el datagrid1 entonces llamar
                          ' a DataGrid, sino llamar solo a DataGrid2

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
        Case 3 'INSERTAR Cabecera Pedido
            If DatosOk Then InsertarCabecera
        Case 4  'MODIFICAR Cabecera Pedido
            If DatosOk Then
                ActualizarClienteVarios Text1(4).Text, Text1(6).Text
                If ModificaDesdeFormulario(Me, 1) Then
                    TerminaBloquear
                    PosicionarData
                End If
            End If
            
         Case 5 'InsertarModificar LINEA
            'Actualizar el registro en la tabla de lineas 'sliped'
            If ModificaLineas = 1 Then 'INSERTAR lineas Pedidos
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
            
        Case 6 'PASAR Pedido a ALBARAN
            'Comprobar que la cantidad a servir es mayor que cero
            'Por si acaso no ha puesto valores
            If txtAux(3).Text = "" Or txtAux(9).Text = "" Then
                PonerFoco txtAux(9)
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

             
            If Not Servidas_vs_Disponibles Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            
             'Si hay servidas
             SQL = "SELECT SUM(servidas) as servidas from sliped WHERE "
             SQL = SQL & Replace(ObtenerWhereCP, NombreTabla, NomTablaLineas)
             If RegistrosAListar(SQL) = 0 Then 'No hay cantidad en linea para el Albaran
                SQL = "La cantidad total a servir en el Albaran es cero." & vbCrLf
                SQL = SQL & vbCrLf & "Introduzca la cantidad a servir."
                MsgBox SQL, vbExclamation
             Else
                If SePuedeServirPedido Then GenerarAlbaran False
             End If
    End Select
    Screen.MousePointer = vbDefault
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Function Servidas_vs_Disponibles() As Boolean
Dim C As String
Dim C1 As Currency

    Servidas_vs_Disponibles = True
    Set miRsAux = New ADODB.Recordset
    C = "Select * from sliped WHERE " & Replace(ObtenerWhereCP, NombreTabla, NomTablaLineas)
    miRsAux.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    C = ""
    While Not miRsAux.EOF
        Precio = ""
        'Cantidad
        C1 = DBLet(miRsAux!servidas, "N")
        C1 = C1 - DBLet(miRsAux!Cantidad, "N")
        If C1 > 0 Then
            Precio = miRsAux!codartic & " " & miRsAux!NomArtic
            Precio = Precio & "     Cantidad : " & C1
        End If
        'Cantidad
        If miRsAux!codartic <> vParamAplic.ArtReciclado Then
            C1 = DBLet(miRsAux!cajserv, "N")
            C1 = C1 - DBLet(miRsAux!Cajas, "N")
            If C1 > 0 Then
                If Precio = "" Then Precio = miRsAux!codartic & " " & miRsAux!NomArtic
                Precio = Precio & "     Cajas : " & C1
            End If
        End If
            
        If Precio <> "" Then C = C & Precio & vbCrLf
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If C <> "" Then
        C = "Error en cantidades servidas:" & vbCrLf & vbCrLf & C
        MsgBox C, vbExclamation
        Servidas_vs_Disponibles = False
    End If
    
    
ECom:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
    Precio = ""
End Function
Private Sub cmdAux_Click(Index As Integer)
    Select Case Index
        Case 0 'Busqueda de Cod. Almacen
            Set frmAlm = New frmAlmAlPropios
            frmAlm.DatosADevolverBusqueda = "0"
            frmAlm.Show vbModal
            Set frmAlm = Nothing
        Case 1 'Busqueda de Cod. Artic
            Set frmArt = New frmAlmArticulos
            frmArt.DatosADevolverBusqueda2 = "@1@" 'Poner en modo busqueda
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
             Set ElArticulo = Nothing
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
            
        Case 6 'Insertar servidas en Generar Albaran (Pedido --> Albaran)
            Set ElArticulo = Nothing
            InicializarServidas
            PonerModo 2
            CargaTxtAuxServidas False, False
            CargaGrid DataGrid1, Data2, True, False
    End Select
End Sub


Private Sub BotonAnyadir()
'Añadir registro en tabla de cabecera de Pedidos: scaped (Cabecera)
Dim NomTraba As String

    LimpiarCampos 'Vacía los TextBox
    
    
    
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3

    'Poner el nombre del trabajador que esta conectado
    Text1(3).Text = PonerTrabajadorConectado(NomTraba)
    Text2(3).Text = NomTraba

    Me.chkServirCom.Value = 1

    Text1(1).Text = Format(Now, "dd/mm/yyyy") 'Fecha Oferta
    PonerFoco Text1(1)
End Sub


Private Sub BotonAnyadirLinea()
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
       
    'Si el pedido esta vinculado SOLO insertara si la empresa es el avab
    If DBLet(Data1.Recordset!refproduccion, "N") > 0 Then
        If Not vParamAplic.EsAVAB Then
            MsgBox "Pedido bloqueado. No se insertan lineas", vbExclamation
            Exit Sub
        End If
    End If
       
    ModificaLineas = 1 'Ponemos Modo Añadir Linea
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerBotonCabecera False
    lblIndicador.Caption = "INSERTAR"
    
    AnyadirLinea DataGrid1, Data2
    CargaTxtAux True, True
    
    'Poner el Almacen por defecto del Trabajador
    txtAux(0).Text = DevuelveDesdeBDNew(conAri, "straba", "codalmac", "codtraba", Text1(3).Text, "N")
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
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
Dim C2 As String
    C2 = ""
    
    If Not EsHistorico Then C2 = DevuelveListaPedidos
    'If Not vUsu.TrabajadorB Then C2 = " numpedcl in (select distinct(numpedcl) from sliped WHERE codalmac = 1)"
    
'    LimpiarCampos
    If chkVistaPrevia.Value = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia C2
    Else
        LimpiarCampos
        LimpiarDataGrids
        CadenaConsulta = "Select * from " & NombreTabla & " "
        If C2 <> "" Then CadenaConsulta = CadenaConsulta & " WHERE " & C2
        CadenaConsulta = CadenaConsulta & Ordenacion
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
'Prepara el Form para Modificar la cabecera de Pedidos (tabla: scaped)
Dim DeVarios As Boolean

    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    PonerFoco Text1(1)
        
    'Si es Cliente de Varios no se pueden modificar sus datos
    DeVarios = EsClienteVarios(Text1(4).Text)
    BloquearDatosCliente (DeVarios)
End Sub


Private Sub BotonModificarLinea()
'Prepara el Form para Modificar una linea de Pedido (tabla: sliped)
Dim vWhere As String

    On Error GoTo EModificarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub '1= Insertar
    If Data2.Recordset.EOF Then Exit Sub
    
    vWhere = ObtenerWhereCP & " and numlinea=" & Data2.Recordset!numlinea
    vWhere = Replace(vWhere, NombreTabla, NomTablaLineas)
    If Not BloqueaRegistro(NomTablaLineas, vWhere) Then Exit Sub
    
    CargaTxtAux True, False
    ModificaLineas = 2 'Modificar
    'Añadiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR"
    PonerBotonCabecera False
    BloquearTxt Text2(16), False 'Campo Ampliacion Linea
    BloquearTxt txtAux(2), True 'campo nombre articulo
    PonerFoco txtAux(0)
    Me.DataGrid1.Enabled = False
    
    Set ElArticulo = New CArticulo
    ElArticulo.LeerDatos Data2.Recordset!codartic
    
EModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Pedidos (scaped)
' y los registros correspondientes de las tablas de lineas (sliped)
Dim Cad As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    'Esta vinculado
    If DBLet(Data1.Recordset!refproduccion, "N") > 0 Then
        If Not vParamAplic.EsAVAB Then
            MsgBox "Pedido bloqueado.", vbExclamation
            Exit Sub
        End If
    End If




    Cad = "Cabecera de Pedidos." & vbCrLf
    Cad = Cad & "----------------------------------" & vbCrLf & vbCrLf
    Cad = Cad & "Va a eliminar el Pedido:            "
    Cad = Cad & vbCrLf & "Nº:  " & Format(Text1(0).Text, "0000000")
    Cad = Cad & vbCrLf & "Cliente:  " & Format(Text1(4).Text, "000000") & " - " & Text1(5).Text
    Cad = Cad & vbCrLf & vbCrLf & "¿Desea Eliminarlo? "
    
    Screen.MousePointer = vbHourglass
    
    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Data1.Recordset.AbsolutePosition
        
        'Abrir frame de informes para pedir datos antes de grabar en el historico
        CadenaSQL = ""
        Set frmList2 = New frmListadoOfer
        frmList2.OpcionListado = 81
        frmList2.Show vbModal
        Set frmList2 = Nothing
        If CadenaSQL = "" Then Exit Sub
        
        If Not Eliminar() Then Exit Sub
        PosicionarDataTrasEliminar
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Pedido", Err.Description
End Sub


Private Sub BotonEliminarLinea()
'Eliminar una linea Del Pedido. (Tabla: sliped)
Dim SQL As String

    On Error GoTo EEliminarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar

    If Data2.Recordset.EOF Then Exit Sub
            
            
    'Si esta vinculado AVAB-MORALES
    If DBLet(Data1.Recordset!refproduccion, "N") > 0 Then
        If Not vParamAplic.EsAVAB Then
            'NO es AVAB.
            MsgBox "Pedido bloqueado", vbExclamation
            Exit Sub
        End If
    End If
            
    ModificaLineas = 3 'Eliminar
    SQL = "¿Seguro que desea eliminar la línea del Pedido?     "
    SQL = SQL & vbCrLf & "NumLinea:  " & Data2.Recordset!numlinea & vbCrLf
    SQL = SQL & "Almacen:  " & Format(Data2.Recordset!codAlmac, "000")
    SQL = SQL & vbCrLf & "Artículo:  " & Data2.Recordset!codartic & " - " & Data2.Recordset!NomArtic
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Data2.Recordset.AbsolutePosition
        SQL = "Delete from " & NomTablaLineas & " WHERE " & Replace(ObtenerWhereCP, NombreTabla, NomTablaLineas)
        SQL = SQL & " and numlinea=" & Data2.Recordset!numlinea
        conn.Execute SQL
        
        'Si esta vinculado
        If DBLet(Data1.Recordset!refproduccion, "N") > 0 Then
            SQL = "Delete from ariges" & EmprMorales & "." & NomTablaLineas & " WHERE numpedcl = " & Data1.Recordset!refproduccion
            SQL = SQL & " and numlinea=" & Data2.Recordset!numlinea
            EjecutaSQL conAri, SQL, True
        End If
        
        ModificaLineas = 0
        CargaGrid2 DataGrid1, Data2
'        SituarDataTrasEliminar Data2, NumRegElim
        SituarDataPosicion Me.Data2, NumRegElim, SQL
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
        
    Else 'Se llama desde algún Prismatico de otro Form al Mantenimiento de Trabajadores
        If Data1.Recordset.EOF Then
            MsgBox "Ningún registro devuelto.", vbExclamation
            Exit Sub
        End If
        'cad = Data1.Recordset.Fields(0) & "|"
        'cad = cad & Data1.Recordset.Fields(1) & "|"
        Cad = Data1.Recordset.Fields(0)
        RaiseEvent DatoSeleccionado2(Cad)
        Unload Me
    End If
End Sub


Private Sub DataGrid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Ayuda de Etiqueta de precio de salida de la Función de Precios
    If Data2.Recordset.EOF Then Exit Sub
    If (Modo = 2) Or (Modo = 5 And ModificaLineas = 0) Then
        If X > 7750 And X < 8000 Then
            Select Case DataGrid1.Columns(8).Value
                Case "P": Me.DataGrid1.ToolTipText = "P: Promoción"
                Case "E": Me.DataGrid1.ToolTipText = "E: Precio Especial"
                Case "T": Me.DataGrid1.ToolTipText = "T: Tarifa Artículo"
                Case "A": Me.DataGrid1.ToolTipText = "A: Precio Artículo"
                Case "M": Me.DataGrid1.ToolTipText = "M: Manual"
                Case Else
                    Me.DataGrid1.ToolTipText = ""
            End Select
        Else
            Me.DataGrid1.ToolTipText = ""
        End If
    End If
End Sub


Private Sub FijarUdsCaja()

    txtAnterior = DevuelveDesdeBD(conAri, "unicajas", "sartic", "codartic", CStr(Data2.Recordset!codartic), "T")
    If txtAnterior = "" Then txtAnterior = "1"
    If txtAnterior = "0" Then txtAnterior = "1"
    ElArticulo.UnidCaja = CInt(txtAnterior)
    txtAnterior = ""
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim Devuelve As String

    On Error GoTo Error1

    If Modo = 6 And gridCargado Then '6: Pasar Pedido a Albaran no Completo (Introducir las servidas)
        CargaTxtAuxServidas True, True
        txtAux(3).Text = Data2.Recordset!servidas
        txtAux(9).Text = Val(Data2.Recordset.Fields(9))
        FijarUdsCaja
    End If
    
    If Modo = 5 Then 'Poner el valor al camp ampliacion linea '5: modo lineas
        If Not Data2.Recordset.EOF And ModificaLineas <> 1 Then '1: Insertar
            Devuelve = DevuelveDesdeBDNew(conAri, NomTablaLineas, "ampliaci", "numpedcl", Text1(0).Text, "N", , "numlinea", Data2.Recordset!numlinea, "N")
            'Poner descripcion de ampliacion lineas
            Text2(16).Text = Devuelve
        Else
            Text2(16).Text = ""
        End If
    End If
    
Error1:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub Form_Activate()
    If Me.Tag <> "" Then
        Me.Tag = ""
        PonerCampos
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmppal.Icon

    ' ICONITOS DE LA BARRA
    btnAnyadir = 5
    btnPrimero = 21
    With Me.Toolbar1
        .ImageList = frmppal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        .Buttons(5).Image = 3   'Insertar Nuevo
        .Buttons(6).Image = 4   'Modificar
        .Buttons(7).Image = 5   'Borrar
        .Buttons(10).Image = 10 'Mto Lineas Ofertas
        .Buttons(11).Image = 26 'Generar Albaran
        
        'Enero08
        .Buttons(12).Image = 42 'Generar factura desde el pedido
        
        .Buttons(14).Image = 16 'Imprimir Pedido
        .Buttons(15).Image = 35 'Vincular pedido a empresa produccion
        
        If vParamAplic.EsAVAB Then
            .Buttons(16).Image = 40 'Imprimir Orden envasado
            .Buttons(16).ToolTipText = "Imprimir packing list"
        Else
            .Buttons(16).Image = 27 'Imprimir Orden envasado
            .Buttons(16).ToolTipText = "Imprimir orden envasado"
        End If
        
        .Buttons(18).Image = 15  'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With
    Me.SSTab1.Tab = 0
    Me.Toolbar1.Buttons(15).visible = False
    If vUsu.Nivel <= 1 Then
        Me.Toolbar1.Buttons(15).visible = vParamAplic.EsAVAB And Not Me.EsHistorico
    End If
    
    LimpiarCampos   'Limpia los campos TextBox
    
    CargarComboFacturacion
    CodTipoMov = "PEV"
    VieneDeBuscar = False
   
    'Comprobar si es Departamento o Direccion
    If vParamAplic.Departamento Then
        Me.Label1(1).Caption = "Dpto."
    Else
        Me.Label1(1).Caption = "Direc."
    End If
        
    'Si es AVAB longitud NIF y domicilio cmabian
    'If vEmpresa.codempre = EmpresaAVAB Then
    If vParamAplic.EsAVAB Then
        'AVAB
        '.....................
        'NIF
        Text1(6).MaxLength = 50
        Text1(6).Width = 3990
        'Domicilio
        Text1(8).MaxLength = 100
        Text1(8).Height = 675
  
        Text1(31).visible = True
  
    Else
        'MORALES
        Text1(6).MaxLength = 15
        Text1(6).Width = 1590
        Text1(8).MaxLength = 35
        Text1(8).Height = Text1(6).Height


        Text1(31).visible = False   'Para morales NO dejo ver el observa6 ... de momento

    End If
        
        
        
        
    '## A mano
    Me.FrameHco.visible = EsHistorico
    
    If Not EsHistorico Then
        NombreTabla = "scaped"
        NomTablaLineas = "sliped" 'Tabla lineas de Pedido
        Me.Caption = "Pedidos Clientes"
        Ordenacion = " ORDER BY numpedcl "
    Else
        NombreTabla = "schped"
        NomTablaLineas = "slhped"
        CargarTagsHco Me, "scaped", NombreTabla
        'Estos campos solo estan en la tabla del histórico
        Text1(26).Tag = "Fecha Eliminación|F|N|||schped|fechelim|dd/mm/yyyy|N|"
        Text1(27).Tag = "Trabajador Eliminación|N|N|0|9999|schped|trabelim|0000|N|"
        Text1(28).Tag = "Incidencia elim.|T|N|||schped|codincid||N|"
        Me.Caption = "Histórico Pedidos Clientes"
        Ordenacion = " ORDER BY numpedcl,fecpedcl "
    End If
        
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    
    If DatosADevolverBusqueda2 = "" Then
        CodTipoMov = "-1"
    Else
        CodTipoMov = DatosADevolverBusqueda2
    End If
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    Data1.RecordSource = "Select * from " & NombreTabla & " where numpedcl=" & CodTipoMov
    Data1.Refresh
    
    Me.Tag = "" 'Para que no carge los datos
    If DatosADevolverBusqueda2 = "" Then
        PonerModo 0
    Else
        If Data1.Recordset.EOF Then
            PonerModo 1
            Text1(0).BackColor = vbYellow
        Else
            Me.Tag = "P" 'Para que en el activate ponga los campos
            PonerModo 2
        End If
    End If
    CodTipoMov = "PEV"
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
    
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    
    mnTodosLosAlmacenes.Checked = False
    mnTodosLosAlmacenes.visible = vUsu.TrabajadorB
    Me.mnbarra5.visible = vUsu.TrabajadorB
    Caption = "Pedidos clientes"
    If vUsu.TrabajadorB Then Caption = Caption & "     **************"

End Sub


Private Sub LimpiarCampos()
On Error Resume Next

    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.cboFacturacion.ListIndex = -1
    Me.chkVisadoRes.Value = 0
    Me.chkRestoPed.Value = 0
    Me.chkServirCom.Value = 0
    
    Text3(0).Text = "BASE IMP."
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Agentes
Dim Indice As Byte
    Indice = 17
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod agente
    FormateaCampo Text1(Indice)
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nombre agente
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
                Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 3)
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
    FormateaCampo Text1(4)
    HaDevueltoDatos = True
End Sub


Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento C. Postales
Dim Indice As Byte
Dim Devuelve As String

    Indice = 9
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
     'Poblacion
    Text1(Indice + 1).Text = ObtenerPoblacion(Text1(Indice).Text, Devuelve)
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
'Cuando pasa de Pedido -> Albaran
'Aqui devuelve los valores que se introducen desde el Form de Listado de Pedido
'para generar el Albaran
Dim vSQL As String

    'Construimos parte de la SQL para insertar en tabla de Albaranes(scaalb)
    FechaAlb = RecuperaValor(CadenaSeleccion, 4)
    vSQL = ""
    vSQL = " '" & Format(FechaAlb, FormatoFecha) & "' as fechaalb, " 'Fecha Albaran
    vSQL = vSQL & "1 as factursn, " 'facturar s/n    Le pongo un 1
    vSQL = vSQL & "codclien, nomclien, domclien, codpobla, pobclien, proclien, nifclien, "
    vSQL = vSQL & "telclien, coddirec, nomdirec, referenc,  "
    vSQL = vSQL & RecuperaValor(CadenaSeleccion, 1) & " as codtraba, " 'Trabajador de Albaran
    vSQL = vSQL & " codtraba as codtrab1, "
    vSQL = vSQL & RecuperaValor(CadenaSeleccion, 2) & " as codtrab2, " 'Material Preparado por
    vSQL = vSQL & "codagent, codforpa, "
    vSQL = vSQL & RecuperaValor(CadenaSeleccion, 3) & " as codenvio, " 'Cod Envio
    vSQL = vSQL & "dtoppago, dtognral, tipofact, observa01, observa02, observa03, observa04, observa05, "
    vSQL = vSQL & "numofert, fecofert, "  'Nº Oferta, fecha de la Oferta
    vSQL = vSQL & Text1(0).Text & " as numpedcl, '" 'Nº Pedido
    vSQL = vSQL & Format(Text1(1).Text, FormatoFecha) & "' as fecpedcl, '" 'Fecha Pedido
    vSQL = vSQL & Format(Text1(2).Text, FormatoFecha) & "' as fecentre, " 'Fecha Prevista Entrega
    
    vSQL = vSQL & Val(Text1(18).Text) & " as sementre " 'Semana entrega Pedido
    CadenaSQL = vSQL
    
    'Se almacena aqui si el usuario quiere imprimir el Albaran tras generarlo
    ImprimeAlb = CBool(RecuperaValor(CadenaSeleccion, 5))
    
    'Solo para la facturacion
    CtaBancoPropi = RecuperaValor(CadenaSeleccion, 6)
    
    'Albaran VALORADO
    'Precio
    Precio = RecuperaValor(CadenaSeleccion, 7)
End Sub


Private Sub frmList2_DatoSeleccionado(CadenaSeleccion As String)
'devuelve los datos necesarios para grabar en la tabla cabecera del historico
    CadenaSQL = ""
    CadenaSQL = DBSet(RecuperaValor(CadenaSeleccion, 1), "F") & " as fechelim,"
    CadenaSQL = CadenaSQL & RecuperaValor(CadenaSeleccion, 2) & " as trabelim,"
    CadenaSQL = CadenaSQL & DBSet(RecuperaValor(CadenaSeleccion, 3), "T") & " as codincid"
End Sub


Private Sub frmMen_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mensaje de Nº de Serie disponibles
'En cadena seleccion estan concatenados los seleccionados
Dim I As Byte, J As Byte, K As Byte
Dim nSerie As String
Dim SQL As String
Dim Devuelve As String
Dim cadSEL As String
Dim codartic As String
Dim RS As ADODB.Recordset
Dim contador As Integer
Dim numSerie As CNumSerie

    On Error GoTo ErrorNSerie
    
    'Para cada articulo (separado por ., obtener los nº de serie empipados
    I = 0
    J = I + 1
    I = InStr(J, CadenaSeleccion, "·")
    
    While I > 0
        cadSEL = Mid(CadenaSeleccion, J, I - J)
        
        'Para cada valor empipado actualizar la tabla sserie
        K = InStr(1, cadSEL, "|")
        If K > 0 Then
            codartic = Mid(cadSEL, 1, K - 1) 'El primero es el codartic
            cadSEL = Mid(cadSEL, K + 1, Len(cadSEL)) 'Los Nº de serie
            SQL = "select codartic, cantidad, numlinea from slialb "
            SQL = SQL & " WHERE codtipom='ALV' and numalbar= " & Me.cmdAux(1).Tag & " and codartic=" & DBSet(codartic, "T")
            Set RS = New ADODB.Recordset
            RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        
            K = InStr(1, cadSEL, "|")
            contador = RS!Cantidad
            While K > 0
                nSerie = Mid(cadSEL, 1, K - 1)
                cadSEL = Mid(cadSEL, K + 1, Len(cadSEL))
                
                If contador = 0 Then
                    RS.MoveNext
                    If Not RS.EOF Then contador = RS!Cantidad
                End If
                If contador > 0 Then
                    'Actualizar la tabla sserie
                    Set numSerie = New CNumSerie
                    numSerie.Cliente = Val(Text1(4).Text)
                    numSerie.DirDpto = Text1(12).Text
                    numSerie.tipoMov = "ALV"
                    'Obtenemos la fecha del albaran insertado
                    Devuelve = DevuelveDesdeBDNew(conAri, "scaalb", "fechaalb", "codtipom", "ALV", "T", , "numalbar", Me.cmdAux(1).Tag, "N")
                    numSerie.FechaVta = Devuelve
                    numSerie.ObtenFechaFinGarantia codartic, Devuelve

                    numSerie.NumAlbaran = Me.cmdAux(1).Tag
                    numSerie.NumLinAlb = ComprobarCero(RS!numlinea)
                    numSerie.Articulo = codartic
                    numSerie.numSerie = nSerie
                    
                    numSerie.ActualizarNumSerie (True)
                    
                    Set numSerie = Nothing
                End If
                contador = contador - 1
                K = InStr(1, cadSEL, "|")
            Wend
            RS.Close
            Set RS = Nothing
        End If
        J = I + 1
        I = InStr(J, CadenaSeleccion, "·")
    Wend
    
ErrorNSerie:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Actualizar tabla Nº Series", Err.Description
        MsgBox "No se cargaron correctamente los Nº de Serie.", vbExclamation
    End If
End Sub


Private Sub frmNSerie_CargarNumSeries()
'Insertar un registro en la tabla "sserie" por cada uno de los
'Nº de Serie introducidos en la Tabla Temporal
Dim RStmp As ADODB.Recordset
Dim RsAlb As ADODB.Recordset
Dim SQL As String
Dim I As Byte

    On Error GoTo EInsertar
    
    SQL = "SELECT slialb.codartic, numlinea, cantidad "
    SQL = SQL & " FROM slialb INNER JOIN sartic on slialb.codartic=sartic.codartic "
    SQL = SQL & " WHERE (codtipom='ALV' and numalbar=" & Me.cmdAux(1).Tag
    SQL = SQL & " And nseriesn = 1) ORDER BY codartic, numlinea "

    Set RsAlb = New ADODB.Recordset
    RsAlb.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RsAlb.EOF 'Para cada linea del ALbaran
        'Recuperar los Nº Serie de ese articulo cargados en la Temporal
        'Seleccionar los nº de serie cargados en la temporal: tmpnseries
        SQL = "SELECT * FROM tmpnseries WHERE codusu=" & vUsu.Codigo
        SQL = SQL & " AND codartic=" & DBSet(RsAlb!codartic, "T")
        SQL = SQL & " ORDER BY codartic, numlinea "
        Set RStmp = New ADODB.Recordset
        RStmp.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        'If Not RStmp.EOF Then RStmp.MoveFirst
        'Intentar asignar un Nº serie al total de cantidad del articulo
        
        For I = 1 To RsAlb!Cantidad
            If Not RStmp.EOF Then
                InsertarNSerie RStmp!numSerie, RStmp!codartic, RsAlb!numlinea
                RStmp.MoveNext
            End If
        Next I
        RStmp.Close
        Set RStmp = Nothing
        RsAlb.MoveNext
    Wend
    RsAlb.Close
    Set RsAlb = Nothing
    
EInsertar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Insertando Nº Serie", Err.Description
End Sub


Private Sub frmO_DatoSeleccionado(Datos As String)
Dim I As Integer
    For I = 1 To 5
        Text1(I + 18).Text = RecuperaValor(Datos, I)
    Next
    If Text1(31).visible Then Text1(31).Text = RecuperaValor(Datos, 6)
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Trabajadores
Dim Indice As Byte

    Indice = 3
    Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000") 'Cod Trabajador
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Trabajador
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
            Indice = 4
            Set frmC = New frmFacClientes
            frmC.DatosADevolverBusqueda = "0"
            If Not IsNumeric(Text1(4).Text) Then Text1(4).Text = ""
            frmC.Show vbModal
            Set frmC = Nothing
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
        Case 9
            If Text1(4).Text = "" Then
                MsgBox "Ponga el cliente", vbExclamation
                
            Else
                Indice = 19
                Set frmO = New frmFacCopiarObservaciones2
                frmO.PackingList = False
                frmO.IdCliente = CLng(Text1(4).Text)
                frmO.Show vbModal
                Set frmO = Nothing
            End If
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
   Text1_LostFocus CInt(Indice)
   PonerFoco Text1(Indice)
End Sub


Private Sub mnBuscar_Click()
    Me.SSTab1.Tab = 0
    BotonBuscar
End Sub


Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar lineas de Pedido
         BotonEliminarLinea
    Else   'Eliminar Pedido
         BotonEliminar
    End If
End Sub


Private Sub mnGenAlbaran_Click()
'Dim L As Long
'    If Me.Data1.Recordset Is Nothing Then Exit Sub
'    If Me.Data1.Recordset.EOF Then Exit Sub
    
'    L = Val(Data1.Recordset!numpedcl)
'    If Not BloqueoManual("GenPed", CStr(L)) Then Exit Sub
    
    GeneAlbaranClick

 '   DesBloqueoManual "GenPed"

End Sub

Private Sub GeneAlbaranClick()
'Pasar una Pedido a Albaran
Dim Resp As Byte
Dim b As Boolean
Dim cadMen As String

    'Comprobar que hay un Pedido seleccionado
    If Not ComprobarOpcionTraspaso2(False) Then Exit Sub
    
    
    'si no se va a servir completo preguntar como se quiere servir si completo o no
    If Me.chkServirCom = 0 Then
        'Preguntar si se sirve el pedido completo o no
        Resp = MsgBox("¿Servir el pedido completo?", vbYesNoCancel)
        If Resp = vbCancel Then Exit Sub
    
        If Resp = vbYes Then
            AlbCompleto = True 'SERVIR COMPLETO
        Else
            AlbCompleto = False
        End If
    Else
        AlbCompleto = True
    End If
        
    If AlbCompleto Then 'SERVIR COMPLETO
        Screen.MousePointer = vbHourglass
        'comprobar si hay control de stock si se puede servir el pedido
        b = SePuedeServirPedido
        
        If b Then 'Hay suficiente stock
            'Si hay stock generar albaran completo
            GenerarAlbaran False
        Else
            Screen.MousePointer = vbDefault
            'Si no se puede servir mostrar mensaje detallando y bloquear
            cadMen = "No hay suficiente Stock para servir el Pedido. "
            cadMen = cadMen & vbCrLf & "¿Desea Ver Detalle?"
            If MsgBox(cadMen, vbYesNo, "Contol de Stock") = vbYes Then
                'ANTES 01/12/08
                'frmMensajes.cadWHERE = " WHERE numpedcl = " & Text1(0).Text & " "   'And sfamia.instalac = 0 "
                'ahora
                frmMensajes.cadWhere = " WHERE numpedcl = " & Text1(0).Text & " and ctrstock=1 "
                frmMensajes.vCampos = NomTablaLineas
                frmMensajes.OpcionMensaje = 2 'Articulos sin Stock
                frmMensajes.Show vbModal
            End If
            Exit Sub
        End If
        
    Else 'SERVIR INCOMPLETO
        AlbCompleto = False
        Set ElArticulo = New CArticulo
        InicializarServidas
        'Si no se va a servir completo Mostrar lineas para que se indiquen las Servidas
        MsgBox "Introduzca la cantidad  a servir para cada línea.", vbInformation
        Modo = 6
        gridCargado = False
        Me.cmdAceptar.visible = True
        Me.cmdCancelar.visible = True
        PonerModoOpcionesMenu Modo
        CargaGrid DataGrid1, Data2, True, True
        
        FijarUdsCaja
        
        CargaTxtAuxServidas True, True
        PrimeraVez = True
    End If
End Sub


Private Function ComprobarOpcionTraspaso2(Factura As Boolean) As Boolean

    ComprobarOpcionTraspaso2 = False
    
    
    
    
    
   'Comprobar que hay un Pedido seleccionado
    If Text1(0).Text = "" Then Exit Function
    
    CtaBancoPropi = "- No tiene lineas el pedido" & vbCrLf
    If Not (Data2.Recordset Is Nothing) Then
        If Data2.Recordset.RecordCount > 0 Then CtaBancoPropi = ""
    End If
    
    
    'Comprobar que el Pedido esta visado por el Responsable
    If Me.chkVisadoRes = 0 Then CtaBancoPropi = CtaBancoPropi & "- El pedido debe tener el Visado del Responsable." & vbCrLf
        
    
    
    'si no se va a servir completo preguntar como se quiere servir si completo o no
    If Factura Then
        If Me.chkServirCom = 0 Then CtaBancoPropi = CtaBancoPropi & "-Solo se facturan drectamente pedidos completos" & vbCrLf
    End If
        
        
    If CtaBancoPropi <> "" Then
        CtaBancoPropi = "Faltan campos: " & vbCrLf & vbCrLf & CtaBancoPropi
        MsgBox CtaBancoPropi, vbExclamation
        CtaBancoPropi = ""
        Exit Function
    End If
        
        
    'Si el pedido esta vinculado, de momento NO dejo hacer no complet
    If DBLet(Data1.Recordset!refproduccion, "N") > 0 Then
        If Me.chkServirCom = 0 Then
            MsgBox "Pedido bloqueado. Solo servir completo", vbExclamation
            Exit Function
        End If
            
        If Factura Then
            MsgBox "Pedido bloqueado. No puede generar factura directamente", vbExclamation
            Exit Function
        End If
            
        'Si es pedido vinculado, comprobaremos que si es del AVAB
        If vParamAplic.EsAVAB Then
            'Comprobare que el pedido vincnulado en MORALEs ya no esta, y que esta en albaranes
            CtaBancoPropi = DevuelveDesdeBD(conAri, "numpedcl", "ariges" & EmprMorales & ".scaped", "numpedcl", CStr(Data1.Recordset!refproduccion))
            If CtaBancoPropi <> "" Then
                MsgBox "El pedido no ha pasado a albaran en empresa produccion", vbExclamation
                Exit Function
            Else
                CtaBancoPropi = DevuelveDesdeBD(conAri, "numalbar", "ariges" & EmprMorales & ".scaalb", "refproduccion", CStr(Data1.Recordset!numpedcl))
                If CtaBancoPropi = "" Then
                    MsgBox "NO ese encuentra el albaran asociado al pedido", vbExclamation
                    Exit Function
                End If
                
                'Si llega aqui es que el albaran esta. Vamos a comprobar que estan las mismas lineas y tienen Lotes
                If Not comprobarAlbaranViculado Then Exit Function
                
            End If
        End If
    End If
    
    'Compriobaremos que el usuario 1 solo factura en 1
    'y el de 2 que avise
    CtaBancoPropi = ""
    If Not vUsu.TrabajadorB Then
        'NO ES DE B. Comprobaremos que todas las lineas son de almacen 1
        CtaBancoPropi = DevuelveDesdeBD(conAri, "count(*)", "sliped", "codalmac=" & vParamAplic.AlmacenB & " and numpedcl", CStr(Data2.Recordset!numpedcl))
        If CtaBancoPropi = "" Then CtaBancoPropi = "0"
        If Val(CtaBancoPropi) > 0 Then
            MsgBox "Hay lineas de pedido de distintos almacenes", vbExclamation
            CtaBancoPropi = ""
            Exit Function
        Else
            CtaBancoPropi = ""
        End If
    Else
        'Para el trabjador de B avisamos
        CtaBancoPropi = DevuelveDesdeBD(conAri, "count(*)", "sliped", "codalmac<>" & vParamAplic.AlmacenB & " and numpedcl", CStr(Data2.Recordset!numpedcl))
        If CtaBancoPropi = "" Then CtaBancoPropi = "0"
        If Val(CtaBancoPropi) > 0 Then
            If MsgBox("Hay lineas de pedido de distinto almacen al asignado. ¿Continuar?", vbQuestion + vbYesNo) = vbYes Then CtaBancoPropi = ""
        Else
            CtaBancoPropi = ""
        End If
        If CtaBancoPropi <> "" Then
            CtaBancoPropi = ""
            Exit Function
        End If
    End If
        
        
    'Empresa BODEGA Moixent
    ' Si tienen articulos de granel DEBE pasar a albaran para asignar el Hectogrado
    If vParamAplic.QUE_EMPRESA = 2 Then
    
        'HECtogrado
        'Los articulos que llevan hectogrado (Cod. Modelo--> sartic|codtipar| )
        ' el codtipar='05'
        CtaBancoPropi = "select count(*) from sliped,sartic where sliped.codartic=sartic.codartic"
        CtaBancoPropi = CtaBancoPropi & " AND numpedcl=" & Text1(0).Text
        CtaBancoPropi = CtaBancoPropi & " and (codtipar='05' or codfamia =6)"
        
        If RegistrosAListar(CtaBancoPropi) > 0 Then
            
            
            CtaBancoPropi = "Venta de articulos donde debe indicar el hectogrado."
            If Factura Then
                CtaBancoPropi = CtaBancoPropi & vbCrLf & "Pase a albaran primero"
            Else
                CtaBancoPropi = CtaBancoPropi & vbCrLf & " Hágalo en el albaran "
            End If
            
            MsgBox CtaBancoPropi, vbExclamation
            CtaBancoPropi = ""
            If Factura Then Exit Function
        
        
            
        End If
        
    End If
        
        
        
        
    If Factura Then
        'Si llega aqui. Veo si lleva articulos de trazabilidad (casi seguro)
        'y aviso
        CtaBancoPropi = "select count(*) from sliped,sartic where sliped.codartic=sartic.codartic"
        CtaBancoPropi = CtaBancoPropi & " and trazabilidad =1 and numpedcl=" & Text1(0).Text
        If RegistrosAListar(CtaBancoPropi) > 0 Then
            CtaBancoPropi = "Existen artículos de trazabilidad y debería indicar los lotes(an el albarán). ¿Continuar igualmente?"
            If MsgBox(CtaBancoPropi, vbQuestion + vbYesNo) = vbNo Then
                CtaBancoPropi = ""
                Exit Function
            Else
                CtaBancoPropi = ""
            End If
        End If
    End If
    
    'Llegado aqui: bien
    ComprobarOpcionTraspaso2 = True
End Function


Private Sub mnGeneraFactura_Click()

'    Dim L As Long
'    If Me.Data1.Recordset Is Nothing Then Exit Sub
'    If Me.Data1.Recordset.EOF Then Exit Sub
'
'    L = Val(Data1.Recordset!numpedcl)
'    If Not BloqueoManual("GenPed", CStr(L)) Then Exit Sub
    
    GeneraFacturaClick
    
    
'    DesBloqueoManual "GenPed"

End Sub

Private Sub GeneraFacturaClick()

Dim b As Boolean

   'Comprobaciones iniciales
   '----------------------------------------------------------------------------
   If Not ComprobarOpcionTraspaso2(True) Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    'Solo se generan albarenes completos
    AlbCompleto = True
    
    'comprobar si hay control de stock si se puede servir el pedido
    b = SePuedeServirPedido
        
    If b Then 'Hay suficiente stock
        'Si hay stock generar albaran completo
        GenerarAlbaran True
    Else
        Screen.MousePointer = vbDefault
        'Si no se puede servir mostrar mensaje detallando y bloquear
        TituloLinea = "No hay suficiente Stock para servir el Pedido. "
        TituloLinea = TituloLinea & vbCrLf & "¿Desea Ver Detalle?"
        If MsgBox(TituloLinea, vbYesNo, "Contol de Stock") = vbYes Then
            frmMensajes.cadWhere = " WHERE numpedcl = " & Text1(0).Text & " And sfamia.instalac = 0 "
            frmMensajes.vCampos = NomTablaLineas
            frmMensajes.OpcionMensaje = 2 'Articulos sin Stock
            frmMensajes.Show vbModal
        End If
        TituloLinea = ""
    End If


End Sub

Private Sub mnImpOrde_Click()
'Impreme la Orden de Instalacion de un pedido
Dim cadFormula As String, Cadparam As String
Dim Devuelve As String, nomDocu As String
Dim NumParam As Byte

    'Comprobar que hay un pedido seleccionado
    If Text1(0).Text = "" Then
        MsgBox "No hay ningún Pedido seleccionado.", vbInformation
        Exit Sub
    End If

'    'Comprobar que algun Articulo pertenece a la familia de Instalaciones
'    If Not PedidoConInstalaciones Then
'        MsgBox "El Pedido no tiene ningún Artículo que sea Instalación.", vbInformation
'        Exit Sub
'    End If

    '=======================================================================
    '=============== FORMULA    ============================================
    cadFormula = ""
    Cadparam = ""
    NumParam = 0
    
    If Text1(0).Text <> "" Then 'Seleccionar el Pedido
        Devuelve = "{" & NombreTabla & ".numpedcl}=" & Val(Text1(0).Text)
        If Not AnyadirAFormula(cadFormula, Devuelve) Then Exit Sub
    End If
    
    Devuelve = "{sartic.conjunto}=1"
    If Not AnyadirAFormula(cadFormula, Devuelve) Then Exit Sub
    
    
    'Seleccionar solo las lineas de Articulos que son de una familia que es Instalacion
    'Devuelve = "{sfamia.instalac}=1"
    'If Not AnyadirAFormula(cadFormula, Devuelve) Then Exit Sub
    
    If Not PonerParamRPT(9, Cadparam, NumParam, nomDocu) Then Exit Sub

    With frmImprimir
        .NombreRPT = nomDocu
        .FormulaSeleccion = cadFormula
        .OtrosParametros = Cadparam
        .NumeroParametros = NumParam
        .SoloImprimir = False
        .EnvioEMail = False
        .opcion = 39
        .Titulo = ""
        .Show vbModal
    End With
End Sub


Private Sub mnImpPedido_Click()
'Imprime un Pedido
       frmListadoOfer.NumCod = Text1(0).Text   'Nº de Pedido
       frmListadoOfer.CodClien = Text1(4).Text 'cliente del pedido
       If EsHistorico Then
            frmListadoOfer.FecEntre = Text1(1).Text   'Fecha de Pedido
            AbrirListadoOfer (239) '239: Informe de Pedidos (Historico
       Else
            AbrirListadoOfer (38) '38: Informe de Pedidos
       End If
End Sub

Private Sub mnLineas_Click()
    BotonMtoLineas 0, "Pedidos"
End Sub


Private Sub mnModificar_Click()
    If Modo = 5 Then 'Modificar lineas
         BotonModificarLinea
    Else   'Modificar Pedido
         If BLOQUEADesdeFormulario(Me) Then BotonModificar
    End If
End Sub


Private Sub mnNuevo_Click()
    If Modo = 5 Then 'Añadir lineas
         BotonAnyadirLinea
    Else 'Añadir Cabecera de Pedidos
         Me.SSTab1.Tab = 0
         BotonAnyadir
    End If
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

Private Sub mnTodosLosAlmacenes_Click()
    mnTodosLosAlmacenes.Checked = Not mnTodosLosAlmacenes.Checked
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub


Private Sub Text1_Change(Index As Integer)
    If Index = 9 Then HaCambiadoCP = True
End Sub

'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    If Index = 31 Then Exit Sub
    txtAnterior = Text1(Index).Text
    kCampo = Index
    If Index = 9 Then HaCambiadoCP = False 'CPostal
    ConseguirFoco Text1(Index), Modo
    If Index = 3 And Text1(3).Text <> "" Then PonerFoco Text1(4)
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Index <> 31 Then KEYdown KeyCode
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
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
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
       
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
                'Obtener la semana de Entrega
                Text1(18).Text = CalculaSemana(CDate(Text1(2).Text))
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
                Else 'Insertando
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
                Text2(12).Text = ""
                Exit Sub
            End If
            
            Text1(Index).Text = Format(Text1(Index).Text, "000")
            
            If PonerDptoEnCliente Then
            'Comprobar que el cliente seleccionada tiene esa direccion
                'Comprobar que el cliente tiene mantenimientos en esa direc/dpto
                Devuelve = DevuelveDesdeBDNew(conAri, "scaman", "nummante", "codclien", Text1(4).Text, "N", , "coddirec", Text1(12).Text, "N")
                If Devuelve <> "" Then MsgBox "El cliente tiene Mantenimientos.", vbInformation
            Else
                PonerFoco Text1(Index)
            End If
            
        Case 13 'Referencia Obligatoria
            If Trim(Text1(4).Text) <> "" Then ComprobarRefObligatoria
            
        Case 14 'Forma de Pago
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sforpa", "nomforpa")
            Else
                Text2(Index).Text = ""
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
Dim C3 As String
    cadB = ObtenerBusqueda(Me, False)
    If cadB = "" Then Exit Sub
    C3 = DevuelveListaPedidos
    cadB = cadB & " AND " & C3
   
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
        Cad = Cad & ParaGrid(Text1(0), 15, "Nº Pedido")
        Cad = Cad & ParaGrid(Text1(1), 20, "Fecha Ped.")
        Cad = Cad & ParaGrid(Text1(4), 15, "Cliente")
        Cad = Cad & ParaGrid(Text1(5), 50, "Nombre Cliente")
        Tabla = NombreTabla
        If EsHistorico Then
            Titulo = "Histórico de Pedidos"
            Devuelve = "0|1|"
        Else
            Titulo = "Pedidos"
            Devuelve = "0|"
        End If
        
    Else
        If vParamAplic.Departamento Then
            Titulo = "Dptos Cliente: "
            Desc = "Dpto."
        Else
            Titulo = "Direc. Cliente: "
            Desc = "Direc."
        End If
        Titulo = Titulo & Text1(4).Text & " - " & Text1(5).Text
        Cad = Cad & "Cod. " & Desc & "|sdirec|coddirec|N||15·"
        Cad = Cad & "Desc. " & Desc & "|sdirec|nomdirec|T||35·"
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
        frmB.vConexionGrid = conAri 'Conexión a BD: Ariges
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
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        If Modo = 1 Then
            Me.cboFacturacion.ListIndex = -1
            PonerFoco Text1(kCampo)
'            Text1(0).BackColor = vbYellow
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
'Carga las Pestañas con las tablas de lineas del Trabajador seleccionado para mostrar
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
    
    'Poner Nombre del Trabajador
    Text2(3).Text = PonerNombreDeCod(Text1(3), conAri, "straba", "nomtraba")
    'Poner Desc. del Dpto/Direc.
    Text2(12).Text = DevuelveDesdeBDNew(conAri, "sdirec", "nomdirec", "codclien", Text1(4).Text, "N", , "coddirec", Text1(12).Text, "N")
    'Poner el Nombre del Agente
    Text2(17).Text = PonerNombreDeCod(Text1(17), conAri, "sagent", "nomagent")
    'Poner la Desc. de la Forma de Pago
    Text2(14).Text = PonerNombreDeCod(Text1(14), conAri, "sforpa", "nomforpa")
       
    PonerCamposLineas 'Pone los datos de las tablas de lineas de Ofertas
    
    If EsHistorico Then
        'poner datos de eliminacion
        Text2(27).Text = PonerNombreDeCod(Text1(27), conAri, "straba", "nomtraba", "codtraba")
        Text2(28).Text = PonerNombreDeCod(Text1(28), conAri, "sincid", "nomincid", "codincid")
    End If
    
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
    
    If Modo = 6 Then Me.lblIndicador.Caption = "Insertar Cant. Servidas"
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    b = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda2 <> "" Then
        cmdRegresar.visible = b
    Else
        cmdRegresar.visible = False
    End If
        
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
        
        
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    'Campo Numero de Albaran siempre bloqueado, excepto si estamos en modo de busqueda
    b = (Modo <> 1)
    BloquearTxt Text1(0), b, True
    'Bloquear los campos de Oferta
    BloquearTxt Text1(24), b
    BloquearTxt Text1(25), b
    'Referncia produccion
    BloquearTxt Text1(32), b


    'Campo Semana Se calcula automat., siempre bloqueado
    BloquearTxt Text1(18), True
    
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
    Me.chkVisadoRes.Enabled = b
    Me.chkServirCom.Enabled = b
    Me.chkRecogeClien.Enabled = b
    
    
    'Si no es modo lineas Boquear los TxtAux
    For I = 0 To txtAux.Count - 1
        BloquearTxt txtAux(I), (Modo <> 5)
    Next I
    BloquearTxt Text2(16), (Modo <> 5)
    
    
    '---------------------------------------------
    b = (Modo <> 0 And Modo <> 2 And Modo <> 5)
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
    b = (Modo = 5)
    Me.Label1(35).visible = b
    Me.Text2(16).visible = b
    BloquearTxt Text2(16), True
       
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
    
    
    

    
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
       
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario

EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Function DatosOk() As Boolean
'Comprueba si los datos de la cabecera son correctos antes de Insertar o Modificar el
'Pedido
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
        Devuelve = DevuelveDesdeBDNew(conAri, "sclien", "referobl", "codclien", Text1(4).Text, "N")
        If Devuelve = "1" And Text1(13).Text = "" Then 'Referencia Obligatoria
            MsgBox "La Referencia del Cliente es Obligatoria.", vbInformation
            PonerFoco Text1(13)
            b = False
        End If
    End If
    If Not b Then Exit Function
          
          
          
          
    If Modo = 4 Then
        If Not vParamAplic.EsAVAB Then
            'En morales
            If DBLet(Data1.Recordset!refproduccion, "N") > 0 Then
                If Val(Text1(4).Text) <> 1 Then
                    MsgBox "Pedido bloqueado. Codigo cliente(REF1)", vbExclamation
                    Exit Function
                End If
            End If
        End If
    End If
          
          
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLinea() As Boolean
'Comprueba si los datos de una linea son correctos antes de Insertar o Modificar
'una linea del Pedido
Dim b As Boolean
Dim I As Byte
Dim vArtic As CArticulo

    On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    b = True

    'Comprobar que los campos NOT NULL tienen valor
    For I = 0 To txtAux.Count - 1
        If I <> 11 Then
            If txtAux(I).Text = "" Then
                MsgBox "El campo " & txtAux(I).Tag & " no puede ser nulo", vbExclamation
                b = False
                PonerFoco txtAux(I)
                Exit Function
            End If
        End If
    Next I
        
    'Comprobar que existe el articulo en el almacen seleccionado
    Set vArtic = New CArticulo
    vArtic.Codigo = txtAux(1).Text
    If Not vArtic.ExisteEnAlmacen(txtAux(0).Text) Then
        b = False
        PonerFoco txtAux(1)
    End If
    Set vArtic = Nothing
    
    
    
    
    'Si esta vinculado a MORALES si es modificar, NO puede cambiar el articulo
    If ModificaLineas = 2 And b Then
        If DBLet(Data1.Recordset!refproduccion, "N") > 0 Then
            'OK esta vinculado AVAB-MORALES
            If Data2.Recordset!codartic <> txtAux(1).Text Then
                MsgBox "No puede cambiar la referencia de una linea en un pedido bloqueado", vbExclamation
                b = False
            End If
        End If
    End If
    
    
    
    'Contamos las lineas
    
    
    DatosOkLinea = b

EDatosOkLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub Text2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 16 And KeyCode = 40 Then 'campo Ampliacion linea y Flecha hacia abajo
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYdown KeyCode
    End If
End Sub


Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    'campo Ampliación linea y ENTER
    If Index = 16 And KeyAscii = 13 Then PonerFocoBtn Me.cmdAceptar
End Sub

Private Sub Text2_LostFocus(Index As Integer)
    If Index = 16 And (Text2(Index).Locked = False) Then Text2(Index).Text = UCase(Text2(Index).Text)
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Buscar
            mnBuscar_Click
        Case 2  'Todos
            mnVerTodos_Click
        Case 5  'Nuevo
            mnNuevo_Click
        Case 6  'Modificar
            mnModificar_Click
        Case 7  'Borrar
            mnEliminar_Click
            
        Case 10  'Lineas
            mnLineas_Click
        Case 11 'Generar Albaran
            mnGenAlbaran_Click
        Case 12
            'Genera la factura directamente
            mnGeneraFactura_Click
            
        Case 14 'Imprimir Pedido
            mnImpPedido_Click
            
        Case 15
             'Vincular pedido a producion
             VincularPedidoEnEmpresaProduccion
             
        Case 16 'Imprimir Orden Instalacion
             If vParamAplic.EsAVAB Then
                PackingList
             Else
                mnImpOrde_Click
            End If
            
        Case 18    'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub


Private Sub PonerOpcionesMenu()
Dim J As Byte

    PonerOpcionesMenuGeneral Me
      
    J = Val(Me.mnGenAlbaran.HelpContextID)
    If J < vUsu.Nivel Then Me.mnGenAlbaran.Enabled = False
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub
    
    
Private Function InsertarLinea() As Boolean
'Inserta un registro en la tabla de lineas de Pedido: slipre
Dim SQL As String
'Dim ImpReciclado As Single
Dim ImpReciclado As Currency
Dim numlinea As String, vWhere As String

    On Error GoTo EInsertarLinea

    InsertarLinea = False
    SQL = ""

    If DatosOkLinea() Then 'Lineas de Pedidos
        'Conseguir el siguiente numero de linea
        vWhere = Replace(ObtenerWhereCP, NombreTabla, NomTablaLineas)
        numlinea = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", vWhere)
        'Construir la sentencia SQL
'        vWhere = ObtenerWhereCP
        SQL = "INSERT INTO " & NomTablaLineas
        SQL = SQL & "(numpedcl,numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, servidas, precioar, dtoline1, dtoline2, importel, origpre,cajas,PrecioLitro,palets) "
        SQL = SQL & "VALUES (" & Val(Text1(0).Text) & ", " & numlinea & ", " & Val(txtAux(0).Text) & ","
        SQL = SQL & DBSet(txtAux(1).Text, "T") & ", " & DBSet(txtAux(2).Text, "T") & ", " & DBSet(Text2(16).Text, "T") & ", "
        SQL = SQL & DBSet(txtAux(3).Text, "N") & ", 0,"
        SQL = SQL & DBSet(txtAux(4).Text, "N") & ", " & DBSet(txtAux(6).Text, "N") & ", "
        SQL = SQL & DBSet(txtAux(7).Text, "N") & ", " 'Dto2
        SQL = SQL & DBSet(txtAux(8).Text, "N") & ", "
        SQL = SQL & DBSet(txtAux(5).Text, "T") & ","
        'cajas,PrecioLitro
        SQL = SQL & DBSet(txtAux(9).Text, "N") & ","
        SQL = SQL & DBSet(txtAux(10).Text, "N") & ","
        'Palets
        SQL = SQL & DBSet(txtAux(11).Text, "N", "S") & ")"
    End If
    
    If SQL <> "" Then
        conn.Execute SQL
        InsertarLinea = True
        
        
        'Si esta vinculado a un pedido
        If vParamAplic.EsAVAB Then
             If DBLet(Data1.Recordset!refproduccion, "N") > 0 Then
                 Espera 0.2
                 'Esta vinculado
                 Set miRsAux = New ADODB.Recordset
                 SQL = "Select * from sliped WHERE " & vWhere & " AND numlinea = " & numlinea
                 miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                 'NO PUEDE dar error
                 SQL = DevueleveSQLLineasPedido
                 SQL = "  VALUES (" & Data1.Recordset!refproduccion & SQL
                 SQL = "INSERT INTO ariges" & EmprMorales & ".sliped(numpedcl,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,servidas,precioar,dtoline1,dtoline2,origpre,importel,cajas,PrecioLitro,palets) " & SQL
                 EjecutaSQL conAri, SQL, True
                 Set miRsAux = Nothing
            End If
        End If
        
        If ClienteConTasaReciclado Then
            If ArticuloConTasaReciclado2(txtAux(1).Text, ImpReciclado) Then
                'Insertamos la linea del reciclado
                vWhere = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", vParamAplic.ArtReciclado, "T")
                SQL = "INSERT INTO " & NomTablaLineas
                SQL = SQL & "(numpedcl,numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, servidas, precioar,"
                SQL = SQL & "dtoline1, dtoline2, importel, origpre) "
                SQL = SQL & "VALUES (" & Val(Text1(0).Text) & ", " & numlinea + 1 & ", " & Val(txtAux(0).Text) & ","
                SQL = SQL & DBSet(vParamAplic.ArtReciclado, "T") & "," & DBSet(vWhere, "T") & ", Null, "
                SQL = SQL & DBSet(txtAux(3).Text, "N") & ", 0," 'Cantidad. La misma
                SQL = SQL & DBSet(ImpReciclado, "N") & ",0,0,"
                'Importe linea
                ImpReciclado = ImporteFormateado(txtAux(3).Text) * ImpReciclado
                SQL = SQL & DBSet(ImpReciclado, "N") & ", 'A')"
                conn.Execute SQL
                    
                
            End If
        End If
        
    End If
    Exit Function
    
EInsertarLinea:
    MuestraError Err.Number, "Insertar Lineas Pedido" & vbCrLf & Err.Description
End Function


Private Function ModificarLinea() As Boolean
'Modifica un registro en la tabla de lineas de Pedido: sliped
Dim SQL As String
Dim Aux As String
Dim UpdateaEnVinculada As Boolean

    On Error GoTo EModificarLinea

    ModificarLinea = False
    SQL = ""
    
    If DatosOkLinea() Then
        'Creamos la sentencia SQL
        SQL = "UPDATE " & NomTablaLineas & " Set codalmac = " & txtAux(0).Text & ", codartic=" & DBSet(txtAux(1).Text, "T") & ", "
        SQL = SQL & "nomartic=" & DBSet(txtAux(2).Text, "T") & ", ampliaci=" & DBSet(Text2(16).Text, "T") & ", "
        SQL = SQL & " cantidad = " & DBSet(txtAux(3).Text, "N") & ", "
        SQL = SQL & " precioar = " & DBSet(txtAux(4).Text, "N") & ", "
        SQL = SQL & "dtoline1= " & DBSet(txtAux(6).Text, "N") & ", dtoline2= " & DBSet(txtAux(7).Text, "N") & ", "
        SQL = SQL & "importel= " & DBSet(txtAux(8).Text, "N") & ", "
        SQL = SQL & "origpre=" & DBSet(txtAux(5).Text, "T") & ", "
                'cajas,PrecioLitro
        SQL = SQL & "cajas=" & DBSet(txtAux(9).Text, "N") & ", "
        SQL = SQL & "PrecioLitro=" & DBSet(txtAux(10).Text, "N") & ", "
        SQL = SQL & "Palets=" & DBSet(txtAux(11).Text, "N", "S")


        SQL = SQL & " WHERE " & Replace(ObtenerWhereCP, NombreTabla, NomTablaLineas) & " AND numlinea=" & Data2.Recordset!numlinea
        
        
        UpdateaEnVinculada = False
        If DBLet(Data1.Recordset!refproduccion, "N") > 0 Then
            'Si ha cambiado cajas/palets/cantidad entonces tengo que updatear en el vinculado
            If ImporteFormateado(txtAux(11).Text) <> DBLet(Data2.Recordset!palets, "N") Then UpdateaEnVinculada = True
            If ImporteFormateado(txtAux(3).Text) <> Data2.Recordset!Cantidad Then UpdateaEnVinculada = True
            If ImporteFormateado(txtAux(9).Text) <> Data2.Recordset!Cajas Then UpdateaEnVinculada = True
        End If
        
    End If
    
    If SQL <> "" Then
        conn.Execute SQL
        ModificarLinea = True
        
        If UpdateaEnVinculada Then UpdatearLineaPedidoEnOtraEmpresa
       
    End If
    Exit Function
    
EModificarLinea:
    MuestraError Err.Number, "Modificar Lineas Pedido" & vbCrLf & Err.Description
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
        Me.lblIndicador.Caption = "Líneas " & TituloLinea
        PonerFocoBtn Me.cmdRegresar
    End If
    
    'Habilitar las opciones correctas del menu según Modo
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu 'Habilitar las opciones correctas del menu según Nivel de Acceso
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean, Optional conServidas As Boolean)
'IN: enlaza= si carga el grid con valores de la tabla o lo muestra vacio si no enlaza
'    conServidas=si enlaza, se muestra la columna de servidas solo cuando se va a generar el Albaran no completo
Dim b As Boolean
Dim SQL As String

    On Error GoTo ECargaGrid

    b = DataGrid1.Enabled
    
    SQL = MontaSQLCarga(enlaza, conServidas)
    CargaGridGnral vDataGrid, vData, SQL, PrimeraVez
    
'    If PrimeraVez Or conServidas Then
    If conServidas Then
        vDataGrid.ClearFields
        vDataGrid.ReBind
        vDataGrid.Refresh
    End If
    
    CargaGrid2 vDataGrid, vData, conServidas
    vDataGrid.ScrollBars = dbgAutomatic
    
    b = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2) '5:Modo Mto Lineas (Insertando o Modificando linea)
    vDataGrid.Enabled = Not b
    PrimeraVez = False
    gridCargado = True
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, Optional conServidas As Boolean)
Dim I As Byte

    On Error GoTo ECargaGrid

    vData.Refresh
    vDataGrid.Columns(0).visible = False
    vDataGrid.Columns(1).visible = False

    Select Case vDataGrid.Name
        Case "DataGrid1" 'Cod. Almacen
                vDataGrid.Columns(2).Caption = "Alm."
                If conServidas Then
                    vDataGrid.Columns(2).Width = 450
                Else
                    vDataGrid.Columns(2).Width = 500
                End If
                vDataGrid.Columns(2).NumberFormat = "000"
                
                vDataGrid.Columns(3).Caption = "Articulo"
                If conServidas Then
                    vDataGrid.Columns(3).Width = 1600
                Else
                    vDataGrid.Columns(3).Width = 1700
                End If
                
                vDataGrid.Columns(4).Caption = "Desc. Artículo"
                If conServidas Then
                    vDataGrid.Columns(4).Width = 3000
                Else
                    vDataGrid.Columns(4).Width = 3300
                End If
                
                vDataGrid.Columns(5).visible = False
                
                
                'Abril 2009
                'Se añade cajas y precio litro
                'Mayo 2011
                'Se añade palets
                vDataGrid.Columns(6).Caption = "Palets"
                vDataGrid.Columns(6).Width = 650
                vDataGrid.Columns(6).Alignment = dbgRight
                
                vDataGrid.Columns(7).Caption = "Cajas"
                vDataGrid.Columns(7).Width = 650
                vDataGrid.Columns(7).Alignment = dbgRight
                                
                                
                vDataGrid.Columns(8).Caption = "Cantidad"
                vDataGrid.Columns(8).Width = 850
                vDataGrid.Columns(8).Alignment = dbgRight
                vDataGrid.Columns(8).NumberFormat = FormatoImporte
                
                If conServidas Then
                    'Cargar el grid con la columna de cantidad servida
                    vDataGrid.Columns(9).Caption = "Caja ser."
                    vDataGrid.Columns(9).Width = 800
                    vDataGrid.Columns(9).Alignment = dbgRight
                    
                    
                    
                    vDataGrid.Columns(10).Caption = "Servidas"
                    vDataGrid.Columns(10).Width = 800
                    vDataGrid.Columns(10).Alignment = dbgRight
                    vDataGrid.Columns(10).NumberFormat = FormatoImporte
                    
                    I = 11
                Else
                    I = 9
                End If
                
                vDataGrid.Columns(I).Caption = "Precio Ud"
                vDataGrid.Columns(I).Width = 1000
                vDataGrid.Columns(I).Alignment = dbgRight
                vDataGrid.Columns(I).NumberFormat = FormatoPrecio
                
                'nuevo precio litro para formatos menores
                I = I + 1
                vDataGrid.Columns(I).Caption = "Precio Lit"
                vDataGrid.Columns(I).Width = 1000
                vDataGrid.Columns(I).Alignment = dbgRight
                vDataGrid.Columns(I).NumberFormat = FormatoPrecio
                
                
                
                
                
                vDataGrid.Columns(I + 1).Caption = "OP"
                vDataGrid.Columns(I + 1).Width = 350
                vDataGrid.Columns(I + 1).Alignment = dbgCenter
                
                    
                vDataGrid.Columns(I + 2).Caption = "Dto.1"
                If conServidas Then
                    vDataGrid.Columns(I + 2).Width = 550
                Else
                    vDataGrid.Columns(I + 2).Width = 600
                End If
                vDataGrid.Columns(I + 2).Alignment = dbgRight
                vDataGrid.Columns(I + 2).NumberFormat = FormatoDescuento
                
                vDataGrid.Columns(I + 3).Caption = "Dto.2"
                If conServidas Then
                    vDataGrid.Columns(I + 3).Width = 550
                Else
                    vDataGrid.Columns(I + 3).Width = 600
                End If
                vDataGrid.Columns(I + 3).Alignment = dbgRight
                vDataGrid.Columns(I + 3).NumberFormat = FormatoDescuento
            
                vDataGrid.Columns(I + 4).Caption = "Importe Línea"
                If conServidas Then
                    vDataGrid.Columns(I + 4).Width = 550
                Else
                    vDataGrid.Columns(I + 4).Width = 1400
                End If
                vDataGrid.Columns(I + 4).Alignment = dbgRight
                vDataGrid.Columns(I + 4).NumberFormat = FormatoImporte
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
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single
Dim I As Byte

    On Error Resume Next

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For I = 0 To txtAux.Count - 1 'TextBox
            txtAux(I).Top = 290
            txtAux(I).visible = visible
        Next I
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
                Select Case I
                Case 0 To 2
                    txtAux(I).Text = DataGrid1.Columns(I + 2).Text
                Case 3, 4
                    txtAux(I).Text = DataGrid1.Columns(I + 5).Text
                Case 10
                    txtAux(I).Text = DataGrid1.Columns(10).Text
                Case 11
                    txtAux(I).Text = DataGrid1.Columns(6).Text
                Case 5 To 8
                    '5 a 8
                    txtAux(I).Text = DataGrid1.Columns(I + 6).Text
                Case 9
                    txtAux(I).Text = DataGrid1.Columns(7).Text
                End Select
                txtAux(I).Locked = False
            Next I
        End If
               
        'El Campo de Origen del precio se actualiza por programa al modificar el precio
        BloquearTxt txtAux(5), True
        'El campo Importe es calculado y lo bloqueamos.
        BloquearTxt txtAux(8), True
    

        'Fijamos altura(Height) y posición Top
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
        
        'MAyo 2011
        txtAux(11).Left = txtAux(2).Left + txtAux(2).Width + 10
        txtAux(11).Width = DataGrid1.Columns(6).Width - 10
        'Abril 2009
        'Cajas
        txtAux(9).Left = txtAux(11).Left + txtAux(11).Width + 10
        txtAux(9).Width = DataGrid1.Columns(7).Width - 10
        
        
        'Cantidad
        txtAux(3).Left = txtAux(9).Left + txtAux(9).Width + 10
        txtAux(3).Width = DataGrid1.Columns(8).Width - 10
        
        'Precio UD
        txtAux(4).Left = txtAux(3).Left + txtAux(3).Width + 10
        txtAux(4).Width = DataGrid1.Columns(9).Width - 10
        
        'Precio Litr
        txtAux(10).Left = txtAux(4).Left + txtAux(4).Width + 10
        txtAux(10).Width = DataGrid1.Columns(10).Width - 10
        
        'Origen PRE
        txtAux(5).Left = txtAux(10).Left + txtAux(10).Width + 10
        txtAux(5).Width = DataGrid1.Columns(11).Width - 10
        
        '  Dto 1 Dto2, importe lin
        For I = 6 To 8
            txtAux(I).Left = txtAux(I - 1).Left + txtAux(I - 1).Width + 10
            txtAux(I).Width = DataGrid1.Columns(I + 6).Width - 10
        Next I
        
        'Los ponemos Visibles o No
        '--------------------------
        For I = 0 To txtAux.Count - 1
            txtAux(I).visible = visible
        Next I
        cmdAux(0).visible = visible
        cmdAux(1).visible = visible
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaTxtAuxServidas(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
'Carga el TxtAux(3) con el campo servidas de la tabla sliped
Dim alto As Single
Dim I As Byte   'Cantidad serv
Dim C As Byte   'CAJAS serv
    On Error Resume Next

    I = 3
    C = 9
    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        txtAux(I).Top = 290
        txtAux(I).visible = visible
        txtAux(I).BackColor = vbWhite
        txtAux(C).Top = 290
        txtAux(C).visible = visible
        txtAux(C).BackColor = vbWhite

    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid1
            txtAux(I).Text = ""
            txtAux(C).Text = ""
            BloquearTxt txtAux(I), False
            BloquearTxt txtAux(C), False
            txtAux(I).BackColor = &HC0C0C0
            txtAux(C).BackColor = &HC0C0C0
        End If

        'Fijamos altura(Height) y posición Top
        '-------------------------------
        If DataGrid1.Row < 0 Then
            alto = DataGrid1.Top + 230
        Else
            alto = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + 10
        End If

        txtAux(I).Top = alto
        txtAux(I).Height = DataGrid1.RowHeight
        txtAux(C).Top = alto
        txtAux(C).Height = DataGrid1.RowHeight
        'Fijamos anchura y posicion Left
        '--------------------------------
        'Cantidad servida
'        alto = DataGrid1.Left + 330 + DataGrid1.Columns(2).Width + DataGrid1.Columns(3).Width
'        alto = alto + DataGrid1.Columns(4).Width + DataGrid1.Columns(6).Width
'        txtAux(i).Left = alto + 10
'        txtAux(i).Width = DataGrid1.Columns(7).Width - 30

        alto = DataGrid1.Columns(9).Left + DataGrid1.Left
        txtAux(C).Left = alto + 10
        txtAux(C).Width = DataGrid1.Columns(9).Width

        alto = DataGrid1.Columns(10).Left + DataGrid1.Left
        txtAux(I).Left = alto + 10
        txtAux(I).Width = DataGrid1.Columns(10).Width



        'Los ponemos Visibles o No
        '--------------------------
        txtAux(I).visible = visible
        txtAux(C).visible = visible
        PonerFoco txtAux(C)
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub TxtAux_Change(Index As Integer)
    If Index = 4 And ModificaLineas = 2 Then 'Precio y Modo Modificar Lineas
        txtAux(5).Text = "M"
    End If
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
Dim cadkey As Integer

    cadkey = ObtenerCadKey(kCampo, Index)
    kCampo = Index
    ConseguirFocoLin txtAux(Index), cadkey
    txtAnterior = txtAux(Index).Text
End Sub




Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Modo <> 6 Then 'Modo6: Pasar de Pedido a Albaran
        If Not (Index = 0 And KeyCode = 38) Then KEYdown KeyCode
    Else 'Modo lineas
        Select Case KeyCode
            Case 38 'Desplazamieto Fecha Hacia Arriba
'                    If DataGrid1.Row > 0 Then
'                        DataGrid1.Row = DataGrid1.Row - 1
'                        CargaTxtAuxServidas True, True
'                    Else
'                        PonerFoco txtAux(3)
'                    End If
'                    txtAux(3).Text = Data2.Recordset!servidas
'                    txtAux(9).Text = Data2.Recordset!cajserv
'                    FijarUdsCaja
'                    ConseguirFocoLin txtAux(3)

            Case 40 'Desplazamiento Flecha Hacia Abajo
'                    If DataGrid1.Row < Data2.Recordset.RecordCount - 1 Then
                    If Index = 3 Then PonerServidas
'                    MoverSigRegisros
'                    If Data2.Recordset.AbsolutePosition <= Data2.Recordset.RecordCount - 1 Then
'                        DataGrid1.Row = DataGrid1.Row + 1
'                        CargaTxtAuxServidas True, True
'                    Else
'                        PonerFocoBtn Me.cmdAceptar
'                    End If
'                    txtAux(3).Text = Data2.Recordset!servidas
'                    ConseguirFocoLin txtAux(3)
        End Select
    End If
End Sub


Private Sub MoverSigRegistro()
    On Error GoTo EMover
    
    If Data2.Recordset.EOF Then Exit Sub
    If Data2.Recordset.AbsolutePosition <= Data2.Recordset.RecordCount - 1 Then
        DataGrid1.Row = DataGrid1.Row + 1
        CargaTxtAuxServidas True, True
    Else
        PonerFocoBtn Me.cmdAceptar
    End If
    txtAux(3).Text = Data2.Recordset!servidas
    txtAux(9).Text = Data2.Recordset!cajserv
    FijarUdsCaja
    ConseguirFocoLin txtAux(9)
EMover:
    If Err.Number <> 0 Then MuestraError Err.Description, "Mover registro.", Err.Description
End Sub


Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If Modo <> 6 Then
        KEYpress KeyAscii
    Else 'Modo 6: Pasar el Pedido a Albaran
        If KeyAscii = 13 Then  'ENTER
            If Index = 3 Then
                PonerServidas
                ConseguirFoco txtAux(3), Modo
            Else
                PonerFoco txtAux(3)
        
            End If
        End If
    End If
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim Devuelve As String, cadMen As String
Dim codTarif As String
Dim CPrecioFact As CPreciosFact
Dim vCStock As CStock
Dim NumCajas As Long, RestoUnid As Long
Dim OrigP As String 'De donde viene el precio
Dim b As Boolean

    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then
        If Modo = 6 Then
            
            If Index = 9 Or Index = 3 Then
                
                If Index = 9 Then
                    If txtAnterior = txtAux(9).Text Then Exit Sub
                    If Not PonerFormatoEntero(txtAux(9)) Then txtAux(9).Text = ""
                    If txtAux(9).Text = "" Then
                        txtAux(3).Text = ""
                    Else
                        'Ponemos las uds
                        
                        CantidadCajasServidas False
                        'Guardar
                        'PonerServidas
                    End If
                    
                Else
                    If txtAnterior = txtAux(3).Text Then Exit Sub
                    If Not PonerFormatoDecimal(txtAux(3), 3) Then txtAux(3).Text = ""
                    If txtAux(3).Text = "" Then
                        txtAux(9).Text = ""
                    Else
                        CantidadCajasServidas True
                        PonerFoco txtAux(9)
                    End If
                End If
                Exit Sub
            End If
        End If
    End If
    
    Select Case Index
        Case 0 'Cod Almacen
            'Comprobar que existe el almacen
            Devuelve = ""
            If vUsu.TrabajadorB Then
                If txtAux(0).Text <> vParamAplic.AlmacenB Then Devuelve = "U"
            Else
                If txtAux(0).Text = vParamAplic.AlmacenB Then Devuelve = "U"
            End If
            If Devuelve <> "" Then
                MsgBox "Usuario / almacen incorrecto", vbExclamation
                txtAux(0).Text = ""
                PonerFoco txtAux(0)
                Exit Sub
            End If
            Devuelve = PonerAlmacen(txtAux(Index).Text)
            txtAux(Index).Text = Devuelve
            If Devuelve = "" Then PonerFoco txtAux(Index)

        Case 1 'Cod. Articulo
            If txtAux(1).Text = "" Then 'Cod Artic
                txtAux(2).Text = "" 'Nom Artic
                Exit Sub
            End If
            If txtAux(0).Text = "" Then 'Cod Almacen
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
                
                
                If ElArticulo Is Nothing Then Set ElArticulo = New CArticulo
                
                If ElArticulo.Codigo <> txtAux(1).Text Then ElArticulo.LeerDatos txtAux(1).Text
                
                'Por si acaso ha cambiado el articulo
                If Devuelve <> txtAux(1).Text Then
                    PrecioUdLitro True
                    CantidadCajas True
                End If
                
                If Not b Then
'                    If txtAux(2).Locked Then PonerFoco txtAux(3)
        
                   
                Else
                    PonerFoco txtAux(0)
                End If
            Else
                b = False
                PonerFoco txtAux(Index)
            End If
        
        Case 2 'desc Articulo
            If txtAux(Index).Locked = False Then txtAux(Index).Text = UCase(txtAux(Index).Text)
            
        Case 3 'CANTIDAD
            If PonerFormatoDecimal(txtAux(Index), 1) Then  'Tipo 1: Decimal(12,2)
                If Modo = 5 Then 'Mantenimiento lineas
                    'Comprobar si hay suficiente stock
                    Set vCStock = New CStock
                    If Not InicializarCStock(vCStock, "S") Then Exit Sub
                    If vCStock.MueveStock Then
                        If Not vCStock.MoverStock(False) Then
                            Set vCStock = Nothing
                            Exit Sub
                        End If
                    End If
                    
                    
                    b = False
                    If ModificaLineas = 1 Then 'insertar linea
                        b = True
                    ElseIf ModificaLineas = 2 Then 'modificar linea
                        If Data2.Recordset!codartic <> txtAux(1).Text Then b = True
                    End If
                    
                    If b Then 'Modo Insertar en Mto Lineas
                        'Obtener el precio correspondiente y los descuentos
                        'Comprobar si el articulo se vende por cajas antes de entrar a la función
                        Devuelve = DevuelveDesdeBDNew(conAri, "sartic", "unicajas", "codartic", txtAux(1).Text, "T")
                        If Devuelve <> "" Then
                            Set CPrecioFact = New CPreciosFact
                            'Si se puede vender por cajas(devuelve>1) poner numero de cajas en una linea con el
                            'precio de caja, y otra linea con el resto unidades un precio unidad
                            NumCajas = CPrecioFact.ObtenerNumCajas(vCStock.Cantidad, Devuelve)
                            RestoUnid = CLng(vCStock.Cantidad) - NumCajas * CLng(Devuelve)
                            'Obtenemos la Tarifa del Cliente
                            codTarif = DevuelveDesdeBDNew(conAri, "sclien", "codtarif", "codclien", Text1(4).Text, "N")
                            CPrecioFact.CodigoLista = codTarif
                            CPrecioFact.CodigoArtic = vCStock.codartic
                            CPrecioFact.CodigoClien = Text1(4).Text
                            PorCaja = (NumCajas > 0)
                            Precio = CPrecioFact.ObtenerPrecio(PorCaja, Text1(1).Text, OrigP)
                            'Si PorCaja vuelve de ObtenerPrecio a false se calcula con precio unidad aunque NumCajas>0
                            'Ya que a regresado con pvp del Articulo
                            If PorCaja And NumCajas > 0 And RestoUnid > 0 Then
                                cadMen = "El Artículo puede venderse por Cajas (" & Devuelve & "uds. por Caja)." & vbCrLf
                                cadMen = cadMen & vbCrLf & "Inserte dos Lineas:   "
                                cadMen = cadMen & vbCrLf & "   Linea 1:  " & NumCajas * CInt(Devuelve) & " uds a Precio Caja"
                                cadMen = cadMen & vbCrLf & "   Linea 2:  " & CInt(vCStock.Cantidad) - NumCajas * CInt(Devuelve) & " uds a Precio Unidad"
                                MsgBox cadMen, vbInformation
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
                                
                                
                                'Pondere el foco en precio litro si es mayor que un litro
                                RestoUnid = 4
                                If Not (ElArticulo Is Nothing) Then
                                    If ElArticulo.LitrosxUd > 1 Then RestoUnid = 10
                                End If
                                PonerFoco txtAux(RestoUnid)
                                RestoUnid = 0
                            End If
    '                        ConseguirFoco txtAux(Index + 1), Modo
                            Set CPrecioFact = Nothing
                        End If
                    End If
                    ConseguirFocoLin txtAux(4)
    '            End If
                Set vCStock = Nothing
            End If
        End If
            
        Case 4 'PRECIO
             If txtAux(Index).Text <> "" Then
                PonerFormatoDecimal txtAux(Index), 2 'Tipo 2: Decimal(10,4)
                If ModificaLineas = 1 Then
                    'Precio=valor devuelto por la funcion de precios
                    If CSng(txtAux(Index).Text) <> CSng(ComprobarCero(Precio)) Then txtAux(5).Text = "M"
                End If
            End If
            PrecioUdLitro True
        Case 6, 7 'Descuentos
            PonerFormatoDecimal txtAux(Index), 4 'Tipo 4: Decimal(4,2)
        Case 8 'Importe Linea
            PonerFormatoDecimal txtAux(Index), 1 'Tipo 3: Decimal(12,2)
            
            
        'Abril 2009
        Case 9
               'cajas
            txtAnterior = txtAux(3).Text
            If txtAux(9).Text <> "" Then
                If Not PonerFormatoEntero(txtAux(9)) Then
                    txtAux(9).Text = ""
                    PonerFoco txtAux(9)
                Else
                    CantidadCajas False
                    If txtAnterior <> txtAux(3).Text Then PonerFoco txtAux(3)
                    
                End If
            Else
                txtAux(3).Text = ""
            End If
            
        Case 10
            PonerFoco txtAux(6)
            txtAnterior = txtAux(4).Text
            If txtAux(10).Text <> "" Then
                If Not PonerFormatoDecimal(txtAux(10), 2) Then
                    txtAux(10).Text = ""
                    PonerFoco txtAux(10)
                Else
                   
                    PrecioUdLitro False
                    If txtAnterior <> txtAux(4).Text Then PonerFoco txtAux(4)
    
                End If
            Else
                txtAux(4).Text = ""
            End If
        Case 11
            If Not PonerFormatoEntero(txtAux(Index)) Then
                If txtAux(Index).Text <> "" Then PonerFoco txtAux(Index)
                txtAux(Index).Text = ""
                
                
            Else
                'Si tiene articulo y NO tiene las cajas puestas
                If txtAux(9).Text = "" And txtAux(1).Text <> "" Then
                    Devuelve = DevuelveDesdeBD(conAri, "pal_udbas * pal_udalt", "sarti4", "codartic", txtAux(1).Text, "T")
                    If Devuelve = "" Then Devuelve = "0"
                    RestoUnid = Val(Devuelve) * CInt(txtAux(Index).Text)
                    txtAux(9).Text = RestoUnid
                End If
            End If
            
    End Select
    
    If Modo = 5 Then 'Modo Lineas
         If (Index = 3 Or Index = 4 Or Index = 6 Or Index = 7) Then 'Cant., Precio, dto1, dto2
            If txtAux(1).Text = "" Then Exit Sub 'Cod artic
            txtAux(8).Text = CalcularImporte(txtAux(3).Text, txtAux(4).Text, txtAux(6).Text, txtAux(7).Text, vParamAplic.TipoDtos)
            PonerFormatoDecimal txtAux(8), 1
        End If
    End If
End Sub


Private Sub BotonMtoLineas(numTab As Integer, Cad As String)

        If DBLet(Data1.Recordset!refproduccion, "N") > 0 Then MsgBox "Pedido bloqueado.", vbExclamation
            

        Me.SSTab1.Tab = numTab
        TituloLinea = Cad
        ModificaLineas = 0
        If vParamAplic.ArtReciclado <> "" Then
            ClienteConTasaReciclado = Val(DevuelveDesdeBD(conAri, "tasareciclado", "sclien", "codclien", Text1(4).Text)) = 1
        Else
            ClienteConTasaReciclado = False
        End If
        
        PonerModo 5
        'Si tiene lineas, mostrare la ampliacion en el campo de ampliacion
        PonerAmpliacionAlcambiarAlineas
        PonerBotonCabecera True
End Sub


Private Sub PonerAmpliacionAlcambiarAlineas()
Dim C As String
    C = ""
    If Not Data2.Recordset Is Nothing Then
        If Not Data2.Recordset.EOF Then C = DevuelveDesdeBDNew(conAri, NomTablaLineas, "ampliaci", "numpedcl", Text1(0).Text, "N", , "numlinea", Data2.Recordset!numlinea, "N")
    End If
    'Poner descripcion de ampliacion lineas
    Text2(16).Text = C
End Sub


Private Function Eliminar() As Boolean
Dim b As Boolean
Dim SQL As String
Dim MenError As String
Dim vTipoMov As CTiposMov

    On Error GoTo FinEliminar

        conn.BeginTrans
        SQL = ObtenerWhereCP
        
        'CadenaSQL: datos introducidos en el form de eliminacion
        b = ActualizarElTraspaso(MenError, SQL, CodTipoMov, CadenaSQL)

        If b Then
            If DBLet(Data1.Recordset!refproduccion, "N") > 0 Then
                conn.Execute "DELETE FROM ariges" & EmprMorales & ".sliped where numpedcl= " & Val(Data1.Recordset!refproduccion)
                conn.Execute "DELETE FROM ariges" & EmprMorales & ".scaped where numpedcl= " & Val(Data1.Recordset!refproduccion)
            End If
        End If
        If b Then
            'Devolvemos contador, si no estamos actualizando
            Set vTipoMov = New CTiposMov
            b = vTipoMov.DevolverContador(CodTipoMov, Data1.Recordset.Fields(0).Value)
            Set vTipoMov = Nothing
        End If
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Pedido" & vbCrLf & MenError, Err.Description
        b = False
    End If
    If Not b Then
        conn.RollbackTrans
        Eliminar = False
    Else
        conn.CommitTrans
        Eliminar = True
    End If
End Function


Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next
    CargaGrid DataGrid1, Data2, False
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PosicionarData()
'Despues de hacer refresh del Data, volver a situar el Data en el registro que estaba
Dim Indicador As String
Dim vWhere As String

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = Replace(ObtenerWhereCP, NombreTabla & ".", "")
         If SituarData(Data1, vWhere, Indicador) Then
             PonerModo 2
             PonerCampos
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


Private Sub PosicionarDataTrasEliminar()
'Despues Eliminar y hacer refresh del Data, situar el Data en el registro siguiente
    If SituarDataTrasEliminar(Data1, NumRegElim) Then
        PonerCampos
    Else
        LimpiarCampos
        LimpiarDataGrids
        PonerModo 0
    End If
End Sub


Private Function ObtenerWhereCP() As String
'Obtiene la where de la Clave Primaria de la tabla de Cabecera: scaped
Dim SQL As String

    On Error Resume Next
    
    SQL = NombreTabla & ".numpedcl= " & Val(Text1(0).Text)
    If EsHistorico Then SQL = SQL & " AND " & NomTablaLineas & ".fecpedcl=" & DBSet(Text1(1).Text, "F")
    ObtenerWhereCP = SQL
    
    If Err.Number <> 0 Then Err.Clear
End Function


Private Function MontaSQLCarga(enlaza As Boolean, Optional conServidas As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data2
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim SQL As String
    
    SQL = "SELECT numpedcl, numlinea, codalmac, codartic, nomartic, ampliaci,"
    'Cantidad , ""
    SQL = SQL & " palets, cajas,cantidad,"
    If conServidas Then SQL = SQL & "cajserv,servidas, "
    SQL = SQL & "precioar, preciolitro,"
    SQL = SQL & " origpre, dtoline1, dtoline2,importel "
    SQL = SQL & " FROM " & NomTablaLineas
    If enlaza Then
        SQL = SQL & " WHERE " & Replace(ObtenerWhereCP, NombreTabla, NomTablaLineas)
        If EsHistorico Then SQL = SQL & " and fecpedcl='" & Format(Text1(1).Text, FormatoFecha) & "'"
    Else
        SQL = SQL & " WHERE numpedcl = -1"
    End If
    SQL = SQL & " Order by numpedcl, numlinea"
    MontaSQLCarga = SQL
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el Modo en que estemos
Dim b As Boolean

        b = (Modo = 2) Or (Modo = 5 And ModificaLineas = 0)
        'Me.mnOpciones.Enabled = (b Or Modo = 0)
        'Insertar
        Toolbar1.Buttons(5).Enabled = (b Or Modo = 0) And Not EsHistorico
        Me.mnNuevo.Enabled = (b Or Modo = 0) And Not EsHistorico
        'Modificar
        Toolbar1.Buttons(6).Enabled = b And Not EsHistorico
        Me.mnModificar.Enabled = b And Not EsHistorico
        'eliminar
        Toolbar1.Buttons(7).Enabled = b And Not EsHistorico
        Me.mnEliminar.Enabled = b And Not EsHistorico
            
        b = (Modo = 2)
        'Mantenimiento lineas
        Toolbar1.Buttons(10).Enabled = b And Not EsHistorico
        Me.mnLineas.Enabled = b And Not EsHistorico
        'Generar Albaran desde Pedido
        Toolbar1.Buttons(11).Enabled = b And Not EsHistorico
        Me.mnGenAlbaran.Enabled = b And Not EsHistorico
        Toolbar1.Buttons(12).Enabled = b And Not EsHistorico
        Me.mnGeneraFactura.Enabled = b And Not EsHistorico
                
        
        
        
        'Imprimir orden de instalacion
        Me.Toolbar1.Buttons(15).Enabled = Not EsHistorico
        Me.mnImpOrde.Enabled = Not EsHistorico
        
        b = (Modo >= 3) Or Modo = 1
        'Buscar
        Toolbar1.Buttons(1).Enabled = Not b
        Me.mnBuscar.Enabled = Not b
        'Ver Todos
        Toolbar1.Buttons(2).Enabled = Not b
        Me.mnVerTodos.Enabled = Not b
End Sub


Private Sub CargarComboFacturacion()
'### Combo Tipo Facturación
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


Private Function InsertarPedido(vSQL As String, vTipoMov As CTiposMov) As Boolean
'Insertar la Cabecera de un Pedido, tabla: scaped
Dim MenError As String
Dim bol As Boolean, Existe As Boolean
Dim cambiaSQL As Boolean

    On Error GoTo EInsertarOferta
    
    bol = True
    
    cambiaSQL = False
    'Comprobar si mientras tanto se incremento el contador de Pedidos
    'para ello vemos si existe un Pedido con ese contador y si existe lo incrementamos
    Do
        MenError = DevuelveDesdeBDNew(conAri, NombreTabla, "numpedcl", "numpedcl", Text1(0).Text, "N")
        If MenError <> "" Then
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
    MenError = "Insertando en la tabla Cabecera de Pedidos (" & NombreTabla & ")."
    conn.Execute vSQL, , adCmdText
    
    'Actualizar los datos del cliente si es de varios
    If EsDeVarios Then
        MenError = "Actualizando el Cliente de Varios (sclvar)."
        bol = ActualizarClienteVarios(Text1(4).Text, Text1(6).Text)
    End If
    
    MenError = "Actualizando el contador del Pedido."
'    bol = vTipoMov.IncrementarContador("REG")
    vTipoMov.IncrementarContador (CodTipoMov)

EInsertarOferta:
        If Err.Number <> 0 Then
            MenError = "Insertando Pedido." & vbCrLf & "----------------------------" & vbCrLf & MenError
            MuestraError Err.Number, MenError, Err.Description
            bol = False
        End If
        If bol Then
            conn.CommitTrans
            InsertarPedido = True
        Else
            conn.RollbackTrans
            InsertarPedido = False
        End If
End Function


Private Sub LimpiarDatosCliente()
'Limpia los campos del Form con datos del cliente
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
'        Text2(8).Text = ""
        Me.cboFacturacion.ListIndex = -1
    End If
End Sub
    

Private Function PedidoConInstalaciones() As Boolean
'Comprobar si en las lineas del Pedido hay algun articulo que sea Instalacion
'Si no hay niguna linea que sea instalacion no se imprimira la Orden de Instalacion
Dim SQL As String
Dim RS As ADODB.Recordset

    On Error GoTo EInstalac

    PedidoConInstalaciones = False
    SQL = "SELECT sliped.codartic, sliped.numlinea,scaped.numpedcl, sfamia.instalac "
    SQL = SQL & " FROM ((sliped INNER JOIN scaped ON sliped.numpedcl=scaped.numpedcl) "
    SQL = SQL & " INNER JOIN sartic ON sliped.codartic=sartic.codartic) INNER JOIN "
    SQL = SQL & " sfamia ON sartic.codfamia=sfamia.codfamia "
    SQL = SQL & " WHERE scaped.numpedcl = " & Val(Text1(0).Text) & " And sfamia.instalac = 1"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        PedidoConInstalaciones = False
    Else
        PedidoConInstalaciones = True
    End If
    RS.Close
    Set RS = Nothing
    
EInstalac:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar si hay Articulos que son Instalaciones.", Err.Description
End Function


Private Function InicializarCStockAlbar(ByRef vCStock As CStock, TipoM As String, Optional numlinea As String, Optional ByRef RS As ADODB.Recordset) As Boolean
'Para comprobar stock al pasar de Pedido a Albaran de Venta
On Error Resume Next
    
    vCStock.tipoMov = TipoM
    vCStock.DetaMov = "ALV"
    vCStock.Trabajador = CLng(Text1(4).Text) 'En codigope ponemos el Cliente
    vCStock.Documento = Text1(0).Text
    vCStock.codartic = RS!codartic
    vCStock.codAlmac = CInt(RS!codAlmac)
    
    If AlbCompleto Then
        vCStock.Cantidad = CSng(RS!Cantidad)
        If RS.Fields.Count > 3 Then 'Si no se selecciona el campo importe de la tabla es que solo vamos a comprobar stock y no se necesita
            vCStock.Importe = CCur(RS!ImporteL)
        End If
    Else
        vCStock.Cantidad = CSng(RS!servidas)
        'Si se va a Insertar en alguna linea obtener el importe
        'Si solo vamos a comprobar stock no hace falta el importe
        If RS.Fields.Count > 4 Then
            vCStock.Importe = CCur(CalcularImporte(RS!servidas, RS!precioar, RS!dtoline1, RS!dtoline2, vParamAplic.TipoDtos))
        End If
    End If
    
    vCStock.LineaDocu = CInt(ComprobarCero(numlinea))
    
    If Err.Number <> 0 Then
        MsgBox "No se han podido inicializar la clase para actualizar Stock", vbExclamation
        InicializarCStockAlbar = False
    Else
        InicializarCStockAlbar = True
    End If
End Function


Private Function PasarPedidoAAlbaran(vSQL As String, NumAlb As String) As Boolean
'IN -> vSQL: cadena para el Select con los datos obtenidos en frmList
'OUT -> numAlb: Nº de Albaran de Venta que se ha insertado
Dim bol As Boolean
Dim MenError As String
Dim Devuelve As String
Dim SQL As String
Dim RS As ADODB.Recordset
Dim cCli As CCliente

    On Error GoTo EGenPedido

    bol = False
        
    'Aqui empieza transaccion
    conn.BeginTrans
    
    'Insertar en tablas de Albaranes el Pedido (scaalb, slialb)
    bol = InsertarAlbaran(vSQL, MenError, NumAlb)
    
    'Actualizar Stock en salmac, e introducir movimiento en smoval
    If bol Then
        MenError = "Error al insertar movimientos de stock."
        bol = InsertarMovStock(NumAlb)
    End If
    
    If bol Then
        If AlbCompleto Then  'Si se inserta Albaran
            'Borrar el Pedido de las tablas de Pedidos (scaped, sliped)
            MenError = "Eliminar pedido."
            bol = EliminarPedido(CLng(Text1(0).Text))
        Else
            'Actualizar la cantidad=cantidad-servidas y servidas= 0 en sliped
            bol = ActualizarPedido()
            'Marcar Resto de pedido: restoped=1
            If bol Then bol = ActualizarCabPedido
        End If
        
        If bol Then
            'Actualizar el campo fechamov (ult. movimiento) de la tabla de clientes (sclien)
            'si la fecha es posterior a la que tiene
            Set cCli = New CCliente
            If cCli.LeerDatos(Text1(4).Text) Then
                bol = cCli.ActualizaUltFecMovim(FechaAlb)
            Else
                bol = False
            End If
            Set cCli = Nothing
            
'            devuelve = DevuelveDesdeBDNew(conAri, "sclien", "fechamov", "codclien", Text1(4).Text, "N")
'            If CDate(FechaAlb) > CDate(devuelve) Then
'                MenError = "Actualizando Fecha Movimiento del Cliente."
'                devuelve = "UPDATE sclien SET fechamov=" & DBSet(FechaAlb, "F")
'                devuelve = devuelve & " WHERE codclien=" & Text1(4).Text
'                Conn.Execute devuelve, , adCmdText
'            End If
        End If
    End If
    
    If bol Then
        'Comprobar si Hay Nº SERIE en compras, si hay Mostrar los Nº Serie y seleccionar
        'sino, pedir los Nº de serie de aquellos articulos que lo requieran
        ComprobarNSeriesLineas (NumAlb)
            
        If Not AlbCompleto Then
            'Eliminar las filas del pedido que se servieron completas (sliped)
            SQL = "DELETE FROM sliped WHERE " & Replace(ObtenerWhereCP, NombreTabla, NomTablaLineas) & " AND cantidad=0"
            conn.Execute SQL
            
            'Comprobar que si no quedan lineas en el pedido se elimine la cabecera del pedido
            SQL = "select codalmac,codartic FROM sliped WHERE numpedcl=" & Text1(0).Text
            Set RS = New ADODB.Recordset
            RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If RS.EOF Then 'No hay lineas de pedido --> Eliminar la cabecera
                SQL = "DELETE FROM " & NombreTabla & " WHERE numpedcl=" & Text1(0).Text
                conn.Execute SQL
            End If
            RS.Close
            Set RS = Nothing
        End If
    End If
    
    

    
EGenPedido:
    If Err.Number <> 0 Or Not bol Then
        MenError = "Pasando Pedido a Albaran." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
        bol = False
    End If
    If bol Then
        conn.CommitTrans
    Else
        conn.RollbackTrans
    End If
    PasarPedidoAAlbaran = bol
End Function



Private Function InsertarAlbaran(vSQL As String, MenError As String, NumAlb As String) As Boolean
'Devuelve el mensaje de error si se produce
Dim bol As Boolean, Existe As Boolean
Dim Devuelve As String
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim Codtipom As String

    On Error GoTo EInsertarAlbaran
    
    bol = False
    InsertarAlbaran = bol
    
    'Obtener el Contador de PEDIDO
    
    If Not vUsu.TrabajadorB Then
        Codtipom = "ALV"
    Else
        Devuelve = DevuelveDesdeBDNew(conAri, "sliped", "count(*)", "codalmac <> " & vParamAplic.AlmacenB & " AND numpedcl  ", Text1(0).Text, "N")
        If Devuelve = "" Then Devuelve = "0"
        If Val(Devuelve) > 1 Then
            Devuelve = "Existen lineas(" & Devuelve & ") asignadas a otro almacen. No deberia continuar."
            'devuelve = devuelve & vbCrLf & vbCrLf & "¿Continuar de igual modo ?"
            MsgBox Devuelve, vbExclamation
            Exit Function
        End If
        Codtipom = "ALZ"
    End If
        
        
        
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(Codtipom) Then
        'Comprobar si mientras tanto se incremento el contador de Pedidos
        'para ello vemos si existe una oferta con ese contador y si existe la incrementamos
        Do
            NumAlb = vTipoMov.ConseguirContador(Codtipom)
            Devuelve = DevuelveDesdeBDNew(conAri, "scaalb", "numalbar", "codtipom", Codtipom, "T", , "numalbar", NumAlb, "N")
            If Devuelve <> "" Then
                'Ya existe el contador incrementarlo
                Existe = True
                vTipoMov.IncrementarContador (Codtipom)
                NumAlb = vTipoMov.ConseguirContador(Codtipom)
            Else
                Existe = False
            End If
        Loop Until Not Existe
            
    Else 'No existe el tipo de Movimiento
        Set vTipoMov = Nothing
        Exit Function
    End If
    
    'Acabar la sql con el contador seleccionado
    Devuelve = vSQL
    vSQL = "INSERT INTO scaalb (codtipom,numalbar,fechaalb,factursn,codclien,nomclien,domclien,codpobla,pobclien,proclien,nifclien,telclien,"
    vSQL = vSQL & "coddirec,nomdirec,referenc,codtraba,codtrab1,codtrab2,codagent,codforpa,codenvio,dtoppago,dtognral,tipofact,"
    vSQL = vSQL & "observa01,observa02,observa03,observa04,observa05,numofert,fecofert,numpedcl,fecpedcl,fecentre,sementre,observa6,refproduccion) "
    vSQL = vSQL & "SELECT '" & Codtipom & "' as codtipom, " & NumAlb & " as numalbar, " & Devuelve & ",observa6,refproduccion "
    vSQL = vSQL & " FROM " & NombreTabla & " WHERE numpedcl=" & Text1(0).Text

    'Insertar Cabecera
    MenError = "Error al insertar en la tabla Cabecera de Albaranes (scaalb )."
    conn.Execute vSQL, , adCmdText
    
    'Insertar Lineas Pedido
    MenError = "Error al insertar en la tabla Lineas de Albaran (slialb)."
    If Not InsertarLineasAlbaran(Codtipom, NumAlb) Then Exit Function
    
    MenError = "Error al actualizar el contador del ALbaran."
'    bol = vTipoMov.IncrementarContador("REG")
    vTipoMov.IncrementarContador (Codtipom)
    Set vTipoMov = Nothing
    
    
    
    'Si todo ha ido bien
    'Y estamos en AVAB y es pedido vinculado, entonces traeremos los numeros de lote

    If vParamAplic.EsAVAB Then
        If DBLet(Data1.Recordset!refproduccion, "N") > 0 Then
            MenError = "Error al insertar en la tabla lotes lineas de Albaran (slialblotes)."
            '--------------------------
            'Traemos los numeros de lote desde slialblotes
            Devuelve = DevuelveDesdeBD(conAri, "numalbar", "ariges" & EmprMorales & ".scaalb", "refproduccion", CStr(Data1.Recordset!numpedcl))
            
            
            Devuelve = " from ariges" & EmprMorales & ".slialblotes where codtipom='ALV' and numalbar = " & Devuelve
            Devuelve = "SELECT '" & Codtipom & "' as codtipom, " & NumAlb & ",numlinea,linea,numlote,cantidad " & Devuelve
            Devuelve = "insert into slialblotes " & Devuelve
            If Not EjecutaSQL(conAri, Devuelve, True) Then
                Devuelve = MenError & vbCrLf & Err.Description
                If MsgBox(Devuelve, vbQuestion + vbYesNo) = vbNo Then Err.Raise 513, MenError
            End If
        End If
    End If

    
    
    
    
    
    
    bol = True
    
EInsertarAlbaran:
        If Err.Number <> 0 Then bol = False
        InsertarAlbaran = bol
End Function


Private Function InsertarLineasAlbaran(TipoM As String, NumAlb As String) As Boolean
'Inserta en la tabla de lineas de albaran (slialb)
Dim SQL As String
Dim RS As ADODB.Recordset
Dim ImpLinea As String
Dim Cajas As Long
Dim Aux As Integer
Dim Pales As Integer
    On Error Resume Next

    'ENERO 2008.   codprove en slialb para traza de proveedores en lineas

    If AlbCompleto Then
        'Insertar en la tabla de Pedido, los registros seleccionados de la tabla de Ofertas
        SQL = ""
        SQL = "SELECT '" & TipoM & "', " & NumAlb & " as numalbar, numlinea, codalmac,"
        SQL = SQL & NomTablaLineas & ".codartic, " & NomTablaLineas & ".nomartic, ampliaci, "
        SQL = SQL & "cantidad, precioar, dtoline1, dtoline2, importel, origpre "
        'traza
        SQL = SQL & ",codprove,cajas,PrecioLitro,palets "
        SQL = SQL & " FROM " & NomTablaLineas & ",sartic WHERE " & NomTablaLineas & ".codartic = sartic.codartic"
        SQL = SQL & " AND numpedcl=" & Text1(0).Text
        
        ImpLinea = "INSERT INTO slialb (codtipom,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,"
        ImpLinea = ImpLinea & "cantidad,precioar,dtoline1,dtoline2,importel,origpre,codprovex,cajas,PrecioLitro,palets)"
        
        SQL = ImpLinea & SQL
        conn.Execute SQL
        ImpLinea = ""
    Else
        'TRAZA con codprove   ENERO 2008
        SQL = "select sliped.*,codprove,unicajas from sliped,sartic WHERE  sliped.codartic=sartic.codartic "
        SQL = SQL & " AND " & Replace(ObtenerWhereCP, NombreTabla, NomTablaLineas)
        SQL = SQL & " AND servidas>0"
        Set RS = New ADODB.Recordset
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not RS.EOF 'Para cada linea de pedido insertar una de albaran si servidas >0
            If RS!servidas > 0 Then
                Cajas = DBLet(RS!Unicajas, "N")
                If Cajas = 0 Then Cajas = 1
                If (RS!servidas Mod Cajas) > 0 Then
                    Aux = 1
                Else
                    Aux = 0
                End If
                Cajas = (RS!servidas \ Cajas) + Aux
                
                If DBLet(RS!palets, "N") > 0 Then
                    Pales = RS!Cajas \ RS!palets
                    'Caben por palet
                    If Pales <> 0 Then Pales = ((Cajas - 1) \ Pales) + 1
                Else
                    Pales = 0
                End If
                
                'AQUI
                'aqui aqui aqui
                
                ImpLinea = CalcularImporte(RS!servidas, RS!precioar, RS!dtoline1, RS!dtoline2, vParamAplic.TipoDtos)
                SQL = "INSERT INTO slialb (codtipom,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,"
                SQL = SQL & "cantidad,precioar,dtoline1,dtoline2,importel,origpre,codprovex,cajas,PrecioLitro,palets) "
                SQL = SQL & " VALUES('" & TipoM & "', " & NumAlb & ", " & RS!numlinea & " , "
                SQL = SQL & RS!codAlmac & ", " & DBSet(RS!codartic, "T") & ", " & DBSet(RS!NomArtic, "T") & ", " & DBSet(RS!ampliaci, "T") & ", "
                SQL = SQL & DBSet(RS!servidas, "N") & ", " & DBSet(RS!precioar, "N") & ", " & DBSet(RS!dtoline1, "N") & ", " & DBSet(RS!dtoline2, "N") & ", "
                SQL = SQL & DBSet(ImpLinea, "N") & ", " & DBSet(RS!origpre, "T") & "," & RS!codProve & "," & DBSet(Cajas, "N")
                SQL = SQL & "," & DBSet(RS!PrecioLitro, "N") & ","
                
                'Febrero 2012
                'Pongo los palets tb
                'SQL = SQL & DBSet(RS!palets, "N") & ")"
                SQL = SQL & DBSet(Pales, "N") & ")"
                conn.Execute SQL
            End If
            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
    End If
    
    If Err.Number <> 0 Then
         'Hay error , almacenamos y salimos
        InsertarLineasAlbaran = False
    Else
        InsertarLineasAlbaran = True
    End If
End Function


Private Function EliminarPedido(numPed As Long) As Boolean
'Eliminar las lineas y la Cabecera de un Pedido. Tablas: scaped, sliped
Dim SQL As String

    On Error GoTo EEliminarPed

     SQL = " WHERE  numpedcl=" & numPed

    'Lineas de Pedido
    conn.Execute "Delete from " & NomTablaLineas & SQL

    'Cabecera
    conn.Execute "Delete from " & NombreTabla & SQL

EEliminarPed:
    If Err.Number <> 0 Then
        EliminarPedido = False
    Else
        EliminarPedido = True
    End If
End Function


Private Function ActualizarPedido() As Boolean
'Actualiza la tabla de lineas de pedido (sliped)
'cantidad=cantidad-servidas y servidas=0
Dim SQL As String
Dim RS As ADODB.Recordset
Dim ImpLinea As String
Dim Pales As Integer

    On Error GoTo EActPedido
    
    SQL = "select codalmac, codartic, cantidad,servidas, precioar,dtoline1,dtoline2,numpedcl,palets,cajserv,cajas,numlinea from sliped "
    SQL = SQL & " WHERE " & Replace(ObtenerWhereCP, NombreTabla, NomTablaLineas)
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF 'Para cada linea
        ImpLinea = CalcularImporte(RS!Cantidad - RS!servidas, RS!precioar, RS!dtoline1, RS!dtoline2, vParamAplic.TipoDtos)
        SQL = "UPDATE sliped SET cantidad=cantidad-servidas, "
        Pales = 0
        If RS!codartic = vParamAplic.ArtReciclado Then
            SQL = SQL & "cajas=0"
            
        Else
            SQL = SQL & "cajas=cajas-cajserv "
            
            
            If DBLet(RS!palets, "N") > 0 Then
                If DBLet(RS!cajserv, "N") > 0 Then
                    Pales = RS!Cajas \ RS!palets
                
                    'Caben por palet
                    If Pales <> 0 Then Pales = Abs(((RS!Cajas - RS!cajserv) - 1) \ Pales) + 1
                End If
            End If
        
            
        End If
        SQL = SQL & ",servidas=0,cajserv=0 , importel=" & DBSet(ImpLinea, "N")
        
        'Si ha cambiado servidas entonces actualizo palets
        If DBLet(RS!servidas, "N") > 0 Then SQL = SQL & ",palets= " & Pales
        
        SQL = SQL & " WHERE codalmac=" & RS!codAlmac & " AND codartic=" & DBSet(RS!codartic, "T")
        'Para que no cambie los importes. Abril 2008
        SQL = SQL & " AND numpedcl= " & RS!numpedcl
        'marzo 2012
        SQL = SQL & " AND numlinea= " & RS!numlinea
        conn.Execute SQL
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing

EActPedido:
    If Err.Number <> 0 Then
        ActualizarPedido = False
    Else
        ActualizarPedido = True
    End If
End Function


Private Function ActualizarCabPedido() As Boolean
Dim SQL As String

    On Error Resume Next

    SQL = "UPDATE scaped SET restoped=1 " & " WHERE " & ObtenerWhereCP
    conn.Execute SQL
    
    If Err.Number <> 0 Then
        ActualizarCabPedido = False
    Else
        ActualizarCabPedido = True
    End If
End Function


Private Function InsertarMovStock(NumAlb As String) As Boolean
Dim vCStock As CStock
Dim b As Boolean
Dim RS As ADODB.Recordset
Dim SQL As String

    On Error Resume Next

    InsertarMovStock = False
    
    Set vCStock = New CStock
    b = True
    
    SQL = "select * from sliped WHERE " & Replace(ObtenerWhereCP, NombreTabla, NomTablaLineas)
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    vCStock.Fechamov = FechaAlb
    
    'para cada linea del Pedido Insertar en smoval y Actualizar Stock en salmac
    While (Not RS.EOF) And b
        'si hay control de stock
'        SQL = DevuelveDesdeBDNew(conAri, "sartic", "ctrstock", "codartic", RS!codartic, "T")
'        If Val(SQL) = 1 Then
            If Not InicializarCStockAlbar(vCStock, "S", CStr(RS!numlinea), RS) Then Exit Function
            
            'Para que meta el NUMERO DE ALBARAN COMO TOCA
            '
            'vCStock.Documento = NumAlb
            vCStock.Documento = Format(NumAlb, "0000000")
            If vCStock.Cantidad <> 0 Then
                'en actualizar stock comprobamos si el articulo tiene control de stock
                    b = vCStock.ActualizarStock
            End If
'        End If
        RS.MoveNext
    Wend
    Set vCStock = Nothing
    RS.Close
    Set RS = Nothing
    
    InsertarMovStock = b
    
End Function


Private Sub ImprimirAlbaran(opcion As Integer, NumAlbar As String, ImprimeValorado As Boolean)
Dim cadFormula As String
Dim Cadparam As String
Dim NumParam As Byte
Dim Cadselect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim Devuelve As String
Dim NomTabla As String
Dim Codtipom As String


    cadFormula = ""
    Cadparam = ""
    Cadselect = ""
    NumParam = 0
    NomTabla = "scaalb"
   
    '===================================================
    '============ PARAMETROS ===========================
    If (opcion = 45) Then
        If vUsu.TrabajadorB Then
            indRPT = 29   'Albaranes B
            Codtipom = "ALZ"
        Else
            indRPT = 10 'Albaran Clientes
            Codtipom = "ALV"
        End If
    End If
    If Not PonerParamRPT(indRPT, Cadparam, NumParam, nomDocu) Then
        Exit Sub
    End If

    'Añadir el codigo de usuario como parametro para link con tabla Temporal (tmptiposiva) en el Report
    'tabla temporal para el calculo del bruto total para cada tipo de IVA
    Cadparam = Cadparam & "pCodUsu=" & vUsu.Codigo & "|"
    NumParam = NumParam + 1
    
    'Nombre fichero .rpt a Imprimir
    frmImprimir.NombreRPT = nomDocu
                
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Nº de Albaran
    '---------------------------------------------------
    If NumAlbar <> "" Then
        'Cod Tipo Movimiento
        Devuelve = "{" & NomTabla & ".codtipom}='" & Codtipom & "'" 'Val(txtCodigo(0).Text)
        If Not AnyadirAFormula(cadFormula, Devuelve) Then Exit Sub
        'Nº Albaran
        Devuelve = "{" & NomTabla & ".numalbar}=" & NumAlbar
        If Not AnyadirAFormula(cadFormula, Devuelve) Then Exit Sub
        'select para insertar en tabla temporal
        Cadselect = QuitarCaracterACadena(cadFormula, "{")
        Cadselect = QuitarCaracterACadena(Cadselect, "}")
    End If
   
    '=========================================================================

    'Aqui sabemos que valor tiene CodClien y añadimos a los parametros el tipo de IVA
    'que se aplica a ese cliente
    Devuelve = DevuelveDesdeBDNew(conAri, "scaalb", "codclien", "codtipom", Codtipom, "T", , "numalbar", NumAlbar, "N")
    If Devuelve <> "" Then
        Devuelve = DevuelveDesdeBDNew(conAri, "sclien", "tipoiva", "codclien", Devuelve, "N")
        
        If vUsu.TrabajadorB Then
            Devuelve = "2"
        Else
            If Devuelve = "3" Then Devuelve = "2" 'El intracomunitario lo trato como exento
        End If
        
        If Devuelve <> "" Then
            Cadparam = Cadparam & "pTipoIVA=" & Devuelve & "|"
            NumParam = NumParam + 1
        End If
    End If
     
     
    'Si se imprimen importes y/o
    If ImprimeValorado Then
        Devuelve = "0"
    Else
        Devuelve = "2"
    End If
    ' 0 "Todo"
    ' 1 "Cantidad y Precio"
    ' 2 "Cantidad"
    Cadparam = Cadparam & "Albarcon= " & Devuelve & "|"
    NumParam = NumParam + 1

     
     
     
     With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = Cadparam
            .NumeroParametros = NumParam
            .SoloImprimir = False
            .EnvioEMail = False
            .opcion = opcion
            .Titulo = "Albaran de Cliente"
            .Show vbModal
    End With
End Sub


Private Function InicializarCStock(ByRef vCStock As CStock, TipoM As String, Optional numlinea As String) As Boolean
On Error Resume Next

    vCStock.tipoMov = TipoM
    If Modo = 6 Then 'Pasar Pedido a Albaran
        vCStock.DetaMov = "ALV"
    Else
        vCStock.DetaMov = CodTipoMov
    End If
    
    vCStock.Trabajador = CLng(Text1(4).Text) 'ponemos el cliente del pedido
    vCStock.Documento = Text1(0).Text 'Nº Pedido
    vCStock.Fechamov = Text1(1).Text
    
    If ModificaLineas = 1 Or ModificaLineas = 2 Then '1=Insertar, 2=Modificar
        vCStock.codartic = txtAux(1).Text
        vCStock.codAlmac = CInt(txtAux(0).Text)
        vCStock.Cantidad = CSng(ComprobarCero(txtAux(3).Text))
        vCStock.Importe = CCur(ComprobarCero(txtAux(8).Text))
    Else
        vCStock.codartic = Data2.Recordset!codartic
        vCStock.codAlmac = CInt(Data2.Recordset!codAlmac)
        If Modo = 6 Then 'Pasar Pedido a Albaran
            vCStock.Cantidad = CSng(ComprobarCero(txtAux(3).Text))
        Else
            vCStock.Cantidad = CSng(Data2.Recordset!Cantidad)
        End If
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


Private Function ActualizarServidas() As Boolean
'Actualiza el campo "servidas" de la tabla "sliped"
Dim SQL As String

    On Error Resume Next
    
    SQL = "0"
    If txtAux(3).Text <> "" Then
        If InStr(1, txtAux(3).Text, ",") > 0 Then
            If InStr(1, txtAux(3).Text, ".") Then
                'Importeformateado
                SQL = TransformaComasPuntos(CStr(ImporteFormateado(txtAux(3).Text)))
            Else
                SQL = TransformaComasPuntos(txtAux(3).Text)
            End If
        Else
            SQL = txtAux(3).Text
        End If
    End If

    SQL = "UPDATE sliped SET servidas= " & SQL
    
    'cajaserv
    txtAnterior = "0"
    If txtAux(9).Text <> "" Then
        If IsNumeric(txtAux(9).Text) Then txtAnterior = Val(txtAux(9).Text)
    End If
    SQL = SQL & ", cajserv = " & txtAnterior
    txtAnterior = ""
    
    
    SQL = SQL & " WHERE " & Replace(ObtenerWhereCP, NombreTabla, NomTablaLineas) & " AND numlinea=" & Data2.Recordset!numlinea
    conn.Execute SQL
    
    If Err.Number <> 0 Then
        ActualizarServidas = False
    Else
        ActualizarServidas = True
    End If
End Function


Private Sub PonerServidas()
Dim NumFila As Integer
Dim cadMen As String
Dim vStock As String
Dim Servir As Boolean
'    NumFila = DataGrid1.Row

    If txtAux(3).Text = "" Or txtAux(9).Text = "" Then Exit Sub

    NumFila = Data2.Recordset.AbsolutePosition
    ActualizarServidas
    CargaGrid2 DataGrid1, Data2, True
    SituarDataPosicion Data2, CLng(NumFila), ""
    
'    DataGrid1.Row = NumFila
    Servir = SePuedeServir(vStock)
    If Not Servir Then
        cadMen = "No hay suficiente Stock para servir la cantidad solicitada."
        cadMen = cadMen & vbCrLf & "(Stock= " & vStock & ")" & vbCrLf
        cadMen = cadMen & vbCrLf & "¿Continuar?"
        If MsgBox(cadMen, vbInformation + vbYesNo) = vbYes Then Servir = True
    End If
    If Servir Then
        If CSng(txtAux(3).Text) > Data2.Recordset!Cantidad Then
            cadMen = "La cantitad a servir debe ser menor o igual a al cantidad del pedido."
            cadMen = cadMen & vbCrLf
            MsgBox cadMen, vbInformation
            txtAux(9).Text = ""
            txtAux(3).Text = ""
            PonerFoco txtAux(9)
        Else
'            TxtAux_KeyDown 3, 40, 0
            MoverSigRegistro
        End If
    Else
        txtAux(9).Text = ""
        txtAux(3).Text = ""
        PonerFoco txtAux(9)
    End If
    
End Sub


Private Function SePuedeServir(vStock As String) As Boolean
'Si se puede servir una determinada linea del pedido cuando se esta introduciendo
'la cantidad a servir para cada codalmac,codartic
'OUT -> vStock: stock existente
Dim SQL As String
Dim RS As ADODB.Recordset
Dim Dif As Integer
Dim vCStock As CStock

    On Error GoTo EServir

    Set vCStock = New CStock
    vCStock.codartic = Data2.Recordset!codartic
    If Not vCStock.MueveStock Then
        SePuedeServir = True
        Set vCStock = Nothing
        Exit Function
    End If
    Set vCStock = Nothing

    
    SePuedeServir = False
    SQL = " SELECT sliped.codalmac, sliped.codartic, canstock , sum(servidas), canstock - SUM(servidas) as Dif "
    SQL = SQL & " FROM sliped INNER JOIN salmac ON sliped.codalmac=salmac.codalmac AND sliped.codartic=salmac.codartic "
    SQL = SQL & " WHERE " & Replace(ObtenerWhereCP, NombreTabla, NomTablaLineas) & " AND sliped.codAlmac = " & Data2.Recordset!codAlmac & " AND sliped.codartic=" & DBSet(Data2.Recordset!codartic, "T")
    SQL = SQL & " GROUP by sliped.codalmac, sliped.codartic "

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        Dif = RS!Dif
        SePuedeServir = (RS!Dif >= 0)
        vStock = CStr(RS!CanStock)
    End If
    RS.Close
    Set RS = Nothing

EServir:
    If Err.Number <> 0 Then SePuedeServir = False
End Function


Private Sub GenerarAlbaran(PasarTambienAFacturar As Boolean)
 Dim numPed As Long 'Nº Pedido
Dim NumAlb As String 'Nº Albaran
Dim SQL As String
Dim ImprimeFactura As Boolean
Dim AlbaranValorado As Boolean

    'Pedir: Operador de Albaran, Material Preparado por y forma de envio
    SQL = DevuelveDesdeBD(conAri, "albarcon", "sclien", "codclien", Text1(4).Text, "N")
    AlbaranValorado = SQL = "1"
    
    Precio = "" 'AQUI GUARDARE la variabale de albaran valorado
    CadenaSQL = ""
    Set frmList = New frmListadoPed
    If PasarTambienAFacturar Then
        frmList.OpcionListado = 1000
    Else
        frmList.OpcionListado = 43
        frmList.chkAlbValorado.Value = Abs(AlbaranValorado)
    End If
    frmList.NumCod = CodTipoMov
    frmList.Show vbModal
    Set frmList = Nothing
    If CadenaSQL = "" Then Exit Sub
    AlbaranValorado = Precio = "1"
    Precio = ""
  
    NumRegElim = Data1.Recordset.AbsolutePosition
    numPed = Data1.Recordset!numpedcl
    'Si es factura el albaran NO se imprime, y se imprimira si lo ha marcado, la factura
    
    If PasarTambienAFacturar Then
        ImprimeFactura = ImprimeAlb
        ImprimeAlb = False 'El albaran NO se imprime generanod la factura
    End If
    
    'CadenaSQL, se obtiene desde frmList
    lblIndicador.Caption = "Gen. albaran"
    lblIndicador.Refresh
    
    If PasarPedidoAAlbaran(CadenaSQL, NumAlb) Then
'        'Comprobar si Hay Nº SERIE en compras, si hay Mostrar los Nº Serie y seleccionar
'        'sino, pedir los Nº de serie de aquellos articulos que lo requieran
'        ComprobarNSeriesLineas (NumAlb)
        Espera 0.4
        If Not PasarTambienAFacturar Then
            MsgBox "El Pedido de Venta Nº: " & Format(numPed, "0000000") & vbCrLf & vbCrLf & "ha generado el Albaran Nº: " & Format(NumAlb, "0000000"), vbInformation
        Else
            'Ahora genero la factura a partir del ALBARAN
            lblIndicador.Caption = "Gen FACTURA"
            DoEvents
            
            'Genero la factura del albaran que se ha generado
            'Montare un cadselect
            'Tipo movimiento = "ALV"
            'Numero albaran  = NumAlb
            'Fecha factura=fecha albaran = FechaAlb
            
            
            CadenaSQL = "scaalb.codtipom ='AL"
            If vUsu.TrabajadorB Then
                CadenaSQL = CadenaSQL & "Z"
            Else
                CadenaSQL = CadenaSQL & "V"
            End If
            CadenaSQL = CadenaSQL & "' AND scaalb.numalbar = " & NumAlb
            Precio = "SELECT scaalb.*,sclien.nomclien FROM scaalb INNER JOIN sclien ON scaalb.codclien=sclien.codclien "
            Precio = Precio & " WHERE " & CadenaSQL
            If vUsu.TrabajadorB Then
                TraspasoAlbaranesFacturas Precio, CadenaSQL, FechaAlb, CtaBancoPropi, Nothing, lblIndicador, ImprimeFactura, "ALZ", ""
            Else
                TraspasoAlbaranesFacturas Precio, CadenaSQL, FechaAlb, CtaBancoPropi, Nothing, lblIndicador, ImprimeFactura, "ALV", ""
            End If
        End If
            
        
        
        
        PonerModo 2
        If AlbCompleto Then
            'Se habra eliminado el pedido de (scaped, sliped)
            PosicionarDataTrasEliminar
        Else
            SQL = DevuelveDesdeBDNew(conAri, "scaped", "numpedcl", "numpedcl", Text1(0).Text, "N")
            If SQL = "" Then 'Ya no existe le pedido lo hemos eliminado
                PosicionarDataTrasEliminar
            Else
                PosicionarData
                CargaGrid DataGrid1, Data2, True, False
            End If
        End If
        Screen.MousePointer = vbDefault
        CargaTxtAuxServidas False, False
    
        'Imprimer albaran si se solicitó
        If ImprimeAlb Then ImprimirAlbaran 45, NumAlb, AlbaranValorado
        
'    Else 'Si no se ha pasado el Pedido a Albaran
        
    End If
End Sub


Private Function SePuedeServirPedido() As Boolean
'Si se puede servir el Pedido solicitado (parcial o completo) y pasar a albaran
Dim vCStock As CStock
Dim SQL As String
Dim b As Boolean
Dim RS As ADODB.Recordset

    On Error Resume Next

    'Verificar si hay stock para aquellas familias que no son instalacion
    Set vCStock = New CStock
    b = True
    
    If AlbCompleto Then
        SQL = "SELECT codalmac, codartic, SUM(cantidad) as cantidad from sliped "
    Else
        SQL = "SELECT codalmac, codartic, SUM(servidas) as servidas from sliped "
    End If
    SQL = SQL & " WHERE " & Replace(ObtenerWhereCP, NombreTabla, NomTablaLineas)
    SQL = SQL & " GROUP by codalmac, codartic"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'para cada linea del Pedido comprobar el stock si no es instalacion
    While (Not RS.EOF) And b
        If Not InicializarCStockAlbar(vCStock, "S", , RS) Then
            b = False
            Screen.MousePointer = vbDefault
            Set vCStock = Nothing
            RS.Close
            Set RS = Nothing
            Exit Function
        End If
        
        'Comprobar si se puede mover stock (hay stock, o si no hay pero no control de stock)
        If AlbCompleto Then
            If vCStock.MueveStock Then b = vCStock.MoverStock(False, True)
        Else
            If vCStock.MueveStock Then b = vCStock.MoverStock(False)
        End If
        RS.MoveNext
    Wend
    
    Set vCStock = Nothing
    RS.Close
    Set RS = Nothing
    SePuedeServirPedido = b
    
    If Err.Number <> 0 Then SePuedeServirPedido = False
End Function


Private Sub InicializarServidas()
'Pone el campo servidas a 0 en la tabla lineas de pedido (sliped)
Dim SQL As String

    SQL = "UPDATE " & NomTablaLineas & " SET servidas= 0 "
    SQL = SQL & ", cajserv=0"
    SQL = SQL & " WHERE " & Replace(ObtenerWhereCP, NombreTabla, NomTablaLineas)
    conn.Execute SQL
End Sub


Private Sub ComprobarNSeriesLineas(NumAlb As String)
'Al pasar de PEDIDO a ALBARAN
'control de Nº Series si hay algun articulo en las lineas de pedido que requiere Nº de serie
'Si NO se realiza control Nº series en compras pedirlos ahora
'Si se realiza control Nº Series en compras verificar que efectivamente estan introducidos
'y mostrarlos para seleccionarlos
Dim SQL As String
Dim RSLineas As ADODB.Recordset
Dim cadWhere As String
        
    On Error GoTo ECompNSerie
    
    cadWhere = " WHERE codtipom='ALV' and "
    cadWhere = cadWhere & " numalbar=" & NumAlb
    
    'Seleccionamos aquellas lineas de albaran que tienen Nº de Serie
    SQL = "SELECT slialb.codartic, sum(cantidad) as cantidad, slialb.numlinea "
    SQL = SQL & " FROM slialb INNER JOIN sartic on slialb.codartic=sartic.codartic "
    SQL = SQL & cadWhere & " And nseriesn = 1 "
    SQL = SQL & " GROUP BY codartic ORDER BY Codartic "

    Set RSLineas = New ADODB.Recordset
    RSLineas.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RSLineas.EOF Then
        'Comprobar si NO Hay Nº SERIE en Compras y si no se realizo alli
        'Mostrar ahora ventana para pedir los Nº Serie de la cantidad introducida
        Me.cmdAux(1).Tag = NumAlb
        If Not vParamAplic.NumSeries Then
            PedirNSeries RSLineas
        Else 'Se realizo contro en COMPRAS, Mostramos los Nº y seleccionamos
            MostrarNSeries RSLineas
        End If
    End If
    RSLineas.Close
    Set RSLineas = Nothing
    
ECompNSerie:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobanco Nº Serie.", Err.Description
End Sub


Private Sub PedirNSeries(ByRef RS As ADODB.Recordset)
    On Error GoTo EPedirNSeries
    
    'Visualizar en pantalla el Grid, y rellenar los Nº Serie
    PedirNSeriesGnral RS, True

    Set frmNSerie = New frmRepCargarNSerie
    frmNSerie.DeVentas = True 'Se llama desde Alb. de Venta
    frmNSerie.Show vbModal
    Set frmNSerie = Nothing
        
EPedirNSeries:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub MostrarNSeries(ByRef RSLineas As ADODB.Recordset)
'Si los Nº de serie se introdujeron en ALBARAN COMPRAS se muestran
'los Nº de serie de los articulos comprados y se seleccionan tantos como cantidad de la linea
Dim SQL As String
Dim Campos As String
   
    SQL = MostrarNSeriesGnral(RSLineas, Campos)
    
    Set frmMen = New frmMensajes
    frmMen.cadWhere = SQL
    frmMen.cadWHERE2 = ""
    frmMen.OpcionMensaje = 4 'Nº Series Articulo
    frmMen.vCampos = Campos
    frmMen.Show vbModal
    Set frmMen = Nothing
End Sub


Private Sub InsertarCabecera()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim SQL As String

    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
        Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
        SQL = CadenaInsertarDesdeForm(Me)
        If SQL <> "" Then
            If InsertarPedido(SQL, vTipoMov) Then
'                            PosicionarData
                CadenaConsulta = "Select * from " & NombreTabla & " WHERE " & ObtenerWhereCP & Ordenacion
                PonerCadenaBusqueda
                PonerModo 2
                'Ponerse en Modo Insertar Lineas
                BotonMtoLineas 1, "Pedidos"
                BotonAnyadirLinea
            End If
        End If
        Text1(0).Text = Format(Text1(0).Text, "0000000")
    End If
    Set vTipoMov = Nothing
    Me.SSTab1.Tab = 0
End Sub


Private Function InsertarNSerie(numSerie As String, codartic As String, numlinea As String) As Boolean
'Inserta o Actualiza en la tabla sserie, si al pasar Pedido -> Albaran
'existen lineas con control de Nº Serie
'Dim CadValues As String, cadValuesU As String, CadValuesI As String
Dim Devuelve As String
Dim TieneMan As String * 1
Dim NumAlbar As String
Dim nSerie As CNumSerie
Dim b As Boolean

    On Error GoTo EInsertarNSerie
    
    'Comprobar que el cliente tiene mantenimientos en esa direc/dpto
    TieneMan = "0"
    Devuelve = DevuelveDesdeBDNew(conAri, "scaman", "nummante", "codclien", Text1(4).Text, "N", , "coddirec", Text1(12).Text, "N")
    'El cliente tiene Mantenimientos
    If Devuelve <> "" Then TieneMan = "1"
    
    Set nSerie = New CNumSerie
    nSerie.Cliente = CInt(Text1(4).Text)
    nSerie.DirDpto = Text1(12).Text
    nSerie.conMante = TieneMan
    nSerie.tipoMov = CodTipoMov
    
    Devuelve = DevuelveDesdeBDNew(conAri, "scaalb", "fechaalb", "codtipom", "ALV", "T", , "numalbar", Me.cmdAux(1).Tag, "N")
    If Devuelve <> "" Then nSerie.FechaVta = Devuelve
    
    nSerie.NumAlbaran = Me.cmdAux(1).Tag
    nSerie.NumLinAlb = numlinea


    'obtenemos los dias de garantia del articulo
    nSerie.ObtenFechaFinGarantia codartic, Text1(1).Text
    
     'Comprobar si existe en la tabla sserie
     NumAlbar = "numalbar" 'Nº albaran de Venta
     Devuelve = DevuelveDesdeBDNew(conAri, "sserie", "numserie", "numserie", numSerie, "T", NumAlbar, "codartic", codartic, "T")
     If Devuelve <> "" Then 'EXISTE en tabla sserie
        If NumAlbar = "" Then b = nSerie.ActualizarNumSerie(True)
     Else
        b = nSerie.InsertarNumSerie
    End If
    InsertarNSerie = True
    Set nSerie = Nothing
         
EInsertarNSerie:
    If Err.Number <> 0 Then b = False
    InsertarNSerie = b
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
                    If Text1(5).Text = Data1.Recordset!nomclien Then
                        Set vCliente = Nothing
                        Exit Sub
                    End If
                End If
            End If
            
'            If (Not EsDeVarios) Or (EsDeVarios And modo = 3) Then
            Text1(4).Text = vCliente.Codigo
            FormateaCampo Text1(4)
            If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar
                Text1(5).Text = vCliente.Nombre  'Nom clien
                Text1(8).Text = vCliente.Domicilio
                Text1(9).Text = vCliente.CPostal
                Text1(10).Text = vCliente.Poblacion
                Text1(11).Text = vCliente.Provincia
                Text1(6).Text = vCliente.NIF
                Text1(7).Text = vCliente.TfnoClien
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
            If Observaciones <> "" Then MsgBox Observaciones, vbInformation, "Observaciones del cliente"
                           
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



Private Sub CalcularDatosFactura()
Dim I As Integer
Dim cadWhere As String, SQL As String
Dim vFactu As CFactura

    'Limpiar en el form los datos calculados de la factura
    'y volvemos a recalcular
    For I = 33 To 56
         Text3(I).Text = ""
    Next I
    
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
        Text3(56).Text = vFactu.BaseImp
        Text3(55).Text = vFactu.TotalFac
        
        
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



Private Sub PrecioUdLitro(DePrecioAPrecioLitros As Boolean)
Dim ListrosUd As Currency
Dim V As Currency
    
    ListrosUd = 1
    If Not (ElArticulo Is Nothing) Then
        'If ElArticulo.LitrosxUd > 1 Then ListrosUd = ElArticulo.LitrosxUd
        If ElArticulo.LitrosxUd > 0 Then ListrosUd = ElArticulo.LitrosxUd
        If ElArticulo.Codigo = vParamAplic.ArtReciclado Then Exit Sub
    End If
        
    
    If DePrecioAPrecioLitros Then
        If txtAux(4).Text = "" Then
            txtAux(10).Text = ""
        Else
            V = ImporteFormateado(txtAux(4).Text)
            V = Round(V / ListrosUd, 4)
            txtAux(10).Text = Format(V, FormatoPrecio)
        End If
    Else
        'Ha metido precio UD
        If txtAux(10).Text = "" Then
            txtAux(4).Text = ""
        Else
            V = ImporteFormateado(txtAux(10).Text)
            V = V * ListrosUd
            txtAux(4).Text = Format(V, FormatoPrecio)
        End If
    End If
End Sub



Private Sub CantidadCajas(DeCantidadACajas As Boolean)
Dim CajaUd As Integer
Dim V As Long
    CajaUd = 1
    If Not (ElArticulo Is Nothing) Then
        If ElArticulo.UnidCaja > 1 Then CajaUd = ElArticulo.UnidCaja
    End If
        

    If DeCantidadACajas Then
        If txtAux(3).Text = "" Then
            txtAux(9).Text = ""
        Else
            V = Val(txtAux(3).Text)
            txtAux(9).Text = V \ CajaUd
        End If
    Else
        'Ha metido cajas. Nos vamos a cantidad
        If txtAux(9).Text = "" Then
            txtAux(3).Text = ""
        Else
            V = Val(txtAux(9).Text)
            txtAux(3).Text = Format(V * CajaUd, FormatoCantidad)
        End If
    End If
End Sub




Private Sub CantidadCajasServidas(DeCantidadACajas As Boolean)
Dim CajaUd As Integer
Dim V As Long
    CajaUd = 1
    If Not (ElArticulo Is Nothing) Then
        If ElArticulo.UnidCaja > 1 Then CajaUd = ElArticulo.UnidCaja
       
    End If
    
    
    If Not (Me.Data2.Recordset Is Nothing) Then
        If Not Data2.Recordset.EOF Then
            If Data2.Recordset!codartic = vParamAplic.ArtReciclado Then Exit Sub
        End If
    End If

    If DeCantidadACajas Then
        If txtAux(3).Text = "" Then
            txtAux(9).Text = ""
        Else
            V = Val(txtAux(3).Text)
            txtAux(9).Text = V \ CajaUd
        End If
    Else
        'Ha metido cajas. Nos vamos a cantidad
        If txtAux(9).Text = "" Then
            txtAux(3).Text = ""
        Else
            V = Val(txtAux(9).Text)
            txtAux(3).Text = Format(V * CajaUd, FormatoCantidad)
        End If
    End If
End Sub




Private Function DevuelveListaPedidos() As String
Dim R As ADODB.Recordset
Dim Cad As String
Dim CadenaFinal As String
Dim Anyade As Boolean
    'Pedidos vacios
    Set R = New ADODB.Recordset
    CadenaFinal = ""
    'Busco los pedidos sin lineas..
    Cad = "select numpedcl,presupuesto from scaped,straba where "
    Cad = Cad & " scaped.codtraba=straba.codtraba  and not numpedcl in"
    Cad = Cad & " (select distinct(numpedcl) from sliped)"
    R.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    While Not R.EOF
        Anyade = False
        If vUsu.TrabajadorB Then
            If Me.mnTodosLosAlmacenes.Checked Then
                'Tiene marcado TODOS
                Anyade = True
            Else
                If Val(R!presupuesto) = 1 Then Anyade = True
            End If
        Else
            If Val(R!presupuesto) = 0 Then Anyade = True
        End If
        If Anyade Then Cad = Cad & " ," & R!numpedcl
        R.MoveNext
    Wend
    R.Close
    CadenaFinal = Cad
    
    'Los que tienen lineas
    Anyade = False 'no añadimos where
    Cad = ""
    If vUsu.TrabajadorB Then
        If Me.mnTodosLosAlmacenes.Checked Then
            'TOOOODOS
            'NO añado nada
            
            
            
        Else
            Cad = " = "
            Anyade = True
        End If
    Else
        Cad = " <> "
        Anyade = True
    End If
            
    If Anyade Then Cad = " WHERE codalmac " & Cad & vParamAplic.AlmacenB
        
    
        
   
    Cad = "select numpedcl from scaped where numpedcl in (select distinct(numpedcl) from sliped " & Cad & ")"
    R.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    While Not R.EOF
        Cad = Cad & " ," & R!numpedcl
        R.MoveNext
    Wend
    R.Close
    If Anyade Then
        If Cad = "" Then Cad = "  -1"
    End If
    CadenaFinal = CadenaFinal & Cad
    
    
    
    If Len(CadenaFinal) > 0 Then
        CadenaFinal = Mid(CadenaFinal, 3)
        CadenaFinal = " numpedcl in (" & CadenaFinal & ")"
    End If
    
    
    
    DevuelveListaPedidos = CadenaFinal
    
    
End Function



Private Sub VincularPedidoEnEmpresaProduccion()
Dim Cad As String
    
    
    'Por si acaso
    If Modo <> 2 Then Exit Sub
    If Me.Data1.Recordset.EOF Then Exit Sub
    If Not vParamAplic.EsAVAB Then Exit Sub
    
    'Si no tienen lineas
    If Me.Data2.Recordset.EOF Then Exit Sub
    If Me.Data2.Recordset.RecordCount = 0 Then Exit Sub
    
    If vUsu.Nivel > 1 Then
        MsgBox "No tiene permisos para realizar al accion", vbExclamation
        Exit Sub
    End If
    
    'Lo mas importante. NO tienen que estar YA vinculado
    If DBLet(Me.Data1.Recordset!refproduccion, "N") > 0 Then
        MsgBox "Ya esta vinculado con la empresa de produccion: " & Data1.Recordset!refproduccion, vbExclamation
        Exit Sub
    End If
    

    'Minima comprobacion
    Set miRsAux = New ADODB.Recordset
    Cad = "select * from sliped where numpedcl=" & Me.Data1.Recordset!numpedcl & " and not codartic in (select codartic from ariges" & EmprMorales & ".sartic)"
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    While Not miRsAux.EOF
        Cad = Cad & vbCrLf & miRsAux!codartic & "    " & miRsAux!NomArtic
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    If Cad <> "" Then
        Cad = "Error en articulos.  Referencias NO existen en produccion: " & vbCrLf & Cad
        MsgBox Cad, vbExclamation
        Exit Sub
    End If

    If MsgBox("Seguro que desea crear el pedido en la empresa de producción?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    
    Screen.MousePointer = vbHourglass
    conn.BeginTrans
    If RealizarElPedidoEnMorales Then
        conn.CommitTrans
        PosicionarData
    Else
        'Ya he ehco le rollback en la rutina de error
        'Conn.RollbackTrans
    End If
    Screen.MousePointer = vbDefault
    
End Sub


Private Function RealizarElPedidoEnMorales() As Boolean
Dim SQL As String
Dim Aux As String
Dim Col As Collection
Dim ContadorMor As Long
Dim I As Integer

    On Error GoTo ERealizarElPedidoEnMorales
    RealizarElPedidoEnMorales = False
    Set miRsAux = New ADODB.Recordset
    
    
    'Para las lineas
    SQL = "Select * from sliped WHERE numpedcl=" & Data1.Recordset!numpedcl & " ORDER BY numlinea"
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Set Col = New Collection
    While Not miRsAux.EOF
        'Generaremos el SQL a fatla el numero de pedido en morales
        SQL = DevueleveSQLLineasPedido()
        Col.Add SQL

        miRsAux.MoveNext
    Wend
    miRsAux.Close
    'Vamos con la cabecera
          
        
    SQL = DevuelveDesdeBD(conAri, "contador", "ariges" & EmprMorales & ".stipom", "codtipom", "PEV", "T")
    If SQL = "" Then SQL = "0"
    ContadorMor = Val(SQL) + 1
    SQL = "UPDATE ariges" & EmprMorales & ".stipom set contador = " & ContadorMor & " WHERE codtipom='PEV'"
    conn.Execute SQL
    
    'Llegado aqui va OK
    'numpedcl,fecpedcl,fecentre,sementre,visadore,codclien,nomclien,domclien,codpobla,pobclien,proclien,nifclien,telclien,coddirec,nomdirec,referenc,codtraba,codagent,codforpa,dtoppago,dtognral,tipofact,observa01,observa02,observa03,observa04,observa05,servcomp,restoped,numofert,fecofert,observap1,observap2,recogecl,observa6,refproduccion
    SQL = "INSERT INTO ariges" & EmprMorales & ".scaped(numpedcl,fecpedcl,fecentre,sementre,visadore,codclien,nomclien,domclien,codpobla,pobclien,proclien,nifclien,telclien,coddirec,nomdirec,referenc,codtraba,codagent,codforpa,dtoppago,dtognral,tipofact,observa01,observa02,observa03,observa04,observa05,servcomp,restoped,numofert,fecofert,observap1,observap2,recogecl,observa6,refproduccion) "
    SQL = SQL & " SELECT " & ContadorMor & ",fecpedcl,fecentre,sementre,visadore,"
    'codclien,nomclien,domclien,codpobla,pobclien,proclien,nifclien,telclien,
    'Leeemos los datos del cliente 1
    Aux = "select codclien,nomclien,domclien,codpobla,pobclien,proclien,nifclien,telclie1 from ariges" & EmprMorales & ".sclien where codclien = 1"
    miRsAux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'No puede haber error
    SQL = SQL & DBSet(miRsAux!CodClien, "N") & "," & DBSet(miRsAux!nomclien, "T") & "," & DBSet(miRsAux!domclien, "T") & ","
    SQL = SQL & DBSet(miRsAux!codpobla, "T") & "," & DBSet(miRsAux!pobclien, "T") & "," & DBSet(miRsAux!proclien, "T") & ","
    SQL = SQL & DBSet(miRsAux!nifClien, "T") & "," & DBSet(miRsAux!telclie1, "T") & ","
    miRsAux.Close
    
    'Resto datos pedidod                          codagent codforpa
    SQL = SQL & "coddirec,nomdirec,referenc,codtraba,1,2,dtoppago,dtognral,tipofact,"
    'observa01,observa02,observa03,observa04,observa05,"
    SQL = SQL & "null,null,null,null,null,"
    '                                               observap1,observap2,recogecl,observa6
    SQL = SQL & "servcomp,restoped,numofert,fecofert,null,null,recogecl,null," & Data1.Recordset!numpedcl  'en el de mora grabamos el numped de aqui
    SQL = SQL & " FROM scaped WHERE numpedcl = " & Data1.Recordset!numpedcl
    
    If EjecutaSQL(conAri, SQL, True) Then
        'Vamos con las lineas
        SQL = ""
        For I = 1 To Col.Count
            Aux = ContadorMor & Col.Item(I)
            SQL = SQL & ", (" & Aux & ""
        Next
        SQL = Mid(SQL, 2)
        

        'numpedcl,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,servidas,precioar,dtoline1,dtoline2,importel,origpre,cajas,PrecioLitro,cajserv,palets
        Aux = "INSERT INTO ariges" & EmprMorales & ".sliped(numpedcl,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,servidas,precioar,dtoline1,dtoline2,origpre,importel,cajas,PrecioLitro,palets) VALUES "
        SQL = Aux & SQL
        If Not EjecutaSQL(conAri, SQL, False) Then
            SQL = "Error insertando lineas"
            Err.Raise 513, SQL
        End If
        
        
        
        'Updateamos el pedido en AVAB con la referencia de morales
        SQL = "UPDATE scaped set refproduccion = " & ContadorMor & " WHERE numpedcl = " & Data1.Recordset!numpedcl
        conn.Execute SQL
    End If
ERealizarElPedidoEnMorales:
    If Err.Number <> 0 Then
        'Hago YA el rollbacktrans, para que desbloquee pedidos
        SQL = Err.Description
        ContadorMor = Err.Number
        conn.RollbackTrans
        MuestraError ContadorMor, "Realizar ped. produccion" & vbCrLf & SQL
        
    Else
        RealizarElPedidoEnMorales = True
    End If
    
    Set Col = Nothing
    Set miRsAux = Nothing
End Function


'Devuelve un string que sera un SQL
'Implictamente el mirsaux esta gargado con la linea que vamos a insertar, si no estan en los txtaux
Private Function DevueleveSQLLineasPedido() As String
Dim CPmor As CPrecioEnMorales
Dim PorCaja  As Boolean
Dim NumCajas  As Long
Dim RestoUnid As Long
Dim Precio As String
Dim OrigP As String
Dim FormatoArt As String
Dim Aux2 As Currency
Dim SQL As String
Dim Devuelve As String

    FormatoArt = "litrosunidad"

    Devuelve = DevuelveDesdeBDNew(conAri, "ariges" & EmprMorales & ".sartic", "unicajas", "codartic", miRsAux!codartic, "T", FormatoArt)
    If Devuelve = "" Then
        SQL = "Error leyendo datos artículo en empresa producción"
        Err.Raise 513, SQL
    End If
    If FormatoArt = "" Or FormatoArt = "0" Then FormatoArt = "1"
    Set CPmor = New CPrecioEnMorales


    'NUEVO
    SQL = "," & miRsAux!numlinea & "," & miRsAux!codAlmac & "," & DBSet(miRsAux!codartic, "T") & ", " & DBSet(miRsAux!NomArtic, "T") & ", " & DBSet(miRsAux!ampliaci, "T") & ", "
    SQL = SQL & DBSet(miRsAux!Cantidad, "N") & ", 0,"
    'El precio lo calculo ahora
        
    
        
        NumCajas = CPmor.ObtenerNumCajas(miRsAux!Cantidad, Devuelve)
        RestoUnid = CLng(miRsAux!Cantidad) - NumCajas * CLng(Devuelve)
        'Obtenemos la Tarifa del Cliente
        CPmor.CodigoListaMor = Val(DevuelveDesdeBDNew(conAri, "ariges" & EmprMorales & ".sclien", "codtarif", "codclien", "1", "N"))   'EL UNO ES AVAB
        CPmor.CodigoArticMor = miRsAux!codartic
        CPmor.CodigoClienMor = 1
        PorCaja = (NumCajas > 0)
        Precio = CPmor.ObtenerPrecioMor(PorCaja, Data1.Recordset!fecpedcl, OrigP)
        'Si PorCaja vuelve de ObtenerPrecio a false se calcula con precio unidad aunque NumCajas>0
        'Ya que a regresado con pvp del Articulo
        If PorCaja And NumCajas > 0 And RestoUnid > 0 Then
            SQL = "El Artículo puede venderse por Cajas (" & Devuelve & "uds. por Caja)." & vbCrLf
            SQL = SQL & vbCrLf & "Inserte dos Lineas:   "
            SQL = SQL & vbCrLf & "   Linea 1:  " & NumCajas * CInt(Devuelve) & " uds a Precio Caja"
            SQL = SQL & vbCrLf & "   Linea 2:  " & CInt(miRsAux!Cantidad) - NumCajas * CInt(Devuelve) & " uds a Precio Unidad"
            Err.Raise 513, SQL
            
        Else
            'Veo cuantas cajas son
            If Devuelve = "0" Then Devuelve = "1"
            NumCajas = miRsAux!Cajas
        
            'precioar, dtoline1, dtoline2, origpre,importel,cajas,PrecioLitro,palets
            SQL = SQL & DBSet(Precio, "N", "N") & "," & DBSet(CPmor.Descuento1, "N") & "," & DBSet(CPmor.Descuento2, "N") & ",'" & OrigP & "',"
            'Importe
            OrigP = CalcularImporte(CStr(miRsAux!Cantidad), Precio, CPmor.Descuento1, CPmor.Descuento2, vParamAplic.TipoDtos)
            SQL = SQL & DBSet(OrigP, "N", "N") & "," & DBSet(NumCajas, "N") & ","
            'Precio por litro
           
            If CCur(FormatoArt) <> 1 Then
                'Formato NO es 1 Litro
                
                Aux2 = CCur(Precio) / CCur(FormatoArt)
            Else
                Aux2 = CCur(Precio)
            End If
            SQL = SQL & DBSet(Aux2, "N", "N") & "," & DBSet(miRsAux!palets, "N") & ")"
        End If
    '                        ConseguirFoco txtAux(Index + 1), Modo
        Set CPmor = Nothing
    
    DevueleveSQLLineasPedido = SQL
    
End Function


Private Sub UpdatearLineaPedidoEnOtraEmpresa()
Dim Aux As String
Dim SQL As String
Dim PrecioArticulo As Currency

    'Comprobacion UNO. El codartic aqui y alli es el mismo
    Aux = "numpedcl = " & Data1.Recordset!refproduccion & " AND numlinea"
    NumRegElim = EmprMorales
    If Not vParamAplic.EsAVAB Then NumRegElim = EmprAVAB
    SQL = "precioar"
    Aux = DevuelveDesdeBD(conAri, "codartic", "ariges" & NumRegElim & ".sliped", Aux, Data2.Recordset!numlinea, "N", SQL)
    If SQL = "" Or SQL = "precioar" Then SQL = "0"
    PrecioArticulo = CCur(SQL)
    SQL = ""
    If Aux = "" Then
        SQL = "No se encuentra la linea en la empresa vinculada"
    Else
        If Aux <> txtAux(1).Text Then
            SQL = "Linea en empresa vinculada distinta de la actual"
        Else
            If Aux <> Data2.Recordset!codartic Then SQL = "Codigo articulo distinto en empresa vinculada"
        End If
    End If
    
    If SQL <> "" Then
        MsgBox SQL, vbExclamation
        Exit Sub
    End If
    
    'OK vamos p'alla
                    'precio         *   cantidad
    PrecioArticulo = PrecioArticulo * ImporteFormateado(txtAux(3).Text)
    PrecioArticulo = Round2(PrecioArticulo, 2)
    SQL = "UPDATE ariges" & NumRegElim & "." & NomTablaLineas & " SET cantidad = " & DBSet(txtAux(3).Text, "N") & ", "
    SQL = SQL & "importel= " & DBSet(PrecioArticulo, "N") & ", "
    'cajas,PrecioLitro
    SQL = SQL & "cajas=" & DBSet(txtAux(9).Text, "N") & ", "
    SQL = SQL & "Palets=" & DBSet(txtAux(11).Text, "N", "S")
    Aux = "numpedcl = " & Data1.Recordset!refproduccion & " AND numlinea=" & Data2.Recordset!numlinea
    SQL = SQL & " WHERE " & Aux
    conn.Execute SQL
        
End Sub



Private Function comprobarAlbaranViculado() As Boolean
Dim RT As ADODB.Recordset
Dim RN As ADODB.Recordset
Dim SQL As String
Dim Fin As Boolean
Dim TotalCantidad As Currency
Dim canti As Currency
Dim Aux As String

On Error GoTo EcomprobarAlbaranViculado
    comprobarAlbaranViculado = False
    
    'Comprobaremos que las lineas son las mismas
    SQL = "Select codartic,cantidad,numlinea from sliped where numpedcl=" & Data1.Recordset!numpedcl & " ORDER BY numlinea"
    Set RT = New ADODB.Recordset
    RT.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    SQL = "Select slialb.*,conjunto  from ariges" & EmprMorales & ".slialb,sartic where slialb.codartic=sartic.codartic and codtipom='ALV' and numalbar = " & CtaBancoPropi & " ORDER BY numlinea"
    Set RN = New ADODB.Recordset
    RN.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    SQL = ""
    Fin = False
    TotalCantidad = 0
    While Not Fin
        If RT.EOF Then
            Fin = True
        Else
            If RN.EOF Then
                'NO es EOF el pedido y si el albaran de morales
                SQL = "No hay mas lineas en albaran de produccion"
                Fin = True
            Else
                If RT!codartic <> RN!codartic Then
                    SQL = SQL & " Articulo distinto(pedido/albaran): " & RT!codartic & "/" & RN!codartic
                Else
                    If RT!numlinea <> RN!numlinea Then
                        SQL = SQL & " Nº Linea distinto(pedido/albaran): " & RT!codartic & "     " & RT!numlinea & "/" & RN!numlinea
                    Else
                        If RT!Cantidad <> RN!Cantidad Then
                            SQL = SQL & " Distinta cantidad(pedido/albaran): " & RT!codartic & "     " & RT!Cantidad & "/" & RN!Cantidad
                        Else
                            'OK
                            If RN!Conjunto = 1 Then TotalCantidad = TotalCantidad + RT!Cantidad
                        End If
                    End If
                End If
                RT.MoveNext
                RN.MoveNext
            End If
        End If
    Wend
    RT.Close
    RN.Close
    
    'Comprobaremos los lotes
    If TotalCantidad > 0 Then
            Aux = "Select * from ariges" & EmprMorales & ".slialblotes where codtipom='ALV' and numalbar = " & CtaBancoPropi
            RN.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            canti = 0
            While Not RN.EOF
                canti = canti + RN!Cantidad
                RN.MoveNext
            Wend
            RN.Close
            If canti <> TotalCantidad Then SQL = SQL & vbCrLf & "Cantidades lotes asignadas incorrectas(lotes/albaran): " & canti & " / " & TotalCantidad
            
    End If
    
    
    If SQL <> "" Then
        SQL = "Error albaran produccion: " & vbCrLf & vbCrLf & SQL & vbCrLf & "¿Continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then comprobarAlbaranViculado = True
    Else
        comprobarAlbaranViculado = True
    End If
EcomprobarAlbaranViculado:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    CtaBancoPropi = ""
    Set RT = Nothing
    Set RN = Nothing
End Function


Private Sub PackingList()

    If EsHistorico Then Exit Sub
    If Me.Data1.Recordset Is Nothing Then Exit Sub
    If Me.Data1.Recordset.EOF Then Exit Sub
    
    If CodTipoMov <> "PEV" Then
        MsgBox "Solo pedidos de venta", vbExclamation
        Exit Sub
    End If
    
    


    
    


    
    
    

    If Not PonerParamRPT(34, "", CByte(NumRegElim), CadenaSQL) Then Exit Sub
    'El nombre sera el que tiene acabado en ALB
      
    NumRegElim = InStr(1, CadenaSQL, ".")
    CadenaSQL = Mid(CadenaSQL, 1, NumRegElim - 1) & "PED.rpt"

    
    
    
    With frmImprimir
            .FormulaSeleccion = "{sliped.numpedcl}=" & Data1.Recordset!numpedcl
            .OtrosParametros = ""
            .NumeroParametros = 0
            .SoloImprimir = False
            .EnvioEMail = False
            .NombreRPT = CadenaSQL
            .Titulo = "Packing List"
            .opcion = 53
            .Show vbModal
    End With
    
    
End Sub


