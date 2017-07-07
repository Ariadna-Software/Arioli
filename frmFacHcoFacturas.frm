VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmFacHcoFacturas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Histórico de Facturas Clientes"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   4335
   ClientWidth     =   14460
   Icon            =   "frmFacHcoFacturas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmFacHcoFacturas.frx":000C
   ScaleHeight     =   7185
   ScaleWidth      =   14460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   710
      Left            =   120
      TabIndex        =   123
      Top             =   400
      Width           =   12615
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   315
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   5
         Left            =   8010
         MaxLength       =   40
         TabIndex        =   6
         Tag             =   "Nombre Cliente|T|N|||scafac|nomclien||N|"
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwww aq"
         Top             =   240
         Width           =   4350
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   7125
         MaxLength       =   6
         TabIndex        =   5
         Tag             =   "Cod. Cliente|N|N|0|999999|scafac|codclien|000000|N|"
         Text            =   "Text1"
         Top             =   240
         Width           =   780
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   1290
         TabIndex        =   1
         Tag             =   "Tipo Factura|T|N|||scafac|codtipom||S|"
         Text            =   "Text3"
         Top             =   315
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   2670
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Fecha Factura|F|N|||scafac|fecfactu|dd/mm/yyyy|S|"
         Top             =   315
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
         Tag             =   "Nº Factura|N|N|||scafac|numfactu|0000000|S|"
         Text            =   "Text1 7"
         Top             =   315
         Width           =   980
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Contabilizado"
         Height          =   375
         Left            =   4440
         TabIndex        =   4
         Tag             =   "Contabilizado|N|N|0|1|scafac|intconta||N|"
         Top             =   195
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         Height          =   255
         Index           =   0
         Left            =   6240
         TabIndex        =   127
         Top             =   240
         Width           =   495
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   6855
         ToolTipText     =   "Buscar cliente"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Fact."
         Height          =   255
         Index           =   29
         Left            =   2670
         TabIndex        =   126
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Factura"
         Height          =   255
         Index           =   28
         Left            =   240
         TabIndex        =   125
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Fact."
         Height          =   255
         Index           =   27
         Left            =   1320
         TabIndex        =   124
         Top             =   120
         Width           =   795
      End
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   9960
      Top             =   6840
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
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   10200
      Top             =   6840
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
   Begin TabDlg.SSTab SSTab2 
      Height          =   5400
      Left            =   120
      TabIndex        =   36
      Tag             =   "Fecha Oferta|F|N|||scapre|fecentre|dd/mm/yyyy|N|"
      Top             =   1095
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   9525
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   5
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
      TabPicture(0)   =   "frmFacHcoFacturas.frx":0A0E
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FrameCliente"
      Tab(0).Control(1)=   "FrameFactura"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Albaranes"
      TabPicture(1)   =   "frmFacHcoFacturas.frx":0A2A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtAux(11)"
      Tab(1).Control(1)=   "cmdaux"
      Tab(1).Control(2)=   "txtAux(10)"
      Tab(1).Control(3)=   "txtAux(9)"
      Tab(1).Control(4)=   "Text3(15)"
      Tab(1).Control(5)=   "Text3(14)"
      Tab(1).Control(6)=   "txtAux3(2)"
      Tab(1).Control(7)=   "txtAux3(1)"
      Tab(1).Control(8)=   "txtAux3(0)"
      Tab(1).Control(9)=   "txtAux(5)"
      Tab(1).Control(10)=   "txtAux(3)"
      Tab(1).Control(11)=   "txtAux(2)"
      Tab(1).Control(12)=   "txtAux(1)"
      Tab(1).Control(13)=   "txtAux(0)"
      Tab(1).Control(14)=   "Text2(3)"
      Tab(1).Control(15)=   "Text3(3)"
      Tab(1).Control(16)=   "Text3(4)"
      Tab(1).Control(17)=   "Text3(5)"
      Tab(1).Control(18)=   "Text3(7)"
      Tab(1).Control(19)=   "Text3(6)"
      Tab(1).Control(20)=   "Text3(8)"
      Tab(1).Control(21)=   "Text2(0)"
      Tab(1).Control(22)=   "Text3(0)"
      Tab(1).Control(23)=   "Text2(1)"
      Tab(1).Control(24)=   "Text3(1)"
      Tab(1).Control(25)=   "Text2(2)"
      Tab(1).Control(26)=   "Text3(2)"
      Tab(1).Control(27)=   "txtAux(4)"
      Tab(1).Control(28)=   "txtAux(6)"
      Tab(1).Control(29)=   "txtAux(7)"
      Tab(1).Control(30)=   "txtAux(8)"
      Tab(1).Control(31)=   "DataGrid1"
      Tab(1).Control(32)=   "DataGrid2"
      Tab(1).Control(33)=   "imgBuscar(6)"
      Tab(1).Control(34)=   "imgBuscar(9)"
      Tab(1).Control(35)=   "imgBuscar(8)"
      Tab(1).Control(36)=   "Label1(40)"
      Tab(1).Control(37)=   "Label1(22)"
      Tab(1).Control(38)=   "Label1(18)"
      Tab(1).Control(39)=   "Label1(6)"
      Tab(1).Control(40)=   "Label1(2)"
      Tab(1).Control(41)=   "Label1(21)"
      Tab(1).Control(42)=   "Label1(24)"
      Tab(1).Control(43)=   "Label1(23)"
      Tab(1).Control(44)=   "Label1(9)"
      Tab(1).Control(45)=   "imgBuscar(7)"
      Tab(1).ControlCount=   46
      TabCaption(2)   =   "  Datos carga"
      TabPicture(2)   =   "frmFacHcoFacturas.frx":0A46
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label1(73)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label1(72)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label1(71)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label1(70)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label1(69)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label1(68)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label1(67)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label1(66)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label1(65)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Label1(64)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Label1(63)"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Label1(62)"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Label1(61)"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Label1(60)"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "Label1(59)"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "Label1(58)"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "Label1(57)"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "Label1(53)"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "Label1(56)"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "Label1(55)"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "Label1(54)"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "Label1(48)"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "chkCarga(3)"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "chkCarga(2)"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "chkCarga(1)"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "Text3(30)"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "chkCarga(0)"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).Control(27)=   "Text3(32)"
      Tab(2).Control(27).Enabled=   0   'False
      Tab(2).Control(28)=   "Text3(31)"
      Tab(2).Control(28).Enabled=   0   'False
      Tab(2).Control(29)=   "Text3(20)"
      Tab(2).Control(29).Enabled=   0   'False
      Tab(2).Control(30)=   "Text3(29)"
      Tab(2).Control(30).Enabled=   0   'False
      Tab(2).Control(31)=   "Text3(25)"
      Tab(2).Control(31).Enabled=   0   'False
      Tab(2).Control(32)=   "Text3(33)"
      Tab(2).Control(32).Enabled=   0   'False
      Tab(2).Control(33)=   "Text3(26)"
      Tab(2).Control(33).Enabled=   0   'False
      Tab(2).Control(34)=   "Text3(24)"
      Tab(2).Control(34).Enabled=   0   'False
      Tab(2).Control(35)=   "Text3(19)"
      Tab(2).Control(35).Enabled=   0   'False
      Tab(2).Control(36)=   "Text3(18)"
      Tab(2).Control(36).Enabled=   0   'False
      Tab(2).Control(37)=   "Text3(28)"
      Tab(2).Control(37).Enabled=   0   'False
      Tab(2).Control(38)=   "Text3(22)"
      Tab(2).Control(38).Enabled=   0   'False
      Tab(2).Control(39)=   "Text3(21)"
      Tab(2).Control(39).Enabled=   0   'False
      Tab(2).Control(40)=   "Text3(17)"
      Tab(2).Control(40).Enabled=   0   'False
      Tab(2).Control(41)=   "Text3(27)"
      Tab(2).Control(41).Enabled=   0   'False
      Tab(2).Control(42)=   "Text3(23)"
      Tab(2).Control(42).Enabled=   0   'False
      Tab(2).Control(43)=   "Text3(34)"
      Tab(2).Control(43).Enabled=   0   'False
      Tab(2).ControlCount=   44
      TabCaption(3)   =   "Observaciones"
      TabPicture(3)   =   "frmFacHcoFacturas.frx":0A62
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Text1(46)"
      Tab(3).Control(1)=   "Text3(16)"
      Tab(3).Control(2)=   "Text3(13)"
      Tab(3).Control(3)=   "Text3(12)"
      Tab(3).Control(4)=   "Text3(11)"
      Tab(3).Control(5)=   "Text3(10)"
      Tab(3).Control(6)=   "Text3(9)"
      Tab(3).Control(7)=   "FrObserva2"
      Tab(3).Control(8)=   "Label1(46)"
      Tab(3).Control(9)=   "imgBuscar(10)"
      Tab(3).Control(10)=   "Label1(47)"
      Tab(3).ControlCount=   11
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   34
         Left            =   5040
         MaxLength       =   10
         TabIndex        =   194
         Tag             =   "O1|T|S|||scafac1|Hora||N|"
         Top             =   960
         Width           =   945
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Index           =   23
         Left            =   3720
         MaxLength       =   60
         TabIndex        =   171
         Tag             =   "O1|T|S|||scafac1|TransEmpresa||N|"
         Text            =   "Text15"
         Top             =   1920
         Width           =   3030
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Index           =   27
         Left            =   3720
         MaxLength       =   30
         TabIndex        =   170
         Tag             =   "O1|T|S|||scafac1|TransConductor||N|"
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
         Top             =   2640
         Width           =   3405
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   17
         Left            =   3720
         MaxLength       =   10
         TabIndex        =   169
         Tag             =   "Fecha carga|F|S|||scafac1|FechaCarga|dd/mm/yyyy|N|"
         Top             =   960
         Width           =   1185
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   21
         Left            =   11040
         MaxLength       =   10
         TabIndex        =   168
         Tag             =   "BrutoKg|N|S|1||scafac1|TransBruto|#,##0||"
         Top             =   960
         Width           =   1065
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   22
         Left            =   12360
         MaxLength       =   10
         TabIndex        =   167
         Tag             =   "TaraKg|N|S|1||scafac1|TransTara|#,##0||"
         Top             =   960
         Width           =   1065
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Index           =   28
         Left            =   7320
         MaxLength       =   30
         TabIndex        =   166
         Tag             =   "O1|T|S|||scafac1|TransCondDNI||N|"
         Text            =   " "
         Top             =   2640
         Width           =   2205
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   18
         Left            =   6240
         MaxLength       =   10
         TabIndex        =   165
         Tag             =   "O1|T|S|||scafac1|Muestra||N|"
         Top             =   960
         Width           =   1545
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   19
         Left            =   7920
         MaxLength       =   10
         TabIndex        =   164
         Tag             =   "Deposito|N|S|1|20|scafac1|Deposito|0000|N|"
         Top             =   960
         Width           =   825
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Index           =   24
         Left            =   7320
         MaxLength       =   6
         TabIndex        =   163
         Tag             =   "O1|T|S|||scafac1|TransMatricula||N|"
         Text            =   "Text15"
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   26
         Left            =   11760
         MaxLength       =   10
         TabIndex        =   162
         Tag             =   "Bocas|N|S|1|100|scafac1|TransNumBocas|00|N|"
         Top             =   1920
         Width           =   1305
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Index           =   33
         Left            =   3720
         MaxLength       =   30
         TabIndex        =   161
         Tag             =   "O1|T|S|||scafac1|TransObsPrecintos||N|"
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
         Top             =   4920
         Width           =   7245
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Index           =   25
         Left            =   9360
         MaxLength       =   10
         TabIndex        =   160
         Tag             =   "O1|T|S|||scafac1|TransMatRemolque||N|"
         Text            =   "Text15"
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Index           =   29
         Left            =   9840
         MaxLength       =   30
         TabIndex        =   159
         Tag             =   "O1|T|S|||scafac1|TransDestino||N|"
         Text            =   " "
         Top             =   2640
         Width           =   3165
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   20
         Left            =   9000
         MaxLength       =   40
         TabIndex        =   158
         Tag             =   "O1|T|S|||scafac1|TransAcidez||N|"
         Top             =   960
         Width           =   1785
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   31
         Left            =   4680
         MaxLength       =   10
         TabIndex        =   157
         Tag             =   "Deposito|N|S|1|50|scafac1|TransLacradasCoop|00|N|"
         Top             =   4320
         Width           =   825
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   32
         Left            =   5880
         MaxLength       =   10
         TabIndex        =   156
         Tag             =   "Deposito|N|S||50|scafac1|TransLacradasCompr|00|N|"
         Top             =   4320
         Width           =   825
      End
      Begin VB.CheckBox chkCarga 
         Caption         =   "Ticket báscula"
         Height          =   375
         Index           =   0
         Left            =   7440
         TabIndex        =   155
         Top             =   4200
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Height          =   555
         Index           =   30
         Left            =   3720
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   154
         Tag             =   "O1|T|S|||scafac1|TransMercancia||N|"
         Text            =   "frmFacHcoFacturas.frx":0A7E
         Top             =   3240
         Width           =   8205
      End
      Begin VB.CheckBox chkCarga 
         Caption         =   "CMR"
         Height          =   375
         Index           =   1
         Left            =   9240
         TabIndex        =   153
         Top             =   4200
         Width           =   975
      End
      Begin VB.CheckBox chkCarga 
         Caption         =   "Certificado limpieza"
         Height          =   375
         Index           =   2
         Left            =   10440
         TabIndex        =   152
         Top             =   4200
         Width           =   2055
      End
      Begin VB.CheckBox chkCarga 
         Caption         =   "Otros"
         Height          =   375
         Index           =   3
         Left            =   12600
         TabIndex        =   151
         Top             =   4200
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   1995
         Index           =   46
         Left            =   -72360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   149
         Tag             =   "Obs|T|S|||scafac|packingobs|||"
         Text            =   "frmFacHcoFacturas.frx":0A9D
         Top             =   3360
         Width           =   8985
      End
      Begin VB.TextBox Text3 
         Height          =   660
         Index           =   16
         Left            =   -72360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   147
         Tag             =   "Observación 6|T|S|||scafac1|observa6||N|"
         Text            =   "frmFacHcoFacturas.frx":0AB2
         Top             =   2520
         Width           =   8940
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   13
         Left            =   -72360
         MaxLength       =   80
         TabIndex        =   146
         Tag             =   "Observación 5|T|S|||scafac1|observa5||N|"
         Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
         Top             =   2040
         Width           =   8940
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   12
         Left            =   -72360
         MaxLength       =   80
         TabIndex        =   145
         Tag             =   "Observación 4|T|S|||scafac1|observa4||N|"
         Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
         Top             =   1710
         Width           =   8940
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   11
         Left            =   -72360
         MaxLength       =   80
         TabIndex        =   144
         Tag             =   "Observación 3|T|S|||scafac1|observa3||N|"
         Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
         Top             =   1380
         Width           =   8940
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   10
         Left            =   -72360
         MaxLength       =   80
         TabIndex        =   143
         Tag             =   "Observación 2|T|S|||scafac1|observa2||N|"
         Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
         Top             =   1050
         Width           =   8940
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   9
         Left            =   -72360
         MaxLength       =   80
         TabIndex        =   142
         Tag             =   "Observación 1|T|S|||scafac1|observa1||N|"
         Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
         Top             =   720
         Width           =   8940
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   11
         Left            =   -71880
         MaxLength       =   30
         TabIndex        =   109
         Text            =   "palets"
         Top             =   4080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdaux 
         Caption         =   "+"
         Height          =   320
         Left            =   -66360
         TabIndex        =   117
         Top             =   4320
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   10
         Left            =   -66120
         MaxLength       =   30
         TabIndex        =   139
         Text            =   "Dto2"
         Top             =   4320
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   9
         Left            =   -66720
         MaxLength       =   30
         TabIndex        =   116
         Tag             =   "Dto 2|N|N|||slifac|codprovex|0||"
         Text            =   "Dto2"
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   15
         Left            =   -63840
         MaxLength       =   7
         TabIndex        =   129
         Tag             =   "Nº Terminal|N|S|||scafac1|numtermi||N|"
         Text            =   "Text1 7"
         Top             =   1800
         Width           =   885
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   14
         Left            =   -62640
         MaxLength       =   7
         TabIndex        =   128
         Tag             =   "Nº Venta|N|S|||scafac1|numventa|0000000|N|"
         Text            =   "Text1 7"
         Top             =   1800
         Width           =   1185
      End
      Begin VB.TextBox txtAux3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   2
         Left            =   -72960
         MaxLength       =   30
         TabIndex        =   122
         Tag             =   "Fecha Albaran|F|N|||scafac1|fechaalb|dd/mm/yyyy|N|"
         Text            =   "fecalbar"
         Top             =   2160
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   -73920
         MaxLength       =   15
         TabIndex        =   121
         Tag             =   "Nº Albaran|N|N|||scafac1|numalbar|0000000|N|"
         Text            =   "numalbar"
         Top             =   2160
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   -74640
         MaxLength       =   7
         TabIndex        =   120
         Tag             =   "Tipo Albaran|T|N|||scafac1|codtipoa||N|"
         Text            =   "codtipoa"
         Top             =   2160
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   5
         Left            =   -69960
         MaxLength       =   5
         TabIndex        =   111
         Text            =   "origp"
         Top             =   4320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   3
         Left            =   -71760
         MaxLength       =   12
         TabIndex        =   108
         Tag             =   "Cantidad|N|N|0||slifac|cantidad|#,###,###,##0.00|N|"
         Text            =   "cantidad"
         Top             =   4320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   2
         Left            =   -72840
         MaxLength       =   12
         TabIndex        =   107
         Tag             =   "Nombre Art.|T|N|||slifac|nomartic||N|"
         Text            =   "nomartic"
         Top             =   4320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   -73680
         MaxLength       =   12
         TabIndex        =   106
         Tag             =   "Art.|T|N|||slifac|codartic||N|"
         Text            =   "codartic"
         Top             =   4320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   -74640
         MaxLength       =   12
         TabIndex        =   105
         Tag             =   "Almacen|N|N|0|999|slifac|codalmac|000|N|"
         Text            =   "almacen"
         Top             =   4320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   3
         Left            =   -68160
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   104
         Text            =   "Text2"
         Top             =   2160
         Width           =   3525
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   -68880
         MaxLength       =   30
         TabIndex        =   30
         Tag             =   "Cod. Envío|N|N|0|999|scafac1|codenvio|000|N|"
         Text            =   "Text1"
         Top             =   2160
         Width           =   660
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   -63840
         MaxLength       =   7
         TabIndex        =   98
         Tag             =   "Nº Oferta|N|S|||scafac1|numofert|0000000|N|"
         Text            =   "Text1 7"
         Top             =   1800
         Width           =   885
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   -62640
         MaxLength       =   10
         TabIndex        =   97
         Tag             =   "Fecha Oferta|F|S|||scafac1|fecofert|dd/mm/yyyy|N|"
         Top             =   1800
         Width           =   1185
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   7
         Left            =   -63240
         MaxLength       =   10
         TabIndex        =   96
         Tag             =   "Fecha Pedido|F|S|||scafac1|fecpedcl|dd/mm/yyyy|N|"
         Top             =   960
         Width           =   1185
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   6
         Left            =   -64200
         MaxLength       =   7
         TabIndex        =   95
         Tag             =   "Nº Pedido|N|S|||scafac1|numpedcl|0000000|N|"
         Text            =   "Text1 7"
         Top             =   960
         Width           =   885
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   8
         Left            =   -61920
         MaxLength       =   10
         TabIndex        =   94
         Tag             =   "Semana Entrega|N|S|||scafac1|sementre||N|"
         Top             =   960
         Width           =   705
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   0
         Left            =   -68160
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   93
         Text            =   "Text2"
         Top             =   480
         Width           =   3525
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   -68880
         MaxLength       =   30
         TabIndex        =   27
         Tag             =   "Trabajador Albaran|N|N|0|9999|scafac1|codtraba|0000|N|"
         Text            =   "Text1"
         Top             =   480
         Width           =   660
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   1
         Left            =   -68160
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   92
         Text            =   "Text2"
         Top             =   1040
         Width           =   3525
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   -68880
         MaxLength       =   30
         TabIndex        =   28
         Tag             =   "Trabajador pedido|N|S|0|9999|scafac1|codtrab1|0000|N|"
         Text            =   "Text1"
         Top             =   1040
         Width           =   660
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   2
         Left            =   -68160
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   91
         Text            =   "Text2"
         Top             =   1600
         Width           =   3525
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   -68880
         MaxLength       =   30
         TabIndex        =   29
         Tag             =   "Preparador materia|N|N|0|9999|scafac1|codtrab2|0000|N|"
         Text            =   "Text1"
         Top             =   1600
         Width           =   660
      End
      Begin VB.Frame FrameFactura 
         Height          =   1980
         Left            =   -74400
         TabIndex        =   61
         Top             =   3240
         Width           =   12975
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   44
            Left            =   6720
            MaxLength       =   5
            TabIndex        =   135
            Tag             =   "% IVA 1|N|S|0|99.90|scafac|porciva3re|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1605
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   43
            Left            =   7320
            MaxLength       =   15
            TabIndex        =   134
            Tag             =   "Importe IVA 1|N|S|||scafac|imporiv3re|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1605
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   42
            Left            =   6720
            MaxLength       =   5
            TabIndex        =   133
            Tag             =   "% IVA 1|N|S|0|99.90|scafac|porciva2re|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1275
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   41
            Left            =   7320
            MaxLength       =   15
            TabIndex        =   132
            Tag             =   "Importe IVA 1|N|S|||scafac|imporiv2re|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1275
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   40
            Left            =   6720
            MaxLength       =   5
            TabIndex        =   131
            Tag             =   "% IVA 1|N|S|0|99.90|scafac|porciva1re|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   960
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   39
            Left            =   7320
            MaxLength       =   15
            TabIndex        =   130
            Tag             =   "Importe IVA 1|N|S|||scafac|imporiv1re|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   960
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFC0&
            Height          =   285
            Index           =   38
            Left            =   9720
            MaxLength       =   15
            TabIndex        =   86
            Tag             =   "Total Factura|N|N|||scafac|totalfac|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   37
            Left            =   5160
            MaxLength       =   15
            TabIndex        =   81
            Tag             =   "Importe IVA 3|N|S|||scafac|imporiv3|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1605
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   31
            Left            =   4560
            MaxLength       =   5
            TabIndex        =   80
            Tag             =   "% IVA 3|N|S|0|99.90|scafac|porciva3|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1605
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   28
            Left            =   2400
            MaxLength       =   4
            TabIndex        =   79
            Tag             =   "Cod. IVA 3|N|S|0|9999|scafac|codigiv3|0000|N|"
            Text            =   "Text1 7"
            Top             =   1605
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   34
            Left            =   3000
            MaxLength       =   15
            TabIndex        =   78
            Tag             =   "Base Imponible 3|N|S|||scafac|baseimp3|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1605
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   36
            Left            =   5160
            MaxLength       =   15
            TabIndex        =   77
            Tag             =   "Importe IVA 2|N|S|||scafac|imporiv2|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1275
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   30
            Left            =   4560
            MaxLength       =   5
            TabIndex        =   76
            Tag             =   "% IVA 2|N|S|0|99.90|scafac|porciva2|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1275
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   27
            Left            =   2400
            MaxLength       =   4
            TabIndex        =   75
            Tag             =   "Cod. IVA 2|N|S|0|9999|scafac|codigiv2|0000|N|"
            Text            =   "Text1 7"
            Top             =   1275
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   33
            Left            =   3000
            MaxLength       =   15
            TabIndex        =   74
            Tag             =   "Base Imponible 2 |N|S|||scafac|baseimp2|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1275
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   35
            Left            =   5160
            MaxLength       =   15
            TabIndex        =   73
            Tag             =   "Importe IVA 1|N|N|||scafac|imporiv1|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   960
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   29
            Left            =   4560
            MaxLength       =   5
            TabIndex        =   72
            Tag             =   "% IVA 1|N|S|0|99.90|scafac|porciva1|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   960
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   26
            Left            =   2400
            MaxLength       =   4
            TabIndex        =   71
            Tag             =   "Cod. IVA 1|N|S|0|9999|scafac|codigiv1|0000|N|"
            Text            =   "Text1 7"
            Top             =   960
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   32
            Left            =   3000
            MaxLength       =   15
            TabIndex        =   70
            Tag             =   "Base Imponible 1|N|N|||scafac|baseimp1|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   960
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   25
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   65
            Text            =   "Text1 7"
            Top             =   320
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   24
            Left            =   3960
            MaxLength       =   15
            TabIndex        =   64
            Tag             =   "Imp. Dto Gn|N|N|0||scafac|impdtogr|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   320
            Width           =   1365
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   23
            Left            =   2160
            MaxLength       =   15
            TabIndex        =   63
            Tag             =   "Imp. Dto PP|N|N|0||scafac|impdtopp|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   320
            Width           =   1365
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   22
            Left            =   240
            MaxLength       =   15
            TabIndex        =   62
            Tag             =   "Imp.Bruto|N|N|||scafac|brutofac|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   360
            Width           =   1485
         End
         Begin VB.Label Label1 
            Caption         =   "Importe RE"
            Height          =   195
            Index           =   44
            Left            =   7560
            TabIndex        =   138
            Top             =   720
            Width           =   825
         End
         Begin VB.Label Label1 
            Caption         =   "% RE"
            Height          =   255
            Index           =   43
            Left            =   6720
            TabIndex        =   137
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Importe IVA"
            Height          =   195
            Index           =   37
            Left            =   5520
            TabIndex        =   136
            Top             =   720
            Width           =   825
         End
         Begin VB.Line Line1 
            X1              =   2280
            X2              =   2280
            Y1              =   960
            Y2              =   1920
         End
         Begin VB.Label Label1 
            Caption         =   "Desglose IVA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   42
            Left            =   960
            TabIndex        =   119
            Top             =   1320
            Width           =   1155
         End
         Begin VB.Label Label1 
            Caption         =   "% IVA"
            Height          =   255
            Index           =   41
            Left            =   4560
            TabIndex        =   118
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "TOTAL FACTURA"
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
            Left            =   9720
            TabIndex        =   89
            Top             =   1320
            Width           =   1530
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
            Index           =   38
            Left            =   9360
            TabIndex        =   88
            Top             =   1560
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
            TabIndex        =   87
            Top             =   2160
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "Base impo. IVA"
            Height          =   255
            Index           =   33
            Left            =   3120
            TabIndex        =   85
            Top             =   720
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
            TabIndex        =   84
            Top             =   240
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
            TabIndex        =   83
            Top             =   240
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
            TabIndex        =   82
            Top             =   240
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "Base Imponible"
            Height          =   255
            Index           =   14
            Left            =   5880
            TabIndex        =   69
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Dto Gn"
            Height          =   255
            Index           =   12
            Left            =   4080
            TabIndex        =   68
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Dto PP"
            Height          =   255
            Index           =   11
            Left            =   2280
            TabIndex        =   67
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Bruto"
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   66
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   4
         Left            =   -70920
         MaxLength       =   12
         TabIndex        =   110
         Tag             =   "Precio|N|N|0|999999.0000|slifac|precioar|###,##0.0000|N|"
         Text            =   "Precio"
         Top             =   4320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   6
         Left            =   -69240
         MaxLength       =   5
         TabIndex        =   112
         Tag             =   "Dto 1|N|N|0|99.90|slifac|dtoline1|#0.00|N|"
         Text            =   "Dto1"
         Top             =   4320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   7
         Left            =   -68520
         MaxLength       =   30
         TabIndex        =   113
         Tag             =   "Dto 2|N|N|0|99.90|slifac|dtolinea|#0.00|N|"
         Text            =   "Dto2"
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   8
         Left            =   -67920
         MaxLength       =   12
         TabIndex        =   115
         Tag             =   "Importe|N|N|0||slifac|importel|#,###,###,##0.00|N|"
         Text            =   "Importe"
         Top             =   4320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame FrameCliente 
         Caption         =   "Datos Cliente"
         ForeColor       =   &H00972E0B&
         Height          =   2775
         Left            =   -74400
         TabIndex        =   38
         Top             =   360
         Width           =   12975
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   47
            Left            =   6840
            MaxLength       =   4
            TabIndex        =   20
            Tag             =   "IBAN|T|S|||scafac|iban|||"
            Text            =   "Text1 7"
            Top             =   2235
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   45
            Left            =   11040
            MaxLength       =   10
            TabIndex        =   140
            Tag             =   "Aportacion|N|S|||scafac|aportacion|#,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   2280
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   21
            Left            =   9120
            MaxLength       =   10
            TabIndex        =   24
            Tag             =   "Cuenta Bancaria|T|S|||scafac|cuentaba|0000000000|N|"
            Text            =   "Text1 7"
            Top             =   2235
            Width           =   1245
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   20
            Left            =   8640
            MaxLength       =   2
            TabIndex        =   23
            Tag             =   "Digito Control|T|S|||scafac|digcontr|00|N|"
            Text            =   "Text1 7"
            Top             =   2235
            Width           =   405
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   19
            Left            =   8040
            MaxLength       =   4
            TabIndex        =   22
            Tag             =   "Sucursal|N|S|0|9999|scafac|codsucur|0000|N|"
            Text            =   "Text1 7"
            Top             =   2235
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   18
            Left            =   7440
            MaxLength       =   4
            TabIndex        =   21
            Tag             =   "Banco|N|S|0|9999|scafac|codbanco|0000|N|"
            Text            =   "Text1 7"
            Top             =   2235
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   3
            Left            =   1125
            MaxLength       =   20
            TabIndex        =   12
            Text            =   "Text1 Text1 Text1 Te"
            Top             =   2220
            Width           =   1680
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   13
            Left            =   7470
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   15
            Tag             =   "Direccion/Dpto.|T|S|||scafac|nomdirec||N|"
            Text            =   "Text1"
            Top             =   285
            Width           =   3405
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   12
            Left            =   6885
            MaxLength       =   3
            TabIndex        =   14
            Tag             =   "Direccion/Dpto.|N|S|0|999|scafac|coddirec|000|N|"
            Text            =   "Text1"
            Top             =   285
            Width           =   540
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   11
            Left            =   1125
            MaxLength       =   30
            TabIndex        =   11
            Tag             =   "Provincia|T|N|||scafac|proclien||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text22"
            Top             =   1830
            Width           =   2445
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   9
            Left            =   1125
            MaxLength       =   6
            TabIndex        =   9
            Tag             =   "CPostal|T|N|||scafac|codpobla||N|"
            Text            =   "Text15"
            Top             =   1470
            Width           =   630
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   10
            Left            =   1755
            MaxLength       =   30
            TabIndex        =   10
            Tag             =   "Población|T|N|||scafac|pobclien||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
            Top             =   1470
            Width           =   3405
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   7
            Left            =   3480
            MaxLength       =   20
            TabIndex        =   13
            Tag             =   "teléfono Cliente|T|S|||scafac|telclien||N|"
            Text            =   "12345678911234567899"
            Top             =   2220
            Width           =   1605
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   6
            Left            =   1125
            MaxLength       =   15
            TabIndex        =   7
            Tag             =   "NIF Cliente|T|N|||scafac|nifclien||N|"
            Text            =   "123456789"
            Top             =   285
            Width           =   1590
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   14
            Left            =   6885
            MaxLength       =   4
            TabIndex        =   17
            Tag             =   "Cod. Agente|N|N|0|9999|scafac|codagent|0000|N|"
            Text            =   "Text1"
            Top             =   645
            Width           =   540
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   14
            Left            =   7470
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   44
            Text            =   "Text2"
            Top             =   645
            Width           =   3405
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   15
            Left            =   6885
            MaxLength       =   3
            TabIndex        =   19
            Tag             =   "Forma de Pago|N|N|0|999|scafac|codforpa|000|N|"
            Text            =   "Text1"
            Top             =   960
            Width           =   540
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   15
            Left            =   7470
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   40
            Text            =   "Text2"
            Top             =   990
            Width           =   3405
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   16
            Left            =   12045
            MaxLength       =   5
            TabIndex        =   16
            Tag             =   "Descuento P.Pago|N|N|0|99.90|scafac|dtoppago|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   270
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   17
            Left            =   12045
            MaxLength       =   5
            TabIndex        =   18
            Tag             =   "Descuento General|N|N|0|99.90|scafac|dtognral|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   630
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Height          =   675
            Index           =   8
            Left            =   1125
            MaxLength       =   35
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Tag             =   "Domicilio|T|N|||scafac|domclien||N|"
            Text            =   "frmFacHcoFacturas.frx":0B03
            Top             =   645
            Width           =   4030
         End
         Begin VB.Label Label1 
            Caption         =   "Aportación"
            Height          =   255
            Index           =   45
            Left            =   11040
            TabIndex        =   141
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   6600
            ToolTipText     =   "Buscar agente"
            Top             =   645
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta"
            Height          =   255
            Index           =   8
            Left            =   8160
            TabIndex        =   60
            Top             =   2040
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "DC"
            Height          =   255
            Index           =   5
            Left            =   7680
            TabIndex        =   59
            Top             =   2040
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "Sucursal"
            Height          =   255
            Index           =   4
            Left            =   6960
            TabIndex        =   58
            Top             =   2040
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "IBAN"
            Height          =   255
            Index           =   3
            Left            =   6360
            TabIndex        =   57
            Top             =   2280
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Ref. Cliente"
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   51
            Top             =   2220
            Width           =   855
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   2
            Left            =   855
            ToolTipText     =   "Buscar población"
            Top             =   1485
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Direc./Dpto"
            Height          =   255
            Index           =   1
            Left            =   5700
            TabIndex        =   50
            Top             =   285
            Width           =   855
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   3
            Left            =   6600
            ToolTipText     =   "Buscar direc./dpto"
            Top             =   285
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Provincia"
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   49
            Top             =   1830
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Población"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   48
            Top             =   1470
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Tfno:"
            Height          =   255
            Index           =   19
            Left            =   3000
            TabIndex        =   47
            Top             =   2280
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "NIF"
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   46
            Top             =   285
            Width           =   615
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   855
            ToolTipText     =   "Buscar cliente varios"
            Top             =   300
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Agente"
            Height          =   255
            Index           =   34
            Left            =   5700
            TabIndex        =   45
            Top             =   645
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Forma Pago"
            Height          =   255
            Index           =   15
            Left            =   5700
            TabIndex        =   43
            Top             =   990
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. P.P"
            Height          =   255
            Index           =   25
            Left            =   11340
            TabIndex        =   42
            Top             =   270
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. Gral"
            Height          =   255
            Index           =   26
            Left            =   11355
            TabIndex        =   41
            Top             =   630
            Width           =   735
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   5
            Left            =   6600
            ToolTipText     =   "Buscar forma de pago"
            Top             =   990
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Domicilio"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   39
            Top             =   645
            Width           =   735
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmFacHcoFacturas.frx":0B27
         Height          =   2745
         Left            =   -74760
         TabIndex        =   56
         Top             =   2550
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   4842
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
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frmFacHcoFacturas.frx":0B3C
         Height          =   1945
         Left            =   -74760
         TabIndex        =   90
         Top             =   520
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   3440
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
         Caption         =   "Albaranes de la Factura"
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
      Begin VB.Frame FrObserva2 
         BorderStyle     =   0  'None
         Caption         =   "Obs. en factura"
         ForeColor       =   &H00972E0B&
         Height          =   2775
         Left            =   -74880
         TabIndex        =   193
         Tag             =   "Observación 4|T|S|||scafac1|observa4||N|"
         Top             =   480
         Width           =   13695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hora"
         Height          =   195
         Index           =   48
         Left            =   5040
         TabIndex        =   195
         Top             =   720
         Width           =   345
      End
      Begin VB.Label Label1 
         Caption         =   "Datos almazara"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   54
         Left            =   120
         TabIndex        =   192
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Transporte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   55
         Left            =   120
         TabIndex        =   191
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   56
         Left            =   120
         TabIndex        =   190
         Top             =   4080
         Width           =   2535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fec. carga"
         Height          =   195
         Index           =   53
         Left            =   3720
         TabIndex        =   189
         Top             =   720
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Muestra"
         Height          =   195
         Index           =   57
         Left            =   6240
         TabIndex        =   188
         Top             =   720
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Depósito"
         Height          =   195
         Index           =   58
         Left            =   7920
         TabIndex        =   187
         Top             =   720
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Bruto(kg)"
         Height          =   195
         Index           =   59
         Left            =   11040
         TabIndex        =   186
         Top             =   720
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tara  (kg)"
         Height          =   195
         Index           =   60
         Left            =   12360
         TabIndex        =   185
         Top             =   720
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Index           =   61
         Left            =   3720
         TabIndex        =   184
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Matrícula"
         Height          =   195
         Index           =   62
         Left            =   7320
         TabIndex        =   183
         Top             =   1680
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Bocas precintadas"
         Height          =   195
         Index           =   63
         Left            =   11760
         TabIndex        =   182
         Top             =   1680
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre conductor"
         Height          =   195
         Index           =   64
         Left            =   3720
         TabIndex        =   181
         Top             =   2400
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DNI conductor"
         Height          =   195
         Index           =   65
         Left            =   7320
         TabIndex        =   180
         Top             =   2400
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Precintos"
         Height          =   195
         Index           =   66
         Left            =   3720
         TabIndex        =   179
         Top             =   4680
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Matrí. remolque"
         Height          =   195
         Index           =   67
         Left            =   9360
         TabIndex        =   178
         Top             =   1680
         Width           =   1110
      End
      Begin VB.Label Label1 
         Caption         =   "Mercancia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   68
         Left            =   120
         TabIndex        =   177
         Top             =   3120
         Width           =   2535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Destino"
         Height          =   195
         Index           =   69
         Left            =   9840
         TabIndex        =   176
         Top             =   2400
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Acidez"
         Height          =   195
         Index           =   70
         Left            =   9000
         TabIndex        =   175
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Coop"
         Height          =   195
         Index           =   71
         Left            =   4800
         TabIndex        =   174
         Top             =   4080
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Comprador"
         Height          =   195
         Index           =   72
         Left            =   5880
         TabIndex        =   173
         Top             =   4080
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lacradas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   73
         Left            =   3720
         TabIndex        =   172
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Obs"
         Height          =   255
         Index           =   46
         Left            =   -74880
         TabIndex        =   150
         Top             =   3480
         Width           =   2175
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   10
         Left            =   -72600
         ToolTipText     =   "Buscar cliente varios"
         Top             =   3480
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones en factura"
         Height          =   255
         Index           =   47
         Left            =   -74880
         TabIndex        =   148
         Top             =   720
         Width           =   2055
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   -69120
         ToolTipText     =   "Buscar trabajador"
         Top             =   480
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   9
         Left            =   -69120
         ToolTipText     =   "Buscar forma de envio"
         Top             =   2160
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   -69120
         ToolTipText     =   "Buscar trabajador"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Oferta"
         Height          =   255
         Index           =   40
         Left            =   -63840
         TabIndex        =   103
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Oferta"
         Height          =   255
         Index           =   22
         Left            =   -62400
         TabIndex        =   102
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Pedido"
         Height          =   255
         Index           =   18
         Left            =   -63240
         TabIndex        =   101
         Top             =   615
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Pedido"
         Height          =   255
         Index           =   6
         Left            =   -64080
         TabIndex        =   100
         Top             =   615
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Sem. Entrega"
         Height          =   255
         Index           =   2
         Left            =   -62040
         TabIndex        =   99
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Trabajador Albaran"
         Height          =   255
         Index           =   21
         Left            =   -70560
         TabIndex        =   55
         Top             =   525
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo  Envío"
         Height          =   195
         Index           =   24
         Left            =   -70560
         TabIndex        =   54
         Top             =   2160
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "Prepar. Material"
         Height          =   255
         Index           =   23
         Left            =   -70560
         TabIndex        =   53
         Top             =   1680
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Trabajador Pedido"
         Height          =   255
         Index           =   9
         Left            =   -70560
         TabIndex        =   52
         Top             =   1080
         Width           =   1425
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   -69120
         ToolTipText     =   "Buscar trabajador"
         Top             =   1080
         Width           =   240
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   16
      Left            =   2400
      Locked          =   -1  'True
      MaxLength       =   60
      TabIndex        =   114
      Text            =   "Text2 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwqa"
      Top             =   6780
      Visible         =   0   'False
      Width           =   6885
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   120
      TabIndex        =   32
      Top             =   6615
      Width           =   2175
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   240
         TabIndex        =   33
         Top             =   180
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   13050
      TabIndex        =   26
      Top             =   6720
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   11760
      TabIndex        =   25
      Top             =   6720
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   14460
      _ExtentX        =   25506
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
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "0"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Lineas Factura"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Factura"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "packing List"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   8400
         TabIndex        =   35
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   13050
      TabIndex        =   31
      Top             =   6720
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data3 
      Height          =   330
      Left            =   3000
      Top             =   1080
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
   Begin VB.Label Label1 
      Caption         =   "Ampliación Línea"
      Height          =   255
      Index           =   35
      Left            =   2400
      TabIndex        =   37
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
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
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
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
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
Attribute VB_Name = "frmFacHcoFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========== VBLES PUBLICAS ====================
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'Si se llama de la busqueda en el frmAlmMovimArticulos se accede
'a las tablas del Albaran o de Facturas de movimiento seleccionado (solo consulta)
Public hcoCodMovim As String 'cod. movim
Public hcoCodTipoM As String 'Codigo detalle de Movimiento(ALC)
Public hcoFechaMov As String 'fecha del movimiento

'========== VBLES PRIVADAS ====================
Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
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
Private WithEvents frmFE As frmFacFormasEnvio  'Form Formas de Envio
Attribute frmFE.VB_VarHelpID = -1
Private WithEvents frmP As frmComProveedores
Attribute frmP.VB_VarHelpID = -1
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
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------

Dim ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim CodTipoMov As String
'Codigo tipo de movimiento en función del valor en la tabla de parámetros: stipom

Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

Dim PrimeraVez As Boolean

Dim EsCabecera As Boolean
'Para saber en MandaBusquedaPrevia si busca en la tabla scapla o en la tabla sdirec


Dim EsDeVarios As Boolean
'Si el cliente mostrado es de Varios o No

'SQL de la tabla principal del formulario
Private CadenaConsulta As String

Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private NomTablaLineas As String 'Nombre de la Tabla de lineas
Private Kcampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1


Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos
Private HaCambiadoCP As Boolean
'Para saber si tras haber vuelto de prismaticos ha cambiado el valor del CPostal


Private BuscaChekc As String

Private Sub Check1_Click()
     If Modo = 1 Then CheckCadenaBusqueda Check1, BuscaChekc
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim I As Integer

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda

        Case 4  'MODIFICAR
            If DatosOk Then
               If ModificarFactura Then
                    Espera 0.2
                    TerminaBloquear
                    PosicionarData
                    FormatoDatosTotales
                    I = data3.Recordset.AbsolutePosition
                    PonerCamposLineas
                    SituarDataPosicion data3, CLng(I), ""
                End If
            End If
            
         Case 5 'InsertarModificar LINEAS
            If ModificaLineas = 2 Then 'MODIFICAR lineas
                If ModificarLinea Then
                    TerminaBloquear
                    CargaGrid DataGrid1, Data2, True
                    ModificaLineas = 0
                    PonerBotonCabecera True
                    BloquearTxt Text2(16), True
            
                    LLamaLineas Modo, 0, "DataGrid1"
                    PosicionarData
                Else
                    TerminaBloquear
                End If
                Me.DataGrid1.Enabled = True
            End If
    End Select
    Screen.MousePointer = vbDefault

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdAux_Click()
    Set frmP = New frmComProveedores
    frmP.DatosADevolverBusqueda = "0|1|"
    frmP.Show vbModal
    Set frmP = Nothing

End Sub

Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 0, 1 'Busqueda, Insertar
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            LimpiarDataGrids
            PonerModo 0
            LLamaLineas Modo, 0, "DataGrid2"
            PonerFoco Text1(0)
            
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
            
        Case 5 'Lineas Detalle
            TerminaBloquear
            BloquearTxt Text2(16), True
            If ModificaLineas = 1 Then 'INSERTAR
                ModificaLineas = 0
                DataGrid1.AllowAddNew = False
                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
            End If
            ModificaLineas = 0
            LLamaLineas Modo, 0, "DataGrid1"
            PonerBotonCabecera True
            Me.DataGrid1.Enabled = True
    End Select
End Sub


Private Sub BotonBuscar()
Dim anc As Single

    'Buscar
    If Modo <> 1 Then
        BuscaChekc = ""
        LimpiarCampos
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
        PonerModo 1
        
        'poner los txtaux para buscar por lineas de albaran
        anc = DataGrid2.Top
        If DataGrid2.Row < 0 Then
            anc = anc + 440
        Else
            anc = anc + DataGrid2.RowTop(DataGrid2.Row) + 20
        End If
        LLamaLineas Modo, anc, "DataGrid2"
        
        
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
Dim C As String
        'Si no es del B NO puede ver el almacen
    If Not vUsu.TrabajadorB Then
        C = "  scafac.codtipom <> 'FAZ'"
    Else
        C = ""
    End If


    If chkVistaPrevia.Value = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia C
    Else
        LimpiarCampos
        LimpiarDataGrids
        CadenaConsulta = "Select scafac.* "
        CadenaConsulta = CadenaConsulta & "from " & NombreTabla
        If C <> "" Then CadenaConsulta = CadenaConsulta & " WHERE " & C
'        CadenaConsulta = CadenaConsulta & " WHERE scafac.codtipom='" & CodTipoMov & "'"
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

    'solo se puede modificar la factura si no esta contabilizada
    If FactContabilizada2(True) Then
        TerminaBloquear
        Exit Sub
    End If
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    PonerFocoChk Me.Check1
         
    'Si es Cliente de Varios no se pueden modificar sus datos
    DeVarios = EsClienteVarios(Text1(4).Text)
    BloquearDatosCliente (DeVarios)
End Sub


Private Sub BotonModificarLinea()
'Modificar una linea
Dim vWhere As String
Dim anc As Single
Dim J As Byte

    On Error GoTo EModificarLinea


     'solo se puede modificar la factura si no esta contabilizada
    If FactContabilizada2(False) Then
        TerminaBloquear
        Exit Sub
    End If

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then  '1= Insertar
        TerminaBloquear
        Exit Sub
    End If
    
    If Data2.Recordset.EOF Then
        TerminaBloquear
        Exit Sub
    End If
    
    vWhere = ObtenerWhereCP(False)
    vWhere = vWhere & " AND codtipoa='" & data3.Recordset.Fields!Codtipoa & "' AND numalbar=" & data3.Recordset.Fields!NumAlbar
    vWhere = vWhere & " and numlinea=" & Data2.Recordset!numlinea
    If Not BloqueaRegistro(NomTablaLineas, vWhere) Then
        TerminaBloquear
        Exit Sub
    End If

    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        If DataGrid1.Bookmark - DataGrid1.FirstRow > 0 Then
            J = DataGrid1.Bookmark - DataGrid1.FirstRow
            DataGrid1.Scroll 0, J
            DataGrid1.Refresh
        End If
    End If
    
'    anc = ObtenerAlto(Me.DataGrid1)
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 210
    Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 10
    End If

    For J = 0 To 2
        txtAux(J).Text = DataGrid1.Columns(J + 5).Text
    Next J
    Text2(16).Text = DataGrid1.Columns(8).Text
    For J = 4 To 10
        txtAux(J).Text = DataGrid1.Columns(J + 7).Text
    Next J
    txtAux(11).Text = DataGrid1.Columns(9).Text
    txtAux(3).Text = DataGrid1.Columns(10).Text
    
    
    ModificaLineas = 2 'Modificar
    LLamaLineas ModificaLineas, anc, "DataGrid1"
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR"
    PonerBotonCabecera False
    BloquearTxt Text2(16), False 'Campo Ampliacion Linea
    PonerFoco txtAux(11)
    Me.DataGrid1.Enabled = False

EModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub LLamaLineas(xModo As Byte, Optional alto As Single, Optional grid As String)
Dim jj As Integer
Dim b As Boolean

    Select Case grid
        Case "DataGrid1"
            DeseleccionaGrid Me.DataGrid1
            'PonerModo xModo + 1
    
            b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Lineas
    
            For jj = 0 To txtAux.Count - 1
                If jj = 4 Or (jj >= 6 And jj <= 11) Then
                    txtAux(jj).Height = DataGrid1.RowHeight
                    txtAux(jj).Top = alto
                    txtAux(jj).visible = b
                End If
            Next jj
            cmdAux.Top = alto
            cmdAux.visible = b
            
        Case "DataGrid2"
            DeseleccionaGrid Me.DataGrid2
            b = (xModo = 1)
             For jj = 0 To txtAux3.Count - 1
                txtAux3(jj).Height = DataGrid2.RowHeight
                txtAux3(jj).Top = alto
                txtAux3(jj).visible = b
            Next jj
    End Select
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Facturas (scafac)
' y los registros correspondientes de las tablas cab. albaranes (scafac1)
' y las lineas de la factura (slifac)
Dim Cad As String
'Dim vTipoMov As CTiposMov

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    'solo se puede modificar la factura si no esta contabilizada
    If FactContabilizada2(False) Then Exit Sub
    
    Cad = "Cabecera de Facturas." & vbCrLf
    Cad = Cad & "-------------------------------------      " & vbCrLf & vbCrLf
    Cad = Cad & "Va a eliminar la Factura:            "
    Cad = Cad & vbCrLf & "Tipo:  " & Text1(1).Text
    Cad = Cad & vbCrLf & "Nº Fact.:  " & Format(Text1(0).Text, "0000000")
    Cad = Cad & vbCrLf & "Fecha:  " & Format(Text1(2).Text, "dd/mm/yyyy")

    Cad = Cad & vbCrLf & vbCrLf & " ¿Desea Eliminarla? "

    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
'        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
'        NumPedElim = Data1.Recordset.Fields(1).Value
        CodTipoMov = Text1(1).Text
        
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
        
'        'Devolvemos contador, si no estamos actualizando
'        Set vTipoMov = New CTiposMov
'        vTipoMov.DevolverContador CodTipoMov, NumPedElim
'        Set vTipoMov = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
    
EEliminar:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminar Albaran", Err.Description
End Sub


'Private Sub cmdObserva_Click()
'    If Modo <> 2 And Modo <> 4 Then Exit Sub
'    If Me.FrameObserva.visible = False Then
'        Me.DataGrid1.visible = False
'        Me.FrameObserva.visible = True
'        Me.cmdObserva.Picture = frmppal.imgListComun.ListImages(18).Picture
''        CargarICO Me.cmdObserva, "volver.ico"
'        Me.cmdObserva.ToolTipText = "volver lineas albaran"
'        BloqueaText3
'    Else
'        Me.DataGrid1.visible = True
'        Me.FrameObserva.visible = False
'        Me.cmdObserva.Picture = frmppal.imgListComun.ListImages(41).Picture
''        CargarICO Me.cmdObserva, "message.ico"
'        Me.cmdObserva.ToolTipText = "ver observaciones albaran"
'    End If
'End Sub

Private Sub BloqueaText3()
Dim I As Byte
    'bloquear los Text3 que son las lineas de scafac1
    For I = 0 To 3
        BloquearTxt Text3(I), (Modo <> 4)
    Next I
    'If Me.FrameObserva.visible Then
    If True Then
        For I = 9 To 13
            BloquearTxt Text3(I), (Modo <> 4)
        Next I
        BloquearTxt Text3(16), (Modo <> 4)  'es de la 6ª observacion
    End If
    For I = 4 To 8
        BloquearTxt Text3(I), True
    Next I
    
    'datos venta TPV
    BloquearTxt Text3(14), True
    BloquearTxt Text3(15), True
    
    If vParamAplic.QUE_EMPRESA = 4 Then
        For I = 17 To 34
            BloquearTxt Text3(I), Modo <> 4
        Next
        For I = 0 To 3
            Me.chkCarga(I).Enabled = Modo = 4
        Next
        
        'chkCarga(0).Value
        
    End If
    
    
End Sub


Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim Cad As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
        DesBloqueoManual "scafac"
        PonerModo 2
        DataGrid2.Enabled = True
        If Not Data1.Recordset.EOF Then _
            Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount

    Else 'Se llama desde algún Prismatico de otro Form al Mantenimiento de Trabajadores
        If Data1.Recordset.EOF Then
            MsgBox "Ningún registro devuelto.", vbExclamation
            Exit Sub
        End If
        Cad = Data1.Recordset.Fields(0) & "|"
        Cad = Cad & Data1.Recordset.Fields(1) & "|"
        RaiseEvent DatoSeleccionado(Cad)
        Unload Me
    End If
End Sub


Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Combo1_LostFocus()
    Text1(1).Text = Mid(Combo1.List(Combo1.ListIndex), 1, 3)
End Sub

Private Sub DataGrid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Ayuda de Etiqueta de precio de salida de la Función de Precios
On Error Resume Next

    If Data2.Recordset.EOF Then Exit Sub
    If (Modo = 2) Or (Modo = 5 And ModificaLineas = 0) Then
        Me.DataGrid1.ToolTipText = ""
        If X > 7790 And X < 8170 Then
            Select Case DataGrid1.Columns(11).Value
                Case "P": Me.DataGrid1.ToolTipText = "P: Promoción"
                Case "E": Me.DataGrid1.ToolTipText = "E: Precio Especial"
                Case "T": Me.DataGrid1.ToolTipText = "T: Tarifa Artículo"
                Case "A": Me.DataGrid1.ToolTipText = "A: Precio Artículo"
                Case "M": Me.DataGrid1.ToolTipText = "M: Manual"
'                Case Else
'                    Me.DataGrid1.ToolTipText = ""
            End Select
'        Else
'            Me.DataGrid1.ToolTipText = ""
        End If
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo Error1

    If Not Data2.Recordset.EOF Then
        If ModificaLineas <> 1 Then Text2(16).Text = DBLet(Data2.Recordset.Fields!ampliaci)
    Else
        Text2(16).Text = ""
    End If

Error1:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub DataGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim I As Byte

    If Not data3.Recordset.EOF Then
        'Trabajador Albaran
        Text3(0).Text = data3.Recordset.Fields!CodTraba
        Text3_LostFocus (0)
        'Trabajador pedido
        Text3(1).Text = DBLet(data3.Recordset.Fields!codtrab1, "T")
        Text3_LostFocus (1)
        'Trab. Prepara Material
        Text3(2).Text = data3.Recordset.Fields!codtrab2
        Text3_LostFocus (2)
        Text3(3).Text = data3.Recordset.Fields!codenvio
        Text3_LostFocus (3)
        
        'oferta
        Text3(4).Text = DBLet(data3.Recordset.Fields!NumOfert, "N")
        If Text3(4).Text <> "0" Then
            FormateaCampo Text3(4)
        Else
            Text3(4).Text = ""
        End If
        Text3(5).Text = DBLet(data3.Recordset.Fields!fecofert, "F")
        'pedido
        Text3(6).Text = DBLet(data3.Recordset.Fields!numpedcl, "N")
        If Text3(6).Text <> "0" Then
            FormateaCampo Text3(6)
        Else
            Text3(6).Text = ""
        End If
        Text3(7).Text = DBLet(data3.Recordset.Fields!fecpedcl, "F")
        If Text3(7).Text <> "" Then FormateaCampo Text3(7)
        Text3(8).Text = DBLet(data3.Recordset.Fields!sementre, "N")
        If Text3(8).Text = "0" Then Text3(8).Text = ""
        'venta
        Text3(15).Text = DBLet(data3.Recordset.Fields!NumTermi, "N")
        Text3(14).Text = DBLet(data3.Recordset.Fields!NumVenta, "N")
        FormateaCampo Text3(14)
'        If Text3(14).Text = "0" Then Text3(14).Text = ""
'        If Text3(15).Text = "0" Then Text3(15).Text = ""
        
        'Observaciones
        Text3(9).Text = DBLet(data3.Recordset.Fields!observa1, "T")
        Text3(10).Text = DBLet(data3.Recordset.Fields!observa2, "T")
        Text3(11).Text = DBLet(data3.Recordset.Fields!observa3, "T")
        Text3(12).Text = DBLet(data3.Recordset.Fields!observa4, "T")
        Text3(13).Text = DBLet(data3.Recordset.Fields!observa5, "T")
        Text3(16).Text = DBLet(data3.Recordset.Fields!observa6, "T")
        
        If vParamAplic.QUE_EMPRESA = 4 Then
             Text3(17).Text = DBLet(data3.Recordset.Fields!FechaCarga, "F")
            Text3(18).Text = DBLet(data3.Recordset.Fields!Muestra, "T")
            Text3(19).Text = DBLet(data3.Recordset.Fields!Deposito, "T")
            Text3(20).Text = DBLet(data3.Recordset.Fields!TransAcidez, "T")
            Text3(21).Text = DBLet(data3.Recordset.Fields!TransBruto, "T")
            Text3(22).Text = DBLet(data3.Recordset.Fields!TransTara, "T")
            Text3(23).Text = DBLet(data3.Recordset.Fields!TransEmpresa, "T")
            Text3(24).Text = DBLet(data3.Recordset.Fields!TransMatricula, "T")
            Text3(25).Text = DBLet(data3.Recordset.Fields!TransMatRemolque, "T")
            
            '
            '
           
            Text3(26).Text = DBLet(data3.Recordset.Fields!TransNumBocas, "T")
            Text3(27).Text = DBLet(data3.Recordset.Fields!TransConductor, "F")
            Text3(28).Text = DBLet(data3.Recordset.Fields!TransCondDNI, "T")
            Text3(29).Text = DBLet(data3.Recordset.Fields!TransDestino, "T")
            Text3(30).Text = DBLet(data3.Recordset.Fields!TransMercancia, "T")
            Text3(31).Text = DBLet(data3.Recordset.Fields!TransLacradasCoop, "T")
            Text3(32).Text = DBLet(data3.Recordset.Fields!TransLacradasCompr, "T")
            Text3(33).Text = DBLet(data3.Recordset.Fields!TransObsPrecintos, "T")
            Text3(34).Text = ""
            If Not IsNull(data3.Recordset.Fields!hora) Then Text3(34).Text = Format(data3.Recordset.Fields!hora, "hh:mm")
            
            Me.chkCarga(0).Value = Abs(Val(DBLet(data3.Recordset!TransTicketBas, "N")))
            Me.chkCarga(1).Value = Abs(Val(DBLet(data3.Recordset!TransCMR, "N")))
            Me.chkCarga(2).Value = Abs(Val(DBLet(data3.Recordset!TransCertLim, "N")))
            Me.chkCarga(3).Value = Abs(Val(DBLet(data3.Recordset!TransOtros, "N")))
            
            
            
            Text3_LostFocus 17
            Text3_LostFocus 21
            Text3_LostFocus 22
            Text3_LostFocus 31
            Text3_LostFocus 32
            
        End If
        
        'Datos de la tabla slipre
        CargaGrid DataGrid1, Data2, True
    Else
        For I = 0 To Text3.Count - 1
            Text3(I).Text = ""
        Next I
        For I = 0 To 3
            Text2(I).Text = ""
        Next I
        
        If vParamAplic.QUE_EMPRESA = 4 Then
            For I = 0 To 3
                Me.chkCarga(I).Value = 0
            Next I
        End If
        
        'Datos de la tabla slipre
        CargaGrid DataGrid1, Data2, False
    End If
    
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    
    'Viene de DblClick en frmAlmMovimArticulos y carga el form con los valores
    If hcoCodMovim <> "" And Not Data1.Recordset.EOF Then PonerCadenaBusqueda
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmppal.Icon
    
     'Icono de busqueda
    For Kcampo = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(Kcampo).Picture = frmppal.imgListComun.ListImages(19).Picture
    Next Kcampo

    ' ICONITOS DE LA BARRA
    btnPrimero = 21
    With Me.Toolbar1
        .ImageList = frmppal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        .Buttons(5).Image = 4   'Modificar
        .Buttons(6).Image = 5   'Borrar
        .Buttons(9).Image = 10 'Mto Lineas Ofertas
        .Buttons(10).Image = 16 'Imprimir Pedido
        .Buttons(11).Image = 40 'Imprimir packing list
        .Buttons(18).Image = 15  'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With
    'Me.Toolbar1.Buttons(11).visible = vEmpresa.codempre = EmpresaAVAB
    Me.Toolbar1.Buttons(11).visible = vParamAplic.EsAVAB  'Voy a dar el PACKING
    
    
    
    
    Me.SSTab2.TabVisible(2) = vParamAplic.QUE_EMPRESA = 4
    Me.SSTab2.Tab = 0
    
    LimpiarCampos   'Limpia los campos TextBox
    CargaCombo
    Label1(46).Caption = "Observaciones packing list"
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
  
        Text3(16).visible = True
       
    Else
        'MORALES
        Text1(6).MaxLength = 15
        Text1(6).Width = 1590
        Text1(8).MaxLength = 35
        Text1(8).Height = Text1(6).Height


        Text3(16).visible = False   'Para morales NO dejo ver el observa6 ... de momento
        Label1(46).Caption = "Observaciones auxiliares"
    End If
    
    
    
    'cargar icono de observaciones de los albaranes de factura
    'Me.cmdObserva.Picture = frmppal.imgListComun.ListImages(41).Picture
'    CargarICO Me.cmdObserva, "message.ico"
    'Me.FrameObserva.visible = False
    'Me.cmdObserva.ToolTipText = "ver observaciones albaran"
    
    VieneDeBuscar = False
    
    'Comprobar si es Departamento o Direccion
    If vParamAplic.Departamento Then
        Me.Label1(1).Caption = "Dpto."
    Else
        Me.Label1(1).Caption = "Direc."
    End If
        
        
    Me.Label1(45).visible = vParamAplic.ctaAportacion <> ""
    Text1(45).visible = vParamAplic.ctaAportacion <> ""
        
        
    '## A mano
    NombreTabla = "scafac"
    NomTablaLineas = "slifac" 'Tabla lineas de Facturacion
    Ordenacion = " ORDER BY scafac.codtipom, scafac.numfactu, scafac.fecfactu "
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    CadenaConsulta = "Select * from " & NombreTabla
    If hcoCodMovim <> "" Then
        'Se llama desde Dobleclick en frmAlmMovimArticulos
        'como tenemos aqui el nº de albaran, buscar a que factura corresponde
        'en la scafac1
        CadenaConsulta = CadenaConsulta & ObtenerSelFactura
'        CadenaConsulta = CadenaConsulta & " WHERE codtipom='" & hcoCodTipoM & "' AND numalbar= " & hcoCodMovim
    Else
        CadenaConsulta = CadenaConsulta & " where numfactu=-1"
    End If
    
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
        
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
    
    If hcoCodMovim = "" Then
        If DatosADevolverBusqueda = "" Then
            PonerModo 0
        Else
            BotonBuscar
        End If
'        CargaGrid DataGrid1, Data2, False
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
        PrimeraVez = False
    Else
        PonerModo 0
    End If
End Sub


Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.Check1.Value = 0
    Me.Combo1.ListIndex = -1
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
'Form Agentes
Dim Indice As Byte
    Indice = 14
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod agente
    FormateaCampo Text1(Indice)
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom agente
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim cadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        If EsCabecera Then 'Llama desde VerTodos del Form
            cadB = ""
            Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 1)
            cadB = Aux
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 2)
            cadB = cadB & " and " & Aux
            Aux = ValorDevueltoFormGrid(Text1(2), CadenaDevuelta, 3)
            cadB = cadB & " and " & Aux
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        Else 'Llama desde Prismatico Direcciones/Departamentos
            Text1(12).Text = Format(RecuperaValor(CadenaDevuelta, 1), "000")
            Text1(13).Text = RecuperaValor(CadenaDevuelta, 2)
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Clientes
    Text1(4).Text = RecuperaValor(CadenaSeleccion, 1)  'Cod Clien
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


Private Sub frmFE_DatoSeleccionado(CadenaSeleccion As String)
'Formas de Envio
Dim Indice As Byte
    Indice = 29
    Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Forma Envio
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Forma Envio
End Sub

Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Formas de Pago
Dim Indice As Byte
    Indice = 15
    Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Forma Pago
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Forma Pago
End Sub


Private Sub frmO_DatoSeleccionado(Datos As String)
    Text1(46).Text = Datos
End Sub

Private Sub frmP_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(9).Text = RecuperaValor(CadenaSeleccion, 1)
    txtAux(10).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Trabajadores
Dim Indice As Byte
    Indice = Val(Me.imgBuscar(3).Tag)
    Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000") 'Cod Trabajador
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Trabajador
End Sub


Private Sub imgBuscar_Click(Index As Integer)
Dim Indice As Byte

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cod. Cliente
            PonerFoco Text1(4)
            Set frmC = New frmFacClientes
            frmC.DatosADevolverBusqueda = "0"
            frmC.Show vbModal
            Set frmC = Nothing
            Indice = 5
            PonerFoco Text1(Indice)
            
        Case 1 'NIF para cliente de Varios
            Set frmCV = New frmFacClientesV
            frmCV.DatosADevolverBusqueda = "0"
            frmCV.Show vbModal
            Set frmCV = Nothing
            Indice = 6
            PonerFoco Text1(Indice)
            
        Case 2 'Cod. Postal
            Set frmCP = New frmCPostal
            frmCP.DatosADevolverBusqueda = "0"
            frmCP.Show vbModal
            Set frmCP = Nothing
            Indice = 9
            VieneDeBuscar = True
            PonerFoco Text1(Indice)
        
        Case 3 'Cod. Direc.
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
             PonerFoco Text1(Indice)
             
        Case 4 'Agente
            Indice = 14
            PonerFoco Text1(Indice)
            Set frmA = New frmFacAgentesCom
            frmA.DatosADevolverBusqueda = "0"
            frmA.Show vbModal
            Set frmA = Nothing
            
         Case 5 'Forma de Pago
            Indice = 15
            PonerFoco Text1(Indice)
            Set frmFP = New frmFacFormasPago
            frmFP.DatosADevolverBusqueda = "0"
            frmFP.Show vbModal
            Set frmFP = Nothing
            
        Case 6, 7, 8 'Realizada Por Trabajador (Pedido, Albaran, Preparador Material
            Indice = Index - 6
            Me.imgBuscar(3).Tag = Indice
            Set frmT = New frmAdmTrabajadores
            frmT.DatosADevolverBusqueda = "0"
            frmT.Show vbModal
            Set frmT = Nothing
            PonerFoco Text3(Indice)
       
        Case 9 'Cod Envio
            Indice = 3
            PonerFoco Text3(Indice)
            Set frmFE = New frmFacFormasEnvio
            frmFE.DatosADevolverBusqueda = "0"
            frmFE.Show vbModal
            Set frmFE = Nothing
            PonerFoco Text3(Indice)
        Case 10
            Indice = 0
            If Text1(46).Text <> "" Then
                BuscaChekc = "Ya tiene introducidas " & Label1(46).Caption & vbCrLf
                BuscaChekc = BuscaChekc & "Serán reemplazadas.      ¿Continuar?"
                If MsgBox(BuscaChekc, vbQuestion + vbYesNo) = vbNo Then Indice = 1
                BuscaChekc = ""
            End If
            If Indice = 0 Then
                Set frmO = New frmFacCopiarObservaciones2
                frmO.PackingList = True
                frmO.IdCliente = CLng(Text1(4).Text)
                frmO.Show vbModal
                Set frmO = Nothing
            End If
    End Select
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub mnBuscar_Click()
    Me.SSTab2.Tab = 0
    BotonBuscar
End Sub


Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar lineas de Pedido
'         BotonEliminarLinea
        'INSERTAR NUEVA LINEA DE FACTURA
        If vUsu.Nivel < 1 Then ModificarAñadirLineasNUEVO True
    Else   'Eliminar Pedido
         BotonEliminar
    End If
End Sub


Private Sub mnImprimir_Click()
'Imprimir Factura
    
    If Data1.Recordset.EOF Then Exit Sub
    
    If Data1.Recordset!Codtipom = "FTI" Then 'ticket de venta del TPV
        BotonImprimirTicket
    Else

        If CInt(DBLet(data3.Recordset!NumTermi, "N")) > 0 Then
            'Es factura del TPV
            BotonImprimir 63
        Else
            'Impresion normal
            BotonImprimir (53) '53: Informe de Facturas
        End If

    End If
End Sub


Private Sub mnLineas_Click()

    'If Me.FrameObserva.visible Then cmdObserva_Click
    BotonMtoLineas 1, "Facturas"
End Sub


Private Sub mnModificar_Click()
    If Modo = 5 Then 'Modificar lineas
        'bloquea la tabla cabecera de factura: scafac
        If BLOQUEADesdeFormulario(Me) Then
            'bloquear la tabla cabecera de albaranes de la factura: scafac1
            If BloqueaAlbxFac Then
                'ANTES
                'If BloqueaLineasFac Then BotonModificarLinea
                'AHORA
                ModificarAñadirLineasNUEVO False
                
            End If
        End If
         
    Else   'Modificar Pedido
        'bloquea la tabla cabecera de factura: scafac
        If BLOQUEADesdeFormulario(Me) Then
            'bloquear la tabla cabecera de albaranes de la factura: scafac1
            If BloqueaAlbxFac Then BotonModificar
        End If
    End If
End Sub


Private Function BloqueaAlbxFac() As Boolean
'bloquea todos los albaranes de la factura
Dim SQL As String

    On Error GoTo EBloqueaAlb
    
    BloqueaAlbxFac = False
    'bloquear cabecera albaranes x factura
    SQL = "select * FROM scafac1 "
    SQL = SQL & ObtenerWhereCP(True) & " FOR UPDATE"
    conn.Execute SQL, , adCmdText
    BloqueaAlbxFac = True

EBloqueaAlb:
    If Err.Number <> 0 Then BloqueaAlbxFac = False
End Function


Private Function BloqueaLineasFac() As Boolean
'bloquea todas las lineas de la factura
Dim SQL As String

    On Error GoTo EBloqueaLin

    BloqueaLineasFac = False
    'bloquear cabecera albaranes x factura
    SQL = "select * FROM slifac "
    SQL = SQL & ObtenerWhereCP(True) & " FOR UPDATE"
    conn.Execute SQL, , adCmdText
    BloqueaLineasFac = True

EBloqueaLin:
    If Err.Number <> 0 Then BloqueaLineasFac = False
End Function


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

'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    Kcampo = Index
    If Index = 9 Then HaCambiadoCP = False 'CPostal
    If Index = 1 And Modo = 1 Then
        SendKeys "{tab}"
        Exit Sub
    End If
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
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
        
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
          
    'Si queremos hacer algo ..
    Select Case Index
            'NUEVO ENERO 2012
        Case 0
                '0000000  numero factura
                If Text1(Index).Text = "" Then Exit Sub
                
                NumRegElim = InStr(1, Text1(Index).Text, ".")
                If NumRegElim > 0 Then
                    Devuelve = Mid(Text1(Index).Text, NumRegElim + 1)
                    Text1(Index).Text = Mid(Text1(Index).Text, 1, NumRegElim - 1)
                    NumRegElim = Len(Text1(Index).Text) + Len(Devuelve)
                    If NumRegElim < 7 Then
                        NumRegElim = 7 - NumRegElim
                        Text1(Index).Text = Text1(Index).Text & String(NumRegElim, "0") & Devuelve
                    End If
                    
                End If
    
    
        Case 2 'Fecha factura
                If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
                
        Case 3, 27, 28 'Cod Vendedor
'                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "straba", "nomtraba", "codtraba")

        Case 4 'Cod. Cliente
            If Modo = 1 Then 'Modo=1 Busqueda
                '-- Laura 12/01/2007
                'Text1(5).Text = PonerNombreDeCod(Text1(Index), conAri, NombreTabla, "nomclien")
                Text1(5).Text = PonerNombreDeCod(Text1(Index), conAri, "sclien", "nomclien")
                '--
            Else
                PonerDatosCliente (Text1(Index).Text)
            End If
        
        Case 6 'NIF
            If Not EsDeVarios Then Exit Sub
            If Modo = 4 Then 'Modificar
                'si no se ha modificado el nif del cliente no hacer nada
                If Text1(6).Text = Data1.Recordset!nifClien Then Exit Sub
            End If
            PonerDatosClienteVario (Text1(Index).Text)
        
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
            If PonerFormatoEntero(Text1(Index)) Then
                'Comprobar que el cliente seleccionada tiene esa direccion
                If PonerDptoEnCliente Then
                    'Comprobar que el cliente tiene mantenimientos en esa direc/dpto
                    Devuelve = DevuelveDesdeBDNew(conAri, "scaman", "nummante", "codclien", Text1(4).Text, "N", , "coddirec", Text1(12).Text, "N")
                    If Devuelve <> "" And Text1(Index).Locked = False Then
                        Devuelve = "El cliente tiene Mantenimientos."
                        MsgBox Devuelve, vbInformation
                    End If
                Else
                    PonerFoco Text1(Index)
                End If
            Else
                Text1(Index + 1).Text = ""
            End If
            
        Case 14 'Cod. Agente
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sagent", "nomagent")
            Else
                Text2(Index).Text = ""
            End If
        
        Case 15 'Forma de Pago
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sforpa", "nomforpa")
            Else
                Text2(Index).Text = ""
            End If
            
        Case 16, 17 'Descuentos
            If PonerFormatoDecimal(Text1(Index), 4) Then   'Tipo 4: Decimal(4,2)
                If Modo = 4 Then CalcularDatosFactura
            End If
            
        Case 18 To 21 'banco, sucursal
            PonerFormatoEntero Text1(Index)
        Case 29 'Cod envio
'            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "senvio", "nomenvio")
    End Select
End Sub


Private Sub HacerBusqueda()
Dim cadB As String
Dim cadAux As String
    
    '--- Laura 12/01/2007
    cadAux = Text1(5).Text
    If Text1(4).Text <> "" Then Text1(5).Text = ""
    '---
    
    Text1(3).Text = ""
    If Combo1.ListIndex >= 0 Then Text1(3).Text = Mid(Trim(Combo1.List(Combo1.ListIndex)), 1, 3)
        
    
        
    
    
    cadB = ObtenerBusqueda(Me, False, BuscaChekc)
    
    'Si no es del B NO puede ver el almacen
    If Not vUsu.TrabajadorB Then
        If cadB <> "" Then cadB = cadB & " AND"
        cadB = cadB & "  scafac.codtipom <> 'FAZ'"
    End If
    
    '--- Laura 12/01/2007
    Text1(5).Text = cadAux
    '---
    
    If chkVistaPrevia = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then
        'Se muestran en el mismo form
'        cadB = cadB & " and scafac.codtipom='" & CodTipoMov & "'" 'Solo seleccionamos los del Movimiento, aqui los FAV
        CadenaConsulta = "select scafac.* from " & NombreTabla & " INNER JOIN scafac1 ON scafac.codtipom=scafac1.codtipom AND scafac.numfactu=scafac1.numfactu AND scafac.fecfactu=scafac1.fecfactu "
        CadenaConsulta = CadenaConsulta & " WHERE " & cadB & " GROUP BY scafac.codtipom,scafac.numfactu,scafac.fecfactu " & Ordenacion
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
        Cad = Cad & ParaGrid(Text1(1), 10, "Tipo Fac.")
        Cad = Cad & ParaGrid(Text1(0), 15, "Nº Factura")
        Cad = Cad & ParaGrid(Text1(2), 15, "Fecha Fac.")
        Cad = Cad & ParaGrid(Text1(4), 10, "Cliente")
        Cad = Cad & ParaGrid(Text1(5), 50, "Nombre Cliente")
        Tabla = NombreTabla & " INNER JOIN scafac1 ON scafac.codtipom=scafac1.codtipom AND scafac.numfactu=scafac1.numfactu AND scafac.fecfactu=scafac1.fecfactu "
        'CadenaConsulta = "select scafac.* from " & NombreTabla & " INNER JOIN scafac1 ON scafac.codtipom=scafac1.codtipom AND scafac.numfactu=scafac1.numfactu AND scafac.fecfactu=scafac1.fecfactu "
        'CadenaConsulta = CadenaConsulta & " WHERE " & cadB & " GROUP BY scafac.codtipom,scafac.numfactu,scafac.fecfactu " & Ordenacion
        
        Titulo = "Facturas"
        Devuelve = "0|1|2|"
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
        frmB.vConexionGrid = conAri  'Conexión a BD: Ariges
        If Not EsCabecera Then frmB.Label1.FontSize = 11
'        frmB.vBuscaPrevia = chkVistaPrevia
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        If EsCabecera Then
            PonerCadenaBusqueda
            Text1(0).Text = Format(Text1(0).Text, "0000000")
        End If
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
            PonerFoco Text1(Kcampo)
'            Text1(0).BackColor = vbYellow
        End If
        Exit Sub
    Else
        Data1.Recordset.MoveFirst
        PonerModo 2
        LLamaLineas Modo, 0, "DataGrid2"
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
'Carga el grid de los AlbaranesxFactura, es decir, la tabla scafac1 de la factura seleccionada
Dim b As Boolean
Dim b2 As Boolean

    On Error GoTo EPonerLineas

    Screen.MousePointer = vbHourglass
    
    'Datos de la tabla albaranes x factura: scafac1
    CargaGrid DataGrid2, data3, True
    
    'Comprobar si el albaran de la factura viene de una venta de ticket del TPV
    b = False
    b2 = False
    If Not data3.Recordset.EOF Then
        If Not IsNull(data3.Recordset!NumVenta) Then
            b = True
            If data3.Recordset!Codtipom = "FAV" And data3.Recordset!Codtipoa <> "FTI" Then b2 = True
        End If
    End If
    
    'Visualizar los campos de Oferta y Pedido si es una Factura q no es de venta TPV
    'o visulaizar numventa, numtermi si es una Factura de venta del TPV
    Label1(6).Caption = "Nº Pedido"
    Label1(18).Caption = "Fecha Pedido"
    If b Then
        If b2 Then
            Label1(6).Caption = "Nº Ticket"
            Label1(18).Caption = "Fecha Ticket"
        End If
        Label1(40).Caption = "Nº Terminal"
        Label1(22).Caption = "Nº Venta"
    Else
        Label1(40).Caption = "Nª Oferta"
        Label1(22).Caption = "Fecha Oferta"
    End If
    'sem. entrega
    Label1(2).visible = Not (b And b2)
    Text3(8).visible = Not (b And b2)
    'OFERTA
    Text3(4).visible = Not b
    Text3(5).visible = Not b
    'VENTA
    Text3(14).visible = b
    Text3(15).visible = b
    
    
    'Poner la referencia del cliente
    If Not data3.Recordset.EOF Then Text1(3).Text = DBLet(data3.Recordset.Fields!referenc, "T")
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
EPonerLineas:
    MuestraError Err.Number, "PonerCamposLineas"
    PonerModo 2
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
Dim BrutoFac As Single

    On Error Resume Next

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    
    'Poner la base imponible (impbruto - dtoppago - dtognral
    BrutoFac = CSng(Text1(22).Text) - CSng(Text1(23).Text) - CSng(Text1(24).Text)
    Text1(25).Text = Format(BrutoFac, FormatoImporte)
    
    FormatoDatosTotales
    
    'poner descripcion campos
    Modo = 4
    Text1_LostFocus (12) 'direc./dpto
    Text1_LostFocus (14) 'agente
    Text1_LostFocus (15) 'forma de pago
    Modo = 2
    
    PonerCamposLineas 'Pone los datos de las tablas de lineas de Ofertas
    
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
    '## No tiene el boton modificar y no utiliza la funcion general
    ActualizarToolbar Modo, Kmodo
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    b = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
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
    'si estamos en modificar bloquea las compos que son clave primaria
    BloquearText1 Me, Modo
    If Modo = 4 Then
        'MOdificando olo dejo modifcar el codfra a administrdores
        Text1(4).Locked = vUsu.Nivel > 1
        
    End If
    Me.Check1.Enabled = (Modo = 1)
    
    b = (Modo <> 1)
    'Campos Nº Factura bloqueado y en azul
    BloquearTxt Text1(0), b, True
    BloquearTxt Text1(3), b 'referencia
    
    'Importes siempre bloqueados, excepto para busquedas. ivas y aportacion tb bloqueado
    For I = 22 To 45
        BloquearTxt Text1(I), (Modo <> 1)
    Next I
    'Campo B.Imp y Imp. IVA siempre en azul
    BloquearTxt Text1(25), True
    Text1(25).BackColor = &HFFFFC0
    
    If Modo <> 1 Then
        Text1(35).BackColor = &HFFFFC0
        Text1(36).BackColor = &HFFFFC0
        Text1(37).BackColor = &HFFFFC0
'    Text1(38).BackColor = &HC0C0FF    'Total factura
        Text1(38).BackColor = &HC0FFC0
    End If
    
    'bloquear los Text3 que son las lineas de scafac1
    BloqueaText3
    
    'Si no es modo lineas Boquear los TxtAux
    For I = 0 To txtAux.Count - 1
        BloquearTxt txtAux(I), (Modo <> 5)
    Next I
    BloquearTxt txtAux(8), True
    BloquearTxt txtAux(10), True
    
    'Si no es modo Busqueda Bloquear los TxtAux3 (son los txtaux de los albaranes de factura)
    For I = 0 To txtAux3.Count - 1
        BloquearTxt txtAux3(I), (Modo <> 1)
    Next I
    
    'ampliacion linea
    b = (Modo = 5) And Me.DataGrid1.visible
    'Modo Linea de Albaranes
    Me.Label1(35).visible = b
    Me.Text2(16).visible = b
    BloquearTxt Text2(16), (Modo <> 5) Or (Modo = 5 And ModificaLineas <> 1)

    Me.Combo1.visible = (Modo = 1)

    '---------------------------------------------
    b = (Modo <> 0 And Modo <> 2 And Modo <> 5)
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    
    For I = 0 To 5
        Me.imgBuscar(I).Enabled = b
    Next I
    For I = 6 To 10
        Me.imgBuscar(I).Enabled = b And (Modo <> 1)
    Next I
    
    Me.imgBuscar(1).visible = False
                    
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
'Comprobar que los datos de la cabecera son correctos antes de Insertar o Modificar
'la cabecera del Pedido
Dim b As Boolean

    On Error GoTo EDatosOK

    DatosOk = False
    
    ComprobarDatosTotales
    
    'comprobamos datos OK de la tabla scafac
    b = CompForm(Me, 1) 'Comprobar formato datos ok de la cabecera: opcion=1
    If Not b Then Exit Function
    
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLinea() As Boolean
Dim b As Boolean
Dim I As Byte

    On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    b = True

    For I = 0 To txtAux.Count - 1
        If I = 4 Or I = 6 Or I = 7 Then
            If txtAux(I).Text = "" Then
                MsgBox "El campo " & txtAux(I).Tag & " no puede ser nulo", vbExclamation
                b = False
                PonerFoco txtAux(I)
                Exit Function
            End If
        End If
    Next I
            
            
    'PRoveedor
    If txtAux(9).Text <> "" And txtAux(10).Text = "" Then
        MsgBox "Codigo proveedor incorrecto", vbExclamation
        PonerFoco txtAux(9)
        b = False
        Exit Function
    End If
            
    DatosOkLinea = b
    
EDatosOkLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub Text2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 16 And KeyCode = 40 Then 'campo Amliacion Linea y Flecha hacia abajo
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYdown KeyCode
    End If
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 16 And KeyAscii = 13 Then 'campo Amliacion Linea y ENTER
        PonerFocoBtn Me.cmdAceptar
    End If
End Sub


Private Sub Text3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text3_LostFocus(Index As Integer)
    
    Select Case Index
        Case 0, 1, 2 'trabajador
            Text2(Index).Text = PonerNombreDeCod(Text3(Index), conAri, "straba", "nomtraba", "codtraba", "Cod. Trabajador", "N")
        Case 3 'cod. envio
            Text2(Index).Text = PonerNombreDeCod(Text3(Index), conAri, "senvio", "nomenvio", "codenvio", "Cod. Envio", "N")
            If Screen.ActiveControl.TabIndex <> 27 Then PonerFocoBtn Me.cmdAceptar
            
        Case 13 'observa 5
            PonerFocoBtn Me.cmdAceptar
            
            
            
        Case 17
            
            If Text3(Index).Text <> "" Then
                PonerFormatoFecha Text3(Index)
            End If
        
        Case 21, 22, 31, 32
            If Text3(Index).Text <> "" Then
                If Index < 30 Then
                    If Not PonerFormatoDecimal(Text3(Index), 3) Then Text3(Index).Text = ""
                Else
                    If Not PonerFormatoEntero(Text3(Index)) Then Text3(Index).Text = ""
                End If
            End If
    End Select
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Buscar
            mnBuscar_Click
        Case 2  'Todos
            BotonVerTodos

        Case 5  'Modificar
            mnModificar_Click
        Case 6  'Borrar
            mnEliminar_Click
        Case 9  'Lineas
            mnLineas_Click
        Case 10 'Imprimir Albaran
            mnImprimir_Click
        Case 11
            PackingList
        Case 18    'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub ActualizarToolbar(Modo As Byte, Kmodo As Byte)
'Modo: Modo antiguo
'Kmodo: Modo que se va a poner

    If (Modo = 5) And (Kmodo <> 5) Then
        'El modo antigu era modificando las lineas
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
'        Toolbar1.Buttons(5).Image = 3
'        Toolbar1.Buttons(5).ToolTipText = "Nuevo Albaran"
        '-- Modificar
        Toolbar1.Buttons(5).Image = 4
        Toolbar1.Buttons(5).ToolTipText = "Modificar Factura"
        '-- eliminar
        Toolbar1.Buttons(6).Image = 5
        Toolbar1.Buttons(6).ToolTipText = "Eliminar Factura"
    End If
    If Kmodo = 5 Then
        'Ponemos nuevos dibujitos y tal y tal
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
'        Toolbar1.Buttons(5).Image = 12
'        Toolbar1.Buttons(5).ToolTipText = "Nueva linea"
        '-- Modificar
        Toolbar1.Buttons(5).Image = 47  '13
        Toolbar1.Buttons(5).ToolTipText = "CAMBIAR LINEA FACTURA"
        '-- eliminar  --->>  AHORA NUEVA linea
        Toolbar1.Buttons(6).Image = 12
        Toolbar1.Buttons(6).ToolTipText = "NUEVA linea factura"
        'Toolbar1.Buttons(6).Image = 14
        'Toolbar1.Buttons(6).ToolTipText = "Eliminar linea factura"
        
    End If
End Sub
    
    
Private Function ModificarLinea() As Boolean
'Modifica un registro en la tabla de lineas de Albaran: slialb
Dim SQL As String
Dim vWhere As String
Dim b As Boolean

    On Error GoTo EModificarLinea

    ModificarLinea = False
    If Data2.Recordset.EOF Then Exit Function
    
    vWhere = ObtenerWhereCP(True)
    vWhere = vWhere & " AND codtipoa='" & data3.Recordset.Fields!Codtipoa & "' "
    vWhere = vWhere & " AND numalbar=" & data3.Recordset.Fields!NumAlbar
    vWhere = vWhere & " AND numlinea=" & Data2.Recordset.Fields!numlinea
    
    If DatosOkLinea() Then
        SQL = "UPDATE slifac SET "
        SQL = SQL & " ampliaci=" & DBSet(Text2(16).Text, "T") & ", "
        SQL = SQL & "precioar = " & DBSet(txtAux(4).Text, "N") & ", "
        SQL = SQL & "dtoline1= " & DBSet(txtAux(6).Text, "N") & ", dtoline2= " & DBSet(txtAux(7).Text, "N") & ", "
        
        
        If vParamAplic.QUE_EMPRESA = 2 Then
            'MOIXENT, hay que ver el hectogrado
        
        Else
            SQL = SQL & "importel = " & DBSet(txtAux(8).Text, "N") & ", "
        End If
        
        SQL = SQL & "origpre='" & txtAux(5) & "',"
        'TRAZA
        SQL = SQL & " codprovex= " & DBSet(txtAux(9).Text, "N", "S")
        'Los palets tb dejo cambiarlos
        SQL = SQL & ", palets= " & DBSet(txtAux(11).Text, "N", "S")
        SQL = SQL & vWhere
    End If
    
    If SQL <> "" Then
        'actualizar la factura y vencimientos
        b = ModificarFactura(SQL)
        
        ModificarLinea = b
    End If
    
EModificarLinea:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Modificar Lineas Factura" & vbCrLf & Err.Description
        b = False
    End If
    ModificarLinea = b
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
    'Habilitar las opciones correctas del menu segun Modo
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
    DataGrid2.Enabled = Not b
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
Dim b As Boolean
Dim Opcion As Byte
Dim SQL As String

    On Error GoTo ECargaGrid

    b = DataGrid1.Enabled
    If vDataGrid.Name = "DataGrid1" Then
        Opcion = 1
    Else
        Opcion = 2
    End If
    SQL = MontaSQLCarga(enlaza, Opcion)
    CargaGridGnral vDataGrid, vData, SQL, PrimeraVez
    
    vDataGrid.RowHeight = 270
    
    CargaGrid2 vDataGrid, vData
    vDataGrid.ScrollBars = dbgAutomatic
    
     b = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2)
     vDataGrid.Enabled = Not b
    
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim tots As String
    
    On Error GoTo ECargaGrid

    Select Case vDataGrid.Name
        Case "DataGrid1" 'Cod. Almacen
            'SQL = "SELECT codtipom, numfactu, fecfactu, numalbar, numlinea,
            'codalmac, codartic, nomartic, ampliaci, cantidad, precioar, origpre, dtoline1, dtoline2, importel "
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux(0)|T|Alm.|500|;S|txtAux(1)|T|Artículo|1570|;S|txtAux(2)|T|Nombre Art.|3275|;"
            tots = tots & "N||||0|;"
            tots = tots & "S|txtAux(11)|T|Palets|600|;"
            tots = tots & "S|txtAux(3)|T|Cantidad|900|;S|txtAux(4)|T|Precio|1100|;S|txtAux(5)|T|OP|400|;S|txtAux(6)|T|Dto 1|600|;S|txtAux(7)|T|Dto 2|600|;S|txtAux(8)|T|Importe|1250|;"
            'TRAZA
            tots = tots & "S|txtAux(9)|T|Prov.|700|;S|cmdaux|B|||;S|txtAux(10)|T|Nom. prov.|1500|;"
            arregla tots, DataGrid1, Me
            DataGrid1.Columns(8).Alignment = dbgRight
            DataGrid1.Columns(11).Alignment = dbgRight
            DataGrid1.Columns(12).Alignment = dbgCenter
            DataGrid1.Columns(13).Alignment = dbgRight
            DataGrid1.Columns(14).Alignment = dbgRight
            DataGrid1.Columns(15).Alignment = dbgRight
                       
         Case "DataGrid2" 'albaranes x articulo
'             SQL = "SELECT codtipom,numfactu,fecfactu,codtipoa,numalbar, fechaalb,"
             'numpedcl,fecpedcl,sementre,numofert,fecofert, referenc, codenvio,codtraba, codtrab1, codtrab2,observa1,observa2,observa3,observa4,observa5,numtermi,numventa  "
            tots = "N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux3(0)|T|Tipo|600|;S|txtAux3(1)|T|Albaran|1100|;S|txtAux3(2)|T|Fecha|1200|;"
            tots = tots & "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
            tots = tots & "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
            tots = tots & "N||||0|;" 'observa6
            
            If vParamAplic.QUE_EMPRESA = 4 Then
               
                 tots = tots & String(21, "€") 'campos cooperativa
                 tots = Replace(tots, "€", "N||||0|;")
            End If
            arregla tots, DataGrid2, Me
                     
            DataGrid2_RowColChange 1, 1
    End Select
    
    vDataGrid.HoldFields
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then Stop: MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub TxtAux_Change(Index As Integer)
    If Index = 6 And ModificaLineas = 2 Then 'Precio y Modo Borrar Lineas
        txtAux(5).Text = "M"
    End If
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux(Index)
End Sub

Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtAux_LostFocus(Index As Integer)

    'Quitar espacios en blanco
    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub
    
    Select Case Index
        Case 4 'Precio
             'Tipo 2: Decimal(10,4)
             If txtAux(Index).Text <> "" Then PonerFormatoDecimal txtAux(Index), 2
            
        Case 6, 7 'Descuentos
            PonerFormatoDecimal txtAux(Index), 4 'Tipo 4: Decimal(4,2)
            If Index = 7 Then PonerFoco Me.Text2(16)
            
        Case 8 'Importe Linea
            PonerFormatoDecimal txtAux(Index), 3 'Tipo 3: Decimal(10,2)
        Case 9
              txtAux(9).Text = Trim(txtAux(9).Text)
              txtAux(10).Tag = ""
              If txtAux(9).Text <> "" Then
                    If Not IsNumeric(txtAux(9).Text) Then
                        MsgBox "Campo proveedor debe ser numérico", vbExclamation
                    Else
                        txtAux(10).Tag = DevuelveDesdeBD(conAri, "nomprove", "sprove", "codprove", txtAux(9).Text)
                        If txtAux(10).Tag = "" Then
                            MsgBox "No existe proveedor: " & txtAux(9).Text, vbExclamation
                            txtAux(9).Text = ""
                            PonerFoco txtAux(9)
                        End If
                    End If
                End If
                txtAux(10).Text = txtAux(10).Tag
                txtAux(10).Tag = ""
        Case 11
            If Not PonerFormatoEntero(txtAux(Index)) Then
                txtAux(Index).Text = ""
                PonerFoco txtAux(Index)
            End If
    End Select
    
    If (Index = 3 Or Index = 4 Or Index = 6 Or Index = 7) Then 'Cant., Precio, Dto1, Dto2
        If txtAux(1).Text = "" Then Exit Sub
        txtAux(8).Text = CalcularImporte(txtAux(3).Text, txtAux(4).Text, txtAux(6).Text, txtAux(7).Text, vParamAplic.TipoDtos)
        PonerFormatoDecimal txtAux(8), 1
    End If
End Sub


Private Sub BotonMtoLineas(numTab As Integer, Cad As String)
    Me.SSTab2.Tab = numTab
    If Me.DataGrid1.visible Then
        If Me.Data2.Recordset.RecordCount < 1 Then
            MsgBox "La factura no tiene lineas.", vbInformation
            Exit Sub
        End If
        TituloLinea = Cad
    End If
    
    If FactContabilizada2(False) Then
        TerminaBloquear
        Exit Sub
    End If
    
    
    If Not BloqueoManual("scafac", Text1(0).Text & Data1.Recordset!Codtipom & "|" & Text1(2).Text) Then Exit Sub
    
    
    ModificaLineas = 0
    PonerModo 5
    PonerBotonCabecera True
End Sub


Private Function Eliminar() As Boolean
Dim SQL As String, LEtra As String
Dim b As Boolean
Dim vTipoMov As CTiposMov
    
    On Error GoTo FinEliminar

    b = False
    If Data1.Recordset.EOF Then Exit Function
        
    conn.BeginTrans
        
    'Eliminar en las tablas de la Contabilidad
    '------------------------------------------
    LEtra = ObtenerLetraSerie(Data1.Recordset!Codtipom)
    
    If LEtra <> "" Then
        SQL = " numserie='" & LEtra & "' AND codfaccl=" & Data1.Recordset.Fields!NumFactu
        SQL = SQL & " AND anofaccl=" & Year(Data1.Recordset.Fields!Fecfactu)
        
        'Lineas
        ConnConta.Execute "Delete from linfact WHERE " & SQL
        
        'cabecera
        ConnConta.Execute "Delete from cabfact WHERE " & SQL
        
        'cobros
        SQL = " numserie='" & LEtra & "' AND codfaccl=" & Data1.Recordset.Fields!NumFactu
        SQL = SQL & " AND fecfaccl='" & Format(Data1.Recordset.Fields!Fecfactu, FormatoFecha) & "'"
        ConnConta.Execute "Delete from scobro WHERE " & SQL
        b = True
    Else
        b = False
    End If

    'Eliminar en tablas de factura de Ariges
    '------------------------------------------
    If b Then
        SQL = " " & ObtenerWhereCP(True)
    
        'Lineas de facturas (slifac)
        conn.Execute "Delete from slifac " & SQL
    
        'Lineas de cabeceras de albaranes de la factura
        conn.Execute "Delete from scafac1 " & SQL
        
        'Eliminar los vencimientos
        conn.Execute "Delete from svenci " & SQL
        
        'Cabecera de facturas (scafac)
        conn.Execute "Delete from " & NombreTabla & SQL
        
        'Decrementar contador si borramos la ult. factura
        Set vTipoMov = New CTiposMov
        vTipoMov.DevolverContador Data1.Recordset!Codtipom, Val(Text1(0).Text)
        Set vTipoMov = Nothing
    End If
    
    b = True
    
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Factura", Err.Description
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

    CargaGrid DataGrid2, data3, False
    CargaGrid DataGrid1, Data2, False
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = "(" & ObtenerWhereCP(False) & ")"
         If SituarDataMULTI(Data1, vWhere, Indicador) Then
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
        CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Function ObtenerWhereCP(conWhere As Boolean) As String
Dim SQL As String

    On Error Resume Next
    
    SQL = " codtipom= '" & Text1(1).Text & "' and numfactu= " & Val(Text1(0).Text) & " and fecfactu='" & Format(Text1(2).Text, FormatoFecha) & "' "
    If conWhere Then SQL = " WHERE " & SQL
    ObtenerWhereCP = SQL
    
    If Err.Number <> 0 Then MuestraError Err.Number, "Obteniendo cadena WHERE.", Err.Description
End Function


Private Function MontaSQLCarga(enlaza As Boolean, Opcion As Byte) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim SQL As String
    
    If Opcion = 1 Then
        SQL = "SELECT codtipom, numfactu, fecfactu, numalbar, numlinea, codalmac, codartic, nomartic, ampliaci, palets,cantidad, precioar, origpre, dtoline1, dtoline2, importel ,codprovex, nomprove"
        SQL = SQL & " FROM slifac left join sprove on codprovex=codprove " 'lineas de factura
    ElseIf Opcion = 2 Then
        SQL = "SELECT codtipom,numfactu,fecfactu,codtipoa,numalbar, fechaalb, numpedcl,fecpedcl,sementre,numofert,fecofert, referenc, codenvio,codtraba, codtrab1, codtrab2,observa1,observa2,observa3,observa4,observa5,numtermi,numventa,observa6  "
        If vParamAplic.QUE_EMPRESA = 4 Then
            SQL = SQL & ",FechaCarga ,Muestra ,Deposito ,TransEmpresa ,TransMatricula ,TransConductor ,TransCondDNI "
            SQL = SQL & ",TransNumBocas ,TransBruto ,TransTara ,TransObsPrecintos,TransMatRemolque ,TransMercancia "
            SQL = SQL & ",TransAcidez ,TransDestino ,TransLacradasCoop ,TransLacradasCompr ,TransTicketBas "
            SQL = SQL & ",TransCMR ,TransCertLim ,TransOtros,hora"
        
        End If
        SQL = SQL & " FROM scafac1 " 'cabeceras albaranes de la factura
    End If
    
    If enlaza Then
        SQL = SQL & " " & ObtenerWhereCP(True)
        If Opcion = 1 Then SQL = SQL & " AND numalbar=" & data3.Recordset.Fields!NumAlbar
    Else
        SQL = SQL & " WHERE numfactu = -1"
    End If
    SQL = SQL & " ORDER BY codtipom, numfactu, fecfactu,numalbar "
    If Opcion = 1 Then SQL = SQL & ", numlinea "
    MontaSQLCarga = SQL
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean


        
        Toolbar1.Buttons(2).Enabled = b
        
        b = (Modo = 2) Or (Modo = 5 And ModificaLineas = 0)
        
        'Modificar
        Toolbar1.Buttons(5).Enabled = b
        Me.mnModificar.Enabled = b
        
        '31 Agosto 2011. Sera btn para añadir NUEVA linea
        Toolbar1.Buttons(6).Enabled = b
        Me.mnEliminar.Enabled = b
                
        'Toolbar1.Buttons(6).Enabled = (Modo = 2)
        'Me.mnEliminar.Enabled = (Modo = 2)
            
        b = (Modo = 2)
        'Mantenimiento lineas
        Toolbar1.Buttons(9).Enabled = b
        Me.mnLineas.Enabled = b
        'Imprimir
        Toolbar1.Buttons(10).Enabled = b
        Me.mnImprimir.Enabled = b
        
        
        Toolbar1.Buttons(11).Enabled = b
        
        
        b = (Modo >= 3) Or Modo = 1
        'Buscar
        Toolbar1.Buttons(1).Enabled = Not b
        Me.mnBuscar.Enabled = Not b
        'Ver Todos
        Toolbar1.Buttons(2).Enabled = Not b
        Me.mnVerTodos.Enabled = Not b
End Sub



Private Sub PonerDatosCliente(CodClien As String, Optional nifClien As String)
Dim vCliente As CCliente
Dim Observaciones As String
Dim b As Boolean

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
                If Modo = 3 Then
                    b = True
                ElseIf Modo = 4 Then
                     If (Val(Text1(4).Text) <> Val(Data1.Recordset!CodClien)) Then b = True
                End If
                If b Then
                    LimpiarDatosCliente
                    Set vCliente = Nothing
                    Exit Sub
                End If
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
            
            'insertar
            If Modo = 3 Then Text1(15).Text = vCliente.ForPago

            Observaciones = DBLet(vCliente.Observaciones)
            If Observaciones <> "" Then
                MsgBox Observaciones, vbInformation, "Observaciones del cliente"
            End If
                
            'cuenta bancaria
            Text1(18).Text = vCliente.Banco
            FormateaCampo Text1(18)
            Text1(19).Text = vCliente.Sucursal
            FormateaCampo Text1(19)
            Text1(20).Text = vCliente.DigControl
            Text1(21).Text = vCliente.CuentaBan
            
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
    Text1(7).Text = DBLet(vCliente.TfnoClien, "T")
            
    If Not b Then PonerFoco Text1(6)
    Set vCliente = Nothing
End Sub



Private Sub LimpiarDatosCliente()
Dim I As Byte
    
    For I = 4 To 13
        Text1(I).Text = ""
    Next I
    If (Modo = 3 Or Modo = 4) Then PonerFoco Text1(4)
End Sub
    
    
Private Sub BotonImprimir(OpcionListado As Byte)
Dim cadFormula As String
Dim Cadparam As String
Dim NumParam As Byte
Dim Cadselect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim Devuelve As String
Dim ImprimeDirecto As Boolean

    If Text1(0).Text = "" Then
        MsgBox "Debe seleccionar una Factura para Imprimir.", vbInformation
        Exit Sub
    End If
    
    cadFormula = ""
    Cadparam = ""
    Cadselect = ""
    NumParam = 0
    
    '===================================================
    '============ PARAMETROS ===========================
    If (OpcionListado = 53) Then
        If Text1(1).Text = "FAZ" Then
            'Factura B
            indRPT = 30
        Else
            indRPT = 12 'Facturas Clientes
        End If
    Else
        '-----------------------------------------------
        indRPT = 18 'Facturas Clientes TPV
    End If
    If Not PonerParamRPT(indRPT, Cadparam, NumParam, nomDocu, ImprimeDirecto) Then Exit Sub
      
      
      
    'PUNTO VERDE
    '--------------------------------------------------------------------------
    If vParamAplic.ArtReciclado <> "" Then
        Cadparam = Cadparam & "PuntoVerde= """ & vParamAplic.ArtReciclado & """|"
        NumParam = NumParam + 1
    End If
      
    'Nombre fichero .rpt a Imprimir
    If Not ImprimeDirecto Then frmImprimir.NombreRPT = nomDocu
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Nº de Factura
    '---------------------------------------------------
    If Text1(0).Text <> "" Then
        'Cod Tipo Movimiento
        Devuelve = "{" & NombreTabla & ".codtipom}='" & Text1(1).Text & "'"
        If Not AnyadirAFormula(cadFormula, Devuelve) Then Exit Sub
        
        'Nº Factura
        Devuelve = "{" & NombreTabla & ".numfactu}=" & Val(Text1(0).Text)
        If Not AnyadirAFormula(cadFormula, Devuelve) Then Exit Sub
        
        Cadselect = cadFormula
        
        'Fecha Factura
        Devuelve = "{" & NombreTabla & ".fecfactu}= Date(" & Year(Text1(2).Text) & "," & Month(Text1(2).Text) & "," & Day(Text1(2).Text) & ")"
        If Not AnyadirAFormula(cadFormula, Devuelve) Then Exit Sub
        'Fecha Factura en cadSelect
        Devuelve = "{" & NombreTabla & ".fecfactu}= '" & Format(Text1(2).Text, FormatoFecha) & "'"
        If Not AnyadirAFormula(Cadselect, Devuelve) Then Exit Sub
    End If
   
    If Not HayRegParaInforme(NombreTabla, Cadselect) Then Exit Sub
     
     
     If ImprimeDirecto Then
        'Imrpime directo
        If MsgBox("Imprimir la factura?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        ImprimirDirectoFact Cadselect
     Else
         With frmImprimir
                .FormulaSeleccion = cadFormula
                .OtrosParametros = Cadparam
                .NumeroParametros = NumParam
                .SoloImprimir = False
                .EnvioEMail = False
                .Opcion = OpcionListado
                .Titulo = ""
                .Show vbModal
        End With
    End If
End Sub



Private Sub BotonImprimirTicket()
Dim MIPATH As String
Dim cadImpresion As String, SQL As String
Dim NomImpre As String
Dim NomImpTi As String
Dim bImpre As Boolean

    cadImpresion = "{scafac.codtipom}='" & Text1(1).Text & "' and {scafac.numfactu}=" & Text1(0).Text
    SQL = cadImpresion & " and {scafac.fecfactu}=" & DBSet(Text1(2).Text, "F")
    cadImpresion = cadImpresion & " and {scafac.fecfactu}=Date(" & Year(CDate(Text1(2).Text)) & "," & Month(CDate(Text1(2).Text)) & "," & Day(CDate(Text1(2).Text)) & ")"
    
    If Not HayRegParaInforme("scafac", SQL) Then Exit Sub
    
'    'Obtener que terminal es
'     'Terminal con el que trabajaremos, leemos el nombre del ordenador
'    SQL = ComputerName 'Nombre PC conectado por Terminal Server / local
'    SQL = DevuelveDesdeBDNew(conAri, "spatpvt", "numtermi", "nombrepc", SQL, "T")
'    If Not IsNumeric(SQL) Then
'        MsgBox "No se ha podido establecer la impresora de ticket." & vbCrLf & "Debe configurar primero los parámetros del TPV.", vbExclamation
'    Else
'        bImpre = True
'    End If
'
'    If bImpre Then
'         'Establecemos la impresora de ticket
'         NomImpTi = NombreImpresoraTicket(CInt(SQL))
'         If NomImpTi <> "" Then
'            If Printer.DeviceName <> NomImpTi Then
'                'guardamos la impresora que habia
'                NomImpre = Printer.DeviceName
'                'establecemos la de ticket
'                EstablecerImpresora NomImpTi
'            End If
'        End If
'    End If


    


    MIPATH = App.Path & "\Informes\"
'    cadImpresion = cadImpresion & " and {scafac.fecfactu}=Date(" & Year(RSVenta!fecventa) & "," & Month(RSVenta!fecventa) & "," & Day(RSVenta!fecventa) & ")"
    With frmVisReport
        .FormulaSeleccion = cadImpresion
        .SoloImprimir = False
        .OtrosParametros = ""
        .NumeroParametros = 0
        .MostrarTree = False
        .Informe = MIPATH & "rTPVTicket.rpt"
        .ConSubInforme = False
        .Opcion = 93
        .ExportarPDF = False
        .Show vbModal
   End With
   
'   If bImpre Then
'        'volver la impresora a la predeterminada
'        EstablecerImpresora NomImpre
'   End If
   
End Sub




Private Function ModificaAlbxFac() As Boolean
Dim SQL As String
Dim b As Boolean
    
    On Error GoTo EModificaAlb
    
    ModificaAlbxFac = False
    'comprobar datos OK de la scafac1
     b = CompForm(Me, 2) 'Comprobar formato datos ok de la cabecera alb: opcion=2
    If Not b Then Exit Function
    
    SQL = "UPDATE scafac1 SET codenvio=" & Text3(3).Text & ", "
    SQL = SQL & "codtraba=" & Text3(0).Text & ", "
    SQL = SQL & "codtrab1=" & DBSet(Text3(1).Text, "N", "S") & ", " 'Trab. pedido
    SQL = SQL & "codtrab2=" & Text3(2).Text 'Trab. Prep. Material
    'If Me.FrameObserva.visible Then
    If True Then
        SQL = SQL & ", observa1=" & DBSet(Text3(9).Text, "T")
        SQL = SQL & ", observa2=" & DBSet(Text3(10).Text, "T")
        SQL = SQL & ", observa3=" & DBSet(Text3(11).Text, "T")
        SQL = SQL & ", observa4=" & DBSet(Text3(12).Text, "T")
        SQL = SQL & ", observa5=" & DBSet(Text3(13).Text, "T")
        SQL = SQL & ", observa6=" & DBSet(Text3(16).Text, "T")
    End If
    
    If vParamAplic.QUE_EMPRESA = 4 Then
    
        SQL = SQL & ",FechaCarga  = " & DBSet(Text3(17).Text, "F", "S")
        SQL = SQL & ",Muestra  = " & DBSet(Text3(18).Text, "T")
        SQL = SQL & ",Deposito  = " & DBSet(Text3(19).Text, "T")
        SQL = SQL & ",TransAcidez  = " & DBSet(Text3(20).Text, "T")
        SQL = SQL & ",TransBruto  = " & DBSet(Text3(21).Text, "N", "S")
        SQL = SQL & ",TransTara  = " & DBSet(Text3(22).Text, "N", "S")
        SQL = SQL & ",TransEmpresa  = " & DBSet(Text3(23).Text, "T")
        SQL = SQL & ",TransMatricula  = " & DBSet(Text3(24).Text, "T")
        SQL = SQL & ",TransMatRemolque  = " & DBSet(Text3(25).Text, "T")
        SQL = SQL & ",TransNumBocas  = " & DBSet(Text3(26).Text, "T")
        SQL = SQL & ",TransConductor  = " & DBSet(Text3(27).Text, "T")
        SQL = SQL & ",TransCondDNI  = " & DBSet(Text3(28).Text, "T")
        SQL = SQL & ",TransDestino  = " & DBSet(Text3(29).Text, "T")
        SQL = SQL & ",TransMercancia  = " & DBSet(Text3(30).Text, "T")
        SQL = SQL & ",TransLacradasCoop  = " & DBSet(Text3(31).Text, "T")
        SQL = SQL & ",TransLacradasCompr  = " & DBSet(Text3(32).Text, "T")
        SQL = SQL & ",TransObsPrecintos  = " & DBSet(Text3(33).Text, "T")
        SQL = SQL & ",TransTicketBas  = " & Abs(chkCarga(0).Value)
        SQL = SQL & ",TransCMR  = " & Abs(chkCarga(1).Value)
        SQL = SQL & ",TransCertLim  = " & Abs(chkCarga(2).Value)
        SQL = SQL & ",TransOtros  = " & Abs(chkCarga(3).Value)

    
    End If
    
    
    
    SQL = SQL & ObtenerWhereCP(True)
    SQL = SQL & " AND codtipoa='" & data3.Recordset.Fields!Codtipoa & "' AND numalbar=" & data3.Recordset.Fields!NumAlbar
    conn.Execute SQL
    ModificaAlbxFac = True
    
EModificaAlb:
    If Err.Number <> 0 Then MuestraError Err.Number, "Modificar Albaranes de factura", Err.Description
End Function


Private Function ModificarFactura(Optional sqlLineas As String) As Boolean
'si se ha modificado la linea de slifac, añadir a la transaccion la modificación de la linea y recalcular
Dim bol As Boolean
Dim MenError As String
Dim SQL As String, LEtra As String
Dim vFactura As CFactura
Dim recalcular As Boolean

    On Error GoTo EModFact

    
    'Comprobar si hay que recalcular la factura
    recalcular = False
    If sqlLineas <> "" Then
        'comprobamos si se ha modificado la linea del albaran (precio y descuentos)
        recalcular = True
    ElseIf CInt(Data1.Recordset!codforpa) <> CInt(Text1(15).Text) Then
        'si se ha cambiado la forma de pago
        recalcular = True
    ElseIf CSng(Data1.Recordset!DtoPPago) <> CSng(DBSet(Text1(16).Text, "N")) Then
        'si se ha cambiado el dto ppago
        recalcular = True
    ElseIf CSng(Data1.Recordset!DtoGnral) <> CSng(DBSet(Text1(17).Text, "N")) Then
        'si se ha cambiado el descuento general
        recalcular = True
        
    'NO LLEVA BONIFICACIONES.  Junio 2010
    'ElseIf CInt(Data1.Recordset!CodClien) <> CInt(Text1(4).Text) Then
    '    'si se ha cambiado el cliente (bonificara o no)
    '    recalcular = True
    ElseIf CSng(Data1.Recordset!TotalFac) <> CSng(Text1(38).Text) Then
        recalcular = True
    End If
    
    
    bol = True
    conn.BeginTrans
    ConnConta.BeginTrans
    
    If recalcular Then
        If sqlLineas <> "" Then
            'actualizar el importe de la linea modificada
            MenError = "Modificando lineas de Factura."
            conn.Execute sqlLineas
        End If
        
        'recalcular las bases imponibles x IVA
        MenError = "Recalcular importes IVA"
    '    bol = RecalcularFactura
        bol = CalcularDatosFactura
        
    '    bol = True
    End If
    
    If bol Then
'        ComprobarDatosTotales
        
        'modificamos la scafac
        MenError = "Modificando cabecera de factura"
        bol = ModificaDesdeFormulario(Me, 1)
        
        If bol Then
            'Si es cliente de varios actualizar datos cliente en tabla:sclvar
            MenError = "Modificando datos cliente varios"
            bol = ActualizarClienteVarios(Text1(4).Text, Text1(6).Text)
        End If
        
        If bol Then
            MenError = "Modificando albaranes de factura"
            'modificar la tabla: scafac1
            bol = ModificaAlbxFac
            
            If bol And recalcular Then 'si se ha modificado la factura
                MenError = "Actualizando en Tesoreria"
                'borrar los vencimientos de ariges.svenci
                'y eliminar de tesoreria conta.scobros los registros de la factura(si existen en Tesoreria)
                
                'Eliminar los vencimientos
                '----------------------------------------
                SQL = ObtenerWhereCP(True)
                conn.Execute "Delete from svenci " & SQL
                
                'Eliminar de Tesoreria
                '----------------------------------------
'                SQL = ObtenerLetraSerie(Text1(1).Text)
'                SQL = "SELECT COUNT(*) FROM scobro WHERE numserie='" & SQL & "' and codfaccl=" & Text1(0).Text
'                SQL = SQL & " AND fecfaccl=" & DBSet(Text1(2).Text, "F")
'
'                If RegistrosAListar(SQL, conConta) Then
                    'antes de Eliminar en las tablas de la Contabilidad
                Set vFactura = New CFactura
                If vFactura.LeerDatosN(Text1(1).Text, Text1(0).Text, Text1(2).Text) Then
                Else
                  bol = False
                End If
              
                If bol Then
                    'Eliminar de la scobro
                    
                    'Eliminar de la scobro
                    If vParamAplic.ContabilidadNueva Then
                        SQL = " cobros WHERE numserie='" & vFactura.LetraSerie & "' AND numfactu=" & Data1.Recordset.Fields!NumFactu
                        SQL = SQL & " AND fecfactu='" & Format(Data1.Recordset.Fields!Fecfactu, FormatoFecha) & "'"
                    Else
                        SQL = " scobro WHERE numserie='" & vFactura.LetraSerie & "' AND codfaccl=" & Data1.Recordset.Fields!NumFactu
                        SQL = SQL & " AND fecfaccl='" & Format(Data1.Recordset.Fields!Fecfactu, FormatoFecha) & "'"
                    End If
                    ConnConta.Execute "Delete from " & SQL
                    
                    
                    
                    
                    
                    
                    
                    bol = True

                    'Volvemos a Insertar los Vencimientos de la Factura. Tabla: svenci
                    'Grabar en TESORERIA. Tabla de Contabilidad: sconta.scobros
                    If bol Then
                        vFactura.Agente = Text1(14).Text
                        bol = vFactura.InsertarEnTesoreria("", MenError)
                    End If
                End If
                Set vFactura = Nothing
            End If
'            End If
        End If
    End If

EModFact:
     If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        bol = False
    End If
    If bol Then
        conn.CommitTrans
        ConnConta.CommitTrans
        ModificarFactura = True
    Else
        conn.RollbackTrans
        ConnConta.RollbackTrans
        ModificarFactura = False
        MenError = "Actualizando Factura." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
    End If
End Function


Private Function CalcularDatosFactura() As Boolean
Dim I As Integer
Dim vFactu As CFactura
Dim FacOK As Boolean
Dim CambiaIVA As Boolean

    'Limpiar en el form los datos calculados de la factura
    'y volvemos a recalcular
    For I = 22 To 38
         Text1(I).Text = ""
    Next I
    
    
    Set vFactu = New CFactura
    vFactu.DtoPPago = CCur(ComprobarCero(Text1(16).Text))
    vFactu.DtoGnral = CCur(ComprobarCero(Text1(17).Text))
    vFactu.Cliente = Text1(4).Text
    vFactu.Codtipom = CStr(Me.Data1.Recordset!Codtipom)
    CambiaIVA = False
    If CDate(Text1(2).Text) < CDate("01/09/2012") Then CambiaIVA = True
    
    'If Modo = 5 And vFactu.Codtipom = "" Then vFactu.Codtipom = Text1(1).Text
    
    If vFactu.CalcularDatosFactura(ObtenerWhereCP(False), NombreTabla, NomTablaLineas, CambiaIVA) Then
        FacOK = True
        Text1(22).Text = vFactu.BrutoFac
        Text1(23).Text = vFactu.ImpPPago
        Text1(24).Text = vFactu.ImpGnral
        Text1(25).Text = vFactu.BaseImp
        Text1(26).Text = QuitarCero(vFactu.TipoIVA1)
        Text1(27).Text = QuitarCero(vFactu.TipoIVA2)
        Text1(28).Text = QuitarCero(vFactu.TipoIVA3)
        Text1(29).Text = vFactu.PorceIVA1
        Text1(30).Text = vFactu.PorceIVA2
        Text1(31).Text = vFactu.PorceIVA3
        Text1(32).Text = vFactu.BaseIVA1
        Text1(33).Text = vFactu.BaseIVA2
        Text1(34).Text = vFactu.BaseIVA3
        Text1(35).Text = vFactu.ImpIVA1
        Text1(36).Text = vFactu.ImpIVA2
        Text1(37).Text = vFactu.ImpIVA3
        Text1(38).Text = vFactu.TotalFac
        FormatoDatosTotales
    Else
        FacOK = False
        MuestraError Err.Number, "Calculando Totales", Err.Description
    End If
    Set vFactu = Nothing
    CalcularDatosFactura = FacOK
End Function


Private Sub FormatoDatosTotales()
Dim I As Byte

    For I = 22 To 25
        Text1(I).Text = QuitarCero(Text1(I).Text)
        Text1(I).Text = Format(Text1(I).Text, FormatoImporte)
    Next I
    
    'Desglose B.Imponible por IVA
    For I = 32 To 34
        If Text1(I).Text <> "" Then
             If CSng(Text1(I).Text) = 0 And Text1(I - 6).Text = "" Then
                Text1(I).Text = QuitarCero(Text1(I).Text)
                Text1(I - 3).Text = QuitarCero(Text1(I - 3).Text)
                Text1(I - 6).Text = QuitarCero(Text1(I - 6).Text)
                Text1(I + 3).Text = QuitarCero(Text1(I).Text)
            Else
                Text1(I).Text = Format(Text1(I).Text, FormatoImporte)
                Text1(I - 3) = Format(Text1(I - 3).Text, FormatoDescuento)
    '            Text3(i - 6) = Format(Text3(i - 6).Text, "000")
                Text1(I + 3).Text = Format(Text1(I + 3).Text, FormatoImporte)
            End If
        End If
    Next I
End Sub



Private Sub ComprobarDatosTotales()
Dim I As Byte

    For I = 22 To 25
        Text1(I).Text = ComprobarCero(Text1(I).Text)
    Next I
End Sub


Private Function FactContabilizada2(DatosCabecera As Boolean) As Boolean
Dim LEtra As String, numasien As String
    
    On Error GoTo EContab

    'comprabar que se puede modificar/eliminar la factura
    If Me.Check1.Value = 1 Then 'si esta contabilizada
        'comprobar en la contabilidad si esta contabilizada
        LEtra = ObtenerLetraSerie(Text1(1).Text)
        If LEtra <> "" Then
            If vParamAplic.ContabilidadNueva Then
                'Aunque en la nueva contabiliad SIEMPRE esta con apunte.
                numasien = DevuelveDesdeBDNew(conConta, "factcli", "numasien", "numserie", LEtra, "T", , "numfactu", Text1(0).Text, "N", "anofactu", Year(Text1(2).Text), "N")

            Else
                numasien = DevuelveDesdeBDNew(conConta, "cabfact", "numasien", "numserie", LEtra, "T", , "codfaccl", Text1(0).Text, "N", "anofaccl", Year(Text1(2).Text), "N")
            End If
            If Val(ComprobarCero(numasien)) <> 0 Then
                FactContabilizada2 = True
              
                LEtra = ""
                If Not DatosCabecera Then

                Else
                    If vUsu.Nivel <= 1 Then LEtra = "La factura esta contabilizada. ¿Continuar?"
                End If
                
                If LEtra = "" Then
                    'NO PUEDO SEGUIR
                    LEtra = "La factura esta contabilizada y no se puede modificar."
                    MsgBox LEtra, vbExclamation
                Else
                    If MsgBox(LEtra, vbQuestion + vbYesNo) = vbYes Then LEtra = ""
                End If
                If LEtra <> "" Then Exit Function
                FactContabilizada2 = False
            Else
                MsgBox "Factura contabilizada", vbExclamation
                'Agosto 2011
                'No se ha encontrado la factura Y si que esta contbilizada
                'De momento si no es super-usuario tapoco dejo continuar
                If (vUsu.Codigo Mod 1000) = 0 Then
                    FactContabilizada2 = False
                Else
                    FactContabilizada2 = True
                End If
            End If
        Else
'            MsgBox "Las factura de venta no tienen asignada una letra de serie", vbInformation
            FactContabilizada2 = True
            Exit Function
        End If
    Else
        FactContabilizada2 = False
    End If
    
EContab:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar Factura contabilizada", Err.Description
End Function


Private Sub TxtAux3_GotFocus(Index As Integer)
    ConseguirFoco txtAux3(Index), Modo
End Sub

Private Sub TxtAux3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index <> 0 And KeyCode <> 38 Then KEYdown KeyCode
End Sub

Private Sub TxtAux3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub TxtAux3_LostFocus(Index As Integer)
    If Not PerderFocoGnral(txtAux3(Index), Modo) Then Exit Sub
End Sub


Private Sub BloquearDatosCliente(bol As Boolean)
Dim I As Byte

    'bloquear/desbloquear campos de datos segun sea de varios o no
    If Modo <> 5 Then
        Me.imgBuscar(1).visible = bol
        Me.imgBuscar(1).Enabled = bol
        Me.imgBuscar(2).Enabled = bol
        
        
        BloquearTxt Text1(5), Not bol  'EL NOMBRE NUNCA DEJO
        For I = 5 To 11 'si no es de varios no se pueden modificar los datos
            'Nuevo para poder cambiar los datos
            'Diciembre 2009
            Debug.Print I & "  " & Text1(I).Tag
            'BloquearTxt Text1(i), Not bol
            BloquearTxt Text1(I), False
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



Private Function ObtenerSelFactura() As String
Dim Cad As String
Dim RS As ADODB.Recordset

    On Error Resume Next

    Cad = ""
    '******************************************************
    'laura: esto se puede comentar, ya no hay movimiento FTI en la smoval
    If hcoCodTipoM = "FTI" Then
        'no hay albaran directamente va a factura de ticket
        
        'ver si lo encontramos como factura: codtipom, numfactu,fecfactu
        Cad = "SELECT COUNT(*) FROM scafac "
        Cad = Cad & " WHERE codtipom='" & hcoCodTipoM & "' AND numfactu= " & hcoCodMovim & " AND fecfactu=" & DBSet(hcoFechaMov, "F")
        If RegistrosAListar(Cad) > 0 Then
            Cad = " WHERE codtipom='" & hcoCodTipoM & "' AND numfactu= " & hcoCodMovim & " AND fecfactu=" & DBSet(hcoFechaMov, "F")
        Else
            Cad = ""
        End If
    Else
        If hcoCodTipoM = "FAM" Then
            Cad = "  WHERE codtipom='" & hcoCodTipoM & "' AND numfactu= " & hcoCodMovim & " AND fecfactu=" & DBSet(hcoFechaMov, "F")
        End If
    End If
    '******************************************************
        
    If Cad = "" Then
        'En la smoval estaba e mov. de ALbaran
        Cad = "SELECT codtipom,numfactu,fecfactu FROM scafac1 "
        Cad = Cad & " WHERE codtipoa=" & DBSet(hcoCodTipoM, "T") & " AND numalbar=" & hcoCodMovim & " AND fechaalb=" & DBSet(hcoFechaMov, "F")
        
        Set RS = New ADODB.Recordset
        RS.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then 'where para la factura
            Cad = " WHERE codtipom='" & RS!Codtipom & "' AND numfactu= " & RS!NumFactu & " AND fecfactu=" & DBSet(RS!Fecfactu, "F")
        Else
            Cad = " WHERE numfactu=-1"
        End If
        RS.Close
        Set RS = Nothing
    End If
    ObtenerSelFactura = Cad
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
    vClien = Nothing
End Function


Private Sub CargaCombo()
Dim RS As ADODB.Recordset
Dim SQL As String
Dim I As Byte
    
    Combo1.Clear
    
    SQL = "SELECT codtipom,nomtipom FROM stipom WHERE codtipom LIKE 'F%'"
    If Not vUsu.TrabajadorB Then SQL = SQL & " AND codtipom <> 'FAZ'"
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        SQL = RS!nomtipom
        SQL = Replace(SQL, "Factura", "")
        Combo1.AddItem RS!Codtipom & "-" & SQL
        Combo1.ItemData(Combo1.NewIndex) = I
        I = I + 1
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
End Sub


Private Sub PackingList()
Dim cadFormula As String
Dim Cadparam As String
Dim NumParam As Byte
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim Devuelve As String
Dim ImprimeDirecto As Boolean

    If Me.Data1.Recordset Is Nothing Then Exit Sub
    If Me.Data1.Recordset.EOF Then Exit Sub
    
    If Me.Data1.Recordset!Codtipom <> "FAV" Then
        MsgBox "Solo facturas de venta", vbExclamation
        Exit Sub
    End If
    
    
    'Vamos a meter el lote y demas
    PonerDatosLote
    
    
    
    cadFormula = ""
    Cadparam = ""
    
    NumParam = 0
    

    If Not PonerParamRPT(34, Cadparam, NumParam, nomDocu, ImprimeDirecto) Then Exit Sub
      
      
    
    
    'Cod Tipo Movimiento
    Devuelve = "{slifac.codtipom}='" & Text1(1).Text & "'"
    If Not AnyadirAFormula(cadFormula, Devuelve) Then Exit Sub
    
    'Nº Factura
    Devuelve = "{slifac.numfactu}=" & Val(Text1(0).Text)
    If Not AnyadirAFormula(cadFormula, Devuelve) Then Exit Sub
    
    
    'Fecha Factura
    Devuelve = "{slifac.fecfactu}= Date(" & Year(Text1(2).Text) & "," & Month(Text1(2).Text) & "," & Day(Text1(2).Text) & ")"
    If Not AnyadirAFormula(cadFormula, Devuelve) Then Exit Sub


    Cadparam = Cadparam & "codusu=" & vUsu.Codigo & "|"
    NumParam = NumParam + 1
    
    With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = Cadparam
            .NumeroParametros = NumParam
            .SoloImprimir = False
            .EnvioEMail = False
            .NombreRPT = nomDocu
            .Titulo = "Packing List"
            .Opcion = 53
            .Show vbModal
    End With
    
    
End Sub



Private Sub PonerDatosLote()
Dim SQL As String
Dim RS As ADODB.Recordset
Dim Aux As String
Dim F As Date
Dim masDeUnLinea As String
Dim codartic As String
Dim DAV As String


'Septiembre 2013
'En lugar de unidades pondemos CAJAS
Dim Unicajas As Long
Dim uniTexto As String

    On Error GoTo EPonerlotes
    
    
    
    SQL = "DELETE FROM tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL
    
    If Not vParamAplic.EsAVAB Then Exit Sub
    
    Set RS = New ADODB.Recordset
    
    
    SQL = "select * from slifaclotes " & ObtenerWhereCP(True)
    RS.Open SQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    masDeUnLinea = "|"
    Aux = ""
    While Not RS.EOF
        Aux = "OK"
        If RS!linea > 1 Then masDeUnLinea = masDeUnLinea & Format(RS!NumAlbar, "0000") & Format(RS!numlinea, "000") & "|"
        RS.MoveNext
    Wend
    If masDeUnLinea = "|" Then masDeUnLinea = ""
    If Aux <> "" Then RS.MoveFirst
    Aux = ""
    While Not RS.EOF
        
        
    
        SQL = ObtenerWhereCP(False) & " AND numalbar=" & RS!NumAlbar & " and codtipoa='" & RS!Codtipoa & "' AND numlinea "
        SQL = DevuelveDesdeBD(conAri, "codartic", "slifac", SQL, CStr(RS!numlinea))

        
        If SQL = "" Then
            MsgBox "No se encuentra el articulo para el lote: " & RS!NUmlote, vbExclamation
        Else
            codartic = SQL
            
            BuscaChekc = ""
            If masDeUnLinea <> "" Then
                BuscaChekc = Format(RS!NumAlbar, "0000") & Format(RS!numlinea, "000") & "|"
                If InStr(1, masDeUnLinea, BuscaChekc) = 0 Then BuscaChekc = ""
            End If
            
            
            'Sept. Pondremos las cajas si tienen(deberia tener valor) unicajas
            'If BuscaChekc <> "" Then BuscaChekc = "(" & RS!Cantidad & ")"
            SQL = DevuelveDesdeBD(conAri, "unicajas", "sartic", "codartic", codartic)
            Unicajas = 1
            If Val(SQL) > 1 Then Unicajas = Val(SQL)
            If Unicajas < 2 Then
                If BuscaChekc <> "" Then BuscaChekc = "(" & RS!Cantidad & ")"
            Else
                If BuscaChekc <> "" Then
                    Unicajas = ((RS!Cantidad - 1) \ Unicajas) + 1
                    BuscaChekc = "[" & Unicajas & "]"
                End If
            End If
            
            'El lote sin la fecprod
            NumRegElim = InStr(RS!NUmlote, " ")
            If NumRegElim > 0 Then
                SQL = Mid(RS!NUmlote, 1, NumRegElim)
            Else
                SQL = Mid(RS!NUmlote, 5)
            End If
            BuscaChekc = SQL & BuscaChekc 'Aqui tendre ej: 9945(23) para el nº lote 9945 2011/10/21
            
            'Como es AVAB, para saber la fecha de produccion tendremos que irnos a ariges1 (morales)
          
            'La fecha de caducidad esta en la tabla de produccion (tambien estara en la de lotes
            'Con lo cual YA no es sumando 2 años a la de produccion. Habra que buscarla en la BD
            SQL = "numfactu = " & RS!NumFactu & " AND codtipom='" & RS!Codtipom & "' AND fecfactu = " & DBSet(RS!Fecfactu, "F")
            SQL = SQL & " AND numalbar=" & RS!NumAlbar & " and codtipoa='" & RS!Codtipoa & "' AND numlinea "
            
            

            SQL = DevuelveDesdeBD(conAri, "codartic", "slifac", SQL, CStr(RS!numlinea))
            If SQL = "" Then
                MsgBox "No se encuentra el articulo para el lote: " & RS!NUmlote, vbExclamation
            Else
            

                
                DAV = SQL  'codartic
            
                SQL = " codartic = " & DBSet(DAV, "T") & " AND numlote=" & DBSet(RS!NUmlote, "T") & " AND 1"
                SQL = DevuelveDesdeBD(conAri, "numalbar", "ariges" & EmprMorales & ".spartidas", SQL, "1")
            End If
            
            'En numalbar tendre el NUMero de produccion
            If SQL <> "" Then
                DAV = Mid(SQL, 1, 2)
                SQL = Mid(SQL, 3)
                
                If DAV = "PR" Then
                    'PRODUCCION ANTIGUA
                    'SQL=Numero de produccion
                    DAV = "feccaduca"
                    SQL = DevuelveDesdeBD(conAri, "fecproduccion", "ariges" & EmprMorales & ".sordprod", "codigo", SQL, "N", DAV)
                    
                Else
                    'NUEVA PRODUCCION
                    'Cp.NumAlbar = "NP" & Format(Me.CodProduccion, "00000") & Format(Me.idLiProd, "00")
                    'NUEVA PRODUCCION
                    'Cp.NumAlbar = "NP" & Format(Me.CodProduccion, "00000") & Format(Me.idLiProd, "00")
                    SQL = "codigo = " & Mid(SQL, 1, 5) & " AND idlin = " & Mid(SQL, 6) & " AND 1"
                    DAV = "feccaduca"
                    SQL = DevuelveDesdeBD(conAri, "fhinicio", "ariges" & EmprMorales & ".prodlin", SQL, "1", "N", DAV)
                End If
                
                If SQL <> "" Then
                    'OK ha conseguido la fecha de produccion. Verificamos la de caducidad

                    'Verifico la fecha de caducidad
                    If DAV <> "" Then
                        If Not IsDate(DAV) Then
                            MsgBox "Error obteniendo caducidad. Lote: " & RS!NUmlote
                            DAV = DateAdd("yyyy", 2, CDate(SQL))
                        End If
                    End If

                'else
                '     'Si no la obtiene la fecha de produccion continua por ahi abajo ya que SQL sera ""
                End If
                
                
                
                
            End If
            
            
            If SQL = "" Then
                'MAL. No se encuentra en la uno el lote del articulo. No pongo fechas
                SQL = vUsu.Codigo & "," & RS!NumAlbar & "," & RS!numlinea & "," & RS!linea & ","
                
                NumRegElim = InStr(RS!NUmlote, " ")
                If NumRegElim > 0 Then
                    DAV = Mid(RS!NUmlote, 1, NumRegElim)
                Else
                    'Nuevos numeros de lote
                    DAV = Mid(RS!NUmlote, 5)
                End If
                SQL = SQL & DBSet(RS!Cantidad, "N", "N") & "," & DBSet(Trim(DAV), "T") & ",'',''"
            Else
                'OK, todo OK
                'tmpinformes(codusu,codigo1,campo1,campo2,importe1,nombre1,nombre2,nombre3)

                    
                F = CDate(SQL)
                SQL = vUsu.Codigo & "," & RS!NumAlbar & "," & RS!numlinea & "," & RS!linea & ","
                
                
                
                SQL = SQL & DBSet(RS!Cantidad, "N", "N") & "," & DBSet(BuscaChekc, "T") & ",'" & Format(F, "dd/mm/yyyy")
                F = CDate(DAV)
                SQL = SQL & "','" & Format(F, "dd/mm/yyyy") & "'"
                
            End If
            Aux = Aux & ", (" & SQL & ")"
        End If
        RS.MoveNext
    Wend
    RS.Close
    BuscaChekc = ""
    If Aux <> "" Then
        Aux = Mid(Aux, 2)
        SQL = "INSERT INTO tmpinformes(codusu,codigo1,campo1,campo2,importe1,nombre1,nombre2,nombre3) VALUES "
        SQL = SQL & Aux
        conn.Execute SQL
    End If
    
EPonerlotes:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set RS = Nothing
End Sub



Private Sub ModificarAñadirLineasNUEVO(Nueva As Boolean)
Dim Nlinea As Integer
Dim vWhere As String

    If data3.Recordset.EOF Then Exit Sub  'ALBARANES
    If Not Nueva Then
        If Data2.Recordset.EOF Then Exit Sub  'LINEAS
    End If
    
    'AGOSTO 2011
    'Abre un frm para poder modificar cualquier cosa de la linea... incluso borrarla
    CadenaDesdeOtroForm = ""
    Nlinea = -1
    With frmFacHcoLineaCambiar
        .Caption = Mid(Text1(4).Text & Space(10), 1, 10) & Text1(5).Text 'Los 10 primeros son el codclien
        .Codtipoa = data3.Recordset!Codtipoa
        .NumAlbar = data3.Recordset!NumAlbar
        .Codtipom = Data1.Recordset!Codtipom
        .NumFactu = Data1.Recordset!NumFactu
        .Fecfactu = Data1.Recordset!Fecfactu
        If Not Nueva Then
            Nlinea = Data2.Recordset!numlinea
            .numlinea = Nlinea
        Else
            .numlinea = -1
            .Caption = .Caption & "  [NUEVA]"
        End If
        .Show vbModal
    End With
    
    If CadenaDesdeOtroForm <> "" Then
        'Recalculo totatels
        
        'truco del almenduco. Para forzar todos los recalculos, pongo un update inirrealizable
        ModificarFactura "UPDATE sactiv set codactiv=-1 WHERE codactiv=-2"
            Espera 0.2
            TerminaBloquear
            PosicionarData
            FormatoDatosTotales
            DesBloqueoManual "scafac"
            PonerCamposLineas
            If Nlinea >= 0 Then SituarDataPosicion data3, CLng(Nlinea), ""
    End If
End Sub
