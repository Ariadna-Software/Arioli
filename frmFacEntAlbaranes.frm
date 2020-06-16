VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmFacEntAlbaranes 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7125
   ClientLeft      =   -495
   ClientTop       =   915
   ClientWidth     =   14805
   Icon            =   "frmFacEntAlbaranes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   14805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   15
      Left            =   9480
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   77
      Top             =   6780
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   16
      Left            =   2400
      Locked          =   -1  'True
      MaxLength       =   60
      TabIndex        =   76
      Text            =   "Text2 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwqa"
      Top             =   6780
      Visible         =   0   'False
      Width           =   6885
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   120
      TabIndex        =   57
      Top             =   6615
      Width           =   2175
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   240
         TabIndex        =   58
         Top             =   180
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   13440
      TabIndex        =   55
      Top             =   6720
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   12240
      TabIndex        =   54
      Top             =   6720
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   240
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   59
      Top             =   0
      Width           =   14805
      _ExtentX        =   26114
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   23
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
            Object.ToolTipText     =   "Lineas Albaran"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nº Series"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Generar Factura"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Marcar facturar"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir packing list"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Albaran"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
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
         Left            =   9960
         MaxLength       =   15
         TabIndex        =   122
         Text            =   "BASE IMP."
         Top             =   100
         Width           =   1490
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   56
         Left            =   11520
         MaxLength       =   15
         TabIndex        =   121
         Text            =   "Text1 7"
         Top             =   80
         Width           =   1530
      End
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   8520
         TabIndex        =   60
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   240
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   5220
      Left            =   120
      TabIndex        =   61
      Tag             =   "Fecha Oferta|F|N|||scapre|fecentre|dd/mm/yyyy|N|"
      Top             =   1275
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   9208
      _Version        =   393216
      Style           =   1
      Tabs            =   4
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
      TabPicture(0)   =   "frmFacEntAlbaranes.frx":000C
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
      Tab(0).Control(13)=   "txtAux(9)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtAux(10)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdAux(9)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtAux(12)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtAux(11)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtAux(13)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "Otros Datos"
      TabPicture(1)   =   "frmFacEntAlbaranes.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text1(41)"
      Tab(1).Control(1)=   "Text1(39)"
      Tab(1).Control(2)=   "FrameFacRec"
      Tab(1).Control(3)=   "FrameHco"
      Tab(1).Control(4)=   "Text1(29)"
      Tab(1).Control(5)=   "Text2(29)"
      Tab(1).Control(6)=   "Text1(28)"
      Tab(1).Control(7)=   "Text2(28)"
      Tab(1).Control(8)=   "Text1(27)"
      Tab(1).Control(9)=   "Text2(27)"
      Tab(1).Control(10)=   "Text1(2)"
      Tab(1).Control(11)=   "Text1(25)"
      Tab(1).Control(12)=   "Text1(26)"
      Tab(1).Control(13)=   "Text1(24)"
      Tab(1).Control(14)=   "Text1(23)"
      Tab(1).Control(15)=   "Text1(22)"
      Tab(1).Control(16)=   "Text1(21)"
      Tab(1).Control(17)=   "Text1(20)"
      Tab(1).Control(18)=   "Text1(19)"
      Tab(1).Control(19)=   "Text1(18)"
      Tab(1).Control(20)=   "Text1(38)"
      Tab(1).Control(21)=   "imgBuscar(12)"
      Tab(1).Control(22)=   "Label1(24)"
      Tab(1).Control(23)=   "imgBuscar(9)"
      Tab(1).Control(24)=   "Label1(23)"
      Tab(1).Control(25)=   "imgBuscar(8)"
      Tab(1).Control(26)=   "Label1(9)"
      Tab(1).Control(27)=   "imgBuscar(7)"
      Tab(1).Control(28)=   "Label1(12)"
      Tab(1).Control(29)=   "Label1(11)"
      Tab(1).Control(30)=   "Label1(10)"
      Tab(1).Control(31)=   "Label1(5)"
      Tab(1).Control(32)=   "Label1(3)"
      Tab(1).Control(33)=   "Label1(45)"
      Tab(1).ControlCount=   34
      TabCaption(2)   =   "Datos carga"
      TabPicture(2)   =   "frmFacEntAlbaranes.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "imgBuscar(13)"
      Tab(2).Control(1)=   "Label1(54)"
      Tab(2).Control(2)=   "imgFecha(44)"
      Tab(2).Control(3)=   "Label1(55)"
      Tab(2).Control(4)=   "Label1(56)"
      Tab(2).Control(5)=   "Label1(53)"
      Tab(2).Control(6)=   "Label1(57)"
      Tab(2).Control(7)=   "Label1(58)"
      Tab(2).Control(8)=   "Label1(59)"
      Tab(2).Control(9)=   "Label1(60)"
      Tab(2).Control(10)=   "Label1(61)"
      Tab(2).Control(11)=   "Label1(62)"
      Tab(2).Control(12)=   "Label1(63)"
      Tab(2).Control(13)=   "Label1(64)"
      Tab(2).Control(14)=   "Label1(65)"
      Tab(2).Control(15)=   "Label1(66)"
      Tab(2).Control(16)=   "Label1(67)"
      Tab(2).Control(17)=   "Label1(68)"
      Tab(2).Control(18)=   "Label1(69)"
      Tab(2).Control(19)=   "Label1(70)"
      Tab(2).Control(20)=   "Label1(71)"
      Tab(2).Control(21)=   "Label1(72)"
      Tab(2).Control(22)=   "Label1(73)"
      Tab(2).Control(23)=   "Label1(74)"
      Tab(2).Control(24)=   "Text1(43)"
      Tab(2).Control(25)=   "Text1(44)"
      Tab(2).Control(26)=   "Text1(45)"
      Tab(2).Control(27)=   "Text1(46)"
      Tab(2).Control(28)=   "Text1(47)"
      Tab(2).Control(29)=   "Text1(48)"
      Tab(2).Control(30)=   "Text1(49)"
      Tab(2).Control(31)=   "Text1(50)"
      Tab(2).Control(32)=   "Text1(51)"
      Tab(2).Control(33)=   "Text1(52)"
      Tab(2).Control(34)=   "Text1(53)"
      Tab(2).Control(35)=   "Text1(54)"
      Tab(2).Control(36)=   "Text1(55)"
      Tab(2).Control(37)=   "Text1(56)"
      Tab(2).Control(38)=   "Text1(57)"
      Tab(2).Control(39)=   "Text1(58)"
      Tab(2).Control(40)=   "chkCarga(0)"
      Tab(2).Control(41)=   "Text1(59)"
      Tab(2).Control(42)=   "chkCarga(1)"
      Tab(2).Control(43)=   "chkCarga(2)"
      Tab(2).Control(44)=   "chkCarga(3)"
      Tab(2).Control(45)=   "Text1(60)"
      Tab(2).ControlCount=   46
      TabCaption(3)   =   "Totales"
      TabPicture(3)   =   "frmFacEntAlbaranes.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "FrameFactura"
      Tab(3).ControlCount=   1
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   60
         Left            =   -69840
         MaxLength       =   10
         TabIndex        =   33
         Tag             =   "Hora|H|S|||scaalb|Hora|hh:mm|N|"
         Top             =   720
         Width           =   1065
      End
      Begin VB.CheckBox chkCarga 
         Caption         =   "Otros"
         Height          =   375
         Index           =   3
         Left            =   -62280
         TabIndex        =   52
         Tag             =   "Facturar|N|S|||scaalb|TransOtros||N|"
         Top             =   3960
         Width           =   975
      End
      Begin VB.CheckBox chkCarga 
         Caption         =   "Certificado limpieza"
         Height          =   375
         Index           =   2
         Left            =   -64440
         TabIndex        =   51
         Tag             =   "Facturar|N|N|||scaalb|TransCertLim||N|"
         Top             =   3960
         Width           =   2055
      End
      Begin VB.CheckBox chkCarga 
         Caption         =   "CMR"
         Height          =   375
         Index           =   1
         Left            =   -65640
         TabIndex        =   50
         Tag             =   "Facturar|N|N|||scaalb|TransCMR||N|"
         Top             =   3960
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   555
         Index           =   59
         Left            =   -71160
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   46
         Tag             =   "O1|T|S|||scaalb|TransMercancia||N|"
         Text            =   "frmFacEntAlbaranes.frx":007C
         Top             =   3000
         Width           =   8205
      End
      Begin VB.CheckBox chkCarga 
         Caption         =   "Ticket báscula"
         Height          =   375
         Index           =   0
         Left            =   -67440
         TabIndex        =   49
         Tag             =   "Facturar|N|N|||scaalb|TransTicketBas||N|"
         Top             =   3960
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   58
         Left            =   -69000
         MaxLength       =   10
         TabIndex        =   48
         Tag             =   "Deposito|N|S||50|scaalb|TransLacradasCompr|00|N|"
         Top             =   4080
         Width           =   825
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   57
         Left            =   -70200
         MaxLength       =   10
         TabIndex        =   47
         Tag             =   "Deposito|N|S|1|50|scaalb|TransLacradasCoop|00|N|"
         Top             =   4080
         Width           =   825
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   56
         Left            =   -64920
         MaxLength       =   40
         TabIndex        =   36
         Tag             =   "O1|T|S|||scaalb|TransAcidez||N|"
         Top             =   720
         Width           =   1785
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   55
         Left            =   -65040
         MaxLength       =   30
         TabIndex        =   45
         Tag             =   "O1|T|S|||scaalb|TransDestino||N|"
         Text            =   " "
         Top             =   2400
         Width           =   3165
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   54
         Left            =   -65520
         MaxLength       =   20
         TabIndex        =   41
         Tag             =   "O1|T|S|||scaalb|TransMatRemolque||N|"
         Text            =   "Text15"
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   53
         Left            =   -71160
         MaxLength       =   255
         TabIndex        =   53
         Tag             =   "O1|T|S|||scaalb|TransObsPrecintos||N|"
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
         Top             =   4680
         Width           =   8205
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   52
         Left            =   -63120
         MaxLength       =   10
         TabIndex        =   42
         Tag             =   "Bocas|N|S|1|100|scaalb|TransNumBocas|00|N|"
         Top             =   1680
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   51
         Left            =   -67560
         MaxLength       =   20
         TabIndex        =   40
         Tag             =   "O1|T|S|||scaalb|TransMatricula||N|"
         Text            =   "Text15"
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   50
         Left            =   -66120
         MaxLength       =   10
         TabIndex        =   35
         Tag             =   "Deposito|T|S|||scaalb|Deposito||N|"
         Top             =   720
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   49
         Left            =   -68640
         MaxLength       =   100
         TabIndex        =   34
         Tag             =   "O1|T|S|||scaalb|Muestra||N|"
         Top             =   720
         Width           =   2385
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   48
         Left            =   -67560
         MaxLength       =   30
         TabIndex        =   44
         Tag             =   "O1|T|S|||scaalb|TransCondDNI||N|"
         Text            =   " "
         Top             =   2400
         Width           =   2205
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   47
         Left            =   -61800
         MaxLength       =   10
         TabIndex        =   38
         Tag             =   "TaraKg|N|S|1||scaalb|TransTara|#,##0||"
         Top             =   720
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   46
         Left            =   -63000
         MaxLength       =   10
         TabIndex        =   37
         Tag             =   "BrutoKg|N|S|1||scaalb|TransBruto|#,##0||"
         Top             =   720
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   45
         Left            =   -71160
         MaxLength       =   10
         TabIndex        =   32
         Tag             =   "Fecha carga|F|S|||scaalb|FechaCarga|dd/mm/yyyy|N|"
         Top             =   720
         Width           =   1185
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   44
         Left            =   -71160
         MaxLength       =   30
         TabIndex        =   43
         Tag             =   "O1|T|S|||scaalb|TransConductor||N|"
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
         Top             =   2400
         Width           =   3405
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   43
         Left            =   -71160
         MaxLength       =   60
         TabIndex        =   39
         Tag             =   "O1|T|S|||scaalb|TransEmpresa||N|"
         Text            =   "Text15"
         Top             =   1680
         Width           =   3030
      End
      Begin VB.Frame FrameFactura 
         Height          =   3300
         Left            =   -74280
         TabIndex        =   149
         Top             =   840
         Width           =   10575
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
            Left            =   6120
            MaxLength       =   15
            TabIndex        =   172
            Text            =   "Text1 7"
            Top             =   2640
            Width           =   1845
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   48
            Left            =   4680
            MaxLength       =   15
            TabIndex        =   171
            Text            =   "Text1 7"
            Top             =   2160
            Width           =   1245
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   42
            Left            =   4080
            MaxLength       =   5
            TabIndex        =   170
            Text            =   "Text1 7"
            Top             =   2160
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   39
            Left            =   2040
            MaxLength       =   4
            TabIndex        =   169
            Text            =   "Text1 7"
            Top             =   2160
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   45
            Left            =   2760
            MaxLength       =   15
            TabIndex        =   168
            Text            =   "Text1 7"
            Top             =   2160
            Width           =   1125
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   47
            Left            =   4680
            MaxLength       =   15
            TabIndex        =   167
            Text            =   "Text1 7"
            Top             =   1800
            Width           =   1245
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   41
            Left            =   4080
            MaxLength       =   5
            TabIndex        =   166
            Text            =   "Text1 7"
            Top             =   1800
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   38
            Left            =   2040
            MaxLength       =   4
            TabIndex        =   165
            Text            =   "Text1 7"
            Top             =   1800
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   44
            Left            =   2760
            MaxLength       =   15
            TabIndex        =   164
            Text            =   "Text1 7"
            Top             =   1800
            Width           =   1125
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   46
            Left            =   4680
            MaxLength       =   15
            TabIndex        =   163
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   1245
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   40
            Left            =   4080
            MaxLength       =   5
            TabIndex        =   162
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   37
            Left            =   2040
            MaxLength       =   4
            TabIndex        =   161
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   43
            Left            =   2760
            MaxLength       =   15
            TabIndex        =   160
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   1125
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   36
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   159
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
            TabIndex        =   158
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
            TabIndex        =   157
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
            TabIndex        =   156
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   49
            Left            =   6120
            MaxLength       =   5
            TabIndex        =   155
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   52
            Left            =   6720
            MaxLength       =   15
            TabIndex        =   154
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   1245
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   50
            Left            =   6120
            MaxLength       =   5
            TabIndex        =   153
            Text            =   "Text1 7"
            Top             =   1800
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   53
            Left            =   6720
            MaxLength       =   15
            TabIndex        =   152
            Text            =   "Text1 7"
            Top             =   1800
            Width           =   1245
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   51
            Left            =   6120
            MaxLength       =   5
            TabIndex        =   151
            Text            =   "Text1 7"
            Top             =   2160
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   54
            Left            =   6720
            MaxLength       =   15
            TabIndex        =   150
            Text            =   "Text1 7"
            Top             =   2160
            Width           =   1245
         End
         Begin VB.Label Label1 
            Caption         =   "Cod."
            Height          =   255
            Index           =   42
            Left            =   2040
            TabIndex        =   187
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "% IVA"
            Height          =   255
            Index           =   41
            Left            =   4080
            TabIndex        =   186
            Top             =   1200
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "TOTAL ALBARAN"
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
            Index           =   39
            Left            =   4200
            TabIndex        =   185
            Top             =   2655
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
            TabIndex        =   184
            Top             =   2160
            Width           =   135
         End
         Begin VB.Line Line1 
            X1              =   2040
            X2              =   8040
            Y1              =   1065
            Y2              =   1065
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. IVA"
            Height          =   255
            Index           =   33
            Left            =   4800
            TabIndex        =   183
            Top             =   1200
            Width           =   735
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
            TabIndex        =   182
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
            TabIndex        =   181
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
            TabIndex        =   180
            Top             =   480
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "Base Imponible"
            Height          =   255
            Index           =   2
            Left            =   5760
            TabIndex        =   179
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Dto Gn"
            Height          =   255
            Index           =   18
            Left            =   3960
            TabIndex        =   178
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Dto PP"
            Height          =   255
            Index           =   22
            Left            =   2160
            TabIndex        =   177
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Bruto"
            Height          =   255
            Index           =   27
            Left            =   240
            TabIndex        =   176
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Base Imponible"
            Height          =   255
            Index           =   28
            Left            =   2760
            TabIndex        =   175
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. RE"
            Height          =   255
            Index           =   6
            Left            =   6960
            TabIndex        =   174
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "% RE"
            Height          =   255
            Index           =   48
            Left            =   6120
            TabIndex        =   173
            Top             =   1200
            Width           =   495
         End
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   13
         Left            =   5880
         MaxLength       =   12
         TabIndex        =   67
         Tag             =   "PrecioLitro"
         Text            =   "Palets"
         Top             =   3600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   1125
         Index           =   41
         Left            =   -72840
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         Tag             =   "o6|T|S|||scaalb|observa6||N|"
         Top             =   3840
         Width           =   7845
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   11
         Left            =   12240
         MaxLength       =   12
         TabIndex        =   68
         Tag             =   "Cajas"
         Text            =   "Cajas"
         Top             =   3960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   12
         Left            =   13200
         MaxLength       =   12
         TabIndex        =   71
         Tag             =   "PrecioLitro"
         Text            =   "PrecioLitro"
         Top             =   3960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Index           =   9
         Left            =   11880
         TabIndex        =   144
         ToolTipText     =   "Buscar artículo"
         Top             =   3960
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   10
         Left            =   12000
         TabIndex        =   143
         Tag             =   "Importe"
         Text            =   "nomprove"
         Top             =   3960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   9
         Left            =   11040
         MaxLength       =   12
         TabIndex        =   75
         Tag             =   "Importe"
         Text            =   "codprove"
         Top             =   3960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   39
         Left            =   -70800
         MaxLength       =   7
         TabIndex        =   142
         Tag             =   "Nº Venta|N|S|||scaalb|numventa|0000000|N|"
         Text            =   "Text1 7"
         Top             =   675
         Width           =   885
      End
      Begin VB.Frame FrameFacRec 
         Caption         =   "Datos Factura a rectificar "
         Height          =   1815
         Left            =   -68160
         TabIndex        =   134
         Top             =   480
         Width           =   2775
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   37
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   139
            Tag             =   "Tipo Mov. Factura|T|S|||scaalb|codtipmf||N|"
            Top             =   360
            Width           =   1185
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   36
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   137
            Tag             =   "Nº. Factura|N|S|0||scaalb|numfactu|0000000|N|"
            Top             =   780
            Width           =   1185
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   35
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   135
            Tag             =   "Fecha Factura|F|S|||scaalb|fecfactu|dd/mm/yyyy|N|"
            Top             =   1200
            Width           =   1185
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Mov."
            Height          =   255
            Index           =   47
            Left            =   240
            TabIndex        =   140
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Nº Factura"
            Height          =   255
            Index           =   46
            Left            =   240
            TabIndex        =   138
            Top             =   780
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Fact."
            Height          =   255
            Index           =   44
            Left            =   240
            TabIndex        =   136
            Top             =   1200
            Width           =   855
         End
      End
      Begin VB.Frame FrameHco 
         Height          =   2055
         Left            =   -68280
         TabIndex        =   123
         Top             =   360
         Width           =   4455
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   33
            Left            =   795
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   131
            Text            =   "Text2"
            Top             =   1560
            Width           =   3525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   33
            Left            =   135
            MaxLength       =   30
            TabIndex        =   130
            Text            =   "Text1"
            Top             =   1560
            Width           =   660
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   32
            Left            =   795
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   128
            Text            =   "Text2"
            Top             =   840
            Width           =   3525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   32
            Left            =   135
            MaxLength       =   30
            TabIndex        =   127
            Text            =   "Text1"
            Top             =   840
            Width           =   660
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   31
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   125
            Top             =   240
            Width           =   1185
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   11
            Left            =   1080
            Picture         =   "frmFacEntAlbaranes.frx":009B
            ToolTipText     =   "Buscar incidencia"
            Top             =   1320
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Incidencia"
            Height          =   255
            Index           =   40
            Left            =   120
            TabIndex        =   132
            Top             =   1320
            Width           =   825
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   10
            Left            =   1080
            Picture         =   "frmFacEntAlbaranes.frx":019D
            ToolTipText     =   "Buscar trabajador"
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Trabajador"
            Height          =   255
            Index           =   38
            Left            =   120
            TabIndex        =   129
            Top             =   615
            Width           =   825
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Eliminación"
            Height          =   255
            Index           =   37
            Left            =   120
            TabIndex        =   126
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label1 
            Height          =   255
            Index           =   29
            Left            =   360
            TabIndex        =   124
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   29
         Left            =   -72825
         MaxLength       =   30
         TabIndex        =   25
         Tag             =   "Cod. Envío|N|N|0|999|scaalb|codenvio|000|N|"
         Text            =   "Text1"
         Top             =   2040
         Width           =   660
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   29
         Left            =   -72120
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   112
         Text            =   "Text2"
         Top             =   2040
         Width           =   3525
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   28
         Left            =   -72825
         MaxLength       =   30
         TabIndex        =   24
         Tag             =   "Preparador Material|N|N|0|9999|scaalb|codtrab2|0000|N|"
         Text            =   "Text1"
         Top             =   1680
         Width           =   660
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   28
         Left            =   -72120
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   110
         Text            =   "Text2"
         Top             =   1680
         Width           =   3525
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   27
         Left            =   -72820
         MaxLength       =   30
         TabIndex        =   23
         Tag             =   "Trabajador pedido|N|S|0|9999|scaalb|codtrab1|0000|N|"
         Text            =   "Text1"
         Top             =   1320
         Width           =   660
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   27
         Left            =   -72120
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   108
         Text            =   "Text2"
         Top             =   1320
         Width           =   3525
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   -72285
         MaxLength       =   10
         TabIndex        =   106
         Tag             =   "Semana Entrega|N|S|||scaalb|sementre||N|"
         Top             =   675
         Width           =   705
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   25
         Left            =   -74520
         MaxLength       =   7
         TabIndex        =   103
         Tag             =   "Nº Pedido|N|S|||scaalb|numpedcl|0000000|N|"
         Text            =   "Text1 7"
         Top             =   675
         Width           =   885
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   26
         Left            =   -73545
         MaxLength       =   10
         TabIndex        =   102
         Tag             =   "Fecha Pedido|F|S|||scaalb|fecpedcl|dd/mm/yyyy|N|"
         Top             =   675
         Width           =   1185
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   24
         Left            =   -69840
         MaxLength       =   10
         TabIndex        =   98
         Tag             =   "Fecha Oferta|F|S|||scaalb|fecofert|dd/mm/yyyy|N|"
         Top             =   675
         Width           =   1185
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   23
         Left            =   -70800
         MaxLength       =   7
         TabIndex        =   97
         Tag             =   "Nº Oferta|N|S|||scaalb|numofert|0000000|N|"
         Text            =   "Text1 7"
         Top             =   675
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
         TabIndex        =   78
         Tag             =   "OF"
         Text            =   "OF"
         Top             =   3960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Frame FrameCliente 
         Height          =   2550
         Left            =   240
         TabIndex        =   82
         Top             =   315
         Width           =   14055
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   42
            Left            =   9945
            MaxLength       =   7
            TabIndex        =   146
            Tag             =   "RefProduccion|N|S|0||scaalb|refproduccion|0||"
            Text            =   "Text1 7"
            Top             =   2160
            Width           =   900
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   40
            Left            =   9840
            MaxLength       =   7
            TabIndex        =   19
            Tag             =   "Descuento General|N|S|||scaalb|aportacion|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   1020
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   34
            Left            =   11940
            MaxLength       =   30
            TabIndex        =   20
            Tag             =   "Cant. Km|N|S|0|99999|scaalb|cantidkm||N|"
            Text            =   "Text1"
            Top             =   480
            Width           =   950
         End
         Begin VB.CheckBox chkFacturarKm 
            Caption         =   "Facturar Km"
            Height          =   375
            Left            =   3720
            TabIndex        =   22
            Tag             =   "Facturar Km|N|N|||scaalb|facturkm||N|"
            Top             =   1680
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   13
            Left            =   1125
            MaxLength       =   20
            TabIndex        =   13
            Tag             =   "Referencia Cliente|T|S|||scaalb|referenc||N|"
            Text            =   "Text1 Text1 Text1 Te"
            Top             =   2160
            Width           =   1605
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   12
            Left            =   7470
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   94
            Tag             =   "Direccion/Dpto.|T|S|||scaped|nomdirec||N|"
            Text            =   "Text2"
            Top             =   165
            Width           =   3405
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   12
            Left            =   6885
            MaxLength       =   30
            TabIndex        =   14
            Tag             =   "Direccion/Dpto.|N|S|0|999|scaalb|coddirec|000|N|"
            Text            =   "Text1"
            Top             =   165
            Width           =   540
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   11
            Left            =   1125
            MaxLength       =   30
            TabIndex        =   12
            Tag             =   "Provincia|T|N|||scaalb|proclien||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text22"
            Top             =   1695
            Width           =   2445
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   9
            Left            =   1125
            MaxLength       =   6
            TabIndex        =   10
            Tag             =   "CPostal|T|N|||scaalb|codpobla||N|"
            Text            =   "Text15"
            Top             =   1335
            Width           =   630
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   10
            Left            =   1755
            MaxLength       =   30
            TabIndex        =   11
            Tag             =   "Población|T|N|||scaalb|pobclien||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
            Top             =   1335
            Width           =   3405
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   7
            Left            =   3360
            MaxLength       =   20
            TabIndex        =   8
            Tag             =   "teléfono Cliente|T|S|||scaalb|telclien||N|"
            Text            =   "12345678911234567899"
            Top             =   2160
            Width           =   1845
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   6
            Left            =   1125
            MaxLength       =   15
            TabIndex        =   7
            Tag             =   "NIF Cliente|T|N|||scaalb|nifclien||N|"
            Text            =   "123456789"
            Top             =   165
            Width           =   1110
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   17
            Left            =   6885
            MaxLength       =   30
            TabIndex        =   15
            Tag             =   "Cod. Agente|N|N|0|9999|scaalb|codagent|0000|N|"
            Text            =   "Text1"
            Top             =   513
            Width           =   540
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   17
            Left            =   7470
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   89
            Text            =   "Text2"
            Top             =   513
            Width           =   3405
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   14
            Left            =   6885
            MaxLength       =   30
            TabIndex        =   16
            Tag             =   "Forma de Pago|N|N|0|999|scaalb|codforpa|000|N|"
            Text            =   "Text1"
            Top             =   861
            Width           =   540
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   14
            Left            =   7470
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   84
            Text            =   "Text2"
            Top             =   861
            Width           =   3405
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   15
            Left            =   6885
            MaxLength       =   7
            TabIndex        =   17
            Tag             =   "Descuento P.Pago|N|N|0|99.90|scaalb|dtoppago|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1209
            Width           =   540
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   16
            Left            =   6885
            MaxLength       =   7
            TabIndex        =   18
            Tag             =   "Descuento General|N|N|0|99.90|scaalb|dtognral|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   540
         End
         Begin VB.ComboBox cboFacturacion 
            Height          =   315
            Left            =   6840
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Tag             =   "Tipo Facturación|N|N|||scaalb|tipofact||N|"
            Top             =   2160
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   8
            Left            =   1125
            MaxLength       =   35
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Tag             =   "Domicilio|T|N|||scaalb|domclien||N|"
            Text            =   "frmFacEntAlbaranes.frx":029F
            Top             =   513
            Width           =   4030
         End
         Begin VB.Label Label1 
            Caption         =   "Ref. producción"
            Height          =   255
            Index           =   51
            Left            =   8520
            TabIndex        =   147
            Top             =   2160
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "APORTACION TERMINAL"
            Height          =   255
            Index           =   49
            Left            =   7560
            TabIndex        =   145
            Top             =   1575
            Width           =   2175
         End
         Begin VB.Label Label1 
            Caption         =   "Km a facturar"
            Height          =   255
            Index           =   43
            Left            =   11940
            TabIndex        =   133
            Top             =   255
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Ref. Cliente"
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   101
            Top             =   2160
            Width           =   855
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   6
            Left            =   855
            Picture         =   "frmFacEntAlbaranes.frx":02C3
            ToolTipText     =   "Buscar población"
            Top             =   1350
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Direc./Dpto"
            Height          =   255
            Index           =   1
            Left            =   5700
            TabIndex        =   96
            Top             =   165
            Width           =   855
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   2
            Left            =   6600
            Picture         =   "frmFacEntAlbaranes.frx":03C5
            ToolTipText     =   "Buscar direc./dpto"
            Top             =   165
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Provincia"
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   95
            Top             =   1695
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Población"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   93
            Top             =   1335
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Tfno"
            Height          =   255
            Index           =   19
            Left            =   2805
            TabIndex        =   92
            Top             =   2160
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "NIF"
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   91
            Top             =   165
            Width           =   615
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   855
            Picture         =   "frmFacEntAlbaranes.frx":04C7
            ToolTipText     =   "Buscar cliente varios"
            Top             =   165
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Agente"
            Height          =   255
            Index           =   34
            Left            =   5700
            TabIndex        =   90
            Top             =   513
            Width           =   615
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   5
            Left            =   6600
            Picture         =   "frmFacEntAlbaranes.frx":05C9
            ToolTipText     =   "Buscar agente"
            Top             =   516
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Forma Pago"
            Height          =   255
            Index           =   15
            Left            =   5700
            TabIndex        =   88
            Top             =   861
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. P.Pago"
            Height          =   255
            Index           =   25
            Left            =   5700
            TabIndex        =   87
            Top             =   1215
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. Gral"
            Height          =   255
            Index           =   26
            Left            =   5715
            TabIndex        =   86
            Top             =   1575
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Facturac."
            Height          =   255
            Index           =   4
            Left            =   5640
            TabIndex        =   85
            Top             =   2160
            Width           =   1095
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   6600
            Picture         =   "frmFacEntAlbaranes.frx":06CB
            ToolTipText     =   "Buscar forma de pago"
            Top             =   867
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Domicilio"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   83
            Top             =   513
            Width           =   735
         End
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Index           =   1
         Left            =   2640
         TabIndex        =   81
         ToolTipText     =   "Buscar artículo"
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
         TabIndex        =   80
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
         TabIndex        =   66
         Tag             =   "Nombre Artículo"
         Text            =   "nomArtic"
         Top             =   3960
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
         TabIndex        =   74
         Tag             =   "Importe"
         Text            =   "Importe"
         Top             =   3960
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
         TabIndex        =   73
         Tag             =   "Descuento 2"
         Text            =   "Dto2"
         Top             =   3960
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
         TabIndex        =   72
         Tag             =   "Descuento 1"
         Text            =   "Dto1"
         Top             =   3960
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
         TabIndex        =   70
         Tag             =   "Precio"
         Text            =   "123,456.7879"
         Top             =   3960
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
         TabIndex        =   69
         Tag             =   "Cantidad"
         Text            =   "1,234,567,891.25"
         Top             =   3960
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
         TabIndex        =   65
         Tag             =   "Código Artículo"
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
         Left            =   240
         MaxLength       =   15
         TabIndex        =   64
         Tag             =   "Código Almacen"
         Text            =   "codalmac"
         Top             =   3840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   22
         Left            =   -72840
         MaxLength       =   80
         TabIndex        =   30
         Tag             =   "Observación 5|T|S|||scaalb|observa05||N|"
         Top             =   3480
         Width           =   7845
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   21
         Left            =   -72840
         MaxLength       =   80
         TabIndex        =   29
         Tag             =   "Observación 4|T|S|||scaalb|observa04||N|"
         Top             =   3240
         Width           =   7845
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   20
         Left            =   -72840
         MaxLength       =   80
         TabIndex        =   28
         Tag             =   "Observación 3|T|S|||scaalb|observa03||N|"
         Top             =   3000
         Width           =   7845
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   19
         Left            =   -72840
         MaxLength       =   80
         TabIndex        =   27
         Tag             =   "Observación 2|T|S|||scaalb|observa02||N|"
         Top             =   2760
         Width           =   7845
      End
      Begin VB.TextBox Text1 
         Height          =   280
         Index           =   18
         Left            =   -72840
         MaxLength       =   80
         TabIndex        =   26
         Tag             =   "Observación 1|T|S|||scaalb|observa01||N|"
         Top             =   2520
         Width           =   7845
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmFacEntAlbaranes.frx":07CD
         Height          =   2040
         Left            =   240
         TabIndex        =   79
         Top             =   3000
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   3598
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
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   38
         Left            =   -69840
         MaxLength       =   10
         TabIndex        =   141
         Tag             =   "Nº terminal|N|S|||scaalb|numtermi||N|"
         Top             =   675
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hora"
         Height          =   195
         Index           =   74
         Left            =   -69720
         TabIndex        =   209
         Top             =   480
         Width           =   345
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
         Left            =   -71160
         TabIndex        =   208
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Comprador"
         Height          =   195
         Index           =   72
         Left            =   -69000
         TabIndex        =   207
         Top             =   3840
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Coop"
         Height          =   195
         Index           =   71
         Left            =   -70080
         TabIndex        =   206
         Top             =   3840
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Acidez"
         Height          =   195
         Index           =   70
         Left            =   -64920
         TabIndex        =   205
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Destino"
         Height          =   195
         Index           =   69
         Left            =   -65040
         TabIndex        =   204
         Top             =   2160
         Width           =   540
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
         Left            =   -74760
         TabIndex        =   203
         Top             =   2880
         Width           =   2535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Matrí. remolque"
         Height          =   195
         Index           =   67
         Left            =   -65520
         TabIndex        =   202
         Top             =   1440
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Precintos"
         Height          =   195
         Index           =   66
         Left            =   -71160
         TabIndex        =   201
         Top             =   4440
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DNI conductor"
         Height          =   195
         Index           =   65
         Left            =   -67560
         TabIndex        =   200
         Top             =   2160
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre conductor"
         Height          =   195
         Index           =   64
         Left            =   -71160
         TabIndex        =   199
         Top             =   2160
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Bocas precintadas"
         Height          =   195
         Index           =   63
         Left            =   -63120
         TabIndex        =   198
         Top             =   1440
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Matrícula"
         Height          =   195
         Index           =   62
         Left            =   -67560
         TabIndex        =   197
         Top             =   1440
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Index           =   61
         Left            =   -71160
         TabIndex        =   196
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tara  (kg)"
         Height          =   195
         Index           =   60
         Left            =   -61800
         TabIndex        =   195
         Top             =   480
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Bruto(kg)"
         Height          =   195
         Index           =   59
         Left            =   -63000
         TabIndex        =   194
         Top             =   480
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Depósito"
         Height          =   195
         Index           =   58
         Left            =   -66120
         TabIndex        =   193
         Top             =   480
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Muestra"
         Height          =   195
         Index           =   57
         Left            =   -68640
         TabIndex        =   192
         Top             =   480
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fec. carga"
         Height          =   195
         Index           =   53
         Left            =   -71160
         TabIndex        =   191
         Top             =   480
         Width           =   765
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
         Left            =   -74760
         TabIndex        =   190
         Top             =   3840
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
         Left            =   -74760
         TabIndex        =   189
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   44
         Left            =   -70320
         Picture         =   "frmFacEntAlbaranes.frx":07E2
         ToolTipText     =   "Buscar fecha"
         Top             =   480
         Width           =   240
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
         Left            =   -74760
         TabIndex        =   188
         Top             =   600
         Width           =   2535
      End
      Begin VB.Image imgBuscar 
         Enabled         =   0   'False
         Height          =   240
         Index           =   13
         Left            =   -70440
         Picture         =   "frmFacEntAlbaranes.frx":086D
         ToolTipText     =   "Buscar población"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   12
         Left            =   -73200
         Picture         =   "frmFacEntAlbaranes.frx":096F
         ToolTipText     =   "Buscar trabajador"
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. Envío"
         Height          =   255
         Index           =   24
         Left            =   -74520
         TabIndex        =   113
         Top             =   2055
         Width           =   1425
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   9
         Left            =   -73095
         Picture         =   "frmFacEntAlbaranes.frx":0A71
         ToolTipText     =   "Buscar forma de envio"
         Top             =   2055
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Preparador Material"
         Height          =   255
         Index           =   23
         Left            =   -74520
         TabIndex        =   111
         Top             =   1695
         Width           =   1425
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   -73095
         Picture         =   "frmFacEntAlbaranes.frx":0B73
         ToolTipText     =   "Buscar trabajador"
         Top             =   1695
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Trabajador Pedido"
         Height          =   255
         Index           =   9
         Left            =   -74520
         TabIndex        =   109
         Top             =   1340
         Width           =   1420
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   -73100
         Picture         =   "frmFacEntAlbaranes.frx":0C75
         ToolTipText     =   "Buscar trabajador"
         Top             =   1330
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Sem. Entrega"
         Height          =   255
         Index           =   12
         Left            =   -72285
         TabIndex        =   107
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Pedido"
         Height          =   255
         Index           =   11
         Left            =   -74520
         TabIndex        =   105
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Pedido"
         Height          =   255
         Index           =   10
         Left            =   -73545
         TabIndex        =   104
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Oferta"
         Height          =   255
         Index           =   5
         Left            =   -69840
         TabIndex        =   100
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Oferta"
         Height          =   255
         Index           =   3
         Left            =   -70800
         TabIndex        =   99
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones"
         Height          =   255
         Index           =   45
         Left            =   -74520
         TabIndex        =   63
         Top             =   2520
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   13410
      TabIndex        =   56
      Top             =   6720
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   114
      Top             =   360
      Width           =   11415
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   6960
         MaxLength       =   30
         TabIndex        =   5
         Tag             =   "Cod. Cliente|N|N|0|999999|scaalb|codclien|000000|N|"
         Text            =   "Text1"
         Top             =   480
         Width           =   760
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   5
         Left            =   7780
         MaxLength       =   40
         TabIndex        =   6
         Tag             =   "Nombre Cliente|T|N|||scaalb|nomclien||N|"
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwww aq"
         Top             =   495
         Width           =   3360
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   6960
         MaxLength       =   30
         TabIndex        =   4
         Tag             =   "Realizada Por|N|N|0|9999|scaalb|codtraba|0000|N|"
         Text            =   "Text1"
         Top             =   120
         Width           =   760
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   3
         Left            =   7780
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   119
         Text            =   "Text2"
         Top             =   120
         Width           =   3360
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   30
         Left            =   1275
         TabIndex        =   1
         Tag             =   "Tipo Albaran|T|N|||scaalb|codtipom||S|"
         Text            =   "Text3"
         Top             =   345
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha Albaran|F|N|||scaalb|fechaalb|dd/mm/yyyy|N|"
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
         Tag             =   "Nº Albaran|N|S|0||scaalb|numalbar|0000000|S|"
         Text            =   "Text1 7"
         Top             =   345
         Width           =   885
      End
      Begin VB.CheckBox chkFacturar 
         Caption         =   "Facturar"
         Height          =   375
         Left            =   3345
         TabIndex        =   3
         Tag             =   "Facturar|N|N|||scaalb|factursn||N|"
         Top             =   360
         Width           =   975
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   6660
         Picture         =   "frmFacEntAlbaranes.frx":0D77
         ToolTipText     =   "Buscar cliente"
         Top             =   510
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         Height          =   255
         Index           =   0
         Left            =   5595
         TabIndex        =   120
         Top             =   495
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Realizada Por"
         Height          =   255
         Index           =   21
         Left            =   5595
         TabIndex        =   118
         Top             =   165
         Width           =   1050
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   6660
         Picture         =   "frmFacEntAlbaranes.frx":0E79
         ToolTipText     =   "Buscar trabajador"
         Top             =   150
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fec. Alb."
         Height          =   255
         Index           =   14
         Left            =   2040
         TabIndex        =   117
         Top             =   150
         Width           =   735
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   2835
         Picture         =   "frmFacEntAlbaranes.frx":0F7B
         ToolTipText     =   "Buscar fecha"
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Albaran"
         Height          =   255
         Index           =   50
         Left            =   240
         TabIndex        =   116
         Top             =   150
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo"
         Height          =   255
         Index           =   8
         Left            =   1275
         TabIndex        =   115
         Top             =   150
         Width           =   735
      End
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Servicios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   360
      Index           =   75
      Left            =   12480
      TabIndex        =   210
      Top             =   600
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.Label Label1 
      Caption         =   "Grado"
      Height          =   255
      Index           =   52
      Left            =   9480
      TabIndex        =   148
      Top             =   6600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Ampliación Línea"
      Height          =   255
      Index           =   35
      Left            =   2400
      TabIndex        =   62
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
Attribute VB_Name = "frmFacEntAlbaranes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========== VBLES PUBLICAS ====================
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'Si se llama de la busqueda en el frmAlmMovimArticulos se accede
'a las tablas del Albaran  de Venta de Facturas de movimiento seleccionado (solo consulta)
Public hcoCodMovim As String 'cod. movim
Public hcoCodTipoM As String 'Codigo detalle de Movimiento(ALV,ALR,ALS)

Public EsHistorico As Boolean 'Si es true abrir el formulario con la tabla de
                              'de historico schalb, y solo en modo de consulta
                              
Public RecuperarFactu As Boolean 'si esta recuperando facturas al generar las facturas no coger contaror
                                 'pedirlas por teclado
                                 

'========== VBLES PRIVADAS ====================
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
Private WithEvents frmFE As frmFacFormasEnvio  'Form Formas de Envio
Attribute frmFE.VB_VarHelpID = -1

'Private WithEvents frmNSerie As frmRepCargarNSerie  'Form Cargar nº Series
Private WithEvents frmMen As frmMensajes  'Form Mensajes
Attribute frmMen.VB_VarHelpID = -1
Private WithEvents frmList As frmListadoOfer
Attribute frmList.VB_VarHelpID = -1

Private WithEvents frmProv As frmComProveedores
Attribute frmProv.VB_VarHelpID = -1

Private WithEvents frmO As frmFacCopiarObservaciones2
Attribute frmO.VB_VarHelpID = -1

Private WithEvents frmVh As frmFacVehiculos
Attribute frmVh.VB_VarHelpID = -1


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


Dim ModificaLineas As Byte
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

Private Ordenacion As String 'Para el ORDER BY de la consulta
Private NombreTabla As String  'Nombre de la tabla o de la
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

Dim cadList As String 'cadena para pasar al historico
Dim motivo As String 'cadena para el motivo si es factura Rectificativa


Dim PulsadoMas2 As Boolean

Dim txtAnterior As String

Dim ClienteConTasaReciclado As Boolean  'Cuando pasamos a las lineas pondremos esta variab


'Para las lineas. Tanto nueva, como modificando
Dim ElArticulo As CArticulo

Dim txtArt As String  'Para cuando pincha para seleccionar articulo
Dim RN As ADODB.Recordset

Private Sub cboFacturacion_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub chkCarga_KeyPress(Index As Integer, KeyAscii As Integer)
 KEYpress KeyAscii
End Sub

Private Sub chkFacturar_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub chkFacturar_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub chkFacturarKm_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim PrimeraLin As Boolean 'Si se inserta la primera linea no esta creado el datagrid1 entonces llamar
                          ' a DataGrid, sino llamar solo a DataGrid2
Dim numlinea As String

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
        Case 3 'INSERTAR
            If DatosOk Then
                InsertarCabecera
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificarCabAlbaran Then
                    TerminaBloquear
                    PosicionarData
                End If
            End If
            
         Case 5 'InsertarModificar LINEA
            'Actualizar el registro en la tabla de lineas 'slialb'
            If ModificaLineas = 1 Then 'INSERTAR lineas Albaran
                PrimeraLin = False
                If Data2.Recordset.EOF = True Then PrimeraLin = True
                If InsertarLinea(numlinea, False) Then
                    'Comprobar si el Articulo tiene control de Nº de Serie
                    ComprobarNSeriesLineas numlinea
                    If PrimeraLin Then
                        'Para LaVall
                        If vParamAplic.QUE_EMPRESA = 4 Then
                            cadList = "codtipom=" & DBSet(Me.Data1.Recordset!Codtipom, "T") & " AND numalbar "
                            cadList = DevuelveDesdeBD(conAri, "transmercancia", "scaalb", cadList, CStr(Data1.Recordset!NumAlbar))
                            If cadList = "" Then
                                cadList = "UPDATE  scaalb set transmercancia = " & DBSet(txtAux(2).Text, "T")
                                cadList = cadList & "codtipom=" & DBSet(Me.Data1.Recordset!Codtipom, "T") & " AND numalbar =" & Data1.Recordset!NumAlbar
                                EjecutaSQL conAri, cadList
                                Text1(59).Text = txtAux(2).Text
                            End If
                            
                            
                            cadList = ""
                        End If
                        CargaGrid DataGrid1, Data2, True
                    Else
                        CargaGrid2 DataGrid1, Data2
                    End If
                    BotonAnyadirLinea
                End If
                
            ElseIf ModificaLineas = 2 Then 'MODIFICAR lineas
                If ModificarLinea Then
                    numlinea = Data2.Recordset!numlinea
                    'Comprobar si el Articulo tiene control de Nº de Serie
                    ComprobarNSeriesLineas numlinea
                    TerminaBloquear
                    
                    CargaTxtAux False, False
                    CargaGrid2 DataGrid1, Data2
                    ModificaLineas = 0
                    PonerBotonCabecera True
                    BloquearTxt Text2(16), True
                    BloquearTxt Text2(15), True
                    Me.DataGrid1.Enabled = True
                    
                    SituarData Data2, "numlinea = " & numlinea, Me.lblIndicador.Caption
                End If
                
            End If
            CalcularDatosFactura
    End Select
    Screen.MousePointer = vbDefault

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Function ModificarCabAlbaran() As Boolean
Dim b As Boolean
Dim SQL As String

    On Error GoTo EModificaAlb
    conn.BeginTrans
    
    'Si es cliente de varios actualizar datos cliente en tabla:sclvar
    b = ActualizarClienteVarios(Text1(4).Text, Text1(6).Text)
    
    If b Then
        b = ModificaDesdeFormulario(Me, 1)
        If b Then
            SQL = "UPDATE scaalb SET nomdirec=" & DBSet(Text2(12).Text, "T") & " WHERE codtipom=" & DBSet(CodTipoMov, "T") & " and numalbar=" & Data1.Recordset!NumAlbar
            conn.Execute SQL
        End If

        If b Then
            'comprobar si se ha cambiado el cliente
            'o si se ha cambiado la fecha del albaran
            'If (CInt(Me.Data1.Recordset!CodClien) <> CInt(Text1(4).Text)) Or (CDate(Data1.Recordset!FechaAlb) <> CDate(Text1(1).Text)) Then
            'DAVID.   No es un CINT. Tiene que ser un clng o val
            If (Val(Me.Data1.Recordset!CodClien) <> Val(Text1(4).Text)) Or (CDate(Data1.Recordset!FechaAlb) <> CDate(Text1(1).Text)) Then
                'si hay numeros de serie en ese albaran, actualizamos el cliente
                'al nuevo cliente
                SQL = "UPDATE sserie SET codclien=" & DBSet(Text1(4).Text, "N") & ","
                SQL = SQL & " fechavta=" & DBSet(Text1(1).Text, "F")
                SQL = SQL & " WHERE codtipom='" & CodTipoMov & "'" & " AND numalbar=" & Data1.Recordset!NumAlbar & " and fechavta=" & DBSet(Data1.Recordset!FechaAlb, "F")
                conn.Execute SQL
                
                'Modificar el cliente en la smoval
                SQL = "UPDATE smoval SET codigope=" & DBSet(Text1(4).Text, "N") & ","
                SQL = SQL & " fechamov=" & DBSet(Text1(1).Text, "F")
                'MODIF   DAVID   13 OCTUBRE 2009
                'SQL = SQL & ", horamovi= concat(" & DBSet(Text1(1).Text, "F") & ",hour(horamovi),':',minute(horamovi),':',second(horamovi))"
                SQL = SQL & ", horamovi= '" & Format(Text1(1).Text, FormatoFecha) & " "
                If Me.Text1(60).Text = "" Then
                    SQL = SQL & Format(Now, "hh:nn:ss") & "'"
                Else
                    SQL = SQL & Text1(60).Text & "'"
                End If
                SQL = SQL & " WHERE detamovi='" & CodTipoMov & "'" & " AND document="
                'ANTES
                'SQL = SQL & DBSet(CStr(Data1.Recordset!NumAlbar), "T") & " and fechamov=" & DBSet(Data1.Recordset!FechaAlb, "F")
                'AHORA
                SQL = SQL & "'" & Text1(0).Text & "' and fechamov=" & DBSet(Data1.Recordset!FechaAlb, "F")
                conn.Execute SQL
            End If
            
            
            'Si ha cambiado la fecha , actualizo en la tabla
            'de lineas de repartos de rutas
            If CDate(Text1(1).Text) <> CDate(Data1.Recordset!FechaAlb) Then
                SQL = "UPDATE srepartol set FechaAlb=" & DBSet(Text1(1).Text, "F")
                SQL = SQL & " where codtipom='" & Text1(30).Text & "' and numalbar=" & Text1(0).Text
                SQL = SQL & " AND fechaalb = " & DBSet(Data1.Recordset!FechaAlb, "F")
                conn.Execute SQL
            End If
            
        End If
    End If
    
EModificaAlb:
    If Err.Number <> 0 Then b = False
    If b Then
        conn.CommitTrans
    Else
        conn.RollbackTrans
    End If
    ModificarCabAlbaran = b
    If Err.Number <> 0 Then MuestraError Err.Number, "Modificar cabecera Albaran.", Err.Description
End Function




Private Sub cmdAux_Click(Index As Integer)
Dim b As Boolean
Dim HaCambiado As Boolean

    HaCambiado = False
    Select Case Index
        Case 0 'Busqueda de Cod. Almacen
            Set frmAlm = New frmAlmAlPropios
            frmAlm.DatosADevolverBusqueda = "0"
            frmAlm.Show vbModal
            Set frmAlm = Nothing
        Case 1 'Busqueda de Cod. Artic
            b = True
            If CodTipoMov = "ART" Then
                If MsgBox("¿Desea traer líneas de la factura que va a rectificar?", vbQuestion + vbYesNo) = vbYes Then
                
                    'si es Albaran de Factura rectificativa cargar un listview con todas las
                    'lineas de la factura y marcar las que queremos seleccionar para
                    'cargarlas en las lineas del Albaran rectificativo
                    b = False
                    Set frmMen = New frmMensajes
                    frmMen.cadWhere = " codtipom=" & DBSet(Text1(37).Text, "T") & " and numfactu=" & Text1(36).Text & " and fecfactu=" & DBSet(Text1(35).Text, "F")
                    frmMen.OpcionMensaje = 11 'Lineas Factura a Rectificar
                    frmMen.Show vbModal
                    Set frmMen = Nothing
                    CargaGrid Me.DataGrid1, Me.Data2, True
                    cmdCancelar_Click
                End If
            End If
            
            If b Then
                
                Set frmArt = New frmAlmArticulos
                frmArt.DatosADevolverBusqueda2 = "@1@" 'Poner en Modo busqueda
                frmArt.DeConsulta = True
                frmArt.ParaVenta = True
                frmArt.Show vbModal
                Set frmArt = Nothing
                If txtArt <> "" Then HaCambiado = True
'                txtAux_LostFocus (1)
            End If
    Case 9
        Set frmProv = New frmComProveedores
        frmProv.DatosADevolverBusqueda = "1"
        frmProv.Show vbModal
        Set frmProv = Nothing
    End Select
    PonerFoco txtAux(Index)
    If HaCambiado Then
        '
        txtAux(1).Text = RecuperaValor(txtArt, 1) 'Cod Artic
        txtAux(2).Text = RecuperaValor(txtArt, 2) 'Nom Artic
    
        
    End If
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
            BloquearTxt Text2(15), True
            DataGrid1.Columns(4).Caption = "Artículo"
            If ModificaLineas = 1 Then 'INSERTAR
                ModificaLineas = 0
                DataGrid1.AllowAddNew = False
                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
            End If
            ModificaLineas = 0
            PonerBotonCabecera True
            Me.DataGrid1.Enabled = True
    End Select
End Sub


Private Sub BotonAnyadir()
'Añadir registro en tabla de cabecera de Pedidos: scaped (Cabecera)
Dim NomTraba As String
Dim cad As String
Dim RS As ADODB.Recordset

    LimpiarCampos 'Vacía los TextBox
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    
    'Si es Albaran para factura RECTIFICATIVA pedir la Factura que se va
    'a Rectificar y si existe en el historico, tabla "scafac", entonces dejamos
    'que inserte el Albaran Rectificativo, si no salimos
    If CodTipoMov = "ART" Then
        cadList = ""

        Set frmList = New frmListadoOfer
        frmList.OpcionListado = 225
        frmList.Show vbModal
        Set frmList = Nothing
        If cadList = "" Then Exit Sub
        
        'cargar los datos de la factura recuperada en el formulario
        NomTraba = "select codtipom as codtipmf,numfactu,fecfactu,codclien,nomclien,domclien,scafac.codpobla,pobclien,proclien,nifclien,telclien,"
        NomTraba = NomTraba & "coddirec,nomdirec,scafac.codagent,nomagent,scafac.codforpa, nomforpa,dtoppago,dtognral "
        NomTraba = NomTraba & " from (scafac inner join sforpa on scafac.codforpa=sforpa.codforpa) "
        NomTraba = NomTraba & " inner join sagent on scafac.codagent=sagent.codagent where " & cadList
        
        Set RS = New ADODB.Recordset
        RS.Open NomTraba, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        
        PonerModo 3
        
        If Not RS.EOF Then
            Text1(4).Text = RS!CodClien
            FormateaCampo Text1(4)
            Text1(5).Text = RS!nomclien
            Text1(6).Text = RS!nifClien
            Text1(7).Text = DBLet(RS!telclien, "T")
            Text1(8).Text = RS!domclien
            Text1(9).Text = RS!codpobla
            Text1(10).Text = RS!pobclien
            Text1(11).Text = DBLet(RS!proclien, "T")
            Text1(12).Text = DBLet(RS!CodDirec, "T")
            FormateaCampo Text1(12)
            Text2(12).Text = DBLet(RS!nomdirec, "T")
            Text1(14).Text = RS!codforpa
            FormateaCampo Text1(14)
            Text2(14).Text = RS!nomforpa
            Text1(15).Text = DBLet(RS!DtoPPago, "N")
            FormateaCampo Text1(15)
            Text1(16).Text = DBLet(RS!DtoGnral, "N")
            FormateaCampo Text1(16)
            Text1(17).Text = DBLet(RS!codagent, "T")
            FormateaCampo Text1(17)
            Text2(17).Text = RS!nomagent
            Text1(37).Text = RS!codtipmf
            Text1(36).Text = DBLet(RS!NumFactu, "N")
            FormateaCampo Text1(36)
            Text1(35).Text = RS!Fecfactu
            
            'Observacion 1   'DAVID
            'Text1(18).Text = "RECTIFICA A FACTURA: " & RS!codtipmf & ", " & RS!NumFactu & ", " & RS!FecFactu
            Text1(18).Text = RS!NumFactu & ", " & RS!Fecfactu
            'Observacion 2
            Text1(19).Text = motivo
            
            NomTraba = "tipofact"
            cad = DevuelveDesdeBD(conAri, "clivario", "sclien", "codclien", Text1(4).Text, "N", NomTraba)
            If cad = "0" Then BloquearDatosCliente (False)
            
            cad = DevuelveDesdeBD(conAri, "clivario", "sclien", "codclien", Text1(4).Text, "N", NomTraba)
            If cad = "0" Then BloquearDatosCliente (False)
            
            
            
            
            'recuperamos el tipo de facturacion del cliente
            Me.cboFacturacion.ListIndex = CInt(NomTraba)
        Else
            cad = "N" 'para que la busqueda de despues no de error
        End If
        RS.Close
        
        
        
        'DAVID
        'Para que meta la letra de serie, NO el tipo moviemiento
        RS.Open "SELECT * FROM stipom WHERE codtipom='" & cad & "'"
        If Not RS.EOF Then cad = DBLet(RS!LetraSer, "T")
        RS.Close
        If cad = "" Then cad = CodTipoMov
        Text1(18).Text = "RECTIFICA A FACTURA: " & cad & ", " & Text1(18).Text
        
        
        
        'Traeremos resto datos
        'cad=replace(cadlist,"scafac.","scafac1"
        
        
        
        Set RS = Nothing
    Else
        'Añadiremos el boton de aceptar y demas objetos para insertar
        PonerModo 3
        
        
        If vParamAplic.QUE_EMPRESA = 2 Then
            ' Observaciones fijas para la empresa 2 MOIXENT
            Text1(22).Text = "Hora salida: " & Format(Now, "hh:mm:ss") & ".    Grado alcohólico a 20º de temperatura"
        End If
    End If
    
    NomTraba = ""
    'Poner el nombre del trabajador que esta conectado
    Text1(3).Text = PonerTrabajadorConectado(NomTraba)
    Text2(3).Text = NomTraba

    'El preparador del material lo hacemos tb al trabajador actual
    Text1(28).Text = Text1(3).Text
    Text2(28).Text = Text2(3).Text

    Text1(1).Text = Format(Now, "dd/mm/yyyy") 'Fecha Albaran
    Text1(30).Text = CodTipoMov
    Me.chkFacturar.Value = 1
    PonerFoco Text1(1)
End Sub


Private Sub BotonAnyadirLinea()
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
       
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
    Text2(15).Text = ""
    
    
    'Descuentos a 0
    txtAux(6).Text = 0
    txtAux(7).Text = 0
    
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
    
    If Me.EsHistorico = False Then
        'Hacer busquedar del tipo de movimiento de albaran en el que estamos
        Text1(30).Text = CodTipoMov
        BloquearTxt Text1(30), True
    End If
End Sub


Private Sub BotonVerTodos()
Dim Aux As String
Dim cad As String

'    LimpiarCampos
    Aux = ""

    
    If chkVistaPrevia.Value = 1 Then
        EsCabecera = True
        cad = " codtipom='" & CodTipoMov & "'"
        If Aux <> "" Then cad = cad & " AND " & Aux
            
        MandaBusquedaPrevia cad
    Else
        LimpiarCampos
        LimpiarDataGrids
        CadenaConsulta = "Select * from " & NombreTabla
        If EsHistorico = False Then
            CadenaConsulta = CadenaConsulta & " WHERE codtipom='" & CodTipoMov & "'"
            If Aux <> "" Then CadenaConsulta = CadenaConsulta & " AND " & Aux
        Else
            
        End If
        CadenaConsulta = CadenaConsulta & " " & Ordenacion
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

    'Añadiremos el boton de aceptar y demas objetos para insertar
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
    
    If Data2.Recordset.EOF Then Exit Sub
    
    'bloqueamos el registro a modificar
    vWhere = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas) & " and numlinea=" & Data2.Recordset!numlinea
    If Not BloqueaRegistro(NomTablaLineas, vWhere) Then Exit Sub
    
    CargaTxtAux True, False
    
    'si es factura rectificativa y es una linea de la factura que rectificamos
    'solo podremos modificar la cantidad el resto de campos bloqueados
    If CodTipoMov = "ART" Then '(Albaran Rectificativo)
        vWhere = "codtipom='" & Text1(37).Text & "' and numfactu=" & Text1(36).Text & " and fecfactu=" & DBSet(Text1(35).Text, "F")
        vWhere = vWhere & " and codartic=" & DBSet(txtAux(1).Text, "T")
        vWhere = "SELECT COUNT(*) FROM slifac WHERE " & vWhere
        If RegistrosAListar(vWhere) > 0 Then
            'modificamos una linea de factura a rectificar y solo podemos modificar cantidad
            BloquearTxt txtAux(0), True
            BloquearTxt txtAux(1), True
            BloquearTxt txtAux(2), True
            BloquearTxt txtAux(4), True
            BloquearTxt txtAux(6), True
            BloquearTxt txtAux(7), True
            Me.cmdAux(0).Enabled = False
            Me.cmdAux(1).Enabled = False
        End If
    End If
    
    
    ModificaLineas = 2 'Modificar
    'Añadiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR"
    PonerBotonCabecera False
    
    BloquearTxt Text2(16), False 'Campo Ampliacion Linea
    If vParamAplic.QUE_EMPRESA = 2 Then
        'Moixent bodega
        vWhere = "if(codtipar='05',1,0)+if(codfamia=6,1,0)"  'Los 05 o la familia 6
        vWhere = DevuelveDesdeBD(conAri, vWhere, "sartic", "codartic", txtAux(1).Text, "T")
        'If vWhere <> "05" Or vWhere = "04" Then vWhere = ""  'solo el granel
        If vWhere <> "" Then
            If Val(vWhere) = 0 Then vWhere = ""
        End If
        BloquearTxt Text2(15), vWhere = ""  'Campo hectogrado
    End If
    
    BloquearTxt txtAux(2), True
    PonerFoco txtAux(0)
    Me.DataGrid1.Enabled = False


    Set ElArticulo = New CArticulo
    ElArticulo.LeerDatos Data2.Recordset!codArtic

EModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Mantenimientos (scaman)
' y los registros correspondientes de las tablas de lineas (sliman y slima1)
Dim cad As String
Dim NumAlbElim As Long

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    If DBLet(Data1.Recordset!refproduccion, "N") > 0 Then
        cad = String(40, "*") & vbCrLf & vbCrLf
        cad = cad & "ALBARAN BLOQUEADO " & vbCrLf & vbCrLf & cad & vbCrLf & vbCrLf
    Else
        cad = ""
    End If
    cad = cad & "Cabecera de Albaranes." & vbCrLf
    cad = cad & "------------------------------------       " & vbCrLf & vbCrLf
    cad = cad & "Va a eliminar el Albaran:            "
    cad = cad & vbCrLf & "Tipo:  " & Text1(30).Text
    cad = cad & vbCrLf & "Nº:  " & Format(Text1(0).Text, "0000000")
    cad = cad & vbCrLf & "Fecha:  " & Text1(1).Text
'    cad = cad & vbCrLf & "Cliente:  " & Format(Text1(4).Text, "000000") & " - " & Text1(5).Text
    cad = cad & vbCrLf & vbCrLf & " ¿Desea Eliminarlo? "
      
    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Abrir frame de informes para pedir datos antes de grabar en el historico
        cadList = ""
        Set frmList = New frmListadoOfer
        frmList.OpcionListado = 80
        frmList.Show vbModal
        Set frmList = Nothing
        If cadList = "" Then Exit Sub
        
        Screen.MousePointer = vbHourglass
        
        NumRegElim = Data1.Recordset.AbsolutePosition
        NumAlbElim = Data1.Recordset.Fields(1).Value
        CodTipoMov = Text1(30).Text
        
        If Not Eliminar(NumAlbElim) Then
            Screen.MousePointer = vbDefault
            Exit Sub
         Else
            PosicionarDataTrasEliminar
        End If
        
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Albaran", Err.Description
End Sub


Private Sub BotonEliminarLinea()
'Eliminar una linea De Mantenimiento. Tabla: slima1
Dim SQL As String

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar

    If Data2.Recordset.EOF Then Exit Sub
        
    'Julio 2014
    'Si es una venta directa de aceite, pertenece a un deposito, tiene que pasar a la edicion de lotes y eliminar la asignacion
    If vParamAplic.Produccion Then
        Set miRsAux = New ADODB.Recordset
        SQL = "Select * from slialblotes " & Replace(ObtenerWhereCP(True), NombreTabla, "slialblotes")
        SQL = SQL & " and numlinea=" & Data2.Recordset!numlinea
        miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        While Not miRsAux.EOF
            SQL = SQL & ", " & DBSet(miRsAux!numLote, "T")
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        'Tiene los lotes
        If SQL <> "" Then
            SQL = Mid(SQL, 2)
            SQL = "(" & SQL & ")"
            SQL = "select * from proddepositos WHERE numlote in " & SQL
            miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            SQL = ""
            While Not miRsAux.EOF
                SQL = SQL & vbCrLf & "Lote: " & miRsAux!numLote & " --> Deposito " & miRsAux!NumDeposito
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            
        End If
        Set miRsAux = Nothing
        If SQL <> "" Then
            SQL = "Elimine manualmente los lotes asignados al articulo" & vbCrLf & SQL
            MsgBox SQL, vbExclamation
            Exit Sub
        End If
        
    End If
    ModificaLineas = 3 'Eliminar
    SQL = "¿Seguro que desea eliminar la línea de Albaran?     "
    SQL = SQL & vbCrLf & "NumLinea:  " & Data2.Recordset!numlinea & vbCrLf
    SQL = SQL & "Almacen:  " & Format(Data2.Recordset!codAlmac, "000")
    SQL = SQL & vbCrLf & "Artículo:  " & Data2.Recordset!codArtic & " - " & Data2.Recordset!NomArtic
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Data2.Recordset.AbsolutePosition
        If EliminarLinea Then
            ModificaLineas = 0
            CargaGrid2 DataGrid1, Data2
            SituarDataTrasEliminar Data2, NumRegElim
            CalcularDatosFactura
        End If
'        CancelaADODC
    End If
    PonerFocoBtn Me.cmdRegresar

EEliminarLinea:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Lineas Mantenimientos", Err.Description
End Sub


Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim cad As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
        If Not LineasRecicladoCorrectas Then Exit Sub
        PonerModo 2
        If Not Data1.Recordset.EOF Then _
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
        cad = Data1.Recordset.Fields(0) & "|"
        cad = cad & Data1.Recordset.Fields(1) & "|"
        RaiseEvent DatoSeleccionado(cad)
        Unload Me
    End If
End Sub


Private Sub DataGrid1_DblClick()
    If Data2.Recordset Is Nothing Then Exit Sub
    If Data2.Recordset.EOF Then Exit Sub
    If Modo = 5 And ModificaLineas = 0 Then
        'Modo lineas sin insertar ni modificar
         LanzaLote -1
    End If
    TerminaBloquear
End Sub

Private Sub DataGrid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Ayuda de Etiqueta de precio de salida de la Función de Precios
    If Data2.Recordset Is Nothing Then Exit Sub
    If Data2.Recordset.EOF Then Exit Sub
    If (Modo = 2) Or (Modo = 5 And ModificaLineas = 0) Then
        Me.DataGrid1.ToolTipText = ""
        If X > 9164 And X < 9520 Then
            If IsNull(Me.Data2.Recordset!origpre) Then Exit Sub
            Select Case DataGrid1.Columns(11).Value
                Case "P": Me.DataGrid1.ToolTipText = "P: Promoción"
                Case "E": Me.DataGrid1.ToolTipText = "E: Precio Especial"
                Case "T": Me.DataGrid1.ToolTipText = "T: Tarifa Artículo"
                Case "A": Me.DataGrid1.ToolTipText = "A: Precio Artículo"
                Case "M": Me.DataGrid1.ToolTipText = "M: Manual"
                Case "O": Me.DataGrid1.ToolTipText = "O: Tarifa-oferta"
'                Case Else
'                    Me.DataGrid1.ToolTipText = ""
            End Select
'        Else
'            Me.DataGrid1.ToolTipText = ""
        End If
    End If
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim RS As ADODB.Recordset
Dim SQL As String

    On Error GoTo Error1

    If Not Data2.Recordset.EOF And ModificaLineas <> 1 Then '1: Insertar
        SQL = "select ampliaci,hectogrado from " & NomTablaLineas & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " and numlinea=" & Data2.Recordset!numlinea
        Set RS = New ADODB.Recordset
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then
            Text2(16).Text = DBLet(RS.Fields(0).Value, "T")
            If vParamAplic.QUE_EMPRESA = 2 Then
                If RS.Fields(1).Value = 1 Then
                    'Cuando es UNO no lo pinto, no ha lugar
                    Text2(15).Text = ""
                Else
                    Text2(15).Text = DBLet(RS.Fields(1).Value, "N") * 100
                    PonerFormatoDecimal Text2(15), 3
                End If
            Else
                Text2(15).Text = ""
            End If
            
        End If
        RS.Close
        Set RS = Nothing
    Else
        Text2(16).Text = ""
        Text2(15).Text = ""
    End If
    
Error1:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    'Viene de DblClick en frmAlmMovimArticulos y carga el form con los valores
    If hcoCodMovim <> "" And Not Data1.Recordset.EOF And Modo <> 5 Then PonerCadenaBusqueda
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmppal.Icon

    ' ICONITOS DE LA BARRA
    btnAnyadir = 5
    btnPrimero = 20
    With Me.Toolbar1
        .ImageList = frmppal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        .Buttons(5).Image = 3   'Insertar Nuevo
        .Buttons(6).Image = 4   'Modificar
        .Buttons(7).Image = 5   'Borrar
        .Buttons(10).Image = 10 'Mto Lineas Ofertas
        .Buttons(11).Image = 33 'Nº Serie si lineas con articulos de control Nº serie
        .Buttons(12).Image = 26 'GEnerar factura
        .Buttons(13).Image = 30 'Marcar a facturar
        
        .Buttons(15).Image = 40 'packing list
        
        .Buttons(16).Image = 16 'Imprimir

        .Buttons(17).Image = 15  'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With
    
    
    Me.SSTab1.TabVisible(2) = vParamAplic.QUE_EMPRESA = 4
    Me.SSTab1.Tab = 0
    
    Me.Toolbar1.Buttons(15).visible = vParamAplic.EsAVAB Or vParamAplic.QUE_EMPRESA = 4
    If vParamAplic.QUE_EMPRESA = 4 Then Toolbar1.Buttons(15).ToolTipText = "impresion transporte"
    
    
    
    LimpiarCampos   'Limpia los campos TextBox
    
    CargarComboFacturacion
    VieneDeBuscar = False
    CodTipoMov = hcoCodTipoM
    
    Label1(75).visible = False
    If Not EsHistorico And hcoCodTipoM = "ALS" Then Label1(75).visible = True
    
    
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
  
        Text1(41).visible = True
  
    Else
        'MORALES
        Text1(6).MaxLength = 15
        Text1(6).Width = 1590
        Text1(8).MaxLength = 35
        Text1(8).Height = Text1(6).Height


        Text1(41).visible = False   'Para morales NO dejo ver el observa6 ... de momento

    End If
    
    
    
    
    
    If CodTipoMov = "ALR" Then
        Me.Caption = "Albaranes Reparación"
        Label1(3).visible = False
        Label1(5).visible = False
        Text1(23).visible = False
        Text1(24).visible = False
        Label1(12).visible = False
        Text1(2).visible = False
    End If
   Caption = "Albaranes Clientes"
   If hcoCodTipoM = "ALZ" Then Caption = Caption & "      *********"
   
    'Comprobar si es Departamento o Direccion
    If vParamAplic.Departamento Then
        Me.Label1(1).Caption = "Dpto."
    Else
        Me.Label1(1).Caption = "Direc."
    End If
        
    '## A mano
    Me.FrameHco.visible = EsHistorico
    Me.FrameFacRec.visible = (CodTipoMov = "ART")
    
    
    
    'Aportacion a terminal
    Label1(49).visible = hcoCodTipoM = "ALM" And vParamAplic.ctaAportacion <> ""
    Text1(40).visible = hcoCodTipoM = "ALM" And vParamAplic.ctaAportacion <> ""
    
    If Not EsHistorico Then
        NombreTabla = "scaalb"
        NomTablaLineas = "slialb" 'Tabla lineas de Albaranes
        Ordenacion = " ORDER BY codtipom, numalbar "
        If CodTipoMov = "ALV" Then
            Me.Caption = "Albaranes Clientes"
        ElseIf CodTipoMov = "ALM" Then
            Me.Caption = "Albaranes de Mostrador"
            If Me.RecuperarFactu Then Me.Caption = "Albaranes de Mostrador (Recuperar facturas)"
        ElseIf CodTipoMov = "ART" Then
            Me.Caption = "Albaranes Rectificativos"
        ElseIf CodTipoMov = "ALI" Then
            Me.Caption = "Albaranes internos"
        End If
    Else
        NombreTabla = "schalb"
        NomTablaLineas = "slhalb"
        CargarTagsHco Me, "scaalb", NombreTabla
        'Estos campos solo estan en la tabla del histórico
        Text1(31).Tag = "Fecha Eliminación|F|N|||schalb|fechelim|dd/mm/yyyy|N|"
        Text1(32).Tag = "Trabajador Eliminación|N|N|0|9999|schalb|trabelim|0000|N|"
        Text1(33).Tag = "Incidencia elim.|T|N|||schalb|codincid||N|"
        Me.Caption = "Histórico Albaranes Clientes"
        Ordenacion = " ORDER BY codtipom, numalbar,fechaalb "
    End If
 
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    CadenaConsulta = "Select * from " & NombreTabla
    If hcoCodMovim <> "" Then
        'Se llama desde Dobleclick en frmAlmMovimArticulos
        CadenaConsulta = CadenaConsulta & " WHERE codtipom='" & hcoCodTipoM & "' AND numalbar= " & hcoCodMovim
    Else
        CadenaConsulta = CadenaConsulta & " where numalbar=-1"
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
            PonerModo 1
            Text1(0).BackColor = vbYellow
        End If
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
    End If
End Sub


Private Sub LimpiarCampos()
On Error Resume Next

    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.cboFacturacion.ListIndex = -1
    Me.chkFacturar.Value = 0
    Me.chkFacturarKm.Value = 0
    Me.chkCarga(0).Value = 0: Me.chkCarga(1).Value = 0: Me.chkCarga(2).Value = 0: Me.chkCarga(3).Value = 0
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
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Agente
    FormateaCampo Text1(Indice)
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom agente
End Sub

Private Sub frmAlm_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Almacenes Propios
    txtAux(0).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Almacen
End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Articulos
    txtArt = CadenaSeleccion
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim cadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        If EsCabecera Then 'Llama desde VerTodos del Form
            cadB = ""
            Aux = ValorDevueltoFormGrid(Text1(30), CadenaDevuelta, 1)
            cadB = Aux
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 2)
            cadB = cadB & " and " & Aux
            
            If EsHistorico Then
                Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 3)
                cadB = cadB & " and " & Aux
            End If
            
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
            Text1(0).Text = Format(RecuperaValor(CadenaDevuelta, 2), "0000000")
            
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
Dim devuelve As String

    Indice = 9
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
    Text1(Indice + 1).Text = ObtenerPoblacion(Text1(Indice).Text, devuelve) 'Poblacion
    'provincia
    Text1(Indice + 2).Text = devuelve
    
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
    Indice = 14
    Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Forma Pago
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Forma Pago
End Sub


Private Sub frmList_DatoSeleccionado(CadenaSeleccion As String)
'devuelve los datos necesarios para grabar en la tabla del historico
' o para recuperar una factura que vamos a Rectificar

    cadList = ""
    
    If frmList.OpcionListado = 225 Then  'Factura Rectificativa
        If CadenaSeleccion <> "" Then
            'codtipom
            cadList = " codtipom='" & RecuperaValor(CadenaSeleccion, 1) & "' and numfactu="
            'numfactu
            cadList = cadList & RecuperaValor(CadenaSeleccion, 2) & " and fecfactu="
            'fecfactu
            cadList = cadList & DBSet(RecuperaValor(CadenaSeleccion, 3), "F")
            
            'campos observaciones
            motivo = "MOTIVO: " & RecuperaValor(CadenaSeleccion, 4)
        End If
        
    Else 'Para recoger los Datos de Eliminacion que se introdujeron
        cadList = DBSet(RecuperaValor(CadenaSeleccion, 1), "F") & " as fechelim,"
        cadList = cadList & RecuperaValor(CadenaSeleccion, 2) & " as trabelim,"
        cadList = cadList & DBSet(RecuperaValor(CadenaSeleccion, 3), "T") & " as codincid"
    End If
End Sub


Private Sub frmMen_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mensaje de Nº de Serie disponibles
'En cadena seleccion estan concatenados los seleccionados

    If frmMen.OpcionMensaje = 11 Then
        'En cadenaseleccion tenemos la WHERE que selecciona las lineas de la factura
        'que nos queremos traer para generar un albaran de rectificacion
        'Insertaremos estas lineas en la tabla slialb, y luego se podran eliminar,modificar,etc. (son de apoyo)
         InsertarLineasFactu (CadenaSeleccion)
    Else
        If Text1(30).Text = "ART" Then
            'Albaran de factura rectificativa
            If Not QuitarNumSeriesAlbVenta(CadenaSeleccion) Then MsgBox "Los nº de serie a rectificar no se han actualizado correctamente.", vbExclamation
        Else
            If Not AsignarNumSeriesAlbVenta(CadenaSeleccion) Then
                MsgBox "Los nº de serie del albaran no se han actualizado correctamente.", vbExclamation
            End If
        End If
    End If
End Sub


Private Sub frmNSerie_CargarNumSeries()
Dim CadValues As String, cadValuesU As String
Dim devuelve As String
Dim TieneMan As String * 1

    'Estamos en VENTAS e insertamos datos venta vacios
    If ModificaLineas = 4 Then
        CargarNumSeries
    Else
        'Viene de insertar Nº de series al insertar una linea

        'Comprobar que el cliente tiene mantenimientos en esa direc/dpto
        TieneMan = "0"
        devuelve = DevuelveDesdeBDNew(conAri, "scaman", "nummante", "codclien", Text1(4).Text, "N", , "coddirec", Text1(12).Text, "N")
        'El cliente tiene Mantenimientos
        If devuelve <> "" Then TieneMan = "1"
        
        'cadena para INSERT
        'Estamos en VENTAS e insertamos datos de Cliente
        CadValues = ""
        CadValues = CadValues & Text1(4).Text & ", " & DBSet(Text1(12).Text, "T") & ", " & TieneMan & ", " & DBSet(devuelve, "T") & ", "
        CadValues = CadValues & ValorNulo & ", " & ValorNulo & ", " 'Fecha ult. Repar y Fin Garantia
        'Datos Venta
        CadValues = CadValues & DBSet(Text1(30).Text, "T") & ", " & ValorNulo & ", '" & Format(Text1(1).Text, FormatoFecha) & "', " & Text1(0).Text & ", " & Me.cmdAux(0).Tag & ", "
        'Rellenar los datos COMPRA del Proveedor a NULO
        CadValues = CadValues & ValorNulo & ", " & ValorNulo & ", " & ValorNulo & ", " & ValorNulo
        
        'cadena para UPDATE
        cadValuesU = " codclien=" & Text1(4).Text & ", coddirec=" & DBSet(Text1(12).Text, "T") & ", "
        cadValuesU = cadValuesU & " tieneman=" & TieneMan & ", nummante=" & DBSet(devuelve, "T") & ", codtipom=" & DBSet(Text1(30).Text, "T")
        cadValuesU = cadValuesU & ", fechavta='" & Format(Text1(1).Text, FormatoFecha) & "' "
        cadValuesU = cadValuesU & ", numalbar=" & Text1(0).Text & ", numline1=" & Me.cmdAux(0).Tag
        InsertarNSeries txtAux(1).Text, CadValues, cadValuesU, True
    End If
End Sub


Private Sub frmO_DatoSeleccionado(Datos As String)
Dim I As Integer
    For I = 1 To 5
        Text1(I + 17).Text = RecuperaValor(Datos, I)
    Next
    If Text1(41).visible Then Text1(41).Text = RecuperaValor(Datos, 6)
End Sub

Private Sub frmProv_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(9).Text = RecuperaValor(CadenaSeleccion, 1)
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Trabajadores
Dim Indice As Byte
    Indice = Val(Me.imgBuscar(3).Tag)
    Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000") 'Cod Trabajador
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Trabajador
End Sub


Private Sub frmVh_DatoSeleccionado(CadenaSeleccion As String)
    motivo = CadenaSeleccion
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim Indice As Byte

    If Modo = 2 Or Modo = 0 Then Exit Sub
    
    'Trabajador albaran
    If Index = 3 Then
        If Text1(3).Text <> "" Then Exit Sub
    End If
    
    TerminaBloquear
    Screen.MousePointer = vbHourglass

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
             
        Case 3, 7, 8 'Realizada Por Trabajador (Pedido, Albaran, Preparador Material
            If Index = 7 Then
                Indice = 27
            ElseIf Index = 8 Then
                Indice = 28
            Else
                Indice = Index
            End If
            Me.imgBuscar(3).Tag = Indice
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
            
        Case 9 'Cod Envio
            Indice = 29
            PonerFoco Text1(Indice)
            Set frmFE = New frmFacFormasEnvio
            frmFE.DatosADevolverBusqueda = "0"
            frmFE.Show vbModal
            Set frmFE = Nothing
            
        Case 12
            'Observaciones FRA
            Indice = 18
            Set frmO = New frmFacCopiarObservaciones2
            frmO.PackingList = False
            frmO.IdCliente = CLng(Text1(4).Text)
            frmO.Show vbModal
            Set frmO = Nothing
    
    
        Case 13
            'Vehiculos para LA vall
            Indice = 43
            motivo = ""
            Set frmVh = New frmFacVehiculos
            frmVh.DatosADevolverBusqueda = "0"
            frmVh.Show vbModal
            Set frmVh = Nothing
            If motivo <> "" Then
                Set miRsAux = New ADODB.Recordset
                motivo = "select * from svehiculos where codigo=" & RecuperaValor(motivo, 1)
                miRsAux.Open motivo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not miRsAux.EOF Then
                    'matricula  Empresa   Conductor   DNIConductor
                    If Not IsNull(miRsAux!Empresa) Then Text1(43).Text = miRsAux!Empresa
                    If Not IsNull(miRsAux!matricula) Then Text1(51).Text = miRsAux!matricula
                    If Not IsNull(miRsAux!conductor) Then Text1(44).Text = miRsAux!conductor
                    If Not IsNull(miRsAux!DNIConductor) Then Text1(48).Text = miRsAux!DNIConductor
                    If Not IsNull(miRsAux!MatriculaRemolque) Then Text1(54).Text = miRsAux!MatriculaRemolque
                End If
                miRsAux.Close
                Set miRsAux = Nothing
                motivo = ""
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
   PonerFoco Text1(Indice)
End Sub


Private Sub mnBuscar_Click()
    Me.SSTab1.Tab = 0
    BotonBuscar
End Sub


Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar lineas de Pedido
         BotonEliminarLinea
    Else   'Eliminar Albaran
         If Not ComprobarSiNoEstaEnOrdenCarga Then Exit Sub
         BotonEliminar
    End If
End Sub


Private Sub mnImprimir_Click()
'Imprimir Albaran
    BotonImprimir (45) '45: Informe de Albaranes
End Sub


Private Sub mnLineas_Click()

    If Not ComprobarSiNoEstaEnOrdenCarga Then Exit Sub

    'Si esta vinculado, tampo deberia cambiar cantidades
    If Not ComprobarVinculado Then Exit Sub

    BotonMtoLineas 0, "Albaranes"
End Sub


Private Sub mnModificar_Click()
    If Modo = 5 Then 'Modificar lineas
         BotonModificarLinea
    Else   'Modificar albaran
         If Not ComprobarSiNoEstaEnOrdenCarga Then Exit Sub
         If BLOQUEADesdeFormulario(Me) Then BotonModificar
    End If
End Sub


Private Sub mnNuevo_Click()
    If Modo = 5 Then 'Añadir lineas
         BotonAnyadirLinea
    Else 'Añadir Cabecera
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

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub


Private Sub Text1_Change(Index As Integer)
    If Index = 9 Then HaCambiadoCP = True        'Cod. Postal

End Sub

'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    If Index = 41 Then Exit Sub
    txtAnterior = Text1(Index).Text
    kCampo = Index
    If Index = 9 Then HaCambiadoCP = False 'CPostal
    If Modo > 2 And Index = 3 Then
        If Text1(3).Text <> "" Then PonerFoco Text1(4)
    End If
    If Not (Index = 30 And Modo = 1) Then
        
        ConseguirFoco Text1(Index), Modo
        If Text1(Index).MultiLine Then Text1(Index).SelLength = 0
    End If
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Index <> 41 Then KEYdown KeyCode
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
Dim devuelve As String
Dim campo As String
        
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
          
          
    'Por si no ha cambiado nada
    If txtAnterior = Text1(Index).Text Then Exit Sub
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 1, 45 'Fecha Albaran - fecha carga
                If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
                
        Case 3, 27, 28 'Cod Vendedor
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "straba", "nomtraba", "codtraba")
                If Index = 3 And Modo = 3 Then
                    Text1(28).Text = Text1(Index).Text
                    Text2(28).Text = Text2(Index).Text
                End If
            Else
                Text2(Index).Text = ""
            End If
            
        Case 4 'Cod. Cliente
            If PonerFormatoEntero(Text1(Index)) Then
                If Modo = 1 Then 'Modo=1 Busqueda
                    Text1(5).Text = PonerNombreDeCod(Text1(Index), conAri, "sclien", "nomclien")
                Else 'If Modo = 3 Then 'Modo Insertar
                    'si es ART-Albaran de factura Rectificativa ya he cargado los
                    'datos de la factura
                    If CodTipoMov <> "ART" Then
                        PonerDatosCliente (Text1(Index).Text)
                    Else
                        campo = "nomclien"
                        devuelve = DevuelveDesdeBD(conAri, "clivario", "sclien", "codclien", Text1(4).Text, "N", campo)
                        If campo <> Text1(5).Text Then PonerDatosCliente Text1(Index).Text
                    End If
                End If
            Else
                LimpiarDatosCliente
            End If
            
        Case 6 'NIF
'            If Not EsDeVarios Then Exit Sub
'            'si no se ha modificado el nif del cliente no hacer nada (Modo 4=Modificar)
'            If (Modo = 4) Then
'                If (Text1(6).Text = Data1.Recordset!nifClien) Then Exit Sub
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
                 Text1(Index + 1).Text = ObtenerPoblacion(Text1(Index).Text, devuelve)
                 Text1(Index + 2).Text = devuelve
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
                'Comprobar que el cliente tiene mantenimientos en esa direc/dpto
                devuelve = DevuelveDesdeBDNew(conAri, "scaman", "nummante", "codclien", Text1(4).Text, "N", , "coddirec", Text1(12).Text, "N")
                If devuelve <> "" Then MsgBox "El cliente tiene Mantenimientos.", vbInformation
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
            If PonerFormatoDecimal(Text1(Index), 4) Then   'Tipo 4: Decimal(4,2)
                If Modo = 4 Then CalcularDatosFactura
            End If
            
        Case 17 'Cod. Agente
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sagent", "nomagent")
            Else
                Text2(Index).Text = ""
            End If
            
        Case 29 'Cod envio
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "senvio", "nomenvio")
            Else
                Text2(Index).Text = ""
            End If
        Case 40
            PonerFormatoDecimal Text1(Index), 3
                    
        
        Case 52, 46, 47, 57, 58
            If Not PonerFormatoEntero(Text1(Index)) Then Text1(Index).Text = ""
            
        Case 60
            PonerFormatoHora Text1(Index)
    End Select
End Sub


Private Sub HacerBusqueda()
Dim cadB As String

    'Poner el valor del combo Tipos de Movimiento Asociado
'    If Me.cboTipomov.ListIndex <> -1 Then
'        Text1(30).Text = ObtenerCodTipom
'    End If

    cadB = ObtenerBusqueda(Me, False)
    
    

    
    If chkVistaPrevia = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then
        'Se muestran en el mismo form
        If Me.EsHistorico = False Then
            cadB = cadB & " and codtipom='" & CodTipoMov & "'" 'Solo seleccionamos los del Movimiento, aqui los ALV
        End If
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim Tabla As String
Dim Titulo As String
Dim Desc As String, devuelve As String
    'Llamamos a al form
    '##A mano
    cad = ""
    If EsCabecera Then
        cad = cad & ParaGrid(Text1(30), 10, "Tipo Alb.")
        cad = cad & ParaGrid(Text1(0), 15, "Nº Albaran")
        cad = cad & ParaGrid(Text1(1), 15, "Fecha Ped.")
        cad = cad & ParaGrid(Text1(4), 10, "Cliente")
        cad = cad & ParaGrid(Text1(5), 50, "Nombre Cliente")
        Tabla = NombreTabla
        Titulo = "Albaranes"
        
        If EsHistorico Then
            Titulo = "Histórico de Albaranes"
            devuelve = "0|1|2|"
        Else
            Titulo = "Albaranes"
            devuelve = "0|1|"
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
        cad = cad & "Cod. " & Desc & "|sdirec|coddirec|N|000|15·"
        cad = cad & "Desc. " & Desc & "|sdirec|nomdirec|T||55·"
        Tabla = "sdirec"
        devuelve = "0|1|"
    End If
           
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = Tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = devuelve
'        frmB.vDevuelve = devuelve
        frmB.vTitulo = Titulo
        frmB.vselElem = 1
        frmB.vConexionGrid = conAri  'Conexión a BD: Ariges
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
            Text1(0).BackColor = vbYellow
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

    Screen.MousePointer = vbHourglass
    On Error GoTo EPonerLineas

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
Dim b As Boolean

    On Error Resume Next

    If Data1.Recordset.EOF Then Exit Sub
    
     'Si es un Albaran de Ticket visualizamos unos datos y sino otros
    b = (Data1.Recordset!EsTicket = 1)
    Me.Toolbar1.Buttons(11).Enabled = (Not b) And (Not EsHistorico)
    
    'sem. entrega pedido
    Label1(12).visible = Not b
    Text1(2).visible = Not b
    'num oferta
    Text1(23).visible = Not b
    'fecha oferta
    Text1(24).visible = Not b
    'nº terminal
    Text1(38).visible = b
    'nº venta
    Text1(39).visible = b
    
    If b Then
    'El albaran se genero a partir de un ticket
        Me.Label1(11).Caption = "Nº Ticket"
        Me.Label1(10).Caption = "Fecha Ticket"
        Me.Label1(9).Caption = "Trabajador Ticket"
    
        'ocultamos los datos de la oferta
        Me.Label1(3).Caption = "Nº Venta"
        Label1(5).Caption = "Nº Terminal"
    Else
        Me.Label1(11).Caption = "Nº Pedido"
        Me.Label1(10).Caption = "Fecha Pedido"
        Me.Label1(9).Caption = "Trabajador Pedido"

        'Mostramos los datos de la oferta
        Me.Label1(3).Caption = "Nº Oferta"
        Label1(5).Caption = "Fecha Oferta"
    End If
    
    
    PonerCamposForma Me, Data1
    
    Text2(3).Text = PonerNombreDeCod(Text1(3), conAri, "straba", "nomtraba", "codtraba")
    Text2(27).Text = PonerNombreDeCod(Text1(27), conAri, "straba", "nomtraba", "codtraba")
    Text2(28).Text = PonerNombreDeCod(Text1(28), conAri, "straba", "nomtraba", "codtraba")
    Text2(29).Text = PonerNombreDeCod(Text1(29), conAri, "senvio", "nomenvio")
    Text2(12).Text = DevuelveDesdeBDNew(conAri, "sdirec", "nomdirec", "codclien", Text1(4).Text, "N", , "coddirec", Text1(12).Text, "N")
    Text2(17).Text = PonerNombreDeCod(Text1(17), conAri, "sagent", "nomagent")
    Text2(14).Text = PonerNombreDeCod(Text1(14), conAri, "sforpa", "nomforpa")
    
    PonerCamposLineas 'Pone los datos de las tablas de lineas de Ofertas
    
    If EsHistorico Then
        'poner datos de eliminacion
        Text2(32).Text = PonerNombreDeCod(Text1(32), conAri, "straba", "nomtraba", "codtraba")
        Text2(33).Text = PonerNombreDeCod(Text1(33), conAri, "sincid", "nomincid", "codincid")
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
    BloquearText1 Me, Modo
    'Campo Nº Albaran y Tipo Movim. siempre bloqueado, excepto si estamos en modo de busqueda
    b = (Modo <> 1)
    BloquearTxt Text1(0), b, True
    BloquearTxt Text1(30), b
    'Bloquear los campos de Oferta
    BloquearTxt Text1(23), b
    BloquearTxt Text1(24), b
    'Bloquear los campos de Pedido
    For I = 25 To 27
        BloquearTxt Text1(I), b
    Next I
    BloquearTxt Text1(2), b
    'bloquea los datos de venta del TPV (si hay)
    BloquearTxt Text1(38), b
    BloquearTxt Text1(39), b
    
    'Bloquea los campos de Factura (si visibles, ed, si es Rectificativa)
    For I = 35 To 37
        BloquearTxt Text1(I), b
    Next I
  
    '-----  Datos Totales de Factura siempre bloqueado
    For I = 33 To 56
        BloquearTxt Text3(I), True
    Next I
    
    'Referencia produccion tb esta bloqueado
    BloquearTxt Text1(42), b
    
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
    Me.chkFacturar.Enabled = b
    Me.chkFacturarKm.Enabled = b
   
    If vParamAplic.QUE_EMPRESA = 4 Then
        For I = 0 To 3
            Me.chkCarga(I).Enabled = b
        Next
    End If
   
   
    'Si no es modo lineas Boquear los TxtAux
    For I = 0 To txtAux.Count - 1
        BloquearTxt txtAux(I), (Modo <> 5)
    Next I
    BloquearTxt Text2(16), (Modo <> 5)
    BloquearTxt Text2(15), True  'siempre bloqueado. Cuando ponga el articulo habilitara o no
    
    '---------------------------------------------
    b = (Modo <> 0 And Modo <> 2 And Modo <> 5)
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    Me.imgFecha(0).Enabled = b
    Me.imgFecha(44).Enabled = b
    
    For I = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(I).Enabled = b
    Next I
    Me.imgBuscar(1).visible = False
    Me.imgBuscar(7).Enabled = (Modo = 1)
    
              
    'Modo Linea de Albaranes
    Me.Label1(35).visible = (Modo = 5)
    Me.Text2(16).visible = (Modo = 5)
    HectogradoVisible Modo = 5
    BloquearTxt Text2(16), True
    BloquearTxt Text2(15), True
       
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
Dim devuelve As String

    On Error GoTo EDatosOK

    DatosOk = False
    
    'Asignarle el valor del Combo Tipo de Movimiento al texto oculto text1(30)
'    Text1(30).Text = ObtenerCodTipom
    
    b = CompForm(Me, 1) 'Comprobar formato datos ok
    If Not b Then Exit Function
    
    'Comprobar si la referencia del cliente es obligatoria que tenga valor
     If Trim(Text1(4).Text) <> "" Then
        devuelve = DevuelveDesdeBDNew(conAri, "sclien", "referobl", "codclien", Text1(4).Text, "N")
        If devuelve = "1" And Text1(13).Text = "" Then 'Referencia Obligatoria
            MsgBox "La Referencia del Cliente es Obligatoria.", vbInformation
            PonerFoco Text1(13)
            b = False
        End If
    End If
    If Not b Then Exit Function
    
    If Modo = 4 Then
        If DBLet(Data1.Recordset!refproduccion, "N") > 0 Then
            If Not vParamAplic.EsAVAB Then
                'Estamos; en; MORALES.El; CodClien; NO; puede; cambiar
                If Val(Data1.Recordset!CodClien) <> Val(Text1(4).Text) Then
                     MsgBox "Albarán bloqueado. No puede cambiar cliente", vbExclamation
                     b = False
                End If
            End If
        End If
    End If
    DatosOk = b
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLinea(ByRef vCStock As cStock) As Boolean
Dim b As Boolean
Dim I As Byte
    
    On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    b = True

    For I = 0 To 12 'txtAux.Count - 1  'EL 13 no lo meto
        If txtAux(I).Text = "" And I <> 5 Then
            'El campo 5= origpre puede ser nulo (en alb.repar)
            MsgBox "El campo " & txtAux(I).Tag & " no puede ser nulo", vbExclamation
            b = False
            PonerFoco txtAux(I)
            Exit Function
        End If
    Next I
    
    If vParamAplic.QUE_EMPRESA = 2 Then
        'Bodega
        If Not Text2(15).Locked Then
            'Esta habilitado el hectogrado
            If Text2(15).Text = "" Then
                b = False
                MsgBox "Debe indicar hectogrado", vbExclamation
            Else
                If Not PonerFormatoDecimal(Text2(15), 3) Then b = False
            End If
            If Not b Then Exit Function
        End If
    End If
    'Abril 2009
    If Not vUsu.TrabajadorB Then
        If Val(txtAux(0).Text) = vParamAplic.AlmacenB Then
            MsgBox "Almacen incorrecto(2)", vbExclamation
            Exit Function
        End If
    End If
    
    
    'Comprobar si se mueve mover stock (hay stock, o si no hay pero no control de stock)
    If vCStock.MueveStock Then
        b = vCStock.MoverStock(False)
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
    If Index = 16 And KeyAscii = 13 Then  'campo Amliacion Linea y ENTER
        If vParamAplic.QUE_EMPRESA = 2 Then
            KEYpressGnral KeyAscii, Modo, False
        Else
            PonerFocoBtn Me.cmdAceptar
        End If
    End If
    If Index = 15 Then KEYpressGnral KeyAscii, Modo, False
End Sub

Private Sub Text2_LostFocus(Index As Integer)
    If Index = 16 Then
        If (Text2(Index).Locked = False) Then Text2(Index).Text = UCase(Text2(Index).Text)
        
    ElseIf Index = 15 Then
        'Hectogrado
        If Text2(Index).Text <> "" Then
            If Not PonerFormatoDecimal(Text2(Index), 3) Then
                Text2(Index).Text = ""
                PonerFoco Text2(Index)
            Else
                txtAux_LostFocus 7  'para que recalcule
            End If
        End If
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim b As Boolean

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
        Case 11 'Control Nº Series
            BotonNSeries
        Case 12 'Generar Factura Mostrador
                'o Factura Rectificativa (FRT)
            If Me.RecuperarFactu Then
                BotonRecuperarFactura
                
            Else
            
            
                'Septiebmre2009
                If Data2.Recordset Is Nothing Then Exit Sub
                If Data2.Recordset.RecordCount = 0 Then
                    MsgBox "No tiene lineas de albarán", vbExclamation
                    Exit Sub
                End If
            
                
            
                'procedimiento normal
                If Data1.Recordset!Codtipom = "ART" Then
                    'Comprobar nº serie de las facturas rectificativas
                    DevolverNumSeries
                End If
                    
                If Not ComprobarVinculado Then Exit Sub
                
                If Not ComprobarNUmerosDeLote Then Exit Sub
                     
                'Comprobar que esta marcada para facturar
'                If Data1.Recordset!codTipoM <> "ALM" Then Exit Sub
                If Me.chkFacturar.Value = 1 Then
                    NumRegElim = Data1.Recordset.AbsolutePosition
                    
                    'Facturacion de Albaran de Mostrador
                    frmListadoPed.CodClien = CodTipoMov  'utilizamos esta vble para pasarle el tipo de movimiento
                    frmListadoPed.NumCod = Text1(0).Text  'utilizamos esta vble para pasarle el nº albaran
                    frmListadoPed.EstaRecupFact = False
                    AbrirListadoPed (222)
                    
                    PosicionarDataTrasEliminar
                Else
                    MsgBox "El Albaran no esta marcado para facturar", vbInformation
                End If
            End If
        Case 13
            'DAVID
            'Marca los albaranes que esten como NO facturar a facturar
            If Modo = 5 And vParamAplic.PackVtaARticulo And hcoCodTipoM = "ALS" Then
                CadenaDesdeOtroForm = ""
                frmListado.OpcionListado = 100
                frmListado.Show vbModal
                If CadenaDesdeOtroForm <> "" Then
                    Screen.MousePointer = vbHourglass
                    HazInsercionPack
                    Screen.MousePointer = vbDefault
                End If
            Else
                
                Screen.MousePointer = vbHourglass
                MarcarAlbaranes
                Screen.MousePointer = vbDefault
            End If
        Case 15
            'Traer numeros de lote
            'frmAVABasignLotes.Show vbModal
            If Me.Data1.Recordset Is Nothing Then Exit Sub
            If Me.Data1.Recordset.EOF Then Exit Sub
            
            If vParamAplic.QUE_EMPRESA = 4 Then
                CadenaDesdeOtroForm = Data1.Recordset!NumAlbar
                frmListado2.Opcion = 32
                frmListado2.Show vbModal
            Else
                PackingList
            End If
            
        Case 16 'Imprimir Albaran
            mnImprimir_Click
        Case 17    'Salir
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

  
'DesdeRecuperaParaRectificativa:  Para que no inserte el punto verde
Private Function InsertarLinea(numlinea As String, DesdeRecuperaParaRectificativa As Boolean) As Boolean
'Inserta un registro en la tabla de lineas de Albaranes: slialb
Dim SQL As String
Dim vWhere As String
Dim b As Boolean
Dim vCStock As cStock
Dim ImpReciclado As Currency
Dim DentroTRANS As Boolean
Dim Hecto As Currency

    InsertarLinea = False
    SQL = ""
    DentroTRANS = False
    
    'Conseguir el siguiente numero de linea
    vWhere = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas)
    numlinea = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", vWhere)
    Me.cmdAux(0).Tag = numlinea 'Aqui almaceno el Nº linea que acabo de Insertar
    
    Set vCStock = New cStock
    If Not InicializarCStock(vCStock, "S", numlinea) Then Exit Function
    
    If DatosOkLinea(vCStock) Then 'Lineas de Albaranes
        'Inserta en tabla "slialb"
        SQL = "INSERT INTO " & NomTablaLineas
        SQL = SQL & "(codtipom, numalbar,numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel, origpre,codprovex,cajas,PrecioLitro,palets,hectogrado) "
        SQL = SQL & "VALUES ('" & Text1(30).Text & "', " & Val(Text1(0).Text) & ", " & numlinea & ", " & Val(txtAux(0).Text) & ","
        SQL = SQL & DBSet(txtAux(1).Text, "T") & ", " & DBSet(txtAux(2).Text, "T") & ", " & DBSet(Text2(16).Text, "T") & ", "
        SQL = SQL & DBSet(txtAux(3).Text, "N") & ", "
        SQL = SQL & DBSet(txtAux(4).Text, "N") & ", " & DBSet(txtAux(6).Text, "N") & ", "
        SQL = SQL & DBSet(txtAux(7).Text, "N") & ","
        SQL = SQL & DBSet(txtAux(8).Text, "N") & ", "
        SQL = SQL & DBSet(txtAux(5).Text, "T", "N") & ","
        SQL = SQL & DBSet(txtAux(9).Text, "N", "N") & ","
        SQL = SQL & DBSet(txtAux(11).Text, "N", "N") & ","
        SQL = SQL & DBSet(txtAux(12).Text, "N", "N") & ","
        SQL = SQL & DBSet(txtAux(13).Text, "N", "N") & ","
        'hectogrado
        Hecto = 1
        If vParamAplic.QUE_EMPRESA = 2 Then
            If Not Text2(15).Locked Then
                Hecto = ImporteFormateado(Text2(15).Text)
                Hecto = Hecto / 100
            End If
        End If
        SQL = SQL & DBSet(Hecto, "N", "N") & ")"
     Else
        Exit Function
     End If
    
    If SQL <> "" Then
        On Error GoTo EInsertarLinea
        conn.BeginTrans
        DentroTRANS = True
        
        'insertar la linea
        conn.Execute SQL
        
        'si hay control de stock para el articulo actualizar en salmac e insertar en smoval
        'en actualizar stock comprobamos si el articulo tiene control de stock
        b = vCStock.ActualizarStock
        
        
        
        
        'Si ha actualizado el sctock
        If b Then
            If ClienteConTasaReciclado And Not DesdeRecuperaParaRectificativa Then
                If ArticuloConTasaReciclado2(txtAux(1).Text, ImpReciclado) Then
                    'Insertamos la linea del reciclado
                 
                    vWhere = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", vParamAplic.ArtReciclado, "T")
                    SQL = "INSERT INTO " & NomTablaLineas
                    SQL = SQL & "(codtipom, numalbar,numlinea, codalmac, codartic, nomartic, ampliaci, cantidad,  precioar,"
                    SQL = SQL & "dtoline1, dtoline2, importel, origpre) "
                    SQL = SQL & "VALUES ('" & Text1(30).Text & "', " & Val(Text1(0).Text) & ", " & numlinea + 1 & ", " & Val(txtAux(0).Text) & ","
                    SQL = SQL & DBSet(vParamAplic.ArtReciclado, "T") & "," & DBSet(vWhere, "T") & ", Null, "
                    SQL = SQL & DBSet(txtAux(3).Text, "N") & "," 'Cantidad. La misma
                    SQL = SQL & DBSet(ImpReciclado, "N") & ",0,0,"
                    'Importe linea
                    ImpReciclado = ImporteFormateado(txtAux(3).Text) * ImpReciclado
                    SQL = SQL & DBSet(ImpReciclado, "N") & ", 'A')"
                    conn.Execute SQL
                        
                    
                End If 'articulo con sunida reciclado
            End If  'Cliente con tasa reciclado
        End If 'ok actualiza stock
        
        
    
    End If
    Set vCStock = Nothing
    
    If b Then
        conn.CommitTrans
        InsertarLinea = True
    Else
        conn.RollbackTrans
         InsertarLinea = False
    End If
    
    'Si ha ido bien abrimos
    LanzaLote CInt(numlinea)
    
    Exit Function
    
EInsertarLinea:
    If Err.Number <> 0 Then
        InsertarLinea = False
        If DentroTRANS Then conn.RollbackTrans
        MuestraError Err.Number, "Insertar Lineas Albaran" & vbCrLf & Err.Description
    End If

End Function

'>0  Nueva linea. Cogeremos los datos del txtaux
' -1:   dblclick en el datagrid
Private Sub LanzaLote(linea As Integer)
Dim numlinea As Integer

    'Cuando la produccion nueva este en marcha NO se asigna los lotes desde aqui
    If vParamAplic.ProduccionNueva Then
        If MsgBox("No deberia asignarlos desde aqui. Continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If

    'Articulos que NO llevan lote
    If linea < 0 Then
        If Data2.Recordset!codArtic = vParamAplic.ArtReciclado Then Exit Sub
        If Not EsArticuloTrazabilidad(CStr(Data2.Recordset!codArtic)) Then Exit Sub
    Else
        If Not ElArticulo.Trazabilidad Then Exit Sub
    End If

        
    frmFacLotes.vFecha = Data1.Recordset!FechaAlb
    frmFacLotes.vNumalbar = Data1.Recordset!NumAlbar
    frmFacLotes.vCodtipom = Data1.Recordset!Codtipom
    '
    If linea < 0 Then
        frmFacLotes.vCodAlmac = Data2.Recordset!codAlmac
        frmFacLotes.vCantidad = Data2.Recordset!Cantidad
        frmFacLotes.vNumlinea = Data2.Recordset!numlinea
        frmFacLotes.vCodArtic = Data2.Recordset!codArtic
    Else
        frmFacLotes.vCodAlmac = CInt(txtAux(0).Text)
        frmFacLotes.vCantidad = ImporteFormateado(txtAux(3).Text)
        frmFacLotes.vNumlinea = linea
        frmFacLotes.vCodArtic = txtAux(1).Text
    End If
    frmFacLotes.Show vbModal
End Sub

Private Function ModificarLinea() As Boolean
'Modifica un registro en la tabla de lineas de Albaran: slialb
Dim SQL As String
Dim vCStock As cStock
Dim b As Boolean
Dim ImpReciclado As Currency
Dim HaCambiadoCantidad As Boolean


    On Error GoTo EModificarLinea

    ModificarLinea = False
    SQL = ""
    HaCambiadoCantidad = False
    '## LAURA 15/11/2006
    'si se ha modificado el articulo eliminar de la smoval y reestablecer stock
    'Inicilizar la clase para Actualizar los stocks
    
    
    
    Set vCStock = New cStock
    If Not InicializarCStock(vCStock, "S") Then Exit Function
    
    If DatosOkLinea(vCStock) Then
        '#### LAURA 15/11/2006
        conn.BeginTrans
        
'        Set vCStock = New CStock
        'iniciamos la clase con los valores anteriores para deshacer lo q insertamos antes
        b = InicializarCStock(vCStock, "E")
        If b Then
            b = vCStock.DevolverStock2 'eliminamos de smoval y devolvemos stock valores anteriores
            If b Then
                'si se ha modificado el articulo
                If CStr(Data2.Recordset!codArtic) <> txtAux(1).Text Then
                    'si la linea tenia numero de serie vaciar los campos correspondien al albaran venta
                    SQL = "UPDATE sserie SET codclien=" & ValorNulo & ",codtipom=" & ValorNulo & ", fechavta=" & ValorNulo & ",numalbar=" & ValorNulo & ",numline1=" & ValorNulo
                    SQL = SQL & " WHERE codartic=" & DBSet(Data2.Recordset!codArtic, "T") & " and codtipom='" & CodTipoMov & "' and fechavta=" & DBSet(Data1.Recordset!FechaAlb, "F")
                    SQL = SQL & " AND numalbar=" & Data1.Recordset!NumAlbar & " AND numline1=" & Data2.Recordset!numlinea
                    conn.Execute SQL
                End If
            End If
            'ahora leemos los valores nuevos
            If b Then b = InicializarCStock(vCStock, "S")
            'insertamos en smoval y actualizamos stock a los valores nuevos
            vCStock.Cantidad = CSng(ComprobarCero(txtAux(3).Text))
            If b Then b = vCStock.ActualizarStock
    
            'actualizar la linea de Albaran
            If b Then
                SQL = "UPDATE " & NomTablaLineas & " Set codalmac = " & txtAux(0).Text & ", codartic=" & DBSet(txtAux(1).Text, "T") & ", "
                SQL = SQL & "nomartic=" & DBSet(txtAux(2).Text, "T") & ", ampliaci=" & DBSet(Text2(16).Text, "T") & ", "
                SQL = SQL & "cantidad= " & DBSet(txtAux(3).Text, "N") & ", "
                SQL = SQL & "precioar= " & DBSet(txtAux(4).Text, "N") & ", " 'precio
                SQL = SQL & "dtoline1= " & DBSet(txtAux(6).Text, "N") & ", dtoline2= " & DBSet(txtAux(7).Text, "N") & ", "
                SQL = SQL & "importel= " & DBSet(txtAux(8).Text, "N") & ", " 'Importe
                SQL = SQL & "origpre=" & DBSet(txtAux(5).Text, "T", "S") & ","
                SQL = SQL & "codprovex=" & DBSet(txtAux(9).Text, "N", "N") & ","
                'Abril 2009
                SQL = SQL & "cajas=" & DBSet(txtAux(11).Text, "N", "N") & ","
                SQL = SQL & "PrecioLitro=" & DBSet(txtAux(12).Text, "N", "N") & ","
                'Palets
                SQL = SQL & "Palets=" & DBSet(txtAux(13).Text, "N", "N") & ","
                
                'Hectogrado
                ImpReciclado = 1
                If vParamAplic.QUE_EMPRESA = 2 Then
                    If Not Text2(15).Locked Then
                        ImpReciclado = ImporteFormateado(Text2(15).Text)
                        ImpReciclado = ImpReciclado / 100
                    End If
                End If
                SQL = SQL & "hectogrado=" & DBSet(ImpReciclado, "N") & ""
                ImpReciclado = 0  'reestablzco
                
                SQL = SQL & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND numlinea=" & Data2.Recordset!numlinea
                conn.Execute SQL
                
                
                
                'Veo si ha cambiado la cantidad
                HaCambiadoCantidad = vCStock.Cantidad <> DBLet(Data2.Recordset!Cantidad, "N")
                
                'Llegado aqui, si tiene Punto verde(tasa ecologica)
                'Y el cliente tiene tasa recliclado
                If ClienteConTasaReciclado Then
                    If ArticuloConTasaReciclado2(txtAux(1).Text, ImpReciclado) Then
                        
                       'Si el articulo siguiente es PV entoces lo updatearemos
                       SQL = Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND numlinea"
                       'QUITO EL WHERE
                       SQL = Mid(SQL, 8)
                       NumRegElim = Val(DBLet(Data2.Recordset!numlinea, "N")) + 1
                       SQL = DevuelveDesdeBD(conAri, "codartic", "slialb", SQL, CStr(NumRegElim))
                       'En SQL tengo el codarti de la linea SIGUIENTE
                       'SI es punto verde de parametros, supondremos que esta vinculado con la linea que estamos modificando
                       If SQL = vParamAplic.ArtReciclado Then
                       
                            SQL = "UPDATE " & NomTablaLineas & " SET "
                            SQL = SQL & "cantidad= " & DBSet(txtAux(3).Text, "N") & ", "
                            SQL = SQL & "precioar= " & DBSet(ImpReciclado, "N") & ", " 'precio
                            ImpReciclado = ImporteFormateado(txtAux(3).Text) * ImpReciclado
                            SQL = SQL & "importel= " & DBSet(ImpReciclado, "N")  'Importe
                            'WHERE
                            SQL = SQL & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND numlinea=" & NumRegElim
                            conn.Execute SQL
                      End If  'linea siguiente con codarti=puntoverde
                    End If  'articulo con reciclado
                End If ' de cliente con tasa reciclado
                
            End If
'        If SQL <> "" Then
'
'            vCStock.Cantidad = CSng(txtAux(3).Text)
'            b = vCStock.ModificarStock(Data2.Recordset!Cantidad)
'        End If
        End If
        
        
        
    Else
        Set vCStock = Nothing
        Exit Function
    End If
    Set vCStock = Nothing
    
EModificarLinea:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Modificar Lineas Albaran" & vbCrLf & Err.Description
        b = False
    End If
    If b Then
        conn.CommitTrans
        ModificarLinea = True
        If HaCambiadoCantidad Then LanzaLote Val(Data2.Recordset!numlinea)
    Else
        conn.RollbackTrans
         ModificarLinea = False
    End If
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
        Me.cmdRegresar.Cancel = True
        Me.lblIndicador.Caption = "Líneas " & TituloLinea
        PonerFocoBtn Me.cmdRegresar
    Else
        Me.cmdCancelar.Cancel = True
    End If
    
    'Habilitar las opciones correctas del menu segun Modo
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
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
    
    b = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2)
    vDataGrid.Enabled = Not b
    vDataGrid.ScrollBars = dbgAutomatic
    PrimeraVez = False
    Exit Sub
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim I As Byte
    
    On Error GoTo ECargaGrid

    vData.Refresh

    vDataGrid.Columns(0).visible = False
    vDataGrid.Columns(1).visible = False
    vDataGrid.Columns(2).visible = False

    Select Case vDataGrid.Name
        Case "DataGrid1" 'Cod. Almacen
                I = 3
                vDataGrid.Columns(I).Caption = "Alm."
                vDataGrid.Columns(I).Width = 450
                vDataGrid.Columns(I).NumberFormat = "000"
                I = 4
                vDataGrid.Columns(I).Caption = "Articulo"
                vDataGrid.Columns(I).Width = 1550
                I = 5
                vDataGrid.Columns(I).Caption = "Desc. Artículo"
                vDataGrid.Columns(I).Width = 3300

                I = 6
                vDataGrid.Columns(I).visible = False
                
                
                'JUNIO 2011
                'Palets
                I = 7
                vDataGrid.Columns(I).Caption = "Palets"
                vDataGrid.Columns(I).Width = 620
                vDataGrid.Columns(I).Alignment = dbgRight
                
                I = 8
                vDataGrid.Columns(I).Caption = "Cajas"
                vDataGrid.Columns(I).Width = 750
                vDataGrid.Columns(I).Alignment = dbgRight
                
                I = 9
                vDataGrid.Columns(I).Caption = "Unidades"
                vDataGrid.Columns(I).Width = 850
                vDataGrid.Columns(I).Alignment = dbgRight
                vDataGrid.Columns(I).NumberFormat = FormatoImporte
                
                
                
                I = 10
                vDataGrid.Columns(I).Caption = "Precio Ud"
                vDataGrid.Columns(I).Width = 930
                vDataGrid.Columns(I).Alignment = dbgRight
                vDataGrid.Columns(I).NumberFormat = FormatoPrecio
                
                I = 11
                vDataGrid.Columns(I).Caption = "Precio L"
                vDataGrid.Columns(I).Width = 930
                vDataGrid.Columns(I).Alignment = dbgRight
                vDataGrid.Columns(I).NumberFormat = FormatoPrecio
                
                
                I = 12
                vDataGrid.Columns(I).Caption = "OP"
                vDataGrid.Columns(I).Width = 350
                vDataGrid.Columns(I).Alignment = dbgCenter
                
                I = 13
                vDataGrid.Columns(I).Caption = "Dto1"
                vDataGrid.Columns(I).Width = 570
                vDataGrid.Columns(I).Alignment = dbgRight
                vDataGrid.Columns(I).NumberFormat = FormatoDescuento
                
                I = 14
                vDataGrid.Columns(I).Caption = "Dto2"
                vDataGrid.Columns(I).Width = 570
                vDataGrid.Columns(I).Alignment = dbgRight
                vDataGrid.Columns(I).NumberFormat = FormatoDescuento
                
                I = 15
                vDataGrid.Columns(I).Caption = "Importe lin"
                vDataGrid.Columns(I).Width = 1050
                vDataGrid.Columns(I).Alignment = dbgRight
                vDataGrid.Columns(I).NumberFormat = FormatoImporte
                
                I = 16
                vDataGrid.Columns(I).Caption = "Prov"
                vDataGrid.Columns(I).Width = 600
                vDataGrid.Columns(I).Alignment = dbgRight
                
                I = 17
                vDataGrid.Columns(I).Caption = "Nom. prove"
                vDataGrid.Columns(I).Width = 1000
                
                
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



    'Cmabio importante.
    'Se ha puesto caja(11) antes que cantidad(3)
    'Era mas importante matener los index que tocar esto
    'Con lo caul, estamos tocando esto
    






    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For I = 0 To txtAux.Count - 1 'TextBox
            txtAux(I).Top = 290
            txtAux(I).visible = visible
        Next I
        cmdAux(0).visible = visible
        cmdAux(1).visible = visible
        cmdAux(9).visible = visible
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
                Case 0, 1, 2
                    txtAux(I).Text = DataGrid1.Columns(I + 3).Text
                
                'Son cantidad, precio,caja, precioltro
                Case 3
                    txtAux(I).Text = DataGrid1.Columns(9).Text
                Case 4
                    txtAux(I).Text = DataGrid1.Columns(10).Text
                Case 11
                    txtAux(I).Text = DataGrid1.Columns(8).Text
                Case 12
                    txtAux(I).Text = DataGrid1.Columns(11).Text
                    
                Case 13
                    txtAux(I).Text = DataGrid1.Columns(7).Text
                Case Else
                    'Cajas y precio litro. Datagrid 8 y 10
                    txtAux(I).Text = DataGrid1.Columns(I + 7).Text
                End Select
                txtAux(I).Locked = False
            Next I
        End If
        
        cmdAux(0).Enabled = True
        cmdAux(1).Enabled = True
        cmdAux(9).Enabled = True
               
        'El Campo de Origen del precio se actualiza por programa al modificar el precio
        BloquearTxt txtAux(5), True
        'El campo Importe es calculado y lo bloqueamos.
        '##### Laura Recuperar facturas ALZIRA
'        BloquearTxt txtAux(8), True
        BloquearTxt txtAux(8), Not (Me.RecuperarFactu)
        '#####
        
        'Fijamos altura(Height) y posición Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid1, 10)
        
        For I = 0 To txtAux.Count - 1
            txtAux(I).Top = alto
            txtAux(I).Height = DataGrid1.RowHeight
        Next I
        cmdAux(0).Top = alto
        cmdAux(1).Top = alto
        cmdAux(9).Top = alto
        cmdAux(0).Height = DataGrid1.RowHeight
        cmdAux(1).Height = DataGrid1.RowHeight
        cmdAux(9).Height = DataGrid1.RowHeight
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'Cod. Almac
        txtAux(0).Left = DataGrid1.Left + 330
        txtAux(0).Width = DataGrid1.Columns(3).Width - 160
        cmdAux(0).Left = txtAux(0).Left + txtAux(0).Width - 40
        'Cod Artic
        txtAux(1).Left = cmdAux(0).Left + cmdAux(0).Width + 20
        txtAux(1).Width = DataGrid1.Columns(4).Width - 160
        cmdAux(1).Left = txtAux(1).Left + txtAux(1).Width - 50
        'Nom Artic
        txtAux(2).Left = cmdAux(1).Left + cmdAux(1).Width
        txtAux(2).Width = DataGrid1.Columns(5).Width - 10
        
        'Junio 2011
        'Palets
        txtAux(13).Left = txtAux(2).Left + txtAux(2).Width + 10
        txtAux(13).Width = DataGrid1.Columns(7).Width - 10
        
        
        
        'CAMBIO mentado arriba
        'Cajas
        txtAux(11).Left = txtAux(13).Left + txtAux(13).Width + 10
        txtAux(11).Width = DataGrid1.Columns(8).Width - 10
        
        'Abril 2009
        'Hemos añadido cajas y precio litro
        txtAux(3).Left = txtAux(11).Left + txtAux(11).Width + 10
        txtAux(3).Width = DataGrid1.Columns(9).Width - 10
        
        
        txtAux(4).Left = txtAux(3).Left + txtAux(3).Width + 10
        txtAux(4).Width = DataGrid1.Columns(10).Width - 10
        
        txtAux(12).Left = txtAux(4).Left + txtAux(4).Width + 10
        txtAux(12).Width = DataGrid1.Columns(11).Width - 10
        
        
        
        
        'Precio, Dto1, Dto2, Precio
        txtAux(5).Left = txtAux(12).Left + txtAux(12).Width + 10
        txtAux(5).Width = DataGrid1.Columns(12).Width - 10
        
        For I = 6 To 10
            txtAux(I).Left = txtAux(I - 1).Left + txtAux(I - 1).Width + 10
            txtAux(I).Width = DataGrid1.Columns(7 + I).Width - 10
        Next I
        
        'El boton 3 lo superpongo un poquito
        cmdAux(9).Left = txtAux(10).Left - 15
        
        'Los ponemos Visibles o No
        '--------------------------
        For I = 0 To txtAux.Count - 1
            txtAux(I).visible = visible
        Next I
        cmdAux(0).visible = visible
        cmdAux(1).visible = visible
        cmdAux(9).visible = visible
    End If
End Sub


Private Sub TxtAux_Change(Index As Integer)
    If Index = 4 And ModificaLineas = 2 Then 'Precio y Modo Borrar Lineas
        txtAux(5).Text = "M"
    End If
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
Dim cadkey As Integer
        
    cadkey = ObtenerCadKey(kCampo, Index)
    kCampo = Index
    ConseguirFocoLin txtAux(Index), cadkey
'    If VaciarTxtAnterior Then
'        VaciarTxtAnterior = False
'        txtAnterior = ""
'    Else
'        txtAnterior = txtAux(Index).Text
'    End If
    
End Sub

Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Index = 0 And KeyCode = 38 Then Exit Sub 'campo almacen y flecha arriba
    If Index = 1 Then
        If Modo = 5 And ModificaLineas = 1 Then
            'Insertando linea albaran
            
            If KeyCode = 43 Or KeyCode = 107 Then
                KeyCode = 0
                PulsadoMas2 = True
                
                cmdAux_Click 1
            
            Else
                'Ha pulsado F2
                If KeyCode = 113 Then Me.DataGrid1.Columns(4).Caption = "EAN"
            End If
        End If
    End If
    KEYdown KeyCode
End Sub


Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim devuelve As String, cadMen As String
Dim codTarif As String
Dim CPrecioFact As CPreciosFact
Dim NumCajas As Long, RestoUnid As Long
Dim OrigP As String 'De donde viene el precio
Dim Cantidad As String
Dim vCStock As cStock
Dim b As Boolean
Dim okArticulo As Boolean
Dim ImporteConhectogrado As Currency


    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub
    
    
    If PulsadoMas2 Then
        'Para que cuando pulse el mas abra el form
        PulsadoMas2 = False
        txtAux(Index).Text = Mid(txtAux(Index).Text, 1, Len(txtAux(Index).Text) - 1)
        Exit Sub
    End If
    
    'NO ha cambiado nada
    If txtAnterior = txtAux(Index).Text Then
        '
       ' Exit Sub
    End If
    
    Select Case Index
        Case 0 'Cod ALMACEN
            'Comprobar que existe el almacen
            devuelve = PonerAlmacen(txtAux(Index).Text)
            txtAux(Index).Text = devuelve
            If devuelve = "" Then PonerFoco txtAux(Index)

        Case 1 'Cod. ARTICULO
            If txtAux(Index).Text = "" Then
                txtAux(2).Text = ""
                Exit Sub
            End If
        
            If txtAux(0).Text = "" Then
                MsgBox "Debe seleccionar un almacen.", vbInformation
                PonerFoco txtAux(0)
                Exit Sub
            End If
        
            devuelve = ""
            If ModificaLineas = 2 Then
                If Not Data2.Recordset.EOF Then devuelve = Data2.Recordset!codArtic
            End If
            Cantidad = txtAux(9).Text
            
            If Me.DataGrid1.Columns(4).Caption = "EAN" Then
                'Ha pulsado F2, para meter, en lugar del codigo del articulo, el EAN
                okArticulo = PonerArticuloEAN(txtAux(1), txtAux(2), txtAux(0).Text, CodTipoMov, ModificaLineas, devuelve, , Cantidad)
            Else
                okArticulo = PonerArticulo(txtAux(1), txtAux(2), txtAux(0).Text, CodTipoMov, ModificaLineas, devuelve, , Cantidad)
            End If
            If Not okArticulo Then
                If Me.DataGrid1.Columns(4).Caption = "EAN" Then txtAux(1).Text = ""
                PonerFoco txtAux(Index)
            Else
            
                'Leemos el articulo
                If ElArticulo Is Nothing Then Set ElArticulo = New CArticulo
                
                If ElArticulo.Codigo <> txtAux(1).Text Then ElArticulo.LeerDatos txtAux(1).Text
                
                If vParamAplic.QUE_EMPRESA = 2 Then
                    devuelve = "if(codtipar='05',1,0)+if(codfamia=6,1,0)" 'Los 05 o la familia 6
                    devuelve = DevuelveDesdeBD(conAri, devuelve, "sartic", "codartic", ElArticulo.Codigo, "T")
                    If devuelve = "" Then devuelve = 0
                    BloquearTxt Text2(15), Val(devuelve) = 0
                End If
                
                'Por si acaso ha cambiado el articulo
                PrecioUdLitro True
                CantidadCajas True
                
                
                b = (Me.ActiveControl.Name = "txtAux")
                If b Then b = (Me.ActiveControl.Index = 0)
                If Not b Then
'                    If txtAux(2).Locked Then PonerFoco txtAux(3)
                Else
                    PonerFoco txtAux(0)
                End If
                
                'Si  ha cambiado el articulo, el proveedore
                If txtAux(9).Text = "" Then
                    txtAux(9).Text = Cantidad
                    'Fuerzo el lostfocus para que carge el proveedor
                    txtAux_LostFocus 9
                End If
            End If
        
        Case 2 'Nombre Articulo
           If txtAux(Index).Locked = False Then txtAux(Index).Text = UCase(txtAux(Index).Text)
        
        Case 3 'CANTIDAD

        
            If PonerFormatoDecimal(txtAux(Index), 1) Then  'Tipo 1: Decimal(12,2)
                'Si es factura rectifica la cantidad solo puede ser negativa
                If CodTipoMov = "ART" Then
                    If CCur(txtAux(Index)) >= 0 Then
                        MsgBox "En facturas rectificativas la cantidad debe ser negativa.", vbExclamation
                        PonerFoco txtAux(Index)
                        Exit Sub
                    End If
                End If
            
                'Ponemos las cajas
                CantidadCajas True
            
                'Comprobar si hay suficiente stock
                Set vCStock = New cStock
                If Not InicializarCStock(vCStock, "S") Then Exit Sub
                If vCStock.MueveStock Then 'Comprobar si el articulo mueve stock: tiene control de stock y no es instalacion
                  If Not vCStock.MoverStock(False) Then
                    PonerFoco txtAux(Index)
                    Set vCStock = Nothing
                    Exit Sub
                  End If
                End If
                
                b = False
                If Modo = 5 Then 'Modo lineas
                    If ModificaLineas = 1 Then 'insertar linea
                        b = True
                    ElseIf ModificaLineas = 2 Then 'modificar linea
                        If Data2.Recordset!codArtic <> txtAux(1).Text Then b = True
                    End If
                End If
                
                If b Then 'Modo Insertar en Mto Lineas
                    'Obtener el precio correspondiente y los descuentos
                    'Comprobar si el articulo se vende por cajas antes de entrar a la función
                    devuelve = DevuelveDesdeBDNew(conAri, "sartic", "unicajas", "codartic", txtAux(1).Text, "T")
                    If devuelve <> "" Then
                        Set CPrecioFact = New CPreciosFact
                        'Si se puede vender por cajas(devuelve>1) poner numero de cajas en una linea con el
                        'precio de caja, y otra linea con el resto unidades un precio unidad
                        Cantidad = txtAux(Index).Text
                        NumCajas = CPrecioFact.ObtenerNumCajas(Cantidad, devuelve)
                        RestoUnid = CLng(ComprobarCero(Cantidad)) - NumCajas * CInt(devuelve)
                        'Obtenemos la Tarifa del Cliente
                        codTarif = DevuelveDesdeBDNew(conAri, "sclien", "codtarif", "codclien", Text1(4).Text, "N")
                        CPrecioFact.CodigoLista = codTarif
                            
                        CPrecioFact.CodigoArtic = txtAux(1).Text
                        CPrecioFact.CodigoClien = Text1(4).Text
                        PorCaja = (NumCajas > 0)
                        Precio = CPrecioFact.ObtenerPrecio(PorCaja, Text1(1).Text, OrigP)
                        'Si PorCaja vuelve de ObtenerPrecio a false se calcula con precio unidad aunque NumCajas>0
                        'Ya que a regresado con pvp del Articulo
                        If PorCaja And NumCajas > 0 And RestoUnid > 0 Then
                            cadMen = "El Artículo puede venderse por Cajas (" & devuelve & "uds. por Caja)." & vbCrLf
                            cadMen = cadMen & vbCrLf & "Inserte dos Lineas:   "
                            cadMen = cadMen & vbCrLf & "   Linea 1:  " & NumCajas * CInt(devuelve) & " uds a Precio Caja"
                            cadMen = cadMen & vbCrLf & "   Linea 2:  " & CInt(Cantidad) - NumCajas * CInt(devuelve) & " uds a Precio Unidad"
                            MsgBox cadMen, vbInformation
                        Else
                            If (txtAux(4).Text = "") Or (txtAux(4).Text <> "" And ModificaLineas = 2 And b) Then
                                txtAux(4).Text = Precio
                                txtAux(5).Text = OrigP 'De donde viene el precio
                                PrecioUdLitro True
                            End If
                            PonerFormatoDecimal txtAux(4), 2
                            If txtAux(6).Text = "" Then txtAux(6).Text = CPrecioFact.Descuento1
                            PonerFormatoDecimal txtAux(6), 4
                            If txtAux(7).Text = "" Then txtAux(7).Text = CPrecioFact.Descuento2
                            PonerFormatoDecimal txtAux(7), 4
                            
                            
                            'Pondere el foco en precio litro si es mayor que un litro
                            RestoUnid = 4
                            If Not (ElArticulo Is Nothing) Then
                                If ElArticulo.LitrosxUd > 1 Then RestoUnid = 12
                            End If
                            
                            PonerFoco txtAux(RestoUnid)
                            RestoUnid = 0
                        End If
'                            PonerFoco txtAux(Index + 1)
                        Set CPrecioFact = Nothing
                    End If
                End If
                Set vCStock = Nothing
            End If
            
            
        Case 4 'Precio
             If txtAux(Index).Text <> "" Then
                PonerFormatoDecimal txtAux(Index), 2 'Tipo 2: Decimal(10,4)
                If ModificaLineas = 1 Then
                    If CSng(txtAux(Index).Text) <> CSng(ComprobarCero(Precio)) Then txtAux(5).Text = "M"
                End If
            End If
            PrecioUdLitro True
            
        Case 6, 7 'Descuentos
            PonerFormatoDecimal txtAux(Index), 4 'Tipo 4: Decimal(4,2)
        Case 8 'Importe Linea
            PonerFormatoDecimal txtAux(Index), 1 'Tipo 3: Decimal(12,2)
            
            
        Case 9
            'Cod proveeee
            If txtAux(9).Text = "" Then
                devuelve = ""
            Else
                If Not IsNumeric(txtAux(9).Text) Then
                    MsgBox "Campo proveedor debe ser numérico", vbExclamation
                    devuelve = ""
                Else
                        
                    devuelve = DevuelveDesdeBDNew(conAri, "sprove", "nomprove", "codprove", txtAux(9).Text, "N")
                    If devuelve = "" Then MsgBox "No existe el proveedor: " & txtAux(9).Text, vbExclamation
                End If
                If devuelve = "" Then
                    txtAux(9).Text = ""
                    PonerFoco txtAux(9)
                End If
            End If
            txtAux(10).Text = devuelve
            
    Case 11
        'cajas
        txtAnterior = txtAux(3).Text
        If txtAux(11).Text <> "" Then
            If Not PonerFormatoEntero(txtAux(11)) Then
                txtAux(11).Text = ""
                PonerFoco txtAux(11)
            Else
                CantidadCajas False
                If txtAnterior <> txtAux(3).Text Then PonerFoco txtAux(3)
                
            End If
        Else
            txtAux(3).Text = ""
        End If
        
    Case 12
        
        txtAnterior = txtAux(4).Text
        If txtAux(12).Text <> "" Then
            If Not PonerFormatoDecimal(txtAux(12), 2) Then
                txtAux(12).Text = ""
                PonerFoco txtAux(12)
            Else
               
                PrecioUdLitro False
                If txtAnterior <> txtAux(4).Text Then PonerFoco txtAux(4)

            End If
        Else
            txtAux(4).Text = ""
        End If
       
    Case 13
    
        If txtAux(Index).Text = "" Then Exit Sub
    
       If Not PonerFormatoEntero(txtAux(Index)) Then
            txtAux(Index).Text = ""
            PonerFoco txtAux(Index)
            Exit Sub
        End If
        
        
        'Si tiene articulo y NO tiene las cajas puestas
        If txtAux(11).Text = "" And txtAux(1).Text <> "" Then
            devuelve = DevuelveDesdeBD(conAri, "pal_udbas * pal_udalt", "sarti4", "codartic", txtAux(1).Text, "T")
            If devuelve = "" Then devuelve = "0"
            RestoUnid = Val(devuelve) * CInt(txtAux(Index).Text)
            txtAux(11).Text = RestoUnid
        End If

        
    End Select
    
    
    
     If (Index = 3 Or Index = 4 Or Index = 6 Or Index = 7) Then 'Cant., Precio, Dto1, Dto2
'        If Trim(TxtAux(3).Text) = "" Or Trim(TxtAux(4).Text) = "" Then Exit Sub
'        If Trim(TxtAux(6).Text) = "" Or Trim(TxtAux(7).Text) = "" Then Exit Sub
        If txtAux(1).Text = "" Then Exit Sub
        devuelve = CalcularImporte(txtAux(3).Text, txtAux(4).Text, txtAux(6).Text, txtAux(7).Text, vParamAplic.TipoDtos)
        
        ImporteConhectogrado = 1
        If vParamAplic.QUE_EMPRESA = 2 Then
            If Not Text2(15).Locked Then
                If Text2(15).Text <> "" Then
                    ImporteConhectogrado = ImporteFormateado(Text2(15).Text)
                    ImporteConhectogrado = ImporteConhectogrado / 100
                End If
             End If
        End If
        txtAux(8).Text = ImporteFormateado(devuelve) * ImporteConhectogrado
        PonerFormatoDecimal txtAux(8), 1
    End If
End Sub


Private Sub BotonMtoLineas(numTab As Integer, cad As String)
    Me.SSTab1.Tab = numTab
    TituloLinea = cad
    ModificaLineas = 0
    
        If vParamAplic.ArtReciclado <> "" Then
            ClienteConTasaReciclado = Val(DevuelveDesdeBD(conAri, "tasareciclado", "sclien", "codclien", Text1(4).Text)) = 1
        Else
            ClienteConTasaReciclado = False
        End If
    
    
    PonerModo 5
    PonerBotonCabecera True
End Sub


Private Function Eliminar(NumAlbElim As Long) As Boolean
Dim SQL As String
Dim b As Boolean
Dim vTipoMov As CTiposMov
Dim MenError As String

    On Error GoTo FinEliminar
    conn.BeginTrans
    
    SQL = ObtenerWhereCP(False)
    
    'Reestablecer el stock en la tabla salmac a partir de todas las lineas del albaran
    MenError = "Restableciendo stocks de almacen."
    b = ReestablecerStock
    

    If b Then
        'Los numeros de lote
        'Conn.Execute "DELETE FROM slialblotes WHERE numalbar=" & Data1.Recordset!NumAlbar & "  AND codtipom='" & CodTipoMov & "'"
        If Not Data2.Recordset Is Nothing Then
            If Data2.Recordset.RecordCount > 0 Then
                Data2.Recordset.MoveFirst
                While Not Data2.Recordset.EOF
                    'Los numeros de lote
                    If EsArticuloTrazabilidad(CStr(Data2.Recordset!codArtic)) Then _
                         EliminarLineaProcesoLotaje CStr(Data2.Recordset!codArtic), CInt(Data2.Recordset!numlinea), CInt(Data2.Recordset!codAlmac)
                
                
                    Data2.Recordset.MoveNext
                Wend
            End If
        End If
    
        'eliminamos de albaranes y pasamos al historico
        b = ActualizarElTraspaso(MenError, SQL, CodTipoMov, cadList)
        
        If b Then
            MenError = "Actualizando numeros de serie."
            'Actualizar los posibles num. serie de ese albaran. vaciar los campos
            SQL = "UPDATE  sserie SET codclien=" & ValorNulo & ", codtipom=" & ValorNulo & ","
            SQL = SQL & " fechavta=" & ValorNulo & ", numalbar=" & ValorNulo & ", numline1=" & ValorNulo
            SQL = SQL & " WHERE codtipom='" & CodTipoMov & "' AND numalbar=" & Data1.Recordset!NumAlbar & " AND fechavta=" & DBSet(Data1.Recordset!FechaAlb, "F")
            conn.Execute SQL
            
            
            'Devolvemos contador, si no estamos actualizando
            Set vTipoMov = New CTiposMov
            b = CBool(vTipoMov.DevolverContador(CodTipoMov, NumAlbElim))
            Set vTipoMov = Nothing
        End If
        

    End If
        
FinEliminar:
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, MenError, Err.Description
    End If
    If Not b Then
        conn.RollbackTrans
    Else
        conn.CommitTrans
    End If
    Eliminar = b
End Function


Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next

    CargaGrid DataGrid1, Data2, False
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = "(" & ObtenerWhereCP(False) & ")"
         vWhere = Replace(vWhere, NombreTabla & ".", "")
         If SituarDataMULTI(Data1, vWhere, Indicador) Then
'         If SituarDataGral(Data1, Text1(30).Text, "T", Text1(0).Text, "N", Indicador) Then
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
    
    SQL = " " & NombreTabla & ".codtipom= '" & Text1(30).Text & "' and " & NombreTabla & ".numalbar= " & Val(Text1(0).Text)
    If EsHistorico Then SQL = SQL & " AND " & NombreTabla & ".fechaalb=" & DBSet(Text1(1).Text, "F")
    If conWhere Then SQL = " WHERE " & SQL
    ObtenerWhereCP = SQL
    
    If Err.Number <> 0 Then Err.Clear
End Function


Private Function MontaSQLCarga(enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim SQL As String
    
    
    'Enero 2008. David
    'Para la trazabilidad con repescto al codproveedor en las lineas
    'Abril 2009
    'Aceites.  Cajas, y precio litro
    SQL = "SELECT codtipom, numalbar, numlinea, codalmac, codartic, nomartic, ampliaci,"
    'SQL = SQL & "cantidad, cajas, precioar, preciolitro,"
    'SQL = SQL & "  cajas,cantidad, precioar, preciolitro,"
    SQL = SQL & " palets, cajas,cantidad, precioar, preciolitro,"
    SQL = SQL & "origpre, dtoline1, dtoline2, importel ,codprovex,nomprove "
    SQL = SQL & " FROM " & NomTablaLineas
    'traza
    SQL = SQL & " LEFT JOIN sprove on codprovex=codprove "
    If enlaza Then
        SQL = SQL & " " & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
        If EsHistorico Then SQL = SQL & " and fechaalb='" & Format(Text1(1).Text, FormatoFecha) & "'"
    Else
        SQL = SQL & " WHERE numalbar = -1"
    End If
    SQL = SQL & " Order by codtipom, numalbar, numlinea"
    MontaSQLCarga = SQL
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean

        b = ((Modo = 2) Or (Modo = 5 And ModificaLineas = 0))
        'Insertar
        Toolbar1.Buttons(5).Enabled = (b Or Modo = 0) And Not EsHistorico
        Me.mnNuevo.Enabled = (b Or Modo = 0) And Not EsHistorico
        'Modificar
        Toolbar1.Buttons(6).Enabled = b And Not EsHistorico
        Me.mnModificar.Enabled = b And Not EsHistorico
        'eliminar
        Toolbar1.Buttons(7).Enabled = b And Not EsHistorico
        Me.mnEliminar.Enabled = b And Not EsHistorico
            
        b = (Modo = 2) And Not EsHistorico
        'Mantenimiento lineas
        Toolbar1.Buttons(10).Enabled = b
        Me.mnLineas.Enabled = b
        
        'Nº Series
        Toolbar1.Buttons(11).Enabled = b And Not EsHistorico
        
        'Generar Factura
        'DAVID###
        'Antes:
        'Toolbar1.Buttons(12).Enabled = b And (CodTipoMov = "ALM" Or CodTipoMov = "ART")
        'Ahora.  Cualquier tipo se puede generar la factura
        Toolbar1.Buttons(12).Enabled = b
        
        
        If Modo = 5 And vParamAplic.PackVtaARticulo And hcoCodTipoM = "ALS" Then
            Toolbar1.Buttons(13).ToolTipText = "Añadir PACK"
        Else
            Toolbar1.Buttons(13).ToolTipText = "Marcar facturar"
        End If
        
        'Imprimir
        Toolbar1.Buttons(15).Enabled = ((Modo = 2) And CodTipoMov <> "ALM")
        Me.mnImprimir.Enabled = ((Modo = 2) And CodTipoMov <> "ALM")
        
        
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


Private Function InsertarOferta(vSQL As String, vTipoMov As CTiposMov) As Boolean
Dim MenError As String
Dim bol As Boolean, Existe As Boolean
Dim cambiaSQL As Boolean
Dim devuelve As String

    On Error GoTo EInsertarOferta
    
    bol = True
    
    cambiaSQL = False
    'Comprobar si mientras tanto se incremento el contador de Pedidos
    'para ello vemos si existe una oferta con ese contador y si existe la incrementamos
    Do
        devuelve = DevuelveDesdeBDNew(conAri, NombreTabla, "numalbar", "codtipom", Text1(30).Text, "T", , "numalbar", Text1(0).Text, "N")
        If devuelve <> "" Then
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
    MenError = "Error al insertar en la tabla Cabecera de Albaranes (" & NombreTabla & ")."
    conn.Execute vSQL, , adCmdText
    
    'Actualizar los datos del cliente si es de varios
    If EsDeVarios Then
        'Si es cliente de varios actualizar datos cliente en tabla:sclvar
        MenError = "Modificando datos cliente varios"
        bol = ActualizarClienteVarios(Text1(4).Text, Text1(6).Text)
    End If
           
    If bol Then
        'Actualizar el campo fechamov (ult. movimiento) de la tabla de clientes (sclien)
        MenError = "Actualizando Fecha Movimiento del Cliente."
        bol = ActualizarFecMovCliente
        
        MenError = "Error al actualizar el contador del Pedido."
    '    bol = vTipoMov.IncrementarContador("REG")
        vTipoMov.IncrementarContador (CodTipoMov)
    End If
    
EInsertarOferta:
        If Err.Number <> 0 Then
            MenError = "Insertando Albaran." & vbCrLf & "----------------------------" & vbCrLf & MenError
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

    For I = 4 To 17
        Text1(I).Text = ""
    Next I
    Text2(12).Text = ""
    Text2(14).Text = ""
    Text2(17).Text = ""
    Me.cboFacturacion.ListIndex = -1
End Sub
    

Private Function EliminarLinea() As Boolean
Dim vCStock As cStock
Dim SQL As String
Dim b As Boolean
Dim ImpReciclado As Currency



    EliminarLinea = False
    
    'Construir la SQL para eliminar la linea de la tabla "slialb"
    SQL = "Delete from " & NomTablaLineas & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
    SQL = SQL & " and numlinea=" & Data2.Recordset!numlinea
    
    
    
    'Inicilizar la clase para Actualizar los stocks
    Set vCStock = New cStock
    If Not InicializarCStock(vCStock, "E") Then Exit Function
    
    On Error GoTo EEliminarLinea
    
    conn.BeginTrans
    conn.Execute SQL 'Eliminar linea
    b = vCStock.DevolverStock2
    Set vCStock = Nothing

    If b Then
        'Ha borrado la linea y ha devuelvto correctamente el sctock
                   'Llegado aqui, si tiene Punto verde(tasa ecologica)
                'Y el cliente tiene tasa recliclado
                If ClienteConTasaReciclado Then
                    SQL = CStr(Data2.Recordset!codArtic)
                    If ArticuloConTasaReciclado2(SQL, ImpReciclado) Then
                        
                       'Si el articulo siguiente es PV entoces lo updatearemos
                       SQL = Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND numlinea"
                       'QUITO EL WHERE
                       SQL = Mid(SQL, 8)
                       NumRegElim = Val(DBLet(Data2.Recordset!numlinea, "N")) + 1
                       SQL = DevuelveDesdeBD(conAri, "codartic", "slialb", SQL, CStr(NumRegElim))
                       'En SQL tengo el codarti de la linea SIGUIENTE
                       'SI es punto verde de parametros, supondremos que esta vinculado con la linea que estamos modificando
                       If SQL = vParamAplic.ArtReciclado Then
                       
                            SQL = "DELETE FROM " & NomTablaLineas
                            'WHERE
                            SQL = SQL & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND numlinea=" & NumRegElim
                            conn.Execute SQL
                      End If  'linea siguiente con codarti=puntoverde
                    End If  'articulo con reciclado
                End If ' de cliente con tasa reciclado
                
    End If


    'si la linea tenia numero de serie vaciar los campos correspondien al albaran venta
    SQL = "UPDATE sserie SET codclien=" & ValorNulo & ",codtipom=" & ValorNulo & ", fechavta=" & ValorNulo & ",numalbar=" & ValorNulo & ",numline1=" & ValorNulo
    SQL = SQL & " WHERE codartic=" & DBSet(Data2.Recordset!codArtic, "T") & " and codtipom='" & CodTipoMov & "' and fechavta=" & DBSet(Data1.Recordset!FechaAlb, "F")
    SQL = SQL & " AND numalbar=" & Data1.Recordset!NumAlbar & " AND numline1=" & Data2.Recordset!numlinea
    conn.Execute SQL
    
    
    'Los numeros de lote
    If EsArticuloTrazabilidad(CStr(Data2.Recordset!codArtic)) Then _
        EliminarLineaProcesoLotaje CStr(Data2.Recordset!codArtic), CInt(Data2.Recordset!numlinea), CInt(Data2.Recordset!codAlmac)
        
        
    
    
EEliminarLinea:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Linea Albaran " & vbCrLf & Err.Description
        b = False
    End If
    
    If b Then
        conn.CommitTrans
        EliminarLinea = True
    Else
        conn.RollbackTrans
         EliminarLinea = False
    End If
End Function



Private Function EliminarLineaProcesoLotaje(codArtic As String, numlinea As Integer, codAlmac As Integer) As Boolean
Dim cP As cPartidas
Dim J As Integer
Dim Can As Currency
Dim Lista As Collection
Dim RL As ADODB.Recordset
Dim cLot As cLotaje
Dim SQL As String

On Error GoTo EEliminarLineaProcesoLotaje
        EliminarLineaProcesoLotaje = False
        SQL = "SELECT * FROM slialblotes WHERE numalbar=" & Data1.Recordset!NumAlbar & " AND numlinea=" & Data2.Recordset!numlinea & " AND codtipom='" & CodTipoMov & "'"
        Set RL = New ADODB.Recordset
        Set Lista = New Collection
        
        RL.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RL.EOF
            'Reestablecemos la cantidad en partidas
            Lista.Add RL!numLote & "|" & RL!Cantidad & "|"
            
            'Sig
            RL.MoveNext
        Wend
        RL.Close
        
        
        If vParamAplic.Produccion Then
            Set cP = New cPartidas
            For J = 1 To Lista.Count
                SQL = RecuperaValor(Lista.Item(J), 2)
                Can = CCur(SQL)
                SQL = RecuperaValor(Lista.Item(J), 1)
                If cP.LeerDesdeArticulo(CStr(codArtic), codAlmac, SQL) Then
                    cP.IncrementarCantidad Can
                Else
                    MsgBox "Partida no encontrada: " & codArtic & " " & SQL, vbExclamation
                End If
            Next
        End If
        
        SQL = "DELETE FROM slialblotes WHERE numalbar=" & Data1.Recordset!NumAlbar & " AND numlinea=" & numlinea & " AND codtipom='" & CodTipoMov & "'"
        conn.Execute SQL
        

        Set cLot = New cLotaje
        cLot.DetaMov = CodTipoMov
        cLot.Fechamov = Data1.Recordset!FechaAlb
        cLot.codAlmac = codAlmac
        cLot.codArtic = codArtic
        cLot.Documento = Data1.Recordset!NumAlbar
        cLot.LineaDocu = numlinea
        cLot.EliminarMovimArticulosLotaje True
        
    
    
EEliminarLineaProcesoLotaje:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set RL = Nothing
    Set cLot = Nothing
    Set Lista = Nothing
End Function


Private Function InicializarCStock(ByRef vCStock As cStock, TipoM As String, Optional numlinea As String) As Boolean
    On Error Resume Next

    vCStock.tipoMov = TipoM
    vCStock.DetaMov = CodTipoMov
    vCStock.Trabajador = CLng(Text1(4).Text) 'guardamos el cliente del albaran
    vCStock.Documento = Text1(0).Text 'Nº Albaran
    vCStock.Fechamov = Text1(1).Text 'Fecha del Albaran
    
    If Text1(60).Text <> "" Then vCStock.HoraMov = Text1(1).Text & " " & Text1(60).Text
        
    
    '1=Insertar, 2=Modificar
    If ModificaLineas = 1 Or (ModificaLineas = 2 And TipoM = "S") Then
        vCStock.codArtic = txtAux(1).Text
        vCStock.codAlmac = CInt(txtAux(0).Text)
        If ModificaLineas = 1 Then '1=Insertar
            vCStock.Cantidad = CSng(ComprobarCero(txtAux(3).Text))
        Else '2=Modificar(Debe haber en stock la diferencia)
            If Data2.Recordset!codArtic = txtAux(1).Text Then
                vCStock.Cantidad = CSng(ComprobarCero(txtAux(3).Text)) - Data2.Recordset!Cantidad
            Else
                vCStock.Cantidad = CSng(ComprobarCero(txtAux(3).Text))
            End If
        End If
        vCStock.Importe = CCur(ComprobarCero(txtAux(8).Text))
    Else
        vCStock.codArtic = Data2.Recordset!codArtic
        vCStock.codAlmac = CInt(Data2.Recordset!codAlmac)
        vCStock.Cantidad = CSng(Data2.Recordset!Cantidad)
        vCStock.Importe = CCur(Data2.Recordset!ImporteL)
    End If
    If ModificaLineas = 1 Then
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


Private Function ReestablecerStock() As Boolean
Dim vCStock As cStock
Dim b As Boolean

    On Error GoTo ERestablecer
    
    ReestablecerStock = False
    b = True
    If Data2.Recordset.RecordCount > 0 Then
       Data2.Refresh
       Data2.Recordset.MoveFirst
    
       'Para cada linea de albaran reestablecer el stock
       While (Not Data2.Recordset.EOF) And b
           Set vCStock = New cStock
           If InicializarCStock(vCStock, "E", Data2.Recordset!numlinea) Then
               'Actualiza el stock en salmac y borra de smoval
               If Not vCStock.DevolverStock2() Then b = False
           Else
               b = False
           End If
           Data2.Recordset.MoveNext
           Set vCStock = Nothing
       Wend
    End If
    
ERestablecer:
    If Err.Number <> 0 Then b = False
    ReestablecerStock = b
End Function


Private Sub BotonImprimir(OpcionListado As Byte)
Dim cadFormula As String
Dim Cadparam As String
Dim NumParam As Byte
Dim Cadselect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim ImpresionDirecta As Boolean
    If Text1(0).Text = "" Then
        MsgBox "Debe seleccionar un Albaran para Imprimir.", vbInformation
        Exit Sub
    End If
    
    cadFormula = ""
    Cadparam = ""
    Cadselect = ""
    NumParam = 0
    
    '===================================================
    '============ PARAMETROS ===========================
    If (OpcionListado = 45) Then
        If hcoCodTipoM = "ALZ" Then
            indRPT = 29   'Albaranes B
        Else
            If EsHistorico Then
                indRPT = 11 'Hist. Albaranes clientes
            Else
                indRPT = 10 'Albaran Clientes
            End If
        End If
    End If
    
    If Not PonerParamRPT(indRPT, Cadparam, NumParam, nomDocu, ImpresionDirecta) Then Exit Sub
   
    'Añadir el codigo de usuario como parametro para link con tabla Temporal (tmptiposiva) en el Report
    'tabla temporal para el calculo del bruto total para cada tipo de IVA
    Cadparam = Cadparam & "pCodUsu=" & vUsu.Codigo & "|"
    NumParam = NumParam + 1
    
    
    
    'PUNTO VERDE
    Cadparam = Cadparam & "PuntoVerde= """ & vParamAplic.ArtReciclado & """|"
    NumParam = NumParam + 1
    
    
    'Si se imprimen importes y/o
    devuelve = DevuelveDesdeBD(conAri, "albarcon", "sclien", "codclien", Text1(4).Text, "N")
    If devuelve = "" Then devuelve = "0"
    ' 0 "Todo"
    ' 1 "Cantidad y Precio"
    ' 2 "Cantidad"
    Cadparam = Cadparam & "Albarcon=" & devuelve & "|"
    NumParam = NumParam + 1

    
    
    
    'Nombre fichero .rpt a Imprimir
    If Not ImpresionDirecta Then frmImprimir.NombreRPT = nomDocu
        
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Nº de Albaran
    '---------------------------------------------------
    If Text1(0).Text <> "" Then
        'Cod Tipo Movimiento
        devuelve = "{" & NombreTabla & ".codtipom}='" & CodTipoMov & "'"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        'Nº Albaran
        devuelve = "{" & NombreTabla & ".numalbar}=" & Val(Text1(0).Text)
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        Cadselect = cadFormula
        
        If EsHistorico Then
            'El campo fecha tambien es clave primaria
            devuelve = Text1(1).Text
            devuelve = "{" & NombreTabla & ".fechaalb}=Date(" & Year(devuelve) & "," & Month(devuelve) & "," & Day(devuelve) & ")"
            If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
            
            devuelve = "{" & NombreTabla & ".fechaalb}='" & Format(Text1(1).Text, FormatoFecha) & "'"
            If Not AnyadirAFormula(Cadselect, devuelve) Then Exit Sub
        End If
        
    End If
   
    '=========================================================================
    'Aqui sabemos que valor tiene CodClien y añadimos a los parametros el tipo de IVA
    'que se aplica a ese cliente
    devuelve = DevuelveDesdeBDNew(conAri, "sclien", "tipoiva", "codclien", Text1(4).Text, "N")
    If devuelve <> "" Then
        'Albaranes en B
        If Me.hcoCodTipoM = "ALZ" Or hcoCodTipoM = "ALI" Then
            devuelve = "2"
        Else
            If devuelve = "3" Then devuelve = "2" 'El intracomunitario lo trato como exento
        End If
        Cadparam = Cadparam & "pTipoIVA=" & devuelve & "|"
        NumParam = NumParam + 1
    End If

        
    '==============================================================
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    devuelve = NombreTabla & " INNER JOIN " & NomTablaLineas & " ON "
    devuelve = devuelve & NombreTabla & ".codtipom=" & NomTablaLineas & ".codtipom AND " & NombreTabla & ".numalbar= " & NomTablaLineas & ".numalbar "
    If EsHistorico Then devuelve = devuelve & " AND " & NombreTabla & ".fechaalb= " & NomTablaLineas & ".fechaalb "
    If Not HayRegParaInforme(devuelve, Cadselect) Then Exit Sub
    
    
    If ImpresionDirecta Then
        'Imrpimie directamente. Tipo 4tonda
        If MsgBox("¿Imprimir el albarán?", vbQuestion + vbYesNo) = vbYes Then ImprimirDirectoAlb Cadselect
    Else
        With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = Cadparam
            .NumeroParametros = NumParam
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = OpcionListado
            .Titulo = "Albaran de Cliente"
            .ConSubInforme = True
            .Show vbModal
        End With
    End If
End Sub


Private Sub MostrarNSeries(ByRef RSLineas As ADODB.Recordset, Optional Dif As String, Optional cadSEL As String)
'Si los Nº de serie se introdujeron en ALBARAN COMPRAS se muestran
'los Nº de serie de los articulos comprados y se seleccionan tantos como cantidad de la linea
'Dif: si se ha modificado la cantidad pasamos la difencia con lo que habia
Dim SQL As String
Dim Campos As String

    On Error GoTo EMostrarNSeries

    If Text1(30).Text = "ART" Then
        SQL = MostrarNSeriesGnral(RSLineas, Campos, True)
    Else
        SQL = MostrarNSeriesGnral(RSLineas, Campos)
    End If
    
   If SQL <> "" Then
        Set frmMen = New frmMensajes
        frmMen.cadWhere = SQL
        
        If Dif <> "" Then
            SQL = " WHERE (codtipom=" & DBSet(CodTipoMov, "T") & " and "
            SQL = SQL & " numalbar=" & Text1(0).Text & " and numline1=" & Data2.Recordset!numlinea & ")"
            frmMen.cadWhere2 = Dif & "|" & SQL & "|"
        Else
            If cadSEL <> "" Then
                'seleccionar lineas de nº serie de la factura a rectificar
                frmMen.cadWhere2 = cadSEL
            Else
                frmMen.cadWhere2 = ""
            End If
        End If
        frmMen.OpcionMensaje = 4 'Nº Series Articulo
        frmMen.vCampos = Campos
        frmMen.Show vbModal
        Set frmMen = Nothing
    End If
    
EMostrarNSeries:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub PedirNSeries(ByRef RS As ADODB.Recordset)
Dim SQL As String

    On Error GoTo EPedirNSeries

        SQL = "El artículo tienen control de Nº de Serie." & vbCrLf & vbCrLf
        SQL = SQL & "Introduzca los Nº De Serie." & vbCrLf
        MsgBox SQL, vbInformation
        PedirNSeriesGnral RS, False
        
       ' Set frmNSerie = New frmRepCargarNSerie
       ' frmNSerie.DeVentas = True 'Se llama desde Alb. de Venta
       ' frmNSerie.Show vbModal
       ' Set frmNSerie = Nothing
        
EPedirNSeries:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Sub InsertarCabecera()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim SQL As String
    On Error GoTo EInsertarCab
    
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
        Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
        SQL = CadenaInsertarDesdeForm(Me)
        If SQL <> "" Then
            If InsertarOferta(SQL, vTipoMov) Then
                CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                PonerCadenaBusqueda
                PonerModo 2
                'Ponerse en Modo Insertar Lineas
                BotonMtoLineas 0, "Albaranes"
                BotonAnyadirLinea
            End If
        End If
        Text1(0).Text = Format(Text1(0).Text, "0000000")
    End If
    Set vTipoMov = Nothing
    Me.SSTab1.Tab = 0
    
EInsertarCab:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Sub ComprobarNSeriesLineas(numlinea As String)
'Al pasar de PEDIDO a ALBARAN
'control de Nº Series si hay algun articulo en las lineas de pedido que requiere Nº de serie
'Si NO se realiza control Nº series en compras pedirlos ahora
'Si se realiza control Nº Series en compras verificar que efectivamente estan introducidos
'y mostrarlos para seleccionarlos
Dim SQL As String
Dim RSLineas As ADODB.Recordset
Dim cadWhere As String
Dim Dif As Single

    'Comprobar si el Articulo tiene control de Nº de Serie
    SQL = DevuelveDesdeBDNew(conAri, "sartic", "nseriesn", "codartic", txtAux(1).Text, "T")
    
    If SQL = "1" Then 'Hay NºSerie para el Articulo
        'si estamos insertando
        If Modo = 5 Then
            If ModificaLineas = 1 Then 'Insertar
                'Comprobar que la cantidad comprada es >0
                If ComprobarCero(txtAux(3).Text) <= 0 Then Exit Sub
            ElseIf ModificaLineas = 2 Then 'Modificar
                'si se ha modificado la cantidad, habrá que quitar algun nº serie
                'de los seleccionado o anyadir alguno mas
                Dif = CSng(txtAux(3).Text) - CSng(Data2.Recordset!Cantidad)
                If Dif = 0 Then Exit Sub
                If Text1(30).Text = "ART" Then Exit Sub
'                    Dif = CSng(Data2.Recordset!Cantidad) - CSng(txtAux(3).Text)
            End If
        End If
        
        cadWhere = " WHERE codtipom=" & DBSet(CodTipoMov, "T") & " and "
        cadWhere = cadWhere & " numalbar=" & Text1(0).Text & " and numlinea=" & numlinea
    
        'Seleccionamos aquellas lineas de albaran que tienen Nº de Serie
        SQL = "SELECT slialb.codartic, sum(cantidad) as cantidad, numlinea "
        SQL = SQL & " FROM slialb INNER JOIN sartic on slialb.codartic=sartic.codartic "
        SQL = SQL & cadWhere & " And nseriesn = 1 "
        SQL = SQL & " GROUP BY codartic ORDER BY Codartic "

        Set RSLineas = New ADODB.Recordset
        RSLineas.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        Me.cmdAux(1).Tag = Text1(0).Text 'Num Albaran
        Me.cmdAux(0).Tag = numlinea 'Num Linea
        
        'Comprobar si NO Hay Nº SERIE en Compras y si no se realizo alli
        'Mostrar ahora ventana para pedir los Nº Serie de la cantidad introducida
        If Not vParamAplic.NumSeries And ModificaLineas = 1 Then
            PedirNSeries RSLineas
        Else 'Se realizo contro en COMPRAS, Mostramos los Nº y seleccionamos
            If ModificaLineas = 1 Then 'Insertando la linea
                'Comprobar que efectivamente estan en tabla sserie los NºSerie del Articulo
                ' y que no esten asignados ya a otro albaran de venta
                SQL = " select distinct count(numserie) from sserie where codartic=" & DBSet(txtAux(1).Text, "T") & " and (numalbar='' or isnull(numalbar))"
                '=== Laura 17/01/2007
                'y que no este asignados a una factura
                SQL = SQL & " and (numfactu='' or isnull(numfactu))"
                '===
                If RegistrosAListar(SQL) = 0 Then 'No hay Nº de Serie para elegir
                    PedirNSeries RSLineas
                Else
                    MostrarNSeries RSLineas
                End If
            ElseIf ModificaLineas = 2 Then
                SQL = " select distinct count(numserie) from sserie " & Replace(cadWhere, "numlinea", "numline1")
                If RegistrosAListar(SQL) > 0 Then
                    MostrarNSeries RSLineas, CStr(Dif)
                End If
            End If
        End If

        RSLineas.Close
        Set RSLineas = Nothing
    End If
End Sub


Private Sub BotonNSeries()
Dim cadWhere As String, SQL As String
Dim RSLineas As ADODB.Recordset

    If Me.Data1.Recordset!EsTicket Then
        MsgBox "Albaranes provenientes de Ticket no tienen control de Nº Serie.", vbInformation
        Exit Sub
    End If

    'Si es Albaran para Factura rectificativa (ART)
    If CodTipoMov = "ART" Then
'      'Si es una Factura Venta(FAV) generada desde un ticket del TPV entonces
'      'no hay numseries
'      SQL = DevuelveDesdeBDNew(conAri, "scafac1", "codtipoa", "codtipom", Data1.Recordset!codtipmf, "T", , "numfactu", Data1.Recordset!NumFactu, "N", "fecfactu", Data1.Recordset!FecFactu, "F")
'      If SQL = "FTI" Then
'        MsgBox "Facturas provenientes de Ticket no tienen control de Nº Serie.", vbInformation
'        Exit Sub
'      Else
        Exit Sub
'      End If
    End If
    
    
    
    ModificaLineas = 4

    cadWhere = " WHERE codtipom='" & Text1(30).Text & "'"
    cadWhere = cadWhere & " and numalbar=" & Text1(0).Text
    
    'Seleccionamos aquellas lineas de albaran que tienen Nº de Serie
    SQL = "SELECT numlinea, slialb.codartic, sum(cantidad) as cantidad "
    SQL = SQL & " FROM slialb INNER JOIN sartic on slialb.codartic=sartic.codartic "
    SQL = SQL & cadWhere & " And nseriesn = 1 "
    SQL = SQL & " GROUP BY codartic ORDER BY Codartic "

    Set RSLineas = New ADODB.Recordset
    RSLineas.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RSLineas.EOF Then
        'Comprobar si NO Hay Nº SERIE en Compras y si no se realizo alli
        'Mostrar ahora ventana para pedir los Nº Serie de la cantidad introducida
        PedirNSeriesT RSLineas
    Else
        MsgBox "No hay ninguna linea de Articulo con Control de Nº Serie", vbInformation
    End If
    RSLineas.Close
    Set RSLineas = Nothing
    ModificaLineas = 0
End Sub


Private Sub PedirNSeriesT(ByRef RS As ADODB.Recordset)
Dim RSseries As ADODB.Recordset
Dim SQL As String
Dim linea As Integer

    On Error GoTo EPedirNSeries


        PedirNSeriesGnral RS, False
        RS.MoveFirst
        While Not RS.EOF
            linea = 0
            'Cargar los Nº de serie asignados
            SQL = "SELECT numserie, codartic FROM sserie "
            SQL = SQL & " WHERE codtipom='" & Text1(30).Text & "' and "
            SQL = SQL & "numalbar=" & Text1(0).Text
            SQL = SQL & " and numline1=" & RS!numlinea
            SQL = SQL & " ORDER BY codartic "
            Set RSseries = New ADODB.Recordset
            RSseries.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RSseries.EOF
                linea = linea + 1
                SQL = "UPDATE tmpnseries SET numserie=" & DBSet(RSseries!numSerie, "T")
                SQL = SQL & " WHERE codartic=" & DBSet(RS!codArtic, "T")
                SQL = SQL & " and numlinealb=" & RS!numlinea
                SQL = SQL & " and numlinea=" & linea
                conn.Execute SQL
                RSseries.MoveNext
            Wend
            RS.MoveNext
        Wend
        RSseries.Close
        Set RSseries = Nothing
     '   Set frmNSerie = New frmRepCargarNSerie
      '  frmNSerie.DeVentas = True 'Se llama desde Alb. de Venta
      '  frmNSerie.NumAlb = Text1(0).Text
      '  frmNSerie.Show vbModal
      '  Set frmNSerie = Nothing
EPedirNSeries:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Sub CargarNumSeries()
'Insertar un registro en la tabla "sserie" por cada uno de los
'Nº de Serie introducidos en la Tabla Temporal o actualizarlo
Dim RStmp As ADODB.Recordset
Dim SQL As String
Dim b As Boolean

    On Error GoTo ECargar
    
    conn.BeginTrans
    
    'Limpiar primero los Nº de serie asignados al ALV y luego volver a cargarlos
    SQL = "UPDATE sserie SET codtipom=" & ValorNulo & ", numalbar=" & ValorNulo & ", fechavta="
    SQL = SQL & ValorNulo & ", numline1=" & ValorNulo
    SQL = SQL & " WHERE codtipom=" & DBSet(Text1(30).Text, "T") & " and numalbar=" & Text1(0).Text & " AND year(fechavta)=" & Year(Text1(1).Text)
    conn.Execute SQL
    
    'Recuperar los Nº Serie de ese articulo cargados en la Temporal
    'Seleccionar los nº de serie cargados en la temporal: tmpnseries
    SQL = "SELECT * FROM tmpnseries WHERE codusu=" & vUsu.Codigo
    SQL = SQL & " ORDER BY codartic, numlinealb, numlinea "
    Set RStmp = New ADODB.Recordset
    RStmp.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                
    b = True
    While Not RStmp.EOF And b
        b = InsertarNSerie(RStmp!numSerie, RStmp!codArtic, RStmp!numlinealb)
        RStmp.MoveNext
    Wend
    RStmp.Close
    Set RStmp = Nothing
    
ECargar:
    If Err.Number <> 0 Then b = False
    If b Then
        conn.CommitTrans
    Else
        conn.RollbackTrans
    End If
End Sub


Private Function InsertarNSerie(numSerie As String, codArtic As String, numlinea As String) As Boolean
'Inserta o Actualiza en la tabla sserie, si al pasar Pedido -> Albaran
'existen lineas con control de Nº Serie
Dim devuelve As String
Dim TieneMan As String * 1
Dim NumAlbar As String
Dim nSerie As CNumSerie
Dim b As Boolean

    On Error GoTo EInsertarNSerie

    'Comprobar que el cliente tiene mantenimientos en esa direc/dpto
    TieneMan = "0"
    devuelve = DevuelveDesdeBDNew(conAri, "scaman", "nummante", "codclien", Text1(4).Text, "N", , "coddirec", Text1(12).Text, "N")
    'El cliente tiene Mantenimientos
    If devuelve <> "" Then TieneMan = "1"

    Set nSerie = New CNumSerie
    nSerie.numSerie = numSerie
    nSerie.Articulo = codArtic
    
    nSerie.Cliente = CLng(Text1(4).Text)
    nSerie.DirDpto = Text1(12).Text
    nSerie.conMante = TieneMan
    nSerie.tipoMov = CodTipoMov
    nSerie.FechaVta = Text1(1).Text
    nSerie.NumAlbaran = Text1(0).Text
    nSerie.NumLinAlb = numlinea
    nSerie.ObtenFechaFinGarantia codArtic, Text1(1).Text
        
'    'fin garantia= fecha albaran + dias de garantia
'    If Text1(1).Text <> "" Then
'        'obtenemos los dias de garantia del articulo
'        devuelve = DevuelveDesdeBDNew(conAri, "sartic", "garantia", "codartic", codArtic, "T")
'        nSerie.FinGarantia = CStr(CDate(Text1(1).Text) + CInt(ComprobarCero(devuelve)))
'    End If
    
    'Comprobar si existe en la tabla sserie
     NumAlbar = "numalbar" 'Nº albaran de Venta
     devuelve = DevuelveDesdeBDNew(conAri, "sserie", "numserie", "numserie", numSerie, "T", NumAlbar, "codartic", codArtic, "T")
     If devuelve <> "" Then 'EXISTE en tabla sserie
        If NumAlbar = "" Then b = nSerie.ActualizarNumSerie(True)
     Else
        b = nSerie.InsertarNumSerie
    End If
    InsertarNSerie = True
    Set nSerie = Nothing
    
EInsertarNSerie:
    If Err.Number <> 0 Then b = False
    If b Then
        InsertarNSerie = True
    Else
        InsertarNSerie = False
    End If
End Function




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
            
'            If (Not EsDeVarios) Or (EsDeVarios And modo = 3) Then
            Text1(4).Text = vCliente.Codigo
            FormateaCampo Text1(4)
            If (Modo = 3) Or (Modo = 4) Then
                Text1(5).Text = vCliente.Nombre  'Nom clien
                Text1(8).Text = vCliente.Domicilio
                Text1(9).Text = vCliente.CPostal
                Text1(10).Text = vCliente.Poblacion
                Text1(11).Text = vCliente.Provincia
                Text1(6).Text = vCliente.NIF
                Text1(7).Text = DBLet(vCliente.TfnoClien, "T")
            End If
            
            If Modo = 3 Or Modo = 4 Then 'insertar
                Text1(14).Text = vCliente.ForPago
                Text2(14).Text = PonerNombreDeCod(Text1(14), conAri, "sforpa", "nomforpa")
                Text1(15).Text = Format(vCliente.DtoPPago, FormatoDescuento)
                Text1(16).Text = Format(vCliente.DtoGnral, FormatoDescuento)
                Text1(17).Text = vCliente.Agente
                Text2(17).Text = PonerNombreDeCod(Text1(17), conAri, "sagent", "nomagent")
                Text1(34).Text = vCliente.Kilometros
                Me.cboFacturacion.ListIndex = vCliente.TipoFactu
                Text1(29).Text = vCliente.FEnvio
                Text2(29).Text = PonerNombreDeCod(Text1(29), conAri, "senvio", "nomenvio")
            End If
                
'                SituacionCliente = RS.Fields!codsitua

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
    If b Then Text1(5).Text = vCliente.Nombre         'Nom clien
    Text1(8).Text = vCliente.Domicilio
    Text1(9).Text = vCliente.CPostal
    Text1(10).Text = vCliente.Poblacion
    Text1(11).Text = vCliente.Provincia
    Text1(7).Text = DBLet(vCliente.TfnoClien, "T")
            
'    If Not b Then PonerFoco Text1(6)
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


Private Function ActualizarFecMovCliente() As Boolean
Dim vCliente As CCliente
Dim b As Boolean

    On Error GoTo EActFecha

    ActualizarFecMovCliente = False
    Set vCliente = New CCliente
    vCliente.Codigo = Text1(4).Text
    b = vCliente.ActualizaUltFecMovim(Text1(1).Text)
    Set vCliente = Nothing
    
EActFecha:
    If Err.Number <> 0 Then b = False
    ActualizarFecMovCliente = b
End Function


Private Sub CalcularDatosFactura()
Dim I As Integer
Dim cadWhere As String, SQL As String
Dim vFactu As CFactura
Dim CambiarValoresIVA As Boolean

    'Limpiar en el form los datos calculados de la factura
    'y volvemos a recalcular
    For I = 33 To 56
         Text3(I).Text = ""
    Next I

    'Comprobar que hay lineas de albaran para calcular totales
    cadWhere = ObtenerWhereCP(False)
    SQL = "Select count(*) from " & NomTablaLineas & " Where " & Replace(cadWhere, NombreTabla, NomTablaLineas)
    If RegistrosAListar(SQL) = 0 Then Exit Sub
    
    Set vFactu = New CFactura
    vFactu.DtoPPago = CCur(ComprobarCero(Text1(15).Text))
    vFactu.DtoGnral = CCur(ComprobarCero(Text1(16).Text))
    vFactu.Cliente = Text1(4).Text
    If hcoCodTipoM = "ALZ" Or hcoCodTipoM = "ALI" Then vFactu.Codtipom = hcoCodTipoM
    
    'Si el albaran es rectificativo y la fecha es
    CambiarValoresIVA = False
    If hcoCodTipoM = "ART" Then CambiarValoresIVA = CDate(Text1(35).Text) < CDate("01/09/2012")
    
    If vFactu.CalcularDatosFactura(cadWhere, NombreTabla, NomTablaLineas, CambiarValoresIVA) Then
        Text3(33).Text = vFactu.BrutoFac
        Text3(34).Text = vFactu.ImpPPago
        Text3(35).Text = vFactu.ImpGnral
        Text3(36).Text = vFactu.BaseImp
        Text3(37).Text = vFactu.TipoIVA1
        Text3(38).Text = vFactu.TipoIVA2
        Text3(39).Text = vFactu.TipoIVA3
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
                '---- Laura: Modificado 27/09/2006
'                Text3(i + 3).Text = QuitarCero(Text3(i).Text)
                Text3(I + 3).Text = QuitarCero(Text3(I + 3).Text)
                '----
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



 Private Sub InsertarLineasFactu(cadWhere)
'cadSerie = "INSERT INTO slialb(codtipom,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel,origpre) "
'cadSerie = cadSerie & " SELECT '" & Text1(30).Text & "' as codtipom," & Text1(0).Text & " as numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel,origpre FROM slifac WHERE " & CadenaSeleccion
 Dim RS As ADODB.Recordset
 Dim SQL As String
 Dim I As Integer
 Dim cadI As String
 Dim NumLin As String
 Dim LitrosUd As Currency
 
    On Error GoTo EInsFactu
    Screen.MousePointer = vbHourglass
    
    If cadWhere <> "" Then
        'Obtenemos el numero de linea a insertar
'        SQL = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas)
'        SQL = SugerirCodigoSiguienteStr("slialb", "numlinea", SQL)
'        i = Int(SQL)
    
        cadI = ""
        
        'MAyo 2009
        'SQL = "SELECT * FROM slifac WHERE " & cadWhere
        SQL = "select slifac.*,unicajas,LitrosUnidad from slifac, sartic where "
        SQL = SQL & " sartic.codartic=slifac.codartic "
        SQL = SQL & " AND " & cadWhere
    
    
        Set RS = New ADODB.Recordset
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
            txtAux(0).Text = RS!codAlmac
            txtAux(1).Text = RS!codArtic
            txtAux(2).Text = RS!NomArtic
            Text2(16).Text = DBLet(RS!ampliaci, "T")
            txtAux(3).Text = CStr(RS!Cantidad * -1)
            txtAux(4).Text = RS!precioar
            txtAux(5).Text = DBLet(RS!origpre, "T")
            txtAux(6).Text = RS!dtoline1
            txtAux(7).Text = RS!dtoline2
            txtAux(8).Text = CStr(RS!ImporteL * -1)
            txtAux(9).Text = DBLet(RS!Codprovex, "N")
            
            'Cajas e importe litros
            I = DBLet(RS!Unicajas, "N")
            If I = 0 Then I = 1
            I = RS!Cantidad \ I
            txtAux(11).Text = -I
            
            'Precio por litro
            LitrosUd = DBLet(RS!LitrosUnidad, "N")
            If LitrosUd <= 1 Then
                LitrosUd = RS!precioar
            Else
                LitrosUd = (RS!precioar / LitrosUd)
                LitrosUd = Round2(LitrosUd, 4)
            End If
            txtAux(12).Text = CStr(LitrosUd)
            'para no tener que traer ahora el proveedor pongo en txt(10) un texto
            txtAux(10).Text = "*"
            
            If InsertarLinea(NumLin, True) Then
            
            End If
            
'            SQL = "('" & Text1(30).Text & "'," & Text1(0).Text & "," & i & ","  'codtipoa,numalbar,numlinea
'            SQL = SQL & DBSet(RS!codAlmac, "N") & "," & DBSet(RS!codArtic, "T") & "," & DBSet(RS!NomArtic, "T") & "," & DBSet(RS!ampliaci, "T") & ","
'            SQL = SQL & DBSet(RS!cantidad * -1, "N") & "," & DBSet(RS!precioar, "N") & "," & DBSet(RS!dtoline1, "N") & "," & DBSet(RS!dtoline2, "N") & ","
'            SQL = SQL & DBSet(RS!ImporteL * -1, "N") & "," & DBSet(RS!origpre, "T") & ")"
'            If cadI = "" Then
'                cadI = SQL
'            Else
'                cadI = cadI & "," & SQL
'            End If
'            i = i + 1
            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
        
        CalcularDatosFactura
        
'        If cadI <> "" Then
'            SQL = "INSERT INTO slialb(codtipom,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel,origpre) VALUES "
'            SQL = SQL & cadI
'            Conn.Execute SQL
'        End If
    End If
    Screen.MousePointer = vbDefault
    
EInsFactu:
    If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        MuestraError Err.Number, "Lineas Factura", Err.Description
    End If
End Sub


'Private Sub InsertarLineasFactu_old(cadWHERE)
''cadSerie = "INSERT INTO slialb(codtipom,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel,origpre) "
''cadSerie = cadSerie & " SELECT '" & Text1(30).Text & "' as codtipom," & Text1(0).Text & " as numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel,origpre FROM slifac WHERE " & CadenaSeleccion
' Dim RS As ADODB.Recordset
' Dim SQL As String
' Dim i As Integer
' Dim cadI As String
'
'    On Error GoTo EInsFactu
'    Screen.MousePointer = vbHourglass
'
'    If cadWHERE <> "" Then
'        'Obtenemos el numero de linea a insertar
'        SQL = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas)
'        SQL = SugerirCodigoSiguienteStr("slialb", "numlinea", SQL)
'        i = Int(SQL)
'
'        cadI = ""
'
'        SQL = "SELECT * FROM slifac WHERE " & cadWHERE
'
'        Set RS = New ADODB.Recordset
'        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        While Not RS.EOF
'            SQL = "('" & Text1(30).Text & "'," & Text1(0).Text & "," & i & ","  'codtipoa,numalbar,numlinea
'            SQL = SQL & DBSet(RS!codAlmac, "N") & "," & DBSet(RS!codArtic, "T") & "," & DBSet(RS!NomArtic, "T") & "," & DBSet(RS!ampliaci, "T") & ","
'            SQL = SQL & DBSet(RS!Cantidad * -1, "N") & "," & DBSet(RS!precioar, "N") & "," & DBSet(RS!dtoline1, "N") & "," & DBSet(RS!dtoline2, "N") & ","
'            SQL = SQL & DBSet(RS!ImporteL * -1, "N") & "," & DBSet(RS!origpre, "T") & ")"
'            If cadI = "" Then
'                cadI = SQL
'            Else
'                cadI = cadI & "," & SQL
'            End If
'            i = i + 1
'            RS.MoveNext
'        Wend
'        RS.Close
'        Set RS = Nothing
'
'        If cadI <> "" Then
'            SQL = "INSERT INTO slialb(codtipom,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel,origpre) VALUES "
'            SQL = SQL & cadI
'            Conn.Execute SQL
'        End If
'    End If
'    Screen.MousePointer = vbDefault
'
'EInsFactu:
'    If Err.Number <> 0 Then
'        Screen.MousePointer = vbDefault
'        MuestraError Err.Number, "Lineas Factura", Err.Description
'    End If
'End Sub



Private Function AsignarNumSeriesAlbVenta(cadSEL As String) As Boolean
Dim I As Integer
Dim cant As Integer
Dim cadSerie As String
Dim nSerie As CNumSerie
Dim devuelve As String
Dim b As Boolean
    
    'Para cada valor empipado actualizar la tabla sserie
    
    
    cant = CInt(ComprobarCero(txtAux(3).Text))
    
    On Error GoTo ErrorNSerie
    conn.BeginTrans
        
    If ModificaLineas = 2 Then 'Venimos de modificar la cantidad de una linea
        'Borramos los numeros de serie que tenia asignada la linea del albaran
        Set nSerie = New CNumSerie
        nSerie.tipoMov = CodTipoMov
        nSerie.NumAlbaran = Text1(0).Text
        nSerie.NumLinAlb = ComprobarCero(Me.cmdAux(0).Tag)
        b = nSerie.BorrarNumSeriesAlbVta
        Set nSerie = Nothing
    Else
        b = True
    End If
        
    If b Then
        Set nSerie = New CNumSerie
        nSerie.Articulo = txtAux(1).Text
        nSerie.Cliente = CLng(Text1(4).Text)
        nSerie.DirDpto = Text1(12).Text
        nSerie.tipoMov = CodTipoMov
        nSerie.FechaVta = Text1(1).Text
        If nSerie.FechaVta <> "" Then
            devuelve = DevuelveDesdeBDNew(conAri, "sartic", "garantia", "codartic", txtAux(1).Text, "T")
            nSerie.FinGarantia = CStr(CDate(nSerie.FechaVta) + CInt(ComprobarCero(devuelve)))
        End If
        nSerie.NumAlbaran = Text1(0).Text
        nSerie.NumLinAlb = ComprobarCero(Me.cmdAux(0).Tag)
                
        For I = 1 To cant
            cadSerie = RecuperaValor(cadSEL, I + 1)
            If cadSerie <> "" Then
                nSerie.numSerie = cadSerie
                If nSerie.ActualizarNumSerie(True) = False And b Then b = False
            End If
        Next I
        Set nSerie = Nothing
    End If
ErrorNSerie:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Actualizar tabla Nº Series", Err.Description
        b = False
    End If
    If b Then
        conn.CommitTrans
    Else
        conn.RollbackTrans
    End If
    AsignarNumSeriesAlbVenta = b
End Function




Private Sub DevolverNumSeries()
Dim SQL As String
Dim cadWhere As String
Dim RS As ADODB.Recordset

    On Error GoTo EDevNumSerie
        
    cadWhere = ObtenerWhereCP(True)
    SQL = "select slialb.codartic,abs(cantidad) as cantidad,numlinea"
    SQL = SQL & " from slialb inner join scaalb on slialb.codtipom=scaalb.codtipom and scaalb.numalbar=slialb.numalbar "
    SQL = SQL & " inner join sserie on slialb.codartic=sserie.codartic and sserie.numfactu=scaalb.numfactu and sserie.codclien=scaalb.codclien "
    '-- LAURA: 02/07/2007
'    SQL = SQL & " inner join scafac1 on scafac1.codtipom=scaalb.codtipmf and scafac1.numfactu=scaalb.numfactu and scafac1.fecfactu=scaalb.fecfactu "
'    SQL = SQL & " inner join sserie on scafac1.codtipoa=sserie.codtipom and scafac1.numalbar=sserie.numalbar and scafac1.fechaalb=sserie.fechavta "
    SQL = SQL & cadWhere & " and scaalb.numfactu=" & CStr(Me.Data1.Recordset!NumFactu)
'    If Me.Data1.Recordset!codtipmf = "FAV" Then SQL = SQL & " AND codtipom='ALV'"
    '--

    
    
    
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Hay articulos con nº de serie en las lineas del albaran rectificativo
    'que hay que quitar de los nº de serie que tenia asignados
    'estamos devolviendo nº serie y pedimos los que vamos a devolver y a estos
    'le limpiamos los campos de venta de la tabla sserie
    If Not RS.EOF Then
        SQL = "select sserie.numserie, sserie.codartic, sartic.nomartic"
        SQL = SQL & " from slialb inner join scaalb on slialb.codtipom=scaalb.codtipom and scaalb.numalbar=slialb.numalbar "
        '-- LAURA: 02/07/2007
'        SQL = SQL & " inner join scafac1 on scafac1.codtipom=scaalb.codtipmf and scafac1.numfactu=scaalb.numfactu and scafac1.fecfactu=scaalb.fecfactu "
'        SQL = SQL & " inner join sserie on scafac1.codtipoa=sserie.codtipom and scafac1.numalbar=sserie.numalbar and scafac1.fechaalb=sserie.fechavta "
        SQL = SQL & " inner join sserie on slialb.codartic=sserie.codartic and sserie.numfactu=scaalb.numfactu  and sserie.codclien=scaalb.codclien "
        '--
        SQL = SQL & " inner join sartic on sserie.codartic=sartic.codartic "
        SQL = SQL & cadWhere & " and scaalb.numfactu=" & CStr(Me.Data1.Recordset!NumFactu)
    
        MostrarNSeries RS, , SQL
    End If
    RS.Close
    Set RS = Nothing
    
EDevNumSerie:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Actualizando Nº Serie.", Err.Description
    End If
End Sub




Private Function QuitarNumSeriesAlbVenta(cadSEL As String) As Boolean
Dim I As Integer
Dim numSerie As String
Dim codArtic As String
Dim nSerie As CNumSerie
Dim Grupo As String
Dim b As Boolean
    
    'Para cada valor empipado actualizar la tabla sserie
   
    On Error GoTo ErrorNSerie
    
    b = True
    While cadSEL <> ""
        I = InStr(1, cadSEL, "·")
        If I > 0 Then
            Grupo = Mid(cadSEL, 1, I - 1)
            cadSEL = Mid(cadSEL, I + 1, Len(cadSEL))
            If Grupo <> "" Then
                codArtic = RecuperaValor(Grupo, 1)
                numSerie = RecuperaValor(Grupo, 2)
                
                Set nSerie = New CNumSerie
                nSerie.numSerie = numSerie
                nSerie.Articulo = codArtic
                b = b And nSerie.ActualizarNumSerie(True)
                Set nSerie = Nothing
            End If
        End If
    Wend
   
ErrorNSerie:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Actualizar tabla Nº Series", Err.Description
        Set nSerie = Nothing
        b = False
    End If
    QuitarNumSeriesAlbVenta = b
End Function


Private Sub BotonRecuperarFactura()
'Genera una factura a partir del Albaran de Mostrador
'pero sin coger contador de factura lo pide en un form

    'Comprobar que esta marcada para facturar
    If Me.chkFacturar.Value = 1 Then
        NumRegElim = Data1.Recordset.AbsolutePosition
        
        'Facturacion de Albaran de Mostrador
        frmListadoPed.CodClien = CodTipoMov  'utilizamos esta vble para pasarle el tipo de movimiento
        frmListadoPed.NumCod = Text1(0).Text  'utilizamos esta vble para pasarle el nº albaran
        frmListadoPed.EstaRecupFact = Me.RecuperarFactu
        AbrirListadoPed (222)
        
        PosicionarDataTrasEliminar
    Else
        MsgBox "El Albaran no esta marcado para facturar", vbInformation
    End If
End Sub


Private Sub MarcarAlbaranes()

        'Lanzara un desde hasta y los marcara
        frmListado.NumCod = hcoCodTipoM
        CadenaDesdeOtroForm = ""
        AbrirListado 82
        If CadenaDesdeOtroForm = "OK" Then
            'OK. Cambiadas las marcas. Refrescamos y situamos
            Screen.MousePointer = vbHourglass
            DoEvents
            PonerCadenaBusqueda
            PosicionarData
            Screen.MousePointer = vbDefault
        End If
        
End Sub


Private Sub CantidadCajas(DeCantidadACajas As Boolean)
Dim CajaUd As Integer
Dim V As Long
Dim V2 As Currency
    CajaUd = 1
    If Not (ElArticulo Is Nothing) Then
        If ElArticulo.UnidCaja > 1 Then CajaUd = ElArticulo.UnidCaja
    End If
        

    If DeCantidadACajas Then
        If txtAux(3).Text = "" Then
            txtAux(11).Text = ""
        Else
            V2 = ImporteFormateado(txtAux(3).Text)
            txtAux(11).Text = V2 \ CajaUd
        End If
    Else
        'Ha metido cajas. Nos vamos a cantidad
        If txtAux(11).Text = "" Then
            txtAux(3).Text = ""
        Else
        
            V = Val(txtAux(11).Text)
            txtAux(3).Text = Format(V * CajaUd, FormatoCantidad)
        End If
    End If
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
            txtAux(12).Text = ""
        Else
            V = ImporteFormateado(txtAux(4).Text)
            V = Round(V / ListrosUd, 4)
            txtAux(12).Text = Format(V, FormatoPrecio)
        End If
    Else
        'Ha metido precio UD
        If txtAux(12).Text = "" Then
            txtAux(4).Text = ""
        Else
            V = ImporteFormateado(txtAux(12).Text)
            V = V * ListrosUd
            txtAux(4).Text = Format(V, FormatoPrecio)
        End If
    End If
End Sub




Private Function LineasRecicladoCorrectas() As Boolean

Dim cad As String
Dim canti As Currency
Dim ConReciclado As Boolean
Dim Referencia As String
Dim Fin1 As Boolean

    On Error GoTo ELineasRecicladoCorrectas
    LineasRecicladoCorrectas = True
    If Not ClienteConTasaReciclado Then Exit Function
    cad = "select slialb.codartic,slialb.nomartic,cantidad,tasareciclado from slialb,sartic,sunida where"
    cad = cad & " slialb.codartic=sartic.codartic and  sunida.CodUnida = sartic.CodUnida"
    cad = cad & " And codtipom='" & Text1(30).Text & "' AND NumAlbar = " & Text1(0).Text & " ORDER BY numlinea"
    Set RN = New ADODB.Recordset
    RN.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cad = "" 'aqui meteremos los fallos
    Referencia = ""
    If Not RN.EOF Then
        'Tiene lineas
        Do
            
        
        
            If Referencia = "" Then
                
                
                    If DBLet(RN!tasareciclado, "N") > 0 Then
                        Referencia = RN!codArtic
                        txtAnterior = DBLet(RN!NomArtic, "T")
                        canti = RN!Cantidad
                    Else
                        If RN!codArtic = vParamAplic.ArtReciclado Then cad = cad & "PUNTO VERDE sin pertencer a ningun articulo" & vbCrLf
                    End If
                    RN.MoveNext
              
            Else
            
                'Ya teniamos un articulo con tasa reciclado
                If RN!codArtic <> vParamAplic.ArtReciclado Then
                    'MAL. Tenia que tener linea del punto verde
                    cad = cad & Referencia & "  " & txtAnterior & "   SIN PUNTO VERDE" & vbCrLf
                    
                    
                    'Ponemos apuntando a el
                    If DBLet(RN!tasareciclado, "N") > 0 Then
                        Referencia = RN!codArtic
                        txtAnterior = DBLet(RN!NomArtic, "T")
                        canti = RN!Cantidad
                    Else
                        Referencia = ""
                    End If
                    
                Else
                    'OK despues de la linea del articulo esta el punto verde.
                    'Coinciden las cantidades?
                    If DBLet(RN!Cantidad, "N") <> canti Then
                        cad = cad & Referencia & "  " & txtAnterior & "   Cantidades distintas" & vbCrLf
                    Else
                        'OK. Todo perfecto. Tiene pverde y es la misma cantidad
                        
                     End If
                     Referencia = ""
                End If
                'Mv al siguiente
                RN.MoveNext
        
            End If
        
        
        Loop While Not RN.EOF

    End If
    RN.Close
    
    'La ultima no teine punto verde
    If Referencia <> "" Then cad = cad & Referencia & "  " & txtAnterior & "   SIN PUNTO VERDE" & vbCrLf
    
    If cad <> "" Then
        cad = cad & vbCrLf & vbCrLf & "Continuar?"
        If MsgBox("Error comprobando tasa reciclado" & vbCrLf & vbCrLf & cad, vbQuestion + vbYesNo) = vbNo Then LineasRecicladoCorrectas = False
    End If
ELineasRecicladoCorrectas:
    If Err.Number <> 0 Then MuestraError Err.Number, "Proceso: " & LineasRecicladoCorrectas, Err.Description
    Set RN = Nothing
End Function




Private Function ComprobarNUmerosDeLote() As Boolean
Dim SQL As String
Dim Ca As Currency
Dim Col As Collection
    On Error GoTo EComprobarNUmerosDeLote

    SQL = ""
    If EsHistorico Then SQL = "N"
    'If vEmpresa.codempre = EmpresaAVAB Then SQL = "N"
    'If vParamAplic.EsAVAB Then SQL = "N"  'Ahora tb voy a mirar que el AVAB meta el lotaje
    
    If SQL <> "" Then
        ComprobarNUmerosDeLote = True
        Exit Function
    End If
    
    ComprobarNUmerosDeLote = False
    
    Set RN = New ADODB.Recordset
    SQL = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas)
    SQL = NomTablaLineas & ",sartic WHERE " & NomTablaLineas & ".codartic = sartic.codartic AND " & SQL
    SQL = SQL & " AND trazabilidad = 1 ORDER BY numlinea"
    SQL = "Select numlinea,slialb.codartic,sartic.nomartic,cantidad from " & SQL
    RN.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set Col = New Collection
    While Not RN.EOF
        SQL = ""
        Col.Add RN!numlinea & "|" & CStr(RN!Cantidad) & "|" & RN!codArtic & " " & DBLet(RN!NomArtic, "T") & "|"
        RN.MoveNext
    Wend
    RN.Close
    
    If Col.Count > 0 Then
        'Hay articulos con trazabilidad
        SQL = Replace(ObtenerWhereCP(True), NombreTabla, "slialblotes")
        SQL = "select numlinea,sum(cantidad) from slialblotes " & SQL & " group by 1"
        RN.Open SQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
        
        
        txtAnterior = ""
        For NumRegElim = 1 To Col.Count
            'Vemos la linea
            SQL = "numlinea = " & RecuperaValor(Col.Item(NumRegElim), 1)
            RN.Find SQL, , adSearchForward, 1
            SQL = ""
            If RN.EOF Then
                SQL = "No se encuentra la entrada en lotes. "
            Else
                Ca = CCur(RecuperaValor(Col.Item(NumRegElim), 2))
                If RN.Fields(1) <> Ca Then SQL = "Cantidad distinta lineas/lotes: " & Ca & " / " & RN.Fields(1)
            End If
            If SQL <> "" Then txtAnterior = txtAnterior & RecuperaValor(Col.Item(NumRegElim), 3) & " => " & SQL & vbCrLf
                
        Next
        RN.Close
        If txtAnterior <> "" Then
            txtAnterior = "Articulos TRAZABILIDAD" & vbCrLf & String(30, "=") & vbCrLf & txtAnterior
            txtAnterior = txtAnterior & vbCrLf & "¿Continuar?"
            If MsgBox(txtAnterior, vbQuestion + vbYesNo) = vbNo Then GoTo EComprobarNUmerosDeLote
        End If
        txtAnterior = ""
    End If
    ComprobarNUmerosDeLote = True
EComprobarNUmerosDeLote:
    If Err.Number <> 0 Then MuestraError Err.Number, "Funcion: ComprobarNUmerosDeLote"
    Set RN = Nothing
    Set Col = Nothing
End Function


Private Function ComprobarSiNoEstaEnOrdenCarga() As Boolean
Dim Aux As String
'Si esta en carga el albaran, ya han generado una serie de datos de carga con codigos de barra y demas. NO deberiamos dejar pasar
    ComprobarSiNoEstaEnOrdenCarga = True  'Sigue modificando
    If hcoCodTipoM <> "ALV" Then Exit Function
    
    If vParamAplic.ProduccionNueva Then
        Aux = "codtipom=1 and numalbar"
        Aux = DevuelveDesdeBD(conAri, "count(*)", "srepartolot", Aux, Me.Text1(0).Text)
        If Aux = "" Then Aux = "0"
        If Val(Aux) > 0 Then
            MsgBox "Esta ya en proceso de carga", vbExclamation
            If vUsu.Nivel > 1 Then ComprobarSiNoEstaEnOrdenCarga = False
        End If
    End If
End Function


Private Function ComprobarVinculado() As Boolean
Dim cad As String

    ComprobarVinculado = True
    If vParamAplic.EsAVAB Then Exit Function
    If hcoCodTipoM <> "ALV" Then Exit Function
    
    If DBLet(Data1.Recordset!refproduccion, "N") > 0 Then
        'Veremos si todavia este en pedido
        cad = DevuelveDesdeBD(conAri, "numpedcl", "ariges" & EmprAVAB & ".scaped", "numpedcl", Data1.Recordset!refproduccion)
        If cad = "" Then
            'YA se ha pasado a ALBARAN. Los lotes pueden
            cad = "El pedido en empresa exportación YA ha sido generado. Los cambios no se reflejaran" & vbCrLf & vbCrLf & "¿Continuar?"
            If MsgBox(cad, vbQuestion + vbYesNo) = vbNo Then ComprobarVinculado = False
        Else
            MsgBox "Los cambios NO se reflejaran en el pedido ", vbExclamation
        End If
    End If
    
    
End Function






Private Sub PackingList()
Dim cadFormula As String
Dim Cadparam As String
Dim NumParam As Byte
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim ImprimeDirecto As Boolean

    
    
    If Me.Data1.Recordset!Codtipom <> "ALV" Then
        MsgBox "Solo albaranes de venta", vbExclamation
        Exit Sub
    End If
    
    
    'Vamos a meter el lote y demas
    PonerDatosLote
    
    
    
    cadFormula = ""
    Cadparam = ""
    
    NumParam = 0
    

    If Not PonerParamRPT(34, Cadparam, NumParam, nomDocu, ImprimeDirecto) Then Exit Sub
    'El nombre sera el que tiene acabado en ALB
      
    NumParam = InStr(1, nomDocu, ".")
    devuelve = Mid(nomDocu, 1, NumParam - 1) & "ALB.rpt"
    nomDocu = devuelve
    
    
    'Cod Tipo Movimiento
    devuelve = "{slialb.codtipom}='" & CodTipoMov & "'"
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    
    'Nº Factura
    devuelve = "{slialb.numalbar}=" & Val(Text1(0).Text)
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    
    



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
Dim DAV As String

    On Error GoTo EPonerlotes
    
    
    
    SQL = "DELETE FROM tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL
    
    If Not vParamAplic.EsAVAB Then Exit Sub
    
    Set RS = New ADODB.Recordset
    
    
    SQL = ObtenerWhereCP(True)
    SQL = Replace(SQL, "scaalb.", "")
    SQL = "select * from slialblotes " & SQL
    RS.Open SQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    masDeUnLinea = "|"
    Aux = ""
    While Not RS.EOF
        Aux = "OK" 'para saber que tiene registros
        If RS!linea > 1 Then masDeUnLinea = masDeUnLinea & Format(RS!NumAlbar, "0000") & Format(RS!numlinea, "000") & "|"
        RS.MoveNext
    Wend
    If masDeUnLinea = "|" Then masDeUnLinea = ""
    If Aux <> "" Then RS.MoveFirst
    Aux = ""
    While Not RS.EOF
        SQL = " numalbar=" & RS!NumAlbar & " and codtipom='" & RS!Codtipom & "' AND numlinea "
        SQL = DevuelveDesdeBD(conAri, "codartic", "slialb", SQL, CStr(RS!numlinea))
        If SQL = "" Then
            MsgBox "No se encuentra el articulo para el lote: " & RS!numLote, vbExclamation
        Else
            TituloLinea = SQL  'codartic
            
            motivo = ""
            If masDeUnLinea <> "" Then
                motivo = Format(RS!NumAlbar, "0000") & Format(RS!numlinea, "000") & "|"
                If InStr(1, masDeUnLinea, motivo) = 0 Then motivo = ""
            End If
            If motivo <> "" Then motivo = "(" & RS!Cantidad & ")"
            
            'El lote sin la fecprod
            NumRegElim = InStr(RS!numLote, " ")
            If NumRegElim > 0 Then
                SQL = Mid(RS!numLote, 1, NumRegElim)
            Else
                SQL = Mid(RS!numLote, 5)
               
            End If
            motivo = SQL & motivo 'Aqui tendre ej: 9945(23) para el nº lote 9945 2011/10/21
            
            'Como es AVAB, para saber la fecha de produccion tendremos que irnos a ariges1 (morales)
            'MARZO 2012
            'La fecha de caducidad esta en la tabla de produccion (tambien estara en la de lotes
            'Con lo cual YA no es sumando 2 años a la de produccion. Habra que buscarla en la BD

            
                SQL = " codartic = " & DBSet(TituloLinea, "T") & " AND numlote=" & DBSet(RS!numLote, "T") & " AND 1"
                SQL = DevuelveDesdeBD(conAri, "numalbar", "ariges" & EmprMorales & ".spartidas", SQL, "1")
            
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
                    SQL = "codigo = " & Mid(SQL, 1, 5) & " AND idlin = " & Mid(SQL, 6) & " AND 1"
                    DAV = "feccaduca"
                    SQL = DevuelveDesdeBD(conAri, "fhinicio", "ariges" & EmprMorales & ".prodlin", SQL, "1", "N", DAV)
                End If
                
                If SQL <> "" Then
                    'OK ha conseguido la fecha de produccion. Verificamos la de caducidad

                    'Verifico la fecha de caducidad
                    If DAV <> "" Then
                        If Not IsDate(DAV) Then
                            MsgBox "Error obteniendo caducidad. Lote: " & RS!numLote
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
                If Mid(RS!numLote, 1, 4) = "0000" Then
                    'NO imprimo todas los numeros de lote
                    SQL = SQL & DBSet(RS!Cantidad, "N", "N") & "," & DBSet(Mid(RS!numLote, 5), "T") & ",'',''"
                Else
                    SQL = SQL & DBSet(RS!Cantidad, "N", "N") & "," & DBSet(RS!numLote, "T") & ",'',''"
                End If
            Else
                'OK, todo OK
                'tmpinformes(codusu,codigo1,campo1,campo2,importe1,nombre1,nombre2,nombre3)

                    
                F = CDate(SQL)
                SQL = vUsu.Codigo & "," & RS!NumAlbar & "," & RS!numlinea & "," & RS!linea & ","
                
                
                
                SQL = SQL & DBSet(RS!Cantidad, "N", "N") & "," & DBSet(motivo, "T") & ",'" & Format(F, "dd/mm/yyyy")
                F = CDate(DAV)
                SQL = SQL & "','" & Format(F, "dd/mm/yyyy") & "'"
                
            End If
            Aux = Aux & ", (" & SQL & ")"
        End If
        RS.MoveNext
    Wend
    RS.Close
    motivo = ""
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



Private Sub HectogradoVisible(Si As Boolean)
Dim b As Boolean
    b = Si
    If b Then
        If vParamAplic.QUE_EMPRESA <> 2 Then b = False
    End If
    Me.Label1(52).visible = b
    Me.Text2(15).visible = b
End Sub



Private Sub HazInsercionPack()
Dim OK As Boolean
Dim vsTk As cStock

    If (ModificaLineas = 1 Or ModificaLineas = 2) Then
        MsgBox "Esta editando linea", vbExclamation
        Exit Sub
    End If


    Set vsTk = New cStock

        vsTk.tipoMov = "S"
        vsTk.DetaMov = CodTipoMov
        vsTk.Trabajador = CLng(Text1(4).Text) 'guardamos el cliente del albaran
        vsTk.Documento = Text1(0).Text 'Nº Albaran
        vsTk.Fechamov = Text1(1).Text 'Fecha del Albaran
        vsTk.HoraMov = Text1(1).Text & " " & Format(Now, "hh:nn:ss")
         
     



    conn.BeginTrans
    OK = INsertaArticulosPAck(vsTk)
    If OK Then
        conn.CommitTrans
        'Cargalineas
        CargaGrid DataGrid1, Data2, True
    Else
        conn.RollbackTrans
    End If

End Sub


Private Function INsertaArticulosPAck(ByRef vCStock As cStock) As Boolean
 Dim cad As String
 Dim Ampliacion As String
Dim SQL As String
    On Error GoTo eINsertaArticulosPAck
    NumRegElim = 0
    cad = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas)
    NumRegElim = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", cad)
    
    
    

    cad = "SELECT salmac.codartic,nomartic,canstock,sarti6.cantidad,sartic.preciouc,codProve from salmac,sartic,sarti6"
    cad = cad & " where salmac.codartic=Sartic.codartic and codalmac=1 and"
    cad = cad & " salmac.codartic =sarti6.codarti1 AND  sarti6.codartic =" & DBSet(CadenaDesdeOtroForm, "T") & " ORDER BY numlinea"
    
    
    
    
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cad = ""
    While Not miRsAux.EOF
        
        'sOTCK
            
            
            vCStock.codAlmac = 1
            vCStock.Cantidad = miRsAux!Cantidad
            vCStock.Importe = CCur(DBLet(miRsAux!PrecioUC, "N") * vCStock.Cantidad)
            vCStock.LineaDocu = NumRegElim
            vCStock.codArtic = miRsAux!codArtic
        
            'La linea
            'Inserta en tabla "slialb"
            SQL = "INSERT INTO " & NomTablaLineas
            SQL = SQL & "(codtipom, numalbar,numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel,"
            SQL = SQL & " origpre,codprovex,cajas,PrecioLitro,palets,hectogrado) "
            SQL = SQL & "VALUES ('" & Text1(30).Text & "', " & Val(Text1(0).Text) & ", " & NumRegElim & ", " & vCStock.codAlmac & ","
            Ampliacion = "scaalb.codtipom=slialb.codtipom  and scaalb.numalbar=slialb.numalbar and scaalb.Codtipom='ALS'"
            Ampliacion = Ampliacion & " AND codartic =" & DBSet(miRsAux!codArtic, "T") & " AND 1"
            Ampliacion = DevuelveDesdeBD(conAri, "ampliaci", "scaalb,slialb", Ampliacion, " 1 ORDER BY fechaalb desc")
            
            SQL = SQL & DBSet(miRsAux!codArtic, "T") & ", " & DBSet(miRsAux!NomArtic, "T") & ", " & DBSet(Ampliacion, "T") & ", "
            '           cantidad,                   precioar,        dtoline1, dtoline2,        importel,
            SQL = SQL & DBSet(CInt(vCStock.Cantidad), "N") & "," & DBSet(DBLet(miRsAux!PrecioUC, "N"), "N", "N") & ", 0,0," & DBSet(DBLet(vCStock.Importe, "N"), "N", "N") & ","
            '               origpre,codprovex,cajas
            SQL = SQL & DBSet("A", "T") & ", " & DBSet(miRsAux!codProve, "N") & ",0,"
            '       ,PrecioLitro,palets,hectogrado) "
            SQL = SQL & "1,0,"
            'hectogrado
            'Hecto = 1
    '        If vParamAplic.QUE_EMPRESA = 2 Then
    '            If Not Text2(15).Locked Then
    '                Hecto = ImporteFormateado(Text2(15).Text)
    '                Hecto = Hecto / 100
    '            End If
    '        End If
            'SQL = SQL & DBSet(Hecto, "N", "N") & ")"
            SQL = SQL & DBSet(1, "N", "N") & ")"
            If vCStock.ActualizarStock Then
                conn.Execute SQL
            Else
                Err.Raise 513, , "Error actualizando stock"
            End If
        
            NumRegElim = NumRegElim + 1
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    INsertaArticulosPAck = True
    
eINsertaArticulosPAck:
  If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
  Set miRsAux = Nothing
  Set vCStock = Nothing
End Function
