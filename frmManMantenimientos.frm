VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmManMantenimientos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimientos"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11490
   Icon            =   "frmManMantenimientos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   11490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   910
      Left            =   120
      TabIndex        =   126
      Top             =   425
      Width           =   11175
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   8100
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "N� Mantenimiento|T|N|||scaman|nummante||S|"
         Text            =   "Text1"
         Top             =   160
         Width           =   1365
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   8100
         MaxLength       =   15
         TabIndex        =   3
         Tag             =   "Fecha Inicio|F|N|||scaman|fechaini|dd/mm/yyyy|N|"
         Text            =   "Text1"
         Top             =   510
         Width           =   1365
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   1
         Left            =   2240
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   128
         Text            =   "Text2"
         Top             =   510
         Width           =   3765
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   0
         Left            =   2240
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   127
         Text            =   "Text2"
         Top             =   160
         Width           =   3765
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   1400
         MaxLength       =   3
         TabIndex        =   1
         Tag             =   "C�d. Direcci�n|N|S|0|999|scaman|coddirec|000|N|"
         Text            =   "Text1"
         Top             =   510
         Width           =   765
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   1400
         MaxLength       =   6
         TabIndex        =   0
         Tag             =   "C�digo Cliente|N|N|0|999999|scaman|codclien|000000|S|"
         Text            =   "Text"
         Top             =   160
         Width           =   750
      End
      Begin VB.Label Label1 
         Caption         =   "N� Mantenim."
         Height          =   255
         Index           =   13
         Left            =   6840
         TabIndex        =   132
         Top             =   160
         Width           =   1335
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   7800
         Picture         =   "frmManMantenimientos.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   510
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Inicio"
         Height          =   255
         Index           =   14
         Left            =   6840
         TabIndex        =   131
         Top             =   510
         Width           =   975
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1120
         ToolTipText     =   "Buscar direc./dpto"
         Top             =   520
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1120
         Picture         =   "frmManMantenimientos.frx":0097
         ToolTipText     =   "Buscar cliente"
         Top             =   170
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. Direc."
         Height          =   255
         Index           =   1
         Left            =   160
         TabIndex        =   130
         Top             =   510
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "C�d. Cliente"
         Height          =   255
         Index           =   0
         Left            =   160
         TabIndex        =   129
         Top             =   160
         Width           =   855
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   21
      Left            =   3840
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   60
      Text            =   "Text2"
      Top             =   6960
      Visible         =   0   'False
      Width           =   4365
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   120
      TabIndex        =   43
      Top             =   6855
      Width           =   2535
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   240
         TabIndex        =   47
         Top             =   180
         Width           =   2115
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10290
      TabIndex        =   39
      Top             =   6960
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9000
      TabIndex        =   38
      Top             =   6960
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   2640
      Top             =   6960
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
      TabIndex        =   48
      Top             =   0
      Width           =   11490
      _ExtentX        =   20267
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   20
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
            Object.ToolTipText     =   "Revisiones"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Hist�rico"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Componentes"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "�ltimo"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   8400
         TabIndex        =   49
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   8520
      Top             =   1440
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
      Height          =   5520
      Left            =   120
      TabIndex        =   50
      Top             =   1335
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   9737
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Datos b�sicos"
      TabPicture(0)   =   "frmManMantenimientos.frx":0199
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(34)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(15)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(36)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(7)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "imgBuscar(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "imgBuscar(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(54)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(4)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(6)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(9)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text1(4)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text1(5)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text2(4)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text2(5)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cboTipoPago"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "chkBaterias"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text1(6)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Text1(7)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Text1(34)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Text1(35)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Text1(36)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Text1(37)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).ControlCount=   24
      TabCaption(1)   =   "Observaciones"
      TabPicture(1)   =   "frmManMantenimientos.frx":01B5
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text1(33)"
      Tab(1).Control(1)=   "Text1(32)"
      Tab(1).Control(2)=   "Label1(5)"
      Tab(1).Control(3)=   "Label1(3)"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Revisiones"
      TabPicture(2)   =   "frmManMantenimientos.frx":01D1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DataGrid1"
      Tab(2).Control(1)=   "TxtAux1(2)"
      Tab(2).Control(2)=   "TxtAux1(1)"
      Tab(2).Control(3)=   "TxtAux1(0)"
      Tab(2).Control(4)=   "cmdAux"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Hist�rico"
      TabPicture(3)   =   "frmManMantenimientos.frx":01ED
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Text2(28)"
      Tab(3).Control(1)=   "Text2(29)"
      Tab(3).Control(2)=   "Text2(30)"
      Tab(3).Control(3)=   "Text2(31)"
      Tab(3).Control(4)=   "Text2(32)"
      Tab(3).Control(5)=   "Text2(33)"
      Tab(3).Control(6)=   "Text2(41)"
      Tab(3).Control(7)=   "Text2(42)"
      Tab(3).Control(8)=   "Text2(43)"
      Tab(3).Control(9)=   "Text2(44)"
      Tab(3).Control(10)=   "Text2(45)"
      Tab(3).Control(11)=   "Text2(46)"
      Tab(3).Control(12)=   "Text2(47)"
      Tab(3).Control(13)=   "Text2(34)"
      Tab(3).Control(14)=   "Text2(22)"
      Tab(3).Control(15)=   "Text2(23)"
      Tab(3).Control(16)=   "Text2(24)"
      Tab(3).Control(17)=   "Text2(25)"
      Tab(3).Control(18)=   "Text2(26)"
      Tab(3).Control(19)=   "Text2(27)"
      Tab(3).Control(20)=   "Text2(35)"
      Tab(3).Control(21)=   "Text2(36)"
      Tab(3).Control(22)=   "Text2(37)"
      Tab(3).Control(23)=   "Text2(38)"
      Tab(3).Control(24)=   "Text2(39)"
      Tab(3).Control(25)=   "Text2(40)"
      Tab(3).Control(26)=   "imgFlecha(1)"
      Tab(3).Control(27)=   "imgFlecha(0)"
      Tab(3).Control(28)=   "Label1(53)"
      Tab(3).Control(29)=   "Label1(52)"
      Tab(3).Control(30)=   "Label1(51)"
      Tab(3).Control(31)=   "Label1(50)"
      Tab(3).Control(32)=   "Label1(49)"
      Tab(3).Control(33)=   "Label1(48)"
      Tab(3).Control(34)=   "Label1(47)"
      Tab(3).Control(35)=   "Label1(46)"
      Tab(3).Control(36)=   "Label1(45)"
      Tab(3).Control(37)=   "Label1(44)"
      Tab(3).Control(38)=   "Label1(43)"
      Tab(3).Control(39)=   "Label1(42)"
      Tab(3).Control(40)=   "Label1(41)"
      Tab(3).Control(41)=   "Label1(40)"
      Tab(3).Control(42)=   "Label1(39)"
      Tab(3).Control(43)=   "Label1(38)"
      Tab(3).Control(44)=   "Label1(37)"
      Tab(3).ControlCount=   45
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   37
         Left            =   3480
         MaxLength       =   30
         TabIndex        =   8
         Tag             =   "P|T|S|||scaman|attetiqu||N|"
         Text            =   "WWWWWWWWWWWWWWW"
         Top             =   840
         Width           =   2805
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   36
         Left            =   1680
         MaxLength       =   60
         TabIndex        =   11
         Tag             =   "P|T|S|||scaman|concefac||N|"
         Text            =   "WWWWWWWWW0WWWWWWWWW0WWWWWWWWW0WWWWWWWWW0WWWWWWWWW0WWWWWWWW60"
         Top             =   1320
         Width           =   8085
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   35
         Left            =   7080
         MaxLength       =   15
         TabIndex        =   13
         Tag             =   "P|T|S|||scaman|producto||N|"
         Text            =   "WWWWWWWWWWWWWWW"
         Top             =   1800
         Width           =   2805
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   34
         Left            =   1680
         MaxLength       =   35
         TabIndex        =   12
         Tag             =   "P|T|S|||scaman|persconta||N|"
         Text            =   "WWWWWWWWW0WWWWWWWWW0WWWWWWWWW0"
         Top             =   1800
         Width           =   4245
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   435
         Left            =   -72480
         TabIndex        =   133
         Top             =   3600
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   28
         Left            =   -68400
         MaxLength       =   15
         TabIndex        =   95
         Text            =   "Text2"
         Top             =   1200
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   29
         Left            =   -68400
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   94
         Text            =   "Text2"
         Top             =   1560
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   30
         Left            =   -68400
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   93
         Text            =   "Text2"
         Top             =   1920
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   31
         Left            =   -68400
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   92
         Text            =   "Text2"
         Top             =   2280
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   32
         Left            =   -68400
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   91
         Text            =   "Text2"
         Top             =   2640
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   33
         Left            =   -68400
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   90
         Text            =   "Text2"
         Top             =   3000
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   41
         Left            =   -66720
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   89
         Text            =   "Text2"
         Top             =   1200
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   42
         Left            =   -66720
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   88
         Text            =   "Text2"
         Top             =   1560
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   43
         Left            =   -66720
         MaxLength       =   15
         TabIndex        =   87
         Text            =   "Text2"
         Top             =   1920
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   44
         Left            =   -66720
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   86
         Text            =   "Text2"
         Top             =   2280
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   45
         Left            =   -66720
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   85
         Text            =   "Text2"
         Top             =   2640
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   46
         Left            =   -66720
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   84
         Text            =   "Text2"
         Top             =   3000
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   47
         Left            =   -66720
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   83
         Text            =   "Text2"
         Top             =   3480
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   34
         Left            =   -68400
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   82
         Text            =   "Text2"
         Top             =   3480
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   22
         Left            =   -73395
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   73
         Text            =   "Text2"
         Top             =   1200
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   23
         Left            =   -73395
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   72
         Text            =   "Text2"
         Top             =   1560
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   24
         Left            =   -73395
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   71
         Text            =   "Text2"
         Top             =   1920
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   25
         Left            =   -73395
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   70
         Text            =   "Text2"
         Top             =   2280
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   26
         Left            =   -73395
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   69
         Text            =   "Text2"
         Top             =   2640
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   27
         Left            =   -73395
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   68
         Text            =   "Text2"
         Top             =   3000
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   35
         Left            =   -71760
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   67
         Text            =   "Text2"
         Top             =   1200
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   36
         Left            =   -71760
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   66
         Text            =   "Text2"
         Top             =   1560
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   37
         Left            =   -71760
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   65
         Text            =   "Text2"
         Top             =   1920
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   38
         Left            =   -71760
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   64
         Text            =   "Text2"
         Top             =   2280
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   39
         Left            =   -71760
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   63
         Text            =   "Text2"
         Top             =   2640
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   40
         Left            =   -71760
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   62
         Text            =   "Text2"
         Top             =   3000
         Width           =   1485
      End
      Begin VB.TextBox TxtAux1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   -74520
         MaxLength       =   15
         TabIndex        =   44
         Tag             =   "Fecha Rev.|F|N|||slima1|fecharev||N|"
         Text            =   "F. Revision"
         Top             =   3600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox TxtAux1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   -73200
         MaxLength       =   4
         TabIndex        =   45
         Tag             =   "Cod. Traba|N|N|0|9999|slima1|codtraba|0000|N|"
         Text            =   "codtra"
         Top             =   3600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox TxtAux1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   555
         Index           =   2
         Left            =   -72360
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   46
         Tag             =   "Observaciones|T|N|||slima1|observac||N|"
         Text            =   "frmManMantenimientos.frx":0209
         Top             =   3540
         Visible         =   0   'False
         Width           =   8055
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   7
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   5
         Tag             =   "Anticipado Sig.|N|S|0||scaman|anticip2|##,###,##0.00|N|"
         Text            =   "Text1"
         Top             =   840
         Width           =   1365
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   6
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   4
         Tag             =   "Anticipado Act.|N|S|0||scaman|anticip1|##,###,##0.00|N|"
         Text            =   "Text1"
         Top             =   480
         Width           =   1365
      End
      Begin VB.TextBox Text1 
         Height          =   1755
         Index           =   33
         Left            =   -74040
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   41
         Tag             =   "Observ. T�cnico|T|S|||scaman|obsertec||N|"
         Text            =   "frmManMantenimientos.frx":02B2
         Top             =   3360
         Width           =   9645
      End
      Begin VB.TextBox Text1 
         Height          =   1755
         Index           =   32
         Left            =   -74040
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   40
         Tag             =   "Observ. Comercial|T|S|||scaman|observac||N|"
         Text            =   "frmManMantenimientos.frx":02BA
         Top             =   840
         Width           =   9645
      End
      Begin VB.CheckBox chkBaterias 
         Caption         =   "Baterias"
         Height          =   255
         Left            =   5280
         TabIndex        =   7
         Tag             =   "Bater�as|N|N|||scaman|baterias||N|"
         Top             =   480
         Width           =   975
      End
      Begin VB.ComboBox cboTipoPago 
         Height          =   315
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Tag             =   "Tipo de Pago|N|N|||scaman|tipopago||N|"
         Top             =   495
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   5
         Left            =   8160
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   55
         Text            =   "Text2"
         Top             =   860
         Width           =   2805
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   4
         Left            =   8160
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   54
         Text            =   "Text2"
         Top             =   500
         Width           =   2805
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   7605
         MaxLength       =   30
         TabIndex        =   10
         Tag             =   "Forma de Pago|N|N|0|999|scaman|codforpa|000|N|"
         Text            =   "Text1"
         Top             =   860
         Width           =   525
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   7590
         MaxLength       =   30
         TabIndex        =   9
         Tag             =   "Tipo Contrato|T|N|||scaman|codtipco||N|"
         Text            =   "Text1"
         Top             =   500
         Width           =   540
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmManMantenimientos.frx":02C0
         Height          =   4620
         Left            =   -74640
         TabIndex        =   59
         Top             =   555
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   8149
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
      Begin VB.Frame Frame2 
         Height          =   3195
         Left            =   120
         TabIndex        =   106
         Top             =   2160
         Width           =   10740
         Begin VB.ComboBox cmbMes 
            Height          =   315
            Left            =   2640
            Style           =   2  'Dropdown List
            TabIndex        =   137
            Tag             =   "Ultimo mes facturado|N|N|1||scaman|ulmesfac||N|"
            Top             =   2760
            Width           =   2175
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   20
            Left            =   6720
            MaxLength       =   15
            TabIndex        =   20
            Tag             =   "Julio Actual|N|S|0||scaman|mes07act|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   480
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   22
            Left            =   6720
            MaxLength       =   15
            TabIndex        =   21
            Tag             =   "Agosto Actual|N|S|0||scaman|mes08act|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   840
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   24
            Left            =   6720
            MaxLength       =   15
            TabIndex        =   22
            Tag             =   "Septiembre Actual|N|S|0||scaman|mes09act|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   1200
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   26
            Left            =   6720
            MaxLength       =   15
            TabIndex        =   23
            Tag             =   "Octubre Actual|N|S|0||scaman|mes10act|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   1560
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   28
            Left            =   6720
            MaxLength       =   15
            TabIndex        =   24
            Tag             =   "Noviembre Actual|N|S|0||scaman|mes11act|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   1920
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   30
            Left            =   6720
            MaxLength       =   15
            TabIndex        =   25
            Tag             =   "Diciembre Actual|N|S|0||scaman|mes12act|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   2280
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   21
            Left            =   8400
            MaxLength       =   15
            ScrollBars      =   2  'Vertical
            TabIndex        =   32
            Tag             =   "Julio Siguiente|N|S|0||scaman|mes07sig|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   480
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   23
            Left            =   8400
            MaxLength       =   15
            TabIndex        =   33
            Tag             =   "Agosto Siguiente|N|S|0||scaman|mes08sig|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   840
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   25
            Left            =   8400
            MaxLength       =   15
            TabIndex        =   34
            Tag             =   "Septiembre Siguiente|N|S|0||scaman|mes09sig|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   1200
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   27
            Left            =   8400
            MaxLength       =   15
            TabIndex        =   35
            Tag             =   "Octubre Siguiente|N|S|0||scaman|mes10sig|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   1560
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   29
            Left            =   8400
            MaxLength       =   15
            TabIndex        =   36
            Tag             =   "Noviembre Siguiente|N|S|0||scaman|mes11sig|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   1920
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   31
            Left            =   8400
            MaxLength       =   15
            TabIndex        =   37
            Tag             =   "Diciembre Siguiente|N|S|0||scaman|mes12sig|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   2280
            Width           =   1485
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   3
            Left            =   8400
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   116
            Text            =   "Text2"
            Top             =   2760
            Width           =   1485
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   2
            Left            =   6720
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   115
            Text            =   "Text2"
            Top             =   2760
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   8
            Left            =   1605
            MaxLength       =   15
            TabIndex        =   14
            Tag             =   "Enero Actual|N|S|0||scaman|mes01act|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   480
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   10
            Left            =   1605
            MaxLength       =   15
            TabIndex        =   15
            Tag             =   "Febrero Actual|N|S|0||scaman|mes02act|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   840
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   12
            Left            =   1605
            MaxLength       =   15
            TabIndex        =   16
            Tag             =   "Marzo Actual|N|S|0||scaman|mes03act|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   1200
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   14
            Left            =   1605
            MaxLength       =   15
            TabIndex        =   17
            Tag             =   "Abril Actual|N|S|0||scaman|mes04act|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   1560
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   16
            Left            =   1605
            MaxLength       =   15
            TabIndex        =   18
            Tag             =   "Mayo Actual|N|S|0||scaman|mes05act|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   1920
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   18
            Left            =   1605
            MaxLength       =   15
            TabIndex        =   19
            Tag             =   "Junio Actual|N|S|0||scaman|mes06act|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   2280
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   9
            Left            =   3240
            MaxLength       =   15
            TabIndex        =   26
            Tag             =   "Enero Siguiente|N|S|0||scaman|mes01sig|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   480
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   11
            Left            =   3240
            MaxLength       =   15
            TabIndex        =   27
            Tag             =   "Febrero Siguiente|N|S|0||scaman|mes02sig|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   840
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   13
            Left            =   3240
            MaxLength       =   15
            TabIndex        =   28
            Tag             =   "Marzo Siguiente|N|S|0||scaman|mes03sig|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   1200
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   15
            Left            =   3240
            MaxLength       =   15
            TabIndex        =   29
            Tag             =   "Abril Siguiente|N|S|0||scaman|mes04sig|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   1560
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   17
            Left            =   3240
            MaxLength       =   15
            TabIndex        =   30
            Tag             =   "Mayo Siguiente|N|S|0||scaman|mes05sig|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   1920
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   19
            Left            =   3240
            MaxLength       =   15
            TabIndex        =   31
            Tag             =   "Junio Siguiente|N|S|0||scaman|mes06sig|##,###,##0.00|N|"
            Text            =   "Text1"
            Top             =   2280
            Width           =   1485
         End
         Begin VB.Label Label1 
            Caption         =   "Ultimo mes facturado"
            Height          =   195
            Index           =   8
            Left            =   960
            TabIndex        =   138
            Top             =   2840
            Width           =   1485
         End
         Begin VB.Label Label1 
            Caption         =   "Julio"
            Height          =   255
            Index           =   24
            Left            =   5880
            TabIndex        =   125
            Top             =   585
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Agosto"
            Height          =   255
            Index           =   25
            Left            =   5880
            TabIndex        =   124
            Top             =   900
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Septiembre"
            Height          =   255
            Index           =   26
            Left            =   5880
            TabIndex        =   123
            Top             =   1260
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Octubre"
            Height          =   255
            Index           =   27
            Left            =   5880
            TabIndex        =   122
            Top             =   1620
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Noviembre"
            Height          =   255
            Index           =   28
            Left            =   5880
            TabIndex        =   121
            Top             =   1965
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Diciembre"
            Height          =   255
            Index           =   29
            Left            =   5880
            TabIndex        =   120
            Top             =   2325
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "ACTUAL"
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
            Index           =   32
            Left            =   6840
            TabIndex        =   119
            Top             =   240
            Width           =   1485
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "SIGUIENTE"
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
            Index           =   33
            Left            =   8400
            TabIndex        =   118
            Top             =   240
            Width           =   1485
         End
         Begin VB.Label Label1 
            Caption         =   "TOTAL"
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
            Index           =   21
            Left            =   5880
            TabIndex        =   117
            Top             =   2760
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Enero"
            Height          =   255
            Index           =   16
            Left            =   960
            TabIndex        =   114
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Febrero"
            Height          =   255
            Index           =   17
            Left            =   960
            TabIndex        =   113
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Marzo"
            Height          =   255
            Index           =   19
            Left            =   960
            TabIndex        =   112
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Abril"
            Height          =   255
            Index           =   20
            Left            =   960
            TabIndex        =   111
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Mayo"
            Height          =   255
            Index           =   22
            Left            =   960
            TabIndex        =   110
            Top             =   1920
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Junio"
            Height          =   255
            Index           =   23
            Left            =   960
            TabIndex        =   109
            Top             =   2280
            Width           =   615
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "ACTUAL"
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
            Index           =   30
            Left            =   1680
            TabIndex        =   108
            Top             =   240
            Width           =   1365
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "SIGUIENTE"
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
            Index           =   31
            Left            =   3240
            TabIndex        =   107
            Top             =   240
            Width           =   1485
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Att ETIQ."
         Height          =   255
         Index           =   9
         Left            =   2640
         TabIndex        =   139
         Top             =   885
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Concepto factura"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   136
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Producto"
         Height          =   255
         Index           =   4
         Left            =   6240
         TabIndex        =   135
         Top             =   1840
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Persona contacto"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   134
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Image imgFlecha 
         Height          =   480
         Index           =   1
         Left            =   -70680
         Picture         =   "frmManMantenimientos.frx":02D5
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image imgFlecha 
         Height          =   480
         Index           =   0
         Left            =   -73440
         Picture         =   "frmManMantenimientos.frx":0717
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Anticipado 2"
         Height          =   255
         Index           =   54
         Left            =   240
         TabIndex        =   105
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Julio"
         Height          =   255
         Index           =   53
         Left            =   -69240
         TabIndex        =   104
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Agosto"
         Height          =   255
         Index           =   52
         Left            =   -69240
         TabIndex        =   103
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Septiembre"
         Height          =   255
         Index           =   51
         Left            =   -69240
         TabIndex        =   102
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Octubre"
         Height          =   255
         Index           =   50
         Left            =   -69240
         TabIndex        =   101
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Noviembre"
         Height          =   255
         Index           =   49
         Left            =   -69240
         TabIndex        =   100
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Diciembre"
         Height          =   255
         Index           =   48
         Left            =   -69240
         TabIndex        =   99
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "ACTUAL"
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
         Index           =   47
         Left            =   -68400
         TabIndex        =   98
         Top             =   840
         Width           =   1485
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "SIGUIENTE"
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
         Index           =   46
         Left            =   -66720
         TabIndex        =   97
         Top             =   840
         Width           =   1485
      End
      Begin VB.Label Label1 
         Caption         =   "TOTAL"
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
         Index           =   45
         Left            =   -69240
         TabIndex        =   96
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Enero"
         Height          =   255
         Index           =   44
         Left            =   -74280
         TabIndex        =   81
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Febrero"
         Height          =   255
         Index           =   43
         Left            =   -74280
         TabIndex        =   80
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Marzo"
         Height          =   255
         Index           =   42
         Left            =   -74280
         TabIndex        =   79
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Abril"
         Height          =   255
         Index           =   41
         Left            =   -74280
         TabIndex        =   78
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Mayo"
         Height          =   255
         Index           =   40
         Left            =   -74280
         TabIndex        =   77
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Junio"
         Height          =   255
         Index           =   39
         Left            =   -74280
         TabIndex        =   76
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "ACTUAL"
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
         Index           =   38
         Left            =   -72840
         TabIndex        =   75
         Top             =   840
         Width           =   885
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "SIGUIENTE"
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
         Index           =   37
         Left            =   -71760
         TabIndex        =   74
         Top             =   840
         Width           =   1125
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   7320
         ToolTipText     =   "Buscar forma de pago"
         Top             =   885
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   7320
         ToolTipText     =   "Buscar tipo contrato"
         Top             =   510
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Anticipado 1"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   58
         Top             =   495
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Observaci�n T�cnico"
         Height          =   255
         Index           =   5
         Left            =   -74040
         TabIndex        =   57
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Observaci�n Comercial"
         Height          =   255
         Index           =   3
         Left            =   -74040
         TabIndex        =   56
         Top             =   540
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Pago"
         Height          =   255
         Index           =   36
         Left            =   2640
         TabIndex        =   53
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Forma Pago"
         Height          =   255
         Index           =   15
         Left            =   6360
         TabIndex        =   52
         Top             =   855
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Contrato"
         Height          =   255
         Index           =   34
         Left            =   6360
         TabIndex        =   51
         Top             =   495
         Width           =   975
      End
   End
   Begin MSAdodcLib.Adodc Data3 
      Height          =   330
      Left            =   7560
      Top             =   7080
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
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   10290
      TabIndex        =   42
      Top             =   6960
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "T�cnico"
      Height          =   255
      Index           =   35
      Left            =   3165
      TabIndex        =   61
      Top             =   6960
      Visible         =   0   'False
      Width           =   615
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
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnMtoLineas 
      Caption         =   "&Mantenimiento Lineas"
      Begin VB.Menu mnRevisiones 
         Caption         =   "&Revisiones"
         HelpContextID   =   2
         Shortcut        =   ^R
      End
      Begin VB.Menu mnHistorico 
         Caption         =   "&Hist�rico"
         Shortcut        =   ^H
      End
   End
End
Attribute VB_Name = "frmManMantenimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmC As frmFacClientes 'Form Mto Clientes
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmTC As frmManTiposContrato 'Form Mto Tipos Contrato
Attribute frmTC.VB_VarHelpID = -1
Private WithEvents frmFP As frmFacFormasPago 'Form Mto Formas de Pago
Attribute frmFP.VB_VarHelpID = -1
Private WithEvents frmT As frmAdmTrabajadores
Attribute frmT.VB_VarHelpID = -1

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

Dim NumTabMto As Byte
'Indica que numero de Tab que esta en modo Mantenimiento

Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

Dim PrimeraVez As Boolean

Dim EsCabecera As Boolean
'Para saber en MandaBusquedaPrevia si busca en la tabla scaman o en la tabla sdirec

'SQL de la tabla principal del formulario
Private CadenaConsulta As String

Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private NomTablaLineas As String 'Nombre de la Tabla de lineas del Mantenimiento en que estemos
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Dim btnAnyadir As Byte
'Variable que indica el n�mero del Boton  Anyadir en la Toolbar1
Dim btnPrimero As Byte
'Variable que indica el n�mero del Boton  PrimerRegistro en la Toolbar1


Private Sub cboTipoPago_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkBaterias_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub chkBaterias_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub cmdAceptar_Click()
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
              If InsertarDesdeForm(Me) Then
                    Colorines
                    PosicionarData
              End If
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me, 1) Then
                    Colorines
                    TerminaBloquear
                    PosicionarData
                End If
            End If
            
         Case 5 'InsertarModificar linea
            'Actualizar el registro en la tabla de lineas 'slima1' (Revisiones)
            If ModificaLineas = 1 Then 'INSERTAR lineas
                If InsertarLinea Then 'Revisiones
                    If Me.SSTab1.Tab = 2 Then CargaGrid DataGrid1, Data2, True
                    BotonAnyadirLinea
                End If
            ElseIf ModificaLineas = 2 Then 'MODIFICAR lineas
                If ModificarLinea Then
                    TerminaBloquear
                    PonerBotonCabecera True
                    ModificaLineas = 0
                    If Me.SSTab1.Tab = 2 Then 'Habilidades
                        LLamaLineas 10
                        CargaGrid2 DataGrid1, Data2
                    End If
                    PonerFocoBtn Me.cmdRegresar
                End If
            End If
    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdAux_Click()
    'Abre Formulario de Mantenimiento de Trabajadores
    Set frmT = New frmAdmTrabajadores
    frmT.DatosADevolverBusqueda = "0|1|"
    frmT.Show vbModal
    Set frmT = Nothing
    PonerFoco Me.TxtAux1(1)
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
            If Me.SSTab1.Tab = 2 Then 'Revisiones
                If ModificaLineas = 1 Then 'INSERTAR
                    ModificaLineas = 0
                    DataGrid1.AllowAddNew = False
                    If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
                End If
                ModificaLineas = 0
                LLamaLineas 10
                DataGrid1.Enabled = True
            End If
            PonerBotonCabecera True
    End Select
End Sub


Private Sub BotonAnyadir()
'A�adir registro en tabla de trabajadores: straba (Cabecera)

    LimpiarCampos 'Vac�a los TextBox
    LimpiarCamposHistorico
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    Colorines
    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    Text1(3).Text = Format(Now, "dd/mm/yyyy")
    PonerFoco Text1(0)
    cmbMes.ListIndex = 0
End Sub


Private Sub BotonAnyadirLinea()
Dim anc As Single
Dim I As Byte

    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
        
    If NumTabMto <> Me.SSTab1.Tab Then
        MsgBox "No puede A�adir. Esta en Modo Mantenimiento de otra linea.", vbExclamation
        Exit Sub
    End If
    
    ModificaLineas = 1 'Ponemos Modo A�adir Linea
    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerBotonCabecera False
    lblIndicador.Caption = "INSERTAR"
    
    If Me.SSTab1.Tab = 2 Then 'Revisiones
        AnyadirLinea DataGrid1, Data2
        For I = 0 To Me.TxtAux1.Count - 1
            Me.TxtAux1(I).Text = ""
        Next I
        anc = ObtenerAlto(Me.DataGrid1) + 10
        LLamaLineas anc
        BloquearTxt TxtAux1(0), False
        PonerFoco TxtAux1(0)
    End If
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
'Ver todos
    LimpiarCampos
    LimpiarCamposHistorico
    LimpiarDataGrids
    
    If chkVistaPrevia.Value = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    If Data1.Recordset.EOF Then Exit Sub
    Text2(21).Text = ""
    DesplazamientoData Data1, Index
    PonerCampos
End Sub


Private Sub BotonModificar()
    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    PonerFoco Text1(3)
End Sub


Private Sub BotonModificarLinea()
'Modificar una linea
Dim vWhere As String
Dim anc As Single

    On Error GoTo EModificarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub '1= Insertar
    
    If NumTabMto <> Me.SSTab1.Tab Then
        MsgBox "No puede Modificar. Esta en Modo Mantenimiento de otra linea.", vbExclamation
        Exit Sub
    End If
    
    ModificaLineas = 2 'Modificar
    
    If Me.SSTab1.Tab = 2 Then 'Revisiones
         If Data2.Recordset.EOF Then Exit Sub
          vWhere = ObtenerWhereCP(False) & " and fecharev='" & Format(Data2.Recordset!FechaRev, FormatoFecha) & "'"
         If Not BloqueaRegistro(NomTablaLineas, vWhere) Then Exit Sub
         anc = ObtenerAlto(Me.DataGrid1) + 10
         'Llamamos al form
         Me.TxtAux1(0).Text = DataGrid1.Columns(2).Text
         Me.TxtAux1(1).Text = DataGrid1.Columns(3).Text
         Me.TxtAux1(2).Text = DataGrid1.Columns(4).Text
         LLamaLineas anc
         DataGrid1.Enabled = False
         BloquearTxt TxtAux1(0), True
         PonerFoco TxtAux1(1)
    End If
    
   
    'A�adiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR"
    PonerBotonCabecera False
    
EModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Mantenimientos (scaman)
' y los registros correspondientes de las tablas de lineas (sliman y slima1)
Dim cad As String
Dim B As Boolean


    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    'ANTES DE CREAR EL HCO DE MANTENIMIENTOS ANULADOS
    
'    cad = "Cabecera de Mantenimientos." & vbCrLf
'    cad = cad & "-----------------------------------" & vbCrLf & vbCrLf
'    cad = cad & "Va a eliminar el Mantenimiento:            "
'    cad = cad & vbCrLf & "Cliente:  " & Format(Text1(0).Text, "000000") & " - " & Text2(0).Text
''    cad = cad & vbCrLf & "Direc.:  " & Format(Text1(1).Text, "000") & " - " & Text2(1).Text
'    cad = cad & vbCrLf & "N� Mante.:  " & Text1(2).Text
'    cad = cad & vbCrLf & vbCrLf & " �Desea Eliminarlo? "
'
'    'Borramos
'    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
'        'Hay que eliminar
'        On Error GoTo EEliminar
'        Screen.MousePointer = vbHourglass
'        NumRegElim = Data1.Recordset.AbsolutePosition
'        If Not Eliminar Then
'            Exit Sub
'        ElseIf SituarDataTrasEliminar(Data1, NumRegElim) Then
'            PonerCampos
'        Else
'            LimpiarCampos
'            'Poner los grid sin apuntar a nada
'            LimpiarDataGrids
'            PonerModo 0
'        End If
'    End If


'       AHORA que hay hco
    CadenaDesdeOtroForm = ""
    frmListado.OpcionListado = 99
    frmListado.Show vbModal
        
    cad = ""
    If CadenaDesdeOtroForm <> "" Then cad = RecuperaValor(CadenaDesdeOtroForm, 1)
    If cad = "OK" Then
        
        Conn.BeginTrans
        B = False
        NumRegElim = Data1.Recordset.AbsolutePosition
        cad = Traspasa_A_Anulados
        If cad = "" Then
            If Eliminar Then B = True
        End If
        
        If B Then
            Conn.CommitTrans
            Espera 0.3
            If SituarDataTrasEliminar(Data1, NumRegElim) Then
                PonerCampos
            Else
                LimpiarCampos
                'Poner los grid sin apuntar a nada
                LimpiarDataGrids
                PonerModo 0
            End If
    
        Else
            Conn.RollbackTrans
        End If
                    
    End If
    
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Mantenimiento", Err.Description
End Sub


Private Sub BotonEliminarLinea()
'Eliminar una linea De Mantenimiento. Tabla: slima1
Dim SQL As String
Dim FechaRev As Date

    On Error GoTo EEliminarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar

     If NumTabMto <> Me.SSTab1.Tab Then
        MsgBox "No puede eliminar. Esta en Modo Mantenimiento de otra linea.", vbExclamation
        Exit Sub
    End If

    If Me.SSTab1.Tab = 2 Then 'Revisiones
        If Data2.Recordset.EOF Then Exit Sub
        FechaRev = Data2.Recordset!FechaRev
    End If
    
    ModificaLineas = 3 'Eliminar
    SQL = "�Seguro que desea eliminar la l�nea de " & TituloLinea & "?      " & vbCrLf
    SQL = SQL & vbCrLf & "Fec. Rev.: " & FechaRev
    SQL = SQL & vbCrLf & " T�cnico: " & Format(Data2.Recordset!CodTraba, "0000") & " - " & Text2(21).Text
    
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        SQL = "Delete from " & NomTablaLineas & ObtenerWhereCP(True)
        SQL = SQL & " and fecharev='" & Format(FechaRev, FormatoFecha) & "'"
        Conn.Execute SQL
        ModificaLineas = 0
        If Me.SSTab1.Tab = 2 Then CargaGrid2 DataGrid1, Data2 'Revisiones
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
        PonerModo 2
        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    Else 'Se llama desde alg�n Prismatico de otro Form al Mantenimiento de Trabajadores
        If Data1.Recordset.EOF Then
            MsgBox "Ning�n registro devuelto.", vbExclamation
            Exit Sub
        End If
        cad = Data1.Recordset.Fields(0) & "|"
        cad = cad & Data1.Recordset.Fields(1) & "|"
        RaiseEvent DatoSeleccionado(cad)
        Unload Me
    End If
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo Error1

    If Not Data2.Recordset.EOF And ModificaLineas <> 1 Then 'No en Insertar
        'Poner descripcion del Trabajador
        Text2(21).Text = DevuelveDesdeBDNew(conAri, "straba", "nomtraba", "codtraba", Data2.Recordset!CodTraba.Value, "N")
    Else
        Text2(21).Text = ""
    End If
    
Error1:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    'Icono de busqueda
    For kCampo = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(kCampo).Picture = frmPpal.imgListComun.ListImages(19).Picture
    Next kCampo

    ' ICONITOS DE LA BARRA
    btnAnyadir = 5
    btnPrimero = 17
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Bot�n Buscar
        .Buttons(2).Image = 2   'Bot�n Todos
        .Buttons(5).Image = 3   'Insertar Nuevo
        .Buttons(6).Image = 4   'Modificar
        .Buttons(7).Image = 5   'Borrar
        .Buttons(10).Image = 10 'Mto Lineas Revisiones
        .Buttons(11).Image = 38 'Mto Lineas Hist�rico
        .Buttons(12).Image = 34 'Componentes
        .Buttons(14).Image = 15  'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 '�ltimo
    End With
    Me.SSTab1.Tab = 0
          
    CargaComboMes
    LimpiarCampos   'Limpia los campos TextBox
    CargarComboTipoPago
   
    'Comprobar si es Departamento o Direccion
    If vParamAplic.Departamento Then
        Me.Label1(1).Caption = "Cod. Dpto"
    Else
        Me.Label1(1).Caption = "Cod. Direc."
    End If
        
    '## A mano
    NombreTabla = "scaman"
    Ordenacion = " ORDER BY codclien, nummante "
         
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = Conn
    Data1.RecordSource = "Select * from " & NombreTabla & " where codclien=-1"
    Data1.Refresh
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1
        Text1(0).BackColor = vbYellow
    End If
    
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
    NomTablaLineas = "slima1" 'Tabla lineas de Revisiones de MAntenimientos
    Data2.ConnectionString = Conn
    Data2.RecordSource = "Select * from " & NomTablaLineas & " where codclien=-1"
    Data2.Refresh
    CargaGrid DataGrid1, Data2, False
    
    'Cargamos inicialmente el DATA3 a nada
    Data3.ConnectionString = Conn
    Data3.RecordSource = "select * from sliman where codclien=-1"
    Data3.Refresh
End Sub


Private Sub LimpiarCampos()
On Error Resume Next

    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.cboTipoPago.ListIndex = -1
    Me.chkBaterias.Value = False
    Me.cmbMes.ListIndex = -1
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
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
'            Aux = Aux & " and " & ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
            Aux = Aux & " and " & ValorDevueltoFormGrid(Text1(2), CadenaDevuelta, 3)
            cadB = Aux
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
        Else 'Llama desde Prismatico Direcciones/Departamentos
            Text1(1).Text = Format(RecuperaValor(CadenaDevuelta, 1), "000")
            Text2(1).Text = RecuperaValor(CadenaDevuelta, 2)
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Clientes
    Text1(0).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000") 'Cod Clien
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Clien
End Sub


Private Sub frmF_Selec(vFecha As Date) 'Calendario Fechas
    Text1(3).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Formas de Pago
    Text1(5).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Forma Pago
    Text2(5).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Forma Pago
End Sub


Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Trabajadores
    Me.TxtAux1(1).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod trabajador
    FormateaCampo Me.TxtAux1(1)
    Text2(21).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom trabajador
End Sub

Private Sub frmTC_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Tipos de Contrato
    Text1(4).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Tipo Contrato
    Text2(4).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Tipo Contrato
End Sub


Private Sub imgBuscar_Click(Index As Integer)
Dim indice As Byte

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cod. Cliente
            Set frmC = New frmFacClientes
            frmC.DatosADevolverBusqueda = "0"
            frmC.Show vbModal
            Set frmC = Nothing
            indice = Index
            
        Case 1 'Cod. Direc.
             'Mostrar las Direc. o Dptos del cliente seleccionado
             If Trim(Text1(0).Text) = "" Then
                MsgBox "Debe seleccionar un cliente.", vbInformation
                Screen.MousePointer = vbDefault
                Exit Sub
             Else
                EsCabecera = False
                MandaBusquedaPrevia " codclien= " & Val(Text1(0).Text)
                indice = 1
             End If
             
        Case 2 'Cod. Tipo Contrato
            Set frmTC = New frmManTiposContrato
            frmTC.DatosADevolverBusqueda = "0"
            frmTC.Show vbModal
            Set frmTC = Nothing
            indice = 4
            
        Case 3 'Cod. Formas de Pago
            Set frmFP = New frmFacFormasPago
            frmFP.DatosADevolverBusqueda = "0"
            frmFP.Show vbModal
            Set frmFP = Nothing
            indice = 5
    End Select
    PonerFoco Text1(indice)
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFecha_Click(Index As Integer) 'Abre calendario Fechas
Dim indice As Byte

   If Modo = 2 Or Modo = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass
   Set frmF = New frmCal
   frmF.Fecha = Now
   indice = Index + 3
   
   PonerFormatoFecha Text1(indice)
   If Text1(indice).Text <> "" Then frmF.Fecha = CDate(Text1(indice).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco Text1(indice)
End Sub


Private Sub imgFlecha_Click(Index As Integer)
'Desplazarse por los dos registros siguientes del hist�rico
On Error GoTo EDesplazamiento
    
    Select Case Index
        Case 0 'Anterior
            If Data3.Recordset.BOF Then Exit Sub
            If (Data3.Recordset.AbsolutePosition Mod 2) = 0 Then 'numero par
                Data3.Recordset.Move -3
            Else 'numero impar
                Data3.Recordset.Move -2
            End If
            If Data3.Recordset.BOF Then Exit Sub
        Case 1 'Siguiente
            If Data3.Recordset.EOF Then Exit Sub
            If Not Data3.Recordset.BOF Then
                Data3.Recordset.MoveNext
            Else
                Data3.Recordset.Move 3
            End If
    End Select
    PonerCamposHistorico
    
EDesplazamiento:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub mnBuscar_Click()
    Me.SSTab1.Tab = 0
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar lineas de Mantenimiento
         BotonEliminarLinea
    Else   'Eliminar Mantenimiento
         BotonEliminar
    End If
End Sub


Private Sub mnHistorico_Click()
Dim B As Boolean
    
    'Cargar el data3 con los datos de la tabla "sliman"
    NomTablaLineas = "sliman"
    Me.SSTab1.Tab = 3
    'ASignamos un SQL al DATA3
'    Data3.ConnectionString = Conn
    Data3.RecordSource = "Select anomante,mes01man,mes02man,mes03man,mes04man,mes05man,mes06man,mes07man,mes08man,mes09man,mes10man,mes11man,mes12man from " & NomTablaLineas & ObtenerWhereCP(True)
    Data3.CursorType = adOpenStatic
    Data3.Refresh
    If Data3.Recordset.EOF Then
        MsgBox "No existen datos en el Hist�rico para ese cliente y Direc./Dpto.", vbInformation
        Exit Sub
    Else
        B = Data3.Recordset.RecordCount > 2
        Me.imgFlecha(0).visible = B
        Me.imgFlecha(1).visible = B
    End If
    PonerCamposHistorico
End Sub


Private Sub mnModificar_Click()
    If Modo = 5 Then 'Modificar lineas
         BotonModificarLinea
    Else   'Modificar Mantenimiento
         If SSTab1.Tab <> 0 And Me.SSTab1.Tab <> 1 Then Me.SSTab1.Tab = 0
         If BLOQUEADesdeFormulario(Me) Then BotonModificar
    End If
End Sub


Private Sub mnNuevo_Click()
    If Modo = 5 Then 'A�adir lineas
         BotonAnyadirLinea
    Else 'A�adir Mantenimiento
         Me.SSTab1.Tab = 0
         BotonAnyadir
    End If
End Sub

Private Sub mnRevisiones_Click()
'Abre Mantenimiento de lineas  de Revisiones
    BotonMtoLineas 2, "Revisiones"
    NomTablaLineas = "slima1"
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


Private Sub SSTab1_Click(PreviousTab As Integer)
Dim B As Boolean
    
    B = (Me.SSTab1.Tab = 2)
    'Poner Visible el Nombre del T�cnico si estamos en Mantenimiento Lineas
    Me.Text2(21).visible = B
    Me.Label1(35).visible = B
End Sub


'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Not Text1(Index).MultiLine Then KEYdown KeyCode
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not Text1(Index).MultiLine Then KEYpress KeyAscii
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
Dim cadMen As String
Dim cadDpto As String
        
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si queremos hacer algo ..
    Select Case Index
        Case 0 'Cod. Clien
            devuelve = PonerNombreDeCod(Text1(Index), conAri, "sclien", "nomclien")
            Text2(Index).Text = devuelve

        Case 1 'Cod. Direc
            If Text1(Index).Text = "" Then
                Text2(Index).Text = ""
            Else
                'Comprobar que el cliente seleccionada tiene esa direccion
                If Text1(0).Text <> "" Then 'HAY cliente
                    devuelve = DevuelveDesdeBDNew(conAri, "sdirec", "coddirec", "codclien", Text1(0).Text, "N", , "coddirec", Text1(1).Text, "N")
                    If devuelve = "" Then 'NO EXISTE el dpto para el cliente
                        Text2(Index).Text = ""
                        If vParamAplic.Departamento Then
                            cadDpto = " el Departamento "
                        Else
                            cadDpto = "la Direcci�n "
                        End If
                        cadMen = "No existe" & cadDpto & Text1(1).Text & " para el cliente: " & vbCrLf
                        cadMen = cadMen & vbCrLf & Text1(0).Text & " - " & Text2(0).Text
                        MsgBox cadMen, vbInformation
                        
                        Text1(Index).Text = ""
                        PonerFoco Text1(Index)
                    Else 'EXISTE el dpto
                        '---- Laura: modificado 27/09/2006
                        'devuelve = PonerNombreDeCod(Text1(Index), conAri, "sdirec", "nomdirec")
                        devuelve = DevuelveDesdeBDNew(conAri, "sdirec", "nomdirec", "codclien", Text1(0).Text, "N", , "coddirec", Text1(1).Text, "N")
                        Text2(Index).Text = devuelve
                        FormateaCampo Text1(Index)
                    End If
                Else
                    Text1(Index).Text = ""
                    If vParamAplic.Departamento Then
                        cadDpto = "Departamentos. "
                    Else
                        cadDpto = "Direcciones. "
                    End If
                    MsgBox "Debe introducir un Cliente, antes de seleccionar sus " & cadDpto, vbInformation
                End If
            End If
            
        Case 2 'N� Mantenimiento
            'Comprobar si ya existe un registro con esa clave Primaria si Insertando
            If Modo = 3 And Text1(0).Text <> "" And Text1(2).Text <> "" Then
                devuelve = "select count(*) from scaman " & ObtenerWhereCP(True)
                If RegistrosAListar(devuelve) > 0 Then
                    MsgBox "Ya existe el registro en la Tabla de Mantenimientos.", vbExclamation
                    PonerFoco Text1(2)
                End If
            End If

        Case 3 'Fecha Inicio
            'Si no es modo de Busqueda poner el formato
            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
        Case 4 'Tipo Contrato
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "stipco", "nomtipco")
        Case 5 'Forma de Pago
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sforpa", "nomforpa")
        Case 6, 7 'Anticipado
            'Formato 3: Decimal(10,2)
            If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 3
        Case 8 To 31 'Meses
            If Text1(Index).Text <> "" Then
                'Formato 3: Decimal(10,2)
                PonerFormatoDecimal Text1(Index), 3
                If EsNumerico(Text1(Index).Text) Then
                    SumarTotal
                Else
                    PonerFoco Text1(Index)
                End If
            End If
        Case 33
            PonerFocoBtn Me.cmdAceptar
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
Dim cad As String
Dim Tabla As String
Dim Titulo As String
Dim Desc As String, devuelve As String
    'Llamamos a al form
    '##A mano
    cad = ""
    If EsCabecera Then
        cad = cad & ParaGrid(Text1(0), 10, "Cliente")
        cad = cad & "Desc. Cliente|sclien|nomclien|T||36�"
        If vParamAplic.Departamento Then
            Desc = "Dpto."
        Else
            Desc = "Direc."
        End If
        cad = cad & ParaGrid(Text1(1), 7, Desc)
        cad = cad & "Desc. " & Desc & "|sdirec|nomdirec|T||35�"
        cad = cad & ParaGrid(Text1(2), 13, "N� Mant.")
        
        Tabla = "(" & NombreTabla & " LEFT JOIN sclien ON " & NombreTabla & ".codclien=sclien.codclien" & ")"
        Tabla = Tabla & " LEFT JOIN sdirec ON " & NombreTabla & ".codclien=sdirec.codclien and "
        Tabla = Tabla & NombreTabla & ".coddirec=sdirec.coddirec"
        Titulo = "Mantenimientos"
        devuelve = "0|2|4|"
    Else
        If vParamAplic.Departamento Then
            Titulo = "Departamentos Cliente: "
            Desc = "Dpto."
        Else
            Titulo = "Direcciones Cliente: "
            Desc = "Direc."
        End If
        Titulo = Titulo & Text1(0).Text & " - " & Text2(0).Text
        cad = cad & "Cod. " & Desc & "|sdirec|coddirec|N|000|15�"
        cad = cad & "Desc. " & Desc & "|sdirec|nomdirec|T||60�"
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
'        frmB.vDevuelve = "0|1|"
        frmB.vDevuelve = devuelve
        frmB.vTitulo = Titulo
        frmB.vselElem = 1
        frmB.vConexionGrid = conAri 'Conexi�n a BD: Ariges
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
        PonerFoco Text1(0)
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
   
    'Revisiones - Datos de la tabla slima1
    CargaGrid DataGrid1, Data2, True
    
    PrimeraVez = False
    Screen.MousePointer = vbDefault
    Exit Sub
    
EPonerLineas:
    MuestraError Err.Number, "PonerCamposLineas"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
Dim cCli As CCliente
Dim nomDirDpto As String
    On Error Resume Next

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    
    'Colorines de los meses facturados
    Colorines
    
    'Desc. Tipo Contrato
    Text2(4).Text = PonerNombreDeCod(Text1(4), conAri, "stipco", "nomtipco")
    'Desc. Forma PAgo
    Text2(5).Text = PonerNombreDeCod(Text1(5), conAri, "sforpa", "nomforpa")
    'Cod. Cliente
    Text2(0).Text = PonerNombreDeCod(Text1(0), conAri, "sclien", "nomclien")
    
    '---- Laura: modificado 28/09/2006
    'Cod Dir./Dpto
    'Text2(1).Text = DevuelveDesdeBDNew(Text1(1), conAri, "sdirec", "nomdirec")
    Text2(1).Text = ""
    If Text1(1).Text <> "" Then
        Set cCli = New CCliente
        cCli.Codigo = Text1(0).Text
        If cCli.DptoCliente(Text1(1).Text, nomDirDpto) Then Text2(1).Text = nomDirDpto
        Set cCli = Nothing
    End If
    '----

       
    PonerCamposLineas 'Pone los datos de las tablas de lineas asociadas al trabajador
    LimpiarCamposHistorico
    SumarTotal
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
Dim B As Boolean
    
    On Error GoTo EPonerModo

    'Actualiza Iconos Insertar,Modificar,Eliminar
    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    B = (Modo = 2)

    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = B
    Else
        cmdRegresar.visible = False
    End If
        
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, B, NumReg
        
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar adem�s limpia los campos Text1
    BloquearText1 Me, Modo
              
    '---------------------------------------------
    B = Modo <> 0 And Modo <> 2 And Modo <> 5
    cmdCancelar.visible = B
    cmdAceptar.visible = B
    Me.cmbMes.Enabled = B
    
    For I = 0 To Me.imgFecha.Count - 1
        Me.imgFecha(I).Enabled = B
    Next I
    
'    For i = 0 To Me.imgBuscar.Count - 1
'        Me.imgBuscar(i).Enabled = b
'    Next i
   For I = 0 To Me.imgBuscar.Count - 1
        BloquearImg Me.imgBuscar(I), Not B
    Next I
    
    
    If Modo = 4 Then 'Modificar. Bloquear clave Primaria
        Me.imgBuscar(0).Enabled = False
'        Me.imgBuscar(1).Enabled = False
    End If
    
    Me.chkVistaPrevia.visible = (Modo <> 5)
    Me.cboTipoPago.Enabled = (Modo = 3) Or (Modo = 4) Or Modo = 1
    Me.chkBaterias.Enabled = (Modo = 3) Or (Modo = 4)
       
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu seg�n modo
    PonerOpcionesMenu   'Activar opciones de menu seg�n nivel
                        'de permisos del usuario
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Function DatosOk() As Boolean
Dim B As Boolean
    
    On Error GoTo EDatosOK

    DatosOk = False
    B = True
    B = CompForm(Me, 1) 'Comprobar formato datos ok
    If Not B Then Exit Function
    
    DatosOk = B
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLinea() As Boolean
Dim B As Boolean

    On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    B = True

    If Me.SSTab1.Tab = 2 Then 'Fecha Revision
        If Trim(TxtAux1(0).Text) = "" Then
            MsgBox "El campo Fecha Revisi�n no puede ser nulo", vbExclamation
            B = False
        End If
        
        If Trim(TxtAux1(1).Text) = "" Then 'Tecnico
            MsgBox "El campo Cod. T�cnico no puede ser nulo", vbExclamation
            B = False
        End If
    End If
    
    DatosOkLinea = B
    
EDatosOkLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


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
        Case 10  'Lineas Revisiones
            mnRevisiones_Click
        Case 11 'L�neas Hist�rico
            mnHistorico_Click
        Case 12 'Componentes
            BotonComponentes
        Case 14    'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub


Private Sub PonerOpcionesMenu()
Dim J As Byte

    On Error Resume Next
    
    PonerOpcionesMenuGeneral Me
    
    J = Val(Me.mnRevisiones.HelpContextID)
    If J < vUsu.Nivel Then Me.mnRevisiones.Enabled = False
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub

    
Private Function InsertarLinea() As Boolean
'Inserta un registro en la tabla de Revisiones: slima1
Dim SQL As String

    On Error GoTo EInsertarLinea

    InsertarLinea = False
    SQL = ""
    If DatosOkLinea And Me.SSTab1.Tab = 2 Then 'Revisiones
        SQL = "INSERT INTO slima1 "
        SQL = SQL & "(codclien, nummante, fecharev, codtraba, observac) "
        SQL = SQL & "VALUES (" & Val(Text1(0).Text) & ", "
        SQL = SQL & DBSet(Text1(2).Text, "T") & ", " & DBSet(TxtAux1(0).Text, "F") & ", " & TxtAux1(1).Text & ", "
        SQL = SQL & QuitarCaracterEnter(DBSet(TxtAux1(2).Text, "T")) & ")"
     End If
    
    If SQL <> "" Then
        Conn.Execute SQL
        InsertarLinea = True
    End If
    Exit Function
    
EInsertarLinea:
    MuestraError Err.Number, "Insertar Lineas Mantenimiento" & vbCrLf & Err.Description
End Function


Private Function ModificarLinea() As Boolean
'Modifica un registro en la tabla de Revisiones: slima1
Dim SQL As String

    On Error GoTo EModificarLinea

    ModificarLinea = False
    SQL = ""
    If DatosOkLinea And Me.SSTab1.Tab = 2 Then 'Habilidades
        SQL = "UPDATE slima1 Set codtraba = " & TxtAux1(1).Text & ", observac='" & QuitarCaracterEnter(TxtAux1(2).Text) & "'"
        SQL = SQL & ObtenerWhereCP(True) & " AND fecharev='" & Format(Data2.Recordset!FechaRev, FormatoFecha) & "'"
    End If

    If SQL <> "" Then
        Conn.Execute SQL
        ModificarLinea = True
    End If
    Exit Function
    
EModificarLinea:
    MuestraError Err.Number, "Modificar Lineas Trabajador" & vbCrLf & Err.Description
End Function


Private Sub PonerBotonCabecera(B As Boolean)
'Pone el boton de Regresar a la Cabecera si pasamos a MAntenimiento de Lineas
'o Pone los botones de Aceptar y cancelar en Insert,update o delete lineas
On Error Resume Next
    Me.cmdAceptar.visible = Not B
    Me.cmdCancelar.visible = Not B
    Me.cmdRegresar.visible = B
    Me.cmdRegresar.Caption = "Cabecera"
    If B Then
        Me.lblIndicador.Caption = "L�neas " & TituloLinea
        PonerFocoBtn Me.cmdRegresar
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
Dim B As Boolean
Dim SQL As String

    On Error GoTo ECargaGrid

    B = vDataGrid.Enabled
    
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral vDataGrid, vData, SQL, PrimeraVez
    vDataGrid.RowHeight = 470
    CargaGrid2 vDataGrid, vData
   
        
    vDataGrid.ScrollBars = dbgAutomatic
    vDataGrid.Enabled = B
    
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim tots As String

    On Error GoTo ECargaGrid
  
    vData.Refresh
    
    'codclien,nummante
    tots = "N||||0|;N||||0|;"
    'fecharev,codtraba,observa
    tots = tots & "S|TxtAux1(0)|T|Fecha Rev.|1100|;S|TxtAux1(1)|T|T�cnico|900|;S|cmdAux|B||0|;S|TxtAux1(2)|T|Observaciones|8100|;"
    arregla tots, vDataGrid, Me
    
     vDataGrid.Columns(3).NumberFormat = "0000"
     vDataGrid.Columns(4).WrapText = True
     
     vDataGrid.RowHeight = 470

     vDataGrid.Enabled = (Modo = 0) Or (Modo = 2) Or (Modo = 5 And ModificaLineas = 0)
     Exit Sub
     
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub LLamaLineas(alto As Single)
Dim jj As Byte
Dim B As Boolean

    DeseleccionaGrid Me.DataGrid1
    
    B = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2) 'Insertar o Modificar Lineas
   
    For jj = 0 To Me.TxtAux1.Count - 1
        Me.TxtAux1(jj).Height = DataGrid1.RowHeight
        Me.TxtAux1(jj).Top = alto
        Me.TxtAux1(jj).visible = B
    Next jj

    Me.cmdAux.Height = DataGrid1.RowHeight
    Me.cmdAux.Top = alto
    Me.cmdAux.visible = B
End Sub


Private Sub txtAux1_GotFocus(Index As Integer)
    If Index <> 2 Then ConseguirFoco TxtAux1(Index), Modo
End Sub

Private Sub txtAux1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub txtAux1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        If Index <> 2 Then
            PonerFoco TxtAux1(Index + 1)
        Else
            PonerFocoBtn Me.cmdAceptar
        End If
    End If
End Sub


Private Sub TxtAux1_LostFocus(Index As Integer)
Dim devuelve As String

    TxtAux1(Index).Text = Trim(TxtAux1(Index))
    If TxtAux1(Index).Text = "" Then Exit Sub
    Select Case Index
        Case 0 'Fecha Revision
            PonerFormatoFecha TxtAux1(Index)
            
        Case 1 'Cod. Trabajador
            If PonerFormatoEntero(Me.TxtAux1(Index)) Then
                devuelve = DevuelveDesdeBDNew(1, "straba", "nomtraba", "codtraba", TxtAux1(1).Text, "N")
                Text2(21).Text = devuelve
                FormateaCampo Me.TxtAux1(1)
                If devuelve = "" Then
                    MsgBox "No existe el T�cnico: " & TxtAux1(Index).Text, vbExclamation
                    PonerFoco TxtAux1(Index)
                End If
            End If
    End Select
End Sub


Private Sub BotonMtoLineas(numTab As Integer, cad As String)
        Me.SSTab1.Tab = numTab
        NumTabMto = numTab
        TituloLinea = cad
        PonerModo 5
        PonerBotonCabecera True
End Sub


Private Function Eliminar() As Boolean
Dim SQL As String

    On Error GoTo FinEliminar

    'Conn.BeginTrans   ESTA EN EL SUB SUPERIOR
    SQL = " WHERE  codclien=" & Data1.Recordset!CodClien
'        SQL = SQL & " AND coddirec=" & Data1.Recordset!CodDirec
    SQL = SQL & " AND nummante='" & DevNombreSQL(Data1.Recordset!numMante) & "'"

    'Lineas Mantenimiento (Hist�rico)
    Conn.Execute "Delete from sliman " & SQL
    'Lineas Revisiones
    Conn.Execute "Delete from slima1 " & SQL
    
    'Cabecera
    Conn.Execute "Delete from scaman " & SQL

FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
        'Conn.RollbackTrans
        Eliminar = False
    Else
       ' Conn.CommitTrans
        Eliminar = True
    End If
End Function


Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ning�n registro
On Error Resume Next
    CargaGrid DataGrid1, Data2, False
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargarComboTipoPago()
'### Combo Tipo Pago
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Mensual, 1-Trimestral, 2-Semestral, 3-Anual

    cboTipoPago.Clear
    cboTipoPago.AddItem "Mensual"
    cboTipoPago.ItemData(cboTipoPago.NewIndex) = 0
    
    cboTipoPago.AddItem "Trimestral"
    cboTipoPago.ItemData(cboTipoPago.NewIndex) = 1
    
    cboTipoPago.AddItem "Semestral"
    cboTipoPago.ItemData(cboTipoPago.NewIndex) = 2
    
    cboTipoPago.AddItem "Anual"
    cboTipoPago.ItemData(cboTipoPago.NewIndex) = 3
    
End Sub


Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    vWhere = "(" & ObtenerWhereCP(False) & ")"
    If SituarDataMULTI(Data1, vWhere, Indicador) Then
        PonerModo 2
        lblIndicador.Caption = Indicador
   Else
'        LimpiarCampos
        'Poner los grid sin apuntar a nada
'        LimpiarDataGrids
        PonerModo 0
    End If
End Sub


Private Function ObtenerWhereCP(conW As Boolean) As String
Dim SQL As String

    SQL = ""
    If conW Then SQL = " WHERE "
    SQL = SQL & " codclien=" & Text1(0).Text
    SQL = SQL & " AND nummante=" & DBSet(Text1(2).Text, "T")
    ObtenerWhereCP = SQL
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
    
    SQL = "SELECT * FROM slima1"
    If enlaza Then
        SQL = SQL & " " & ObtenerWhereCP(True)
    Else
        SQL = SQL & " WHERE codclien = -1"
    End If
    SQL = SQL & Ordenacion
    MontaSQLCarga = SQL
End Function


Private Sub SumarTotal()
Dim total As Currency
Dim I As Byte

    'Calcular el total actual
    total = 0
    For I = 8 To 30 Step 2
        If Text1(I).Text <> "" Then total = total + CCur(Text1(I).Text)
    Next I
    If total <> 0 Then
        Text2(2).Text = CStr(total)
        PonerFormatoDecimal Text2(2), 3
    Else
        Text2(2).Text = ""
    End If
    
    'Calcular el total Siguiente
    total = 0
    For I = 9 To 31 Step 2
        If Text1(I).Text <> "" Then total = total + CCur(Text1(I).Text)
    Next I
    If total <> 0 Then
        Text2(3).Text = CStr(total)
        PonerFormatoDecimal Text2(3), 3
    Else
        Text2(3).Text = ""
    End If
End Sub


Private Sub SumarTotalHistorico()
Dim total As Currency
Dim I As Byte

    'Calcular el total actual
    total = 0
    For I = 22 To 33
        If Text2(I).Text <> "" Then total = total + CCur(Text2(I).Text)
    Next I
    If total <> 0 Then
        Text2(I).Text = CStr(total)
        PonerFormatoDecimal Text2(I), 3
    Else
        Text2(I).Text = ""
    End If
    
    'Calcular el total Siguiente
    total = 0
    For I = 35 To 46
        If Text2(I).Text <> "" Then total = total + CCur(Text2(I).Text)
    Next I
    If total <> 0 Then
        Text2(I).Text = CStr(total)
        PonerFormatoDecimal Text2(I), 3
    Else
        Text2(I).Text = ""
    End If
End Sub



Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar seg�n el modo en que estemos
Dim B As Boolean

    B = (Modo = 2) Or (Modo = 5)
    'Insertar
    Toolbar1.Buttons(5).Enabled = (B Or Modo = 0)
    Me.mnNuevo.Enabled = (B Or Modo = 0)
    'Modificar
    Toolbar1.Buttons(6).Enabled = B
    Me.mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(7).Enabled = B
    Me.mnEliminar.Enabled = B
    
    
    B = (Modo = 2)
    'Mantenimiento lineas Revisiones
    Toolbar1.Buttons(10).Enabled = B
    Me.mnRevisiones.Enabled = B
    'Lineas Hist�rico
    Toolbar1.Buttons(11).Enabled = B
    Me.mnHistorico.Enabled = B
    Me.mnOpciones.Enabled = B Or (Modo = 0)
    Me.mnMtoLineas.Enabled = B Or (Modo = 0)
    'Componentes
    Me.Toolbar1.Buttons(12).Enabled = B
    
    
    B = (Modo >= 3) Or Modo = 1
    'Buscar
    Toolbar1.Buttons(1).Enabled = Not B
    Me.mnBuscar.Enabled = Not B
    'Ver Todos
    Toolbar1.Buttons(2).Enabled = Not B
    Me.mnVerTodos.Enabled = Not B
End Sub


Private Sub PonerCamposHistorico()
Dim I As Integer
    
    On Error Resume Next
    
    If Data3.Recordset.EOF Then
        Data3.Recordset.MoveLast
        Exit Sub
    End If
    'Pone 2 a�os (2 registros) cada vez
    'Primer A�o
    '----------------------------------------------------------------------------
    Me.Label1(38).Caption = Data3.Recordset.Fields(0).Value
    Me.Label1(47).Caption = Me.Label1(38).Caption
    
    For I = 1 To 12
        
        'Text2(22).Text = Format(Data3.Recordset.Fields(4).Value, FormatoCantidad)
        Text2(21 + I).Text = Format(Data3.Recordset.Fields(I).Value, FormatoCantidad)
    Next I
    
       
    'Segundo A�o
    '----------------------------------------------------------------------------
    Data3.Recordset.MoveNext
    If Not Data3.Recordset.EOF Then
        'Poner el a�o siguiente
        Me.Label1(37).Caption = Data3.Recordset.Fields(0).Value
        Me.Label1(46).Caption = Me.Label1(37).Caption
        
        
        For I = 1 To 12
            
            'Text2(22).Text = Format(Data3.Recordset.Fields(4).Value, FormatoCantidad)
            Text2(34 + I).Text = Format(Data3.Recordset.Fields(I).Value, FormatoCantidad)
        Next I
        
        
        
    Else
        Data3.Recordset.MoveLast
        Me.Label1(37).Caption = ""
        Me.Label1(46).Caption = ""
        For I = 4 To 15
            Text2(I + 31).Text = ""
        Next I
    End If
    SumarTotalHistorico
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub LimpiarCamposHistorico()
Dim I As Byte

    Me.imgFlecha(0).visible = False
    Me.imgFlecha(1).visible = False
    Me.Label1(38).Caption = "A�o"
    Me.Label1(37).Caption = "A�o"
    Me.Label1(46).Caption = "A�o"
    Me.Label1(47).Caption = "A�o"
    For I = 4 To 15
        Text2(I + 18).Text = ""
        Text2(I + 31).Text = ""
    Next I
    'Limpiar el total del Hist�rico
    Text2(34).Text = ""
    Text2(47).Text = ""
End Sub


Private Sub BotonComponentes()
'Muestra un form de Mensaje para seleccionar el tipo de resumen que queremos mostrar:
'Por Mantenimiento, Por Departamento, Por Cliente
Dim vWhere As String

    'Codigo Cliente
    If Text1(0).Text = "" Then Exit Sub
    vWhere = " WHERE codclien = " & Text1(0).Text
    
    frmMensajes.cadWhere = vWhere
    
    'cadClien= codclien - nomclien
    vWhere = Text1(0).Text & " - " & Text2(0).Text
    'vCampos= Mantenimiento|coddirec|Desc. coddirec| cadCliente
    frmMensajes.vCampos = Text1(2).Text & "|" & Text1(1).Text & "|" & Text2(1).Text & "|" & vWhere & "|"
    frmMensajes.OpcionMensaje = 5 'Componentes
    frmMensajes.Show vbModal
End Sub



Private Sub CargaComboMes()
Dim F As Date
    Me.cmbMes.Clear
    'Sin especificar
        cmbMes.AddItem "    "
        cmbMes.ItemData(0) = 0
    For kCampo = 1 To 12
        NombreTabla = Format(CDate("01/" & kCampo & "/2000"), "mmmm")
        NombreTabla = UCase(Mid(NombreTabla, 1, 1)) & Mid(NombreTabla, 2)
        cmbMes.AddItem NombreTabla
        cmbMes.ItemData(kCampo) = kCampo
    Next kCampo
    
End Sub


Private Sub Colorines()
Dim I As Integer
Dim J As Integer
    
    I = cmbMes.ListIndex
    For J = 1 To I
        Text1(((J - 1) * 2) + 8).ForeColor = vbRed
    Next J
    For J = I + 1 To 12
        Text1(((J - 1) * 2) + 8).ForeColor = vbBlack
    Next J
End Sub



Private Function Traspasa_A_Anulados() As String
Dim I As Integer
Dim C As String
Dim F As Date

    On Error GoTo ETraspasa_A_Anulados
    
    Traspasa_A_Anulados = "MAL"
    
    'Se trata de inserta en cabceras y lineas de mantenimienos anulados
    'Preparo datos desde otroform para insercion
    F = CDate(RecuperaValor(CadenaDesdeOtroForm, 2))
    C = "'" & Format(F, FormatoFecha) & "','"
    C = C & DevNombreSQL(RecuperaValor(CadenaDesdeOtroForm, 3)) & "'," & Val(RecuperaValor(CadenaDesdeOtroForm, 4)) & ","
    CadenaDesdeOtroForm = C
    'Los campos son los mismos MAS :
    '           FECHABAJA,usuario(texto),codincid
    C = ""
    For I = 0 To Data1.Recordset.Fields.Count - 1
        C = C & "," & Data1.Recordset.Fields(I).Name
    Next I
    C = Mid(C, 2)
    C = "INSERT INTO scamana(fechabaj,usuario,codincid," & C & ") SELECT " & CadenaDesdeOtroForm & C & " FROM scaman WHERE codclien="
    C = C & Data1.Recordset!CodClien & " AND nummante = '" & DevNombreSQL(Data1.Recordset!numMante) & "'"
    Conn.Execute C
    
    'Las lineas de hco de importes
    If Not (Data3.Recordset Is Nothing) Then
        C = ""
        For I = 0 To Data3.Recordset.Fields.Count - 1
            C = C & "," & Data3.Recordset.Fields(I).Name
        Next I
        CadenaDesdeOtroForm = "'" & Format(F, FormatoFecha) & "'"
        C = "INSERT INTO slimana(fechabaj" & C & ") SELECT " & CadenaDesdeOtroForm & C & " FROM sliman WHERE codclien="
        C = C & Data1.Recordset!CodClien & " AND nummante = '" & DevNombreSQL(Data1.Recordset!numMante) & "'"
        Conn.Execute C
    
    
    End If
    
    Traspasa_A_Anulados = ""
    Exit Function
ETraspasa_A_Anulados:
    MuestraError Err.Number, "Traspasa_A_Anulados"
End Function
