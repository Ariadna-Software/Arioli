VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmComProveedores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proveedores"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10860
   Icon            =   "frmComProveedores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   10860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Digitos 1er nivel|N|N|||empresa|numdigi1|||"
   Begin VB.Frame Frame4 
      Height          =   670
      Left            =   240
      TabIndex        =   71
      Top             =   480
      Width           =   10455
      Begin VB.CheckBox chkProveV 
         Caption         =   "Proveedor de Varios"
         Height          =   195
         Left            =   8400
         TabIndex        =   76
         Tag             =   "Proveedor Varios|N|N|||sprove|provario||N|"
         Top             =   220
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   3345
         MaxLength       =   40
         TabIndex        =   74
         Tag             =   "Nombre Proveedor|T|N|||sprove|nomprove||N|"
         Text            =   "Text1"
         Top             =   220
         Width           =   4245
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   72
         Tag             =   "C�digo Proveedor|N|N|0|999999|sprove|codprove|000000|S|"
         Text            =   "Text1"
         Top             =   220
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   75
         Top             =   220
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. Proveedor"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   73
         Top             =   220
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   8400
      Top             =   5520
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin VB.Frame Frame3 
      Height          =   540
      Left            =   240
      TabIndex        =   65
      Top             =   5475
      Width           =   3000
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   66
         Top             =   240
         Width           =   2715
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4220
      Left            =   240
      TabIndex        =   36
      Top             =   1200
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   7435
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Datos b�sicos"
      TabPicture(0)   =   "frmComProveedores.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(6)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(4)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(5)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(7)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(8)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(9)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(11)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(10)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "imgCuentas(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(20)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(12)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "imgCuentas(2)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(14)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(13)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(19)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "imgCuentas(0)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "imgFecha(0)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "imgFecha(1)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "imgCuentas(3)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label1(21)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label1(62)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "imgCuentas(4)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Text1(6)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Text1(4)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Text1(3)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Text1(2)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Text1(7)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Text1(8)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Text1(9)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Text1(10)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "cboTipoDto"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Text1(14)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Text1(15)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Text1(16)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Text1(17)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Text1(18)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Text1(13)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Text1(12)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Text2(1)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Text2(2)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Text1(11)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Text1(5)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "cboTipoProv"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Text2(0)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "Text1(29)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "Text2(29)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "checkAlbFac(0)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "Text1(31)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "checkAlbFac(1)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).ControlCount=   51
      TabCaption(1)   =   "Datos Contacto"
      TabPicture(1)   =   "frmComProveedores.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2(10)"
      Tab(1).Control(1)=   "Label2(11)"
      Tab(1).Control(2)=   "imgWeb"
      Tab(1).Control(3)=   "Text1(27)"
      Tab(1).Control(4)=   "Text1(28)"
      Tab(1).Control(5)=   "Frame1"
      Tab(1).Control(6)=   "Frame2(13)"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Documentos"
      TabPicture(2)   =   "frmComProveedores.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label3"
      Tab(2).Control(1)=   "imgFecha(2)"
      Tab(2).Control(2)=   "Label2(0)"
      Tab(2).Control(3)=   "lw1"
      Tab(2).Control(4)=   "Toolbar2"
      Tab(2).Control(5)=   "Text1(30)"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Homologaci�n"
      TabPicture(3)   =   "frmComProveedores.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "imgFecha(3)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label1(15)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label2(1)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label2(12)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "ImgHomo(0)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "ImgHomo(1)"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "ImgHomo(2)"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "lwH"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Text1(32)"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "Text1(33)"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).ControlCount=   10
      Begin VB.TextBox Text1 
         Height          =   2985
         Index           =   33
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   87
         Tag             =   "Observaciones|T|S|||sprove|obshomo|||"
         Text            =   "frmComProveedores.frx":007C
         Top             =   1080
         Width           =   4695
      End
      Begin VB.TextBox Text1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Index           =   32
         Left            =   -72720
         MaxLength       =   10
         TabIndex        =   85
         Tag             =   "Fecha homo.|F|S|||sprove|fechahomo|dd/mm/yyyy||"
         Text            =   "Text1"
         Top             =   510
         Width           =   1335
      End
      Begin VB.CheckBox checkAlbFac 
         Alignment       =   1  'Right Justify
         Caption         =   "Homologado"
         Height          =   195
         Index           =   1
         Left            =   6840
         TabIndex        =   84
         Tag             =   "s|N|N|||sprove|homologado||N|"
         Top             =   2800
         Width           =   1870
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   31
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   12
         Tag             =   "Iban|T|S|||sprove|iban|||"
         Text            =   "Text1"
         Top             =   2505
         Width           =   615
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   30
         Left            =   -66000
         TabIndex        =   80
         Text            =   "Text4"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CheckBox checkAlbFac 
         Alignment       =   1  'Right Justify
         Caption         =   "Albaran x Factura"
         Height          =   195
         Index           =   0
         Left            =   6840
         TabIndex        =   18
         Tag             =   "s|N|N|||sprove|albaranxfactura||N|"
         Top             =   3120
         Width           =   1870
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   29
         Left            =   2400
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   77
         Text            =   "Text2"
         Top             =   3720
         Width           =   3165
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   29
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   20
         Tag             =   "Cod. Situaci�n|N|N|0|99|sprove|codsitua|0|N|"
         Text            =   "Te"
         Top             =   3720
         Width           =   645
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   0
         Left            =   3030
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   38
         Text            =   "Text2"
         Top             =   2880
         Width           =   3495
      End
      Begin VB.ComboBox cboTipoProv 
         Height          =   315
         ItemData        =   "frmComProveedores.frx":0083
         Left            =   8520
         List            =   "frmComProveedores.frx":0085
         TabIndex        =   6
         Tag             =   "Tipo de Proveedor|N|N|||sprove|tipprove||N|"
         Text            =   "Combo1"
         Top             =   495
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   5
         Left            =   3360
         MaxLength       =   30
         TabIndex        =   3
         Tag             =   "Poblaci�n|T|N|||sprove|pobprove||N|"
         Text            =   "Text1"
         Top             =   1245
         Width           =   2550
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   8520
         MaxLength       =   5
         TabIndex        =   11
         Tag             =   "Dto. General|N|S|0|99.90|sprove|dtognral|#0.00||"
         Text            =   "Text1"
         Top             =   2355
         Width           =   735
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   2
         Left            =   6600
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   42
         Text            =   "Text2"
         Top             =   3720
         Width           =   3735
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   1
         Left            =   2400
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   40
         Text            =   "Text2"
         Top             =   3300
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   12
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   17
         Tag             =   "Cuenta Contable|T|N|||sprove|codmacta|||"
         Text            =   "Text1"
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   13
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   19
         Tag             =   "Forma Pago|N|N|0|999|sprove|codforpa|000|N|"
         Text            =   "Text1"
         Top             =   3300
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   18
         Left            =   4200
         MaxLength       =   10
         TabIndex        =   16
         Tag             =   "Cuenta Bancaria|T|S|||sprove|cuentaba|0000000000||"
         Text            =   "Text1"
         Top             =   2505
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   17
         Left            =   3735
         MaxLength       =   2
         TabIndex        =   15
         Tag             =   "Digito Control|T|S|||sprove|digcontr|00||"
         Text            =   "Text1"
         Top             =   2505
         Width           =   360
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   16
         Left            =   3060
         MaxLength       =   4
         TabIndex        =   14
         Tag             =   "Sucursal|N|S|0|9999|sprove|codsucur|0000||"
         Text            =   "Text1"
         Top             =   2505
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   15
         Left            =   2400
         MaxLength       =   4
         TabIndex        =   13
         Tag             =   "Banco|N|S|0|9999|sprove|codbanco|0000||"
         Text            =   "Text1"
         Top             =   2505
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   14
         Left            =   5880
         MaxLength       =   4
         TabIndex        =   21
         Tag             =   "Banco Propio|N|N|0|9999|sprove|codbanpr|0000||"
         Text            =   "Text1"
         Top             =   3720
         Width           =   615
      End
      Begin VB.ComboBox cboTipoDto 
         Height          =   315
         Left            =   8520
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Tag             =   "Tipo Descuento|N|N|||sprove|tipodtos||N|"
         Top             =   1605
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   8520
         MaxLength       =   5
         TabIndex        =   10
         Tag             =   "Dto. Pronto Pago|N|S|0|99.90|sprove|dtoppago|#0.00||"
         Text            =   "Text1"
         Top             =   1980
         Width           =   735
      End
      Begin VB.TextBox Text1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   8520
         MaxLength       =   10
         TabIndex        =   8
         Tag             =   "Fecha �ltima compra|F|S|||sprove|fechamov|dd/mm/yyyy||"
         Text            =   "Text1"
         Top             =   1245
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   8520
         MaxLength       =   10
         TabIndex        =   7
         Tag             =   "Fecha de Alta|F|N|||sprove|fecprove|dd/mm/yyyy||"
         Text            =   "Text1"
         Top             =   870
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         Caption         =   "Compras"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   2080
         Index           =   13
         Left            =   -69720
         TabIndex        =   53
         Top             =   450
         Width           =   4935
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   23
            Left            =   120
            MaxLength       =   40
            TabIndex        =   29
            Tag             =   "Persona de Contacto Compras|T|S|||sprove|perprov2|||"
            Text            =   "Text1"
            Top             =   600
            Width           =   4440
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   24
            Left            =   120
            MaxLength       =   40
            TabIndex        =   30
            Tag             =   "eMail Compras|T|S|||sprove|maiprov2|||"
            Text            =   "Text1"
            Top             =   1200
            Width           =   4440
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   25
            Left            =   840
            MaxLength       =   15
            TabIndex        =   31
            Tag             =   "Tel�fono Compras|T|S|||sprove|telprov2|||"
            Text            =   "Text1"
            Top             =   1640
            Width           =   1560
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   26
            Left            =   3000
            MaxLength       =   15
            TabIndex        =   32
            Tag             =   "Fax Compras|T|S|||sprove|faxprov2|||"
            Text            =   "Text1"
            Top             =   1640
            Width           =   1560
         End
         Begin VB.Image ImgMail 
            Height          =   240
            Index           =   1
            Left            =   600
            Tag             =   "-1"
            ToolTipText     =   "Enviar e-mail"
            Top             =   945
            Width           =   240
         End
         Begin VB.Label Label2 
            Caption         =   "Persona de Contacto"
            Height          =   240
            Index           =   6
            Left            =   120
            TabIndex        =   57
            Top             =   360
            Width           =   3495
         End
         Begin VB.Label Label2 
            Caption         =   "E-mail"
            Height          =   240
            Index           =   7
            Left            =   120
            TabIndex        =   56
            Top             =   960
            Width           =   3495
         End
         Begin VB.Label Label2 
            Caption         =   "Tel�fono"
            Height          =   240
            Index           =   8
            Left            =   120
            TabIndex        =   55
            Top             =   1640
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Fax"
            Height          =   240
            Index           =   9
            Left            =   2640
            TabIndex        =   54
            Top             =   1640
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Administraci�n"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   2080
         Left            =   -74760
         TabIndex        =   48
         Top             =   450
         Width           =   4935
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   22
            Left            =   3000
            MaxLength       =   15
            TabIndex        =   28
            Tag             =   "Fax Administraci�n|T|S|||sprove|faxprov1|||"
            Text            =   "Text1"
            Top             =   1640
            Width           =   1560
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   21
            Left            =   840
            MaxLength       =   15
            TabIndex        =   27
            Tag             =   "Telefono Administraci�n|T|S|||sprove|telprov1|||"
            Text            =   "Text1"
            Top             =   1640
            Width           =   1560
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   20
            Left            =   120
            MaxLength       =   40
            TabIndex        =   26
            Tag             =   "eMail Administraci�n|T|S|||sprove|maiprov1|||"
            Text            =   "Text1"
            Top             =   1200
            Width           =   4440
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   19
            Left            =   120
            MaxLength       =   40
            TabIndex        =   25
            Tag             =   "Persona de Contacto Administraci�n|T|S|||sprove|perprov1|||"
            Text            =   "Text1"
            Top             =   600
            Width           =   4440
         End
         Begin VB.Image ImgMail 
            Height          =   240
            Index           =   0
            Left            =   600
            Tag             =   "-1"
            ToolTipText     =   "Enviar e-mail"
            Top             =   945
            Width           =   240
         End
         Begin VB.Label Label2 
            Caption         =   "Fax"
            Height          =   240
            Index           =   5
            Left            =   2640
            TabIndex        =   52
            Top             =   1640
            Width           =   255
         End
         Begin VB.Label Label2 
            Caption         =   "Tel�fono"
            Height          =   240
            Index           =   4
            Left            =   120
            TabIndex        =   51
            Top             =   1640
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "E-mail"
            Height          =   240
            Index           =   3
            Left            =   120
            TabIndex        =   50
            Top             =   960
            Width           =   3495
         End
         Begin VB.Label Label2 
            Caption         =   "Persona de Contacto"
            Height          =   240
            Index           =   2
            Left            =   120
            TabIndex        =   49
            Top             =   360
            Width           =   3495
         End
      End
      Begin VB.TextBox Text1 
         Height          =   580
         Index           =   28
         Left            =   -74760
         MaxLength       =   200
         TabIndex        =   33
         Tag             =   "Observaciones|T|S|||sprove|observac|||"
         Text            =   "Text2 "
         Top             =   2850
         Width           =   9975
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   27
         Left            =   -74760
         MaxLength       =   40
         TabIndex        =   34
         Tag             =   "Web|T|S|||sprove|wwwprove|||"
         Text            =   "Text1"
         Top             =   3735
         Width           =   6000
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   7
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   5
         Tag             =   "N.I.F.|T|N|||sprove|nifprove|||"
         Text            =   "Text1"
         Top             =   1995
         Width           =   2070
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   2
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   0
         Tag             =   "Nombre Comercial|T|N|||sprove|nomcomer||N|"
         Text            =   "Text1"
         Top             =   510
         Width           =   4245
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   3
         Left            =   1680
         MaxLength       =   35
         TabIndex        =   1
         Tag             =   "Domicilio|T|S|||sprove|domprove||N|"
         Text            =   "Text1"
         Top             =   885
         Width           =   4230
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   4
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   2
         Tag             =   "CPostal|T|N|||sprove|codpobla||N|"
         Text            =   "Text1"
         Top             =   1245
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   6
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   4
         Tag             =   "Provincia|T|N|||sprove|proprove|||"
         Text            =   "Text1"
         Top             =   1620
         Width           =   3270
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   1710
         Left            =   -74760
         TabIndex        =   79
         Top             =   480
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   3016
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Appearance      =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Ofertas"
               Object.Tag             =   "0"
               Style           =   2
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Pedidos"
               Object.Tag             =   "1"
               Style           =   2
               Value           =   1
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Albaranes"
               Object.Tag             =   "2"
               Style           =   2
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "facturas"
               Object.Tag             =   "3"
               Style           =   2
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Precios especiales"
               Object.Tag             =   "4"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lw1 
         Height          =   3495
         Left            =   -74040
         TabIndex        =   81
         Top             =   480
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   6165
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lwH 
         Height          =   3135
         Left            =   -69840
         TabIndex        =   89
         Top             =   960
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   5530
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Acci�n"
            Object.Width           =   6068
         EndProperty
      End
      Begin VB.Image ImgHomo 
         Height          =   255
         Index           =   2
         Left            =   -68280
         ToolTipText     =   "Eliminar"
         Top             =   720
         Width           =   255
      End
      Begin VB.Image ImgHomo 
         Height          =   255
         Index           =   1
         Left            =   -68640
         ToolTipText     =   "Modificar"
         Top             =   720
         Width           =   255
      End
      Begin VB.Image ImgHomo 
         Height          =   255
         Index           =   0
         Left            =   -69000
         ToolTipText     =   "Nueva"
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Acciones "
         Height          =   240
         Index           =   12
         Left            =   -69840
         TabIndex        =   90
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Observaciones"
         Height          =   240
         Index           =   1
         Left            =   -74880
         TabIndex        =   88
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha homologaci�n"
         Height          =   195
         Index           =   15
         Left            =   -74880
         TabIndex        =   86
         Top             =   480
         Width           =   1710
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   -73000
         Picture         =   "frmComProveedores.frx":0087
         ToolTipText     =   "Buscar fecha"
         Top             =   480
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Index           =   0
         Left            =   -67080
         TabIndex        =   83
         Top             =   480
         Width           =   2385
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   -66360
         Picture         =   "frmComProveedores.frx":0112
         ToolTipText     =   "Buscar fecha"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Left            =   -66960
         TabIndex        =   82
         Top             =   1200
         Width           =   735
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   4
         Left            =   1360
         ToolTipText     =   "Buscar situaci�n"
         Top             =   3750
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. Situaci�n"
         Height          =   255
         Index           =   62
         Left            =   240
         TabIndex        =   78
         Top             =   3800
         Width           =   1095
      End
      Begin VB.Image imgWeb 
         Height          =   255
         Left            =   -74280
         Picture         =   "frmComProveedores.frx":019D
         Stretch         =   -1  'True
         Tag             =   "-1"
         ToolTipText     =   "Abrir web"
         Top             =   3475
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Banco Proveedor"
         Height          =   195
         Index           =   21
         Left            =   255
         TabIndex        =   70
         Top             =   2505
         Width           =   1320
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   3
         Left            =   1380
         Tag             =   "-1"
         ToolTipText     =   "Buscar poblaci�n"
         Top             =   1275
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   8235
         Picture         =   "frmComProveedores.frx":0727
         ToolTipText     =   "Buscar fecha"
         Top             =   1245
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   8235
         Picture         =   "frmComProveedores.frx":07B2
         ToolTipText     =   "Buscar fecha"
         Top             =   840
         Width           =   240
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   0
         Left            =   1380
         Tag             =   "-1"
         ToolTipText     =   "Buscar cuenta contable"
         Top             =   2925
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Proveedor"
         Height          =   255
         Index           =   19
         Left            =   6855
         TabIndex        =   68
         Top             =   495
         Width           =   1110
      End
      Begin VB.Label Label1 
         Caption         =   "Dto. General"
         Height          =   195
         Index           =   13
         Left            =   6855
         TabIndex        =   67
         Top             =   2400
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "Banco Propio"
         Height          =   195
         Index           =   14
         Left            =   5880
         TabIndex        =   64
         Top             =   3480
         Width           =   1080
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   2
         Left            =   6960
         ToolTipText     =   "Buscar banco propio"
         Top             =   3480
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Dto. Pronto Pago"
         Height          =   195
         Index           =   12
         Left            =   6855
         TabIndex        =   63
         Top             =   2040
         Width           =   1320
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Descuento"
         Height          =   255
         Index           =   20
         Left            =   6855
         TabIndex        =   62
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   1
         Left            =   1380
         ToolTipText     =   "Buscar forma de pago"
         Top             =   3337
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Forma de Pago"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   61
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Cta Contable"
         Height          =   195
         Index           =   11
         Left            =   255
         TabIndex        =   60
         Top             =   2880
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Ult. Compra"
         Height          =   195
         Index           =   9
         Left            =   6855
         TabIndex        =   59
         Top             =   1245
         Width           =   1320
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha de Alta"
         Height          =   195
         Index           =   8
         Left            =   6855
         TabIndex        =   58
         Top             =   870
         Width           =   1080
      End
      Begin VB.Label Label2 
         Caption         =   "Observaciones"
         Height          =   240
         Index           =   11
         Left            =   -74760
         TabIndex        =   47
         Top             =   2630
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Web"
         Height          =   240
         Index           =   10
         Left            =   -74760
         TabIndex        =   46
         Top             =   3520
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "N.I.F."
         Height          =   255
         Index           =   7
         Left            =   255
         TabIndex        =   45
         Top             =   1995
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre Comercial"
         Height          =   255
         Index           =   2
         Left            =   255
         TabIndex        =   44
         Top             =   510
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Poblaci�n"
         Height          =   255
         Index           =   5
         Left            =   2535
         TabIndex        =   43
         Top             =   1245
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Domicilio"
         Height          =   255
         Index           =   3
         Left            =   255
         TabIndex        =   41
         Top             =   885
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. postal"
         Height          =   255
         Index           =   4
         Left            =   255
         TabIndex        =   39
         Top             =   1245
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Provincia"
         Height          =   240
         Index           =   6
         Left            =   255
         TabIndex        =   37
         Top             =   1620
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5640
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   8115
      TabIndex        =   22
      Top             =   5640
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   10860
      _ExtentX        =   19156
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Todos"
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
            Object.ToolTipText     =   "Borrar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "salir"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "�ltimo"
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   7080
         TabIndex        =   69
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   9240
      TabIndex        =   24
      Top             =   5640
      Visible         =   0   'False
      Width           =   1035
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
End
Attribute VB_Name = "frmComProveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)


Private HaDevueltoDatos As Boolean
Private NombreTabla As String  'Nombre de la tabla o de la
Private Kcampo As Integer
Private Ordenacion As String
Private CadenaConsulta As String

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmCP As frmCPostal 'Codigos Postales
Attribute frmCP.VB_VarHelpID = -1

Private WithEvents frmFP As frmFacFormasPago 'Form Formas de Pago en menu Facturacion
Attribute frmFP.VB_VarHelpID = -1
Private WithEvents frmBP As frmFacBancosPropios
Attribute frmBP.VB_VarHelpID = -1
Private WithEvents frmS As frmFacSituaciones
Attribute frmS.VB_VarHelpID = -1

Dim btnPrimero As Byte
'Variable que indica el n�mero del Boton  PrimerRegistro en la Toolbar1

Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos

Dim Modo As Byte

Private BuscaChekc As String



Private Sub cboTipoDto_KeyPress(KeyAscii As Integer)
    KEYpress (KeyAscii)
End Sub


Private Sub cboTipoProv_KeyPress(KeyAscii As Integer)
    KEYpress (KeyAscii)
End Sub


Private Sub checkAlbFac_Click(Index As Integer)
     If Modo = 1 Then CheckCadenaBusqueda checkAlbFac(Index), BuscaChekc
End Sub

Private Sub checkAlbFac_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkProveV_Click()
 If Modo = 1 Then CheckCadenaBusqueda chkProveV, BuscaChekc
End Sub

Private Sub chkProveV_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim I As Integer


    Screen.MousePointer = vbHourglass
    On Error GoTo Error1


    Select Case Modo
    Case 1 'BUSCAR
        HacerBusqueda
   
    Case 2, 4 'MODIFICAR
        If DatosOk Then
            If Data1.Recordset.EOF Then
                I = InsertarDesdeForm(Me)
            Else
                I = ModificaDesdeFormulario(Me, 1)
                TerminaBloquear
                PosicionarData
            End If
        End If

    Case 3 'INSERTAR
        If DatosOk Then
            If InsertarDesdeForm(Me) Then
                ComprobarCrearCuenta
                PosicionarData
            End If
        End If
    End Select
    


Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub ComprobarCrearCuenta()
        'Si pone en la cuenta contable, crear nueva cta contable
         If Text2(0).Text = vbCrearNuevaCta Then
            If Not InsertarCuentaCble(Text1(12).Text, "", Text1(0).Text) Then
                MsgBox "Se ha producido un error insertando la cuenta: " & Text1(0).Text & ". Compruebelo", vbExclamation
                Exit Sub
            Else
                Text1_LostFocus 12
            End If
        End If
End Sub

Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 1, 3 'Busqueda, Insertar
            LimpiarCampos
            PonerModo 0
        Case 2
            PonerCampos
            lblIndicador.Caption = ""
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    End Select
    
End Sub


Private Sub cmdRegresar_Click()
Dim Cad As String

    If Data1.Recordset.EOF Then
        MsgBox "Ning�n registro devuelto.", vbExclamation
        Exit Sub
    End If

    Cad = Data1.Recordset.Fields(0) & "|"
    Cad = Cad & Data1.Recordset.Fields(1) & "|"
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If Modo = 1 Then PonerFoco Text1(0)
End Sub

Private Sub Form_Load()
    'Icono del formularios
    Me.Icon = frmppal.Icon
   
   
    'Icono de busqueda
    For Kcampo = 0 To Me.imgCuentas.Count - 1
        Me.imgCuentas(Kcampo).Picture = frmppal.imgListComun.ListImages(19).Picture
    Next Kcampo
   
    'Icono de e-mail
    For Kcampo = 0 To Me.ImgMail.Count - 1
        Me.ImgMail(Kcampo).Picture = frmppal.imgListComun.ListImages(20).Picture
    Next Kcampo
   
    For Kcampo = 0 To Me.ImgHomo.Count - 1
        ImgHomo(Kcampo).Picture = frmppal.ImageListB.ListImages(Kcampo + 1).Picture
    Next Kcampo

   
   
    'Lista imagen
    btnPrimero = 12
    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmppal.imgListComun
        .Buttons(1).Image = 1  'Buscar
        .Buttons(2).Image = 2  'Ver Todos
        .Buttons(5).Image = 3  'A�adir
        .Buttons(6).Image = 4  'Modificar
        .Buttons(7).Image = 5  'Borrar
        .Buttons(10).Image = 15 'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Siguiente
        .Buttons(btnPrimero + 2).Image = 8 'Anterior
        .Buttons(btnPrimero + 3).Image = 9 'Ultimo
    End With
    
    'Documentos
    ImagenesNavegacion
    
    'Solo si puede tener REA, entonces mostraremos el check este
    checkAlbFac(0).visible = vParamAplic.IVA_REA > 0
    
    
    limpiar Me
    Me.SSTab1.Tab = 0
    VieneDeBuscar = False

    NombreTabla = "sprove"
    Ordenacion = " ORDER BY codprove"
    Data1.ConnectionString = conn
    Data1.RecordSource = "Select * from " & NombreTabla & " where codprove=-1"
    Data1.Refresh
    
    Toolbar1.Buttons(6).Enabled = Not Data1.Recordset.EOF
    Toolbar1.Buttons(7).Enabled = Not Data1.Recordset.EOF
     
    CargarComboTipoProveedor
    CargarComboTipoDto
      
      
    'Ponemos los datos del listview
    imgFecha(2).Tag = vEmpresa.FechaIni
    CargaColumnas 1
      
      
     '=======Modif.
     If DatosADevolverBusqueda = "" Then
        PonerModo 0
     Else
        PonerModo 1
     End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim cadB As String
Dim Aux As String
Dim Indice As Byte
      
    If CadenaDevuelta <> "" Then
        If Val(imgCuentas(0).Tag) >= 0 Then
            'Se llama desde un bot�n de busqueda de los campos
            'Cuenta Contable, Forma Pago, Banco Propio
            'Recuperar solo el campo c�digo y Descripci�n
            HaDevueltoDatos = True
            Screen.MousePointer = vbHourglass
    
            Indice = Val(Me.imgCuentas(0).Tag)
            Text1(Indice + 12).Text = RecuperaValor(CadenaDevuelta, 1)
            Text2(Indice).Text = RecuperaValor(CadenaDevuelta, 2)

        Else
            'Recupera todo el registro de Proveedor
            HaDevueltoDatos = True
            Screen.MousePointer = vbHourglass
            'Sabemos que campos son los que nos devuelve
            'Creamos una cadena consulta y ponemos los datos
            cadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
            cadB = Aux
    
            'Se muestran en el mismo form
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
            Screen.MousePointer = vbDefault
        End If
    End If
End Sub


Private Sub frmBP_DatoSeleccionado(CadenaSeleccion As String)
'Banco Propio
    Text1(14).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(14)
    Text2(2).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento C. Postales
Dim Indice As Byte
Dim Devuelve As String

    Indice = 4
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
    'Poblacion
    Text1(Indice + 1).Text = ObtenerPoblacion(Text1(Indice).Text, Devuelve)
    'provincia
    Text1(Indice + 2).Text = Devuelve
End Sub

Private Sub frmF_Selec(vFecha As Date)
Dim Indice As Byte
    
    Indice = CByte(Val(imgFecha(0).Tag))
    Text1(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
'Forma de Pago
    Text1(13).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    Text2(1).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub



Private Sub frmS_DatoSeleccionado(CadenaSeleccion As String)
    Text1(29).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(29).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgCuentas_Click(Index As Integer)
Dim Indice As Byte

    If Modo = 2 Or Modo = 5 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cuenta Contable
            imgCuentas(0).Tag = Index
            'Conexi�n a BD: Conta, Tabla: Cuentas
            MandaBusquedaPrevia "apudirec='S'"
            imgCuentas(0).Tag = -1 'Abre el frmBuscaGrid para la conexi�n
                                   'de la BD: Ariges
            Indice = 12
        Case 1 'Forma de Pago
            Set frmFP = New frmFacFormasPago
            frmFP.DatosADevolverBusqueda = "0"
            frmFP.Show vbModal
            Set frmFP = Nothing
            Indice = 13
        Case 2 'Banco Propio
            Set frmBP = New frmFacBancosPropios
            frmBP.DatosADevolverBusqueda = "0"
            frmBP.Show vbModal
            Set frmBP = Nothing
            Indice = 14
        Case 3 'Cod. Postal
            Set frmCP = New frmCPostal
            frmCP.DatosADevolverBusqueda = "0"
            frmCP.Show vbModal
            Set frmCP = Nothing
            VieneDeBuscar = True
            Indice = 4
        Case 4
            Set frmS = New frmFacSituaciones
            frmS.DatosADevolverBusqueda = "0"
            frmS.Show vbModal
            Set frmS = Nothing
    End Select
    PonerFoco Text1(Indice)
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFecha_Click(Index As Integer)
Dim Indice As Integer

   If Modo = 2 Or Modo = 5 Or Modo = 0 Then
        If Index <> 2 Then Exit Sub
    End If
   Screen.MousePointer = vbHourglass
   
   
   If Index < 2 Then
        Indice = 8 + Index
   ElseIf Index = 2 Then
        'text1)30)
        Indice = 30
   Else
        Indice = 32
   End If
   imgFecha(0).Tag = Indice
   'FECHA
   Set frmF = New frmCal
   frmF.Fecha = Now
   
   PonerFormatoFecha Text1(Indice)
   If Text1(Indice).Text <> "" Then frmF.Fecha = CDate(Text1(Indice).Text)

   frmF.Show vbModal
   Set frmF = Nothing
   If Index <> 2 Then
        PonerFoco Text1(Indice)
    Else
        imgFecha(2).Tag = Text1(30).Text
        CargaDatosLW
    End If
End Sub


Private Sub ImgHomo_Click(Index As Integer)
Dim C As String

    If Modo <> 2 Then Exit Sub
    
    If Me.checkAlbFac(1).Value = 0 Then
        MsgBox "No esta homologado", vbExclamation
        Exit Sub
    End If
    
    
    If Index > 0 Then
        If lwH.SelectedItem Is Nothing Then Exit Sub
    End If
    
    If Index < 2 Then
        'Nuevo modificar
        If Index = 0 Then
            CadenaDesdeOtroForm = ""
        Else
            CadenaDesdeOtroForm = "codprove = " & Me.Text1(0).Text & " AND secuencial"
            CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "obshomo", "sproveHomo", CadenaDesdeOtroForm, Mid(lwH.SelectedItem.Key, 2))
            CadenaDesdeOtroForm = lwH.SelectedItem.Text & "|" & CadenaDesdeOtroForm & "|"
        End If
        
        frmVarios.Opcion = 11
        frmVarios.Show vbModal
        If CadenaDesdeOtroForm = "" Then Exit Sub
        'codprove,secuencial,usuario,fecha,obshomo
        If Index = 0 Then
            C = DevuelveDesdeBD(conAri, "max(secuencial)", "sproveHomo", "codprove", Text1(0).Text)
            C = Val(C) + 1
        Else
            C = Mid(lwH.SelectedItem.Key, 2)
        End If
        C = "sproveHomo SET codprove =" & Text1(0).Text & ",secuencial = " & C
        C = C & ",fecha =" & DBSet(RecuperaValor(CadenaDesdeOtroForm, 1), "F")
        C = C & ",obshomo =" & DBSet(RecuperaValor(CadenaDesdeOtroForm, 2), "F")
        If Index = 0 Then
            C = C & ",usuario=" & DBSet(vUsu.Login, "T")
            C = "INSERT INTO " & C
        Else
            C = C & " WHERE codprove=" & Text1(0).Text & " AND secuencial = " & Mid(lwH.SelectedItem.Key, 2)
            C = "UPDATE " & C
        End If
        
        
        
    Else
        'Eliminar
        C = lwH.SelectedItem.SubItems(1)
        If Len(C) > 50 Then C = Mid(C, 50, "....")
        C = "Accion: " & C & vbCrLf
        C = "Fecha " & lwH.SelectedItem.Text & vbCrLf & C
        
        C = "Seguro que desea eliminar la acci�n comercial: " & vbCrLf & C
        If MsgBox(C, vbQuestion + vbYesNo) <> vbYes Then Exit Sub
        
        C = "DELETE FROM sproveHomo WHERE codprove=" & Text1(0).Text & " AND secuencial = " & Mid(lwH.SelectedItem.Key, 2)
        
        
        
    End If
    
    
    If EjecutaSQL(conAri, C, True) Then
        If Index <> 1 Then
            CargaHomologaciones
        Else
            lwH.SelectedItem.Text = Format(RecuperaValor(CadenaDesdeOtroForm, 1), "dd/mm/yyyy")
            lwH.SelectedItem.SubItems(1) = Replace(RecuperaValor(CadenaDesdeOtroForm, 2), vbCrLf, " ")
        End If
    End If
End Sub

Private Sub ImgMail_Click(Index As Integer)
'Abrir Outlook para enviar e-mail
Dim dirMail As String

    If Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    If Index = 0 Then
        dirMail = Text1(20).Text
    ElseIf Index = 1 Then
        dirMail = Text1(24).Text
    End If
    If LanzaMailGnral(dirMail) Then Espera 2
    Screen.MousePointer = vbDefault
End Sub

Private Sub imgWeb_Click()
'Abrimos el explorador de windows con la pagina Web del cliente
    If Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
'    If LanzaHome("websoporte") Then espera 2
    If LanzaHomeGnral(Text1(27).Text) Then Espera 2
    Screen.MousePointer = vbDefault
End Sub


Private Sub lw1_DblClick()
Dim Seleccionado As Long
    If Modo <> 2 Then Exit Sub
    If lw1.ListItems.Count = 0 Then Exit Sub
    If lw1.SelectedItem Is Nothing Then Exit Sub


    If Me.DatosADevolverBusqueda <> "" Then
        'De momento NO dejo continuar
        MsgBox "Esta buscando un proveedor. No puede ver los documentos.", vbExclamation
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Llegados aqui
    Select Case CByte(RecuperaValor(lw1.Tag, 1))
    Case 2
        'ALBARANES
        'Set frmAlb = New frmFacEntAlbaranes
        
        frmComEntAlbaranes.cadSelAlbaranes = " numalbar='" & DevNombreSQL(lw1.SelectedItem.Text) & _
            "' AND fechaalb= '" & Format(lw1.SelectedItem.SubItems(1), "yyyy-mm-dd") & _
            "' AND codprove = " & Data1.Recordset!codProve
        
        frmComEntAlbaranes.Show vbModal
        frmComEntAlbaranes.cadSelAlbaranes = ""
    Case 0
        'OFERTAS
        'Set frmOfe = New frmFacEntOfertas
        'frmOfe.DatosOferta = lw1.SelectedItem.Text
        'frmOfe.Show vbModal
        'Set frmOfe = Nothing
    Case 1
        'PEDIDOS
        frmComEntPedidos.MostrarDatos = lw1.SelectedItem.Text
        frmComEntPedidos.EsHistorico = False
        frmComEntPedidos.Show vbModal
    Case 3
        'FACTURAS
        'Este no necesitamos crear instancias
        AbrirFacturaLW
        
        
    End Select
        
    'Pase lo que pase, por si acaso, cargamos el lw
    lw1.SetFocus
    Seleccionado = lw1.SelectedItem.Index
    CargaDatosLW
    lw1.SelectedItem.Selected = False
    Set lw1.SelectedItem = Nothing
    If lw1.ListItems.Count >= Seleccionado Then
            lw1.ListItems(Seleccionado).Selected = True
            lw1.ListItems(Seleccionado).EnsureVisible
    End If

End Sub


Private Sub lwH_DblClick()
    If Modo = 2 Then ImgHomo_Click (1)
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnModificar_Click()
    If BLOQUEADesdeFormulario(Me) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
     Kcampo = Index
     ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 33 Then KEYpress (KeyAscii)
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
On Error Resume Next

    'Si no estamos Insertando o Modificando no hacemos nada
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    Select Case Index
        Case 0 'proveedor
            PonerFormatoEntero Text1(Index)
            
        Case 4 'Cod. Postal
            If Text1(Index).Locked Then Exit Sub
            If Text1(Index).Text <> "" And Not VieneDeBuscar Then
                Text1(Index + 1).Text = ObtenerPoblacion(Text1(Index).Text, Devuelve)
                Text1(Index + 2).Text = Devuelve
            Else
                Text1(Index + 1).Text = ""
                Text1(Index + 2).Text = ""
            End If
            VieneDeBuscar = False
            
        Case 7 'NIF
            If Text1(Index).Text <> "" Then
                Text1(Index).Text = UCase(Text1(Index).Text)
                ValidarNIF Text1(Index).Text
            End If
       
        Case 8, 9, 32 'Fechas
            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
            
        Case 10, 11 'Descuentos
            'Formato tipo 4: Decimal(4,2)
            If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 4
            
        Case 12 'Cta Contable
            Text2(0).Text = PonerNombreCuenta(Text1(Index), Modo, Text1(0).Text)
            
        Case 13 ' Forma Pago
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(1).Text = PonerNombreDeCod(Text1(Index), conAri, "sforpa", "nomforpa", "codforpa")
            Else
                Text2(1).Text = ""
            End If
            
        Case 14 'Banco Propio
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(2).Text = PonerNombreDeCod(Text1(Index), conAri, "sbanpr", "nombanpr", "codbanpr")
            Else
                Text2(2).Text = ""
            End If
            
        Case 15, 16, 17, 18, 31 'cuenta bancaria
            If Index = 31 Then
                Me.Text1(31).Text = UCase(Me.Text1(31).Text)
                If Me.Text1(31).Text <> "" Then IBAN_Correcto Me.Text1(31).Text
            Else
                PonerFormatoEntero Text1(Index)
            End If
            If Index = 18 Or Index = 31 Then
                
                Devuelve = Text1(15).Text & Text1(16).Text & Text1(17).Text & Text1(18).Text
                If Len(Devuelve) = 20 Then
                    'OK. Calculamos el IBAN
                    If Text1(31).Text = "" Then
                        'NO ha puesto IBAN
                        If DevuelveIBAN2("ES", Devuelve, Devuelve) Then Text1(31).Text = "ES" & Devuelve
                    Else
                        CadenaConsulta = CStr(Mid(Text1(31).Text, 1, 2))
                        If DevuelveIBAN2(CadenaConsulta, Devuelve, Devuelve) Then
                            If Mid(Text1(31).Text, 3) <> Devuelve Then
                                
                                MsgBox "Codigo IBAN distinto del calculado [" & CadenaConsulta & Devuelve & "]", vbExclamation
                                
                            End If
                        End If
                    End If
                End If
                CadenaConsulta = ""
            
           End If
        Case 27
            PonerFocoBtn Me.cmdAceptar
            
        Case 29
            If PonerFormatoEntero(Text1(29)) Then
                Text2(29).Text = PonerNombreDeCod(Text1(29), conAri, "ssitua", "nomsitua", "codsitua")
            Else
                Text2(29).Text = ""
            End If
    End Select
End Sub



'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
Private Sub PonerModo(Kmodo As Byte)
Dim I As Byte
Dim b As Boolean
Dim NumReg As Byte

    Modo = Kmodo
    PonerIndicador lblIndicador, Kmodo
    
    'Modo 2. Hay datos y estamos visualizandolos
    b = (Kmodo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = b
    Else
        cmdRegresar.visible = False
    End If
    
    'Poner botones de desplazamiento visible si Modo 2
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    
    '----------------------------------------------------------------
    b = (Kmodo >= 3) Or Modo = 1 'Modo: Insertar/Modificar o Busqueda
    Me.cboTipoProv.Enabled = b
    Me.cboTipoDto.Enabled = b
    Me.chkProveV.Enabled = b 'proveedor varios
    checkAlbFac(0).Enabled = b           'Solo si al aplicacion lleva REA veremos este check
    checkAlbFac(1).Enabled = b
    cmdAceptar.visible = b
    cmdCancelar.visible = b
    
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar adem�s limpia los campos Text1
    BloquearText1 Me, Modo
    'Fecha ult. compra bloqueada pq se modifica por programa
    BloquearTxt Text1(9), (Modo <> 1)
    'La fecha esta NUNCA se puede escribir
    Text1(30).Enabled = False
    Text1(30).Text = Me.imgFecha(2).Tag
    
    
    chkVistaPrevia.Enabled = (Modo <= 2)
    
    'Poner longitud de los campos
    PonerLongCampos
    
    PonerModoOpcionesMenu 'Activar opciones de menu seg�n modo
    PonerOpcionesMenu   'Activar opciones de menu seg�n nivel
                        'de permisos del usuario
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de b�squeda o no
'para los campos que permitan introducir criterios m�s largos del tama�o del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Sub PonerModoOpcionesMenu()
Dim b As Boolean
    
    b = (Modo = 2) Or (Modo = 0) Or (Modo = 1)
    'Insertar
    Toolbar1.Buttons(5).Enabled = b
    Me.mnNuevo.Enabled = b
      
    b = (Modo = 2)
    'Modificar
    Toolbar1.Buttons(6).Enabled = b
    mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(7).Enabled = b
    mnEliminar.Enabled = b

    b = (Modo >= 3) 'Modo: Insertar/Modificar
    'Buscar
    Toolbar1.Buttons(1).Enabled = Not b
    Me.mnBuscar.Enabled = Not b
    'VerTodos
    Toolbar1.Buttons(2).Enabled = Not b
    Me.mnVerTodos.Enabled = Not b
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub

Private Sub PonerCampos()
    
    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    Text2(0).Text = PonerNombreCuenta(Text1(12), Modo)

    'Rellenar Text2 con nombre asociado al codigo
    Text2(1).Text = DevuelveDesdeBDNew(conAri, "sforpa", "nomforpa", "codforpa", Text1(13).Text, "N")
    Text2(2).Text = DevuelveDesdeBDNew(conAri, "sbanpr", "nombanpr", "codbanpr", Text1(14).Text, "N")
        
    'Poner la situacion
    Modo = 3   'peque�a trampa para que haga el losfocus
    Text1_LostFocus 29
    Modo = 2
    
    
    
    Me.Refresh
    DoEvents
    CargaDatosLW
    
    
    'Homologaciones
    CargaHomologaciones
    

    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount

End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
            
    DatosOk = False
    b = CompForm(Me, 1)
    If Not b Then Exit Function
        
    'Validar que la cuenta bancaria es correcta
    If Trim(Text1(15).Text & Text1(16).Text & Text1(17).Text & Text1(18).Text) <> "" Then
        b = Comprueba_CuentaBan(Text1(15).Text & Text1(16).Text & Text1(17).Text & Text1(18).Text)
    End If
        
    'Si esta homologado, tiene que tener fecha homologacion
    If Me.checkAlbFac(1).Value = 1 Then
        If Me.Text1(32).Text = "" Then
            MsgBox "Si esta homologado debe indicar la fecha de homologaci�n", vbExclamation
            b = False
        End If
    End If
    DatosOk = b
End Function


Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 1 'Buscar
        mnBuscar_Click
    Case 2 'Recuperar Todos
        mnVerTodos_Click
    Case 5  'Insertar Nuevo
        mnNuevo_Click
    Case 6  'Modificar
        mnModificar_Click
    Case 7  'Borrar
        mnEliminar_Click
    Case 10 'Salir
        mnSalir_Click
    Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento:Primero,Anterior,Siguiente,Ultimo
        Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub

Private Sub BotonAnyadir()
    LimpiarCampos
    PonerModo 3 'Modo 3: Insertar
    'Obtenemos la siguiente numero de codigo de Proveedor
    Text1(0).Text = SugerirCodigoSiguienteStr("sprove", "codprove")
    Text1(8).Text = Format(Now, "dd/mm/yyyy")
    Me.cboTipoProv.ListIndex = 0
    
    Text1(29).Text = vParamAplic.PorDefecto_Situ
    Text1_LostFocus 29
    
    PonerFoco Text1(0)   'Ponemos el foco
End Sub


Private Sub BotonEliminar()
Dim Cad As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    '### a mano
    Cad = "�Seguro que desea eliminar el Proveedor?"
    Cad = Cad & vbCrLf & "Cod. : " & Data1.Recordset.Fields(0)
    Cad = Cad & vbCrLf & "Nombre: " & Data1.Recordset.Fields(1)

    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        Data1.Recordset.Delete
        If SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            PonerModo 0
        End If
    End If
Error2:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then
            Data1.Recordset.CancelUpdate
            MuestraError Err.Number, "Eliminar Proveedor", Err.Description
        End If
End Sub


Private Sub BotonModificar()
    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    PonerFoco Me.Text1(2)
End Sub

Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        BuscaChekc = ""
        PonerModo 1
        PonerFoco Text1(0)
        Text1(0).BackColor = vbYellow
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(Kcampo).Text = ""
            PonerFoco Text1(Kcampo)
            Text1(Kcampo).BackColor = vbYellow
        End If
    End If
End Sub


Private Sub BotonVerTodos()
    LimpiarCampos
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub

Private Sub CargarComboTipoProveedor()
'###
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Nacional, 1-Intracomunitario, 2-Extranjero  3- Regimen especial agrario
'4- Estimaci�n directa

    cboTipoProv.Clear
    cboTipoProv.AddItem "Nacional"
    cboTipoProv.ItemData(cboTipoProv.NewIndex) = 0
    
    cboTipoProv.AddItem "Intracomunitario"
    cboTipoProv.ItemData(cboTipoProv.NewIndex) = 1
    
    cboTipoProv.AddItem "Extranjero"
    cboTipoProv.ItemData(cboTipoProv.NewIndex) = 2
    
    cboTipoProv.AddItem "R.E.A."
    cboTipoProv.ItemData(cboTipoProv.NewIndex) = 3
    
    cboTipoProv.AddItem "Estimaci�n directa"
    cboTipoProv.ItemData(cboTipoProv.NewIndex) = 4
End Sub

Private Sub CargarComboTipoDto()
'### Combo Tipo Descuento
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Aditivo, 1-Resto

    cboTipoDto.Clear
    cboTipoDto.AddItem "Aditivo"
    cboTipoDto.ItemData(cboTipoDto.NewIndex) = 0
    
    cboTipoDto.AddItem "Resto"
    cboTipoDto.ItemData(cboTipoDto.NewIndex) = 1
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    cboTipoDto.ListIndex = -1
    cboTipoProv.ListIndex = -1
    Me.chkProveV.Value = 0
    Me.checkAlbFac(0).Value = 0
    Me.checkAlbFac(1).Value = 0
    Me.lw1.ListItems.Clear
    Me.lwH.ListItems.Clear
End Sub


Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False, BuscaChekc)
    
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia cadB
    Else
        'Se muestran en el mismo form
        If cadB <> "" Then
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
        End If
    End If
End Sub


Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
Dim Tabla As String
Dim Titulo As String
Dim Conexion As String

    'Llamamos a al form
    Cad = ""
    Select Case Val(Me.imgCuentas(0).Tag)
        Case 0 'Se llama a Busqueda desde el campo: Cuenta Contable
            '#A MANO: Porque busca en la Tabla: Cuentas
            'de la BDatos de Contabilidad
            Cad = Cad & "C�digo|cuentas|codmacta|T||30�Denominacion|cuentas|nommacta|T||70�"
            Tabla = "cuentas"
            Titulo = "Cuentas"
            Conexion = conConta    'Conexi�n a BD: Conta
        Case Else 'Se llama a Busqueda desde el registro Proveedor
            Cad = Cad & ParaGrid(Text1(0), 20, "C�digo")
            Cad = Cad & ParaGrid(Text1(1), 40, "Nombre")
            Cad = Cad & ParaGrid(Text1(2), 41, "Nombre Comercial")
            Tabla = "sprove"
            Titulo = "Proveedores"
            Conexion = conAri    'Conexi�n a BD: Ariges
    End Select
        
    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = Tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = Titulo
        frmB.vselElem = 1
        frmB.vConexionGrid = Conexion
        frmB.vCargaFrame = (Conexion = 2)
'        frmB.vBuscaPrevia = chkVistaPrevia
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
        If HaDevueltoDatos Then
            PonerFoco Text1(Kcampo + 1)
'                If (Not adodc1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'                    cmdRegresar_Click
        Else   'de ha devuelto datos, es decir NO ha devuelto datos
            PonerFoco Text1(Kcampo)
        End If
    End If
End Sub


Private Sub PonerCadenaBusqueda()
    On Error GoTo EEPonerBusq

    Screen.MousePointer = vbHourglass
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    If Data1.Recordset.EOF Then
        MsgBox "No hay ning�n registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        PonerModo 2
        Data1.Recordset.MoveFirst
        PonerCampos
        PonerFocoBtn Me.cmdRegresar
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub Desplazamiento(Index As Integer)
'Para desplazarse por los registros de control Data
    DesplazamientoData Data1, Index
    PonerCampos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub



Private Sub PosicionarData()
Dim Cad As String, Indicador As String

    Cad = "(codprove=" & Text1(0).Text & ")"
    If SituarData(Data1, Cad, Indicador) Then
       PonerModo 2
       lblIndicador.Caption = Indicador
    Else
'       LimpiarCampos
        PonerModo 0
    End If
End Sub





'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
'
'
'       El listview tendra los datos de albaranes, facturas... que tenga el cliente
'       Con lo cual, a partir de un click tendremos que ser capaces de situarnos en
'       el formulario correspondiente
'
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------


Private Sub ImagenesNavegacion()
    With Me.Toolbar2
        .ImageList = frmppal.ImgListPpal
        '.Buttons(1).Image = 5
        .Buttons(3).Image = 9
        .Buttons(5).Image = 10
        .Buttons(7).Image = 11
        '.Buttons(8).Image = 5
    End With
    
    Set lw1.SmallIcons = frmppal.ImgListPpal
End Sub


Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Tag = "" Then Exit Sub
    Label2(0).Caption = ""
    'Levantamos todos los botones y dejamos pulsado el de ahora
    For NumRegElim = 1 To Toolbar2.Buttons.Count
        If Toolbar2.Buttons(NumRegElim).Tag <> "" Then
            If Toolbar2.Buttons(NumRegElim).Index <> Button.Index Then Toolbar2.Buttons(NumRegElim).Value = tbrUnpressed
        End If
    Next NumRegElim
    
    'Hacemos las acciones
    CargaColumnas CByte(Button.Tag)
    If Modo = 2 Then CargaDatosLW
End Sub




Private Sub CargaColumnas(OpcionList As Byte)
Dim Columnas As String
Dim Ancho As String
Dim Alinea As String
Dim Formato As String
Dim Ncol As Integer
Dim C As ColumnHeader

    Select Case OpcionList
    Case 2
        'ALBARANES
        
        Label2(0).Caption = "Albaranes"
        Columnas = "Numero|Fecha||Importe|"
        Ancho = "2500|1500|0|2500|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|1|"
        'Formatos
        Formato = "|dd/mm/yyyy||" & FormatoImporte & "|"
        Ncol = 4
               
               
    Case 3
        
        Label2(0).Caption = "Facturas"
        Columnas = "Numero|Fecha|F. recepcion|Importe|"
        Ancho = "1800|1300|1300|2000|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|1|"
        'Formatos
        Formato = "|dd/mm/yyyy|dd/mm/yyyy|" & FormatoImporte & "|"
        Ncol = 4
               
    Case 1
        'PEDIDOS
        
        Label2(0).Caption = "Pedidos"
        Columnas = "Visado"
        
        Columnas = "Numero|Fecha|Importe|"
        Ancho = "2000|2000|1800|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|1|"
        'Formatos
        Formato = "00000000|dd/mm/yyyy|" & FormatoImporte & "|"
        Ncol = 3
    'Case 2
        '
    End Select
    
    
    'Fecha incio busquedas
    Text1(30).Text = Format(imgFecha(2).Tag, "dd/mm/yyyy")
    'Guardo la opcion en el tag
    lw1.Tag = OpcionList & "|" & Ncol & "|"
    
    lw1.ColumnHeaders.Clear
    
    For NumRegElim = 1 To Ncol
         Set C = lw1.ColumnHeaders.Add()
         C.Text = RecuperaValor(Columnas, CInt(NumRegElim))
         C.Width = RecuperaValor(Ancho, CInt(NumRegElim))
         C.Alignment = Val(RecuperaValor(Alinea, CInt(NumRegElim)))
         C.Tag = RecuperaValor(Formato, CInt(NumRegElim))
    Next NumRegElim
End Sub


Private Sub CargaDatosLW()
Dim C As String
Dim bs As Byte
    bs = Screen.MousePointer
    C = Me.lblIndicador.Caption
    lblIndicador.Caption = "Leyendo " & Label2(0).Caption
    lblIndicador.Refresh
    CargaDatosLW2
    Me.lblIndicador.Caption = C
    Screen.MousePointer = bs
End Sub



Private Sub CargaDatosLW2()
Dim Cad As String
Dim RS As ADODB.Recordset
Dim It As ListItem
Dim ElIcono As Integer
Dim GroupBy As String
Dim BuscaChekc
    On Error GoTo ECargaDatosLW
    
    
    If Modo <> 2 Then Exit Sub
    
    For NumRegElim = 1 To Toolbar2.Buttons.Count
        If Toolbar2.Buttons(NumRegElim).Value = tbrPressed Then
            ElIcono = Toolbar2.Buttons(NumRegElim).Image
            Exit For
        End If
    Next
    
    
    'Fecha incio busquedas
    Text1(30).Text = Format(imgFecha(2).Tag, "dd/mm/yyyy")
    
    
    Select Case CByte(RecuperaValor(lw1.Tag, 1))
    Case 2
        'ALBARANES
        Cad = "select c.numalbar,c.fechaalb,c.codprove as codprove,sum(importel) from scaalp c,slialp l where c.codprove=l.codprove and c.numalbar=l.numalbar"
        
        'QUE NO MUESTRE EL "B"
        If Not vUsu.TrabajadorB Then Cad = Cad & " AND presupuesto = 0"
        
        GroupBy = "1,2,3"
        BuscaChekc = "c.fechaalb"
        
    Case 0
        'OFERTAS
        'cad = "select c.numofert,c.fecofert,fecentre,if(aceptado=1,""SI"","" "") ,sum(importel) from scapre c,slipre l where"
        'cad = cad & " c.numofert=l.numofert "
        'GroupBy = "1,2"
        'BuscaChekc = "fecofert"
    Case 1
        'PEDIDOS
        Cad = "select c.numpedpr,c.fecpedpr,sum(importel) from scappr c,slippr l where "
         Cad = Cad & " c.numpedpr=l.numpedpr  "
        BuscaChekc = "fecpedpr"
        GroupBy = "1"
    Case 3
        Cad = "select numfactu,fecfactu,fecrecep,totalfac from scafpc c WHERE 1=1"
        'QUE NO MUESTRE EL "B"
        If Not vUsu.TrabajadorB Then Cad = Cad & " AND presupuesto = 0"
        BuscaChekc = "fecfactu"
        GroupBy = "1,2,3"
    End Select
    
    
    'La fecha
    
    'EL where del codclien
    Cad = Cad & " and c.codprove=" & Data1.Recordset!codProve
    
    'La fecha
    Cad = Cad & " and " & BuscaChekc & " >='" & Format(imgFecha(2).Tag, FormatoFecha) & "'"
    
    
    'El group by
    Cad = Cad & " GROUP BY " & GroupBy
    
    'El ORDER BY
    Cad = Cad & " ORDER BY " & BuscaChekc & " DESC"
    BuscaChekc = ""
    
    lw1.ListItems.Clear
    Set RS = New ADODB.Recordset
    RS.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        Set It = lw1.ListItems.Add()
        If lw1.ColumnHeaders(1).Tag <> "" Then
            It.Text = Format(RS.Fields(0), lw1.ColumnHeaders(1).Tag)
        Else
            It.Text = RS.Fields(0)
        End If
        'El resto de cmpos
        For NumRegElim = 2 To CInt(RecuperaValor(lw1.Tag, 2))
            If IsNull(RS.Fields(NumRegElim - 1)) Then
                It.SubItems(NumRegElim - 1) = " "
            Else
                If lw1.ColumnHeaders(NumRegElim).Tag <> "" Then
                    It.SubItems(NumRegElim - 1) = Format(RS.Fields(NumRegElim - 1), lw1.ColumnHeaders(NumRegElim).Tag)
                Else
                    It.SubItems(NumRegElim - 1) = RS.Fields(NumRegElim - 1)
                End If
            End If
        Next
        It.SmallIcon = ElIcono
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    
    Exit Sub
ECargaDatosLW:
    MuestraError Err.Number
    Set RS = Nothing
    
End Sub



Private Sub AbrirFacturaLW()
Dim s As String
    
    Set miRsAux = New ADODB.Recordset

    s = "select numalbar,fechaalb from scafpa where numfactu='" & DevNombreSQL(lw1.SelectedItem.Text)
    s = s & "' and fecfactu='" & Format(lw1.SelectedItem.SubItems(1), FormatoFecha) & "' ORDER BY numalbar desc"
    miRsAux.Open s, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    s = ""
    If Not miRsAux.EOF Then
        s = DevNombreSQL(miRsAux.Fields(0)) & "|" & miRsAux.Fields(1) & "|"
    End If
    miRsAux.Close
    Set miRsAux = Nothing

    
    If s <> "" Then
        With frmComHcoFacturas
            .hcoCodMovim = RecuperaValor(s, 1)
            .hcoFechaMovim = RecuperaValor(s, 2)
            .hcoCodProve = Data1.Recordset!codProve
            .Show vbModal
        End With
    Else
        MsgBox "No se han encontrado los albaranes de la factura", vbExclamation
    End If
End Sub



Private Sub CargaHomologaciones()
Dim It As ListItem

    lwH.ListItems.Clear
    If Me.checkAlbFac(1).Value = 0 Then Exit Sub
    
    Set miRsAux = New ADODB.Recordset

    miRsAux.Open "select fecha,obshomo,secuencial from  sproveHomo where codprove=" & Me.Text1(0).Text & " ORDER BY fecha desc,secuencial desc", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set It = lwH.ListItems.Add(, "S" & miRsAux!secuencial)
        It.Text = Format(miRsAux!Fecha, "dd/mm/yyyy")
        
        It.SubItems(1) = Replace(miRsAux!obshomo, vbCrLf, " ")
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
End Sub



