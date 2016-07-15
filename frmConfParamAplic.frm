VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmConfParamAplic 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parámetros de la Aplicación"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   9420
   Icon            =   "frmConfParamAplic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   9420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   8115
      TabIndex        =   74
      Top             =   7560
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   240
      TabIndex        =   77
      Top             =   7440
      Width           =   3000
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   120
         TabIndex        =   78
         Top             =   210
         Width           =   2280
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6960
      TabIndex        =   73
      Top             =   7560
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8115
      TabIndex        =   75
      Top             =   7560
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   76
      Top             =   0
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   8160
      Top             =   360
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      Height          =   6735
      Left            =   240
      TabIndex        =   79
      Top             =   600
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   11880
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Datos Varios"
      TabPicture(0)   =   "frmConfParamAplic.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(6)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(5)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(4)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(14)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "imgBuscar(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(55)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text1(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "FrameDiasMante"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "FrameOpciones"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Text1(5)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text1(4)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "FramePrecioKm"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text1(1)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cboTipodtos"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cboOrdenDtos"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text2(1)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text1(58)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "Datos Facturación"
      TabPicture(1)   =   "frmConfParamAplic.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame10"
      Tab(1).Control(1)=   "Frame9"
      Tab(1).Control(2)=   "Text1(50)"
      Tab(1).Control(3)=   "chkTicketsAgrupads"
      Tab(1).Control(4)=   "Frame6"
      Tab(1).Control(5)=   "Frame4"
      Tab(1).Control(6)=   "Frame3"
      Tab(1).Control(7)=   "Frame2"
      Tab(1).Control(8)=   "Label1(51)"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Internet"
      TabPicture(2)   =   "frmConfParamAplic.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameEMail"
      Tab(2).Control(1)=   "FrameSoporte"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Datos Contabilidad "
      TabPicture(3)   =   "frmConfParamAplic.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Text1(52)"
      Tab(3).Control(1)=   "Text2(52)"
      Tab(3).Control(2)=   "cboObsFactura"
      Tab(3).Control(3)=   "Text2(48)"
      Tab(3).Control(4)=   "Text1(48)"
      Tab(3).Control(5)=   "Text1(49)"
      Tab(3).Control(6)=   "Text2(47)"
      Tab(3).Control(7)=   "Text1(47)"
      Tab(3).Control(8)=   "Text1(46)"
      Tab(3).Control(9)=   "Text2(46)"
      Tab(3).Control(10)=   "Frame8"
      Tab(3).Control(11)=   "Text1(23)"
      Tab(3).Control(12)=   "Text1(22)"
      Tab(3).Control(13)=   "Text1(21)"
      Tab(3).Control(14)=   "Text1(20)"
      Tab(3).Control(15)=   "Label1(47)"
      Tab(3).Control(16)=   "imgBuscar(45)"
      Tab(3).Control(17)=   "Label1(53)"
      Tab(3).Control(18)=   "Label1(52)"
      Tab(3).Control(19)=   "imgBuscar(41)"
      Tab(3).Control(20)=   "Label1(50)"
      Tab(3).Control(21)=   "Label1(49)"
      Tab(3).Control(22)=   "Label2(6)"
      Tab(3).Control(23)=   "Label2(7)"
      Tab(3).Control(24)=   "imgBuscar(40)"
      Tab(3).Control(25)=   "Label1(48)"
      Tab(3).Control(26)=   "imgBuscar(39)"
      Tab(3).Control(27)=   "Label1(19)"
      Tab(3).Control(28)=   "Label1(18)"
      Tab(3).Control(29)=   "Label1(17)"
      Tab(3).Control(30)=   "Label1(15)"
      Tab(3).ControlCount=   31
      TabCaption(4)   =   "Valores por defecto / AVISOS"
      TabPicture(4)   =   "frmConfParamAplic.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame5"
      Tab(4).Control(1)=   "Frame7"
      Tab(4).ControlCount=   2
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   320
         Index           =   58
         Left            =   6720
         MaxLength       =   10
         TabIndex        =   191
         Tag             =   "PesoEtiqueta|N|S|0||spara1|PesoEtiqueta|0.0000||"
         Text            =   "Text1"
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Frame Frame10 
         Caption         =   "Reciclado / Punto verde"
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
         Height          =   855
         Left            =   -74880
         TabIndex        =   180
         Top             =   4920
         Width           =   8655
         Begin VB.TextBox Text1 
            Height          =   320
            Index           =   53
            Left            =   2040
            MaxLength       =   16
            TabIndex        =   29
            Tag             =   "Reci. |T|S|||spara1|ArtReciclado|||"
            Text            =   "Text1 "
            Top             =   297
            Width           =   1815
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   53
            Left            =   3840
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   181
            Text            =   "Text2"
            Top             =   300
            Width           =   4665
         End
         Begin VB.Label Label1 
            Caption         =   "Artículo "
            Height          =   195
            Index           =   54
            Left            =   600
            TabIndex        =   182
            Top             =   360
            Width           =   780
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   53
            Left            =   1680
            Tag             =   "-1"
            ToolTipText     =   "Buscar artículo"
            Top             =   337
            Width           =   240
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   320
         Index           =   52
         Left            =   -72480
         MaxLength       =   2
         TabIndex        =   45
         Tag             =   "IVAexento|N|S|0||spara1|IvaIntracom|||"
         Text            =   "Text1"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   52
         Left            =   -71760
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   178
         Text            =   "Text2"
         Top             =   1680
         Width           =   3105
      End
      Begin VB.ComboBox cboObsFactura 
         Height          =   315
         Left            =   -71520
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Tag             =   "Orden Descuentos|N|S|||spara1|obsfactura|||"
         Top             =   960
         Width           =   3135
      End
      Begin VB.Frame Frame9 
         Caption         =   "Aportación en facturas"
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
         Height          =   735
         Left            =   -74880
         TabIndex        =   173
         Top             =   4080
         Width           =   8655
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   51
            Left            =   2520
            MaxLength       =   10
            TabIndex        =   28
            Tag             =   "Cta aportacion|N|S|||spara1|ctaaportacion|||"
            Text            =   "3"
            Top             =   240
            Width           =   1260
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   51
            Left            =   3840
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   174
            Text            =   "Text2"
            Top             =   240
            Width           =   4185
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   42
            Left            =   1800
            Tag             =   "-1"
            ToolTipText     =   "Buscar cuenta contable"
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label2 
            Caption         =   "Cuenta"
            Height          =   255
            Index           =   8
            Left            =   1080
            TabIndex        =   175
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   320
         Index           =   50
         Left            =   -68640
         MaxLength       =   2
         TabIndex        =   31
         Tag             =   "NºConta|N|S|1|99|spara1|conta_B|||"
         Text            =   "Text1"
         Top             =   6120
         Width           =   735
      End
      Begin VB.CheckBox chkTicketsAgrupads 
         Caption         =   "Contabilizar ticket TPV agrupados"
         Height          =   375
         Left            =   -74160
         TabIndex        =   30
         Tag             =   "Tickets agrupadsos|N|N|||spara1|conttickagrupado|||"
         Top             =   6000
         Width           =   3135
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   48
         Left            =   -71760
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   170
         Text            =   "Text2"
         Top             =   1320
         Width           =   3105
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   320
         Index           =   48
         Left            =   -72480
         MaxLength       =   2
         TabIndex        =   44
         Tag             =   "IVAexento|N|S|0||spara1|ivaexento|||"
         Text            =   "Text1"
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   49
         Left            =   -67560
         MaxLength       =   5
         TabIndex        =   48
         Tag             =   "Nº Contabilidad|N|S|||spara1|porreten|||"
         Text            =   "3"
         Top             =   2520
         Width           =   660
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   47
         Left            =   -71160
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   166
         Text            =   "Text2"
         Top             =   2520
         Width           =   3105
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   47
         Left            =   -72480
         MaxLength       =   10
         TabIndex        =   47
         Tag             =   "Cta retencion|N|S|||spara1|ctareten|||"
         Text            =   "3"
         Top             =   2520
         Width           =   1260
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   320
         Index           =   46
         Left            =   -72480
         MaxLength       =   2
         TabIndex        =   46
         Tag             =   "REA|N|S|0||spara1|iva_rea|||"
         Text            =   "Text1"
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   46
         Left            =   -71760
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   164
         Text            =   "Text2"
         Top             =   2040
         Width           =   3105
      End
      Begin VB.Frame Frame8 
         Caption         =   "IVA recargo de equivalencia"
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
         Height          =   3375
         Left            =   -74880
         TabIndex        =   149
         Top             =   3000
         Width           =   8055
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   57
            Left            =   4920
            MaxLength       =   2
            TabIndex        =   56
            Tag             =   "IVA1|N|S|0|99|spara1|iva_oldre2|||"
            Text            =   "Text1"
            Top             =   1920
            Width           =   495
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   57
            Left            =   5520
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   188
            Text            =   "Text2"
            Top             =   1920
            Width           =   2385
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   56
            Left            =   4920
            MaxLength       =   2
            TabIndex        =   54
            Tag             =   "IVA1|N|S|0|99|spara1|iva_old2|||"
            Text            =   "Text1"
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   56
            Left            =   5520
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   187
            Text            =   "Text2"
            Top             =   1560
            Width           =   2385
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   55
            Left            =   1200
            MaxLength       =   2
            TabIndex        =   55
            Tag             =   "IVA1|N|S|0|99|spara1|iva_oldre1|||"
            Text            =   "Text1"
            Top             =   1920
            Width           =   495
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   55
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   185
            Text            =   "Text2"
            Top             =   1920
            Width           =   2265
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   54
            Left            =   1200
            MaxLength       =   2
            TabIndex        =   53
            Tag             =   "IVA1|N|S|0|99|spara1|iva_old1|||"
            Text            =   "Text1"
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   54
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   183
            Text            =   "Text2"
            Top             =   1560
            Width           =   2265
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   42
            Left            =   3480
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   158
            Text            =   "Text2"
            Top             =   2880
            Width           =   3105
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   45
            Left            =   3480
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   157
            Text            =   "Text2"
            Top             =   2520
            Width           =   3105
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   41
            Left            =   5520
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   156
            Text            =   "Text2"
            Top             =   960
            Width           =   2385
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   44
            Left            =   5520
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   155
            Text            =   "Text2"
            Top             =   600
            Width           =   2385
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   40
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   154
            Text            =   "Text2"
            Top             =   960
            Width           =   2265
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   43
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   153
            Text            =   "Text2"
            Top             =   600
            Width           =   2265
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   42
            Left            =   2880
            MaxLength       =   2
            TabIndex        =   58
            Tag             =   "IVRE3|N|S|0|99|spara1|ivare3eq|||"
            Text            =   "Text1"
            Top             =   2880
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   41
            Left            =   4920
            MaxLength       =   2
            TabIndex        =   52
            Tag             =   "IVRE2|N|S|0|99|spara1|ivare2eq|||"
            Text            =   "Text1"
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   40
            Left            =   1200
            MaxLength       =   2
            TabIndex        =   51
            Tag             =   "IVRE1|N|S|0|99|spara1|ivare1eq|||"
            Text            =   "Text1"
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   43
            Left            =   1200
            MaxLength       =   2
            TabIndex        =   49
            Tag             =   "IVA1|N|S|0|99|spara1|ivare1|||"
            Text            =   "Text1"
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   44
            Left            =   4920
            MaxLength       =   2
            TabIndex        =   50
            Tag             =   "IVA2|N|S|0|99|spara1|ivare2|||"
            Text            =   "Text1"
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   45
            Left            =   2880
            MaxLength       =   2
            TabIndex        =   57
            Tag             =   "IVA3|N|S|0|99|spara1|ivare3|||"
            Text            =   "Text1"
            Top             =   2520
            Width           =   615
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   50
            Left            =   4680
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   1920
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   49
            Left            =   4680
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   1560
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   48
            Left            =   960
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   1920
            Width           =   240
         End
         Begin VB.Label Label2 
            Caption         =   "Ant. RE"
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   186
            Top             =   1920
            Width           =   735
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   47
            Left            =   960
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   1560
            Width           =   240
         End
         Begin VB.Label Label2 
            Caption         =   "Antiguo"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   184
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "R.E."
            Height          =   255
            Index           =   5
            Left            =   1920
            TabIndex        =   163
            Top             =   2880
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "Normal"
            Height          =   255
            Index           =   4
            Left            =   1920
            TabIndex        =   162
            Top             =   2520
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "Normal"
            Height          =   255
            Index           =   2
            Left            =   2760
            TabIndex        =   161
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "R.E."
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   160
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "Normal"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   159
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "General"
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
            Index           =   45
            Left            =   240
            TabIndex        =   152
            Top             =   240
            Width           =   675
         End
         Begin VB.Label Label1 
            Caption         =   "Reducido"
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
            Index           =   44
            Left            =   4200
            TabIndex        =   151
            Top             =   240
            Width           =   825
         End
         Begin VB.Label Label1 
            Caption         =   "Super-Reducido"
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
            Index           =   46
            Left            =   240
            TabIndex        =   150
            Top             =   2520
            Width           =   1380
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   36
            Left            =   960
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   600
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   33
            Left            =   960
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   960
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   37
            Left            =   4680
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   600
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   34
            Left            =   4680
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   960
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   38
            Left            =   2520
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   2520
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   35
            Left            =   2520
            Tag             =   "-1"
            ToolTipText     =   "Buscar I.V.A."
            Top             =   2880
            Width           =   240
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Avisos"
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
         Height          =   2415
         Left            =   -74880
         TabIndex        =   137
         Top             =   3360
         Width           =   8535
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   33
            Left            =   3600
            MaxLength       =   2
            TabIndex        =   66
            Tag             =   "ped. cli|N|S|0||spara1|avipedcli|||"
            Text            =   "3"
            Top             =   315
            Width           =   780
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   34
            Left            =   6720
            MaxLength       =   2
            TabIndex        =   67
            Tag             =   "ped.pro.|N|S|0||spara1|avipedpro|||"
            Text            =   "3"
            Top             =   315
            Width           =   780
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   35
            Left            =   3600
            MaxLength       =   2
            TabIndex        =   68
            Tag             =   "alb.cli.|N|S|0||spara1|avialbcli|||"
            Text            =   "3"
            Top             =   720
            Width           =   780
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   36
            Left            =   6720
            MaxLength       =   2
            TabIndex        =   69
            Tag             =   "alb.pro.|N|S|0||spara1|avialbpro|||"
            Text            =   "3"
            Top             =   720
            Width           =   780
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   37
            Left            =   3600
            MaxLength       =   2
            TabIndex        =   70
            Tag             =   "avi.mante|N|S|0||spara1|avimanteni|||"
            Text            =   "3"
            Top             =   1275
            Width           =   780
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   38
            Left            =   3600
            MaxLength       =   2
            TabIndex        =   72
            Tag             =   "avi.avisos|N|S|0||spara1|aviavios|||"
            Text            =   "3"
            Top             =   1995
            Width           =   780
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   39
            Left            =   3600
            MaxLength       =   2
            TabIndex        =   71
            Tag             =   "avi.repa.|N|S|0||spara1|avirepara|||"
            Text            =   "3"
            Top             =   1635
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Pedidos clientes"
            Height          =   195
            Index           =   33
            Left            =   2040
            TabIndex        =   148
            Top             =   360
            Width           =   1155
         End
         Begin VB.Label Label1 
            Caption         =   "Pedidos proveedores"
            Height          =   195
            Index           =   34
            Left            =   4680
            TabIndex        =   147
            Top             =   360
            Width           =   1500
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Albaranes clientes"
            Height          =   195
            Index           =   35
            Left            =   2040
            TabIndex        =   146
            Top             =   765
            Width           =   1290
         End
         Begin VB.Label Label1 
            Caption         =   "Albaranes proveedores"
            Height          =   195
            Index           =   36
            Left            =   4680
            TabIndex        =   145
            Top             =   765
            Width           =   1635
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Mantenimientos"
            Height          =   195
            Index           =   37
            Left            =   2040
            TabIndex        =   144
            Top             =   1320
            Width           =   1110
         End
         Begin VB.Label Label1 
            Caption         =   "Reparaciones"
            Height          =   195
            Index           =   38
            Left            =   2040
            TabIndex        =   143
            Top             =   1680
            Width           =   1155
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Avisos "
            Height          =   195
            Index           =   39
            Left            =   2040
            TabIndex        =   142
            Top             =   2040
            Width           =   510
         End
         Begin VB.Label Label1 
            Caption         =   "Dias desde la fecha"
            Height          =   195
            Index           =   40
            Left            =   120
            TabIndex        =   141
            Top             =   360
            Width           =   7275
         End
         Begin VB.Label Label1 
            Caption         =   "No facturados"
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
            Index           =   41
            Left            =   4680
            TabIndex        =   140
            Top             =   1320
            Width           =   2955
         End
         Begin VB.Label Label1 
            Caption         =   "Pendientes de reparar sin motivo de reparación"
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
            Index           =   42
            Left            =   4680
            TabIndex        =   139
            Top             =   1680
            Width           =   3555
         End
         Begin VB.Label Label1 
            Caption         =   "Abiertos"
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
            Index           =   43
            Left            =   4680
            TabIndex        =   138
            Top             =   2040
            Width           =   2955
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Telefonía"
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
         Height          =   855
         Left            =   -74880
         TabIndex        =   134
         Top             =   3120
         Width           =   8655
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   32
            Left            =   3840
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   135
            Text            =   "Text2"
            Top             =   300
            Width           =   4665
         End
         Begin VB.TextBox Text1 
            Height          =   320
            Index           =   32
            Left            =   2040
            MaxLength       =   16
            TabIndex        =   27
            Tag             =   "Recar |T|S|||spara1|codartictel|||"
            Text            =   "Text1 "
            Top             =   297
            Width           =   1815
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   32
            Left            =   1680
            Tag             =   "-1"
            ToolTipText     =   "Buscar artículo"
            Top             =   337
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Artículo a facturar"
            Height          =   195
            Index           =   32
            Left            =   240
            TabIndex        =   136
            Top             =   360
            Width           =   1380
         End
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   1
         Left            =   3360
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   133
         Text            =   "Text2"
         Top             =   1560
         Width           =   4065
      End
      Begin VB.Frame Frame5 
         Caption         =   "Clientes"
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
         Height          =   2895
         Left            =   -74880
         TabIndex        =   118
         Top             =   360
         Width           =   8535
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   31
            Left            =   120
            MaxLength       =   3
            TabIndex        =   65
            Tag             =   "Agente|N|S|0|999|spara1|defagente|000||"
            Text            =   "Tex"
            Top             =   2520
            Width           =   735
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   31
            Left            =   960
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   131
            Text            =   "Text2"
            Top             =   2520
            Width           =   3105
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   30
            Left            =   4440
            MaxLength       =   3
            TabIndex        =   64
            Tag             =   "Tarifa|N|S|0|999|spara1|deftarifa|000||"
            Text            =   "Tex"
            Top             =   1800
            Width           =   735
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   30
            Left            =   5280
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   124
            Text            =   "Text2"
            Top             =   1800
            Width           =   3105
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   29
            Left            =   120
            MaxLength       =   3
            TabIndex        =   63
            Tag             =   "Situacion|N|S|0|999|spara1|defstituacion|000||"
            Text            =   "Tex"
            Top             =   1800
            Width           =   735
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   29
            Left            =   960
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   123
            Text            =   "Text2"
            Top             =   1800
            Width           =   3105
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   28
            Left            =   4440
            MaxLength       =   3
            TabIndex        =   62
            Tag             =   "Ruta|N|S|0|999|spara1|defruta|000||"
            Text            =   "Tex"
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   28
            Left            =   5280
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   122
            Text            =   "Text2"
            Top             =   1080
            Width           =   3105
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   27
            Left            =   120
            MaxLength       =   3
            TabIndex        =   61
            Tag             =   "Zona|N|S|0|999|spara1|defzona|000||"
            Text            =   "Tex"
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   27
            Left            =   960
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   121
            Text            =   "Text2"
            Top             =   1080
            Width           =   3105
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   26
            Left            =   4440
            MaxLength       =   3
            TabIndex        =   60
            Tag             =   "Envio|N|S|0|999|spara1|defenvio|000||"
            Text            =   "Tex"
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   26
            Left            =   5280
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   120
            Text            =   "Text2"
            Top             =   480
            Width           =   3105
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   25
            Left            =   120
            MaxLength       =   3
            TabIndex        =   59
            Tag             =   "Actividad|N|S|0||spara1|defactividad|000||"
            Text            =   "Tex"
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   25
            Left            =   960
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   119
            Text            =   "Text2"
            Top             =   480
            Width           =   3105
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   31
            Left            =   840
            ToolTipText     =   "Buscar agente"
            Top             =   2280
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   30
            Left            =   4920
            ToolTipText     =   "Buscar tarifa"
            Top             =   1560
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   29
            Left            =   840
            ToolTipText     =   "Buscar situacion"
            Top             =   1560
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   25
            Left            =   840
            Tag             =   "-1"
            ToolTipText     =   "Buscar actividad"
            Top             =   240
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   28
            Left            =   4920
            ToolTipText     =   "Buscar ruta"
            Top             =   840
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   27
            Left            =   840
            ToolTipText     =   "Buscar zona"
            Top             =   840
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   26
            Left            =   4920
            ToolTipText     =   "Buscar forma de envio"
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Agente"
            Height          =   255
            Index           =   31
            Left            =   120
            TabIndex        =   132
            Top             =   2280
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Tarifa"
            Height          =   255
            Index           =   30
            Left            =   4440
            TabIndex        =   130
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Situación"
            Height          =   255
            Index           =   29
            Left            =   120
            TabIndex        =   129
            Top             =   1560
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Ruta"
            Height          =   195
            Index           =   28
            Left            =   4440
            TabIndex        =   128
            Top             =   840
            Width           =   345
         End
         Begin VB.Label Label1 
            Caption         =   "Zona"
            Height          =   255
            Index           =   27
            Left            =   120
            TabIndex        =   127
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Envio"
            Height          =   195
            Index           =   26
            Left            =   4440
            TabIndex        =   126
            Top             =   240
            Width           =   405
         End
         Begin VB.Label Label1 
            Caption         =   "Actividad"
            Height          =   255
            Index           =   25
            Left            =   120
            TabIndex        =   125
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Cheques  regalo"
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
         Height          =   735
         Left            =   -74880
         TabIndex        =   115
         Top             =   2280
         Width           =   8655
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   24
            Left            =   3960
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   116
            Text            =   "Text2"
            Top             =   240
            Width           =   4550
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   24
            Left            =   3000
            MaxLength       =   3
            TabIndex        =   26
            Tag             =   "Forma de pago para cheque regalo |N|S|0|999|spara1|codforpa|000||"
            Text            =   "Tex"
            Top             =   237
            Width           =   735
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   24
            Left            =   2640
            Tag             =   "-1"
            ToolTipText     =   "Buscar forma pago"
            Top             =   270
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Forma de pago "
            Height          =   255
            Index           =   24
            Left            =   1320
            TabIndex        =   117
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   23
         Left            =   -72720
         MaxLength       =   30
         TabIndex        =   39
         Tag             =   "Servidor Contabilidad|T|S|||spara1|serconta|||"
         Text            =   "3wwwwwwwwwwwwwwwwwwwwwwwwwwwww"
         Top             =   555
         Width           =   1620
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   22
         Left            =   -66600
         MaxLength       =   2
         TabIndex        =   42
         Tag             =   "Nº Contabilidad|N|S|||spara1|numconta|||"
         Text            =   "3"
         Top             =   555
         Width           =   300
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   21
         Left            =   -70440
         MaxLength       =   20
         TabIndex        =   40
         Tag             =   "Usuario Contabilidad|T|S|||spara1|usuconta|||"
         Text            =   "3wwwwwwwwwwwwwwwwwww"
         Top             =   555
         Width           =   900
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   20
         Left            =   -68880
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   41
         Tag             =   "Password Contabilidad|T|S|||spara1|pasconta|||"
         Text            =   "3"
         Top             =   555
         Width           =   1140
      End
      Begin VB.ComboBox cboOrdenDtos 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "Orden Descuentos|N|N|||spara1|ordendto|||"
         Top             =   2220
         Width           =   1815
      End
      Begin VB.Frame Frame3 
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
         Height          =   855
         Left            =   -74880
         TabIndex        =   107
         Top             =   1320
         Width           =   8655
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   19
            Left            =   7320
            MaxLength       =   2
            TabIndex        =   25
            Tag             =   "Mes a no girar|N|S|0|12|spara1|mesnogir|||"
            Text            =   "Text1"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   18
            Left            =   4080
            MaxLength       =   2
            TabIndex        =   24
            Tag             =   "Dia 3 de pago compras|N|S|0|31|spara1|diapago3|||"
            Text            =   "Text1"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   17
            Left            =   3360
            MaxLength       =   2
            TabIndex        =   23
            Tag             =   "Dia 2 de pago compras|N|S|0|31|spara1|diapago2|||"
            Text            =   "Text1"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   16
            Left            =   2640
            MaxLength       =   2
            TabIndex        =   22
            Tag             =   "Dia 1 de pago compras|N|S|0|31|spara1|diapago1|||"
            Text            =   "Text1"
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Mes a no girar"
            Height          =   255
            Index           =   13
            Left            =   6120
            TabIndex        =   109
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Días de pago"
            Height          =   255
            Index           =   11
            Left            =   1440
            TabIndex        =   108
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Desplazamientos"
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
         Height          =   855
         Left            =   -74880
         TabIndex        =   104
         Top             =   360
         Width           =   8655
         Begin VB.TextBox Text1 
            Height          =   320
            Index           =   15
            Left            =   2040
            MaxLength       =   16
            TabIndex        =   21
            Tag             =   "Artículo para facturar desplazamientos |T|S|||spara1|codartid|||"
            Text            =   "Text1 "
            Top             =   327
            Width           =   1815
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   15
            Left            =   3840
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   105
            Text            =   "Text2"
            Top             =   330
            Width           =   4665
         End
         Begin VB.Label Label1 
            Caption         =   "Artículo a facturar"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   106
            Top             =   360
            Width           =   1455
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   15
            Left            =   1680
            Tag             =   "-1"
            ToolTipText     =   "Buscar artículo"
            Top             =   360
            Width           =   240
         End
      End
      Begin VB.Frame FrameSoporte 
         ForeColor       =   &H00972E0B&
         Height          =   1635
         Left            =   -74760
         TabIndex        =   99
         Top             =   3840
         Width           =   8355
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   12
            Left            =   1860
            MaxLength       =   100
            TabIndex        =   36
            Tag             =   "Web de Soporte|T|S|||spara1|websoporte|||"
            Text            =   "3"
            Top             =   300
            Width           =   6060
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   13
            Left            =   1860
            MaxLength       =   100
            TabIndex        =   37
            Tag             =   "Mail de Soporte|T|S|||spara1|mailsoporte|||"
            Text            =   "3"
            Top             =   690
            Width           =   6060
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   14
            Left            =   1860
            MaxLength       =   100
            TabIndex        =   38
            Tag             =   "Version Web|T|S|||spara1|webversion|||"
            Text            =   "3"
            Top             =   1080
            Width           =   6060
         End
         Begin VB.Label Label1 
            Caption         =   "Web de soporte"
            Height          =   195
            Index           =   9
            Left            =   300
            TabIndex        =   103
            Top             =   360
            Width           =   1140
         End
         Begin VB.Label Label1 
            Caption         =   "Mail soporte"
            Height          =   195
            Index           =   12
            Left            =   300
            TabIndex        =   102
            Top             =   720
            Width           =   1080
         End
         Begin VB.Label Label1 
            Caption         =   "Web check version"
            Height          =   195
            Index           =   16
            Left            =   300
            TabIndex        =   101
            Top             =   1140
            Width           =   1500
         End
         Begin VB.Label Label8 
            Caption         =   "Soporte"
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
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   100
            Top             =   0
            Width           =   840
         End
      End
      Begin VB.Frame FrameEMail 
         Height          =   2055
         Left            =   -74760
         TabIndex        =   93
         Top             =   840
         Width           =   8355
         Begin VB.TextBox Text1 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   11
            Left            =   1440
            MaxLength       =   30
            PasswordChar    =   "*"
            TabIndex        =   35
            Tag             =   "Password SMTP|T|S|||spara1|smtppass|||"
            Text            =   "3"
            Top             =   1560
            Width           =   4260
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   10
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   34
            Tag             =   "Usuario SMTP|T|S|||spara1|smtpuser|||"
            Text            =   "3"
            Top             =   1180
            Width           =   4260
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   9
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   33
            Tag             =   "Servidor SMTP|T|S|||spara1|smtphost|||"
            Text            =   "3"
            Top             =   800
            Width           =   5700
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   8
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   32
            Tag             =   "Direccion e-mail|T|S|||spara1|diremail|||"
            Text            =   "3"
            Top             =   420
            Width           =   5700
         End
         Begin VB.Label Label8 
            Caption         =   "Envio E-Mail"
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
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   98
            Top             =   0
            Width           =   1320
         End
         Begin VB.Label Label1 
            Caption         =   "Password"
            Height          =   195
            Index           =   23
            Left            =   300
            TabIndex        =   97
            Top             =   1620
            Width           =   840
         End
         Begin VB.Label Label1 
            Caption         =   "Usuario"
            Height          =   195
            Index           =   22
            Left            =   300
            TabIndex        =   96
            Top             =   1260
            Width           =   780
         End
         Begin VB.Label Label1 
            Caption         =   "Servidor SMTP"
            Height          =   195
            Index           =   21
            Left            =   300
            TabIndex        =   95
            Top             =   840
            Width           =   1140
         End
         Begin VB.Label Label1 
            Caption         =   "E-Mail"
            Height          =   195
            Index           =   20
            Left            =   300
            TabIndex        =   94
            Top             =   480
            Width           =   900
         End
      End
      Begin VB.ComboBox cboTipodtos 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Tag             =   "Tipo Descuentos|N|N|||spara1|tipodtos|||"
         Top             =   2880
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   320
         Index           =   1
         Left            =   2640
         MaxLength       =   3
         TabIndex        =   2
         Tag             =   "Código Tarifa PVP|N|N|||spara1|codtarif|000||"
         Text            =   "Text1"
         Top             =   1560
         Width           =   615
      End
      Begin VB.Frame FramePrecioKm 
         Caption         =   "Precio Km"
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
         Height          =   1215
         Left            =   4080
         TabIndex        =   84
         Top             =   2040
         Width           =   2415
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   2
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   5
            Tag             =   "Precio Km desplaz. Clientes|N|S|0|9999.0000|spara1|preukmcl|#,##0.0000||"
            Text            =   "Text1"
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   3
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   6
            Tag             =   "Precio Km desplaz. Técnicos|N|S|0|9999.0000|spara1|preukmtc|#,##0.0000||"
            Text            =   "Text1"
            Top             =   660
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "A Clientes"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   86
            Top             =   255
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Técnicos"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   85
            Top             =   660
            Width           =   660
         End
      End
      Begin VB.TextBox Text1 
         Height          =   320
         Index           =   4
         Left            =   3360
         MaxLength       =   35
         TabIndex        =   0
         Tag             =   "Nombre Director Gerente|T|S|||spara1|nomgeren|||"
         Text            =   "Text1"
         Top             =   540
         Width           =   4095
      End
      Begin VB.TextBox Text1 
         Height          =   320
         Index           =   5
         Left            =   3360
         MaxLength       =   35
         TabIndex        =   1
         Tag             =   "Nombre responsable Admon|T|S|||spara1|nomadmin|||"
         Text            =   "Text1"
         Top             =   900
         Width           =   4095
      End
      Begin VB.Frame FrameOpciones 
         Caption         =   "Opciones"
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
         Height          =   3255
         Left            =   240
         TabIndex        =   83
         Top             =   3360
         Width           =   6255
         Begin VB.CheckBox chkEmpresaExportador 
            Caption         =   "Empresa exportadora"
            Height          =   375
            Left            =   3120
            TabIndex        =   189
            Tag             =   "AVAB|N|N|||spara1|EsEmpresaExportadora|||"
            Top             =   1680
            Width           =   2775
         End
         Begin VB.CheckBox chkNuevaProducion 
            Caption         =   "Nuevo sistema de produccion"
            Height          =   375
            Left            =   3120
            TabIndex        =   20
            Tag             =   "Descriptores|N|N|||spara1|ProduccionNueva|||"
            Top             =   1320
            Width           =   2775
         End
         Begin VB.CheckBox chkDescriptores 
            Caption         =   "Usa descriptores especiales"
            Height          =   375
            Left            =   3120
            TabIndex        =   19
            Tag             =   "Descriptores|N|N|||spara1|descriptores|||"
            Top             =   960
            Width           =   2775
         End
         Begin VB.CheckBox chkProduccion 
            Caption         =   "Tiene produccion"
            Height          =   375
            Left            =   3120
            TabIndex        =   18
            Tag             =   "Tiene produccion|N|N|||spara1|produccion|||"
            Top             =   600
            Width           =   2775
         End
         Begin VB.CheckBox chkHayServicio 
            Caption         =   "Hay Servicios"
            Height          =   375
            Left            =   240
            TabIndex        =   13
            Tag             =   "Hay Servicios|N|N|||spara1|hayservicio|||"
            Top             =   1680
            Width           =   2175
         End
         Begin VB.CheckBox chkCajacomp 
            Caption         =   "Cajas completas precios"
            Height          =   375
            Left            =   240
            TabIndex        =   9
            Tag             =   "Cajas Completas Precios|N|N|||spara1|cajacomp|||"
            Top             =   240
            Width           =   2175
         End
         Begin VB.CheckBox chkHaymante 
            Caption         =   "Realiza Mantenimientos"
            Height          =   375
            Left            =   240
            TabIndex        =   10
            Tag             =   "Mantenimientos|N|N|||spara1|haymante|||"
            Top             =   600
            Width           =   2175
         End
         Begin VB.CheckBox chkHayrepar 
            Caption         =   "Realiza Reparaciones"
            Height          =   375
            Left            =   240
            TabIndex        =   11
            Tag             =   "Reparaciones|N|N|||spara1|hayrepar|||"
            Top             =   960
            Width           =   2175
         End
         Begin VB.CheckBox chkHayfrecu 
            Caption         =   "Hay Frecuencias"
            Height          =   375
            Left            =   240
            TabIndex        =   12
            Tag             =   "Hay Frecuencias|N|N|||spara1|hayfrecu|||"
            Top             =   1320
            Width           =   2175
         End
         Begin VB.CheckBox chkHaydepar 
            Caption         =   "Departamentos (o Dirección)"
            Height          =   375
            Left            =   240
            TabIndex        =   15
            Tag             =   "Departamento/Direc.|N|N|||spara1|haydepar|||"
            Top             =   2400
            Width           =   2775
         End
         Begin VB.CheckBox chkctrstock 
            Caption         =   "Realiza control de Stock"
            Height          =   375
            Left            =   240
            TabIndex        =   16
            Tag             =   "Control de Stock|N|N|||spara1|ctrstock|||"
            Top             =   2760
            Width           =   2775
         End
         Begin VB.CheckBox chkInventar 
            Caption         =   "Realiza Inventario por Proveedor"
            Height          =   375
            Left            =   3120
            TabIndex        =   17
            Tag             =   "Inventarios por Proveedor|N|N|||spara1|inventar|||"
            Top             =   240
            Width           =   2775
         End
         Begin VB.CheckBox chkHaynserie 
            Caption         =   "Hay Nº Serie en Compras"
            Height          =   375
            Left            =   240
            TabIndex        =   14
            Tag             =   "Hay Nº Serie en Compras|N|N|||spara1|haynserie|||"
            Top             =   2040
            Width           =   2175
         End
      End
      Begin VB.Frame FrameDiasMante 
         Caption         =   "Días Reparación"
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
         Height          =   1215
         Left            =   6720
         TabIndex        =   80
         Top             =   2040
         Width           =   1935
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   6
            Left            =   960
            MaxLength       =   4
            TabIndex        =   8
            Tag             =   "Dias Repar. sin Mantenimiento|N|N|0|9999|spara1|diasnoman|||"
            Text            =   "Text"
            Top             =   680
            Width           =   855
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   7
            Left            =   960
            MaxLength       =   4
            TabIndex        =   7
            Tag             =   "Dias Repar. con Mantenimiento|N|N|0|9999|spara1|diassiman|||"
            Text            =   "Text1"
            Top             =   300
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Sin Mto"
            Height          =   195
            Index           =   7
            Left            =   240
            TabIndex        =   82
            Top             =   675
            Width           =   540
         End
         Begin VB.Label Label1 
            Caption         =   "Con Mto"
            Height          =   195
            Index           =   8
            Left            =   240
            TabIndex        =   81
            Top             =   300
            Width           =   600
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   4260
         MaxLength       =   15
         TabIndex        =   91
         Tag             =   "Código Parámetros Aplic|N|N|||spara1|codigo||S|"
         Text            =   "Text1"
         Top             =   540
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Peso etiqueta (gr)"
         Height          =   195
         Index           =   55
         Left            =   6720
         TabIndex        =   192
         Top             =   3840
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "IVA intracomunitario"
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
         Index           =   47
         Left            =   -74760
         TabIndex        =   179
         Top             =   1680
         Width           =   1725
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   45
         Left            =   -72840
         Tag             =   "-1"
         ToolTipText     =   "Buscar I.V.A."
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Integracion facturas. Observaciones "
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
         Index           =   53
         Left            =   -74760
         TabIndex        =   177
         Top             =   960
         Width           =   3165
      End
      Begin VB.Label Label1 
         Caption         =   "R.E.A."
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
         Index           =   52
         Left            =   -74760
         TabIndex        =   176
         Top             =   2040
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Conta presupuestos *"
         Height          =   255
         Index           =   51
         Left            =   -70440
         TabIndex        =   172
         Top             =   6120
         Width           =   2175
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   41
         Left            =   -72840
         Tag             =   "-1"
         ToolTipText     =   "Buscar I.V.A."
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "IVA exento"
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
         Index           =   50
         Left            =   -74760
         TabIndex        =   171
         Top             =   1320
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "Contabilidad"
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
         Index           =   49
         Left            =   -74760
         TabIndex        =   169
         Top             =   600
         Width           =   1185
      End
      Begin VB.Label Label2 
         Caption         =   "%"
         Height          =   195
         Index           =   6
         Left            =   -67800
         TabIndex        =   168
         Top             =   2520
         Width           =   120
      End
      Begin VB.Label Label2 
         Caption         =   "Cuenta"
         Height          =   255
         Index           =   7
         Left            =   -73680
         TabIndex        =   167
         Top             =   2520
         Width           =   615
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   40
         Left            =   -72840
         Tag             =   "-1"
         ToolTipText     =   "Buscar cuenta contable"
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Retención"
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
         Index           =   48
         Left            =   -74760
         TabIndex        =   165
         Top             =   2520
         Width           =   1155
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   39
         Left            =   -72840
         Tag             =   "-1"
         ToolTipText     =   "Buscar I.V.A."
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   2085
         Tag             =   "-1"
         ToolTipText     =   "Buscar tarifa"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Servidor"
         Height          =   195
         Index           =   19
         Left            =   -73440
         TabIndex        =   114
         Top             =   600
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "Nº conta"
         Height          =   195
         Index           =   18
         Left            =   -67560
         TabIndex        =   113
         Top             =   600
         Width           =   660
      End
      Begin VB.Label Label1 
         Caption         =   "Usuario"
         Height          =   195
         Index           =   17
         Left            =   -71040
         TabIndex        =   112
         Top             =   600
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "Pass."
         Height          =   195
         Index           =   15
         Left            =   -69360
         TabIndex        =   111
         Top             =   600
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Orden Descuentos"
         Height          =   255
         Index           =   14
         Left            =   480
         TabIndex        =   110
         Top             =   2220
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de Descuentos"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   90
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Código Tarifa de PVP"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   89
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre del Director Gerente"
         Height          =   255
         Index           =   4
         Left            =   480
         TabIndex        =   88
         Top             =   540
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre responsable Administración"
         Height          =   255
         Index           =   5
         Left            =   480
         TabIndex        =   87
         Top             =   900
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Código"
         Height          =   255
         Index           =   6
         Left            =   3720
         TabIndex        =   92
         Top             =   1020
         Width           =   495
      End
   End
   Begin VB.Label Label3 
      Caption         =   "EXPORTADORA,MOIX SOLO DESDE YOG"
      Height          =   495
      Left            =   3360
      TabIndex        =   190
      Top             =   7440
      Width           =   2895
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   1
         Shortcut        =   ^M
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmConfParamAplic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmMtoArt As frmAlmArticulos
Attribute frmMtoArt.VB_VarHelpID = -1
Private WithEvents frmFP As frmFacFormasPago 'Form Mto Formas de Pago
Attribute frmFP.VB_VarHelpID = -1


Private WithEvents frmA As frmFacActividades
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmE As frmFacFormasEnvio
Attribute frmE.VB_VarHelpID = -1
Private WithEvents frmZ As frmFacZonas
Attribute frmZ.VB_VarHelpID = -1
Private WithEvents frmR As frmFacRutas
Attribute frmR.VB_VarHelpID = -1
Private WithEvents frmAC As frmFacAgentesCom
Attribute frmAC.VB_VarHelpID = -1
Private WithEvents frmT As frmFacTarifas
Attribute frmT.VB_VarHelpID = -1
Private WithEvents frmS As frmFacSituaciones
Attribute frmS.VB_VarHelpID = -1


Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1


Private NombreTabla As String  'Nombre de la tabla o de la
Private CadenaConsulta As String



Dim PrimeraVez As Boolean
Dim Modo As Byte
'0: Inicial
'2: Visualizacion
'3: Añadir
'4: Modificar






Private Sub cboObsFactura_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cboOrdenDtos_KeyPress(KeyAscii As Integer)
      KEYpress KeyAscii
End Sub

Private Sub cboTipodtos_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub chkCajacomp_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkCajacomp_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkctrstock_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkctrstock_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub




Private Sub chkEmpresaExportador_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkEmpresaExportador_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkHaydepar_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkHaydepar_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkHayfrecu_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkHayfrecu_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub chkHaymante_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkHaymante_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub chkHaynserie_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkHaynserie_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub chkHayrepar_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkHayrepar_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkHayServicio_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkHayServicio_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkInventar_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkInventar_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub chkInventar_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub




Private Sub chkNuevaProducion_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkNuevaProducion_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkProduccion_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkProduccion_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkDescriptores_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkDescriptores_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkTicketsAgrupads_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkTicketsAgrupads_KeyPress(KeyAscii As Integer)
  KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim actualiza As Boolean
Dim kms As Currency

    If Modo = 4 Then 'MODIFICAR
        If DatosOk Then
            vParamAplic.TipoDtos = Me.cboTipodtos.ListIndex
            vParamAplic.OrdenDtos = Me.cboOrdenDtos.ListIndex
            vParamAplic.ObsFactura = Me.cboObsFactura.ListIndex
            vParamAplic.CodTarifa = Text1(1).Text
            vParamAplic.NomGerente = Text1(4).Text
            vParamAplic.NomRespAdmin = Text1(5).Text
            kms = ImporteFormateado(ComprobarCero(Text1(2).Text))
            vParamAplic.PrecioKmClientes = CSng(CStr(kms))
            kms = ImporteFormateado(ComprobarCero(Text1(3).Text))
            vParamAplic.PrecioKmTecnicos = CSng(CStr(kms))
            vParamAplic.CajasCompletas = Me.chkCajacomp.Value
            vParamAplic.Mantenimientos = Me.chkHaymante.Value
            vParamAplic.Reparaciones = Me.chkHayrepar.Value
            vParamAplic.Frecuencias = Me.chkHayfrecu.Value
            vParamAplic.Servicios = Me.chkHayServicio.Value
            vParamAplic.Departamento = Me.chkHaydepar.Value
            vParamAplic.ControlStock = Me.chkctrstock.Value
            vParamAplic.InventarioxProv = Me.chkInventar.Value
            vParamAplic.NumSeries = Me.chkHaynserie.Value  'Hay Nº Serie en Compras?
            vParamAplic.DiasSiMante = Me.Text1(7).Text 'Dias Rep. con Mantenimiento
            vParamAplic.DiasNoMante = Me.Text1(6).Text 'Dias Rep. sin Mantenimiento
            
            'Articulo para facturar mantenimientos
            vParamAplic.ArticDesplaz = Me.Text1(15).Text
            'dias de pago para compras
            vParamAplic.DiaPago1 = CByte(DBLet(ComprobarCero(Text1(16).Text), "N"))
            vParamAplic.DiaPago2 = CByte(DBSet(Text1(17).Text, "N"))
            vParamAplic.DiaPago3 = CByte(DBSet(Text1(18).Text, "N"))
            vParamAplic.MesNoGirar = CByte(DBSet(Text1(19).Text, "N"))
            vParamAplic.ForPagoChequeRegalo = Me.Text1(24).Text
            
            vParamAplic.DireMail = Text1(8).Text 'Direccion email
            vParamAplic.SMTPhost = Text1(9).Text
            vParamAplic.SMTPuser = Text1(10).Text
            vParamAplic.SMTPpass = Text1(11).Text
            vParamAplic.WebSoporte = Text1(12).Text
            vParamAplic.MailSoporte = Text1(13).Text
            vParamAplic.WebVersion = Text1(14).Text
            
            'Datos contabilidad
            vParamAplic.ServidorConta = Text1(23).Text
            vParamAplic.UsuarioConta = Text1(21).Text
            vParamAplic.PasswordConta = Text1(20).Text
            vParamAplic.NumeroConta = ComprobarCero(Text1(22).Text)
            
            'Valores por defecto
            vParamAplic.PorDefecto_Activ = ComprobarCero(Text1(25).Text)
            vParamAplic.PorDefecto_Envio = ComprobarCero(Text1(26).Text)
            vParamAplic.PorDefecto_Zona = ComprobarCero(Text1(27).Text)
            vParamAplic.PorDefecto_Ruta = ComprobarCero(Text1(28).Text)
            vParamAplic.PorDefecto_Situ = ComprobarCero(Text1(29).Text)
            vParamAplic.PorDefecto_Tarifa = ComprobarCero(Text1(30).Text)
            vParamAplic.PorDefecto_Agente = ComprobarCero(Text1(31).Text)
            
            'Telefonia
            vParamAplic.CodarticTfnia = Me.Text1(32).Text
            
            'Los avisos
            vParamAplic.avipedcli = ComprobarCero(Text1(33).Text)
            vParamAplic.avipedpro = ComprobarCero(Text1(34).Text)
            vParamAplic.avialbcli = ComprobarCero(Text1(35).Text)
            vParamAplic.avialbpro = ComprobarCero(Text1(36).Text)
            vParamAplic.avimanteni = ComprobarCero(Text1(37).Text)
            vParamAplic.aviavisos = ComprobarCero(Text1(38).Text)
            vParamAplic.avirepara = ComprobarCero(Text1(39).Text)
            
            
            'Los tipos de IVA
            vParamAplic.TipoIVAre1 = ComprobarCero(Text1(40).Text)
            vParamAplic.TipoIVAre2 = ComprobarCero(Text1(41).Text)
            vParamAplic.TipoIVAre3 = ComprobarCero(Text1(42).Text)
             
            vParamAplic.TipoIVA1 = ComprobarCero(Text1(43).Text)
            vParamAplic.TipoIVA2 = ComprobarCero(Text1(44).Text)
            vParamAplic.TipoIVA3 = ComprobarCero(Text1(45).Text)
            
            
            
            'REtencion y REA
            vParamAplic.IVA_REA = ComprobarCero(Text1(46).Text)
            vParamAplic.CtaReten = ComprobarCero(Text1(47).Text)
            vParamAplic.PorReten = ComprobarCero(Text1(49).Text)
            
            'IVA exento
            vParamAplic.IVA_Exento2 = ComprobarCero(Text1(48).Text)
            vParamAplic.IVA_Intracomunitario = ComprobarCero(Text1(52).Text)

            
            'Tickets acgrupados
            vParamAplic.ContabilizarTicketAgrupados = Me.chkTicketsAgrupads.Value
            
            vParamAplic.ContabilidadB = ComprobarCero(Text1(50).Text)
            vParamAplic.ctaAportacion = Text1(51).Text
            
            vParamAplic.Produccion = Me.chkProduccion.Value
            vParamAplic.Descriptores = Me.chkDescriptores.Value
            
            vParamAplic.ArtReciclado = Text1(53).Text
            
            
            
            'Los tipos de IVA ANTIGUOS
            vParamAplic.OLDIVA1 = ComprobarCero(Text1(54).Text)
            vParamAplic.OLDIVA2 = ComprobarCero(Text1(56).Text)
            
             
            vParamAplic.OLDIVAre1 = ComprobarCero(Text1(55).Text)
            vParamAplic.OLDIVAre2 = ComprobarCero(Text1(57).Text)
            
            vParamAplic.ProduccionNueva = Me.chkNuevaProducion.Value
            
            'vParamAplic.EsAVAB  = Me.chkEmpresaExportador
            
            
            actualiza = vParamAplic.Modificar(Text1(0).Text)
            TerminaBloquear

            If actualiza Then  'Inserta o Modifica
                'Abrir la conexion a la conta q hemos modificado
                CerrarConexionConta
                If vParamAplic.NumeroConta <> 0 Then
                    If Not AbrirConexionConta(False) Then End
                End If
                PonerModo 2
                PonerFocoBtn Me.cmdSalir
            End If
        End If
    End If
End Sub


Private Sub cmdCancelar_Click()
    TerminaBloquear
    If Data1.Recordset.EOF Then
        PonerModo 0
        LimpiarCampos
    Else
        PonerCampos
        PonerModo 2
    End If
End Sub


Private Sub cmdSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerModo 0
    Else
        If Modo <> 4 Then PonerCadenaBusqueda
        PonerFoco Text1(0)
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
Dim Im
    'Icono del formulario
    Me.Icon = frmppal.Icon

    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmppal.imgListComun
'        .Buttons(1).Image = 3   'Anyadir
        .Buttons(1).Image = 4   'Modificar
        .Buttons(4).Image = 15  'Salir
    End With
    
    'cargar iconos de busqueda
    For Each Im In Me.imgBuscar
        Im.Picture = frmppal.imgListComun.ListImages(19).Picture
    Next
    'imgBuscar(1).Picture = frmPpal.imgListComun.ListImages(19).Picture
    'imgBuscar(15).Picture = frmPpal.imgListComun.ListImages(19).Picture
   '
   ' For NumRegElim = 24 To 42
   '     Me.imgBuscar(NumRegElim).Picture = frmPpal.imgListComun.ListImages(19).Picture
   ' Next NumRegElim
    
    
    

    LimpiarCampos   'Limpia los campos TextBox
    CargarComboTipoDtos
    CargarComboOrdenDtos
    CargaComoboObsFactura
    Me.SSTab1.Tab = 0
    
    NombreTabla = "spara1"
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    CadenaConsulta = "Select * from " & NombreTabla
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    
    Label3.visible = (vUsu.Codigo Mod 1000) = 0
    
    PonerModo 0

End Sub


Private Sub PonerCadenaBusqueda()
On Error GoTo EEPonerBusq
    Screen.MousePointer = vbHourglass

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        LimpiarCampos
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        Me.Toolbar1.Buttons(1).Enabled = False 'Modificar
    Else
        Data1.Recordset.MoveFirst
        PonerCampos
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    Screen.MousePointer = vbDefault
End Sub






Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
    'ENVIO
    Text1(25).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(25).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmAC_DatoSeleccionado(CadenaSeleccion As String)
    'agentes
    Text1(31).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(31).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    CadenaDesdeOtroForm = CadenaDevuelta
End Sub

Private Sub frmE_DatoSeleccionado(CadenaSeleccion As String)
    'ENVIO
    Text1(26).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(26).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Formas de Pago
Dim Indice As Byte

    Indice = 24
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Forma Pago
    FormateaCampo Text1(Indice)
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Forma Pago
End Sub

Private Sub frmMtoArt_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Articulos
    
    Text1(CInt(imgBuscar(1).Tag)).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod articulo
    Text2(CInt(imgBuscar(1).Tag)).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre articulo
End Sub

Private Sub frmR_DatoSeleccionado(CadenaSeleccion As String)
    'RUTA
    Text1(28).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(28).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmS_DatoSeleccionado(CadenaSeleccion As String)
    'SITUACION
    Text1(29).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(29).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
    'TARIFA
    If Not IsNumeric(Me.imgBuscar(1).Tag) Then Exit Sub
    
    If CInt(Me.imgBuscar(1).Tag) = 1 Then
        Text1(1).Text = RecuperaValor(CadenaSeleccion, 1)
        Text2(1).Text = RecuperaValor(CadenaSeleccion, 2)
    Else
        Text1(30).Text = RecuperaValor(CadenaSeleccion, 1)
        Text2(30).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmZ_DatoSeleccionado(CadenaSeleccion As String)
    'ZONA
    Text1(27).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(27).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim i As Integer
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 15, 32, 53 'cod. articulo
            Me.imgBuscar(1).Tag = Index
            Set frmMtoArt = New frmAlmArticulos
            frmMtoArt.DatosADevolverBusqueda2 = "@1@"
            frmMtoArt.Show vbModal
            Set frmMtoArt = Nothing
            
        Case 24 'forma de pago
            If Modo = 4 Then TerminaBloquear
            Set frmFP = New frmFacFormasPago
            frmFP.DatosADevolverBusqueda = "0"
            frmFP.Show vbModal
            Set frmFP = Nothing
            If Modo = 4 Then
                If Not BLOQUEADesdeFormulario(Me) Then cmdCancelar_Click
            End If
    
        Case 25 'Codigo Actividad
            Set frmA = New frmFacActividades
            frmA.DatosADevolverBusqueda = "0"
            If Not IsNumeric(Text1(Index)) Then Text1(Index).Text = ""
            frmA.Show vbModal
            Set frmA = Nothing
            
        Case 26  'Cod. Envio
            Set frmE = New frmFacFormasEnvio
            frmE.DatosADevolverBusqueda = "0"
            If Not IsNumeric(Text1(Index)) Then Text1(Index).Text = ""
            frmE.Show vbModal
            Set frmE = Nothing
            
        Case 27  'Cod. Zona
            Set frmZ = New frmFacZonas
            frmZ.DatosADevolverBusqueda = "0"
            If Not IsNumeric(Text1(Index)) Then Text1(Index).Text = ""
            frmZ.Show vbModal
            Set frmZ = Nothing
            
        Case 28  'Cod. Ruta
            Set frmR = New frmFacRutas
            frmR.DatosADevolverBusqueda = "0"
            If Not IsNumeric(Text1(Index)) Then Text1(Index).Text = ""
            frmR.Show vbModal
            Set frmR = Nothing
            
        Case 4  'Cod. Forma de Pago
            Set frmFP = New frmFacFormasPago
            frmFP.DatosADevolverBusqueda = "0"
            If Not IsNumeric(Text1(Index)) Then Text1(Index).Text = ""
            frmFP.Show vbModal
            Set frmFP = Nothing
            
            
        Case 31 'Código de Agente
            Set frmAC = New frmFacAgentesCom
            frmAC.DatosADevolverBusqueda = "0"
            If Not IsNumeric(Text1(Index)) Then Text1(Index).Text = ""
            frmAC.Show vbModal
            Set frmAC = Nothing
            
        Case 1, 30 'Código de Tarifa
            Me.imgBuscar(1).Tag = Index
            Set frmT = New frmFacTarifas
            frmT.DatosADevolverBusqueda = "0"
            If Not IsNumeric(Text1(Index)) Then Text1(Index).Text = ""
            frmT.Show vbModal
            Set frmT = Nothing
            
        Case 29 'Código de Situación
            Set frmS = New frmFacSituaciones
            frmS.DatosADevolverBusqueda = "0"
             If Not IsNumeric(Text1(Index)) Then Text1(Index).Text = ""
            frmS.Show vbModal
            Set frmS = Nothing
            
            
        Case 33 To 42, 45, 47 To 50 'Todos los ivas y la Cta de retencion, y cuenta aportacion TERMINAL
            CadenaDesdeOtroForm = ""
                        
            BuscaBuscaGRid2 (Index <> 40 And Index <> 42)
            If CadenaDesdeOtroForm <> "" Then
                If Index = 42 Then
                    i = 9 'Para la cta aportacion
                Else
                    i = 7
                End If
                Text1(Index + i).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
                Text2(Index + i).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
            End If
        
            
    End Select
    PonerFoco Text1(Index)
End Sub


Private Sub BuscaBuscaGRid2(EsIVa As Boolean)


    Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        If EsIVa Then
            'Busco IVAS
            frmB.vCampos = "Código|tiposiva|codigiva|N||20·Denominacion|tiposiva|nombriva|T||70·"
            frmB.vTabla = "tiposiva"
            frmB.vTitulo = "Tipos de IVA"
        Else
                
            frmB.vCampos = "Código|cuentas|codmacta|T||20·Denominacion|cuentas|nommacta|T||70·"
            frmB.vTabla = "cuentas"
            frmB.vTitulo = "Cta contable"
            frmB.vSQL = "apudirec = 'S'"
        
        End If
        frmB.vDevuelve = "0|1|"
        frmB.vselElem = 1
        frmB.vConexionGrid = conConta

        frmB.vCargaFrame = False
      
        frmB.Show vbModal
        Set frmB = Nothing


    Screen.MousePointer = vbDefault

End Sub


Private Sub mnModificar_Click()
    If BLOQUEADesdeFormulario(Me) Then BotonModificar
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub SSTab1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress (KeyAscii)
End Sub


Private Sub Text1_LostFocus(Index As Integer)
'    If Text1(Index).Text = "" Then Exit Sub

    'Quitar espacios en blanco
    Text1(Index).Text = Trim(Text1(Index).Text)
    
    Select Case Index
        Case 1 'tarifa de PVP
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "starif", "nomlista", "codlista", , "N")
            Else
                Text2(Index).Text = ""
            End If
        
        Case 2 'Km desplaz clientes
            PonerFormatoDecimal Text1(Index), 5 'Tipo 4: Decimal(8,4)
        Case 3 'Km desplaz tecnicos
            PonerFormatoDecimal Text1(Index), 5 'Tipo 4: Decimal(8,4)
            
'        Case 6, 7 'Dias Reparacion con/sin mantenimiento
'            If Not EsNumerico(Text1(Index).Text) Then
'                Text1(Index).Text = ""
'                PonerFoco Text1(Index)
'            End If
        Case 14
            'PonerFocoBtn Me.cmdAceptar
            
        Case 15, 32, 53 'cod. artic
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sartic", "nomartic", "codartic", "Articulo")
            If Text2(Index).Text = "" Then Text1(Index).Text = ""
        Case 22 'nº conta
            'PonerFocoBtn Me.cmdAceptar
            
        Case 24 'FORMA DE PAGO
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sforpa", "nomforpa")
            Else
                Text2(Index).Text = ""
            End If
            'PonerFocoBtn Me.cmdAceptar
            
            
        Case 25 To 31
            'Campos por defecto
            If Text1(Index).Text = "" Then
                Text2(Index).Text = ""
            Else
                Select Case Index
                Case 25
                    Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sactiv", "nomactiv", "codactiv")
                Case 26
                    Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "senvio", "nomenvio", "codenvio")
                Case 27
                    Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "szonas", "nomzonas", "Codzonas")
                Case 28
                    Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "srutas", "nomrutas", "codrutas")
                Case 29
                    Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "ssitua", "nomsitua", "codsitua")
                Case 30
                    Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "starif", "nomlista", "codlista")
                Case 31
                    Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sagent", "nomagent", "codagent")
                End Select
            End If
            
        Case 40 To 46, 48, 52, 54 To 57
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conConta, "tiposiva", "nombriva", "codigiva")
                If Text2(Index).Text = "" And Modo > 2 Then PonerFoco Text1(Index)
            Else
                Text2(Index).Text = ""
                
            End If
        Case 47, 51
            'Cta retencion y Cta aportacion al terminal
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conConta, "cuentas", "nommacta", "codmacta")
            Else
                Text2(Index).Text = ""
            End If
        Case 49
            'pORCE RETENCION
            PonerFormatoDecimal Text1(48), 4
        Case 50
            PonerFormatoEntero Text1(Index)
            
        Case 53
            PonerFormatoDecimal Text1(53), 3
    End Select
End Sub


Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 6, 7, 16, 17, 18
            If Text1(Index).Text <> "" Then
                If Not EsNumerico(Text1(Index).Text) Then
                    Cancel = True
                    ConseguirFoco Text1(Index), Modo
                End If
            End If
    End Select
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
'        Case 1  'Anyadir
'            BotonAnyadir
        Case 1  'Modificar
            mnModificar_Click
        Case 4 'Salir
            mnSalir_Click
    End Select
End Sub


'Private Sub BotonAnyadir()
'    LimpiarCampos
'    PonerModo 3
'
'    Text1(0).Text = SugerirCodigoSiguienteStr("scryst", "codcryst")
'    PonerFoco Text1(0)
'End Sub


Private Sub BotonModificar()
    PonerModo 4
    
    Select Case Me.SSTab1.Tab
        Case 0:    PonerFoco Text1(4)
        Case 1: PonerFoco Text1(15)
        Case 2: PonerFoco Text1(8)
        Case 3: PonerFoco Text1(23)
    End Select
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean

    On Error GoTo ErrOK

    If Me.Text1(7).Text = "" Then Me.Text1(7).Text = "0"
    If Me.Text1(6).Text = "" Then Me.Text1(6).Text = "0"

    DatosOk = False
    b = CompForm(Me, 1)
    
    '--- forma de pago de CHEQUE regalo
    'comprobar q el tipo de la forma de pago es EFECTIVO
    If b And Text1(24).Text <> "" Then
        If DevuelveDesdeBDNew(conAri, "sforpa", "tipforpa", "codforpa", Text1(24).Text, "N") <> "0" Then
            MsgBox "La forma de pago del cheque debe ser del tipo EFECTIVO", vbExclamation
            b = False
        End If
    End If
    
    If Text1(47).Text = "" Xor Text1(49).Text = "" Then
        MsgBox "Cta retención o % retención vacios", vbExclamation
        Exit Function
    End If
    
    DatosOk = b
    Exit Function
    
ErrOK:
    MuestraError Err.Number, "Comprobar datos", Err.Description
End Function


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerBotonCabecera(b As Boolean)
    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    Me.cmdSalir.visible = b
    If b Then Me.lblIndicador.Caption = ""
End Sub


Private Sub PonerCampos()
On Error GoTo EPonerCampos

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    
    
    
    'poner descripcion del articulo
    Text2(15).Text = PonerNombreDeCod(Text1(15), conAri, "sartic", "nomartic", "codartic", "Artículos")
    Text2(32).Text = PonerNombreDeCod(Text1(32), conAri, "sartic", "nomartic", "codartic", "Artículos")
    
    'poner descripcion de la forma de pago
    Text2(24).Text = PonerNombreDeCod(Text1(24), conAri, "sforpa", "nomforpa", "codforpa")
    
    'poner descripcion de la tarifa de PVP
    Text2(1).Text = PonerNombreDeCod(Text1(1), conAri, "starif", "nomlista", "codlista", , "N")
    
    
    
    For NumRegElim = 25 To 57
        If NumRegElim < 49 Or NumRegElim > 50 Then
            If Text1(NumRegElim).Text <> "" Then Text1_LostFocus CInt(NumRegElim)
        End If
    Next NumRegElim
    NumRegElim = 0
    
    
    
    BloquearChecks Me, Modo
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
End Sub


'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim b As Boolean
   
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
      
    '------------------------------------------------------
    'Modo insertar o modificar
    b = (Kmodo >= 3) '-->Luego not b sera kmodo<3
    If cmdCancelar.visible Then
        cmdCancelar.Cancel = True
    Else
        cmdCancelar.Cancel = False
    End If
    PonerBotonCabecera Not b
       
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1 y bloquea la clave primaria
    BloquearText1 Me, Modo
    
    'Bloquear los checkbox
    BloquearChecks Me, Modo

    'Bloquear el combobox
    b = Modo = 4
    Me.cboTipodtos.Enabled = b
    Me.cboOrdenDtos.Enabled = b
    Me.cboObsFactura.Enabled = b
    
    'Bloquear imagen de Busqueda

    Dim img As Image
    For Each img In Me.imgBuscar
        BloquearImg img, Not b
    Next
'    BloquearImg Me.imgBuscar(1), (Modo <> 4)
'    BloquearImg Me.imgBuscar(15), (Modo <> 4)
'    For NumRegElim = 24 To 42
'        BloquearImg Me.imgBuscar(NumRegElim), (Modo <> 4)
'    Next NumRegElim
'    NumRegElim = 0
    
    PonerModoOpcionesMenu 'Activar opciones de menu según el Modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub




Private Sub PonerModoOpcionesMenu()
Dim b As Boolean
    b = (Modo = 3) Or (Modo = 4)
    Me.Toolbar1.Buttons(1).Enabled = Not b 'Modificar
    Me.mnModificar.Enabled = Not b
'    Me.Toolbar1.Buttons(2).Enabled = (Not b) 'Modificar
End Sub


Private Sub CargarComboTipoDtos()
'### Combo Tipo Descuento
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Aditivo, 1-sobre Resto

    cboTipodtos.Clear
    cboTipodtos.AddItem "Aditivo"
    cboTipodtos.ItemData(cboTipodtos.NewIndex) = 0
    
    cboTipodtos.AddItem "sobre Resto"
    cboTipodtos.ItemData(cboTipodtos.NewIndex) = 1
End Sub


Private Sub CargarComboOrdenDtos()
'### Combo Tipo Descuento
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Aditivo, 1-sobre Resto

    Me.cboOrdenDtos.Clear
    Me.cboOrdenDtos.AddItem "Familia/Marca"
    cboOrdenDtos.ItemData(cboOrdenDtos.NewIndex) = 0
    
    cboOrdenDtos.AddItem "Marca/Familia"
    cboOrdenDtos.ItemData(cboOrdenDtos.NewIndex) = 1
End Sub

Private Sub CargaComoboObsFactura()
'## Cuando contabilice, que valor pondra en el campo observaciones del
'   la factura, tanto cliente como de proveedores

    Me.cboObsFactura.Clear
    Me.cboObsFactura.AddItem "Sin observaciones"
    cboObsFactura.ItemData(cboObsFactura.NewIndex) = 0
    
    cboObsFactura.AddItem "Número factura"
    cboObsFactura.ItemData(cboObsFactura.NewIndex) = 1

    cboObsFactura.AddItem "Fecha integración"
    cboObsFactura.ItemData(cboObsFactura.NewIndex) = 2

End Sub
