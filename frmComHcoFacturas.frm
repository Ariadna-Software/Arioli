VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmComHcoFacturas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hist�rico de Facturas Proveedores"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   4335
   ClientWidth     =   11745
   Icon            =   "frmComHcoFacturas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmComHcoFacturas.frx":000C
   ScaleHeight     =   6495
   ScaleWidth      =   11745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   17
      Left            =   7545
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   91
      Text            =   "ABCDKFJADKSFJAK"
      Top             =   6060
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Frame Frame2 
      Height          =   710
      Left            =   120
      TabIndex        =   97
      Top             =   385
      Width           =   11415
      Begin VB.CheckBox Check2 
         Caption         =   "Pre"
         Enabled         =   0   'False
         Height          =   255
         Left            =   5280
         TabIndex        =   111
         Tag             =   "presupuesto|N|N|0|1|scafpc|presupuesto||N|"
         Top             =   300
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   31
         Left            =   2640
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha Recepci�n|F|N|||scafpc|fecrecep|dd/mm/yyyy|N|"
         Top             =   315
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   3
         Left            =   8010
         MaxLength       =   40
         TabIndex        =   5
         Tag             =   "Nombre Proveedor|T|N|||scafpc|nomprove||N|"
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwww aq"
         Top             =   240
         Width           =   3150
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   7200
         MaxLength       =   6
         TabIndex        =   4
         Tag             =   "Cod. Proveedor|N|N|0|999999|scafpc|codprove|000000|S|"
         Text            =   "Text1"
         Top             =   240
         Width           =   780
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   1470
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Factura|F|N|||scafpc|fecfactu|dd/mm/yyyy|S|"
         Top             =   315
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   0
         Left            =   240
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "N� Factura|T|N|||scafpc|numfactu||S|"
         Text            =   "Text1 7"
         Top             =   315
         Width           =   1125
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Contabilizado"
         Height          =   255
         Left            =   3840
         TabIndex        =   3
         Tag             =   "Contabilizado|N|N|0|1|scafpc|intconta||N|"
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "F. Recepci�n"
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   104
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Proveedor"
         Height          =   255
         Index           =   0
         Left            =   6120
         TabIndex        =   100
         Top             =   300
         Width           =   735
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   6960
         Picture         =   "frmComHcoFacturas.frx":0A0E
         ToolTipText     =   "Buscar proveedor"
         Top             =   270
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "F. Factura"
         Height          =   255
         Index           =   29
         Left            =   1470
         TabIndex        =   99
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "N� Factura"
         Height          =   255
         Index           =   28
         Left            =   240
         TabIndex        =   98
         Top             =   120
         Width           =   855
      End
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   7320
      Top             =   3720
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
      Left            =   7560
      Top             =   3720
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
      Height          =   4800
      Left            =   120
      TabIndex        =   24
      Tag             =   "Fecha Oferta|F|N|||scapre|fecentre|dd/mm/yyyy|N|"
      Top             =   1090
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   8467
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
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
      TabPicture(0)   =   "frmComHcoFacturas.frx":0B10
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FrameFactura"
      Tab(0).Control(1)=   "FrameCliente"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Albaranes"
      TabPicture(1)   =   "frmComHcoFacturas.frx":0B2C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "imgBuscar(6)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(9)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "imgBuscar(5)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1(21)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label1(6)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label1(18)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "FrameObserva"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "DataGrid2"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "DataGrid1"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtAux(7)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtAux(6)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtAux(5)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtAux(4)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Text3(1)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Text2(1)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Text3(0)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Text2(0)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Text3(2)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Text3(3)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "cmdObserva"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "txtAux(0)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "txtAux(1)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "txtAux(2)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "txtAux(3)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "txtAux3(0)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "txtAux3(1)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).ControlCount=   26
      Begin VB.TextBox txtAux3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   2040
         MaxLength       =   30
         TabIndex        =   96
         Tag             =   "Fecha Albaran|F|N|||scafpa|fechaalb|dd/mm/yyyy|N|"
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
         Index           =   0
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   95
         Tag             =   "N� Albaran|N|N|||scafpa|numalbar|0|N|"
         Text            =   "numalbar"
         Top             =   2160
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   3
         Left            =   3240
         MaxLength       =   12
         TabIndex        =   86
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
         Left            =   2160
         MaxLength       =   12
         TabIndex        =   85
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
         Left            =   1320
         MaxLength       =   12
         TabIndex        =   84
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
         Left            =   360
         MaxLength       =   12
         TabIndex        =   83
         Tag             =   "Almacen|N|N|0|999|slifac|codalmac|000|N|"
         Text            =   "almacen"
         Top             =   4320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdObserva 
         Height          =   375
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   520
         Width           =   375
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   6015
         MaxLength       =   10
         TabIndex        =   73
         Tag             =   "Fecha Pedido|F|S|||scafpc|fecpedpr|dd/mm/yyyy|N|"
         Top             =   1875
         Width           =   1185
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   5040
         MaxLength       =   7
         TabIndex        =   72
         Tag             =   "N� Pedido|N|S|||scafpa|numpedpr|0000000|N|"
         Text            =   "Text1 7"
         Top             =   1875
         Width           =   885
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   0
         Left            =   7620
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   71
         Text            =   "Text2"
         Top             =   720
         Width           =   3525
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   6960
         MaxLength       =   30
         TabIndex        =   17
         Tag             =   "Trabajador Albaran|N|S|0|9999|scafpa|codtrab2|0000|N|"
         Text            =   "Text1"
         Top             =   720
         Width           =   660
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   1
         Left            =   7620
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   70
         Text            =   "Text2"
         Top             =   1080
         Width           =   3525
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   6960
         MaxLength       =   30
         TabIndex        =   18
         Tag             =   "Trabajador pedido|N|S|0|9999|scafpa|codtrab1|0000|N|"
         Text            =   "Text1"
         Top             =   1080
         Width           =   660
      End
      Begin VB.Frame FrameFactura 
         Height          =   2220
         Left            =   -74760
         TabIndex        =   39
         Top             =   2430
         Width           =   11055
         Begin VB.Frame FrmRetencionSocios 
            Height          =   855
            Left            =   240
            TabIndex        =   106
            Top             =   1200
            Width           =   3615
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   33
               Left            =   1920
               MaxLength       =   15
               TabIndex        =   110
               Tag             =   "Imp Ret|N|S|||scafpc|impret|#,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   360
               Width           =   1485
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   32
               Left            =   960
               MaxLength       =   15
               TabIndex        =   107
               Tag             =   "% Ret|N|S|||scafpc|porret|#0.00|N|"
               Text            =   "Text1 7"
               Top             =   360
               Width           =   525
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "%"
               Height          =   195
               Index           =   5
               Left            =   1560
               TabIndex        =   109
               Top             =   360
               Width           =   120
            End
            Begin VB.Label Label1 
               Caption         =   "Retenci�n"
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   108
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0FF&
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
            Index           =   30
            Left            =   9360
            MaxLength       =   15
            TabIndex        =   64
            Tag             =   "Total Factura|N|N|||scafpc|totalfac|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1680
            Width           =   1605
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   29
            Left            =   7560
            MaxLength       =   15
            TabIndex        =   59
            Tag             =   "Importe IVA 3|N|S|||scafpc|impoiva3|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1725
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   23
            Left            =   5040
            MaxLength       =   5
            TabIndex        =   58
            Tag             =   "% IVA 3|N|S|0|99.90|scafpc|porciva3|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1725
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   20
            Left            =   4320
            MaxLength       =   3
            TabIndex        =   57
            Tag             =   "Cod. IVA 3|N|S|0|999|scafpc|tipoiva3|000|N|"
            Text            =   "Text1 7"
            Top             =   1725
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   26
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   56
            Tag             =   "Base Imponible 3|N|S|||scafpc|baseiva3|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1725
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   28
            Left            =   7560
            MaxLength       =   15
            TabIndex        =   55
            Tag             =   "Importe IVA 2|N|S|||scafpc|impoiva2|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1395
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   22
            Left            =   5040
            MaxLength       =   5
            TabIndex        =   54
            Tag             =   "& IVA 2|N|S|0|99.90|scafpc|porciva2|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1395
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   19
            Left            =   4320
            MaxLength       =   3
            TabIndex        =   53
            Tag             =   "Cod. IVA 2|N|S|0|999|scafpc|tipoiva2|000|N|"
            Text            =   "Text1 7"
            Top             =   1395
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   25
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   52
            Tag             =   "Base Imponible 2 |N|S|||scafpc|baseiva2|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1395
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   27
            Left            =   7560
            MaxLength       =   15
            TabIndex        =   51
            Tag             =   "Importe IVA 1|N|N|||scafpc|impoiva1|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1080
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   21
            Left            =   5040
            MaxLength       =   5
            TabIndex        =   50
            Tag             =   "% IVA 1|N|S|0|99.90|scafpc|porciva1|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1080
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   18
            Left            =   4320
            MaxLength       =   3
            TabIndex        =   49
            Tag             =   "Cod. IVA 1|N|S|0|999|scafpc|tipoiva1|000|N|"
            Text            =   "Text1 7"
            Top             =   1080
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   24
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   48
            Tag             =   "Base Imponible 1|N|N|||scafpc|baseiva1|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1080
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   17
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   43
            Text            =   "Text1 7"
            Top             =   435
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   16
            Left            =   3960
            MaxLength       =   15
            TabIndex        =   42
            Tag             =   "Imp. Dto Gn|N|N|||scafpc|impgnral|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   435
            Width           =   1365
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   15
            Left            =   2160
            MaxLength       =   15
            TabIndex        =   41
            Tag             =   "Imp. Dto PP|N|N|||scafpc|impppago|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   435
            Width           =   1365
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   14
            Left            =   240
            MaxLength       =   15
            TabIndex        =   40
            Tag             =   "Imp.Bruto|N|N|||scafpc|brutofac|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   435
            Width           =   1485
         End
         Begin VB.Label Label1 
            Caption         =   "Cod. IVA"
            Height          =   255
            Index           =   42
            Left            =   4320
            TabIndex        =   94
            Top             =   870
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "% IVA"
            Height          =   255
            Index           =   41
            Left            =   5040
            TabIndex        =   93
            Top             =   870
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
            ForeColor       =   &H000000C0&
            Height          =   255
            Index           =   39
            Left            =   9330
            TabIndex        =   68
            Top             =   1440
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
            Left            =   9120
            TabIndex        =   67
            Top             =   1680
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
            TabIndex        =   66
            Top             =   2160
            Width           =   135
         End
         Begin VB.Line Line1 
            X1              =   4320
            X2              =   7320
            Y1              =   825
            Y2              =   825
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
            Index           =   37
            Left            =   7320
            TabIndex        =   65
            Top             =   960
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. IVA"
            Height          =   255
            Index           =   33
            Left            =   7680
            TabIndex        =   63
            Top             =   870
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
            TabIndex        =   62
            Top             =   360
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
            TabIndex        =   61
            Top             =   360
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
            TabIndex        =   60
            Top             =   360
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "Base Imponible"
            Height          =   255
            Index           =   14
            Left            =   5880
            TabIndex        =   47
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Dto Gn"
            Height          =   255
            Index           =   12
            Left            =   4080
            TabIndex        =   46
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Dto PP"
            Height          =   255
            Index           =   11
            Left            =   2280
            TabIndex        =   45
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Bruto"
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   44
            Top             =   240
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
         Left            =   4080
         MaxLength       =   12
         TabIndex        =   87
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
         Index           =   5
         Left            =   5760
         MaxLength       =   5
         TabIndex        =   88
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
         Index           =   6
         Left            =   6480
         MaxLength       =   30
         TabIndex        =   89
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
         Index           =   7
         Left            =   7080
         MaxLength       =   12
         TabIndex        =   92
         Tag             =   "Importe|N|N|0||slifac|importel|#,###,###,##0.00|N|"
         Text            =   "Importe"
         Top             =   4320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame FrameCliente 
         Caption         =   "Datos Proveedor"
         ForeColor       =   &H00972E0B&
         Height          =   1880
         Left            =   -74760
         TabIndex        =   26
         Top             =   495
         Width           =   11055
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   13
            Left            =   7530
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   102
            Text            =   "Text2"
            Top             =   240
            Width           =   3285
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   13
            Left            =   6945
            MaxLength       =   4
            TabIndex        =   101
            Tag             =   "Trabajador|N|N|0|9999|scafpc|codtraba|0000|N|"
            Text            =   "Text1"
            Top             =   240
            Width           =   540
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   9
            Left            =   1125
            MaxLength       =   30
            TabIndex        =   11
            Tag             =   "Provincia|T|N|||scafpc|proprove||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text22"
            Top             =   1350
            Width           =   2445
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   7
            Left            =   1125
            MaxLength       =   6
            TabIndex        =   9
            Tag             =   "CPostal|T|N|||scafpc|codpobla||N|"
            Text            =   "Text15"
            Top             =   960
            Width           =   630
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   8
            Left            =   1755
            MaxLength       =   30
            TabIndex        =   10
            Tag             =   "Poblaci�n|T|N|||scafpc|pobprove||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
            Top             =   990
            Width           =   3405
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   5
            Left            =   3195
            MaxLength       =   20
            TabIndex        =   7
            Tag             =   "tel�fono proveedor|T|S|||scafpc|telprove||N|"
            Text            =   "12345678911234567899"
            Top             =   285
            Width           =   1965
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   4
            Left            =   1125
            MaxLength       =   15
            TabIndex        =   6
            Tag             =   "NIF proveedor|T|N|||scafpc|nifprove||N|"
            Text            =   "123456789"
            Top             =   285
            Width           =   1110
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   10
            Left            =   6945
            MaxLength       =   3
            TabIndex        =   12
            Tag             =   "Forma de Pago|N|N|0|999|scafpc|codforpa|000|N|"
            Text            =   "Text1"
            Top             =   645
            Width           =   540
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   10
            Left            =   7530
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   28
            Text            =   "Text2"
            Top             =   645
            Width           =   3285
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   11
            Left            =   7560
            MaxLength       =   5
            TabIndex        =   13
            Tag             =   "Descuento P.Pago|N|N|0|99.90|scafpc|dtoppago|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1230
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   12
            Left            =   8925
            MaxLength       =   5
            TabIndex        =   14
            Tag             =   "Descuento General|N|N|0|99.90|scafpc|dtognral|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1230
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   6
            Left            =   1125
            MaxLength       =   35
            TabIndex        =   8
            Tag             =   "Domicilio|T|N|||scafpc|domprove||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aq"
            Top             =   645
            Width           =   4030
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   6660
            Picture         =   "frmComHcoFacturas.frx":0B48
            ToolTipText     =   "Buscar trabajador"
            Top             =   270
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Trabajador"
            Height          =   255
            Index           =   1
            Left            =   5730
            TabIndex        =   103
            Top             =   240
            Width           =   855
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   2
            Left            =   855
            Picture         =   "frmComHcoFacturas.frx":0C4A
            ToolTipText     =   "Buscar poblaci�n"
            Top             =   1005
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Provincia"
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   35
            Top             =   1350
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Poblaci�n"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   34
            Top             =   990
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Tel�fono"
            Height          =   255
            Index           =   19
            Left            =   2445
            TabIndex        =   33
            Top             =   285
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "NIF"
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   32
            Top             =   285
            Width           =   615
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   855
            Picture         =   "frmComHcoFacturas.frx":0D4C
            ToolTipText     =   "Buscar proveedor varios"
            Top             =   300
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Forma Pago"
            Height          =   255
            Index           =   15
            Left            =   5730
            TabIndex        =   31
            Top             =   645
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. P.P"
            Height          =   255
            Index           =   25
            Left            =   6900
            TabIndex        =   30
            Top             =   1230
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. Gral"
            Height          =   255
            Index           =   26
            Left            =   8235
            TabIndex        =   29
            Top             =   1230
            Width           =   735
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   6660
            Picture         =   "frmComHcoFacturas.frx":0E4E
            ToolTipText     =   "Buscar forma de pago"
            Top             =   675
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Domicilio"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   27
            Top             =   645
            Width           =   735
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmComHcoFacturas.frx":0F50
         Height          =   2030
         Left            =   240
         TabIndex        =   38
         Top             =   2620
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   3572
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
         Bindings        =   "frmComHcoFacturas.frx":0F65
         Height          =   1995
         Left            =   240
         TabIndex        =   69
         Top             =   520
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   3519
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
      Begin VB.Frame FrameObserva 
         Caption         =   "Observaciones"
         ForeColor       =   &H00972E0B&
         Height          =   2055
         Left            =   240
         TabIndex        =   76
         Tag             =   "Observaci�n 4|T|S|||scafac1|observa4||N|"
         Top             =   2520
         Width           =   10575
         Begin VB.TextBox Text3 
            Height          =   300
            Index           =   8
            Left            =   720
            MaxLength       =   80
            TabIndex        =   81
            Tag             =   "Observaci�n 5|T|S|||scafpa|observa5||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   1560
            Width           =   8940
         End
         Begin VB.TextBox Text3 
            Height          =   300
            Index           =   7
            Left            =   720
            MaxLength       =   80
            TabIndex        =   80
            Tag             =   "Observaci�n 4|T|S|||scafpa|observa4||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   1230
            Width           =   8940
         End
         Begin VB.TextBox Text3 
            Height          =   300
            Index           =   6
            Left            =   720
            MaxLength       =   80
            TabIndex        =   79
            Tag             =   "Observaci�n 3|T|S|||scafpa|observa3||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   900
            Width           =   8940
         End
         Begin VB.TextBox Text3 
            Height          =   300
            Index           =   5
            Left            =   720
            MaxLength       =   80
            TabIndex        =   78
            Tag             =   "Observaci�n 2|T|S|||scafpa|observa2||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   570
            Width           =   8940
         End
         Begin VB.TextBox Text3 
            Height          =   300
            Index           =   4
            Left            =   720
            MaxLength       =   80
            TabIndex        =   77
            Tag             =   "Observaci�n 1|T|S|||scafpa|observa1||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   240
            Width           =   8940
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Pedido"
         Height          =   255
         Index           =   18
         Left            =   6015
         TabIndex        =   75
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "N� Pedido"
         Height          =   255
         Index           =   6
         Left            =   5040
         TabIndex        =   74
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Trabajador Albaran"
         Height          =   255
         Index           =   21
         Left            =   5040
         TabIndex        =   37
         Top             =   765
         Width           =   1455
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   6615
         Picture         =   "frmComHcoFacturas.frx":0F7A
         ToolTipText     =   "Buscar trabajador"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Trabajador Pedido"
         Height          =   255
         Index           =   9
         Left            =   5040
         TabIndex        =   36
         Top             =   1095
         Width           =   1425
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   6600
         Picture         =   "frmComHcoFacturas.frx":107C
         ToolTipText     =   "Buscar trabajador"
         Top             =   1095
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
      TabIndex        =   90
      Text            =   "Text2 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwqa"
      Top             =   6060
      Visible         =   0   'False
      Width           =   5085
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   5895
      Width           =   2175
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   240
         TabIndex        =   21
         Top             =   180
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10530
      TabIndex        =   16
      Top             =   6000
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9360
      TabIndex        =   15
      Top             =   6000
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
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
            Object.Tag             =   "0"
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir Factura"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "�ltimo"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   8400
         TabIndex        =   23
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   10530
      TabIndex        =   19
      Top             =   6000
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
      Caption         =   "N� Lote"
      Height          =   255
      Index           =   3
      Left            =   7545
      TabIndex        =   105
      Top             =   5880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Ampliaci�n L�nea"
      Height          =   255
      Index           =   35
      Left            =   2400
      TabIndex        =   25
      Top             =   5880
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
         Enabled         =   0   'False
         Shortcut        =   ^I
         Visible         =   0   'False
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
Attribute VB_Name = "frmComHcoFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========== VBLES PUBLICAS ====================
Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'Si se llama de la busqueda en el frmAlmMovimArticulos se accede
'a las tablas del Albaran o de Facturas de movimiento seleccionado (solo consulta)
Public hcoCodMovim As String 'cod. movim
Public hcoFechaMovim As Date 'Fecha del Movim
Public hcoCodProve As Integer 'Codigo de Proveedor

'========== VBLES PRIVADAS ====================
Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmCP As frmCPostal 'Codigos Postales
Attribute frmCP.VB_VarHelpID = -1

Private WithEvents frmProv As frmComProveedores  'Form Mto Proveedores
Attribute frmProv.VB_VarHelpID = -1
Private WithEvents frmPV As frmComProveV  'Form Mto Proveedores Varios
Attribute frmPV.VB_VarHelpID = -1
Private WithEvents frmFP As frmFacFormasPago 'Form Mto Formas de Pago
Attribute frmFP.VB_VarHelpID = -1
Private WithEvents frmT As frmAdmTrabajadores  'Form Mto Trabajadores
Attribute frmT.VB_VarHelpID = -1


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
'1.- A�adir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

Dim PrimeraVez As Boolean

'Si el cliente mostrado es de Varios o No
Dim EsDeVarios As Boolean


'SQL de la tabla principal del formulario
Private CadenaConsulta As String

Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private NomTablaLineas As String 'Nombre de la Tabla de lineas
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Dim btnPrimero As Byte
'Variable que indica el n�mero del Boton  PrimerRegistro en la Toolbar1

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

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 4  'MODIFICAR
            If DatosOk Then
               If ModificarFactura Then
                    TerminaBloquear
'                    PosicionarData
               Else
                    '---- Laura 24/10/2006
                    'como no hemos modificado dejamos la fecha como estaba ya que ahora se puede modificar
                    Text1(1).Text = Me.Data1.Recordset!Fecfactu
               End If
               PosicionarData
            End If
            
         Case 5 'InsertarModificar LINEAS
            'Actualizar el registro en la tabla de lineas 'slialb'
            If ModificaLineas = 1 Then 'INSERTAR lineas Albaran
'                PrimeraLin = False
'                If Data2.Recordset.EOF = True Then PrimeraLin = True
'                If InsertarLinea(NumLinea) Then
'                    'Comprobar si el Articulo tiene control de N� de Serie
'                    ComprobarNSeriesLineas NumLinea
'                    If PrimeraLin Then
'                        CargaGrid DataGrid1, Data2, True
'                    Else
'                        CargaGrid2 DataGrid1, Data2
'                    End If
'                    BotonAnyadirLinea
'                End If
            ElseIf ModificaLineas = 2 Then 'MODIFICAR lineas
                If ModificarLinea Then
                    TerminaBloquear
                    NumRegElim = Data2.Recordset.AbsolutePosition
                    
                    CargaGrid2 DataGrid1, Data2
                    ModificaLineas = 0
                    PonerBotonCabecera True
                    BloquearTxt Text2(16), True
                    BloquearTxt Text2(17), True
           
                    LLamaLineas Modo, 0, "DataGrid1"
                    PosicionarData
                    If (Not Data2.Recordset.EOF) And (Not Data2.Recordset.BOF) Then
                        SituarDataPosicion Data2, NumRegElim, ""
                    End If
                End If
                Me.DataGrid1.Enabled = True
                Me.DataGrid2.Enabled = True
            End If
    End Select
    Screen.MousePointer = vbDefault

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
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
            BloquearTxt Text2(17), True
'            If ModificaLineas = 1 Then 'INSERTAR
'                ModificaLineas = 0
'                DataGrid1.AllowAddNew = False
'                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
'            End If
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
        LimpiarCampos
        BuscaChekc = ""
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
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
Dim C As String
    C = " presupuesto ="
    If vUsu.TrabajadorB Then
        C = C & "1"
    Else
        C = C & "0"
    End If
        

    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia C
    Else
        LimpiarCampos

        LimpiarDataGrids
        CadenaConsulta = "Select scafpc.* "
        CadenaConsulta = CadenaConsulta & "from " & NombreTabla
        CadenaConsulta = CadenaConsulta & " WHERE " & C & Ordenacion
        

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
    If FactContabilizada Then
        TerminaBloquear
        Exit Sub
    End If
    
    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    PonerFocoChk Me.Check1
        
    'Si es proveedor de Varios no se pueden modificar sus datos
    DeVarios = EsProveedorVarios(Text1(2).Text)
    BloquearDatosProve (DeVarios)
End Sub


Private Sub BotonModificarLinea()
'Modificar una linea
Dim vWhere As String
Dim anc As Single
Dim J As Byte
On Error GoTo EModificarLinea




    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then
        TerminaBloquear
        Exit Sub '1= Insertar
    End If
    
    If Data2.Recordset.EOF Then
        TerminaBloquear
        Exit Sub
    End If
    
    vWhere = ObtenerWhereCP(False)
    vWhere = vWhere & " AND numalbar='" & data3.Recordset.Fields!NumAlbar & "'"
    vWhere = vWhere & " and numlinea=" & Data2.Recordset!numlinea
    If Not BloqueaRegistro(NomTablaLineas, vWhere) Then
        TerminaBloquear
        Exit Sub
    End If

    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        J = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, J
        DataGrid1.Refresh
    End If
    
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 210
    Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 20
    End If

    For J = 0 To 2
        txtAux(J).Text = DataGrid1.Columns(J + 5).Text
    Next J
    Text2(16).Text = DataGrid1.Columns(J + 5).Text
    For J = J + 1 To 8
        txtAux(J - 1).Text = DataGrid1.Columns(J + 5).Text
    Next J
    Text2(17).Text = DataGrid1.Columns(14).Text
    
    ModificaLineas = 2 'Modificar
    LLamaLineas ModificaLineas, anc, "DataGrid1"
    
    
    
        
    
    'A�adiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR LINEAS"
    PonerBotonCabecera False
    BloquearTxt Text2(16), False 'Campo Ampliacion Linea
    BloquearTxt Text2(17), True 'Campo Ampliacion Linea
'    PonerFoco txtAux(4)
    PonerFoco Text2(16)
    Me.DataGrid1.Enabled = False

EModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub LLamaLineas(xModo As Byte, Optional alto As Single, Optional grid As String)
Dim jj As Integer
Dim b As Boolean
'
'    Select Case grid
'        Case "DataGrid1"
'            DeseleccionaGrid Me.DataGrid1
'            'PonerModo xModo + 1
'
'            b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Lineas
'
        If grid = "DataGrid1" Then
            
            alto = ObtenerAlto(DataGrid1, 15)
            b = (xModo = 2)
            If b Then DeseleccionaGrid Me.DataGrid1
            
            
            
            For jj = 4 To txtAux.Count - 1
                If jj >= 4 And jj <= 7 Then
                    txtAux(jj).Height = DataGrid1.RowHeight
                    txtAux(jj).Top = alto
                    txtAux(jj).visible = b
                    If b Then
                        txtAux(jj).Left = DataGrid1.Columns(jj + 6).Left + 255
                        txtAux(jj).Width = DataGrid1.Columns(jj + 6).Width - 15
                        BloquearTxt txtAux(jj), Data1.Recordset!intconta = 1
                    End If
                    
                End If
            Next jj
            
            
            
        End If
        
        
        If grid = "DataGrid2" Then
            DeseleccionaGrid Me.DataGrid2
            b = (xModo = 1)
             For jj = 0 To txtAux3.Count - 1
                txtAux3(jj).Height = DataGrid2.RowHeight
                txtAux3(jj).Top = alto
                txtAux3(jj).visible = b
            Next jj
        End If
'    End Select
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Facturas (scafac)
' y los registros correspondientes de las tablas cab. albaranes (scafac1)
' y las lineas de la factura (slifac)
Dim cad As String
'Dim NumPedElim As Long
On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    'solo se puede modificar la factura si no esta contabilizada
    If FactContabilizada Then Exit Sub
    
    cad = "Cabecera de Facturas." & vbCrLf
    cad = cad & "-----------------------------------" & vbCrLf & vbCrLf
    cad = cad & "Va a eliminar la Factura:            "
    cad = cad & vbCrLf & "Proveedor:  " & Text1(2).Text & " - " & Text1(3).Text
    cad = cad & vbCrLf & "N� Fact.:  " & Text1(0).Text
    cad = cad & vbCrLf & "Fecha:  " & Format(Text1(1).Text, "dd/mm/yyyy")

    cad = cad & vbCrLf & vbCrLf & " �Desea Eliminarla? "
    
    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
       ' NumPedElim = Data1.Recordset.Fields(1).Value   MAAAAl. Ya que el n�fac prov es ALFANUMERICO
        
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
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Albaran", Err.Description
End Sub


Private Sub cmdObserva_Click()
    If Modo <> 2 And Modo <> 4 Then Exit Sub
    If Me.FrameObserva.visible = False Then
        Me.DataGrid1.visible = False
        Me.FrameObserva.visible = True
        Me.cmdObserva.Picture = frmppal.imgListComun.ListImages(18).Picture
'        CargarICO Me.cmdObserva, "volver.ico"
        Me.cmdObserva.ToolTipText = "volver lineas albaran"
        BloqueaText3
    Else
        Me.DataGrid1.visible = True
        Me.FrameObserva.visible = False
'        CargarICO Me.cmdObserva, "message.ico"
        Me.cmdObserva.Picture = frmppal.imgListComun.ListImages(41).Picture
        Me.cmdObserva.ToolTipText = "ver observaciones albaran"
    End If
End Sub


Private Sub BloqueaText3()
Dim I As Byte
    'bloquear los Text3 que son las lineas de scafpa
    For I = 0 To 1
        BloquearTxt Text3(I), (Modo <> 4)
    Next I
    If Me.FrameObserva.visible Then
        For I = 4 To 8
            BloquearTxt Text3(I), (Modo <> 4)
        Next I
    End If
    'numpedpr, fecpedpr siempre bloqueados
    For I = 2 To 3
        BloquearTxt Text3(I), True
    Next I
End Sub


Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim cad As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
        PonerModo 2
        DataGrid2.Enabled = True
        If Not Data1.Recordset.EOF Then _
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

    If Not Data2.Recordset.EOF Then
        If ModificaLineas <> 1 Then
            Text2(16).Text = DBLet(Data2.Recordset.Fields!ampliaci)
            Text2(17).Text = DBLet(Data2.Recordset.Fields!numlotes)
        End If
    Else
        Text2(16).Text = ""
        Text2(17).Text = ""
    End If

Error1:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Sub DataGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim I As Byte

    If Not data3.Recordset.EOF Then
        Text3(0).Text = DBLet(data3.Recordset.Fields!codtrab2, "T")
        Text3_LostFocus (0)
        Text3(1).Text = DBLet(data3.Recordset.Fields!codtrab1, "T")
        Text3_LostFocus (1)

        Text3(2).Text = DBLet(data3.Recordset.Fields!numpedpr, "N")
        If Text3(2).Text <> "0" Then
            FormateaCampo Text3(2)
        Else
            Text3(2).Text = ""
        End If
        Text3(3).Text = DBLet(data3.Recordset.Fields!fecpedpr, "F")
        
        'Observaciones
        Text3(4).Text = DBLet(data3.Recordset.Fields!observa1, "T")
        Text3(5).Text = DBLet(data3.Recordset.Fields!observa2, "T")
        Text3(6).Text = DBLet(data3.Recordset.Fields!observa3, "T")
        Text3(7).Text = DBLet(data3.Recordset.Fields!observa4, "T")
        Text3(8).Text = DBLet(data3.Recordset.Fields!observa5, "T")
        
        'Datos de la tabla slipre
        CargaGrid DataGrid1, Data2, True
    Else
        For I = 0 To Text3.Count - 1
            Text3(I).Text = ""
        Next I
        Text2(0).Text = ""
        Text2(1).Text = ""
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

    ' ICONITOS DE LA BARRA
    btnPrimero = 15
    With Me.Toolbar1
        .ImageList = frmppal.imgListComun
        .Buttons(1).Image = 1   'Bot�n Buscar
        .Buttons(2).Image = 2   'Bot�n Todos
        .Buttons(5).Image = 4   'Modificar
        .Buttons(6).Image = 5   'Borrar
        .Buttons(9).Image = 10 'Mto Lineas Ofertas
        .Buttons(10).Image = 16 'Imprimir
        .Buttons(12).Image = 15  'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 '�ltimo
    End With
    Me.SSTab1.Tab = 0
      
      
    'El frame FrmRetencionSocios es el que tendra el % retencion aplicable a los socios
    'Solo sera visible SI (y solo si) en paretros.ctaretnecion <>""
    FrmRetencionSocios.visible = vParamAplic.CtaReten <> ""
      
    LimpiarCampos   'Limpia los campos TextBox
     
    'cargar icono de observaciones de los albaranes de factura
'    CargarICO Me.cmdObserva, "message.ico"
    Me.cmdObserva.Picture = frmppal.imgListComun.ListImages(41).Picture
    Me.FrameObserva.visible = False
    Me.cmdObserva.ToolTipText = "ver observaciones albaran"
    
    VieneDeBuscar = False
            
    '## A mano
    NombreTabla = "scafpc"
    NomTablaLineas = "slifpc" 'Tabla lineas de Facturacion
    Ordenacion = " ORDER BY scafpc.fecrecep desc ,scafpc.codprove, scafpc.numfactu "
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    CadenaConsulta = "Select * from " & NombreTabla
    
    If hcoCodMovim <> "" Then
        'Se llama desde Dobleclick en frmAlmMovimArticulos
'        CadenaConsulta = CadenaConsulta & " WHERE numalbar='" & hcoCodMovim & "' AND fechaalb= """ & Format(hcoFechaMovim, "yyyy-mm-dd") & """"
'        CadenaConsulta = CadenaConsulta & " AND codprove=" & hcoCodProve
        CadenaConsulta = CadenaConsulta & ObtenerSelFactura
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
            PonerModo 1
        End If
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
        PrimeraVez = False
    Else
         PonerModo 0
    End If
    
    
    Check2.visible = vUsu.TrabajadorB
    
End Sub


Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.Check1.Value = 0
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
        cadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        cadB = Aux
        Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
        cadB = cadB & " and " & Aux
        Aux = ValorDevueltoFormGrid(Text1(2), CadenaDevuelta, 3)
        cadB = cadB & " and " & Aux
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB
        CadenaConsulta = CadenaConsulta & " GROUP BY scafpc.codprove, scafpc.numfactu, scafpc.fecfactu "
        CadenaConsulta = CadenaConsulta & " " & Ordenacion
        PonerCadenaBusqueda
'        Text1(0).Text = RecuperaValor(CadenaDevuelta, 2)
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento C. Postales
Dim Indice As Byte
Dim devuelve As String

        Indice = 7
        Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
        Text1(Indice + 1).Text = ObtenerPoblacion(Text1(Indice).Text, devuelve)  'Poblacion
        'provincia
        Text1(Indice + 2).Text = devuelve
End Sub

Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Formas de Pago
Dim Indice As Byte
    Indice = 10
    Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Forma Pago
    Text2(10).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Forma Pago
End Sub


Private Sub frmProv_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Proveedores
    Text1(2).Text = RecuperaValor(CadenaSeleccion, 1)  'Cod Prove
End Sub

Private Sub frmPV_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Proveedores Varios
Dim Indice As Byte

    Indice = 4
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'NIF
    Text1(Indice - 1).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Clien
    PonerDatosProveVario (Text1(Indice).Text)
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Trabajadores
Dim Indice As Byte

    Indice = Val(Me.imgBuscar(4).Tag)
    If Indice = 4 Then
        Indice = Indice + 9
        Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000") 'Cod Trabajador
        Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Trabajador
    Else
        Text3(Indice - 5).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000") 'Cod Trabajador
        Text2(Indice - 5).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Trabajador
    End If
End Sub


Private Sub imgBuscar_Click(Index As Integer)
Dim Indice As Byte

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cod. Proveedor
            PonerFoco Text1(2)
            Set frmProv = New frmComProveedores
            frmProv.DatosADevolverBusqueda = "0"
            frmProv.Show vbModal
            Set frmProv = Nothing
            Indice = 2
            PonerFoco Text1(Indice)
            
        Case 1 'NIF para proveedor de Varios
            Set frmPV = New frmComProveV
            frmPV.DatosADevolverBusqueda = "0"
            frmPV.Show vbModal
            Set frmPV = Nothing
            Indice = 7
            PonerFoco Text1(Indice)
            
        Case 2 'Cod. Postal
            Set frmCP = New frmCPostal
            frmCP.DatosADevolverBusqueda = "0"
            frmCP.Show vbModal
            Set frmCP = Nothing
            Indice = 7
            VieneDeBuscar = True
            PonerFoco Text1(Indice)
      
         Case 3 'Forma de Pago
            Indice = 10
            PonerFoco Text1(Indice)
            Set frmFP = New frmFacFormasPago
            frmFP.DatosADevolverBusqueda = "0"
            frmFP.Show vbModal
            Set frmFP = Nothing
            
        Case 4, 5, 6 'Realizada Por Trabajador (Pedido, Albaran, Preparador Material
            Me.imgBuscar(4).Tag = Index
            Set frmT = New frmAdmTrabajadores
            frmT.DatosADevolverBusqueda = "0"
            frmT.Show vbModal
            Set frmT = Nothing
            If Index = 4 Then
                PonerFoco Text1(13)
            Else
                PonerFoco Text3(Index - 5)
            End If
    End Select
    
    Screen.MousePointer = vbDefault
End Sub




Private Sub mnBuscar_Click()
    Me.SSTab1.Tab = 0
    BotonBuscar
End Sub


Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar lineas de Pedido
'         BotonEliminarLinea
    Else   'Eliminar Pedido
         BotonEliminar
    End If
End Sub


Private Sub mnImprimir_Click()
'Imprimir Factura
'    BotonImprimir (53) '53: Informe de Facturas
End Sub


Private Sub mnLineas_Click()
    If Data1.Recordset.EOF Then Exit Sub
    
    'Si son facturas de liquidacion de soccios NO dejamos modificarlas
    If Me.FrmRetencionSocios.visible Then
        If DBLet(Data1.Recordset!PorRet, "N") > 0 Then
            MsgBox "Factura liquidaci�n socios. No puede modificarse", vbExclamation
            Exit Sub
        End If
    End If

    BotonMtoLineas 1, "Facturas"
End Sub


Private Sub mnModificar_Click()
    If Data1.Recordset.EOF Then Exit Sub
    
    'Si son facturas de liquidacion de soccios NO dejamos modificarlas
    If Me.FrmRetencionSocios.visible Then
        If DBLet(Data1.Recordset!PorRet, "N") > 0 Then
            MsgBox "Factura liquidaci�n socios. No puede modificarse", vbExclamation
            Exit Sub
        End If
    End If


    If Modo = 5 Then 'Modificar lineas
        'bloquea la tabla cabecera de factura: scafac
        If BLOQUEADesdeFormulario(Me) Then
            'bloquear la tabla cabecera de albaranes de la factura: scafac1
            If BloqueaAlbxFac Then
                If BloqueaLineasFac Then BotonModificarLinea
            End If
        End If
         
    Else   'Modificar Pedido
        'bloquea la tabla cabecera de factura: scafpc
        If BLOQUEADesdeFormulario(Me) Then
            'bloquear la tabla cabecera de albaranes de la factura: scafpa
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
    SQL = "select * FROM scafpa "
    SQL = SQL & ObtenerWhereCP(True) & " FOR UPDATE"
    conn.Execute SQL, , adCmdText
    BloqueaAlbxFac = True

EBloqueaAlb:
    If Err.Number <> 0 Then BloqueaAlbxFac = False
End Function


Private Function BloqueaLineasFac() As Boolean
'bloquea TODAS las lineas de la factura
Dim SQL As String
    
    On Error GoTo EBloqueaLin

    BloqueaLineasFac = False
    
    'bloquear cabecera albaranes x factura
    SQL = "select * FROM slifpc "
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
    kCampo = Index
    If Index = 9 Then HaCambiadoCP = False 'CPostal
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
Dim devuelve As String
        
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 1, 31 'Fecha factura,fecha recepcion
            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
                
        Case 13 'Cod trabajador
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "straba", "nomtraba", "codtraba")

        Case 2 'Cod. prove
            If Modo = 1 Then 'Modo=1 Busqueda
                Text1(Index + 1).Text = PonerNombreDeCod(Text1(Index), conAri, "sprove", "nomprove")
            Else
                PonerDatosProveedor (Text1(Index).Text)
            End If
        
        Case 4 'NIF
            If Not EsDeVarios Then Exit Sub
            If Modo = 4 Then 'Modificar
                'si no se ha modificado el nif del cliente no hacer nada
                If Text1(4).Text = DBLet(Data1.Recordset!nifProve, "T") Then
                    Exit Sub
                End If
            End If
            PonerDatosProveVario (Text1(Index).Text)
        
        Case 7 'Cod. Postal
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
        
        
        Case 10 'Forma de Pago
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sforpa", "nomforpa")
            Else
                Text2(Index).Text = ""
            End If
            
        Case 11, 12 'Descuentos
            If Modo = 4 Then 'comprobar que el dato a cambiado
                If Index = 11 Then
                    If CCur(Text1(Index).Text) = CCur(Data1.Recordset!DtoPPago) Then Exit Sub
                ElseIf Index = 12 Then
                    If CCur(Text1(Index).Text) = CCur(Data1.Recordset!DtoGnral) Then Exit Sub
                End If
            End If
            
            If Modo = 3 Or Modo = 4 Then
                If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 4 'Tipo 4: Decimal(4,2)
                If Not ActualizarDatosFactura Then
                   If Index = 11 Then Text1(Index).Text = Data1.Recordset!DtoPPago
                   If Index = 12 Then Text1(Index).Text = Data1.Recordset!DtoGnral
                   FormateaCampo Text1(Index)
                End If
            End If
    End Select
End Sub


Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False, BuscaChekc)
    If cadB <> "" Then cadB = cadB & " AND presupuesto = "
    If vUsu.TrabajadorB Then
        cadB = cadB & " 1"
    Else
        cadB = cadB & " 0"
    End If
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select " & NombreTabla & ".* from " & NombreTabla & " LEFT OUTER JOIN scafpa ON " & NombreTabla & ".codprove=scafpa.codprove AND " & NombreTabla & ".numfactu=scafpa.numfactu AND " & NombreTabla & ".fecfactu=scafpa.fecfactu "
        CadenaConsulta = CadenaConsulta & " WHERE " & cadB
        CadenaConsulta = CadenaConsulta & " GROUP BY scafpc.codprove, scafpc.numfactu, scafpc.fecfactu "
        CadenaConsulta = CadenaConsulta & " " & Ordenacion
        
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim Tabla As String
Dim Titulo As String
Dim devuelve As String
    
    'Llamamos a al form
    '##A mano
    cad = ""
'        cad = cad & ParaGrid(Text1(1), 10, "Tipo Fac.")
        cad = cad & ParaGrid(Text1(0), 18, "N� Factura")
        cad = cad & ParaGrid(Text1(1), 15, "Fecha Fac.")
        cad = cad & ParaGrid(Text1(2), 12, "Prov.")
        cad = cad & ParaGrid(Text1(3), 55, "Nombre Prov")
        Tabla = NombreTabla
        Titulo = "Facturas"
        devuelve = "0|1|2|"
           
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
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri  'Conexi�n a BD: Ariges
'        If Not EsCabecera Then frmB.Label1.FontSize = 11
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
        If Modo = 1 Then PonerFoco Text1(0)
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
'Carga el grid de los AlbaranesxFactura, es decir, la tabla scafpc de la factura seleccionada
On Error GoTo EPonerLineas

    Screen.MousePointer = vbHourglass
    
    'Datos de la tabla albaranes x factura: scafpa
    CargaGrid DataGrid2, data3, True
   
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
    BrutoFac = CSng(Text1(14).Text) - CSng(Text1(15).Text) - CSng(Text1(16).Text)
    Text1(17).Text = Format(BrutoFac, FormatoImporte)
    
    'poner descripcion campos
    Text2(10).Text = PonerNombreDeCod(Text1(10), conAri, "sforpa", "nomforpa")
    Text2(13).Text = PonerNombreDeCod(Text1(13), conAri, "straba", "nomtraba", "codtraba")
    
    
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
    'Si estamos en Insertar adem�s limpia los campos Text1
    'si estamos en modificar bloquea las compos que son clave primaria
    BloquearText1 Me, Modo
    '---- laura 24/10/2006: si ponemos las claves de la tabla con ON UPDATE CASCADE
    'podemos permitir modificar la fecha de la factura que es clave primaria
'    If Modo = 4 Then BloquearTxt Text1(1), False
    
    
    Me.Check1.Enabled = (Modo = 1 Or Modo = 3 Or Modo = 4)
    
    b = (Modo <> 1)
    'Campos N� Factura bloqueado y en azul
    BloquearTxt Text1(0), b, True
    BloquearTxt Text1(3), b 'referencia
    
    FrmRetencionSocios.Enabled = Not b
    
    
    'Importes siempre bloqueados
    For I = 14 To 30
        BloquearTxt Text1(I), (Modo <> 1)
    Next I

    'Campo B.Imp y Imp. IVA siempre en azul
    Text1(17).BackColor = &HFFFFC0
    Text1(27).BackColor = &HFFFFC0
    Text1(28).BackColor = &HFFFFC0
    Text1(29).BackColor = &HFFFFC0
    Text1(30).BackColor = &HC0C0FF
    
    'bloquear los Text3 que son las lineas de scafac1
    BloqueaText3
    
    'Si no es modo lineas Boquear los TxtAux
    For I = 0 To txtAux.Count - 1
        BloquearTxt txtAux(I), (Modo <> 5)
    Next I
    BloquearTxt txtAux(8), True
    
    'Si no es modo Busqueda Bloquear los TxtAux3 (son los txtaux de los albaranes de factura)
    For I = 0 To txtAux3.Count - 1
        BloquearTxt txtAux3(I), (Modo <> 1)
    Next I
    
    'ampliacion linea
    b = (Modo = 5) And Me.DataGrid1.visible
    'Modo Linea de Albaranes
    Me.Label1(35).visible = b
    Me.Label1(3).visible = b
    Me.Text2(16).visible = b
    Me.Text2(17).visible = b
    BloquearTxt Text2(16), (Modo <> 5) Or (Modo = 5 And ModificaLineas <> 1)
    BloquearTxt Text2(17), (Modo <> 5) Or (Modo = 5 And ModificaLineas <> 1)

    '---------------------------------------------
    b = (Modo <> 0 And Modo <> 2 And Modo <> 5)
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    
    For I = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(I).Enabled = b
    Next I
    Me.imgBuscar(0).Enabled = (Modo = 1)
    Me.imgBuscar(1).visible = False
                    
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
'Comprobar que los datos de la cabecera son correctos antes de Insertar o Modificar
'la cabecera del Pedido
Dim b As Boolean
On Error GoTo EDatosOK

    DatosOk = False
    
    'Para que no den errores los 0's de los importes de dtos
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
        If I = 4 Or I = 5 Or I = 6 Then
            If txtAux(I).Text = "" Then
                MsgBox "El campo " & txtAux(I).Tag & " no puede ser nulo", vbExclamation
                b = False
                PonerFoco txtAux(I)
                Exit Function
            End If
        End If
    Next I
            
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
    If Index = 17 And KeyAscii = 13 Then 'campo n� de lote y ENTER
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYpress KeyAscii
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
        Case 0, 1 'trabajador
            Text2(Index).Text = PonerNombreDeCod(Text3(Index), conAri, "straba", "nomtraba", "codtraba", "Cod. Trabajador", "N")
        Case 8 'observa 5
            PonerFocoBtn Me.cmdAceptar
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
        Case 12    'Salir
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
        Toolbar1.Buttons(5).Image = 13
        Toolbar1.Buttons(5).ToolTipText = "Modificar linea factura"
        '-- eliminar
        Toolbar1.Buttons(6).Image = 14
        Toolbar1.Buttons(6).ToolTipText = "Eliminar linea factura"
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
    vWhere = vWhere & " AND numalbar='" & data3.Recordset.Fields!NumAlbar & "'"
    vWhere = vWhere & " AND numlinea=" & Data2.Recordset.Fields!numlinea
    
    If DatosOkLinea() Then
        SQL = "UPDATE " & NomTablaLineas & " SET "
        SQL = SQL & " ampliaci=" & DBSet(Text2(16).Text, "T") & ", "
        SQL = SQL & "precioar= " & DBSet(txtAux(4).Text, "N") & ", "
        SQL = SQL & "dtoline1= " & DBSet(txtAux(5).Text, "N") & ", dtoline2= " & DBSet(txtAux(6).Text, "N") & ", "
        SQL = SQL & "importel= " & DBSet(txtAux(7).Text, "N")
        SQL = SQL & ", numlotes=" & DBSet(Text2(17).Text, "T")
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
        Me.lblIndicador.Caption = "L�neas " & TituloLinea
        PonerFocoBtn Me.cmdRegresar
    End If
    'Habilitar las opciones correctas del menu segun Modo
'    PonerModoOpcionesMenu (Modo)
'    PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
    DataGrid2.Enabled = Not b
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
Dim b As Boolean
Dim Opcion As Byte
Dim SQL As String

    On Error GoTo ECargaGrid

'    b = DataGrid1.Enabled

    If vDataGrid.Name = "DataGrid1" Then
        Opcion = 1
    Else
        Opcion = 2
    End If
    
    SQL = MontaSQLCarga(enlaza, Opcion)
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
Dim tots As String

    On Error GoTo ECargaGrid
    
    vData.Refresh
    Select Case vDataGrid.Name
        Case "DataGrid1" 'Lineas de Albaran
            'SQL = "SELECT codtipom, numfactu, fecfactu, numalbar, numlinea,
            'codalmac, codartic, nomartic, ampliaci, cantidad, precioar, origpre, dtoline1, dtoline2, importel "
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux(0)|T|Alm.|520|;S|txtAux(1)|T|Art�culo|1750|;S|txtAux(2)|T|Nombre Art.|3150|;"
            tots = tots & "N||||0|;S|txtAux(3)|T|Cantidad|950|;S|txtAux(4)|T|Precio|1300|;S|txtAux(5)|T|Dto 1|600|;S|txtAux(6)|T|Dto 2|600|;S|txtAux(7)|T|Importe|1350|;N||||0|;"
            arregla tots, DataGrid1, Me
            DataGrid1.Columns(9).Alignment = dbgRight
            DataGrid1.Columns(10).Alignment = dbgRight
            DataGrid1.Columns(12).Alignment = dbgRight
            DataGrid1.Columns(13).Alignment = dbgRight
                       
         Case "DataGrid2" 'albaranes x articulo
            'SQL = "SELECT codtipom,numfactu,fecfactu,codtipoa,numalbar, fechaalb,"
            'numpedcl,fecpedcl,sementre,numofert,fecofert, referenc, codenvio,codtraba, codtrab1, codtrab2,observa1,observa2,observa3,observa4,observa5  "
            tots = "N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux3(0)|T|Albaran|1400|;S|txtAux3(1)|T|Fecha|1300|;"
            tots = tots & "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
            tots = tots & "N||||0|;"
            arregla tots, DataGrid2, Me
        
            DataGrid2_RowColChange 1, 1
    End Select
    
    vDataGrid.HoldFields
    Exit Sub
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), Modo
End Sub

Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtAux_LostFocus(Index As Integer)

    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub
    
    Select Case Index
        Case 4 'Precio
            If txtAux(Index).Text <> "" Then
                If Not PonerFormatoDecimal(txtAux(Index), 2) Then txtAux(Index).Text = "" 'Tipo 2: Decimal(10,4)
            End If
            
        Case 5, 6 'Descuentos
            PonerFormatoDecimal txtAux(Index), 4 'Tipo 4: Decimal(4,2)
            If Index = 6 Then PonerFoco Me.Text2(16)
            
        Case 7 'Importe Linea
            PonerFormatoDecimal txtAux(Index), 1 'Tipo 3: Decimal(12,2)
    End Select
    
    If (Index = 3 Or Index = 4 Or Index = 6) Then  'Cant., Precio, Dto1, Dto2
        If txtAux(1).Text = "" Then Exit Sub
        txtAux(7).Text = CalcularImporte(txtAux(3).Text, txtAux(4).Text, txtAux(5).Text, txtAux(6).Text, vParamAplic.TipoDtos)
        PonerFormatoDecimal txtAux(7), 1
    End If
End Sub


Private Sub BotonMtoLineas(numTab As Integer, cad As String)
    Me.SSTab1.Tab = numTab
    
    If Me.DataGrid1.visible Then 'Lineas de Albaranes
        If Me.Data2.Recordset.RecordCount < 1 Then
            MsgBox "La factura no tiene lineas.", vbInformation
            Exit Sub
        End If
        TituloLinea = cad
        
        ModificaLineas = 0
        PonerModo 5
        PonerBotonCabecera True
    End If
End Sub


Private Function Eliminar() As Boolean
Dim SQL As String
Dim cta As String
Dim b As Boolean

    On Error GoTo FinEliminar

        b = False
        Eliminar = False
        If Data1.Recordset.EOF Then Exit Function
        
        conn.BeginTrans
        ConnConta.BeginTrans
        
        'Eliminar en la tabla pagos de la Contabilidad: spagop
        '------------------------------------------------
        cta = DevuelveDesdeBDNew(conAri, "sprove", "codmacta", "codprove", Text1(2).Text, "N")
        If vParamAplic.ContabilidadNueva Then
            SQL = "codmacta"
        Else
            SQL = "ctaprove"
        End If
        SQL = SQL & " ='" & cta & "' AND numfactu='" & Data1.Recordset.Fields!NumFactu & "'"
        SQL = SQL & " AND fecfactu='" & Format(Data1.Recordset.Fields!Fecfactu, FormatoFecha) & "'"
        If vParamAplic.ContabilidadNueva Then SQL = SQL & " AND numserie='" & SerieFraPro & "'"
        
        ConnConta.Execute "Delete from " & IIf(vParamAplic.ContabilidadNueva, "pagos", "spagop") & " WHERE " & SQL
        b = True
        
        
        'Eliminar en tablas de factura de Ariges: scafpc, scafpa, slifpc
        '---------------------------------------------------------------
        If b Then
            SQL = " " & ObtenerWhereCP(True)
        
            'Lineas de facturas (slifpc)
            conn.Execute "Delete from " & NomTablaLineas & SQL
        
            'Lineas de cabeceras de albaranes de la factura
            conn.Execute "Delete from scafpa " & SQL
            
            'Cabecera de facturas (scafpc)
            conn.Execute "Delete from " & NombreTabla & SQL
        End If
        
        'Eliminar los movimientos generados por el albaran que genero la factura
        '-----------------------------------------------------------------------
        If b Then
        
        End If
        
'        b = True
        
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Factura", Err.Description
        b = False
    End If
    If Not b Then
        conn.RollbackTrans
        ConnConta.RollbackTrans
    Else
        conn.CommitTrans
        ConnConta.CommitTrans
    End If
    Eliminar = b
End Function


Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ning�n registro
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
             If Modo <> 5 Then
                PonerModo 2
                PonerCampos
             End If
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
    SQL = "codprove= " & Text1(2).Text & " and numfactu= '" & Text1(0).Text & "' and fecfactu='" & Format(Text1(1).Text, FormatoFecha) & "' "
    If conWhere Then SQL = " WHERE " & SQL
    ObtenerWhereCP = SQL
End Function


Private Function MontaSQLCarga(enlaza As Boolean, Opcion As Byte) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Bas�ndose en la informaci�n proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim SQL As String
    
    If Opcion = 1 Then
        SQL = "SELECT codprove, numfactu, fecfactu, numalbar, numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel,numlotes "
        SQL = SQL & " FROM slifpc " 'lineas de factura
    ElseIf Opcion = 2 Then
        SQL = "SELECT codprove,numfactu,fecfactu,numalbar, fechaalb, numpedpr,fecpedpr,codtrab1, codtrab2,observa1,observa2,observa3,observa4,observa5  "
        SQL = SQL & " FROM scafpa " 'cabeceras albaranes de la factura
    End If
    
    If enlaza Then
        SQL = SQL & " " & ObtenerWhereCP(True)
        'lineas factura proveedor
        If Opcion = 1 Then SQL = SQL & " AND numalbar=" & DBSet(data3.Recordset.Fields!NumAlbar, "T")
    Else
        SQL = SQL & " WHERE numfactu = -1"
    End If
    SQL = SQL & " ORDER BY codprove, numfactu, fecfactu,numalbar "
    If Opcion = 1 Then SQL = SQL & ", numlinea "
    MontaSQLCarga = SQL
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar seg�n el modo en que estemos
Dim b As Boolean

        b = (Modo = 2) Or (Modo = 5 And ModificaLineas = 0)
        'Modificar
        Toolbar1.Buttons(5).Enabled = b
        Me.mnModificar.Enabled = b
        'eliminar
        Toolbar1.Buttons(6).Enabled = (Modo = 2)
        Me.mnEliminar.Enabled = (Modo = 2)
            
        b = (Modo = 2)
        'Mantenimiento lineas
        Toolbar1.Buttons(9).Enabled = b
        Me.mnLineas.Enabled = b
        'Imprimir
'        Toolbar1.Buttons(10).Enabled = b
'        Me.mnImprimir.Enabled = b
        
        b = (Modo >= 3) Or Modo = 1
        'Buscar
        Toolbar1.Buttons(1).Enabled = Not b
        Me.mnBuscar.Enabled = Not b
        'Ver Todos
        Toolbar1.Buttons(2).Enabled = Not b
        Me.mnVerTodos.Enabled = Not b
End Sub


Private Sub PonerDatosProveedor(codProve As String, Optional nifProve As String)
Dim vProve As CProveedor
Dim Observaciones As String
    
    On Error GoTo EPonerDatos
    
    If codProve = "" Then
        LimpiarDatosProve
        Exit Sub
    End If

    Set vProve = New CProveedor
    'si se ha modificado el proveedor volver a cargar los datos
    If vProve.Existe(codProve) Then
        If vProve.LeerDatos(codProve) Then
           
            EsDeVarios = vProve.DeVarios
            BloquearDatosProve (EsDeVarios)
        
            If Modo = 4 And EsDeVarios Then 'Modificar
                'si no se ha modificado el proveedor no hacer nada
                If CLng(Text1(2).Text) = CLng(Data1.Recordset!codProve) Then
                    Set vProve = Nothing
                    Exit Sub
                End If
            End If
        
            Text1(2).Text = vProve.Codigo
            FormateaCampo Text1(2)
            
            If (Modo = 3) Or (Modo = 4) Then
                Text1(3).Text = vProve.Nombre  'Nom prove
                Text1(6).Text = vProve.Domicilio
                Text1(7).Text = vProve.CPostal
                Text1(8).Text = vProve.Poblacion
                Text1(9).Text = vProve.Provincia
                Text1(4).Text = vProve.NIF
                Text1(5).Text = DBLet(vProve.TfnoAdmon, "T")
            End If
            
            Observaciones = DBLet(vProve.Observaciones)
            If Observaciones <> "" Then
                MsgBox Observaciones, vbInformation, "Observaciones del proveedor"
            End If
        End If
    Else
        LimpiarDatosProve
    End If
    Set vProve = Nothing

EPonerDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poner Datos Proveedor", Err.Description
End Sub


Private Sub PonerDatosProveVario(nifProve As String)
'Poner el los campos Text el valor del proveedor
Dim vProve As CProveedor
Dim b As Boolean
   
    If nifProve = "" Then Exit Sub
   
    Set vProve = New CProveedor
    b = vProve.LeerDatosProveVario(nifProve)
    If b Then
        Text1(3).Text = vProve.Nombre   'Nom proveedor
        Text1(6).Text = vProve.Domicilio
        Text1(7).Text = vProve.CPostal
        Text1(8).Text = vProve.Poblacion
        Text1(9).Text = vProve.Provincia
        Text1(5).Text = DBLet(vProve.TfnoAdmon, "T")
    End If
    Set vProve = Nothing
End Sub


Private Sub LimpiarDatosProve()
Dim I As Byte

    For I = 3 To 9
        Text1(I).Text = ""
    Next I
End Sub
   

Private Function ModificaAlbxFac() As Boolean
Dim SQL As String
Dim b As Boolean
On Error GoTo EModificaAlb
    
    ModificaAlbxFac = False
    If data3.Recordset.EOF Then Exit Function
    
    'comprobar datos OK de la scafac1
     b = CompForm(Me, 2) 'Comprobar formato datos ok de la cabecera alb: opcion=2
    If Not b Then Exit Function
    
    SQL = "UPDATE scafpa SET codtrab2=" & DBSet(Text3(0).Text, "N", "S") & ", "
    SQL = SQL & "codtrab1=" & DBSet(Text3(1).Text, "N", "S")
    If Me.FrameObserva.visible Then
        SQL = SQL & ", observa1=" & DBSet(Text3(4).Text, "T")
        SQL = SQL & ", observa2=" & DBSet(Text3(5).Text, "T")
        SQL = SQL & ", observa3=" & DBSet(Text3(6).Text, "T")
        SQL = SQL & ", observa4=" & DBSet(Text3(7).Text, "T")
        SQL = SQL & ", observa5=" & DBSet(Text3(8).Text, "T")
    End If
    SQL = SQL & ObtenerWhereCP(True)
    SQL = SQL & " AND numalbar=" & DBSet(data3.Recordset.Fields!NumAlbar, "T")
    conn.Execute SQL
    ModificaAlbxFac = True
    
EModificaAlb:
If Err.Number <> 0 Then MuestraError Err.Number, "Modificar Albaranes de factura", Err.Description
End Function



Private Function ModificarFactura(Optional sqlLineas As String) As Boolean
'si se ha modificado la linea de slifac, a�adir a la transaccion la modificaci�n de la linea y recalcular
Dim bol As Boolean
Dim MenError As String
Dim SQL As String
Dim vFactu As CFacturaCom
On Error GoTo EModFact

    bol = False
    conn.BeginTrans
    ConnConta.BeginTrans
    
    If sqlLineas <> "" Then
        'actualizar el importe de la linea modificada
        MenError = "Modificando lineas de Factura."
        conn.Execute sqlLineas
    End If
    
    'recalcular las bases imponibles x IVA
    MenError = "Recalcular importes IVA"
    bol = ActualizarDatosFactura
    
    If bol Then
        'modificamos la scafpc
        MenError = "Modificando cabecera de factura"
        bol = ModificaDesdeFormulario(Me, 1)
        
        If bol Then
            'Si es proveedor de varios actualizar datos proveedor en tabla:sprvar
            MenError = "Modificando datos proveedor varios"
            bol = ActualizarProveVarios(Text1(2).Text, Text1(4).Text)
        End If
        
        If bol Then
            MenError = "Modificando albaranes de factura"
            'modificar la tabla: scafpa
            bol = ModificaAlbxFac
            
            If bol Then 'si se ha modificado la factura
                MenError = "Actualizando en Tesoreria"
                'y eliminar de tesoreria conta.spagop los registros de la factura
                
                'antes de Eliminar en las tablas de la Contabilidad
                Set vFactu = New CFacturaCom
                bol = vFactu.LeerDatos3(Text1(2).Text, Text1(0).Text, Text1(1).Text)
                
                If bol Then
                    'Eliminar de la spagop
                    If vParamAplic.ContabilidadNueva Then
                        SQL = " codmacta='" & vFactu.CtaProve & "' AND numfactu='" & Data1.Recordset.Fields!NumFactu & "'"
                        SQL = SQL & " AND fecfactu='" & Format(Data1.Recordset.Fields!Fecfactu, FormatoFecha) & "'"
                    
                        ConnConta.Execute "Delete from pagos WHERE " & SQL
                    Else
                        SQL = " ctaprove='" & vFactu.CtaProve & "' AND numfactu='" & Data1.Recordset.Fields!NumFactu & "'"
                        SQL = SQL & " AND fecfactu='" & Format(Data1.Recordset.Fields!Fecfactu, FormatoFecha) & "'"
                    
                        ConnConta.Execute "Delete from spagop WHERE " & SQL
                    End If
                    
                    'Volvemos a grabar en TESORERIA. Tabla de Contabilidad: sconta.spagop
                    If bol Then
                        'Aqui veriamos si ponemos datos de la cuenta bancaria del proveedor
                        bol = vFactu.InsertarEnTesoreria(MenError)
                    End If
                End If
                Set vFactu = Nothing
            End If
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



Private Function FactContabilizada() As Boolean
Dim cta As String, numasien As String
On Error GoTo EContab

    'comprabar que se puede modificar/eliminar la factura
    If Me.Check1.Value = 1 Then 'si esta contabilizada
        'comprobar en la contabilidad si esta contabilizada
        cta = DevuelveDesdeBDNew(conAri, "sprove", "codmacta", "codprove", Text1(2).Text, "N")
        If cta <> "" Then
            numasien = DevuelveDesdeBDNew(conConta, "cabfactprov", "numasien", "codmacta", cta, "T", , "numfacpr", Text1(0).Text, "T", "fecfacpr", Text1(1).Text, "F")
            If numasien <> "" Then
                FactContabilizada = True
                MsgBox "La factura esta contabilizada y no se puede modificar ni eliminar.", vbInformation
                Exit Function
            Else
                FactContabilizada = False
            End If
        Else
            FactContabilizada = True
            Exit Function
        End If
    Else
        FactContabilizada = False
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


Private Sub BloquearDatosProve(bol As Boolean)
Dim I As Byte

    'bloquear/desbloquear campos de datos segun sea de varios o no
    If Modo <> 5 Then
        Me.imgBuscar(1).visible = bol 'NIF
        Me.imgBuscar(1).Enabled = bol 'NIF
        Me.imgBuscar(2).Enabled = bol 'poblacion
        
        For I = 3 To 9 'si no es de varios no se pueden modificar los datos
            BloquearTxt Text1(I), Not bol
        Next I
    End If
End Sub


Private Function ActualizarProveVarios(Prove As String, NIF As String) As Boolean
'Modifica los datos de la tabla de Proveedores Varios
Dim vProve As CProveedor
On Error GoTo EActualizarCV

    ActualizarProveVarios = False
    
    Set vProve = New CProveedor
    If EsProveedorVarios(Prove) Then
        vProve.NIF = NIF
        vProve.Nombre = Text1(3).Text
        vProve.Domicilio = Text1(6).Text
        vProve.CPostal = Text1(7).Text
        vProve.Poblacion = Text1(8).Text
        vProve.Provincia = Text1(9).Text
        vProve.TfnoAdmon = Text1(5).Text
        vProve.ActualizarProveV (NIF)
    End If
    Set vProve = Nothing
    
    ActualizarProveVarios = True
    
EActualizarCV:
    If Err.Number <> 0 Then
        ActualizarProveVarios = False
    Else
        ActualizarProveVarios = True
    End If
End Function


Private Function ObtenerSelFactura() As String
'Cuando venimos desde dobleClick en Movimientos de Articulos para Albaranes ya
'Facturados, abrimos este form pero cargando los datos de la factura
'correspendiente al albaran que se selecciono
Dim cad As String
Dim RS As ADODB.Recordset
On Error Resume Next

    cad = "SELECT codprove,numfactu,fecfactu FROM scafpa "
    cad = cad & " WHERE codprove=" & DBSet(hcoCodProve, "N") & " AND numalbar=" & DBSet(hcoCodMovim, "T")
    cad = cad & " AND fechaalb=" & DBSet(hcoFechaMovim, "F")

    Set RS = New ADODB.Recordset
    RS.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then 'where para la factura
        cad = " WHERE codprove=" & RS!codProve & " AND numfactu= '" & RS!NumFactu & "' AND fecfactu=" & DBSet(RS!Fecfactu, "F")
    Else
        cad = " where numfactu=-1"
    End If
    RS.Close
    Set RS = Nothing

    ObtenerSelFactura = cad
End Function



Private Function ActualizarDatosFactura() As Boolean
Dim vFactu As CFacturaCom
Dim cadSEL As String
Dim FacturaSinIVA As Boolean
Dim Aux As String
Dim CambiarIVA As Boolean

    Set vFactu = New CFacturaCom
    cadSEL = ObtenerWhereCP(False)
    cadSEL = "slifpc." & cadSEL
    vFactu.DtoPPago = CCur(Text1(11).Text)
    vFactu.DtoGnral = CCur(Text1(12).Text)
    
    'Si tiene RETENCION
    If Me.FrmRetencionSocios.visible Then
        vFactu.PorRet = ImporteFormateado(Text1(32).Text)
        vFactu.ImpRet2 = ImporteFormateado(Text1(33).Text)
    End If

    'Abril. Lo comento
    'Aux = DevuelveDesdeBD(conAri, "tipprove", "sprove", "codprove", Text1(2).Text)
    'FacturaSinIVA = False
    'If Aux = "1" Or Aux = "2" Then FacturaSinIVA = True
    CambiarIVA = False
    If CDate(Text1(1).Text) < CDate("01/09/2012") Then CambiarIVA = True



    vFactu.FijarTipoIvaProveedor Val(Text1(2).Text)
    If vFactu.CalcularDatosFactura2(cadSEL, "scafpa", "slifpc", CambiarIVA, False) Then
        Text1(14).Text = vFactu.BrutoFac
        Text1(15).Text = vFactu.ImpPPago
        Text1(16).Text = vFactu.ImpGnral
        Text1(17).Text = vFactu.BaseImp
        Text1(18).Text = vFactu.TipoIVA1
        Text1(19).Text = vFactu.TipoIVA2
        Text1(20).Text = vFactu.TipoIVA3
        Text1(21).Text = vFactu.PorceIVA1
        Text1(22).Text = vFactu.PorceIVA2
        Text1(23).Text = vFactu.PorceIVA3
        Text1(24).Text = vFactu.BaseIVA1
        Text1(25).Text = vFactu.BaseIVA2
        Text1(26).Text = vFactu.BaseIVA3
        Text1(27).Text = vFactu.ImpIVA1
        Text1(28).Text = vFactu.ImpIVA2
        Text1(29).Text = vFactu.ImpIVA3
        Text1(30).Text = vFactu.TotalFac
        If Me.FrmRetencionSocios.visible Then
            If vFactu.PorRet = 0 Then
                Text1(32).Text = ""
                Text1(33).Text = ""
            Else
                Text1(32).Text = vFactu.PorRet
                Text1(33).Text = vFactu.ImpRet2
            End If
        End If
        
        FormatoDatosTotales
        
        ActualizarDatosFactura = True
    Else
        ActualizarDatosFactura = False
        MuestraError Err.Number, "Recalculando Factura", Err.Description
    End If
    Set vFactu = Nothing
End Function


Private Sub FormatoDatosTotales()
Dim I As Byte

    For I = 14 To 17
'        Text1(I).Text = QuitarCero(Text1(I).Text)
        FormateaCampo Text1(I)
    Next I
    
    For I = 24 To 26
        If Text1(I).Text <> "" Then
            'Si la Base Imp. es 0
            If CSng(Text1(I).Text) = 0 Then
                Text1(I).Text = QuitarCero(Text1(I).Text)
                Text1(I - 3).Text = QuitarCero(Text1(I - 3).Text)
                Text1(I - 6).Text = QuitarCero(Text1(I - 6).Text)
                Text1(I + 3).Text = QuitarCero(Text1(I + 3).Text)
            Else
                FormateaCampo Text1(I)
                FormateaCampo Text1(I - 3)
                FormateaCampo Text1(I - 6)
                FormateaCampo Text1(I + 3)
            End If
        Else 'No hay Base Imponible
            Text1(I - 3).Text = QuitarCero(Text1(I - 3).Text)
            Text1(I - 6).Text = QuitarCero(Text1(I - 6).Text)
            Text1(I + 3).Text = ""
        End If
    Next I
    
    If Me.FrmRetencionSocios.visible Then
        FormateaCampo Text1(32)
        FormateaCampo Text1(33)
    End If
End Sub



Private Sub ComprobarDatosTotales()
Dim I As Byte

    For I = 14 To 17
        Text1(I).Text = ComprobarCero(Text1(I).Text)
    Next I
End Sub
