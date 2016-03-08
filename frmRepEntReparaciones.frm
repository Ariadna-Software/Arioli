VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRepEntReparaciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reparaciones"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   540
   ClientWidth     =   11760
   ClipControls    =   0   'False
   Icon            =   "frmRepEntReparaciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   120
      TabIndex        =   54
      Tag             =   "A|N|S|||scarep|contestado||S|"
      ToolTipText     =   "Descliente"
      Top             =   1080
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   11456
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Datos basicos "
      TabPicture(0)   =   "frmRepEntReparaciones.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameOtros"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrameClientes"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Presupuesto / S.A.T."
      TabPicture(1)   =   "frmRepEntReparaciones.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text1(27)"
      Tab(1).Control(1)=   "Text1(26)"
      Tab(1).Control(2)=   "Text1(25)"
      Tab(1).Control(3)=   "Check1"
      Tab(1).Control(4)=   "Combo1"
      Tab(1).Control(5)=   "Text1(16)"
      Tab(1).Control(6)=   "Text1(17)"
      Tab(1).Control(7)=   "Text1(18)"
      Tab(1).Control(8)=   "Text1(19)"
      Tab(1).Control(9)=   "Text1(20)"
      Tab(1).Control(10)=   "Text1(21)"
      Tab(1).Control(11)=   "Text2(21)"
      Tab(1).Control(12)=   "Text1(22)"
      Tab(1).Control(13)=   "imgBuscar(8)"
      Tab(1).Control(14)=   "Label9(5)"
      Tab(1).Control(15)=   "Label9(4)"
      Tab(1).Control(16)=   "imgFecha(5)"
      Tab(1).Control(17)=   "Label11(1)"
      Tab(1).Control(18)=   "Label12(1)"
      Tab(1).Control(19)=   "Label12(0)"
      Tab(1).Control(20)=   "Line2"
      Tab(1).Control(21)=   "Line1"
      Tab(1).Control(22)=   "Label1(12)"
      Tab(1).Control(23)=   "Label11(0)"
      Tab(1).Control(24)=   "Label2(1)"
      Tab(1).Control(25)=   "Label9(1)"
      Tab(1).Control(26)=   "imgFecha(2)"
      Tab(1).Control(27)=   "Label9(2)"
      Tab(1).Control(28)=   "imgFecha(3)"
      Tab(1).Control(29)=   "Label1(13)"
      Tab(1).Control(30)=   "Label9(3)"
      Tab(1).Control(31)=   "imgFecha(4)"
      Tab(1).Control(32)=   "Label2(2)"
      Tab(1).Control(33)=   "Label2(3)"
      Tab(1).ControlCount=   34
      TabCaption(2)   =   "Lineas"
      TabPicture(2)   =   "frmRepEntReparaciones.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdAux(1)"
      Tab(2).Control(1)=   "cmdAux(0)"
      Tab(2).Control(2)=   "txtAux(2)"
      Tab(2).Control(3)=   "txtAux(7)"
      Tab(2).Control(4)=   "txtAux(6)"
      Tab(2).Control(5)=   "txtAux(5)"
      Tab(2).Control(6)=   "txtAux(4)"
      Tab(2).Control(7)=   "txtAux(3)"
      Tab(2).Control(8)=   "txtAux(1)"
      Tab(2).Control(9)=   "txtAux(0)"
      Tab(2).Control(10)=   "DataGrid1"
      Tab(2).Control(11)=   "Label1(16)"
      Tab(2).ControlCount=   12
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Index           =   1
         Left            =   -72600
         TabIndex        =   47
         Top             =   5820
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Index           =   0
         Left            =   -74280
         TabIndex        =   46
         Top             =   5820
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   2
         Left            =   -72360
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   38
         Tag             =   "Nombre Artículo"
         Text            =   "nomArtic"
         Top             =   5820
         Visible         =   0   'False
         Width           =   3285
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   7
         Left            =   -65160
         MaxLength       =   12
         TabIndex        =   43
         Tag             =   "Importe"
         Text            =   "Importe"
         Top             =   5940
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   6
         Left            =   -65760
         MaxLength       =   30
         TabIndex        =   42
         Tag             =   "Descuento 2"
         Text            =   "Dto2"
         Top             =   5940
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   5
         Left            =   -66360
         MaxLength       =   5
         TabIndex        =   41
         Tag             =   "Descuento 1"
         Text            =   "Dto1"
         Top             =   5940
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
         Left            =   -67800
         MaxLength       =   12
         TabIndex        =   40
         Tag             =   "Precio"
         Text            =   "123,456.7879"
         Top             =   5940
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
         Left            =   -69000
         MaxLength       =   16
         TabIndex        =   39
         Tag             =   "Cantidad"
         Text            =   "1,234,567,891.25"
         Top             =   5940
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   -74040
         MaxLength       =   18
         TabIndex        =   37
         Tag             =   "Código Artículo"
         Text            =   "Artic Artic Artic5"
         Top             =   5760
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   -74880
         MaxLength       =   15
         TabIndex        =   36
         Tag             =   "Código Almacen"
         Text            =   "codalmac"
         Top             =   5760
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   1155
         Index           =   27
         Left            =   -72195
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Tag             =   "F|T|S|||scarep|observasat|||"
         Text            =   "frmRepEntReparaciones.frx":0060
         Top             =   5160
         Width           =   8280
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   26
         Left            =   -72195
         MaxLength       =   10
         TabIndex        =   28
         Tag             =   "Fecha Entrega SAT|F|S|||scarep|fecentresat|dd/mm/yyyy|N|"
         Text            =   "Text1"
         Top             =   4680
         Width           =   1080
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   25
         Left            =   -72195
         MaxLength       =   7
         TabIndex        =   27
         Tag             =   "Imp reparacion SAT|N|S|||scarep|importesat|0.00||"
         Text            =   "Text1"
         Top             =   4080
         Width           =   1320
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Cliente avisado"
         Height          =   195
         Left            =   -67080
         TabIndex        =   23
         Tag             =   "A|N|S|||scarep|avisocli||S|"
         Top             =   1650
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmRepEntReparaciones.frx":0066
         Left            =   -65700
         List            =   "frmRepEntReparaciones.frx":0073
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Tag             =   "Aceptado|N|S|||scarep|contestado||N|"
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   16
         Left            =   -72240
         MaxLength       =   7
         TabIndex        =   18
         Tag             =   "Imp pres1|N|S|||scarep|imppresu1|0.00||"
         Text            =   "Text1"
         Top             =   1080
         Width           =   1320
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   17
         Left            =   -72240
         MaxLength       =   7
         TabIndex        =   19
         Tag             =   "Imp presupuesto 2|N|S|||scarep|impresu2|0.00||"
         Text            =   "Text1"
         Top             =   1560
         Width           =   1320
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   18
         Left            =   -68640
         MaxLength       =   10
         TabIndex        =   20
         Tag             =   "Fecha presupuesto|F|S|||scarep|fecha|dd/mm/yyyy|N|"
         Text            =   "Text1"
         Top             =   1080
         Width           =   1080
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   19
         Left            =   -68640
         MaxLength       =   10
         TabIndex        =   21
         Tag             =   "Fecha aprobacion|F|S|||scarep|fechaaprob|dd/mm/yyyy|N|"
         Text            =   "Text1"
         Top             =   1560
         Width           =   1080
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   20
         Left            =   -72195
         MaxLength       =   10
         TabIndex        =   25
         Tag             =   "Fecha envio|F|S|||scarep|fecenviosat|dd/mm/yyyy|N|"
         Text            =   "Text1"
         Top             =   3120
         Width           =   1080
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   21
         Left            =   -72195
         MaxLength       =   4
         TabIndex        =   24
         Tag             =   "Servicio SAT|N|S|||scarep|codman|000|N|"
         Text            =   "Text1"
         Top             =   2640
         Width           =   1080
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   300
         Index           =   21
         Left            =   -71040
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   99
         Text            =   "Text2"
         Top             =   2640
         Width           =   3885
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   22
         Left            =   -72195
         MaxLength       =   15
         TabIndex        =   26
         Tag             =   "Nº Reparación|T|S|||scarep|resguardosat|||"
         Text            =   "Text1"
         Top             =   3600
         Width           =   3720
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   3585
         Left            =   120
         TabIndex        =   77
         Top             =   2760
         Width           =   11175
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   300
            Index           =   24
            Left            =   2835
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   96
            Text            =   "Text2"
            Top             =   2160
            Width           =   3285
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   24
            Left            =   2280
            MaxLength       =   4
            TabIndex        =   14
            Tag             =   "Trabajo realizado|N|S|||scarep|codtrabajo|00|N|"
            Text            =   "Te"
            Top             =   2160
            Width           =   525
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   300
            Index           =   23
            Left            =   2835
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   94
            Text            =   "Text2"
            Top             =   720
            Width           =   3285
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   23
            Left            =   2280
            MaxLength       =   4
            TabIndex        =   11
            Tag             =   "Tipo averia|N|S|||scarep|codavi|00|N|"
            Text            =   "Te"
            Top             =   720
            Width           =   525
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   0
            Left            =   9855
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   82
            Tag             =   "Tipo Albaran|T|S|||schrep|codtipom||N|"
            Text            =   "Text2"
            Top             =   3240
            Width           =   615
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   14
            Left            =   10560
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   81
            Tag             =   "Fecha Alb|F|S|||schrep|fechaalb|dd/mm/yyyy|N|"
            Text            =   "Text2"
            Top             =   3240
            Width           =   1065
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   15
            Left            =   8880
            Locked          =   -1  'True
            MaxLength       =   7
            TabIndex        =   80
            Tag             =   "Nº Albaran|T|S|||schrep|numalbar|0000000|N|"
            Text            =   "Text2"
            Top             =   3240
            Width           =   890
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   14
            Left            =   2280
            MaxLength       =   80
            TabIndex        =   15
            Tag             =   "Texto Reparación 1|T|S|||scarep|textore1||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   2520
            Width           =   6495
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   13
            Left            =   2280
            MaxLength       =   80
            TabIndex        =   16
            Tag             =   "Texto Reparación 2|T|S|||scarep|textore2||N|"
            Text            =   "Text1"
            Top             =   2880
            Width           =   6495
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   12
            Left            =   2280
            MaxLength       =   80
            TabIndex        =   17
            Tag             =   "Texto Reparación 3|T|S|||scarep|textore3||N|"
            Text            =   "Text1"
            Top             =   3240
            Width           =   6495
         End
         Begin VB.TextBox Text1 
            Height          =   300
            Index           =   8
            Left            =   2280
            MaxLength       =   80
            TabIndex        =   9
            Tag             =   "Material con el que entra|T|S|||scarep|material||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   0
            Width           =   6495
         End
         Begin VB.TextBox Text1 
            Height          =   300
            Index           =   9
            Left            =   2280
            MaxLength       =   80
            TabIndex        =   10
            Tag             =   "Avería detectada|T|S|||scarep|tipoaver||N|"
            Text            =   "Text1"
            Top             =   360
            Width           =   6495
         End
         Begin VB.TextBox Text1 
            Height          =   300
            Index           =   10
            Left            =   2280
            MaxLength       =   80
            TabIndex        =   12
            Tag             =   "Situación de la Reparación|T|S|||scarep|motivore||N|"
            Text            =   "Text1"
            Top             =   1200
            Width           =   6495
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   11
            Left            =   2280
            MaxLength       =   2
            TabIndex        =   13
            Tag             =   "Motivo Pendiente Rep.|N|S|||scarep|codmotre|00|N|"
            Text            =   "Te"
            Top             =   1560
            Width           =   525
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   300
            Index           =   11
            Left            =   2835
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   79
            Text            =   "Text2"
            Top             =   1560
            Width           =   3285
         End
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            Height          =   285
            Index           =   15
            Left            =   9480
            MaxLength       =   80
            TabIndex        =   78
            Tag             =   "Viso|T|S|||scarep|numaviso||N|"
            Text            =   "Text1"
            Top             =   0
            Width           =   1575
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   7
            Left            =   1440
            Picture         =   "frmRepEntReparaciones.frx":0091
            Top             =   2160
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Trabajo realizado"
            Height          =   195
            Index           =   15
            Left            =   120
            TabIndex        =   98
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   6
            Left            =   1080
            Picture         =   "frmRepEntReparaciones.frx":0193
            Top             =   720
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo averia"
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   95
            Top             =   720
            Width           =   975
         End
         Begin VB.Image imgVerAlbaran 
            Height          =   240
            Left            =   9600
            Picture         =   "frmRepEntReparaciones.frx":0295
            Top             =   2880
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Aviso"
            Height          =   255
            Index           =   11
            Left            =   8880
            TabIndex        =   93
            Top             =   0
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Alb."
            Height          =   255
            Index           =   8
            Left            =   9855
            TabIndex        =   91
            Top             =   3000
            Width           =   630
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Alb."
            Height          =   255
            Index           =   22
            Left            =   10560
            TabIndex        =   90
            Top             =   3000
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Nº Albaran"
            Height          =   255
            Index           =   10
            Left            =   8880
            TabIndex        =   89
            Top             =   3000
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Tipo de Avería detectada"
            Height          =   255
            Left            =   120
            TabIndex        =   87
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label8 
            Caption         =   "Situación de la Reparación"
            Height          =   255
            Left            =   120
            TabIndex        =   86
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label Label1 
            Caption         =   "Motivo Pendiente Rep."
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   85
            Top             =   1590
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Texto Reparación"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   84
            Top             =   2520
            Width           =   1695
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   5
            Left            =   1920
            Picture         =   "frmRepEntReparaciones.frx":0C97
            Top             =   1575
            Width           =   240
         End
         Begin VB.Label Label6 
            Caption         =   "Material con el que entra"
            Height          =   255
            Left            =   120
            TabIndex        =   83
            Top             =   0
            Width           =   1815
         End
      End
      Begin VB.Frame FrameClientes 
         Caption         =   "Datos Clientes"
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
         Height          =   2265
         Left            =   120
         TabIndex        =   64
         Top             =   360
         Width           =   5775
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   6
            Left            =   960
            MaxLength       =   6
            TabIndex        =   3
            Tag             =   "Cod. Cliente|N|N|0|999999|scarep|codclien|000000|N|"
            Text            =   "Text1"
            Top             =   360
            Width           =   740
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   6
            Left            =   1740
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   71
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwww aq"
            Top             =   360
            Width           =   3780
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   7
            Left            =   1540
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   70
            Text            =   "Text2"
            Top             =   1845
            Width           =   3525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   7
            Left            =   960
            MaxLength       =   3
            TabIndex        =   4
            Tag             =   "Direccion/Dpto.|N|S|0|999|scarep|coddirec|000|N|"
            Text            =   "Text1"
            Top             =   1845
            Width           =   540
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   300
            Index           =   12
            Left            =   960
            Locked          =   -1  'True
            MaxLength       =   6
            TabIndex        =   69
            Text            =   "Text15"
            Top             =   1500
            Width           =   630
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   300
            Index           =   13
            Left            =   1640
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   68
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
            Top             =   1500
            Width           =   3880
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   300
            Index           =   9
            Left            =   3675
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   67
            Text            =   "12345678911234567899"
            Top             =   720
            Width           =   1845
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   300
            Index           =   8
            Left            =   960
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   66
            Text            =   "123456789"
            Top             =   720
            Width           =   1470
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   300
            Index           =   10
            Left            =   960
            Locked          =   -1  'True
            MaxLength       =   35
            TabIndex        =   65
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aq"
            Top             =   1080
            Width           =   4560
         End
         Begin VB.Label Label1 
            Caption         =   "Cliente"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   76
            Top             =   360
            Width           =   495
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   660
            ToolTipText     =   "Buscar cliente"
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Direc."
            Height          =   255
            Index           =   2
            Left            =   150
            TabIndex        =   75
            Top             =   1845
            Width           =   495
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   4
            Left            =   720
            ToolTipText     =   "Buscar direc./dpto"
            Top             =   1890
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Teléfono"
            Height          =   255
            Index           =   19
            Left            =   2925
            TabIndex        =   74
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "NIF"
            Height          =   255
            Index           =   20
            Left            =   150
            TabIndex        =   73
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Domicilio"
            Height          =   255
            Index           =   7
            Left            =   150
            TabIndex        =   72
            Top             =   1080
            Width           =   735
         End
      End
      Begin VB.Frame FrameOtros 
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
         Height          =   2265
         Left            =   6000
         TabIndex        =   55
         Top             =   360
         Width           =   5295
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   2
            Left            =   3840
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   88
            Text            =   "123456789"
            Top             =   780
            Width           =   1230
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   3
            Left            =   3840
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   92
            Text            =   "1234567891"
            Top             =   1500
            Width           =   1230
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   4
            Left            =   3840
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   97
            Text            =   "1234567891"
            Top             =   1140
            Width           =   1230
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   2
            Left            =   1365
            MaxLength       =   7
            TabIndex        =   6
            Tag             =   "Nº Reparación|N|S|0|9999999|scarep|numrepar|0000000|S|"
            Text            =   "Text1"
            Top             =   780
            Width           =   1080
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   4
            Left            =   1365
            MaxLength       =   10
            TabIndex        =   8
            Tag             =   "Fecha Repar|F|N|||scarep|fecentre|dd/mm/yyyy|N|"
            Text            =   "Text1"
            Top             =   1500
            Width           =   1080
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   3
            Left            =   1365
            MaxLength       =   10
            TabIndex        =   7
            Tag             =   "Fecha Entrada|F|N|||scarep|fecrepar|dd/mm/yyyy|N|"
            Text            =   "Text1"
            Top             =   1140
            Width           =   1080
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   5
            Left            =   2030
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   56
            Text            =   "Text2"
            Top             =   360
            Width           =   3040
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   5
            Left            =   1365
            MaxLength       =   4
            TabIndex        =   5
            Tag             =   "Operador|N|N|0|9999|scarep|codtraba|0000|N|"
            Text            =   "Te"
            Top             =   360
            Width           =   645
         End
         Begin VB.Label Label1 
            Caption         =   "Nº Mantenim."
            Height          =   255
            Index           =   0
            Left            =   2880
            TabIndex        =   63
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Ult. Repar."
            Height          =   255
            Index           =   3
            Left            =   2880
            TabIndex        =   62
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Fin Garantia"
            Height          =   255
            Index           =   4
            Left            =   2880
            TabIndex        =   61
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Nº Reparac."
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   60
            Top             =   780
            Width           =   975
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   1
            Left            =   1095
            Picture         =   "frmRepEntReparaciones.frx":0D99
            ToolTipText     =   "Buscar fecha"
            Top             =   1500
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Fec. Entrada"
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   1140
            Width           =   975
         End
         Begin VB.Label Label9 
            Caption         =   "Fec. Repar."
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   58
            Top             =   1500
            Width           =   975
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   0
            Left            =   1095
            Picture         =   "frmRepEntReparaciones.frx":0E24
            ToolTipText     =   "Buscar fecha"
            Top             =   1140
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Operador"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   57
            Top             =   360
            Width           =   735
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   2
            Left            =   1035
            ToolTipText     =   "Buscar trabajador"
            Top             =   375
            Width           =   240
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   5355
         Left            =   -74640
         TabIndex        =   109
         Top             =   600
         Visible         =   0   'False
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   9446
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
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
         Index           =   8
         Left            =   -72600
         Picture         =   "frmRepEntReparaciones.frx":0EAF
         Top             =   2640
         Width           =   240
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   16
         Left            =   -74640
         TabIndex        =   115
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label Label9 
         Caption         =   "Observaciones"
         Height          =   195
         Index           =   5
         Left            =   -73680
         TabIndex        =   114
         Top             =   5160
         Width           =   1155
      End
      Begin VB.Label Label9 
         Caption         =   "Fec.entrega"
         Height          =   255
         Index           =   4
         Left            =   -73680
         TabIndex        =   113
         Top             =   4680
         Width           =   975
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   5
         Left            =   -72600
         Picture         =   "frmRepEntReparaciones.frx":0FB1
         ToolTipText     =   "Buscar fecha"
         Top             =   4680
         Width           =   240
      End
      Begin VB.Label Label11 
         Caption         =   "Imp. reparación"
         Height          =   195
         Index           =   1
         Left            =   -73680
         TabIndex        =   112
         Top             =   4110
         Width           =   1095
      End
      Begin VB.Label Label12 
         Height          =   255
         Index           =   1
         Left            =   -67080
         TabIndex        =   111
         Top             =   1590
         Width           =   255
      End
      Begin VB.Label Label12 
         Caption         =   "Aceptado"
         Height          =   255
         Index           =   0
         Left            =   -67080
         TabIndex        =   110
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         X1              =   -71400
         X2              =   -63720
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   -73200
         X2              =   -63600
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label1 
         Caption         =   "Presupuesto"
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
         Height          =   255
         Index           =   12
         Left            =   -74760
         TabIndex        =   108
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Importe 1º"
         Height          =   255
         Index           =   0
         Left            =   -73680
         TabIndex        =   107
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Importe 2º"
         Height          =   255
         Index           =   1
         Left            =   -73680
         TabIndex        =   106
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Fecha presupuesto"
         Height          =   255
         Index           =   1
         Left            =   -70440
         TabIndex        =   105
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   -68880
         Picture         =   "frmRepEntReparaciones.frx":103C
         ToolTipText     =   "Buscar fecha"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Fecha aprobación"
         Height          =   195
         Index           =   2
         Left            =   -70440
         TabIndex        =   104
         Top             =   1590
         Width           =   1290
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   -68880
         Picture         =   "frmRepEntReparaciones.frx":10C7
         ToolTipText     =   "Buscar fecha"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Servicio de asistencia técnica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   13
         Left            =   -74760
         TabIndex        =   103
         Top             =   2160
         Width           =   3135
      End
      Begin VB.Label Label9 
         Caption         =   "Fec.envio"
         Height          =   255
         Index           =   3
         Left            =   -73680
         TabIndex        =   102
         Top             =   3120
         Width           =   975
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   4
         Left            =   -72600
         Picture         =   "frmRepEntReparaciones.frx":1152
         ToolTipText     =   "Buscar fecha"
         Top             =   3120
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Servicio SAT"
         Height          =   255
         Index           =   2
         Left            =   -73680
         TabIndex        =   101
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Nº Resguardo"
         Height          =   195
         Index           =   3
         Left            =   -73680
         TabIndex        =   100
         Top             =   3600
         Width           =   1005
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   50
      Top             =   410
      Width           =   11535
      Begin VB.CheckBox chkPresupuesto 
         Caption         =   "Presupuesto"
         Enabled         =   0   'False
         Height          =   195
         Left            =   9720
         TabIndex        =   2
         Tag             =   "Presupuesto|N|N|||scarep|presupue||N|"
         Top             =   200
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   3915
         MaxLength       =   16
         TabIndex        =   1
         Tag             =   "Cod. Artículo|T|N|||scarep|codartic||N|"
         Text            =   "Text1"
         Top             =   200
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   1
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   52
         Text            =   "Text2"
         Top             =   200
         Width           =   3885
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   0
         Tag             =   "Nº Serie|T|N|||scarep|numserie||N|"
         Text            =   "Text1"
         Top             =   200
         Width           =   1590
      End
      Begin VB.Label Label5 
         Caption         =   "Artículo"
         Height          =   255
         Left            =   3000
         TabIndex        =   53
         Top             =   200
         Width           =   615
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   3645
         ToolTipText     =   "Buscar artículo"
         Top             =   200
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Nº Serie"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   200
         Width           =   615
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   795
         Picture         =   "frmRepEntReparaciones.frx":11DD
         ToolTipText     =   "Buscar Nº Serie"
         Top             =   200
         Width           =   240
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   16
      Left            =   2640
      Locked          =   -1  'True
      MaxLength       =   60
      TabIndex        =   48
      Text            =   "Text2 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwqa"
      Top             =   7740
      Visible         =   0   'False
      Width           =   6765
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9480
      TabIndex        =   44
      Top             =   7710
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10560
      TabIndex        =   45
      Top             =   7710
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   10560
      TabIndex        =   30
      Top             =   7710
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   120
      TabIndex        =   34
      Top             =   7560
      Width           =   2295
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   240
         TabIndex        =   35
         Top             =   180
         Width           =   1875
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   11760
      _ExtentX        =   20743
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
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver Todos"
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
            Object.ToolTipText     =   "Lineas"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Confirmar Reparación"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
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
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   315
         Left            =   6120
         TabIndex        =   33
         Top             =   0
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   8760
      Top             =   3360
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
      Left            =   9960
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
   Begin VB.Label Label1 
      Caption         =   "Ampliación Línea"
      Height          =   255
      Index           =   35
      Left            =   2640
      TabIndex        =   49
      Top             =   7560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "Cargando datos ........."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   32
      Top             =   7680
      Visible         =   0   'False
      Width           =   3495
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
Attribute VB_Name = "frmRepEntReparaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public ControlRep As Boolean 'Para saber si se llama en el menu ppal desde
                             'Mantenimiento de Reparaciones o desde Control de Reparaciones
Public EsHistorico As Boolean 'Si es true abrir el formulario con la tabla de
                              'de historico schrep, y solo en modo de consulta
Public EntradaEquipo As String 'SI desde avisos le han dado a meter equipo.


Private WithEvents frmB As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmA As frmAlmArticulos  'Form Mantenimiento Articulos
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmAlm As frmAlmAlPropios   'Form Almacenes Propios
Attribute frmAlm.VB_VarHelpID = -1
Private WithEvents frmCli As frmFacClientes 'Form Mantenimiento Clientesº
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmNSeries As frmRepNumSerie 'Form Mantenimiento Nº Series
Attribute frmNSeries.VB_VarHelpID = -1
Private WithEvents frmTraba As frmAdmTrabajadores  'Form Mantenimiento Trabajadores
Attribute frmTraba.VB_VarHelpID = -1
Private WithEvents frmMoti As frmRepMotivosPend  'Form Mantenimiento Motivos Ptes. Rep.
Attribute frmMoti.VB_VarHelpID = -1

Private WithEvents frmTpAve As frmtipave
Attribute frmTpAve.VB_VarHelpID = -1
Private WithEvents frmSAT   As frmManSat
Attribute frmSAT.VB_VarHelpID = -1
Private WithEvents frmTraRea As frmManTraReali
Attribute frmTraRea.VB_VarHelpID = -1

Private WithEvents frmList As frmListadoPed 'Listados para pasar de Pedido -> Albaran
Attribute frmList.VB_VarHelpID = -1


Private Modo As Byte
Private ModoAnterior As Byte

Private ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Private HaDevueltoDatos As Boolean


Dim NombreTabla As String 'Nombre de la Tabla Cabecera
Dim NomTablaLineas As String 'Nombre de la Tabla de lineas

Dim Ordenacion As String
Dim kCampo As Integer

Dim btnAnyadir As Byte
'Variable que indica el número del Boton  Anyadir en la Toolbar1
Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1

Dim PrimeraVez As Boolean
Dim PrimeraVezForm As Boolean
Dim EsCabecera As Boolean
'Para saber en MandaBusquedaPrevia si busca en la tabla sserie o en la tabla sdirec

Dim CodTipoMov As String
'Codigo tipo de movimiento en función del valor en la tabla de parámetros: stipom

Dim CadenaConsulta As String
Dim CadenaSQL As String 'Para crear consulta de Generar Albaran a partir del Pedido
Dim CadenaSQLHco As String
Dim ImprimeAlb As Boolean 'Para saber cuando vuelve de Generar ALbaran si se ha solicitado Imprimir Albaran o no
Dim FechaAlb As String

Dim PorCaja As Boolean
'Para Saber si se ha salido con precio caja y hay que calcular el importe de la
'linea aplicando el precio de la caja. Si PorCaja=false se aplicaca el precio de unidad

Dim Precio As String 'Precio de la linea de Articulo
Dim Indice As Byte
'Dim PrimeraVez As Boolean



Private Sub Check1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkPresupuesto_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub chkPresupuesto_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub chkVistaPrevia_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub cmdAceptar_Click()
Dim PrimeraLin As Boolean 'Si se inserta la primera linea no esta creado el datagrid1 entonces llamar
                          ' a DataGrid, sino llamar solo a DataGrid2
    On Error GoTo Error1
    
    Screen.MousePointer = vbHourglass
    Select Case Modo
        Case 1 'BUSCAR
            HacerBusqueda
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarCabecera Then
                    If EntradaEquipo <> "" Then
                        'Viene de entrada equipo
                        CadenaDesdeOtroForm = "OK"
                        Unload Me
                        Exit Sub
                    End If
                End If
            End If
        Case 4 'MODIFICAR
            If DatosOk Then
                 'El campo numaviso lo tengo que dejar con el valor que tiene
                 Text1(15).Text = DBLet(Me.Data1.Recordset!numaviso, "T")
                 If ModificaDesdeFormulario(Me, 1) Then
                     TerminaBloquear
                     PosicionarData
                 End If
                 'Vuelvo a
                 'Mostrar SOLO el numero de aviso, no la fecha de donde venia
                 If Me.Text1(15).Text <> "" Then Text1(15).Text = RecuperaValor(Text1(15).Text, 1)

             End If
             
        Case 5 'InsertarModificar LINEA
            'Actualizar el registro en la tabla de lineas 'slirep'
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
    End Select
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
            Indice = 20
            Set frmA = New frmAlmArticulos
            frmA.DatosADevolverBusqueda2 = "@1@" 'Poner en Modo Busqueda
            frmA.Show vbModal
            Set frmA = Nothing
    End Select
    PonerFoco txtAux(Index)
End Sub


Private Sub cmdCancelar_Click()
    On Error GoTo ECancelar

    Select Case Modo
        Case 1 'BUSCAR
            LimpiarCampos
            PonerModo 0
        Case 3 'INSERTAR
            If ModoAnterior = 0 Then
                LimpiarCampos
                PonerModo 0
            Else
                PonerModo 2
                PonerCampos
            End If
                
        Case 4  'MODIFICAR
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
            
         Case 5 'LINEAS Detalle
            TerminaBloquear
            CargaTxtAux False, False
            If ModificaLineas = 1 Then 'INSERTAR
                DataGrid1.AllowAddNew = False
                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
            End If
            BloquearTxt Text2(16), True
            ModificaLineas = 0
            PonerBotonCabecera True
            Me.DataGrid1.Enabled = True
    End Select
    PonerFoco Text1(0)
    
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdRegresar_Click()
'Este es el boton de cabecera

    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
        PonerModo 2
        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        If DataGrid1.Row >= 0 Then
            DeseleccionaGrid DataGrid1
            DataGrid1.Bookmark = 1
        End If
        
        'DataGrid1.visible = False
        

    End If
End Sub


Private Sub HabilitarFrames(b As Boolean)
    Me.Frame3.Enabled = Not b
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Modo = 5 Then 'Poner el valor al camp ampliacion linea '5: modo lineas
        If Not Data2.Recordset.EOF And ModificaLineas <> 1 Then '1: Insertar
            'Poner descripcion de ampliacion lineas
            Text2(16).Text = DevuelveDesdeBDNew(conAri, NomTablaLineas, "ampliaci", "numrepar", Text1(2).Text, "N", , "numlinea", Data2.Recordset!numlinea, "N")
        Else
            Text2(16).Text = ""
        End If
    End If
End Sub

Private Sub Form_Activate()
    If PrimeraVezForm Then
        PrimeraVezForm = False
        DoEvents
        Screen.MousePointer = vbHourglass
        '--------------------------------
        
        
        
        If ControlRep Then
            'Cargamos el DATAç
            DataGrid1.visible = True
            CargaGrid DataGrid1, Data2, False
        End If
        
        CargaDatosAviso

    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    PrimeraVezForm = True

    'Icono del formulario
    Me.Icon = frmppal.Icon
    
    'Icono de busqueda
    For kCampo = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(kCampo).Picture = frmppal.imgListComun.ListImages(19).Picture
    Next kCampo
    

    'ICONOS de La toolbar
    btnAnyadir = 5
    btnPrimero = 17 'Posicion del Boton Primero en la toolbar (+ 3 siguientes)
    With Toolbar1
        .ImageList = frmppal.imgListComun
        'ASignamos botones
        .Buttons(1).Image = 1   'Buscar
        .Buttons(2).Image = 2 'Ver Todos
        .Buttons(5).Image = 3 'Añadir
        .Buttons(6).Image = 4 'Modificar
        .Buttons(7).Image = 5 'Eliminar
        .Buttons(10).Image = 10 'Mto Lineas
        .Buttons(11).Image = 26 'Confirmar Reparación
        .Buttons(12).Image = 16 'Imprimir
        .Buttons(14).Image = 15 'Salir
        .Buttons(btnPrimero).Image = 6 'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Ultimo
    End With
    
    LimpiarCampos   'Limpia los campos TextBox
    DataGrid1.visible = False

    'Ocultar los Textos de Reparacion si no en Control de Rep
    Label1(6).visible = ControlRep
    Text1(12).visible = ControlRep
    Text1(13).visible = ControlRep
    Text1(14).visible = ControlRep
    
    
    'Trabajo realizado  si es control reparacion o en HCO
    kCampo = 0
    If ControlRep Or EsHistorico Then kCampo = 1
    Label1(15).visible = (kCampo = 1)
    Me.imgBuscar(7).visible = (kCampo = 1)
    Text1(24).visible = (kCampo = 1)
    Text2(24).visible = (kCampo = 1)
    
    'La solapa de las lineas
    SSTab1.TabVisible(2) = ControlRep
    


    'Si es Histórico no aparece codmotre
    Label1(5).visible = Not EsHistorico
    imgBuscar(5).visible = Not EsHistorico
    Text1(11).visible = Not EsHistorico
    Text2(11).visible = Not EsHistorico
    
    'Si es hco no tiene el dato de numaviso
    Text1(15).visible = Not EsHistorico
    Label1(11).visible = Not EsHistorico
    
    'Si es Histórico no aparece fecentre 'Fecha Prev. entrega Repar
    'David: Hemos metido el campo en la BD
    'Label4.visible = Not EsHistorico
    'imgFecha(1).visible = Not EsHistorico
    'Text1(4).visible = Not EsHistorico
    
    'Si es Histórico no aparece Presupuesto
    Me.chkPresupuesto.visible = Not EsHistorico

    'Campos que solo aparecen en el Histórico
    Text2(0).visible = EsHistorico
    Text2(14).visible = EsHistorico
    Text2(15).visible = EsHistorico
    Label1(8).visible = EsHistorico
    Label1(22).visible = EsHistorico
    Label1(10).visible = EsHistorico
    imgVerAlbaran.visible = EsHistorico
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    CodTipoMov = "REP"
    PrimeraVez = True
    
    If Not EsHistorico Then
        NombreTabla = "scarep" 'Tabla Cabecera Reparaciones
        NomTablaLineas = "slirep" 'Tabla Lineas Reparaciones
        Me.Caption = "Reparaciones"
    Else
        NombreTabla = "schrep"
        NomTablaLineas = "slhrep"
        CargarTagsHco Me, "scarep", NombreTabla
        'Leer estos datos almacenados en la tabla del Historico
        Text2(1).Tag = "Cod. Artículo|T|N|||schrep|nomartic||N|"
        Text2(2).Tag = "Ult. Reparac|F|S|||schrep|ultrepar|dd/mm/yyyy|N|"
        Text2(3).Tag = "Fin Garantia|F|N|||schrep|fingaran|dd/mm/yyyy|N|"
        Text2(4).Tag = "Nº Mantenim.|N|S|||schrep|nummante||N|"
        Me.Caption = "Histórico Reparaciones"
        
        'Datos Albaran
        Label1(10).Left = 240
        Text2(15).Left = 240
        Label1(8).Left = 1240
        Text2(0).Left = 1240
        Label1(22).Left = 1980
        Text2(14).Left = 1980
        imgVerAlbaran.Top = Text2(15).Top + 30
        imgVerAlbaran.Left = 1980 + Text2(14).Width + 120
    End If
    
    Ordenacion = " ORDER BY numrepar "
    CadenaConsulta = "Select * from " & NombreTabla & " WHERE numrepar = -1" 'No recupera datos
    Data1.ConnectionString = Conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1
        Text1(0).BackColor = vbYellow
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento Articulos
    If Indice = 1 Then
        Text1(1).Text = RecuperaValor(CadenaSeleccion, 1)
        Text2(1).Text = RecuperaValor(CadenaSeleccion, 2)
    Else
        txtAux(1).Text = RecuperaValor(CadenaSeleccion, 1)
        txtAux(2).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub


Private Sub frmAlm_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Almacenes Propios
    txtAux(0).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Almacen
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
'Formulario para Busqueda
Dim cadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        If EsCabecera Then 'Llama desde VerTodos del Form
            'Estamos en Cabecera
            'Recupera todo el registro de Nº Serie
            'Sabemos que campos son los que nos devuelve
            'Creamos una cadena consulta y ponemos los datos
            cadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
            cadB = Aux
            Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
            cadB = cadB & " and " & Aux
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
        Else  'Llama desde Prismatico Direcciones/Departamentos
            Text1(7).Text = Format(RecuperaValor(CadenaDevuelta, 1), "000")
            Text2(7).Text = RecuperaValor(CadenaDevuelta, 2)
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento Clientes
    Text1(6).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    Text2(6).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    Indice = Val(Me.imgFecha(0).Tag)
    Text1(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmList_DatoSeleccionado(CadenaSeleccion As String)
'Cuando pasa de Reparacion -> Albaran
'Aqui devuelve los valores que se introducen desde el Form de Listado de Pedido
'para generar el Albaran
Dim vSQL As String
Dim RS As ADODB.Recordset
Dim cad1 As String, cad2 As String

    'Seleccionar algunos campos del Cliente
    vSQL = "Select proclien, codagent, codforpa, dtoppago, dtognral, tipofact "
    vSQL = vSQL & " FROM sclien WHERE codclien=" & Text1(6).Text
    Set RS = New ADODB.Recordset
    RS.Open vSQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    cad1 = RecuperaValor(CadenaSeleccion, 1) 'trab. albaran
    cad2 = RecuperaValor(CadenaSeleccion, 2) 'trab. prepara material
    FechaAlb = RecuperaValor(CadenaSeleccion, 4)

    'Construimos parte de la SQL para insertar en tabla de Albaranes
    vSQL = ""
    vSQL = " '" & Format(FechaAlb, FormatoFecha) & "', " 'Fecha Albaran
    vSQL = vSQL & "0, " 'facturar s/n
    vSQL = vSQL & Text1(6).Text & ", " & DBSet(Text2(6).Text, "T") & ", " 'nomclien
    vSQL = vSQL & DBSet(Text2(10).Text, "T") & ", " & DBSet(Text2(12).Text, "T") & ", " & DBSet(Text2(13).Text, "T") & ", " 'domclien, codpobla, pobclien
    vSQL = vSQL & DBSet(RS!proclien, "T") & ", '" & Text2(8).Text & "', '" & Text2(9).Text & "', " 'proclien, nifclien, telclien "
    vSQL = vSQL & DBSet(Text1(7).Text, "N", "S") & ", " & DBSet(Text2(7).Text, "T") & ", " ' nomdirec
    vSQL = vSQL & ValorNulo & ", " & cad1 & ", "  'referenc, codtraba(ped), "
    vSQL = vSQL & DBSet(Text1(5).Text, "N", "S") & ", " 'Trabajador de pedido
    vSQL = vSQL & cad2 & ", " 'Material Preparado por
    vSQL = vSQL & DBSet(RS!codagent, "N") & ", " & DBSet(RS!codforpa, "N") & ", " '"codagent, codforpa, "
    vSQL = vSQL & RecuperaValor(CadenaSeleccion, 3) & ", " 'Cod Envio
    vSQL = vSQL & DBSet(RS!DtoPPago, "N") & ", " & DBLet(RS!DtoGnral, "N") & ", " & DBLet(RS!TipoFact, "N") & ", " '" '"dtoppago, dtognral, tipofact,
    
    'ANTIGUAS OBSERVACIONES. 19 JUN 07
    'vSQL = vSQL & DBSet(Text1(8).Text, "T") & ", " & DBSet(Text1(9).Text, "T") & ", " & DBSet(Text1(10).Text, "T") & ", " 'observa01, observa02, observa03,
    'vSQL = vSQL & DBSet(Text1(14).Text, "T") & ", " & DBSet(Text1(13).Text, "T") & ", " 'observa04, observa05, "
    
    'AHORA
    vSQL = vSQL & DBSet(Text1(14).Text, "T") & ", " & DBSet(Text1(13).Text, "T") & ", " & DBSet(Text1(12).Text, "T") & ", " 'observa01, observa02, observa03,
    vSQL = vSQL & DBSet("Número serie: " & Text1(0).Text, "T") & ", " & DBSet("Artículo: " & Text1(1).Text & " - " & Text2(1).Text, "T") & ", " 'observa04, observa05, "
    
    vSQL = vSQL & ValorNulo & ", " & ValorNulo & ", "  'Nº Oferta, fecha de la Oferta
    vSQL = vSQL & Text1(2).Text & ", '"  'Nº Pedido
    vSQL = vSQL & Format(Text1(3).Text, FormatoFecha) & "', " & ValorNulo 'Fecha Pedido, Semana entrega
    'vSQL = vSQL & Text1(18).Text 'Semana entrega Pedido
    CadenaSQL = vSQL
    
    RS.Close
    Set RS = Nothing
    
    CadenaSQLHco = cad1 & ", " & cad2 & ", material, tipoaver, motivore, textore1, textore2, textore3 "
    
    'Se almacena aqui si el usuario quiere imprimir el Albaran tras generarlo
    ImprimeAlb = CBool(RecuperaValor(CadenaSeleccion, 5))
End Sub


Private Sub frmMoti_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento Motivos Pendientes Rep.
    Text1(11).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    Text2(11).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmNSeries_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento Nº Serie
    Text1(0).Text = RecuperaValor(CadenaSeleccion, 1) 'num serie
    Text1(1).Text = RecuperaValor(CadenaSeleccion, 2) 'cod artic
    Text2(1).Text = RecuperaValor(CadenaSeleccion, 3) ' desc artic
    'DAVID.
    'Si me va a devolver VACIO no lo borro por si , y solo si, viene de los avisos
    If EntradaEquipo = "" Then
        'mantenimiento normal
        Text1(6).Text = Format(RecuperaValor(CadenaSeleccion, 4), "000000") 'cod cliente
    Else
        If RecuperaValor(CadenaSeleccion, 4) = "" Then
            'NO HACEMOS NADA. NO vaciamos el campo codcliente
        Else
            Text1(6).Text = Format(RecuperaValor(CadenaSeleccion, 4), "000000") 'cod cliente
        End If
    End If
End Sub

Private Sub frmSAT_DatoSeleccionado(CadenaSeleccion As String)
    PonValoresDatoSeleccionado 21, CadenaSeleccion
End Sub

Private Sub frmTpAve_DatoSeleccionado(CadenaSeleccion As String)
    PonValoresDatoSeleccionado 23, CadenaSeleccion
End Sub

Private Sub frmTraba_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento Trabajadores
    'Text1(5).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    'Text2(5).Text = RecuperaValor(CadenaSeleccion, 2)
    PonValoresDatoSeleccionado 5, CadenaSeleccion
End Sub

Private Sub PonValoresDatoSeleccionado(Indice As Integer, CadenaSeleccion As String)
    Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmTraRea_DatoSeleccionado(CadenaSeleccion As String)
PonValoresDatoSeleccionado 24, CadenaSeleccion
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim cadMen As String

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Nº Serie
            Set frmNSeries = New frmRepNumSerie
            frmNSeries.DatosADevolverBusqueda = "0"
            frmNSeries.Show vbModal
            Set frmNSeries = Nothing
            Indice = 0
            
        Case 1 'Codigo Articulo
            Indice = 1
            Set frmA = New frmAlmArticulos
            frmA.DatosADevolverBusqueda2 = "@1@" 'Abrir en Modo busqueda
            frmA.Show vbModal
            Set frmA = Nothing
        
        Case 2 'Cod. Trabajador (Operador)
            Set frmTraba = New frmAdmTrabajadores
            frmTraba.DatosADevolverBusqueda = "0"
            frmTraba.Show vbModal
            Set frmTraba = Nothing
            Indice = 5
        
        Case 3 'Cod. Cliente
            Set frmCli = New frmFacClientes
            frmCli.DatosADevolverBusqueda = "0"
            frmCli.Show vbModal
            Set frmCli = Nothing
            Indice = 6
            
        Case 4 'Direc/Dpto del Cliente
            'Mostrar las Direc. o Dptos del cliente seleccionado
            If Trim(Text1(6).Text) = "" Then
               If vParamAplic.Departamento Then
                   cadMen = "Departamentos"
               Else
                   cadMen = "Direcciones"
               End If
               MsgBox "Debe seleccionar un cliente para mostrar sus " & cadMen & ".", vbInformation
               Screen.MousePointer = vbDefault
               Exit Sub
            Else
               EsCabecera = False
               MandaBusquedaPrevia " codclien= " & Val(Text1(6).Text)
               Indice = 7
            End If
             
        Case 5 'Cod. Motivo Pendiente Rep.
            Set frmMoti = New frmRepMotivosPend
            frmMoti.DatosADevolverBusqueda = "0"
            frmMoti.Show vbModal
            Set frmMoti = Nothing
            Indice = 11
            

        Case 6
            Set frmTpAve = New frmtipave
            frmTpAve.DatosADevolverBusqueda = "0"
            frmTpAve.Show vbModal
            Set frmTpAve = Nothing
            Indice = 10 'Para que ponga el foco en el siguiente
        Case 7
            Set frmTraRea = New frmManTraReali
            frmTraRea.DatosADevolverBusqueda = "0"
            frmTraRea.Show vbModal
            Set frmTraRea = Nothing
            Indice = 24 'Para que ponga el foco en el siguiente
        Case 8
            Set frmSAT = New frmManSat
            frmSAT.DatosADevolverBusqueda = "0"
            frmSAT.Show vbModal
            Set frmSAT = Nothing
            Indice = 20 'Para que ponga el foco en el siguiente
    End Select
    
    PonerFoco Text1(Indice)
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFecha_Click(Index As Integer)

   If Modo = 2 Or Modo = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass
   
   Set frmF = New frmCal
   frmF.Fecha = Now
   
   Select Case Index
        Case 0: Indice = 3 'Fecha Reparacion
        Case 1: Indice = 4 'Fecha Entrega
        Case 2 To 4
            Indice = Index + 16
        Case 5
            Indice = 26
   End Select
   imgFecha(0).Tag = Indice

   PonerFormatoFecha Text1(Indice)
   If Text1(Indice).Text <> "" Then frmF.Fecha = CDate(Text1(Indice).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco Text1(Indice)
End Sub


Private Sub imgVerAlbaran_Click()
    If Text2(15).Text <> "" Then
    
    
        CadenaSQL = DevuelveDesdeBDNew(conAri, "scaalb", "numalbar", "codtipom", Text2(0).Text, "T", , "numalbar", Text2(15).Text, "N")
        If CadenaSQL <> "" Then 'existe el Albaran
             With frmFacEntAlbaranes
                .hcoCodMovim = Format(Text2(15).Text, , "0000000")
                .hcoCodTipoM = Text2(0).Text ' Comprobar esto
                .RecuperarFactu = False
                .Show vbModal
            End With
        Else 'No existe en albaran, abrir Historico Factura
            With frmFacHcoFacturas
                .hcoCodMovim = Format(Text2(15).Text, , "0000000")
                .hcoCodTipoM = Text2(0).Text
                .hcoFechaMov = CDate(Text2(14).Text)
                .Show vbModal
            End With
        End If
    End If
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub


Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar Linea
        BotonEliminarLinea
    Else
        BotonEliminar
    End If
End Sub


Private Sub mnModificar_Click()
    If Modo = 5 Then 'Modifica linea
        BotonModificarLinea
    Else
        If BLOQUEADesdeFormulario(Me) Then BotonModificar
    End If
End Sub


Private Sub mnNuevo_Click()
    If Modo = 5 Then 'Añadir linea
        BotonAnyadirLinea
    Else
        BotonAnyadir
    End If
End Sub


Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    If (Modo = 5) Then 'Modo 5: Mto Lineas
        '1:Insertar linea, 2: Modificar
        If (ModificaLineas = 1 Or ModificaLineas = 2) Then cmdCancelar_Click
        cmdRegresar_Click
        Exit Sub
    End If
    CadenaDesdeOtroForm = ""
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Index = 0 And KeyCode = 38 Then Exit Sub 'Primer campo, fecla arriba
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 27 Then KEYpress KeyAscii
End Sub


Private Sub Text1_LostFocus(Index As Integer)
Dim Devuelve As String
Dim totArtic As Integer

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0 'Nº serie
            If Text1(Index).Text = "" Then Exit Sub
            If Modo = 1 Or Modo = 4 Then Exit Sub
            totArtic = ArticulosDelNSerie(Text1(Index).Text)
            If totArtic = 0 Then
                'No se encontro ningun registro en la tabla sserie para ese valor de Nº de serie
                MsgBox "No existe el Nº de Serie: " & Text1(Index).Text
                Exit Sub
            ElseIf totArtic = 1 Then
                'Solo hay un articulo que tiene ese nº de serie: Recuperar datos de
                'la tabla sserie
                Text1(1).Text = DevuelveDesdeBDNew(conAri, "sserie", "codartic", "numserie", Text1(0).Text, "T")
                Text2(1).Text = PonerNombreDeCod(Text1(1), conAri, "sartic", "nomartic")
                CargarDatosNSerie Text1(0).Text, Text1(1).Text
                ComprobarReparaciones Modo, Text1(0).Text, Text1(1).Text
            Else
                'hay varios artículos que tienen este nº de serie, hasta que no se
                'seleccione el codartic no se pueden recuperar los datos de la tabla sserie
                If Text1(1).Text <> "" Then
                    CargarDatosNSerie Text1(0).Text, Text1(1).Text
                    ComprobarReparaciones Modo, Text1(0).Text, Text1(1).Text
                Else
                    MsgBox "Hay varios artículos con ese Nº de Serie, seleccione uno.", vbInformation
                End If
            End If
            
        Case 1 'Codigo Articulo
            'Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sartic", "nomartic")
            PonerDatosCodigoDescripcion Index
            
        'Fechas Reparacion, Fecha Entrega      :  Fecha presupu,aprobacion   : SAT: envio entrega
        Case 3, 4, 18, 19, 20, 26
            If Text1(Index).Text = "" Then Exit Sub
            PonerFormatoFecha Text1(Index)

            'Comprobar que Fecha Rep. es posterior a la de Entrada
            If Index <= 4 Then
                If Not EsFechaIgualPosterior(Text1(3).Text, Text1(4).Text, True, "La Fecha de Reparación debe ser posterior a la Fecha de Entrada.") Then
                    Text1(Index).Text = ""
                    PonerFoco Text1(Index)
                    Exit Sub
                End If
            End If
                
        Case 5 'Cod Trabajador
'            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "straba", "nomtraba", "codtraba")
            PonerDatosCodigoDescripcion Index
        Case 6 'Cliente
            If PonerFormatoEntero(Text1(Index)) Then
                If Modo = 1 Then 'Modo=1 Busqueda
'                    Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sclien", "nomclien")
                    PonerDatosCodigoDescripcion Index
                Else 'Insertando
                    PonerDatosCliente Text1(Index).Text, False
                End If
            Else
                LimpiarDatosCliente
            End If
            
        Case 7 'Direc/dpto del cliente
            If Text1(Index).Text = "" Then
                Text2(Index).Text = ""
                Exit Sub
            End If
            If Text1(6).Text = "" Then
                MsgBox "Debe seleccionar un cliente.", vbInformation
                Text1(Index).Text = ""
                PonerFoco Text1(6)
                Exit Sub
            End If
            Text1(Index).Text = Format(Text1(Index).Text, "000")
            
            'Comprobar que el cliente seleccionado tiene esa direccion o dpto
            Devuelve = DevuelveDesdeBDNew(conAri, "sdirec", "nomdirec", "codclien", Text1(6).Text, "N", , "coddirec", Text1(7).Text, "N")
            Text2(Index).Text = Devuelve 'Nombre direc. o dpto
            If Devuelve = "" Then 'No existe el dpto
                If vParamAplic.Departamento Then
                    Devuelve = " el Departamento "
                Else
                    Devuelve = " la Dirección "
                End If
                Devuelve = "No existe" & Devuelve & Text1(Index).Text & " para el cliente: "
                Devuelve = Devuelve & Text1(6).Text & " - " & Text2(6).Text
                MsgBox Devuelve, vbInformation
                Text1(Index).Text = ""
                PonerFoco Text1(Index)
            End If
            
        Case 11 'Motivo pendiente reparacion
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "smotre", "nommotre", "codmotre")
            
        Case 16, 17, 25
            PonerFormatoDecimal Text1(Index), 1 'Tipo 2: Decimal(10,4)
        'Case 21
        '    'Servicio ASISTENCIA TECNICA
        '    Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "smansat", "nomsat", "codsat")
        'Case 23
        '    'Tipo averia
        '    'Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "stipave", "nomave", "codave")
        '
        'Case 24
        '    'Trabajao realizado
        '    'Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "smantr", "nomtrabajo", "codtrabajo")
            
            
        Case 21, 23, 24
            PonerDatosCodigoDescripcion Index
    End Select
End Sub



Private Sub PonerDatosCodigoDescripcion(Index As Integer)


    Select Case Index
        Case 1
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sartic", "nomartic")
            
        Case 5
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "straba", "nomtraba", "codtraba")
            
        Case 6
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sclien", "nomclien")
            
        Case 21
            'Servicio ASISTENCIA TECNICA
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "smansat", "nomsat", "codsat")
            
        Case 23
            'Tipo averia
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "stipave", "nomave", "codave")
            
        Case 24
            'Trabajao realizado
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "smantr", "nomtrabajo", "codtrabajo")

    
    End Select
End Sub



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

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 1 'Busqueda
            mnBuscar_Click
        Case 2 'Ver Todos
            mnVerTodos_Click
        Case 5 'Nuevo
            mnNuevo_Click
        Case 6  'Modificar
            mnModificar_Click
        Case 7 'Eliminar
            mnEliminar_Click
        Case 10 'Mto Lineas
             BotonMtoLineas
        Case 11 'Confirmar Reparación
             BotonConfirmarRep
        Case 12 'Imprimir
            If (Not ControlRep) And (Not EsHistorico) Then BotonImprimir (62)
        Case 14  'Salir
             mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas de Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then
        CadenaDesdeOtroForm = ""
        Unload Me
    End If
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim I As Byte
Dim b As Boolean
Dim NumReg As Byte

    On Error GoTo EPonerModo

    'Actualiza Iconos Insertar,Modificar,Eliminar
    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    'Visualizar flechas de desplazamiento en la toolbar si modo=2
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, (Modo = 2), NumReg
    
        
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = (Modo = 2)
    Else
        cmdRegresar.visible = False
    End If
        
    '-------------------------------------------
    'Bloquear todos los Text Box que se llamen Text1
    BloquearText1 Me, Modo
    
    'Nº Reparacion siempre bloqueado, es contador, salvo en Modo=Buscar
    If Modo <> 1 Then BloquearTxt Text1(2), True, True
    
                
       
    '------------------------------------
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    b = b And Modo < 5
    Me.chkPresupuesto.Enabled = b ' (Modo = 3) Or (Modo = 4) 'Insertar o Modificar
    Me.Check1.Enabled = b
    Me.Combo1.Enabled = b
    b = ((Modo = 3 Or Modo = 4) And (ControlRep = False)) Or Modo = 1
'    For i = 0 To Me.imgBuscar.Count - 1
'        Me.imgBuscar(i).Enabled = b
'    Next i
     For I = 0 To Me.imgBuscar.Count - 1
        BloquearImg Me.imgBuscar(I), Not b
    Next I
    Me.imgBuscar(1).Enabled = (Modo = 1)
    'La imagen del TRABAJO REALIZADO no se tiene que mostrar a no ser que haya entrado como reparacion
    Me.imgBuscar(7).visible = Me.imgBuscar(7).visible And ControlRep
    
    For I = 0 To Me.imgFecha.Count - 1
        Me.imgFecha(I).Enabled = b 'Si es insertar o modificar
    Next I
    
    If Modo = 1 Then 'Busqueda
        Text1(1).TabIndex = 1
    Else
        Text1(1).TabIndex = 18
    End If
    
    'Modo Linea de Ofertas
    b = (Modo = 5)
    Me.Label1(35).visible = b
    Me.Text2(16).visible = b
    BloquearTxt Text2(16), True
    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
    
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    
    PonerModoOpcionesMenu   'Activar opciones de menu según Modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
EPonerModo:
    If Err.Number <> 0 Then
        MsgBox Err.Number & ": " & Err.Description, vbExclamation
    End If
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Sub PonerModoOpcionesMenu()
Dim b As Boolean
Dim I As Byte

    'Modo 2. Hay datos y estamos visualizandolos
    b = ((Not ControlRep) Or (ControlRep And Modo = 5)) And (Not EsHistorico)
    Toolbar1.Buttons(5).visible = b
    Me.mnNuevo.visible = b
    Toolbar1.Buttons(7).visible = b
    Me.mnEliminar.visible = b
    Toolbar1.Buttons(6).visible = Not EsHistorico
    Me.mnModificar.visible = Not EsHistorico
    Toolbar1.Buttons(8).visible = Not EsHistorico
    Toolbar1.Buttons(9).visible = Not EsHistorico
    Me.mnBarra2.visible = Not EsHistorico
    
    For I = 10 To 11
        Toolbar1.Buttons(I).visible = ControlRep
    Next I
    
    b = (Modo = 2) Or (Modo = 5 And ModificaLineas = 0)
    'Modificar
    Toolbar1.Buttons(6).Enabled = b
    Me.mnModificar.Enabled = b
    'Insertar
    Toolbar1.Buttons(5).Enabled = (b Or Modo = 0)
    Me.mnNuevo.Enabled = (b Or Modo = 0)
        
    'eliminar
    Toolbar1.Buttons(7).Enabled = b
    Me.mnEliminar.Enabled = b
    
    
    If ControlRep Then
        b = (Modo = 2)
        'Mto Lineas
        Toolbar1.Buttons(10).Enabled = b
        'Confirmación Reparación
        Toolbar1.Buttons(11).Enabled = b
    End If
    
    
    '-------------------------------------
    b = (Modo >= 3) Or Modo = 1
    'Buscar
    Toolbar1.Buttons(1).Enabled = Not b
    Me.mnBuscar.Enabled = Not b
    'Ver Todos
    Toolbar1.Buttons(2).Enabled = Not b
    Me.mnVerTodos.Enabled = Not b
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
    Me.chkPresupuesto.Value = 0
    Me.Check1.Value = 0
    Me.Combo1.ListIndex = -1
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    DesplazamientoData Data1, Index
    PonerCampos
End Sub


Private Sub BotonBuscar()
'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        If ControlRep Then CargaGrid DataGrid1, Data2, False
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
    If chkVistaPrevia.Value = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub BotonAnyadir()
Dim NomTraba As String

    LimpiarCampos 'Vacía los TextBox
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    ModoAnterior = Modo 'Para el botón Cancelar en Modo Insertar
    PonerModo 3
    
    'Bloquear algunos campos
    BloquearTxt Text1(1), True
    
    Text1(3).Text = Format(Now, "dd/mm/yyyy")
    Text1(5).Text = PonerTrabajadorConectado(NomTraba)
    Text2(5).Text = NomTraba
    
    PonerFoco Text1(0)
End Sub


Private Sub BotonModificar()
Dim I As Byte
    'Escondemos el navegador y ponemos Modo Modificar
    PonerModo 4
    
    'Como el campo Nº Repar. es clave primaria, NO se puede modificar
    BloquearTxt Text1(2), True, True
    BloquearTxt Text1(1), True
    
    If ControlRep Then
        Me.chkPresupuesto.Enabled = False
        Text1(0).Locked = True
        For I = 3 To 7
            Text1(I).Locked = True
        Next I
        Me.imgBuscar(5).Enabled = True
        PonerFoco Text1(8)
    Else
        PonerFoco Text1(0)
    End If
    
End Sub


Private Sub BotonModificarLinea()
'Prepara el Form para Modificar una linea de Reparacion (tabla: slirep)
Dim vWhere As String

    On Error GoTo EModificarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub '1= Insertar
    
    If Data2.Recordset.EOF Then Exit Sub
    
    vWhere = Mid(ObtenerWhereCP, 7) & " and numlinea=" & Data2.Recordset!numlinea
    If Not BloqueaRegistro(NomTablaLineas, vWhere) Then Exit Sub
    CargaTxtAux True, False
    ModificaLineas = 2 'Modificar
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR"
    PonerBotonCabecera False
    BloquearTxt Text2(16), False 'Campo Ampliacion Linea
    PonerFoco txtAux(0)
    Me.DataGrid1.Enabled = False
    
EModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub BotonEliminar()
Dim SQL As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
        
    SQL = ""
    SQL = SQL & "Va a Eliminar la Reparación: " & Text1(2).Text & vbCrLf
    SQL = SQL & vbCrLf & "Nº Serie: " & Text1(0).Text
    SQL = SQL & vbCrLf & "Artic. : " & Text1(1).Text & " - " & Text2(1).Text
    SQL = SQL & vbCrLf & vbCrLf & "¿Desea continuar? "
    
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        NumRegElim = Data1.Recordset.AbsolutePosition
        If Not Eliminar Then Exit Sub
        PosicionarDataTrasEliminar
    End If
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Nº Reparación", Err.Description
        Data1.Recordset.CancelUpdate
    End If
End Sub


Private Function Eliminar() As Boolean
Dim SQL As String

    On Error GoTo FinEliminar

    SQL = " WHERE numrepar=" & Data1.Recordset!numrepar
    
    'Eliminar las Lineas
    Conn.Execute "Delete from " & NomTablaLineas & SQL
    
    'Eliminar Cabecera
    Conn.Execute "Delete  from " & NombreTabla & SQL
               
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
        Eliminar = False
    Else
        Eliminar = True
    End If
End Function


Private Sub BotonEliminarLinea()
'Eliminar una linea De la Reparacion. (Tabla: slirep)
Dim SQL As String

    On Error GoTo EEliminarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar

    If Data2.Recordset.EOF Then Exit Sub
            
    ModificaLineas = 3 'Eliminar
    SQL = "¿Seguro que desea eliminar la línea de la Reparación?     "
    SQL = SQL & vbCrLf & "NumLinea:  " & Data2.Recordset!numlinea & vbCrLf
    SQL = SQL & "Almacen:  " & Format(Data2.Recordset!codAlmac, "000")
    SQL = SQL & vbCrLf & "Artículo:  " & Data2.Recordset!codArtic & " - " & Data2.Recordset!NomArtic
    
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        NumRegElim = Data2.Recordset.AbsolutePosition
        SQL = "Delete from " & NomTablaLineas & ObtenerWhereCP
        SQL = SQL & " and numlinea=" & Data2.Recordset!numlinea
        Conn.Execute SQL
        
        ModificaLineas = 0
        CargaGrid2 DataGrid1, Data2
        SituarDataTrasEliminar Data2, NumRegElim
'        CancelaADODC
    End If
    PonerFocoBtn Me.cmdRegresar
    
EEliminarLinea:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Lineas Reparación", Err.Description
End Sub


Private Sub BotonMtoLineas()

    'Por si acaso esta puesto el modo incorrecto
    If EsHistorico Or Not ControlRep Then Exit Sub
    

    SSTab1.Tab = 2

    ModificaLineas = 0
    PonerModo 5
    
 
   
    'Me.DataGrid1.visible = True
    'Esto antes estaba descomentado.  21 Abril de 2008
    'CargaGrid DataGrid1, Data2, True
    
    PonerBotonCabecera True
End Sub


Private Sub BotonConfirmarRep()
'Confirmar Reparacion
Dim b As Boolean
Dim cadMen As String, vWhere As String

    If MsgBox("¿Desea Cerrar la Orden de Reparación y Generar Albaran?", vbQuestion + vbYesNoCancel + vbDefaultButton2) = vbYes Then
        Screen.MousePointer = vbHourglass
        b = SePuedeServirPedido(cadMen)
        If b Then 'Hay suficiente stock
            'Si hay stock generar albaran completo
            GenerarAlbaran
        ElseIf cadMen <> "" Then
            MsgBox cadMen, vbExclamation
        Else
            Screen.MousePointer = vbDefault
            'Si no se puede servir mostrar mensaje detallando y bloquear
            cadMen = "No hay suficiente Stock para servir la Reparación. "
            cadMen = cadMen & vbCrLf & "¿Desea Ver Detalle?"
            If MsgBox(cadMen, vbYesNo, "Contol de Stock") = vbYes Then
                vWhere = " WHERE numrepar = " & Text1(2).Text & " And sfamia.instalac = 0 "
                frmMensajes.cadWhere = vWhere
                frmMensajes.vCampos = NomTablaLineas
                frmMensajes.OpcionMensaje = 2 'Articulos sin Stock
                frmMensajes.Show vbModal
            End If
            Exit Sub
        End If
        'Pedir Datos para el Albaran: Operador, Fecha, Reparado por
        
        Screen.MousePointer = vbDefault
    End If
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean

    DatosOk = False
    b = CompForm(Me, 1)
    If Not b Then Exit Function
    If Modo = 3 Then
        If EntradaEquipo <> "" Then
            If Val(Text1(6).Text) <> RecuperaValor(EntradaEquipo, 3) Then
                MensajeNoCoinciden Text1(6).Text, True
                b = MsgBox(CadenaDesdeOtroForm, vbQuestion + vbYesNo) = vbYes
                CadenaDesdeOtroForm = ""
                If Not b Then Exit Function
            End If
        End If
    End If
    DatosOk = True
End Function



Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim Tabla As String
Dim Titulo As String, Desc As String
Dim selElem As Byte

    'Llamamos a al form
    cad = ""
    If EsCabecera Then
    'Estamos en Modo de Cabeceras
    'Registro de la tabla de cabeceras: sserie
        cad = cad & ParaGrid(Text1(0), 20, "Nº Serie")
        cad = cad & ParaGrid(Text1(1), 25, "Artic.")
        cad = cad & "Desc. Artic.|sartic|nomartic|T||40·"
        cad = cad & ParaGrid(Text1(2), 15, "Num Rep.")
'        cad = cad & "Desc. Tipo|stipar|nomtipar|T||20·"
    
        Tabla = "(" & NombreTabla & " LEFT JOIN sartic ON " & NombreTabla & ".codartic=sartic.codartic" & ")"
'        tabla = tabla & " LEFT JOIN stipar ON " & NombreTabla & ".codtipar=stipar.codtipar"
        If EsHistorico Then
            Titulo = "Histórico Reparaciones"
        Else
            Titulo = "Reparaciones"
        End If
        selElem = 2
   Else
        If vParamAplic.Departamento Then
            Titulo = "Dptos Cliente: "
            Desc = "Dpto."
        Else
            Titulo = "Direc. Cliente: "
            Desc = "Direc."
        End If
        Titulo = Titulo & Text1(6).Text & " - " & Text2(6).Text 'Cod y Desc. Cliente
        cad = cad & "Cod. " & Desc & "|sdirec|coddirec|N||20·"
        cad = cad & "Desc. " & Desc & "|sdirec|nomdirec|T||40·"
        Tabla = "sdirec"
        selElem = 1
    End If
           
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = Tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = Titulo
        frmB.vselElem = selElem
        frmB.vConexionGrid = conAri 'Conexion a BD Ariges
'        frmB.vBuscaPrevia = chkVistaPrevia
        '#
        If Not EsCabecera Then frmB.Label1.FontSize = 11
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
        If HaDevueltoDatos Then
            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            If Modo = 5 Then
'                PonerFoco txtAux(0)
'            Else
                PonerFoco Text1(kCampo)
'            End If
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False)
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then 'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        If Modo = 1 Then
            MsgBox "No hay ningún registro en la tabla " & NombreTabla & " para ese criterio de Búsqueda.", vbInformation
        Else
            MsgBox "No hay ningún registro en la tabla " & NombreTabla & ".", vbInformation
        End If
        Screen.MousePointer = vbDefault
        PonerModo Modo
        PonerFoco Text1(0)
        Exit Sub
    Else
        PonerModo 2
        Data1.Recordset.MoveFirst
        PonerCampos
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
    
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
Dim Devuelve As String
Dim TieneMan As String

    On Error GoTo EPonerCampos
    
    
    If Data1.Recordset.EOF Then Exit Sub
    
    'Por si acaso, como puede ser NULL
    Combo1.ListIndex = -1
    
    PonerCamposForma Me, Data1 'Los Text1
            
    'Poner el nombre del cod. Articulo
    'Text2(1).Text = PonerNombreDeCod(Text1(1), conAri, "sartic", "nomartic")
    'Poner el nombre del Trabajador (Operador)
    'Text2(5).Text = PonerNombreDeCod(Text1(5), conAri, "straba", "nomtraba")
    'Poner el nombre del cod. Cliente
    'Text2(6).Text = PonerNombreDeCod(Text1(6), conAri, "sclien", "nomclien")
    
    PonerDatosCodigoDescripcion 1
    PonerDatosCodigoDescripcion 5
    PonerDatosCodigoDescripcion 6
    PonerDatosCodigoDescripcion 21
    PonerDatosCodigoDescripcion 23
    PonerDatosCodigoDescripcion 24
    
    
    PonerDatosCliente Text1(6).Text
    
    If EsHistorico Then
        'Poner datos Albaran
        Text2(15).Text = DBLet(Me.Data1.Recordset!NumAlbar, "T")
        FormateaCampo Text2(15)
        Text2(14).Text = DBLet(Me.Data1.Recordset!FechaAlb, "F")
        Text2(0).Text = DBLet(Me.Data1.Recordset!codTipoM, "T")
    End If
        
    
    'Poner el nombre del cod. Direc./Dpto
    Devuelve = DevuelveDesdeBDNew(conAri, "sdirec", "nomdirec", "codclien", Text1(6).Text, "N", , "coddirec", Text1(7).Text, "N")
    Text2(7).Text = Devuelve
    If Not EsHistorico Then
        'Poner la fecha fin Garantia y ult. repar
        Devuelve = "ultrepar"
        Text2(3).Text = DevuelveDesdeBDNew(conAri, "sserie", "fingaran", "numserie", Text1(0).Text, "T", Devuelve, "codartic", Text1(1).Text, "T")
        Text2(2).Text = Devuelve
        'Poner el num mantenimiento
        TieneMan = "tieneman"
        Text2(4).Text = DevuelveDesdeBDNew(conAri, "sserie", "nummante", "numserie", Text1(0).Text, "T", TieneMan, "codartic", Text1(1).Text, "T")
        If TieneMan = "0" Then
            Text2(4).Text = ""
        Else
            If Text2(4).Text = "" Then Text2(4).Text = "SIN ESPC."
        End If
    End If
    'Poner la descripcion del Motivo Pendiente Reparac.
    Text2(11).Text = PonerNombreDeCod(Text1(11), conAri, "smotre", "nommotre")
        
        
        
    'Mostraremos SOLO el numero de aviso, no la fecha de donde venia
    If Me.Text1(15).Text <> "" Then Text1(15).Text = RecuperaValor(Text1(15).Text, 1)
    
    
    If ControlRep Then
        'Cargamos el DATA
        CargaGrid DataGrid1, Data2, True
    End If
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    vWhere = "(numrepar=" & Val(Text1(2).Text) & ")"
    If SituarData(Data1, vWhere, Indicador) Then
        PonerModo 2
        lblIndicador.Caption = Indicador
    Else
        LimpiarCampos
        PonerModo 0
    End If
End Sub


Private Sub LimpiarDatosCliente()
Dim I As Byte

    For I = 6 To 13
        Text2(I).Text = ""
    Next I
    Text1(6).Text = ""
    Text1(7).Text = ""

    If (Modo = 3 Or Modo = 4) Then PonerFoco Text1(6)
End Sub


Private Function ArticulosDelNSerie(numSerie As String) As Integer
'Recupera si para ese numero de Serie hay varios articulos que lo tienen
'RETURN -> Nº de articulos diferentes que tienen ese numserie
Dim RS As ADODB.Recordset
Dim SQL As String

    On Error Resume Next

    SQL = "select distinct count(codartic) FROM sserie "
    SQL = SQL & "WHERE numserie='" & numSerie & "'"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        ArticulosDelNSerie = RS.Fields(0).Value
    Else
        ArticulosDelNSerie = 0
    End If
    RS.Close
    Set RS = Nothing
    If Err.Number <> 0 Then Err.Clear
End Function


Private Sub CargarDatosNSerie(numSerie As String, codArtic As String)
Dim SQL As String
Dim RS As ADODB.Recordset

    SQL = "Select codclien, coddirec, tieneman, nummante, ultrepar, fingaran "
    SQL = SQL & "FROM sserie WHERE numserie=" & DBSet(numSerie, "T") & " and "
    SQL = SQL & " codartic=" & DBSet(codArtic, "T")

    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
    
        'Si viene del formulario de AVISO
        'y estamos insertando
        If Modo = 3 Then
        
            If EntradaEquipo <> "" Then
                'Los datos del cliente me viene de reparacion
                If IsNull(RS!CodClien) Then
                    RS.Close
                    Set RS = Nothing
                    Exit Sub
                End If
            
                SQL = RecuperaValor(EntradaEquipo, 3)
                If Val(RS!CodClien) <> Val(SQL) Then
                    MensajeNoCoinciden CStr(Val(RS!CodClien)), False
                    MsgBox CadenaDesdeOtroForm, vbExclamation
                    CadenaDesdeOtroForm = ""
                End If
            End If
        End If
        Text1(6).Text = Format(RS!CodClien, "000000")
        Text1(7).Text = Format(DBLet(RS!CodDirec), "000")
        If Text1(7).Text <> "" Then Text2(7).Text = DevuelveDesdeBDNew(conAri, "sdirec", "nomdirec", "codclien", Text1(6).Text, "N", , "coddirec", Text1(7).Text, "N")
        Text2(2).Text = DBLet(RS!ultrepar, "F")
        Text2(3).Text = DBLet(RS!fingaran, "F")
        
        'Poner fecha prevista reparacion en funcion del param. de la aplicacion (diassiman,diasnoman)
        'dependiendo de si el numserie,codartic tiene mantenimiento (ver tabla sserie)
        If RS!TieneMan = "1" Then
            Text2(4).Text = DBLet(RS!numMante, "T")
            If Text2(4).Text = "" Then Text2(4).Text = "SIN ESPEC."
            Text1(4).Text = Format(Now + vParamAplic.DiasSiMante, "dd/mm/yyyy")
        Else
            Text1(4).Text = Format(Now + vParamAplic.DiasNoMante, "dd/mm/yyyy")
        End If
        
        'Cargar los datos del Cliente
        PonerDatosCliente (Text1(6).Text), True
    End If
    RS.Close
    Set RS = Nothing
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
                    
            If Modo = 4 Then
                'si no se ha modificado el cliente no hacer nada
                If CLng(Text1(6).Text) = CLng(Data1.Recordset!CodClien) Then
           '         If Text2(6).Text = Data1.Recordset!nomclien Then
                        Set vCliente = Nothing
                        Exit Sub
           '         End If
                End If
            End If
            
            Text1(6).Text = vCliente.Codigo
            FormateaCampo Text1(6)
            If (Modo = 3) Or (Modo = 4) Or (Modo = 2) Then 'Insertar o Modificar
                Text2(6).Text = vCliente.Nombre  'Nom clien
                Text2(10).Text = vCliente.Domicilio
                Text2(12).Text = vCliente.CPostal
                Text2(13).Text = vCliente.Poblacion
                Text2(8).Text = vCliente.NIF
                Text2(9).Text = vCliente.TfnoClien
            End If

            Observaciones = DBLet(vCliente.Observaciones)
            If Observaciones <> "" And (Modo = 3 Or Modo = 4) Then
                MsgBox Observaciones, vbInformation, "Observaciones del cliente"
            End If
                           
            'Comprobar si el cliente tiene cobros pendientes
            ComprobarCobrosCliente CodClien, Text1(3).Text
        End If
    Else
        LimpiarDatosCliente
    End If
    Set vCliente = Nothing

EPonerDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poner Datos Cliente", Err.Description
End Sub


Private Function InsertarCabecera() As Boolean
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim SQL As String

    On Error GoTo EInsertarCab
    InsertarCabecera = False
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
        Text1(2).Text = vTipoMov.ConseguirContador(CodTipoMov)
        SQL = CadenaInsertarDesdeForm(Me)
        If SQL <> "" Then
            If InsertarRepar(SQL, vTipoMov) Then
                InsertarCabecera = True
                CadenaConsulta = "Select * from " & NombreTabla & " WHERE numrepar=" & Text1(2).Text '& ObtenerWhereCP(True) & Ordenacion
                PonerCadenaBusqueda
'                PonerModo 2
'                PosicionarData
            End If
        End If
        Text1(2).Text = Format(Text1(2).Text, "0000000")
    End If
    
    Set vTipoMov = Nothing
    
EInsertarCab:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function



Private Function InsertarRepar(vSQL As String, vTipoMov As CTiposMov) As Boolean
Dim MenError As String
Dim bol As Boolean, Existe As Boolean
Dim cambiaSQL As Boolean
Dim Devuelve As String

    On Error GoTo EInsertar
    
    bol = True
    
    cambiaSQL = False
    'Comprobar si mientras tanto se incremento el contador de Pedidos
    'para ello vemos si existe una oferta con ese contador y si existe la incrementamos
    Do
        Devuelve = DevuelveDesdeBDNew(conAri, NombreTabla, "numrepar", "numrepar", Text1(2).Text, "N")
        If Devuelve <> "" Then
            'Ya existe el contador incrementarlo
            Existe = True
            vTipoMov.IncrementarContador (CodTipoMov)
            Text1(2).Text = vTipoMov.ConseguirContador(CodTipoMov)
            cambiaSQL = True
        Else
            Existe = False
        End If
    Loop Until Not Existe
    If cambiaSQL Then vSQL = CadenaInsertarDesdeForm(Me)
    
    
    'Aqui empieza transaccion
    Conn.BeginTrans
    MenError = "Error al insertar en la tabla de Reparaciones (" & NombreTabla & ")."
    Conn.Execute vSQL, , adCmdText
    
    
    MenError = "Error al actualizar el contador del Pedido."
    vTipoMov.IncrementarContador (CodTipoMov)

EInsertar:
    If Err.Number <> 0 Then
        MenError = "Insertando Reparación." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
        bol = False
    End If
    If bol Then
        Conn.CommitTrans
        InsertarRepar = True
    Else
        Conn.RollbackTrans
        InsertarRepar = False
    End If
End Function


Private Function ObtenerWhereCP() As String
Dim SQL As String

    SQL = " WHERE  numrepar= " & Text1(2).Text
    ObtenerWhereCP = SQL
End Function


Private Sub BotonImprimir(OpcionListado As Byte)
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
Dim Devuelve As String


    If Text1(2).Text = "" Then 'Nº Reparacion
        MsgBox "Debe seleccionar una Reparación para Imprimir.", vbInformation
        Exit Sub
    End If
    
    cadFormula = ""
    cadParam = ""
    numParam = 0
    
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1

    'Añadir el parametro con el Nº de mantenimiento si hay
    If Trim(Text2(4).Text) <> "" Then
        cadParam = cadParam & "pMantenimiento=""" & Text2(4).Text & """|"
        numParam = numParam + 1
    End If
      
    'Añadir el parametro si esta en garantia o no
    If Trim(Text2(3).Text) <> "" Then
        If Format(Now, "dd/mm/yyyy") > Format(Text2(3).Text, "dd/mm/yyyy") Then
            cadParam = cadParam & "pGarantia=""NO""|"
        Else
            cadParam = cadParam & "pGarantia=""SI""|"
        End If
        numParam = numParam + 1
    End If
      
    'Nombre fichero .rpt a Imprimir
    If Not PonerParamRPT(24, cadParam, numParam, Devuelve) Then
        Exit Sub
    End If

    'Nombre fichero .rpt a Imprimir
    frmImprimir.NombreRPT = Devuelve
    'frmImprimir.NombreRPT = "rRepResguardo.rpt"
    Devuelve = ""
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Nº de Reparacion
    '---------------------------------------------------
    If Text1(2).Text <> "" Then
        'Nº Reparacion
        Devuelve = "{" & NombreTabla & ".numrepar}=" & Val(Text1(2).Text)
        If Not AnyadirAFormula(cadFormula, Devuelve) Then Exit Sub
    End If
    
     With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        .opcion = OpcionListado
        .Titulo = ""
        .Show vbModal
    End With
End Sub


Private Sub PonerBotonCabecera(b As Boolean)
'Pone el boton de Regresar a la Cabecera si pasamos a MAntenimiento de Lineas
'o Pone los botones de Aceptar y cancelar en Insert,update o delete lineas
On Error Resume Next
    
    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    Me.cmdRegresar.visible = b
    Me.cmdRegresar.Caption = "Cabecera"
    If b Then
        cmdRegresar.Cancel = True
        Me.lblIndicador.Caption = "Líneas Reparaciones"
        PonerFocoBtn Me.cmdRegresar
    Else
        cmdCancelar.Cancel = True
    End If
    
    'Habilitar las opciones correctas del menu según Modo
    
    PonerModoOpcionesMenu
    PonerOpcionesMenu 'Habilitar las opciones correctas del menu según Nivel de Acceso
    If Err.Number <> 0 Then Err.Clear
End Sub



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
                If I < 3 Then
                    txtAux(I).Text = DataGrid1.Columns(I + 2).Text
                Else
                    txtAux(I).Text = DataGrid1.Columns(I + 3).Text
                End If
                BloquearTxt txtAux(I), False
            Next I
            'El campo Nom Artic lo bloqueamos inicialmente
            BloquearTxt txtAux(2), True
        End If
            
        'El campo Importe es calculado y lo bloqueamos.
        BloquearTxt txtAux(7), True

        'Fijamos altura(Height) y posición Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid1, 20)
        alto = alto '+ SSTab1.Top
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
        txtAux(0).Left = DataGrid1.Left + 330 '+ SSTab1.Left
        txtAux(0).Width = DataGrid1.Columns(2).Width - 160
        cmdAux(0).Left = txtAux(0).Left + txtAux(0).Width - 40
        'Cod Artic
        txtAux(1).Left = cmdAux(0).Left + cmdAux(0).Width + 20
        txtAux(1).Width = DataGrid1.Columns(3).Width - 180
        cmdAux(1).Left = txtAux(1).Left + txtAux(1).Width - 30
        'Nom Artic
        txtAux(2).Left = cmdAux(1).Left + cmdAux(1).Width
        txtAux(2).Width = DataGrid1.Columns(4).Width - 10
        'Cantidad
        txtAux(3).Left = txtAux(2).Left + txtAux(2).Width + 10
        txtAux(3).Width = DataGrid1.Columns(6).Width - 10
        'Precio, Dto1, Dto2, Precio
        For I = 4 To txtAux.Count - 1
            txtAux(I).Left = txtAux(I - 1).Left + txtAux(I - 1).Width + 10
            txtAux(I).Width = DataGrid1.Columns(I + 3).Width - 10
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
    txtAux(0).Text = DevuelveDesdeBDNew(conAri, "straba", "codalmac", "codtraba", Text1(5).Text, "N")
    If txtAux(0).Text <> "" Then txtAux(0).Text = Format(txtAux(0).Text, "000")
    'Campo Ampliacion Linea
    Text2(16).Text = ""
    BloquearTxt Text2(16), False
    
    PonerFoco txtAux(0)
    Me.DataGrid1.Enabled = False
    PrimeraVez = True
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
'IN: enlaza= si carga el grid con valores de la tabla o lo muestra vacio si no enlaza
'    conServidas=si enlaza, se muestra la columna de servidas solo cuando se va a generar el Albaran no completo
Dim b As Boolean
Dim SQL As String

    On Error GoTo ECargaGrid

    b = DataGrid1.Enabled
    
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral vDataGrid, vData, SQL, PrimeraVez
    
    CargaGrid2 vDataGrid, vData
    
    b = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2) '5:Modo Mto Lineas (Insertando o Modificando linea)
    vDataGrid.Enabled = Not b
    vDataGrid.ScrollBars = dbgAutomatic

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

    
        'Cod. Almacen
        vDataGrid.Columns(2).Caption = "Alm."

        vDataGrid.Columns(2).Width = 500

        vDataGrid.Columns(2).NumberFormat = "000"
        
        vDataGrid.Columns(3).Caption = "Articulo"

        vDataGrid.Columns(3).Width = 1800

        
        vDataGrid.Columns(4).Caption = "Desc. Artículo"
        vDataGrid.Columns(4).Width = 3400

        vDataGrid.Columns(5).visible = False
        
        vDataGrid.Columns(6).Caption = "Cantidad"
        vDataGrid.Columns(6).Width = 850
        vDataGrid.Columns(6).Alignment = dbgRight
        vDataGrid.Columns(6).NumberFormat = FormatoImporte
        
        I = 7
        vDataGrid.Columns(I).Caption = "Precio"
        vDataGrid.Columns(I).Width = 950
        vDataGrid.Columns(I).Alignment = dbgRight
        vDataGrid.Columns(I).NumberFormat = FormatoPrecio
        
            
        I = I + 1
        vDataGrid.Columns(I).Caption = "Dto.1"
        vDataGrid.Columns(I).Width = 600
        vDataGrid.Columns(I).Alignment = dbgRight
        vDataGrid.Columns(I).NumberFormat = FormatoDescuento
                
        I = I + 1
        vDataGrid.Columns(I).Caption = "Dto.2"
        vDataGrid.Columns(I).Width = 600

        vDataGrid.Columns(I).Alignment = dbgRight
        vDataGrid.Columns(I).NumberFormat = FormatoDescuento
    
        I = I + 1
        vDataGrid.Columns(I).Caption = "Importe Línea"
'        If conServidas Then
'            vDataGrid.Columns(i).Width = 1250
'        Else
            vDataGrid.Columns(I).Width = 1400
'        End If
        vDataGrid.Columns(I).Alignment = dbgRight
        vDataGrid.Columns(I).NumberFormat = FormatoImporte
    
        For I = 0 To vDataGrid.Columns.Count - 1
            vDataGrid.Columns(I).Locked = True
            vDataGrid.Columns(I).AllowSizing = False
        Next I
        vDataGrid.HoldFields
        Exit Sub
        
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub



Private Function MontaSQLCarga(enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data2
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim SQL As String
    
    SQL = "SELECT numrepar, numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, "
    'If conServidas Then SQL = SQL & "servidas, "
    SQL = SQL & "precioar, dtoline1, dtoline2,importel "
    SQL = SQL & " FROM " & NomTablaLineas
    If enlaza Then
        SQL = SQL & " " & ObtenerWhereCP
    Else
        SQL = SQL & " WHERE numrepar = -1"
    End If
    SQL = SQL & " Order by numrepar, numlinea"
    MontaSQLCarga = SQL
End Function


Private Sub txtAux_GotFocus(Index As Integer)
Dim cadkey As Integer

    cadkey = ObtenerCadKey(kCampo, Index)
    kCampo = Index
    ConseguirFocoLin txtAux(Index), cadkey
End Sub

Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 0 And KeyCode = 38 Then Exit Sub 'en almacen y flecha h. arriba
    KEYdown KeyCode
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim Devuelve As String, cadMen As String
Dim codTarif As String
Dim vCStock As CStock
Dim CPrecioFact As CPreciosFact
Dim NumCajas As Integer, RestoUnid As Integer
Dim OrigP As String 'De donde viene el precio
Dim b As Boolean

    'Quitar espacios en blanco
    txtAux(Index).Text = Trim(txtAux(Index))
    
    If txtAux(Index).Text = "" And (Index <> 1) Then Exit Sub
    
    If Screen.ActiveForm.ActiveControl.Name = "cmdCancelar" Then Exit Sub
    
     Select Case Index
        Case 0 'Cod Almacen
            'Comprobar que existe el almacen
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
                If Not Data2.Recordset.EOF Then Devuelve = Data2.Recordset!codArtic
            End If
            
            If Not PonerArticulo(txtAux(1), txtAux(2), txtAux(0).Text, CodTipoMov, ModificaLineas, Devuelve) Then
                PonerFoco txtAux(Index)
            Else
                b = (Me.ActiveControl.Name = "txtAux")
                If b Then b = (Me.ActiveControl.Index = 0)
                If Not b Then
'                    If txtAux(2).Locked Then PonerFoco txtAux(3)
                Else
                    PonerFoco txtAux(0)
                End If
            End If
            
        Case 3 'CANTIDAD
            If PonerFormatoDecimal(txtAux(Index), 1) Then  'Tipo 1: Decimal(12,2)
                'Comprobar si hay suficiente stock
                Set vCStock = New CStock
                If Not InicializarCStock(vCStock, "S") Then Exit Sub '"S"=Salida de Stock
                vCStock.MoverStock (False)
                If Not PrimeraVez Then Exit Sub
                PrimeraVez = False
                If (Modo = 5 And ModificaLineas = 1) Then 'Modo Insertar en Mto Lineas
                    'Ver si esta en Garantia el Aparato
                    'Si el Articulo esta en garantia pregunta si se facturara la linea o no
                    'Si facturar -> precioar=Precio
                    'Si no facturar -> precioar=0
                    If EsFechaPosterior(Text1(3).Text, Text2(3).Text, False) Then
                       If MsgBox("El aparato esta en Garantía.¿Facturar la linea de Reparacion?", vbYesNo) = vbNo Then
                            txtAux(4).Text = "0,00"
                            txtAux(5).Text = "0,00"
                            txtAux(6).Text = "0,00"
                            Set vCStock = Nothing
                            Exit Sub
                       End If
                    End If
                        
                    'Si el aparato tiene Mantenimiento no se cobra la linea de Reparación? Preguntar
                    If Text2(4).Text <> "" Then
                        If MsgBox("El aparato tiene Mantenimiento.¿Facturar la linea de Reparación?", vbYesNo) = vbNo Then
                            txtAux(4).Text = "0,00"
                            txtAux(5).Text = "0,00"
                            txtAux(6).Text = "0,00"
                            Set vCStock = Nothing
                            Exit Sub
                        End If
                    End If
                        
                    'Obtener el precio correspondiente y los descuentos
                    'Comprobar si el articulo se vende por cajas antes de entrar a la función
                    Devuelve = DevuelveDesdeBDNew(conAri, "sartic", "unicajas", "codartic", txtAux(1).Text, "T")
                    If Devuelve <> "" Then
                        Set CPrecioFact = New CPreciosFact
                        'Si se puede vender por cajas(devuelve>1) poner numero de cajas en una linea con el
                        'precio de caja, y otra linea con el resto unidades un precio unidad
                        NumCajas = CPrecioFact.ObtenerNumCajas(vCStock.Cantidad, Devuelve)
                        RestoUnid = CInt(vCStock.Cantidad) - NumCajas * CInt(Devuelve)
                        'Obtenemos la Tarifa del Cliente
                        codTarif = DevuelveDesdeBDNew(conAri, "sclien", "codtarif", "codclien", Text1(6).Text, "N")
                        CPrecioFact.CodigoLista = codTarif
                        CPrecioFact.CodigoArtic = vCStock.codArtic
                        CPrecioFact.CodigoClien = Text1(6).Text
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
                            If txtAux(4).Text = "" Then
                                txtAux(4).Text = Precio
                            End If
                            PonerFormatoDecimal txtAux(4), 2
                            If txtAux(5).Text = "" Then txtAux(5).Text = CPrecioFact.Descuento1
                            PonerFormatoDecimal txtAux(5), 4
                            If txtAux(6).Text = "" Then txtAux(6).Text = CPrecioFact.Descuento2
                            PonerFormatoDecimal txtAux(6), 4
                        End If
                        Set CPrecioFact = Nothing
                    End If
                End If
                Set vCStock = Nothing
            End If
            
        Case 4 'Precio
            PonerFormatoDecimal txtAux(Index), 2 'Tipo 2: Decimal(10,4)
            
        Case 5, 6 'Descuentos
            PonerFormatoDecimal txtAux(Index), 4 'Tipo 4: Decimal(4,2)
            
        Case 7 'Importe Linea
            PonerFormatoDecimal txtAux(Index), 3 'Tipo 3: Decimal(10,2)
    End Select
    
    If Modo = 5 Then 'Modo Lineas
        If (Index = 3 Or Index = 4 Or Index = 5 Or Index = 6) Then 'Cant., Precio, dto1, dto2
            If txtAux(1).Text = "" Then Exit Sub 'Cod artic
            txtAux(7).Text = CalcularImporte(txtAux(3).Text, txtAux(4).Text, txtAux(5).Text, txtAux(6).Text, vParamAplic.TipoDtos)
            PonerFormatoDecimal txtAux(7), 1
            'If Index = 6 Then PonerFocoBtn cmdAceptar
        End If
    End If
End Sub


Private Function InicializarCStock(ByRef vCStock As CStock, TipoM As String, Optional numlinea As String) As Boolean
On Error Resume Next

    vCStock.tipoMov = TipoM
    vCStock.DetaMov = CodTipoMov
    vCStock.Trabajador = CLng(Text1(6).Text) 'guardamos el cliente
    vCStock.Documento = Text1(2).Text 'Nº Albaran
    If ModificaLineas = 1 Or ModificaLineas = 2 Then '1=Insertar, 2=Modificar
        vCStock.codArtic = txtAux(1).Text
        vCStock.codAlmac = CInt(txtAux(0).Text)
        If ModificaLineas = 1 Then '1=Insertar
            vCStock.Cantidad = CSng(ComprobarCero(txtAux(3).Text))
        Else '2=Modificar(Debe haber en stock la diferencia)
            vCStock.Cantidad = CSng(ComprobarCero(txtAux(3).Text)) - Data2.Recordset!Cantidad
        End If
        vCStock.Importe = CCur(ComprobarCero(txtAux(7).Text))
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


Private Function InsertarLinea() As Boolean
'Inserta un registro en la tabla de lineas de Pedido: slirep
Dim SQL As String
Dim numlinea As String, vWhere As String

    On Error GoTo EInsertarLinea

    InsertarLinea = False
    SQL = ""

    If DatosOkLinea() Then 'Lineas de Pedidos
        'Conseguir el siguiente numero de linea
        vWhere = Mid(ObtenerWhereCP, 7)
        numlinea = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", vWhere)
        'Construir la sentencia SQL
        SQL = "INSERT INTO " & NomTablaLineas
        SQL = SQL & "(numrepar, numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel) "
        SQL = SQL & "VALUES (" & DBSet(Text1(2).Text, "N") & ", " & numlinea & ", " & Val(txtAux(0).Text) & ","
        SQL = SQL & DBSet(txtAux(1).Text, "T") & ", " & DBSet(txtAux(2).Text, "T") & ", " & DBSet(Text2(16).Text, "T") & ", "
        SQL = SQL & DBSet(txtAux(3).Text, "N", "N") & ", " 'cantidad
        SQL = SQL & DBSet(txtAux(4).Text, "N", "N") & ", " 'precio
        SQL = SQL & DBSet(txtAux(5).Text, "N", "N") & ", " 'Dto1
        SQL = SQL & DBSet(txtAux(6).Text, "N", "N") & ", " ' Dto2
        SQL = SQL & DBSet(txtAux(7).Text, "N", "N") & ")" 'Importe linea
    End If
    
    If SQL <> "" Then
        Conn.Execute SQL
        InsertarLinea = True
    End If
    Exit Function
    
EInsertarLinea:
    MuestraError Err.Number, "Insertar Lineas Reparación" & vbCrLf & Err.Description
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
        
        If txtAux(I).Text = "" Then
            If I = 5 Or I = 6 Then
                'LOS DESCUENTOS
                'Si los descuentos estan a blancos, pinto el cero yo
                txtAux(I).Text = "0"
            Else
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
       
    DatosOkLinea = b
    
EDatosOkLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function ModificarLinea() As Boolean
'Modifica un registro en la tabla de lineas de Reparaciones: slirep
Dim SQL As String
Dim vCStock As CStock
Dim b As Boolean

    On Error GoTo EModificarLinea

    ModificarLinea = False
    SQL = ""
    Set vCStock = New CStock
    If Not InicializarCStock(vCStock, "S") Then Exit Function

    If DatosOkLinea() Then
        SQL = "UPDATE " & NomTablaLineas & " Set codalmac = " & txtAux(0).Text & ", codartic=" & DBSet(txtAux(1).Text, "T") & ", "
        SQL = SQL & "nomartic=" & DBSet(txtAux(2).Text, "T") & ", ampliaci=" & DBSet(Text2(16).Text, "T") & ", "
        SQL = SQL & "cantidad = " & DBSet(txtAux(3).Text, "N", "N") & ", "
        SQL = SQL & "precioar = " & DBSet(txtAux(4).Text, "N", "N") & ", "
        SQL = SQL & "dtoline1= " & DBSet(txtAux(5).Text, "N", "N") & ", dtoline2= " & DBSet(txtAux(6).Text, "N", "N") & ", "
        SQL = SQL & "importel=" & DBSet(txtAux(7).Text, "N", "N") & " "
        SQL = SQL & ObtenerWhereCP & " AND numlinea=" & Data2.Recordset!numlinea

        If SQL <> "" Then
            Conn.BeginTrans
            Conn.Execute SQL
            vCStock.Cantidad = CSng(txtAux(3).Text)
            b = vCStock.ModificarStock2(Data2.Recordset!Cantidad)
        End If
    End If
    Set vCStock = Nothing
    
EModificarLinea:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Modificar Lineas Reparación" & vbCrLf & Err.Description
        b = False
    End If
    If b Then
        Conn.CommitTrans
    Else
        Conn.RollbackTrans
    End If
    ModificarLinea = b
End Function


Private Function SePuedeServirPedido(Optional cadErr As String) As Boolean
'Comprobar Si se puede servir la Reparacion solicitada y pasar a albaran
Dim vCStock As CStock
Dim SQL As String
Dim b As Boolean
Dim RS As ADODB.Recordset

    On Error GoTo EServir

    SePuedeServirPedido = False
    'Verificar si hay stock para aquellas familias que no son instalacion
    Set vCStock = New CStock
    
    SQL = "SELECT codalmac, codartic, SUM(cantidad) as cantidad from " & NomTablaLineas
    SQL = SQL & ObtenerWhereCP
    SQL = SQL & " GROUP by codalmac, codartic"
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    'Si no hay lineas para pasar al albaran no seguimos
    If RS.EOF Then
        cadErr = "No hay lineas para generar el Albaran."
        b = False
        GoTo EServir
    End If
    
    'para cada linea de la Reparacion comprobar el stock si no es instalacion
    b = True
    While (Not RS.EOF) And b
        If Not InicializarCStockAlbar(vCStock, "S", , RS) Then
            cadErr = "No se pudo inicializar la clase Stock"
            b = False
            GoTo EServir
'            Exit Function
        End If
        'Comprobar si se puede mover stock (hay stock, o si no hay pero no control de stock)
        cadErr = ""
        If vCStock.MueveStock Then
            If Not vCStock.MoverStock(False, True) Then b = False
        End If
        RS.MoveNext
    Wend
    Set vCStock = Nothing
    RS.Close
    Set RS = Nothing
    SePuedeServirPedido = b
    
EServir:
    If Err.Number <> 0 Then
        b = False
        Set vCStock = Nothing
        RS.Close
        Set RS = Nothing
    End If
    
    SePuedeServirPedido = b
End Function


Private Function InicializarCStockAlbar(ByRef vCStock As CStock, TipoM As String, Optional numlinea As String, Optional ByRef RS As ADODB.Recordset) As Boolean
'Para comprobar stock al pasar de Reparacion a Albaran de Reparacion
On Error Resume Next
    
    vCStock.tipoMov = TipoM
    vCStock.DetaMov = "ALR"
    vCStock.Trabajador = CLng(Text1(6).Text) 'guardamos el cliente
    vCStock.Documento = Text1(2).Text
    vCStock.codArtic = RS!codArtic
    vCStock.codAlmac = CInt(RS!codAlmac)
    
    vCStock.Cantidad = CSng(RS!Cantidad)
    'Si no se selecciona el campo importe de la tabla es que solo vamos a comprobar stock y no se necesita
    If RS.Fields.Count > 3 Then vCStock.Importe = CCur(RS!ImporteL)
    
    vCStock.LineaDocu = CInt(ComprobarCero(numlinea))
    
    If Err.Number <> 0 Then
        MsgBox "No se han podido inicializar la clase para actualizar Stock", vbExclamation
        InicializarCStockAlbar = False
    Else
        InicializarCStockAlbar = True
    End If
End Function


Private Sub GenerarAlbaran()
Dim numRep As Long 'Nº Reparacion
Dim NumAlb As Long 'Nº Albaran

    'Pedir: Operador de Albaran, Material Preparado por y forma de envio
    Set frmList = New frmListadoPed
    frmList.NumCod = CodTipoMov
    frmList.OpcionListado = 43
    frmList.Show vbModal
    Set frmList = Nothing
    
    NumRegElim = Data1.Recordset.AbsolutePosition
    numRep = Data1.Recordset!numrepar

    If PasarPedidoAAlbaran(CadenaSQL, NumAlb) Then
        MsgBox "La Reparación Nº: " & Format(numRep, "0000000") & " ha generado " & vbCrLf & vbCrLf & "el Albaran de Reparación Nº: " & Format(NumAlb, "0000000"), vbInformation
        PonerModo 2
        'Se habra eliminado el pedido de (scarep, slirep)
        PosicionarDataTrasEliminar
    End If
    Screen.MousePointer = vbDefault
    
    'Imprimer albaran si se solicitó
'    If ImprimeAlb Then
'            ImprimirAlbaran 45, NumAlb
'    End If
End Sub


Private Sub PosicionarDataTrasEliminar()
'Despues Eliminar y hacer refresh del Data, situar el Data en el registro siguiente
    If SituarDataTrasEliminar(Data1, NumRegElim) Then
        PonerCampos
    Else
        LimpiarCampos
        If ControlRep Then
            'Cargamos el DATA
            CargaGrid DataGrid1, Data2, False
        End If
        PonerModo 0
    End If
End Sub


Private Function PasarPedidoAAlbaran(vSQL As String, NumAlb As Long) As Boolean
'IN -> vSQL: cadena para el Select con los datos obtenidos en frmList
'OUT -> numAlb: Nº de Albaran de Venta que se ha insertado
Dim bol As Boolean
Dim MenError As String
    
    On Error GoTo EGenPedido

    bol = False
        
    'Aqui empieza transaccion
    Conn.BeginTrans
    
    'Insertar en tablas de Albaranes el Pedido (scaalb, slialb)
    MenError = "Insertando el tablas de albaranes. (scaalb,slialb)"
    bol = InsertarAlbaran(vSQL, MenError, NumAlb)
    
    'Actualizar Stock en salmac, e introducir movimiento en smoval
    If bol Then
        MenError = "Actualizando movimientos de stock."
        bol = InsertarMovStock(NumAlb)
    End If
    
    If bol Then
        MenError = "Pasando al histórico de reparaciones."
        'Pasar al Historico de Reparaciones: schrep
        bol = InsertarCabeceraHcoRep(NumAlb)
         
        'Borrar la Reparacion de las tablas de Reparaciones (scarep, slirep)
        MenError = "Eliminando en tablas de reparaciones.(scarep,slirep)"
        If bol Then bol = Eliminar()
    End If
    
    
    
    'Si correcto y tiene numnero de aviso, cierro el aviso
    If bol Then
        If Text1(15).Text <> "" Then
            'LLEVA REPARCION
            Text1(15).Text = Data1.Recordset!numaviso
            MenError = "Actualizando avisos."
            CadenaDesdeOtroForm = "UPDATE scaavi SET situacio = 3"
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & " WHERE numaviso =" & RecuperaValor(Text1(15).Text, 1)
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & " AND  fechaavi = '" & Format(RecuperaValor(Text1(15).Text, 2), FormatoFecha) & "'"
            Conn.Execute CadenaDesdeOtroForm
        End If

    End If
    
EGenPedido:
    If Err.Number <> 0 Then bol = False
    
    If bol Then
        Conn.CommitTrans
    Else
        Conn.RollbackTrans
        MenError = "Pasando Reparación a Albaran." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
        CadenaDesdeOtroForm = ""
    End If
    PasarPedidoAAlbaran = bol
End Function


Private Function InsertarAlbaran(vSQL As String, MenError As String, NumAlb As Long) As Boolean
'Devuelve el mensaje de error si se produce
Dim bol As Boolean, Existe As Boolean
Dim Devuelve As String, SQL As String
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim codTipoM As String

    On Error GoTo EInsertarAlbaran
    
    bol = False
    InsertarAlbaran = bol
    
    'Obtener el Contador de ALBARAN de Reparacion
    codTipoM = "ALR"
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(codTipoM) Then
        'Comprobar si mientras tanto se incremento el contador de Pedidos
        'para ello vemos si existe una oferta con ese contador y si existe la incrementamos
        Do
            NumAlb = vTipoMov.ConseguirContador(codTipoM)
            Devuelve = DevuelveDesdeBDNew(conAri, "scaalb", "numalbar", "codtipom", codTipoM, "T", , "numalbar", CStr(NumAlb), "N")
            If Devuelve <> "" Then
                'Ya existe el contador incrementarlo
                Existe = True
                vTipoMov.IncrementarContador (codTipoM)
                NumAlb = vTipoMov.ConseguirContador(codTipoM)
            Else
                Existe = False
            End If
        Loop Until Not Existe
            
    Else 'No existe el tipo de Movimiento
        Set vTipoMov = Nothing
        Exit Function
    End If
    
    'Acabar la sql con el contador seleccionado
    SQL = "INSERT INTO scaalb (codtipom, numalbar, fechaalb, factursn, codclien, nomclien, domclien, codpobla, pobclien, proclien, "
    SQL = SQL & "nifclien, telclien, coddirec, nomdirec, referenc, codtraba, codtrab1, codtrab2, codagent, codforpa, codenvio, "
    SQL = SQL & "dtoppago, dtognral, tipofact, observa01, observa02, observa03, observa04, observa05, numofert, fecofert, numpedcl, fecpedcl, sementre) "
    SQL = SQL & " VALUES ('" & codTipoM & "', " & NumAlb & "," & vSQL & ")"
    
    'Insertar Cabecera
    MenError = "Error al insertar en la tabla Cabecera de Albaranes (scaalb )."
    Conn.Execute SQL, , adCmdText
    
    'Insertar Lineas de Albaran
    MenError = "Error al insertar en la tabla Lineas de Albaran (slialb)."
    If Not InsertarLineasAlbaran(codTipoM, NumAlb) Then Exit Function
    
    MenError = "Error al actualizar el contador del Albaran."
    vTipoMov.IncrementarContador (codTipoM)
    Set vTipoMov = Nothing
    bol = True
    
EInsertarAlbaran:
    If Err.Number <> 0 Then bol = False
    InsertarAlbaran = bol
End Function


Private Function InsertarMovStock(NumAlb As Long) As Boolean
Dim vCStock As CStock
Dim b As Boolean
Dim RS As ADODB.Recordset
Dim SQL As String

    On Error GoTo EInsMov

    InsertarMovStock = False
    
    Set vCStock = New CStock
    b = True
    
    SQL = "select * from " & NomTablaLineas & ObtenerWhereCP
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'para cada linea del Pedido Insertar en smoval y Actualizar Stock en salmac
    While (Not RS.EOF) And b
        If Not InicializarCStockAlbar(vCStock, "S", CStr(RS!numlinea), RS) Then Exit Function
        vCStock.Documento = CStr(NumAlb)
         'en actualizar stock comprobamos si el articulo tiene control de stock
        'If vCStock.Cantidad <> 0 Then
            b = vCStock.ActualizarStock
        RS.MoveNext
    Wend
    Set vCStock = Nothing
    RS.Close
    Set RS = Nothing
    
'    InsertarMovStock = b
    
EInsMov:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Insertando movimiento de stock.", Err.Description
        b = False
    End If
    InsertarMovStock = b
End Function



Private Function InsertarLineasAlbaran(TipoM As String, NumAlb As Long) As Boolean
'Inserta en la tabla de lineas de albaran (slialb)
Dim SQL As String

    On Error GoTo EInsertarLin

    'Insertar en la tabla de Pedido, los registros seleccionados de la tabla de Ofertas
    'Cambio por el rollo de la trazabilidad
    SQL = ""
    SQL = "SELECT '" & TipoM & "' as codtipom, " & NumAlb & " as numalbar, numlinea, codalmac,  s.codartic, s.nomartic , ampliaci, "
    SQL = SQL & "cantidad, precioar, dtoline1, dtoline2, importel, '' as origpre,codprove"
    SQL = SQL & " FROM " & NomTablaLineas & " s,sartic WHERE s.codartic=sartic.codartic AND numrepar=" & Text1(2).Text
    SQL = "INSERT INTO slialb " & SQL
    Conn.Execute SQL
    InsertarLineasAlbaran = True

EInsertarLin:
'    If Err.Number <> 0 Then
'        InsertarLineasAlbaran = False
'    Else
'        InsertarLineasAlbaran = True
'    End If
    InsertarLineasAlbaran = Not (Err.Number <> 0)
End Function


Private Function InsertarCabeceraHcoRep(NumAlb As Long) As Boolean
'Insertar en la Tabla Cabecera de Historico
Dim SQL As String
Dim Aux As String

    On Error Resume Next
    
    
    SQL = "SELECT numrepar, fecrepar,fecentre," & NombreTabla & ".numserie, " & NombreTabla & ".codartic, sartic.nomartic, "
    'fecha fin garantia: fingaran, ultrepar
    SQL = SQL & DBSet(Text2(3).Text, "F") & " as fingaran, " & DBSet(Text2(2).Text, "F", "S") & " as ultrepar, "
    SQL = SQL & "codclien, coddirec, " & DBSet(Text2(4).Text, "T") & " as nunmante, " 'nummante
    SQL = SQL & "codtraba, " & CadenaSQLHco & ", "
    SQL = SQL & "'ALR' as codtipom, " & NumAlb & " as numalbar, " & DBSet(FechaAlb, "F") & " as fechaalb "
    
    'Modifiaciones 1 OCTUBRE 2007
    'Añadimos SAT tipo averia y presupuestos
    Aux = ",codman,codavi,codtrabajo,imppresu1,impresu2,contestado,fecha,fechaaprob,avisocli,fecenviosat,resguardosat,importesat,fecentresat,observasat"
    SQL = SQL & Aux
    
    SQL = SQL & " FROM " & NombreTabla & " INNER JOIN sartic ON " & NombreTabla & ".codartic=sartic.codartic "
    SQL = SQL & ObtenerWhereCP
    
    SQL = "INSERT INTO schrep (numrepar,fecrepar,fecentre,numserie,codartic,nomartic,fingaran,ultrepar,codclien,coddirec,nummante,codtraba,codtrab1,codtrab2,material,tipoaver,motivore,textore1,textore2,textore3,codtipom,numalbar,fechaalb" & Aux & ") " & SQL
    Conn.Execute SQL
    
'    If Err.Number <> 0 Then
'         'Hay error , almacenamos y salimos
'        InsertarCabeceraHcoRep = False
'    Else
'        InsertarCabeceraHcoRep = True
'    End If
    InsertarCabeceraHcoRep = Not (Err.Number <> 0)
End Function



Private Sub CargaDatosAviso()
    On Error GoTo ECargaDatosAviso
    
    
    
    If EntradaEquipo = "" Then Exit Sub
    
    BotonAnyadir
            
    'Ahora pongo los campos de la entradequipo
    'Numero aviso
    Text1(15).Text = RecuperaValor(EntradaEquipo, 1) & "|" & RecuperaValor(EntradaEquipo, 2) & "|"
    'Cliente
    Text1(6).Text = RecuperaValor(EntradaEquipo, 3)
    Text2(6).Text = RecuperaValor(EntradaEquipo, 4)
    'NIF
    Text2(8).Text = RecuperaValor(EntradaEquipo, 7)
    'Tfno
    Text2(9).Text = RecuperaValor(EntradaEquipo, 8)
    'Domicilio
    Text2(10).Text = RecuperaValor(EntradaEquipo, 9)
    'Codpostal
    Text2(12).Text = RecuperaValor(EntradaEquipo, 10)
    'Pobla
    Text2(13).Text = RecuperaValor(EntradaEquipo, 11)
    'Dpto
    Text1(7).Text = RecuperaValor(EntradaEquipo, 5)
    Text2(7).Text = RecuperaValor(EntradaEquipo, 6)
    
    
    Exit Sub
ECargaDatosAviso:
    MuestraError Err.Number, "CargaDatosAviso"

End Sub


Private Sub MensajeNoCoinciden(Equipo As String, Pregunta As Boolean)

    CadenaDesdeOtroForm = "############"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & CadenaDesdeOtroForm
    CadenaDesdeOtroForm = vbCrLf & vbCrLf & CadenaDesdeOtroForm & CadenaDesdeOtroForm & vbCrLf & vbCrLf
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & " No coinciden el cliente del aviso (" & RecuperaValor(EntradaEquipo, 3) & ") con el del numero de serie (" & Equipo & ")" & CadenaDesdeOtroForm
    If Pregunta Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & vbCrLf & vbCrLf & "¿Continuar?"
End Sub
