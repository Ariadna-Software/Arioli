VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOliTarifaOferta 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   12660
   Icon            =   "frmOliTarifaOferta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   12660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Data3 
      Height          =   330
      Left            =   2640
      Top             =   7680
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
      Height          =   5895
      Left            =   120
      TabIndex        =   36
      Top             =   1680
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   10398
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Artículos"
      TabPicture(0)   =   "frmOliTarifaOferta.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DataGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtAux(9)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtAux(8)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtAux(12)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtAux(11)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtAux(10)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtAux(7)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtAux(6)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtAux(5)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtAux(4)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtAux(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtAux(3)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtAux(2)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdAux(1)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Materias primas"
      TabPicture(1)   =   "frmOliTarifaOferta.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DataGrid2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Observaciones"
      TabPicture(2)   =   "frmOliTarifaOferta.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Text1(5)"
      Tab(2).ControlCount=   1
      Begin VB.TextBox Text1 
         Height          =   4875
         Index           =   5
         Left            =   -72600
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   42
         Tag             =   "Obs|T|S|||olitarifaoferta|observaciones|||"
         Top             =   600
         Width           =   8505
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   4935
         Left            =   -73440
         TabIndex        =   40
         Top             =   600
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   8705
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Index           =   1
         Left            =   1680
         TabIndex        =   39
         ToolTipText     =   "Buscar artículo"
         Top             =   3960
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   38
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
         Index           =   3
         Left            =   4320
         MaxLength       =   16
         TabIndex        =   7
         Tag             =   "pivu"
         Text            =   "pivu"
         Top             =   3960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   360
         MaxLength       =   18
         TabIndex        =   6
         Tag             =   "Código Artículo"
         Text            =   "codartic"
         Top             =   3960
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   4
         Left            =   5400
         MaxLength       =   16
         TabIndex        =   8
         Tag             =   "pivl"
         Text            =   "pivl"
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
         Index           =   5
         Left            =   6480
         MaxLength       =   16
         TabIndex        =   9
         Tag             =   "coste1"
         Text            =   "coste1"
         Top             =   3960
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   6
         Left            =   7320
         MaxLength       =   16
         TabIndex        =   10
         Tag             =   "coste2"
         Text            =   "coste2"
         Top             =   3960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   7
         Left            =   7920
         MaxLength       =   16
         TabIndex        =   11
         Tag             =   "Coste3"
         Text            =   "coste3"
         Top             =   3960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   10
         Left            =   9360
         MaxLength       =   16
         TabIndex        =   14
         Tag             =   "margen"
         Text            =   "margen"
         Top             =   3960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   11
         Left            =   9840
         MaxLength       =   16
         TabIndex        =   15
         Tag             =   "pvfu"
         Text            =   "pvfu"
         Top             =   3960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   12
         Left            =   10680
         MaxLength       =   16
         TabIndex        =   16
         Tag             =   "pvfl"
         Text            =   "pvfl"
         Top             =   3960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   8
         Left            =   8520
         MaxLength       =   16
         TabIndex        =   12
         Tag             =   "Coste4"
         Text            =   "coste4"
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
         Index           =   9
         Left            =   8880
         MaxLength       =   16
         TabIndex        =   13
         Tag             =   "Coste5"
         Text            =   "coste5"
         Top             =   3960
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmOliTarifaOferta.frx":0060
         Height          =   5280
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   9313
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
               ColumnAllowSizing=   -1  'True
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   900
      Left            =   120
      TabIndex        =   24
      Top             =   480
      Width           =   12255
      Begin VB.Frame FrameTarifa 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   4920
         TabIndex        =   5
         Top             =   120
         Width           =   6975
         Begin VB.TextBox Text1 
            BackColor       =   &H80000013&
            Height          =   315
            Index           =   4
            Left            =   240
            TabIndex        =   30
            Tag             =   "Tarifa|N|S|||olitarifaoferta|tarifa|||"
            Text            =   "Text1 7"
            Top             =   240
            Width           =   1245
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   315
            Index           =   4
            Left            =   1560
            TabIndex        =   29
            Text            =   "Text2"
            Top             =   240
            Width           =   4215
         End
         Begin VB.Label Label1 
            Caption         =   "Tarifa"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   31
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.Frame FrameCliente 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   4920
         TabIndex        =   32
         Top             =   120
         Width           =   7095
         Begin VB.TextBox Text1 
            BackColor       =   &H80000013&
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   3
            Tag             =   "Cliente|N|S|||olitarifaoferta|codclien|||"
            Text            =   "Text1 7"
            Top             =   240
            Width           =   1245
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            Left            =   1440
            TabIndex        =   33
            Text            =   "Text2"
            Top             =   240
            Width           =   4215
         End
         Begin VB.CheckBox Check1 
            Caption         =   " "
            Height          =   255
            Left            =   5760
            TabIndex        =   4
            Tag             =   "Cliente|N|N|||olitarifaoferta|aceptada|||"
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Aceptada"
            Height          =   255
            Index           =   3
            Left            =   6240
            TabIndex        =   41
            Top             =   270
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Cliente"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   34
            Top             =   0
            Width           =   735
         End
         Begin VB.Image imgCuentas 
            Height          =   240
            Index           =   0
            Left            =   840
            Picture         =   "frmOliTarifaOferta.frx":0075
            Tag             =   "-1"
            ToolTipText     =   "Buscar articulo"
            Top             =   0
            Width           =   240
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha fin|F|N|||olitarifaoferta|fechafin|dd/mm/yyyy|N|"
         Top             =   360
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha inicio|F|N|||olitarifaoferta|fechaini|dd/mm/yyyy|N|"
         Top             =   360
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   0
         Left            =   240
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "Codigo|N|S|0||olitarifaoferta|codigo|0000000|S|"
         Text            =   "Text1 7"
         Top             =   360
         Width           =   1365
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   4320
         Picture         =   "frmOliTarifaOferta.frx":0177
         ToolTipText     =   "Buscar fecha"
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "F. fin"
         Height          =   255
         Index           =   0
         Left            =   3360
         TabIndex        =   28
         Top             =   165
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "F. inicio"
         Height          =   255
         Index           =   14
         Left            =   1920
         TabIndex        =   26
         Top             =   165
         Width           =   735
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   2760
         Picture         =   "frmOliTarifaOferta.frx":0202
         ToolTipText     =   "Buscar fecha"
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo"
         Height          =   255
         Index           =   50
         Left            =   240
         TabIndex        =   25
         Top             =   165
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   7680
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
      Left            =   11130
      TabIndex        =   18
      Top             =   7800
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9960
      TabIndex        =   17
      Top             =   7800
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   2400
      Top             =   7920
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
      TabIndex        =   22
      Top             =   0
      Width           =   12660
      _ExtentX        =   22331
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
            Object.ToolTipText     =   "Lineas ofertas"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "AVAB"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir "
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Enviar cliente"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
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
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   6600
         TabIndex        =   23
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   2280
      Top             =   7920
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
      Left            =   11130
      TabIndex        =   19
      Top             =   7800
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   35
      Top             =   7680
      Width           =   4935
   End
   Begin VB.Label Label2 
      Caption         =   "Productos ofertados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   1440
      Width           =   3375
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
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
      End
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnORden 
      Caption         =   "Ordenacion"
      Begin VB.Menu mnORden2 
         Caption         =   "Codigo articulo"
         Index           =   0
      End
      Begin VB.Menu mnORden2 
         Caption         =   "Nombre articulo"
         Index           =   1
      End
      Begin VB.Menu mnORden2 
         Caption         =   "Categoria, marca,formato"
         Index           =   2
      End
      Begin VB.Menu mnORden2 
         Caption         =   "Formato,Marca,categoria"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmOliTarifaOferta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Solo para cuando viene desde pantalla smoval
Public DatosADevolverBusqueda2 As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado2(CadenaSeleccion As String)


Public EsTarifa As Boolean

Private WithEvents frmB As frmBuscaGrid  'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1

Private WithEvents frmArt As frmAlmArticulos
Attribute frmArt.VB_VarHelpID = -1

Private WithEvents frmCl As frmFacClientes
Attribute frmCl.VB_VarHelpID = -1

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


Private ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

Dim PrimeraVez As Boolean



'SQL de la tabla principal del formulario
Private CadenaConsulta As String


Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla de Cabecera
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Dim btnAnyadir As Byte
'Variable que indica el número del Boton  Anyadir en la Toolbar1
Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1






Dim gridCargado As Boolean 'Saber si el grid esta cargado cuando se ejecuta DataGrid1_RowColChange


Dim vArti As CArticulo  'Para las lineas
'================================================================================






Private Sub cmdAceptar_Click()
'Dim SQL As String
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
                
                If ModificaDesdeFormulario(Me, 1) Then
                    'ActualizarLineasPedido
                    TerminaBloquear
                    PosicionarData
                End If
            End If
            
         Case 5 'InsertarModificar LINEA
            
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
            
            
                    Dim Linea As Integer
                    Dim Cad As String
                    Linea = DataGrid1.FirstRow
                    Cad = "codartic = '" & Data2.Recordset!codartic & "'"
                    
                    
                    TerminaBloquear
                    CargaTxtAux False, False
                    CargaGrid2 DataGrid1, Data2
                    
                    
                    ModificaLineas = 0
                    PonerBotonCabecera True
                    Data2.Recordset.Find Cad
                    
                    On Error Resume Next
                    Linea = Linea - DataGrid1.FirstRow
                    DataGrid1.Scroll 0, Linea
                    If Err.Number <> 0 Then
                        MsgBox Err.Number & ": " & Err.Description, vbExclamation
                        Err.Clear
                    End If
                        
                End If
                Me.DataGrid1.Enabled = True
            End If
            
            

    End Select
    
    
Error1:
   
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
     Screen.MousePointer = vbDefault
End Sub


Private Sub cmdAux_Click(Index As Integer)

    Set frmArt = New frmAlmArticulos
    frmArt.DatosADevolverBusqueda2 = "@1@" 'Poner en modo busqueda
    frmArt.Show vbModal
    Set frmArt = Nothing

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
'Añadir registro en tabla de cabecera de Pedidos: scaped (Cabecera)
Dim NomTraba As String

    LimpiarCampos 'Vacía los TextBox
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3

 

    'Poner el nombre del trabajador que esta conectado
    'Text1(3).Text = PonerTrabajadorConectado(NomTraba)
    'Text2(3).Text = NomTraba

    Text1(1).Text = Format(Now, "dd/mm/yyyy hh:mm") 'Fecha Oferta
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
Dim C As String
'    LimpiarCampos
    If Me.EsTarifa Then
        C = " codigo < 100000"
    Else
        C = " codigo > 100000"
    End If
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia C
    Else
        LimpiarCampos
        LimpiarDataGrids
        CadenaConsulta = "Select * from " & NombreTabla & " WHERE " & C & Ordenacion
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
    cmdCancelar.Cancel = True

End Sub


Private Sub BotonModificarLinea()
'Prepara el Form para Modificar una linea de Pedido (tabla: sliped)
Dim vWhere As String

    On Error GoTo EModificarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub '1= Insertar
    If Data2.Recordset.EOF Then Exit Sub
    
  
    Set vArti = Nothing
    Set vArti = New CArticulo
    If Not vArti.LeerDatos(CStr(Data2.Recordset!codartic)) Then
        Set vArti = Nothing
        Exit Sub
    End If
        
    
    
    
    CargaTxtAux True, False
    ModificaLineas = 2 'Modificar
    'Añadiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR"
    PonerBotonCabecera False

    'BloquearTxt txtAux(2), True 'campo nombre articulo
    PonerFoco txtAux(3)
    Me.DataGrid1.Enabled = False
    
    
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


    
    Cad = "----------------------------------" & vbCrLf & vbCrLf
    Cad = Cad & "Va a pasar a HCO la tarifa - oferta:"
    Cad = Cad & vbCrLf & "Nº:  " & Format(Text1(0).Text, "0000000")
    Cad = Cad & vbCrLf & "Fechas:  " & Format(Text1(1).Text, "dd/mm/yyyy") & " - " & Format(Text1(3).Text, "dd/mm/yyyy")
    If EsTarifa Then
        Cad = Cad & vbCrLf & "Tarifa:  " & Text1(4).Text & " - " & Text2(4).Text
    Else
        Cad = Cad & vbCrLf & "Cliente:  " & Text1(2).Text & " - " & Text2(2).Text
    End If
    Cad = Cad & vbCrLf & vbCrLf & "¿Desea continuar? "
    
    Screen.MousePointer = vbHourglass
    
    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Data1.Recordset.AbsolutePosition
        
        'Abrir frame de informes para pedir datos antes de grabar en el historico
        
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

    ModificaLineas = 3 'Eliminar
    SQL = "¿Seguro que desea eliminar el articulo de la tarifa-oferta?     "
    SQL = SQL & vbCrLf
    SQL = SQL & vbCrLf & "Artículo:  " & Data2.Recordset!codartic & " - " & Data2.Recordset!NomArtic
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Data2.Recordset.AbsolutePosition
        SQL = " WHERE codartic = " & DBSet(Data2.Recordset!codartic, "T")
        SQL = SQL & " and codigo=" & Data1.Recordset!Codigo
        

        'Las lineas
        Conn.Execute "DELETE FROM olitarifaofertalin " & SQL
        

        
        ModificaLineas = 0
        CargaGrid2 DataGrid1, Data2
        SituarDataPosicion Me.Data2, NumRegElim, SQL
        

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
    btnPrimero = 20
    With Me.Toolbar1
        .ImageList = frmppal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        .Buttons(5).Image = 3   'Insertar Nuevo
        .Buttons(6).Image = 4   'Modificar
        .Buttons(7).Image = 5   'Borrar
        .Buttons(10).Image = 10 'Mto Lineas Ofertas
        .Buttons(11).Image = 27 '
        
        'Enero08
        .Buttons(12).Image = 21 'Cerrar orden produccion
        
        
        .Buttons(14).Image = 16 'Imprimir Pedido
        .Buttons(15).Image = 40 'Envio cliente

        .Buttons(17).Image = 15  'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With

    
    Me.FrameCliente.visible = Not EsTarifa
    Me.FrameTarifa.visible = EsTarifa
    Toolbar1.Buttons(11).visible = False
    'If Not EsTarifa And (vEmpresa.codempre <> EmpresaAVAB) Then
    If EmprAVAB > 0 Then   'Si no hay empresa AVAB ni lo ponemos
        If Not EsTarifa And (Not vParamAplic.EsAVAB) Then
            If vUsu.Nivel = 0 Then Toolbar1.Buttons(11).visible = True
        End If
    End If
    
    CheckOrden True
    
    If EsTarifa Then
        NombreTabla = "**"
        Ordenacion = ""
        Label3.ForeColor = vbWhite
        
    Else
        'TOs
        NombreTabla = "-ofertas (TOs)"
        Ordenacion = NombreTabla
        Label3.ForeColor = vbBlue
    End If
    
    Label3.Caption = "Tarifas" & Ordenacion
    Caption = "Mantenimiento tarifas " & NombreTabla
    
    
    
    LimpiarCampos   'Limpia los campos TextBox
   

    NombreTabla = "olitarifaoferta"
    Ordenacion = " ORDER BY codigo "
  
        
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    

    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = Conn
    CadenaConsulta = "Select * from " & NombreTabla & " where codigo= "
    If DatosADevolverBusqueda2 = "" Then
        CadenaConsulta = CadenaConsulta & "-1"
    Else
        CadenaConsulta = CadenaConsulta & DatosADevolverBusqueda2
    End If
    Data1.RecordSource = CadenaConsulta
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

    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
    
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
End Sub


Private Sub LimpiarCampos()
On Error Resume Next

    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""

    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    CheckOrden False
    Set vArti = Nothing
End Sub





'Private Sub frmAlm_DatoSeleccionado(CadenaSeleccion As String)
'    Text1(4).Text = RecuperaValor(CadenaSeleccion, 1)
'    Text2(1).Text = RecuperaValor(CadenaSeleccion, 2)
'End Sub

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
        
            cadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
            cadB = Aux
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
            Text1(0).Text = Format(RecuperaValor(CadenaDevuelta, 1), "0000000")
 
    End If
    Screen.MousePointer = vbDefault
End Sub









Private Sub frmCl_DatoSeleccionado(CadenaSeleccion As String)
    Text1(2).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(2).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmF_Selec(vFecha As Date) 'Calendario Fechas
Dim Indice As Byte
    Indice = CByte(Me.imgFecha(0).Tag)
    Text1(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub







'Private Sub frmPe_DatoSeleccionado2(CadenaSeleccion As String)
'    Text1(4).Text = CadenaSeleccion
'End Sub

Private Sub imgCuentas_Click(Index As Integer)
If Modo = 2 Or Modo = 0 Then Exit Sub

    If Index = 0 Then
    
        Set frmCl = New frmFacClientes
        frmCl.DatosADevolverBusqueda = "0|1|"
        frmCl.Show vbModal
        Set frmCl = Nothing
    
    
    
        'articulo
'            EsCabecera = True
'            Set frmArt = New frmAlmArticulos
'            frmArt.DatosADevolverBusqueda2 = "@1@" 'Poner en modo busqueda
'            frmArt.Show vbModal
'            Set frmArt = Nothing
'            PonerFoco Text1(2)
    
    End If
End Sub

'Private Sub imgBuscar_Click(Index As Integer)
'Dim Indice As Byte
'
'    If Modo = 2 Or Modo = 0 Then Exit Sub
'    Screen.MousePointer = vbHourglass
'    Set frmPe = New frmFacEntPedidos
'    frmPe.DatosADevolverBusqueda2 = "0"
'    frmPe.Show vbModal
'    Set frmPe = Nothing
'
'
'
'    Screen.MousePointer = vbDefault
'
'
'End Sub


Private Sub imgFecha_Click(Index As Integer) 'Abre calendario Fechas
Dim Indice As Byte

   If Modo = 2 Or Modo = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass
   Set frmF = New frmCal
   frmF.Fecha = Now
   If Index = 0 Then
        Indice = 1
   Else
        Indice = 3
   End If
   Me.imgFecha(0).Tag = Indice
   
   PonerFormatoFecha Text1(Indice)
   If Text1(Indice).Text <> "" Then frmF.Fecha = CDate(Text1(Indice).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco Text1(Indice)
End Sub


Private Sub mnBuscar_Click()

    BotonBuscar
End Sub


Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar lineas de Pedido
         BotonEliminarLinea
    Else   'Eliminar Pedido
         BotonEliminar
    End If
End Sub






'Private Sub mnImpOrde_Click()
''Impreme la Orden de Instalacion de un pedido
'Dim cadFormula As String, cadParam As String
'Dim devuelve As String, nomDocu As String
'Dim numParam As Byte
'
'    'Comprobar que hay un pedido seleccionado
'    If Text1(0).Text = "" Then
'        MsgBox "No hay ningún Pedido seleccionado.", vbInformation
'        Exit Sub
'    End If
'
'    'Comprobar que algun Articulo pertenece a la familia de Instalaciones
'    If Not PedidoConInstalaciones Then
'        MsgBox "El Pedido no tiene ningún Artículo que sea Instalación.", vbInformation
'        Exit Sub
'    End If
'
'    '=======================================================================
'    '=============== FORMULA    ============================================
'    cadFormula = ""
'    cadParam = ""
'    numParam = 0
'
'    If Text1(0).Text <> "" Then 'Seleccionar el Pedido
'        devuelve = "{" & NombreTabla & ".numpedcl}=" & Val(Text1(0).Text)
'        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
'    End If
'
'    'Seleccionar solo las lineas de Articulos que son de una familia que es Instalacion
'    devuelve = "{sfamia.instalac}=1"
'    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
'
'    If Not PonerParamRPT(9, cadParam, numParam, nomDocu) Then Exit Sub
'
'    With frmImprimir
'        .NombreRPT = nomDocu
'        .FormulaSeleccion = cadFormula
'        .OtrosParametros = cadParam
'        .NumeroParametros = numParam
'        .SoloImprimir = False
'        .EnvioEMail = False
'        .Opcion = 39
'        .Titulo = ""
'        .Show vbModal
'    End With
'End Sub




Private Sub mnLineas_Click()
    BotonMtoLineas
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
         BotonAnyadir
    End If
End Sub


Private Sub mnORden2_Click(Index As Integer)
Dim I As Integer
    For I = 0 To mnORden2.Count - 1
        mnORden2(I).Checked = I = Index
    Next
    If Modo = 2 Then
        PonerCampos
    ElseIf Modo = 5 And ModificaLineas = 0 Then
        PonerCamposLineas   'Pone los datos de las tablas de lineas de Ofertas
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




'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
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

       
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    'If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
       
    'Si queremos hacer algo ..
    Select Case Index
        Case 1, 3 'Fecha Oferta, Fecha Entrega
            If Text1(Index).Text = "" Then Exit Sub
            
            PonerFormatoFecha Text1(Index)
            
        Case 2, 4
            
            If Text1(Index).Text = "" Then
                Text2(Index).Text = ""
            Else
                If Not PonerFormatoEntero(Text1(Index)) Then
                    Text1(Index).Text = ""
                    Text2(Index).Text = ""
                Else
                    If Index = 2 Then
                        Text2(Index).Text = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", Text1(2).Text)
                    Else
                        Text2(Index).Text = DevuelveDesdeBD(conAri, "nomlista", "starif", "codlista", Text1(4).Text)
                    End If
                    If Text2(Index).Text = "" Then
                        MsgBox "No existe : " & Text1(Index).Text, vbExclamation
                        Text1(Index).Text = ""
                        PonerFoco Text1(Index)
                    End If
                End If
            End If
    End Select
End Sub


Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False)
    If cadB <> "" Then cadB = cadB & " AND "
    cadB = cadB & " (olitarifaoferta.codigo  "
    If EsTarifa Then
        cadB = cadB & " < "
    Else
        cadB = cadB & " > "
    End If
    cadB = cadB & "100000 )"
    
    If chkVistaPrevia = 1 Then
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

        Screen.MousePointer = vbHourglass
        
        
        
        Cad = ParaGrid(Text1(0), 14, "Nº ")
        Cad = Cad & ParaGrid(Text1(1), 15, "F. Ini ")
        Cad = Cad & ParaGrid(Text1(3), 15, "F. Fin ")
        If EsTarifa Then
            Cad = Cad & ParaGrid(Text1(4), 15, "Tarifa")
            Cad = Cad & "Nombre|starif|nomlista|T||37·"
        Else
            Cad = Cad & ParaGrid(Text1(2), 15, "Cliente")
            Cad = Cad & "Nombre|sclien|nomclien|T||37·"
        End If
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        If EsTarifa Then
            frmB.vTabla = NombreTabla & " inner join starif on " & NombreTabla & ".tarifa = starif.codlista"
        Else
            frmB.vTabla = NombreTabla & " inner join sclien on " & NombreTabla & ".codclien = sclien.codclien"
        End If
        frmB.vSQL = cadB
        HaDevueltoDatos = False
    
        frmB.vDevuelve = "0|"
        frmB.vTitulo = "Tarifas - Ofertas"
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri 'Conexión a BD: Ariges
        
        frmB.Show vbModal
        Set frmB = Nothing
        
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass

    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        If EsTarifa Then
            MsgBox "Ninguna tarifa", vbExclamation
        Else
            MsgBox "No hay ningún registro en la tabla  " & NombreTabla, vbInformation
        End If
        Screen.MousePointer = vbDefault
        If Modo = 1 Then
          
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
        
'    'Total
'    b = DataGrid1.Enabled
'    DataGrid1.Enabled = False
'    Tot = 0
'    If Not Data2.Recordset.EOF Then
'        While Not Data2.Recordset.EOF
'            Tot = Tot + Data2.Recordset!kilos
'            Data2.Recordset.MoveNext
'        Wend
'        Data2.Recordset.MoveFirst
'    End If
'    txtTotal.Tag = Tot
'    txtTotal.Text = Format(Tot, FormatoPrecio)
'    DataGrid1.Enabled = b
    
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
    
    'Para que haga el losffocus bien
    Modo = 3
    If EsTarifa Then
        Text1_LostFocus 4
    Else
        Text1_LostFocus 2
    End If
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
Dim B As Boolean

    On Error GoTo EPonerModo

    'Actualiza Iconos Insertar,Modificar,Eliminar
    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
    
    'Fianelmente.

    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    B = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda2 <> "" Then
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
        
        

    'Campo Numero de Albaran siempre bloqueado, excepto si estamos en modo de busqueda
    B = (Modo <> 1)
    BloquearTxt Text1(0), B, True

    B = Modo = 0 Or Modo = 2 Or Modo >= 5
    BloquearTxt Text1(1), B
    BloquearTxt Text1(3), B
    BloquearTxt Text1(5), B
    
    If EsTarifa Then
        BloquearTxt Text1(4), Modo <> 1
        BloquearTxt Text1(2), True
    Else
        BloquearTxt Text1(4), True
        BloquearTxt Text1(2), B
    End If
    
    'Si no es modo lineas Boquear los TxtAux
    For I = 1 To txtAux.Count
        BloquearTxt txtAux(I), (Modo <> 5)
    Next I
  
    
    
    '---------------------------------------------
    B = (Modo <> 0 And Modo <> 2 And Modo <> 5)
    cmdCancelar.visible = B
    cmdAceptar.visible = B
    
    'Las imagenes añadimos el modo 6
    B = B And Modo <> 6
    For I = 0 To Me.imgFecha.Count - 1
        Me.imgFecha(I).Enabled = B
    Next I

    'El check de aceptada
    B = Modo = 1 Or Modo = 3 Or Modo = 4
    Me.Check1.Enabled = B

    Me.chkVistaPrevia.Enabled = (Modo <= 2)
    
    'Los kilos totatels
    B = Modo = 2 Or Modo = 4 Or Modo = 5

    
    'Abrir un coupage cerrado solo para admon
    B = False
    If Modo = 1 Then
        B = True
    Else
        If Modo = 4 Then B = vUsu.Nivel < 1
    End If

    
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
Dim B As Boolean
Dim devuelve As String

    On Error GoTo EDatosOK

    DatosOk = False
    B = CompForm(Me, 1) 'Comprobar formato datos ok
    If Not B Then Exit Function

    B = False
    
    'La fecha fin no ppude ser mayor que la fecha inicio
    If CDate(Text1(1).Text) > CDate(Text1(3).Text) Then
        MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
        Exit Function
    End If
    
    devuelve = ""
    If Me.EsTarifa Then
        If Text1(4).Text = "" Then devuelve = "Tarifa"
    Else
        If Text1(2).Text = "" Then devuelve = "Cliente"
    End If
    If devuelve <> "" Then
        MsgBox "Campo " & devuelve & " no pude estar en blanco", vbExclamation
        Exit Function
    End If
    
    'hAY QUE comprobar que no hay ninguna oferta a este cliente dentro de estas fechas
    'Si es modificar habra que quitar esta oferta
    'Falta###
    If EsTarifa Then
        devuelve = "((fechaini <= '" & Format(Text1(1).Text, FormatoFecha) & "' AND fechafin >='" & Format(Text1(1).Text, FormatoFecha) & "') or "
        devuelve = devuelve & " (fechaini <= '" & Format(Text1(3).Text, FormatoFecha) & "' AND fechafin >='" & Format(Text1(3).Text, FormatoFecha) & "'))"
        If Modo = 4 Then devuelve = devuelve & " AND codigo <> " & Text1(0).Text & " AND codclien is null "
        devuelve = devuelve & " AND tarifa "
    
        devuelve = DevuelveDesdeBD(conAri, "codigo", NombreTabla, devuelve, Text1(4).Text)
        If devuelve <> "" Then
            MsgBox "La oferta se solapa con la oferta :" & devuelve, vbExclamation
            Exit Function
        End If
    End If
    B = True
    DatosOk = B
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLinea() As Boolean
'Comprueba si los datos de una linea son correctos antes de Insertar o Modificar
'una linea del Pedido
Dim B As Boolean
Dim I As Byte
Dim C As String

    On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    B = True

    'Comprobar que los campos NOT NULL tienen valor
    For I = 1 To txtAux.Count
        If txtAux(I).Text = "" Then
            MsgBox "El campo " & txtAux(I).Tag & " no puede ser nulo", vbExclamation
            B = False
            PonerFoco txtAux(I)
            Exit Function
        End If
    Next I
        

    DatosOkLinea = B

EDatosOkLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function









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
            
        Case 10, 11, 12
            '10 Lineas 12, 14
            'IMPRIMIR (14)    y cerrar(12) orden produccion
            ' 11 Subir al AVAB
            '--------------------------------------------------------------------
        
            If Data1.Recordset.EOF Then
                MsgBox "Seleccione una orden de TO", vbExclamation
                Exit Sub
            End If
            
            


                
                If Button.Index = 10 Then
                        mnLineas_Click
                Else
                    If Me.Data2.Recordset.EOF Then
                        MsgBox "No tiene lineas", vbExclamation
                        Exit Sub
                    End If
                    If Button.Index = 11 Then
                        'SI noes uno NO es el AVAB
                        If EsTarifa Then Exit Sub
                        If Val(Me.Data1.Recordset!CodClien) <> 1 Then
                            MsgBox "No es el AVAB", vbExclamation
                            Exit Sub
                        End If
                        
                        'Llegado aqui...
                        'Lanazamos
                        frmOliCrearTO1.SegundoParametro = Data1.Recordset!Codigo
                        frmOliCrearTO1.vOpcion = 2   'Llevar al AVAB
                        frmOliCrearTO1.Show vbModal
                        
                    End If
                End If '=10
        Case 14
'                'Imprimir orden prod
'                With frmImprimir
'                    .ConSubInforme = False
'                    '.FormulaSeleccion = "{olicoupage.codigo} = " & Data1.Recordset!Codigo
'                    .NombreRPT = "rToArticulo.rpt"
'                    .OtrosParametros = "|pNomEmpre=""" & vParam.NombreEmpresa & """|"
'                    .NumeroParametros = 1
'                    .Titulo = "Listado"
'                    .opcion = 2003 'Esta libre
'
'                    .Show vbModal
'                End With

                If mnORden2(2).Checked Or mnORden2(3).Checked Then MsgBox "El orden en el listado de la  TOs será por nombre", vbExclamation
                
                
                CadenaDesdeOtroForm = 0
                If mnORden2(1).Checked Then CadenaDesdeOtroForm = 1
                
                NumRegElim = Abs(EsTarifa)
                frmListado2.Opcion = 17
                frmListado2.Show vbModal

        Case 15 'Imprimir Orden Instalacion
                If Data1.Recordset.EOF Then
                    MsgBox "Seleccione algun dato", vbExclamation
                    Exit Sub
                End If
                CadenaDesdeOtroForm = Data1.Recordset!Codigo
                frmListado2.Opcion = 21
                frmListado2.Show vbModal
          
        Case 17    'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub


Private Sub PonerOpcionesMenu()
Dim J As Byte

    PonerOpcionesMenuGeneral Me
      
'    J = Val(Me.mnGenAlbaran.HelpContextID)
'    If J < vUsu.Nivel Then Me.mnGenAlbaran.Enabled = False
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub
    
    
Private Function InsertarLinea() As Boolean
'Inserta un registro en la tabla de lineas de Pedido: slipre
Dim SQL As String


    On Error GoTo EInsertarLinea

    InsertarLinea = False
    SQL = ""

    If DatosOkLinea() Then 'Lineas de Pedidos
        'Conseguir el siguiente numero de linea
        SQL = "insert into `olitarifaofertalin` (`codigo`,`codartic`,`pivu`,`pivl`,`coste1`,`coste2`,`coste3`,coste4,coste5,"
        SQL = SQL & "`margen`,`pfvu`,`pfvl`) values (" & Data1.Recordset!Codigo & ","
        'codartic pivu
        SQL = SQL & DBSet(txtAux(1).Text, "T") & "," & DBSet(txtAux(3).Text, "N") & ","
        'pivl coste1
        SQL = SQL & DBSet(txtAux(4).Text, "N") & "," & DBSet(txtAux(5).Text, "N") & ","
        'coste2  coste 3 coste4
        SQL = SQL & DBSet(txtAux(6).Text, "N") & "," & DBSet(txtAux(7).Text, "N") & "," & DBSet(txtAux(8).Text, "N") & ","
        'coste5 margen
        SQL = SQL & DBSet(txtAux(9).Text, "N") & "," & DBSet(txtAux(10).Text, "N") & ","
        'pfvu  pfvl
        SQL = SQL & DBSet(txtAux(11).Text, "N") & "," & DBSet(txtAux(12).Text, "N") & ")"
    End If
    
    If SQL <> "" Then
        Conn.Execute SQL
        
        InsertarLinea = True
    End If
    Exit Function
    
EInsertarLinea:
    MuestraError Err.Number, "Insertar Lineas Produccion" & vbCrLf & Err.Description
End Function


Private Function ModificarLinea() As Boolean
'Modifica un registro en la tabla de lineas de Pedido: sliped
Dim SQL As String

    On Error GoTo EModificarLinea

    ModificarLinea = False
    SQL = ""
    
    If DatosOkLinea() Then
        'Creamos la sentencia SQL
        SQL = "UPDATE olitarifaofertalin set pivu = " & DBSet(txtAux(3).Text, "N")
        SQL = SQL & " , pivl = " & DBSet(txtAux(4).Text, "N")
        SQL = SQL & " , coste1 = " & DBSet(txtAux(5).Text, "N")
        SQL = SQL & " , coste2 = " & DBSet(txtAux(6).Text, "N")
        SQL = SQL & " , coste3 = " & DBSet(txtAux(7).Text, "N")
        SQL = SQL & " , coste4 = " & DBSet(txtAux(8).Text, "N")
        SQL = SQL & " , coste5 = " & DBSet(txtAux(9).Text, "N")
        SQL = SQL & " , margen = " & DBSet(txtAux(10).Text, "N")
        SQL = SQL & " , pfvu = " & DBSet(txtAux(11).Text, "N")
        SQL = SQL & " , pfvl = " & DBSet(txtAux(12).Text, "N")
        SQL = SQL & " WHERE codigo =" & Data1.Recordset!Codigo
        SQL = SQL & " AND codartic =" & DBSet(Data2.Recordset!codartic, "T")
                

    
        
    End If
    
    If SQL <> "" Then
        Conn.Execute SQL
        ModificarLinea = True
    End If
    Exit Function
    
EModificarLinea:
    MuestraError Err.Number, "Modificar Lineas Pedido" & vbCrLf & Err.Description
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
        Me.cmdRegresar.Cancel = True
        Me.lblIndicador.Caption = "Líneas " & TituloLinea
        PonerFocoBtn Me.cmdRegresar
    Else
        cmdCancelar.Cancel = True
    End If
    
    'Habilitar las opciones correctas del menu según Modo
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu 'Habilitar las opciones correctas del menu según Nivel de Acceso
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
'IN: enlaza= si carga el grid con valores de la tabla o lo muestra vacio si no enlaza
'    conServidas=si enlaza, se muestra la columna de servidas solo cuando se va a generar el Albaran no completo
Dim B As Boolean
Dim SQL As String

    On Error GoTo ECargaGrid

    B = DataGrid1.Enabled
    
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral vDataGrid, vData, SQL, PrimeraVez
    

    
    CargaGrid2 vDataGrid, vData
    vDataGrid.ScrollBars = dbgAutomatic
    
    
    
    'Aqui cargaremos tb el de las materias primas
    SQL = "select olitarifaofertalin2.codartic,nomartic,costereal,costesimul from olitarifaofertalin2,sartic "
    SQL = SQL & " where olitarifaofertalin2.codartic=sartic.codartic AND codigo = "
    If enlaza Then
        SQL = SQL & Val(Text1(0).Text)
    Else
        SQL = SQL & "-1"
    End If
    CargaGridGnral Me.DataGrid2, Me.Data3, SQL, PrimeraVez
    CargaGrid2 Me.DataGrid2, Data3
    
    B = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2) '5:Modo Mto Lineas (Insertando o Modificando linea)
    vDataGrid.Enabled = Not B
    PrimeraVez = False
    gridCargado = True
    
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim I As Byte

    On Error GoTo ECargaGrid

    vData.Refresh

    Select Case vDataGrid.Name
        Case "DataGrid1" 'Cod. Almacen
                
                
                'sartic.codartic ,nomartic,pivu, pivl,coste1,coste2,coste3,4,5,margen,pfvu,pfvl"
                
                vDataGrid.Columns(0).Caption = "Articulo"
                vDataGrid.Columns(0).Width = 1500

                
                vDataGrid.Columns(1).Caption = "Desc. Artículo"
                vDataGrid.Columns(1).Width = 3000
                
                I = 2
                vDataGrid.Columns(I).Caption = "PIVU"
                vDataGrid.Columns(I).Width = 800
                vDataGrid.Columns(I).Alignment = dbgRight
                vDataGrid.Columns(I).NumberFormat = FormatoPrecio
                
                I = 3
                vDataGrid.Columns(I).Caption = "PIVL"
                vDataGrid.Columns(I).Width = 800
                vDataGrid.Columns(I).Alignment = dbgRight
                vDataGrid.Columns(I).NumberFormat = FormatoPrecio
                                
                I = 4
                vDataGrid.Columns(I).Caption = "%1"
                vDataGrid.Columns(I).Width = 600
                vDataGrid.Columns(I).Alignment = dbgRight
                vDataGrid.Columns(I).NumberFormat = FormatoPorcen
                
                I = 5
                vDataGrid.Columns(I).Caption = "%2"
                vDataGrid.Columns(I).Width = 600
                vDataGrid.Columns(I).Alignment = dbgRight
                vDataGrid.Columns(I).NumberFormat = FormatoPorcen
                
                I = 6
                vDataGrid.Columns(I).Caption = "%3"
                vDataGrid.Columns(I).Width = 600
                vDataGrid.Columns(I).Alignment = dbgRight
                vDataGrid.Columns(I).NumberFormat = FormatoPorcen
                                
                                
                I = 7
                vDataGrid.Columns(I).Caption = "Coste 1"
                vDataGrid.Columns(I).Width = 800
                vDataGrid.Columns(I).Alignment = dbgRight
                vDataGrid.Columns(I).NumberFormat = FormatoPrecio
                                
                                
                I = 8
                vDataGrid.Columns(I).Caption = "Coste 2"
                vDataGrid.Columns(I).Width = 800
                vDataGrid.Columns(I).Alignment = dbgRight
                vDataGrid.Columns(I).NumberFormat = FormatoPrecio
                
                I = 9
                vDataGrid.Columns(I).Caption = "Margen"
                vDataGrid.Columns(I).Width = 600
                vDataGrid.Columns(I).Alignment = dbgRight
                vDataGrid.Columns(I).NumberFormat = FormatoPorcen
                
                
                
                I = 10
                vDataGrid.Columns(I).Caption = "PFVU"
                vDataGrid.Columns(I).Width = 800
                vDataGrid.Columns(I).Alignment = dbgRight
                vDataGrid.Columns(I).NumberFormat = FormatoPrecio
                
                I = 11
                vDataGrid.Columns(I).Caption = "PFVL"
                vDataGrid.Columns(I).Width = 800
                vDataGrid.Columns(I).Alignment = dbgRight
                vDataGrid.Columns(I).NumberFormat = FormatoPrecio
                
                'Los campos no se ven
                'margecom,LitrosUnidad
                vDataGrid.Columns(12).visible = False
                vDataGrid.Columns(13).visible = False
                
        Case "DataGrid2" 'Cod. Almacen
                
                

                
                vDataGrid.Columns(0).Caption = "Articulo"
                vDataGrid.Columns(0).Width = 1500

                
                vDataGrid.Columns(1).Caption = "Desc. Artículo"
                vDataGrid.Columns(1).Width = 3900
                
                I = 2
                vDataGrid.Columns(I).Caption = "Coste"
                vDataGrid.Columns(I).Width = 1500
                vDataGrid.Columns(I).Alignment = dbgRight
                vDataGrid.Columns(I).NumberFormat = FormatoPrecio
                
                I = 3
                vDataGrid.Columns(I).Caption = "Simulacion"
                vDataGrid.Columns(I).Width = 1500
                vDataGrid.Columns(I).Alignment = dbgRight
                vDataGrid.Columns(I).NumberFormat = FormatoPrecio
                
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

    'On Error Resume Next
    On Error GoTo Quitar
    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For I = 1 To txtAux.Count 'TextBox
            txtAux(I).Top = 290
            txtAux(I).visible = visible
        Next I
        cmdaux(1).visible = visible
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid1
            For I = 1 To txtAux.Count
                txtAux(I).Text = ""
                If I <> 2 Then BloquearTxt txtAux(I), False
            Next I
            BloquearTxt txtAux(2), True
        Else 'Vamos a modificar
            For I = 1 To txtAux.Count
                txtAux(I).Text = DataGrid1.Columns(I - 1).Text
                BloquearTxt txtAux(I), I < 3
            Next I
        End If
        

    

        'Fijamos altura(Height) y posición Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid1, 10)
        
        For I = 1 To txtAux.Count
            txtAux(I).Top = alto
            txtAux(I).Height = DataGrid1.RowHeight
        Next I
        'cmdAux(0).Top = alto
        cmdaux(1).Top = alto
        'cmdAux(0).Height = DataGrid1.RowHeight
        cmdaux(1).Height = DataGrid1.RowHeight
        

        txtAux(1).Left = DataGrid1.Left + 360
        If limpiar Then
            'NUEVO
            I = 160
        Else
            I = 30
        End If
        txtAux(1).Width = DataGrid1.Columns(0).Width - I
        
        
        cmdaux(1).Left = txtAux(1).Left + txtAux(1).Width - 35
        'Nom Artic
        If limpiar Then
            txtAux(2).Left = cmdaux(1).Left + cmdaux(1).Width
            txtAux(2).Width = DataGrid1.Columns(1).Width - 90
        Else
           txtAux(2).Left = DataGrid1.Columns(1).Left + 150
           txtAux(2).Width = DataGrid1.Columns(1).Width - 30
        End If
        
        'El resto de txt
        For I = 3 To txtAux.Count
           txtAux(I).Left = DataGrid1.Columns(I - 1).Left + 150
           txtAux(I).Width = DataGrid1.Columns(I - 1).Width - 30
        Next
        
        'Los ponemos Visibles o No
        '--------------------------
        For I = 1 To txtAux.Count
            txtAux(I).visible = visible

        Next I
        
        cmdaux(1).visible = limpiar
    End If
Quitar:
    If Err.Number <> 0 Then Err.Clear
End Sub





Private Sub txtAux_GotFocus(Index As Integer)
Dim cadkey As Integer

    cadkey = ObtenerCadKey(kCampo, Index)
    kCampo = Index
    ConseguirFocoLin txtAux(Index), cadkey
End Sub


Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Modo <> 6 Then 'Modo6: Pasar de Pedido a Albaran
        If Not (Index = 0 And KeyCode = 38) Then KEYdown KeyCode
    End If
End Sub




Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If Modo <> 6 Then
        KEYpress KeyAscii
    Else 'Modo 6: Pasar el Pedido a Albaran
        If KeyAscii = 13 Then 'ENTER
'            PonerServidas
'            ConseguirFoco txtAux(3), Modo
        End If
    End If
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim devuelve As String
Dim Calcular As Boolean
Dim Im As Currency
Dim LitrosUd As Currency

    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub
    
    
    Calcular = False
    Select Case Index
'        Case 0 'Cod Almacen
'            'Comprobar que existe el almacen
'            devuelve = PonerAlmacen(txtAux(Index).Text)
'            txtAux(Index).Text = devuelve
'            'If devuelve = "" Then PonerFoco txtAux(Index)

        Case 1 'Cod. Articulo
            If txtAux(1).Text = "" Then 'Cod Artic
                txtAux(2).Text = "" 'Nom Artic
                Set vArti = Nothing
                Exit Sub
            End If

            If Not vArti Is Nothing Then
                If vArti.Codigo <> txtAux(1).Text Then
                    'Para que lo vuelva a leer
                    Set vArti = Nothing
                Else
                    'Es el mismo articulo. NO hago nada
                    Exit Sub
                End If
            End If
            
            Set vArti = New CArticulo
            txtAux(2).Text = ""
            
            If vArti.LeerDatos(txtAux(1).Text) Then
                vArti.MostrarStatusArtic False
                txtAux(2).Text = vArti.Nombre
                txtAux(3).Text = Format(vArti.PrecioVenta, FormatoPrecio)
                txtAux(4).Text = Format(Round2(vArti.PrecioVenta / vArti.LitrosxUd, 4), FormatoPrecio)
                txtAux(10).Text = Format(vArti.MargenComercial, FormatoPorcen)
                Calcular = True
            Else
                MsgBox "No existe el artículo", vbExclamation
                txtAux(1).Text = ""
                PonerFoco txtAux(1)
            End If
            

    
            
        Case 2 'desc Articulo
            If txtAux(Index).Locked = False Then txtAux(Index).Text = UCase(txtAux(Index).Text)
            
        Case 3, 4, 8, 9, 11, 12 '4 decimmaels
            If txtAux(Index).Text <> "" Then
                If PonerFormatoDecimal(txtAux(Index), 2) Then   'Tipo 2: 10,4
    
                Else
                    txtAux(Index).Text = ""
                    PonerFoco txtAux(Index)
                End If
            End If
            If Index < 11 Then Calcular = True
            
        Case 5, 6, 7, 10
            'Formato porcentaje
            If txtAux(Index).Text <> "" Then
                If PonerFormatoDecimal(txtAux(Index), 4) Then   'Tipo 4: Formato porcentaje
    
                Else
                    txtAux(Index).Text = ""
                    PonerFoco txtAux(Index)
                End If
            End If
            Calcular = True
    End Select
    
    If Calcular Then
            If vArti Is Nothing Then
                LitrosUd = 1
            Else
                LitrosUd = vArti.LitrosxUd
                If LitrosUd = 0 Then LitrosUd = 1
            End If
    
            Im = CalculaImporteLineaTO(ImporteFormateado(txtAux(3).Text), ImporteFormateado(txtAux(5).Text), ImporteFormateado(txtAux(6).Text), ImporteFormateado(txtAux(7).Text) _
                , ImporteFormateado(txtAux(8).Text), ImporteFormateado(txtAux(9).Text), ImporteFormateado(txtAux(10).Text), LitrosUd)
    
            
            txtAux(11).Text = Format(Im, FormatoPrecio)
            'Precio por Litro
            Im = Round2(Im / LitrosUd, 4)
            txtAux(12).Text = Format(Im, FormatoPrecio)
    End If
End Sub


Private Sub BotonMtoLineas()
       
        ModificaLineas = 0
        PonerModo 5
        PonerBotonCabecera True
End Sub


Private Function Eliminar() As Boolean
Dim B As Boolean
Dim Cad As String


    On Error GoTo FinEliminar
        Eliminar = False
        
        'La pasamos a hco
        TituloLinea = "codigo,codclien,fechaini,fechafin,aceptada,tarifa,observaciones"
        Cad = "INSERT INTO olitarifaofertah(" & TituloLinea & ",fecelim) SELECT " & TituloLinea & ",curdate()  from olitarifaoferta where codigo =" & Text1(0).Text
        If Not EjecutaSQL(conAri, Cad, True) Then Exit Function
        
        TituloLinea = "codigo,codartic,pivu,pivl,coste1,coste2,coste3,coste4,coste5,margen,pfvu,pfvl"
        Cad = "INSERT INTO olitarifaofertalinh(" & TituloLinea & ",fecelim) SELECT " & TituloLinea & ",curdate() from olitarifaofertalin where codigo =" & Text1(0).Text
        If Not EjecutaSQL(conAri, Cad, True) Then Exit Function
        
        TituloLinea = "codigo,codartic,costereal,costesimul"
        Cad = "INSERT INTO olitarifaofertalin2h(" & TituloLinea & ",fecelim) SELECT " & TituloLinea & ",curdate() from olitarifaofertalin2 where codigo =" & Text1(0).Text
        If Not EjecutaSQL(conAri, Cad, True) Then Exit Function

        Conn.BeginTrans
        Conn.Execute "Delete from olitarifaofertalin2 where codigo =" & Text1(0).Text
        Conn.Execute "Delete from olitarifaofertalin where codigo =" & Text1(0).Text
        Conn.Execute "Delete from olitarifaoferta where codigo =" & Text1(0).Text
        B = True
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar tarifa-oferta" & vbCrLf, Err.Description
        B = False
    End If
    If Not B Then
        Conn.RollbackTrans
        Eliminar = False
    Else
        Conn.CommitTrans
        Eliminar = True
    End If
    TituloLinea = ""
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
    
    SQL = NombreTabla & ".codigo= " & Val(Text1(0).Text)
    ObtenerWhereCP = SQL
    
    If Err.Number <> 0 Then Err.Clear
End Function


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
    
    SQL = "select sartic.codartic ,nomartic,pivu, pivl,coste1,coste2,coste3"
    SQL = SQL & ",coste4,coste5,margen"
    SQL = SQL & ",pfvu,pfvl,margecom,LitrosUnidad from olitarifaofertalin,sartic "
    SQL = SQL & " where olitarifaofertalin.codartic=sartic.codartic AND "
    
    If enlaza Then
        SQL = SQL & Replace(ObtenerWhereCP, NombreTabla, "olitarifaofertalin")
    Else
        SQL = SQL & " codigo = -1"
    End If
    SQL = SQL & " Order by "  'podriamos poner codartic tb
    
    
    
    If mnORden2(1).Checked Then
        SQL = SQL & " nomartic"
    ElseIf mnORden2(2).Checked Then
        SQL = SQL & " codfamia,codmarca,codunida"
    ElseIf mnORden2(3).Checked Then
        SQL = SQL & " codunida,codmarca,codfamia"
    Else
        
        SQL = SQL & " sartic.codartic"
        
    End If
    
    MontaSQLCarga = SQL
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el Modo en que estemos
Dim B As Boolean

        B = (Modo = 2) Or (Modo = 5 And ModificaLineas = 0)
        'Me.mnOpciones.Enabled = (b Or Modo = 0)

        'Modificar
        Toolbar1.Buttons(6).Enabled = B
        Me.mnModificar.Enabled = B
        'eliminar
        Toolbar1.Buttons(7).Enabled = B
        Me.mnEliminar.Enabled = B
            
            
        'Insertar
        'Si son las tarifas no las puede insertar desde aqui
        If EsTarifa Then
            'No esta habilitado el nuevo
            If Modo <> 5 Then B = False
            
        Else
            B = (B Or Modo = 0)
        
        End If
        Me.Toolbar1.Buttons(5).Enabled = B
        Me.mnNuevo.Enabled = B
        
        B = (Modo = 2)
        'Mantenimiento lineas
        Toolbar1.Buttons(10).Enabled = B
        Me.mnLineas.Enabled = B
        
        If Toolbar1.Buttons(11).visible Then Toolbar1.Buttons(11).Enabled = B
        
        
        'Generar Albaran desde Pedido
'        Toolbar1.Buttons(11).Enabled = b
'        Me.mnGenAlbaran.Enabled = b
        
'        Toolbar1.Buttons(12).Enabled = b
'        Me.mnGeneraFactura.Enabled = b
        
        Toolbar1.Buttons(13).Enabled = B
        
        
        
      
        B = (Modo >= 3) Or Modo = 1
        'Buscar
        Toolbar1.Buttons(1).Enabled = Not B
        Me.mnBuscar.Enabled = Not B
        'Ver Todos
        Toolbar1.Buttons(2).Enabled = Not B
        Me.mnVerTodos.Enabled = Not B
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
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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






Private Function EliminarPedido(numPed As Long) As Boolean
'Eliminar las lineas y la Cabecera de un Pedido. Tablas: scaped, sliped
Dim SQL As String

    On Error GoTo EEliminarPed

     SQL = " WHERE  numpedcl=" & numPed

    'Lineas de Pedido
   ' conn.Execute "Delete from " & NomTablaLineas & sql

    'Cabecera
    Conn.Execute "Delete from " & NombreTabla & SQL

EEliminarPed:
    If Err.Number <> 0 Then
        EliminarPedido = False
    Else
        EliminarPedido = True
    End If
End Function








Private Sub InsertarCabecera()
    
    Text1(0).Text = SugerirCodigoSiguienteStr(NombreTabla, "codigo")
    If InsertarDesdeForm(Me) Then
    
       
            CadenaConsulta = "Select * from " & NombreTabla & " WHERE codigo = " & Text1(0).Text & Ordenacion
            PonerCadenaBusqueda
            'Ponerse en Modo Insertar Lineas
            BotonMtoLineas
            BotonAnyadirLinea
    
    End If

End Sub

'Private Sub ActualizarLineasPedido()
'Dim SQL As String
'    If OpcionConElPedido = 0 Then Exit Sub
'
'    'Si tiene que coger pero no tiene pedido (NO DEBERIA PASAR)
'    If Text1(4).Text = "" Then Exit Sub
'
'    If OpcionConElPedido = 2 Then
'        'Eliminamos los que hubieren
'        SQL = "DELETE FROM sliordpr where codigo = " & Text1(0).Text
'        Conn.Execute SQL
'    End If
'    SQL = "INSERT IGNORE INTO sliordpr(codigo,codalmac,codartic,cantidad)"
'    SQL = SQL & "select " & Text1(0).Text & ",codalmac,codartic,sum(cantidad) from sliped"
'    SQL = SQL & " Where numpedcl = " & Text1(4).Text
'    SQL = SQL & " group by 1,2,3"
'    Conn.Execute SQL
'
'End Sub




Private Sub CheckOrden(Leer As Boolean)
Dim B As Byte
Dim Rc As Byte
Dim NF As Integer

    On Error GoTo eCheckOrden
    NF = FreeFile
    NombreTabla = App.Path & "\OrdenTo.xdf"
    If Leer Then
        Rc = 0
        If Dir(NombreTabla, vbArchive) <> "" Then
            Open NombreTabla For Input As #NF
            Input #NF, NombreTabla
            Close #NF
            If Val(NombreTabla) > 0 Then
                If Val(NombreTabla) <= Me.mnORden2.Count - 1 Then Rc = CByte(Val(NombreTabla))
            End If
        End If
        Me.mnORden2(Rc).Checked = True
    Else
        B = 0
        For Rc = 0 To mnORden2.Count - 1
           If mnORden2(Rc).Checked Then B = Rc
        Next
    
        Open NombreTabla For Output As #NF
        Print #NF, B
        Close #NF
        
    End If
    
    Exit Sub
eCheckOrden:
    Err.Clear
End Sub


