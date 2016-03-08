VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAlmArticulos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Artículos"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11505
   Icon            =   "frmAlmArticulos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   11505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkConso 
      Caption         =   "Stock todos los almacenes"
      Height          =   195
      Left            =   3000
      TabIndex        =   160
      Top             =   7080
      Width           =   2655
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Generar denominación"
      Height          =   375
      Left            =   5880
      TabIndex        =   132
      Top             =   6960
      Visible         =   0   'False
      Width           =   2535
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5700
      Left            =   240
      TabIndex        =   45
      Top             =   1080
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   10054
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      Tab             =   2
      TabsPerRow      =   7
      TabHeight       =   520
      TabCaption(0)   =   "Datos básicos   "
      TabPicture(0)   =   "frmAlmArticulos.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "ChkTrazabilidad"
      Tab(0).Control(1)=   "Frame5"
      Tab(0).Control(2)=   "Text1(30)"
      Tab(0).Control(3)=   "FrameLitrosUd"
      Tab(0).Control(4)=   "chkCtrStock"
      Tab(0).Control(5)=   "txtSumaStock"
      Tab(0).Control(6)=   "Text1(10)"
      Tab(0).Control(7)=   "cboStatus"
      Tab(0).Control(8)=   "Text1(9)"
      Tab(0).Control(9)=   "Text1(8)"
      Tab(0).Control(10)=   "Text1(11)"
      Tab(0).Control(11)=   "Text1(12)"
      Tab(0).Control(12)=   "Text1(6)"
      Tab(0).Control(13)=   "Text2(4)"
      Tab(0).Control(14)=   "Text2(0)"
      Tab(0).Control(15)=   "Text2(1)"
      Tab(0).Control(16)=   "Text2(5)"
      Tab(0).Control(17)=   "Text2(2)"
      Tab(0).Control(18)=   "Text1(4)"
      Tab(0).Control(19)=   "Text1(7)"
      Tab(0).Control(20)=   "Text1(3)"
      Tab(0).Control(21)=   "Text1(2)"
      Tab(0).Control(22)=   "Text1(5)"
      Tab(0).Control(23)=   "Text2(3)"
      Tab(0).Control(24)=   "chkConjunto"
      Tab(0).Control(25)=   "chkSeries"
      Tab(0).Control(26)=   "FrameDatosAlmacen2"
      Tab(0).Control(27)=   "imgPreciosProv"
      Tab(0).Control(28)=   "imgCuentas(0)"
      Tab(0).Control(29)=   "Label1(36)"
      Tab(0).Control(30)=   "lblSumaStocks"
      Tab(0).Control(31)=   "imgFecha(0)"
      Tab(0).Control(32)=   "Label1(16)"
      Tab(0).Control(33)=   "Label1(4)"
      Tab(0).Control(34)=   "Label1(3)"
      Tab(0).Control(35)=   "Label1(2)"
      Tab(0).Control(36)=   "Label1(19)"
      Tab(0).Control(37)=   "Label1(20)"
      Tab(0).Control(38)=   "Label1(9)"
      Tab(0).Control(39)=   "imgCuentas(4)"
      Tab(0).Control(40)=   "imgCuentas(5)"
      Tab(0).Control(41)=   "imgCuentas(1)"
      Tab(0).Control(42)=   "imgCuentas(2)"
      Tab(0).Control(43)=   "Label1(5)"
      Tab(0).Control(44)=   "Label1(6)"
      Tab(0).Control(45)=   "Label1(8)"
      Tab(0).Control(46)=   "Label1(7)"
      Tab(0).Control(47)=   "Label1(17)"
      Tab(0).Control(48)=   "imgCuentas(3)"
      Tab(0).ControlCount=   49
      TabCaption(1)   =   "Otros Datos"
      TabPicture(1)   =   "frmAlmArticulos.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2(3)"
      Tab(1).Control(1)=   "Label2(2)"
      Tab(1).Control(2)=   "Label2(11)"
      Tab(1).Control(3)=   "Label2(1)"
      Tab(1).Control(4)=   "Text1(21)"
      Tab(1).Control(5)=   "Text1(20)"
      Tab(1).Control(6)=   "Text1(19)"
      Tab(1).Control(7)=   "FrameServicios"
      Tab(1).Control(8)=   "Text1(28)"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Componentes"
      TabPicture(2)   =   "frmAlmArticulos.frx":0044
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Line4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label5(0)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label5(1)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label5(2)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label5(3)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label5(4)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label5(5)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Line5"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label5(6)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Label5(7)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Data2"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "DataGrid1"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "cmdAux"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "txtAux(0)"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "txtAux(1)"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "txtAux2"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "txtAux(3)"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "txtAux(4)"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "txtAux(5)"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "txtConjunto(0)"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "txtConjunto(1)"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "txtConjunto(2)"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "txtConjunto(3)"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "txtConjunto(4)"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "txtConjunto(5)"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "cmdActualizarImportes1(0)"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "cmdActualizarImportes1(1)"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).Control(27)=   "txtConjunto(6)"
      Tab(2).Control(27).Enabled=   0   'False
      Tab(2).Control(28)=   "txtConjunto(7)"
      Tab(2).Control(28).Enabled=   0   'False
      Tab(2).Control(29)=   "cmdImprCompo"
      Tab(2).Control(29).Enabled=   0   'False
      Tab(2).ControlCount=   30
      TabCaption(3)   =   "Control instalación / producción"
      TabPicture(3)   =   "frmAlmArticulos.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtAux(2)"
      Tab(3).Control(1)=   "DataGrid2"
      Tab(3).Control(2)=   "Data3"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Stocks"
      TabPicture(4)   =   "frmAlmArticulos.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "DataGrid3"
      Tab(4).Control(1)=   "FrameArtxAlmac"
      Tab(4).Control(2)=   "Text3(2)"
      Tab(4).Control(3)=   "Text2(8)"
      Tab(4).Control(4)=   "Text3(0)"
      Tab(4).Control(5)=   "cmdAlma"
      Tab(4).ControlCount=   6
      TabCaption(5)   =   "Documentos"
      TabPicture(5)   =   "frmAlmArticulos.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label2(0)"
      Tab(5).Control(1)=   "lw1"
      Tab(5).Control(2)=   "Frame4"
      Tab(5).Control(3)=   "FrameDisponible"
      Tab(5).ControlCount=   4
      TabCaption(6)   =   "Ficha técnica"
      TabPicture(6)   =   "frmAlmArticulos.frx":00B4
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Label1(37)"
      Tab(6).Control(1)=   "Label1(38)"
      Tab(6).Control(2)=   "Label1(42)"
      Tab(6).Control(3)=   "FrFichaTec(3)"
      Tab(6).Control(4)=   "FrFichaTec(8)"
      Tab(6).Control(5)=   "FrFichaTec(4)"
      Tab(6).Control(6)=   "Text5(0)"
      Tab(6).Control(7)=   "Text5(1)"
      Tab(6).Control(8)=   "cboTipoArt"
      Tab(6).Control(9)=   "cmdIMpriFT"
      Tab(6).Control(10)=   "FrFichaTec(7)"
      Tab(6).Control(11)=   "FrFichaTec(6)"
      Tab(6).Control(12)=   "FramePalet"
      Tab(6).Control(13)=   "cmdImg"
      Tab(6).ControlCount=   14
      Begin VB.CommandButton cmdImg 
         Height          =   375
         Left            =   -65160
         Picture         =   "frmAlmArticulos.frx":00D0
         Style           =   1  'Graphical
         TabIndex        =   230
         ToolTipText     =   "Editar imagenes adjuntas ficha técnica"
         Top             =   5160
         Width           =   375
      End
      Begin VB.Frame FramePalet 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   3015
         Left            =   -74520
         TabIndex        =   204
         Top             =   1320
         Visible         =   0   'False
         Width           =   10095
         Begin VB.TextBox Text5 
            Height          =   285
            Index           =   16
            Left            =   5520
            TabIndex        =   206
            Text            =   "Text5"
            Top             =   480
            Width           =   2175
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Index           =   9
            Left            =   3720
            TabIndex        =   212
            Text            =   "Text5"
            Top             =   2400
            Width           =   1095
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Index           =   8
            Left            =   2160
            TabIndex        =   211
            Text            =   "Text5"
            Top             =   2400
            Width           =   1095
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Index           =   2
            Left            =   720
            TabIndex        =   205
            Text            =   "Text5"
            Top             =   480
            Width           =   4215
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Index           =   3
            Left            =   720
            TabIndex        =   207
            Text            =   "Text5"
            Top             =   1560
            Width           =   4455
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Index           =   4
            Left            =   5520
            TabIndex        =   208
            Text            =   "Text5"
            Top             =   1560
            Width           =   1095
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Index           =   5
            Left            =   6960
            TabIndex        =   209
            Text            =   "Text5"
            Top             =   1560
            Width           =   1095
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Index           =   6
            Left            =   5280
            TabIndex        =   213
            Text            =   "Text5"
            Top             =   2400
            Width           =   4455
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Index           =   7
            Left            =   720
            TabIndex        =   210
            Text            =   "Text5"
            Top             =   2400
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Codigo DUN"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   255
            Index           =   52
            Left            =   5520
            TabIndex        =   225
            Top             =   180
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Peso bruto"
            Height          =   255
            Index           =   60
            Left            =   3720
            TabIndex        =   223
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Peso neto"
            Height          =   255
            Index           =   59
            Left            =   2160
            TabIndex        =   222
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Paletización"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   255
            Index           =   46
            Left            =   120
            TabIndex        =   221
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Flejado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   255
            Index           =   39
            Left            =   120
            TabIndex        =   220
            Top             =   180
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo"
            Height          =   255
            Index           =   40
            Left            =   720
            TabIndex        =   203
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Peso vacio"
            Height          =   255
            Index           =   41
            Left            =   720
            TabIndex        =   217
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Cajas base"
            Height          =   255
            Index           =   43
            Left            =   5520
            TabIndex        =   216
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Cajas altura"
            Height          =   255
            Index           =   44
            Left            =   6960
            TabIndex        =   215
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Marcado"
            Height          =   255
            Index           =   45
            Left            =   5280
            TabIndex        =   214
            Top             =   2160
            Width           =   855
         End
      End
      Begin VB.Frame FrFichaTec 
         Caption         =   "Embalaje"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   2415
         Index           =   6
         Left            =   -74520
         TabIndex        =   186
         Top             =   1320
         Visible         =   0   'False
         Width           =   6615
         Begin VB.CheckBox chkFichTec 
            Caption         =   "Precinto"
            Height          =   255
            Index           =   4
            Left            =   5040
            TabIndex        =   176
            Top             =   1920
            Width           =   1095
         End
         Begin VB.CheckBox chkFichTec 
            Caption         =   "Cola"
            Height          =   255
            Index           =   3
            Left            =   5040
            TabIndex        =   175
            Top             =   1440
            Width           =   1095
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Index           =   21
            Left            =   240
            TabIndex        =   174
            Text            =   "Text5"
            Top             =   1920
            Width           =   4335
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Index           =   20
            Left            =   2640
            TabIndex        =   173
            Text            =   "Text5"
            Top             =   1200
            Width           =   1935
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Index           =   19
            Left            =   240
            TabIndex        =   172
            Text            =   "Text5"
            Top             =   1200
            Width           =   1935
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Index           =   18
            Left            =   2640
            TabIndex        =   171
            Text            =   "Text5"
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Index           =   17
            Left            =   240
            TabIndex        =   170
            Text            =   "Text5"
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Sellado"
            Height          =   195
            Index           =   58
            Left            =   4800
            TabIndex        =   199
            Top             =   1080
            Width           =   645
         End
         Begin VB.Shape Shape1 
            Height          =   1095
            Left            =   4800
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Marcado"
            Height          =   255
            Index           =   57
            Left            =   240
            TabIndex        =   198
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Unidades /Caja"
            Height          =   195
            Index           =   56
            Left            =   2640
            TabIndex        =   197
            Top             =   960
            Width           =   1110
         End
         Begin VB.Label Label1 
            Caption         =   "Peso vacia"
            Height          =   255
            Index           =   55
            Left            =   240
            TabIndex        =   196
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Volumen"
            Height          =   255
            Index           =   54
            Left            =   2640
            TabIndex        =   195
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Medidas"
            Height          =   255
            Index           =   53
            Left            =   240
            TabIndex        =   194
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame FrFichaTec 
         Caption         =   "Rejilla"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   975
         Index           =   7
         Left            =   -74520
         TabIndex        =   226
         Top             =   1440
         Visible         =   0   'False
         Width           =   6015
         Begin VB.TextBox Text5 
            Height          =   285
            Index           =   24
            Left            =   1680
            TabIndex        =   227
            Text            =   "Text5"
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label6 
            Caption         =   "Ira en el campo caj_medid"
            Height          =   255
            Index           =   4
            Left            =   3840
            TabIndex        =   229
            Top             =   480
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Label Label6 
            Caption         =   "Medidas"
            Height          =   255
            Index           =   3
            Left            =   600
            TabIndex        =   228
            Top             =   480
            Width           =   1695
         End
      End
      Begin VB.CommandButton cmdIMpriFT 
         Height          =   375
         Left            =   -64680
         Picture         =   "frmAlmArticulos.frx":065A
         Style           =   1  'Graphical
         TabIndex        =   224
         ToolTipText     =   "Imprimir ficha técnica"
         Top             =   5160
         Width           =   375
      End
      Begin VB.ComboBox cboTipoArt 
         Height          =   315
         ItemData        =   "frmAlmArticulos.frx":0BE4
         Left            =   -74520
         List            =   "frmAlmArticulos.frx":0BE6
         Style           =   2  'Dropdown List
         TabIndex        =   218
         Tag             =   "Tipo articulo|N|S|||sartic|tipartic|||"
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   1
         Left            =   -70200
         TabIndex        =   163
         Text            =   "Text5"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   0
         Left            =   -71760
         TabIndex        =   162
         Text            =   "Text5"
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdImprCompo 
         Height          =   375
         Left            =   10080
         Picture         =   "frmAlmArticulos.frx":0BE8
         Style           =   1  'Graphical
         TabIndex        =   161
         ToolTipText     =   "Imprimir listado componentes"
         Top             =   4920
         Width           =   375
      End
      Begin VB.CheckBox ChkTrazabilidad 
         Caption         =   "Trazabilidad"
         Height          =   315
         Left            =   -66120
         TabIndex        =   23
         Tag             =   "trazabilidad|N|N|0|1|sartic|trazabilidad||N|"
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox txtConjunto 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   1560
         TabIndex        =   157
         Text            =   "Text5"
         Top             =   5040
         Width           =   975
      End
      Begin VB.TextBox txtConjunto 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   240
         TabIndex        =   156
         Text            =   "Text5"
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Frame Frame5 
         Caption         =   "Frame5"
         Height          =   1575
         Left            =   -64200
         TabIndex        =   153
         Top             =   720
         Visible         =   0   'False
         Width           =   3975
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   26
            Left            =   2280
            MaxLength       =   12
            TabIndex        =   154
            Tag             =   "Precio anual matenimiento|N|S|0|999999.00|sartic|preanuman|###,##0.00|N|"
            Text            =   "Text1"
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Pre. anual mantenimiento"
            Height          =   255
            Index           =   34
            Left            =   720
            TabIndex        =   155
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   30
         Left            =   -65400
         MaxLength       =   15
         TabIndex        =   28
         Tag             =   "Factor conversion|N|N|||sartic|factorconversion|0.0000|N|"
         Text            =   "Tex"
         Top             =   4680
         Width           =   1005
      End
      Begin VB.Frame FrameLitrosUd 
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
         ForeColor       =   &H00972E0B&
         Height          =   375
         Left            =   -66120
         TabIndex        =   150
         Top             =   4080
         Width           =   2055
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   29
            Left            =   960
            MaxLength       =   15
            TabIndex        =   27
            Text            =   "Tex"
            Top             =   0
            Width           =   1035
         End
         Begin VB.Label Label1 
            Caption         =   "Litros  x  UD"
            Height          =   255
            Index           =   35
            Left            =   0
            TabIndex        =   151
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdActualizarImportes1 
         Height          =   375
         Index           =   1
         Left            =   8760
         Picture         =   "frmAlmArticulos.frx":1172
         Style           =   1  'Graphical
         TabIndex        =   149
         ToolTipText     =   "Modificar componente"
         Top             =   4920
         Width           =   375
      End
      Begin VB.CommandButton cmdActualizarImportes1 
         Height          =   375
         Index           =   0
         Left            =   9240
         Picture         =   "frmAlmArticulos.frx":1B74
         Style           =   1  'Graphical
         TabIndex        =   148
         ToolTipText     =   "Actualizar importes"
         Top             =   4920
         Width           =   375
      End
      Begin VB.TextBox txtConjunto 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   7320
         TabIndex        =   146
         Text            =   "Text5"
         Top             =   5040
         Width           =   1215
      End
      Begin VB.TextBox txtConjunto 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   6000
         TabIndex        =   144
         Text            =   "Text5"
         Top             =   5040
         Width           =   1215
      End
      Begin VB.TextBox txtConjunto 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   4680
         TabIndex        =   142
         Text            =   "Text5"
         Top             =   5040
         Width           =   1215
      End
      Begin VB.TextBox txtConjunto 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   9600
         TabIndex        =   140
         Text            =   "Text5"
         Top             =   4680
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtConjunto 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2760
         TabIndex        =   138
         Text            =   "Text5"
         Top             =   5040
         Width           =   1215
      End
      Begin VB.TextBox txtConjunto 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   9960
         TabIndex        =   136
         Text            =   "Text5"
         Top             =   4680
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtAux 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   5
         Left            =   7800
         TabIndex        =   135
         Tag             =   "C|N|S|||||###,##0.0000||"
         Text            =   "Dato2"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.TextBox txtAux 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   4
         Left            =   9120
         TabIndex        =   134
         Tag             =   "C|N|S|||||###,##0.0000||"
         Text            =   "Dato2"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.TextBox txtAux 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   3
         Left            =   6360
         TabIndex        =   133
         Tag             =   "C|N|S|||||###,##0.0000||"
         Text            =   "Dato2"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   28
         Left            =   -74760
         MaxLength       =   60
         TabIndex        =   130
         Tag             =   "Taux|T|S|||sartic|txtauxdocumento|||"
         Top             =   5040
         Width           =   6015
      End
      Begin VB.Frame FrameDisponible 
         Height          =   2295
         Left            =   -66960
         TabIndex        =   118
         Top             =   3120
         Width           =   2655
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   3
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   126
            Text            =   "Text4"
            Top             =   1800
            Width           =   1335
         End
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   2
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   125
            Text            =   "Text4"
            Top             =   1200
            Width           =   1335
         End
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   1
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   124
            Text            =   "Text4"
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   0
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   123
            Text            =   "Text4"
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "Disponible"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   122
            Top             =   1860
            Width           =   855
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   2520
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Label Label4 
            Caption         =   "Stock"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   121
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "Pedidos"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   120
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "Reservas"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   119
            Top             =   720
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   3735
         Left            =   -74880
         TabIndex        =   116
         Top             =   360
         Width           =   855
         Begin MSComctlLib.Toolbar Toolbar2 
            Height          =   2370
            Left            =   0
            TabIndex        =   117
            Top             =   0
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   4180
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Appearance      =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   9
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Tarifas"
                  Object.Tag             =   "0"
                  Style           =   2
                  Value           =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Precios especiales"
                  Object.Tag             =   "1"
                  Style           =   2
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Promociones"
                  Object.Tag             =   "2"
                  Style           =   2
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Pedidos"
                  Object.Tag             =   "3"
                  Style           =   2
               EndProperty
               BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Object.ToolTipText     =   "Precios especiales"
                  Style           =   2
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame FrameServicios 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   2055
         Left            =   -68640
         TabIndex        =   105
         Top             =   360
         Width           =   4575
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   22
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   112
            Text            =   "Text2"
            Top             =   480
            Width           =   3645
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   22
            Left            =   120
            MaxLength       =   3
            TabIndex        =   111
            Tag             =   "Cod. Categoría|T|S|||sartic|codcateg||N|"
            Text            =   "Tex"
            Top             =   480
            Width           =   645
         End
         Begin VB.Frame Frame3 
            Caption         =   "Registro fitosanitarios"
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
            Left            =   120
            TabIndex        =   106
            Top             =   960
            Width           =   4060
            Begin VB.TextBox Text1 
               Height          =   315
               Index           =   23
               Left            =   240
               MaxLength       =   15
               TabIndex        =   108
               Tag             =   "Nº serie|T|S|||sartic|numserie||N|"
               Text            =   "Tex"
               Top             =   430
               Width           =   1965
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   315
               Index           =   24
               Left            =   2400
               MaxLength       =   10
               TabIndex        =   107
               Tag             =   "Fecha vigencia|F|S|||sartic|fecvigen||N|"
               Text            =   "Tex"
               Top             =   430
               Width           =   1400
            End
            Begin VB.Image imgFecha 
               Height          =   240
               Index           =   3
               Left            =   3555
               Picture         =   "frmAlmArticulos.frx":20FE
               ToolTipText     =   "Buscar fecha"
               Top             =   185
               Width           =   240
            End
            Begin VB.Label Label1 
               Caption         =   "Nº"
               Height          =   255
               Index           =   31
               Left            =   240
               TabIndex        =   110
               Top             =   230
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   "Fecha vigencia"
               Height          =   255
               Index           =   32
               Left            =   2400
               TabIndex        =   109
               Top             =   230
               Width           =   1215
            End
         End
         Begin VB.Image imgCuentas 
            Height          =   240
            Index           =   8
            Left            =   1320
            Picture         =   "frmAlmArticulos.frx":2688
            ToolTipText     =   "Buscar familia"
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cod. Categoría"
            Height          =   255
            Index           =   21
            Left            =   120
            TabIndex        =   113
            Top             =   240
            Width           =   1125
         End
      End
      Begin VB.CommandButton cmdAlma 
         Caption         =   "+"
         Height          =   255
         Left            =   -74040
         TabIndex        =   104
         Top             =   3600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Index           =   0
         Left            =   -74760
         MaxLength       =   8
         TabIndex        =   84
         Tag             =   "Código Almacen|N|N|||salmac|codalmac|0|S|"
         Text            =   "Text3"
         Top             =   3600
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   8
         Left            =   -73920
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   102
         Text            =   "Text2"
         Top             =   3600
         Visible         =   0   'False
         Width           =   3045
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   -70800
         MaxLength       =   16
         TabIndex        =   85
         Tag             =   "Cantidad Stock|N|N|||salmac|canstock|#,###,###,##0.00|N|"
         Text            =   "Text3"
         Top             =   3600
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Frame FrameArtxAlmac 
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
         Height          =   3615
         Left            =   -68640
         TabIndex        =   82
         Top             =   960
         Width           =   4455
         Begin VB.TextBox Text3 
            Height          =   315
            Index           =   1
            Left            =   240
            MaxLength       =   15
            TabIndex        =   86
            Tag             =   "Ubicación|T|N|||salmac|ubialmac||N|"
            Text            =   "Text3"
            Top             =   480
            Width           =   990
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   3
            Left            =   2280
            MaxLength       =   16
            TabIndex        =   88
            Tag             =   "Stock Mínimo|N|S|||salmac|stockmin|#,###,###,##0.00|N|"
            Text            =   "Text3"
            Top             =   1200
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   4
            Left            =   240
            MaxLength       =   16
            TabIndex        =   89
            Tag             =   "Punto de Pedido|N|S|||salmac|puntoped|#,###,###,##0.00|N|"
            Text            =   "Text3"
            Top             =   1920
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   6
            Left            =   240
            MaxLength       =   16
            TabIndex        =   87
            Tag             =   "Stock inventario|N|S|||salmac|stockinv|#,###,###,##0.00|N|"
            Text            =   "Text3"
            Top             =   1200
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   7
            Left            =   240
            MaxLength       =   10
            TabIndex        =   91
            Tag             =   "Fecha inventario|F|S|||salmac|fechainv|dd/mm/yyyy|N|"
            Text            =   "Text3"
            Top             =   2640
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   8
            Left            =   2280
            MaxLength       =   8
            TabIndex        =   92
            Tag             =   "Hora Inventario|H|S|||salmac|horainve|hh:mm:ss|N|"
            Text            =   "Text3"
            Top             =   2640
            Width           =   1485
         End
         Begin VB.CheckBox chkInventario 
            Height          =   195
            Left            =   240
            TabIndex        =   94
            Top             =   3240
            Width           =   255
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   5
            Left            =   2280
            MaxLength       =   16
            TabIndex        =   90
            Tag             =   "Stock Máximo|N|S|||salmac|stockmax|#,###,###,##0.00|N|"
            Text            =   "Text3"
            Top             =   1920
            Width           =   1485
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   6
            Left            =   1320
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   83
            Text            =   "Text2"
            Top             =   480
            Width           =   2925
         End
         Begin VB.Label Label3 
            Caption         =   "Realizando Inventario"
            Height          =   255
            Left            =   600
            TabIndex        =   93
            Top             =   3240
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Ubicación"
            Height          =   255
            Index           =   23
            Left            =   240
            TabIndex        =   101
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Stock Mínimo"
            Height          =   255
            Index           =   25
            Left            =   2280
            TabIndex        =   100
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Punto de Pedido"
            Height          =   255
            Index           =   26
            Left            =   240
            TabIndex        =   99
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Stock Inventario"
            Height          =   255
            Index           =   28
            Left            =   240
            TabIndex        =   98
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Inventario"
            Height          =   255
            Index           =   29
            Left            =   240
            TabIndex        =   97
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Hora Inventario"
            Height          =   255
            Index           =   30
            Left            =   2280
            TabIndex        =   96
            Top             =   2400
            Width           =   1095
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   2
            Left            =   1440
            Picture         =   "frmAlmArticulos.frx":308A
            ToolTipText     =   "Buscar fecha"
            Top             =   2400
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Stock Máximo"
            Height          =   255
            Index           =   27
            Left            =   2280
            TabIndex        =   95
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Image imgCuentas 
            Height          =   240
            Index           =   6
            Left            =   3600
            Picture         =   "frmAlmArticulos.frx":3614
            ToolTipText     =   "Buscar almacen"
            Top             =   3240
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgCuentas 
            Height          =   240
            Index           =   7
            Left            =   1080
            Picture         =   "frmAlmArticulos.frx":4016
            ToolTipText     =   "Buscar ubicación"
            Top             =   180
            Width           =   240
         End
      End
      Begin VB.TextBox Text1 
         Height          =   975
         Index           =   19
         Left            =   -74760
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Tag             =   "Texto para Ventas|T|S|||sartic|textoven|||"
         Top             =   840
         Width           =   6015
      End
      Begin VB.TextBox Text1 
         Height          =   855
         Index           =   20
         Left            =   -74760
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   33
         Tag             =   "Texto para compras|T|S|||sartic|textocom|||"
         Top             =   2160
         Width           =   6015
      End
      Begin VB.TextBox Text1 
         Height          =   855
         Index           =   21
         Left            =   -74760
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Tag             =   "Control de instalación|T|S|||sartic|controli|||"
         Top             =   3465
         Width           =   6015
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   2
         Left            =   -74040
         MaxLength       =   60
         TabIndex        =   68
         Text            =   "Dat"
         Top             =   2880
         Visible         =   0   'False
         Width           =   7035
      End
      Begin VB.TextBox txtAux2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   290
         Left            =   1800
         TabIndex        =   66
         Text            =   "Dato2"
         Top             =   3180
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox txtAux 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   1
         Left            =   4680
         TabIndex        =   65
         Tag             =   "C|N|N|||||###,##0.00000||"
         Text            =   "Dato2"
         Top             =   3180
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   0
         Left            =   720
         TabIndex        =   64
         Text            =   "Dat"
         Top             =   3180
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Left            =   1560
         TabIndex        =   63
         Top             =   3180
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.CheckBox chkCtrStock 
         Caption         =   "¿Control de stock?"
         Height          =   315
         Left            =   -66120
         TabIndex        =   26
         Tag             =   "Control de stock|N|N|0|1|sartic|ctrstock||N|"
         Top             =   3720
         Width           =   1815
      End
      Begin VB.TextBox txtSumaStock 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         Height          =   315
         Left            =   -66000
         Locked          =   -1  'True
         MaxLength       =   13
         TabIndex        =   58
         Text            =   "Text1"
         Top             =   5280
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   10
         Left            =   -66120
         MaxLength       =   10
         TabIndex        =   19
         Tag             =   "Fecha de Alta|F|N|||sartic|fecaltas|dd/mm/yyyy|N|"
         Text            =   "Text1"
         Top             =   1215
         Width           =   1335
      End
      Begin VB.ComboBox cboStatus 
         Height          =   315
         Left            =   -66120
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Tag             =   "Situación Artículo|N|N|||sartic|codstatu||N|"
         Top             =   1575
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   9
         Left            =   -66120
         MaxLength       =   18
         TabIndex        =   18
         Tag             =   "Código Asociación|T|S|||sartic|codtelem||N|"
         Text            =   "Text1"
         Top             =   855
         Width           =   1830
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   8
         Left            =   -66120
         MaxLength       =   13
         TabIndex        =   17
         Tag             =   "Código de Barras|T|S|||sartic|codigoea||N|"
         Text            =   "Text1"
         Top             =   495
         Width           =   1830
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   11
         Left            =   -66120
         MaxLength       =   8
         TabIndex        =   21
         Tag             =   "Días de garantia|N|N|0|99999|sartic|garantia||N|"
         Text            =   "Text1"
         Top             =   1935
         Width           =   990
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   12
         Left            =   -66120
         MaxLength       =   8
         TabIndex        =   22
         Tag             =   "Unidades por caja|N|N|||sartic|unicajas||N|"
         Text            =   "Text1"
         Top             =   2295
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   6
         Left            =   -72840
         MaxLength       =   2
         TabIndex        =   7
         Tag             =   "Cod. Tipo Artículo|T|N|||sartic|codtipar||N|"
         Text            =   "Te"
         Top             =   1935
         Width           =   765
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   4
         Left            =   -72000
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   42
         Text            =   "Text2"
         Top             =   1935
         Width           =   3285
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   0
         Left            =   -72000
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   35
         Text            =   "Text2"
         Top             =   495
         Width           =   3285
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   1
         Left            =   -72000
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   36
         Text            =   "Text2"
         Top             =   855
         Width           =   3285
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   5
         Left            =   -72000
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   44
         Text            =   "Text2"
         Top             =   2295
         Width           =   3285
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   2
         Left            =   -72000
         Locked          =   -1  'True
         MaxLength       =   35
         TabIndex        =   37
         Text            =   "Text2"
         Top             =   1215
         Width           =   3285
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   -72840
         MaxLength       =   4
         TabIndex        =   5
         Tag             =   "Cod. Marca|N|N|0|9999|sartic|codmarca|0000|N|"
         Text            =   "Text1"
         Top             =   1215
         Width           =   765
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   7
         Left            =   -72840
         MaxLength       =   1
         TabIndex        =   8
         Tag             =   "Tipo de IVA|N|N|0|9|sartic|codigiva||N|"
         Text            =   "T"
         Top             =   2295
         Width           =   765
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   -72840
         MaxLength       =   4
         TabIndex        =   4
         Tag             =   "Cod. Familia|N|N|0|9999|sartic|codfamia|0000|N|"
         Text            =   "Text1"
         Top             =   855
         Width           =   765
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   -72840
         MaxLength       =   6
         TabIndex        =   3
         Tag             =   "Cod. Proveedor|N|N|0|999999|sartic|codprove|000000|N|"
         Text            =   "Text1"
         Top             =   495
         Width           =   765
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   -72840
         MaxLength       =   2
         TabIndex        =   6
         Tag             =   "Cod. Tipo Unidad|N|N|0|99|sartic|codunida|00|N|"
         Text            =   "Text1"
         Top             =   1575
         Width           =   765
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   3
         Left            =   -72000
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   40
         Text            =   "Text2"
         Top             =   1575
         Width           =   3285
      End
      Begin VB.CheckBox chkConjunto 
         Caption         =   "Tiene componentes"
         Height          =   315
         Left            =   -66120
         TabIndex        =   25
         Tag             =   "¿Es conjunto?|N|N|0|1|sartic|conjunto||N|"
         Top             =   3360
         Width           =   1935
      End
      Begin VB.CheckBox chkSeries 
         Caption         =   "¿Control Nº Serie?"
         Height          =   315
         Left            =   -66120
         TabIndex        =   24
         Tag             =   "¿Control nº serie?|N|N|0|1|sartic|nseriesn||N|"
         Top             =   3000
         Width           =   1815
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   4170
         Left            =   -74280
         TabIndex        =   69
         Top             =   960
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   7355
         _Version        =   393216
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
      Begin MSAdodcLib.Adodc Data3 
         Height          =   330
         Left            =   -66360
         Top             =   720
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3675
         Left            =   240
         TabIndex        =   67
         Top             =   600
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   6482
         _Version        =   393216
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
      Begin MSAdodcLib.Adodc Data2 
         Height          =   330
         Left            =   8760
         Top             =   840
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
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
      Begin VB.Frame FrameDatosAlmacen2 
         Caption         =   "Datos Relacionados con Almacen"
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
         TabIndex        =   74
         Top             =   2760
         Width           =   8655
         Begin VB.TextBox txtPVPIVA 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Height          =   315
            Left            =   7320
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   128
            Text            =   "Text1"
            Top             =   1680
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   27
            Left            =   7440
            MaxLength       =   10
            TabIndex        =   14
            Tag             =   "Fecha último cambio P.V.P.|F|S|||sartic|ultfecpvp|dd/mm/yyyy|N|"
            Text            =   "Text1"
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   25
            Left            =   1680
            MaxLength       =   6
            TabIndex        =   16
            Tag             =   "Margen comercia|N|S|0|999.00|sartic|margecom|##0.00|N|"
            Text            =   "Text1"
            Top             =   2280
            Width           =   1200
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   18
            Left            =   4440
            MaxLength       =   10
            TabIndex        =   13
            Tag             =   "Fecha última compra|F|S|||sartic|ultfecco|dd/mm/yyyy|N|"
            Text            =   "Text1"
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   17
            Left            =   1680
            MaxLength       =   12
            TabIndex        =   15
            Tag             =   "Precio Venta al público|N|N|0|999999.0000|sartic|preciove|###,##0.0000|N|"
            Text            =   "Text1"
            Top             =   1650
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   16
            Left            =   4440
            MaxLength       =   12
            TabIndex        =   10
            Tag             =   "Precio Standard|N|S|0|999999.0000|sartic|preciost|###,##0.0000|N|"
            Text            =   "Text1"
            Top             =   270
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   15
            Left            =   1800
            MaxLength       =   12
            TabIndex        =   12
            Tag             =   "Precio Ultima Compra|N|S|0|999999.0000|sartic|preciouc|###,##0.0000|N|"
            Text            =   "Text1"
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   14
            Left            =   7440
            MaxLength       =   12
            TabIndex        =   11
            Tag             =   "Precio Medio Acumulado|N|S|0|999999.0000|sartic|precioma|###,##0.0000|N|"
            Text            =   "Text1"
            Top             =   270
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   13
            Left            =   1800
            MaxLength       =   12
            TabIndex        =   9
            Tag             =   "Precio Medio Ponderado|N|S|0|999999.0000|sartic|preciomp|###,##0.0000|N|"
            Text            =   "Text1"
            Top             =   270
            Width           =   1095
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00C0C0C0&
            X1              =   120
            X2              =   8520
            Y1              =   2040
            Y2              =   2040
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00C0C0C0&
            X1              =   120
            X2              =   8520
            Y1              =   1530
            Y2              =   1530
         End
         Begin VB.Label Label1 
            Caption         =   "P.V.P. + IVA"
            Height          =   255
            Index           =   24
            Left            =   5760
            TabIndex        =   129
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Últ. fec. cambio P.V.P."
            Height          =   255
            Index           =   22
            Left            =   5760
            TabIndex        =   127
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Margen Comercial"
            Height          =   255
            Index           =   33
            Left            =   240
            TabIndex        =   81
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   1
            Left            =   4200
            Picture         =   "frmAlmArticulos.frx":4A18
            ToolTipText     =   "Buscar fecha"
            Top             =   960
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Actualiza. coste"
            Height          =   255
            Index           =   15
            Left            =   3000
            TabIndex        =   80
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "P.V.P."
            Height          =   255
            Index           =   14
            Left            =   480
            TabIndex        =   79
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Pr.Standard"
            Height          =   255
            Index           =   13
            Left            =   3120
            TabIndex        =   78
            Top             =   300
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Precio coste"
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   77
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Pr. Med  Acumulado"
            Height          =   255
            Index           =   11
            Left            =   5760
            TabIndex        =   76
            Top             =   300
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Pr Med. Ponderado"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   75
            Top             =   300
            Width           =   1575
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid3 
         Height          =   4695
         Left            =   -74760
         TabIndex        =   103
         Top             =   600
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   8281
         _Version        =   393216
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
      Begin MSComctlLib.ListView lw1 
         Height          =   5055
         Left            =   -74040
         TabIndex        =   114
         Top             =   360
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   8916
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
      Begin VB.Frame FrFichaTec 
         Caption         =   "Etiqueta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   2055
         Index           =   4
         Left            =   -74520
         TabIndex        =   185
         Top             =   1440
         Visible         =   0   'False
         Width           =   6015
         Begin VB.TextBox Text5 
            Height          =   285
            Index           =   10
            Left            =   1800
            TabIndex        =   179
            Text            =   "Text5"
            Top             =   1320
            Width           =   3855
         End
         Begin VB.CheckBox chkFichTec 
            Caption         =   "Reimprime "
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   178
            Top             =   1320
            Width           =   1335
         End
         Begin VB.CheckBox chkFichTec 
            Caption         =   "Imprime EAN"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   177
            Top             =   600
            Width           =   2655
         End
         Begin VB.Label Label6 
            Caption         =   "Texto reimpresion"
            Height          =   255
            Index           =   0
            Left            =   1800
            TabIndex        =   188
            Top             =   1080
            Width           =   1695
         End
      End
      Begin VB.Frame FrFichaTec 
         Caption         =   "Retráctil"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   2055
         Index           =   8
         Left            =   -74520
         TabIndex        =   200
         Top             =   1440
         Visible         =   0   'False
         Width           =   8895
         Begin VB.TextBox Text5 
            Height          =   285
            Index           =   23
            Left            =   1800
            TabIndex        =   180
            Text            =   "Text5"
            Top             =   480
            Width           =   1575
         End
         Begin VB.CheckBox chkFichTec 
            Caption         =   "Reimprime "
            Height          =   255
            Index           =   5
            Left            =   360
            TabIndex        =   187
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Index           =   22
            Left            =   1800
            TabIndex        =   182
            Text            =   "Text5"
            Top             =   1320
            Width           =   6735
         End
         Begin VB.Label Label6 
            Caption         =   "Medidas"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   202
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "Texto reimpresion"
            Height          =   255
            Index           =   1
            Left            =   1800
            TabIndex        =   201
            Top             =   1080
            Width           =   1695
         End
      End
      Begin VB.Frame FrFichaTec 
         Caption         =   "Tapón"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   2295
         Index           =   3
         Left            =   -74520
         TabIndex        =   184
         Top             =   1440
         Visible         =   0   'False
         Width           =   7455
         Begin VB.TextBox Text5 
            Height          =   285
            Index           =   15
            Left            =   1440
            TabIndex        =   169
            Text            =   "Text5"
            Top             =   1800
            Width           =   5895
         End
         Begin VB.CheckBox chkFichTec 
            Caption         =   "Serigrafia"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   168
            Top             =   1800
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Index           =   14
            Left            =   4080
            TabIndex        =   167
            Text            =   "Text5"
            Top             =   1200
            Width           =   1455
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Index           =   13
            Left            =   360
            TabIndex        =   166
            Text            =   "Text5"
            Top             =   1200
            Width           =   2895
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Index           =   12
            Left            =   4080
            TabIndex        =   165
            Text            =   "Text5"
            Top             =   480
            Width           =   2655
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Index           =   11
            Left            =   360
            TabIndex        =   164
            Text            =   "Text5"
            Top             =   480
            Width           =   2895
         End
         Begin VB.Label Label1 
            Caption         =   "Texto serigrafia"
            Height          =   255
            Index           =   51
            Left            =   1440
            TabIndex        =   193
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Color"
            Height          =   255
            Index           =   50
            Left            =   4080
            TabIndex        =   192
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Faldón"
            Height          =   255
            Index           =   49
            Left            =   360
            TabIndex        =   191
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Medidas"
            Height          =   255
            Index           =   48
            Left            =   4080
            TabIndex        =   190
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Aplicacion"
            Height          =   255
            Index           =   47
            Left            =   360
            TabIndex        =   189
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Image imgPreciosProv 
         Height          =   240
         Left            =   -68640
         Picture         =   "frmAlmArticulos.frx":4FA2
         ToolTipText     =   "Precios proveedores"
         Top             =   480
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo producto"
         Height          =   255
         Index           =   42
         Left            =   -74520
         TabIndex        =   219
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Peso bruto"
         Height          =   255
         Index           =   38
         Left            =   -70200
         TabIndex        =   183
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Peso neto"
         Height          =   255
         Index           =   37
         Left            =   -71760
         TabIndex        =   181
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Formato"
         Height          =   255
         Index           =   7
         Left            =   1560
         TabIndex        =   159
         Top             =   4800
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Componentes"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   158
         Top             =   4800
         Width           =   1215
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   0
         Left            =   -73140
         Picture         =   "frmAlmArticulos.frx":59A4
         ToolTipText     =   "Buscar familia"
         Top             =   480
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Factor conversión"
         Height          =   255
         Index           =   36
         Left            =   -66120
         TabIndex        =   152
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         X1              =   4680
         X2              =   8520
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Label Label5 
         Caption         =   "Diferencia"
         Height          =   255
         Index           =   5
         Left            =   7320
         TabIndex        =   147
         Top             =   4800
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "PVP real"
         Height          =   255
         Index           =   4
         Left            =   6000
         TabIndex        =   145
         Top             =   4800
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "PVP articulo"
         Height          =   255
         Index           =   3
         Left            =   4680
         TabIndex        =   143
         Top             =   4800
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Diferencia"
         Height          =   255
         Index           =   2
         Left            =   9600
         TabIndex        =   141
         Top             =   4440
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Coste real"
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   139
         Top             =   4800
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Coste articulo"
         Height          =   255
         Index           =   0
         Left            =   9960
         TabIndex        =   137
         Top             =   4440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         X1              =   240
         X2              =   4080
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Label Label2 
         Caption         =   "Texto auxiliar documentos"
         Height          =   240
         Index           =   1
         Left            =   -74760
         TabIndex        =   131
         Top             =   4800
         Width           =   2655
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
         ForeColor       =   &H00000080&
         Height          =   300
         Index           =   0
         Left            =   -66960
         TabIndex        =   115
         Top             =   480
         Width           =   2865
      End
      Begin VB.Label Label2 
         Caption         =   "Texto para Ventas"
         Height          =   240
         Index           =   11
         Left            =   -74760
         TabIndex        =   73
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Texto para Compras"
         Height          =   240
         Index           =   2
         Left            =   -74760
         TabIndex        =   72
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Control de Instalación"
         Height          =   240
         Index           =   3
         Left            =   -74760
         TabIndex        =   71
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label lblSumaStocks 
         Caption         =   "Suma Stock Almacenes"
         Height          =   195
         Left            =   -66000
         TabIndex        =   59
         Top             =   5040
         Width           =   1695
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   -66600
         Picture         =   "frmAlmArticulos.frx":63A6
         ToolTipText     =   "Buscar fecha"
         Top             =   1215
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha de Alta"
         Height          =   255
         Index           =   16
         Left            =   -67800
         TabIndex        =   57
         Top             =   1215
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Situación Artículo"
         Height          =   255
         Index           =   4
         Left            =   -67800
         TabIndex        =   56
         Top             =   1575
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo Asociación"
         Height          =   255
         Index           =   3
         Left            =   -67800
         TabIndex        =   55
         Top             =   855
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo de Barras"
         Height          =   255
         Index           =   2
         Left            =   -67800
         TabIndex        =   54
         Top             =   495
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Dias de Garantia"
         Height          =   255
         Index           =   19
         Left            =   -67800
         TabIndex        =   53
         Top             =   1935
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Unidades por Caja"
         Height          =   255
         Index           =   20
         Left            =   -67800
         TabIndex        =   52
         Top             =   2295
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. Tipo Artículo"
         Height          =   255
         Index           =   9
         Left            =   -74520
         TabIndex        =   51
         Top             =   1935
         Width           =   1335
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   4
         Left            =   -73140
         Picture         =   "frmAlmArticulos.frx":6930
         ToolTipText     =   "Buscar tipo artículo"
         Top             =   1935
         Width           =   240
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   5
         Left            =   -73140
         Picture         =   "frmAlmArticulos.frx":7332
         ToolTipText     =   "Buscar tipo IVA"
         Top             =   2295
         Width           =   240
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   1
         Left            =   -73140
         Picture         =   "frmAlmArticulos.frx":7D34
         ToolTipText     =   "Buscar familia"
         Top             =   855
         Width           =   240
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   2
         Left            =   -73140
         Picture         =   "frmAlmArticulos.frx":8736
         ToolTipText     =   "Buscar marca"
         Top             =   1215
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cod.  Proveedor"
         Height          =   255
         Index           =   5
         Left            =   -74520
         TabIndex        =   50
         Top             =   495
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. Familia"
         Height          =   255
         Index           =   6
         Left            =   -74520
         TabIndex        =   49
         Top             =   855
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de I.V.A."
         Height          =   255
         Index           =   8
         Left            =   -74520
         TabIndex        =   48
         Top             =   2295
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. Marca"
         Height          =   255
         Index           =   7
         Left            =   -74520
         TabIndex        =   47
         Top             =   1215
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. Tipo Unidad"
         Height          =   255
         Index           =   17
         Left            =   -74520
         TabIndex        =   46
         Top             =   1575
         Width           =   1335
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   3
         Left            =   -73140
         Picture         =   "frmAlmArticulos.frx":9138
         ToolTipText     =   "Buscar tipo unidad"
         Top             =   1575
         Width           =   240
      End
   End
   Begin VB.Frame Frame2 
      Height          =   620
      Left            =   240
      TabIndex        =   60
      Top             =   410
      Width           =   11055
      Begin VB.ComboBox cboArticuloVarios 
         Height          =   315
         Left            =   9600
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Tag             =   "Artículo de Varios|N|N|||sartic|artvario||N|"
         Top             =   210
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   4040
         MaxLength       =   40
         TabIndex        =   1
         Tag             =   "Denominación Artículo|T|N|||sartic|nomartic||N|"
         Text            =   "Text1"
         Top             =   210
         Width           =   4245
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   1040
         MaxLength       =   16
         TabIndex        =   0
         Tag             =   "Código Artículo|T|N|||sartic|codartic||S|"
         Text            =   "Text1"
         Top             =   210
         Width           =   1710
      End
      Begin VB.Label Label1 
         Caption         =   "Artículo Varios"
         Height          =   255
         Index           =   18
         Left            =   8490
         TabIndex        =   70
         Top             =   220
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Denominación"
         Height          =   255
         Index           =   1
         Left            =   2950
         TabIndex        =   62
         Top             =   220
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Código Art."
         Height          =   255
         Index           =   0
         Left            =   200
         TabIndex        =   61
         Top             =   225
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   240
      TabIndex        =   38
      Top             =   6840
      Width           =   2655
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   120
         TabIndex        =   39
         Top             =   180
         Width           =   2235
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10200
      TabIndex        =   30
      Top             =   6960
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9000
      TabIndex        =   29
      Top             =   6960
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   6120
      Top             =   6960
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      TabIndex        =   41
      Top             =   0
      Width           =   11505
      _ExtentX        =   20294
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
            ImageIndex      =   1
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
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Stocks Almacenes"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Componentes"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Instalaciones"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
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
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkProdVenta 
         Caption         =   "Producto venta"
         Height          =   195
         Left            =   9480
         TabIndex        =   231
         Top             =   120
         Width           =   1575
      End
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   7320
         TabIndex        =   43
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data4 
      Height          =   330
      Left            =   5520
      Top             =   6960
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      Left            =   10200
      TabIndex        =   31
      Top             =   6960
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
   Begin VB.Menu mnMantenimientos 
      Caption         =   "&Mantenimientos"
      Begin VB.Menu mnMtoStocksAlm 
         Caption         =   "&Stocks Almacenes"
         Shortcut        =   ^K
      End
      Begin VB.Menu mnMtoConjuntos 
         Caption         =   "&Conjuntos"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnMtoInstalaciones 
         Caption         =   "&Instalaciones"
         Shortcut        =   ^I
      End
   End
End
Attribute VB_Name = "frmAlmArticulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Public DatosADevolverBusqueda2 As String    'Tendra el nº de text que quiere que devuelva, empipados
Public DeConsulta As Boolean 'Muestra Form para consulta, solo buscar y ver todos activos

Public Event DatoSeleccionado(CadenaSeleccion As String)
Public ParaVenta As Boolean


Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1

Private WithEvents frmP As frmComProveedores
Attribute frmP.VB_VarHelpID = -1
Private WithEvents frmM As frmAlmMarcas 'Marcas de Artículos
Attribute frmM.VB_VarHelpID = -1
Private WithEvents frmTU As frmAlmTipoUnidad
Attribute frmTU.VB_VarHelpID = -1
Private WithEvents frmTA As frmAlmTipoArticulo
Attribute frmTA.VB_VarHelpID = -1
Private WithEvents frmFA As frmAlmFamiliaArticulo
Attribute frmFA.VB_VarHelpID = -1
Private WithEvents frmA As frmAlmAlPropios 'Almacenes Propios
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmUbic As frmAlmUbicaciones 'ubicaciones de almacen
Attribute frmUbic.VB_VarHelpID = -1
Private WithEvents frmCat As frmAlmCategorias 'categorias articulo (control de lotes(S/N))
Attribute frmCat.VB_VarHelpID = -1



Private Modo As Byte
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar un registro
'   5.-  Mantenimiento Lineas de Articulos x Almacen
'   6.-  Mantenimiento Lineas de Componentes de Conjuntos
'   7.-  Mantenimiento Lineas de Control de Instalaciones
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
Private ModoAnterior As Byte

Private ModoFrame As Byte
'ModoFrame: 0.-Inicio, 3.-Insertar, 4.-Modificar

Private CadenaConsulta As String
'SQL de la tabla principal del formulario

Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private kCampo As Integer

Dim btnAnyadir As Byte
'Variable que indica el número del Boton  Anyadir en la Toolbar1

Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1

'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Dim ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim PrimeraVez As Boolean

Private TagText3 As String

'NUEVO: JULIO 2007. PARA BUSCAR POR CHECKS TB
'------------------------------------------------
Private BuscaChekc As String

'NUevo: Nov 2008
' Hay un campo para pintar el PVP con el IVA
' Guardaremos el tipo de iva y el % (para no tener que recaluclarlo cada ve
Private mPorIva As String

Private PriVezForm As Boolean

Private MostrarSolapa As Boolean

Private HayQueRecalcularPesos As Boolean


Private Sub cboArticuloVarios_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cboStatus_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Check1_Click()

End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)

End Sub

Private Sub cboTipoArt_Click()
Dim N As Integer
    ''''0   GENERICA
''''1   ACEITE
''''2   ENVASE
''''3   TAPON
''''4   E.FRONTAL
''''5   E.DORSAL
''''6   EMBALAJE
''''7   REJILLA
''''8   RETRACTIL

    If Modo > 2 Then
        'Ha cambiado
         If Me.chkConjunto.Value = 1 Then
            If cboTipoArt.ItemData(cboTipoArt.ListIndex) <> 1 Then
                MsgBox "Articulo tiene componentes", vbExclamation
                PonerCamposTipoArt 1
                N = 0
            End If
        Else
            If cboTipoArt.ListIndex < 0 Then
                N = 0
            Else
                N = cboTipoArt.ItemData(cboTipoArt.ListIndex)
            End If
        End If
        
        PonerFramesFichaTecnica2 N, chkConjunto.Value = 1
    
    End If
End Sub

Private Sub PonerCamposTipoArt(KItemData As Integer)
Dim J As Integer
    For J = 0 To cboTipoArt.ListCount - 1
        If cboTipoArt.ItemData(J) = KItemData Then
            Me.cboTipoArt.ListIndex = J
            Exit For
        End If
    Next
    

End Sub


Private Sub chkConjunto_Click()

 If Modo = 1 Then
    CheckCadenaBusqueda chkConjunto, BuscaChekc
 Else
    If Modo = 3 Then
        'INSERTANDO
        If Me.cboTipoArt.ListIndex >= 0 Then
            'Es aceite
            
            PonerFramesFichaTecnica2 cboTipoArt.ItemData(cboTipoArt.ListIndex), chkConjunto.Value = 1
        
        End If
        If Me.chkConjunto.Value = 1 Then
            'Pongo el flejado
            Text5(2).Text = "FILM EXTENSIBLE MAQUINA FLEJADORA"
            Text5(3).Text = "EUROPALET 80x120"
        Else
            Text5(2).Text = ""
            Text5(3).Text = ""
        End If
        
    End If
 End If
End Sub

Private Sub chkConjunto_GotFocus()
     ConseguirfocoChk Modo
End Sub

Private Sub chkConjunto_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkCtrStock_Click()
 If Modo = 1 Then CheckCadenaBusqueda chkctrstock, BuscaChekc

End Sub

Private Sub chkctrstock_GotFocus()
     ConseguirfocoChk Modo
End Sub

Private Sub chkctrstock_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub chkFichTec_Click(Index As Integer)
Dim Ot As Integer
    If Index = 3 Or Index = 4 Then
        Ot = 3
        If Index = 3 Then Ot = 4
        If chkFichTec(Index).Value = 0 Then
            chkFichTec(Ot).Value = 1
        Else
            chkFichTec(Ot).Value = 0
        End If
    End If
        
End Sub

Private Sub chkFichTec_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkInventario_Click()
    If Modo = 1 Then CheckCadenaBusqueda chkInventario, BuscaChekc
End Sub

Private Sub chkInventario_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkInventario_LostFocus()
    PonerFocoBtn Me.cmdAceptar
End Sub

Private Sub chkSeries_Click()
 If Modo = 1 Then CheckCadenaBusqueda chkSeries, BuscaChekc
 
End Sub

Private Sub chkSeries_GotFocus()
     ConseguirfocoChk Modo
End Sub

Private Sub chkSeries_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub






Private Sub ChkTrazabilidad_Click()
     If Modo = 1 Then CheckCadenaBusqueda ChkTrazabilidad, BuscaChekc
End Sub

Private Sub ChkTrazabilidad_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub ChkTrazabilidad_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim cad As String, Indicador As String
Dim bol As Boolean

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
'              If InsertarDesdeForm(Me) Then
'                InsetarArticulosPorAlmacen
'                InsertarPreciosPorTarifa
                If InsertarArticulo Then
'                    MsgBox "Los precios del artículo por tarifa se han introducido correctamente.", vbInformation
                    PosicionarData
                End If
'              End If
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me, 1) Then
                    TerminaBloquear
                    'FICHA TECNICA
                    If Not vParamAplic.EsAVAB Then ModificarEnFichaTecnica
                    'comprobar si se ha modificado el precio ult. compra
                    'y preguntar si modificar PVP y Tarifas
                    If CCur(DBLet(Data1.Recordset!precioUC, "N")) <> ImporteFormateado(Text1(15).Text) Then
                        ActualizarPreciosVenta
                    ElseIf CCur(DBLet(Data1.Recordset!preciove, "N")) <> ImporteFormateado(Text1(17).Text) Then
                        'Comprobar si se ha modificado el precio de venta PVP y preguntar
                        'si se quieren actualizar las tarifas de precios
                        ActualizarPreciosPorTarifa
                    ElseIf CCur(DBLet(Data1.Recordset!margecom, "N")) <> ImporteFormateado(Text1(25).Text) Then
                        'comprobar si se ha modificado el margen comercial
                        'y preguntar si modificar PVP y Tarifas
                         ActualizarPreciosVenta
                    End If
                    
'                    DesBloqueaRegistroForm Text1(0)
                    PosicionarData
                End If
            End If
                
         Case 5 'InsertarModificar linea  '----------------
         
            'Actualizar el registro en la tabla de lineas 'salmac' (Artículos x Almacen)
            If InsertarModificarLinea Then
'                DesBloqueaRegistroForm Text1(0)
      
                NumRegElim = Data4.Recordset.AbsolutePosition
                TerminaBloquear
                LLamaLineas2 0, 0, 4
                DataGrid3.AllowAddNew = False
                CargaGrid Me.DataGrid3, Me.Data4, True
                SituarDataPosicion Data4, NumRegElim, Indicador
                
                lblIndicador.Caption = Indicador
                PonerModoFrame 0
                PonerSumaStocks
                
               
                
            End If
            
          Case 6, 7 '6: InsertarModificar Conjuntos
                    '7: InsertarModificar Instalaciones
             If Modo = 6 Then bol = InsertarModificarConjunto
             If Modo = 7 Then bol = InsertarModificarInstalacion
             If bol Then
                TerminaBloquear
                If Modo = 6 Then 'Conjunto
                  txtAux(0).visible = False
                  txtAux(1).visible = False
                  txtAux2.visible = False
                  cmdAux.visible = False
                  CargaGrid Me.DataGrid1, Me.Data2, True
                Else 'Instalacion
                    txtAux(2).visible = False
                    CargaGrid Me.DataGrid2, Me.Data3, True
                End If
                If ModificaLineas = 2 Then 'Modificar
                    DesBloqueaRegistroForm Text1(0)
                    If Modo = 6 Then
                        Data2.Recordset.Find (Data2.Recordset.Fields(1).Name & " =" & CInt(Me.cmdAceptar.Tag))
                    Else
                        Data3.Recordset.Find (Data3.Recordset.Fields(1).Name & " =" & CInt(Me.cmdAceptar.Tag))
                    End If
                    PonerBotonCabecera True
'                    Me.lblIndicador.Caption = ""
                    PonerFocoBtn Me.cmdAceptar
                    ModificaLineas = 0
                ElseIf ModificaLineas = 1 Then 'Insertar
                    BotonAnyadirConjunto2
                End If
            End If
    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub




Private Sub cmdActualizarImportes1_Click(Index As Integer)
Dim frmAr As frmAlmArticulos

    If Modo <> 6 Then Exit Sub
    
    If ModificaLineas <> 0 Then
        MsgBox "Esta cambiando datos", vbExclamation
        Exit Sub
    End If
    
    If Index = 0 Then
        If txtConjunto(1).Text = "" Or txtConjunto(1).Text = "" Then
            MsgBox "Falta importes calculados", vbExclamation
            Exit Sub
        End If
        BuscaChekc = "¿Desea cambiar los importes PVP y UPC del árticulo principal?"
        If MsgBox(BuscaChekc, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    
    If Index = 0 Then
        'ACtualizar importes
    
        'Haremos lo siguiente
        If BLOQUEADesdeFormulario(Me) Then
            'Fijaremos los nuevos importes
             
             If ModificarImportesDesdeConjuntos Then
                    TerminaBloquear
                    Text1(15).Text = Me.txtConjunto(1).Text
                    Text1(17).Text = Me.txtConjunto(4).Text
                    'comprobar si se ha modificado el precio ult. compra
                    'y preguntar si modificar PVP y Tarifas
                    If CCur(DBLet(Data1.Recordset!precioUC, "N")) <> ImporteFormateado(Text1(15).Text) Then ActualizarPreciosVenta
                    'Comprobar si se ha modificado el precio de venta PVP y preguntar
                    'si se quieren actualizar las tarifas de precios
                    If CCur(DBLet(Data1.Recordset!preciove, "N")) <> ImporteFormateado(Text1(17).Text) Then ActualizarPreciosPorTarifa
                    

                    PosicionarData
            End If
        End If
    Else
        'VER ARTICULO LINEA
        Set frmAr = New frmAlmArticulos
        frmAr.DeConsulta = True
        frmAr.DatosADevolverBusqueda2 = "::" & DevNombreSQL(Data2.Recordset!codarti1)
        frmAr.Show vbModal
        Set frmAr = Nothing
        
        'Por si acaso ha cambiado
        'recargo el grid
        '--------------------------------------------------------------------------------------
        NumRegElim = Data2.Recordset.AbsolutePosition - 1
        
        CancelaADODC Me.Data2
        CargaGrid Me.DataGrid1, Me.Data2, True
        CancelaADODC Me.Data2
        ponerDatosConjuntos
        If NumRegElim > 0 Then Data2.Recordset.Move NumRegElim, 1
        
    End If
    BuscaChekc = ""
End Sub

Private Function ModificarImportesDesdeConjuntos() As Boolean
    On Error GoTo EM
    ModificarImportesDesdeConjuntos = False
    BuscaChekc = "UPDATE sartic set precioUC = " & TransformaComasPuntos(CStr(ImporteFormateado(Me.txtConjunto(1).Text)))
    BuscaChekc = BuscaChekc & " , preciove =" & TransformaComasPuntos(CStr(ImporteFormateado(Me.txtConjunto(4).Text)))
    BuscaChekc = BuscaChekc & " WHERE codartic = '" & DevNombreSQL(Data1.Recordset!codartic) & "'"
    Conn.Execute BuscaChekc
    ModificarImportesDesdeConjuntos = True
    Exit Function
EM:
    MuestraError Err.Number, Err.Description
End Function


Private Sub cmdAlma_Click()
    imgCuentas_Click 6
End Sub

Private Sub cmdAux_Click()
    MandaBusquedaPrevia " conjunto=0 "
    PonerFoco txtAux(1)
End Sub

Private Sub cmdCancelar_Click()
On Error Resume Next
    Select Case Modo
        Case 1 'Busqueda
            LimpiarCampos
            PonerModo 0
            PonerFoco Text1(0)
            
        Case 3 'Insertar
            If ModoAnterior = 0 Then
                LimpiarCampos
                PonerModo 0
            Else
                PonerModo 2
                PonerCampos
            End If
                
        Case 4  'Modificar
            TerminaBloquear
'            DesBloqueaRegistroForm Text1(0)
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
        
        
        'QUITAR### TODO ESTE COMENTARIO ELIMINADO LAS LINEAS
'        Case 5 'Lineas Detalle
''            DesBloqueoManual NombreTabla
'            TerminaBloquear
'            PonerModoFrame 0
'            PonerCamposAlmacenes2
'            ModificaLineas = 0
'            PonerFoco Text3(1)
        
        Case 5, 6, 7 'Lineas Conjuntos, Lineas Instalaciones
            ModificaLineas = 0
'            DesBloqueoManual NombreTabla
            TerminaBloquear
            Select Case Modo
            Case 6
                txtAux(0).visible = False
                txtAux(1).visible = False
                txtAux2.visible = False
                cmdAux.visible = False
                DataGrid1.AllowAddNew = False
                If Not (ModificaLineas = 2) Then 'Modificar
                    If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
                End If
                DataGrid1.Enabled = True
            Case 7
                txtAux(2).visible = False
                DataGrid2.AllowAddNew = False
                If Not (ModificaLineas = 2) Then 'Modificar
                    If Not Data3.Recordset.EOF Then Data3.Recordset.MoveFirst
                End If
                DataGrid2.Enabled = True
            Case 5
                
                DataGrid3.AllowAddNew = False
                DataGrid2.Enabled = False
                LLamaLineas2 0, 0, 4
                NumRegElim = Data4.Recordset.AbsolutePosition
                CargaGrid DataGrid3, Data4, True
                SituarDataPosicion Data4, NumRegElim, Me.lblIndicador.Caption
                If Not Data4.Recordset.EOF Then PonerCamposAlmacenes2
                 
            End Select
            PonerBotonCabecera True
            PonerFocoBtn Me.cmdRegresar
    End Select
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub BotonAnyadir()
    LimpiarCampos 'Vacía los TextBox
    ModoAnterior = Modo 'Para el botón Cancelar en Modo Insertar
    PonerModo 3
    PonerFramesFichaTecnica2 -1, False
    Me.SSTab1.Tab = 0
    
    'Poner valores por defecto
    Me.chkctrstock.Value = 1 'por defecto hay control de stock
    Me.ChkTrazabilidad.Value = 1
    Me.Text1(10).Text = Format(Now, "dd/mm/yyyy") 'fecha alta
    Me.cboArticuloVarios.ListIndex = 0
    Me.cboStatus.ListIndex = 0
    cboTipoArt.ListIndex = -1
    Me.Text1(11).Text = "0"
    Me.Text1(12).Text = "1"
    Me.Text1(30).Text = "1,0000"
    PonerFoco Text1(0)
End Sub


Private Sub BotonAnyadirLinea()
Dim vWhere As String
    
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
    Me.SSTab1.Tab = 4
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModoFrame 3    '3: Insertar
    ModificaLineas = 1 'Insertar

    'Obtenemos la siguiente numero de Artículo
    vWhere = "codartic=" & DBSet(Text1(0).Text, "T")
    Text3(0).Text = SugerirCodigoSiguienteStr("salmac", "codalmac", vWhere)
    lblIndicador.Caption = "INSERTAR ALMACEN"
    PonerFoco Text3(0)
End Sub





Private Sub BotonAnyadirConjunto2()
Dim NumF As String
Dim vWhere As String
Dim anc As Single
    
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
    ModificaLineas = 1
    PonerBotonCabecera False
    
    vWhere = "codartic=" & DBSet(Text1(0).Text, "T")
'    ancIni = 200
    Select Case Modo
    Case 6
        NumF = SugerirCodigoSiguienteStr("sarti1", "numlinea", vWhere)
        Me.SSTab1.Tab = 2
        lblIndicador.Caption = "INSERTAR CONJUNTO"
    Case 7
        NumF = SugerirCodigoSiguienteStr("sarti2", "numlinea", vWhere)
        Me.SSTab1.Tab = 3
        lblIndicador.Caption = "INSERTAR INSTALACIÓN"
    Case 5
        'STOCK
        Me.SSTab1.Tab = 4
        lblIndicador.Caption = "INSERTAR STOCK"
        NumF = 1
    End Select
    cmdAceptar.Tag = NumF
    
    Select Case Modo
    'If Modo = 6 Then 'Conjuntos
    Case 6
        txtAux(0).Text = ""
        txtAux2.Text = ""
        txtAux(1).Text = ""
        'Situamos el grid al final
        AnyadirLinea DataGrid1, Data2

        anc = ObtenerAlto(DataGrid1, 20)
        LLamaLineas2 anc, 1, 2
        
        BloquearTxt txtAux(0), False
        Me.cmdAux.Enabled = True
        PonerFoco txtAux(0)
        
    Case 7
        'INSTALACIONES
        Me.txtAux(2).Text = ""
        AnyadirLinea DataGrid2, Data3
        anc = ObtenerAlto(DataGrid2, 20)
        LLamaLineas2 anc, 1, 3
        PonerFoco txtAux(2)
    Case 5
        'STOCK
        PonerDatosForaGrid True
        PonerModoFrame 3
        AnyadirLinea DataGrid3, Data4
        anc = ObtenerAlto(DataGrid3, 20)
        LLamaLineas2 anc, 1, 4
        PonerFoco Text3(0)
        BloquearTxt Text3(0), False
    End Select
End Sub


Private Sub BotonBuscar()
'Buscar
    LimpiarCampos
    If Modo <> 1 Then 'Modo 1: Busqueda
        BuscaChekc = ""
        SSTab1.Tab = 0
        PonerModo 1
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Text1(0).BackColor = vbYellow
        'Si es de buscqueda , buscamos solo activos
        If DeConsulta Then Me.cboStatus.ListIndex = 0
    Else
        If DeConsulta Then
            If cboStatus.ListIndex < 0 Then cboStatus.ListIndex = 0
        End If
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
Dim c As String
  
    c = ""
    If DeConsulta Then c = "codstatu = 0"
    If Me.chkProdVenta.Value = 1 Then
        If c <> "" Then c = c & " AND "
        c = c & " conjunto=1 "
    End If
    LimpiarCampos
    If chkVistaPrevia.Value = 1 Then

    
        MandaBusquedaPrevia c
    Else
    
        If c <> "" Then c = " WHERE " & c
        c = "Select * from " & NombreTabla & c
        
        CadenaConsulta = c & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
    Select Case Modo
        Case 5 'Modo Mantenimiento de Almacenes (Lineas)
            If Data4.Recordset.EOF Then Exit Sub
            DesplazamientoData Data4, Index
            PonerCamposAlmacenes2
        Case Else 'Datos de Cabecera
            If Data1.Recordset.EOF Then Exit Sub
            DesplazamientoData Data1, Index
            PonerCampos
            PonerModoOpcionesMenu (Modo) 'Poner opciones de menu según modo
            PonerOpcionesMenu   'Activar opciones de menu según nivel
                                'de permisos del usuario
    End Select
End Sub


Private Sub BotonModificar()
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    PonerFoco Text1(1)
End Sub



Private Sub BotonModificarConjunto(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim anc As Single
Dim i As Integer

    If vData.Recordset.EOF Then Exit Sub
    If vData.Recordset.RecordCount < 1 Then Exit Sub
   
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub

    Screen.MousePointer = vbHourglass
    PonerBotonCabecera False
         
    If vDataGrid.Bookmark < vDataGrid.FirstRow Or vDataGrid.Bookmark > (vDataGrid.FirstRow + vDataGrid.VisibleRows - 1) Then
        i = vDataGrid.Bookmark - vDataGrid.FirstRow
        vDataGrid.Scroll 0, i
        vDataGrid.Refresh
    End If
    PonerFocoBtn Me.cmdAceptar
    vDataGrid.Enabled = False

    anc = ObtenerAlto(vDataGrid, 20)
    
    If Modo = 5 Then
        cmdAceptar.Tag = vData.Recordset!codAlmac
    Else
        cmdAceptar.Tag = vData.Recordset!numlinea
    End If
    Select Case Modo
    Case 6
    ' If Modo = 6 Then 'Componentes de Conjunto
        Me.lblIndicador.Caption = "MODIFICAR CONJUNTO"
        Me.SSTab1.Tab = 2
         'Llamamos al form
        txtAux(0).Text = DataGrid1.Columns(2).Text
        BloquearTxt txtAux(0), True
        Me.txtAux2.Text = DataGrid1.Columns(3).Text
        txtAux(1).Text = DataGrid1.Columns(4).Text
        LLamaLineas2 anc, 2, 2
        PonerFoco txtAux(1)
        If ModificaLineas = 2 Then cmdAux.Enabled = False
    'Poner el foco
    'ElseIf Modo = 7 Then
    Case 7
        Me.lblIndicador.Caption = "MODIFICAR INSTALACIÓN"
        Me.SSTab1.Tab = 3
        txtAux(2).Text = DataGrid2.Columns(2).Text
        LLamaLineas2 anc, 2, 3
        PonerFoco txtAux(2)
        
    Case 5
        
        
        PonerModoFrame 4 'ModoFrame=4 -> Modificar
        Me.lblIndicador.Caption = "MODIFICAR ALMACEN"
        LLamaLineas2 anc, 2, 4
        BloquearTxt Text3(0), True
        Text3(0).Text = Data4.Recordset!codAlmac
        Text3(2).Text = Data4.Recordset!CanStock
        Text2(8).Text = Data4.Recordset!nomalmac
        PonerFoco Text3(1)
        
    End Select
    Screen.MousePointer = vbDefault
End Sub


Private Sub BotonEliminar()
Dim SQL As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    'No esta bloqueado
    If Val(Data1.Recordset!codstatu) = 1 Then
        MsgBox "Articulo bloqueado", vbExclamation
        Exit Sub
    End If
    
    
    'Tiene stock
    If ImporteFormateado(txtSumaStock.Text) <> 0 Then
        MsgBox "El articulo tiene stock", vbExclamation
        Exit Sub
    End If
    

    
    BuscaChekc = lblIndicador.Caption
    SQL = SePuedeEliminarArticulo(CStr(Data1.Recordset!codartic), lblIndicador)
    lblIndicador.Caption = BuscaChekc
    BuscaChekc = ""
    If SQL <> "" Then
        SQL = "No se puede eliminar el articulo: " & Data1.Recordset!codartic & vbCrLf & vbCrLf & SQL
        MsgBox SQL, vbExclamation
        Exit Sub
    End If
    SQL = "Cabecera de Artículos." & vbCrLf
    SQL = SQL & "---------------------------        " & vbCrLf & vbCrLf
    SQL = SQL & "Va a eliminar el Artículo:"
    SQL = SQL & vbCrLf & "Cod. Artic. :   " & Data1.Recordset.Fields(0)
    SQL = SQL & vbCrLf & "Descripción :   " & Data1.Recordset.Fields(1)
    SQL = SQL & vbCrLf & vbCrLf & " ¿Desea Eliminarlo? "
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        TerminaBloquear
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        If Not Eliminar Then
            Screen.MousePointer = vbDefault
            Exit Sub
        ElseIf SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerModo 2
            PonerCampos
        Else
            LimpiarCampos
            PonerModo 0
        End If
    End If
    Screen.MousePointer = vbDefault
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Articulo", Err.Description
End Sub


Private Sub BotonEliminarLinea()
'Eliminar una linea De ArticulosxAlmacen
Dim cad As String

     On Error GoTo Error2

    If Data4.Recordset.EOF Then Exit Sub
    If Data4.Recordset.RecordCount < 1 Then Exit Sub
    
    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar
    ModificaLineas = 3 'Eliminar
    
    '### a mano
    cad = "Seguro que desea eliminar de la BD el registro:"
    cad = cad & vbCrLf & "Cod. Artículo: " & Data1.Recordset.Fields(0)
    cad = cad & vbCrLf & "Cod. Almacen: " & Data4.Recordset.Fields(1)

    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
       
        Screen.MousePointer = vbHourglass
        NumRegElim = Data4.Recordset.AbsolutePosition
        
        cad = "DELETE FROM salmac where codartic = '" & DevNombreSQL(Data1.Recordset.Fields(0)) & "' AND codalmac = " & Data4.Recordset!codAlmac
        Conn.Execute cad
        
        CargaGrid Me.DataGrid3, Me.Data4, True
        If Data4.Recordset.EOF Then
            'Solo habia un registro
            LimpiarCamposAlmacenes
            PonerModoFrame 0
        Else
            SituarDataPosicion Me.Data4, NumRegElim, cad
            PonerCamposAlmacenes2
        End If
        ModificaLineas = 0
    End If
    Screen.MousePointer = vbDefault
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        Data4.Recordset.CancelUpdate
        MuestraError Err.Number, "Eliminar Linea de Articulo", Err.Description
    End If
End Sub


Private Sub BotonEliminarConjunto()
Dim SQL As String
    On Error GoTo Error2
    
    'Ciertas comprobaciones
    If Data2.Recordset.EOF Then Exit Sub
    
    SQL = "Seguro que desea eliminar el Componente de Conjunto:"
    SQL = SQL & vbCrLf & "Código: " & Data2.Recordset!codarti1
    SQL = SQL & vbCrLf & "Descripción: " & Data2.Recordset.Fields(3)
    
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        SQL = "Delete from sarti1 where codartic=" & DBSet(Data2.Recordset!codartic, "T")
        SQL = SQL & " and numlinea=" & Data2.Recordset!numlinea
        SQL = SQL & " and codarti1=" & DBSet(Data2.Recordset!codarti1, "T")
        Conn.Execute SQL
        
        'If vEmpresa.codemprevParamAplic.EsAVAB <> EmpresaAVAB Then ActualizaComponentesAVAB SQL
        If Not vParamAplic.EsAVAB Then ActualizaComponentesAVAB SQL
        
        CancelaADODC Me.Data2
        CargaGrid Me.DataGrid1, Me.Data2, True
        CancelaADODC Me.Data2
        
        
        
        
    End If
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Componente de Conjunto", Err.Description
End Sub


Private Sub BotonEliminarInstalacion()
Dim SQL As String
    On Error GoTo Error2

    'Ciertas comprobaciones
    If Data3.Recordset.EOF Then Exit Sub
    
    SQL = "Seguro que desea eliminar el control de instalación:"
    SQL = SQL & vbCrLf & "Linea: " & Data3.Recordset!numlinea
    SQL = SQL & vbCrLf & "Descripción: " & Data3.Recordset!licontro
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        SQL = "Delete from sarti2 where codartic=" & DBSet(Data3.Recordset!codartic, "T")
        SQL = SQL & " and numlinea=" & Data3.Recordset!numlinea
        Conn.Execute SQL
        CancelaADODC Me.Data3
        CargaGrid Me.DataGrid2, Me.Data3, True
        CancelaADODC Me.Data3
    End If
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Control de Instalaciones", Err.Description
End Sub


Private Sub BotonArticulosxAlmac()

    If vUsu.Nivel > 0 Then Exit Sub

    If MsgBox("NO deberia modificar nada aqui.  ¿Continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

    Screen.MousePointer = vbHourglass
    On Error GoTo ErrorArticAlmac
    
    
    
    
    
    
    
    
    
    Screen.MousePointer = vbHourglass
    'RESTAURO LOS tag's
    AccionesSobreTagText3_ False, False

    Me.SSTab1.Tab = 4
    PonerModo (5)
    PonerBotonCabecera True
    PonerFocoBtn Me.cmdRegresar
    Screen.MousePointer = vbDefault
    
    
  
    
    
    
    
    
    
'ANTEs ------------------------------------------------------
'
'    'Crear las lineas de Articulos x Almacen para el artículo
'    Me.SSTab1.Tab = 0
'
'    'ASignamos un SQL al DATA4
''    Me.Data4.ConnectionString = Conn
''    Data4.RecordSource = "Select * from salmac where codartic = '" & Text1(0).Text & "';"
''    Data4.Refresh
'
'    If Data4.Recordset.RecordCount <= 0 Then
'        MsgBox "No hay ningún registro en la tabla salmac", vbInformation
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    Else
'        'Poner el modo en el formulario
'        PonerModo (5) 'Modo 5: Modificar lineas
'        PonerModoFrame 0 'TextBox Bloqueados inicialmente
'
'        'Data4.Recordset.MoveFirst
'        'PonerCamposAlmacenes
'        'PonerFocoBtn Me.cmdRegresar
'        Screen.MousePointer = vbDefault
'    End If
    Exit Sub
ErrorArticAlmac:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub

Private Sub BotonConjuntos()
    On Error GoTo ErrorConjuntos
    
    Screen.MousePointer = vbHourglass
    Me.SSTab1.Tab = 2
    
    PonerModo (6)
    
    PonerBotonCabecera True
    HayQueRecalcularPesos = False
    DataGrid1.Enabled = True
    Me.DataGrid1.SetFocus
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrorConjuntos:
    MuestraError Err.Number, "Conjuntos"
    Screen.MousePointer = vbDefault
End Sub


Private Sub BotonInstalaciones()
    On Error GoTo ErrorInstala

    Screen.MousePointer = vbHourglass
    Me.SSTab1.Tab = 3
    PonerModo (7)
    PonerBotonCabecera True
    PonerFocoBtn Me.cmdRegresar
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrorInstala:
    MuestraError Err.Number, "Instalaciones", Err.Description
    Screen.MousePointer = vbDefault
End Sub


Private Sub cmdGenerar_Click()
    Dim Aux As String
    
    If vParamAplic.ContabilizacionMoixent Then
        'MOIXENT, No lleva la categoria
        Aux = Text2(2) & " " & Text2(4) & " " & Text2(3)
    Else
        Aux = Text2(2) & " " & Text2(1) & " " & Text2(4) & " " & Text2(3)
    End If
    Text1(1).Text = Replace(Left(Aux, 40), "*", "")
    Text1(0).Text = SugerirCodAutomatico(Text1(4), Text1(3), Text1(6), Text1(5))
End Sub

Private Sub cmdImg_Click()
    If Data1.Recordset.EOF Then Exit Sub
    If Val(DBLet(Data1.Recordset!tipartic, "N")) < 1 Then
        'Seguro que no tiene imganes
        MsgBox "Tipo articulo incorrecto(=0)", vbExclamation
        Exit Sub
    End If
        
    If Val(DBLet(Data1.Recordset!tipartic, "N")) = 1 Then
        'Seguro que no tiene imganes
        MsgBox "Tipo articulo: Producto venta.", vbExclamation
        Exit Sub
    End If
        
    frmFichaTecIMG_.EsArticulo = True
    frmFichaTecIMG_.vDatos = Text1(0).Text & "|" & Text1(1).Text & "|"
    frmFichaTecIMG_.Show vbModal
End Sub

Private Sub cmdImprCompo_Click()
    If Data1.Recordset.EOF Then Exit Sub
    If Data2.Recordset.EOF Then Exit Sub
     
 
     With frmImprimir
        .FormulaSeleccion = "{sartic.codartic}=""" & Text1(0).Text & """"
        .OtrosParametros = "|pEmpresa=""" & vEmpresa.nomempre & """|"
        .NumeroParametros = 1
        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = 96
        .Show vbModal
    End With
  
End Sub

Private Sub cmdIMpriFT_Click()
    If vParamAplic.EsAVAB Then Exit Sub
    If Me.chkConjunto.Value = 0 Then
        MsgBox "La ficha tecnica es de producto venta", vbExclamation
        Exit Sub
    End If
    frmFichaTecnicaImp.vCodartic = Text1(0).Text & "|" & Text1(1).Text & "|"
    frmFichaTecnicaImp.Show vbModal
End Sub

Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim cad As String
Dim C1 As Integer
Dim C2 As Currency
Dim PesoNetoVacio As Boolean

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Or Modo = 6 Or Modo = 7 Then

        If Modo = 6 And Not vParamAplic.EsAVAB Then
        
            'No ha puesto el peso del aceite
            PesoNetoVacio = False
            If Text5(0).Text = "" Then
                PesoNetoVacio = True
                'Si tiene conjuntos
                If DBLet(Data1.Recordset!Conjunto, "N") > 0 And DBLet(Data1.Recordset!LitrosUnidad, "N") > 0 Then
                    cad = "FactorConversion < 1 And sarti1.codarti1 = sartic.codartic And sarti1.codartic"
                    cad = DevuelveDesdeBD(conAri, "factorconversion*cantidad", "sarti1,sartic", cad, Data1.Recordset!codartic, "T")
                    If cad <> "" Then
                        C2 = CCur(cad)
                        Text5(0).Text = Format(C2, FormatoPrecio)
                    End If
                End If
            End If
            'Cajas Palet
            C1 = Val(ComprobarCero(Text5(4).Text))
            C2 = Val(ComprobarCero(Text5(5).Text))
            C1 = C1 * C2
            'El peso del aceite
            cad = ComprobarCero(Text5(0).Text)
            C2 = CCur(cad)
            RecalcularPesoArticulo CStr(Data1.Recordset!codartic), DBLet(Data1.Recordset!UniCajas, "N"), C1, C2, PesoNetoVacio
            Espera 0.1
            
            PonerDatosFichaTecnica2 DBLet(Data1.Recordset!tipartic, "N")
            
        End If
        'modo 5: Lineas Articulos x Almacen
        'modo 6: Lineas Conjuntos
        'modo 7: Lineas Instalaciones
        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        If DataGrid1.Row >= 0 Then
            DeseleccionaGrid Me.DataGrid1
            DataGrid1.Bookmark = 1
        End If
        If DataGrid2.Row >= 0 Then
            DeseleccionaGrid Me.DataGrid2
            DataGrid2.Bookmark = 1
        End If
        PonerModo 2

    Else 'Se llamo desde un botón de Prismático
        If Data1.Recordset.EOF Then
            MsgBox "Ningún registro devuelto.", vbExclamation
            Exit Sub
        End If
        If DeConsulta Then
            If cboStatus.ListIndex > 0 Then
                MsgBox "Articulo " & cboStatus.Text, vbExclamation
                Exit Sub
            End If
        End If
            
        cad = Data1.Recordset.Fields(0) & "|"
        cad = cad & Data1.Recordset.Fields(1) & "|"
        cad = cad & Data1.Recordset.Fields(8).Value & "|"
        cad = cad & Text2(4).Text & "|"
        RaiseEvent DatoSeleccionado(cad)
        Unload Me
    End If
End Sub


Private Sub Data4_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If Modo = 5 And ModificaLineas > 0 Then Exit Sub
    If Not Data4.Recordset.EOF Then
        If Not PrimeraVez Then PonerCamposAlmacenes2
    End If
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        If ModificaLineas = 0 Then
            PonerFocoBtn Me.cmdRegresar
        Else
            SendKeys "{tab}"
        End If
    End If
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If Modo = 1 Then PonerFoco Text1(1)
    If PriVezForm Then
        PriVezForm = False
        'He abierto el form queriendo cargar un articulo
        If Mid(DatosADevolverBusqueda2, 1, 2) = "::" Then
            DatosADevolverBusqueda2 = Mid(DatosADevolverBusqueda2, 3)
            CadenaConsulta = "Select * from " & NombreTabla & " where codartic='" & DatosADevolverBusqueda2 & "'"
            PonerCadenaBusqueda
         End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Load()
    

  
    
    PriVezForm = True
    'Icono del formulario
    Me.Icon = frmppal.Icon
 
    ' ICONITOS DE LA BARRA
    btnAnyadir = 6
    btnPrimero = 17 'Posicion del Boton Primero en la toolbar (+ 3 siguientes)
    With Me.Toolbar1
        .ImageList = frmppal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        .Buttons(6).Image = 3   'Insertar Nuevo
        .Buttons(7).Image = 4   'Modificar
        .Buttons(8).Image = 5   'Borrar
        .Buttons(10).Image = 10 'Stocks Almacenes
        .Buttons(11).Image = 11 'Conjuntos
        .Buttons(12).Image = 36 'Instalaciones
        .Buttons(14).Image = 16  'Imprimir
        .Buttons(15).Image = 15  'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With

    
    'Como en un futuro se parametrizaran el numero de decimales...
    'NUEVO Oct 2010. NEcesit 4 decimales
    'Text1(29).Tag = "Litros x Ud|N|S|||sartic|LitrosUnidad|" & FormatoCantidad & "|N|"
    Text1(29).Tag = "Litros x Ud|N|S|||sartic|LitrosUnidad|" & FormatoPrecio & "|N|"
    
    Me.SSTab1.Tab = 0
    Me.SSTab1.TabVisible(2) = False
    Me.SSTab1.TabVisible(3) = False
    'Me.FrameDatosAlmacen.Left = 360
    'Me.FrameDatosAlmacen.Top = 2780
    'Me.FrameArtxAlmac2.visible = False
    
    LimpiarCampos   'Limpia los campos TextBox
    PrimeraVez = True
    
        
    'Si no tiene servicios nmo muesotr el frame
    FrameServicios.visible = vParamAplic.Servicios
    FrameLitrosUd.visible = vParamAplic.Descriptores
    

    
    
    
    Dim B As Boolean
    B = False
    If vParamAplic.EsAVAB Then
        If vUsu.Nivel >= 1 Then B = True
    End If
    If B Then
        FrameDatosAlmacen2.Left = 13000
    End If
    
    B = False
    If Not vParamAplic.EsAVAB Then
        'Morales solo
        If vParamAplic.QUE_EMPRESA < 1 Then B = True
    End If
    'SSTab1.TabVisible(6) = Not vParamAplic.EsAVAB
    SSTab1.TabVisible(6) = B
    
    'Si hay algun combo los cargamos
    CargarComboStatus
    CargarComboArticuloVarios
    CargarCombo_Tabla cboTipoArt, "stipfamia", "tipfamia", "desctipfamia"
          
    Me.chkConso.visible = vUsu.TrabajadorB
    
    
    'El tag de los stocks
    AccionesSobreTagText3_ True, True
    
    'Pone el Tag del primer botón de busqueda de cuentas a -1
    'Si tag =-1 abre busqueda en la tabla: sartic, BD: Ariges
    'Si tag>0 abre busqueda en la tabla asociada al indice.
    imgCuentas(0).Tag = "-1"
         
    '## A mano
    NombreTabla = "sartic"
    Ordenacion = " ORDER BY codartic"
  
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = Conn
    Data1.RecordSource = "Select * from " & NombreTabla & " where codartic='-1' "
    Data1.Refresh

    If DatosADevolverBusqueda2 = "" Then
        PonerModo 0
        PonerCamposLineas False
    Else
        If DatosADevolverBusqueda2 = "@1@" Then 'Poner Modo Busqueda
            BotonBuscar
        Else 'Poner Modo Insertar
            If Mid(DatosADevolverBusqueda2, 1, 2) = "::" Then
                'Abrimos el articulo poniendo un articulo especificado a continuacion
                
                'Lo haremos en el ACTIVATE
            Else
                PonerModo 3
                Text1(0).Text = DatosADevolverBusqueda2
            End If
        End If
    End If
    '-- Descriptores especiales y botón de composición (Rafa VRS 4.0.9)
    If vParamAplic.Descriptores Then
        'cmdGenerar.visible = True  estara en poner modo
        Label1(6) = "Cod. Categoria"
        Label1(9) = "Cod. Modelo"
        Label1(17) = "Cod. Formato"
        '-- Aqui cambiamos los tag para evitar lios.
        CambiaTagDescriptores Text1(3), "Cod. Categoria"
        CambiaTagDescriptores Text1(5), "Cod. Formato"
        CambiaTagDescriptores Text1(6), "Cod. Modelo"
    Else
        cmdGenerar.visible = False
    End If
    '--
  
    ImagenesNavegacion
    Me.chkProdVenta.Value = Abs(Me.ParaVenta)
    CargaColumnas 0
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox

    lblIndicador.Caption = ""
    CargaGrid Me.DataGrid3, Me.Data4, False  'Desenlazamos el GRID
    'Aqui va el especifico de cada form es
    Me.chkConjunto.Value = 0
    Me.chkSeries.Value = 0
    Me.chkctrstock.Value = 0
    Me.ChkTrazabilidad.Value = 0
    Me.cboArticuloVarios.ListIndex = -1
    Me.cboStatus.ListIndex = -1
    cboTipoArt.ListIndex = -1
End Sub


Private Sub LimpiarCamposAlmacenes()
Dim i As Byte
    Text3(0).BackColor = vbRed
    For i = 0 To Text3.Count - 1
        Text3(i).Text = ""
    Next i
    Text2(8).Text = ""
    Me.chkInventario.Value = 0
    lblIndicador.Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    ParaVenta = False  'por si acaso
End Sub


Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
'Almacenes Propios
    Text3(0).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text3(0)
    Text2(8).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim cadB As String
Dim Aux As String
Dim Indice As Integer
      
    If CadenaDevuelta <> "" Then
        If Val(imgCuentas(0).Tag) >= 0 Then
            'Se llama desde el botón de busqueda del campo Tipos de IVA
            'Recuperar solo el campo código y Descripción
            HaDevueltoDatos = True
            Screen.MousePointer = vbHourglass
    
            Indice = Val(Me.imgCuentas(0).Tag)
            Text1(Indice + 2).Text = RecuperaValor(CadenaDevuelta, 1)
            Text2(Indice).Text = RecuperaValor(CadenaDevuelta, 2)
        Else
            HaDevueltoDatos = True
            Screen.MousePointer = vbHourglass
            
            If Modo <> 6 Then
                'Recupera todo el registro de Artículos
                'Sabemos que campos son los que nos devuelve
                'Creamos una cadena consulta y ponemos los datos
                cadB = ""
                Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
                cadB = Aux
                CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
                PonerCadenaBusqueda
            Else
                'Llamamos desde el boton auxiliar de Conjuntos
                txtAux(0).Text = RecuperaValor(CadenaDevuelta, 1)
                txtAux2.Text = RecuperaValor(CadenaDevuelta, 2)
            End If
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmCat_DatoSeleccionado(CadenaSeleccion As String)
'Categoria del Articulo
    Text1(22).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(22).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmF_Selec(vFecha As Date) 'Calendario Fechas

    Select Case Val(imgFecha(0).Tag)
        Case 0
            Text1(10).Text = Format(vFecha, "dd/mm/yyyy")
        Case 1
            Text1(18).Text = Format(vFecha, "dd/mm/yyyy")
        Case 2
            Text3(7).Text = Format(vFecha, "dd/mm/yyyy")
            
        Case 3
            Text1(24).Text = Format(vFecha, "dd/mm/yyyy")
    End Select
End Sub


Private Sub frmFA_DatoSeleccionado(CadenaSeleccion As String)
'Familia de Articulo
    Text1(3).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(3)
    Text2(1).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmM_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Marcas
    Text1(4).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(4)
    Text2(2).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmP_DatoSeleccionado(CadenaSeleccion As String)
'Proveedores
    Text1(2).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(2)
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmTA_DatoSeleccionado(CadenaSeleccion As String)
'Tipo de Articulo
    Text1(6).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(4).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmTU_DatoSeleccionado(CadenaSeleccion As String)
'Tipo de Unidad
    Text1(5).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(5)
    Text2(3).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmUbic_DatoSeleccionado(CadenaSeleccion As String)
'Mto Ubicaciones de almacen
    Text3(1).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(6).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub imgCuentas_Click(Index As Integer)

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Codigo Proveedor
            Set frmP = New frmComProveedores
            frmP.DatosADevolverBusqueda = "0"
            frmP.Show vbModal
            Set frmP = Nothing
        Case 1  'Cod. Familia
            Set frmFA = New frmAlmFamiliaArticulo
            frmFA.DatosADevolverBusqueda = "0"
            frmFA.Show vbModal
            Set frmFA = Nothing
        Case 2  'Cod. Marca
            Set frmM = New frmAlmMarcas
            frmM.DatosADevolverBusqueda = "0"
            frmM.Show vbModal
            Set frmM = Nothing
        Case 3  'Cod. Tipo Unidad
            Set frmTU = New frmAlmTipoUnidad
            frmTU.DatosADevolverBusqueda = "0"
            frmTU.Show vbModal
            Set frmTU = Nothing
        Case 4  'Cod. Tipo Articulo
            Set frmTA = New frmAlmTipoArticulo
            frmTA.DatosADevolverBusqueda = "0"
            frmTA.Show vbModal
            Set frmTA = Nothing
            
        Case 5  'Tipos de IVA. Tabla de la BD Contabilidad
            imgCuentas(0).Tag = Index
            MandaBusquedaPrevia ""
            imgCuentas(0).Tag = -1
            
        Case 6 'Código de Almacen
            Set frmA = New frmAlmAlPropios
            frmA.DatosADevolverBusqueda = "0"
            frmA.Show vbModal
            Set frmA = Nothing
            
        Case 7 'cod. ubicaciones
            Set frmUbic = New frmAlmUbicaciones
            frmUbic.DatosADevolverBusqueda = "0"
            frmUbic.Show vbModal
            Set frmUbic = Nothing
            
        Case 8 'cod. categoria
            Set frmCat = New frmAlmCategorias
            frmCat.DatosADevolverBusqueda = "0"
            frmCat.Show vbModal
            Set frmCat = Nothing
    End Select
    
    If Index = 6 Then
        PonerFoco Text3(0)
    ElseIf Index = 7 Then
        PonerFoco Text3(1)
    ElseIf Index = 8 Then
        PonerFoco Text1(22)
    Else
        PonerFoco Text1(Index + 2)
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFecha_Click(Index As Integer)
Dim Indice As Byte

   If Modo = 2 Or Modo = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass
   imgFecha(0).Tag = Index
   
   Set frmF = New frmCal
   frmF.Fecha = Now
   
   Select Case Index
     Case 0, 1, 3
        If Index = 0 Then
            Indice = 10
        ElseIf Index = 1 Then
            Indice = 18
        Else
            Indice = 24
        End If
        PonerFormatoFecha Text1(Indice)
        If Text1(Indice).Text <> "" Then frmF.Fecha = CDate(Text1(Indice).Text)

     Case 2
        PonerFormatoFecha Text3(7)
         If Text3(7).Text <> "" Then frmF.Fecha = CDate(Text3(7).Text)
   End Select
   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
End Sub



Private Sub imgPreciosProv_Click()
    frmAlmArtiProv.articulo2 = Text1(0).Text
    frmAlmArtiProv.Caption = "Proveedores para " & Text1(1).Text
    frmAlmArtiProv.Top = Me.Top + 2620
    frmAlmArtiProv.Left = Me.Left + 3300
    frmAlmArtiProv.Show vbModal
End Sub

Private Sub lw1_DblClick()
Dim Seleccionado As Long
    If Modo <> 2 Then Exit Sub
    If lw1.ListItems.Count = 0 Then Exit Sub
    If lw1.SelectedItem Is Nothing Then Exit Sub


    If Me.DatosADevolverBusqueda2 <> "" Then
        'De momento NO dejo continuar
        MsgBox "Esta buscando un cliente. No puede ver los documentos.", vbExclamation
        Exit Sub
    End If
    
    
    Screen.MousePointer = vbHourglass
    
    'Llegados aqui
    Select Case CByte(RecuperaValor(lw1.Tag, 1))
    Case 0, 1, 2
        
    Case 3
        If lw1.SelectedItem.SmallIcon = 6 Then
            'PEDIDO CLIENTE
            
            
            frmFacEntPedidos.DatosADevolverBusqueda2 = lw1.SelectedItem.Text
            frmFacEntPedidos.EsHistorico = False
            frmFacEntPedidos.Show vbModal
            
        Else
            'PROVEEDOR
            frmComEntPedidos.MostrarDatos = lw1.SelectedItem.Text
            frmComEntPedidos.EsHistorico = False
            frmComEntPedidos.Show vbModal

        End If
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
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    Select Case Modo
        Case 5  'Eliminar lineas Artículos x Almacen
            BotonEliminarLinea
        Case 6 'Eliminar Líneas Conjuntos
            BotonEliminarConjunto
        Case 7 'Eliminar Lineas de Control de Instalacion
            BotonEliminarInstalacion
        Case Else   'Eliminar Artículo
            BotonEliminar
    End Select
End Sub


Private Sub mnModificar_Click()
Dim cad As String
Dim Aux As String
Dim i As Integer

    Select Case Modo
        Case 5  'Modificar lineas Artículos x Almacen
'                cad = Text1(0).Text
'                i = InStr(1, cad, """")
'                If i > 0 Then
'                    Aux = Mid(cad, 1, i)
'                    Aux = Aux & """"
'                    Aux = Aux & Mid(cad, i + 1, Len(cad))
'                End If
'                NombreSQL cad
'                If BloqueoManual(NombreTabla, "'" & cad & "|" & Text3(0).Text & "|'") Then BotonModificarLinea
                Aux = " codartic=" & DBSet(Text1(0).Text, "T")
                If BloqueaRegistro("sartic", Aux) Then BotonModificarConjunto Me.DataGrid3, Me.Data4
                
                
        Case 6 'Modificar Líneas Conjuntos
'                If BloqueoManual(NombreTabla, "|'" & Text1(0).Text & "'|" & txtAux(0).Text & "|") Then
'                    BotonModificarConjunto Me.DataGrid1, Me.Data2
'                End If
                Aux = " codartic=" & DBSet(Text1(0).Text, "T")
                If BloqueaRegistro("sartic", Aux) Then BotonModificarConjunto Me.DataGrid1, Me.Data2
                
                
        Case 7  'Modificar Linea de Control de Instalacion
'                If BloqueoManual(NombreTabla, "|'" & Text1(0).Text & "'|" & cmdAceptar.Tag & "|") Then
'                    BotonModificarConjunto Me.DataGrid2, Me.Data3
'                End If
                Aux = " codartic=" & DBSet(Text1(0).Text, "T")
                If BloqueaRegistro("sartic", Aux) Then BotonModificarConjunto Me.DataGrid2, Me.Data3
                
        Case Else   'Modificar Artículos
            If BLOQUEADesdeFormulario(Me) Then BotonModificar
'            If BloqueaRegistroForm(Me) Then BotonModificar
    End Select
End Sub


Private Sub mnMtoConjuntos_Click()
    BotonConjuntos
End Sub

Private Sub mnMtoInstalaciones_Click()
    BotonInstalaciones
End Sub

Private Sub mnMtoStocksAlm_Click()
    BotonArticulosxAlmac
End Sub

Private Sub mnNuevo_Click()
     Select Case Modo
        'Case 5 'Añadir lineas Artículos x Almacen
         '       BotonAnyadirLinea   'QUITAR EL PROCEDEIMIENTO
        Case 5, 6, 7 'Añadir Líneas Conjuntos
                  'Añadir Linea de Control de Instalacion
                BotonAnyadirConjunto2
        Case Else 'Añadir Artículos
                BotonAnyadir
    End Select
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    If Modo = 5 Then
        '------------------------------------------------------
        'Si esta insertando lineas es una cosa, si no es otra
        cmdCancelar_Click
    Else
        If (Modo = 6) Or (Modo = 7) Then 'Modo 5: Mto Lineas
                        'Modo 6: Conjuntos, Modo 7: Instalaciones
                        
            cmdRegresar_Click
        Else
            Unload Me
        End If
    End If
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub SSTab1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    If (Not Text1(Index).MultiLine) And (Text1(Index).ScrollBars) = 0 Then ConseguirFoco Text1(Index), Modo
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
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
Dim c As String
    'Si modo=1 busqueda y pierde el foco el control del nombre articulo
    'entonces pongo el foco en aceptar, ya que el 99 % de las veces
    'buscare por nomartic
    If Modo = 1 And Index = 1 Then PonerFocoObjeto cmdAceptar



    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
        
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    

    'Si queremos hacer algo ..
    Select Case Index
        Case 0 'Codigo Artículo
            'Comprobar si ya existe el cod de articulo en la tabla
            If Modo = 3 Then 'Insertar
                If ExisteCP(Text1(Index)) Then PonerFoco Text1(Index)
            End If

        Case 2 'Codigo de Proveedor
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index - 2).Text = PonerNombreDeCod(Text1(Index), conAri, "sprove", "nomprove")
            Else
                Text2(Index - 2).Text = ""
            End If
            
        Case 3 'Código de Familia
            If PonerFormatoEntero(Text1(Index)) Then
                c = "tipfamia"
                'Text2(Index - 2).Text = PonerNombreDeCod(Text1(Index), conAri, "sfamia", "nomfamia")
                Text2(Index - 2).Text = DevuelveDesdeBD(conAri, "nomfamia", "sfamia", "codfamia", Text1(3).Text, "N", c)
                If Text2(Index - 2).Text = "" Then Text1(Index).Text = ""
                If Modo = 3 And Not vParamAplic.EsAVAB Then
                    c = Val(c)
                    PonerCamposTipoArt CInt(c)
                    PonerFramesFichaTecnica2 CInt(c), chkConjunto.Value = 1
                End If
            Else
                Text2(Index - 2).Text = ""
            End If
            
        Case 4 'Código de Marca
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index - 2).Text = PonerNombreDeCod(Text1(Index), conAri, "smarca", "nommarca")
            Else
                Text2(Index - 2).Text = ""
            End If
            
        Case 5 'Código Tipo Unidad
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index - 2).Text = PonerNombreDeCod(Text1(Index), conAri, "sunida", "nomunida")
            Else
                Text2(Index - 2).Text = ""
            End If
            
        Case 6 'Codigo Tipo Artículo
            Text2(Index - 2).Text = PonerNombreDeCod(Text1(Index), conAri, "stipar", "nomtipar")
            If Text1(Index).Text <> "" And Text2(Index - 2).Text = "" Then PonerFoco Text1(Index)
            
        Case 7 'Tipo de IVA
            'conConta: BD Contabilidad
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index - 2).Text = PonerNombreDeCod(Text1(Index), conConta, "tiposiva", "nombriva")
            Else
                Text2(Index - 2).Text = ""
            End If
            
        Case 10, 18, 24 'Fecha alta, Fecha última compra, FECHA VIGENCIA
            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)

        Case 11, 12 'numericos
            PonerFormatoEntero Text1(Index)

        Case 13, 14, 15, 16, 17 'Precios
            'Formato tipo 2: Decimal(10,4)
            If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 2
        
        Case 21 'Texto Control de instalación
            If (Modo <> 0) Then PonerFocoBtn Me.cmdAceptar
            
        Case 22 'categoria
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "scateg", "descateg")
            If Text2(Index).Text = "" And Text1(Index) <> "" Then PonerFoco Text1(Index)
            
        Case 25 'Margen comercial
            'Formato 7: Decimal(5,2)
            PonerFormatoDecimal Text1(Index), 7
        Case 26, 29, 30
             'Precio anual mantenimiento.  Lo que ponga en su tag
             ' Listros x Unidad
             If Not PonerFormatoDecimal(Text1(Index), 8) Then
                If Index = 30 Then Text1(Index).Text = ""
             End If
    End Select
End Sub


Private Sub HacerBusqueda()
Dim cadB As String
    
    cadB = ObtenerBusqueda(Me, False, BuscaChekc)
    If chkProdVenta.Value Then
        If cadB <> "" Then cadB = cadB & " AND "
        cadB = cadB & " conjunto=1 "
    End If
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
Dim cad As String
Dim Tabla As String
Dim Titulo As String
Dim Conexion As Byte

    'Llamamos a al form
    '##A mano
    cad = ""
    Select Case Val(Me.imgCuentas(0).Tag)
        Case 5  'Tipo de IVA
            'Se llama a Busqueda desde el campo Tipos IVA
            '#A MANO: Porque busca en la tabla tiposiva
            'de la base de datos de Contabilidad
            cad = cad & "Código|tiposiva|codigiva|N||20·Denominacion|tiposiva|nombriva|T||70·"
            Tabla = "tiposiva"
            Titulo = "Tipos de IVA"
            Conexion = conConta    'Conexión a BD: Conta
        Case Else   'Registro de la tabla de cabeceras: sartic
            cad = cad & ParaGrid(Text1(0), 30, "Código")
            cad = cad & ParaGrid(Text1(1), 70, "Denominación")
            Tabla = "sartic"
            Titulo = "Artículos"
            Conexion = conAri    'Conexión a BD: Ariges
    End Select
           
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
        frmB.vselElem = 1
        frmB.vConexionGrid = Conexion
'        frmB.vBuscaPrevia = VPrevia
        frmB.vCargaFrame = False
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
        If HaDevueltoDatos Then
            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda2 <> "" Then _
                cmdRegresar_Click
        Else   'de ha devuelto datos, es decir NO ha devuelto datos
            PonerFoco Text1(kCampo)
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCadenaBusqueda()

    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
    
    
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        If Modo = 1 Then 'Busqueda
             MsgBox "No hay ningún registro en la tabla " & NombreTabla & " para ese criterio de Búsqueda.", vbInformation
             PonerFoco Text1(0)
        Else
            MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        Data1.Recordset.MoveFirst
        PonerModo 2
        PonerCampos
        PonerCamposAlmacenes2
        'David 28 Nov 2008
        ' Si es conjunto mostrare sus solapa. Para morales
        MostrarSolapa = False
        If vUsu.Nivel = 0 Then
            MostrarSolapa = True
        Else
            If Not vParamAplic.EsAVAB Then MostrarSolapa = True
        End If
        If MostrarSolapa Then
            If Me.chkConjunto.Value = 1 Or Me.chkConjunto.Value = 2 Then PonerModoOpcionesMenu 2
        End If
        
        If DatosADevolverBusqueda2 <> "" Then PonerFocoBtn Me.cmdRegresar
    End If

    Screen.MousePointer = vbDefault
    Exit Sub
    
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
Dim Impor As Currency
Dim TipoFamilia As Integer

    If Data1.Recordset.EOF Then Exit Sub
    
    lblIndicador.Caption = "Datos articulo"
    lblIndicador.Refresh
    PonerCamposForma Me, Data1
    

    
    
    Text2(0).Text = PonerNombreDeCod(Text1(2), conAri, "sprove", "nomprove")
    
    'AQUI LEERE LOS DATOS DE TIPO FAMILIA (envase, apon.....) TipoFamilia
    'Text2(1).Text = PonerNombreDeCod(Text1(3), conAri, "sfamia", "nomfamia")
    mPorIva = "tipfamia"
    Text2(1).Text = DevuelveDesdeBD(conAri, "nomfamia", "sfamia", "codfamia", Text1(3).Text, "N", mPorIva)
    TipoFamilia = Val(mPorIva)
    
    Text2(2).Text = PonerNombreDeCod(Text1(4), conAri, "smarca", "nommarca")
    Text2(3).Text = PonerNombreDeCod(Text1(5), conAri, "sunida", "nomunida")
    Text2(4).Text = PonerNombreDeCod(Text1(6), conAri, "stipar", "nomtipar")
    mPorIva = "porceiva"
    Text2(5).Text = DevuelveDesdeBD(conConta, "nombriva", "tiposiva", "codigiva", Text1(7).Text, "N", mPorIva)
    Text2(22).Text = PonerNombreDeCod(Text1(22), conAri, "scateg", "descateg")
    
    
    lblIndicador.Caption = "Importes"
    lblIndicador.Refresh
    PonerSumaStocks 'Poner la suma total de stocks de los almacenes donde esta el artic
    
    BloquearChecks Me, Modo

    PrimeraVez = False

    PonerCamposLineas True 'Pone los datos de las tablas de lineas de Componentes e Instalaciones
    
    'Lista campos
    CargaDatosLW
    
    'Pongo el PVP con IVA
    If mPorIva = "porceiva" Then mPorIva = 0
    Impor = CCur(mPorIva)
    Impor = Round2((Impor * Data1.Recordset!preciove) / 100, 4) + Data1.Recordset!preciove
    Me.txtPVPIVA.Text = Format(Impor, FormatoPrecio)
    
    
    
    
    
    'Si tiene conjuntos
    MostrarSolapa = False
    If vUsu.Nivel = 0 Then
        MostrarSolapa = True
    Else
        If Not vParamAplic.EsAVAB Then MostrarSolapa = True
    End If
    If MostrarSolapa Then
        If Val(Data1.Recordset!Conjunto) = 1 Then ponerDatosConjuntos
    End If
    
    If Not vParamAplic.EsAVAB Then
        'AHORA LO LLEVA EL PROPIO ARTICULO
        If IsNull(Data1.Recordset!tipartic) Then
            TipoFamilia = -1
        Else
            TipoFamilia = Val(Data1.Recordset!tipartic)
        End If
        lblIndicador.Caption = "Ficha tecnica"
        lblIndicador.Refresh
        PonerDatosFichaTecnica2 TipoFamilia
    End If
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub


Private Sub PonerCamposLineas(enlaza As Boolean)
'Carga las Pestañas con las tablas de lineas de Conjunto o Instalaciones
'segun la pestaña de datos a mostrar
On Error GoTo EPonerLineas

    Screen.MousePointer = vbHourglass
    'Conjuntos
    CargaGrid DataGrid1, Data2, enlaza
    'Instalaciones
    CargaGrid DataGrid2, Data3, enlaza
    'Stocks
    CargaGrid DataGrid3, Data4, enlaza


    Screen.MousePointer = vbDefault
    Exit Sub
    
EPonerLineas:
    MuestraError Err.Number, "PonerCamposLineas"
'    PonerModo 2
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerSumaStocks()
Dim rst As ADODB.Recordset
Dim SQL As String
    
    'NUEVO
    SQL = "select sum(canstock) from salmac where codartic=" & DBSet(Text1(0).Text, "T")
    If vUsu.TrabajadorB Then
        If Me.chkConso.Value = 0 Then SQL = SQL & "  AND salmac.codalmac= " & vParamAplic.AlmacenB
    Else
        SQL = SQL & "  AND salmac.codalmac <> " & vParamAplic.AlmacenB
    End If
   
    Set rst = New ADODB.Recordset
    rst.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If rst.EOF Then
        Me.txtSumaStock.Text = ""
    Else
         Me.txtSumaStock.Text = DBLet(rst.Fields(0).Value, "N")
    End If
    rst.Close
    Set rst = Nothing
    
    
    
    
    
    
    
'
'
'    SQL = DevuelveDesdeBDNew(conAri, "salmac", "codartic", "codartic", Text1(0).Text, "T")
'    If SQL <> "" Then
'        SQL = "select sum(canstock) from salmac where codartic=" & DBSet(Text1(0).Text, "T")
'        Set rst = New ADODB.Recordset
'        rst.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        If Not rst.EOF Then
'            Me.txtSumaStock.Text = rst.Fields(0).Value
'        End If
'        rst.Close
'        Set rst = Nothing
'    Else
'        Me.txtSumaStock.Text = 0
'    End If
End Sub


Private Sub PonerCamposAlmacenes2()
    If Data4.Recordset.EOF Then Exit Sub
    PonerCamposFormaFrame Me, "Text3", Data4
    
    'Rellenar el nombre correspondiente al código de los TextBox de indice 8
    Text2(8).Text = PonerNombreDeCod(Text3(0), conAri, "salmpr", "nomalmac", "codalmac")
    
    'Rellenar el nombre correspondiente al código de ubicacion
    Text2(6).Text = PonerNombreDeCod(Text3(1), conAri, "subica", "nomubica", "codubica")
    
    'El check del inventario
    chkInventario.Value = DBLet(Data4.Recordset!statusin, "N")
    
    '-- Esto permanece para saber donde estamos
'    lblIndicador.Caption = Data4.Recordset.AbsolutePosition & " de " & Data4.Recordset.RecordCount
End Sub


'Private Function ComprobarEsInstalacion() As Boolean
'Dim devuelve As String
'Dim EsInstal As Boolean
'
'    EsInstal = False
'    If Not (vParamAplic.Frecuencias) Then Exit Function ' si no estan activadas las frecuencias no se muestra ná
'    If Text1(3).Text <> "" Then
'        devuelve = DevuelveDesdeBDNew(conAri, "sfamia", "instalac", "codfamia", Text1(3).Text, "N")
'        If devuelve = "1" Then
'            EsInstal = CBool(devuelve)
'        Else
'            EsInstal = False
'        End If
'    End If
'    ComprobarEsInstalacion = EsInstal
'End Function
'
'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte
Dim B As Boolean
Dim NumReg As Byte

    'Actualiza Iconos Insertar,Modificar,Eliminar
    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    'Modo 2. Hay datos y estamos visualizandolos
    B = (Kmodo = 2) Or (Modo = 5) Or (Modo = 6) Or (Modo = 7)
    
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda2 <> "" Then
        cmdRegresar.visible = B
        cmdRegresar.Caption = "&Regresar"
    Else
        cmdRegresar.visible = False
    End If
    
    'Poner Flechas de Desplazamiento Visibles o no
    NumReg = 1
    If (Modo = 2) Then
        If Not Data1.Recordset.EOF Then
            If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
        End If
    ElseIf Modo = 5 Then
        If Not Data4.Recordset.EOF Then
            If Data4.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
        End If
    End If
    B = (Modo = 2) Or (Modo = 5)
    DesplazamientoVisible Me.Toolbar1, btnPrimero, B, NumReg
    
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    'campos Precio medio y ponderado bloqueados, pq son calculados
    BloquearTxt Text1(13), True
    BloquearTxt Text1(14), True
    'fecha ultimo cambio PVP bloqueado pq se actualiza automaticamente
    BloquearTxt Text1(27), True
    
    Me.FrameArtxAlmac.Enabled = (Modo = 5)
    'Me.FrameArtxAlmac2.visible = (Modo = 5)
    If Me.FrameArtxAlmac.Enabled Then
        If Modo = 5 And ModificaLineas = 2 Then BloquearTxt Text3(0), True
         'Me.FrameArtxAlmac.Height = 2010
         'Me.FrameArtxAlmac.Top = 2260
         'Me.FrameArtxAlmac.Left = 360
    End If
    Me.FrameDatosAlmacen2.visible = (Modo <> 5)
        
    B = (Modo = 1 Or Modo = 3 Or Modo = 4) '1:Busqueda, 3:Insertar, 4:Modificar
    cboArticuloVarios.Enabled = B
    cboStatus.Enabled = B
    cboTipoArt.Enabled = B
    'Bloquear los checkbox
    BloquearChecks Me, Modo

    cmdCancelar.visible = B
    cmdAceptar.visible = B
    
    Me.imgFecha(i).Enabled = B
    For i = 0 To 5
        Me.imgCuentas(i).Enabled = B
    Next i
    
    chkVistaPrevia.Enabled = (Modo <= 2)

    'Bton generar denominacion solo en descriptores y en modo insertar
    B = False
    If vParamAplic.Descriptores Then
        If Modo = 3 Then
            B = True
        Else
            If Modo = 4 And vParamAplic.ContabilizacionMoixent Then B = True
        End If
    End If
    'Me.cmdGenerar.visible = vParamAplic.Descriptores And Modo = 3
    Me.cmdGenerar.visible = B

    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    PonerModoOpcionesMenu (Modo) 'Poner opciones de menu según modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
                        
    'Esto esta "a piñon". Modo=2 y
    'Si hay una empresa AVAB(exportadora) es que estamos en
    'Morales, con lo cual, puede imprimir fuichas tecnicas...
    cmdIMpriFT.visible = Modo = 2 And EmprAVAB > 0
    cmdImg.visible = Modo = 2 And EmprAVAB > 0
    
    
    imgPreciosProv.visible = Modo = 2 And Not vParamAplic.EsAVAB And vEmpresa.codempre <> EmprMorales
    
    
    'Los tag's de los campos de sctock NO estaran visibles si
    'inserto,modifico o busco en la PPAL
    If Modo = 1 Or Modo = 3 Or Modo = 4 Then
        AccionesSobreTagText3_ True, False
    Else
        'Los vuelvo a poner
        AccionesSobreTagText3_ False, False
    End If
    
    'El listview
    If Modo <> 2 Then lw1.ListItems.Clear

    B = Modo = 3 Or Modo = 4 '3:Insertar, 4:Modificar
    PonerModoText5 Not B
    'cmdACtualizar importes en conjuntos
    cmdActualizarImportes1(0).visible = Modo = 6 And (ModificaLineas <> 1)
    cmdActualizarImportes1(1).visible = Modo = 6 And (ModificaLineas <> 1)
End Sub


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim B As Boolean
Dim EsInstal As Boolean

    B = (Modo = 2) Or (Modo = 5) Or (Modo = 6) Or (Modo = 7)
    'Insertar
    Toolbar1.Buttons(6).Enabled = (B Or Modo = 0 Or Modo = 1)
    Me.mnNuevo.Enabled = (B Or Modo = 0 Or Modo = 1)
    'Modificar
    Toolbar1.Buttons(7).Enabled = B
    Me.mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(8).Enabled = B
    Me.mnEliminar.Enabled = B
    
    'Imprimir
    Toolbar1.Buttons(14).Enabled = Not DeConsulta

    B = (Modo = 2) And Not DeConsulta
    'Lineas Articulos x Almacen
    Toolbar1.Buttons(10).Enabled = B
    Me.mnMtoStocksAlm.Enabled = B
    'Lineas Conjuntos
    Toolbar1.Buttons(11).Enabled = (B And (Me.chkConjunto.Value = 1 Or Me.chkConjunto.Value = 2))
    Me.mnMtoConjuntos.Enabled = (B And (Me.chkConjunto.Value = 1 Or Me.chkConjunto.Value = 2))
    
    MostrarSolapa = (Me.chkConjunto.Value = 1 Or Me.chkConjunto.Value = 2)
    If vUsu.Nivel <> 0 Then
        If vParamAplic.EsAVAB Then MostrarSolapa = False
    End If
        
    Me.SSTab1.TabVisible(2) = MostrarSolapa
    If Me.SSTab1.TabVisible(2) Then
        Me.cmdActualizarImportes1(0).Enabled = Not DeConsulta And vUsu.Nivel <= 1
        Me.cmdActualizarImportes1(1).Enabled = Not DeConsulta And vUsu.Nivel <= 1
    End If
    
    'Lineas Instalaciones
    'EsInstal = ComprobarEsInstalacion
    EsInstal = True
    B = B And EsInstal
    Toolbar1.Buttons(12).Enabled = B
    Me.mnMtoInstalaciones.Enabled = B
    Me.SSTab1.TabVisible(3) = EsInstal

    B = (Modo = 0) Or (Modo = 2) Or (Modo = 1)
    'Buscar
    Toolbar1.Buttons(1).Enabled = B
    Me.mnBuscar.Enabled = B
    'Modificar
    Toolbar1.Buttons(2).Enabled = B
    Me.mnVerTodos.Enabled = B
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Sub PonerModoFrame(Kmodo As Byte)
Dim i As Byte

    ModoFrame = Kmodo
    
    Select Case ModoFrame
        Case 0  'MODO INICIAL
                For i = 0 To Me.Text3.Count - 1
                    BloquearTxt Text3(i), True
                Next i
                Me.imgFecha(2).Enabled = False
                Me.imgCuentas(6).Enabled = False
                Me.imgCuentas(7).Enabled = False
                Me.chkInventario.Enabled = False
                PonerBotonCabecera True
                
        Case 3  'Modo INSERTAR
                BloquearTxt Text3(0), False
                Text2(8).Text = ""
    End Select
    If ModoFrame = 3 Or ModoFrame = 4 Then
        '3=Insertar,  4=Modificar
        For i = 0 To Me.Text3.Count - 1
             BloquearTxt Text3(i), False
            If ModoFrame = 3 Then Text3(i).Text = ""
        Next i
        chkInventario.Enabled = True
        Me.imgFecha(2).Enabled = True
        Me.imgCuentas(6).Enabled = (ModoFrame = 3)
        Me.imgCuentas(7).Enabled = (ModoFrame = 3 Or ModoFrame = 4)
        PonerFoco Text3(1)
        PonerBotonCabecera False
    End If
End Sub


Private Function DatosOk() As Boolean
Dim B As Boolean


    DatosOk = False
    
    'Comprobamos que el campo dias de garantia si no tiene valor lo
    'ponemos a 0 para q no de error que no puede ser nulo
    If Trim(Me.Text1(11).Text) = "" Then Text1(11).Text = "0"
    
    B = CompForm(Me, 1)
    If Not B Then Exit Function
    
    'Comprobar si ya existe el cod de trabajador en la tabla
    If Modo = 3 Then 'Insertar
        If ExisteCP(Text1(0)) Then
            B = False
        Else
            'No podemos crear este articulo ya que es una constante que utiliza
            If Text1(0).Text = "@1@" Then
                MsgBox "Imposible crear articulo @1@", vbExclamation
                B = False
            End If
            
            If Mid(Text1(0).Text, 1, 2) = "::" Then
                MsgBox "Imposible crear articulo ::", vbExclamation
                B = False
            End If
        
        
        End If
        
        
        
        If B Then
            '---------------------------------
            'Ficha técnica
            If Not vParamAplic.EsAVAB Then
                If Me.cboTipoArt.ListIndex = -1 Then
                    MsgBox "Falta tipo articulo", vbExclamation
                    B = False
                End If
            End If
        End If
    End If
    
    
    'si se ha cambiado el precio venta PVP actualizamos la fecha de
    'ult. cambio PVP
    If Modo = 4 Then 'modo modificar
        'si se ha modificado el ult. precio compra la fecha ult. compra
        'debe tener valor
        If Text1(15).Text <> "" And Trim(Text1(18).Text) = "" Then
            B = False
            MsgBox "Si hay precio de ult. compra la fecha de ult. compra debe tener valor.", vbInformation
        End If
        
        
        'si se ha modificado el precio venta PVP actualizamos campos
        'para guardarlo correctamente
        If CCur(Me.Text1(17).Text) <> CCur(Me.Data1.Recordset!preciove) Then
            Me.Text1(27).Text = Format(Now, "dd/mm/yyyy")
        End If
        
        
        
        'Cuando modificamos, si pasamos un articulo a CADUCADO, entonces comproaremos
        'si tiene sctock. Si es asi NO dejammos continuar
        If Me.cboStatus.ListIndex = 2 And Val(Data1.Recordset!codstatu) < 2 Then
            If Me.chkctrstock.Value = 1 Then
                'Lleva stcok
                'Comprobamos k valor tiene
                BuscaChekc = TotalRegistros("select sum(canstock) from salmac where codartic='" & DevNombreSQL(Text1(0).Text) & "'")
                If Val(BuscaChekc) > 0 Then
                    MsgBox "No podemos pasar un árticulo a caducado teniendo stock.", vbExclamation
                    Exit Function
                End If
            End If
        End If
        
        
        
    End If
    
    'El facto de conversion NO puede ser cero
    If B Then
        If ImporteFormateado(Text1(30).Text) = 0 Then
            MsgBox "El factor de conversion NO puede ser cero", vbExclamation
            B = False
        End If
    End If
    
    DatosOk = B
End Function


Private Function DatosOkConjunto() As Boolean
Dim B As Boolean
Dim devuelve As String

    DatosOkConjunto = False
    B = True
    If txtAux(1).Text = "" Then
         MsgBox "El campo Cantidad no puede ser nulo", vbExclamation, "Artículos"
         B = False
    End If
        
    If Not IsNumeric(txtAux(1).Text) Then
        MsgBox "La cantidad de Artículos tiene que ser numérico", vbExclamation
        B = False
    End If
    If Not B Then Exit Function
    
    'Comprobamos  si existe, solo si estamos insertando (ModificaLineas=1)
    'conAri: conexion a BD Ariges
    devuelve = DevuelveDesdeBDNew(conAri, "sarti1", "codartic", "codartic", Text1(0).Text, "T", , "codarti1", txtAux(0).Text, "T")
    If ModificaLineas = 1 And devuelve <> "" Then
        B = False
        devuelve = "Ya existe el Artículo de Conjunto: " & vbCrLf
        devuelve = devuelve & "Codigo: " & txtAux(0).Text & vbCrLf
        devuelve = devuelve & "Descripción: " & txtAux2.Text
        
        MsgBox devuelve, vbExclamation, "Artículos"
    End If
    If Not B Then Exit Function
    
    'Comprobar que el articulo no tiene conjuntos, solo si estamos insertando (ModificaLineas=1)
    'Si tiene conjuntos no puede ser elemento de conjunto de otro articulo
    If ModificaLineas = 1 And DevuelveDesdeBDNew(conAri, "sartic", "conjunto", "codartic", txtAux(0).Text, "N") = "1" Then
        B = False
        devuelve = "No es un Artículo de Conjunto: " & vbCrLf
        devuelve = devuelve & "Codigo: " & txtAux(0).Text & vbCrLf
        devuelve = devuelve & "Descripción: " & txtAux2.Text & vbCrLf & vbCrLf
        devuelve = devuelve & "¿Continuar?"
        If MsgBox(devuelve, vbQuestion + vbYesNo) = vbYes Then B = True
    End If
    DatosOkConjunto = B
End Function


Private Function DatosOkLinea() As Boolean
Dim B As Boolean
Dim devuelve As String

    DatosOkLinea = False
    B = True
    
    If Trim(Text3(1).Text) = "" Then 'Campo Ubicación
        MsgBox "El campo Ubicación no puede ser nulo", vbExclamation, "Artículos"
        B = False
    End If
    
    'Campo de cantidad de Stock (Son decimales)
    If Trim(Text3(2).Text) = "" Or IsNull(Text3(2).Text) Then
        MsgBox "El campo Cantidad Stock no puede ser nulo", vbExclamation, "Artículos"
        B = False
    End If
    If Not B Then Exit Function
    
    'Comprobamos  si existe
    devuelve = DevuelveDesdeBDNew(conAri, "salmac", "codartic", "codartic", Text1(0).Text, "T", , "codalmac", Text3(0).Text, "N")
    If ModificaLineas = 1 And devuelve <> "" Then
        B = False
        devuelve = "Ya existe el Artículo en el Almacen: " & vbCrLf
        devuelve = devuelve & "Codigo: " & Text3(0).Text & vbCrLf
        devuelve = devuelve & "Descripción: " & Text2(8).Text
        MsgBox devuelve, vbExclamation, "Artículos"
    End If
    
    DatosOkLinea = B
End Function


Private Sub Text3_GotFocus(Index As Integer)
    kCampo = Index
    If ModificaLineas <> 0 Then
        ConseguirFoco Text3(Index), 4
    Else
        ConseguirFoco Text3(Index), 2
    End If
End Sub

Private Sub Text3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        If Index = 8 Then
            PonerFocoBtn Me.cmdAceptar
        Else
            KeyAscii = 0
            SendKeys "{tab}"
        End If
    End If
End Sub


Private Sub Text3_LostFocus(Index As Integer)
    
     If Screen.ActiveForm.ActiveControl.Name = "cmdCancelar" Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0 'Codigo Almacen
             Text2(8).Text = PonerNombreDeCod(Text3(Index), conAri, "salmpr", "nomalmac")
             
        Case 1 'Codigo ubicacion
            Text2(6).Text = PonerNombreDeCod(Text3(Index), conAri, "subica", "nomubica", "codubica")
            If Text2(6).Text = "" And Text3(Index) <> "" Then PonerFoco Text3(Index)
                
        Case 2, 3, 4, 5, 6 'Stocks, Punto Pedido
                'Formato tipo 1: Decimal(12,2)
                If Trim(Text3(Index)) <> "" Then PonerFormatoDecimal Text3(Index), 1
        
        Case 7  'Fecha Inventario
            If Text3(Index).Text <> "" Then PonerFormatoFecha Text3(Index)

        Case 8  'Hora Inventario
            If Trim(Text3(Index).Text) <> "" Then PonerFormatoHora Text3(Index)
    End Select
End Sub



Private Sub Text5_GotFocus(Index As Integer)
    ConseguirFoco Text5(Index), 4
End Sub

Private Sub Text5_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text5_LostFocus(Index As Integer)
    Text5(Index).Text = Trim(Text5(Index).Text)
    If Text5(Index).Text = "" Then Exit Sub
    Select Case Index
    Case 0, 1, 7, 8, 9, 19
        If Not PonerFormatoDecimal(Text5(Index), 2) Then Text5(Index).Text = ""
    Case 17
        'Medidas
        If Not ComprobarMedidas Then
            MsgBox "Campo incorrecto :  NNN*NNN*NNN", vbExclamation
            Text5(17).Text = ""
            PonerFoco Text5(17)
        
        End If
    Case 4, 5, 20
        If Not PonerFormatoEntero(Text5(Index)) Then Text5(Index).Text = ""
    End Select
End Sub


Private Function ComprobarMedidas() As Boolean
Dim Inicio As Integer
Dim J As Integer
Dim CuantosAsteriscos As Integer
Dim Aux As String
Dim Multip As Currency

On Error GoTo EComprobarMedidas
    ComprobarMedidas = False
    Inicio = 1
    CuantosAsteriscos = 0
    Multip = 1
    Do
        J = InStr(Inicio, Text5(17).Text, "*")
        If J = 0 Then
            'No HAY MAS
        
        Else
            Aux = Mid(Text5(17).Text, Inicio, J - Inicio)
            If Not IsNumeric(Aux) Then
                Exit Function
            Else
                
                Multip = Multip * Val(Aux)
                
                
                CuantosAsteriscos = CuantosAsteriscos + 1
            
                Inicio = J + 1
                
                'Es e
                If CuantosAsteriscos = 2 Then
                    Aux = Mid(Text5(17).Text, J + 1)
                    If Not IsNumeric(Aux) Then
                        Exit Function
                    Else
                        Multip = Multip * Val(Aux)
                    End If
                End If
            End If
        End If
    Loop Until J = 0
    
    'Si llega aqui:
    If CuantosAsteriscos <= 1 Then
        'MAL
        
    ElseIf CuantosAsteriscos > 2 Then
        
    Else
        Multip = Round(Multip / 1000000000, 4)
        Text5(18).Text = Format(Multip, FormatoPrecio)
        ComprobarMedidas = True
    End If
        
EComprobarMedidas:
    If Err.Number <> 0 Then Err.Clear
End Function

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Buscar
            mnBuscar_Click
        Case 2  'Todos
            mnVerTodos_Click
        Case 6  'Nuevo
           mnNuevo_Click
        Case 7  'Modificar
            mnModificar_Click
        Case 8  'Borrar
            mnEliminar_Click
            
        Case 10  'Stocks Almacenes
            mnMtoStocksAlm_Click
        Case 11 'Conjuntos
            mnMtoConjuntos_Click
        Case 12 'Instalaciones
            mnMtoInstalaciones_Click
            
        Case 14 'Imprimir Listado de Articulos
            BotonImprimir
        Case 15 'Salir
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


Private Sub CargarComboStatus()
'### Combo Situación Artículo
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Normal, 1-Bloqueado, 2-Caducado

    cboStatus.Clear
    cboStatus.AddItem "Normal"
    cboStatus.ItemData(cboStatus.NewIndex) = 0
    
    cboStatus.AddItem "Bloqueado"
    cboStatus.ItemData(cboStatus.NewIndex) = 1
    
    cboStatus.AddItem "Caducado"
    cboStatus.ItemData(cboStatus.NewIndex) = 2
    
End Sub


Private Sub CargarComboArticuloVarios()
'### Combo Situación Artículo
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-No, 1-Si, 2-Rectificacion
 
    cboArticuloVarios.Clear
    cboArticuloVarios.AddItem "No"
    cboArticuloVarios.ItemData(cboArticuloVarios.NewIndex) = 0
    
    cboArticuloVarios.AddItem "Si"
    cboArticuloVarios.ItemData(cboArticuloVarios.NewIndex) = 1
    
    cboArticuloVarios.AddItem "Rectificación"
    cboArticuloVarios.ItemData(cboArticuloVarios.NewIndex) = 2
    
End Sub

Private Function InsetarArticulosPorAlmacen(EnAVAB As Boolean, cadErr As String) As Boolean
'Inserta en la tabla salmac una fila del artículo que se esta insertando
'para cada uno de los almacenes que existen en la tabla salmpr
Dim vCodartic As String, vCodAlmac As Integer
Dim rsAlmPr As ADODB.Recordset
Dim cad As String


    On Error GoTo EInsEnAlm

    vCodartic = Text1(0).Text
    Set rsAlmPr = New ADODB.Recordset
    cad = "Select codalmac from "
    
    'If EnAVAB Then Cad = Cad & "ariges" & EmpresaAVAB & "."
    If EnAVAB Then cad = cad & "ariges" & EmprAVAB & "."

    
    cad = cad & "salmpr order by codalmac;"
    rsAlmPr.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    While Not rsAlmPr.EOF
        vCodAlmac = rsAlmPr.Fields(0).Value
        cad = "INSERT INTO "
        If EnAVAB Then cad = cad & "ariges" & EmprAVAB & "."
        cad = cad & "salmac (codartic,codalmac,ubialmac,canstock,stockmin,puntoped,stockmax,stockinv,fechainv,"
        cad = cad & "horainve,statusin"
        'ENERO 2010 preciomp precioma preciouc preciost
        cad = cad & ",preciomp, precioma , preciost,preciouc)"
        cad = cad & " VALUES (" & DBSet(vCodartic, "T") & "," & vCodAlmac & ",'',0,0,0,0,0,NULL,NULL,0"
        'ENERO 2010 preciomp precioma preciouc preciost
        cad = cad & "," & DBSet(Text1(13).Text, "N", "N")
        cad = cad & "," & DBSet(Text1(14).Text, "N", "N")
        cad = cad & "," & DBSet(Text1(15).Text, "N", "N")
        cad = cad & "," & DBSet(Text1(16).Text, "N", "N") & ")"
        Conn.Execute cad
        rsAlmPr.MoveNext
    Wend
        
    rsAlmPr.Close
    Set rsAlmPr = Nothing
    InsetarArticulosPorAlmacen = True
    Exit Function
    
EInsEnAlm:
    InsetarArticulosPorAlmacen = False
    'MuestraError Err.Number, "Insertando Artículo en Almacenes.", Err.Description
    cadErr = "Insertando Artículo en Almacenes: " & vbCrLf & Err.Description
End Function
   
   

Private Function InsertarPreciosPorTarifa(Optional cadErr As String) As Boolean
'Insertar en la lista de precios las tarifas para el articulo
Dim SQL As String
Dim RS As ADODB.Recordset
Dim cTar As CTarifaArt
Dim NoOK As Boolean

    On Error GoTo ErrInsPrecio
    
    'comprobamos que el PVP tenga valor y sea >0 para insertar lista de precios
    InsertarPreciosPorTarifa = True
    If Text1(17).Text = "" Then Exit Function
    If Not (CCur(Text1(17).Text) > 0) Then Exit Function
    
    
    
    
    
    
    InsertarPreciosPorTarifa = False
    
    
    
    
    'seleccionar todas las posibles tarifas
    SQL = "SELECT * FROM starif WHERE NOT ISNULL(margecom) "
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'para cada tarifa insertar un linea en la tabla de lista de precios
    'por cada codartic,codtarif
    
    '23 Abril 2008
    'Tb, en funcion sobre donde se aplica el margen se hara una cosa u otra
    ' Sobre PVP o sobre PUC
    'FALTA###
    NoOK = False
    While Not RS.EOF
        Set cTar = New CTarifaArt
        cTar.CodigoArticulo = Text1(0).Text
        cTar.CodigoTarifa = RS!codlista
        'Aqui dependera de una cosa u otra para lo del PVP / UPC
        cTar.PrecioActual = CCur(Text1(17).Text) 'precio venta al publico (pvp)
        If cTar.InsertarPrecios = False Then NoOK = True
        Set cTar = Nothing
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
        
    If NoOK Then
        InsertarPreciosPorTarifa = False
        cadErr = "Los precios del artículo por tarifa NO se han introducido correctamente."
    Else
        InsertarPreciosPorTarifa = True
    End If
        
    Exit Function
    
ErrInsPrecio:
    InsertarPreciosPorTarifa = False
    cadErr = "Insertar precios por tarifa: " & Err.Description
End Function
   
   
Private Function BloquearTarifas(codartic As String) As Boolean
Dim cadWhere As String
    cadWhere = "codartic=" & DBSet(codartic, "T")
    BloquearTarifas = BloqueaRegistro("slista", cadWhere)
End Function
   
   
Private Function ActualizarPreciosVenta() As Boolean
'si se modifica el precio ult. compra a mano preguntar si quiere modificar
'el PVP y las tarifas de venta desde el formulario de actualizar precios
Dim precioUC As Currency 'precio ult. compra (valor actual)
Dim FechaUC As String
Dim newPrecioUC As Currency
Dim bActualizar As Boolean
Dim cad As String

    'Comprobar si se ha modificado el precio desde la ultima compra
    'y preguntar quiere modificar el PVP del articulo aplicandole su margen
    'y el precio de las TArifas aplicandole el margen
    '-- Laura 19/12/2006: el precio de compra es el precio con los descuentos (importe/cantidad)
    precioUC = CCur(DBLet(Me.Data1.Recordset!precioUC, "N"))
    If Not IsNull(Me.Data1.Recordset!ultfecco) Then FechaUC = DBLet(Me.Data1.Recordset!ultfecco, "F")
    newPrecioUC = ImporteFormateado(Text1(15).Text)
    
    bActualizar = False
    cad = ""
    If precioUC <> newPrecioUC Then
        If FechaUC = "" Then
            bActualizar = True
        ElseIf CDate(Text1(18).Text) >= CDate(FechaUC) Then
            bActualizar = True
        Else
            
        End If
        cad = "precio de última compra"
    End If
    
    
    '## LAURA 25/06/2008
    If Not bActualizar Then
        '-- comprobar si se ha modificado el margen comercial y
        '-- en este caso recalcular tambien el PVP y tarifas
        precioUC = CCur(DBLet(Me.Data1.Recordset!margecom, "N")) 'margen actual
        newPrecioUC = ImporteFormateado(Text1(25).Text) 'margen nuevo
        If precioUC <> newPrecioUC Then bActualizar = True
        cad = "margen comercial"
    End If
    '##
    
    
     If bActualizar Then
            'FALTA### ver si esta bien. Ahora NO actualiza tarifas
            If False Then
                If MsgBox("Se ha modificado el " & cad & "." & vbCrLf & "¿Desea actualizar los precios de venta?", vbQuestion + vbYesNo) = vbYes Then
                    'Comprobar que el artículo tiene margen comercial
                    If ArticuloTieneMargen(Text1(0).Text) Then
                        'Llamar al form de actualizar precios venta
                        frmComActPrecios.parCodArtic = Text1(0).Text
                        frmComActPrecios.parNomArtic = Text1(1).Text
                        frmComActPrecios.Show vbModal
                    End If
                End If
            End If
        End If
    
    
    
'        If CDate(Text1(18).Text) >= CDate(FechaUC) Then
'            If MsgBox("Se ha modificado el precio última compra." & vbCrLf & "¿Desea actualizar los precios de venta?", vbQuestion + vbYesNo) = vbYes Then
'                'Comprobar que el artículo tiene margen comercial
'                If ArticuloTieneMargen(txtAux(1).Text) Then
'                    'bloquear las tarifas del articulo para modificar
''                                If BloqueaRegistro("slista", "codartic=" & DBSet(txtAux(1).Text, "T")) Then
'                        'Aplicar margen comercial a los precios
'                        'Modificar precios de venta en articulo y tarifas
'                        frmComActPrecios.parCodArtic = txtAux(1).Text
'                        frmComActPrecios.parNomArtic = txtAux(2).Text
''                            frmcomactprecios.parPrecioUC =
'                        frmComActPrecios.Show vbModal
''                                End If
'                End If
'            End If
'        End If   'Fecha ultima compra
'    End If  'Precio ultima compra


End Function
  
  
  
  
Private Function ActualizarPreciosPorTarifa() As Boolean
Dim QueTipo As Byte
Dim Importe As Currency
Dim Aux As Currency
       'Reutilizo BuscaChekc
       QueTipo = 100
       BuscaChekc = ""
       
       '- ver si se ha modificado el precion venta PVP
       Importe = DBLet(Data1.Recordset!preciove, "N")
       If Importe <> CCur(Text1(17).Text) Then
            BuscaChekc = "-el precio de venta." & vbCrLf
            QueTipo = 0 'que mire tarifas PVP
       End If
        
       '- ver si se ha modificado el precio ultima compra
       Importe = DBLet(Data1.Recordset!precioUC, "N")
       Aux = 0
       If Text1(15).Text <> "" Then Aux = CCur(Text1(15).Text)
       If Importe <> Aux Then
            BuscaChekc = BuscaChekc & "-el precio de ultima compra." & vbCrLf
            If Aux = 0 Then BuscaChekc = BuscaChekc & "*****  Precio ultima compra=  CERO    ****** " & vbCrLf
            If QueTipo = 0 Then
                QueTipo = 2  'Que mire las dos
            Else
                QueTipo = 1  'que mire solo en tarifas U.P.C.
            End If
        End If
            
        '## LAURA 25/06/2008
        'si el tipo es 0 o 2 ya se va a modificar el PVP y no comprobamos margen
'        If QueTipo <> 0 And QueTipo <> 2 Then
'            'si se ha modificado el margen comercial tambien
'            'hay que actualizar el PVP
'            Importe = DBLet(Data1.Recordset!margecom, "N")
'            If Importe <> CCur(Text1(25).Text) Then
'                 BuscaChekc = "-el margen comercial." & vbCrLf
'                 If QueTipo = 1 Then
'                    QueTipo = 2  'Que mire las dos
'                 Else
'                    QueTipo = 0 'que mire tarifas PVP
'                 End If
'            End If
'        End If
        '##
        
            
        If QueTipo <> 100 Then
            'FALTA### Comprobar esto. Lo he quitado (para morales por supuesto)
         '   BuscaChekc = vbCrLf & BuscaChekc & vbCrLf
         '   BuscaChekc = "Se han modificado: " & BuscaChekc & "¿Desea actualizar las tarifas de precios?"
         '   If MsgBox(BuscaChekc, vbQuestion + vbYesNo) = vbYes Then
         '
         '       Screen.MousePointer = vbHourglass
         '       ActualizarPreciosPorTarifaDOS QueTipo
         '       Screen.MousePointer = vbDefault
         '   End If
        End If
    
    
End Function
  
  
                    'QueTipoActualiza : 0. PVP
                    '                   1. UPC
                    '                   2. LOS DOS
Private Function ActualizarPreciosPorTarifaDOS(PVP As Byte, Optional cadErr As String) As Boolean
'Actualiza en la lista de precios las tarifas para el articulo
Dim SQL As String
Dim RS As ADODB.Recordset
Dim cTar As CTarifaArt
Dim NoOK As Boolean
Dim menErr As String
Dim newPrecio As Currency


    On Error GoTo ErrActPrecio
    
    '-- comprobamos que el PVP tenga valor y sea >0 para insertar lista de precios
    ActualizarPreciosPorTarifaDOS = True
    
    
    
    '-- comprobar que para ese articulo en la tabla de tarifas no haya ningun registros
    '   con valor en el campo precio_nuevo
    SQL = "SELECT COUNT(*) FROM slista WHERE codartic=" & DBSet(Text1(0).Text, "T")
    SQL = SQL & " AND not isnull(precionu) and precionu>0"
    If RegistrosAListar(SQL) > 0 Then
        MsgBox "No se pueden actualizar las tarifas del artículo." & vbCrLf & "Tiene precios nuevos.", vbExclamation
        Exit Function
    End If
    
    
    ActualizarPreciosPorTarifaDOS = False
    
    If Not BloquearTarifas(Text1(0).Text) Then
        MsgBox "NO se han actualizado las tarifas de precios.", vbExclamation, "Actualizar precios"
        Exit Function
    End If
    
    
    '-- seleccionar todas las posibles tarifas
    SQL = "SELECT * FROM starif WHERE NOT ISNULL(margecom) "
    If PVP < 2 Then
        'Sera de uno de los tipos
        SQL = SQL & " AND opcionINC = " & CStr(PVP)
    End If
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'para cada tarifa actualizar la linea en la tabla de lista de precios
    'por cada codartic,codtarif
    NoOK = False
    While Not RS.EOF
        If BloquearTarifas(Text1(0).Text) Then
            Set cTar = New CTarifaArt
            If cTar.LeerDatos(Text1(0).Text, RS!codlista) Then
                
                If cTar.TarifaSobre = 0 Then
                    'TARIFAS SOBRE PVP
                    newPrecio = Round2((CCur(Text1(17).Text) * cTar.MargenComercial) / 100, 4)
                    newPrecio = CCur(Text1(17).Text) + newPrecio
                    
                Else
                    'TARIFAS SOBRE UPC
                    newPrecio = Round2((CCur(Text1(15).Text) * cTar.MargenComercial) / 100, 4)
                    newPrecio = CCur(Text1(15).Text) + newPrecio
                End If
                
                If cTar.ActualizarPrecios(Format(Now, "dd/mm/yyyy"), newPrecio, 0, menErr, False) = False Then NoOK = True
            Else
                'si no existe el articulo con esa tarifa la damos de alta
                cTar.CodigoArticulo = Text1(0).Text
                cTar.CodigoTarifa = RS!codlista
                'Si la tarifa es sobre PVP, mando el PVP
                'Si es sobre el UPC mando el UPC
                If DBLet(RS!opcionINC, "N") = 0 Then
                    'PVP
                    cTar.PrecioActual = CCur(Text1(17).Text) 'precio venta al publico (pvp)
                Else
                    cTar.PrecioActual = CCur(Text1(15).Text) 'precio venta al publico (pUC)
                End If
                
                If Not cTar.InsertarPrecios Then NoOK = True
            End If
            Set cTar = Nothing
        Else
            NoOK = True
'            MsgBox "NO se han actualizado correctamente todas las tarifa del artículo.", vbExclamation, "Actualizar precios"
            'Exit Function
        End If
        
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
        
    If NoOK Then
        ActualizarPreciosPorTarifaDOS = False
        cadErr = "NO se han actualizado correctamente todas las tarifa del artículo."
        cadErr = cadErr & vbCrLf & menErr
        MsgBox cadErr, vbExclamation, "Actualizar Precios"
    Else
        ActualizarPreciosPorTarifaDOS = True
    End If
        
    Exit Function
    
ErrActPrecio:
    ActualizarPreciosPorTarifaDOS = False
    cadErr = "Actualizar precios por tarifa: " & Err.Description
    MsgBox cadErr, vbExclamation
End Function
   
    
Private Function InsertarModificarLinea() As Boolean
Dim i As Integer
Dim SQL As String

    On Error GoTo EInsertarModificarLinea

    InsertarModificarLinea = False
    SQL = ""
    Select Case ModificaLineas
    Case 1  'INSERTAr
        'ENERO 2010 preciomp precioma preciouc preciost
        If DatosOkLinea Then 'INSERTAR
            SQL = "INSERT INTO salmac(codartic,codalmac,ubialmac,canstock,stockmin,puntoped,stockmax,stockinv,fechainv,"
            SQL = SQL & "horainve,statusin"
            'ENERO 2010 preciomp precioma preciouc preciost
            SQL = SQL & ",preciomp, precioma, preciost,preciouc) VALUES ("
            SQL = SQL & DBSet(Text1(0).Text, "T") & ", "
            SQL = SQL & Text3(0).Text & ", "
            SQL = SQL & DBSet(Text3(1).Text, "T") & ", "
            
            'Campos Stocks (Son Decimales)
            SQL = SQL & DBSet(Text3(2).Text, "N", "N") & ", "
            For i = 3 To 6
                SQL = SQL & DBSet(Text3(i).Text, "N", "S") & ", "
            Next i
        
            'Campo Fecha
            SQL = SQL & DBSet(Text3(7).Text, "F", "S") & ", "
'            If Trim(Text3(7).Text) <> "" Then
'              SQL = SQL & DBSet(Text3(7).Text, "F") & ", "
'            Else
'              SQL = SQL & "NULL, "
'            End If
        
            If Trim(Text3(8).Text) <> "" Then     'Campo Hora
              SQL = SQL & Format(Text3(8).Text, "hh:mm:ss") & ", "
            Else
              SQL = SQL & "NULL, "
            End If
        
            SQL = SQL & chkInventario.Value
            ''ENERO 2010 preciomp precioma preciouc preciost
            'De momento, al insertar inserto los valores de la ficha. Al modificar NO actualizo los precios
            SQL = SQL & "," & DBSet(Text1(13).Text, "N", "N")
            SQL = SQL & "," & DBSet(Text1(14).Text, "N", "N")
            SQL = SQL & "," & DBSet(Text1(15).Text, "N", "N")
            SQL = SQL & "," & DBSet(Text1(16).Text, "N", "N") & ")"
        End If
        
    Case 2  'MODIFICAR
        If DatosOkLinea Then
            SQL = "UPDATE salmac Set ubialmac = " & DBSet(Text3(1).Text, "T") & ", "
            SQL = SQL & " canstock = " & DBSet(Text3(2).Text, "N") & ", "
            SQL = SQL & " stockmin = " & DBSet(Text3(3).Text, "N", "S") & ", "
            SQL = SQL & " puntoped = " & DBSet(Text3(4).Text, "N", "S") & ", "
            SQL = SQL & " stockmax = " & DBSet(Text3(5).Text, "N", "S") & ", "
            SQL = SQL & " stockinv = " & DBSet(Text3(6).Text, "N", "S")
            If Trim(Text3(7).Text) <> "" Then _
            SQL = SQL & ", fechainv = " & DBSet(Text3(7).Text, "F", "S")
            If Trim(Text3(8).Text) <> "" Then
                SQL = SQL & ", horainve = '" & Format(Text3(8).Text, "hh:mm:ss") & "'"
            Else
                SQL = SQL & ", horainve = " & ValorNulo
            End If
            SQL = SQL & ", statusin = " & (chkInventario.Value)
            SQL = SQL & " WHERE codartic = " & DBSet(Text1(0).Text, "T") & " AND "
            SQL = SQL & " codalmac =" & Val(Text3(0).Text)
            
        End If
    End Select
        
    If SQL <> "" Then
        Conn.Execute SQL
        InsertarModificarLinea = True
    Else
        PonerFoco Text3(1)
    End If
    Exit Function

EInsertarModificarLinea:
    MuestraError Err.Number, "Insertar/Modificar Stocks Almacenes" & vbCrLf & Err.Description
End Function
    
    
Private Function InsertarArticulo() As Boolean
Dim B As Boolean
Dim menErr As String
Dim CadenaSarti4 As String


    On Error GoTo ErrInsArt
    Conn.BeginTrans
    

    
    B = InsertarDesdeForm(Me)
    If Not B Then menErr = "Insertando en tabla articulos"
    'insertar una linea en salmac para cada uno de los almacenes
    If B Then B = InsetarArticulosPorAlmacen(False, menErr)
    
    'insertar una linea de lista de precios para cada tarifa
    If B Then B = InsertarPreciosPorTarifa(menErr)
                
    'Inserta en sarti4
    CadenaSarti4 = ""  'para el AVAB si hiciera falta
    If Not vParamAplic.EsAVAB Then
        If EmprAVAB > 0 And vEmpresa.codempre = EmprMorales Then
            If B Then B = InsertarEnFichaTecnica(menErr, CadenaSarti4)
        End If
    End If
    
    If B Then
        Conn.CommitTrans
    Else
        Conn.RollbackTrans
        MsgBox menErr, vbExclamation
    End If
    
    InsertarArticulo = B
    If Not B Then Exit Function
    
    'Si ha ido bien, y la empresa NO es AVAB entonces intentamos crearlo en la
    'Empresa AVAB
    Espera 0.5
    If Not vParamAplic.EsAVAB And vEmpresa.codempre = EmprMorales Then
        
        'FALTA QUITAR DE AQUI AVAB
        If MsgBox("Crear la referencia en la empresa exportadora?", vbQuestion + vbYesNo) = vbYes Then
            menErr = ""
            
            Conn.BeginTrans
            B = InsertarArticuloAVAB(menErr, CadenaSarti4)
            
            If B Then B = InsetarArticulosPorAlmacen(True, menErr)
            
            If Not B Then MsgBox menErr, vbExclamation
            
            If B Then
                Conn.CommitTrans
            Else
                Conn.RollbackTrans
            End If
            
        End If
    End If
    
    Exit Function
                
ErrInsArt:
    Conn.RollbackTrans
    InsertarArticulo = False
    MuestraError Err.Number, "Insertar artículo.", Err.Description
End Function
    
    
    
    

Public Function InsertarModificarConjunto() As Boolean
Dim SQL As String
On Error GoTo EInsertarModificarLinea

    SQL = ""
    InsertarModificarConjunto = False
    
    If DatosOkConjunto Then
        Select Case ModificaLineas
        Case 1 'Insertar
                SQL = "INSERT INTO sarti1 VALUES ("
                SQL = SQL & DBSet(Text1(0).Text, "T") & ", "
                SQL = SQL & cmdAceptar.Tag & ", "
                SQL = SQL & DBSet(txtAux(0).Text, "T") & ", "
                SQL = SQL & DBSet(txtAux(1).Text, "S") & ") "
        Case 2 'Modificar
                SQL = "UPDATE sarti1 Set codarti1 = " & DBSet(txtAux(0).Text, "T")
                SQL = SQL & ", cantidad = " & DBSet(txtAux(1).Text, "S")
                SQL = SQL & " WHERE codartic =" & DBSet(Text1(0).Text, "T") & " AND "
                SQL = SQL & " numlinea =" & cmdAceptar.Tag
        End Select
    End If
    
    If SQL <> "" Then
        Conn.Execute SQL
        'If vEmpresa.codempre <> EmpresaAVAB Then ActualizaComponentesAVAB SQL   'Actualizar en AVAB
        If Not vParamAplic.EsAVAB Then ActualizaComponentesAVAB SQL    'Actualizar en AVAB
            
        InsertarModificarConjunto = True
        HayQueRecalcularPesos = True
    End If
    Exit Function
    
EInsertarModificarLinea:
    MuestraError Err.Number, "Insertar/Modificar Conjuntos" & vbCrLf & Err.Description
End Function


Public Function InsertarModificarInstalacion() As Boolean
Dim SQL As String
Dim Valor As String

On Error GoTo EInsertarModificarInstalacion
    InsertarModificarInstalacion = False
    Valor = Trim(txtAux(2).Text)
    If Valor = "" Then Valor = " "
    
    If ModificaLineas = 1 Then 'INSERTAR
        SQL = "INSERT INTO sarti2 VALUES ("
        SQL = SQL & DBSet(Text1(0).Text, "T") & ", "
        SQL = SQL & cmdAceptar.Tag & ", "
        SQL = SQL & DBSet(Valor, "T") & ") "
    ElseIf ModificaLineas = 2 Then 'MODIFICAR
        SQL = "UPDATE sarti2 Set licontro = " & DBSet(Valor, "T")
        SQL = SQL & " WHERE codartic =" & DBSet(Text1(0).Text, "T") & " AND "
        SQL = SQL & " numlinea =" & cmdAceptar.Tag
    End If
    
    Conn.Execute SQL
    InsertarModificarInstalacion = True
    Exit Function

EInsertarModificarInstalacion:
    MuestraError Err.Number, "Insertar/Modificar Instalación" & vbCrLf & Err.Description
End Function


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
Dim tots As String
Dim SQL As String

    On Error GoTo ECargaGrid
    
      
    If vDataGrid.Name = "DataGrid1" Then
        SQL = MontaSQLCarga(enlaza, 2)
        CargaGridGnral DataGrid1, Me.Data2, SQL, PrimeraVez
        tots = "N||||0|;N||||0|;S|txtAux(0)|T|Cod. Artículo|1800|;S|cmdAux|B||0|;S|txtAux2|T|Desc. Artículo|4100|;S|txtAux(1)|T|Cantidad|900|" & FormatoCantidad & "|;"
        tots = tots & "S|txtAux(3)|T|PVP|1000|;S|txtAux(4)|T|UPC|1000|;S|txtAux(5)|T|Factor conv.|1000|;"
        'El factor de conversion. NO lo muestro pero el datarecordset lo tiene data2!factorconversion
        'tots = tots & "N||||0|;"
        arregla tots, DataGrid1, Me
        DataGrid1.Columns(4).Alignment = dbgCenter
        DataGrid1.ScrollBars = dbgAutomatic
    ElseIf vDataGrid.Name = "DataGrid2" Then
        SQL = MontaSQLCarga(enlaza, 3)
        CargaGridGnral DataGrid2, Me.Data3, SQL, PrimeraVez
        tots = "N||||0|;N||||0|;S|txtAux(2)|T|Control Instalaciones|7100|;"
        arregla tots, DataGrid2, Me
        DataGrid2.ScrollBars = dbgAutomatic
    ElseIf vDataGrid.Name = "DataGrid3" Then
        SQL = MontaSQLCarga(enlaza, 4)
        CargaGridGnral DataGrid3, Me.Data4, SQL, PrimeraVez
        tots = "S|Text3(0)|T|Cod.Alm|1200|;S|cmdAlma|B||0|;S|Text2(8)|T|Nombre Almacen|2400|;S|Text3(2)|T|Stock|1200|;"
        'Los campos que no se ven que van FUERA DEL GRID
        tots = tots & "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
        arregla tots, DataGrid3, Me
        DataGrid3.ScrollBars = dbgAutomatic
 
        
    End If
    
    
    
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Function MontaSQLCarga(enlaza As Boolean, Opcion As Byte) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el Data
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim SQL As String

    If Opcion = 2 Then
        'cadena SQL para cargar los CONJUNTOS de la tabla sarti1
        'SQL = "SELECT sarti1.codartic,sarti1.numlinea,sarti1.codarti1,sartic.nomartic,sarti1.cantidad "
        'SQL = SQL & " FROM sarti1 INNER JOIN sartic ON sarti1.codarti1=sartic.codartic "
        
        'Marzo 2009
        'YA NO CARGAMOS EL PRECIO TARIFA
        'SQL = "SELECT sarti1.codartic, numlinea, sarti1.codarti1,sartic.nomartic,"
        'SQL = SQL & " sarti1.Cantidad , sartic.preciove, sartic.precioUC, slista.precioac , factorconversion"
        'SQL = SQL & " FROM   sarti1 INNER JOIN sartic ON"
        'SQL = SQL & " sarti1.codarti1 = sartic.codArtic"
        'SQL = SQL & " LEFT OUTER JOIN slista ON sarti1.codarti1=slista.codartic AND slista.codlista = " & vParamAplic.CodTarifa
        SQL = "SELECT sarti1.codartic, numlinea, sarti1.codarti1,sartic.nomartic,"
        SQL = SQL & " sarti1.Cantidad , sartic.preciove, sartic.precioUC, factorconversion"
        SQL = SQL & " FROM   sarti1 INNER JOIN sartic ON"
        SQL = SQL & " sarti1.codarti1 = sartic.codArtic"
        
        SQL = SQL & " where sarti1.codartic="
        If enlaza Then
            SQL = SQL & DBSet(Text1(0).Text, "T")
        Else
            SQL = SQL & "'-1@#'"
        End If
        SQL = SQL & " ORDER BY sarti1.numlinea "
        
        
    ElseIf Opcion = 3 Then 'INSTALACIONES
        SQL = "SELECT sarti2.codartic, sarti2.numlinea, sarti2.licontro "
        SQL = SQL & " FROM sarti2"
        If enlaza Then
            SQL = SQL & " WHERE sarti2.codartic=" & DBSet(Text1(0), "T")
        Else
            SQL = SQL & " WHERE sarti2.codartic= '-1'"
        End If
        SQL = SQL & " ORDER BY sarti2.numlinea"
    
    ElseIf Opcion = 4 Then 'STOCK
        
        SQL = "select salmac.codalmac,nomalmac,canstock,ubialmac,stockmin,puntoped,stockmax,stockinv,fechainv,horainve,statusin  "
        SQL = SQL & " from salmac,salmpr where salmac.codalmac=salmpr.codalmac  "
        
        If vUsu.TrabajadorB Then
            If Me.chkConso.Value = 0 Then SQL = SQL & " AND salmac.codalmac= " & vParamAplic.AlmacenB
        Else
            SQL = SQL & " AND salmac.codalmac <> " & vParamAplic.AlmacenB
        End If
        SQL = SQL & " AND "
        If enlaza Then
            SQL = SQL & " codartic=" & DBSet(Text1(0), "T")
        Else
            SQL = SQL & " codartic= '-1'"
        End If
    
    End If
    
    MontaSQLCarga = SQL
End Function


Private Sub LLamaLineas2(alto As Single, xModo As Byte, Opcion As Byte)
Dim B As Boolean

    ModificaLineas = xModo
    B = (Modo >= 5 Or Modo <= 7) And (ModificaLineas = 1 Or ModificaLineas = 2) 'Insertar o Modificar Lineas

    Select Case Opcion
    Case 2 'CONJUNTOS
        DeseleccionaGrid Me.DataGrid1
        
        txtAux(0).Height = DataGrid1.RowHeight
        txtAux(0).visible = B
        txtAux(0).Top = alto
        txtAux(1).Height = DataGrid1.RowHeight
        txtAux(1).visible = B
        txtAux(1).Top = alto
        txtAux2.Height = DataGrid1.RowHeight
        txtAux2.visible = B
        txtAux2.Top = alto
        cmdAux.visible = B
        cmdAux.Top = alto
        cmdAux.Height = DataGrid1.RowHeight
         
    Case 3 'INSTALACIONES
        DeseleccionaGrid Me.DataGrid2
        txtAux(2).Height = DataGrid2.RowHeight
        txtAux(2).visible = True
        txtAux(2).Top = alto
        
        
    Case 4
        'STOCK
        DeseleccionaGrid Me.DataGrid3
        Text3(0).Height = DataGrid3.RowHeight
        Text3(0).visible = B
        Text3(0).Top = alto
        Text3(2).Height = DataGrid3.RowHeight
        Text3(2).visible = B
        Text3(2).Top = alto
        Text2(8).Height = DataGrid3.RowHeight
        Text2(8).visible = B
        Text2(8).Top = alto
        
        If B Then
            If ModificaLineas = 1 Then
                cmdAlma.visible = B And ModificaLineas = 1
                cmdAlma.Top = alto
                cmdAlma.Height = DataGrid1.RowHeight
            Else
                cmdAlma.visible = False
                Text3(0).Width = DataGrid3.Columns(0).Width
            End If
        Else
            cmdAlma.visible = False
        End If
    End Select
End Sub


Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), Modo
End Sub


Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        If (Index = 2) Or Index = 1 Then
            KeyAscii = 0
            PonerFocoBtn Me.cmdAceptar
            Exit Sub
        End If
    End If
    KEYpress KeyAscii
End Sub


Private Sub txtAux_LostFocus(Index As Integer)

    txtAux(Index).Text = Trim(txtAux(Index).Text)
    If txtAux(Index).Text = "" Then Exit Sub
    
    Select Case Index
        Case 0 'cod Articulo de conjunto
            txtAux2.Text = PonerNombreDeCod(txtAux(Index), conAri, "sartic", "nomartic", "codartic", "Cod. Artículo", "T")
        Case 1
            'Nueva tipo single
            If Not PonerFormatoSingle(txtAux(Index), 5) Then
                txtAux(1).Text = ""
                PonerFoco txtAux(1)
            End If
    End Select
End Sub


Private Sub PonerBotonCabecera(B As Boolean)
On Error Resume Next

    Me.cmdAceptar.visible = Not B
    Me.cmdCancelar.visible = Not B
    Me.cmdRegresar.visible = B
    Me.cmdRegresar.Caption = "&Cabecera"
    If B Then
        Me.lblIndicador.Caption = "Lineas Detalle"
        PonerFocoBtn Me.cmdRegresar
    Else
        Me.lblIndicador.Caption = ""
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Function Eliminar() As Boolean
    Set LOG = New cLOG
    
    

    Conn.BeginTrans
    
    If EliminarArticulo(Data1.Recordset!codartic, lblIndicador) Then
        LOG.Insertar 7, vUsu, Data1.Recordset!codartic & " " & Data1.Recordset!NomArtic
        Conn.CommitTrans
        Eliminar = True
    Else
        Conn.RollbackTrans
        Eliminar = False
        
    End If
    Set LOG = Nothing
    lblIndicador.Caption = ""
End Function


Private Sub PosicionarData()
Dim cad As String, Indicador As String

    cad = "codartic=" & DBSet(Text1(0).Text, "T")
    If SituarData(Data1, cad, Indicador) Then
        PonerModo 2
        PonerCampos
        
        lblIndicador.Caption = Indicador
    ElseIf Not Data1.Recordset.EOF Then
'        Data1.Recordset.MoveFirst
        PonerCampos
        PonerModo 2
    ElseIf Modo = 3 Then
        'Acabamos de insertar un registro y lo seleccionamos en el recordset
        CadenaConsulta = "Select * from " & NombreTabla & " WHERE codartic =" & DBSet(Text1(0).Text, "T")
        Data1.RecordSource = CadenaConsulta
        If SituarData(Data1, cad, Indicador) Then
            PonerModo 2
            PonerCampos
            lblIndicador.Caption = Indicador
        End If
    Else
        LimpiarCampos
        PonerModo 0
    End If
End Sub


Private Sub BotonImprimir()
    
    AbrirListado (6) '6: Informe de Articulos

    
End Sub


Private Sub AccionesSobreTagText3_(Guardar As Boolean, Cargando As Boolean)
Dim i As Integer

  
    If Guardar Then
        If Cargando Then TagText3 = ""
        For i = 0 To Text3.Count - 1
            If Cargando Then TagText3 = TagText3 & Replace(Text3(i).Tag, "|", ";") & "|"
            Text3(i).Tag = ""
        Next i
        
        'AÑADIMOS EL CHECK chkInventario.
        If Cargando Then TagText3 = TagText3 & Replace(chkInventario.Tag, "|", ";") & "|"
        chkInventario.Tag = ""
    Else
        For i = 0 To Text3.Count - 1
            Text3(i).Tag = Replace(RecuperaValor(TagText3, i + 1), ";", "|")
        Next i
        chkInventario.Tag = Replace(RecuperaValor(TagText3, i + 1), ";", "|")
    End If
End Sub


Private Sub PonerDatosForaGrid(ForzarLimpiar As Boolean)
Dim i As Integer
Dim Limp As Boolean

    Limp = True
    If Not ForzarLimpiar Then
        If Not (Data4.Recordset Is Nothing) Then
            If Not Data4.Recordset.EOF Then Limp = False
        End If
    End If
    
    
    If Limp Then

        'Limpiamos
        For i = 0 To Text3.Count - 1
            Text3(i).Text = ""
        Next i
        Text2(6).Text = ""
        Text2(8).Text = ""
        chkInventario.Value = 0
        
    Else
        'EL
    End If
End Sub

'DAVID
'Para poner el foco en un objeto y si da error que no se arrastre
Private Sub PonerFocoObjeto(obj As Object)
    On Error Resume Next
    obj.SetFocus
    If Err.Number <> 0 Then Err.Clear
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
        .Buttons(1).Image = 5
        .Buttons(3).Image = 6
        .Buttons(5).Image = 7
        .Buttons(7).Image = 1
  
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
    CargaColumnas CByte(Button.Tag)
    
    'Hacemos las acciones
    If Modo = 2 Then CargaDatosLW
End Sub





Private Sub CargaColumnas(OpcionList As Byte)
Dim Columnas As String
Dim Ancho As String
Dim Alinea As String
Dim Formato As String
Dim Ncol As Integer
Dim c As ColumnHeader

    Select Case OpcionList
    Case 0
        'TARIFAS
         Label2(0).Caption = "Tarifas"
        Columnas = "Tarifa|Descripcion |Tipo|Importe|"
        Ancho = "800|2900|850|1500|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|2|"
        'Formatos
        Formato = "|||" & FormatoPrecio & "|"
        Ncol = 4
    
    Case 1
        'PRECIOS ESPECIALES
        Label2(0).Caption = "Precios especiales"
        Columnas = "Cod. cli.|Nombre |Precio|"
        Ancho = "1200|3500|1300|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|1|"
        'Formatos
        Formato = "000||" & FormatoImporte & "|"
        Ncol = 3
    Case 2
        Label2(0).Caption = "Promociones"
        Columnas = "Tarifa|Descripcion|F. inicio|F. Fin| Precio|"
        Ancho = "900|2300|1100|1100|1150|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|1|"
        'Formatos
        Formato = "000||dd/mm/yyyy|dd/mm/yyyy|" & FormatoPrecio & "|"
        Ncol = 5
    Case 3
        Label2(0).Caption = "PEDIDOS"
        Columnas = "NºPed|Fecha|Cod.|Nombre|Candtidad|"
        Ancho = "1250|1100|800|2300|1000|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|1|"
        'Formatos
        Formato = "|dd/mm/yyyy|||" & FormatoImporte & "|"
        Ncol = 5
    End Select
    
    Me.FrameDisponible.visible = OpcionList = 3

    'Guardo la opcion en el tag
    lw1.Tag = OpcionList & "|" & Ncol & "|"
    
    lw1.ColumnHeaders.Clear
    
    For NumRegElim = 1 To Ncol
         Set c = lw1.ColumnHeaders.Add()
         c.Text = RecuperaValor(Columnas, CInt(NumRegElim))
         c.Width = RecuperaValor(Ancho, CInt(NumRegElim))
         c.Alignment = Val(RecuperaValor(Alinea, CInt(NumRegElim)))
         c.Tag = RecuperaValor(Formato, CInt(NumRegElim))
    Next NumRegElim
End Sub


Private Sub CargaDatosLW()
Dim c As String
Dim bs As Byte
    bs = Screen.MousePointer
    c = Me.lblIndicador.Caption
    lblIndicador.Caption = "Leyendo " & Label2(0).Caption
    lblIndicador.Refresh
    CargaDatosLW2
    Me.lblIndicador.Caption = c
    Screen.MousePointer = bs
End Sub

Private Sub CargaDatosLW2()
Dim cad As String
Dim RS As ADODB.Recordset
Dim It As ListItem
Dim ElIcono As Integer



    On Error GoTo ECargaDatosLW
    
    If Modo <> 2 Then Exit Sub
    
    For NumRegElim = 1 To Toolbar2.Buttons.Count
        If Toolbar2.Buttons(NumRegElim).Value = tbrPressed Then
            ElIcono = Toolbar2.Buttons(NumRegElim).Image
            Exit For
        End If
    Next
    
    

    
    Select Case CByte(RecuperaValor(lw1.Tag, 1))
    Case 0
        'OFERTAS
        cad = "select l.codlista,nomlista,if(opcionINC=0,""PVP"",""UPC""),precioac from slista l,starif c where c.codlista=l.codlista"

        BuscaChekc = ""
    Case 1
        'Precios especiales
        cad = "select l.codclien,nomclien,precioac from sprees l,sclien s where s.codclien=l.codclien"
        BuscaChekc = ""

        
    Case 2
        'Promociones
        cad = "select l.codlista,nomlista,fechaini,fechafin,precioac from spromo l, starif s where l.codlista=s.codlista"
        BuscaChekc = ""
   
    Case 3
        '*****************************
        'Es una funcion especial
        CargaDatosPedidos
        Exit Sub
    End Select
    
    
    'La fecha
    
    'EL where del codclien
    cad = cad & " and l.codartic='" & DevNombreSQL(Data1.Recordset!codartic) & "'"
    
    
    

    
    'El ORDER BY

    BuscaChekc = ""
    
    lw1.ListItems.Clear
    Set RS = New ADODB.Recordset
    
    RS.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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



Private Sub CargaDatosPedidos()
Dim c As String
Dim Importe As Currency
Dim T As Currency
    
    'Limpiamos
    lw1.ListItems.Clear
    For NumRegElim = 1 To 3
        Text4(NumRegElim).Text = ""
    Next
    
    'Cargamos el primer combo
    Text4(0).Text = txtSumaStock.Text
    T = 0
    If txtSumaStock.Text <> "" Then T = ImporteFormateado(txtSumaStock.Text)
        
        
    
    'Cargamos primero los de cliente
    c = "select scaped.numpedcl,fecpedcl,codclien,nomclien,sum(cantidad) as cuantos"
    c = c & " from scaped,sliped where scaped.numpedcl=sliped.numpedcl  and codartic='"
    c = c & DevNombreSQL(Data1.Recordset!codartic) & "' GROUP BY 1"
    Importe = CargaListPedidos(6, c)
    T = T - Importe
    Text4(1).Text = Format(Importe, FormatoImporte)
    
    'Cargamos los comprados
    c = "select scappr.numpedpr,fecpedpr,codprove,nomprove,sum(cantidad) as cuantos"
    c = c & " from scappr,slippr where scappr.numpedpr=slippr.numpedpr  and codartic='"
    c = c & DevNombreSQL(Data1.Recordset!codartic) & "' group by 1"
    Importe = CargaListPedidos(9, c)
    T = T + Importe
    Text4(2).Text = Format(Importe, FormatoImporte)
    'Disponible
    Text4(3).Text = Format(T, FormatoImporte)
End Sub


Private Function CargaListPedidos(ByRef ElIcono As Integer, cad As String) As Currency
Dim RS As ADODB.Recordset
Dim It As ListItem
Dim Cantidad As Currency

    Set RS = New ADODB.Recordset
    
    Cantidad = 0
    RS.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
        Cantidad = Cantidad + DBLet(RS!Cuantos, "N")
        It.SmallIcon = ElIcono
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    CargaListPedidos = Cantidad
End Function




Private Sub ponerDatosConjuntos()
Dim Im1 As Currency
Dim Im2 As Currency
Dim Aux As Single  'Cambio pq cantidad puede tener 5 decimales
Dim CantidadCOnvertida As Currency

    On Error GoTo EponerDatosConjuntos
    'Signo los valores del articulo del UPC y PVP
    txtConjunto(0).Text = Text1(15).Text
    txtConjunto(3).Text = Text1(17).Text
    
    If Data2.Recordset.RecordCount > 0 Then Me.Data2.Recordset.MoveFirst
        
        
    'Marzo 2009
    'Los importes se calculan POR COSTES
    'El PVP es el coste total (costes parciales + costes formato ) * margen
        
    'Recorrer el RS buscando los importes reales
    While Not Data2.Recordset.EOF
    '
        CantidadCOnvertida = DBLet(Data2.Recordset!FactorConversion, "N")  'del articulo de la linea
        
        'COSTE
        Aux = DBLet(Data2.Recordset!Cantidad, "N") * CantidadCOnvertida
        Aux = Aux * DBLet(Data2.Recordset!precioUC, "N")
        Im1 = Im1 + Aux
        
        Data2.Recordset.MoveNext
    Wend
    If Data2.Recordset.RecordCount > 0 Then Me.Data2.Recordset.MoveFirst
    
    
    'Añadimos los costes derivados por el tipo de formato(unidad)
    'select sum(importe) from sunilin where codunida =1                          codigounidad
    BuscaChekc = DevuelveDesdeBD(conAri, "sum(importe)", "sunilin", "codunida", Text1(5).Text)
    If BuscaChekc = "" Then BuscaChekc = "0"
    Aux = Round2(CCur(BuscaChekc), 3)
    Im1 = Round2(Im1, 3)
    BuscaChekc = "##,##0.000"
    
    
    'Muestro los valores formato
    txtConjunto(7).Text = Format(Aux, BuscaChekc)
    txtConjunto(6).Text = Format(Im1, BuscaChekc)
    'Le sumo los costes del formato, tanto a precio coste como a venta
    Im1 = Im1 + Aux
    
    
    
    Im2 = DBLet(Data1.Recordset!margecom, "N")
    Im2 = Round2((Im1 * Im2) / 100, 3)
    Im2 = Im2 + Im1
    
    txtConjunto(1).Text = Format(Im1, BuscaChekc)
    txtConjunto(4).Text = Format(Im2, BuscaChekc)
    
    
    'Difernecias
    Aux = Round2(ImporteFormateado(txtConjunto(0).Text), 3)
    Im1 = Aux - Im1
    Aux = Round2(ImporteFormateado(txtConjunto(3).Text), 3)
    Im2 = Aux - Im2
    txtConjunto(2).Text = Format(Im1, BuscaChekc)
    txtConjunto(5).Text = Format(Im2, BuscaChekc)
    
    Exit Sub
EponerDatosConjuntos:
    MuestraError Err.Number, Err.Description
End Sub

Private Function InsertarArticuloAVAB(menErr As String, EnSarti4 As String) As Boolean
Dim c As String

On Error GoTo EinsertarArticuloAVAB

    InsertarArticuloAVAB = False
    
    c = "codartic,nomartic,codigoea,codtelem,codfamia,codmarca,codunida,codtipar,codstatu,codigiva,conjunto,artvario,nseriesn,garantia,preciomp,precioma,preciouc,preciost,preciove,ultfecco,ultfecpvp,"
    c = c & "unicajas , fecaltas, textoven, textocom, controli, CtrStock, codcateg, numSerie, fecvigen, margecom, preanuman, txtauxdocumento, LitrosUnidad, FactorConversion, Trazabilidad"
    
    c = "INSERT INTO ariges" & EmprAVAB & ".sartic(codprove," & c & ")" & " Select 5," & c & " FROM sartic"
    c = c & " WHERE codartic = '" & Text1(0).Text & "'"
    Conn.Execute c
    
    
    
    'Insertamos en sarti4
    Conn.Execute EnSarti4
    
    
    InsertarArticuloAVAB = True
    
    Exit Function
EinsertarArticuloAVAB:
    menErr = "Insertando referencia en AVAB" & Err.Description

End Function


Private Sub ActualizaComponentesAVAB(SQL As String)
Dim c As String
    On Error GoTo EActualizaComponentesAVAB
    
    If EmprAVAB < 1 Then Exit Sub
    
    c = DevuelveDesdeBD(conAri, "codartic", "ariges" & EmprAVAB & ".sartic", "codartic", Text1(0).Text)
    If c = "" Then Exit Sub
    
    'Si que esta la referencia ppal dada de alta
    'Intetentamos insertar/modificar/eleiminar (CRUD) en AVAB
    c = Replace(SQL, "sarti1 ", "ariges" & EmprAVAB & ".sarti1 ")
    Conn.Execute c
    
    Exit Sub
EActualizaComponentesAVAB:
    MuestraError Err.Number, "Actualizando componentes en AVAB"
End Sub


''''0   GENERICA
''''1   ACEITE
''''2   ENVASE
''''3   TAPON
''''4   E.FRONTAL
''''5   E.DORSAL
''''6   EMBALAJE
''''7   REJILLA
''''8   RETRACTIL
Private Sub PonerDatosFichaTecnica2(Cual As Integer)
Dim RS As ADODB.Recordset
Dim Tra As Boolean
Dim i As Integer

    On Error GoTo EPonerDatosFichaTecnica
    
    'AQUI POndremos los datos de la ficha tencina
     LimpiarFichas
     Tra = Val(DBLet(Data1.Recordset!Trazabilidad, "N")) = 1 And Val(DBLet(Data1.Recordset!Conjunto, "N")) = 1
     PonerFramesFichaTecnica2 Cual, Tra
    
     Set RS = New ADODB.Recordset
     RS.Open "Select * from sarti4 where codartic = '" & Text1(0).Text & "'", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
     Me.FramePalet.Tag = Cual & "|"   'Que tipo de componente es
     If Not RS.EOF Then
            'flejado`,`pesoneto`,`pesobruto`,`pal_tipop`,`pal_udbas`,`pal_udalt`,`pal_pvaci`,`pal_marca`,
            '`tap_aplic`,`tap_medid`,`tap_faldo`,`tap_serig`,`tap_seriT`,`tap_color`,
            '`eti_imean`,`eti_reimp`,`eti_texto`,
            '`caj_codun`,`caj_medid`,`caj_volum`,`caj_unida`,`caj_marca`,`caj_sella`,ret_serig ret_seriT ret_medid
            'CAmpos COMUNES
            'PEsos y flejado
            Text5(0).Text = MiDBLet(RS!pesoneto, "D")
            Text5(1).Text = MiDBLet(RS!pesobruto, "D")
            Text5(2).Text = MiDBLet(RS!flejado, "T")
            
            Me.FramePalet.Tag = Me.FramePalet.Tag & Text5(0).Text & "|"
            
            'Si tiene trazabilidad mostramos la paletizacion
            If Tra Then
                'Paletizacion
                '`pal_tipop`,`pal_udbas`,`pal_udalt`,`pal_bruto`,`pal_pneto`,`pal_marca`
                Text5(3).Text = MiDBLet(RS!pal_tipop, "T")
                Text5(4).Text = MiDBLet(RS!pal_udbas, "N")
                Text5(5).Text = MiDBLet(RS!pal_udalt, "N")
                Text5(6).Text = MiDBLet(RS!pal_marca, "T")
                Text5(7).Text = MiDBLet(RS!pal_pvaci, "D")
                Text5(8).Text = MiDBLet(RS!pal_pneto, "D")
                Text5(9).Text = MiDBLet(RS!pal_pbruto, "D")
                Text5(16).Text = MiDBLet(RS!caj_codun, "T")
                
                'Alt x ancho
                Me.FramePalet.Tag = Me.FramePalet.Tag & Text5(4).Text & "|"
                Me.FramePalet.Tag = Me.FramePalet.Tag & Text5(5).Text & "|"
                
            End If
            
            
            'TAPONES
            If Cual = 3 Then
                'tap_aplic`,`tap_medid`,`tap_faldo`,`tap_serig`,`tap_seriT`,`tap_color`
                Text5(11).Text = MiDBLet(RS!tap_aplic, "T")
                Text5(12).Text = MiDBLet(RS!tap_medid, "T")
                Text5(13).Text = MiDBLet(RS!tap_faldo, "T")
                Text5(14).Text = MiDBLet(RS!tap_color, "T")
                Me.chkFichTec(2).Value = MiDBLet(RS!tap_serig, "C")
                Text5(15).Text = MiDBLet(RS!tap_seriT, "T")
            End If
            'ETIQUETAS
            If Cual = 4 Or Cual = 5 Then
                '`eti_imean`,`eti_reimp`,`eti_texto`
                Me.chkFichTec(0).Value = MiDBLet(RS!eti_imean, "N")
                Me.chkFichTec(1).Value = MiDBLet(RS!eti_reimp, "N")
                Text5(10).Text = MiDBLet(RS!eti_texto, "T")
            End If
            
            
                        
            'EMBALAJE
            If Cual = 6 Then
                '`caj_codun`,`caj_medid`,`caj_volum`,`caj_unida`,caj_vacia,`caj_marca`,`caj_sella`)
               
                Text5(17).Text = MiDBLet(RS!caj_medid, "T")
                Text5(18).Text = MiDBLet(RS!caj_volum, "T")
                Text5(20).Text = MiDBLet(RS!caj_unida, "N")
                Text5(19).Text = MiDBLet(RS!caj_vacia, "D")
                Text5(21).Text = MiDBLet(RS!caj_marca, "T")
                Me.chkFichTec(4).Value = DBLet(RS!caj_sella, "N")
                If Me.chkFichTec(4).Value = 1 Then
                    Me.chkFichTec(3).Value = 0
                Else
                    Me.chkFichTec(3).Value = 1
                End If
                Me.FramePalet.Tag = Me.FramePalet.Tag & Text5(19).Text & "|"
                
            End If
            If Cual = 7 Then
                Text5(24).Text = MiDBLet(RS!caj_medid, "T")
            End If
                
            'RETRACTIL
            If Cual = 8 Then
                'ret_serig ret_seriT ret_medid
                Text5(23).Text = MiDBLet(RS!ret_medid, "T")
                Me.chkFichTec(5).Value = DBLet(RS!ret_serig, "N")
                Text5(22).Text = MiDBLet(RS!ret_seriT, "T")
            End If
            
            
            
     End If
     RS.Close
     
EPonerDatosFichaTecnica:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set RS = Nothing
End Sub
'TIPo:  Texto  N:numero    C:check
Private Function MiDBLet(ByRef campo As Field, Optional Tipo As String) As Variant
    If IsNull(campo) Then
        If Tipo = "C" Then
            MiDBLet = 0
        Else
            MiDBLet = ""
        End If
    Else
        MiDBLet = campo
        'AÑADIREMOS  N entero,D decimal
        If Tipo = "D" Then MiDBLet = Format(MiDBLet, FormatoPrecio)
    End If
End Function


Private Sub LimpiarFichas()
Dim T As TextBox
Dim J As Integer
    For Each T In Text5
        T.Text = ""
    Next T
    For J = 0 To 5
        Me.chkFichTec(J).Value = 0
    Next
End Sub




''''0   GENERICA
''''1   ACEITE
''''2   ENVASE
''''3   TAPON
''''4   E.FRONTAL
''''5   E.DORSAL
''''6   EMBALAJE
''''7   REJILLA
''''8   RETRACTIL
Private Sub PonerFramesFichaTecnica2(Cual As Integer, Trazabilidad As Boolean)


    FramePalet.visible = Trazabilidad
    'Me.Frame2.visible = Cual = 2
    
    Me.FrFichaTec(3).visible = Cual = 3
    Me.FrFichaTec(4).visible = Cual = 4 Or Cual = 5
    Me.FrFichaTec(6).visible = Cual = 6
    Me.FrFichaTec(7).visible = Cual = 7
    Me.FrFichaTec(8).visible = Cual = 8
End Sub




Private Sub PonerModoText5(Bloquear As Boolean)
Dim T As TextBox
Dim i As Integer

    For Each T In Text5
        BloquearTxt T, Bloquear
    Next T
    For i = 0 To chkFichTec.Count - 1
        chkFichTec(i).Enabled = Not Bloquear
    Next
End Sub


Private Function InsertarEnFichaTecnica(cadErr As String, CadenaSarti4 As String) As Boolean
Dim SQL As String
Dim Aux As String
Dim V As Integer
On Error GoTo EInsertarEnFichaTecnia
    InsertarEnFichaTecnica = False
    
    
    'Vamos a comprobar
    If Me.FramePalet.visible Then
        If Text5(4).Text = "" Then
            'Veremos las uds palet desde otra referencia para el mismo modelo formato
            Aux = Text1(6).Text & Text1(5).Text
            Aux = " codartic like '%" & Aux & "' AND pal_udbas>0 AND 1"
            SQL = "pal_udalt"
            Aux = DevuelveDesdeBD(conAri, "pal_udbas", "sarti4", Aux, "1", "N", SQL)
            If Aux <> "" Then
                Text5(4).Text = Aux
                Text5(5).Text = SQL
            End If
            
            Text5(7).Text = "22"
            
        End If
    End If
    
    
   ' insert into `sarti4` (`codartic`,`flejado`,`pesoneto`,`pesobruto`,
   '`pal_tipop`,`pal_udbas`,`pal_udalt`,`pal_pvaci`,`pal_pneto`,`pal_pbruto`,`pal_marca`, `caj_codun`
   '`tap_aplic`,`tap_medid`,`tap_faldo`,`tap_serig`,`tap_seriT`,`tap_color`,
   '`eti_imean`,`eti_reimp`,`eti_texto`,
   ',`caj_medid`,`caj_volum`,`caj_unida`,`caj_vacia`,`caj_marca`,`caj_sella`,
   '`ret_medid`,`ret_serig`,`ret_seriT`

    SQL = DBSet(Text1(0).Text, "T") & ","
    

    SQL = SQL & DBSet(Text5(2).Text, "T", "S") & "," & DBSet(Text5(0).Text, "N", "S") & "," & DBSet(Text5(1).Text, "N", "S") & ","
    'Palet
    '`pal_tipop`,`pal_udbas`,`pal_udalt`,`pal_pvaci`,`pal_marca`,pal_pneto pal_pbruto
    If Me.FramePalet.visible Then
        SQL = SQL & DBSet(Text5(3).Text, "T", "S") & "," & DBSet(Text5(4).Text, "N", "S") & "," & DBSet(Text5(5).Text, "N", "S") & ","
        SQL = SQL & DBSet(Text5(7).Text, "N", "S") & "," & DBSet(Text5(6).Text, "T", "S") & ","
        SQL = SQL & DBSet(Text5(8).Text, "N", "S") & "," & DBSet(Text5(9).Text, "N", "S") & ","
        SQL = SQL & DBSet(Text5(16).Text, "T", "S") & ","
    Else
        SQL = SQL & "NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,"
    End If
    
    'tapon  '`tap_aplic`,`tap_medid`,`tap_faldo`,`tap_serig`,`tap_seriT`,`tap_color`,
    If Me.FrFichaTec(3).visible Then
        SQL = SQL & DBSet(Text5(11).Text, "T", "S") & "," & DBSet(Text5(12).Text, "T", "S") & "," & DBSet(Text5(13).Text, "T", "S") & ","
        V = Val(chkFichTec(2).Value)
        Aux = "NULL"
        If V = 1 Then Aux = DBSet(Text5(15).Text, "T", "S")
        SQL = SQL & V & "," & Aux & "," & DBSet(Text5(14).Text, "T", "S") & ","

    Else
        SQL = SQL & "NULL,NULL,NULL,NULL,NULL,NULL,"
    End If
    
    
    If Me.FrFichaTec(4).visible Then
        '`eti_imean`,`eti_reimp`,`eti_texto`
        V = Val(Me.chkFichTec(0).Value)
        SQL = SQL & V & ","
        'Si reimprime etiq
        V = Val(chkFichTec(1).Value)
        Aux = "NULL"
        If V = 1 Then Aux = DBSet(Text5(10).Text, "T", "S")
        SQL = SQL & V & "," & Aux & ","
    Else
        SQL = SQL & "NULL,NULL,NULL,"
    End If

    If Me.FrFichaTec(6).visible Then
        '`caj_medid`,`caj_volum`,`caj_unida`,caj_vacia,`caj_marca`,`caj_sella`)
        SQL = SQL & DBSet(Text5(17).Text, "T", "S") & "," & DBSet(Text5(18).Text, "T", "S") & ","
        V = Val(chkFichTec(4).Value)
        SQL = SQL & DBSet(Text5(20).Text, "N", "S") & "," & DBSet(Text5(19).Text, "T", "S") & "," & DBSet(Text5(21).Text, "N", "S") & "," & V & ","
    Else
        'Caja medida lleva la de la rejilla
        If Me.FrFichaTec(7).visible Then
            Aux = DBSet(Text5(24).Text, "T", "S")
        Else
            Aux = "NULL"
        End If
            
        SQL = SQL & Aux & ",NULL,NULL,NULL,NUll,NULL,"
    End If
    
    If Me.FrFichaTec(8).visible Then
            'ret_medid ret_serig ret_seriT
            V = Val(chkFichTec(5).Value)
            Aux = "NULL"
            If V = 1 Then Aux = DBSet(Text5(22).Text, "T", "S")
            
            SQL = SQL & DBSet(Text5(23).Text, "T", "S") & "," & V & "," & Aux
    Else
        SQL = SQL & "NULL,NULL,NULL"
    End If
    
    SQL = " VALUES (" & SQL & ")"
    SQL = "INSERT INTO sarti4 (`codartic`,`flejado`,`pesoneto`,`pesobruto`,`pal_tipop`,`pal_udbas`,`pal_udalt`,`pal_pvaci`,`pal_pneto`,`pal_pbruto`,`pal_marca`,`caj_codun`,`tap_aplic`,`tap_medid`,`tap_faldo`,`tap_serig`,`tap_seriT`,`tap_color`,`eti_imean`,`eti_reimp`,`eti_texto`,`caj_medid`,`caj_volum`,`caj_unida`,`caj_vacia`,`caj_marca`,`caj_sella`,`ret_medid`,`ret_serig`,`ret_seriT`) " & SQL
    Conn.Execute SQL
    
    If Not vParamAplic.EsAVAB Then
        If EmprAVAB > 0 Then
            Aux = Replace(SQL, "INSERT INTO sarti4 (", "INSERT INTO ariges" & EmprAVAB & ".sarti4 (")
            'EjecutaSQL conAri, Aux, True
            CadenaSarti4 = Aux
        End If
    End If
    
    
    InsertarEnFichaTecnica = True
EInsertarEnFichaTecnia:
    If Err.Number <> 0 Then cadErr = "Insertando ficha tecnica. " & Err.Description

End Function



Private Sub ModificarEnFichaTecnica()
Dim SQL As String
Dim Aux As String
Dim V As Integer
Dim C2 As Currency
Dim RecalcularPesos As Boolean

On Error GoTo EInsertarEnFichaTecnia
    
    RecalcularPesos = False
    
    
    
    
    'Si ha cambiado el tipo de articulo
    V = cboTipoArt.ItemData(cboTipoArt.ListIndex)
    Aux = RecuperaValor(Me.FramePalet.Tag, 1)
    If Val(Aux) <> V Then RecalcularPesos = True
    
    'insert into `sarti4` (`codartic`,`flejado`,`pesoneto`,`pesobruto`,
    '`pal_tipop`,`pal_udbas`,`pal_udalt`,`pal_pvaci`,`pal_marca`,pal_pneto,pal_pbruto
    '`tap_aplic`,`tap_medid`,`tap_faldo`,`tap_serig`,`tap_seriT`,`tap_color`,
    '`eti_imean`,`eti_reimp`,`eti_texto`,
    '`caj_codun`,`caj_medid`,`caj_volum`,`caj_unida`,caj_vacia,`caj_marca`,`caj_sella`,
    '`ret_medid`,`ret_serig`,`ret_seriT`) values

    SQL = "UPDATE sarti4 SET flejado="
    SQL = SQL & DBSet(Text5(2).Text, "T", "S") & ",pesoneto=" & DBSet(Text5(0).Text, "N", "S") & ",pesobruto=" & DBSet(Text5(1).Text, "N", "S")
    
    Aux = RecuperaValor(Me.FramePalet.Tag, 2)  'peso neto
    If Aux <> Text5(0).Text Then RecalcularPesos = True
    
    'Palet
    '`pal_tipop`,`pal_udbas`,`pal_udalt`,`pal_pvaci`,`pal_marca`,
    If Me.FramePalet.visible Then
        'Compruebo si hay que recalcular pesos-
        SQL = SQL & ",pal_tipop= " & DBSet(Text5(3).Text, "T", "S") & ",pal_udbas = " & DBSet(Text5(4).Text, "N", "S") & ",pal_udalt=" & DBSet(Text5(5).Text, "N", "S")
        SQL = SQL & ",pal_pvaci= " & DBSet(Text5(7).Text, "N", "S") & ",pal_marca= " & DBSet(Text5(6).Text, "T", "S")
        SQL = SQL & ",pal_pneto= " & DBSet(Text5(8).Text, "N", "S") & ",pal_pbruto= " & DBSet(Text5(9).Text, "N", "S")
        SQL = SQL & ",caj_codun= " & DBSet(Text5(16).Text, "T", "S")
        
        If Not RecalcularPesos Then
            Aux = RecuperaValor(Me.FramePalet.Tag, 3)  'peso neto
            If Aux <> Text5(4).Text Then
                RecalcularPesos = True
            Else
                Aux = RecuperaValor(Me.FramePalet.Tag, 4)  'peso neto
                If Aux <> Text5(5).Text Then RecalcularPesos = True
            End If
        End If
            
        
    End If
    
    'tapon  '`tap_aplic`,`tap_medid`,`tap_faldo`,`tap_serig`,`tap_seriT`,`tap_color`,
    If Me.FrFichaTec(3).visible Then
        SQL = SQL & ",tap_aplic=" & DBSet(Text5(11).Text, "T", "S") & ",tap_medid = " & DBSet(Text5(12).Text, "T", "S") & ","
        SQL = SQL & "tap_faldo=" & DBSet(Text5(13).Text, "T", "S")
        
        V = Val(chkFichTec(2).Value)
        Aux = "NULL"
        If V = 1 Then Aux = DBSet(Text5(15).Text, "T", "S")
        SQL = SQL & ",tap_serig =" & V & ",tap_seriT=" & Aux & ",tap_color=" & DBSet(Text5(14).Text, "T")

    End If
    
    
    If Me.FrFichaTec(4).visible Then
        '`eti_imean`,`eti_reimp`,`eti_texto`
        V = Val(Me.chkFichTec(0).Value)
        SQL = SQL & ", eti_imean = " & V
        'Si reimprime etiq
        V = Val(chkFichTec(1).Value)
        Aux = "NULL"
        If V = 1 Then Aux = DBSet(Text5(10).Text, "T", "S")
        SQL = SQL & ",eti_reimp = " & V & ",eti_texto = " & Aux

    End If

    If Me.FrFichaTec(6).visible Then
        '`caj_codun`,`caj_medid`,`caj_volum`,`caj_unida`,caj_vacia,`caj_marca`,`caj_sella`)
     
        SQL = SQL & ",caj_medid = " & DBSet(Text5(17).Text, "T", "S") & ",caj_volum = " & DBSet(Text5(18).Text, "T", "S")
        V = Val(chkFichTec(4).Value)
        SQL = SQL & ", caj_unida= " & DBSet(Text5(20).Text, "N", "S") & ",caj_vacia =" & DBSet(Text5(19).Text, "N", "S") & ",caj_marca = " & DBSet(Text5(21).Text, "T", "S") & ",caj_sella= " & V

    End If
    
    'REJILLA
    If Me.FrFichaTec(7).visible Then
        SQL = SQL & ",caj_medid = " & DBSet(Text5(24).Text, "T", "S")
    End If
    
    
    If Me.FrFichaTec(8).visible Then
            'ret_medid ret_serig ret_seriT
            V = Val(chkFichTec(5).Value)
            Aux = "NULL"
            If V = 1 Then Aux = DBSet(Text5(22).Text, "T", "S")
            
            SQL = SQL & ",ret_medid = " & DBSet(Text5(23).Text, "T", "S") & ",ret_serig = " & V & ",ret_seriT = " & Aux
    End If
    SQL = SQL & " WHERE codartic = '" & Text1(0).Text & "'"
    Conn.Execute SQL
    If Not vParamAplic.EsAVAB Then
        If EmprAVAB > 0 Then
              Aux = Replace(SQL, "UPDATE sarti4 SET", "UPDATE ariges" & EmprAVAB & ".sarti4 SET")
              EjecutaSQL conAri, Aux, True
              
        End If
    End If
    If RecalcularPesos Then
            Aux = Me.lblIndicador.Caption
            lblIndicador.Caption = "Recalculando pesos"
            lblIndicador.Refresh
            Espera 0.4
            
            If FramePalet.visible Then
                'Es producto venta. Solo recalculo a el mismo
                V = DBLet(ComprobarCero(Text5(4).Text), "N") * DBLet(ComprobarCero(Text5(5).Text), "N")
                C2 = DBLet(ComprobarCero(Text5(0).Text), "N")
                RecalcularPesoArticulo CStr(Data1.Recordset!codartic), DBLet(Data1.Recordset!UniCajas, "N"), V, C2, False
            Else
                RecalcularPesosArtVenta Text1(0).Text
            
            
            End If
                'Es componente. Vere a cuantos tengo que recalcular
'               C1 = Val(ComprobarCero(Text5(4).Text))
'            c2 = Val(ComprobarCero(Text5(5).Text))
'            C1 = C1 * c2
'            'El peso del aceite
'            cad = ComprobarCero(Text5(0).Text)
'            c2 = CCur(cad)
'            RecalcularPesoArticulo CStr(Data1.Recordset!codArtic), DBLet(Data1.Recordset!UniCajas, "N"), C1, c2
            
            lblIndicador.Caption = Aux
            lblIndicador.Refresh
            
    End If
    
    
    
    
    
    
    Exit Sub
EInsertarEnFichaTecnia:
    MsgBox "Error modificando la FICHA TECNICA del articulo" & vbCrLf & Err.Description, vbExclamation

End Sub


Private Sub RecalcularPesosArtVenta(Arti1 As String)
Dim c As String
Dim R As ADODB.Recordset
Dim CajasPalet As Integer
    On Error GoTo ERecalcularPesosArtVenta
    
    c = "Select sartic.codartic,unicajas,pesoneto,pal_udbas,pal_udalt from sarti1,sartic,sarti4 where "
    c = c & "sarti1.codarti1='" & Arti1 & "' and sarti1.codartic =sartic.codartic and "
    c = c & "sarti4.codartic= sartic.codartic "
    Set R = New ADODB.Recordset
    R.Open c, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not R.EOF
        'pal_udbas`,`pal_udalt`
        CajasPalet = CInt(DBLet(R!pal_udbas, "N")) * CInt(DBLet(R!pal_udalt, "N"))
        RecalcularPesoArticulo CStr(R!codartic), CInt(DBLet(R!UniCajas, "N")), CajasPalet, CCur(DBLet(R!pesoneto, "N")), False
        R.MoveNext
    Wend
    R.Close
    
    
    
    Exit Sub
ERecalcularPesosArtVenta:
    MuestraError Err.Description, "RecalcularPesosArtVenta"
End Sub



Private Sub ImprimirFT()
    If vParamAplic.EsAVAB Then Exit Sub
    If Me.chkConjunto.Value = 0 Then Exit Sub
    frmFichaTecnicaImp.vCodartic = Text1(0).Text
    frmFichaTecnicaImp.Show vbModal
End Sub
