VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMensajes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mensajes"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14160
   Icon            =   "frmMensajes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   14160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameVerDatosStockProduccion 
      Height          =   5775
      Left            =   3360
      TabIndex        =   50
      Top             =   960
      Visible         =   0   'False
      Width           =   9855
      Begin VB.CommandButton cmdStockProd 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   8400
         TabIndex        =   54
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CommandButton cmdStockProd 
         Caption         =   "&Continuar"
         CausesValidation=   0   'False
         Height          =   375
         Index           =   0
         Left            =   7080
         TabIndex        =   53
         Top             =   5160
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView5 
         Height          =   3975
         Left            =   240
         TabIndex        =   52
         Top             =   960
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   7011
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Articulo"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Resto"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Ud linea"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Stock"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock produccion"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   915
         Left            =   240
         TabIndex        =   51
         Top             =   240
         Width           =   6015
      End
   End
   Begin VB.Frame FrameAcercaDe 
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6375
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pasaje Ventura Feliu 13, Entlo 2� Izda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   2760
         TabIndex        =   10
         Top             =   2640
         Width           =   3240
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "46007 - VALENCIA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   3000
         TabIndex        =   9
         Top             =   2925
         Width           =   2535
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tfno:  902 88 88 78  -  96 380 55 79"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   2685
         TabIndex        =   8
         Top             =   3195
         Width           =   3165
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax: 963 42 09 38"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   3480
         TabIndex        =   7
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   2160
         Picture         =   "frmMensajes.frx":000C
         Top             =   2640
         Width           =   480
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
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
         Left            =   480
         TabIndex        =   6
         Top             =   1200
         Width           =   4155
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Gesti�n-Producci�n"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   915
         Left            =   0
         TabIndex        =   4
         Top             =   480
         Width           =   6015
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   81.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1725
         Left            =   4920
         TabIndex        =   5
         Top             =   240
         Width           =   1350
      End
   End
   Begin VB.Frame FrameCorreccionPrecios 
      Height          =   6375
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   12975
      Begin VB.ComboBox cmbActualizarTar 
         Height          =   315
         ItemData        =   "frmMensajes.frx":0CD6
         Left            =   7800
         List            =   "frmMensajes.frx":0CD8
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   5960
         Width           =   2175
      End
      Begin VB.CommandButton cmdCorrecotrPrecios 
         Caption         =   "Salir"
         Height          =   375
         Index           =   1
         Left            =   11760
         TabIndex        =   38
         Top             =   5880
         Width           =   975
      End
      Begin VB.CommandButton cmdCorrecotrPrecios 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   10560
         TabIndex        =   37
         Top             =   5880
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   5175
         Left            =   120
         TabIndex        =   36
         Top             =   600
         Width           =   12660
         _ExtentX        =   22331
         _ExtentY        =   9128
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   3246
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Denominaci�n"
            Object.Width           =   5715
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "U.P.Compra"
            Object.Width           =   2188
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "% M"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "PVP"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "%T"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "P.Tarifa"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "PVP Correcto"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Tarifa correc."
            Object.Width           =   2011
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Actualizar:"
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
         Height          =   255
         Index           =   1
         Left            =   6480
         TabIndex        =   42
         Top             =   6000
         Width           =   1215
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   2
         Left            =   10440
         Picture         =   "frmMensajes.frx":0CDA
         Top             =   240
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   3
         Left            =   10800
         Picture         =   "frmMensajes.frx":0E24
         Top             =   240
         Width           =   240
      End
      Begin VB.Label lblIndicadorCorregir 
         Caption         =   "Label3"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   5880
         Width           =   5055
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   8175
      End
   End
   Begin VB.Frame FrameErrores 
      Height          =   5535
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   8415
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "&Cerrar"
         Height          =   375
         Left            =   7080
         TabIndex        =   29
         Top             =   4920
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   4335
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   28
         Text            =   "frmMensajes.frx":0F6E
         Top             =   360
         Width           =   7695
      End
   End
   Begin VB.Frame FrameCobrosPtes 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      Begin VB.CommandButton cmdCancelarCobros 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   7080
         TabIndex        =   25
         Top             =   4440
         Width           =   975
      End
      Begin VB.TextBox txtParam 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   23
         Text            =   "frmMensajes.frx":0F74
         Top             =   240
         Width           =   6615
      End
      Begin VB.CommandButton cmdAceptarCobros 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5400
         TabIndex        =   1
         Top             =   4440
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3135
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   5530
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "�Desea continuar?"
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   26
         Top             =   4440
         Width           =   7215
      End
      Begin VB.Label Label1 
         Caption         =   "Departamento:"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   24
         Top             =   240
         Width           =   7215
      End
   End
   Begin VB.Frame FrameNSeries 
      Height          =   5000
      Left            =   480
      TabIndex        =   11
      Top             =   120
      Width           =   6975
      Begin VB.CommandButton cmdSelTodos 
         Caption         =   "&Todos"
         Height          =   315
         Left            =   720
         TabIndex        =   22
         Top             =   4320
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.CommandButton cmdDeselTodos 
         Caption         =   "&Ninguno"
         Height          =   315
         Left            =   1680
         TabIndex        =   21
         Top             =   4320
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5040
         TabIndex        =   20
         Top             =   4320
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarNSeries 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3960
         TabIndex        =   12
         Top             =   4320
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3015
         Left            =   720
         TabIndex        =   13
         Top             =   840
         Width           =   4020
         _ExtentX        =   7091
         _ExtentY        =   5318
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label7 
         Caption         =   "Empresas en el sistema"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Index           =   1
         Left            =   720
         TabIndex        =   30
         Top             =   240
         Visible         =   0   'False
         Width           =   5295
      End
   End
   Begin VB.Frame FrameComponentes 
      Height          =   3975
      Left            =   0
      TabIndex        =   14
      Top             =   120
      Width           =   5295
      Begin VB.CommandButton cmdAceptarComp 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3240
         TabIndex        =   19
         Top             =   3240
         Width           =   975
      End
      Begin VB.Frame FrameComponentes2 
         Caption         =   "Mostrar Equipos del :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   2175
         Left            =   960
         TabIndex        =   15
         Top             =   720
         Width           =   3255
         Begin VB.OptionButton OptCompXClien 
            Caption         =   "Cliente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   18
            Top             =   1440
            Width           =   2535
         End
         Begin VB.OptionButton OptCompXDpto 
            Caption         =   "Departamento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   17
            Top             =   960
            Width           =   2535
         End
         Begin VB.OptionButton OptCompXMant 
            Caption         =   "Mantenimiento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   16
            Top             =   480
            Width           =   2055
         End
      End
   End
   Begin VB.Frame FrameTraspasoMante 
      Height          =   3135
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Width           =   5415
      Begin VB.TextBox txtMante 
         Height          =   285
         Index           =   0
         Left            =   2640
         TabIndex        =   48
         Top             =   1080
         Width           =   855
      End
      Begin VB.CheckBox chkMante 
         Caption         =   "Copiar importes en siguiente"
         Height          =   255
         Left            =   360
         TabIndex        =   46
         Top             =   1800
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CommandButton cmdMante 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   3960
         TabIndex        =   45
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CommandButton cmdMante 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   2520
         TabIndex        =   44
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "A�o a traspasar"
         Height          =   195
         Left            =   1320
         TabIndex        =   49
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Pasar importes mantenimiento a historico."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   47
         Top             =   360
         Width           =   4695
      End
   End
   Begin VB.Frame FrameEtiqEstant 
      Height          =   7455
      Left            =   0
      TabIndex        =   31
      Top             =   -120
      Width           =   8535
      Begin VB.CommandButton cmdEtiqEstan 
         Caption         =   "Imprimir"
         Height          =   375
         Index           =   1
         Left            =   5520
         TabIndex        =   34
         Top             =   6960
         Width           =   1215
      End
      Begin VB.CommandButton cmdEtiqEstan 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   0
         Left            =   6960
         TabIndex        =   33
         Top             =   6960
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   6495
         Left            =   240
         TabIndex        =   32
         Top             =   360
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   11456
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Descripcion"
            Object.Width           =   6703
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Precio"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Familia"
            Object.Width           =   4410
         EndProperty
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   600
         Picture         =   "frmMensajes.frx":0F7A
         Top             =   6960
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   240
         Picture         =   "frmMensajes.frx":10C4
         Top             =   6960
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmMensajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'====================== VBLES PUBLICAS ================================

Public Event DatoSeleccionado(CadenaSeleccion As String)

Public OpcionMensaje As Byte
'======================================
'==== FACTURACION =====================
' 1 .- Mensaje de Cobros Pendientes
' 2 .- Mensaje de No hay suficiente Stock para pasar de Pedido a Albaran
' 3 .- Mensaje Acerca de...
' 4 .- Listado de los N� de Serie de un Articulo
' 5 .- Seleccionar tipo de Componente a Mostrar en Mant. de N� de Series
' 6 .- Mostrar Prefacturacion de Albaranes
' 7 .- Mostrar Prefacturacion Mantenimientos
' 8 .- Mostrar lista clientes para seleccionar los que queremos imprimir (Etiquetas)
' 9 .- Mostrar lista Proveedores para seleccionar los que queremos imprimir (Etiquetas)
'10 .- Mostrar lista de Errores de las facturas NO contabilizadas
'11 .- Mostrar lista lineas de factura a Rectificar para seleccionar las q queremos traer al Albaran de FAct. Rectificativa
'12 .- Mostrar Albaranes del Rango que no se van a Facturar. (Facturar Albaranes Venta)

'13 .- Mostrar Errores
'14 .- Mostrar Empresas existentes en el sistema



'15 .- Mostrar lista de articulos para imprimir etiquetas estanteria
'16 .- Lista de articulos para corregir importes
'17 .- Etiquetas clientes. LO MISMO QUE EL 8 pero hecho por david
'18 .- Mantenimientos. paso ejercicio siguiente a actual

'20 .- IGual que el 16. Pero los importes son de los articulos que tienen componentes


'25 .- Correcion de precios de AVAB compandose con los de Morales
'26 .- MOxiente. Falta stock componentes


Public cadWhere As String 'Cadena para pasarle la WHERE de la SELECT de los cobros pendientes o de Pedido(para comp. stock)
                          'o CodArtic para seleccionar los N� Series
                          'para cargar el ListView
                          
Public cadWhere2 As String

Public vCampos As String 'Articulo y cantidad Empipados para N� de Series
                         'Tambien para pasar el nombre de la tabla de lineas (sliped, slirep,...)
                         'Dependiendo desde donde llamemos, de Pedidos o Reparaciones


'====================== VBLES LOCALES ================================

Dim PulsadoSalir As Boolean 'Solo salir con el boton de Salir no con aspa del form
Dim PrimeraVez As Boolean

'Para los N� de Serie
Dim TotalArray As Integer
Dim codartic() As String
Dim Cantidad() As Integer



Private Sub cmdAceptarCobros_Click()
    If OpcionMensaje = 12 Then vCampos = "1"
    Unload Me
End Sub


Private Sub cmdAceptarComp_Click()
'Boton Aceptar de Componentes del Mant. de N� de Series en Reparaciones
Dim H As Integer, W As Integer

    ponerFrameComponentesVisible False, H, W
    PonerFrameCobrosPtesVisible True, H, W
    Me.Height = H + 350
    Me.Width = W + 70
    
    If Me.OptCompXMant.Value Then
        'Mostrar Resumen de los N� de Serie del Mantenimiento
        Me.Caption = "Equipos del Mantenimiento"
        CargarListaComponentes (1)
    ElseIf Me.OptCompXDpto.Value Then
        'Mostrar Resumen de los N� de Serie del Departamento
        Me.Caption = "Equipos del Departamento"
        CargarListaComponentes (2)
    ElseIf Me.OptCompXClien.Value Then
        'Mostrar Resumen de los N� de Serie del Cliente
        Me.Caption = "Equipos del Cliente"
        CargarListaComponentes (3)
    End If
    PonerFocoBtn Me.cmdAceptarCobros
End Sub


Private Sub cmdAceptarNSeries_Click()
Dim I As Byte, J As Byte
Dim Seleccionados As Integer
Dim Cad As String, SQL As String
Dim Articulo As String
Dim RS As ADODB.Recordset
Dim C1 As String * 10, C2 As String * 10, C3 As String * 10


    If OpcionMensaje = 4 Then
        'Comprobar que se han seleccionado el n� correcto de  N� de Serie para cada Articulo
        Seleccionados = 0
        Articulo = ""
      
        'Si se ha seleccionado la cantidad correcta de N� de series, empiparlos y
        'devolverlos al form de Albaranes(facturacion)
        Cad = ""
        For J = 0 To TotalArray
            Articulo = codartic(J)
            Cad = Cad & Articulo & "|"
            For I = 1 To ListView2.ListItems.Count
                If ListView2.ListItems(I).Checked Then
                    If Articulo = ListView2.ListItems(I).ListSubItems(1).Text Then
                        If Seleccionados < Abs(Cantidad(J)) Then
                            Seleccionados = Seleccionados + 1
                            Cad = Cad & ListView2.ListItems(I).Text & "|"
                        End If
                   'cad = cad & Data1.Recordset.Fields(1) & "|"
                    End If
                End If
            Next I
            If Seleccionados < Abs(Cantidad(J)) Then
                'Comprobar que si tiene N�s de serie de ese articulos cargados seleccione los
                'que corresponden
                SQL = "SELECT count(sserie.numserie)"
                SQL = SQL & " FROM sserie " 'INNER JOIN sartic ON sserie.codartic=sartic.codartic "
                SQL = SQL & " WHERE sserie.codartic=" & DBSet(Articulo, "T")
                SQL = SQL & " AND (isnull(sserie.numfactu) or sserie.numfactu='') and (isnull(sserie.numalbar) or sserie.numalbar='') "
                SQL = SQL & " ORDER BY sserie.codartic, numserie "
                Set RS = New ADODB.Recordset
                RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                If RS.Fields(0).Value >= Abs(Cantidad(J)) - Seleccionados Then
                    MsgBox "Debe seleccionar " & Cantidad(J) & " N� Series para el articulo " & codartic(J), vbExclamation
                    Exit Sub
                Else
                    'No hay N� Serie y Pedirlos
                End If
                RS.Close
                Set RS = Nothing
            
            End If
            Cad = Cad & "�"
            Seleccionados = 0
        Next J
      
    ElseIf OpcionMensaje = 8 Or OpcionMensaje = 9 Or OpcionMensaje = 17 Then
        'concatenar todos los clientes seleccionados para imprimir etiquetas
        If OpcionMensaje = 17 Then
            
            '----------------------------------------------------------------
            Cad = "insert into tmpnlotes (codusu,numalbar,fechaalb,numlinea,codprove) values ("
            Cad = Cad & vUsu.Codigo & ",1,'2005-04-12',1,"
            
            
            For I = 1 To ListView2.ListItems.Count
                If ListView2.ListItems(I).Checked Then
                    conn.Execute Cad & (ListView2.ListItems(I).Text) & ")"
                    NumRegElim = NumRegElim + 1
                End If
            Next I
            
            
            '----------------------------------------------------------------
            
        Else
            Cad = ""
            For I = 1 To ListView2.ListItems.Count
                If ListView2.ListItems(I).Checked Then
                    Cad = Cad & Val(ListView2.ListItems(I).Text) & ","
                     'cad = cad & Data1.Recordset.Fields(1) & "|"
                End If
            Next I
            If Cad <> "" Then Cad = Mid(Cad, 1, Len(Cad) - 1)
        End If
    ElseIf OpcionMensaje = 11 Then
    'Lineas Factura a rectificar
        'cad = "(" & cadWHERE & ")"
        Cad = ""
        C1 = ""
        C2 = ""
        C3 = ""
        SQL = ""
        For I = 1 To ListView2.ListItems.Count
            If ListView2.ListItems(I).Checked Then
                If SQL = "" Then
                    C1 = DBSet(ListView2.ListItems(I), "T", "N")
                    C2 = ListView2.ListItems(I).ListSubItems(1)
'                    c3 = ListView2.ListItems(i).ListSubItems(2)
                    Cad = "(codtipoa=" & Trim(C1) & " and numalbar=" & Val(C2) & " and numlinea IN (" & ListView2.ListItems(I).ListSubItems(2)

                Else
                    If Trim(DBSet(ListView2.ListItems(I), "T", "N")) = Trim(C1) And Trim(ListView2.ListItems(I).ListSubItems(1)) = Trim(C2) Then
                    'es el mismo albaran y concatenamos lineas
                        Cad = "," & ListView2.ListItems(I).ListSubItems(2)

                    Else
                        If Cad <> "" Then SQL = SQL & ")) "
                        C1 = DBSet(ListView2.ListItems(I), "T", "N")
                        C2 = ListView2.ListItems(I).ListSubItems(1)
'                    c3 = ListView2.ListItems(i).ListSubItems(2)
                        Cad = " or (codtipoa=" & Trim(C1) & " and numalbar=" & Val(C2) & " and numlinea IN (" & ListView2.ListItems(I).ListSubItems(2)
                        
'                       cad=cad &
                    End If
                End If
                SQL = SQL & Cad
'                If cad <> "" Then cad = cad & " OR "
'                cad = cad & "(codtipoa=" & DBSet(ListView2.ListItems(i), "T", "N") & " and numalbar=" & Val(ListView2.ListItems(i).ListSubItems(1)) & " and numlinea=" & ListView2.ListItems(i).ListSubItems(2) & ")"
            Else
'                cad = ""
            End If
        Next I
        If Cad <> "" Then
            SQL = SQL & "))"
            Cad = "(" & cadWhere & ") AND (" & SQL & ")"
        End If
'        If cad <> "" Then cad = "(" & cadWHERE & ") AND (" & cad & ")"
    ElseIf OpcionMensaje = 14 Then
        Cad = RegresarCargaEmpresas
    End If
    
    
    
     'Actualizar la tabla sseries asignando los valores correspondientes a los
      'campos: codclien, coddirec, tieneman, codtipom, numalbar, fechavta, numline1
      'y Salir (Volver a Mto Albaranes Clientes (Facturacion)
      PulsadoSalir = True
      'RaiseEvent CargarNumSeries
      RaiseEvent DatoSeleccionado(Cad)
      Unload Me
End Sub


Private Sub cmdCancelar_Click()
    If OpcionMensaje = 4 Then
        MsgBox "Debe introducir los n� de serie necesarios para el Albaran.", vbInformation
        Exit Sub
    End If
    PulsadoSalir = True
    RaiseEvent DatoSeleccionado("")
    Unload Me
End Sub

Private Sub cmdCancelarCobros_Click()
    vCampos = "0"
    Unload Me
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdCorrecotrPrecios_Click(Index As Integer)
    
    If Index = 0 Then
        
        If Not ActualizarPrecios Then Exit Sub
        
    End If
    Unload Me
End Sub

Private Function ActualizarPrecios() As Boolean
Dim SQL As String
    
    
    
        
        ActualizarPrecios = False
        'Compruebo si ha seleccionado algun articulo de los de precio ultima compra=0
        cadWhere2 = ""
        SQL = ""
        For TotalArray = 1 To Me.ListView4.ListItems.Count
            If ListView4.ListItems(TotalArray).Checked Then
                If ListView4.ListItems(TotalArray).Tag = "" Then
                    SQL = SQL & "M"
                Else
                    cadWhere2 = cadWhere2 & "M"
                End If
            End If
        Next
    
        If SQL <> "" Then
            MsgBox "No puede actualizar los articulos cuyo precio ultima compra sea 0", vbExclamation
            Exit Function
        End If
    
        If cadWhere2 = "" Then
            MsgBox "Seleccione algun articulo para actualizar", vbExclamation
            Exit Function
        End If
    
        'Llegado aqui todo correcto. Hacemos la pregunta de actualizar y a correr
        SQL = "art�culo"
        If Len(cadWhere2) > 1 Then SQL = SQL & "s"
        SQL = "Va a actualizar los precios de " & Len(cadWhere2) & " " & SQL & vbCrLf & vbCrLf & "�Desea continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNo) <> vbYes Then Exit Function
        
        
        'Aqui esta el proceso de actualizacion de articulos
        Me.lblIndicadorCorregir.Caption = "Actualizaci�n precios"
        Me.Refresh
        Espera 0.5
        
       'Para el LOG
       SQL = cadWhere & vbCrLf
       For TotalArray = 1 To Me.ListView4.ListItems.Count
            If ListView4.ListItems(TotalArray).Checked Then
                If ListView4.ListItems(TotalArray).Tag <> "" Then SQL = SQL & ListView4.ListItems(TotalArray).Text & "|"
            End If
        Next
        SQL = Mid(SQL, 1, 237)
        
        '------------------------------------------------------------------------------
        '  LOG de acciones
        Set LOG = New cLOG
        LOG.Insertar 4, vUsu, "Correccion precios: " & "(" & OpcionMensaje & ")" & vbCrLf & SQL
        Set LOG = Nothing
        '-----------------------------------------------------------------------------
        
        
        
        
        
        
        
        
        
        
        For TotalArray = 1 To Me.ListView4.ListItems.Count
            If ListView4.ListItems(TotalArray).Checked Then
                If ListView4.ListItems(TotalArray).Tag <> "" Then
                    
                    'lo metemos en transaccion. Si queremos vamos
                    Me.lblIndicadorCorregir.Caption = ListView4.ListItems(TotalArray).Text
                    Me.lblIndicadorCorregir.Refresh
                    
                                        
                    conn.BeginTrans
                    If ActualizaPrecios(TotalArray) Then
                        conn.CommitTrans
                    Else
                        conn.RollbackTrans
                    End If
                    
                    
                End If
            End If
        Next
    
    
        ActualizarPrecios = True
End Function


Private Function ActualizaPrecios(NumeroItem As Integer) As Boolean

On Error GoTo EActualizaPrecios
    ActualizaPrecios = False
    With ListView4.ListItems(NumeroItem)
        Select Case OpcionMensaje
        Case 16
            'ACtualizador de precio normal
            If Me.cmbActualizarTar.ListIndex <> 2 Then
                cadWhere2 = TransformaComasPuntos(CStr(ImporteFormateado(.SubItems(7))))
                cadWhere2 = "UPDATE sartic set preciove=" & cadWhere2 & " WHERE codartic = '" & ListView4.ListItems(NumeroItem).Tag & "'"
                conn.Execute cadWhere2
            End If
            If Me.cmbActualizarTar.ListIndex <> 1 Then
                cadWhere2 = TransformaComasPuntos(CStr(ImporteFormateado(.SubItems(8))))
                cadWhere2 = "UPDATE slista set precioac=" & cadWhere2 & " WHERE codartic = '" & ListView4.ListItems(NumeroItem).Tag & "' AND codlista =" & vCampos
                conn.Execute cadWhere2
            End If
            
            
        Case 25
            'DESDE EL AVAB
            '--------------------------------------------------------------

            
            vCampos = ""
            If Me.cmbActualizarTar.ListIndex <> 2 Then
                cadWhere2 = TransformaComasPuntos(CStr(ImporteFormateado(.SubItems(7))))
                vCampos = " preciove = " & cadWhere2
            End If
            If Me.cmbActualizarTar.ListIndex <> 1 Then
                cadWhere2 = TransformaComasPuntos(CStr(ImporteFormateado(.SubItems(8))))
                If vCampos <> "" Then vCampos = vCampos & ","
                vCampos = vCampos & " preciouc = " & cadWhere2
            End If
            cadWhere2 = "UPDATE sartic set " & vCampos & " WHERE codartic = '" & ListView4.ListItems(NumeroItem).Tag & "'"
            conn.Execute cadWhere2
            
            
            
            
        Case Else
            'Precio articulos componentes
            '----------------------------
            vCampos = ""
            If Me.cmbActualizarTar.ListIndex <> 2 Then
                cadWhere2 = TransformaComasPuntos(CStr(ImporteFormateado(.SubItems(7))))
                vCampos = " preciove = " & cadWhere2
            End If
            If Me.cmbActualizarTar.ListIndex <> 1 Then
                cadWhere2 = TransformaComasPuntos(CStr(ImporteFormateado(.SubItems(8))))
                If vCampos <> "" Then vCampos = vCampos & ","
                vCampos = vCampos & " preciouc = " & cadWhere2
            End If
            cadWhere2 = "UPDATE sartic set " & vCampos & " WHERE codartic = '" & ListView4.ListItems(NumeroItem).Tag & "'"
            conn.Execute cadWhere2
            
            
                        

        End Select
        
        
    End With
        
    ActualizaPrecios = True
    Exit Function
EActualizaPrecios:
    MuestraError Err.Number, ListView4.ListItems(NumeroItem).Text
End Function


Private Sub cmdDeselTodos_Click()
Dim I As Byte

    For I = 1 To ListView2.ListItems.Count
        ListView2.ListItems(I).Checked = False
    Next I
End Sub




Private Sub cmdEtiqEstan_Click(Index As Integer)
    If Index = 1 Then
        'Cargo la tabla temporal con los datos que qeuremos imprimir
        cadWhere2 = "insert into `tmpnseries` (`codusu`,`codartic`,`numlinea`,`numlinealb`) VALUES "
        cadWhere = ""
        For NumRegElim = 1 To ListView3.ListItems.Count
            '                                                En el tag YA esta grabado
            If ListView3.ListItems(NumRegElim).Checked Then
                cadWhere = cadWhere & ",(" & vUsu.Codigo & "," & ListView3.ListItems(NumRegElim).Tag & ",0)"
                If (NumRegElim Mod 25) = 0 Then
                    conn.Execute cadWhere2 & Mid(cadWhere, 2) & ";"
                    cadWhere = ""
                    DoEvents
                End If
            End If
        Next NumRegElim
        If cadWhere <> "" Then conn.Execute cadWhere2 & Mid(cadWhere, 2) & ";"
    Else
        NumRegElim = 0
    End If
    Unload Me
End Sub

Private Sub cmdMante_Click(Index As Integer)
Dim b As Boolean
    If Index = 0 Then
        
        
        If Val(txtMante(0).Text) = 0 Then
            MsgBox "El campo A�o a traspasar debe ser num�rico", vbExclamation
            Exit Sub
        End If
        
        
        If MsgBox("El proceso es irreversible. Continuar de igual modo?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        
        '-------------------------------------------
        Screen.MousePointer = vbHourglass
        Set miRsAux = New ADODB.Recordset
        conn.BeginTrans
        b = TraspasarMantenimientos
        Set miRsAux = Nothing
        Screen.MousePointer = vbDefault
        If b Then
            conn.CommitTrans
        Else
            conn.RollbackTrans
        End If
        
        
    End If
    Unload Me
End Sub

Private Sub cmdSelTodos_Click()
    Dim I As Byte

    For I = 1 To ListView2.ListItems.Count
        ListView2.ListItems(I).Checked = True
    Next I
End Sub

Private Sub cmdStockProd_Click(Index As Integer)
    If Index = 0 Then
        Set listacod = Nothing
        Set listacod = New Collection
    End If
    Unload Me
End Sub

Private Sub Form_Activate()
Dim OK As Boolean

    
    Select Case OpcionMensaje
        Case 4 'Mostrar N� Series
            If PrimeraVez Then
                PrimeraVez = False
                Me.Refresh
                Screen.MousePointer = vbHourglass
                OK = ObtenerTamanyosArray
                If OK Then OK = SeparaCampos
                If Not OK Then
                    'Error en SQL
                    'Salimos
                    Unload Me
                    Exit Sub
                End If
                CargarListaNSeries
            End If
            
        Case 8, 9, 17 'Etiquetas de clientes/Proveedores
            CargarListaClientes
'        Case 10 'Errores al contabilizar facturas
'            CargarListaErrContab
        Case 11 'Lineas Factura a rectificar
            CargarListaLinFactu
            
        Case 14 'Mostrar Empresas del sistema
            CargarListaEmpresas
            
        Case 15
            'Etiquetas estanteria
            CargarArticulosEstanteria
            
        Case 16, 20, 25
            'Articulos para corregir
            If OpcionMensaje = 16 Then
                CargarArticulosCorreccionPrecio
            ElseIf OpcionMensaje = 25 Then
                'Correcion AVAB
                CargaPVPPreciosDesdeMorales
            Else
                CargaPVPPreciosArticulosConComponentes
            End If
            If Me.ListView4.ListItems.Count = 0 Then
                MsgBox "Ning�n dato para mostrar", vbExclamation
                Unload Me
            End If
        Case 18
            PonerFoco txtMante(0)
            
        Case 26
            CargaListaProduccion
    End Select
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim Cad As String
On Error Resume Next

    Me.FrameCobrosPtes.visible = False
    Me.FrameAcercaDe.visible = False
    Me.FrameNSeries.visible = False
    Me.FrameComponentes.visible = False
    Me.FrameComponentes2.visible = False
    Me.FrameErrores.visible = False
    FrameEtiqEstant.visible = False
    FrameCorreccionPrecios.visible = False
    FrameTraspasoMante.visible = False
    FrameVerDatosStockProduccion.visible = False
    PulsadoSalir = True
    PrimeraVez = True
    
    Select Case OpcionMensaje
        Case 1 'Mensaje de Cobros Pendientes
            PonerFrameCobrosPtesVisible True, H, W
            CargarListaCobrosPtes
            Me.Caption = "Cobros Pendientes"
            PonerFocoBtn Me.cmdAceptarCobros
            
        Case 2 'Mensaje de no hay suficiente Stock
            PonerFrameCobrosPtesVisible True, H, W
            CargarListaArtSinStock (vCampos)
            Me.Caption = "Art�culos sin stock suficiente"
            PonerFocoBtn Me.cmdAceptarCobros
            
        Case 3 'Mensaje ACERCA DE
            CargaImagen
            Me.Caption = "Acerca de ....."
            PonerFrameAcercaDeVisible True, H, W
            Me.lblVersion.Caption = "Versi�n:  " & App.Major & "." & App.Minor & "." & App.Revision & " "
        
        Case 4 'Listado N� Series Articulo
            PonerFrameNSeriesVisible True, H, W
            Me.Caption = "N� Serie"
            Me.Label7(1).Caption = "Seleccione los N� de serie para el Albaran."
            Me.Label7(1).FontSize = 12
            PulsadoSalir = False
            
        Case 5 'Seleccionar tipo de Componente que queremos mostrar en Resumen
                'En mant. de N� Series de Reparacion
            ponerFrameComponentesVisible True, H, W
            Me.Caption = "Componentes"
            Me.OptCompXMant.Value = True
            PonerFocoBtn Me.cmdAceptarComp
        
        Case 6 'Mostrar Prefacturacion de Albaranes
            PonerFrameCobrosPtesVisible True, H, W
            CargarListaPreFacturar
            Me.Caption = "Prefacturaci�n Albaranes"
            Cad = RecuperaValor(vCampos, 1)
            If Cad <> "" Then Cad = Mid(Cad, 1, Len(Cad) - 1)
            Me.txtParam.Text = Cad
            Cad = RecuperaValor(vCampos, 2)
            If Cad <> "" Then
                Cad = Mid(Cad, 1, Len(Cad) - 1)
                If Trim(Me.txtParam.Text) <> "" Then
                    txtParam.Text = Me.txtParam.Text & vbCrLf & Cad
                Else
                    txtParam.Text = Cad
                End If
            End If
            Cad = RecuperaValor(vCampos, 3)
            If Cad <> "" Then
                Cad = Mid(Cad, 1, Len(Cad) - 1)
                If Trim(Me.txtParam.Text) <> "" Then
                    txtParam.Text = Me.txtParam.Text & vbCrLf & Cad
                Else
                    txtParam.Text = Cad
                End If
            End If
            
            PonerFocoBtn Me.cmdAceptarComp
            
        Case 8, 17 'Etiquetas de Clientes
            PonerFrameNSeriesVisible True, H, W
            Me.Caption = "Clientes"
            PulsadoSalir = False
            Me.cmdSelTodos.visible = True
            Me.cmdDeselTodos.visible = True
            
        Case 9 'Etiquetas de Proveedores
            PonerFrameNSeriesVisible True, H, W
            Me.Caption = "Proveedores"
            PulsadoSalir = False
            Me.cmdSelTodos.visible = True
            Me.cmdDeselTodos.visible = True
        
        Case 10 'Errores al contabilizar facturas
            PonerFrameCobrosPtesVisible True, H, W
            CargarListaErrContab
            Me.Caption = "Facturas NO contabilizadas: "
            PonerFocoBtn Me.cmdAceptarCobros
        
        Case 11 'Lineas Factura a Rectificar
            PonerFrameNSeriesVisible True, H, W
            Me.Caption = "Lineas Factura a Rectificar"
            PulsadoSalir = False
            Me.cmdSelTodos.visible = True
            Me.cmdDeselTodos.visible = True
            Me.cmdAceptarNSeries.Left = Me.cmdAceptarNSeries.Left + 1000
            Me.cmdCancelar.Left = Me.cmdCancelar.Left + 1000
        
        Case 12 'Mensaje Albaranes que no se van a Facturar
            PonerFrameCobrosPtesVisible True, H, W
            CargarListaAlbaranes
            Me.Caption = "Facturaci�n Albaranes"
            Me.Label1(0).Caption = "Existen Albaranes que NO se van a Facturar:"
            Me.Label1(0).Top = 260
            Me.Label1(0).Left = 480
            PonerFocoBtn Me.cmdAceptarCobros
            
        Case 13 'Muestra Errores
            H = 6000
            W = 8800
            PonerFrameVisible Me.FrameErrores, True, H, W
            Me.Text1.Text = vCampos
            Me.Caption = "Errores"
        
        Case 14 'Muestra Empresas del sistema
            PonerFrameNSeriesVisible True, H, W
            Me.Caption = "Selecci�n"
            CargarListaEmpresas
        Case 15
            H = FrameEtiqEstant.Height
            W = FrameEtiqEstant.Width
            PonerFrameVisible FrameEtiqEstant, True, H, W
            
        Case 16, 20, 25
            '25: Precios de AVAB en comparacion con los de Morales
            
            Caption = "Correcci�n precios"
            If OpcionMensaje = 25 Then Caption = Caption & "      AVAB"
            H = FrameCorreccionPrecios.Height
            W = FrameCorreccionPrecios.Width
            PonerFrameVisible FrameCorreccionPrecios, True, H, W
            Me.cmdCorrecotrPrecios(1).Cancel = True
            lblIndicadorCorregir.Caption = ""
            CargaComboActualizarPrecios
            ListView4.ColumnHeaders(3).Text = " Coste "
            ListView4.ColumnHeaders(4).Width = 799
            'If OpcionMensaje = 20 Then
            If OpcionMensaje <> 16 Then
                
                ListView4.ColumnHeaders(9).Text = "Coste correc."
                If OpcionMensaje = 20 Then
                    Label2(0).Caption = " Correcci�n de precios de articulos con componentes"
                Else
                    Label2(0).Caption = " Correcci�n de precios de articulos desde proveedor"
                    ListView4.ColumnHeaders(4).Width = 0
                    
                    
                    'Los campos
                    ListView4.ColumnHeaders(3).Text = ""
                    ListView4.ColumnHeaders(5).Text = ""
                    ListView4.ColumnHeaders(3).Width = 0
                    ListView4.ColumnHeaders(5).Width = 0
                    ListView4.ColumnHeaders(7).Width = 0
                    ListView4.ColumnHeaders(8).Width = 0
                    ListView4.ColumnHeaders(9).Width = 0
                    
                End If
            Else
                ListView4.ColumnHeaders(9).Text = "Tarifa correc."
                Label2(0).Caption = " Correcci�n de errores y actualizaci�n de tarifas"
            End If
            
        Case 18
            
            Caption = "Mantenimientos"
            H = FrameTraspasoMante.Height
            W = FrameTraspasoMante.Width
            PonerFrameVisible FrameTraspasoMante, True, H, W

        Case 26
            
             Caption = "Produccion"
            H = FrameVerDatosStockProduccion.Height
            W = FrameVerDatosStockProduccion.Width
            PonerFrameVisible FrameVerDatosStockProduccion, True, H, W

    End Select
    'Me.cmdCancel(indFrame).Cancel = True
    Me.Height = H + 350
    Me.Width = W + 70
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PonerFrameCobrosPtesVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Cobros Pendientes Visible y Ajustado al Formulario, y visualiza los controles
'necesario para el Informe

    H = 4600
        
    Select Case OpcionMensaje
        Case 1
            H = 5000
            W = 8600
            Me.Label1(0).Caption = "CLIENTE: " & vCampos
        Case 2
            W = 8800
            Me.cmdAceptarCobros.Top = 4000
            Me.cmdAceptarCobros.Left = 4200
        Case 5 'Componentes
            W = 6000
            H = 5000
            Me.cmdAceptarCobros.Left = 4000

        Case 6, 7 'Prefacturar Albaranes
            W = 7000
            H = 6000
            Me.cmdAceptarCobros.Top = 5400
            Me.cmdAceptarCobros.Left = 4600

        Case 10, 12 'Errores al contabilizar facturas
            H = 6000
            W = 8400
            Me.cmdAceptarCobros.Top = 5300
            Me.cmdAceptarCobros.Left = 4900
            If OpcionMensaje = 12 Then
                Me.cmdCancelarCobros.Top = 5300
                Me.cmdCancelarCobros.Left = 4600
                Me.cmdAceptarCobros.Left = 3300
                Me.Label1(1).Top = 4800
                Me.Label1(1).Left = 3400
                Me.cmdAceptarCobros.Caption = "&SI"
                Me.cmdCancelarCobros.Caption = "&NO"
            End If
    End Select
            
    PonerFrameVisible Me.FrameCobrosPtes, visible, H, W

    If visible = True Then
        Me.txtParam.visible = (OpcionMensaje = 6 Or OpcionMensaje = 7)
        Me.Label1(0).visible = (OpcionMensaje = 1) Or (OpcionMensaje = 5) Or (OpcionMensaje = 12)
        Me.cmdCancelarCobros.visible = (OpcionMensaje = 12)
        Me.Label1(1).visible = (OpcionMensaje = 12)
    End If
End Sub


Private Sub PonerFrameAcercaDeVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame ACERCA DE visible y Ajustado al Formulario

    Me.FrameAcercaDe.visible = visible
    If visible = True Then
        'Ajustar Tama�o del Frame para ajustar tama�o de Formulario al del Frame
        Me.FrameAcercaDe.Top = -90
        Me.FrameAcercaDe.Left = 0
        Me.FrameAcercaDe.Height = 4555
        Me.FrameAcercaDe.Width = 6600
        
        W = Me.FrameAcercaDe.Width
        H = Me.FrameAcercaDe.Height
    End If
End Sub


Private Sub PonerFrameNSeriesVisible(visible As Boolean, H As Integer, W As Integer)
'Pone el Frame de N� Serie Visible y Ajustado al Formulario, y visualiza los controles
'necesario para el Informe

    H = 5000
   
    If OpcionMensaje = 11 Then 'Lineas Factura a Rectificar
        W = 10900
    ElseIf OpcionMensaje = 14 Then
        W = 6500
        Me.Label7(1).visible = True
    Else
        W = 8500
        Me.Label7(1).visible = False
    End If
    PonerFrameVisible Me.FrameNSeries, visible, H, W
End Sub


Private Sub ponerFrameComponentesVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Componentes Visible y Ajustado al Formulario, y visualiza los controles
'necesario para el Informe

'    Me.FrameComponentes.visible = visible
    Me.FrameComponentes2.visible = visible
    
    H = 4000
    W = 5300
    PonerFrameVisible Me.FrameComponentes, visible, H, W
        
    If visible = True Then
        'Ajustar Tama�o del Frame para ajustar tama�o de Formulario al del Frame
        If vParamAplic.Departamento Then
            Me.OptCompXDpto.Caption = "Departemento"
        Else
            Me.OptCompXDpto.Caption = "Direcci�n"
        End If
    End If
End Sub


Private Sub CargarListaCobrosPtes()
'Muestra la lista Detallada de cobros en un ListView
'Carga los valores de la tabla scobro de la Contabilidad
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String

    SQL = "SELECT numserie, codfaccl, fecfaccl, fecvenci, impvenci, impcobro "
    SQL = SQL & " FROM scobro INNER JOIN sforpa ON scobro.codforpa=sforpa.codforpa "
    SQL = SQL & cadWhere

    Set RS = New ADODB.Recordset
    RS.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
     
    ListView1.Top = 900
    ListView1.Height = 3250
    ListView1.Width = 8100
    ListView1.Left = 160
    
    'Los encabezados
    ListView1.ColumnHeaders.Clear
    
    ListView1.ColumnHeaders.Add , , "N� Serie", 760
    ListView1.ColumnHeaders.Add , , "N� Factura", 1100, 1
    ListView1.ColumnHeaders.Add , , "Fecha Factura", 1250, 2
    ListView1.ColumnHeaders.Add , , "Fecha Venci.", 1200, 2
    ListView1.ColumnHeaders.Add , , "Imp. Venci.(�)", 1250, 1
    ListView1.ColumnHeaders.Add , , "Imp. Cobro(�)", 1250, 1
    ListView1.ColumnHeaders.Add , , "Pte. Cobro(�)", 1250, 1
    
    While Not RS.EOF
        Set ItmX = ListView1.ListItems.Add
        ItmX.Text = RS.Fields(0).Value 'N� Serie
        ItmX.SubItems(1) = RS.Fields(1).Value 'N� Factura
        ItmX.SubItems(2) = RS.Fields(2).Value 'Fecha Factura
        ItmX.SubItems(3) = RS.Fields(3).Value 'Fecha Vencimiento
        ItmX.SubItems(4) = RS.Fields(4).Value 'Importe Vencido
        ItmX.SubItems(5) = DBLet(RS.Fields(5).Value, "N") 'Importe Cobrado
        ItmX.SubItems(6) = RS.Fields(4).Value - DBLet(RS.Fields(5).Value, "N") 'Pendiente de cobro
        If ItmX.SubItems(6) > 0 Then
            ItmX.ListSubItems(6).ForeColor = vbRed
        End If
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
End Sub


Private Sub CargarListaArtSinStock(NomTabla As String)
'Muestra la lista Detallada de Articulos que no tienen stock suficiente en un ListView
Dim RS As ADODB.Recordset
Dim SQL As String
    
    SQL = "SELECT " & NomTabla & ".codalmac," & NomTabla & ".codartic, " & NomTabla & ".nomartic, salmac.canstock as canstock, SUM(cantidad) as cantidad, canstock-SUM(cantidad) as disp ,conjunto "
    SQL = SQL & " FROM ((" & NomTabla & " INNER JOIN sartic ON " & NomTabla & ".codartic=sartic.codartic) INNER JOIN sfamia ON sartic.codfamia=sfamia.codfamia) "
    SQL = SQL & " INNER JOIN salmac ON " & NomTabla & ".codalmac=salmac.codalmac and " & NomTabla & ".codartic=salmac.codartic "
    SQL = SQL & cadWhere 'Where numpedcl = 2 And sfamia.instalac = 0
    SQL = SQL & " GROUP by " & NomTabla & ".codalmac, " & NomTabla & ".codartic "
    

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
     
    Me.ListView1.Top = 500
     
    'Los encabezados
    ListView1.Width = 8400
    ListView1.Height = 3150
    ListView1.ColumnHeaders.Clear
    
    ListView1.ColumnHeaders.Add , , "Alm.", 500
    ListView1.ColumnHeaders.Add , , "Articulo", 1800, 2
    ListView1.ColumnHeaders.Add , , "Dec. Artic", 3300
    ListView1.ColumnHeaders.Add , , "Stock", 950, 2
    ListView1.ColumnHeaders.Add , , "Cantidad", 900, 2
    ListView1.ColumnHeaders.Add , , "No Disp.", 900, 2
    
    While Not RS.EOF
        CargaItemStock RS, ""
        'Si no tiene produccion miraremos si es conjunto
        If Not vParamAplic.Produccion Then
            If RS!Conjunto = 1 Then
                SQL = RS!codAlmac & "|" & RS!codartic & "|" & RS!Cantidad & "|"
                CargaStockConjuntos SQL
            End If
        End If
        RS.MoveNext
    Wend
    RS.Close
    
    Set RS = Nothing

    
    
End Sub
    
Private Sub CargaStockConjuntos(linea As String)
    
        
        Set miRsAux = New ADODB.Recordset
            'Deberiamos cargar los elementos que tiene subconjuntos
            cadWhere2 = "SELECT " & RecuperaValor(linea, 1) & ",sarti1.codarti1,nomartic,"
            cadWhere2 = cadWhere2 & " sarti1.cantidad * " & TransformaComasPuntos(RecuperaValor(linea, 3)) & " as cantidad,"
            cadWhere2 = cadWhere2 & " salmac.canstock as canstock,  canstock-(sarti1.cantidad * " & TransformaComasPuntos(RecuperaValor(linea, 3))
            cadWhere2 = cadWhere2 & ") as disp From sarti1, salmac, sartic"
            cadWhere2 = cadWhere2 & " Where sarti1.codarti1 = salmac.codArtic And sarti1.codarti1 = sartic.codArtic"
            cadWhere2 = cadWhere2 & " and sarti1.codartic='" & DevNombreSQL(RecuperaValor(linea, 2)) & "'"
            
            miRsAux.Open cadWhere2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                CargaItemStock miRsAux, " * "
                miRsAux.MoveNext
            Wend
            miRsAux.Close
        
        cadWhere2 = ""
    Set miRsAux = Nothing
End Sub
 
    
Private Sub CargaItemStock(ByRef R As ADODB.Recordset, ByRef TxtA�adido As String)
Dim ItmX As ListItem
     If R!disp < 0 Then
            Set ItmX = ListView1.ListItems.Add
            ItmX.Text = Format(R.Fields(0).Value, "000") 'Cod Almacen
            If TxtA�adido <> "" Then TxtA�adido = "[" & TxtA�adido & "]"
            ItmX.SubItems(1) = TxtA�adido & " " & R.Fields(1).Value 'Cod Artic
            ItmX.SubItems(2) = R.Fields(2).Value 'Nom Artic
            ItmX.SubItems(3) = R.Fields(3).Value 'Stock
            ItmX.SubItems(4) = R.Fields(4).Value 'Cantidad
            ItmX.SubItems(5) = R.Fields(5).Value 'No Disp
    End If
End Sub


Private Sub CargarListaNSeries()
'Carga las lista con todos los N� de serie encontrados en la tabla:sserie
'para el articulo pasado como parametro en la cadwhere: "codartic='00012'"
'y que esten disponibles: numfactu y numalbar no tengan valor
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String
Dim cadLista As String
Dim Dif As Single

    On Error GoTo ECargarLista

    If cadWhere2 = "" Then
        'Mostramos los n� serie libres para seleccionar la cantidad
        SQL = "SELECT sserie.numserie, sserie.codartic, sartic.nomartic "
        SQL = SQL & "FROM sserie INNER JOIN sartic ON sserie.codartic=sartic.codartic "
        SQL = SQL & cadWhere 'Where codartic='000012'
        'seleccionamos los que no esten asignados a ninguna factura ni albaran
        SQL = SQL & " AND ((isnull(sserie.numfactu) or sserie.numfactu='') and (isnull(sserie.numalbar) or sserie.numalbar='')) "
        SQL = SQL & " ORDER BY sserie.codartic, numserie "
        
    Else 'venimos de modificar la cantidad y seleccionamos los ya asignados
        If InStr(1, cadWhere2, "|") > 0 Then
            Dif = CSng(RecuperaValor(cadWhere2, 1))
            cadWhere2 = RecuperaValor(cadWhere2, 2)
        
            'seleccionamos n� serie del albaran que modificamos
            SQL = "SELECT sserie.numserie, sserie.codartic, sartic.nomartic "
            SQL = SQL & "FROM sserie INNER JOIN sartic ON sserie.codartic=sartic.codartic "
            SQL = SQL & cadWhere2
                
            
            If Dif < 0 Then
                'Si la diferencia de cantidad es < 0, mostrar en la lista los n� serie que
                'tiene la linea de albaran asignado con todos marcados y desmarcar el que no queremos
                
            Else
                'si la diferencia de cantidad es > 0, mostrar en la lista los n� de serie que
                'ya tenia asignados la linea del albaran m�s los libres para seleccionar los que a�adimos de mas
                cadLista = ""
                Set RS = New ADODB.Recordset
                RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not RS.EOF
                    cadLista = cadLista & ", " & RS!numSerie
                    RS.MoveNext
                Wend
                RS.Close
                Set RS = Nothing
                
                'mostrar tambien los n� serie sin asignar
                SQL = SQL & " OR (" & Replace(cadWhere, "WHERE", "") & " and (numalbar=''or isnull(numalbar)))"
            End If
        Else
            'viene de una factura rectificativa, seleccionamos los n� de serie de
            'esa factura y marcamos los que queremos quitar
            SQL = cadWhere2
        End If
    End If
    

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Los encabezados
    ListView2.Width = 7400
    Me.ListView2.Height = 3100
    Me.ListView2.Left = 650
    ListView2.ColumnHeaders.Clear
    
    ListView2.ColumnHeaders.Add , , "N� Serie", 1800
    ListView2.ColumnHeaders.Add , , "Articulo", 1800
    ListView2.ColumnHeaders.Add , , "Desc. Artic", 3650
        
    If RS.EOF Then Unload Me
    
    While Not RS.EOF
         Set ItmX = ListView2.ListItems.Add
         ItmX.Text = RS.Fields(0).Value 'num serie
         If Dif < 0 Then
            ItmX.Checked = True
         ElseIf Dif > 0 Then
            If InStr(1, cadLista, CStr(RS!numSerie)) > 0 Then
                ItmX.Checked = True
            Else
                ItmX.Checked = False
            End If
         Else
            ItmX.Checked = False
         End If
         ItmX.SubItems(1) = RS.Fields(1).Value 'Desc Artic
         ItmX.SubItems(2) = RS.Fields(2).Value 'Nom Artic
         RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    
ECargarLista:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargar N� Series", Err.Description
End Sub


Private Sub CargarListaComponentes(opt As Byte)
'Muestra la lista Detallada de cobros en un ListView
'Carga los valores de la tabla scobro de la Contabilidad
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String
Dim Codigo As String, cadCodigo As String

    Select Case opt
        Case 1 'Mantenimiento
            Codigo = RecuperaValor(vCampos, 1)
            If Codigo = "" Then
                cadCodigo = " isnull(nummante) "
            Else
                cadCodigo = " nummante=" & DBSet(Codigo, "T")
            End If
            SQL = ObtenerSQLcomponentes(cadWhere & " and " & cadCodigo)
            Me.Label1(0).Caption = "Mantenimiento: " & Codigo
            
        Case 2 'Departamento
            Codigo = RecuperaValor(vCampos, 2)
            If Codigo = "" Then
                cadCodigo = "isnull(coddirec)"
            Else
                cadCodigo = " coddirec=" & Codigo
            End If
            SQL = ObtenerSQLcomponentes(cadWhere & " and " & cadCodigo)
            If vParamAplic.Departamento Then
                Me.Caption = "Equipos del Departamento"
                Me.Label1(0).Caption = " Departamento: " & RecuperaValor(vCampos, 3)
            Else
                Me.Caption = "Equipos de la Direcci�n"
                Me.Label1(0).Caption = " Direcci�n: " & Codigo & " " & RecuperaValor(vCampos, 3)
            End If
        
        Case 3 'Cliente
            SQL = ObtenerSQLcomponentes(cadWhere)
            Me.Caption = "Equipos del Cliente"
            Me.Label1(0).Caption = "Cliente: " & RecuperaValor(vCampos, 4)
    End Select
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Los encabezados
    ListView1.Top = 800
    ListView1.Left = 280
    ListView1.Width = 4900
    ListView1.Height = 3250
    ListView1.ColumnHeaders.Clear
    
    ListView1.ColumnHeaders.Add , , "TA", 760
    ListView1.ColumnHeaders.Add , , "Tipo Articulo", 2800
    ListView1.ColumnHeaders.Add , , "Cantidad", 1280, 2
    
    If Not RS.EOF Then
        While Not RS.EOF
            Set ItmX = ListView1.ListItems.Add
            ItmX.Text = RS.Fields(0).Value 'TA
            ItmX.SubItems(1) = RS.Fields(1).Value 'Tipo Articulo
            ItmX.SubItems(2) = RS.Fields(2).Value 'Cantidad
            RS.MoveNext
        Wend
    End If
    RS.Close
    Set RS = Nothing
End Sub


Private Sub CargarListaPreFacturar()
'Muestra la lista Detallada de Albaranes a Factura en un ListView
'Carga los valores de la tabla scobro de la Contabilidad
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String

    On Error GoTo ECargarList
    
    SQL = "CREATE TEMPORARY TABLE tmp ( "
    SQL = SQL & "codforpa SMALLINT(3) UNSIGNED  DEFAULT '0' NOT NULL, "
    SQL = SQL & "numalbar MEDIUMINT(7) UNSIGNED  DEFAULT '0' NOT NULL, "
    SQL = SQL & "dtoppago DECIMAL(4,2) UNSIGNED  DEFAULT '0.0' NOT NULL, "
    SQL = SQL & "dtopgnral DECIMAL(4,2) UNSIGNED  DEFAULT '0.0' NOT NULL, "
    SQL = SQL & "importe DECIMAL(12,2) UNSIGNED  DEFAULT '0.0' NOT NULL, "
    SQL = SQL & "bruto DECIMAL(12,2) UNSIGNED  DEFAULT '0.0' NOT NULL) "
    conn.Execute SQL
     
'     SQL = "LOCK TABLES scaalb READ, slialb READ;"
'     Conn.Execute SQL
     
    SQL = "SELECT scaalb.codforpa, scaalb.numalbar, dtoppago, dtognral, round(sum(importel),2) as importe, round(sum(importel),2) - round(((round(sum(importel),2)*dtoppago)/100),2) - round(((round(sum(importel),2)*dtognral)/100),2) as bruto "
    SQL = SQL & " FROM (scaalb INNER JOIN slialb ON scaalb.codtipom=slialb.codtipom and scaalb.numalbar=slialb.numalbar) "
    SQL = SQL & " WHERE " & cadWhere
    SQL = SQL & " GROUP BY scaalb.numalbar "
    SQL = SQL & " ORDER BY scaalb.codforpa, scaalb.numalbar "

    SQL = " INSERT INTO tmp " & SQL
    conn.Execute SQL
     
    SQL = " SELECT tmp.codforpa, sforpa.nomforpa, sum(tmp.bruto) as bruto"
    SQL = SQL & " FROM tmp, sforpa WHERE tmp.codforpa=sforpa.codforpa "
    SQL = SQL & " GROUP BY tmp.codforpa "
    
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        ListView1.Height = 3850
        ListView1.Width = 5400
        ListView1.Left = 500
        ListView1.Top = 1200
    '    ListView1.GridLines = False
    
        'Los encabezados
        ListView1.ColumnHeaders.Clear
        
        ListView1.ColumnHeaders.Add , , " Forma de Pago", 3300
        ListView1.ColumnHeaders.Add , , "Base Imp.(�)", 2020, 1
     
        While Not RS.EOF
            Set ItmX = ListView1.ListItems.Add
            ItmX.Text = Format(RS!codforpa.Value, "000") & "  " & RS!nomforpa.Value
            
            ItmX.SubItems(1) = RS!bruto
            RS.MoveNext
        Wend
    End If
    RS.Close
    Set RS = Nothing
    
    'Borrar la tabla temporal
    SQL = " DROP TABLE IF EXISTS tmp;"
    conn.Execute SQL

ECargarList:
    If Err.Number <> 0 Then
         'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmp;"
        conn.Execute SQL
'        SQL = "UNLOCK TABLES "
'        Conn.Execute SQL
    End If
End Sub


Private Sub CargarListaClientes()
'Carga las lista con todos los clientes seleccionados en la tabla:sclien
'para imprimir etiquetas, pasando como parametro la cadwhere
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String, Men As String

    On Error GoTo ECargarLista

    Select Case OpcionMensaje
    Case 8
        'CLIENTES
        SQL = "SELECT codclien,nomclien,nifclien "
        SQL = SQL & "FROM sclien "
        If cadWhere <> "" Then SQL = SQL & " WHERE " & cadWhere
        SQL = SQL & " ORDER BY codclien "
        Men = "Cliente"
    Case 9
        'PROVEEDORES
        SQL = "SELECT codprove,nomprove,nifprove "
        SQL = SQL & "FROM sprove "
        If cadWhere <> "" Then SQL = SQL & " WHERE " & cadWhere
        SQL = SQL & " ORDER BY codprove "
        Men = "Proveedor"
    Case 17
        'CLIENTES MANTENIMIENTO
        SQL = cadWhere
    End Select
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        'Los encabezados
        ListView2.Width = 7000
        ListView2.Top = 500
        ListView2.Height = 3620
        ListView2.ColumnHeaders.Clear
        
        ListView2.ColumnHeaders.Add , , Men, 1350
        ListView2.ColumnHeaders.Add , , "Nombre", 4000
        ListView2.ColumnHeaders.Add , , "NIF", 1330
        
        While Not RS.EOF
             Set ItmX = ListView2.ListItems.Add
             ItmX.Text = Format(RS.Fields(0).Value, "000000") 'cod clien/prove
             ItmX.Checked = False
             ItmX.SubItems(1) = RS.Fields(1).Value 'Nom clien/prove
             ItmX.SubItems(2) = RS.Fields(2).Value 'NIF clien/prove
             RS.MoveNext
        Wend
    End If
    RS.Close
    Set RS = Nothing
    
ECargarLista:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Cargar " & Men, Err.Description
        PulsadoSalir = True
        Unload Me
    End If
End Sub



Private Sub CargarListaErrContab()
'Muestra la lista Detallada de Facturas que dieron error al contabilizar
'en un ListView
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String

    On Error GoTo ECargarList

    SQL = " SELECT  * "
    SQL = SQL & " FROM tmpErrFac "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        ListView1.Height = 4500
        ListView1.Width = 7400
        ListView1.Left = 500
        ListView1.Top = 500

        'Los encabezados
        ListView1.ColumnHeaders.Clear

        If RS.Fields(0).Name = "codprove" Then
            'Facturas de Compra
             ListView1.ColumnHeaders.Add , , "Prove.", 700
        Else 'Facturas de Venta
            ListView1.ColumnHeaders.Add , , "Tipo", 600
        End If
        ListView1.ColumnHeaders.Add , , "Factura", 1000, 1
        ListView1.ColumnHeaders.Add , , "Fecha", 1100, 1
        ListView1.ColumnHeaders.Add , , "Error", 4620
    
        While Not RS.EOF
            Set ItmX = ListView1.ListItems.Add
            'El primer campo ser� codtipom si llamamos desde Ventas
            ' y ser� codprove si llamamos desde Compras
            ItmX.Text = RS.Fields(0).Value
            ItmX.SubItems(1) = Format(RS!NumFactu, "0000000")
            ItmX.SubItems(2) = RS!Fecfactu
            ItmX.SubItems(3) = RS!Error
            RS.MoveNext
        Wend
    End If
    RS.Close
    Set RS = Nothing

ECargarList:
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub


Private Sub CargarListaLinFactu()
'Carga las lista con todas las lineas de la factura que estamos rectificando
'seleccionamos las que nos queremos llevar al Albaran de rectificacion
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String

    On Error GoTo ECargarLista

    SQL = "SELECT codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel,origpre"
    SQL = SQL & " FROM slifac "
    If cadWhere <> "" Then SQL = SQL & " WHERE " & cadWhere
    SQL = SQL & " ORDER BY codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        
        ListView2.Top = 500
        ListView2.Left = 380
        ListView2.Width = 10100
        ListView2.Height = 3620
        
        'Los encabezados
        ListView2.ColumnHeaders.Clear
    
        ListView2.ColumnHeaders.Add , , "T.Alb", 660
        ListView2.ColumnHeaders.Add , , "N� Alb", 840
        ListView2.ColumnHeaders.Add , , "Lin.", 450
         ListView2.ColumnHeaders.Item(3).Alignment = lvwColumnRight
        ListView2.ColumnHeaders.Add , , "Alm", 460
        ListView2.ColumnHeaders.Add , , "Artic", 1380
        ListView2.ColumnHeaders.Add , , "Desc. Artic.", 2500
        ListView2.ColumnHeaders.Add , , "Cant.", 600
        ListView2.ColumnHeaders.Item(7).Alignment = lvwColumnRight
        ListView2.ColumnHeaders.Add , , "Precio", 960
        ListView2.ColumnHeaders.Item(8).Alignment = lvwColumnRight
        ListView2.ColumnHeaders.Add , , "Dto 1", 600
        ListView2.ColumnHeaders.Item(9).Alignment = lvwColumnRight
        ListView2.ColumnHeaders.Add , , "Dto 2", 600
        ListView2.ColumnHeaders.Item(10).Alignment = lvwColumnRight
        ListView2.ColumnHeaders.Add , , "Importe", 950
        ListView2.ColumnHeaders.Item(11).Alignment = lvwColumnRight
    
        While Not RS.EOF
             Set ItmX = ListView2.ListItems.Add
             ItmX.Text = RS!Codtipoa 'cod tipo alb
             ItmX.Checked = False
             ItmX.SubItems(1) = Format(RS!NumAlbar, "0000000") 'N� Albaran
             ItmX.SubItems(2) = RS!numlinea 'linea Albaran
             ItmX.SubItems(3) = Format(RS!codAlmac, "000") 'cod almacen
             ItmX.SubItems(4) = RS!codartic 'Cod Articulo
             ItmX.SubItems(5) = RS!NomArtic 'Nombre del Articulo
             ItmX.SubItems(6) = RS!Cantidad
             ItmX.SubItems(7) = Format(RS!precioar, FormatoPrecio)
             ItmX.SubItems(8) = RS!dtoline1
             ItmX.SubItems(9) = RS!dtoline2
             ItmX.SubItems(10) = Format(RS!ImporteL, FormatoImporte)
             RS.MoveNext
        Wend
    End If
    RS.Close
    Set RS = Nothing
    
    'si aparece la barra de desplazamiento ajustar el ancho
    If Me.ListView2.ListItems.Count > 11 Then
        Me.ListView2.ColumnHeaders(5).Width = 1200 'codartic
        Me.ListView2.ColumnHeaders(8).Width = 920  'precio
    End If
   
    
    
ECargarLista:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Cargar Lineas Factura", Err.Description
        PulsadoSalir = True
        Unload Me
    End If
End Sub




Private Sub CargarListaAlbaranes()
'Muestra la lista Detallada de Facturas que dieron error al contabilizar
'en un ListView
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String

    On Error GoTo ECargarList

    SQL = cadWhere 'cadwhere ya le pasamos toda la SQL
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
        ListView1.Height = 3900
        ListView1.Width = 7200
        ListView1.Left = 500
        ListView1.Top = 700

        'Los encabezados
        ListView1.ColumnHeaders.Clear

        ListView1.ColumnHeaders.Add , , "Tipo", 700
        ListView1.ColumnHeaders.Add , , "N� Albaran", 1000, 1
        ListView1.ColumnHeaders.Add , , "Fecha", 1100, 1
        ListView1.ColumnHeaders.Item(3).Alignment = lvwColumnCenter
        ListView1.ColumnHeaders.Add , , "Cod. Cli.", 900
        ListView1.ColumnHeaders.Add , , "Cliente", 3400
    
        While Not RS.EOF
            Set ItmX = ListView1.ListItems.Add
            ItmX.Text = RS.Fields(0).Value
            ItmX.SubItems(1) = Format(RS!NumAlbar, "0000000")
            ItmX.SubItems(2) = RS!FechaAlb
            ItmX.SubItems(3) = Format(RS!CodClien, "000000")
            ItmX.SubItems(4) = RS!nomclien
            RS.MoveNext
        Wend
    End If
    RS.Close
    Set RS = Nothing

ECargarList:
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub



Private Sub CargarListaEmpresas()
'Carga las lista con todas las empresas que hay en el sistema
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String
Dim I As Integer

Dim Prohibidas As String

    On Error GoTo ECargarLista

    VerEmresasProhibidas Prohibidas
    
    SQL = "Select * from usuarios.empresasarioli order by codempre"
    Set ListView2.SmallIcons = frmppal.ImageListB
    ListView2.Width = 5000
    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , "Empresa", 4900
    ListView2.HideColumnHeaders = True
    ListView2.GridLines = False
    ListView2.ListItems.Clear
    
    Set RS = New ADODB.Recordset
    I = -1
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
        SQL = "|" & RS!codempre & "|"
        If InStr(1, Prohibidas, SQL) = 0 Then
            Set ItmX = ListView2.ListItems.Add(, , RS!nomempre, , 5)
            ItmX.Tag = RS!codempre
            If ItmX.Tag = vEmpresa.codempre Then
                ItmX.Checked = True
                I = ItmX.Index
            End If
            ItmX.ToolTipText = RS!AriGes
        End If
        RS.MoveNext
    Wend
    RS.Close
    If I > 0 Then Set ListView2.SelectedItem = ListView2.ListItems(I)

    
ECargarLista:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Cargando datos empresas", Err.Description
        PulsadoSalir = True
        Unload Me
    End If
End Sub


Private Sub VerEmresasProhibidas(ByRef VarProhibidas As String)
Dim SQL As String
Dim RS As ADODB.Recordset

On Error GoTo EVerEmresasProhibidas
    VarProhibidas = "|"
    SQL = "Select codempre from usuarios.usuarioempresasarioli WHERE codusu = " & (vUsu.Codigo Mod 1000)
    SQL = SQL & " order by codempre"
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
          VarProhibidas = VarProhibidas & RS!codempre & "|"
          RS.MoveNext
    Wend
    RS.Close
    Exit Sub
EVerEmresasProhibidas:
    MuestraError Err.Number, Err.Description & vbCrLf & " Consulte soporte t�cnico"
    Set RS = Nothing
End Sub



Private Sub CargaImagen()
On Error Resume Next
    Image2.Picture = LoadPicture(App.Path & "\minilogo.bmp")
    'Image1.Picture = LoadPicture(App.path & "\fondon.gif")
    Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If PulsadoSalir = False Then Cancel = 1
End Sub



Private Function ObtenerTamanyosArray() As Boolean
'Para el frame de los N� de Serie de los Articulos
'En cada indice pone en CodArtic(i) el codigo del articulo
'y en Cantidad(i) la cantidad solicitada de cada codartic
Dim I As Integer, J As Integer

    ObtenerTamanyosArray = False
    'Primero a los campos de la tabla
    TotalArray = -1
    J = 0
    Do
        I = J + 1
        J = InStr(I, vCampos, "�")
        If J > 0 Then TotalArray = TotalArray + 1
    Loop Until J = 0
    
    If TotalArray < 0 Then Exit Function
    
    'Las redimensionaremos
    ReDim codartic(TotalArray)
    ReDim Cantidad(TotalArray)
    
    ObtenerTamanyosArray = True
End Function


Private Function SeparaCampos() As Boolean
'Para el frame de los N� de Serie de los Articulos
Dim Grupo As String
Dim I As Integer
Dim J As Integer
Dim C As Integer 'Contador dentro del array

    SeparaCampos = False
    I = 0
    C = 0
    Do
        J = I + 1
        I = InStr(J, vCampos, "�")
        If I > 0 Then
            Grupo = Mid(vCampos, J, I - J)
            'Y en la martriz
            InsertaGrupo Grupo, C
            C = C + 1
        End If
    Loop Until I = 0
    SeparaCampos = True
End Function


Private Sub InsertaGrupo(Grupo As String, contador As Integer)
Dim J As Integer
Dim Cad As String

    J = 0
    Cad = ""
    
    'Cod Artic
    J = InStr(1, Grupo, "|")
    If J > 0 Then
        Cad = Mid(Grupo, 1, J - 1)
        Grupo = Mid(Grupo, J + 1)
        J = 1
    End If
    codartic(contador) = Cad
    
    'Cantidad
    J = InStr(1, Grupo, "|")
    If J > 0 Then
        Cad = Mid(Grupo, 1, J - 1)
        Grupo = Mid(Grupo, J + 1)
    Else
        Cad = Grupo
        Grupo = ""
    End If
    Cantidad(contador) = Cad
End Sub





Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If Shift And vbCtrlMask > 0 Then
            MsgBox "HOLITA trabajador." & vbCrLf & "Has encontrado el huevo de pascua...., a curraaaaaarrr!!!!", vbExclamation
        End If
    End If
End Sub

Private Sub imgCheck_Click(Index As Integer)
Dim b As Boolean
    If Index < 2 Then
        'En el listview3
        b = Index = 1
        For TotalArray = 1 To ListView3.ListItems.Count
            ListView3.ListItems(TotalArray).Checked = b
            If (TotalArray Mod 50) = 0 Then DoEvents
        Next TotalArray
        
    Else
        'En el listview4
        b = Index = 3
        For TotalArray = 1 To ListView4.ListItems.Count
            If ListView4.ListItems(TotalArray).Tag <> "" Then
                ListView4.ListItems(TotalArray).Checked = b
            Else
                ListView4.ListItems(TotalArray).Checked = False
            End If
            If (TotalArray Mod 50) = 0 Then DoEvents
        Next TotalArray
    End If
End Sub



Private Sub ListView1_DblClick()
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    
    MsgBox ListView1.SelectedItem.SubItems(3), vbExclamation
End Sub

Private Sub OptCompXClien_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub OptCompXDpto_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub OptCompXMant_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub



Private Function RegresarCargaEmpresas() As String
Dim SQL As String
Dim Parametros As String
Dim I As Integer

    CadenaDesdeOtroForm = ""
    
        SQL = ""
        Parametros = ""
        For I = 1 To ListView2.ListItems.Count
            If Me.ListView2.ListItems(I).Checked Then
                SQL = SQL & Me.ListView2.ListItems(I).Text & "|"
                Parametros = Parametros & "1" 'Contador
            End If
        Next I
        CadenaDesdeOtroForm = Len(Parametros) & "|" & SQL
        'Vemos las conta
        SQL = ""
        For I = 1 To ListView2.ListItems.Count
            If Me.ListView2.ListItems(I).Checked Then
                SQL = SQL & Me.ListView2.ListItems(I).Tag & "|"
            End If
        Next I
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & SQL
    
    
        RegresarCargaEmpresas = CadenaDesdeOtroForm

End Function



Private Sub CargarArticulosEstanteria()
Dim SQL As String
Dim RS As ADODB.Recordset
Dim It As ListItem

    SQL = "select sartic.codartic,nomartic,preciove,codigiva,nomfamia from sartic,sfamia where sartic.codfamia=sfamia.codfamia"
    If cadWhere <> "" Then SQL = SQL & " AND " & cadWhere
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    TotalArray = 0
    While Not RS.EOF
        Set It = ListView3.ListItems.Add
        
        'Ponemos el codigo de articulo y el TIPO de IVA
        It.Tag = "'" & DevNombreSQL(RS!codartic) & "'," & RS!codigiva
        It.Text = RS!NomArtic
        It.SubItems(1) = Format(RS!preciove, cadWhere2)
        It.SubItems(2) = RS!nomfamia
        It.Checked = True
        RS.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            DoEvents
        End If
    Wend
    RS.Close
    
    
End Sub




Private Sub CargarArticulosCorreccionPrecio()
Dim SQL As String
Dim RS As ADODB.Recordset
Dim It As ListItem
Dim margen As Currency
Dim MargenT As Currency
Dim ImpPVP As Currency
Dim ImpTar As Currency
Dim Aux As Currency
Dim decimales As Long
Dim PrecioUC As Currency
Dim SoloImporteMenor As Boolean
Dim SobreUPC As Boolean

    'El amrgen a aplicar
    'Si la tarifa es sobre el PVP es el articulo
    'si es sobre UPC entonces es sobre el de la tarifa

    lblIndicadorCorregir = "LEYENDO BD"
    lblIndicadorCorregir.Refresh
    

    
    'Si NUMREGELIM=1 entonces esta marcada la opcion(check) de solo importes menores
    If NumRegElim = 1 Then SoloImporteMenor = True
    
    
    
    'Comprobamos la tarifa donde se aplica, si sobre PVP o sobre ultima compra (%tarifa)
    SQL = DevuelveDesdeBD(conAri, "opcionINC", "starif", "codlista", vCampos)
    SobreUPC = Val(SQL) = 1
            
    
    TotalArray = InStr(1, cadWhere2, ",")
    SQL = Mid(cadWhere2, TotalArray + 1)
    decimales = Len(SQL)
    'Formato
    cadWhere2 = "#,##0." & Mid(cadWhere2, TotalArray + 1)
    
    'Sql
    SQL = " SELECT sartic.nomartic,slista.codartic,sartic.preciove,sartic.preciouc,"
    SQL = SQL & "slista.precioac, slista.codlista, starif.nomlista,"
    SQL = SQL & "sartic.margecom as margenArt,starif.margecom as margetar"
    SQL = SQL & " FROM   (slista INNER JOIN sartic ON slista.codartic=sartic.codartic)"
    SQL = SQL & " INNER JOIN starif  ON slista.codlista=starif.codlista WHERE "

    SQL = SQL & cadWhere '& " AND "
    ''SQL = SQL & " sartic.preciove <> sartic.preciouc + round((sartic.preciouc * if(isnull(sartic.margecom),0,sartic.margecom))/100," & Decimales & ")"
    
    SQL = SQL & " ORDER BY slista.codartic"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    

    '
  
    TotalArray = 0
    
    While Not RS.EOF
        'Calculo los importes
        lblIndicadorCorregir.Caption = RS!codartic
        lblIndicadorCorregir.Refresh
        
        margen = DBLet(RS!margenart, "N") / 100
        MargenT = DBLet(RS!margetar, "N") / 100
        PrecioUC = DBLet(RS!PrecioUC, "N")
        
        Aux = margen * PrecioUC
        ImpPVP = Round2(PrecioUC + Aux, decimales)
        
        'El de la tarifa
        If SobreUPC Then
            Aux = MargenT * PrecioUC
            ImpTar = Round2(PrecioUC + Aux, CLng(decimales))
        Else
        
            Aux = MargenT * ImpPVP
            ImpTar = Round2(ImpPVP + Aux, CLng(decimales))
        End If
        Aux = Round2(RS!preciove, decimales)
        
        SQL = ""
        

        If SoloImporteMenor Then
            If Aux >= ImpPVP Then
                'El primero esta bien
                'Veamos el segundo. En la tarifa
                Aux = Round2(RS!precioac, decimales)
                If Aux < ImpTar Then SQL = "M"
            Else
                SQL = "M"
            End If
        
        
        Else
            If Aux = ImpPVP Then
                'El primero esta bien
                'Veamos el segundo. En la tarifa
                Aux = Round2(RS!precioac, decimales)
                If Aux <> ImpTar Then SQL = "M"
            Else
                SQL = "M"
            End If
        End If
        
        If SQL <> "" Then
            Set It = ListView4.ListItems.Add
            It.Tag = DevNombreSQL(RS!codartic)
            It.ToolTipText = It.Tag
            It.Text = It.Tag
            It.SubItems(1) = RS!NomArtic
            Aux = Round2(PrecioUC, decimales)
            It.SubItems(2) = Format(Aux, cadWhere2)
            
            It.SubItems(3) = Format(margen * 100, FormatoPorcen)
            Aux = Round2(RS!preciove, decimales)
            It.SubItems(4) = Format(Aux, cadWhere2)
            
            It.SubItems(5) = Format(MargenT * 100, FormatoPorcen)
            Aux = Round2(RS!precioac, decimales)
            It.SubItems(6) = Format(Aux, cadWhere2)
            

            It.SubItems(7) = Format(ImpPVP, cadWhere2)
            It.SubItems(8) = Format(ImpTar, cadWhere2)
            
            
            
            If PrecioUC = 0 Then
                'Precio ultima compra =0
                'NOOOOO se puede actualizar la tarifa
                It.Tag = "" 'para no actualizar
                It.Checked = False
                It.Bold = True
                It.ForeColor = vbRed
            Else
                
            End If
            It.Checked = False
        End If
        RS.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            Me.Refresh
            DoEvents
        End If
    Wend
    RS.Close
    cmbActualizarTar.ListIndex = 0
    lblIndicadorCorregir.Caption = ""
End Sub




Private Function TraspasarMantenimientos() As Boolean
    
    On Error GoTo ETraspasarMantenimientos
    TraspasarMantenimientos = False

    

    cadWhere = "Select count(*) from sliman where anomante =" & txtMante(0).Text
    miRsAux.Open cadWhere, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    
    If NumRegElim > 0 Then
        MsgBox "Ya existen datos para el a�o " & txtMante(0).Text, vbExclamation
        Exit Function
    End If
    
    
    
    'Se divide en 4 pasos
    '1.- Introducir una linea en la sliman con los datos para el a�o
        cadWhere = "insert into sliman (anomante,codclien,nummante,mes01man,mes02man,mes03man,mes04man,mes05man,mes06man,mes07man,mes08man,mes09man,mes10man,mes11man,mes12man)"
        cadWhere = cadWhere & " SELECT " & txtMante(0).Text & ",codclien,nummante,mes01act,mes02act,mes03act,mes04act,mes05act,mes06act,mes07act,mes08act,mes09act,mes10act,mes11act,mes12act FROM scaman"
        conn.Execute cadWhere
    '2.- Updatear los campos de actual con siguiente
        cadWhere = ""
        For TotalArray = 1 To 12
            cadWhere = cadWhere & ", mes" & Format(TotalArray, "00") & "act = mes" & Format(TotalArray, "00") & "sig"
        Next TotalArray
        cadWhere = Mid(cadWhere, 2) 'Para quitar la primera coma
        cadWhere = "UPDATE scaman SET " & cadWhere
        conn.Execute cadWhere
        
    '3.- Si no han marcado la opcion copiar datos tengo que resetear a 0
        If chkMante.Value = 0 Then
            'NO SE COPIA, luego hay que resetear
            cadWhere = ""
            For TotalArray = 1 To 12
                cadWhere = cadWhere & ", mes" & Format(TotalArray, "00") & "sig = 0 "
            Next TotalArray
            cadWhere = Mid(cadWhere, 2) 'Para quitar la primera coma
            cadWhere = "UPDATE scaman SET " & cadWhere
            conn.Execute cadWhere
        End If
        
    '4.- Ultimo mes facturado pasa a ser  cero
        conn.Execute "UPDATE scaman SET ulmesfac=0"
        
    TraspasarMantenimientos = True
    
    Exit Function
ETraspasarMantenimientos:
    MuestraError Err.Number
End Function



Private Sub txtMante_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub




Private Sub CargaPVPPreciosArticulosConComponentes()
Dim decimales As Byte
Dim SQL As String
Dim Impor As Currency
Dim IA As Currency
Dim PC As Currency
Dim PCC As Currency
Dim CosteFormato As Currency
Dim Cantidad As Currency  'Lleva factor de conversion
Dim RCostes As ADODB.Recordset
Dim V1 As Single

    Set miRsAux = New ADODB.Recordset
    
    lblIndicadorCorregir = "LEYENDO BD"
    lblIndicadorCorregir.Refresh
    
    'Guardo los costes asociados al formato. Leere toooodos los costes(tampoco seran tatntos)
    Set RCostes = New ADODB.Recordset
    SQL = "select distinct(s2.codunida) from sarti1,sartic as s1,sartic as s2 where sarti1.codartic=s1.codartic and sarti1.codarti1=s2.codartic"
    'Si lleva WHERE
    If cadWhere <> "" Then
        vCampos = Replace(cadWhere, "sartic.", "s1.")
        SQL = SQL & " AND " & vCampos
        vCampos = ""
    End If
    
    'sql = "select codunida,sum(importe) from sunilin   where codunida in (" & sql & ") GROUP bY codunida"
    SQL = "select codunida,sum(importe) from sunilin  GROUP bY codunida"
    RCostes.Open SQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    'Tengo los costes asocaidos al formato del producto en el reco
    
    
    
    'Fomato importe
    TotalArray = InStr(1, cadWhere2, ",")
    SQL = Mid(cadWhere2, TotalArray + 1)
    decimales = Len(SQL)
    'Formato
    cadWhere2 = "#,##0." & Mid(cadWhere2, TotalArray + 1)
    
    
    'Tres columna svamos a ponerlas a tama�o 0
    ListView4.ColumnHeaders(6).Width = 0
    ListView4.ColumnHeaders(7).Width = 0
    
    SQL = "select sarti1.*,s1.nomartic,s1.preciove pre2,s1.margecom,s1.preciouc,"
    SQL = SQL & " sarti1.cantidad,s2.preciove, s2.preciouc coste, s2.factorconversion, s1.codunida"
    SQL = SQL & " from sarti1,sartic as s1,sartic as s2 where sarti1.codartic=s1.codartic and sarti1.codarti1=s2.codartic"
    'Si lleva WHERE
    If cadWhere <> "" Then
        vCampos = Replace(cadWhere, "sartic.", "s1.")
        SQL = SQL & " AND " & vCampos
        vCampos = ""
    End If
    
    SQL = SQL & " ORDER BY sarti1.codartic"
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    SQL = ""

    While Not miRsAux.EOF
        If SQL <> miRsAux!codartic Then
            
        
        
            'Nuevo articulo
            lblIndicadorCorregir = miRsAux!codartic
            lblIndicadorCorregir.Refresh
            If SQL <> "" Then
                'PVP calculado.  Marzo/Abril 2009
                PCC = PCC + CosteFormato      'precio compra calculado
                PCC = Round2(PCC, CLng(decimales))
                
                
                Impor = (PCC * Impor) / 100
                Impor = Impor + PCC
                Impor = Round2(Impor, CLng(decimales))
                'Si precioventa distionto   o pcompra distionto
                If IA <> Impor Or PC <> PCC Then
                    vCampos = vCampos & Format(IA, cadWhere2) & "|" & Format(Impor, cadWhere2) & "|"
                    vCampos = vCampos & Format(PC, cadWhere2) & "|" & Format(PCC, cadWhere2) & "|"
                    InsertarItemARticuloConjunto vCampos
                End If
                    
                
            End If
            
            'Obtengo el coste formato , de la tabla sunida
            SQL = ""
            RCostes.Find "codunida = " & CStr(miRsAux!CodUnida), , adSearchForward, 1
            If Not RCostes.EOF Then
                SQL = DBLet(RCostes.Fields(1), "N")
            End If
            If SQL = "" Then SQL = "0"
            CosteFormato = CCur(SQL)
        
            SQL = miRsAux!codartic
            vCampos = miRsAux!codartic & "|" & miRsAux!NomArtic & "|"
            PC = DBLet(miRsAux!PrecioUC, "N")
            vCampos = vCampos & Format(PC, cadWhere2)
            vCampos = vCampos & "|" & Format(DBLet(miRsAux!margecom, "N"), FormatoPorcen) & "|"
            
            
            IA = miRsAux!pre2
            PCC = 0
            'Impor = 0 + CosteFormato    'pvp calculado
            Impor = DBLet(miRsAux!margecom, "N") 'Precio copste * margencom
        End If
        
        'If miRsAux!codArtic = "009200010306" Then
        
        V1 = miRsAux!FactorConversion * miRsAux!Cantidad
       
'        Impor = Impor + Round2((cantidad * miRsAux!preciove), CLng(decimales))
        Cantidad = DBLet(miRsAux!coste, "N")
        Cantidad = Round2(Cantidad * V1, 4)
        
        PCC = PCC + Cantidad
        miRsAux.MoveNext
    Wend
    If SQL <> "" Then
            PCC = PCC + CosteFormato      'precio compra calculado
            PCC = Round2(PCC, CLng(decimales))
            
            
            Impor = (PCC * Impor) / 100
            Impor = Impor + PCC
            Impor = Round2(Impor, CLng(decimales))
            'Si precioventa distionto   o pcompra distionto
            If IA <> Impor Or PC <> PCC Then
                vCampos = vCampos & Format(IA, cadWhere2) & "|" & Format(Impor, cadWhere2) & "|"
                vCampos = vCampos & Format(PC, cadWhere2) & "|" & Format(PCC, cadWhere2) & "|"
                InsertarItemARticuloConjunto vCampos
            End If
    

    End If
    miRsAux.Close
    RCostes.Close
    Set RCostes = Nothing
    lblIndicadorCorregir = ""
End Sub



Private Sub InsertarItemARticuloConjunto(Datos As String)
Dim It As ListItem

        Set It = ListView4.ListItems.Add
        It.Tag = RecuperaValor(Datos, 1)
        It.ToolTipText = It.Tag
        It.Text = It.Tag
        It.SubItems(1) = RecuperaValor(Datos, 2)  'nomartic
    
        It.SubItems(2) = RecuperaValor(Datos, 3)  'precio UC del articulo
        It.SubItems(3) = RecuperaValor(Datos, 4)  ' Margen
        
        It.SubItems(4) = RecuperaValor(Datos, 5)  'PVP articulo
        It.SubItems(7) = RecuperaValor(Datos, 6)  'PVP calculado
        It.SubItems(8) = RecuperaValor(Datos, 8)  'PUC calculado
        
            
End Sub

Private Sub CargaComboActualizarPrecios()
    cmbActualizarTar.Clear
    
    If OpcionMensaje = 16 Then
        'ART Y TARIFAS
        cmbActualizarTar.Tag = "Art�culos y tarifas|Solo art�culo|Solo tarifas|"
    Else
        cmbActualizarTar.Tag = "PVP y Coste|Solo PVP|Solo Coste|"
    End If
    cmbActualizarTar.AddItem RecuperaValor(cmbActualizarTar.Tag, 1)
    cmbActualizarTar.AddItem RecuperaValor(cmbActualizarTar.Tag, 2)
    cmbActualizarTar.AddItem RecuperaValor(cmbActualizarTar.Tag, 3)
    cmbActualizarTar.Tag = ""
    cmbActualizarTar.ListIndex = 0
End Sub




Private Sub CargaPVPPreciosDesdeMorales()
Dim decimales As Byte
Dim SQL As String
Dim UPC1 As Currency
Dim UPC2 As Currency
Dim PVP1 As Currency
Dim PVP2 As Currency




    On Error GoTo ECargaPVPPreciosDesdeMorales

    Set miRsAux = New ADODB.Recordset
    
    lblIndicadorCorregir = "LEYENDO BD: 1"
    lblIndicadorCorregir.Refresh
    
    
    'Fomato importe
    TotalArray = InStr(1, cadWhere2, ",")
    SQL = Mid(cadWhere2, TotalArray + 1)
    decimales = Len(SQL)
    'Formato
    cadWhere2 = "#,##0." & Mid(cadWhere2, TotalArray + 1)
    
    
    'Tres columna svamos a ponerlas a tama�o 0
    ListView4.ColumnHeaders(6).Width = 0
    ListView4.ColumnHeaders(7).Width = 0
    
    
    
    SQL = "Select sartic.codartic,sartic.nomartic,sartic.preciouc,sartic.preciove,s2.codartic art2,s2.preciouc upc2,s2.preciove pvp2"
    SQL = SQL & " from sartic left join ariges" & "1" & ".sartic s2 ON sartic.codartic=s2.codartic "  'ARIGES 1, A MANO
    
    
    'Si lleva WHERE
    If cadWhere <> "" Then
        'vCampos = Replace(cadWHERE, "sartic.", "s1.")
        vCampos = cadWhere
        SQL = SQL & " WHERE " & vCampos
        vCampos = ""
    End If
    
    SQL = SQL & " ORDER BY sartic.codartic"
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    SQL = ""

    While Not miRsAux.EOF
        If Not IsNull(miRsAux!art2) Then
            
                
            
            
                'Nuevo articulo
                lblIndicadorCorregir = miRsAux!codartic
                lblIndicadorCorregir.Refresh
    
                
  
            
                SQL = miRsAux!codartic
                vCampos = miRsAux!codartic & "|" & miRsAux!NomArtic & "|"
                UPC1 = DBLet(miRsAux!PrecioUC, "N")
                vCampos = vCampos & Format(UPC1, cadWhere2)
                vCampos = vCampos & "|0|"
                PVP1 = DBLet(miRsAux!preciove, "N")
                UPC2 = DBLet(miRsAux!UPC2, "N")
                PVP2 = DBLet(miRsAux!PVP2, "N")
                
            
            
            
                If PVP1 <> PVP2 Or UPC1 <> UPC2 Then
                    PVP1 = Round2(PVP1, 3)
                    PVP2 = Round2(PVP2, 3)
                    UPC1 = Round2(UPC1, 3)
                    UPC2 = Round2(UPC2, 3)
                    If PVP1 <> PVP2 Or UPC1 <> UPC2 Then
                        vCampos = vCampos & Format(PVP1, cadWhere2) & "|" & Format(PVP2, cadWhere2) & "|"
                        vCampos = vCampos & Format(UPC1, cadWhere2) & "|" & Format(UPC2, cadWhere2) & "|"
                        InsertarItemARticuloConjunto vCampos
                    End If
                End If
        
            
            
            
        End If
        miRsAux.MoveNext
    Wend
    
            

    miRsAux.Close
    Set miRsAux = Nothing
    lblIndicadorCorregir = ""
    Exit Sub
ECargaPVPPreciosDesdeMorales:
    MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
End Sub



Private Sub CargaListaProduccion()

Dim It As ListItem
    
    For NumRegElim = 1 To listacod.Count
        Set It = ListView5.ListItems.Add
       vCampos = listacod.Item(NumRegElim)
        It.Text = RecuperaValor(vCampos, 4)
        It.SubItems(1) = RecuperaValor(vCampos, 5)
        It.SubItems(2) = RecuperaValor(vCampos, 1)
        It.SubItems(3) = RecuperaValor(vCampos, 2)
        It.SubItems(4) = RecuperaValor(vCampos, 3)
    Next
    Me.cmdStockProd(1).Cancel = True
End Sub
