VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAlmContab 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Familias de Artículos"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   Icon            =   "frmAlmContab.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Digitos 1er nivel|N|N|||empresa|numdigi1|||"
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   9
      Left            =   2160
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   38
      Text            =   "Text2"
      Top             =   2040
      Width           =   3885
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   8
      Left            =   2160
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   37
      Text            =   "Text2"
      Top             =   1560
      Width           =   3885
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   1
      Left            =   2160
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   36
      Text            =   "Text2"
      Top             =   1080
      Width           =   3885
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   0
      Left            =   2160
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   35
      Text            =   "Text2"
      Top             =   600
      Width           =   3885
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   9
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   3
      Tag             =   "Tipo articulo|T|N|||sfamcontab|codtipar||S|"
      Text            =   "Text"
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   8
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   2
      Tag             =   "Unidad|N|N|0|9999|sfamcontab|codunida|0000|S|"
      Text            =   "Text"
      Top             =   1560
      Width           =   735
   End
   Begin VB.Frame Frame3 
      Caption         =   "Compras "
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
      Left            =   240
      TabIndex        =   28
      Top             =   4800
      Width           =   7695
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   3
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   8
         Tag             =   "Cta. Contable compras|T|N|||sfamcontab|ctacompr||N|"
         Text            =   "Text1"
         Top             =   240
         Width           =   1125
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   5
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   9
         Tag             =   "Cta.Abono Compras|T|N|||sfamcontab|abocompr||N|"
         Text            =   "Text1"
         Top             =   675
         Width           =   1125
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   5
         Left            =   3600
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   30
         Text            =   "Text2"
         Top             =   675
         Width           =   3885
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   3
         Left            =   3600
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   29
         Text            =   "Text2"
         Top             =   240
         Width           =   3885
      End
      Begin VB.Label Label1 
         Caption         =   "Cta. Abono Compras"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   32
         Top             =   675
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Cta. Contable Compras"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   31
         Top             =   270
         Width           =   1815
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   1
         Left            =   2040
         Picture         =   "frmAlmContab.frx":000C
         ToolTipText     =   "Buscar cuenta contable"
         Top             =   315
         Width           =   240
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   3
         Left            =   2040
         Picture         =   "frmAlmContab.frx":010E
         ToolTipText     =   "Buscar cuenta contable"
         Top             =   720
         Width           =   240
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ventas "
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
      Height          =   2055
      Left            =   240
      TabIndex        =   19
      Top             =   2640
      Width           =   7695
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   2
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "Cta. Contable Ventas|T|N|||sfamcontab|ctaventa||N|"
         Text            =   "Text1"
         Top             =   240
         Width           =   1125
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   4
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   5
         Tag             =   "Cta. Abono Ventas|T|N|||sfamcontab|aboventa||N|"
         Text            =   "Text1"
         Top             =   675
         Width           =   1125
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   4
         Left            =   3600
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   23
         Text            =   "Text2"
         Top             =   675
         Width           =   3885
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   2
         Left            =   3600
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   22
         Text            =   "Text2"
         Top             =   240
         Width           =   3885
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   6
         Left            =   3600
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   21
         Text            =   "Text2"
         Top             =   1080
         Width           =   3885
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   7
         Left            =   3600
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   20
         Text            =   "Text2"
         Top             =   1485
         Width           =   3885
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   7
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   7
         Tag             =   "Cta. Alternativa Abonos|T|N|||sfamcontab|abovent1||N|"
         Text            =   "Text1"
         Top             =   1485
         Width           =   1125
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   6
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   6
         Tag             =   "Cta. Alternativa Ventas|T|N|||sfamcontab|ctavent1||N|"
         Text            =   "Text1"
         Top             =   1080
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Cta. Abono Ventas"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   27
         Top             =   675
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Cta. Contable Ventas"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   26
         Top             =   270
         Width           =   1575
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   0
         Left            =   2040
         Picture         =   "frmAlmContab.frx":0210
         Tag             =   "-1"
         ToolTipText     =   "Buscar cuenta contable"
         Top             =   285
         Width           =   240
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   2
         Left            =   2040
         Picture         =   "frmAlmContab.frx":0312
         ToolTipText     =   "Buscar cuenta contable"
         Top             =   705
         Width           =   240
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   5
         Left            =   2040
         Picture         =   "frmAlmContab.frx":0414
         ToolTipText     =   "Buscar cuenta contable"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   4
         Left            =   2040
         Picture         =   "frmAlmContab.frx":0516
         Tag             =   "-1"
         ToolTipText     =   "Buscar cuenta contable"
         Top             =   1125
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cta. Alternativa Ventas"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   25
         Top             =   1110
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Cta. Alternativa Abonos"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   24
         Top             =   1515
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   6570
      TabIndex        =   11
      Top             =   6360
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   1320
      MaxLength       =   25
      TabIndex        =   1
      Tag             =   "Marca|N|N|0||sfamcontab|codmarca||S|"
      Text            =   "Text1"
      Top             =   1080
      Width           =   765
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   0
      Tag             =   "Categoria|N|N|0|9999|sfamcontab|codfamia|0000|S|"
      Text            =   "Text"
      Top             =   600
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   120
      TabIndex        =   13
      Top             =   6195
      Width           =   2535
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   2355
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6570
      TabIndex        =   12
      Top             =   6360
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5400
      TabIndex        =   10
      Top             =   6360
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   270
      Top             =   6435
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
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
      TabIndex        =   17
      Top             =   0
      Width           =   8040
      _ExtentX        =   14182
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
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
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   5880
         TabIndex        =   18
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Image imgCuentas 
      Height          =   240
      Index           =   9
      Left            =   1080
      Picture         =   "frmAlmContab.frx":0618
      ToolTipText     =   "Buscar familia"
      Top             =   2040
      Width           =   240
   End
   Begin VB.Image imgCuentas 
      Height          =   240
      Index           =   8
      Left            =   1080
      Picture         =   "frmAlmContab.frx":101A
      ToolTipText     =   "Buscar familia"
      Top             =   1560
      Width           =   240
   End
   Begin VB.Image imgCuentas 
      Height          =   240
      Index           =   7
      Left            =   1080
      Picture         =   "frmAlmContab.frx":1A1C
      ToolTipText     =   "Buscar familia"
      Top             =   1080
      Width           =   240
   End
   Begin VB.Image imgCuentas 
      Height          =   240
      Index           =   6
      Left            =   1080
      Picture         =   "frmAlmContab.frx":241E
      ToolTipText     =   "Buscar familia"
      Top             =   600
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Modelo"
      Height          =   195
      Index           =   9
      Left            =   360
      TabIndex        =   34
      Top             =   2040
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "Formato"
      Height          =   195
      Index           =   8
      Left            =   360
      TabIndex        =   33
      Top             =   1560
      Width           =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Marca"
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   16
      Top             =   1080
      Width           =   450
   End
   Begin VB.Label Label1 
      Caption         =   "Categoria"
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   15
      Top             =   600
      Width           =   675
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver todos"
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
Attribute VB_Name = "frmAlmContab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmB2 As frmBuscaGrid 'Form para busquedas
Attribute frmB2.VB_VarHelpID = -1

'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'  Variables comunes a todos los formularios
Private Modo As Byte
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
Private ModoAnterior As Byte

Private CadenaConsulta As String
Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private kCampo As Integer

Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1

'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean





Private Sub cmdAceptar_Click()
    
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then PosicionarData
            End If
        
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me, 1) Then
                    TerminaBloquear
                    PosicionarData
                End If
            End If
    End Select
        
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 1 'Busqueda
            LimpiarCampos
            PonerModo 0
        
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
            PonerModo 2
            PonerCampos
    End Select
    PonerFoco Text1(0)
End Sub


Private Sub BotonAnyadir()
    LimpiarCampos
    'Añadiremos el boton de aceptar y demas objetos para insertar
    ModoAnterior = Modo 'Para el botón cancelar en Modo Insertar
    PonerModo 3
    
    PonerFoco Text1(0)
End Sub


Private Sub BotonBuscar()
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Text1(0).BackColor = vbYellow
    Else 'Modo=1 Busqueda
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
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla
        PonerCadenaBusqueda
    End If
End Sub


Private Sub Desplazamiento(Index As Integer)
'Para desplazarse por los registros de control Data
    DesplazamientoData Data1, Index
    PonerCampos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub


Private Sub BotonModificar()
    
    'El registro de codigo 0 no se puede Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    PonerFoco Text1(2)
End Sub


Private Sub BotonEliminar()
Dim Cad As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    If Modo <> 2 Then Exit Sub
    
    'El registro de codigo 0 no se puede Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    
    
    '### a mano
    Cad = "¿Seguro que desea eliminar los datos de contabilizazion?:" & vbCrLf
    For NumRegElim = 0 To 9
        If NumRegElim < 2 Or NumRegElim > 7 Then
            Cad = Cad & vbCrLf & Me.Label1(NumRegElim).Caption & ":  " & Mid(Text1(NumRegElim).Text & Space(15), 1, 15) & "   " & Text2(NumRegElim).Text
        End If
    Next
    

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
        MuestraError Err.Number, "Eliminar Familia de Articulo", Err.Description
    End If
End Sub


Private Sub cmdRegresar_Click()
Dim Cad As String

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If

    Cad = Data1.Recordset.Fields(0) & "|"
    Cad = Cad & Data1.Recordset.Fields(1) & "|"
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmppal.Icon
    If vParamAplic.Descriptores Then Me.Caption = "Categorias Art."
    ' ICONITOS DE LA BARRA
    btnPrimero = 14 'Posicion del Boton Primero en la toolbar (+ 3 siguientes)
    With Me.Toolbar1
        .ImageList = frmppal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        
        .Buttons(5).Image = 3   'Insertar Nuevo
        .Buttons(6).Image = 4   'Modificar
        .Buttons(7).Image = 5   'Borrar
        
        .Buttons(10).Image = 16  ' Imprimir
        .Buttons(11).Image = 15  'Salir
        
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With
    
    LimpiarCampos   'Limpia los campos TextBox
    
    'Pone el Tag del primer botón de busqueda de cuentas a -1
    'Si tag =-1 abre busqueda en la tabla: sfamia, BD: Ariges
    'Si tag>0 abre busqueda en la tabla: Cuentas, BD: Conta
    imgCuentas(0).Tag = "-1"
        
  
    '## A mano
    NombreTabla = "sfamcontab"
    Ordenacion = " ORDER BY 1,2,3,4"
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
       
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = Conn
    Data1.RecordSource = "Select * from " & NombreTabla & " WHERE codfamia=-1"
    Data1.Refresh
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        BotonBuscar
    End If
    
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano


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
            'Se llama desde un botón de busqueda de Cuentas
            'Recuperar solo el campo código y Descripción
            HaDevueltoDatos = True
            Screen.MousePointer = vbHourglass
    
            Indice = Val(Me.imgCuentas(0).Tag)
            Text1(Indice + 2).Text = RecuperaValor(CadenaDevuelta, 1)
            Text2(Indice + 2).Text = RecuperaValor(CadenaDevuelta, 2)
        Else
            'Recupera todo el registro de Banco Propio
            HaDevueltoDatos = True
            Screen.MousePointer = vbHourglass
            'Sabemos que campos son los que nos devuelve
            'Creamos una cadena consulta y ponemos los datos
            cadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
            cadB = Aux
            '   Como la clave principal es unica, con poner el sql apuntando
            '   al valor devuelto sobre la clave ppal es suficiente
            'Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
            'If CadB <> "" Then CadB = CadB & " AND "
            'CadB = CadB & Aux
            'Se muestran en el mismo form
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
            Screen.MousePointer = vbDefault
        End If
    End If
End Sub


Private Sub frmB2_Selecionado(CadenaDevuelta As String)
    HaDevueltoDatos = True
    Text1(Val(Me.imgCuentas(0).Tag)).Text = RecuperaValor(CadenaDevuelta, 1)
    Text2(Val(Me.imgCuentas(0).Tag)).Text = RecuperaValor(CadenaDevuelta, 2)
End Sub

Private Sub imgCuentas_Click(Index As Integer)

    If Modo = 2 Or Modo = 5 Or Modo = 0 Then Exit Sub
 
    Screen.MousePointer = vbHourglass
    imgCuentas(0).Tag = Index
    If Index >= 6 Then
        'ctegor marca formato modelo
        If Modo <> 4 Then MandaBusquedaPrevia2 Index
    Else
        MandaBusquedaPrevia "apudirec='S'"
        PonerFoco Text1(Index + 2)
    End If
    imgCuentas(0).Tag = -1
    Screen.MousePointer = vbDefault
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


Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not (Index = 0 And KeyCode = 38) Then KEYdown KeyCode
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
Dim C As String
Dim B As Boolean
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    'Si queremos hacer algo ..
    Select Case Index
               
        
        Case 0, 1, 8, 9
            'Familia / categoria         Marca    tipo ud
            C = ""
            If Index = 9 Then
                B = Text1(Index).Text <> ""
            Else
                B = PonerFormatoEntero(Text1(Index))
            End If
            
            If B Then
            
                If Index = 0 Then
                    C = DevuelveDesdeBD(conAri, "nomfamia", "sfamia", "codfamia", Text1(Index).Text)
                ElseIf Index = 1 Then
                    C = DevuelveDesdeBD(conAri, "nommarca", "smarca", "codmarca", Text1(Index).Text)
                ElseIf Index = 8 Then
                    C = DevuelveDesdeBD(conAri, "nomunida", "sunida", "codunida", Text1(Index).Text)
                
                Else
                    C = DevuelveDesdeBD(conAri, "nomtipar", "stipar", "codtipar", Text1(Index).Text, "T")
                End If
                    
                If C = "" Then
                    MsgBox "No existe valor en la base de datos", vbExclamation
                    Text1(Index).Text = ""
                    PonerFoco Text1(Index)
                End If
            Else
                Text1(Index).Text = ""
            End If

            Text2(Index).Text = C
            
        
        
    End Select
End Sub


Private Sub HacerBusqueda()
Dim cadB As String
    
    cadB = ObtenerBusqueda(Me, False)

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
Dim CargaF As Boolean 'Para saber si se carga el frame o no en el BuscaGrid
Dim Conexion As Byte

        'Llamamos a al form
        '##A mano
        Cad = ""
        If Val(Me.imgCuentas(0).Tag) >= 0 Then
        'Se llama a Busqueda desde un campo de Cuenta
            '#A MANO: Porque busca en la tabla Cuentas
            'de la base de datos de Contabilidad
            Cad = Cad & "Código|Cuentas|codmacta|T||15·Denominacion|Cuentas|nommacta|T||70·"
            Tabla = "Cuentas"
            Titulo = "Cuentas"
            Conexion = conConta    'Conexión a BD: Conta
            CargaF = True 'Se puede cargar el frame
        Else
            'Busqueda de una Família de Artículo
            Cad = Cad & ParaGrid(Text1(0), 15, "Código")
            Cad = Cad & ParaGrid(Text1(1), 80, "Denominacion")
            Tabla = "sfamia"
            Titulo = "Família de Artículos"
            If vParamAplic.Descriptores Then Titulo = "Categorias Art."
            Conexion = conAri    'Conexión a BD: Ariges
            CargaF = False 'No se carga el frame
        End If
        
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
            frmB.vCargaFrame = CargaF
            '#
            frmB.Show vbModal
            Set frmB = Nothing
            'Si ha puesto valores y tenemos que es formulario de busqueda entonces
            'tendremos que cerrar el form lanzando el evento
            If HaDevueltoDatos Then
                If kCampo < 5 Then PonerFoco Text1(kCampo + 1)
                If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then
                    If Not (Val(Me.imgCuentas(0).Tag) >= 0) Then cmdRegresar_Click
                End If
            Else   'de ha devuelto datos, es decir NO ha devuelto datos
                PonerFoco Text1(kCampo)
'                If Modo = 1 Then
'                    MsgBox "No hay ningún registro en la tabla " & tabla
'                    PonerFoco Text1(0)
'                End If
            End If
        End If
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
    End If

    Screen.MousePointer = vbDefault
    Exit Sub
    
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
Dim I As Byte
    
    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    
    'poner la descripcion de las cuentas
    For I = 2 To 7
        Text2(I).Text = PonerNombreCuenta(Text1(I), Modo)
    Next I
        
    'Modo
    Modo = 3
    Text1_LostFocus 0
    Text1_LostFocus 1
    Text1_LostFocus 8
    Text1_LostFocus 9
    Modo = 2
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim B As Boolean
Dim NumReg As Byte

    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    '-------------------------------------------------
    'Modo 2. Hay datos y estamos visualizandolos
    B = (Kmodo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = B
    Else
        cmdRegresar.visible = False
    End If
    
    'Visualizar flechas de desplazamiento en la toolbar si modo=2
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, B, NumReg
    
    'Poner Boton de Cabecera o Aceptar/Cancelar
    PonerBotonCabecera B Or (Modo = 0)

    'Bloquear Registros si modo distinto de Insertar o Modificar
    BloquearText1 Me, Modo
    

        
    chkVistaPrevia.Enabled = (Modo <= 2)
    
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    
    PonerModoOpcionesMenu 'Activar opciones de menu según MODO
    PonerOpcionesMenu   'Activar opciones de menu según NIVEL
                        'de permisos del usuario
End Sub


Private Sub PonerModoOpcionesMenu()
Dim B As Boolean
On Error Resume Next

    B = (Modo = 2) Or (Modo = 0) Or (Modo = 1)
    'Añadir
    Toolbar1.Buttons(5).Enabled = B
    Me.mnNuevo.Enabled = B
    
    B = (Modo = 2)
    'Modificar
    Toolbar1.Buttons(6).Enabled = B
    mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(7).Enabled = B
    mnEliminar.Enabled = B
    
     '---------------------------------
    B = (Modo >= 3)
    'Buscar
    Toolbar1.Buttons(1).Enabled = Not B
    Me.mnBuscar.Enabled = Not B
    'VerTodos
    Toolbar1.Buttons(2).Enabled = Not B
    Me.mnVerTodos.Enabled = Not B
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Function DatosOk() As Boolean
Dim B As Boolean

    DatosOk = False
    B = CompForm(Me, 1)
    If Not B Then Exit Function
    

    DatosOk = B
End Function


Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
    If Modo = 3 Or Modo = 4 Then
        Select Case Index
            Case 2, 3, 4, 5, 6, 7 'Cuentas
                Text2(Index).Text = PonerNombreCuenta(Text1(Index), Modo, Text1(1).Text)
                If Text1(Index).Text <> "" And Text2(Index).Text = "" Then Cancel = True
        End Select
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1: mnBuscar_Click
        Case 2: mnVerTodos_Click
        Case 5  'Nuevo
                mnNuevo_Click
        Case 6  'Modificar
                mnModificar_Click
        Case 7  'Borrar
                mnEliminar_Click
        Case 10 'Imprimir listado
            BotonImprimir
        Case 11: mnSalir_Click
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


Private Sub PonerBotonCabecera(B As Boolean)
    Me.cmdAceptar.visible = Not B
    Me.cmdCancelar.visible = Not B
    If B Then Me.lblIndicador.Caption = ""
End Sub


Private Sub PosicionarData()
Dim Cad As String, Indicador As String

    Cad = "(codfamia=" & Text1(0).Text & ")"
    If SituarData(Data1, Cad, Indicador) Then
        PonerModo 2
        lblIndicador.Caption = Indicador
    Else
        PonerModo 0
    End If
End Sub


Private Sub BotonImprimir()
Dim cadFormula As String
Dim cadParam As String
Dim NumParam As Byte
   
    cadFormula = ""
    cadParam = ""
    NumParam = 0
    
    '===================================================
    '============ PARAMETROS ===========================
    cadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
    NumParam = NumParam + 1

    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = NumParam
        .SoloImprimir = False
        .EnvioEMail = False
        .opcion = 5
        .Titulo = "Cta contabilizacion"
        .NombreRPT = "rAlmFamContab.rpt"  'Nombre fichero .rpt a Imprimir
        .Show vbModal
    End With
End Sub



Private Sub MandaBusquedaPrevia2(Index As Integer)
Dim Indice As Integer
        'Se llama a Busqueda desde un campo de Cuenta
            '#A MANO: Porque busca en la tabla Cuentas
            'de la base de datos de Contabilidad
            Set frmB2 = New frmBuscaGrid
            
            
            Indice = Index
            Select Case Index
            Case 6
                Indice = 0
                frmB2.vTabla = "sfamia"
                frmB2.vTitulo = "Categoria"
                frmB2.vCampos = "Código||codfamia|T||15·Denominacion||nomfamia|T||70·"
            Case 7
                Indice = 1
                frmB2.vTabla = "smarca"
                frmB2.vTitulo = "Marca"
                frmB2.vCampos = "Código||codmarca|T||15·Denominacion||nommarca|T||70·"
            Case 8
                
                frmB2.vTabla = "sunida"
                frmB2.vTitulo = "Modelo"
                frmB2.vCampos = "Código||codunida|T||15·Denominacion||nomunida|T||70·"
            Case 9
                frmB2.vTabla = "stipar"
                frmB2.vTitulo = "Formato"
                frmB2.vCampos = "Código||codtipar|T||15·Denominacion||nomtipar|T||70·"
            
            End Select
           
            imgCuentas(0).Tag = Indice
            
            
            
        
        
            Screen.MousePointer = vbHourglass

            HaDevueltoDatos = False
            '###A mano
            frmB2.vDevuelve = "0|1|"
            frmB2.vselElem = 1
            frmB2.vConexionGrid = conAri
            frmB2.vCargaFrame = False
            '#
            frmB2.Show vbModal
            Set frmB2 = Nothing
            
        
End Sub


