VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmRegAccCorrec 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acciones correctoras"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   Icon            =   "frmRegAccCorrec.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   8
      Left            =   1200
      MaxLength       =   100
      TabIndex        =   3
      Tag             =   "P|T|S|||srevacccorrect|lugar|||"
      Text            =   "Text1"
      Top             =   1200
      Width           =   5445
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   7
      Left            =   1200
      MaxLength       =   100
      TabIndex        =   8
      Tag             =   "P|T|S|||srevacccorrect|plazo|||"
      Text            =   "Text1"
      Top             =   5520
      Width           =   5445
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   6
      Left            =   1200
      MaxLength       =   30
      TabIndex        =   6
      Tag             =   "Responsable|N|S|||srevacccorrect|idtrabRespon||N|"
      Text            =   "Text1"
      Top             =   4200
      Width           =   645
   End
   Begin VB.TextBox txtTrabDesc 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   6
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "Text5"
      Top             =   4200
      Width           =   3495
   End
   Begin VB.TextBox txtTrabDesc 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   3
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "Text5"
      Top             =   720
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   3
      Left            =   2400
      MaxLength       =   30
      TabIndex        =   2
      Tag             =   "Quien detecta ave.|N|N|||srevacccorrect|idTraba||N|"
      Text            =   "Text1"
      Top             =   720
      Width           =   645
   End
   Begin VB.TextBox Text1 
      Height          =   675
      Index           =   5
      Left            =   1200
      MaxLength       =   255
      TabIndex        =   7
      Tag             =   "A|T|S|||srevacccorrect|AccionCorrectora|||"
      Text            =   "Text1"
      Top             =   4680
      Width           =   5445
   End
   Begin VB.TextBox Text1 
      Height          =   675
      Index           =   4
      Left            =   1200
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   5
      Tag             =   "a|T|S|||srevacccorrect|Causas|||"
      Text            =   "frmRegAccCorrec.frx":000C
      Top             =   3360
      Width           =   5445
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   120
      MaxLength       =   15
      TabIndex        =   0
      Tag             =   "ID|N|N|||srevacccorrect|id||S|"
      Text            =   "Text1"
      Top             =   720
      Width           =   885
   End
   Begin VB.TextBox Text1 
      Height          =   1395
      Index           =   2
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Tag             =   "Desc|T|N|||srevacccorrect|Descripcion|||"
      Text            =   "frmRegAccCorrec.frx":0012
      Top             =   1800
      Width           =   5445
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   5490
      TabIndex        =   10
      Top             =   6120
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   1200
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "Fecha|F|N|||srevacccorrect|fecha|dd/mm/yyyy||"
      Text            =   "Text1"
      Top             =   720
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   120
      TabIndex        =   12
      Top             =   6000
      Width           =   2535
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   210
         Width           =   2115
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5490
      TabIndex        =   11
      Top             =   6120
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4200
      TabIndex        =   9
      Top             =   6120
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   30
      Top             =   6075
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
      TabIndex        =   16
      Top             =   0
      Width           =   6870
      _ExtentX        =   12118
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
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
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "�ltimo"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   5160
         TabIndex        =   17
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Lugar"
      Height          =   375
      Index           =   8
      Left            =   120
      TabIndex        =   26
      Top             =   1200
      Width           =   975
   End
   Begin VB.Image imgFec 
      Height          =   240
      Index           =   1
      Left            =   1680
      Picture         =   "frmRegAccCorrec.frx":0018
      Top             =   480
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Plazo"
      Height          =   375
      Index           =   6
      Left            =   120
      TabIndex        =   25
      Top             =   5520
      Width           =   975
   End
   Begin VB.Image imgTra 
      Height          =   240
      Index           =   6
      Left            =   960
      Picture         =   "frmRegAccCorrec.frx":05A2
      Top             =   4200
      Width           =   240
   End
   Begin VB.Image imgTra 
      Height          =   240
      Index           =   3
      Left            =   3360
      Picture         =   "frmRegAccCorrec.frx":0FA4
      Top             =   480
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Trab. detecta"
      Height          =   195
      Index           =   3
      Left            =   2400
      TabIndex        =   21
      Top             =   480
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "Accion correctora"
      Height          =   615
      Index           =   7
      Left            =   120
      TabIndex        =   20
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Causas posibles"
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   19
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Descripci�n"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   18
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   15
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "ID"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Responsa."
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   24
      Top             =   4200
      Width           =   975
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
Attribute VB_Name = "frmRegAccCorrec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmT As frmAdmTrabajadores
Attribute frmT.VB_VarHelpID = -1


'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin ningun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'  Variables comunes a todos los formularios
Private Modo As Byte

Private CadenaConsulta As String
Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Private btnPrimero As Byte
'Numero del boton donde empiezan las flecha de desplazamiento de registros

Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos


Private Sub cmdAceptar_Click()
Dim cad As String
Dim Indicador As String

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    Select Case Modo
        Case 1  'BUSCAR
            HacerBusqueda
        Case 3  'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then PonerModo 0
            End If
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me, 1) Then
                    TerminaBloquear
                    cad = "(id=" & Text1(0).Text & ")"
                    If SituarData(Data1, cad, Indicador) Then
                        PonerModo 2
                        lblIndicador.Caption = Indicador
                    Else
                        LimpiarCampos
                        PonerModo 0
                    End If
                End If
            End If
    End Select
        
Error1:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 1, 3 '1:Buscar / 3:Insertar
            LimpiarCampos
            PonerModo 0
        Case 4  'Modificar
            lblIndicador.Caption = ""
            TerminaBloquear
            PonerModo 2
            PonerCampos
    End Select
End Sub


Private Sub BotonAnyadir()
    LimpiarCampos
    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    Text1(0).Text = SugerirCodigoSiguienteStr("srevacccorrect", "id")
    Text1(1).Text = Format(Now, "dd/mm/yyyy")
    PonerFoco Text1(1)
End Sub


Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then 'Modo 1: Buscar
        LimpiarCampos
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
    
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
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
    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    'Como el campo 1 es clave primaria, NO se puede modificar
    BloquearTxt Text1(0), True
    PonerFoco Text1(1)
End Sub


Private Sub BotonEliminar()
Dim cad As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    
    cad = "Va a eliminar la accion correctora?" & vbCrLf
    cad = cad & vbCrLf & "C�digo: " & Data1.Recordset.Fields(0)
    cad = cad & vbCrLf & "Fecha: " & Data1.Recordset.Fields(1)
    'desc
    cad = cad & vbCrLf & "Descripcion: " & DBLet(Data1.Recordset!Descripcion, "T")
    cad = cad & vbCrLf & vbCrLf & "�Continuar?"
    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
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
        MuestraError Err.Number, "Eliminar Clientes Varios", Err.Description
    End If
End Sub


Private Sub cmdRegresar_Click()
Dim cad As String

    If Data1.Recordset.EOF Then
        MsgBox "Ning�n registro devuelto.", vbExclamation
        Exit Sub
    End If
    
    cad = Data1.Recordset.Fields(0) & "|"
    cad = cad & Data1.Recordset.Fields(1) & "|"
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmppal.Icon

    ' ICONITOS DE LA BARRA
    btnPrimero = 13 'Boton donde empiezan las Flechas de desplazamiento de Registros
    With Me.Toolbar1
        .ImageList = frmppal.imgListComun
        .Buttons(1).Image = 1   'Bot�n Buscar
        .Buttons(2).Image = 2   'Bot�n Todos
        .Buttons(5).Image = 3   'Insertar Nuevo
        .Buttons(6).Image = 4   'Modificar
        .Buttons(7).Image = 5   'Borrar
        .Buttons(9).Image = 40 'Imprimir
        .Buttons(10).Image = 15  'Salir
        .Buttons(13).Image = 6  'Primero
        .Buttons(14).Image = 7  'Anterior
        .Buttons(15).Image = 8  'Siguiente
        .Buttons(16).Image = 9  '�ltimo
    End With
    
    LimpiarCampos
    VieneDeBuscar = False
    
    '## A mano
    NombreTabla = "srevacccorrect"
    Ordenacion = " ORDER BY id"
        
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = Conn
    Data1.RecordSource = "Select * from " & NombreTabla & " where id=-1"
    Data1.Refresh
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        BotonBuscar
    End If
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox del form
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
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
        'Sabemos que campos son los que nos devuelve
        'Creamos una cadena consulta y ponemos los datos
        cadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        cadB = Aux
        '   Como la clave principal es unica, con poner el sql apuntando
        '   al valor devuelto sobre la clave ppal es suficiente
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub



Private Sub frmC_Selec(vFecha As Date)
    CadenaConsulta = vFecha
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
    CadenaConsulta = CadenaSeleccion
End Sub

Private Sub imgFec_Click(Index As Integer)
    CadenaConsulta = ""
    Set frmC = New frmCal
    frmC.Fecha = Now
    If Text1(Index).Text <> "" Then frmC.Fecha = CDate(Text1(Index).Text)
    frmC.Show vbModal
    Set frmC = Nothing
    If CadenaConsulta <> "" Then
        Text1(Index).Text = Format(CadenaConsulta, "dd/mm/yyyy")
        PonerFoco Text1(3)
        CadenaConsulta = ""
    End If
End Sub

Private Sub imgTra_Click(Index As Integer)
    CadenaConsulta = ""
    Set frmT = New frmAdmTrabajadores
    frmT.DatosADevolverBusqueda = "0|1|"
    frmT.Show vbModal
    Set frmT = Nothing
    If CadenaConsulta <> "" Then
        Text1(Index).Text = RecuperaValor(CadenaConsulta, 1)
        txtTrabDesc(Index).Text = RecuperaValor(CadenaConsulta, 2)
        If Index = 3 Then
            PonerFoco Text1(8)
        Else
            PonerFoco Text1(5)
        End If
        CadenaConsulta = ""
    End If
    
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
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Index <> 2 Then KEYdown KeyCode
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
    
    Text1(Index).Text = Trim(Text1(Index).Text)
    
    Select Case Index
        Case 0 '
            
            If Text1(Index).Text <> "" Then PonerFormatoEntero Text1(Index)
            
            
        Case 1 '
            If Text1(Index).Text <> "" Then
                devuelve = Text1(Index).Text
                If EsFechaOK(devuelve) Then Text1(Index).Text = devuelve
            End If
            
        Case 3, 6
            devuelve = ""
                
            If Text1(Index).Text <> "" Then
                If PonerFormatoEntero(Text1(Index)) Then
                    devuelve = DevuelveDesdeBD(conAri, "nomtraba", "straba", "codtraba", Text1(Index).Text, "N")
                    If devuelve = "" Then
                        MsgBox "No se ha encontrado el trabajador: " & Text1(Index).Text, vbExclamation
                        Text1(Index).Text = ""
                        PonerFoco Text1(Index)
                    End If
            
                Else
                    Text1(Index).Text = ""
                    PonerFoco Text1(Index)
                End If
            End If
            Me.txtTrabDesc(Index).Text = devuelve
            
    End Select
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


Private Sub MandaBusquedaPrevia(cadB As String)
Dim cad As String
        'Llamamos a al form
        '##A mano
        cad = ""
        cad = cad & ParaGrid(Text1(0), 9, "ID")
        cad = cad & ParaGrid(Text1(1), 17, "Fecha")
        cad = cad & ParaGrid(Text1(2), 70, "Descrp.")
        If cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vCampos = cad
            frmB.vTabla = NombreTabla
            frmB.vSQL = cadB
            HaDevueltoDatos = False
            '###A mano
            frmB.vDevuelve = "0|1|" 'Campos de la tabla que devuelve
            frmB.vTitulo = "Acciones correctoras"
            frmB.vselElem = 1
            frmB.vConexionGrid = 1 'Conexi�n a BD: Ariges
'            frmB.vBuscaPrevia = chkVistaPrevia
            frmB.Show vbModal
            Set frmB = Nothing
            'Si ha puesto valores y tenemos que es formulario de busqueda entonces
            'tendremos que cerrar el form lanzando el evento
            If HaDevueltoDatos Then
'                If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'                    cmdRegresar_Click
            Else   'de ha devuelto datos, es decir NO ha devuelto datos
                PonerFoco Text1(kCampo)
            End If
        End If
End Sub


Private Sub PonerCadenaBusqueda()
On Error GoTo EEPonerBusq

    Screen.MousePointer = vbHourglass
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        If Modo = 1 Then 'Busqueda
             MsgBox "No hay ning�n registro en la tabla " & NombreTabla & " para ese criterio de B�squeda.", vbInformation
             PonerFoco Text1(0)
        Else
            MsgBox "No hay ning�n registro en la tabla " & NombreTabla, vbInformation
        End If
         Screen.MousePointer = vbDefault
         Exit Sub
    Else
        PonerModo 2
        'Data1.Recordset.MoveLast
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

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    
    Modo = 3
    Text1_LostFocus 3
    Text1_LostFocus 6
    Modo = 2
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim b As Boolean
Dim NumReg As Byte

    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    '--------------------------------------------------
    'Modo 2. Hay datos y estamos visualizandolos
    b = (Kmodo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = b
    Else
        cmdRegresar.visible = False
    End If
    
    'Visualizar flechas de desplazamiento en la toolbar si modo=2
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar adem�s limpia los campos Text1
    BloquearText1 Me, Modo
    
    
    '-----------------------------------------------------
    'Modo insertar o modificar
    b = (Kmodo >= 3) Or Modo = 1 '-->Luego not b sera kmodo<3
    cmdAceptar.visible = b
    cmdCancelar.visible = b
    BloquearImg Me.imgTra(3), Not b
    BloquearImg Me.imgTra(6), Not b
    BloquearImg Me.imgFec(1), Not b
    
    If cmdCancelar.visible Then
        cmdCancelar.Cancel = True
    Else
        cmdCancelar.Cancel = False
    End If
        
    Me.chkVistaPrevia.Enabled = (Modo <= 2)

    PonerModoOpcionesMenu 'Activar opciones de menu seg�n modo
    PonerOpcionesMenu   'Activar opciones de menu seg�n nivel
                        'de permisos del usuario
End Sub


Private Sub PonerModoOpcionesMenu()
Dim b As Boolean
    
    b = (Modo = 2)
    'Modificar
    Toolbar1.Buttons(6).Enabled = b
    mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(7).Enabled = b
    mnEliminar.Enabled = b
    
    '---------------------------------------------
    b = (Modo >= 3)
     'Insertar
    Toolbar1.Buttons(5).Enabled = Not b
    Me.mnNuevo.Enabled = Not b
    'Buscar
    Toolbar1.Buttons(1).Enabled = Not b
    Me.mnBuscar.Enabled = Not b
    'Ver Todos
    Toolbar1.Buttons(2).Enabled = Not b
    Me.mnVerTodos.Enabled = Not b
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean

    DatosOk = False
    b = CompForm(Me, 1) 'Comprobar datos OK
    If Not b Then Exit Function
        

    DatosOk = b
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
        Case 9
            If Modo <> 2 And Modo <> 0 Then Exit Sub
            If Data1.Recordset.EOF Then Exit Sub
            CadenaConsulta = "{srevacccorrect.id}=" & Text1(0).Text
            
            LlamaImprimirGral CadenaConsulta, "", 0, "morAccCorrectora.rpt", "Accion correctora: " & Text1(0).Text & " - " & Text1(1).Text
            CadenaConsulta = ""
        Case 10  'Salir
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
