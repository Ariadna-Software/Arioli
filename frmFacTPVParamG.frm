VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmFacTPVParamG 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Par�metros generales TPV"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   9270
   Icon            =   "frmFacTPVParamG.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   360
      TabIndex        =   15
      Top             =   5880
      Width           =   3000
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   210
         Width           =   2280
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   360
      TabIndex        =   17
      Top             =   600
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   9128
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Valores por defecto"
      TabPicture(0)   =   "frmFacTPVParamG.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1(4)"
      Tab(0).Control(1)=   "imgBuscar(8)"
      Tab(0).Control(2)=   "Label1(7)"
      Tab(0).Control(3)=   "imgBuscar(9)"
      Tab(0).Control(4)=   "Text1(8)"
      Tab(0).Control(5)=   "Text2(8)"
      Tab(0).Control(6)=   "Text1(9)"
      Tab(0).Control(7)=   "Text2(9)"
      Tab(0).Control(8)=   "chkRapido"
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Cab./Pie Ticket"
      TabPicture(1)   =   "frmFacTPVParamG.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1(2)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(0)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Text1(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Text1(2)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Text1(3)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Text1(4)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Text1(5)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Text1(6)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Text1(7)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "chkBasesImp"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "chkImprtick"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).ControlCount=   11
      Begin VB.CheckBox chkRapido 
         Caption         =   "TPV con entrada r�pida"
         Height          =   255
         Left            =   -74520
         TabIndex        =   26
         Tag             =   "C|T|S|||spatpvg|rapido|||"
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   9
         Left            =   -71280
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   25
         Text            =   "Text2"
         Top             =   1440
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   320
         Index           =   9
         Left            =   -72240
         MaxLength       =   35
         TabIndex        =   1
         Tag             =   "Cta prev. cobro|T|S|||spatpvg|ctabanc1|||"
         Text            =   "Text1"
         Top             =   1440
         Width           =   855
      End
      Begin VB.CheckBox chkImprtick 
         Caption         =   "Imprimir ticket"
         Height          =   375
         Left            =   480
         TabIndex        =   12
         Tag             =   "Imprimir ticket|N|N|||spatpvg|imprtick|||"
         Top             =   4200
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   8
         Left            =   -71760
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   23
         Text            =   "Text2"
         Top             =   900
         Width           =   4095
      End
      Begin VB.CheckBox chkBasesImp 
         Caption         =   "Mostrar bases imponibles"
         Height          =   375
         Left            =   480
         TabIndex        =   11
         Tag             =   "Mostrar bases imponibles|N|N|||spatpvg|basesimp|||"
         Top             =   3720
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   7
         Left            =   1320
         MaxLength       =   40
         TabIndex        =   10
         Tag             =   "Linea pie 2|T|S|||spatpvg|pietick2|||"
         Text            =   "Text1"
         Top             =   3120
         Width           =   5655
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   6
         Left            =   1320
         MaxLength       =   40
         TabIndex        =   9
         Tag             =   "Linea pie 1|T|S|||spatpvg|pietick1|||"
         Text            =   "Text1"
         Top             =   2760
         Width           =   5655
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   5
         Left            =   1320
         MaxLength       =   40
         TabIndex        =   8
         Tag             =   "Linea cabecera 5|T|S|||spatpvg|cabtick5|||"
         Text            =   "Text1"
         Top             =   2160
         Width           =   5655
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   4
         Left            =   1320
         MaxLength       =   40
         TabIndex        =   7
         Tag             =   "Linea cabecera 4|T|S|||spatpvg|cabtick4|||"
         Text            =   "Text1"
         Top             =   1800
         Width           =   5655
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   3
         Left            =   1320
         MaxLength       =   40
         TabIndex        =   6
         Tag             =   "Linea cabecera 3|T|S|||spatpvg|cabtick3|||"
         Text            =   "Text1"
         Top             =   1440
         Width           =   5655
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   2
         Left            =   1320
         MaxLength       =   40
         TabIndex        =   5
         Tag             =   "Linea cabecera 2|T|S|||spatpvg|cabtick2|||"
         Text            =   "Text1"
         Top             =   1080
         Width           =   5655
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   1
         Left            =   1320
         MaxLength       =   40
         TabIndex        =   4
         Tag             =   "Linea cabecera 1|T|S|||spatpvg|cabtick1|||"
         Text            =   "Text1"
         Top             =   720
         Width           =   5655
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   320
         Index           =   8
         Left            =   -72720
         MaxLength       =   35
         TabIndex        =   0
         Tag             =   "Cliente|N|N|||spatpvg|codclien|000000||"
         Text            =   "Text1"
         Top             =   900
         Width           =   855
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   9
         Left            =   -72600
         Tag             =   "-1"
         ToolTipText     =   "Buscar art�culo"
         Top             =   1460
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta prevista de cobro"
         Height          =   255
         Index           =   7
         Left            =   -74520
         TabIndex        =   24
         Top             =   1440
         Width           =   1800
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   -73080
         Tag             =   "-1"
         ToolTipText     =   "Buscar art�culo"
         Top             =   915
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Pie"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   21
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Cabecera"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   19
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente por defecto"
         Height          =   255
         Index           =   4
         Left            =   -74520
         TabIndex        =   18
         Top             =   900
         Width           =   1410
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   600
      MaxLength       =   15
      TabIndex        =   20
      Tag             =   "C�digo Par�metros TPV|N|N|||spatpvg|codigo||S|"
      Text            =   "Text1"
      Top             =   960
      Width           =   645
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7875
      TabIndex        =   3
      Top             =   6045
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6600
      TabIndex        =   2
      Top             =   6045
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7875
      TabIndex        =   13
      Top             =   6045
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   9270
      _ExtentX        =   16351
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
      Left            =   3600
      Top             =   6120
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
   Begin VB.Label Label1 
      Caption         =   "C�digo"
      Height          =   255
      Index           =   6
      Left            =   600
      TabIndex        =   22
      Top             =   720
      Width           =   495
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
Attribute VB_Name = "frmFacTPVParamG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmMtoCli As frmFacClientes
Attribute frmMtoCli.VB_VarHelpID = -1
Private WithEvents frmMtoBanPr As frmFacBancosPropios
Attribute frmMtoBanPr.VB_VarHelpID = -1

Private NombreTabla As String  'Nombre de la tabla o de la
Private Ordenacion As String
Private CadenaConsulta As String

Dim Modo As Byte
'0: Inicial
'2: Visualizacion
'3: A�adir
'4: Modificar






Private Sub chkBasesImp_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkBasesImp_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub chkBasesImp_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub chkImprtick_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkImprtick_KeyDown(KeyCode As Integer, Shift As Integer)
     KEYdown KeyCode
End Sub

Private Sub chkImprtick_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub




Private Sub chkRapido_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkRapido_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub chkRapido_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub cmdAceptar_Click()

    If Modo = 4 Then 'MODIFICAR
        If DatosOk Then
            
            If ModificaDesdeFormulario(Me, 1) Then
                TerminaBloquear

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
    Screen.MousePointer = vbDefault
    If Modo <> 4 Then
        PonerCadenaBusqueda
        PonerFoco Text1(8)
    End If
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    'Icono de busqueda
    Me.imgBuscar(8).Picture = frmPpal.imgListComun.ListImages(19).Picture
    Me.imgBuscar(9).Picture = frmPpal.imgListComun.ListImages(19).Picture
    
    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 3   'Anyadir
        .Buttons(1).Image = 4   'Modificar
        .Buttons(4).Image = 15  'Salir
    End With
    
    LimpiarCampos   'Limpia los campos TextBox
    
    Me.SSTab1.Tab = 0
'
    NombreTabla = "spatpvg"
    Ordenacion = " ORDER BY codigo"
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = Conn
    CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    
'    Label1(3).Caption = PonerNombreImpresora
'    CargaComboPuertos
'    CargaComboVelocidad
    
    PonerModo 0

End Sub




'Private Sub CargaComboPuertos()
'    'Carga la lista del Combo de Puertos
'    Me.cboPuertos.Clear
'
'
'    cboPuertos.AddItem "COM1"
'    cboPuertos.ItemData(cboPuertos.NewIndex) = 1
'
'    cboPuertos.AddItem "COM2"
'    cboPuertos.ItemData(cboPuertos.NewIndex) = 2
'
'    cboPuertos.AddItem "COM3"
'    cboPuertos.ItemData(cboPuertos.NewIndex) = 3
'
'    cboPuertos.AddItem "COM4"
'    cboPuertos.ItemData(cboPuertos.NewIndex) = 4
'
'    cboPuertos.AddItem "COM5"
'    cboPuertos.ItemData(cboPuertos.NewIndex) = 5
'
''    cboPuertos.ListIndex = 1
'End Sub




'Private Sub CargaComboVelocidad()
'    'Carga la lista del Combo de Puertos
'    Me.cboVelocidad.Clear
'
'    cboVelocidad.AddItem "9600"
'    cboVelocidad.ItemData(cboVelocidad.NewIndex) = 0
'
'    cboVelocidad.AddItem "14400"
'    cboVelocidad.ItemData(cboVelocidad.NewIndex) = 1
'
'    cboVelocidad.AddItem "19200"
'    cboVelocidad.ItemData(cboVelocidad.NewIndex) = 2
'
'    cboVelocidad.AddItem "28800"
'    cboVelocidad.ItemData(cboVelocidad.NewIndex) = 3
'
'    cboVelocidad.AddItem "38400"
'    cboVelocidad.ItemData(cboVelocidad.NewIndex) = 4
'
'    cboVelocidad.AddItem "56000"
'    cboVelocidad.ItemData(cboVelocidad.NewIndex) = 5
'
'    cboVelocidad.AddItem "128000"
'    cboVelocidad.ItemData(cboVelocidad.NewIndex) = 6
'
'    cboVelocidad.AddItem "256000"
'    cboVelocidad.ItemData(cboVelocidad.NewIndex) = 7
''    cboPuertos.ListIndex = 1
'End Sub



Private Sub PonerCadenaBusqueda()
    On Error GoTo EEPonerBusq
    
    Screen.MousePointer = vbHourglass

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        LimpiarCampos
        MsgBox "No hay ning�n registro en la tabla " & NombreTabla, vbInformation
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


Private Sub frmMtoBanPr_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Bancos Propios
    Text1(9).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    Text2(9).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoCli_DatoSeleccionado(CadenaSeleccion As String)
    'Mantenimiento de Clientes
    Text1(8).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000") 'Cod cliente
    Text2(8).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre cliente
End Sub

Private Sub imgBuscar_Click(Index As Integer)
    Select Case Index
        Case 8 'cod. cliente
            Set frmMtoCli = New frmFacClientes
            frmMtoCli.DatosADevolverBusqueda = "1"
            frmMtoCli.Show vbModal
            Set frmMtoCli = Nothing
            PonerFoco Text1(Index)
        
        Case 9 'Bancos propios (cta prev. cobro)
            Set frmMtoBanPr = New frmFacBancosPropios
            frmMtoBanPr.DatosADevolverBusqueda = "0|1|"
            frmMtoBanPr.Show vbModal
            Set frmMtoBanPr = Nothing
            PonerFoco Text1(Index)
    End Select
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
    
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    Select Case Index
        Case 8 'Codl.cliente
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sclien", "nomclien", "codclien", "Clientes")
        Case 9
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sbanpr", "nombanpr", "codbanpr", "Bancos Propios")
                If Text1(Index).Text <> "" Then Text1(Index).Text = Format(Text1(Index).Text, "0000")
            Else
                Text2(Index).Text = ""
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
    PonerFoco Text1(8)
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
    DatosOk = False
    b = CompForm(Me, 1)
    DatosOk = b
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
    
    'El combo del numpuerto lo ponermos en el numpuerto -1
    'Me.cboPuertos.ListIndex = Me.cboPuertos.ListIndex - 1
    
    'poner descripcion del articulo
    Text2(8).Text = PonerNombreDeCod(Text1(8), conAri, "sclien", "nomclien", "codclien", "Clientes")
    Text2(9).Text = PonerNombreDeCod(Text1(9), conAri, "sbanpr", "nombanpr", "codbanpr", "Bancos propios")
    
    BloquearChecks Me, Modo
    
'    PosicionarCombo Me.cboVelocidad, Text1(9).Text
    
    
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
    'Si estamos en Insertar adem�s limpia los campos Text1 y bloquea la clave primaria
    BloquearText1 Me, Modo
    
    'Bloquear los checkbox
    BloquearChecks Me, Modo

'    b = (Modo = 0) Or (Modo = 2)
'    Me.FrameImpresoras.Enabled = (Not b)
'    Me.FrameVisor.Enabled = (Not b)
'    Me.cboPuertos.Enabled = (Not b)
'    Me.cboVelocidad.Enabled = (Not b)
    
    'Bloquear imagen de Busqueda
    Me.imgBuscar(8).Enabled = (Modo = 4)
    Me.imgBuscar(9).Enabled = (Modo = 4)
    
    PonerModoOpcionesMenu 'Activar opciones de menu seg�n el Modo
    PonerOpcionesMenu   'Activar opciones de menu seg�n nivel
                        'de permisos del usuario
End Sub


Private Sub PonerModoOpcionesMenu()
Dim b As Boolean
    b = (Modo = 3) Or (Modo = 4)
    Me.Toolbar1.Buttons(1).Enabled = Not b 'Modificar
    Me.mnModificar.Enabled = Not b
End Sub


