VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmVallpalets 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de palets"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   8235
   ClipControls    =   0   'False
   Icon            =   "frmVallpalets.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   8235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8235
      _ExtentX        =   14526
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
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprime Movimiento"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
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
         Left            =   6960
         TabIndex        =   7
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   1080
      TabIndex        =   14
      Tag             =   "Articulo|T|N|||smoval|codartic|||"
      Text            =   "Text1"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   1
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "Text2"
      Top             =   1080
      Width           =   3945
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   0
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "Text2"
      Top             =   600
      Width           =   3945
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   5880
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6915
      TabIndex        =   2
      Top             =   5880
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   6915
      TabIndex        =   3
      Top             =   5880
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   5760
      Width           =   2655
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   9
         Top             =   180
         Width           =   2115
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Tag             =   "Cliente|N|N|0||smoval|codigope|0000||"
      Text            =   "Text1"
      Top             =   600
      Width           =   855
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3120
      Top             =   6000
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
   Begin MSComctlLib.ListView ListView1 
      Height          =   4095
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   7223
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Fecha/Hora"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tipo"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Documento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Salida"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Entrada"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Saldo"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   720
      Picture         =   "frmVallpalets.frx":000C
      ToolTipText     =   "Buscar grupo plantilla"
      Top             =   1080
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   720
      Picture         =   "frmVallpalets.frx":010E
      ToolTipText     =   "Buscar grupo plantilla"
      Top             =   600
      Width           =   240
   End
   Begin VB.Label Label3 
      Caption         =   "Articulo"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Cliente"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   495
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
      Left            =   3240
      TabIndex        =   5
      Top             =   5880
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
         Visible         =   0   'False
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         HelpContextID   =   2
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnLineas 
         Caption         =   "&Lineas"
         HelpContextID   =   2
         Shortcut        =   ^L
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
Attribute VB_Name = "frmVallpalets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1

Private kCampo As Integer

Private Modo As Byte

Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1

Dim ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim CadenaConsulta As String
'Dim PrimeraVez As Boolean

Private HaDevueltoDatos As Boolean

Dim ArticulosPalets As String

Dim PrimeraVez  As Boolean

Private Sub chkVistaPrevia_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
On Error GoTo Error1
    
    Screen.MousePointer = vbHourglass
    Select Case Modo
        Case 1 'BUSQUEDA
            HacerBusqueda
     
            
    End Select
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub






Private Sub cmdCancelar_Click()
On Error GoTo ECancelar

    Select Case Modo
        Case 1 'Buscar
            LimpiarCampos
            PonerModo 0
      
    End Select
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub cmdRegresar_Click()
'Este es el boton Cabecera

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then 'modo 5: Lineas Ofertas
        
        PonerModo 2
        Me.lblIndicador.Caption = ""
    End If
End Sub




Private Sub Form_Activate()
    
    If PrimeraVez Then
        PrimeraVez = False
        'Cargamos los datos que se pueden ver
        CargarDatosEnTemporal
        lblIndicador.Caption = ""
        
         
        
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmppal.Icon
    PrimeraVez = True

    btnPrimero = 17 'Posicion del Boton Primero en la toolbar (+ 3 siguientes)
    With Toolbar1
        .ImageList = frmppal.imgListComun
        'ASignamos botones
        .Buttons(1).Image = 1   'Buscar
        .Buttons(2).Image = 2 'Ver Todos
        .Buttons(5).Image = 3 'Añadir
        .Buttons(6).Image = 4 'Modificar
        .Buttons(7).Image = 5 'Eliminar
        
        .Buttons(10).Image = 40 'impriir
        .Buttons(11).Image = 16 'impriir
        

        
        .Buttons(15).Image = 15 'Salir
        .Buttons(btnPrimero).Image = 6 'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Ultimo
    End With
    
    Me.ListView1.SmallIcons = frmppal.imgListComun
    
    LimpiarCampos   'Limpia los campos TextBox

    'PrimeraVez = True
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
 
    CadenaConsulta = "Select * from smoval WHERE codartic = -1" 'No recupera datos
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
        
    Set miRsAux = New ADODB.Recordset
    'ArticulosPalets = "SELECT codartic from sartic,sfamia where sartic.codfamia=sfamia.codfamia and tipfamia =31" 'palets
    ArticulosPalets = "SELECT codartic from sartic where tipartic =31" 'palets
    miRsAux.Open ArticulosPalets, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    ArticulosPalets = ""
    While Not miRsAux.EOF
        ArticulosPalets = ArticulosPalets & ", " & DBSet(miRsAux!codartic, "T")
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    If ArticulosPalets = "" Then
        ArticulosPalets = "-1"
    Else
        ArticulosPalets = Mid(ArticulosPalets, 2) 'quitamos la primera coma
    End If
    
            
    PonerModo 0
    CargaGrid (Modo = 2)
   
End Sub


Private Sub CargaGrid(enlaza As Boolean)
Dim SQL As String



    On Error GoTo ECarga
    
    Screen.MousePointer = vbHourglass

    ListView1.ListItems.Clear

    If enlaza Then
    
        Set miRsAux = New ADODB.Recordset
        
        
        SQL = ""
        If Val(Data1.Recordset!Codigo1) > 0 Then SQL = "(detamovi IN ('ALV','PAL') and codigope=" & Data1.Recordset!Codigo1 & ")"
        If Val(Data1.Recordset!campo2) > 0 Then
            If SQL <> "" Then SQL = SQL & " or "
            SQL = SQL & " (detamovi IN ('PAL','ALC') and codigope=" & Data1.Recordset!campo2 & ")"
        End If
        SQL = " where codartic =" & DBSet(Data1.Recordset!nombre2, "T") & " AND fechamov>=" & DBSet(vEmpresa.FechaIni, "F") & " AND (" & SQL & ")"
        
        SQL = "select * from smoval " & SQL & " ORDER BY horamovi"
        
        miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then CargaItems
        miRsAux.Close
        
    End If
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos lineas ", Err.Description
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaItems()
Dim It As ListItem
Dim Saldo As Long

    Saldo = 0
    While Not miRsAux.EOF
        
            Set It = ListView1.ListItems.Add()
            It.Text = Format(miRsAux!horamovi, "dd/mm/yyyy    hh:nn")
            It.SubItems(1) = miRsAux!detamovi
            It.SubItems(2) = miRsAux!document
            If miRsAux!tipomovi = 0 Then
                Saldo = Saldo - miRsAux!Cantidad
                It.SubItems(3) = Format(miRsAux!Cantidad, "#,##0")
                It.SubItems(4) = " "
            Else
                Saldo = Saldo + miRsAux!Cantidad
                It.SubItems(4) = Format(miRsAux!Cantidad, "#,##0")
                It.SubItems(3) = " "
            End If
            'It.SmallIcon = 43
            It.SubItems(5) = Format(Saldo, "#,##0")
            If Saldo < 0 Then
                It.ListSubItems(5).ForeColor = vbRed
                It.ListSubItems(5).Bold = True
            End If
        miRsAux.MoveNext
    Wend
    If Not It Is Nothing Then It.EnsureVisible
End Sub


'Private Sub LLamaLineas(alto As Single)
''Pone posicion TOP y LEFT de los controles en el form
'Dim jj As Integer
'Dim b As Boolean
'
'
'    DeseleccionaGrid Me.DataGrid1
'
'    'Fijamos el ancho
'    b = (Modo = 5 And ModificaLineas = 1 Or ModificaLineas = 2)
'
'    For jj = 0 To txtAux.Count - 1
'        txtAux(jj).Height = DataGrid1.RowHeight
'        txtAux(jj).Top = alto
'        txtAux(jj).visible = b
'        If b Then txtAux(jj).Text = ""
'    Next jj
'
'    jj = 0
'    Me.cmdAux(jj).Height = Me.DataGrid1.RowHeight
'    Me.cmdAux(jj).Top = alto
'    Me.cmdAux(jj).visible = b
'End Sub



'Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
''Formulario Mantenimiento Articulos
'    txtAux(0).Text = RecuperaValor(CadenaSeleccion, 1)
'    txtAux(1).Text = RecuperaValor(CadenaSeleccion, 2)
'End Sub



Private Sub frmB_Selecionado(CadenaDevuelta As String)
'Formulario para Busqueda
Dim cadB As String
Dim Aux As String

    If CadenaDevuelta <> "" Then
            HaDevueltoDatos = True
            Screen.MousePointer = vbHourglass
            
            'Estamos en Cabecera
            'Recupera todo el registro de Tarifas de Precios
            'Sabemos que campos son los que nos devuelve
            'Creamos una cadena consulta y ponemos los datos
            cadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
            cadB = Aux
            CadenaConsulta = "select * from smoval WHERE " & cadB & " ORDER BY id"
            PonerCadenaBusqueda
    End If
    Screen.MousePointer = vbDefault
End Sub








Private Sub imgBuscar_Click(Index As Integer)

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    Select Case Index
        Case 0  '
          
             PonerFoco Text1(5)
    End Select
   
    Screen.MousePointer = vbDefault
End Sub



Private Sub mnBuscar_Click()
    BotonBuscar
End Sub





Private Sub mnEliminar_Click()
Dim Cad As String

        If ListView1.SelectedItem Is Nothing Then Exit Sub
        If ListView1.SelectedItem.SubItems(1) <> "PAL" Then Exit Sub
            
        Cad = "¿Desea eliminar el movimiento de palet seleccionado?" & vbCrLf & vbCrLf
        If Trim(ListView1.SelectedItem.SubItems(3)) = "" Then
            Cad = Cad & "- ENTRADA " & vbCrLf & "-Unidades: " & ListView1.SelectedItem.SubItems(4)
        Else
            Cad = Cad & "- SALIDA " & vbCrLf & "-Unidades: " & ListView1.SelectedItem.SubItems(3)
        End If
        
        If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        eliminarMovimiento
        CargaGrid True
End Sub

Private Sub mnNuevo_Click()
    
    CadenaDesdeOtroForm = ""
    If Text1(1).Text <> "" Then CadenaDesdeOtroForm = Text1(1).Text & "|" & Text2(1).Text & "|"
    frmListado2.Opcion = 35
    frmListado2.Show vbModal
    If CadenaDesdeOtroForm <> "" Then
        CadenaConsulta = RecuperaValor(CadenaDesdeOtroForm, 1)
        kCampo = 1
        If Data1.Recordset!nombre2 = CadenaConsulta Then
            'Ok es el mismo
            CadenaConsulta = RecuperaValor(CadenaDesdeOtroForm, 2)
            If Val(Data1.Recordset!Codigo1) = Val(CadenaConsulta) Then
                kCampo = 0
            Else
                CadenaConsulta = RecuperaValor(CadenaDesdeOtroForm, 2)
                If Val(Data1.Recordset!campo2) = Val(CadenaConsulta) Then kCampo = 0
            End If
        End If
        If kCampo = 0 Then
            'Solo hay que cargar estos items
            CargaGrid True
        Else
            'Hay que cargarlo todo y situar el DATA
            
            Screen.MousePointer = vbHourglass
            CargarDatosEnTemporal
            
            
            'Refrescamos
            CadenaConsulta = "Select * from tmpinformes  WHERE codusu =" & vUsu.Codigo & " ORDER BY nombre2,codigo1"
            Data1.RecordSource = CadenaConsulta
            Data1.Refresh
            kCampo = 0
            Modo = 0
            While Modo = 0
                If Data1.Recordset.EOF Then
                    Modo = 1
                    
                Else
                    If Data1.Recordset!nombre2 = Text1(1).Text Then
                        'Ok es el mismo
                        If Val(Data1.Recordset!Codigo1) = Val(Text1(0).Text) Then
                            kCampo = 1
                            Modo = 1
                        Else
                            If Val(Data1.Recordset!campo2) = Val(Text1(0).Text) Then
                                Modo = 1
                                kCampo = 1
                            End If
                        End If
                    End If
                    
                    If Modo = 0 Then Data1.Recordset.MoveNext
                End If
            Wend
            If kCampo = 1 Then
                PonerCampos               'Ok. Ya hemos situado el data
            Else
                LimpiarCampos
                PonerModo 0
            End If
            Screen.MousePointer = vbDefault
            
        End If
    End If
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
'    If (Modo = 5) Then 'Modo 5: Mto Lineas
'        '1:Insertar linea, 2: Modificar
'        If ModificaLineas = 1 Or ModificaLineas = 2 Then cmdCancelar_Click
'        cmdRegresar_Click
'        Exit Sub
'    End If
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Text1_LostFocus(Index As Integer)

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    With Text1(Index)
        Select Case Index
            Case 0 'Codigo Plantilla
                If PonerFormatoEntero(Text1(Index)) Then
                    'comprobar si ya existe el codigo de plantilla
                    If Modo = 3 Then 'Insertar
                        If ExisteCP(Text1(Index)) Then PonerFoco Text1(Index)
                    End If
                End If
                
            Case 2, 5 'Codigo
                If PonerFormatoEntero(Text1(Index)) Then
                    If Index = 2 Then
                        Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "svehiculos", "descripcion", "codigo")
                    Else
                        Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "straba", "nomtraba", "codtraba")
                    End If
                    If Text2(Index).Text = "" Then PonerFoco Text1(Index)
                Else
                    Text2(Index).Text = ""
                End If
            Case 3
                PonerFormatoFecha Text1(Index)
        End Select
    End With
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Cad As String
'    'If Button.Index = 10 Or Button.Index = 11 Or Button.Index = 13 Or Button.Index = 14 Then
'    If Button.Index >= 10 And Button.Index <= 15 Then
'        If Data1.Recordset Is Nothing Then Exit Sub
'        If Data1.Recordset.EOF Then Exit Sub
'    End If


    Select Case Button.Index
        Case 1 'Busqueda
            mnBuscar_Click
        Case 2 'Ver Todos
            mnVerTodos_Click
        Case 5 'Nuevo
            mnNuevo_Click
        Case 6  'Modificar
            
        Case 7 'Eliminar
           
            mnEliminar_Click
            
            
        Case 10
        
        Case 11
            If ListView1.SelectedItem Is Nothing Then Exit Sub
            If ListView1.SelectedItem.SubItems(1) <> "PAL" Then Exit Sub
            
            conn.Execute "Delete from tmprutas where codusu=" & vUsu.Codigo
            Set miRsAux = New ADODB.Recordset
            If Val(Data1.Recordset!Codigo1) = 0 Then
                Cad = "Select codprove codigo, nomprove nombre, domprove direc, pobprove pobla from sprove where codprove=" & Data1.Recordset!campo2
            Else
                Cad = "Select codclien codigo,nomclien nombre, domclien direc, pobclien pobla from sclien where codclien=" & Data1.Recordset!Codigo1
            End If
            miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            'tmprutas(codusu,idruta,idalb,nomclien,domclien,pobclien,codartic,nomartic,cajas,fecha2,codigo)  'codigo=1 o 2
            Cad = "(" & vUsu.Codigo & ",'" & ListView1.SelectedItem.SubItems(2) & "'," & miRsAux!Codigo & ","
            Cad = Cad & DBSet(miRsAux!Nombre, "T") & "," & DBSet(miRsAux!direc, "T") & "," & DBSet(miRsAux!pobla, "T") & ","
            Cad = Cad & DBSet(Text1(1).Text, "T") & "," & DBSet(Text2(1).Text, "T") & ","
            If Trim(ListView1.SelectedItem.SubItems(3)) = "" Then
                kCampo = -1 * CInt(ImporteFormateado(ListView1.SelectedItem.SubItems(4)))
            Else
                kCampo = CInt(ImporteFormateado(ListView1.SelectedItem.SubItems(3)))
            End If
            Cad = Cad & kCampo & "," & DBSet(ListView1.SelectedItem.Text, "F") & ","
             
 
            Cad = Cad & "1)"
            Cad = "INSERT INTO tmprutas(codusu,idruta,idalb,nomclien,domclien,pobclien,codartic,nomartic,cajas,fecha2,codigo) VALUES " & Cad
            conn.Execute Cad
            
            miRsAux.Close
            Set miRsAux = Nothing
            
 
 
 
            'tmprutas(codusu,idruta,codigo,nomclien,domclien,pobclien,codartic,nomartic,cajas,fecha2)
 
 
 
            
            frmVarios.Opcion = 13
            frmVarios.Show vbModal
            kCampo = 0
        Case 15
            mnSalir_Click
        Case 21  'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas de Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub


Private Sub KEYpress(KeyAscii As Integer)
'On Error Resume Next
'    If KeyAscii = 13 Then 'ENTER
'        KeyAscii = 0
'        SendKeys "{tab}"
'    ElseIf KeyAscii = 27 Then 'ESC
'        Select Case Modo
'            Case 0, 2: Unload Me
'            Case 1: cmdCancelar_Click 'Buscar
'            Case 5 'Lineas
'                If ModificaLineas = 0 Then PonerModo 2
'        End Select
'    End If
'    If Err.Number <> 0 Then Err.Clear
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte
Dim b As Boolean
Dim NumReg As Byte

    'Actualiza Iconos Insertar,Modificar,Eliminar
    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, 3
    
    Modo = Kmodo
    PonerIndicador Me.lblIndicador, Modo
    
    '===========================================
    'Modo 2. Hay datos y estamos visualizandolos
    b = (Kmodo = 2)
    'Visualizar flechas de desplazamiento en la toolbar si modo=2
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
       
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    
    '==============================
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    cmdRegresar.visible = (Modo = 5)
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Enabled = b
    Next i
    
    

    
    chkVistaPrevia.Enabled = (Modo <= 2)
     
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos

    '===============================
    PonerModoOpcionesMenu 'Activa las Opciones de menu según Modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub

Private Sub PonerModoOpcionesMenu()
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean

    b = (Modo = 2)
    'Modificar
    Toolbar1.Buttons(4).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(6).Enabled = b
    Me.mnEliminar.Enabled = b
    
'    b = (Modo = 2)
'    'Lineas
'    Toolbar1.Buttons(10).Enabled = b
'    Me.mnLineas.Enabled = b
''
'    b = (Modo >= 3)
'    'Insertar
'    Toolbar1.Buttons(5).Enabled = Not b Or (Modo = 5)
'    Me.mnNuevo.Enabled = Not b Or (Modo = 5)
'    'Buscar
'    Toolbar1.Buttons(1).Enabled = Not b
'    Me.mnBuscar.Enabled = Not b
'    'Ver Todos
'    Toolbar1.Buttons(2).Enabled = Not b
'    Me.mnVerTodos.Enabled = Not b
'
'
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es

End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    DesplazamientoData Data1, Index
    PonerCampos
End Sub


Private Function MontaSQLCarga(enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim SQL As String
Dim Tabla As String
    
    Tabla = "slipla"
    SQL = "SELECT codplant,numlinea," & Tabla & ".codartic, sartic.nomartic, cantidad "
    SQL = SQL & " FROM " & Tabla & " LEFT JOIN sartic ON " & Tabla & ".codartic=sartic.codartic"
    If enlaza Then
        SQL = SQL & " WHERE codplant=" & Text1(0).Text 'Data1.Recordset!codPlant
    Else
        SQL = SQL & " WHERE codplant = -1"
    End If
    SQL = SQL & " ORDER BY " & Tabla & ".numlinea "
    MontaSQLCarga = SQL
End Function


Private Sub BotonBuscar()
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        'Ponemos el grid lineasfacturas enlazando a ningun sitio
        CargaGrid False

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

    LimpiarCampos
    'Ponemos el grid lineasfacturas enlazando a ningun sitio
    CargaGrid False
    
  
    
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia C
    Else
        CadenaConsulta = "Select * from tmpinformes  WHERE codusu =" & vUsu.Codigo & " GROUP BY codigo1,campo2,nombre2 ORDER BY nombre2,codigo1"
        PonerCadenaBusqueda
    End If
End Sub






Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
Dim Tabla As String
Dim Titulo As String

    'Llamamos a al form
    Cad = ""
    'Estamos en Modo de Cabeceras
    'Registro de la tabla de cabeceras: scapla
    Cad = Cad & ParaGrid(Text1(0), 10, "Código")
    Cad = Cad & ParaGrid(Text1(3), 14, "Fecha")
    Cad = Cad & ParaGrid(Text1(1), 43, "Observaciones")
    Cad = Cad & "Conductor||conductor|T||33·"
    
    Tabla = "smoval"
    Titulo = "Control palets"
           
    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = Tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|"
        frmB.vTitulo = Titulo
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri 'Conexión a BD: Ariges
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
        If HaDevueltoDatos Then
''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
''                cmdRegresar_Click
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
    cadB = ""
    If Text1(1).Text <> "" Then cadB = "nombre2= " & DBSet(Text1(1).Text, "T")
    If Text1(0).Text <> "" Then
        If cadB <> "" Then cadB = cadB & " AND "
        cadB = cadB & " ( codigo1 = " & Text1(0).Text & " OR  campo2= " & Text1(0).Text & ")"
    End If
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "Select * from tmpinformes  WHERE codusu =" & vUsu.Codigo & " AND " & cadB & " GROUP BY codigo1,campo2,nombre2 ORDER BY nombre2,codigo1"
        PonerCadenaBusqueda
    End If
End Sub


Private Sub PonerCadenaBusqueda()
Dim cadMen As String
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        cadMen = "No hay ningún registro en la tabla "
        If Modo = 1 Then
            MsgBox cadMen & " para ese criterio de Búsqueda.", vbInformation
        Else
            MsgBox cadMen, vbInformation
        End If
        CargaGrid False
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
On Error GoTo EPonerCampos

    If Data1.Recordset.EOF Then Exit Sub
    
    
    
    If Val(Data1.Recordset!Codigo1) = 0 Then
        Text1(0).Text = Format(Data1.Recordset!campo2, "0000")
        Text2(0).Text = DevuelveDesdeBD(conAri, "nomprove", "sprove", "codprove", Data1.Recordset!campo2)
    Else
        Text2(0).Text = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", Data1.Recordset!Codigo1)
        Text1(0).Text = Format(Data1.Recordset!Codigo1, "0000")
    End If
    If Trim(Text2(0).Text) = "" Then Text2(0).Text = "Error"
    
    Text1(1).Text = Data1.Recordset!nombre2
    Text2(1).Text = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", Data1.Recordset!nombre2, "T")
    If Trim(Text2(1).Text) = "" Then Text2(1).Text = "Error"
    
    CargaGrid True
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

    vWhere = " id = " & Text1(0).Text
    If SituarData(Data1, vWhere, Indicador) Then
        PonerModo 2
        Indicador = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        lblIndicador.Caption = Indicador
    Else
        PonerModo 0
    End If
End Sub





Private Sub Imprimir(SinVentanaOk As Boolean)
Dim i As Integer
Dim Cad As String
Dim SQL As String
Dim Litros As Currency

    On Error GoTo EImprimir

    conn.Execute "DELETE FROM tmprutas where codusu = " & vUsu.Codigo

   '`tmprutas`
   'insert into `tmprutas`
   ' (`codusu`,`idruta`,`codigo`,`idalb`,`nomclien`,`domclien`,`pobclien`,`proclien`,
   '`codartic`,`nomartic`,`cajas`,`fecha2`)


    Set miRsAux = New ADODB.Recordset
    NumRegElim = 0
    For i = 1 To ListView1.ListItems.Count
        lblIndicador.Caption = 1 & " de " & ListView1.ListItems.Count
        lblIndicador.Refresh
        If ListView1.ListItems(i).SmallIcon = 43 Then
            '************************************   ALBARANES
            SQL = "select scaalb.codtipom,scaalb.numalbar,scaalb.fechaalb,slialb.codartic,slialb.nomartic,cantidad,cajas,nomclien,"
            SQL = SQL & "domclien, codpobla, pobclien, proclien,LitrosUnidad from slialb,scaalb,sartic where slialb.codtipom=scaalb.codtipom and scaalb.numalbar=slialb.numalbar"
            SQL = SQL & " AND slialb.codartic=sartic.codartic"
            SQL = SQL & " and slialb.codartic<>'" & vParamAplic.ArtReciclado & "'"  'Que no salgal el punto verde
            'Ahora el albaran en cuetion
            SQL = SQL & " AND scaalb.codtipom='" & ListView1.ListItems(i).Text & "' "
            SQL = SQL & " AND scaalb.numalbar=" & ListView1.ListItems(i).SubItems(1) & " "
            miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            SQL = ""
            Cad = ""
            While Not miRsAux.EOF
               ' (`codusu`,`idruta`,,`idalb`,`nomclien`,`domclien`,`pobclien`,`proclien`,`fecha2`
               '`codigo` `codartic`,`nomartic`,`cajas`,)

                NumRegElim = NumRegElim + 1
                If Cad = "" Then
                    Cad = ", (" & vUsu.Codigo & "," & Text1(0).Text & ",'"
                    Cad = Cad & miRsAux!Codtipom & Format(miRsAux!NumAlbar, "0000000") & "',"
                    Cad = Cad & DBSet(miRsAux!nomclien, "T") & "," & DBSet(miRsAux!domclien, "T") & ","
                    Cad = Cad & DBSet(miRsAux!pobclien, "T") & ",'"
                    'cppos, provinci
                    Cad = Cad & DevNombreSQL(Trim(DBLet(miRsAux!codpobla, "T") & "   " & DBLet(miRsAux!proclien, "T"))) & "','"
                    Cad = Cad & Format(miRsAux!FechaAlb, FormatoFecha) & "',"
                End If
                'Faltan: `codigo` `codartic`,`nomartic`,`cajas`,)
                SQL = SQL & Cad & NumRegElim & "," & DBSet(miRsAux!codartic, "T") & ","
                Litros = DBLet(miRsAux!LitrosUnidad, "N")
                Litros = Litros * miRsAux!Cantidad
                SQL = SQL & DBSet(miRsAux!NomArtic, "T") & "," & miRsAux!Cajas & "," & DBSet(Litros, "N") & ")"
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            If SQL <> "" Then
                'Tiene datos
                Cad = "INSERT INTO tmprutas (`codusu`,`idruta`,`idalb`,`nomclien`,`domclien`,"
                Cad = Cad & "`pobclien`,`proclien`,`fecha2`,`codigo`,`codartic`,`nomartic`,`cajas`,litros) VALUES "
                Cad = Cad & Mid(SQL, 2) 'quito la primera coma

                conn.Execute Cad
                
                
                
                
            End If
        Else
            '************************************   FACTURAS
            SQL = "select scafac.codtipom,scafac.numfactu,scafac.fecfactu,slifac.codartic,slifac.nomartic,cantidad,nomclien,"
            SQL = SQL & "domclien, codpobla, pobclien, proclien,unicajas,LitrosUnidad from slifac,scafac,sartic where"
            SQL = SQL & " slifac.codtipom=scafac.codtipom and scafac.numfactu=slifac.numfactu"
            SQL = SQL & " AND slifac.codartic=sartic.codartic"
            SQL = SQL & " and slifac.codartic<>'" & vParamAplic.ArtReciclado & "'"  'Que no salgal el punto verde
            'Ahora el albaran en cuetion
            SQL = SQL & " AND scafac.codtipom='" & ListView1.ListItems(i).Text & "' "
            SQL = SQL & " AND scafac.numfactu=" & ListView1.ListItems(i).SubItems(1) & " "
            SQL = SQL & " AND scafac.fecfactu='" & Format(ListView1.ListItems(i).SubItems(2), FormatoFecha) & "' "
            SQL = SQL & " AND slifac.numalbar=" & ListView1.ListItems(i).Tag & " "
            miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            SQL = ""
            Cad = ""
            While Not miRsAux.EOF
               ' (`codusu`,`idruta`,,`idalb`,`nomclien`,`domclien`,`pobclien`,`proclien`,`fecha2`
               '`codigo` `codartic`,`nomartic`,`cajas`,)

                NumRegElim = NumRegElim + 1
                If Cad = "" Then
                    Cad = ", (" & vUsu.Codigo & "," & Text1(0).Text & ",'"
                    Cad = Cad & miRsAux!Codtipom & Format(miRsAux!NumFactu, "0000000") & "',"
                    Cad = Cad & DBSet(miRsAux!nomclien, "T") & "," & DBSet(miRsAux!domclien, "T") & ","
                    Cad = Cad & DBSet(miRsAux!pobclien, "T") & ",'"
                    'cppos, provinci
                    Cad = Cad & DevNombreSQL(Trim(DBLet(miRsAux!codpobla, "T") & "   " & DBLet(miRsAux!proclien, "T"))) & "','"
                    Cad = Cad & Format(miRsAux!FecFactu, FormatoFecha) & "',"
                End If
                'Faltan: `codigo` `codartic`,`nomartic`,`cajas`,)
                SQL = SQL & Cad & NumRegElim & "," & DBSet(miRsAux!codartic, "T") & ","
                SQL = SQL & DBSet(miRsAux!NomArtic, "T") & ","
                If DBLet(miRsAux!Unicajas, "N") = 0 Then
                    SQL = SQL & Round(miRsAux!Cantidad, 0)
                Else
                    SQL = SQL & CStr(CInt(miRsAux!Cantidad) \ CInt(miRsAux!Unicajas))
                End If
                Litros = DBLet(miRsAux!LitrosUnidad, "N")
                Litros = Litros * miRsAux!Cantidad
                SQL = SQL & "," & DBSet(Litros, "N")
                SQL = SQL & ")"
                    
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            If SQL <> "" Then
                'Tiene datos
                Cad = "INSERT INTO tmprutas (`codusu`,`idruta`,`idalb`,`nomclien`,`domclien`,"
                Cad = Cad & "`pobclien`,`proclien`,`fecha2`,`codigo`,`codartic`,`nomartic`,`cajas`,litros) VALUES "
                Cad = Cad & Mid(SQL, 2) 'quito la primera coma

                conn.Execute Cad
            End If
        
        
        End If
    
    
    Next
    
    If NumRegElim > 0 Then
            
            Cad = DevuelveNombreReport(40)
            
    
            With frmImprimir
                .FormulaSeleccion = "{tmprutas.codusu} = " & vUsu.Codigo
                .OtrosParametros = ""
                .NumeroParametros = 0
        
                If SinVentanaOk Then
                    .SoloImprimir = True
                    .NumeroDeCopias = 1
                Else
                    .SoloImprimir = False
                End If
                .EnvioEMail = False
                .Opcion = 2016
                .Titulo = Me.Caption
                .NombreRPT = Cad
                .ConSubInforme = True
                .Show vbModal
            End With

    End If
    
    Exit Sub
EImprimir:
    MuestraError Err.Number
End Sub



Private Sub CargarDatosEnTemporal()
    
    lblIndicador.Caption = "Leyendo datos palets"
        lblIndicador.Refresh
        
        conn.Execute "DELETE FROM tmpinformes where codusu =" & vUsu.Codigo
        conn.Execute "DELETE FROM tmppartidas where codusu =" & vUsu.Codigo
        
        CadenaConsulta = "insert into tmppartidas(codusu,codartic,idpartida,idReferencia)"
        CadenaConsulta = CadenaConsulta & " select distinct " & vUsu.Codigo & ",codartic,if(dEtamovi='ALV',codigope,0),if(detamovi='ALC',codigope,0)"
        CadenaConsulta = CadenaConsulta & " from smoval where fechamov >=" & DBSet(vEmpresa.FechaIni, "F") & " AND  codartic in (" & ArticulosPalets & ") "
        CadenaConsulta = CadenaConsulta & " and detamovi in ('ALC','ALV')"
        conn.Execute CadenaConsulta
        
        CadenaConsulta = "insert into tmppartidas(codusu,codartic,idpartida,idReferencia)"
        CadenaConsulta = CadenaConsulta & " select distinct " & vUsu.Codigo & ",codartic,if(tipomovi=0,codigope,0),if(tipomovi=1,codigope,0)"
        CadenaConsulta = CadenaConsulta & " from smoval where fechamov >=" & DBSet(vEmpresa.FechaIni, "F") & " AND  codartic in (" & ArticulosPalets & ") "
        CadenaConsulta = CadenaConsulta & " and detamovi ='PAL'"
        conn.Execute CadenaConsulta
        
        
        CadenaConsulta = "insert into tmpinformes(codusu,nombre2,codigo1,campo2)"
        CadenaConsulta = CadenaConsulta & " select codusu,codartic,idpartida,idreferencia from tmppartidas where codusu = " & vUsu.Codigo & " group by 1,2,3,4"
        conn.Execute CadenaConsulta
        
        Set miRsAux = New ADODB.Recordset
        CadenaConsulta = "select codclien,codprove from sprove,sclien where nifprove =nifclien"
        miRsAux.Open CadenaConsulta, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not miRsAux.EOF
            lblIndicador = miRsAux!CodClien & " - " & miRsAux!codProve
            lblIndicador.Refresh
            'Son 23 updates aprox
            CadenaConsulta = "UPDATE tmpinformes set campo2=" & miRsAux!codProve
            CadenaConsulta = CadenaConsulta & " WHERE codigo1=" & miRsAux!CodClien & " AND campo2=0 AND codusu =" & vUsu.Codigo
            conn.Execute CadenaConsulta
            
           
             CadenaConsulta = "UPDATE tmpinformes set codigo1=" & miRsAux!CodClien
            CadenaConsulta = CadenaConsulta & " WHERE campo2=" & miRsAux!codProve & " AND codigo1=0 AND codusu =" & vUsu.Codigo
            conn.Execute CadenaConsulta
            
            
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        
        
        
        Set miRsAux = Nothing

        


End Sub




Private Sub eliminarMovimiento()
Dim cSt As cStock
    
    
    Set cSt = New cStock
    
    cSt.Documento = Format(Val(ListView1.SelectedItem.SubItems(2)), "0000")
   
    cSt.codAlmac = 1
    cSt.DetaMov = "PAL"
    cSt.codartic = Text1(1).Text
    cSt.Fechamov = CDate(ListView1.SelectedItem.Text)
    'cSt.HoraMov = cSt.Fechamov & " " & txtHora(0).Text
    cSt.Importe = 0
    cSt.LineaDocu = 1
    If Trim(ListView1.SelectedItem.SubItems(3)) = "" Then
        cSt.tipoMov = "S"  'Era una entrada. Ahora es salida
        cSt.Cantidad = Val(ListView1.SelectedItem.SubItems(4))
    Else
        cSt.tipoMov = "E"   'Era una salida. Ahora es entrada
        cSt.Cantidad = Val(ListView1.SelectedItem.SubItems(3))
    End If
    cSt.DevolverStock2
    
    Set cSt = Nothing
End Sub
