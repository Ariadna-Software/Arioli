VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmProdTrazaVer2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Trazabilidad"
   ClientHeight    =   10575
   ClientLeft      =   45
   ClientTop       =   1035
   ClientWidth     =   13845
   Icon            =   "frmProdTrazaVer2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10575
   ScaleWidth      =   13845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Digitos 1er nivel|N|N|||empresa|numdigi1|||"
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Tag             =   "Lote|T|N|||spartidas|numlote|||"
      Text            =   "Text1"
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      Tag             =   "Articulo|T|N|||spartidas|codartic|||"
      Text            =   "Text1"
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   8160
      TabIndex        =   2
      Tag             =   "Partida|N|N|0||spartidas|id|||"
      Text            =   "Text1"
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   9840
      TabIndex        =   3
      Tag             =   "Cantidad|N|N|||spartidas|cantotal|0.00||"
      Text            =   "Text1"
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   375
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Text2"
      Top             =   840
      Width           =   4215
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   12540
      TabIndex        =   5
      Top             =   10035
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   150
      TabIndex        =   7
      Top             =   9915
      Width           =   3135
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   2715
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   12600
      TabIndex        =   6
      Top             =   10035
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   11280
      TabIndex        =   4
      Top             =   10035
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   150
      Top             =   10155
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
      TabIndex        =   9
      Top             =   0
      Width           =   13845
      _ExtentX        =   24421
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
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   5280
         TabIndex        =   10
         Top             =   120
         Width           =   1215
      End
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   5295
      Left            =   240
      TabIndex        =   12
      Top             =   1320
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   9340
      _Version        =   327682
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   3015
      Left            =   240
      TabIndex        =   13
      Top             =   6840
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   5318
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Factura"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Fecha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Cod."
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Cliente"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Cantidad"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Cantidad"
      Height          =   255
      Index           =   4
      Left            =   9840
      TabIndex        =   18
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Partida"
      Height          =   255
      Index           =   3
      Left            =   8160
      TabIndex        =   17
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Descripcion"
      Height          =   255
      Index           =   2
      Left            =   3720
      TabIndex        =   16
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Cod. art."
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   15
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Lote"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   14
      Top             =   600
      Width           =   1695
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Visible         =   0   'False
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
Attribute VB_Name = "frmProdTrazaVer2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1


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

'-------------------------------------------------------------------------
'-------------------------------------------------------------------------

Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1

Private CadenaConsulta As String
Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private Kcampo As Integer



Private DatosImpresionGenerados As Boolean



Private Sub cmdAceptar_Click()
Dim Cad As String, Indicador As String

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    Select Case Modo
        Case 1  'BUSCAR
            HacerBusqueda
        Case 3  'INSERTAR
         
        Case 4  'MODIFICAR
           
    End Select
        
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 1, 3
            LimpiarCampos
            PonerModo 0
        Case 4
            'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
    End Select
End Sub





Private Sub BotonBuscar()
    If Modo <> 1 Then 'Modo 1: Buscar
        LimpiarCampos
        PonerModo 1
        '### A mano
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Text1(0).BackColor = vbYellow
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(Kcampo).Text = ""
            Text1(Kcampo).BackColor = vbYellow
            PonerFoco Text1(Kcampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
    'Ver todos
    LimpiarCampos
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia " " & NombreTabla & ".codartic IN (select codartic from sartic WHERE conjunto=1)"
    Else
        CadenaConsulta = "Select * from " & NombreTabla & " WHERE spartidas.codartic IN (select codartic from sartic WHERE conjunto=1)"
        PonerCadenaBusqueda
    End If
End Sub


Private Sub Desplazamiento(Index As Integer)
'Para desplazarse por los registros de control Data
    DesplazamientoData Data1, Index
    PonerCampos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
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

Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmppal.Icon

    ' ICONITOS DE LA BARRA
    btnPrimero = 13
    With Me.Toolbar1
        .ImageList = frmppal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        .Buttons(5).Image = 3   'Insertar Nuevo
        .Buttons(6).Image = 4   'Modificar
        .Buttons(7).Image = 5   'Borrar
        .Buttons(9).Image = 16  'Salir
        .Buttons(10).Image = 15  'Salir
        .Buttons(13).Image = 6  'Primero
        .Buttons(14).Image = 7  'Anterior
        .Buttons(15).Image = 8  'Siguiente
        .Buttons(16).Image = 9  'Último
    End With
    
    LimpiarCampos
    

    '## A mano
    NombreTabla = "spartidas"
    Ordenacion = " ORDER BY id"
           
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    '## A mano
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
    TreeView1.Nodes.Clear
    ListView2.ListItems.Clear
End Sub




Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub



    


Private Sub frmB_Selecionado(CadenaDevuelta As String)
    CadenaConsulta = CadenaDevuelta
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
   ' BotonEliminar
End Sub

Private Sub mnModificar_Click()
  '  If BLOQUEADesdeFormulario(Me) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
  '  BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbHourglass
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    Kcampo = Index
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
   
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    'Si queremos hacer algo ..
    Select Case Index
        Case 2 'Cod Forma de Pago
           PonerFormatoEntero Text1(Index)
                
           
            
     
       
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
Dim Cad As String

        'Llamamos a al form
        '##A mano
        Cad = ParaGrid(Text1(0), 20, "Lote")
        Cad = Cad & ParaGrid(Text1(1), 19, "Articulo")
        Cad = Cad & "Referencia|sartic|nomartic|T||50·"
        Cad = Cad & ParaGrid(Text1(2), 11, "Partida")
        
        If cadB <> "" Then cadB = " AND " & cadB
        cadB = "sartic.codartic = spartidas.codartic" & cadB
        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vCampos = Cad
            frmB.vTabla = NombreTabla & ",sartic"
            frmB.vSQL = cadB
            
            CadenaConsulta = ""
            frmB.vDevuelve = "3|" 'Campos de la tabla que devuelve
            frmB.vTitulo = "Trazabilidad"
            frmB.vselElem = 0
            frmB.vConexionGrid = 1 'Conexión a BD: Ariges
            frmB.vCargaFrame = False
            frmB.Show vbModal
            Set frmB = Nothing
        
            If CadenaConsulta <> "" Then
                Cad = "id = " & RecuperaValor(CadenaConsulta, 1)
                CadenaConsulta = "select * from " & NombreTabla & " WHERE " & Cad & " " & Ordenacion
                PonerCadenaBusqueda
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
             MsgBox "No hay ningún registro en la tabla " & NombreTabla & " para ese criterio de Búsqueda.", vbInformation
             PonerFoco Text1(0)
        Else
            MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        End If
'         MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
         Screen.MousePointer = vbDefault
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
    PonerCamposForma Me, Me.Data1
  
  
    Text2.Text = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", Text1(1).Text, "T")
  
    PonerCampos2
    DatosImpresionGenerados = False
  
  
  
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount

EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub

Private Sub PonerCampos2()
Dim SQL As String
Dim cP As cPartidas
Dim N
Dim CargaDesdeTmpTraza As Boolean

    PonerCamposForma Me, Data1
    TreeView1.Nodes.Clear
    ListView2.ListItems.Clear
    Set cP = New cPartidas
    conn.Execute "DELETE FROM tmptraza"
    If cP.LeerDesdeArticulo(Text1(1).Text, Data1.Recordset!codalmac, Data1.Recordset!NUmlote) Then
        cP.TrazbilidadDesdeVenta False, False
     
    End If
    
    
    
    Set miRsAux = New ADODB.Recordset
    SQL = DBLet(Data1.Recordset!NumAlbar, "T")
    If SQL <> "" Then
        'AQUI VERE SI ES UN COUPAGE, PRODUCCION u otro
        CargaDesdeTmpTraza = False
        If Val(Data1.Recordset!codProve) = 0 And Mid(SQL, 1, 2) = "NP" Then
            CargaDesdeTmpTraza = True
        Else
            If vParamAplic.QUE_EMPRESA = 4 Then
                If Val(Data1.Recordset!codProve) = 0 And Mid(SQL, 1, 2) = "PR" Then CargaDesdeTmpTraza = True
                If Val(Data1.Recordset!codProve) = 0 And Mid(SQL, 1, 3) = "CUP" Then CargaDesdeTmpTraza = True
                If Val(Data1.Recordset!codProve) = 0 And Mid(SQL, 1, 3) = "TRS" Then CargaDesdeTmpTraza = True
            End If
        End If
        If CargaDesdeTmpTraza Then
                'PRODUCCION
                'Cargar datos produccion
                CargarDatosProduccion
        Else
                SQL = DevuelveAlbaran(Data1.Recordset!NUmlote, Data1.Recordset!codartic)
                
                Set N = TreeView1.Nodes.Add(, , "C" & CStr(TreeView1.Nodes.Count + 1), SQL)
        End If
        
        If TreeView1.Nodes.Count > 1 Then TreeView1.Nodes(2).EnsureVisible
            
    End If
    'Todos cargaran si hay ventas
    CargarDatosVentas
    Set miRsAux = Nothing
    Set cP = Nothing
End Sub




'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim I As Byte
Dim b As Boolean
Dim NumReg As Byte

    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    '--------------------------------------------
    'Modo 2. Hay datos y estamos visualizandolos
    b = (Kmodo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = b
        If Modo = 1 Then Me.lblIndicador.Caption = "BUSQUEDA"
    Else
        cmdRegresar.visible = False
    End If
    
    'Visualizar flechas de desplazamiento en la toolbar si modo=2
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    
    
    '----------------------------------------------
    'Modo insertar o modificar
    b = (Kmodo >= 3) '-->Luego not b sera kmodo<3
    cmdAceptar.visible = b Or Modo = 1
    cmdCancelar.visible = b Or Modo = 1
    If cmdCancelar.visible Then
        cmdCancelar.Cancel = True
    Else
        cmdCancelar.Cancel = False
    End If
    
    BloquearText1 Me, Modo
    
    'Formas de Pago
    'For i = 0 To Text2.Count - 1
    '    BloquearTxt Text2(i), True
    'Next i
    
    
    b = (Modo = 3) 'Insertar
    'Campos Importe Mínimo y % Adelantado
    If b Then
        For I = 8 To 9
            BloquearTxt Text1(I), True
        Next I
    End If

     chkVistaPrevia.Enabled = (Modo <= 2)

    PonerModoOpcionesMenu 'Activar opciones de menu según modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
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
    
    b = (Modo >= 3)
    'Insertar
    Toolbar1.Buttons(5).Enabled = Not b
    Me.mnNuevo.Enabled = Not b
    'Buscar
    Toolbar1.Buttons(1).Enabled = Not b
    Me.mnBuscar.Enabled = Not b
    'VerTodos
    Toolbar1.Buttons(2).Enabled = Not b
    Me.mnVerTodos.Enabled = Not b
End Sub






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
            Imprimir
            
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




Private Sub CargarArbol(padre, NivelPintando As Integer)
Dim N
Dim C As String
Dim Aux As String
Dim contador As Integer
Dim Fin As Boolean
Dim NivelActual As Integer
   
Dim NOdoErroneo As Boolean
Dim Fin2 As Boolean
    
            Fin = False
            Do
                
                If Not miRsAux.EOF Then
                    If NivelPintando = miRsAux!nivle Then
                        
                        C = miRsAux!artic2 & " " & miRsAux!NomArtic & " [" & miRsAux!NUmlote2 & "]"
                        If vParamAplic.QUE_EMPRESA = 4 Then
                            If miRsAux!nivle > 1 Then
                                C = miRsAux!idoperacion & " [" & miRsAux!NUmlote2 & "]"
                                If InStr(1, miRsAux!idoperacion, "Id:") > 0 And InStr(1, miRsAux!idoperacion, "Dep:") > 0 Then
                                  C = "MOLT." & miRsAux!idoperacion & " [" & miRsAux!NUmlote2 & "]"
                                End If
                            End If
                        End If
                         
                        Aux = ""
                        If Mid(miRsAux!idoperacion, 3, 1) = "/" Then
                            If Mid(miRsAux!idoperacion, 6, 1) = "/" Then
                                Aux = "MOLT"
                                C = miRsAux!NUmlote2 & "(" & miRsAux!Cantidad & ") " & miRsAux!nomclien
                            End If
                        End If
                        If Aux = "" Then
                            Aux = DevuelveAlbaran(miRsAux!NUmlote2, miRsAux!artic2)
                            C = DevuelveCadena(C, Aux, NivelPintando)
                        End If
                       
                        NOdoErroneo = False
                        If LCase(Mid(miRsAux!idoperacion, 1, 3)) = "err" Then
                            'NO LO PINTO
                            NOdoErroneo = True
                            'Y ademas, lo borro
                            C = "DELETE from tmptraza where codusu =" & vUsu.Codigo & " AND contador =" & miRsAux!contador
                            conn.Execute C
                               
                            
                        End If
                        
                        If Not NOdoErroneo Then
                            
                            contador = TreeView1.Nodes.Count + 1
                            Set N = TreeView1.Nodes.Add(padre, tvwChild, "C" & contador, C)
                            N.Tag = miRsAux!contador 'Clave
                           
                        End If
                        
                        NivelActual = miRsAux!nivle
                        miRsAux.MoveNext
                    Else
                        'Stop
                    End If
                End If
                If miRsAux.EOF Then
                    Fin = True
                Else
                
                    If miRsAux!nivle > NivelActual Then
                        'Si habia nodo erroneo, entonces tengo que matar los sub nodos ya que no los voy a pintar
                        If Not NOdoErroneo Then
                        CargarArbol N, miRsAux!nivle
                        Fin = False
                        
                        Else
                            Fin2 = False
                            While Not Fin2
                                C = "DELETE from tmptraza where codusu =" & vUsu.Codigo & " AND contador =" & miRsAux!contador
                                conn.Execute C
                                miRsAux.MoveNext
                                
                                If miRsAux.EOF Then
                                    Fin2 = True
                                    Fin = True
                                Else
                                    If miRsAux!nivle > NivelActual Then
                                        'QUe siga en este borrado
                                    Else
                                        Fin2 = True
                                    End If
                                End If
                            Wend
                        End If
                    Else
                        
                        Fin = True
                    End If
                End If
            Loop Until Fin
        
End Sub



Private Sub CargarDatosProduccion()
Dim C As String
Dim N
Dim contador As Integer
Dim Nivel As Integer
Dim padre As String
Dim Aux As String


    

    C = "select tmptraza.*,nomartic from tmptraza,sartic where codartic=artic2 AND codusu =" & vUsu.Codigo
    miRsAux.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Nivel = -1
        
    
    C = miRsAux!artic2 & " " & miRsAux!NomArtic & " [" & miRsAux!NUmlote2 & "]"
    Aux = DevuelveAlbaran(miRsAux!NUmlote2, miRsAux!artic2)
    C = DevuelveCadena(C, Aux, 0)

                
    contador = TreeView1.Nodes.Count + 1
    Set N = TreeView1.Nodes.Add(, , "C" & contador, C)
                
    Nivel = 0
    miRsAux.MoveNext
    If Not miRsAux.EOF Then
        Nivel = miRsAux!nivle
        CargarArbol N, Nivel
        'Llamamos a un recursivo para cargar el arbol
    End If
            
    miRsAux.Close
    
    If Not N Is Nothing Then N.EnsureVisible
    
    'If ElAceite <> "" Then CargaCoupageRecursivo RecuperaValor(ElAceite, 1), RecuperaValor(ElAceite, 2), N.Key, EsCou
    
End Sub


Private Sub CargarDatosVentas()
Dim C As String
Dim It
    C = "select concat(scafac.codtipom,scafac.numfactu) lafact,scafac.fecfactu,codclien,nomclien,cantidad "
    C = C & " from slifaclotes,scafac  where"
    C = C & " slifaclotes.codTipoM = scafac.codTipoM And slifaclotes.NumFactu = scafac.NumFactu"
    C = C & " and slifaclotes.fecfactu=scafac.fecfactu AND numlote=" & DBSet(Text1(0).Text, "T")
    C = C & " ORDER BY fecfactu,lafact"
    miRsAux.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set It = ListView2.ListItems.Add()
        It.Text = miRsAux!lafact
        It.SubItems(1) = miRsAux!FecFactu
        It.SubItems(2) = Format(miRsAux!CodClien, "0000")
        It.SubItems(3) = miRsAux!nomclien
        It.SubItems(4) = Format(miRsAux!Cantidad, "#,##0")

        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    'EN albaranes
    C = "select concat(scaalb.codtipom,scaalb.numalbar) lafact,scaalb.fechaalb,codclien,nomclien,cantidad "
    C = C & " from slialblotes,scaalb where"
    C = C & " slialblotes.codTipoM = scaalb.codTipoM And slialblotes.numalbar = scaalb.numalbar"
    C = C & "  AND numlote=" & DBSet(Text1(0).Text, "T")
    C = C & " ORDER BY fechaalb,lafact"
    miRsAux.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set It = ListView2.ListItems.Add()
        It.Text = miRsAux!lafact
        It.SubItems(1) = miRsAux!FechaAlb
        It.SubItems(2) = Format(miRsAux!CodClien, "0000")
        It.SubItems(3) = miRsAux!nomclien
        It.SubItems(4) = Format(miRsAux!Cantidad, "#,##0")

        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
End Sub



Private Function DevuelveAlbaran(NUmlote As String, vArtic As String) As String
Dim RT As ADODB.Recordset
Dim Cad As String
Dim PalWhere As String  'numalbar
    DevuelveAlbaran = ""
    Set RT = New ADODB.Recordset
    Cad = "select * from spartidas where numlote=" & DBSet(NUmlote, "T") & " and codartic='" & vArtic & "'"
    RT.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    If Not RT.EOF Then
        
        Cad = "select nomprove ,scafpc.numfactu idDoc ,scafpc.fecfactu fecha from scafpc,slifpc where scafpc.codprove=slifpc.codprove and"
        Cad = Cad & " scafpc.numfactu=slifpc.numfactu and scafpc.fecfactu=slifpc.fecfactu"
        Cad = Cad & " AND slifpc.numalbar=" & DBSet(RT!NumAlbar, "T") & " and codartic=" & DBSet(RT!codartic, "T")
        Cad = Cad & " AND scafpc.codprove =" & RT!codProve
        
    End If
    RT.Close
    
        
    If Cad <> "" Then
        RT.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If RT.EOF Then
            
            RT.Close
            
            Cad = Mid(RT.Source, InStr(1, UCase(RT.Source), "WHERE") + 6)
            'Reemplazamos
            
            Cad = Replace(Cad, "scafpc", "scaalp")
            Cad = Replace(Cad, "slifpc", "slialp")
            Cad = Replace(Cad, "fecfactu", "fechaalb")
            Cad = Replace(Cad, "numfactu", "numalbar")
            Cad = " from scaalp,slialp where " & Cad
            Cad = "select nomprove ,scaalp.numalbar idDoc ,scaalp.fechaalb fecha " & Cad
            
            
            RT.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            
            
        End If
        
        If Not RT.EOF Then
            DevuelveAlbaran = "Alb: " & RT!iddoc & "  " & RT!Fecha & "   " & RT!nomprove
            
            
        
            
        End If
        RT.Close
    End If
    
    Set RT = Nothing
End Function


Private Function DevuelveCadena(Cadena As String, cad2 As String, Nivel As Integer) As String
Dim J As Integer
    
        
    DevuelveCadena = cad2
    J = 124 - (Nivel * 5)
    
    J = J - Len(DevuelveCadena) - Len(Cadena)
    If J < 0 Then J = 0
    DevuelveCadena = Cadena & Space(J) & DevuelveCadena
    
End Function



Private Sub Imprimir()
Dim Producida As Currency
Dim CantidadVenta As Currency
Dim vLote As String

    If Modo <> 2 Then Exit Sub
    If Data1.Recordset.EOF Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    conn.Execute "delete from tmpinformes WHERE codusu =" & vUsu.Codigo
    conn.Execute "delete from tmppartidas WHERE codusu =" & vUsu.Codigo
    
    
    
    Set miRsAux = New ADODB.Recordset
    
    CadenaConsulta = "select * from spartidas where id=" & Text1(2).Text
    miRsAux.Open CadenaConsulta, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CadenaConsulta = ""
    Producida = 0
    If Not miRsAux.EOF Then
        CadenaConsulta = miRsAux!NumAlbar
        If Mid(CadenaConsulta, 1, 2) = "NP" Or Mid(CadenaConsulta, 1, 2) = "PR" Then
            vLote = miRsAux!NUmlote
            miRsAux.Close
            
            If Mid(CadenaConsulta, 1, 2) = "NP" Then
                'NUEVA PRODUCCION
                'select * from prodlin where codigo=492 and idlin=7
                
                CadenaConsulta = " AND prodlin.codigo = " & Val(Mid(CadenaConsulta, 3, 5)) & " AND prodlin.idlin = " & Val(Mid(CadenaConsulta, 8, 2))
                CadenaConsulta = "where prodlin.codigo= prodtrazlin.codigo AND prodlin.idlin= prodtrazlin.idlin " & CadenaConsulta
                CadenaConsulta = CadenaConsulta & " AND lotetraza=" & DBSet(vLote, "T")
                CadenaConsulta = "Select codartic,prodtrazlin.cantprodu cantidad FROM prodlin,prodtrazlin " & CadenaConsulta
                
            Else
                'Antigua
                'select codartic,sum(cantlote) cantidad from sliordprlotes where codigo=100090 group by 1
                CadenaConsulta = " WHERE codigo = " & Val(Mid(CadenaConsulta, 3)) & " AND codartic = " & DBSet(Text1(1).Text, "T")
                CadenaConsulta = "select codartic,sum(cantlote) cantidad from sliordprlotes  " & CadenaConsulta
                CadenaConsulta = CadenaConsulta & " GROUP BY 1"
               
            End If
            miRsAux.Open CadenaConsulta, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            CadenaConsulta = ""
            If miRsAux.EOF Then
                CadenaConsulta = "No se encuentra la produccion "
            Else
                If miRsAux!codartic <> Text1(1).Text Then
                    CadenaConsulta = "No coincide articulo "
                Else
                    'OK. Este es el bueno
                    Producida = DBLet(miRsAux!Cantidad, "N")
                    CadenaConsulta = ""
                End If
            End If
            
        Else
            CadenaConsulta = "No es produccion: " & miRsAux!NumAlbar
        End If
    Else
        CadenaConsulta = "No se encuentra la partida"
    End If
    miRsAux.Close
    
    'If CadenaConsulta <> "" Then MsgBox CadenaConsulta, vbExclamation
        
    
    
    
    CantidadVenta = 0
    For NumRegElim = 1 To ListView2.ListItems.Count
        CantidadVenta = CantidadVenta + ImporteFormateado(ListView2.ListItems(NumRegElim).SubItems(4))
    Next
    
    
    
    CadenaConsulta = "INSERT INTO tmppartidas(codusu,idpartida,codartic,numlote,idOperacion,Referencia,cantidad,abs_cantidad,idNumOperacion) VALUES ("
    CadenaConsulta = CadenaConsulta & vUsu.Codigo & "," & Text1(2).Text & ",'" & Text1(1).Text & "',"
    CadenaConsulta = CadenaConsulta & DBSet(Text2.Text, "T") & "," & DBSet(Text1(0).Text, "T")
    CadenaConsulta = CadenaConsulta & ",'" & Producida & "'," & DBSet(Text1(3).Text, "N", "N") & "," & DBSet(CantidadVenta, "N", "N") & ","
    CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "numalbar", "spartidas", "id", Text1(2).Text)
    If Mid(CadenaDesdeOtroForm, 1, 3) = "CUP" Then
        CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, 4)
        CadenaDesdeOtroForm = Val(CadenaDesdeOtroForm)
        CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "deposito", "olicoupage", "codigo", CadenaDesdeOtroForm)
        If CadenaDesdeOtroForm = "" Then
            CadenaDesdeOtroForm = "null"
        Else
            CadenaDesdeOtroForm = "'Deposito " & CadenaDesdeOtroForm & "'"
        End If
    Else
        CadenaDesdeOtroForm = "null"
    End If
        
    CadenaConsulta = CadenaConsulta & CadenaDesdeOtroForm & ")"
    conn.Execute CadenaConsulta
    
      
    If vParamAplic.QUE_EMPRESA = 4 Then
        If Not DatosImpresionGenerados Then
            HacerTrazaLaVall
            DatosImpresionGenerados = True
        End If
    End If
      
      
      

    'tmpinformes campo1,codigo1,nombre2,importe1
    CadenaConsulta = ""
    For NumRegElim = 1 To ListView2.ListItems.Count
        'tmpinformes codusu,codigo1,nombre1,campo1,nombre2,importe1 fecha
        With ListView2.ListItems(NumRegElim)
            CadenaConsulta = CadenaConsulta & ", (" & vUsu.Codigo & "," & .SubItems(2) & "," & DBSet(.Text, "T") & "," & NumRegElim & ","
            CadenaConsulta = CadenaConsulta & DBSet(.SubItems(3), "T") & "," & DBSet(.SubItems(4), "N") & "," & DBSet(.SubItems(1), "F") & ")"
        End With
    Next
    If CadenaConsulta <> "" Then
        CadenaConsulta = Mid(CadenaConsulta, 2)
        conn.Execute "INSERT INTO tmpinformes (codusu,codigo1,nombre1,campo1,nombre2,importe1,fecha1) VALUES " & CadenaConsulta
    End If
    Screen.MousePointer = vbDefault
    CadenaConsulta = "{tmppartidas.codusu}=" & vUsu.Codigo
    LlamaImprimirGral CadenaConsulta, "", 0, "TrazaArtVenta.rpt", "Trazabilidad lote venta "
    CadenaConsulta = Data1.RecordSource
    
End Sub

Private Sub TreeView1_DblClick()
Dim I As Integer
Dim Cad As String

    If TreeView1.Nodes.Count = 0 Then Exit Sub
    If TreeView1.SelectedItem Is Nothing Then Exit Sub
    
    If vParamAplic.QUE_EMPRESA <> 4 Then Exit Sub
    
    'El primero no hago nada, todavia
    If TreeView1.SelectedItem.Index = 1 Then Exit Sub
    
    If Mid(TreeView1.SelectedItem.Text, 1, 3) = "CUP" Then
        I = InStr(TreeView1.SelectedItem.Text, "(")
        If I = 0 Then Exit Sub
        Cad = Mid(TreeView1.SelectedItem.Text, 1, I - 1)
        Cad = Mid(Cad, 4)
         
         With frmAlmCoupage
            .DatosADevolverBusqueda2 = Trim(Cad)
            .Show vbModal
         End With
         
    Else
    
        If Mid(TreeView1.SelectedItem.Text, 1, 3) = "MOL" Then
            MsgBox TreeView1.SelectedItem.Tag
            I = InStr(1, TreeView1.SelectedItem.Text, "Id:")
            If I = 0 Then Exit Sub
            Cad = Mid(TreeView1.SelectedItem.Text, I + 3)
            I = InStr(1, Cad, "Dep")
            If I = 0 Then Exit Sub
            Cad = Mid(Cad, 1, I - 1)
            frmVallAlmazara.DatosADevolverBusqueda2 = Trim(Cad)
            frmVallAlmazara.Show vbModal
        Else
        
        End If
    End If
End Sub



Private Sub HacerTrazaLaVall()
Dim N As Node
Dim Fin As Boolean
Dim I As Integer
Dim Aux As String
Dim Cad As String
Dim Depo As Integer

    Kcampo = 1
    Fin = False
    Do
    
        I = InStr(1, TreeView1.Nodes(Kcampo).Text, " ")
        If I > 0 Then
            Aux = Trim(Mid(TreeView1.Nodes(Kcampo).Text, 1, I))
            Cad = "factorconversion<1 and codartic"
            Cad = DevuelveDesdeBD(conAri, "codartic", "sartic", Cad, Aux, "T")
            If Cad <> "" Then
                Depo = -1
                
                'Veo el deposito
                
                'OK. Este es el primero de la serie
                HacerNodImpresionRecursivo Kcampo, Depo, False
                
                Fin = True
            Else
                Kcampo = Kcampo + 1
            End If
        Else
            Kcampo = Kcampo + 1
            If Kcampo > 3 Then
                MsgBox "No se encuentra aceite", vbExclamation
                Fin = True
            End If
        End If
    Loop Until Fin
    
End Sub




Private Function HacerNodImpresionRecursivo(ByVal QueNodo As Integer, DepositoOrigen As Integer, PuedeSolapar As Boolean) As Boolean
Dim N1
Dim N2
Dim Aux2 As String
Dim J As Integer
Dim Depos1 As Integer
Dim Depos2 As Integer
Dim tipo1 As Byte  '0, nada  '1. Coup  2 Molt
Dim tipo2 As Byte  '0, nada  '1. Coup  2 Molt
Dim Solapar As Boolean
    
Dim DosHijos As Boolean
    
    
    'If InStr(1, TreeView1.Nodes(QueNodo).Text, "UP 239") > 0 Then Stop
    
    If TreeView1.Nodes(QueNodo).Children > 0 Then
        
        
            DosHijos = False
            If TreeView1.Nodes(QueNodo).Children >= 2 Then DosHijos = True
            
           
            Depos1 = -1
            Depos2 = 0
                
            'MI caso
            Set N1 = TreeView1.Nodes(QueNodo).Child
            tipo1 = EsMolturacionOCoupage(N1.Text)
            If tipo1 > 0 Then
                Depos1 = DepositoMolturacionCoupage(N1.Text, tipo1 = 1)
                
                If TreeView1.Nodes(QueNodo).Children > 1 Then
                    Set N2 = N1.Next
                Else
                    Set N2 = Nothing
                    
                End If
                
               
                If Not N2 Is Nothing Then
                    'Debug.Print TreeView1.Nodes(QueNodo).Text
                    'If N1.Text = N2.Text Then Stop
                    'If InStr(1, N2.Text, "UP 115") > 0 Then Stop
                End If
                
                If DosHijos Then
                    tipo2 = EsMolturacionOCoupage(N2.Text)
                    If tipo2 > 0 Then Depos2 = DepositoMolturacionCoupage(N2.Text, tipo2 = 1)
                Else
                    'Le pongo que tiene el mismo deposito y todo igual
                    tipo2 = 2   'Le digo que es una molturacion , de nmentiras, para que solapa ahi abajo
                    Depos2 = Depos1
                    If QueNodo = 1 And Not N2 Is Nothing Then
                        If TreeView1.Nodes(QueNodo).Children > 2 Then
                            If EsMolturacionOCoupage(N2.Next.Text) Then Depos2 = -1
                        End If
                    End If
                End If
            Else
                
            End If
            
            'If N2.Tag = 37 Then Stop
            
            '----------------------------
            
            'Solaparemos SI son dos molturaciones o coupages y molturacion
            Solapar = False
            If tipo1 = 2 And tipo2 = 2 Then
                Solapar = True
            Else
                If tipo1 = 1 And tipo2 = 2 Then
                    Solapar = True
                Else
                    If tipo1 = 2 And tipo2 = 1 Then Solapar = True
                End If
            End If
            
            'Si puede solapar, Y, repito, Y, el deposito son los mismos que el origen
            If Solapar Then
                If Depos1 = Depos2 Then
                    If Depos1 = DepositoOrigen Then
                        Solapar = True
                    Else
                       ' Stop
                         If tipo1 = 2 And tipo2 = 1 Then Stop
                        Solapar = True
                    End If
                Else
                    Solapar = False
                End If
            Else
                'Stop
                If Depos1 = Depos2 Then
                    'Stop
                    Solapar = True
                End If
            End If
            
            
            'Si fuerzo NO juntar, significa que este NO
            
            
            'Si son coupage. Lanzo a ver si podemos resumir lo que cuelga
            If tipo1 = 1 Then
                
                HacerNodImpresionRecursivo N1.Index, Depos1, Solapar
                
            End If
            If tipo2 = 1 Then
                
                If DosHijos Then
                    HacerNodImpresionRecursivo N2.Index, Depos2, Solapar
                Else
                    Stop
                End If
                
            End If
            
            If TreeView1.Nodes(QueNodo).Children > 2 Then
                
                HacerNodImpresionRecursivo N2.Next.Index, Depos2, Solapar
                'If Not Solapar Then HacerNodImpresionRecursivo N2.Next.Index, Depos2, Solapar
                
                If TreeView1.Nodes(QueNodo).Children > 3 Then Stop
                
            End If
                        
            
            If Solapar Then
                
                If Not PuedeSolapar Then Solapar = False
            End If
            
            
            
            'If InStr(1, TreeView1.Nodes(QueNodo).Text, "CUP 34") > 0 Then Stop
            
            'Solapamos
            If Solapar Then
                
                
                
                'Bajamos un nivel, para que no se desplaze
                
                BajarUnNivel N1.Index
                If DosHijos Then BajarUnNivel N2.Index
                If TreeView1.Nodes(QueNodo).Children > 2 Then Stop: N2.Next.Index
                
                Aux2 = TreeView1.Nodes(QueNodo).Tag
                Aux2 = "Delete from tmptraza where codusu =" & vUsu.Codigo & " AND contador =" & Aux2
                conn.Execute Aux2
                
                Espera 0.1
                
                
                
                
            End If
            
        Else
            MsgBox "+ de uno"
            HacerNodImpresionRecursivo = False
        End If
    
End Function

Private Sub BajarUnNivel(QueNodo As Integer)
Dim C As String
Dim N1
    'Bajamos un nivel su subnodo
    
    If TreeView1.Nodes(QueNodo).Children > 0 Then
        Set N1 = TreeView1.Nodes(QueNodo).Child
        Do
           ' Debug.Print "Tag: " & N1.Tag
            BajarUnNivel N1.Index
            Set N1 = N1.Next
        Loop Until N1 Is Nothing
        
        'Este tambien lo bajamos
        C = TreeView1.Nodes(QueNodo).Tag
        C = "Update tmptraza set nivle=nivle -1 WHERE codusu = " & vUsu.Codigo & " AND contador= " & C
        conn.Execute C
        
    Else
        C = "Update tmptraza set nivle=nivle -1 WHERE codusu = " & vUsu.Codigo & " AND contador= " & TreeView1.Nodes(QueNodo).Tag
        conn.Execute C
        
    End If
    
End Sub


Private Function EsMolturacionOCoupage(Texto As String) As Byte
Dim I As Integer
    
    EsMolturacionOCoupage = 0
    If Mid(Texto, 1, 3) = "CUP" Then
        EsMolturacionOCoupage = 1
    Else
        
        If Mid(Texto, 1, 4) = "MOLT" Then EsMolturacionOCoupage = 2
    End If
End Function

Private Function DepositoMolturacionCoupage(Texto As String, Coupage As Boolean) As Integer
Dim J As Integer
Dim Cad As String

    
    DepositoMolturacionCoupage = -2
    If Coupage Then
        'El coupage viene el deposito asin: CUP 26 (Dep:1 ) [MOSTRA51-2016]
        J = InStr(1, Texto, "(Dep:")
        If J > 0 Then
            Cad = Mid(Texto, J + 5)
            J = InStr(1, Cad, " ")
            If J > 0 Then
                Cad = Trim(Mid(Cad, 1, J - 1))
                DepositoMolturacionCoupage = CInt(Val(Cad))
            End If
        End If
    Else
        'Molturacion. El deposito;:  MOLT.Id:1  Dep:1  F:28/10 10:12 [MOSTRA1-2016]
        J = InStr(1, Texto, "Dep:")
        If J > 0 Then
            Cad = Mid(Texto, J + 4)
            J = InStr(1, Cad, " ")
            If J > 0 Then
                Cad = Trim(Mid(Cad, 1, J - 1))
                DepositoMolturacionCoupage = CInt(Val(Cad))
            End If
        End If
        
    End If
End Function
