VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmRegLisRev 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Historico de revisiones"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   ClipControls    =   0   'False
   Icon            =   "frmRegLisRev.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   10410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   3
      Left            =   2400
      MaxLength       =   16
      TabIndex        =   21
      Tag             =   "Puntuacion|N|S|||srevisiones|puntuacion|##0.00||"
      Text            =   "Text1"
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   1995
      Index           =   2
      Left            =   4440
      MaxLength       =   16
      TabIndex        =   20
      Tag             =   "Observa|T|S|||srevisiones|Comentarios|||"
      Text            =   "Text1"
      Top             =   840
      Width           =   5775
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   120
      MaxLength       =   16
      TabIndex        =   17
      Tag             =   "Realizado por|T|N|||srevisiones|realizadopor|||"
      Text            =   "Text1"
      Top             =   1680
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   120
      MaxLength       =   16
      TabIndex        =   15
      Tag             =   "Fecha|F|N|||srevisiones|fecha|dd/mm/yyyy|S|"
      Text            =   "Text1"
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   290
      Index           =   0
      Left            =   1200
      MaxLength       =   30
      TabIndex        =   14
      Text            =   "nom"
      Top             =   4800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   3
      Left            =   5520
      MaxLength       =   15
      TabIndex        =   5
      Tag             =   "Orden|N|N|||srevisionesl|ok|||"
      Text            =   "cantidad"
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   2
      Left            =   4440
      MaxLength       =   16
      TabIndex        =   4
      Tag             =   "Desc|T|N|||srevisionesl|denominacion||N|"
      Text            =   "hora"
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   5640
      Width           =   2505
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   240
         TabIndex        =   13
         Top             =   180
         Width           =   1515
      End
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   1
      Left            =   3360
      MaxLength       =   11
      TabIndex        =   3
      Tag             =   "Orden|N|N|||srevisionesl|orden|0||"
      Text            =   "fecha"
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   120
      MaxLength       =   3
      TabIndex        =   2
      Tag             =   "Cod. area|N|N|0|999|srevisionesl|codigo|000|S|"
      Text            =   "codalmac"
      Top             =   4800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Left            =   960
      TabIndex        =   11
      ToolTipText     =   "Buscar almacen"
      Top             =   4800
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   8040
      TabIndex        =   0
      Top             =   5790
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9195
      TabIndex        =   1
      Top             =   5790
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   9195
      TabIndex        =   10
      Top             =   5790
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10410
      _ExtentX        =   18362
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
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver Todos"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Generar nuevo listado"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Listado"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "�ltimo"
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   315
         Left            =   6600
         TabIndex        =   9
         Top             =   0
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   2880
      Top             =   5640
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmRegLisRev.frx":000C
      Height          =   2370
      Left            =   360
      TabIndex        =   6
      Top             =   3120
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   4180
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BorderStyle     =   0
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
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   2880
      Top             =   5880
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
      Caption         =   "Puntuacion"
      Height          =   255
      Index           =   3
      Left            =   2400
      TabIndex        =   22
      Top             =   600
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "Observaciones"
      Height          =   255
      Index           =   2
      Left            =   4440
      TabIndex        =   19
      Top             =   600
      Width           =   2040
   End
   Begin VB.Label Label1 
      Caption         =   "Realizado por"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   18
      Top             =   1440
      Width           =   1560
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   600
      Width           =   600
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
      Left            =   240
      TabIndex        =   8
      Top             =   8220
      Visible         =   0   'False
      Width           =   3495
   End
End
Attribute VB_Name = "frmRegLisRev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1
'Private WithEvents frmF As frmCal 'Calendario de Fechas


Dim NombreTabla As String
Dim Ordenacion As String
Private Modo As Byte

Dim kCampo As Integer
Dim PrimeraVez As Boolean
Dim btnPrimero2 As Byte 'Variable que indica el N� del Boton  PrimerRegistro en la Toolbar1

Dim CadenaConsulta As String
Dim CadenaBusqueda As String 'Cadena para la consulta de de busqueda en Grid
'Dim cadSeleccion As String 'Cadena de seleccion para FormulaSelection del Informe

Dim EsBusqueda As Boolean
'Para cargar el DataGrid con la consulta de busqueda y no con todos los registros

Private HaDevueltoDatos As Boolean


Private Sub cmdAceptar_Click()
Dim I As Integer
On Error GoTo Error1
    
    Screen.MousePointer = vbHourglass
    Select Case Modo
        Case 1 'Busqueda
            HacerBusqueda
        Case 4 'Modificar
            If DatosOk Then
'                If ModificaDesdeFormulario(Me, 3) Then
                If ModificarLinea Then
                      TerminaBloquear
'                      CadenaBusqueda = Data1.Recordset.Fields(0)
''                      LLamaLineas Modo, 0
                      PonerModo 2
'                      CancelaADODC Me.Data2
'
'                      Data1.Recordset.Find (Data1.Recordset.Fields(0).Name & " =" & I)
'                      CargaGrid True
                  End If
                  DataGrid1.SetFocus
            End If
    End Select
    
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub




Private Function DatosOk() As Boolean
Dim b As Boolean
    
    b = CompForm(Me, 3)
    If Not b Then Exit Function
       
    DatosOk = b
End Function


Private Sub Imprimir()
'Dim cad As String
'Dim numParam As Byte
'
'    'Resto parametros
'    cad = ""
'    cad = cad & "|pNomEmpre=""" & vParam.NombreEmpresa & """|"
'    numParam = 1
'
'    With frmImprimir
'        .NombreRPT = "rAlmMovim.rpt"
'        .OtrosParametros = cad
'        .NumeroParametros = numParam
'        .FormulaSeleccion = cadSeleccion
'        '.SoloImprimir = True
'        .Opcion = 9
'        .Titulo = ""
'        .Show vbModal
'    End With
End Sub





Private Sub cmdCancelar_Click()
On Error GoTo ECancelar

    Select Case Modo
        Case 1 'Busqueda
            LimpiarCampos
            PonerModo 0
            LLamaLineas 10
        Case 4 'Modificar
            PonerModo 2
            LLamaLineas 10
    End Select

ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmppal.Icon
   
    'ICONOS de La toolbar
    btnPrimero2 = 13 'Posicion del Boton Primero en la toolbar (+ 3 siguientes)
    With Toolbar1
        .ImageList = frmppal.imgListComun
        'ASignamos botones
        .Buttons(1).Image = 1   'Buscar
        .Buttons(2).Image = 2 'Ver Todos
        .Buttons(4).Image = 4 'Modificar
        .Buttons(5).Image = 5 'Eliminar
        
        .Buttons(6).Image = 21 'Nuevo
        
        .Buttons(8).Image = 16 'Imprimir
        .Buttons(9).Image = 40 'listado
        
        .Buttons(11).Image = 15 'Salir
        
        .Buttons(btnPrimero2).Image = 6 'Primero
        .Buttons(btnPrimero2 + 1).Image = 7 'Anterior
        .Buttons(btnPrimero2 + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero2 + 3).Image = 9 'Ultimo
    End With
    
    LimpiarCampos   'Limpia los campos TextBox
    PrimeraVez = True
    
    NombreTabla = "srevisiones"
    Ordenacion = " ORDER BY fecha "
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    Data1.ConnectionString = Conn
    CadenaConsulta = "Select * from " & NombreTabla & " WHERE fecha is null "
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    PonerCampos
    PonerModo 0
    
    CargaGrid (Modo = 2)
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid(enlaza As Boolean)
Dim b As Boolean
Dim tots As String
Dim SQL As String
On Error GoTo ECarga

    b = DataGrid1.Enabled
     
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral DataGrid1, Me.Data2, SQL, PrimeraVez
    
    'SELECT shinve.codartic, shinve.codalmac, salmpr.nomalmac, shinve.fechainv, shinve.horainve,existenc
    tots = "N||||0|;S|txtAux(0)|T|Area|800|;S|cmdAux|B||0|;S|txtAux2(0)|T|Ubicacion|3500|;S|txtAux(1)|T|orden|650|;"
    tots = tots & "S|txtAux(2)|T|ASpecto|3500|;S|txtAux(3)|T|Punt|600|;"
    
    arregla tots, DataGrid1, Me
    DataGrid1.Columns(5).Alignment = dbgRight
    
    DataGrid1.ScrollBars = dbgAutomatic

    DataGrid1.Enabled = b
    If Modo = 2 Then DataGrid1.Enabled = True
    PrimeraVez = False
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub


Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
'Almacen Propios
    txtAux(0).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtAux2(0).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub




Private Sub frmB_Selecionado(CadenaDevuelta As String)
'Formulario para Busqueda
Dim cadB As String


    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        
'            cadB = ""
'            cadB = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
'            CadenaConsulta = "select codartic from " & NombreTabla & " WHERE " & cadB & " GROUP BY codartic " & Ordenacion
'            PonerCadenaBusqueda
            
    End If
    Screen.MousePointer = vbDefault
End Sub







Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Text1_LostFocus(Index As Integer)

    If Text1(Index).BackColor = vbYellow Then Text1(Index).BackColor = vbWhite

 
End Sub


Private Sub txtAux_GotFocus(Index As Integer)
    If Modo = 1 Then
        ConseguirFoco txtAux(Index), Modo
    End If
End Sub

Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If Index = 3 And KeyCode = 40 Then
        PonerFocoBtn Me.cmdAceptar
   Else
        KEYdown KeyCode
   End If
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 3 And KeyAscii = 12 Then
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYpress KeyAscii
    End If
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim devuelve As String 'Para mensajes

    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    Exit Sub
    Select Case Index
        Case 0 'cod. almacen
            If txtAux(Index).Text = "" Then
             
            Else
                devuelve = DevuelveDesdeBD(conAri, "nomalmac", "salmpr", "codalmac", txtAux(Index).Text, "N")
'                Text2(1).Text = SQL
                If devuelve = "" Then 'No existe
                    devuelve = "No existe el Almacen" & vbCrLf
                    devuelve = devuelve & "C�digo: " & txtAux(Index).Text
                    MsgBox devuelve, vbExclamation
                    PonerFoco txtAux(Index)
                Else
                    txtAux(Index).Text = Format(txtAux(Index).Text, "000")
                End If
            End If
            
        Case 1 'Fecha Movimiento
             If txtAux(Index).Text <> "" Then PonerFormatoFecha txtAux(Index)
        Case 3
            If txtAux(Index).Text <> "" Then
                If Not PonerFormatoDecimal(txtAux(Index), 1) Then
                    PonerFoco txtAux(Index)
'                Else
'                    PonerFocoBtn Me.cmdAceptar
                End If
'            Else
'                  PonerFocoBtn Me.cmdAceptar
            End If
    End Select
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 'Busqueda
            BotonBuscar
        Case 2 'Ver Todos
            BotonVerTodos
        Case 4 'Modificar
            If BLOQUEADesdeFormulario(Me) Then BotonModificar
            
            
        Case 5
            'Eliminar
            Eliminar
            
        Case 6 'Imprimir
            CadenaDesdeOtroForm = ""
            frmVarios.Opcion = 6
            frmVarios.Show vbModal
            If CadenaDesdeOtroForm <> "" Then
                CadenaConsulta = "Select * from " & NombreTabla & " WHERE fecha = " & CadenaDesdeOtroForm & Ordenacion
                PonerCadenaBusqueda
                CadenaDesdeOtroForm = ""
            End If
        Case 8, 9
            If Data1.Recordset.EOF Then Exit Sub
            If Me.Data2.Recordset.EOF Then Exit Sub
            
            ImprimirL Button.Index = 8
            
        Case 11  'Salir
            Unload Me
        Case btnPrimero2 To btnPrimero2 + 3 'Flechas de Desplazamiento
            Desplazamiento (Button.Index - btnPrimero2)
    End Select
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim I As Byte
Dim b As Boolean
Dim NumReg As Byte

    Modo = Kmodo
    'Modo 2. Hay datos y estamos visualizandolos
    b = (Kmodo = 2)
    
    PonerIndicador Me.lblIndicador, Modo
    
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero2, b, NumReg

   'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar adem�s limpia los campos Text1
    BloquearText1 Me, Modo
    
    b = (Modo = 0) Or (Modo = 2)
    PonerBotonCabecera b
              
    b = Modo <> 0 And Modo <> 2

    
    PonerModoOpcionesMenu 'Activar opciones de menu seg�n Modo
    PonerOpcionesMenu   'Activar opciones de menu seg�n nivel
                        'de permisos del usuario
End Sub


Private Sub PonerModoOpcionesMenu()
Dim b As Boolean

    b = (Modo = 2) Or (Modo = 0)
    'Buscar
    Toolbar1.Buttons(1).Enabled = b
'    Me.mnBuscar.Enabled = b
    'Ver Todos
    Toolbar1.Buttons(2).Enabled = b
'    Me.mnVerTodos.Enabled = b
    
    b = (Modo = 2)

    'Modificar
    Toolbar1.Buttons(5).Enabled = b
'    Me.mnModificar.Enabled = b

    'Imprimir
    Toolbar1.Buttons(8).Enabled = True
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub

Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    DesplazamientoData Data1, Index
    PonerCampos
    CargaGrid True
End Sub


Private Function MontaSQLCarga(enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Bas�ndose en la informaci�n proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim SQL As String

    SQL = "SELECT * from srevisionesl "
    If enlaza Then
        If EsBusqueda And CadenaBusqueda <> "" Then
'            If Data1.Recordset.RecordCount > 1 Then
            'Si devuelve + de 1 registro en el DataGrid poner la info del primer articulo
                SQL = SQL & CadenaBusqueda & " AND fecha=" & DBSet(Text1(0).Text, "F")
'            Else
'                SQL = SQL & CadenaBusqueda
'            End If
        Else
            SQL = SQL & " WHERE fecha = " & DBSet(Text1(0).Text, "F")
        End If
    Else
        SQL = SQL & " WHERE fecha is null"
    End If
    SQL = SQL & " " & Ordenacion & " DESC "
    MontaSQLCarga = SQL
End Function


Private Sub BotonBuscar()
Dim anc As Single
    EsBusqueda = True
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        'Ponemos el grid lineasfacturas enlazando a ningun sitio
        CargaGrid False
            
        anc = ObtenerAlto(Me.DataGrid1)
        LLamaLineas anc
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


Private Sub BotonModificar()
Dim anc As Single
Dim I As Integer
    
    If Data2.Recordset.EOF Then Exit Sub
    If Data2.Recordset.RecordCount < 1 Then Exit Sub

    Screen.MousePointer = vbHourglass
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    
    PonerModo 4
    
    anc = ObtenerAlto(Me.DataGrid1)

    'Llamamos al form
    txtAux(0).Text = DataGrid1.Columns(1).Text
    txtAux(1).Text = DataGrid1.Columns(3).Text
    txtAux(2).Text = DataGrid1.Columns(4).Text
    txtAux(3).Text = DataGrid1.Columns(5).Text
    LLamaLineas anc
   
   'Como es modificar
    PonerFoco txtAux(1)
    Screen.MousePointer = vbDefault
End Sub


Private Sub BotonVerTodos()
'Ver todos
    EsBusqueda = False
    LimpiarCampos
    'Ponemos el grid lineasfacturas enlazando a ningun sitio
    CargaGrid False
    
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
        CargaGrid True
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub PonerBotonCabecera(b As Boolean)

    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    If b Then Me.lblIndicador.Caption = ""
    
End Sub


Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False)
'    cadSeleccion = ObtenerBusqueda(Me, True) 'Para la consulta de report

    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia cadB
    Else
        'Se muestran en el mismo form
        If cadB <> "" Then
            'Cadena para el Data1
            CadenaConsulta = "select codartic from " & NombreTabla & " WHERE " & cadB & " GROUP BY codartic " & Ordenacion
            'Cadena para el Datagrid y el Data2
            CadenaBusqueda = " WHERE " & cadB 'Para cargar la consulta del CargaGrid
        Else
            'obtener todos los articulos
            CadenaConsulta = "select codartic from " & NombreTabla & " GROUP BY codartic " & Ordenacion
            CadenaBusqueda = ""
        End If
        PonerCadenaBusqueda
    End If
End Sub


Private Sub PonerCadenaBusqueda()
Dim I As Byte
Screen.MousePointer = vbHourglass
On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta

    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ning�n registro en la tabla " & NombreTabla & " para ese criterio de b�squeda", vbInformation
        Screen.MousePointer = vbDefault
        PonerFoco Text1(0)
        'Limpiar los Campos Auxiliares
        For I = 0 To txtAux.Count - 1
            txtAux(I).Text = ""
        Next I
        Exit Sub
    Else
        PonerModo 2
        LLamaLineas 10
        PonerCampos
        CargaGrid True
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
    
    PonerCamposForma Me, Data1
   
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub


Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
Dim tabla As String
Dim Titulo As String

    'Llamamos a al form
    Cad = ""
            
    Cad = Cad & "Articulo|shinve|codartic|T||25�Denominacion|sartic|nomartic|T||70�"
    tabla = "(" & NombreTabla & " LEFT JOIN sartic ON " & NombreTabla & ".codartic=sartic.codartic" & ") "
'        tabla = tabla & " GROUP BY shinve.codartic "
    'tabla = "sartic"
    Titulo = "Hist�rico Inventario"
           
    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = Titulo
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri 'Conexion a BD Ariges
'        frmB.vBuscaPrevia = chkVistaPrevia
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
        If HaDevueltoDatos Then
''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
''                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
            PonerFoco Text1(kCampo)
'            Toolbar1.Buttons(5).Enabled = True 'Imprimir
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub LLamaLineas(alto As Single)
Dim jj As Byte
Dim ini As Byte
Dim b As Boolean

    DeseleccionaGrid Me.DataGrid1
    
    b = (Modo = 3 Or Modo = 4 Or Modo = 1) 'Insertar o Modificar Lineas
    b = False
    If Modo = 4 Then 'modificar
        ini = 1
    Else
        ini = 0
    End If
    
    For jj = ini To txtAux.Count - 1
        txtAux(jj).Height = DataGrid1.RowHeight
        txtAux(jj).Top = alto
        txtAux(jj).visible = b
    Next jj

    b = (Modo = 1)
    Me.cmdAux.Height = DataGrid1.RowHeight
    Me.cmdAux.Top = alto
    Me.cmdAux.visible = b
End Sub


Private Function ModificarLinea() As Boolean
Dim SQL As String
On Error GoTo EModificar

    ModificarLinea = False
    SQL = "UPDATE " & NombreTabla & " SET RealizadoPor=" & DBSet(Text1(1).Text, "T")
    SQL = SQL & ", Comentarios= " & DBSet(Text1(2).Text, "T")
    SQL = SQL & ", puntuacion= " & DBSet(Text1(3).Text, "N")
    SQL = SQL & " WHERE fecha=" & DBSet(Text1(0).Text, "F")
    Conn.Execute SQL
    ModificarLinea = True
EModificar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Modificar Linea", Err.Description
End Function

Private Sub ImprimirL(ImprimeRevision As Boolean)
        If Text1(0).Text = "" Then Exit Sub 'No deberia poasar nunca
        With frmImprimir
            If ImprimeRevision Then
                .FormulaSeleccion = " {srevisiones.fecha} = Date(" & Year(Text1(0).Text) & "," & Month(Text1(0).Text) & "," & Day(Text1(0).Text) & ")"
                .Titulo = "Listado revisi�n: " & Text1(0).Text
                .NombreRPT = "morListaRevision.rpt"
            Else
                'Listado con toad las revisiones
                .FormulaSeleccion = ""
                .Titulo = "Listado revisiones "
                .NombreRPT = "morListaRevisionUltFec.rpt"
            End If
            .OtrosParametros = ""
            .NumeroParametros = 0
    
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = 2002
            .ConSubInforme = False
            .Show vbModal
        End With
End Sub


Private Sub Eliminar()
Dim Cad As String

    If Modo <> 2 Then Exit Sub
       'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    Cad = "�Seguro que desea eliminar el la revision con fecha " & Data1.Recordset.Fields(0) & "?"

    
    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        NumRegElim = Data1.Recordset.AbsolutePosition
        Screen.MousePointer = vbHourglass
        Cad = "DELETE from srevisionesl WHERE fecha = " & DBSet(Text1(0).Text, "F")
        Conn.Execute Cad
        Cad = "DELETE from srevisiones WHERE fecha = " & DBSet(Text1(0).Text, "F")
        Conn.Execute Cad
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
