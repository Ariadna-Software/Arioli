VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmProdNueTraza2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Trazabilidad"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9180
   Icon            =   "frmProdNueTraza.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   9180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPal 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   39
      Text            =   "Text1"
      Top             =   3720
      Width           =   1245
   End
   Begin VB.TextBox txtPal 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   38
      Text            =   "Text1"
      Top             =   3720
      Width           =   1245
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   13
      Left            =   1560
      TabIndex        =   10
      Tag             =   "Inicio|T|S|||prodlin|LineaExtraEtiqueta|||"
      Text            =   "Text1"
      Top             =   2760
      Width           =   3765
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   12
      Left            =   240
      TabIndex        =   9
      Tag             =   "Inicio|F|N|||prodlin|feccaduca|dd/mm/yyyy||"
      Text            =   "Text1"
      Top             =   2760
      Width           =   1125
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   11
      Left            =   4560
      TabIndex        =   13
      Tag             =   "UdsTraza|N|S|0||prodtrazlin|can2|0||"
      Text            =   "Text1"
      Top             =   3720
      Width           =   1245
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   10
      Left            =   3600
      TabIndex        =   12
      Tag             =   "Cajas traza|N|S|0||prodtrazlin|caj2|0||"
      Text            =   "Text1"
      Top             =   3720
      Width           =   885
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   8
      Left            =   8520
      TabIndex        =   8
      Tag             =   "Linea prod|N|S|0|9|prodtrazlin|lineaprod|0||"
      Text            =   "Text1"
      Top             =   1920
      Width           =   525
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   1560
      MaxLength       =   4
      TabIndex        =   1
      Tag             =   "Codigo prod|N|N|0|9999|prodlin|idlin|000||"
      Text            =   "1"
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   9
      Left            =   2280
      TabIndex        =   11
      Tag             =   "Trazabilidad|N|N|0||prodtrazlin|lotetraza|0000000||"
      Text            =   "Text1"
      Top             =   3720
      Width           =   1245
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Tag             =   "Codigo prod|N|N|0||prodlin|codigo|00000||"
      Text            =   "Text1"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   2400
      MaxLength       =   30
      TabIndex        =   2
      Tag             =   "Articulo|T|N|||prodlin|codartic|||"
      Text            =   "Text1"
      Top             =   1200
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   3
      Left            =   4080
      TabIndex        =   3
      Tag             =   "Nomartic|T|N|||sartic|nomartic|||"
      Text            =   "Text1"
      Top             =   1200
      Width           =   4965
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   6
      Left            =   6000
      TabIndex        =   6
      Tag             =   "Cajas|N|S|0||prodlin|cajasprod|#,##0||"
      Text            =   "Text1"
      Top             =   1920
      Width           =   765
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   7
      Left            =   7080
      TabIndex        =   7
      Tag             =   "Uds|N|S|0||prodlin|cantprodu|#,##0||"
      Text            =   "Text1"
      Top             =   1920
      Width           =   1245
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Tag             =   "Inicio|F|N|||prodlin|fhinicio|dd/mm/yyyy hh:mm:ss||"
      Text            =   "Text1"
      Top             =   1920
      Width           =   2685
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   5
      Left            =   3120
      TabIndex        =   5
      Tag             =   "Inicio|F|S|||prodlin|fhfin|dd/mm/yyyy hh:mm:ss||"
      Text            =   "Text1"
      Top             =   1920
      Width           =   2685
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   7890
      TabIndex        =   15
      Top             =   7800
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   315
      TabIndex        =   17
      Top             =   7635
      Width           =   3000
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   210
         Width           =   2595
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7890
      TabIndex        =   16
      Top             =   7800
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6720
      TabIndex        =   14
      Top             =   7800
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   270
      Top             =   7875
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
      TabIndex        =   21
      Top             =   0
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
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
            Object.ToolTipText     =   "Imprimir resumen diario"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   6360
         TabIndex        =   22
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   3135
      Left            =   240
      TabIndex        =   36
      Top             =   4440
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   5530
      _Version        =   393217
      Indentation     =   471
      LabelEdit       =   1
      LineStyle       =   1
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Produccion"
      Height          =   195
      Index           =   14
      Left            =   6120
      TabIndex        =   40
      Top             =   3480
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cajas palets"
      Height          =   195
      Index           =   13
      Left            =   7560
      TabIndex        =   37
      Top             =   3480
      Width           =   855
   End
   Begin VB.Line Line3 
      X1              =   6000
      X2              =   6000
      Y1              =   3480
      Y2              =   4080
   End
   Begin VB.Label Label1 
      Caption         =   "F. Caducidad"
      Height          =   195
      Index           =   9
      Left            =   240
      TabIndex        =   35
      Top             =   2520
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "Datos extra"
      Height          =   195
      Index           =   8
      Left            =   1560
      TabIndex        =   34
      Top             =   2520
      Width           =   1410
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00404000&
      Caption         =   "Trazabilidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Index           =   7
      Left            =   240
      TabIndex        =   33
      Top             =   3600
      Width           =   1965
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Producción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   360
      Index           =   6
      Left            =   240
      TabIndex        =   32
      Top             =   480
      Width           =   1620
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   9000
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label Label1 
      Caption         =   "Uds"
      Height          =   255
      Index           =   5
      Left            =   4560
      TabIndex        =   31
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Cajas"
      Height          =   255
      Index           =   4
      Left            =   3600
      TabIndex        =   30
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Linea"
      Height          =   195
      Index           =   3
      Left            =   8520
      TabIndex        =   29
      Top             =   1680
      Width           =   390
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   9000
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Label1 
      Caption         =   "Lin"
      Height          =   255
      Index           =   2
      Left            =   1560
      TabIndex        =   28
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Trazabilidad"
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
      Index           =   18
      Left            =   2280
      TabIndex        =   27
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Left            =   3120
      Tag             =   "-1"
      ToolTipText     =   "Buscar población"
      Top             =   960
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Fin"
      Height          =   195
      Index           =   15
      Left            =   3120
      TabIndex        =   26
      Top             =   1680
      Width           =   210
   End
   Begin VB.Label Label1 
      Caption         =   "Inicio"
      Height          =   255
      Index           =   12
      Left            =   240
      TabIndex        =   25
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Cajas"
      Height          =   255
      Index           =   11
      Left            =   6000
      TabIndex        =   24
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Uds"
      Height          =   255
      Index           =   10
      Left            =   7200
      TabIndex        =   23
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Artículo"
      Height          =   195
      Index           =   1
      Left            =   2400
      TabIndex        =   20
      Top             =   960
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Código"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   19
      Top             =   990
      Width           =   615
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
Attribute VB_Name = "frmProdNueTraza2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public QueTrazabilidad As Long


Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmA As frmAlmArticulos 'Form Mantenimiento Articulos
Attribute frmA.VB_VarHelpID = -1



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

Private CadenaConsulta2 As String
Private kCampo As Integer
'-------------------------------------------------------------------------
Private Ordenasao As String
Private PrimeraVez As Boolean


Private Sub cmdAceptar_Click()
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                
                    If Data1.Recordset.EOF Then 'No estaba cargado Inicialmente
                        Data1.RecordSource = MontaSQL2(False) & Ordenasao
                        Data1.Refresh
                    End If
                    PosicionarData
                End If
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


Private Function MontaSQL2(ParaBusquedaPrevia As Boolean) As String
    If Not ParaBusquedaPrevia Then
        MontaSQL2 = "select prodlin.codigo,prodlin.idlin,prodlin.codartic,nomartic,cantesti,prodlin.cantprodu,fhinicio,fhFin,"
        MontaSQL2 = MontaSQL2 & " prodlin.cajasprod,feccaduca,LineaExtraEtiqueta , prodtrazlin.cantprodu can2, prodtrazlin.Cajasprod caj2,lineaprod,lotetraza"
        MontaSQL2 = MontaSQL2 & " from prodlin,prodtrazlin,sartic"
        MontaSQL2 = MontaSQL2 & " Where prodlin.Codigo = prodtrazlin.Codigo And prodlin.idlin = prodtrazlin.idlin and prodlin.codartic=sartic.codartic"
    Else
        MontaSQL2 = "prodlin.Codigo = prodtrazlin.Codigo And prodlin.idlin = prodtrazlin.idlin and prodlin.codartic=sartic.codartic"
    
    End If
End Function


Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 1, 3 'Busqueda, Insertar
            LimpiarCampos
            PonerModo 0
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
    PonerModo 3
    
    Text1(0).Text = Format(SugerirCodigoSiguienteStr("sbanpr", "codbanpr"), "0000")
    PonerFoco Text1(0)
End Sub

Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        '### A mano
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(9)
        Text1(9).BackColor = vbYellow
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
    LimpiarCampos
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta2 = MontaSQL2(False) & Ordenasao
        PonerCadenaBusqueda2
    End If
End Sub

Private Sub Desplazamiento(Index As Integer)
'Para desplazarse por los registros de control Data
    DesplazamientoData Data1, Index
    PonerCampos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub

Private Sub BotonModificar()
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    PonerFoco Text1(1)
End Sub


Private Sub BotonEliminar()
Dim cad As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    '### a mano
    cad = "¿Seguro que desea eliminar el Banco Propio? " & vbCrLf
    cad = cad & vbCrLf & "Cod. Banco : " & Format(Data1.Recordset.Fields(0), "0000")
    cad = cad & vbCrLf & "Desc. Banco: " & Data1.Recordset.Fields(1)
    
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then     'Borramos
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        cad = "Delete from sbanpr where codbanpr=" & Data1.Recordset!codbanpr
        Conn.Execute cad
'        Data1.Recordset.Delete
        If SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            PonerModo 0
        End If
        Screen.MousePointer = vbDefault
    End If
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Banco Propio", Err.Description
End Sub


Private Sub cmdRegresar_Click()
Dim cad As String

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    
    cad = Data1.Recordset.Fields(0) & "|"
    cad = cad & Data1.Recordset.Fields(1) & "|"
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        If QueTrazabilidad > 0 Then
            chkVistaPrevia = 0
            Modo = 1
            Text1(9).Text = Me.QueTrazabilidad
            BotonBuscar
        End If
        
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmppal.Icon

    'Icono de busqueda
    Me.imgBuscar.Picture = frmppal.imgListComun.ListImages(19).Picture

    PrimeraVez = True

    ' ICONITOS DE LA BARRA
    btnPrimero = 15
    With Me.Toolbar1
        .ImageList = frmppal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        .Buttons(5).Image = 3   'Insertar Nuevo
        .Buttons(6).Image = 4   'Modificar
        .Buttons(7).Image = 5   'Borrar
        
        .Buttons(9).Image = 16  'Botón Imprimir
        .Buttons(10).Image = 40 'Resumen diario produccion
        
        .Buttons(12).Image = 15  'Salir
        
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With
    
    LimpiarCampos   'Limpia los campos TextBox

    Ordenasao = " ORDER BY prodlin.codigo,prodlin.idlin ,lotetraza"

        
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = Conn
    
    
    If QueTrazabilidad = 0 Then
        Data1.RecordSource = MontaSQL2(False) & " AND  prodlin.idlin=-1"
        Data1.Refresh
    
        PonerModo 0
    Else
    
       
    End If
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    TreeView1.Nodes.Clear
    txtPal(0).Text = ""
    txtPal(1).Text = ""
    
    'Aqui va el especifico de cada form es
    '### a mano
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub

Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
    Text1(2).Text = RecuperaValor(CadenaSeleccion, 1)
    Text1(3).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    CadenaConsulta2 = CadenaDevuelta
End Sub

Private Sub imgBuscar_Click()
    If Modo <> 1 Then Exit Sub
    Screen.MousePointer = vbHourglass
    Set frmA = New frmAlmArticulos 'Form Mantenimiento Articulos
    frmA.DatosADevolverBusqueda2 = "@1@" 'Poner Modo Busqueda
    frmA.Show vbModal
    Set frmA = Nothing


    PonerFocoBtn Me.cmdAceptar
    
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
    Screen.MousePointer = vbHourglass
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
Dim devuelve As String
      
    'en el campo ID de norma 34 no se hace Trim ni nada. Lo q pongan
    If Index = 18 Then Exit Sub
      
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
''''    'Si queremos hacer algo ..
''''    Select Case Index
''''         Case 0
''''            If PonerFormatoEntero(Text1(Index)) Then
''''                If Modo = 3 Then 'Insertar
''''                    'Detectamos aki si ya existe y no esperamos hasta boton Aceptar
''''                    If ExisteCP(Text1(Index)) Then PonerFoco Text1(Index)
''''                End If
''''            End If
''''
''''         Case 3 'CPostal
''''            If Text1(Index).Text = "" Then
''''                Text1(Index + 1).Text = ""
''''            ElseIf Not VieneDeBuscar Then
''''                 Text1(Index + 1).Text = ObtenerPoblacion(Text1(Index).Text, devuelve)
''''            End If
''''            VieneDeBuscar = False
''''
''''         Case 10, 11 'codbanco, codsucursal
''''            PonerFormatoEntero Text1(Index)
''''
''''         Case 12, 13 'DC, numero cta
''''            FormateaCampo Text1(Index)
''''
''''
''''    End Select
End Sub


Private Sub HacerBusqueda()
Dim cadB As String


    'Cambio los CHK para que pueda bsucar por fechas
   ' Me.Text1(1).Tag = "Inicio|F|N|||prodpalets|fhinicio|dd/mm/yyyy hh:mm:ss||"
   ' Me.Text1(2).Tag = "Fin|F|S|||prodpalets|fhFin|dd/mm/yyyy hh:mm:ss||"
    


    cadB = ObtenerBusqueda(Me, False)
    
    If cadB <> "" Then
        
        
        
        'Las fechas
        cadB = Replace(cadB, "prodlin.fhinicio", "date(prodlin.fhinicio)")
        cadB = Replace(cadB, "prodlin.fhfin", "date(prodlin.fhfin)")
        
        
        'Por si acaso esta buscando por cantidad o cajas de prod
        ' prodtrazlin.cantprodu can2, prodtrazlin.Cajasprod caj2,
        cadB = Replace(cadB, "can2", "cantprodu")
        cadB = Replace(cadB, "caj2", "Cajasprod")
        

        
    End If
    
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia cadB
    Else
        'Se muestran en el mismo form
        If cadB <> "" Then
            CadenaConsulta2 = MontaSQL2(False) & " AND " & cadB & Ordenasao
            PonerCadenaBusqueda2
        End If
    End If
End Sub


Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String

        'Llamamos a al form
        '##A mano


            'Cod Diag.|tabla|columna|tipo|formato|10·
            cad = "Trazab.|prodtrazlin|lotetraza|N|000000|12·"
            cad = cad & "Referencia|prodlin|codartic|T||22·"
            cad = cad & "Articulo|sartic|nomartic|T||50·"
           
   
        
        
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vCampos = cad
            frmB.vTabla = "prodlin,prodtrazlin,sartic"
            
            cad = MontaSQL2(True)
            If cadB <> "" Then cadB = " AND " & cadB
            frmB.vSQL = cad & cadB
     

            frmB.vDevuelve = "0|"
            frmB.vTitulo = "Lotes trazabilidad"
            frmB.vselElem = 0
            frmB.vConexionGrid = conAri

            frmB.vCargaFrame = False
            CadenaConsulta2 = ""
            frmB.Show vbModal
            Set frmB = Nothing
          
            If CadenaConsulta2 <> "" Then
                CadenaConsulta2 = RecuperaValor(CadenaConsulta2, 1)
                CadenaConsulta2 = " AND lotetraza =" & CadenaConsulta2
                CadenaConsulta2 = MontaSQL2(False) & CadenaConsulta2
                PonerCadenaBusqueda2
            
            End If
          

End Sub


Private Sub PonerCadenaBusqueda2()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta2
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        If Modo = 1 Then 'Busqueda
             MsgBox "No hay ningún registro  para ese criterio de Búsqueda.", vbInformation
             PonerFoco Text1(0)
        Else
            MsgBox "No hay ningún registro en la tabla ", vbInformation
        End If
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
Dim I As Byte
    
    If Data1.Recordset.EOF Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    PonerCamposForma Me, Data1

    lblIndicador.Caption = "Hco"
    lblIndicador.Refresh
    CargarHco
    
    lblIndicador.Caption = "palets"
    lblIndicador.Refresh
    txtPal(0).Text = ""
    txtPal(1).Text = ""
    
                'siempre deberia ser <>""
    If Text1(9).Text <> "" Then
        txtPal(0).Text = DevuelveDesdeBD(conAri, "count(*)", "prodcajas", "lotetraza", Text1(9).Text)
        txtPal(1).Text = DevuelveDesdeBD(conAri, "count(*)", "prodcajasprod", "lotetraza", Text1(9).Text)
    End If
    
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    Screen.MousePointer = vbDefault
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
Private Sub PonerModo(Kmodo As Byte)
Dim b As Boolean
Dim NumReg As Byte
   
    Modo = Kmodo
        
    '----------------------------------------------
    'Modo 2. Hay datos y estamos visualizandolos
    b = (Kmodo = 2)
    PonerIndicador lblIndicador, Modo
    
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
    
    
    '------------------------------------------------------
    'Modo insertar o modificar
    b = (Kmodo >= 3) '-->Luego not b sera kmodo<3
    cmdAceptar.visible = b Or Modo = 1
    cmdCancelar.visible = b Or Modo = 1
    If cmdCancelar.visible Then
        cmdCancelar.Cancel = True
    Else
        cmdCancelar.Cancel = False
    End If
       
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    
    chkVistaPrevia.Enabled = (Modo <= 2)
    
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    
    PonerModoOpcionesMenu 'Activar opciones de menu según el Modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Sub PonerModoOpcionesMenu()
'Activa botones de la Toolbar segun el Modo
Dim b As Boolean

    b = False
    'Lo comento
    'B = (Modo = 2)
    'Modificar
    Toolbar1.Buttons(6).Enabled = b
    mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(7).Enabled = b
    mnEliminar.Enabled = b
    
    '-----------------------------------------
    'B = (Modo >= 3) 'Insertar/Modificar
    b = True
    'Insertar
    Toolbar1.Buttons(5).Enabled = Not b
    Me.mnNuevo.Enabled = Not b
    
    b = (Modo >= 3) 'Insertar/Modificar
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
    b = CompForm(Me, 1)
    If Not b Then Exit Function
    
    If Not Comprueba_CC(Text1(10).Text & Text1(11).Text & Text1(12).Text & Text1(13).Text) Then
        If MsgBox("La cuenta bancaria no es correcta. ¿Continuar?", vbQuestion + vbYesNo) = vbNo Then b = False
    End If
 
    If Modo = 3 Then 'Insertar
        If ExisteCP(Text1(0)) Then b = False
    End If

    DatosOk = b
End Function


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1: mnBuscar_Click  'Buscar
        Case 2: mnVerTodos_Click  'Todos
            
        Case 5: mnNuevo_Click  'Nuevo
        Case 6: mnModificar_Click  'Modificar
        Case 7: mnEliminar_Click  'Borrar
            
        Case 9: Imprimir
        Case 10:
                'Imprime resume diario produccion
                frmListado2.opcion = 27
                frmListado2.Show vbModal
        Case 12
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


Private Sub PosicionarData()
Dim cad As String
Dim Indicador As String

    cad = "(codbanpr=" & Text1(0).Text & ")"
    If SituarData(Data1, cad, Indicador) Then
        PonerModo 2
        lblIndicador.Caption = Indicador
    Else
'        LimpiarCampos
        PonerModo 0
    End If
End Sub







Private Sub CargarHco()
Dim N As Node
Dim idTraza As Long
Dim Cantidad As Currency
Dim L As Byte
Dim C2 As Currency
Dim SQL As String


    

    TreeView1.Nodes.Clear
    idTraza = -1
    SQL = "select prodtrazcompo.*,nomartic,cantprodu,cajasprod from prodtrazlin,prodtrazcompo,sartic where"
    SQL = SQL & " prodtrazcompo.codigo = prodtrazlin.codigo and prodtrazcompo.idlin  = prodtrazlin.idlin  and"
    SQL = SQL & " prodtrazcompo.lineaprod   = prodtrazlin.lineaprod  and    prodtrazcompo.lotetraza = prodtrazlin.lotetraza and"
    SQL = SQL & " prodtrazcompo.codartic = sartic.codartic and prodtrazlin.codigo=" & Text1(0).Text & " and prodtrazlin.idlin= " & Text1(1).Text
    
    'Las cargamos todas y asi tenemos los datos de los productos utilizados
    'SQL = SQL & " and prodtrazlin.lotetraza <>" & Text1(9).Text
    
    
    SQL = SQL & "  order by lotetraza,factorconversion,codartic"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    idTraza = -1
    Cantidad = 0
    While Not miRsAux.EOF
        If idTraza <> miRsAux!lotetraza Then
            idTraza = miRsAux!lotetraza
            Cantidad = Cantidad + DBLet(miRsAux!cantprodu, "N")
            
                
                Set N = TreeView1.Nodes.Add(, , "C" & idTraza)
                SQL = "      Uds: " & Right(Space(8) & Format(miRsAux!cantprodu, "#,##0"), 8)
                N.Text = "LOTE " & Format(idTraza, "00000") & SQL
                SQL = "      Cajas: " & Right(Space(6) & Format(miRsAux!Cajasprod, "#,##0"), 6)
                N.Text = N.Text & SQL
                If idTraza = Val(Text1(9).Text) Then
                    'Es este
                    N.Bold = True
                    N.BackColor = vbBlack
                    N.ForeColor = vbWhite
                    N.Expanded = True
                End If
        End If
      
            Set N = TreeView1.Nodes.Add("C" & idTraza, tvwChild)
            
            SQL = miRsAux!codArtic & " " & miRsAux!NomArtic
            L = Len(SQL)
            If L > 45 Then
                SQL = Mid(SQL, 1, 45)
                L = 1
            Else
                L = 46 - L
            End If
            
            SQL = SQL & Space(CLng(L))
            SQL = SQL & "Lot:" & miRsAux!NumLote & " / "
            C2 = DBLet(miRsAux!cantutili, "N")
            If Int(C2) = C2 Then
                SQL = SQL & Format(C2, "#,##0")
            Else
                
                SQL = SQL & Format(C2, FormatoCantidad)
            End If
            N.Text = SQL
   
        miRsAux.MoveNext
    Wend
    miRsAux.Close
'    If Cantidad > 0 Then
'        Text1(5).Text = Format(Cantidad, FormatoCantidad)
'        'idTraza = Me.cLP.UnidadesCaja
'        If idTraza = 0 Then idTraza = 1
'        idTraza = Cantidad \ idTraza
'        Text1(10).Text = idTraza
'
'    End If
    Set miRsAux = Nothing
End Sub


Private Sub Imprimir()
                
    If Modo = 1 Then Exit Sub
    
    If Data1.Recordset Is Nothing Then Exit Sub
    If Data1.Recordset.EOF Then Exit Sub
    
    
    
 
    CadenaDesdeOtroForm = "{prodcab.codigo}=" & Text1(0).Text & " AND {prodlin.idlin} = " & Text1(1).Text
    LlamaImprimirGral CadenaDesdeOtroForm, "", 0, "produccionNueva.rpt", "Lote trazabilidad "
    CadenaDesdeOtroForm = ""

End Sub

Private Sub TreeView1_DblClick()

    If Me.TreeView1.Nodes.Count = 0 Then Exit Sub
    If Me.TreeView1.SelectedItem Is Nothing Then Exit Sub
    
    'Es un nodo hijo
    If Me.TreeView1.SelectedItem.Child Is Nothing Then Exit Sub
    
    If MsgBox("Ver lote traza " & Me.TreeView1.SelectedItem.Text & "?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    CadenaConsulta2 = Mid(Me.TreeView1.SelectedItem.Text, 1, InStr(1, Me.TreeView1.SelectedItem.Text, "ds:") - 4)
    CadenaConsulta2 = Trim(Mid(CadenaConsulta2, 5))
    CadenaConsulta2 = "lotetraza = " & CadenaConsulta2
    CadenaConsulta2 = MontaSQL2(False) & " AND " & CadenaConsulta2 & Ordenasao
    PonerCadenaBusqueda2


    
    
    
End Sub
