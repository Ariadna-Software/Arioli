VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmVallCierrAlmazara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cierre y asignacion rendimientos de ALMAZARA"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13380
   ClipControls    =   0   'False
   Icon            =   "frmVallCierrAlmazara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   13380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   2
      Left            =   12360
      MaxLength       =   10
      TabIndex        =   6
      Text            =   "0,00"
      Top             =   720
      Width           =   825
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   9600
      MaxLength       =   10
      TabIndex        =   5
      Tag             =   "Hora fin|H|S|||vallalmazaraproceso|HoraFin|hh:nn:ss|N|"
      Top             =   720
      Width           =   1305
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   8280
      MaxLength       =   10
      TabIndex        =   4
      Tag             =   "Hora fin|H|S|||vallalmazaraproceso|HoraFin|hh:nn:ss|N|"
      Top             =   720
      Width           =   1305
   End
   Begin VB.TextBox txtLote 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "Text2"
      Top             =   720
      Width           =   1965
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   1
      Left            =   11760
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "Text2"
      Top             =   5280
      Width           =   1245
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   0
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   5280
      Width           =   1245
   End
   Begin VB.ComboBox cboDepo 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   720
      Width           =   2055
   End
   Begin VB.TextBox txtAux 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   320
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   16
      TabIndex        =   0
      Text            =   "existencia"
      Top             =   5040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   5835
      Width           =   3255
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   240
         TabIndex        =   12
         Top             =   180
         Width           =   2355
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   10800
      TabIndex        =   1
      Top             =   6000
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   12075
      TabIndex        =   2
      Top             =   6000
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   10800
      TabIndex        =   8
      Top             =   6000
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   13380
      _ExtentX        =   23601
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmVallCierrAlmazara.frx":000C
      Height          =   3795
      Left            =   120
      TabIndex        =   9
      Top             =   1305
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   6694
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   120
      Top             =   5520
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
      Caption         =   "%Inc bodega"
      Height          =   195
      Index           =   3
      Left            =   11160
      TabIndex        =   21
      Top             =   720
      Width           =   1755
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha - Hora finalización"
      Height          =   195
      Index           =   2
      Left            =   6360
      TabIndex        =   20
      Top             =   720
      Width           =   1755
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Litros"
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
      Index           =   1
      Left            =   10680
      TabIndex        =   19
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Lote"
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   17
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Total"
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
      Index           =   0
      Left            =   8400
      TabIndex        =   16
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Deposito"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblInfInv 
      Height          =   255
      Left            =   3600
      TabIndex        =   13
      Top             =   6000
      Width           =   2055
   End
End
Attribute VB_Name = "frmVallCierrAlmazara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ID As Long

Private Modo As Byte

Dim Kcampo As Integer

Dim CadenaConsulta As String

Dim gridCargado As Boolean 'Si el DataGrid ya tiene todos los Datos cargados.
                           'Para el RowColChange, si el grid no esta totalmente cargado el CargaTxtAux da error.

Dim PrimeraVez As Boolean

Private HaDevueltoDatos As Boolean
Dim HorasProceso As Currency

Dim TrabajadorParte As Integer

Private Sub cmdAceptar_Click()
Dim Cad As String
Dim Cade1 As String

    On Error GoTo Error1
    
    Screen.MousePointer = vbHourglass
    
    
    lblInfInv.Caption = ""
    Select Case Modo
        Case 2
            'Cerrar proceso. Asignar rendimientos. Cerrar parte
            Cad = DevuelveDesdeBD(conAri, "numalbar", "tmpnlotes", "cantidad=0 and codusu", vUsu.Codigo)
            If Cad <> "" Then
                MsgBox "Lineas sin asignar rendimiento", vbExclamation
                
            Else
                
                If cboDepo.ListIndex < 0 Then
                    MsgBox "Seleccione el deposito destino", vbExclamation
                    Cad = "No"
                Else
                    If Trim(Text1(0).Text) = "" Or Trim(Text1(1).Text) = "" Then
                        MsgBox "Indique la fecha y hora de la finalizacion del proceso", vbExclamation
                        Cad = "NO"
                    Else
                        If CDate(Text1(0).Text) < vParamAplic.FechaActiva Then
                            MsgBox "Menor que fecha activa", vbExclamation
                            Cad = "NO"
                        Else
                            If CDate(Text1(0).Text) >= vParamAplic.FechaActivaMasUno Then
                                MsgBox "Mayor maxima fecha permitida", vbExclamation
                                Cad = "NO"
                            Else
                                Cad = ""
                            End If
                        End If
                    End If
                End If
            End If
            
            
            If Cad = "" Then
                If Not DatosOKCierre Then Cad = "No"
            End If
            If Cad = "" Then
                
                Cad = DevuelveDesdeBD(conAri, "tipoOliva", "vallalmazaraproceso", "id", CStr(ID))    'Oliva arbol o terra
                CadenaConsulta = "Tipo oliva: "
                If Cad = "1" Then
                    CadenaConsulta = CadenaConsulta & "TERRA"
                ElseIf Cad = "2" Then
                    CadenaConsulta = CadenaConsulta & "ARBEQUINA"
                Else
                    CadenaConsulta = CadenaConsulta & "ARBOL"
                End If
                
                
                
                
                
                Cad = DevuelveDesdeBD(conAri, "max(horamovi)", "proddepositoshco", "numdeposito", CStr(cboDepo.ItemData(cboDepo.ListIndex)))
                If Cad <> "" Then
                    If CDate(Cad) > CDate(Text1(0).Text & " " & Text1(1).Text) Then
                        If MsgBox("Fecha menor que movimiento en el deposito. ¿Continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then
                            Screen.MousePointer = vbDefault
                            Exit Sub
                        End If
                    End If
                End If
                                
                Cade1 = ""
                If Me.txtLote.Text <> "VACIO" Then
                    Cad = "numdeposito<>" & Me.cboDepo.ItemData(cboDepo.ListIndex) & " and numlote=" & DBSet(txtLote.Text, "T") & " AND 1"
                    Cad = DevuelveDesdeBD(conAri, "numdeposito", "proddepositos", Cad, "1")
                    If Cad <> "" Then
                        Cade1 = "*** Nuevo Lote(varios depositos) ** "
                    Else
                        Cade1 = DevuelveDesdeBD(conAri, "nomolturar", "proddepositos", "numdeposito", Me.cboDepo.ItemData(cboDepo.ListIndex))
                        If Cade1 = "1" Then
                            Cade1 = "*** Nueva  muestra ** "
                        Else
                            Cade1 = "Mismo lote!!!!"
                        End If
                    End If
                    Cade1 = Space(20) & Cade1
                End If
                'Preguntaremos si cerramos el parte
                Cad = CadenaConsulta & vbCrLf & vbCrLf & "Almazara." & vbCrLf & "Nº Albaranes: " & Data1.Recordset.RecordCount & vbCrLf & "Kg olivas:  " & Text2(0).Text
                Cad = Cad & vbCrLf & "Litros producidos: " & Text2(1).Text & vbCrLf & "Destino: " & cboDepo.Text & vbCrLf
                Cad = Cad & "       " & "Lote :"
                'Veremos si el deposito esta vacio. Crearemos nuevo lote
                'Veremos la capaciadad que hay mas la estimada
                Set miRsAux = New ADODB.Recordset
                miRsAux.Open "Select * from proddepositos where numdeposito=" & cboDepo.ItemData(cboDepo.ListIndex), conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If IsNull(miRsAux!numLote) Then
                    Cad = Cad & "*** Nuevo ****"
                    CadenaConsulta = "0"
                
                Else
                    Cad = Cad & miRsAux!numLote & Cade1 & vbCrLf
                    CadenaConsulta = miRsAux!Litros
                End If
                Cad = Cad & vbCrLf & "Litros deposito: " & Format(CadenaConsulta, FormatoCantidad) & "   Maximo: " & miRsAux!Capacidad
                CadenaConsulta = CCur(CadenaConsulta) + ImporteFormateado(Text2(1).Text)
                If CCur(CadenaConsulta) > miRsAux!Capacidad Then Cad = Cad & vbCrLf & "         --EXCEDE--"
                
                'Fecha hora
                Cad = Cad & vbCrLf & "FIN proceso:    " & Text1(0).Text & " " & Text1(1).Text
                'Duracion
                Cad = Cad & vbCrLf & "Duracion: " & (HorasProceso \ 60) & ":" & Format(HorasProceso Mod 60, "00")
                
                If Text1(2).Text <> 0 Then
                    If ImporteFormateado(Text1(2).Text) <> 0 Then Cad = Cad & vbCrLf & "INCREMENTO BODEGA " & Text1(2).Text
                End If
                miRsAux.Close
                Set miRsAux = Nothing
                Cad = Cad & vbCrLf & vbCrLf & vbCrLf & "¿Cerrar proceso almazara?"
                If MsgBox(Cad, vbQuestion + vbYesNoCancel) = vbYes Then
                    NumRegElim = 0
                    'hacemos el proceso de obtencion de aceite
                    If cerrarProcesoMolturacion Then
                        Me.Refresh
                        Espera 0.5
                        conn.Execute "commit"
                        Espera 0.5
                        'Si hemos creado un coupage, lo cerramos
                        If NumRegElim > 0 Then
                            frmProduVarios.Intercambio = NumRegElim & "||1|"
                            frmProduVarios.Opcion = 5
                            frmProduVarios.Show vbModal
                            
                            Cad = "UPDATE proddepositos SET noMolturar=0 WHERE numdeposito=" & cboDepo.ItemData(cboDepo.ListIndex)  'permitimos una proxima molturacion
                            conn.Execute Cad
                            
                            
                            Espera 0.5
                        End If
                        
                        CadenaDesdeOtroForm = "OK"
                        Unload Me
                    End If
                End If
                
                
             End If
        Case 4 'Modificar Existencia Real (Introducir Valores Reales)
            CargaTxtAux False, False
            PonerModo 2
    End Select
    Screen.MousePointer = vbDefault
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
    Set miRsAux = Nothing
End Sub

Private Sub cmdCancelar_Click()
    On Error GoTo ECancelar
    
     
    lblInfInv.Caption = ""
    Select Case Modo
        Case 1 'Buscar
            LimpiarCampos
            PonerModo 2
            CargaTxtAux False, False
            CargaGrid
        Case 2
            'Salir sin cerrar parte
            If MsgBox("Desea salir sin cerrar el proceso de almazara?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
            Unload Me
        Case 4  ' 4: Modificar
            PonerModo 2
            CargaTxtAux False, False
            CargaGrid
    End Select
    
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Sub cboDepo_Click()
    If cboDepo.ListIndex < 0 Then
        Me.txtLote.Text = ""
    Else
        CadenaConsulta = DevuelveDesdeBD(conAri, "Numlote", "proddepositos", "numDeposito", cboDepo.ItemData(cboDepo.ListIndex))
        If CadenaConsulta = "" Then
            'NUEVO. No existe lote
            txtLote.Text = "VACIO"
            txtLote.FontBold = True
            txtLote.ForeColor = vbRed
        Else
            txtLote.Text = CadenaConsulta
            txtLote.FontBold = False
            txtLote.ForeColor = vbBlack
        End If
    End If
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not Data1.Recordset.EOF And gridCargado And Modo = 4 Then
       CargaTxtAux True, True
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If Me.Tag = "NO" Then
        Me.Tag = ""
        BotonModificar
        PonerFoco txtAux
    End If
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmppal.Icon
    Me.Tag = "NO"
    
    'ICONOS de La toolbar
    With Toolbar1
        .ImageList = frmppal.imgListComun
        'ASignamos botones
        .Buttons(1).Image = 1   'Buscar
        .Buttons(4).Image = 4 'Modificar
        .Buttons(5).Image = 15 'Salir
    End With

   
    LimpiarCampos   'Limpia los campos TextBox
    PrimeraVez = True
    Text1(2).Text = "0.00"

    CadenaDesdeOtroForm = ""
    Set miRsAux = New ADODB.Recordset
    CadenaDesdeOtroForm = "Select numDeposito from proddepositos where numdeposito<>18 and DepositoVtaDirecta<>2 "   'El 100 es para la molturacion (no poner aqui)
    miRsAux.Open CadenaDesdeOtroForm, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    While Not miRsAux.EOF
        
        cboDepo.AddItem "Deposito " & miRsAux!NumDeposito
        cboDepo.ItemData(NumRegElim) = miRsAux!NumDeposito
        NumRegElim = NumRegElim + 1
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    
    CadenaDesdeOtroForm = ""
    PonerModo 2
    CargaGrid
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid()
Dim I As Byte
Dim SQL As String
Dim IncrementoBodega As Currency

On Error GoTo ECarga

    gridCargado = False

    SQL = MontaSQLCarga()
    CargaGridGnral DataGrid1, Me.Data1, SQL, PrimeraVez
    
    PrimeraVez = False
        
    'numalbar,fechaalb,codprove,numlinea,codartic,nomartic,numlotes,cantidad
    DataGrid1.Columns(0).Caption = "NºAlb"
    DataGrid1.Columns(0).Width = 850
    DataGrid1.Columns(1).Caption = "Fecha"
    DataGrid1.Columns(1).Width = 1150
  
    DataGrid1.Columns(2).Caption = "Cod"
    DataGrid1.Columns(2).Width = 800
    DataGrid1.Columns(2).NumberFormat = "0000"
        
    DataGrid1.Columns(3).Caption = "Proveedor"
    DataGrid1.Columns(3).Width = 2600
    
    'Cod artic
    DataGrid1.Columns(4).Caption = "Art."
    DataGrid1.Columns(4).Width = 1100
    
    'Cod artic
    DataGrid1.Columns(5).Caption = "Oliva"
    DataGrid1.Columns(5).Width = 2400
    
    DataGrid1.Columns(6).Caption = "Kilos"
    DataGrid1.Columns(6).Width = 1100
    DataGrid1.Columns(6).Alignment = dbgRight
    
    DataGrid1.Columns(7).Caption = "Rdto"
    DataGrid1.Columns(7).Width = 800
    DataGrid1.Columns(7).Alignment = dbgRight
    DataGrid1.Columns(7).NumberFormat = "00.00"
    
    DataGrid1.Columns(8).Caption = "Real"
    DataGrid1.Columns(8).Width = 800
    DataGrid1.Columns(8).Alignment = dbgRight
    DataGrid1.Columns(8).NumberFormat = "00.00"
    
    DataGrid1.Columns(9).Caption = "Litros"
    DataGrid1.Columns(9).Width = 1100
    DataGrid1.Columns(9).Alignment = dbgRight
    DataGrid1.Columns(9).NumberFormat = "00.00"
    
    For I = 0 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(I).AllowSizing = False
        DataGrid1.Columns(I).Locked = True
    Next I
    
    DataGrid1.ScrollBars = dbgAutomatic
    
    IncrementoBodega = 0
    If Text1(2).Text <> "" Then IncrementoBodega = ImporteFormateado(Text1(2).Text)
    
    SQL = "SELECT sum(numlotes+0),sum( Round(((numlotes + 0) * (Cantidad + " & TransformaComasPuntos(CStr(IncrementoBodega))
    SQL = SQL & ")) / 100, 2)) from tmpnlotes where codusu = " & vUsu.Codigo
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic
    Me.Text2(0).Text = "": Me.Text2(1).Text = ""
    If Not miRsAux.EOF Then
        Me.Text2(0).Text = Format(miRsAux.Fields(0), FormatoCantidad)
        Me.Text2(1).Text = Format(miRsAux.Fields(1), FormatoCantidad)
    End If
    miRsAux.Close
    gridCargado = True
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub


'Esta funcion sustituye a LlamaLineas
Private Sub CargaTxtAux(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        txtAux.Top = 290
        txtAux.visible = visible
    Else
        DeseleccionaGrid Me.DataGrid1
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
                txtAux.Text = Data1.Recordset!Cantidad
                txtAux.Locked = False
        End If

        If DataGrid1.Row < 0 Then
            alto = DataGrid1.Top + 220
        Else
            alto = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + 20
        End If
        
        'Fijamos altura y posición Top
        '-------------------------------
        txtAux.Top = alto
        txtAux.Height = DataGrid1.RowHeight
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        txtAux.Left = DataGrid1.Columns(8).Left + 130 'codalmac
        txtAux.Width = DataGrid1.Columns(8).Width - 10
        
        'Los ponemos Visibles o No
        '--------------------------
        txtAux.visible = visible
    End If
    PonerFoco txtAux
    
    If visible Then
        txtAux.TabIndex = 2
        txtAux.SelStart = 0
        txtAux.SelLength = Len(txtAux.Text)
    Else
        txtAux.TabIndex = 5
    End If
End Sub




Private Sub ImageCombo1_Change()

End Sub



Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), 3
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    If Index = 0 Then
        PonerFormatoFecha Text1(0)
    ElseIf Index = 1 Then
        PonerFormatoHora Text1(1)
    Else
        If Not PonerFormatoDecimal(Text1(2), 4) Then Text1(2).Text = "0"
        CargaGrid
   

    End If
End Sub
Private Sub CalcularIncremento()

End Sub
Private Sub txtAux_GotFocus()
    ConseguirFocoLin txtAux
End Sub

Private Sub TxtAux_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo EKeyD
    If KeyCode = 38 Or KeyCode = 40 Then
        ModificarExistencia
    End If

    Select Case KeyCode
        Case 38 'Desplazamieto Fecha Hacia Arriba
                If DataGrid1.Bookmark > 0 Then
                    DataGrid1.Bookmark = DataGrid1.Bookmark - 1
                End If
        
'                If DataGrid1.Row > 0 Then
'                    DataGrid1.Row = DataGrid1.Row - 1
'                    CargaTxtAux True, True
''                elseif
'                End If
        Case 40 'Desplazamiento Flecha Hacia Abajo
                PasarSigReg
                
    End Select
EKeyD:
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub txtAux_KeyPress(KeyAscii As Integer)

'    KEYpress KeyAscii
    
   If KeyAscii = 13 Then 'ENTER
        ModificarExistencia
        PasarSigReg
   ElseIf KeyAscii = 27 Then
        cmdCancelar_Click 'ESC
   End If
End Sub


Private Sub txtAux_LostFocus()
    If Screen.ActiveControl.Name = "cmdCancelar" Then Exit Sub
    With txtAux
        .Text = Trim(.Text)
        If .Text = "" Then Exit Sub
        'Formato tipo 1: Decimal(12,2)
        
        If Not PonerFormatoDecimal(txtAux, 4) Then .Text = ""
    End With

End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 'Busqueda
            BotonBuscar
        Case 4 'Modificar
            BotonModificar
        Case 5 'Salir
            Unload Me
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
       
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    'Modo 2. Hay datos y estamos visualizandolos
    'b = (Kmodo = 2)
   
    
    b = (Modo = 0)
    PonerBotonCabecera b
   
    If Modo = 2 Then
        Me.lblIndicador.Caption = "Cerrar"
    Else
        Me.lblIndicador.Caption = "Modificar"
    End If

           
    b = Modo <> 0 And Modo <> 2 And Modo <> 4

    b = (Modo = 1)
    Toolbar1.Buttons(1).Enabled = Not b
    Toolbar1.Buttons(4).Enabled = Not b And (Not (Modo = 0 Or Modo = 4))

    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
End Sub


Private Function MontaSQLCarga() As String
Dim Incremento As Currency

'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim SQL As String


    Incremento = 0
    If Text1(2).Text <> "" Then Incremento = ImporteFormateado(Text1(2).Text)
        

    SQL = "SELECT NumAlbar , FechaAlb, tmpnlotes.codProve, nomprove, codartic, NomArtic, numlotes, cantidad "
    If Incremento < 0 Then
        SQL = SQL & " - "
    Else
        SQL = SQL & " + "
    End If
    SQL = SQL & DBSet(Abs(Incremento), "N")
    SQL = SQL & ", Cantidad,"
    SQL = SQL & " Round(((numlotes + 0) * (Cantidad" & IIf(Incremento < 0, "-", "+")
    SQL = SQL & DBSet(Abs(Incremento), "N")
    SQL = SQL & " )) / 100, 2) "
    SQL = SQL & " from tmpnlotes,sprove where tmpnlotes.codprove="
    SQL = SQL & " sprove.codprove and codusu = " & vUsu.Codigo
    
    MontaSQLCarga = SQL
End Function


Private Sub BotonBuscar()

    Modo = 4
    cmdCancelar_Click
    
End Sub


Private Sub BotonModificar()
    If Data1.Recordset.EOF Then Exit Sub
    PonerModo 4
    CargaTxtAux True, True
    PonerFoco txtAux
End Sub


Private Function DatosOk() As Boolean

'Solo se actualiza el campo de Existencia Real
    
    txtAux.Text = Trim(txtAux.Text)
    
    If txtAux.Text <> "" And EsNumerico(txtAux.Text) Then
        If PonerFormatoDecimal(txtAux, 4) Then
            DatosOk = True
        Else
            DatosOk = False
        End If
        'DatosOk = True
    Else
        DatosOk = False
    End If
    
    If Not DatosOk Then Exit Function
    
End Function



Private Sub PonerBotonCabecera(b As Boolean)
    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    If b Then Me.lblIndicador.Caption = ""
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Function ActualizarExistencia(Cantidad As String) As Boolean
Dim C As String
    C = "UPDATE tmpnlotes SET cantidad=" & DBSet(Cantidad, "N")
    C = C & " WHERE codusu =" & vUsu.Codigo & " AND numalbar=" & DBSet(Data1.Recordset!NumAlbar, "T")
    C = C & " AND fechaalb =" & DBSet(Data1.Recordset!FechaAlb, "F") & " AND codprove=" & Data1.Recordset!codProve
    ActualizarExistencia = EjecutaSQL(conAri, C, True)
    
End Function


Private Sub PasarSigReg()
'Nos situamos en el siguiente registro
    If DataGrid1.Bookmark < Data1.Recordset.RecordCount Then
'        DataGrid1.Row = DataGrid1.Row + 1
        DataGrid1.Bookmark = DataGrid1.Bookmark + 1
    ElseIf DataGrid1.Bookmark = Data1.Recordset.RecordCount Then
        PonerFocoBtn Me.cmdAceptar
    End If
End Sub


Private Function ModificarExistencia() As Boolean
Dim NumReg As Long
Dim Indicador As String

    If DatosOk Then
        If ActualizarExistencia(txtAux.Text) Then
            TerminaBloquear
            NumReg = Data1.Recordset.AbsolutePosition
            CargaGrid
            If SituarDataPosicion(Data1, NumReg, Indicador) Then

            End If
            ModificarExistencia = True
        Else
            ModificarExistencia = False
        End If
    End If
End Function


Private Function cerrarProcesoMolturacion() As Boolean
    conn.BeginTrans
    If cerrarProcesoAlm Then
        cerrarProcesoMolturacion = True
        conn.CommitTrans
    Else
        cerrarProcesoMolturacion = False
        conn.RollbackTrans
        NumRegElim = 0
    End If
End Function


Private Function cerrarProcesoAlm() As Boolean
Dim CP1 As cPartidas  'La ppal me la guardo por si hay coupage ir ams rapido
Dim cP2 As cPartidas
Dim cStock As cStock
Dim cD As cDeposito
Dim Kilos As Currency
Dim cLot As cLotaje
Dim NuevaPartida As Boolean
Dim Cad As String
Dim RT As ADODB.Recordset
Dim vT As CTiposMov
Dim articuloTalco As String
Dim NumDeposito As Integer
Dim KilosTotales As Currency
Dim AuxCantidad As Currency
Dim LoteMasDeUnDeposito As Boolean
Dim RendimientoAplicado As Currency

    On Error GoTo eCerrarProcesoAlm
    
    cerrarProcesoAlm = False
    
    'Generara una entrada en el deposito, o añadira la cantidad
    Set CP1 = New cPartidas
    Set cStock = New cStock
    Set cD = New cDeposito
    Set cLot = New cLotaje
    




    NumDeposito = cboDepo.ItemData(cboDepo.ListIndex)
    'Tres casos
    '1.- Vacio. Todo igual
    '2.- Trasegado, movido, producido.. hay una marca en el deposito
    '   que es: NoMolturar     0. Podemos molturar  1. Habra que hacer lo que hacia
    '3. Lo que hacia. Deposito nuevo y couoage
    'Si el deposito YA ha trasegado, movido... entonces creará uno nuevo(como hacia antes) un
    
    LoteMasDeUnDeposito = False
    If txtLote.Text <> "VACIO" Then
            
        Cad = "numdeposito<>" & Me.cboDepo.ItemData(cboDepo.ListIndex) & " and numlote=" & DBSet(txtLote.Text, "T") & " AND 1"
        Cad = DevuelveDesdeBD(conAri, "numdeposito", "proddepositos", Cad, "1")
                  
        If Cad <> "" Then
            LoteMasDeUnDeposito = True
            Cad = "1"
        Else
            Cad = DevuelveDesdeBD(conAri, "NoMolturar", "proddepositos", "numdeposito", CStr(NumDeposito))
        End If
        If Cad = "1" Then
            'No molturar. Hay que hacer muestra nueva
            NumDeposito = 100
        Else
            'Dejamos molturar en el
        
        End If
    End If
    If Not cD.LeerDatos(NumDeposito, True) Then Err.Raise 513, , "Leyendo deposito: " & NumDeposito
    
    
    If cD.NumDeposito = 100 And cD.numLote <> "" Then Err.Raise 513, , "Deposito produccion NO esta vacio"
    
    
    If cD.numLote = "" Then
        'Asignaremos un lote nuevo.. y una nueva partida
        Set vT = New CTiposMov
        vT.Leer "LOV"
        CP1.numLote = vT.ConseguirContador(vT.TipoMovimiento)
        
        CP1.numLote = "MOSTRA" & CP1.numLote & "-"
        If Month(Now) < 10 Then
            CP1.numLote = CP1.numLote & Year(Now)
        Else
            CP1.numLote = CP1.numLote & Year(Now) + 1
        End If
        NuevaPartida = True
       
        cD.idPartida = CP1.Siguiente
        cD.Kilos = 0
        cD.numLote = CP1.numLote
        vT.IncrementarContador vT.TipoMovimiento
        
    Else
        NuevaPartida = False
        If Not CP1.Leer(cD.idPartida) Then Err.Raise 513, , "Leyendo partida: " & cD.idPartida
    End If
    
    Kilos = ImporteFormateado(Text2(1).Text) * 0.916
    
    'Incrementamos la cantidad del deposito
    If NuevaPartida Then
        cD.VariacionKilosDeposito Kilos   'Si tenia o no, da lo mismo. Esta sumando los nuevos
        If Not cD.InsertarEnDeposito2(10, Text1(0).Text & " " & Text1(1).Text, Format(ID, "0000")) Then Err.Raise 513, , "Insertando datos nuevos deposito: " & cD.numLote
    Else
        cD.VariacionKilosDeposito Kilos   'Si tenia o no, da lo mismo. Esta sumando los nuevos
        cD.InsertarEnHco 10, Text1(0).Text & " " & Text1(1).Text, Format(ID, "0000"), Kilos
    End If
        
    'Metemos los moviimientos
    'Tanto en smoval como en smovallotes
    cStock.DetaMov = "MLT" 'Molturacion
    cStock.codAlmac = 1
    
    'Si es de tarra el articulo sera el de parametros en artMolturaTerra , si no en articMolturacion
    Cad = DevuelveDesdeBD(conAri, "tipoOliva", "vallalmazaraproceso", "id", CStr(ID))   'Oliva arbol o terra
    If Cad = "1" Then
        Cad = "artMolturaTerra"
    ElseIf Cad = "2" Then
        Cad = "artMoltArbequina"
    Else
        Cad = "articMolturacion"
    End If
    Cad = DevuelveDesdeBD(conAri, Cad, "vallparam", "1", "1")
    cStock.codartic = Cad
    cStock.Documento = Format(ID, "00000")
    cStock.Fechamov = Text1(0).Text
    cStock.HoraMov = Text1(0).Text & " " & Text1(1).Text
    cStock.Importe = 0
    cStock.LineaDocu = 1
    cStock.tipoMov = "E"
    
    cStock.Trabajador = TrabajadorParte
    cLot.codAlmac = cStock.codAlmac
    cLot.codartic = cStock.codartic
    cLot.DetaMov = cStock.DetaMov
    cLot.Documento = cStock.Documento
    cLot.Fechamov = cStock.Fechamov
    cLot.HoraMov = cStock.HoraMov
    cLot.LineaDocu = cStock.LineaDocu
    cLot.tipoMov = 1 'entrada
    cLot.numLote = CP1.numLote
    cLot.ProvCliTra = cStock.Trabajador
  
    cStock.Cantidad = Kilos
    cLot.Cantidad = Kilos
    
    
    If Not cStock.ActualizarStock Then Err.Raise 513, , "Actualizando stock"
    If Not cLot.InsertarLote Then Err.Raise 513, , "Actualizando stock lotes"
    
    
    
    If NuevaPartida Then
        CP1.Cantidad = 0  'Luego los incremento
        CP1.codAlmac = cStock.codAlmac
        CP1.codartic = cStock.codartic
        CP1.codProve = 1
        CP1.Fecha = cStock.Fechamov
        CP1.NumAlbar = cStock.Documento
        'cp.NUmlote = lo he asginado arriba
        If Not CP1.Insertar Then Err.Raise 513, , "Creando partida "
    End If
    CP1.IncrementarCantidad Kilos
    
    
    
    
    
    
    
    'El orujo
    '------------------------------------------------------
    'Que es el resto de producto
    'Crearemos partida nueva
    Set cP2 = New cPartidas
    
    
    '07/12/16
    ' Al orujo le sumamos el agua del decanter
    'Es litros /hora. Con lo cual habra que multiplicar litroshora x tiempor
    Cad = DevuelveDesdeBD(conAri, "AguaDecanter", "vallalmazaraproceso ", "id", CStr(ID))
    AuxCantidad = 0
    If Cad <> "" Then
        If Cad <> "0" Then
            AuxCantidad = CCur(Cad)
            Kilos = (HorasProceso Mod 60)
            Kilos = Round2((Kilos / 60), 2) 'minutos en decimal
            
            Kilos = (HorasProceso \ 60) + Kilos 'Horas proceso + minutos   DECIAML
            
            AuxCantidad = Round(AuxCantidad * Kilos, 2)
           
        End If
    End If
    
    Kilos = ImporteFormateado(Text2(0).Text) - ImporteFormateado(Text2(1).Text)
    Kilos = Kilos + AuxCantidad
    
    
    cP2.Cantidad = Kilos
    cP2.numLote = "Orujo" & Format(ID, "0000")
    cP2.codAlmac = cStock.codAlmac
    Cad = DevuelveDesdeBD(conAri, "articOrujo", "vallparam", "1", "1")
    cStock.codartic = Cad
    cStock.Cantidad = Kilos
    cStock.LineaDocu = 1
    cP2.codartic = cStock.codartic
    cP2.codProve = 1
    cP2.Fecha = cStock.Fechamov
    cP2.NumAlbar = cStock.Documento
    'cp.NUmlote = lo he asginado arriba
    If Not cP2.Insertar Then Err.Raise 513, , "Creando partida orujo "
    If Not cStock.ActualizarStock Then Err.Raise 513, , "Actualizando stock orujo"
    
    
    
    
    
    
    'Las lineas de oliva, movimiento nuevo para dar de baja
    Cad = "select * from tmpnlotes where codusu=" & vUsu.Codigo
    Set RT = New ADODB.Recordset
    RT.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cStock.Documento = Format(ID, "00000")
    cStock.Importe = 0
    cStock.LineaDocu = 1
    cStock.Trabajador = TrabajadorParte
    
    RendimientoAplicado = 0
    If Text1(2).Text <> "" Then RendimientoAplicado = ImporteFormateado(Text1(2).Text)
      
    
    While Not RT.EOF
        cStock.codartic = RT!codartic
        Kilos = CCur(TransformaPuntosComas((RT!numlotes)))
        cStock.Cantidad = Kilos
        cStock.tipoMov = "S"
        Cad = "Alb: " & RT!NumAlbar & " " & RT!FechaAlb & " .> " & RT!NomArtic
        If Not cStock.ActualizarStock Then Err.Raise 513, , "Actualizando stock: " & Cad
        
        
        
        'El rendimiento aplicado a cada linea
        Cad = " numalbar=" & DBSet(RT!NumAlbar, "T") & " AND codartic = " & DBSet(RT!codartic, "T") & " AND 1"
        Cad = DevuelveDesdeBD(conAri, "entrada", "vallentradacamionlineas", Cad, "1 ORDER BY entrada desc")  'No deberia haber mas de una
        If Cad = "" Then Err.Raise 513, , "NO se encuentra la entrada de camion para el albarán: " & RT!NumAlbar & " Art: " & RT!codartic
         
        
        Cad = " WHERE  numalbar=" & DBSet(RT!NumAlbar, "T") & " AND entrada =" & Cad
        Cad = ",rdtoRea=" & DBSet(RT!Cantidad, "N") & Cad
        Cad = "UPDATE vallentradacamionlineas set rendimiento=" & DBSet(RT!Cantidad + RendimientoAplicado, "N") & Cad
        
        
        
        If Not EjecutaSQL(conAri, Cad, False) Then Err.Raise 513, , "Actualizando albaran entrada camion: " & Cad
        
        
        RT.MoveNext
    Wend
    RT.Close
    
    
    
    
    
    'La dosis es en kilos hora
    '-----------------------------------------------------------------
    articuloTalco = ""
    Cad = "Select dosis , numlote from vallalmazaraproceso where id =" & ID
    RT.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RT.EOF Then
        If Not IsNull(RT!dosis) Then
            articuloTalco = DevuelveDesdeBD(conAri, "arttalco", "vallparam", "1", "1")
            If articuloTalco = "" Then
                MsgBox "Pone dosis talco y no esta configurado proceso continua sin tratar talco", vbExclamation
            Else
                'Los minutos los pasamos a decimal
                Kilos = (HorasProceso Mod 60)
                Kilos = Round2((Kilos / 60), 2) 'minutos en decimal
                
                Kilos = (HorasProceso \ 60) + Kilos 'Horas proceso + minutos   DECIAML
                
                Kilos = Round2((Kilos * RT!dosis), 2)
                Set cP2 = Nothing
                Set cP2 = New cPartidas
                
                If Not cP2.LeerDesdeArticulo(articuloTalco, 1, CStr(RT!numLote)) Then Err.Raise 513, , "Leyendo lote talco: " & articuloTalco & " " & RT!numLote
               
                cP2.Cantidad = Kilos
                cStock.codartic = cP2.codartic
                cStock.Cantidad = Kilos
                cStock.LineaDocu = 1
                cStock.tipoMov = "S"
                'cp.NUmlote = lo he asginado arriba
                cP2.IncrementarCantidad -Kilos
                If Not cStock.ActualizarStock Then Err.Raise 513, , "Actualizando stock talco"
            
                Set cLot = Nothing
                Set cLot = New cLotaje
                                
                  cLot.codAlmac = cStock.codAlmac
                  cLot.codartic = cStock.codartic
                  cLot.DetaMov = cStock.DetaMov
                  cLot.Documento = cStock.Documento
                  cLot.Fechamov = cStock.Fechamov
                  cLot.HoraMov = cStock.HoraMov
                  cLot.LineaDocu = cStock.LineaDocu
                  cLot.tipoMov = 0  'salida
                  cLot.numLote = RT!numLote
                  cLot.ProvCliTra = cStock.Trabajador
                  cLot.SubLinea = 1
                  cLot.Cantidad = cStock.Cantidad
                  
                
                  If Not cLot.InsertarLote Then Err.Raise 513, , "Actualizando stock talco"
                  Set cLot = Nothing
            
        
            End If
        End If
    End If
    RT.Close
    
    
    NumRegElim = 0
    If NumDeposito = 100 Then
        
        If vT Is Nothing Then
            Set vT = New CTiposMov
            vT.Leer "LOV"
        End If
    
        'Ahora tengo que HACER un nuevo COUPAGE automatico entre el deposito 100 y el deposito destino
        Cad = "Select max(codigo) from olicoupage"
        RT.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        NumRegElim = DBLet(RT.Fields(0), "N") + 1
        RT.Close
        
        Set cP2 = Nothing
        Set cP2 = New cPartidas
        cP2.numLote = vT.ConseguirContador(vT.TipoMovimiento)
        cP2.numLote = "MOSTRA" & cP2.numLote & "-"
        If Month(Now) < 10 Then
            cP2.numLote = cP2.numLote & Year(Now)
        Else
            cP2.numLote = cP2.numLote & Year(Now) + 1
        End If
        
        
        Cad = "INSERT INTO olicoupage(codigo,codartic,fecha,descripcion,YaCreado,codalmac,numlote,Deposito) VALUES ("
        Cad = Cad & NumRegElim & ",'" & CP1.codartic & "'," & DBSet(cStock.HoraMov, "FH") & ",'"
        Cad = Cad & "Molturacion. ID: " & CP1.NumAlbar & "',0, " & CP1.codAlmac & "," & DBSet(cP2.numLote, "T") & ","
        Cad = Cad & cboDepo.ItemData(cboDepo.ListIndex) & ")"
        conn.Execute Cad
        
        
        
        
        'Leemos lo que hay en el deposito AHORA
        If Not cD.LeerDatos(cboDepo.ItemData(cboDepo.ListIndex), True) Then Err.Raise 513, , "Leyendo deposito coupage: " & cboDepo.ItemData(cboDepo.ListIndex)
        Kilos = Round(ImporteFormateado(Text2(1).Text) * 0.916, 2) 'De la nueva produccion + lo que habia
        KilosTotales = cD.Kilos + Kilos
        Cad = "INSERT INTO olicoupagelin(codigo,codartic,kilos) VALUES (" & NumRegElim & ","
        Cad = Cad & DBSet(CP1.codartic, "T") & "," & DBSet(KilosTotales, "N") & ")"
        conn.Execute Cad
        
        
        
        Set cP2 = Nothing
        Set cP2 = New cPartidas
        If Not cP2.Leer(cD.idPartida) Then Err.Raise 513, , "Leyendo partida deposito: " & cD.NumDeposito & " ->" & cD.numLote
        
        'Las dos linea del coupage con lotes
        '---------------------------------------------------------
        
        Cad = "INSERT INTO olicoupagelinlotes(codigo,codartic,linea,numlote,cantlote,fincuba,deposito) VALUES ("
        Cad = Cad & NumRegElim & ",'" & cP2.codartic & "',1," & DBSet(cP2.numLote, "T") & "," & DBSet(cD.Kilos, "N")
        
        Cad = Cad & ",1," & cD.NumDeposito & "),(" & NumRegElim & "," & DBSet(CP1.codartic, "T") & ",2,"
        Cad = Cad & DBSet(CP1.numLote, "T") & "," & DBSet(Kilos, "N") & ",1,100)"
        conn.Execute Cad
        
        
        
        
        
        
        
            
        vT.IncrementarContador vT.TipoMovimiento
        
    
    
    End If
    
    
    
    
    'En la tablaproceso guardo el dato de deposito y fecha fin
    Cad = "UPDATE vallalmazaraproceso SET "
    Cad = Cad & " HoraFin =" & DBSet(cStock.HoraMov, "H")
    Cad = Cad & ",  FechaFin =" & DBSet(cStock.Fechamov, "F")
    Cad = Cad & ",  deposito=" & cboDepo.ItemData(cboDepo.ListIndex)
    'Litros producidos y kilos utlizados
    Cad = Cad & ", kilos =" & DBSet(Text2(0).Text, "N")
    Cad = Cad & ", Litros =" & DBSet(Text2(1).Text, "N")
    If articuloTalco = "" Then
        articuloTalco = "NULL"
    Else
        articuloTalco = DBSet(articuloTalco, "T")
    End If
    Cad = Cad & ", articuloTalco =" & articuloTalco
    
    'Lote y articulo
    Cad = Cad & ", loteproducido =" & DBSet(CP1.numLote, "T")
    Cad = Cad & ", artproducido =" & DBSet(CP1.codartic, "T")
     
    
    
    Cad = Cad & " WHERE id=" & Me.ID
    conn.Execute Cad
    
  
    If LoteMasDeUnDeposito Then
        'Los depositos que tenga el numero de lote NO se podra molturar
        Cad = "Update proddepositos set nomolturar=1 where "
        Cad = Cad & "numdeposito<>" & Me.cboDepo.ItemData(cboDepo.ListIndex) & " and numlote=" & DBSet(txtLote.Text, "T")
        conn.Execute Cad
    End If
    
    cerrarProcesoAlm = True
    
eCerrarProcesoAlm:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
    Set CP1 = Nothing
    Set cP2 = Nothing
    Set cStock = Nothing
    Set cD = Nothing
    Set cLot = Nothing
    Set RT = Nothing
    Set vT = Nothing
End Function




Private Function DatosOKCierre() As Boolean
Dim TFin As Date
Dim TIni As Date
Dim TipoOlivaPRoceso As Byte
Dim C As String


    'Hora fecha
    DatosOKCierre = False
    If Not EsFechaOK(Me.Text1(0).Text) Then Exit Function
    If Not EsHoraOK(Me.Text1(1).Text) Then Exit Function
    
    'Vemaos si la hora de inicio >= que la de fin
    TFin = CDate(Text1(0).Text & " " & Text1(1).Text)
    TIni = "01/01/1900 00:00:01"
    TipoOlivaPRoceso = 127
    CadenaDesdeOtroForm = "Select Fecha ,horainicio,tipooliva from vallalmazaraproceso where Id = " & ID
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open CadenaDesdeOtroForm, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        MsgBox "No se encuentra el registro ID: " & ID, vbExclamation
    Else
        If Not IsNull(miRsAux!Fecha) Then
            If Not IsNull(miRsAux!horainicio) Then TIni = miRsAux!Fecha & " " & Format(miRsAux!horainicio, "hh:mm:ss")
        End If
        TipoOlivaPRoceso = miRsAux!tipooliva
    End If
    miRsAux.Close
    
    If TIni = "01/01/1900 00:00:01" Then Exit Function
    
    
    
    'Que los tipos de oliva son igual
    If TipoOlivaPRoceso = 0 Then
        C = "DALT"
    ElseIf TipoOlivaPRoceso = 1 Then
        C = "TERRA"
    Else
        C = "ARBEQUINA"
    End If
    
    C = "codartic <> " & DBSet(C, "T") & " ANd ID"
    C = DevuelveDesdeBD(conAri, "codartic", "vallalmazaraprocesoalb", C, CStr(ID))
    If C <> "" Then
        If MsgBox("Tipo oliva distinto del de destino." & vbCrLf & C & vbCrLf & "¿Continuar?", vbYesNoCancel + vbQuestion) <> vbYes Then Exit Function
    End If
    
    
    
    
    HorasProceso = 0
    If TFin <= TIni Then
        MsgBox "Fecha fin mayor o igual que fecha inicio", vbExclamation
        Exit Function
    Else
        If TFin < vParamAplic.FechaActiva Then
            MsgBox "Fecha fin menor  que fecha activa", vbExclamation
            Exit Function
        Else
            HorasProceso = DateDiff("n", TIni, TFin)
            
            NumRegElim = HorasProceso \ 60
            
            If NumRegElim > 5 Then
                If MsgBox("Total horas proceso: " & NumRegElim & ":" & Format(HorasProceso Mod 60, "00") & ".   ¿Continuar?", vbQuestion + vbYesNo) <> vbYes Then Exit Function
            Else
                
            End If
        End If
    End If
    
    C = DevuelveDesdeBD(conAri, "codtrabaAlm", "vallalmazaraproceso", "id", CStr(ID)) 'Oliva arbol o terra
    If C = "" Then
        MsgBox "Error trabajador en el parte ", vbExclamation
        Exit Function
    End If
    TrabajadorParte = CInt(C)
    C = DevuelveDesdeBD(conAri, "codtraba", "straba", "codtraba", C) 'Oliva arbol o terra
    If C = "" Then
        MsgBox "No existe el trabajador del parte.     Codigo trabajador: " & TrabajadorParte, vbExclamation
        Exit Function
    End If
    
    
    
    
    
    DatosOKCierre = True
End Function
