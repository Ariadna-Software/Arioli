VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmProdTrazaVer 
   Caption         =   "Form1"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13875
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   13875
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.TreeView TreeView1 
      Height          =   2655
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   4683
      _Version        =   327682
      LabelEdit       =   1
      Style           =   7
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
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   240
      Top             =   6960
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      Caption         =   ""
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
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Index           =   1
      Left            =   8400
      TabIndex        =   6
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Index           =   0
      Left            =   6960
      TabIndex        =   5
      Top             =   6960
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   375
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   360
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   9840
      TabIndex        =   3
      Tag             =   "Cantidad|N|N|||spartidas|cantotal|0.00||"
      Text            =   "Text1"
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   8160
      TabIndex        =   2
      Tag             =   "Partida|N|N|0||spartidas|id|||"
      Text            =   "Text1"
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      Tag             =   "Articulo|T|N|||spartidas|codartic|||"
      Text            =   "Text1"
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Tag             =   "Lote|T|N|||spartidas|numlote|||"
      Text            =   "Text1"
      Top             =   360
      Width           =   1695
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   2775
      Left            =   240
      TabIndex        =   7
      Top             =   3960
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
         SubItemIndex    =   4
         Text            =   "Cantidad"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Lote"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmProdTrazaVer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim Modo As Byte    '0 Buscar   1 Ver resultado
Dim i As Integer
Dim cadB As String

Dim IT As ListItem

Private Sub Command1_Click(Index As Integer)
    If Index = 0 Then
        If Modo = 1 Then
            'ACeptar busqueda
            cadB = ObtenerBusqueda(Me, False)
            If cadB <> "" Then
                cadB = "select * from spartidas WHERE " & cadB
                PonerCadenaBusqueda
            End If
        Else
            BotonBuscar
        End If
    Else
        If Modo = 2 Then
            Unload Me
        Else
            cadB = ObtenerBusqueda(Me, False)
            If cadB = "" Then
                Unload Me
            Else
                BotonBuscar
            End If
        End If
    End If
End Sub


Private Sub BotonBuscar()
    limpiar Me

    Me.ListView2.ListItems.Clear
    Me.TreeView1.Nodes.Clear
    PonerModo 1
    PonerFoco Text1(0)
End Sub

Private Sub PonerModo(xModo As Byte)
    
    
    
    Modo = xModo
    For i = 0 To Me.Text1.Count - 1
        BloquearTxt Text1(i), xModo = 2
    Next
    
    
End Sub

Private Sub Form_Load()
    Data1.ConnectionString = Conn
    BotonBuscar
End Sub



Private Sub PonerCadenaBusqueda()
Screen.MousePointer = vbHourglass

    On Error GoTo EEPonerBusq


    Data1.RecordSource = cadB
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro  ", vbInformation
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
    PonerModo 1
    Screen.MousePointer = vbDefault
End Sub

Private Sub PonerCampos()
Dim SQL As String
Dim cP As cPartidas


    PonerCamposForma Me, Data1

    Set cP = New cPartidas
    Conn.Execute "DELETE FROM tmptraza"
    If cP.LeerDesdeArticulo(Text1(1).Text, Data1.Recordset!codalmac, Data1.Recordset!NUmlote) Then
        cP.TrazbilidadDesdeVenta False, False
        
    End If
    
    Text2.Text = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", Text1(1).Text, "T")
    
    Set miRsAux = New ADODB.Recordset
    SQL = DBLet(Data1.Recordset!NumAlbar, "T")
    If SQL <> "" Then
        'AQUI VERE SI ES UN COUPAGE, PRODUCCION u otro
        If Val(Data1.Recordset!CodProve) = 0 And Mid(SQL, 1, 2) = "NP" Then
                'PRODUCCION
                'Cargar datos produccion
                CargarDatosProduccion
        Else
            Stop
        
        End If
        
    
    End If
    'Todos cargaran si hay ventas
    CargarDatosVentas
    Set miRsAux = Nothing
    
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), Modo
End Sub



Private Sub CargarDatosProduccion()
Dim C As String
Dim N
Dim Contador As Integer
Dim Nivel As Integer
Dim Padre As String

    C = "select tmptraza.*,nomartic from tmptraza,sartic where codartic=artic2 AND codusu =" & vUsu.Codigo
    miRsAux.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Nivel = -1
    While Not miRsAux.EOF
        
        
        
        
            'El albaran de compra del lote
            If miRsAux!nivle = 0 Then
                C = miRsAux!artic2 & " " & miRsAux!NomArtic & " [" & miRsAux!NUmlote2 & "]"
                'C = DevuelveCadena(C, miRsAux!cantutili)
                Contador = TreeView1.Nodes.Count + 1
                Set N = TreeView1.Nodes.Add(, , "C" & Contador, C)
                
                
                
                
                PonAlbaran N.Key, miRsAux!NUmlote2, miRsAux!artic2
                            
                Nivel = 0
                
            Else
                If Nivel <> miRsAux!nivle Then
                    Padre = N.Key
                    Nivel = miRsAux!nivle
                End If
                C = miRsAux!artic2 & " " & miRsAux!NomArtic & " [" & miRsAux!NUmlote2 & "]"
                'C = DevuelveCadena(C, miRsAux!cantutili)
                Contador = TreeView1.Nodes.Count + 1
                Set N = TreeView1.Nodes.Add(Padre, tvwChild, "C" & Contador, C)
                PonAlbaran N.Key, miRsAux!NUmlote2, miRsAux!artic2
            End If
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If Not N Is Nothing Then N.EnsureVisible
    
    'If ElAceite <> "" Then CargaCoupageRecursivo RecuperaValor(ElAceite, 1), RecuperaValor(ElAceite, 2), N.Key, EsCou
    
End Sub


Private Sub PonAlbaran(Key1 As String, NUmlote As String, vArtic As String)
Dim RT As ADODB.Recordset
Dim cad As String
Dim N
    Set RT = New ADODB.Recordset
    cad = "select * from spartidas where numlote=" & DBSet(NUmlote, "T") & " and codartic='" & vArtic & "'"
    RT.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cad = ""
    If Not RT.EOF Then
        cad = "select * from slifpc where numalbar=" & DBSet(RT!NumAlbar, "T") & " and codartic=" & DBSet(RT!codartic, "T")
        cad = cad & " AND codprove =" & RT!CodProve
        
    End If
    RT.Close
    
        
    If cad <> "" Then
        RT.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If RT.EOF Then
            RT.Close
            
            cad = Replace(cad, " slifpc ", " slialp ")
            RT.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            
            
        End If
        
        If Not RT.EOF Then
            cad = "Alb: " & RT!NumAlbar & "   " & RT!CodProve & "  " '& RT!nomprove
            Set N = TreeView1.Nodes.Add(Key1, tvwChild, "C" & CStr(TreeView1.Nodes.Count + 1), cad)
            
        
            
        End If
        RT.Close
    End If
    
    Set RT = Nothing
End Sub

Private Function DevuelveCadena(Cadena As String, Cantidad As Currency) As String
Dim J As Integer
    
        
    DevuelveCadena = Format(Cantidad, FormatoCantidad)
    J = 100 - Len(DevuelveCadena) - Len(Cadena)
    If J < 0 Then J = 0
    DevuelveCadena = Cadena & Space(J) & DevuelveCadena
    
End Function



Private Sub CargarDatosVentas()
Dim C As String
    C = "select concat(scafac.codtipom,scafac.numfactu) lafact,scafac.fecfactu,codclien,nomclien,cantidad "
    C = C & " from slifaclotes,scafac  where"
    C = C & " slifaclotes.codTipoM = scafac.codTipoM And slifaclotes.NumFactu = scafac.NumFactu"
    C = C & " and slifaclotes.fecfactu=scafac.fecfactu AND numlote=" & DBSet(Text1(0).Text, "T")
    C = C & " ORDER BY fecfactu,lafact"
    miRsAux.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = ListView2.ListItems.Add()
        IT.Text = miRsAux!lafact
        IT.SubItems(1) = miRsAux!FecFactu
        IT.SubItems(2) = miRsAux!CodClien
        IT.SubItems(3) = miRsAux!nomclien
        IT.SubItems(4) = miRsAux!Cantidad
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).BackColor = vbWhite
End Sub




