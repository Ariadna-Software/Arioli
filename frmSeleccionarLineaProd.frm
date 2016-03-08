VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProdSeleccionarLineaProd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccionar referencia a producir"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir componentes"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   7080
      Width           =   2055
   End
   Begin VB.CommandButton cmdPlusminus 
      Caption         =   "+"
      Height          =   255
      Index           =   2
      Left            =   7560
      TabIndex        =   15
      Top             =   240
      Width           =   255
   End
   Begin VB.CommandButton cmdPlusminus 
      Caption         =   "-"
      Height          =   255
      Index           =   3
      Left            =   7560
      TabIndex        =   14
      Top             =   510
      Width           =   255
   End
   Begin VB.CommandButton cmdPlusminus 
      Caption         =   "-"
      Height          =   255
      Index           =   1
      Left            =   6360
      TabIndex        =   13
      Top             =   510
      Width           =   255
   End
   Begin VB.CommandButton cmdPlusminus 
      Caption         =   "+"
      Height          =   255
      Index           =   0
      Left            =   6360
      TabIndex        =   12
      Top             =   240
      Width           =   255
   End
   Begin VB.TextBox txtNumero 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   5640
      TabIndex        =   10
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtNumero 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   6720
      TabIndex        =   8
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox txtNumeroEntero 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   8040
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtArticulo 
      Height          =   285
      Index           =   0
      Left            =   120
      MaxLength       =   16
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtDescArticulo 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   0
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   480
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Index           =   1
      Left            =   8160
      TabIndex        =   3
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Index           =   0
      Left            =   6960
      TabIndex        =   2
      Top             =   7080
      Width           =   1095
   End
   Begin MSComctlLib.ListView lwp 
      Height          =   6015
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   10610
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Referencia"
         Object.Width           =   2364
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripion"
         Object.Width           =   6526
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Caj.Pal"
         Object.Width           =   1234
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "UdsCaj"
         Object.Width           =   1234
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "SUMA"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "STOCK"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "PDTE"
         Object.Width           =   1341
      EndProperty
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   735
      Left            =   7920
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Palets"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   2
      Left            =   5640
      TabIndex        =   11
      Top             =   240
      Width           =   525
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Cajas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   1
      Left            =   6840
      TabIndex        =   9
      Top             =   240
      Width           =   465
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Unidades"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   0
      Left            =   8280
      TabIndex        =   7
      Top             =   240
      Width           =   780
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Articulo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   660
   End
   Begin VB.Image imgArticulo 
      Height          =   240
      Index           =   0
      Left            =   840
      Picture         =   "frmSeleccionarLineaProd.frx":0000
      Top             =   240
      Width           =   240
   End
End
Attribute VB_Name = "frmProdSeleccionarLineaProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmMtoArticulos As frmAlmArticulos
Attribute frmMtoArticulos.VB_VarHelpID = -1

Dim PrimeraVez As Boolean
Dim SQL As String

Private Sub cmdPlusminus_Click(Index As Integer)
    If Index < 2 Then
        UnaPaletMas Index = 0
    Else
        UnaCajaMas Index = 2
    End If
    PonerFoco Me.txtNumeroEntero(0)
End Sub

Private Sub Command1_Click(Index As Integer)
    
    CadenaDesdeOtroForm = ""
    If Index = 0 Then

           If Me.txtArticulo(0).Text = "" Or txtDescArticulo(0).Text = "" Or Me.txtNumeroEntero(0).Text = "" Then
                MsgBox "Campos obligatorios", vbExclamation
                Exit Sub
            End If
            
            If Val(Me.txtNumeroEntero(0).Text) = 0 Then
                MsgBox "indique cantidad estimada", vbExclamation
                Exit Sub
            End If
            
            CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "conjunto", "sartic", "codartic", Me.txtArticulo(0).Text, "T")
            If CadenaDesdeOtroForm = "" Then CadenaDesdeOtroForm = "0"
            If Val(CadenaDesdeOtroForm) = "0" Then
                MsgBox "Articulo debe ser de produccion", vbExclamation
                Exit Sub
            End If
            
            CadenaDesdeOtroForm = Me.txtArticulo(0).Text & "|" & txtDescArticulo(0).Text & "|" & Me.txtNumeroEntero(0).Text & "|"

    End If
    Unload Me
End Sub

Private Sub Command2_Click()



    'Imprimir componentes
    SQL = ""
    If Me.txtArticulo(0).Text = "" Then
        If Not Me.lwp.SelectedItem Is Nothing Then SQL = lwp.SelectedItem.Text & "|" & lwp.SelectedItem.SubItems(1) & "|"
    Else
        SQL = Me.txtArticulo(0).Text & "|" & Me.txtDescArticulo(0).Text & "|"
    End If
    If SQL = "" Then Exit Sub
    
    
    LlamaImprimirGral "{sartic.codartic}=""" & RecuperaValor(SQL, 1) & """", "", 0, "prevalidacionprod.rpt", RecuperaValor(SQL, 2)
    SQL = ""
    
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        CargaDatos
    End If
End Sub

Private Sub Form_Load()
    
    Me.Icon = frmppal.Icon
    PrimeraVez = True
End Sub


Private Sub CargaDatos()
    
Dim IT As ListItem
Dim Pal_ped As Currency
Dim Pal_stock As Currency
Dim ArticulosDeB As String
Dim B As Boolean



    Set miRsAux = New ADODB.Recordset
    
    
    
    'Enero 2013
    SQL = "select codartic from sliped WHERE codalmac =" & vParamAplic.AlmacenB
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    ArticulosDeB = "|"
    While Not miRsAux.EOF
        ArticulosDeB = ArticulosDeB & miRsAux!codartic & "|"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    
    SQL = "Select * from tmpplanning "
    'SQL = SQL & " WHERE codartic IN  (select codartic from sliped WHERE codalmac <>" & vParamAplic.AlmacenB & ")"
    SQL = SQL & " order by nomartic"
    
    Me.lwp.ListItems.Clear
    
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not miRsAux.EOF

        
        Set IT = lwp.ListItems.Add
        IT.Text = miRsAux!codartic
        IT.SubItems(1) = miRsAux!NomArtic
        If InStr(1, ArticulosDeB, IT.Text & "|") > 0 Then
            'Si tiene lineas de B
            IT.ListSubItems(1).ForeColor = vbRed
            IT.SubItems(1) = "* " & IT.SubItems(1)
        End If
        
        
        
        'cajaspal,unicajas,enproduccion,
        'estas dos no se ven
        IT.SubItems(2) = miRsAux!CajasPal
        IT.ListSubItems(2).ForeColor = vbBlue
        If miRsAux!CajasPal = 0 Then
            IT.ForeColor = vbRed
            IT.ListSubItems(2).ForeColor = vbRed
        End If
        
        IT.SubItems(3) = miRsAux!UniCajas
        IT.ListSubItems(3).ForeColor = vbBlue
        
        'Totales
        'ud_pedidos ,stock
        If IT.ForeColor <> vbRed Then
            Pal_ped = miRsAux!ud_pedidos / miRsAux!UniCajas 'cuantas cajas
            Pal_ped = Round(Pal_ped / miRsAux!CajasPal, 2)
            IT.SubItems(4) = Format(Pal_ped, FormatoImporte)
            
            Pal_stock = miRsAux!stock / miRsAux!UniCajas 'cuantas cajas
            Pal_stock = Round(Pal_stock / miRsAux!CajasPal, 2)
            IT.SubItems(5) = Format(Pal_stock, FormatoImporte)
            
            'diferencia
            Pal_ped = Pal_ped - Pal_stock
            If Pal_ped > 0 Then
                IT.SubItems(6) = Format(Pal_ped, FormatoImporte)
            Else
                IT.SubItems(6) = " "
            End If
        Else
            IT.SubItems(4) = " "
            IT.SubItems(5) = IT.SubItems(4)
            IT.SubItems(6) = IT.SubItems(4)
        End If
            
        miRsAux.MoveNext
    Wend
    miRsAux.Close
End Sub




Private Sub CargaDatosOLD()
Dim C As String
Dim Produccion As Long
Dim N As Node

    On Error GoTo ECargaDatos
    
'
'
'    C = "select prodcab.codigo,feccreacion,prodlin.codartic,nomartic,cantesti,idlin from "
'    C = C & " prodcab,prodlin,sartic where prodcab.codigo =prodlin.codigo and prodlin.codartic=sartic.codartic"
'    If YaProducidas Then
'        C = C & " AND estado=10 ORDER BY codigo,idlin"
'    Else
'        C = C & " and producido=0 and estado=0 ORDER BY codigo,idlin"
'    End If
'    Set miRsAux = New ADODB.Recordset
'    miRsAux.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    Produccion = -1
'    While Not miRsAux.EOF
'        If Produccion <> miRsAux!Codigo Then
'            Produccion = miRsAux!Codigo
'            C = "PROD." & Format(miRsAux!Codigo, "0000000") & "  " & Format(miRsAux!feccreacion, "dd/mm/yyyy")
'            Set N = TreeView1.Nodes.Add(, , "C" & Produccion, C)
'        End If
'        'El articulo
'        C = miRsAux!codArtic & "  " & miRsAux!NomArtic & "  (" & Format(miRsAux!cantesti, FormatoCantidad) & ")"
'        '
'        Set N = TreeView1.Nodes.Add("C" & Produccion, tvwChild, , C)
'        N.Tag = miRsAux!idlin  'linea de produccion
'        N.EnsureVisible
'        miRsAux.MoveNext
'    Wend
'    miRsAux.Close

ECargaDatos:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set miRsAux = Nothing
End Sub


Private Sub frmMtoArticulos_DatoSeleccionado(CadenaSeleccion As String)
    SQL = CadenaSeleccion
End Sub

Private Sub imgArticulo_Click(Index As Integer)

    SQL = ""
    Set frmMtoArticulos = New frmAlmArticulos
    frmMtoArticulos.DatosADevolverBusqueda2 = "@1@" 'Abrimos en Modo Busqueda
    frmMtoArticulos.Show vbModal
    Set frmMtoArticulos = Nothing
    If SQL <> "" Then
        txtArticulo(0).Text = RecuperaValor(SQL, 1)
        Me.txtDescArticulo(0).Text = RecuperaValor(SQL, 2)
        
        
        
    End If
End Sub

Private Sub lwp_DblClick()
Dim Ca As Long
    If Me.lwp.SelectedItem Is Nothing Then Exit Sub
    
    If Trim(lwp.SelectedItem.SubItems(1)) <> "" Then
        
        txtArticulo(0).Text = lwp.SelectedItem.Text
        Me.txtDescArticulo(0).Text = lwp.SelectedItem.SubItems(1)
        
        Me.txtNumero(0).Tag = Val(lwp.SelectedItem.SubItems(2))
        Me.txtNumero(1).Tag = Val(lwp.SelectedItem.SubItems(3))
        
        Ca = 0
        If Trim(lwp.SelectedItem.SubItems(2)) <> "" Then
            If Trim(lwp.SelectedItem.SubItems(6)) <> "" Then
                Ca = Val(lwp.SelectedItem.SubItems(3))
                If Ca = 0 Then Ca = 1
                Ca = Val(lwp.SelectedItem.SubItems(2)) * Ca
                Ca = CCur(lwp.SelectedItem.SubItems(6)) * Ca

                Ca = ((Ca - 1) \ Val(lwp.SelectedItem.SubItems(3))) + 1
                Ca = Ca * Val(lwp.SelectedItem.SubItems(3))
                
                
            End If
        End If
        If Ca = 0 Then
            Me.txtNumeroEntero(0).Text = ""
            Me.txtNumero(0).Text = ""
            Me.txtNumero(1).Text = ""

        Else
            Me.txtNumeroEntero(0).Text = Ca
            PonerCajasPalets
            
        End If
         PonerFoco txtNumeroEntero(0)
    End If
   
End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtArticulo_GotFocus(Index As Integer)
    ConseguirFoco txtArticulo(Index), 3
End Sub

Private Sub txtArticulo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
    
End Sub


Private Sub txtArticulo_LostFocus(Index As Integer)
Dim T As String
    
    If Index = 0 Then
        txtNumero(0).Tag = 0
        txtNumero(1).Tag = 0
    End If
        
    txtArticulo(Index).Text = Trim(txtArticulo(Index).Text)
    If txtArticulo(Index).Text = "" Then
        'EN blanco
        txtDescArticulo(Index).Text = ""
        Exit Sub
    End If
    
    
    T = "codartic"
    SQL = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", txtArticulo(Index).Text, "T", T)
    Me.txtDescArticulo(Index).Text = SQL
    If SQL = "" Then
        MsgBox "No existe el artículo : " & txtArticulo(Index).Text, vbExclamation
    Else
        txtArticulo(Index).Text = T
        PonerPaletsUnicajas
    End If
    
    SQL = ""
    
End Sub





Private Sub txtNumeroEntero_GotFocus(Index As Integer)
    ConseguirFoco txtNumeroEntero(Index), 3
End Sub

Private Sub txtNumeroEntero_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpressGnral KeyAscii, 2, True
End Sub


Private Sub PonerCajasPalets()
Dim C As Long
Dim I As Integer
    C = 0
    If Me.txtNumeroEntero(0).Text <> "" Then C = Val(ImporteFormateado(txtNumeroEntero(0).Text))
    
    Me.txtNumero(0).Text = "": Me.txtNumero(1).Text = ""
    
    If C > 0 Then
        If Me.txtNumero(1).Tag <> 0 Then
            C = ((C - 1) \ txtNumero(1).Tag) + 1
            txtNumero(1).Text = C
            If Me.txtNumero(0).Tag <> 0 And C > 0 Then
                C = ((C - 1) \ txtNumero(0).Tag) + 1
                txtNumero(0).Text = C
            End If
        End If
    End If
    
End Sub


Private Sub PonerPaletsUnicajas()
    SQL = DevuelveDesdeBD(conAri, "pal_udbas * pal_udalt", "sarti4", "codartic", Me.txtArticulo(0).Text, "T")
    If SQL = "" Then SQL = "0"
    Me.txtNumero(0).Tag = Val(SQL) 'uids palet
    
    SQL = DevuelveDesdeBD(conAri, "unicajas", "sartic", "codartic", txtArticulo(0).Text, "T")
    If SQL = "" Then SQL = "0"
    Me.txtNumero(1).Tag = Val(SQL) 'unicajas
    
End Sub

Private Sub txtNumeroEntero_LostFocus(Index As Integer)
        txtNumeroEntero(Index).Text = Trim(txtNumeroEntero(Index).Text)
    If txtNumeroEntero(Index).Text = "" Then Exit Sub
    
    If Not PonerFormatoEntero(txtNumeroEntero(Index)) Then
        txtNumeroEntero(Index).Text = ""
        PonerFoco txtNumeroEntero(Index)
    Else
        PonerCajasPalets
    End If
End Sub

Private Sub UnaPaletMas(Sumar As Boolean)
Dim N As Long

    If Not Sumar Then
        If Val(txtNumero(0).Text) = 0 Then Exit Sub
    End If
    
    'Si no estan fijados los valores del cajas palet, uds caja no hacemos nada
    If Me.txtNumero(0).Tag = 0 Or Me.txtNumero(0).Tag = 0 Then Exit Sub
    
    N = Val(txtNumero(0).Text)
    If Sumar Then
        N = N + 1
    Else
        N = N - 1
    End If
    txtNumero(0).Text = N
    
    'Calculamos cuantas cajas y uds son
    N = N * txtNumero(0).Tag
    txtNumero(1).Text = N
    N = N * txtNumero(1).Tag 'uds por caja
  
    Me.txtNumeroEntero(0).Text = N
End Sub


Private Sub UnaCajaMas(Sumar As Boolean)
Dim N As Long

    If Not Sumar Then
        If Val(txtNumero(1).Text) = 0 Then Exit Sub
    End If
    
    'Si no estan fijados los valores del cajas palet, uds caja no hacemos nada
    If Me.txtNumero(0).Tag = 0 Then Exit Sub
    
    
    
    'Calculamosuds son
    
    N = Val(txtNumero(1).Text)
    
    If Sumar Then
        N = N + 1
    Else
        N = N - 1
    End If
    txtNumero(1).Text = N
    N = N * txtNumero(1).Tag 'uds por caja
  
    Me.txtNumeroEntero(0).Text = N
End Sub



