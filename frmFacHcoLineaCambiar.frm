VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacHcoLineaCambiar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   8445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameStcok 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1935
      Left            =   120
      TabIndex        =   35
      Top             =   3000
      Width           =   8175
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   11
         Left            =   240
         TabIndex        =   37
         Text            =   "Text1"
         Top             =   360
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   1215
         Left            =   3000
         TabIndex        =   36
         Top             =   360
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   2143
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Lin"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Lote"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "REAL"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Image imgLotes 
         Height          =   240
         Index           =   2
         Left            =   3720
         Picture         =   "frmFacHcoLineaCambiar.frx":0000
         ToolTipText     =   "Eliminar"
         Top             =   120
         Width           =   240
      End
      Begin VB.Image imgLotes 
         Height          =   240
         Index           =   1
         Left            =   3360
         Picture         =   "frmFacHcoLineaCambiar.frx":0A02
         ToolTipText     =   "Modificar"
         Top             =   120
         Width           =   240
      End
      Begin VB.Image imgLotes 
         Height          =   240
         Index           =   0
         Left            =   3000
         Picture         =   "frmFacHcoLineaCambiar.frx":1404
         ToolTipText     =   "Añadir"
         Top             =   120
         Width           =   240
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00004000&
         BorderWidth     =   3
         Height          =   1815
         Left            =   120
         Top             =   0
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Lotes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2280
         TabIndex        =   39
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Stock"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   38
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Frame FrameLinea 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   2055
      Left            =   120
      TabIndex        =   22
      Top             =   840
      Width           =   8295
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   12
         Left            =   6840
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   240
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   1080
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   720
         Width           =   6135
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   240
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   840
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1800
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   2760
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   4200
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   4920
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   5640
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   3840
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1560
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "+"
         Height          =   255
         Left            =   840
         TabIndex        =   23
         Top             =   240
         Width           =   210
      End
      Begin VB.Label Label1 
         Caption         =   "Hectogrado"
         Height          =   255
         Index           =   13
         Left            =   6840
         TabIndex        =   41
         Top             =   1320
         Width           =   855
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00004000&
         BorderWidth     =   3
         Height          =   2055
         Left            =   120
         Top             =   0
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label1 
         Caption         =   "Artículo"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   34
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Ampliación"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   33
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Palets *"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   32
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Cajas"
         Height          =   255
         Index           =   3
         Left            =   960
         TabIndex        =   31
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Cantidad"
         Height          =   255
         Index           =   4
         Left            =   1800
         TabIndex        =   30
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Precio"
         Height          =   255
         Index           =   5
         Left            =   2760
         TabIndex        =   29
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Dto1"
         Height          =   195
         Index           =   6
         Left            =   4320
         TabIndex        =   28
         Top             =   1320
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dto2"
         Height          =   195
         Index           =   7
         Left            =   4920
         TabIndex        =   27
         Top             =   1320
         Width           =   345
      End
      Begin VB.Label Label1 
         Caption         =   "Importe"
         Height          =   255
         Index           =   8
         Left            =   5640
         TabIndex        =   26
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Ori."
         Height          =   195
         Index           =   9
         Left            =   3840
         TabIndex        =   25
         Top             =   1320
         Width           =   315
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Siguiente"
      Height          =   375
      Left            =   5640
      TabIndex        =   11
      Top             =   5040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame FrameEliminar 
      Height          =   615
      Left            =   120
      TabIndex        =   18
      Top             =   4920
      Visible         =   0   'False
      Width           =   2055
      Begin VB.CommandButton cmdEliminar 
         Height          =   375
         Left            =   1440
         Picture         =   "frmFacHcoLineaCambiar.frx":1E06
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   180
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Eliminar la linea ?"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7080
      TabIndex        =   17
      Top             =   5040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   240
      TabIndex        =   12
      Top             =   120
      Width           =   8175
      Begin VB.Label Label6 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   6480
         TabIndex        =   40
         Top             =   210
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Albaran: "
         Height          =   255
         Index           =   11
         Left            =   3720
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   4440
         TabIndex        =   15
         Top             =   210
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   840
         TabIndex        =   14
         Top             =   210
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Factura: "
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label lblIndicador 
      Height          =   375
      Left            =   2280
      TabIndex        =   21
      Top             =   5040
      Width           =   2535
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   7200
      Y1              =   2880
      Y2              =   2880
   End
End
Attribute VB_Name = "frmFacHcoLineaCambiar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public NumFactu As Long
Public Codtipom As String
Public Fecfactu As Date
Public NumAlbar As Long
Public Codtipoa As String
Public Numlinea As Integer   '    -1 =  Nueva



Private WithEvents frmL As frmAlmPartidas
Attribute frmL.VB_VarHelpID = -1
Private WithEvents frmArt  As frmAlmArticulos
Attribute frmArt.VB_VarHelpID = -1

Private cA As CArticulo
Private Cajaspalet As Integer
Private vCStock As CStock   '    Para cuando hagamos el movimiento ver si tenemos o no en stock
Private CantidadEnLotes As Currency

Dim RS As ADODB.Recordset
Dim cad As String


Dim Estado As Byte  '0- Modificando lineas
                    '1- Modificando stocks/lotes

Private Sub cmdAceptar_Click()
Dim Diferencia As Currency

    Select Case Estado
    Case 0
        'Pongo estado 1
        If DatosLineaCorrecto Then PonerEstado 1
            


    Case 1
        If ComprobarDatosStocks(Diferencia) Then
            If RealizarUpdate(Diferencia) Then
                CadenaDesdeOtroForm = "OK"
                Unload Me
            End If
    
        End If
    End Select
    
End Sub




Private Function DatosLineaCorrecto() As Boolean
Dim I As Integer

    DatosLineaCorrecto = False
    'OK
    cad = ""
    For I = 0 To Text1.Count - 2
        If I <> 2 Then
            If Text1(I).Text = "" Then
                cad = "Campos obligatorios"
                PonerFoco Text1(I)
                Exit For
            End If
        End If
    Next
    
    If vParamAplic.QUE_EMPRESA = 2 Then
        If Not Text1(12).Locked Then
            'hay que poner valor
            If Text1(12).Text = "" Then cad = cad & vbCrLf & "Ponga el hectogrado"
        End If
    End If
    
    
    If cad <> "" Then
        MsgBox cad, vbExclamation
        Exit Function
    End If
    
    
    
    
    DatosLineaCorrecto = True
    
End Function

Private Sub cmdCancelar_Click()
    Select Case Estado
    Case 0
        'Me salgo
        Unload Me


    Case 1
        
        If Me.Numlinea < 0 Then
            'Es nuevo.
            If Me.ListView2.ListItems.Count > 0 Then
                'Ha puesto LOTES. No le dejo volver atras
                MsgBox "Ya ha introducido lotes para el articulo", vbExclamation
                Exit Sub
            End If
        End If
        
        PonerEstado 0
    
    End Select
End Sub

Private Sub cmdEliminar_Click()
    'Eliminar la linea
    If Estado = 1 Then Exit Sub
    
    'Preguntamos si quiere eliminar linea
    If MsgBox("Seguro que desea eliminar la linea de la factura?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    If MsgBox("Continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    
    Conn.BeginTrans
    If Eliminar Then
        Conn.CommitTrans
        CadenaDesdeOtroForm = "OK"
        Unload Me
        
    Else
        Conn.RollbackTrans
    End If
    
    
    
End Sub

Private Sub Command1_Click()
    'Articulo
    cad = ""
    Set frmArt = New frmAlmArticulos
    frmArt.DatosADevolverBusqueda2 = "@1@" 'Poner en Modo busqueda
    frmArt.DeConsulta = True
    frmArt.ParaVenta = True
    frmArt.Show vbModal
    Set frmArt = Nothing
    If cad <> "" Then
        Text1(0).Text = RecuperaValor(cad, 1)
        Text1_LostFocus 0
        PonerFoco Text1(2)
    End If
End Sub

Private Sub Form_Activate()
Dim T1 As Single



    If Label2.Caption = "" Then
        Screen.MousePointer = vbHourglass
        
        
        Label1(13).visible = vParamAplic.QUE_EMPRESA = 2
        Text1(12).visible = vParamAplic.QUE_EMPRESA = 2
        
        
        T1 = Timer
        Label2.Caption = Codtipom & Format(NumFactu, "000000") & " de " & Format(Fecfactu, "dd/mm/yyyy")
        Label3.Caption = Codtipoa & Format(NumAlbar, "000000")
        
        Set cA = New CArticulo
        Set vCStock = New CStock
        
        'Bloquearemos codartic
        
        
        PonerEstado 0
        
        Me.Shape1.visible = False

        cmdEliminar.Enabled = False  'Para saber si ha cargado datos
        Me.Command1.visible = Numlinea < 0
        BloquearTxt Text1(0), Numlinea >= 0
        
        
        BloquearTxt Text1(12), True
        
        
        DoEvents
        If Numlinea >= 0 Then
            'Tenia datos
            PonerCampos
            PonerFoco Text1(2)
            lblIndicador.Caption = ""
            
            FrameEliminar.visible = Me.cmdAceptar.Enabled
        Else
            'Nuevo
            Text1(6).Tag = 0
            Text1(10).Text = "M"
        End If
       
        
        While Timer - T1 < 1.2
            Espera 0.1
        Wend
        Me.cmdAceptar.visible = True
        Me.cmdCancelar.visible = True
        Me.Shape1.visible = True
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Me.Icon = frmppal.Icon
    
    Shape1.Width = Me.FrameLinea.Width - 340
    Shape2.Width = Shape1.Width
    
    limpiar Me
        
    Label2.Caption = ""
End Sub

Private Function PonWhere() As String

    PonWhere = " WHERE numfactu=" & NumFactu & " AND fecfactu=" & DBSet(Fecfactu, "F") & " AND codtipom = '" & Codtipom
    PonWhere = PonWhere & "' and codtipoa ='" & Codtipoa & "' AND numalbar= " & NumAlbar
End Function


Private Sub PonerCampos()
Dim I As Integer
Dim IT As ListItem

    On Error GoTo ePonerCam
    
    lblIndicador.Caption = "slifac"
    lblIndicador.Refresh
    
    cad = "select codartic,nomartic,ampliaci,palets,0,cantidad,precioar,dtoline1,dtoline2,importel,origpre,codalmac, "
    If vParamAplic.QUE_EMPRESA = 2 Then
        cad = cad & " hectogrado"
        Text1(12).Text = ""
    Else
        cad = cad & " 1"
    End If
    cad = cad & " hecto from slifac "
    cad = cad & PonWhere
    cad = cad & " AND numlinea =" & Numlinea
    Set RS = New ADODB.Recordset
    RS.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
        For I = 0 To RS.Fields.Count - 3  'el ultimo campo es CODALMAC i hectogrado
            Select Case I
            Case 0, 1, 2, 10 'textos
                Text1(I).Text = DBLet(RS.Fields(I), "T")
                
                If I = 0 Then
                    If Not cA.LeerDatos(RS.Fields(0)) Then
                        'ERROR LEY ARTICULO..> NO DEBERIA PASAR NUNCA
                        
                    End If
                    If cA.UnidCaja = 0 Then cA.UnidCaja = 1
                End If
                
                
            Case 3, 4
                    'Enteros
                    Text1(I).Text = Format(DBLet(RS.Fields(I), "N"), "#,##0")
                    
            Case 7, 8
                    'decimales
                    Text1(I).Text = Format(DBLet(RS.Fields(I), "N"), FormatoDescuento)
            Case 6
                    Text1(I).Text = Format(DBLet(RS.Fields(I), "N"), FormatoPrecio)
            Case 5, 9
                    Text1(I).Text = Format(DBLet(RS.Fields(I), "N"), FormatoImporte)
                    
                    
                    If I = 5 Then Text1(4).Text = Format(RS.Fields(I) \ cA.UnidCaja, "#,##0")
                    
                    
            End Select
        
            Text1(I).Tag = DBLet(RS.Fields(I), "T")  'el valor sin formato
        Next I
        
        If vParamAplic.QUE_EMPRESA = 2 Then
            If RS!Hecto <> 1 Then
                Text1(12).Text = DBLet(RS!Hecto) * 100
                PonerFormatoDecimal Text1(12), 3
            End If
            BloquearTxt Text1(12), Text1(12).Text = ""
        End If
        vCStock.codalmac = RS!codalmac
        vCStock.codArtic = RS!codArtic
        vCStock.DetaMov = "S"
        vCStock.LineaDocu = Numlinea
        vCStock.Documento = Format(NumAlbar, "0000000")
        cmdEliminar.Enabled = True 'habilito ACEPTAR
    
    
    
    
    
    RS.Close
    
    'Las cajas
    lblIndicador.Caption = "Cajas y sctock"
    lblIndicador.Refresh
    
    Text1(11).Tag = cA.ExistenciaTotalAlmacenes
    Text1(11).Text = Format(Text1(11).Tag, FormatoCantidad)
    
    cad = DevuelveDesdeBD(conAri, "pal_udbas * pal_udalt", "sarti4", "codartic", cA.Codigo, "T")
    If cad = "" Then cad = "0"
    If Val(cad) = 0 Then cad = "1"
    Cajaspalet = Val(cad)
        
    
    'Buscare la fecha del albaran
    cad = "Select fechaalb from scafac1 " & PonWhere
    RS.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        'ERROR encontrando el albaran
        Err.Raise 513, "Leyendo scafac1", "No se ha encontrado el albaran. Error grave"
           
    Else
        vCStock.Fechamov = RS!FechaAlb
    
    End If
    RS.Close
    
    
    'Buscare el movimiento "
    lblIndicador.Caption = "smoval"
    lblIndicador.Refresh
    
    cad = "select * from smoval where codartic = " & DBSet(vCStock.codArtic, "T") & "  and codalmac  = " & DBSet(vCStock.codalmac, "N") & "  and fechamov  =  " & DBSet(vCStock.Fechamov, "F")
    cad = cad & " and document  = " & DBSet(vCStock.Documento, "T") & " and numlinea  =  " & DBSet(Numlinea, "N") & " and detamovi  = "
    'Detamovi: SERA SIEMPRE 0 'salida
    cad = cad & "0"
    
    RS.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        'ERROR encontrando el movimiento en smoval
        cad = ""
    Else
        cad = "Leyendo smoval"
        'Datos que faltan
        vCStock.HoraMov = RS!HoraMovi
        vCStock.Trabajador = RS!codigope
        
    End If
    RS.Close
    
    If cad = "" Then Err.Raise 513, "No se ha encontrado el movimiento asociado a la linea de factura."
        
    
    'LOTES
    lblIndicador.Caption = "Lotes"
    lblIndicador.Refresh
    
    
    ListView2.Tag = 0 ' Indicara si hemos cambiado cosas de los lotes pondremos un 1
    CantidadEnLotes = 0
    If cA.Trazabilidad Then
        'Aticulo de trazabilidad
        cad = "Select linea,numlote,cantidad from slifaclotes "
        cad = cad & PonWhere
        cad = cad & " AND numlinea =" & Numlinea & " order by LINEA"
        RS.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
            Set IT = ListView2.ListItems.Add()
            IT.Text = RS!linea
            IT.Bold = True  'Ya estaba.
            IT.SubItems(1) = RS!Numlote
            IT.SubItems(2) = Format(RS!Cantidad, FormatoCantidad)
            IT.SubItems(3) = Format(RS!Cantidad, FormatoCantidad)  'Cantidad que habia en un ppio
            
            CantidadEnLotes = CantidadEnLotes + RS!Cantidad
            RS.MoveNext
        Wend
        RS.Close
    
    End If
    
ePonerCam:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "", cad
         Me.cmdEliminar.Enabled = False
        
    End If
    lblIndicador.Caption = "Leyendo datos"
    lblIndicador.Refresh
    
    Set RS = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cA = Nothing
End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
    cad = CadenaSeleccion
End Sub

Private Sub frmL_DatoSeleccionado(CadenaSeleccion As String)
 
    cad = CadenaSeleccion
End Sub

Private Sub imgLotes_Click(Index As Integer)
Dim Aux As Currency
Dim IT As ListItem
Dim CantidadLineas As Currency
Dim Maximo As Currency
Dim J As Integer

    If Index <> 0 Then
        If Me.ListView2.SelectedItem Is Nothing Then
            MsgBox "Seleccione un lote", vbExclamation
            Exit Sub
        End If
    End If
    
    
    CantidadLineas = ImporteFormateado(Text1(5).Text)
    
    
    Select Case Index
    Case 0
        'INSERTAR
        If CantidadLineas - CantidadEnLotes = 0 Then
            MsgBox "Cantidad totalmente asignada", vbExclamation
            Exit Sub
        End If
        cad = ""
        Set frmL = New frmAlmPartidas
        frmL.DatosADevolverBusqueda = cA.Codigo
        frmL.Show vbModal
        Set frmL = Nothing
        
        If cad <> "" Then
            Aux = CCur(RecuperaValor(cad, 2))  'Cantidad desde PARTIDAS
            
            
            'Comprobaremos si ya ha metido el lote este
            cad = RecuperaValor(cad, 1)
            For J = 1 To ListView2.ListItems.Count  'Buscaremos el maximo
                If cad = ListView2.ListItems(J).SubItems(1) Then
                    MsgBox "El lotes ya está en esta linea", vbExclamation
                    Exit Sub
                End If
            Next J
        
        
            'Ha devuelto datos
            Set IT = ListView2.ListItems.Add()
            
            Maximo = 0
            For J = 1 To ListView2.ListItems.Count - 1 'Buscaremos el maximo
                If Val(ListView2.ListItems(J).Text) > Val(Maximo) Then Maximo = Val(ListView2.ListItems(J).Text)
            Next
            IT.Text = Val(Maximo) + 1
            IT.SubItems(1) = cad
            

            Maximo = CantidadLineas - CantidadEnLotes
                
            
            'Si hay mas de lo que necesito, pongo solamente lo que necesito
            If Aux > Maximo Then Aux = Maximo
            IT.SubItems(2) = Format(Aux, FormatoCantidad)
            IT.SubItems(3) = ""
                        
            
            CantidadEnLotes = CantidadEnLotes + Aux
            ListView2.Tag = 1
        End If
        
    Case 1
        'MODIFICAR
        'ModifcarCantidad
        'Va a modificar la cantidad del LOTE
        
        If CantidadLineas - CantidadEnLotes = 0 Then
            If MsgBox("Cantidad totalmente asignada.   ¿Continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
        
        J = 0
        Do
            Aux = CantidadLineas - CantidadEnLotes 'Sugerimos esta cantidad

                cad = "LOTE: " & ListView2.SelectedItem.SubItems(1) & vbCrLf
                cad = cad & "Cantidad actual: " & ListView2.SelectedItem.SubItems(2) & vbCrLf
                cad = cad & vbCrLf & "Introduzca la cantidad"
                Aux = ImporteFormateado(ListView2.SelectedItem.SubItems(2)) + Aux
                cad = InputBox(cad, "Lotaje", CCur(Aux))
                If cad = "" Then
                    'Ha cancelado
                    J = 2
                Else
                    If Not IsNumeric(cad) Then
                        MsgBox "Campo numérico", vbExclamation
                    Else
                        If InStr(1, cad, ",") > 0 Then
                            MsgBox "Introduzca el valor correctamente", vbExclamation
                        
                        Else
                            Maximo = CCur(TransformaPuntosComas(cad))
                            'Ultimas comprobaciones
                            Aux = ImporteFormateado(ListView2.SelectedItem.SubItems(2))
                            Aux = Maximo - Aux  'diferencia entre lo que habia y lo que hay

                          
                            If Aux > 0 Then
                                'Suma mas que las lineas
                                If CantidadLineas - CantidadEnLotes - Aux < 0 Then
                                    MsgBox "Sobrepasa la cantidad lineas", vbExclamation
                                Else
                                    Aux = -1 'Para que siga el proceso
                                End If
                                
                            End If
                            
                            If Aux <= 0 Then
                                'OK vamos a poner la cantidad y poner
                                Aux = ImporteFormateado(ListView2.SelectedItem.SubItems(2))
                                Aux = Maximo - Aux  'diferencia entre lo que habia y lo que hay
                                CantidadEnLotes = CantidadEnLotes + Aux
                            
                                ListView2.SelectedItem.SubItems(2) = Format(Maximo, FormatoCantidad)
                                ListView2.Tag = 1
                                J = 1
                            End If
                            
                        End If
                    End If
                End If
           
        Loop Until J > 0
        
        'Si maximo=2 HA cancelado
        If J = 2 Then Exit Sub
        
        
    
    Case 2
        'ELIMINAR
        cad = "Va a eliminar : " & vbCrLf & "Lote:   " & ListView2.SelectedItem.SubItems(1) & vbCrLf
        cad = cad & "Cantidad: " & ListView2.SelectedItem.SubItems(2) & vbCrLf & vbCrLf & "¿Continuar?"
        If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
            'La cantidad
            
            'Aux = ImporteFormateado(ListView2.SelectedItem.SubItems(3))
            Aux = ImporteFormateado(ListView2.SelectedItem.SubItems(2))
            CantidadEnLotes = CantidadEnLotes - Aux
        
            ListView2.ListItems.Remove ListView2.SelectedItem.Index
            ListView2.Tag = 1
        End If
    End Select
    
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    
    ConseguirFoco Text1(Index), 3
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub


Private Sub Text1_LostFocus(Index As Integer)
Dim b As Boolean
Dim I As Currency

    If Not PerderFocoGnral(Text1(Index), 3) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    With Text1(Index)
        Select Case Index
        Case 0
            'codartic NO se puede modificar
            If Not cA.LeerDatos(Text1(0).Text) Then
                cA.Codigo = ""
                Text1(0).Text = ""
                Text1(1).Text = ""
                Text1(11).Text = ""
                If vParamAplic.QUE_EMPRESA = 2 Then Text1(12).Text = ""
            Else
                Text1(0).Text = cA.Codigo
                Text1(1).Text = cA.Nombre
                Text1(11).Tag = cA.ExistenciaTotalAlmacenes
                Text1(11).Text = Format(Text1(11).Tag, FormatoCantidad)
                
                cad = DevuelveDesdeBD(conAri, "pal_udbas * pal_udalt", "sarti4", "codartic", cA.Codigo, "T")
                If cad = "" Then cad = "0"
                If Val(cad) = 0 Then cad = "1"
                Cajaspalet = Val(cad)
                                
                
                
                
            End If
            
            If vParamAplic.QUE_EMPRESA = 2 Then
                b = True
                If cA.Codigo <> "" Then

                    cad = "if(codtipar='05',1,0)+if(codfamia=6,1,0)" 'Los 05 o la familia 6
                    cad = DevuelveDesdeBD(conAri, cad, "sartic", "codartic", cA.Codigo, "T")
                    If cad = "" Then cad = 0
                    If Val(cad) > 0 Then b = False
                    
                    
                End If
            
                BloquearTxt Text1(12), b
            End If
        Case 3
            If PonerFormatoEntero(Text1(Index)) Then
                'Ok es un entero
                                
                'Palets
                If Text1(4).Text = "" Then
                    Text1(4).Text = Cajaspalet * Val(Text1(3).Text)
                    PonerFoco Text1(4)
                                 
                End If
                
            End If
            
            
        Case 4, 5
            
            If Index = 4 Then
                b = PonerFormatoEntero(Text1(4))
            Else
                b = PonerFormatoDecimal(Text1(4), 3)
            End If
            
            If b Then CantidadCajas Index = 5
            
        Case 6, 7, 8
            If Index = 6 Then
                b = PonerFormatoDecimal(Text1(Index), 2)
            Else
                b = PonerFormatoDecimal(Text1(Index), 4)
            End If
            
            
            If Index = 6 Then
                If Not b Then Text1(Index).Text = ""
                'Si cambia el precio pongo la M
                I = ImporteFormateado(CStr(Text1(6).Text))
                If I <> CCur(Text1(6).Tag) Then
                    Text1(10).Text = "M"
                Else
                    'Ha dejado el precio que estaba
                    Text1(10).Text = Text1(10).Tag
                End If
            End If
            
        Case 12
            'Hectogrado
            PonerFormatoDecimal Text1(12), 3
        End Select
    End With
    
    
    
    If (Index >= 4 Or Index <= 8) Or Index = 12 Then 'Cant., Precio, Dto1, Dto2
        Dim Dev As String
    
        Dev = CalcularImporte(Text1(5).Text, Text1(6).Text, Text1(7).Text, Text1(8).Text, vParamAplic.TipoDtos)
        
          
        I = 1
        If vParamAplic.QUE_EMPRESA = 2 Then
            If Not Text1(12).Locked Then
                If Text1(12).Text <> "" Then
                    I = ImporteFormateado(Text1(12).Text)
                    I = I / 100
                End If
             End If
        End If
        Text1(9).Text = ImporteFormateado(Dev) * I
        
        
        PonerFormatoDecimal Text1(9), 1
    End If
    
End Sub


Private Sub CantidadCajas(DeCantidadACajas As Boolean)
Dim V As Long
Dim V2 As Currency
    
        

    If DeCantidadACajas Then
        If Text1(5).Text = "" Then
            Text1(4).Text = ""
        Else
            V2 = ImporteFormateado(Text1(5).Text)
            Text1(4).Text = V2 \ cA.UnidCaja
        End If
    Else
        'Ha metido cajas. Nos vamos a cantidad
        If Text1(4).Text = "" Then
            Text1(5).Text = ""
        Else
        
            V = Val(Text1(4).Text)
            Text1(5).Text = Format(V * cA.UnidCaja, FormatoCantidad)
        End If
    End If
End Sub




Private Sub PonerEstado(Cual As Byte)
    Estado = Cual
    Me.Shape1.visible = Estado = 0
    Me.FrameLinea.Enabled = Estado = 0
    Me.Shape2.visible = Estado = 1
    Me.FrameStcok.Enabled = Estado = 1
    
    
    Me.imgLotes(0).visible = Estado = 1
    Me.imgLotes(1).visible = Estado = 1
    Me.imgLotes(2).visible = Estado = 1
    
    If Estado = 0 Then
        Me.cmdAceptar.Caption = "&Siguiente"
        Me.cmdCancelar.Caption = "Cancelar"
        Me.Label6.Caption = "1/2"
        
        
        Me.Command1.visible = Me.Numlinea < 0
        
        
    Else
        Me.cmdCancelar.Caption = "Atrás"
        If Me.Numlinea < 0 Then
            Me.cmdAceptar.Caption = "&Insertar"
        Else
            Me.cmdAceptar.Caption = "&Modificar"
        End If
        Label6.Caption = "2/2"
    End If
    FrameEliminar.visible = Estado = 0 And Me.cmdEliminar.visible
    
End Sub



Private Function ComprobarDatosStocks(ByRef Diferencia As Currency) As Boolean
Dim C As Currency
Dim I As Integer
Dim ModificaCantidad  As Boolean

    ComprobarDatosStocks = False

    'Veremos si ha modificado cantidad
    ModificaCantidad = False
    If Me.Numlinea < 0 Then
        Text1(5).Tag = 0
        vCStock.codalmac = 1
        vCStock.codArtic = cA.Codigo
        vCStock.Cantidad = ImporteFormateado(Text1(5).Text)
        vCStock.Importe = ImporteFormateado(Text1(9).Text)
        vCStock.DetaMov = Codtipoa
        vCStock.tipoMov = "S" 'salida
        'vCStock.LineaDocu = Numlinea  numlinea hay que calcularo
        ' fecha la obtendra de un select
        vCStock.Documento = Format(NumAlbar, "0000000")
        vCStock.HoraMov = Now
        vCStock.Trabajador = Val(Mid(Me.Caption, 1, 10))  'cliente
        Set RS = New ADODB.Recordset
        cad = "Select fechaalb from scafac1 "
        cad = cad & PonWhere
        RS.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        'no puede ser eOF
        If RS.EOF Then Err.Raise 513, "No se ha enontrado el albaran" & vbCrLf & cad
        vCStock.Fechamov = RS!FechaAlb
        RS.Close
        
        cad = "Select max(numlinea) from slifac"
        cad = cad & PonWhere
        RS.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        vCStock.LineaDocu = 0
        If Not RS.EOF Then vCStock.LineaDocu = DBLet(RS.Fields(0), "N")
        RS.Close
        vCStock.LineaDocu = vCStock.LineaDocu + 1
        Set RS = Nothing
    End If
    
    
    If ImporteFormateado(Text1(5).Text) <> CCur(Text1(5).Tag) Then
        'ziiiiiiiiiiiiiiii, ha cambiado cantidad
        ModificaCantidad = True
    
    Else
        If ListView2.Tag <> 0 Then ModificaCantidad = True
    End If

    cad = ""
    If ModificaCantidad Then
        'Veremos si hay stock suficiente
        Diferencia = ImporteFormateado(Text1(5).Text) - CCur(Text1(5).Tag)
        
        C = CCur(Text1(11).Tag) - Diferencia
        If C < 0 Then cad = cad & vbCrLf & "No hay cantidad suficiente en stock"
        
        
        
        'si modifica cantidad y NO modica el LOTE, o no suma bastante
        If cA.Trazabilidad Then
            If Val(Me.ListView2.Tag) = 0 Then
                cad = cad & vbCrLf & "NO ha ajustado los lotes"
                
                
            Else
                'Si ha cambiado veo cuanto suma
                C = 0
                For I = 1 To Me.ListView2.ListItems.Count
                    C = C + ImporteFormateado(ListView2.ListItems(I).SubItems(2))
                Next
                
                If ImporteFormateado(Text1(5).Text) <> C Then cad = cad & vbCrLf & "Suma lotes no coincide con el total"
                
            End If
        End If
        
    End If
        
        
        
        
        
    If cad = "" Then
        If Me.Numlinea >= 0 Then
            cad = "Modificar la linea de factura?"
        Else
            cad = "Insertar la linea de factura?"
        End If
    Else
        cad = cad & vbCrLf & vbCrLf & "¿Continuar?"
    End If

    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then ComprobarDatosStocks = True


End Function



Private Function RealizarUpdate(Diferencia As Currency) As Boolean
    RealizarUpdate = False

    Screen.MousePointer = vbHourglass
    lblIndicador.Caption = ""
        
    Conn.BeginTrans
    Set RS = New ADODB.Recordset
    If Actualizar(Diferencia) Then
        Conn.CommitTrans
        RealizarUpdate = True
    Else
        Conn.RollbackTrans
    End If
    Set RS = Nothing
    lblIndicador.Caption = ""
    Screen.MousePointer = vbDefault

End Function



Private Function Actualizar(Diferencia As Currency) As Boolean

    If Numlinea < 0 Then
        Actualizar = ActualizarInsertar(Diferencia)
    Else
        Actualizar = ActualizarModificar(Diferencia)
    End If
End Function

Private Function ActualizarModificar(Diferencia As Currency) As Boolean
Dim SQL As String
Dim Aux As Currency
On Error GoTo EActualizar
    ActualizarModificar = False

    'Vamos con el stock.
    'Updatearemos directamente
    lblIndicador.Caption = "salmac"
    lblIndicador.Refresh
    
    If Diferencia <> 0 Then
        'Diferencia es de signo contrario
        'ya que hemos restado lo que pone en el txtbox menos lo que habia
        Diferencia = Diferencia * -1
        
        SQL = "UPDATE salmac set canstock = canstock "
        If Diferencia > 0 Then SQL = SQL & " + "
        SQL = SQL & DBSet(Diferencia, "N")
        SQL = SQL & " WHERE codartic = " & DBSet(vCStock.codArtic, "T")
        SQL = SQL & " AND codalmac = " & DBSet(vCStock.codalmac, "N")
        Conn.Execute SQL
    End If
    
    
    
    'ACtualizo la smoval
    lblIndicador.Caption = "smoval"
    lblIndicador.Refresh
    
    'en la smovla el numlote VA formateado, en smoval lotes NO
    
    
        
    
    SQL = "UPDATE smoval set impormov = " & DBSet(Text1(9).Text, "N")
    If Diferencia <> 0 Then SQL = SQL & ", cantidad = " & DBSet(Text1(5).Text, "N")
    SQL = SQL & " where codartic = " & DBSet(vCStock.codArtic, "T") & "  and codalmac  = " & DBSet(vCStock.codalmac, "N") & "  and fechamov  =  " & DBSet(vCStock.Fechamov, "F")
    SQL = SQL & " and document  = " & DBSet(vCStock.Documento, "T") & " and numlinea  =  " & DBSet(Numlinea, "N") & " and detamovi  = " & DBSet(vCStock.DetaMov, "T")
    Conn.Execute SQL
        
    lblIndicador.Caption = "slifac"
    lblIndicador.Refresh
    SQL = "UPDATE slifac set ampliaci = " & DBSet(Text1(2).Text, "T")
    SQL = SQL & ",palets = " & DBSet(Text1(3).Text, "N")
    SQL = SQL & ",cantidad = " & DBSet(Text1(5).Text, "N")
    SQL = SQL & ",precioar = " & DBSet(Text1(6).Text, "N")
    SQL = SQL & ",dtoline1 = " & DBSet(Text1(7).Text, "N")
    SQL = SQL & ",dtoline2 = " & DBSet(Text1(8).Text, "N")
    SQL = SQL & ",importel = " & DBSet(Text1(9).Text, "N")
    SQL = SQL & ",origpre  = " & DBSet(Text1(10).Text, "T")
    If vParamAplic.QUE_EMPRESA = 2 Then
        Aux = 1
        If Not Text1(12).Locked Then
            Aux = ImporteFormateado(Text1(12).Text)
            Aux = Round2(Aux / 100, 4)
        End If
        SQL = SQL & ",hectogrado  = " & DBSet(Aux, "N")
        
    End If
    SQL = SQL & PonWhere
    SQL = SQL & " AND numlinea =" & Numlinea
    Conn.Execute SQL
    
    
    'Vale, ya tenemos stocks y slifac
    'Vamos con lotes etc
    If Val(Me.ListView2.Tag) = 0 Then
            'No han tocado los lotes. Salimos diciendo que esta bien
            ActualizarModificar = True
            Exit Function
    End If
    
    
    'Igual, podriamos poner un commitrans y añadir un begina trans
    'quiero decir que podriamos actualizar lo de arriba a la BD y el siguiente proceso si falla.....
    'ya veriamos
        'Ha modificado lotes. Vamos p'alla
    '1º Eliminamos los que habian
    '2º Insertamos los nuevos
    
    EliminarLotes
    
    'Insertamos los nuevos
    NuevosLotes
    
    'llEga aqui... tutto benne
    ActualizarModificar = True
    Exit Function
EActualizar:
    MuestraError Err.Number, Err.Description
End Function



Private Function ActualizarInsertar(Diferencia As Currency) As Boolean
Dim SQL As String
Dim Aux As Currency

On Error GoTo EActualizar2
    ActualizarInsertar = False

    'Vamos con el stock.
    'Updatearemos directamente
    lblIndicador.Caption = "Sctock"
    lblIndicador.Refresh
    
    vCStock.ActualizarStock False
        
    
    
        
    lblIndicador.Caption = "slifac"
    lblIndicador.Refresh
    
    'codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,
    'cantidad,precioar,dtoline1,dtoline2,importel,origpre,precioiv,preciomp,preciost,preciouc,codproveX,palets
    
    SQL = "INSERT INTO slifac(codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,"
    SQL = SQL & "cantidad,precioar,dtoline1,dtoline2,importel,origpre,precioiv,preciomp,preciost,preciouc,codproveX,palets"
    
    If vParamAplic.QUE_EMPRESA = 2 Then SQL = SQL & ",hectogrado "
        
    SQL = SQL & ") VALUES (" & DBSet(Codtipom, "T") & "," & DBSet(NumFactu, "N") & "," & DBSet(Fecfactu, "F") & ","
    SQL = SQL & DBSet(Codtipoa, "T") & "," & DBSet(NumAlbar, "N") & "," & DBSet(vCStock.LineaDocu, "N") & ","
    SQL = SQL & DBSet(vCStock.codalmac, "N") & "," & DBSet(cA.Codigo, "T") & "," & DBSet(Text1(1).Text, "T") & ","
    
    'ampliaci cantidad,precioar
    SQL = SQL & DBSet(Text1(2).Text, "T") & "," & DBSet(Text1(5).Text, "N") & "," & DBSet(Text1(6).Text, "N") & ","
    
    'dtoline1,dtoline2,importel,origpre
    SQL = SQL & DBSet(Text1(7).Text, "N") & "," & DBSet(Text1(8).Text, "N") & "," & DBSet(Text1(9).Text, "N") & "," & DBSet(Text1(10).Text, "T") & ","
    
    'precioiv,preciomp,preciost,preciouc
    SQL = SQL & DBSet(cA.PrecioVenta, "N") & "," & DBSet(cA.PrecioMedPon, "N") & "," & DBSet(cA.PrecioStan, "N") & "," & DBSet(cA.PrecioUltCom, "N") & ","
    
    'codproveX,palets
    SQL = SQL & DBSet(cA.codProve, "N") & "," & DBSet(Text1(3).Text, "N")


    If vParamAplic.QUE_EMPRESA = 2 Then
        
        Aux = 1
        If Not Text1(12).Locked Then
            Aux = ImporteFormateado(Text1(12).Text)
            Aux = Round2(Aux / 100, 2)
        End If
        SQL = SQL & "," & DBSet(Aux, "N")

    End If

    SQL = SQL & ")"
    Conn.Execute SQL
    
    
    
    'Vale, ya tenemos stocks y slifac
    'Vamos con lotes etc
    If Val(Me.ListView2.Tag) = 0 Then
            'No han tocado los lotes. Salimos diciendo que esta bien
            ActualizarInsertar = True
            Exit Function
    End If
   
    'Insertamos los nuevos
    NuevosLotes
    
    'llEga aqui... tutto benne
    ActualizarInsertar = True
    Exit Function
EActualizar2:
    MuestraError Err.Number, Err.Description
End Function




Private Sub EliminarLotes()
Dim Cp As cPartidas
Dim cL As cLotaje


    cad = "Select linea,numlote,cantidad,numlinea from slifaclotes "
    cad = cad & PonWhere
    cad = cad & " AND numlinea =" & Numlinea & " order by LINEA"
    RS.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Set Cp = New cPartidas
    Set cL = New cLotaje
    
    cL.DetaMov = Codtipoa
    cL.Fechamov = Fecfactu
    cL.Documento = NumAlbar
    cL.tipoMov = "0"
    
    
    
    While Not RS.EOF
        If Cp.LeerDesdeArticulo(vCStock.codArtic, vCStock.codalmac, RS!Numlote) Then
                
                Cp.IncrementarCantidad RS!Cantidad 'devuelvo la cantidad
            
                cL.codalmac = vCStock.codalmac
                cL.codArtic = vCStock.codArtic
                cL.LineaDocu = RS!Numlinea
                cL.SubLinea = RS!linea
                cL.Numlote = RS!Numlote
                If cL.Leer Then cL.EliminarMovimArticulosLotaje False
        
        
        Else
            'error
            Err.Raise 513, "Leyendo partida: " & RS!Numlote
        End If
        
        
        
        RS.MoveNext
    Wend
    RS.Close
    
    'Si llega aqui, borro slifaclotes
    'Elimino la linea de slifaclotes
    cad = "DELETE from slifaclotes"
    cad = cad & PonWhere
    cad = cad & " AND numlinea =" & Numlinea
    
    Conn.Execute cad
    
    Set Cp = Nothing
    Set cL = Nothing
    
End Sub




Private Sub NuevosLotes()
Dim Cp As cPartidas
Dim cL As cLotaje
Dim J As Integer


    Set Cp = New cPartidas
    Set cL = New cLotaje
    

    For J = 1 To ListView2.ListItems.Count
        lblIndicador.Caption = "Inserta slifaclot: " & J
        lblIndicador.Refresh
        
        cad = "INSERT INTO slifaclotes(codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,"
        cad = cad & "linea,numlote,cantidad) VALUES (" & DBSet(Codtipom, "T") & ","
        cad = cad & DBSet(NumFactu, "N") & "," & DBSet(Fecfactu, "F") & "," & DBSet(Codtipoa, "T") & ","
        cad = cad & DBSet(NumAlbar, "N") & ","
        If Me.Numlinea < 0 Then
             cL.LineaDocu = vCStock.LineaDocu
        Else
             cL.LineaDocu = Numlinea
        End If
        
        cad = cad & cL.LineaDocu & "," & DBSet(ListView2.ListItems(J).Text, "N") & "," & DBSet(ListView2.ListItems(J).SubItems(1), "T")
        cad = cad & "," & DBSet(ListView2.ListItems(J).SubItems(2), "N") & ")"
        Conn.Execute cad
        
        cL.DetaMov = Codtipoa
        cL.Fechamov = Fecfactu
        cL.Documento = NumAlbar
                    
        cL.codalmac = vCStock.codalmac
        cL.codArtic = vCStock.codArtic
       
        cL.ProvCliTra = Val(Mid(Me.Caption, 1, 10))
    
    
        lblIndicador.Caption = "Leyendo partida: " & J
        lblIndicador.Refresh
        
        If Cp.LeerDesdeArticulo(vCStock.codArtic, vCStock.codalmac, ListView2.ListItems(J).SubItems(1)) Then
                
                Cp.IncrementarCantidad -1 * ImporteFormateado(ListView2.ListItems(J).SubItems(2))
                
                cL.HoraMov = Now
                cL.Numlote = ListView2.ListItems(J).SubItems(1)
                cL.SubLinea = ListView2.ListItems(J).Text
                cL.Cantidad = CSng(ImporteFormateado(ListView2.ListItems(J).SubItems(2)))
                cL.InsertarLote
        
        
        Else
            
            'error
            Err.Raise 513, "Leyendo partida: " & RS!Numlote
        End If
        
    Next
    
    
    Set Cp = Nothing
    Set cL = Nothing
    
End Sub


Private Function Eliminar() As Boolean
Dim Dif As Currency

On Error GoTo EEliminar
    Eliminar = False
    
    Set RS = New ADODB.Recordset
    EliminarLotes
    Set RS = Nothing
    
    Dif = CCur(Text1(5).Tag)
    
    If CCur(Text1(5).Tag) <> 0 Then
        cad = "UPDATE salmac set canstock = canstock "
        If Dif > 0 Then cad = cad & " + "
        cad = cad & DBSet(Dif, "N")
        cad = cad & " WHERE codartic = " & DBSet(vCStock.codArtic, "T")
        cad = cad & " AND codalmac = " & DBSet(vCStock.codalmac, "N")
        Conn.Execute cad
    End If
    
    'smoval
    cad = "DELETE from smoval where codartic = " & DBSet(vCStock.codArtic, "T") & "  and codalmac  = " & DBSet(vCStock.codalmac, "N") & "  and fechamov  =  " & DBSet(vCStock.Fechamov, "F")
    cad = cad & " and document  = " & DBSet(vCStock.Documento, "T") & " and numlinea  =  " & DBSet(Numlinea, "N") & " and detamovi  = "
    'Detamovi: SERA SIEMPRE 0 'salida
    cad = cad & "0"
    Conn.Execute cad
    
        
    
    
    'Borro la linea
    cad = "DELETE from slifac"
    cad = cad & PonWhere
    cad = cad & " AND numlinea =" & Numlinea
    Conn.Execute cad
    
    
    Eliminar = True
    
    Exit Function
EEliminar:
    MuestraError Err.Number
        
End Function

Private Sub Text2_Change()

End Sub
