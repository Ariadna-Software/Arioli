VERSION 5.00
Begin VB.Form frmFacTrazabilidad3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listado"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   8580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameDesdeAlbCompra 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      Begin VB.CheckBox chkCoupage 
         Caption         =   "No mostrar ventas coupages"
         Height          =   255
         Left            =   480
         TabIndex        =   34
         Top             =   1680
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CommandButton cmdCompra 
         Caption         =   "Desde alb compra"
         Height          =   495
         Left            =   4560
         TabIndex        =   11
         Top             =   1800
         Width           =   1815
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   495
         Index           =   0
         Left            =   6600
         TabIndex        =   10
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   320
         Index           =   0
         Left            =   480
         TabIndex        =   4
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   320
         Index           =   0
         Left            =   1560
         TabIndex        =   3
         Text            =   "Text2"
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   320
         Index           =   1
         Left            =   3240
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   1200
         Width           =   3015
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   320
         Index           =   2
         Left            =   6240
         TabIndex        =   1
         Text            =   "Text2"
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label lbIndicador1 
         Caption         =   "Label2"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   2040
         Width           =   4095
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Lote desde albaran compra"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   6495
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   0
         Left            =   120
         Picture         =   "frmFacTrazabilidad3.frx":0000
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label label1 
         Caption         =   "Id"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   8
         Top             =   840
         Width           =   375
      End
      Begin VB.Label label1 
         Caption         =   "Cod. árticulo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   270
         Index           =   0
         Left            =   1560
         TabIndex        =   7
         Top             =   840
         Width           =   1425
      End
      Begin VB.Label label1 
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   6
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label label1 
         Caption         =   "Lote"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   3
         Left            =   6240
         TabIndex        =   5
         Top             =   840
         Width           =   1935
      End
   End
   Begin VB.Frame FrameDesdeVenta 
      Height          =   2535
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   8415
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   320
         Index           =   5
         Left            =   6240
         TabIndex        =   29
         Text            =   "Text2"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   320
         Index           =   4
         Left            =   3240
         TabIndex        =   28
         Text            =   "Text2"
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   320
         Index           =   3
         Left            =   1560
         TabIndex        =   27
         Text            =   "Text2"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   320
         Index           =   1
         Left            =   480
         TabIndex        =   26
         Top             =   1080
         Width           =   975
      End
      Begin VB.CheckBox chkMatePrima 
         Caption         =   "Solo materia prima"
         Height          =   255
         Left            =   480
         TabIndex        =   24
         Top             =   1800
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.ComboBox cboLotes 
         Height          =   315
         Left            =   5400
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   4080
         Width           =   2535
      End
      Begin VB.TextBox txtArticulo 
         Height          =   285
         Index           =   0
         Left            =   120
         MaxLength       =   16
         TabIndex        =   21
         Top             =   4080
         Width           =   1695
      End
      Begin VB.TextBox txtDescArticulo 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "Text5"
         Top             =   4080
         Width           =   3495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Factura"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   19
         Top             =   4560
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Albarán"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   4560
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   495
         Index           =   1
         Left            =   6600
         TabIndex        =   14
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Desde venta"
         Height          =   495
         Left            =   4560
         TabIndex        =   13
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label label1 
         Caption         =   "Lote"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   10
         Left            =   6240
         TabIndex        =   33
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label label1 
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   9
         Left            =   3240
         TabIndex        =   32
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label label1 
         Caption         =   "Cod. árticulo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   270
         Index           =   8
         Left            =   1560
         TabIndex        =   31
         Top             =   720
         Width           =   1425
      End
      Begin VB.Label label1 
         Caption         =   "Id"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   7
         Left            =   480
         TabIndex        =   30
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   120
         Picture         =   "frmFacTrazabilidad3.frx":0A02
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label label1 
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   4
         Left            =   5400
         TabIndex        =   22
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   0
         Left            =   1560
         Picture         =   "frmFacTrazabilidad3.frx":1404
         Top             =   3720
         Width           =   240
      End
      Begin VB.Label label1 
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   6
         Left            =   1920
         TabIndex        =   17
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Label label1 
         Caption         =   "Cod. árticulo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   16
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Lote en venta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   6495
      End
   End
End
Attribute VB_Name = "frmFacTrazabilidad3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Public Opcion As Integer
    '0- Desde compra
    '1- desde venta
 
Private WithEvents frmPa As frmAlmPartidas
Attribute frmPa.VB_VarHelpID = -1
Private WithEvents frmA As frmAlmArticulos
Attribute frmA.VB_VarHelpID = -1
Dim SQL As String


Private Sub cmdCompra_Click()
    Dim cP As cPartidas
    
    If Text1(0).Text = "" Then Exit Sub
    If Not IsNumeric(Text1(0).Text) Then Exit Sub
    
    
     
    
    
    Set cP = New cPartidas
    If cP.Leer(Val(Text1(0).Text)) Then
        Screen.MousePointer = vbHourglass
        conn.Execute "DELETE FROM tmptraza"
        cP.TrazbilidadDesdeCompra Me.lbIndicador1, chkCoupage.Value = 0
        lbIndicador1.Caption = ""
        Screen.MousePointer = vbDefault
        'QUITARA
       '
        'cP.TrazabilidadDesdeCompra
        
            With frmImprimir
                .FormulaSeleccion = "{tmptraza.codusu} = " & vUsu.Codigo
                .OtrosParametros = "|pNomEmpre=""" & vEmpresa.nomempre & """|"
                .NumeroParametros = 1
        
                .SoloImprimir = False
                .EnvioEMail = False
                .Opcion = 2002
                .Titulo = "Trazabilidad(III)"
                .NombreRPT = "TrazaNueva.rpt"
                .ConSubInforme = True
                .Show vbModal
            End With
    Else
        MsgBox "No existe idTraza: " & Text1(0).Text, vbExclamation
    End If
    Set cP = Nothing
End Sub



Private Sub cmdSalir_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Command1_Click()
Dim cP As cPartidas

'    If Me.txtArticulo(0).Text = "" Or Me.txtDescArticulo(0).Text = "" Or Me.cboLotes.ListIndex < 0 Then
'        MsgBox "Ponga los datos de consulta", vbExclamation
'        Exit Sub
'    End If
'
    If Text1(1).Text = "" Or Me.Text2(3).Text = "" Then
        MsgBox "Ponga los datos de consulta", vbExclamation
        Exit Sub
    End If
    
    Set cP = New cPartidas
    If cP.Leer(CLng(Text1(1).Text)) Then
        conn.Execute "DELETE FROM tmptraza"
        cP.TrazbilidadDesdeVenta Me.chkMatePrima.Value = 1, True
        
        'QUITARA
       '
        'cP.TrazabilidadDesdeCompra
        
            With frmImprimir
                .FormulaSeleccion = "{tmptraza.codusu} = " & vUsu.Codigo
                .OtrosParametros = ""
                .NumeroParametros = 0
        
                .SoloImprimir = False
                .EnvioEMail = False
                .Opcion = 2002
                .Titulo = "Trazabilidad(IV)"
                .NombreRPT = "TrazaNuevaVenta.rpt"
                .ConSubInforme = True
                .Show vbModal
            End With
    Else
        MsgBox "No existe valores de trazbilidad", vbExclamation
    End If
    Set cP = Nothing
End Sub
    
    
    
    
    


Private Sub Form_Load()
Dim H As Integer
Dim W As Integer

    limpiar Me
    Me.Icon = frmppal.Icon
    Me.FrameDesdeAlbCompra.visible = False
    Me.FrameDesdeVenta.visible = False
    Select Case Opcion
    Case 0
        PonerFrameVisible FrameDesdeAlbCompra, H, W
        lbIndicador1.Caption = ""
    Case 1
        PonerFrameVisible FrameDesdeVenta, H, W
    End Select
    
    Me.Height = H
    Me.Width = W
    Me.cmdSalir(Opcion).Cancel = True
    
End Sub


Private Sub PonerFrameVisible(ByRef F As Frame, ByRef CH As Integer, CW As Integer)
    F.Top = 0
    F.Left = 120
    F.visible = True
    CH = F.Height + 480
    CW = F.Width + 240
End Sub


Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
    SQL = CadenaSeleccion
End Sub

Private Sub frmPa_DatoSeleccionado(CadenaSeleccion As String)
    SQL = RecuperaValor(CadenaSeleccion, 1)
End Sub

Private Sub Image1_Click(Index As Integer)
    Set frmPa = New frmAlmPartidas
    frmPa.DatosADevolverBusqueda = "*"
    SQL = ""
    frmPa.Show vbModal
    Set frmPa = Nothing
    If SQL <> "" Then
        Text1(Index).Text = SQL
        Text1_LostFocus Index
    End If
End Sub


Private Sub Limpia2()
    Text2(0).Text = "": Text2(2).Text = "": Text2(2).Text = ""
    Text2(3).Text = "": Text2(4).Text = "": Text2(5).Text = ""
End Sub



Private Sub Image2_Click(Index As Integer)
    SQL = ""
    Set frmA = New frmAlmArticulos
    frmA.DatosADevolverBusqueda2 = "@1@" 'Abrimos en Modo Busqueda
    frmA.Show vbModal
    Set frmA = Nothing
    
    If SQL <> "" Then
        If SQL <> Me.txtArticulo(Index).Text Then
            txtArticulo(Index).Text = RecuperaValor(SQL, 1)
            Me.txtDescArticulo(Index).Text = RecuperaValor(SQL, 2)
            CargaLotes
        End If
    End If
End Sub

Private Sub Option1_Click(Index As Integer)
    CargaLotes2
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim cP As cPartidas
Dim FactorConver As String
    If Index = 0 Then
        Me.chkCoupage.Value = 0
        Me.chkCoupage.visible = False
    End If
    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).Tag = Text1(Index).Text Then Exit Sub
    
    If Text1(Index).Text = "" Then
        Limpia2
    Else
        If Not IsNumeric(Text1(Index).Text) Then
            MsgBox "Campo numérico", vbExclamation
            Text1(Index).Text = ""
            Limpia2
        Else
            Set cP = New cPartidas
            If cP.Leer(Val(Text1(Index).Text)) Then
                'OK
                If Index = 0 Then
                    Text2(0).Text = cP.codartic
                    FactorConver = "factorconversion"
                    Text2(1).Text = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", cP.codartic, "T", FactorConver)
                    Text2(2).Text = cP.NUmlote
                    
                    
                    If FactorConver <> "1" Then
                        Me.chkCoupage.Value = 1
                        Me.chkCoupage.visible = True
                    End If
                Else
                    Text2(3).Text = cP.codartic
                    Text2(4).Text = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", cP.codartic, "T")
                    Text2(5).Text = cP.NUmlote
                End If
            Else
                MsgBox "No existe la partida ID= " & Text1(Index).Text, vbExclamation
                Limpia2
                Text1(Index).Text = ""
            End If
            
            
        End If
    End If
    Text1(Index).Tag = Text1(Index).Text
End Sub

Private Sub txtArticulo_GotFocus(Index As Integer)
    ConseguirFoco txtArticulo(Index), 3
End Sub

Private Sub txtArticulo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
    
End Sub


Private Sub txtArticulo_LostFocus(Index As Integer)
Dim T As String

    txtArticulo(Index).Text = Trim(txtArticulo(Index).Text)
    If txtArticulo(Index).Text = "" Then
        'EN blanco
        txtDescArticulo(Index).Text = ""
        Exit Sub
    End If
    
    
    T = "codartic"
    SQL = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", txtArticulo(Index).Text, "T", T)
    If SQL = "" Then
        MsgBox "No existe el artículo : " & txtArticulo(Index).Text, vbExclamation
    Else
        txtArticulo(Index).Text = T
    End If
    Me.txtDescArticulo(Index).Text = SQL
    SQL = ""
    CargaLotes
End Sub

Private Sub CargaLotes()
    Screen.MousePointer = vbHourglass
    CargaLotes2
    Screen.MousePointer = vbDefault
End Sub

Private Sub CargaLotes2()
    Me.cboLotes.Clear
    
    If Me.txtArticulo(0).Text = "" Or Me.txtDescArticulo(0).Text = "" Then Exit Sub
    
    
    'SQL = "select c.numalbar,c.fechaalb,nomclien,l.numlinea,linea,numlote,nomclien,codartic,c.codtipom"
    If Me.Option1(0).Value Then
        SQL = "select distinct(numlote)"
        SQL = SQL & " from scaalb c,slialb l,slialblotes t where"
        SQL = SQL & " c.codtipom=l.codtipom and c.numalbar=l.numalbar and"
        SQL = SQL & " C.codTipoM = T.codTipoM And C.NumAlbar = T.NumAlbar And L.numlinea = T.numlinea"
        SQL = SQL & " AND l.codartic= '" & txtArticulo(0).Text & "'"
        SQL = SQL & " ORDER BY numlote"
    Else
        SQL = "select  distinct(numlote) "
        SQL = SQL & " from scafac c,scafac1 c2,slifac l, slifaclotes t where"
        SQL = SQL & " c.codtipom=c2.codtipom and c.numfactu=c2.numfactu and c.fecfactu=c2.fecfactu and"
        SQL = SQL & " l.codtipom=c2.codtipom and l.numfactu=c2.numfactu and l.fecfactu=c2.fecfactu and l.codtipoa=c2.codtipoa and l.numalbar=c2.numalbar and"
        SQL = SQL & " L.codTipoM = T.codTipoM And L.NumFactu = T.NumFactu And L.FecFactu = T.FecFactu And L.codtipoa = T.codtipoa And L.NumAlbar = T.NumAlbar And L.numlinea = T.numlinea"
        SQL = SQL & " AND l.codartic= '" & txtArticulo(0).Text & "'"
        SQL = SQL & " ORDER BY numlote"
        
    End If
    
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        'Insertamos el lote
        Me.cboLotes.AddItem CStr(miRsAux!NUmlote)
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If cboLotes.ListCount > 0 Then cboLotes.ListIndex = 0
End Sub
