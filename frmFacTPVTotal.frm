VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacTPVTotal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Total Venta"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   7620
   Icon            =   "frmFacTPVTotal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   5
      Left            =   2160
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2032
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   5
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "Text2"
      Top             =   2032
      Width           =   4245
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   2640
      MaxLength       =   15
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   2
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1180
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   2
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "Text2"
      Top             =   1180
      Width           =   4245
   End
   Begin VB.Frame FrameEfectivo 
      Height          =   1815
      Left            =   2160
      TabIndex        =   6
      Top             =   3360
      Width           =   5295
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   3
         Left            =   2160
         TabIndex        =   5
         Text            =   "0.0"
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label label2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   405
         Index           =   1
         Left            =   2160
         TabIndex        =   15
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label label2 
         Appearance      =   0  'Flat
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   405
         Index           =   0
         Left            =   360
         TabIndex        =   14
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label label2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   4
         Left            =   2160
         TabIndex        =   13
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label label2 
         Caption         =   "CAMBIO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   3
         Left            =   360
         TabIndex        =   12
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label label2 
         Caption         =   "ENTREGADO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   2
         Left            =   360
         TabIndex        =   11
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   1
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   2460
      Width           =   3885
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   2640
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2460
      Width           =   735
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   0
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   1606
      Width           =   4245
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1606
      Width           =   855
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   1111
      ButtonWidth     =   2249
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ticket  F5"
            Object.ToolTipText     =   "Generar Ticket"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Albaran  Ctr+F6"
            Object.ToolTipText     =   "Generar Albaran"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Factura  F7"
            Object.ToolTipText     =   "Generar Factura"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salir"
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
   End
   Begin VB.Label label1 
      Caption         =   "Dpto."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   22
      Top             =   2032
      Width           =   975
   End
   Begin VB.Image imgBuscar 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   5
      Left            =   1680
      ToolTipText     =   "Buscar direc./dpto"
      Top             =   2032
      Width           =   360
   End
   Begin VB.Label label1 
      Caption         =   "Cheque regalo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   3
      Left            =   240
      TabIndex        =   20
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Image imgBuscar 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   2
      Left            =   1680
      ToolTipText     =   "Buscar artículo"
      Top             =   1180
      Width           =   360
   End
   Begin VB.Label label1 
      Caption         =   "Operador "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   18
      Top             =   1180
      Width           =   1215
   End
   Begin VB.Image imgBuscar 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   0
      Left            =   1680
      ToolTipText     =   "Buscar cliente"
      Top             =   1606
      Width           =   360
   End
   Begin VB.Label LabelB 
      Caption         =   "F2 = Buscar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1680
      TabIndex        =   16
      Top             =   840
      Width           =   1320
   End
   Begin VB.Image imgBuscar 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   1
      Left            =   2160
      ToolTipText     =   "Buscar forma de pago"
      Top             =   2460
      Width           =   360
   End
   Begin VB.Label label1 
      Caption         =   "Forma de pago"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   2460
      Width           =   1815
   End
   Begin VB.Label label1 
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   1606
      Width           =   975
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnTicket 
         Caption         =   "&Ticket"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnAlbaran 
         Caption         =   "&Albaran"
         Shortcut        =   ^{F6}
      End
      Begin VB.Menu mnFactura 
         Caption         =   "&Factura"
         Shortcut        =   {F7}
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
Attribute VB_Name = "frmFacTPVTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public cadSel As String 'cadena para seleccion de la venta a totalizar
Public Importe As String
'Public NumTermi As Integer

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1

Private PrimeraVez As Boolean
Private HaDevueltoDatos As Boolean


Dim CodTipoMov As String
'Codigo tipo de movimiento en función del valor en la tabla de parámetros: stipom

Dim TipoForPa As Byte 'tipo forma de pago: efectivo, banco,...
Dim codAlmac As Integer 'cod. almacen
Dim NomTraba As String 'nombre trabajador

Dim RSVenta As ADODB.Recordset


Dim SQL As String
Dim cadImpresion As String

'--- Variables generales para nueva impresión ticket (RAFA/ALZIRA 05092006)
Dim vNumTicket As String
Dim vNumAlbTicket As String
Dim vFechaTicket As Date


Private Sub Form_Activate()
    If PrimeraVez Then
        If vParamTPV.Rapida Then
            If vParamAplic.ForPagoChequeRegalo = "" Then
                PonerFoco Text1(3)
            Else
                PonerFoco Text1(4)
            End If
        Else
            PonerFoco Text1(1)
        End If
        PrimeraVez = False
    End If
End Sub

Private Sub Form_Load()
Dim cad As String


'    If cadSel = "" Then Unload Me


     'Icono del formulario
    Me.Icon = frmppal.Icon
    
    'Icono de busqueda
    Me.imgBuscar(0).Picture = frmppal.ImgListPpal.ListImages(17).Picture
    Me.imgBuscar(1).Picture = frmppal.ImgListPpal.ListImages(17).Picture
    Me.imgBuscar(2).Picture = frmppal.ImgListPpal.ListImages(17).Picture
    Me.imgBuscar(5).Picture = frmppal.ImgListPpal.ListImages(17).Picture
    
    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmppal.ImgListPpal
        .Buttons(2).Image = 18   'Generar Ticket
        .Buttons(4).Image = 7   'Generar Albaran
        .Buttons(6).Image = 8   'Generar Factura

        .Buttons(10).Image = 14  'Salir
    End With
    
    
    PrimeraVez = True
    CodTipoMov = "FTI" 'factura ticket
    
    SQL = "SELECT * FROM scaven "
    If cadSel <> "" Then SQL = SQL & " WHERE " & cadSel
    Set RSVenta = New ADODB.Recordset
    RSVenta.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    
    'Almacen por defecto el del trabajador
    If RSVenta!CodTraba <> "" Then
        NomTraba = "nomtraba"
        codAlmac = ComprobarCero(DevuelveDesdeBDNew(conAri, "straba", "codalmac", "codtraba", CStr(RSVenta!CodTraba), "N", NomTraba))
        If codAlmac = 0 Then codAlmac = DevuelveDesdeBDNew(conAri, "salmpr", "min(codalmac)", "", "")
    Else
        codAlmac = DevuelveDesdeBDNew(conAri, "salmpr", "min(codalmac)", "", "")
    End If
            
            
    Me.Label2(1).Caption = Importe
    
    If vParamTPV.Rapida Then
        Me.Label2(4).Caption = "0.00" 'cambio
        Me.Text1(3).Text = "0.00" 'Entregado
    Else
        Me.Label2(4).Caption = "" 'cambio
        Me.Text1(3).Text = "" 'Entregado
    End If
    Me.Text1(4).Text = "" 'Cheque regalo
    
    'Trabajador conectado
    Text1(2).Text = Format(RSVenta!CodTraba, "0000")
    Text2(2).Text = NomTraba
    
    'Poner el cliente de la venta
    If Not IsNull(RSVenta!CodClien) Then
        cad = "codforpa"
        Text1(0).Text = Format(RSVenta!CodClien, "000000")
        Text2(0).Text = DevuelveDesdeBDNew(conAri, "sclien", "nomclien", "codclien", Text1(0).Text, "N", cad) '(RAFA/ALZIRA 31082006)
        ' (RAFA/ALZIRA 31082006) -- Inicio
        Text1(1).Text = Format(CLng(cad), "000")
        cad = "tipforpa"
        Text2(1).Text = DevuelveDesdeBDNew(conAri, "sforpa", "nomforpa", "codforpa", Text1(1).Text, "N", cad)
'        TipoForPa = CByte(Text1(1).Text)
        TipoForPa = CByte(cad)
        ' (RAFA/ALZIRA 31082006) -- Fin
        
        'departamento del cliente
        If Not IsNull(RSVenta!CodDirec) Then
            Text1(5).Text = Format(DBLet(RSVenta!CodDirec, "N"), "000")
            PonerDptoEnCliente
        Else
            Text1(5).Text = ""
            Text2(5).Text = ""
        End If
        
    Else
        'Poner el cliente que hay por defecto en los parametros
        Text1(0).Text = Format(vParamTPV.Cliente, "000000")
        Text2(0).Text = vParamTPV.NomCliente
        'Forma de pago por defecto (RAFA/ALZIRA 31082006)
        Text1(1).Text = Format(vParamTPV.ForPago, "000")
        Text2(1).Text = vParamTPV.NomForPago
        TipoForPa = vParamTPV.TipoForPago
    End If
    
    'Forma de pago por defecto (RAFA/ALZIRA 31082006) Ahora esto se hace en el IF superior
'    Text1(1).Text = Format(vParamTPV.ForPago, "000")
'    Text2(1).Text = vParamTPV.NomForPago
'    TipoForPa = vParamTPV.TipoForPago
    
    Me.FrameEfectivo.Enabled = (TipoForPa = 0)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    RSVenta.Close
    Set RSVenta = Nothing
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
'para busquedas
Dim I As Byte

    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        I = CInt(Me.imgBuscar(0).Tag)
        
        Text1(I).Text = RecuperaValor(CadenaDevuelta, 1)
'        If i <> 5 Then Text1_LostFocus (i)
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgBuscar_Click(Index As Integer)
    If RSVenta.EOF Then Exit Sub
    imgBuscar(0).Tag = Index
    MandaBusquedaPrevia CStr(Index)
'    If Index = 5 Then Text1_LostFocus (5)
End Sub





Private Sub mnAlbaran_Click()
'Pasamos la venta a una albaran de venta generado a partir de un ticket
'en los campos del pedido almacenamos de que ticket viene
Dim NumAlbaran As String

    '---- comprobar datos correctos
    If Not DatosOk Then Exit Sub
    
    If Me.Text1(4).Text <> "" Then
        MsgBox "El cheque regalo no se puede utilizar en Albaranes", vbInformation
        Exit Sub
    End If
    
    '---- Generar el albaran y eliminar la venta
    CodTipoMov = "ALV" 'factura ticket
    If GenerarAlbaran(NumAlbaran) Then
        '---- Imprimir el Albaran
        NumAlbaran = "Se ha generado correctamente el Albaran de venta: " & NumAlbaran & vbCrLf
        If MsgBox(NumAlbaran & "¿Desea imprimirlo?  ", vbQuestion + vbYesNo) = vbYes Then ImprimirAlbaran
        
        'cerrar ventana total y regresar a entrada de ventas
        cadSel = "1"
'        Unload Me
        mnSalir_Click
    End If
End Sub


Private Sub mnFactura_Click()
Dim NumFactura As String
Dim bSalir As Boolean

    
    If Not DatosOk Then Exit Sub
    
    bSalir = False
    Me.Toolbar1.Buttons(6).Enabled = False
    Me.mnFactura.Enabled = False
    
    CodTipoMov = "FAV" 'factura ticket
    
    Screen.MousePointer = vbHourglass
    If GenerarFactura(NumFactura) Then
        Screen.MousePointer = vbDefault
        'Imprimir la factura
        NumFactura = "Se ha generado correctamente la Factura: " & NumFactura & vbCrLf
        If MsgBox(NumFactura & "¿Desea imprimirla?  ", vbQuestion + vbYesNo) = vbYes Then ImprimirFactura
        'cerrar ventana total y regresar a entrada de ventas
        cadSel = "1"
'        Unload Me
        bSalir = True
    End If
    Screen.MousePointer = vbDefault
    Me.Toolbar1.Buttons(6).Enabled = True
    Me.mnFactura.Enabled = True
    
    If bSalir Then mnSalir_Click
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub


Private Sub mnTicket_Click()
Dim Impr As Boolean
Dim curEntregado As Currency
Dim curCambio As Currency


    If Not DatosOk Then Exit Sub
    
    '## LAURA 20/06/2008
    '-- Comprobar que si existe un articulo con registro fitosanitario
    '-- no se puede hacer un ticket y salimos
    If HayArticuloFitosanitario Then Exit Sub
    
    '##
    
    cadImpresion = ""
    CodTipoMov = "FTI" 'factura ticket
    curEntregado = 0
    curCambio = 0
    
    'Si contabilizamos los tickets agrupados, entonces NO podra generar el ticket si
    ' el cliente no es cliente varios
    If vParamAplic.ContabilizarTicketAgrupados Then
        If Val(Text1(0).Text) <> vParamTPV.Cliente Then
            MsgBox "Si agrupa la contabilizacion de los tickets el cliente debe ser: " & vParamTPV.Cliente, vbExclamation
            Exit Sub
        End If
    End If
    
    
    
    If GenerarTicket Then
        
        
        
        'Imprimir el ticket
        'si el parametro imprimir siempre ticket (spatpvg.imprtick=1) se imprimr directamente
        'si no preguntar si se quiere imprimir
''''''
''''''        If DevuelveDesdeBDNew(conAri, "spatpvg", "imprtick", "codigo", "1", "N") = "1" Then
''''''            'ImprimirTicket
''''''            ImprimirTicketDirecto vNumTicket, vNumAlbTicket, RSVenta!fecventa
''''''        Else
''''''            'If MsgBox("¿Desea imprimir el ticket?", vbQuestion + vbYesNo) = vbYes Then ImprimirTicket
''''''            If MsgBox("¿Desea imprimir el ticket?", vbQuestion + vbYesNo) = vbYes Then _
''''''                        ImprimirTicketDirecto vNumTicket, vNumAlbTicket, RSVenta!fecventa
''''''        End If
        Impr = True
        If Not vParamTPV.ImprimiDirecto Then
            If MsgBox("¿Desea imprimir el ticket?", vbQuestion + vbYesNo) = vbNo Then Impr = False
        End If
        'If Impr Then ImprimirTicketDirecto vNumTicket, vNumAlbTicket, RSVenta!fecventa
        
        '# Modificado: LAURA (25/07/2008)
        If Text1(3).Text <> "" Then curEntregado = CCur(Text1(3).Text)
        If Label2(4).Caption <> "" Then curCambio = CCur(Label2(4).Caption)
        If Impr Then ImprimirTicketDirecto vNumTicket, RSVenta!fecventa, curEntregado, curCambio
        '#
        
        'cerrar ventana total y regresar a entrada de ventas
        cadSel = "1"
'        Unload Me
        mnSalir_Click
    End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), 3
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then BotonBuscar (Index)
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Text1_LostFocus(Index As Integer)
Dim ImpCheque As Currency
Dim devuelve As String

    Text1(Index).Text = Trim(Text1(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    

    Select Case Index
        Case 0 'cod cliente
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sclien", "nomclien", "codclien", "cliente", "N")
                Text1(Index).Text = Format(Text1(Index).Text, "000000")
                If Text2(Index).Text = "" Then
                    PonerFoco Text1(Index)
                ElseIf ClienteOK(Text1(Index), RSVenta!CodClien, True) Then
                    If Text1(1).Text = "" Then
                        'recuperar la forma de pago del cliente
                        SQL = DevuelveDesdeBD(conAri, "codforpa", "sclien", "codclien", Text1(Index).Text, "N")
                        Text1(1).Text = SQL
                        Text1_LostFocus (1)
                    End If
                Else
                    Text1(Index).Text = RSVenta!CodClien
                    Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sclien", "nomclien", "codclien", "cliente", "N")
                    Text1(Index).Text = Format(Text1(Index).Text, "000000")
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
                PonerFoco Text1(Index)
            End If
            
        Case 1 'cod forpa
            If PonerFormatoEntero(Text1(Index)) Then
                Text1(Index).Text = Format(Text1(Index).Text, "000")
'                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sforpa", "nomforpa", "codforpa", "Forma de pago", "N")
                Text2(Index).Text = DevuelveDesdeBDNew(conAri, "sforpa", "nomforpa", "codforpa", Text1(Index).Text, "N")
                If Text2(Index).Text = "" Then
                    PonerFoco Text1(Index)
                Else
                    SQL = DevuelveDesdeBD(conAri, "tipforpa", "sforpa", "codforpa", Text1(Index).Text, "N")
                    TipoForPa = CByte(SQL)
                    Me.FrameEfectivo.Enabled = (TipoForPa = 0)
                    If TipoForPa <> 0 Then
                        Me.Label2(4).Caption = ""
                        Me.Text1(3).Text = ""
                    Else
                        'Forpa correcta. SI NO tiene checque regalo lo posicionamos
                        If vParamAplic.ForPagoChequeRegalo = "" Then PonerFoco Text1(3)
                        
                        
                    End If
'                    If Screen.ActiveControl.Name <> "Text1" And Text1(3).Enabled = True Then PonerFoco Text1(3) '(RAFA/ALZIRA 31082006)
                End If
            Else
                If Text2(Index).Text <> "" Then
                    Text2(Index).Text = ""
                    PonerFoco Text1(1)
                Else
'                    PonerFoco Text1(1)
                End If
            End If
        
        Case 2 'cod trabajador
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "straba", "nomtraba", "codtraba", "Operador", "N")
            Text1(Index).Text = Format(Text1(Index).Text, "0000")
            
        Case 3 'Entregado
            If PonerFormatoDecimal(Text1(Index), 1) Then
                'obtener el importe del cheque regalo si hay
                ImpCheque = CCur(ComprobarCero(Text1(4).Text))
                'Obtener el cambio= entregado + cheque_regalo - importe
                Label2(4).Caption = Format(CCur(Text1(Index).Text) + ImpCheque - CCur(Importe), FormatoImporte)
                PonerFoco Text1(1)
            Else
                Label2(4).Caption = ""
            End If
            frmFacTPVEnt.EnviarVisorPuerto Label2(3).Caption, Label2(4).Caption, Label2(0).Caption, Label2(1).Caption
            
        Case 4 'cheque regalo
             If PonerFormatoDecimal(Text1(Index), 1) Then
                If Me.Text1(3).Enabled = False Then PonerFoco Text1(1)
             End If
             
        Case 5 'DIREC./DPTO.
             If PonerFormatoEntero(Text1(Index)) Then
                Text1(Index).Text = Format(Text1(Index).Text, "000")
                'Comprobar que el cliente seleccionado tiene esa direccion
                If PonerDptoEnCliente Then
                    'Comprobar que el cliente tiene mantenimientos en esa direc/dpto
                    devuelve = DevuelveDesdeBDNew(conAri, "scaman", "nummante", "codclien", Text1(0).Text, "N", , "coddirec", Text1(5).Text, "N")
                    If devuelve <> "" Then MsgBox "El cliente tiene Mantenimientos.", vbInformation
                Else
                    Text1(Index).Text = ""
                    Text2(Index).Text = ""
                    PonerFoco Text1(Index)
                End If
             Else
                Text2(Index).Text = ""
'                PonerFoco Text1(Index)
             End If
    End Select
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 3, cerrar
    If KeyAscii = 27 Then cerrar = True
    If cerrar Then Unload Me
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 2 'Generar ticket
            mnTicket_Click
        Case 4 'Generar Albaran
            mnAlbaran_Click
        Case 6 'Generar Factura
            mnFactura_Click
            
        Case 10 'Salir
            mnSalir_Click
    End Select
End Sub



Private Function InicializarCStock(ByRef vCStock As CStock, TipoM As String, ByRef RS As ADODB.Recordset) As Boolean
'On Error Resume Next
On Error Resume Next

    vCStock.tipoMov = TipoM
    vCStock.HoraMov = CStr(RS!horventa)
    vCStock.codArtic = RS!codArtic
    vCStock.codAlmac = codAlmac
    vCStock.Cantidad = CSng(RS!Cantidad)
    '16 Mayo 08
    '----------
    ' El importe de la linea esta en una columna de la BD
    'vCStock.Importe = CCur(RS!Cantidad) * CCur(RS!precioar)
    vCStock.Importe = RS!implineareal
    vCStock.LineaDocu = CInt(RS!numlinea)
        
    If Err.Number <> 0 Then
        MsgBox "No se han podido inicializar la clase para actualizar Stock", vbExclamation
        InicializarCStock = False
    Else
        InicializarCStock = True
    End If
End Function


Private Function ObtenerContadorTicket(NumTicket As String) As Boolean
Dim vTipoMov As CTiposMov

    On Error Resume Next

    CodTipoMov = "FTI"
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
        NumTicket = vTipoMov.ConseguirContador(CodTipoMov)
        If NumTicket <> "-1" Then ObtenerContadorTicket = True
        
        vTipoMov.IncrementarContador (CodTipoMov)
    Else
        ObtenerContadorTicket = False
    End If
    Set vTipoMov = Nothing
    
    If Err.Number <> 0 Then ObtenerContadorTicket = False
End Function



'01/09/06 Laura
Private Function ObtenerContadorAlbTicket(NumAlbTicket As String) As Boolean
Dim vTipoMov As CTiposMov

    On Error Resume Next

    CodTipoMov = "ATI"
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
        NumAlbTicket = vTipoMov.ConseguirContador(CodTipoMov)
        If NumAlbTicket <> "-1" Then ObtenerContadorAlbTicket = True
        
        vTipoMov.IncrementarContador (CodTipoMov)
    Else
        ObtenerContadorAlbTicket = False
    End If
    Set vTipoMov = Nothing
    
    If Err.Number <> 0 Then ObtenerContadorAlbTicket = False
End Function



Private Function ObtenerContadorAlbaran(NumAlb As String) As Boolean
Dim vTipoMov As CTiposMov
Dim Existe As Boolean

    On Error GoTo ErrConAlb

    CodTipoMov = "ALV"
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
        Do
            NumAlb = vTipoMov.ConseguirContador(CodTipoMov)
            vTipoMov.IncrementarContador (CodTipoMov)
            SQL = "select count(*) from scaalb where codtipom='" & CodTipoMov & "' and numalbar=" & NumAlb
            Existe = (RegistrosAListar(SQL) > 0)
        Loop Until Existe = False
        ObtenerContadorAlbaran = True
    Else
        ObtenerContadorAlbaran = False
    End If
    Set vTipoMov = Nothing
    Exit Function
    
ErrConAlb:
    ObtenerContadorAlbaran = False
    MuestraError Err.Number, "Obtener contador albaran", Err.Description
End Function




Private Function ObtenerContadorFactura(NumFactu As String) As Boolean
Dim vTipoMov As CTiposMov
Dim Existe As Boolean

    On Error GoTo ErrConFac

    CodTipoMov = "FAV"
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
        Do
            NumFactu = vTipoMov.ConseguirContador(CodTipoMov)
            vTipoMov.IncrementarContador (CodTipoMov)
            SQL = "select count(*) from scafac where codtipom='" & CodTipoMov & "' and numfactu=" & NumFactu & " and fecfactu=" & DBSet(RSVenta!fecventa, "F")
            Existe = (RegistrosAListar(SQL) > 0)
        Loop Until Existe = False
        ObtenerContadorFactura = True
    Else
        ObtenerContadorFactura = False
    End If
    Set vTipoMov = Nothing
    Exit Function
    
ErrConFac:
    ObtenerContadorFactura = False
    MuestraError Err.Number, "Obtener contador factura", Err.Description
End Function



Private Function InsertarMovAlmacen(NumTicket As String) As Boolean
'PAra tickets, albaranes y facturas
Dim RS As ADODB.Recordset
Dim vCStock As CStock
Dim B As Boolean

    On Error GoTo EInsMov
    
    'Para cada linea de venta insertar el movimiento e actualizar stocks
    Set RS = New ADODB.Recordset
    Set vCStock = New CStock
    
    SQL = Replace(cadSel, "scaven", "sliven")
    SQL = "SELECT * FROM sliven WHERE " & SQL
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    vCStock.Documento = NumTicket
    vCStock.DetaMov = CodTipoMov
    vCStock.Trabajador = CLng(Text1(0).Text) 'sera el cliente
    vCStock.Fechamov = CStr(RSVenta!fecventa)
    
    B = True
    While Not RS.EOF And B
        If Not InicializarCStock(vCStock, "S", RS) Then Exit Function
        If Not vCStock.ActualizarStock(True) Then B = False
        RS.MoveNext
    Wend
    
    Set vCStock = Nothing
    RS.Close
    Set RS = Nothing
    
EInsMov:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Insertando movimientos de almacen.", Err.Description
        B = False
        Set vCStock = Nothing
        RS.Close
        Set RS = Nothing
    End If
    InsertarMovAlmacen = B
End Function



Private Function InsertarHistFactura(NumTicket As String, Optional NumFactu As String, Optional NumAlbTicket As String, Optional MenError As String) As Boolean
Dim B As Boolean
Dim vFactu As CFactura
Dim vClien As CCliente

    On Error GoTo EInsFac
    
    SQL = ""
    'Insertar la cabecera de Factura (scafac)
    Set vFactu = New CFactura
    If NumFactu = "" Then
        vFactu.Codtipom = "FTI"
        vFactu.NumFactu = NumTicket
    Else
        vFactu.Codtipom = "FAV"
        vFactu.NumFactu = NumFactu
    End If
    vFactu.FecFactu = Format(RSVenta!fecventa, "dd/mm/yyyy")
    vFactu.NumTerminal = RSVenta!NumTermi
    vFactu.NumVenta = RSVenta!NumVenta
    
    
    'Guardamos los valores identificativos de la factura generada
    'para imprimirla posteriormente
    cadImpresion = "{scafac.codtipom}='" & vFactu.Codtipom & "' and {scafac.numfactu}=" & vFactu.NumFactu

    vFactu.Cliente = Text1(0).Text
    vFactu.DirDpto = Text1(5).Text
    vFactu.NombreDirDpto = Text2(5).Text
    If vFactu.Cliente <> "" Then
        Set vClien = New CCliente
        If vClien.LeerDatos(vFactu.Cliente) Then
            vFactu.NombreClien = vClien.Nombre
            vFactu.DomicilioClien = vClien.Domicilio
            vFactu.CPostal = vClien.CPostal
            vFactu.Poblacion = vClien.Poblacion
            vFactu.Provincia = vClien.Provincia
            vFactu.NIF = vClien.NIF
            vFactu.Telefono = vClien.TfnoClien
            vFactu.Agente = vClien.Agente
            vFactu.Banco = vClien.Banco
            vFactu.Sucursal = vClien.Sucursal
            vFactu.DigControl = vClien.DigControl
            vFactu.CuentaBan = vClien.CuentaBan
            'Actualizamos fecha ult. movim del cliente si es posterior
            B = vClien.ActualizaUltFecMovim(vFactu.FecFactu)
        Else
            InsertarHistFactura = False
            Exit Function
        End If
        Set vClien = Nothing
    End If
    'obtener letra serie de la factura (para el tipo de movimiento)
    SQL = DevuelveDesdeBDNew(conAri, "stipom", "letraser", "codtipom", vFactu.Codtipom, "T")
    vFactu.LetraSerie = SQL
    
    vFactu.ForPago = Text1(1).Text
    vFactu.TipForPago = TipoForPa
    vFactu.TotalFac = Importe
     
    'Cuenta Prevista de Cobro de las Facturas
    vFactu.BancoPr = vParamTPV.CtaPrevCobro
    vFactu.CuentaPrev = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", vFactu.BancoPr, "N")
   
    If vFactu.CuentaPrev = "" Then
        B = False
        SQL = "La cuenta prevista de cobro no puede ser nula. Parámetos TPV."
    End If
    
    If B Then
        SQL = Text1(2).Text 'Trabajador
        B = B And vFactu.PasarTicketAFactura(cadSel, SQL, NumTicket, NumAlbTicket, Text1(4).Text)
    End If
    
'    If Not b Then MsgBox SQL, vbInformation
    If Not B Then MenError = SQL
    Set vFactu = Nothing
    
EInsFac:
    If Err.Number <> 0 Then
        'MuestraError Err.Number, "Insertando Histórico de Factura.", Err.Description
        MenError = "Insertando Histórico de Factura." & vbCrLf & Err.Description
        B = False
    End If
    InsertarHistFactura = B
End Function



Private Function InsertarAlbaran(NumAlb As String, NumTicket As String, menErr As String) As Boolean
Dim B As Boolean
Dim vClien As CCliente

    On Error GoTo EInsAlb

    'Cabecera de albaran
    '----------------------------------
    SQL = "INSERT INTO scaalb (codtipom,numalbar,fechaalb,factursn,codclien,nomclien,domclien,codpobla,pobclien,proclien,nifclien,telclien,"
    SQL = SQL & "coddirec,nomdirec,referenc,codtraba,codtrab1,codtrab2,codagent,codforpa,codenvio,dtoppago,dtognral,tipofact,"
    SQL = SQL & "observa01,observa02,observa03,observa04,observa05,numofert,fecofert,numpedcl,fecpedcl,fecentre,sementre,esticket,numtermi,numventa) "
    'Abril 2008
    'Pongo la marca de facturar a TRUE: 1
    SQL = SQL & " VALUES ('" & CodTipoMov & "'," & NumAlb & "," & DBSet(RSVenta!fecventa, "F") & ",1," & Text1(0).Text & ","
    
    'Obtenemos los datos del cliente
    Set vClien = New CCliente
    If vClien.Existe(Text1(0).Text) Then
        If vClien.LeerDatos(Text1(0).Text) Then
            SQL = SQL & DBSet(vClien.Nombre, "T", "N") & ", " & DBSet(vClien.Domicilio, "T", "N") & ","
            SQL = SQL & DBSet(vClien.CPostal, "T", "N") & ", " & DBSet(vClien.Poblacion, "T", "N") & "," & DBSet(vClien.Provincia, "T", "N") & ","
            SQL = SQL & DBSet(vClien.NIF, "T", "N") & "," & DBSet(vClien.TfnoClien, "T") & "," & DBSet(Text1(5).Text, "N", "S") & "," & DBSet(Text2(5).Text, "T") & "," & ValorNulo & "," 'coddirec,nomdirec,referenc a nulo
            SQL = SQL & Text1(2).Text & "," & Text1(2).Text & "," & Text1(2).Text & "," 'trabajador
            SQL = SQL & vClien.Agente & "," & Text1(1).Text & "," & vClien.FEnvio & ",0,0," & vClien.TipoFactu & ","
            'observaciones
            'La primera observacion sera el campo de observaciones de la venta
            If IsNull(RSVenta!observa1) Then
                SQL = SQL & ValorNulo
            Else
                SQL = SQL & "'" & DevNombreSQL(RSVenta!observa1) & "'"
            End If
            SQL = SQL & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            'datos oferta: aqui guardamos nº venta
            SQL = SQL & ValorNulo & "," & ValorNulo & ","
            'En los campos de datos del pedido guardamos los datos del ticket
            SQL = SQL & NumTicket & "," & DBSet(RSVenta!fecventa, "F") & "," & ValorNulo & "," & ValorNulo & ",1," & DBSet(RSVenta!NumTermi, "N") & "," & DBSet(RSVenta!NumVenta, "N", "S") & ")" 'esticket=1, terminal
            B = vClien.ActualizaUltFecMovim(RSVenta!fecventa)
        Else
            B = False
        End If
    End If
    Set vClien = Nothing
    
    
    If B Then
        'Insertar Cabecera
'    MenError = "Error al insertar en la tabla Cabecera de Albaranes (scaalb )."
        Conn.Execute SQL, , adCmdText
        
        'Lineas del albaran
        'Inserta en tabla "slialb" todas las lineas de venta
        SQL = "INSERT INTO slialb "
        SQL = SQL & "(codtipom, numalbar,numlinea, codalmac, codartic, nomartic, ampliaci, cantidad,"
        SQL = SQL & "precioar, dtoline1, dtoline2, importel, origpre,codprovex) "
        
        
        'SQL = SQL & " SELECT '" & CodTipoMov & "' as codtipom," & DBSet(NumAlb, "N") & " as numalbar," & "numlinea," & codAlmac & " as codalmac," & "codartic,nomartic," & ValorNulo & " as ampliaci,cantidad,precioar,0 as dtoline1,0 as dtoline2, round(cantidad*precioar,2) as importel,'' as origpre "
        'SQL = SQL & " FROM sliven WHERE " & Replace(cadSel, "scaven", "sliven")
        
        'Neuvo Abril 2008. David
        'Para llevar el codprove a la linea de albaran y que no ponga el 0
        SQL = SQL & " SELECT '" & CodTipoMov & "' as codtipom," & DBSet(NumAlb, "N") & " as numalbar," & "numlinea," & codAlmac & " as codalmac,"
        SQL = SQL & " sliven.codartic,sliven.nomartic," & ValorNulo & " as ampliaci,cantidad,precioar,"
        SQL = SQL & "dto1 as dtoline1,dto2 as dtoline2,"
        'NUEVO###
        'David.    La linea puede llevar dtos, con lo cual hay un
        '           campo en sliven que lleva el importe real de la linea
        ' ANTES:  round(cantidad*precioar,2) as importel
        SQL = SQL & " implineareal as importel,'' as origpre ,codprove"
        SQL = SQL & " FROM sliven,sartic WHERE sliven.codArtic = sartic.codArtic AND " & Replace(cadSel, "scaven", "sliven")
        
        
        Conn.Execute SQL, , adCmdText
    End If

     
    'Eliminar las ventas que se han pasado a albaranes
    If B Then B = EliminarVenta(cadSel)
    
    'Guardamos los valores identificativos de la factura generada
    'para imprimirla posteriormente
    If B Then cadImpresion = "{scaalb.codtipom}='" & CodTipoMov & "' and {scaalb.numalbar}=" & DBSet(NumAlb, "N")

EInsAlb:
    If Err.Number <> 0 Then
        menErr = "Insertando el Albaran: " & vbCrLf & Err.Description
        B = False
    End If
    InsertarAlbaran = B
End Function



Private Function DatosOk() As Boolean
Dim B As Boolean
Dim I As Byte
Dim cad As String

    On Error GoTo EDatosOK
    B = True
    
    'Comprobaciones
    '------------------
    
    'comprobar que los campos tienen valor
    For I = 0 To 2
        If Trim(Me.Text1(I).Text) = "" Then
            If I = 0 Then
                cad = "Cliente"
            ElseIf I = 1 Then
                cad = "Forma de pago"
            ElseIf I = 2 Then
                cad = "Operador"
            End If
            MsgBox "El campo " & cad & " debe tener valor.", vbInformation
            B = False
            Exit For
        End If
    Next I
    
    'comprobar que el trabajador existe
    If B Then
        If DevuelveDesdeBDNew(conAri, "straba", "codtraba", "codtraba", Text1(2).Text, "N") = "" Then
            B = False
            MsgBox "No existe el trabajador " & Text1(2).Text, vbExclamation
        End If
    End If
    
    
    '--- Laura: 11/04/2007
    '--- comprobar q el cliente no esta bloqueado y q si se ha cambiado sea de la mista
    '--- tarifa q para el q se insertaron las lineas
    If B Then
        B = ClienteOK(Text1(0), RSVenta!CodClien, False)
        If Not B Then Text1(0).Text = RSVenta!CodClien
    End If
    '---
    
    
    '--- Laura: 12/04/2007
    '--- comprobar q si es cliente contado el tipo de forma de pago sea efectivo
    If B Then
        'obtenemos tipoforpa correcta por si acaso
        cad = DevuelveDesdeBD(conAri, "tipforpa", "sforpa", "codforpa", Text1(1).Text, "N")
        If cad = "" Then
            B = False
            MsgBox "No existe la forma de pago.", vbExclamation
        Else
            TipoForPa = CByte(cad)
        
            'si se ha definido un cliente como contado en parametros del TPV
            If vParamTPV.Cliente <> "" Then
                If CLng(Text1(0).Text) = CLng(vParamTPV.Cliente) Then 'si es cliente definido como CONTADO
                    'Aceptamos EFECTIVO y  TARJETA DE CREDITO
                    If TipoForPa <> 0 And TipoForPa <> 6 Then 'tiene q tener tipo forpa EFECTIVO or TARJ CREDIT
                        B = False
                        MsgBox "El cliente '" & Text2(0).Text & "' debe tener una Forma de Pago de tipo EFECTIVO.", vbExclamation
                    End If
                End If
            End If
        End If
    End If
    '---
    
    
    If B Then
        If TipoForPa = 0 Then 'Contado
            If Me.Text1(3).Text = "" Then
                MsgBox "Debe introducir la cantidad a pagar.", vbInformation
                B = False
            Else
                If Not vParamTPV.Rapida Then
            
                    If (CCur(Me.Text1(3).Text) + CCur(ComprobarCero(Text1(4).Text))) < CCur(Me.Label2(1).Caption) Then
                        MsgBox "La cantidad entregada debe ser igual o superior al importe total.", vbInformation
                        B = False
                    End If
                End If
            End If
        ElseIf TipoForPa = 4 Then 'Recibo
            'comprueba que el cliente tenga cuenta bancaria OK sino
            'muestra aviso pero deja pasar
            ComprobarCtaBanCliente
        End If
    End If
    
    
    '--- Laura: 18/12/2006
    'direc./dpto del cliente
    If B And Text1(5).Text <> "" Then
        'comprobar q existe el dpto para el cliente
        B = PonerDptoEnCliente
    End If
    
    
    '--- Laura: 01/12/2006
    'si hay cheque regalo
    If B Then
        If Me.Text1(4).Text <> "" Then
            'comprobar q en parametros de la aplicacion el campo codforpa tiene valor
            If vParamAplic.ForPagoChequeRegalo = CCur(Me.Label2(1).Caption) Then
                MsgBox "No se ha introducido la forma de pago del cheque regalo." & vbCrLf & "Configurar parámetros aplicación.", vbInformation, "Comprobar datos"
                B = False
            End If
            'comprobar que el importe del cheque sea >= q total factura
            If CCur(Me.Text1(4).Text) > CCur(Me.Label2(1).Caption) Then
                MsgBox "El importe del cheque regalo no puede ser superior al TOTAL.", vbExclamation
                B = False
            End If
        End If
    End If
    
    DatosOk = B
    Exit Function
    
EDatosOK:
    MuestraError Err.Number, "Comprobando datos.", Err.Description
    DatosOk = False
End Function



Private Function GenerarTicket() As Boolean
Dim B As Boolean
Dim NumTicket As String
'01/09/06
Dim NumAlbTicket As String
Dim MenError As String

    On Error GoTo ETicket
    
    Conn.BeginTrans
    'si el tipo de forma de pago no es efectivo habrá que insertar
    'en la tabla de contabilidad conta.scobro
'    If TipoForPa <> 0 Then
    ConnConta.BeginTrans
    
'    PreparaBloquear
        
    'Obtener el contador de ticket (FTI).
    B = ObtenerContadorTicket(NumTicket)
    
    'Obtener el contador albaran de ticket (ATI).
    If B Then B = ObtenerContadorAlbTicket(NumAlbTicket)
    
    If B Then
        'Actualizar los stocks de todos los articulos comprados
        'Insertar movimiento en smoval
        B = InsertarMovAlmacen(NumAlbTicket)
    
        'Insertar en el historico de facturas: scafac, scafac1,slifac
        'en el campo scafac1.numalbar guardamos el nº de ticket
        If B Then B = InsertarHistFactura(NumTicket, , NumAlbTicket, MenError)
    End If
    vNumTicket = NumTicket ' (RAFA/ALZIRA 05092006)
    vNumAlbTicket = NumAlbTicket ' (RAFA/ALZIRA 05092006)
    
ETicket:
    If Err.Number <> 0 Then
        B = False
        MenError = Err.Description
    End If
    If B Then
        Conn.CommitTrans
        'If TipoForPa <> 0 Then
        ConnConta.CommitTrans
    Else
        Conn.RollbackTrans
        'If TipoForPa <> 0 Then
        ConnConta.RollbackTrans
        MsgBox "ERROR: " & vbCrLf & MenError, vbExclamation, "Generar Ticket"
    End If
    GenerarTicket = B
    TerminaBloquear
    Espera 0.2
End Function



Private Function GenerarAlbaran(NumAlb As String) As Boolean
'La venta se combierte en un albaran.
Dim B As Boolean
Dim NumTicket As String
Dim MenError As String

    On Error GoTo EAlbar
    Conn.BeginTrans
   
    'Obtener el contador de ticket (FTI).
    B = ObtenerContadorAlbTicket(NumTicket)
    
    If B Then
        'Obtener el contador de Albaran (ALV).
        B = ObtenerContadorAlbaran(NumAlb)
        
        If B Then
            'Actualizar los stocks de todos los articulos comprados
            'Insertar movimiento en smoval
            B = InsertarMovAlmacen(NumAlb)
    
            'Insertar en las tablas de Albaranes: scaalb, slialb
            'en el campo scafac1.numalbar guardamos el nº de ticket
            If B Then B = InsertarAlbaran(NumAlb, NumTicket, MenError)
        End If
    End If
    
EAlbar:
    If Err.Number <> 0 Then
        MenError = MenError & vbCrLf & vbCrLf & Err.Description
        B = False
    End If
    If B Then
        Conn.CommitTrans
    Else
        Conn.RollbackTrans
        MsgBox MenError, vbExclamation, "Generar Albaran"
    End If
    GenerarAlbaran = B
    Espera 0.2
End Function



Private Function GenerarFactura(NumFactu As String) As Boolean
Dim B As Boolean
Dim NumTicket As String
Dim MenError As String
    
    On Error GoTo EGenFac
    
    Conn.BeginTrans
    'si el tipo de forma de pago no es efectivo habrá que insertar
    'en la tabla de contabilidad conta.scobro
    '---- Laura: 10/10/2006 siempre se inserta en la scobro aunque sea efectivo
'    If TipoForPa <> 0 Then ConnConta.BeginTrans
    ConnConta.BeginTrans
    
    'Obtener el contador de ticket (ATI).
    B = ObtenerContadorAlbTicket(NumTicket)
    
    If B Then B = ObtenerContadorFactura(NumFactu)
    
    If B Then
        'Actualizar los stocks de todos los articulos comprados
        'Insertar movimiento en smoval
        CodTipoMov = "ATI"
        B = InsertarMovAlmacen(NumTicket)
    
        'Insertar en el historico de facturas: scafac, scafac1,slifac
        'en el campo scafac1.numalbar guardamos el nº de ticket
        If B Then
            CodTipoMov = "FAV"
            B = InsertarHistFactura(NumTicket, NumFactu, , MenError)
        End If
    End If
    
EGenFac:
    If Err.Number <> 0 Then
        MenError = Err.Description
        B = False
    End If
    If B Then
        Conn.CommitTrans
        '---- Laura 10/10/2006: siempre se inserta en la conta.scobro aunque sea efectivo
        'If TipoForPa <> 0 Then ConnConta.CommitTrans
        ConnConta.CommitTrans
    Else
        Conn.RollbackTrans
        '---- Laura 10/10/2006: siempre se inserta en la conta.scobro aunque sea efectivo
        'If TipoForPa <> 0 Then ConnConta.RollbackTrans
        ConnConta.RollbackTrans
        MsgBox "ERROR: " & MenError & vbCrLf, vbExclamation, "Generar Factura"
    End If
    GenerarFactura = B
    Espera 0.2
End Function



Private Sub ImprimirTicket()
Dim MIPATH As String
'Dim NomImpre As String

    On Error GoTo EImpTick
    
    SQL = cadImpresion & " and {scafac.fecfactu}=" & DBSet(RSVenta!fecventa, "F")
    If Not HayRegParaInforme("scafac", SQL) Then Exit Sub


    MIPATH = App.Path & "\Informes\"
    cadImpresion = cadImpresion & " and {scafac.fecfactu}=Date(" & Year(RSVenta!fecventa) & "," & Month(RSVenta!fecventa) & "," & Day(RSVenta!fecventa) & ")"
    
'    'Establecemos la impresora de ticket
'    If vParamTPV.NomImpresora <> "" Then
'        If Printer.DeviceName <> vParamTPV.NomImpresora Then
'            'guardamos la impresora que habia
'            NomImpre = Printer.DeviceName
'            'establecemos la de ticket
'            EstablecerImpresora vParamTPV.NomImpresora
'        End If
'    End If

    With frmVisReport
        .FormulaSeleccion = cadImpresion
        .SoloImprimir = True
        .OtrosParametros = ""
        .NumeroParametros = 0
        .MostrarTree = False
        .Informe = MIPATH & "rTPVTicket.rpt"
        .ConSubInforme = False
        .opcion = 93
        .ExportarPDF = False
        .Show vbModal
    End With
    
    'volver la impresora a la predeterminada
'    EstablecerImpresora NomImpre
    
EImpTick:
    If Err.Number <> 0 Then MuestraError Err.Number, "Imprimir ticket.", Err.Description
End Sub



Private Sub ImprimirFactura()
Dim MIPATH As String
Dim cadParam As String, nomDocu As String
Dim numParam As Byte
Dim ImprimeDirecto As Boolean

    On Error Resume Next

    SQL = cadImpresion & " and {scafac.fecfactu}=" & DBSet(RSVenta!fecventa, "F")
    If Not HayRegParaInforme("scafac", SQL) Then Exit Sub
     

    
    
    '===================================================
    '============ PARAMETROS ===========================
    
    If Not PonerParamRPT(18, cadParam, numParam, nomDocu, ImprimeDirecto) Then Exit Sub
        
    '===================================================
    If ImprimeDirecto Then
        cadImpresion = cadImpresion & " and {scafac.fecfactu}='" & Format(RSVenta!fecventa, FormatoFecha) & "'"
        SQL = cadImpresion
        SQL = Replace(SQL, "{", "")
        SQL = Replace(SQL, "}", "")
        ImprimirDirectoFact SQL
    Else
        MIPATH = App.Path & "\Informes\"
        cadImpresion = cadImpresion & " and {scafac.fecfactu}=Date(" & Year(RSVenta!fecventa) & "," & Month(RSVenta!fecventa) & "," & Day(RSVenta!fecventa) & ")"
    
    
         With frmVisReport
            .FormulaSeleccion = cadImpresion
            .SoloImprimir = True ' (RAFA/ALZIRA 31082006)
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .MostrarTree = False
            .Informe = MIPATH & nomDocu
            .ConSubInforme = True
            .opcion = 53
            .ExportarPDF = False
            .NumCopias2 = 2 ' (RAFA/ALZIRA 31082006)
            .Show vbModal
        End With
    End If
    If Err.Number <> 0 Then MuestraError Err.Number, "Imprimir Factura.", Err.Description
End Sub




Private Sub ImprimirAlbaran()
Dim MIPATH As String
Dim cadParam As String, nomDocu As String
Dim numParam As Byte
Dim ImprimeDirecto As Boolean

    SQL = cadImpresion '& " and {scafac.fecfactu}=" & DBSet(RSVenta!fecventa, "F")
    If Not HayRegParaInforme("scaalb", SQL) Then Exit Sub
     
    MIPATH = App.Path & "\Informes\"
    
    '===================================================
    '============ PARAMETROS ===========================
    
    If Not PonerParamRPT(10, cadParam, numParam, nomDocu, ImprimeDirecto) Then Exit Sub
    
    'Añadir el codigo de usuario como parametro para link con tabla Temporal (tmptiposiva) en el Report
    'tabla temporal para el calculo del bruto total para cada tipo de IVA
    cadParam = cadParam & "pCodUsu=" & vUsu.Codigo & "|"
    numParam = numParam + 1
    
    
    '=========================================================================
    'Aqui sabemos que valor tiene CodClien y añadimos a los parametros el tipo de IVA
    'que se aplica a ese cliente
    SQL = DevuelveDesdeBDNew(conAri, "sclien", "tipoiva", "codclien", Text1(0).Text, "N")
    If SQL <> "" Then
        cadParam = cadParam & "pTipoIVA=" & SQL & "|"
        numParam = numParam + 1
    End If
    
    
    '===================================================
    If ImprimeDirecto Then
        SQL = cadImpresion
        SQL = Replace(SQL, "{", "")
        SQL = Replace(SQL, "}", "")
        ImprimirDirectoAlb SQL
    
    Else
    
         With frmVisReport
            .FormulaSeleccion = cadImpresion
            .SoloImprimir = True ' (RAFA/ALZIRA 31082006)
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .MostrarTree = False
            .Informe = MIPATH & nomDocu
            .ConSubInforme = True
            .opcion = 45
            .ExportarPDF = False
            .NumCopias2 = 2 ' (RAFA/ALZIRA 31082006)
            .Show vbModal
        End With
    End If
End Sub




Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim Tabla As String
Dim Titulo As String
Dim Desc As String
Dim devuelve As String
Dim I As Byte

    'Llamamos a al form
    '##A mano
    cad = ""
    
    Select Case cadB
        Case "0" 'Cliente
            Tabla = "sclien"
            Titulo = "Clientes"
            devuelve = "0|1|2|3|"
            cad = cad & "Cod.Cli.|sclien|codclien|N|000000|13·"
            cad = cad & "Nom. Cliente|sclien|nomclien|T||47·"
            cad = cad & "Nom. Comer|sclien|nomcomer|T||25·"
            cad = cad & "NIF|sclien|nifclien|T||15·"
            cadB = ""
            
        Case "1" 'Forma pago
            Tabla = "sforpa inner join stippa on sforpa.tipforpa=stippa.tipforpa "
            Titulo = "Formas de Pago"
            devuelve = "0|1|2|"
            cad = cad & "Cod.For.|sforpa|codforpa|N|000|14·"
            cad = cad & "Nom. Forma pago|sforpa|nomforpa|T||50·"
            cad = cad & "Tipo|sforpa|tipforpa|N||12·"
            cad = cad & "Desc Tip.|stippa|destippa|T||23·"
            'cad = cad & "Desc Tip.|sforpa|case tipforpa when 0 then ""Efectivo"" when 1 then ""Transferencia""  when 2 then ""Talón"" when 3 then ""Pagaré"" when 4 then ""Recibo bancario"" when 5 then ""Confirming"" end as desctipo|T||23·"
            cadB = ""
             
        Case "2" 'Trabajador
            Tabla = "straba"
            Titulo = "Operadores"
            devuelve = "0|1|2|"
            cad = cad & "Cod.Op.|straba|codtraba|N|0000|25·"
            cad = cad & "Nom. Operador.|straba|nomtraba|T||55·"
            cad = cad & "NIF|straba|niftraba|T||15·"
            cadB = ""
             
        Case "5" 'direc./dpto del cliente
            If vParamAplic.Departamento Then
                Titulo = "Dptos Cliente: "
                Desc = "Dpto."
            Else
                Titulo = "Direc. Cliente: "
                Desc = "Direc."
            End If
            Titulo = Titulo & Text1(0).Text & " - " & Text2(0).Text
            cad = cad & "Cod. " & Desc & "|sdirec|coddirec|N|000|15·"
            cad = cad & "Desc. " & Desc & "|sdirec|nomdirec|T||65·"
            Tabla = "sdirec"
            devuelve = "0|1|"
            cadB = "codclien=" & Text1(0).Text
    End Select
   
   
    
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = Tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = devuelve
        frmB.vTitulo = Titulo
        frmB.vselElem = 1
        frmB.vConexionGrid = conAri  'Conexión a BD: Ariges
        '#
        If Tabla = "sdirec" Then frmB.Label1.FontSize = 11
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
'        If HaDevueltoDatos Then
'''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'''                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            PonerFoco Text1(kCampo)
        'End If
        I = CInt(Me.imgBuscar(0).Tag)
        Text1_LostFocus (I)
        PonerFoco Text1(I)


    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub BotonBuscar(Ind As Integer)
    imgBuscar_Click (Ind)
End Sub







Private Function PonerDptoEnCliente() As Boolean
Dim vClien As CCliente
Dim NomDpto As String

    Set vClien = New CCliente
    vClien.Codigo = Text1(0).Text
    
    'si existe el departamento para el cliente
    If vClien.DptoCliente(Text1(5).Text, NomDpto) Then
        Text2(5).Text = NomDpto
        PonerDptoEnCliente = True
    Else
        PonerDptoEnCliente = False
    End If
    Set vClien = Nothing
End Function



Private Function ComprobarCtaBanCliente() As Boolean
Dim cCli As CCliente
Dim MenError As String

    Set cCli = New CCliente
    cCli.Codigo = Text1(0).Text
    
    If cCli.LeerDatos(Text1(0).Text) Then
        ComprobarCtaBanCliente = cCli.ComprobarCtaBancaria(MenError)
        If Not ComprobarCtaBanCliente Then MsgBox MenError & vbCrLf & "Contacte con Administración.", vbInformation
    End If
    
    Set cCli = Nothing
End Function


Private Function HayArticuloFitosanitario() As Boolean
'comprueba si entre las lineas de venta insertadas hay algun articulo
'q tiene registro fitosanitario (en ese caso no se puede crear Ticket)
'(OUT) -> true si encuentra algun articulo fitosanitario
Dim SQL As String
Dim RS As ADODB.Recordset
Dim I As Integer

    On Error GoTo ErrFito
    
    SQL = "SELECT distinct sliven.nomartic,numserie FROM sliven"
    SQL = SQL & " inner join sartic on sliven.codartic=sartic.codartic"
    SQL = SQL & " WHERE " & Replace(cadSel, "scaven", "sliven")
    SQL = SQL & " and not isnull(numserie) and trim(numserie)<>''"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If RS.EOF Then
        'no hay articulos con registro fitosanitario
        HayArticuloFitosanitario = False
    Else
        'hay articulos fitosanitarios
        HayArticuloFitosanitario = True
        
        '- seleccionamos algunos articulos para mostrar en el mensaje
        I = 1
        SQL = ""
        While Not RS.EOF And I < 3
            If SQL <> "" Then SQL = SQL & vbCrLf
            SQL = SQL & DBLet(RS!NomArtic, "T") & " (" & DBLet(RS!numSerie, "T") & ")"
            
            I = I + 1
            RS.MoveNext
        Wend
        If I >= 3 And Not RS.EOF Then SQL = SQL & vbCrLf & "..."
        
        '- mostramos mensaje de error
        SQL = "NO se puede crear un Ticket ya que existen articulos con registro Fitosanitario: " & vbCrLf & SQL
        MsgBox SQL, vbExclamation
    End If
    
    RS.Close
    Set RS = Nothing
    Exit Function
    
ErrFito:
    MuestraError Err.Number, "Comprobar articulos fitosanitarios", Err.Description
End Function



Private Function ClienteOK(newCli As String, oldCli As String, Optional mostrarObs As Boolean) As Boolean
'(IN) newCli: cliente nuevo q queremos poner
'(IN) oldCli: cliente guardado actualmente si existe
Dim cCli As CCliente
Dim devuelve As String

    On Error GoTo ErrCliOK
    ClienteOK = False
    
    If newCli <> "" Then newCli = CStr(Val(newCli))
    Set cCli = New CCliente
    If cCli.LeerDatos(newCli) Then
        '-- Si el cliente esta bloqueado no permitimos este cliente para la venta
        If cCli.ClienteBloqueado Then
            Set cCli = Nothing
            Exit Function
        End If
        
        '-- si se ha modificado el cliente y si hay lineas de articulos:
        '   comprobar q el nuevo cliente tiene la misma tarifa q el cliente anterior
        '   sino no permitimos el nuevo cliente para la venta
        If (oldCli <> "") And (newCli <> oldCli) Then
            'obtener la tarifa del cliente actual
            devuelve = DevuelveDesdeBDNew(conAri, "sclien", "codtarif", "codclien", oldCli, "N")
            If devuelve <> CStr(cCli.Tarifa) Then
                devuelve = "No se puede seleccionar el cliente " & newCli & " "
                devuelve = devuelve & "ya que tiene distinta tarifa de precios." & vbCrLf
                devuelve = devuelve & "Seleccione un cliente de la misma tarifa o elimine la venta."
                MsgBox devuelve, vbExclamation, "Comprobar cliente"
                Set cCli = Nothing
                Exit Function
            End If
        End If
        ClienteOK = True
        
        'mostrar las observaciones del cliente
        If mostrarObs Then cCli.MostrarObservaciones
    End If
    
    Set cCli = Nothing
    Exit Function
    
ErrCliOK:
    MuestraError Err.Number, "Comprobar cliente correcto.", Err.Description
End Function

