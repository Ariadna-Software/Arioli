VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmImprimir 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impresión listados"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   Icon            =   "frmImprimir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pg1 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2340
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdConfigImpre 
      Caption         =   "Sel. &impresora"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   2340
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5340
      TabIndex        =   1
      Top             =   2340
      Width           =   1275
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Default         =   -1  'True
      Height          =   375
      Left            =   3900
      TabIndex        =   0
      Top             =   2340
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   6435
      Begin VB.CheckBox chkEMAIL 
         Caption         =   "Enviar e-mail"
         Height          =   195
         Left            =   4920
         TabIndex        =   8
         Top             =   180
         Width           =   1335
      End
      Begin VB.CheckBox chkSoloImprimir 
         Caption         =   "Previsualizar"
         Height          =   255
         Left            =   420
         TabIndex        =   5
         Top             =   180
         Width           =   1275
      End
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Sin definir"
      Top             =   180
      Width           =   6315
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   7
      Top             =   1320
      Width           =   5535
   End
End
Attribute VB_Name = "frmImprimir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public opcion As Integer
    'Equivale a OpcionListado en frmListado
    'SI ES MAYOR QUE 2000 es ke viene de frmListado2
    
Public FormulaSeleccion As String 'Formula de Seleccion para Crystal Report

Public OtrosParametros As String   ' El grupo acaba en |
                                   ' param1=valor1|param2=valor2|
Public NumeroParametros As Integer
'Cuantos parametros hay.  EMPRESA(EMP) no es parametro. Es fijo en todos los informes

Public SoloImprimir As Boolean
Public EnvioEMail As Boolean

Public NombreRPT As String 'Nombre del fichero de crystal Report .Rpt
Public Titulo As String 'Titulo informe a mostrar en el text1

Public NombreSubRptConta As String 'Nombre del subreport si va conectado a la BDatos Contabilidad

Public ConSubInforme As Boolean 'Para saber si hay subinformes y hay que enlazar las
                                 'tablas a la BD correspondiente

'Julio 2012
Public NumeroDeCopias As Integer


Private MostrarTree As Boolean

Private MIPATH As String
Private Lanzado As Boolean
Private PrimeraVez As Boolean


Private ImpresoraSeleccionada As String


'Private ReestableceSoloImprimir As Boolean
Private Sub chkEMAIL_Click()
    If chkEmail.Value = 1 Then Me.chkSoloImprimir.Value = 0
End Sub

Private Sub chkSoloImprimir_Click()
    If Me.chkSoloImprimir.Value = 1 Then Me.chkEmail.Value = 0
End Sub


Private Sub cmdConfigImpre_Click()
    Screen.MousePointer = vbHourglass
    'Me.CommonDialog1.Flags = cdlPDPageNums
    

    CommonDialog1.PrinterDefault = True
    CommonDialog1.ShowPrinter
    
    PonerNombreImpresora
    Screen.MousePointer = vbDefault
End Sub


Private Sub cmdImprimir_Click()

    If Me.chkSoloImprimir.Value = 1 And Me.chkEmail.Value = 1 Then
        MsgBox "Si desea enviar por mail no debe marcar vista preliminar", vbExclamation
        Exit Sub
    End If
    
    Imprime
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        If SoloImprimir Then
           
            Imprime
            Unload Me
            
        ElseIf Me.EnvioEMail Then
            Me.Hide
            DoEvents
            chkEmail.Value = 1
            Imprime
            Unload Me
        End If
        Espera 0.1
        CommitConexion
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim cad As String

    PrimeraVez = True
    Lanzado = False
    
    PonerBtnSelEmpre
    
    'CargaICO
    cad = Dir(App.Path & "\impre.dat", vbArchive)
    
    'ReestableceSoloImprimir = False
    If cad = "" Then
        chkSoloImprimir.Value = 0
    Else
        chkSoloImprimir.Value = 1
        'ReestableceSoloImprimir = True
    End If
    cmdImprimir.Enabled = True
    
    If SoloImprimir Then
        chkSoloImprimir.Value = 0
        Me.Frame2.Enabled = False
        chkSoloImprimir.visible = False
    Else
        Frame2.Enabled = True
        chkSoloImprimir.visible = True
    End If
    PonerNombreImpresora
    MostrarTree = False

'A partir del infome 26, se trabajaba sobre la b de datos de informes(USUARIOS)

    MIPATH = App.Path & "\Informes\"
'    ConSubInforme = False


    If opcion >= 2000 Then
        'LISTADOS QUE VIENE de frmlistado2
        Select Case opcion
        Case 2001 'Confirmacion de Pedido
            Text1.Text = "Reparaciones efectuadas"
            ConSubInforme = False
            MostrarTree = True
            NombreRPT = "rRepEfectuadas.rpt"
        Case 2002
            Text1.Text = "Listado reparacion x tecnico"
            
        
        Case 2003
            'Esta libre. Lo utlizo para la impresion del justificante del pago de la regarga
            Text1.Text = "Justificante recarga móviles"
            
        Case 2004
            Text1.Text = "Listado Recarga móviles"
            ConSubInforme = False
            MostrarTree = True
            NombreRPT = "rRecargaMov.rpt"
        Case 2006
            Text1.Text = "Listados ventas por proveedor"
            MostrarTree = True
            ConSubInforme = False
            'El nombre lo dejo que venga del form listado2
            
        Case 2009
            Text1.Text = "Facturas proveedor"
            
            
        Case 2010
            Text1.Text = "Albarán proveedor"
        
        Case 2014
            Text1.Text = "Listado tickets agrupados"
      
        Case 2015
            Text1.Text = "Informe traza."
        
        Case 2016
            'Listado tarifas ofertas
            MostrarTree = True
            
            
        Case 2021
            'Oferta bonito
            NombreRPT = DevuelveNombreReport(36)
            Text1.Text = "Oferta clientes"
            
        Case 2022, 2023
            'Acciones realizadas
        '       cadTitulo = "Listado LOG "
    'cadNomRPT = "rListLog" ,tr

            NombreRPT = "rListLogOli"
            Text1.Text = "Acciones realizadas"
            MostrarTree = True
            
            If opcion = 2023 Then
                NombreRPT = NombreRPT & "tr"
                Text1.Text = Text1.Text & "(Trabajador)"
            End If
            NombreRPT = NombreRPT & ".rpt"
        End Select
    Else
        'Normal. Los de antes
                If opcion <= 40 Then
                    Select Case opcion
                    
                    
                    '---------------- Algunos listados basicos
                    Case 5
                        'Tipos de contrato de mantenimiento
                        Text1.Text = "Tipo contrato mantimiento"
                        
                    Case 18 'Informe Stocks Maximos o Minimos
                        Text1.Text = "Stocks Máximos-Mínimos"
                
                    Case 31 'Listado de Ofertas
                        Text1.Text = "Listado de Ofertas"
                        ConSubInforme = True
                    Case 32 'Listado Recordatorio de Ofertas
                        Text1.Text = "Recordatorio de Ofertas"
                        ConSubInforme = True
                    Case 33 'Listado Valoracion de Ofertas
                        Text1.Text = "Valoracion de Ofertas"
                
                    Case 35 'Listado Historico de Ofertas
                        Text1.Text = "Histórico de Ofertas"
                        ConSubInforme = True
                    Case 36 'Listado Ofertas Pendientes y Traspaso a Historico
                        Text1.Text = "Ofertas Pendientes"
                        NombreRPT = "rFacOfePtes.rpt"
                
                    Case 39 'Orden de Instalacion
                        Text1.Text = "Orden de Instalación"
                        ConSubInforme = True
                    Case 40 'Confirmacion de Pedido
                        Text1.Text = "Confirmación de Pedido"
                        ConSubInforme = True
                    Case Else
                        Text1.Text = "Opcion incorrecta"
                        Me.cmdImprimir.Enabled = False
                    End Select
                ElseIf opcion < 100 Then
                    Select Case opcion
                    Case 41 'Informe de Pedidos por Articulo
                        Text1.Text = "Pedidos por Articulo"
                        NombreRPT = "rFacPedxArtic.rpt"
                    Case 42 'Informe de Disponibilidad de Stocks
                        Text1.Text = "Disponibilidad de Stocks"
                        NombreRPT = "rFacPedDispStocks.rpt"
                        ConSubInforme = True
                    Case 44 'Informe de Pedidos por Cliente
                        Text1.Text = "Pedidos por Cliente"
                        NombreRPT = "rFacPedxClien.rpt"
                    Case 47 'Listado de Clientes
                        Text1.Text = "Listado de Cliente"
                        NombreRPT = "rFacClientes.rpt"
                    Case 48 'Informe Altas Nuevos Clientes
                        Text1.Text = "Altas Nuevos Clientes"
                    Case 49 'Informe de Albaranes por Articulo
                        Text1.Text = "" ' dejamos la cadena vacía para que use Titulo [SERVICIOS]
                        NombreRPT = "rFacAlbxArtic.rpt"
                    Case 53 'Factura cliente
                        Text1.Text = "Factura Cliente"
                        ConSubInforme = True
                    Case 54 'Listado Descuentos Familia/Marca
                        Text1.Text = "Listado Descuentos Familia/Marca"
                        NombreRPT = "rFacDtosFM.rpt"
                    Case 58 'Listado Proveedor
                        Text1.Text = "Listado Proveedores"
                        ConSubInforme = False
                         NombreRPT = "rComProve.rpt"
                    Case 60 'Informe Equipos con Nº Serie
                        Text1.Text = "Equipos con Nº Serie"
                        ConSubInforme = True
                    Case 61 'Informe Motivos Pend. Rep.
                        NombreRPT = "rRepMotivosPend.rpt"
                        Text1.Text = "Motivos Pend. Rep."
                    Case 62 'Listado Resguardo Reparacion
                        Text1.Text = "Resguardo Reparación"
                    
                    Case 63 'FACTURAs del TPV
                        Text1.Text = "Facturas formato TPV"
                        ConSubInforme = True
                    
                    Case 65 'Informe Motivos Baja equipos
                        NombreRPT = "rRepMotivosBaja.rpt"
                        Text1.Text = "Motivos Baja equipos"
                    
                    Case 78
                        Titulo = "Carta renovación mantenimientos"
                    
                    
                    Case 79
                        Titulo = "Etiquetas de mantenimiento"
                        NombreRPT = "rManClienEtiq.rpt"
                    
                    
                    Case 95
                        Titulo = "Ficha técnica"
                        ConSubInforme = True
                        NombreRPT = DevuelveNombreReport(37) ' "morFichaTecnica.rp"
                    Case 96
                        'COMPONENTES
                        Titulo = "Componentes"
                        ConSubInforme = False
                        NombreRPT = "morArticCompo.rpt"
                    Case Else
                        Text1.Text = "Opcion incorrecta"
                        Me.cmdImprimir.Enabled = False
                    End Select
                End If
End If
    If Titulo <> "" Then
        Text1.Text = Titulo
        Me.cmdImprimir.Enabled = True
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Function Imprime() As Boolean

    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""
    
    With frmVisReport
        .FormulaSeleccion = Me.FormulaSeleccion
        .SoloImprimir = (Me.chkSoloImprimir.Value = 0)
        .OtrosParametros = OtrosParametros
        .NumeroParametros = NumeroParametros
        .MostrarTree = MostrarTree
        .Informe = MIPATH & NombreRPT


        'Julio 2012
        If Me.NumeroDeCopias = 0 Then NumeroDeCopias = 1
        .NumCopias2 = NumeroDeCopias
        
        .ConSubInforme = ConSubInforme
        .opcion = opcion
        .ExportarPDF = (chkEmail.Value = 1)
        .Show vbModal
    End With
    
    If Me.chkEmail.Value = 1 Then
        If CadenaDesdeOtroForm <> "" Then 'se exporto el informe OK (.pdf)
            
            If Me.EnvioEMail Then  'se llamo desde envio masivo
'                frmEMail.Show vbModal
                
            Else 'informe normal, pero que se selecciono enviar e-mail
                frmEMail.opcion = 0
                frmEMail.Show vbModal
            End If
            CadenaDesdeOtroForm = ""
        End If
    End If
    
    Unload Me
  
End Function


Private Sub Form_Unload(Cancel As Integer)
    If Me.chkEmail.Value = 1 Then Me.chkSoloImprimir.Value = 1
    'If ReestableceSoloImprimir Then SoloImprimir = False
    OperacionesArchivoDefecto
    NombreSubRptConta = ""
    NumeroDeCopias = 1
    
    EstablecerImpresoraAnterior
    
End Sub

Private Sub OperacionesArchivoDefecto()
Dim crear  As Boolean
On Error GoTo ErrOperacionesArchivoDefecto

    crear = (Me.chkSoloImprimir.Value = 1)
    'crear = crear And ReestableceSoloImprimir
    If Not crear Then
        Kill App.Path & "\impre.dat"
        Else
            FileCopy App.Path & "\Vacio.dat", App.Path & "\impre.dat"
    End If
ErrOperacionesArchivoDefecto:
        If Err.Number <> 0 Then Err.Clear
End Sub




Private Sub Text1_DblClick()
    Frame2.Tag = Val(Frame2.Tag) + 1
    If Val(Frame2.Tag) > 2 Then
        Frame2.Enabled = True
        chkSoloImprimir.visible = True
    End If
End Sub

Private Sub PonerNombreImpresora()
On Error Resume Next

    If PrimeraVez Then
        ImpresoraSeleccionada = ""
        ImpresoraSeleccionada = Printer.DeviceName
    End If
    
    Label1.Caption = Printer.DeviceName
    
    If Err.Number <> 0 Then
        Label1.Caption = "No hay impresora instalada"
        
        Err.Clear
    End If
End Sub


'Private Sub CargaICO()
'    On Error Resume Next
'  '  Image1.Picture = LoadPicture(App.Path & "\iconos\printer.ico")
'    If Err.Number <> 0 Then Err.Clear
'End Sub


Private Sub EstablecerImpresoraAnterior()
Dim P As Printer
    On Error GoTo eEstablecerImpresoraAnterior
    If ImpresoraSeleccionada = "" Then Exit Sub 'por algun motivo da error
    If ImpresoraSeleccionada <> Printer.DeviceName Then
        
        For Each P In Printers
           If P.DeviceName = ImpresoraSeleccionada Then
           ' La define como predeterminada del sistema.
               Set Printer = P
               ' Sale del bucle.
               Exit For
            End If
        Next

    End If
    
Exit Sub
eEstablecerImpresoraAnterior:
    Err.Clear
End Sub



Private Sub PonerBtnSelEmpre()
    On Error Resume Next
    
    
    
    Me.cmdConfigImpre.visible = True
    If Dir(App.Path & "\SeleImpr.dat", vbArchive) <> "" Then Me.cmdConfigImpre.visible = False
    If Err.Number <> 0 Then Err.Clear
End Sub
