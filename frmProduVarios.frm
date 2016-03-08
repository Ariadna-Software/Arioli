VERSION 5.00
Begin VB.Form frmProduVarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multi form para muchas cosas de produccion"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   9300
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrCoupage 
      Height          =   1695
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5295
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   1
         Left            =   840
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   3960
         TabIndex        =   9
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdCoupage 
         Caption         =   "Hacer"
         Height          =   375
         Left            =   2640
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   480
         Picture         =   "frmProduVarios.frx":0000
         Top             =   720
         Width           =   240
      End
      Begin VB.Label lbFec 
         AutoSize        =   -1  'True
         Caption         =   "Hacer coupage"
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
         Left            =   360
         TabIndex        =   10
         Top             =   240
         Width           =   4620
      End
   End
   Begin VB.Frame FrameFiltrado 
      Height          =   3615
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   8775
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   2
         Left            =   1200
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox chkFiltrado 
         Caption         =   "Depósito 9"
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   32
         Top             =   3000
         Width           =   1815
      End
      Begin VB.CheckBox chkFiltrado 
         Caption         =   "Depósito 8"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   31
         Top             =   3000
         Width           =   1815
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   4
         Left            =   7560
         TabIndex        =   34
         Top             =   2880
         Width           =   975
      End
      Begin VB.CommandButton cmdFiltrar 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   6360
         TabIndex        =   33
         Top             =   2880
         Width           =   975
      End
      Begin VB.ComboBox cboDeposito 
         Height          =   315
         Index           =   4
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1920
         Width           =   6495
      End
      Begin VB.ComboBox cboDeposito 
         Height          =   315
         Index           =   3
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label lbFec 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Index           =   11
         Left            =   240
         TabIndex        =   39
         Top             =   960
         Width           =   495
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   840
         Picture         =   "frmProduVarios.frx":008B
         Top             =   975
         Width           =   240
      End
      Begin VB.Label lbFec 
         AutoSize        =   -1  'True
         Caption         =   "Depósitos auxiliares"
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
         Index           =   10
         Left            =   240
         TabIndex        =   38
         Top             =   2640
         Width           =   1710
      End
      Begin VB.Label lbFec 
         AutoSize        =   -1  'True
         Caption         =   "Proceso filtrado aceite"
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
         Index           =   9
         Left            =   2400
         TabIndex        =   37
         Top             =   360
         Width           =   4170
      End
      Begin VB.Label lbFec 
         AutoSize        =   -1  'True
         Caption         =   "Origen"
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
         Index           =   8
         Left            =   240
         TabIndex        =   36
         Top             =   1680
         Width           =   555
      End
      Begin VB.Label lbFec 
         AutoSize        =   -1  'True
         Caption         =   "Destino"
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
         Index           =   7
         Left            =   6840
         TabIndex        =   35
         Top             =   1680
         Width           =   645
      End
   End
   Begin VB.Frame FrameTrasiego 
      Height          =   2415
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   8775
      Begin VB.ComboBox cboDeposito 
         Height          =   315
         Index           =   1
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   960
         Width           =   1695
      End
      Begin VB.ComboBox cboDeposito 
         Height          =   315
         Index           =   0
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   960
         Width           =   6495
      End
      Begin VB.CommandButton cmdtrasiego 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   6480
         TabIndex        =   15
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   2
         Left            =   7560
         TabIndex        =   14
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label lbFec 
         AutoSize        =   -1  'True
         Caption         =   "Destino"
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
         Index           =   4
         Left            =   6840
         TabIndex        =   20
         Top             =   720
         Width           =   645
      End
      Begin VB.Label lbFec 
         AutoSize        =   -1  'True
         Caption         =   "Origen"
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
         Index           =   3
         Left            =   240
         TabIndex        =   18
         Top             =   720
         Width           =   555
      End
      Begin VB.Label lbFec 
         AutoSize        =   -1  'True
         Caption         =   "Trasiego"
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
         Index           =   2
         Left            =   3360
         TabIndex        =   16
         Top             =   240
         Width           =   1230
      End
   End
   Begin VB.Frame FrameVaciado 
      Height          =   2175
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   7695
      Begin VB.CommandButton cmdVaciadoDeposito 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   5160
         TabIndex        =   26
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   3
         Left            =   6360
         TabIndex        =   25
         Top             =   1440
         Width           =   975
      End
      Begin VB.ComboBox cboDeposito 
         Height          =   315
         Index           =   2
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   960
         Width           =   7095
      End
      Begin VB.Label lbFec 
         AutoSize        =   -1  'True
         Caption         =   "Forzar vaciado depósito"
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
         Index           =   6
         Left            =   1920
         TabIndex        =   24
         Top             =   240
         Width           =   3420
      End
      Begin VB.Label lbFec 
         AutoSize        =   -1  'True
         Caption         =   "Depósito"
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
         Index           =   5
         Left            =   240
         TabIndex        =   23
         Top             =   720
         Width           =   750
      End
   End
   Begin VB.Frame FrCierreOrdenProduccion 
      Height          =   1695
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5295
      Begin VB.TextBox txtMeses 
         Height          =   285
         Left            =   4320
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   705
         Width           =   375
      End
      Begin VB.CommandButton cmdCierreOrdProd 
         Caption         =   "Cerrar orden"
         Height          =   375
         Left            =   2880
         TabIndex        =   2
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   4200
         TabIndex        =   3
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   705
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Meses caducidad"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   12
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lbFec 
         AutoSize        =   -1  'True
         Caption         =   "Cierre orden de producción"
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
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   2280
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   840
         Picture         =   "frmProduVarios.frx":0116
         Top             =   720
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmProduVarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public opcion As Byte
    '0  .-Cierrer de una orden de produccion
    '1  .-Hacer coupage
        
    '2  .- trasiego
    '3  .- Vaciado
    '4  .- Filtrado
    
Public Intercambio As String
    '0 : codiog|fecha creacion
    '1:  codigo|fecha|almacen
    
    
'Para evitar hacer una select cad vez que lle alguna linea para el stock
Private TrabajadorConectado_ As Integer
    
    
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1


Dim cad As String  'multi proposito
Dim i As Integer

Private Sub chkFiltrado_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdCierreOrdProd_Click()
    If txtFecha(0).Text = "" Then Exit Sub
    If txtMeses.Text = "" Then
        MsgBox "Indique los meses para la fecha de caducidad", vbExclamation
        PonerFoco txtMeses
        Exit Sub
    End If
    cad = RecuperaValor(Intercambio, 2)
    If CDate(cad) > CDate(txtFecha(0).Text) Then
        cad = "Va a producir con fecha anterior a la creacion del parte de produccion." & vbCrLf & vbCrLf & "Creacion: " & cad
        cad = cad & vbCrLf & "Cierre: " & txtFecha(0).Text
        cad = cad & vbCrLf & "Caducidad. Meses: " & txtMeses.Text & "    "
        cad = cad & "EXP: " & Format(DateAdd("m", Val(txtMeses.Text), CDate(txtFecha(0).Text)), "mm/yyyy") & vbCrLf & vbCrLf & "¿Continuar?"
        If MsgBox(cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Else
        cad = "¿Seguro que desea cerrar la orden de producción " & RecuperaValor(Intercambio, 1) & " - " & RecuperaValor(Intercambio, 2)
        If MsgBox(cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    If CerrarOrdenProduccion(True) Then
        If CerrarOrdenProduccion(False) Then Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdCoupage_Click()
    If txtFecha(1).Text = "" Then Exit Sub
    cad = "¿Seguro que desea hacer el coupage " & RecuperaValor(Intercambio, 1) & " - " & RecuperaValor(Intercambio, 2)
    If MsgBox(cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    If RealizarCoupage(True) Then
        If RealizarCoupage(False) Then
            'Si ha ido bien, y el articulo es UNO de los que se tiene que actualizar el upc
            ActualizarPrecio
            '---------
            Unload Me
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub cmdFiltrar_Click()
Dim C1 As cDeposito
Dim c2 As cDeposito
Dim CC As CTiposMov
Dim FechaHora As Date

    cad = ""
    If txtFecha(2).Text = "" Then cad = "-Fecha"
    If cboDeposito(3).ListIndex < 0 Or cboDeposito(4).ListIndex < 0 Then cad = cad & "-Deposito"
    If cad <> "" Then
        MsgBox cad, vbExclamation
        Exit Sub
    End If
        
    
    For i = 0 To 1
        cad = ""
        If Me.chkFiltrado(i).Value = 1 Then
            'El deposito 8 no puede ser destino ni estar lleno
            NumRegElim = cboDeposito(3).ItemData(cboDeposito(3).ListIndex)
            If NumRegElim = 8 + i Then cad = "Deposito " & NumRegElim & " no puede ser destino ya que se utiliza como intermedio"
        End If
        
        If cad = "" Then
            If Me.chkFiltrado(i).Value = 1 Then
                Set C1 = New cDeposito
                If C1.LeerDatos(8 + i, False) Then
                    If C1.NUmlote <> "" Then cad = "Deposito intermedio  no esta vacio"
                End If
                Set C1 = Nothing
            End If
        End If
        If cad <> "" Then
            MsgBox cad, vbExclamation
            Exit Sub
        End If
    Next
    
    cad = "Va a realizar el filtrado: " & vbCrLf & "Origen: " & cboDeposito(4).Text
    cad = cad & vbCrLf & "Destino: " & cboDeposito(3).Text & vbCrLf & vbCrLf
    If Me.chkFiltrado(0).Value = 1 Then cad = cad & vbCrLf & "Deposito auxiliar 8"
    If Me.chkFiltrado(1).Value = 1 Then cad = cad & vbCrLf & "Deposito auxiliar 9"
    If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    
    cad = "select horamovi from proddepositoshco  where horamovi>=" & DBSet(txtFecha(2).Text, "F")
    'menor que el dia siguiente
    cad = cad & " AND horamovi<" & DBSet(DateAdd("d", 1, CDate(txtFecha(2).Text)), "F")
    cad = cad & " AND tipoaccion=8 order by horamovi desc"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    FechaHora = CDate(txtFecha(2).Text & " " & "07:00:00")
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux!horamovi) Then
            FechaHora = miRsAux!horamovi
            FechaHora = DateAdd("s", 1, FechaHora)
        End If
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    
    
    'Hacemos el trasiego.
    cad = ""
    Set CC = New CTiposMov
    If CC.ConseguirContador("TRO") Then
    
        Set C1 = New cDeposito
        Set c2 = New cDeposito
        
        'Un poco a mano. Vamos a ver cual es la ultima fecha hora de filtrado del dia
                
        
        If C1.LeerDatos(cboDeposito(4).ItemData(cboDeposito(4).ListIndex), False) Then
            If c2.LeerDatos(cboDeposito(3).ItemData(cboDeposito(3).ListIndex), False) Then
                C1.HacerFiltrado c2, Me.chkFiltrado(0).Value = 1, Me.chkFiltrado(1).Value = 1, CC.contador, FechaHora
                
                CC.IncrementarContador CC.TipoMovimiento
                cad = "OK"
            End If
        End If
    
    
    End If
    Set CC = Nothing
    
    Set C1 = Nothing
    Set c2 = Nothing


    If cad <> "" Then Unload Me
End Sub

Private Sub cmdtrasiego_Click()
Dim C1 As cDeposito
Dim c2 As cDeposito

    If cboDeposito(0).ListIndex < 0 Or cboDeposito(1).ListIndex < 0 Then Exit Sub
    
    
    cad = "Va a realizar el trasiego: " & vbCrLf & " Origen: " & cboDeposito(0).Text
    cad = cad & vbCrLf & "Destino: " & cboDeposito(1).Text
    If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    
    
    'Hacemos el trasiego.
    Set C1 = New cDeposito
    Set c2 = New cDeposito
    
    If C1.LeerDatos(cboDeposito(0).ItemData(cboDeposito(0).ListIndex), False) Then
        If c2.LeerDatos(cboDeposito(1).ItemData(cboDeposito(1).ListIndex), False) Then
            
            
            C1.HacerTrasiego c2
            cad = ""
                        
        End If
    End If
    
    Set C1 = Nothing
    Set c2 = Nothing
    If cad = "" Then Unload Me
    
End Sub

Private Sub cmdVaciadoDeposito_Click()
Dim C1 As cDeposito

    TrabajadorConectado_ = Val(PonerTrabajadorConectado(cad))
    If MsgBox("Continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    cad = "N"
    Set C1 = New cDeposito
    If C1.LeerDatos(cboDeposito(2).ItemData(cboDeposito(2).ListIndex), False) Then
        If C1.Kilos > 0 Then
            'DEBERIAOS REGULARIZAR
            
            RegularizarFinLote_Partida C1
        End If
        C1.QuitarAsignacionDeposito_ 2
        cad = ""
    End If
    Set C1 = Nothing
    If cad = "" Then Unload Me
End Sub

Private Sub Form_Load()
       
    Me.Icon = frmppal.Icon
    FrCierreOrdenProduccion.visible = False
    FrCoupage.visible = False
    FrameTrasiego.visible = False
    FrameVaciado.visible = False
    FrameFiltrado.visible = False
    limpiar Me
    TrabajadorConectado_ = Val(PonerTrabajadorConectado(cad))
    Select Case opcion
    Case 0
        PonerFrameVisible FrCierreOrdenProduccion
        Me.Caption = "Cierre orden producción"
        lbFec(0).Caption = lbFec(0).Caption & ": " & RecuperaValor(Intercambio, 1) & " " & RecuperaValor(Intercambio, 2)
        Me.txtMeses.Text = "18"
    Case 1
        PonerFrameVisible Me.FrCoupage
        Me.Caption = "Hacer coupage"
        lbFec(1).Caption = lbFec(1).Caption & ": " & RecuperaValor(Intercambio, 1) & " " & RecuperaValor(Intercambio, 2)
        
    Case 2
        PonerFrameVisible FrameTrasiego
        Me.Caption = "Realizar trasiego"
        CargaComobosTrasiegos 0, 1
        
    Case 3
        PonerFrameVisible FrameVaciado
        Me.Caption = "Vaciar deposito"
        CargaComobosTrasiegos 2, 2
    Case 4
        PonerFrameVisible FrameFiltrado
        Me.Caption = "Filtrado"
        CargaComobosTrasiegos 3, 4
        Me.txtFecha(2).Text = Format(Now, "dd/mm/yyyy")
    End Select
    cmdCancelar(opcion).Cancel = True
End Sub



Private Sub PonerFrameVisible(ByRef Fr As Frame)

    Fr.visible = True
    Fr.Top = 30
    Fr.Left = 30
    Me.Width = Fr.Width + 180
    Me.Height = Fr.Height + 520
    
End Sub


Private Sub frmC_Selec(vFecha As Date)
    txtFecha(i).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgFecha_Click(Index As Integer)
    'El index tiene que ser el mismo que el del txtfecha al que acompaña
    Set frmC = New frmCal
    frmC.Fecha = Now
    i = Index
    If txtFecha(Index).Text <> "" Then frmC.Fecha = CDate(txtFecha(Index).Text)
    frmC.Show vbModal
    Set frmC = Nothing
    
End Sub

Private Sub txtFecha_GotFocus(Index As Integer)
    ConseguirFoco txtFecha(Index), 3
End Sub

Private Sub txtFecha_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtFecha_LostFocus(Index As Integer)
Dim T As String
    txtFecha(Index).Text = Trim(txtFecha(Index).Text)
    If txtFecha(Index).Text <> "" Then
        T = txtFecha(Index).Text
        If EsFechaOK(T) Then
            txtFecha(Index).Text = T
        Else
            MsgBox "Fecha con formato incorrecto: " & txtFecha(Index).Text, vbExclamation
            txtFecha(Index).Text = ""
            PonerFoco txtFecha(Index)
        End If
    End If
    
End Sub

Private Function CerrarOrdenProduccion(SoloComprobar As Boolean) As Boolean
Dim vCStock As CStock
Dim B As Boolean

    'ACciones a realizar
    'Comprobar stock sublineas, ya que es la que van a disminuir la cantidad
    'Damos de alta en stock (y smoval) las lienas ppales
    'Damos de baja   "        "        las sublineas
    CerrarOrdenProduccion = False
    Set miRsAux = New ADODB.Recordset
    Set vCStock = New CStock
    
    'Veamos las sub lineas  si tienen stock. Antes comprobabamos cantidad x sarti1.cntidad
    'Cad = "select codarti1,codalmac,sarti1.cantidad multiplicador,sum(sliordpr.cantidad) cantilinea from sliordpr,sarti1 where "
    'Cad = Cad & " sliordpr.codartic=sarti1.codartic and  codigo=" & RecuperaValor(Me.Intercambio, 1) & " group by 1,2,3"
    'AHora hay una tabla para los componentes
'    Cad = "select codarti2,sliordpr.codalmac,sliordpr2.cantidad cantilinea from sliordpr,sliordpr2 where"
'    Cad = Cad & " sliordpr.codartic=sliordpr2.codartic and sliordpr.codalmac=sliordpr2.codalmac and"
'    Cad = Cad & " sliordpr.codigo=1 group by 1,2"
'
    cad = "select sliordpr2.*,sartic.factorconversion from sliordpr2,sartic where sliordpr2.codarti2=sartic.codartic and codigo=" & RecuperaValor(Me.Intercambio, 1)
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    B = False
    
    If Not SoloComprobar Then Conn.BeginTrans

    
    While Not miRsAux.EOF

        B = False
        If InicializarCStock(vCStock, "S", True) Then
            
            If vCStock.MueveStock Then
                If SoloComprobar Then
                    B = vCStock.MoverStock(False)
                Else
                    'Estamos ejecutando la actualizacion
                    '---------------------------------------------
                    'si hay control de stock para el articulo actualizar en salmac e insertar en smoval
                    'en actualizar stock comprobamos si el articulo tiene control de stock
                    B = vCStock.ActualizarStock(False)
                End If
            Else
                B = True
            End If
        End If
                             
        
        If Not B Then
            While Not miRsAux.EOF
                miRsAux.MoveNext  'para que no siga
            Wend
        Else
            'Al siguiente
            miRsAux.MoveNext
        End If
    Wend
    miRsAux.Close
  
    
    If Not B Then
        Set miRsAux = Nothing
        Set vCStock = Nothing
        If Not SoloComprobar Then Conn.RollbackTrans
        Exit Function 'Si no puede inicializar los stocks, de las sublineas salimos
    End If
    
    
    
    
    'SSi ha ido bien comprobamos los LOTES
    If Not RealizarProduccionLOTES(SoloComprobar) Then
    
            Set miRsAux = Nothing
            Set vCStock = Nothing
            If Not SoloComprobar Then Conn.RollbackTrans
            Exit Function 'Si no puede inicializar los stocks, de las sublineas salimos
    End If
    Set miRsAux = Nothing
    Set miRsAux = New ADODB.Recordset
    
    'AHora comprobamos los stcosk de las entraddas , de las lineas            factor=1
    cad = "select codartic codarti2,codalmac,sum(sliordpr.cantidad) cantidad,1 factorconversion,numlote from sliordpr where "
    cad = cad & " codigo=" & RecuperaValor(Me.Intercambio, 1) & " group by 1,2"
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    B = False
    While Not miRsAux.EOF
        B = False
        If InicializarCStock(vCStock, "E", False) Then   'Las lineas son de netrada
            If vCStock.MueveStock Then
                If SoloComprobar Then
                   ' B = vCStock.MoverStock(False, True)
                   B = True
                Else
                    B = vCStock.ActualizarStock(False)
                End If
            Else
                B = True
            End If
        End If
        
        If Not B Then
            While Not miRsAux.EOF
                miRsAux.MoveNext  'para que no siga
            Wend
        Else
            If Not SoloComprobar Then
                '-------------------------- LOTES
                
                'Si ha puesto numero de lote
                
            End If
            'Al siguiente
            miRsAux.MoveNext
        End If
    Wend
    miRsAux.Close
    
    
    If Not B Then
        Set miRsAux = Nothing
        Set vCStock = Nothing
        If Not SoloComprobar Then Conn.RollbackTrans
        Exit Function 'Si no puede inicializar los stocks, de las sublineas salimos
    End If
    
    
    
    
    'Acutailizaremos algnas cosas como la fecha de baja
    If Not SoloComprobar Then
        Conn.CommitTrans
        cad = "UPDATE sordprod set fecproduccion = " & DBSet(txtFecha(0).Text, "F")
        'Marzo 2012. Caducidad
        cad = cad & ",feccaduca  = " & DBSet(DateAdd("m", Val(txtMeses.Text), CDate(txtFecha(0).Text)), "F")
        cad = cad & " WHERE  codigo=" & RecuperaValor(Me.Intercambio, 1)
        Conn.Execute cad
    End If
    
    CerrarOrdenProduccion = True
    
    Set miRsAux = Nothing
    Set vCStock = Nothing
    
    
End Function






'-------------------------------------------------------------------------
' Realizar el coupage
'
Private Function RealizarCoupage(SoloComprobar As Boolean) As Boolean
Dim vCStock As CStock
Dim B As Boolean
Dim CantidadTotalAProducir As Currency 'Cuatro decimales

    
    'ACciones a realizar
    'Comprobar stock sublineas, ya que es la que van a disminuir la cantidad
    'Damos de alta en stock (y smoval) las lienas ppales
    'Damos de baja   "        "        las sublineas
    RealizarCoupage = False
    Set miRsAux = New ADODB.Recordset
    Set vCStock = New CStock
    
    
    If Not SoloComprobar Then Conn.BeginTrans

    
    'Los mezclantes
    'Como no lleva factor conversion. Necesito los precios para los calculos de importes
    cad = "select olicoupagelin.*,preciouc, preciomp from olicoupagelin,sartic where olicoupagelin.codartic=sartic.codartic and "
    cad = cad & "  codigo = " & RecuperaValor(Me.Intercambio, 1)
    'cad = "select * from olicoupagelin where codigo = " & RecuperaValor(Me.Intercambio, 1)
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    B = False
    
    CantidadTotalAProducir = 0
    While Not miRsAux.EOF
        B = False
        If InicializarCStockCoupage(vCStock, "S", False) Then
            
            If vCStock.MueveStock Then
                If SoloComprobar Then
                    B = vCStock.MoverStock(False)
                Else
                    'Estamos ejecutando la actualizacion
                    '---------------------------------------------
                    'si hay control de stock para el articulo actualizar en salmac e insertar en smoval
                    'en actualizar stock comprobamos si el articulo tiene control de stock
                    B = vCStock.ActualizarStock(False)
                End If
            Else
                B = True
            End If
            CantidadTotalAProducir = CantidadTotalAProducir + miRsAux!Kilos
        End If
        
        If Not B Then
            While Not miRsAux.EOF
                miRsAux.MoveNext  'para que no siga
            Wend
        Else
            'Al siguiente
            miRsAux.MoveNext
        End If
    Wend
    miRsAux.Close
    
    
    If Not B Then
        Set miRsAux = Nothing
        Set vCStock = Nothing
        If Not SoloComprobar Then Conn.RollbackTrans
        Exit Function 'Si no puede inicializar los stocks, de las sublineas salimos
    End If
    
    
    'SSi ha ido bien comprobamos los LOTES
    If Not RealizarCoupageLOTES(SoloComprobar, CantidadTotalAProducir) Then
    
            Set miRsAux = Nothing
            Set vCStock = Nothing
            If Not SoloComprobar Then Conn.RollbackTrans
            Exit Function 'Si no puede inicializar los stocks, de las sublineas salimos
    End If
    Set miRsAux = Nothing
    Set miRsAux = New ADODB.Recordset
    
    
    
    
    
    
    'AHora comprobamos los stcosk de las entraddas , de las lineas
    cad = TransformaComasPuntos(CStr(CantidadTotalAProducir))
    
    cad = "select olicoupage.codartic," & cad & " kilos,preciouc from olicoupage,sartic where"
    cad = cad & " olicoupage.codartic=sartic.codartic and codigo=" & RecuperaValor(Me.Intercambio, 1)
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    B = False
    
    While Not miRsAux.EOF
        B = False
        If InicializarCStockCoupage(vCStock, "E", False) Then   'Las lineas son de netrada
        
            If vCStock.MueveStock Then
                If SoloComprobar Then
                    'B = vCStock.MoverStock(False)
                    B = True
                    
                    
                    
                    
                Else
                    B = vCStock.ActualizarStock(False)
                End If
            Else
                B = True
            End If
        End If
        
        If Not B Then
            While Not miRsAux.EOF
                miRsAux.MoveNext  'para que no siga
            Wend
        Else
            'Al siguiente
            miRsAux.MoveNext
        End If
    Wend
    miRsAux.Close
    
    
    If Not B Then
        Set miRsAux = Nothing
        Set vCStock = Nothing
        If Not SoloComprobar Then Conn.RollbackTrans
        Exit Function 'Si no puede inicializar los stocks, de las sublineas salimos
    End If
    
    
    
    
    'Acutailizaremos algnas cosas como la fecha de baja
    If Not SoloComprobar Then
        Conn.CommitTrans
        cad = "UPDATE olicoupage set YaCreado = 1"
        cad = cad & " WHERE  codigo=" & RecuperaValor(Me.Intercambio, 1)
        Conn.Execute cad
    End If
    
    
        
    
    
    
    RealizarCoupage = True
    
    Set miRsAux = Nothing
    Set vCStock = Nothing
    
    
End Function






'No le paso el recodset pq es mirsaux que es comun
Private Function InicializarCStock(ByRef vCStock As CStock, TipoM As String, Sublineas As Boolean) As Boolean
Dim CantidadNecesaria As Single
Dim MateriaPrima As Boolean
    On Error Resume Next

    vCStock.tipoMov = TipoM
    vCStock.DetaMov = "PRO"
    vCStock.Trabajador = TrabajadorConectado_
    vCStock.Documento = RecuperaValor(Intercambio, 1)
    vCStock.Fechamov = txtFecha(0).Text '
    vCStock.codAlmac = CInt(miRsAux!codAlmac)
    
    CantidadNecesaria = miRsAux!FactorConversion
    MateriaPrima = False
    If CantidadNecesaria <> 1 Then MateriaPrima = True
    
    'mAYO 2010.   eL FACTOR CONVERSION VIENE ya grabado en sliorpr2
    '           quiero decir que no hay que volver a multiplcarlo
    'If CantidadNecesaria <> 1 Then Stop
    CantidadNecesaria = 1  'YA hemos grabado la sliordpr
    
    If Sublineas Then
        If vCStock.codAlmac = 2 And Not MateriaPrima Then
            'Es el del B
            'Solo el aceite vendra de las garrafas de B. Lo demas todo del limpio
             vCStock.codAlmac = 1
        End If
    End If
    vCStock.codartic = miRsAux!codarti2
    
   
    If CantidadNecesaria = 0 Then CantidadNecesaria = 1 'PARA QUE NO DE ERROR
    CantidadNecesaria = Round2(miRsAux!Cantidad * CantidadNecesaria, 5)
    vCStock.Cantidad = CantidadNecesaria
    vCStock.Importe = 0
    vCStock.LineaDocu = 0

    If Err.Number <> 0 Then
        MsgBox "No se han podido inicializar la clase para actualizar Stock:" & Err.Description, vbExclamation
        InicializarCStock = False
    Else
        InicializarCStock = True
    End If
End Function

Private Function InicializarCStockCoupage(ByRef vCStock As CStock, TipoM As String, ParaLotes As Boolean) As Boolean
'Dim CantidadNecesaria As Single   'No lleva factor conversion, ya que esta en KILOS que es el stcok
Dim Impor As Currency
    On Error Resume Next

    vCStock.tipoMov = TipoM
    vCStock.DetaMov = "CUP"  'Coupages
   ' vCStock.Trabajador = PonerTrabajadorConectado(cad)
    vCStock.Trabajador = TrabajadorConectado_
    vCStock.Documento = RecuperaValor(Intercambio, 1)
    vCStock.Fechamov = txtFecha(1).Text '
    
   
    vCStock.codartic = miRsAux!codartic
    vCStock.codAlmac = RecuperaValor(Intercambio, 3)
'    CantidadNecesaria = miRsAux!FactorConversion
'    If CantidadNecesaria = 0 Then CantidadNecesaria = 1 'PARA QUE NO DE ERROR
'    CantidadNecesaria = Round2(miRsAux!kilos / CantidadNecesaria, 5)
'    vCStock.Cantidad = CantidadNecesaria
    vCStock.Cantidad = miRsAux!Kilos
    If Not ParaLotes Then
        Impor = DBLet(miRsAux!precioUC, "N")
        Impor = Round2(Impor * vCStock.Cantidad, 4)
        vCStock.Importe = Impor
    Else
        vCStock.Importe = 0
    End If
    vCStock.LineaDocu = 0

    If Err.Number <> 0 Then
        MsgBox "No se han podido inicializar la clase para actualizar Stock", vbExclamation
        InicializarCStockCoupage = False
    Else
        InicializarCStockCoupage = True
    End If
End Function





Private Function RealizarProduccionLOTES(SoloComprobar As Boolean) As Boolean
Dim ErroresEnPartidas As String
Dim LotesNecesartios As Collection
Dim CantidadNecesaria As Currency
Dim AuxPartida As String
Dim Err_x_Articulo As String
Dim MiNumeroLote As String
Dim cP As cPartidas   'Para los numeros de lote
Dim Rc As Byte
Dim vvCstock As CStock
Dim B As Boolean
Dim RL As ADODB.Recordset
Dim CantidadQueLLevo As Currency
Dim Aux As String
Dim TieneLotesMP As Boolean
Dim II As Integer
Dim Cant2 As Currency
Dim CL As cLotaje
Dim LoteReal As String  'Con fecha
    On Error GoTo ERealizarProduccionLOTES

    RealizarProduccionLOTES = False
    ErroresEnPartidas = ""
    AuxPartida = ""
    Set cP = New cPartidas

    If Not SoloComprobar Then

        Set CL = New cLotaje
        CL.DetaMov = "PRO"
        CL.Documento = RecuperaValor(Intercambio, 1)
        CL.Fechamov = CDate(Me.txtFecha(0).Text)
        CL.HoraMov = CDate(Me.txtFecha(0).Text & " " & Format(Now, "hh:nn:ss"))
        CL.ProvCliTra = TrabajadorConectado_
        CL.LineaDocu = 0
        CL.SubLinea = 0
    End If
        


    cad = "select sliordpr2.*,sartic.factorconversion,trazabilidad,nomartic from sliordpr2,sartic where "
    cad = cad & " sliordpr2.codarti2=sartic.codartic and codigo=" & RecuperaValor(Me.Intercambio, 1)
    cad = cad & " AND trazabilidad = 1" 'Solo miraremos los que lleven trazabilidad
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    B = False
    


    AuxPartida = ""
    Set vvCstock = New CStock
    While Not miRsAux.EOF
        If Err_x_Articulo <> miRsAux!codartic Then
            'Han habido errores en el articulo anterior.
            If AuxPartida <> "" Then
                cad = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", DevNombreSQL(Err_x_Articulo), "T")
                AuxPartida = "-  " & Err_x_Articulo & "  " & cad & AuxPartida & vbCrLf
                ErroresEnPartidas = ErroresEnPartidas & AuxPartida & vbCrLf
            End If
            Err_x_Articulo = miRsAux!codartic
            AuxPartida = ""
        End If

        B = False
        If InicializarCStock(vvCstock, "E", False) Then   'Las lineas son de netrada
  
            CantidadNecesaria = vvCstock.Cantidad  'Para tenerla despues en los lotes

            '// NUmeros de LOTE
            ' Las materias primas (en ppio solo ellas) pueden forzarle los lotes
            'en el mantenimiento de produccion. Con lo cual, si se lo han asignado comprobare
            'que de lo que le asignan tengo disponible. Si no se lo asigno YO
            Set LotesNecesartios = New Collection
            'De momento solo para las MATERIAS PRIMAS
            ' factorconversion<>1
            If miRsAux!FactorConversion = 1 Then
                TieneLotesMP = False
            Else
                Aux = "Select * from sliordpr2lotes WHERE  codigo = " & RecuperaValor(Intercambio, 1)
                Aux = Aux & " AND codalmac =" & vvCstock.codAlmac & " AND codArtic = " & DBSet(miRsAux!codartic, "T")
                Aux = Aux & " AND codArti2 = " & DBSet(vvCstock.codartic, "T")
                Set RL = New ADODB.Recordset
                RL.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If RL.EOF Then
                    TieneLotesMP = False
                Else
                    TieneLotesMP = True
                    CantidadQueLLevo = CantidadNecesaria
                    'Para cada lote especficiado veremos SI existe el lote o no en partidas
                    While Not RL.EOF
                        'ANTES MAYO 2010
                        'Cant2 = Round2(miRsAux!FactorConversion * RL!cantlote, 5)
                        'AHORA. Mayo 2010.  YA he grabado la sliord2 con el factor conversion multimplicado NO debo volver a miultiplicarlo
                        Cant2 = RL!cantlote
                        
                        
                        CantidadQueLLevo = CantidadQueLLevo - Cant2
                    
                        Aux = RL!NUmlote & "|"
                        RL.MoveNext
                        If RL.EOF Then
                            'Es la utlima. Ajusto los decimales
                            If CantidadQueLLevo > 0 Then Cant2 = Cant2 + CantidadQueLLevo
                        End If
                        Aux = Aux & Cant2 & "|"
                        LotesNecesartios.Add Aux
                        
                        
                        
                    Wend
                    RL.Close
                    CantidadQueLLevo = 0
                End If
            End If
            
            If TieneLotesMP Then
                
                'Los busco en partidas
                For II = 1 To LotesNecesartios.Count
                    Aux = LotesNecesartios(II)
                    Cant2 = CCur(RecuperaValor(Aux, 2))
                    Aux = RecuperaValor(Aux, 1)
                    Aux = "  AND numlote = '" & DevNombreSQL(Aux) & "'"
                    Aux = " AND codalmac =" & vvCstock.codAlmac & Aux
                    Aux = " where codartic = " & DBSet(vvCstock.codartic, "T") & Aux
                    Aux = "Select id,cantotal from spartidas " & Aux
                    RL.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    
                    If RL.EOF Then
                        'NO existe el registro en partidas para ese LOTE - articulo
                        cad = "NO existe LOTE: " & RecuperaValor(LotesNecesartios(II), 1)
                        If Not SoloComprobar Then
                            'FALTA###
                           
                            cP.Cantidad = -1 * CantidadNecesaria
                            cP.codAlmac = vvCstock.codAlmac
                            cP.codartic = vvCstock.codartic
                            cP.codProve = 0
                            cP.Fecha = vvCstock.Fechamov
                            
                            cP.NumAlbar = "PR" & miRsAux!Codigo
                            cP.NUmlote = cP.NumAlbar
                            cP.Insertar



                            InsertarMovientosLotesProduccion CL, cP, cP.Cantidad, miRsAux!codartic

                        End If
                        
                    Else
                        'SI que existe el LOTE veamos si tiene suficiente
                        If RL!cantotal < Cant2 Then
                            'No tengo suficiente
                            'FALTA
                            cad = "No tengo suficiente. (" & LotesNecesartios(II) & ")"

    
                        Else
                            'Todo OK
                            cad = ""
                            
                        End If
                        'Si estamos ya realizando la produccion actualizamos tablas
                        
                        If Not SoloComprobar Then
                            
                            cP.Leer Val(RL!ID)
                            cP.IncrementarCantidad -1 * Cant2
                            
                            InsertarMovientosLotesProduccion CL, cP, -1 * Cant2, miRsAux!codartic
                        End If
                    End If
                    RL.Close
                    If SoloComprobar Then
                        If cad <> "" Then
                                cad = Space(19) & "-- " & vvCstock.codartic & "  " & Mid(miRsAux!NomArtic & Space(45), 1, 45) & cad
                                AuxPartida = AuxPartida & vbCrLf & cad
                        End If
                    
                    End If
                Next   'LotesNecesartios.Count
            
            Else
                
                'Asi es como estaba antes
                Rc = cP.RecuperarLotes(vvCstock.codartic, vvCstock.codAlmac, CantidadNecesaria, LotesNecesartios)
            
                If Rc = 2 Then
                    'No tengo el articulo dado de alta
                    cad = "NO hay ningun lote "
                    
                    'Si estoyNO es solo comprobar, entonces NO dejo que continue en este caso
                    If Not SoloComprobar Then
                        'Realmente deberia salir
                      
                        
                        'FALTA####
                        'Deberian existir. Como No existe lo damos de alta
                        
                        cP.Cantidad = -1 * CantidadNecesaria
                        cP.codAlmac = vvCstock.codAlmac
                        cP.codartic = vvCstock.codartic
                        cP.codProve = 0
                        cP.Fecha = vvCstock.Fechamov
                        
                        cP.NumAlbar = "PR" & miRsAux!Codigo
                        cP.NUmlote = cP.NumAlbar
                        If cP.Insertar Then
                            B = True
                            Insertar_sliordpr2lotes cP, 1, CantidadNecesaria
                        End If
                        InsertarMovientosLotesProduccion CL, cP, cP.Cantidad, miRsAux!codartic
                        
                        
                    End If
                ElseIf Rc = 1 Then
                
                    cad = "NO hay suficiente cantidad"
                    
                    If Not SoloComprobar Then
                        
                        cP.IncrementarCantidad -1 * CantidadNecesaria
                        Insertar_sliordpr2lotes cP, 1, CantidadNecesaria
                        InsertarMovientosLotesProduccion CL, cP, -1 * CantidadNecesaria, miRsAux!codartic
                    End If
                Else
                    'Ahora si
                    cad = ""
                    B = True
                    
                End If
                If SoloComprobar Then
                        If cad <> "" Then
                            cad = Space(19) & "-- " & vvCstock.codartic & "  " & Mid(miRsAux!NomArtic & Space(45), 1, 45) & cad
                            AuxPartida = AuxPartida & vbCrLf & cad
                        End If
                
                Else
                    'Estamos ejecutando
                    If B Then
                      For i = 1 To LotesNecesartios.Count
                            cad = LotesNecesartios(i)
                            
                            'ACciones a realizar. Disminnuir cantidad en LOTES
                            NumRegElim = RecuperaValor(cad, 1)
                            CantidadNecesaria = CCur(RecuperaValor(cad, 2))
                            
                            If Not cP.Leer(NumRegElim) Then
                                'MAAAAAAl
                                MsgBox "Error grave partidas/lotes: " & NumRegElim, vbExclamation
                            Else
                                CantidadNecesaria = -1 * CantidadNecesaria
                                cP.IncrementarCantidad CantidadNecesaria
                            
                            
                                'ACtualizar la fila con el numero de lote asignado
                                Insertar_sliordpr2lotes cP, i, Abs(CantidadNecesaria)
                                
                                InsertarMovientosLotesProduccion CL, cP, CantidadNecesaria, miRsAux!codartic
                                
                                
                            End If  'de cp.leer
                        Next
                    End If  'De B
                End If 'Solo comprobar
            End If  'Tiene lotes MP
            


            
            Set LotesNecesartios = Nothing
        End If 'DE incializa stock
        miRsAux.MoveNext
    Wend
    miRsAux.Close
        
    If AuxPartida <> "" Then
        cad = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", DevNombreSQL(Err_x_Articulo), "T")
        AuxPartida = "-  " & Err_x_Articulo & "   " & cad & AuxPartida
        ErroresEnPartidas = ErroresEnPartidas & AuxPartida
    End If

    If ErroresEnPartidas <> "" Then ErroresEnPartidas = "Error en numeros de lote. " & vbCrLf & String(75, "=") & vbCrLf & ErroresEnPartidas


    AuxPartida = ""
    
        
    cad = "select codartic codarti2,codalmac,sum(sliordpr.cantidad) cantidad,1 factorconversion,numlote from sliordpr where "
    cad = cad & " codigo=" & RecuperaValor(Me.Intercambio, 1) & " group by 1,2"
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not miRsAux.EOF
        B = False
            If InicializarCStock(vvCstock, "E", False) Then   'Las lineas son de netrada
                
                    'AHora veremos los numeros de lote
                    'EL nUMERO DE LOTE NO puede ser NULO
                    CantidadNecesaria = vvCstock.Cantidad  'Para tenerla despues en los lotes
                    cad = "select codalmac,codartic,numlote,cantlote from sliordprlotes where "
                    cad = cad & " codigo=" & RecuperaValor(Me.Intercambio, 1)
                    cad = cad & " AND codartic= '" & miRsAux!codarti2 & "'"
                    
                    CantidadQueLLevo = 0
                    Set RL = New ADODB.Recordset
                    RL.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    While Not RL.EOF
                        CantidadQueLLevo = CantidadQueLLevo + RL!cantlote
                        If Not SoloComprobar Then
                                Set cP = New cPartidas
                                'Vemos si ya existe
                                LoteReal = RL!NUmlote & " " & Format(txtFecha(0).Text, "yyyy/mm/dd")
                                If cP.LeerDesdeArticulo(miRsAux!codarti2, miRsAux!codAlmac, LoteReal) Then
                                    'Ya existia(por algun motivo)
                                    cP.IncrementarCantidad RL!cantlote
                                    
                                Else
                                    cP.Cantidad = RL!cantlote
                                    cP.codAlmac = vvCstock.codAlmac
                                    cP.codartic = vvCstock.codartic
                                    cP.codProve = 0
                                    cP.Fecha = CDate(txtFecha(0).Text)
                                    cP.NumAlbar = "PR" & RecuperaValor(Me.Intercambio, 1)
                                    cP.NUmlote = LoteReal
                                    If Not cP.Insertar Then
                                        cad = "Error insertando partidas/lotes: " & cP.codartic
                                        MsgBox cad, vbExclamation
                                    End If
                                    
                                End If
                                
                                'En movimientos lote
                                CL.tipoMov = 1
                                CL.Cantidad = cP.Cantidad
                                CL.codAlmac = cP.codAlmac
                                CL.codartic = cP.codartic
                                CL.codarti2 = ""
                                CL.NUmlote = cP.NUmlote
                                If Not CL.InsertarLote Then Err.Raise vbObjectError + 513, , "Error insertando en mov lotes: " & cP.codartic
                                Set cP = Nothing
                                
                                
                                'MAYO 2010
                                'UPDATEO el LOTE que antes era de 4 digitos
                                'a otro que sera los 4 mas la fecha
                                cad = "UPDATE sliordprlotes set numlote=" & DBSet(LoteReal, "T")
                                cad = cad & " where codigo=" & RecuperaValor(Me.Intercambio, 1)
                                cad = cad & " AND codartic= '" & miRsAux!codarti2 & "'"
                                cad = cad & " AND numlote= '" & RL!NUmlote & "'"
                                Conn.Execute cad
                        End If
                        RL.MoveNext
                   Wend
                   RL.Close
                   If CantidadQueLLevo <> CantidadNecesaria Then
                        If Not SoloComprobar Then AuxPartida = AuxPartida & vvCstock.codartic & ":   necesaria/lotes: " & Format(CantidadNecesaria, FormatoCantidad) & " / " & Format(CantidadQueLLevo, FormatoCantidad) & vbCrLf
                   End If
            End If 'Ini stock
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        If AuxPartida <> "" Then   'Si han habido errores en comprobar cantidades lotes los añado
            AuxPartida = vbCrLf & vbCrLf & "Articulos producidos: " & vbCrLf & AuxPartida
            ErroresEnPartidas = ErroresEnPartidas & AuxPartida
        End If
        B = True
        
        If SoloComprobar Then
            If ErroresEnPartidas <> "" Then
                ErroresEnPartidas = ErroresEnPartidas & vbCrLf & vbCrLf & "¿Continuar?"
                If MsgBox(ErroresEnPartidas, vbQuestion + vbYesNo) = vbNo Then B = False
            End If
        End If
    
        RealizarProduccionLOTES = B


    
ERealizarProduccionLOTES:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set RL = Nothing
    Set miRsAux = Nothing
    Set vvCstock = Nothing
 
End Function


Private Sub InsertarMovientosLotesProduccion(ByRef cLot As cLotaje, cPar As cPartidas, Cantidad As Currency, ArticuloProduccion As String)

    
    
    cLot.tipoMov = 0  'Salida
    cLot.Cantidad = Abs(Cantidad)
    cLot.codAlmac = cPar.codAlmac
    cLot.codartic = cPar.codartic
    cLot.codarti2 = ArticuloProduccion
    cLot.NUmlote = cPar.NUmlote

    If Not cLot.InsertarLote Then Err.Raise vbObjectError + 513, , "Error insertando en mov lotes: " & cPar.codartic
    
End Sub


Private Sub Insertar_sliordpr2lotes(ByRef Par As cPartidas, LineaLote As Integer, Cantidad As Currency)
Dim SQL As String

    
    SQL = "insert into sliordpr2lotes (`codigo`,`codalmac`,`codartic`,`codarti2`,"
    SQL = SQL & "`linea`,`numlote`,`cantlote`) values ( "

    SQL = SQL & RecuperaValor(Intercambio, 1) & ","
    'En misraux tengo los datos que necesito
    SQL = SQL & miRsAux!codAlmac & ",'" & miRsAux!codartic & "','" & miRsAux!codarti2 & "',"
    SQL = SQL & LineaLote & ",'" & DevNombreSQL(Par.NUmlote) & "'," & TransformaComasPuntos(CStr(Cantidad)) & ")"
    EjecutaSQL conAri, SQL, True
    
End Sub






'------------------------  LOTES COUPAGE
Private Function RealizarCoupageLOTES(SoloComprobar As Boolean, CantidadMezcla As Currency) As Boolean
Dim ErroresEnPartidas As String
Dim CantidadNecesaria As Currency
Dim AuxPartida As String
Dim Err_x_Articulo As String
Dim MiNumeroLote As String
Dim cP As cPartidas   'Para los numeros de lote
Dim Rc As Byte
Dim vvCstock As CStock
Dim B As Boolean
'Si lleva marca de fin depoisto
Dim RegularizacionDeposito As Currency
Dim cDEP As cDeposito

Dim T1 As Single

Dim CantidadQueLLevo As Currency
Dim CL As cLotaje

    On Error GoTo ERealizarCUPLOTES

    RealizarCoupageLOTES = False
    

    If Not SoloComprobar Then

        Set CL = New cLotaje
        CL.DetaMov = "CUP"
        CL.Documento = RecuperaValor(Intercambio, 1)
        CL.Fechamov = CDate(Me.txtFecha(1).Text)
        CL.HoraMov = CDate(Me.txtFecha(1).Text & " " & Format(Now, "hh:nn:ss"))
        CL.ProvCliTra = TrabajadorConectado_
        CL.LineaDocu = 0
        CL.SubLinea = 0
    End If
    'Por si acaso no ha puesto numero de lotes. DEBERIA HABERLOS PUESTO
    cad = "select olicoupagelin.codartic,kilos,olicoupagelinlotes.codartic artlote,numlote,cantlote"
    'Juni 2014
    cad = cad & " ,fincuba,deposito"
    cad = cad & " FROM olicoupagelin left join olicoupagelinlotes on"
    cad = cad & " olicoupagelin.codArtic = olicoupagelinlotes.codArtic"
    cad = cad & " and olicoupagelin.codigo= olicoupagelinlotes.codigo WHERE  olicoupagelin.codigo ="
    cad = cad & RecuperaValor(Me.Intercambio, 1) & " ORDER BY codartic"
    miRsAux.Open cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    B = False
    cad = ""
    ErroresEnPartidas = ""
    'Comprobaremos que todos traen el numero de lote puesto y que los
    While Not miRsAux.EOF
        If IsNull(miRsAux!artlote) Then
            AuxPartida = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", miRsAux!codartic, "T")
            cad = cad & miRsAux!codartic & "   " & AuxPartida
        Else
            If MiNumeroLote <> miRsAux!codartic Then
                If MiNumeroLote <> "" Then
                    If CantidadQueLLevo <> CantidadNecesaria Then
                        AuxPartida = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", MiNumeroLote, "T")
                        ErroresEnPartidas = ErroresEnPartidas & MiNumeroLote & "   " & AuxPartida & vbCrLf
                    End If
                End If
                MiNumeroLote = miRsAux!codartic
                CantidadNecesaria = miRsAux!Kilos
                CantidadQueLLevo = miRsAux!cantlote
            Else
                'Dos lineas del mismo articulo
                CantidadQueLLevo = CantidadQueLLevo + miRsAux!cantlote
            End If
        End If
        miRsAux.MoveNext
        
        
        
        
    Wend
    
    
    'La utlima linea
    If MiNumeroLote <> "" Then
        If CantidadQueLLevo <> CantidadNecesaria Then
            AuxPartida = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", MiNumeroLote, "T")
            ErroresEnPartidas = ErroresEnPartidas & MiNumeroLote & "   " & AuxPartida & vbCrLf
        End If
    End If
    
    If cad <> "" Or ErroresEnPartidas <> "" Then
        If cad <> "" Then cad = "Lineas articulo sin indicar numero de lote: " & vbCrLf & String(60, "-") & vbCrLf & cad
        If ErroresEnPartidas <> "" Then cad = cad & vbCrLf & vbCrLf & "Articulos lineas sin coincidir cantidades lotes: " & vbCrLf & String(70, "-") & vbCrLf & ErroresEnPartidas
        miRsAux.Close
        MsgBox cad, vbExclamation
        Exit Function
    End If
        
    miRsAux.MoveFirst
    MiNumeroLote = ""
    AuxPartida = ""
    ErroresEnPartidas = ""
    Set cP = New cPartidas
    Set vvCstock = New CStock
    Set cDEP = New cDeposito
    
    While Not miRsAux.EOF
        If Err_x_Articulo <> miRsAux!codartic Then
            'Han habido errores en el articulo anterior.
            If AuxPartida <> "" Then
                AuxPartida = "-  " & Err_x_Articulo & vbCrLf & AuxPartida & vbCrLf
                ErroresEnPartidas = ErroresEnPartidas & AuxPartida & vbCrLf
            End If
            Err_x_Articulo = miRsAux!codartic
            AuxPartida = ""
        End If

        RegularizacionDeposito = 0
        B = False
        If InicializarCStockCoupage(vvCstock, "E", True) Then    'Las lineas son de netrada
    
            CantidadNecesaria = CCur(miRsAux!cantlote)
            B = True
            '// NUmeros de LOTE
            cad = ""
            If cP.LeerDesdeArticulo(vvCstock.codartic, vvCstock.codAlmac, miRsAux!NUmlote) Then
            
                If cP.Cantidad >= CantidadNecesaria Then
                    'PERFECTO. NO HAgo nada
                    If Val(miRsAux!fincuba) = 1 Then
                        'Regulzarizaremos el deposito
                        RegularizacionDeposito = cP.Cantidad - CantidadNecesaria
                    End If
                Else
                    If Val(miRsAux!fincuba) = 0 Then
                        'No es fin deposito, no puede seguir
                        cad = "NO hay suficiente cantidad"
                    Else
                        'OK, es fin deposito y habria que "REGULARIZARLO"
                        ' es decir meter una linea para dejar la cantidad del deposito a cero,
                        ' LA PARTIDA a cero
                        ' y una vez acabado el proceso dejar el deposito preparado para llenarlo de nuevo
                        RegularizacionDeposito = cP.Cantidad - CantidadNecesaria
                    End If
                     
                End If
            Else
                'NO existe lote. De momento dejo continuar
                B = False
                cad = "NO hay ningun lote "
                
            End If
    
        
            If SoloComprobar Then
                If cad <> "" Then
                    cad = cad & " (" & miRsAux!NUmlote & ")"
                    cad = Space(15) & "-- " & vvCstock.codartic & "  " & cad
                    AuxPartida = AuxPartida & vbCrLf & cad
                End If
            
            Else
                'Por si acaso es FIN deposito
                RegularizacionDeposito = cP.Cantidad - CantidadNecesaria
            
                CantidadNecesaria = -1 * CantidadNecesaria  'En negativo
                
                'Incrementamos los kilos
                cDEP.LeerDatos miRsAux!Deposito, False
                cDEP.VariacionKilosDeposito CantidadNecesaria
                
                
                
                If Not B Then
                    'NO existe. Lo creo
                    cP.Cantidad = CantidadNecesaria
                    cP.codAlmac = vvCstock.codAlmac
                    cP.codartic = vvCstock.codartic
                    cP.codProve = 0
                    cP.Fecha = CDate(txtFecha(1).Text)
                    cP.NumAlbar = "CUP" & RecuperaValor(Me.Intercambio, 1)
                    cP.NUmlote = DBLet(miRsAux!NUmlote, "T")
                    If cP.NUmlote Then cP.NUmlote = cP.NumAlbar
                    
                    If Not cP.Insertar Then
                        cad = "Error insertando partidas/lotes: " & cP.codartic
                        Err.Raise vbObjectError + 513, , cad
                    End If
        
                Else
                    'Si existe. Lo decremento
                    cP.IncrementarCantidad CantidadNecesaria
                                    
                End If
                'Insertamos en la linea de smoval
                CL.tipoMov = 0
                CL.Cantidad = Abs(CantidadNecesaria)
                CL.codAlmac = vvCstock.codAlmac
                CL.codartic = vvCstock.codartic
                CL.NUmlote = cP.NUmlote
                CL.InsertarLote
                
                'JUNIO 2014
                'Regulzarizacion FIN DEPOSITO
                If Val(miRsAux!fincuba) = 1 Then
                    
                    If RegularizacionDeposito <> 0 Then
                        Espera 1.25 'PAra que el apunte lo haga un poco despues en la smoval
                        'Regulzarizaremos el deposito
                        
                        
                        
                        'Un linea mas en smoval
                        vvCstock.DetaMov = "DEP"
                        
                        
    
                        CL.DetaMov = "DEP"  'FIN DEPOSITO
                        CL.HoraMov = CDate(Me.txtFecha(1).Text & " " & Format(Now, "hh:nn:ss"))
                        CL.tipoMov = 1  '0 entrada 1 salida
                        vvCstock.tipoMov = "E"
                        If RegularizacionDeposito > 0 Then
                            CL.tipoMov = 0
                            vvCstock.tipoMov = "S"
                        End If
                        CL.LineaDocu = cDEP.NumDeposito
                        vvCstock.LineaDocu = CL.LineaDocu
                        CL.Cantidad = Abs(RegularizacionDeposito)
                        CL.InsertarLote
                                                                                           
                        cP.FinPartida   'POndra a cero la cantidad
                        
                        
                        'Cantidad
                        
                        If vvCstock.Cantidad > 0 Then vvCstock.Importe = (vvCstock.Importe / vvCstock.Cantidad) * CL.Cantidad
                        vvCstock.Cantidad = CL.Cantidad
                        vvCstock.ActualizarStock False
                        
                        
                        'Dejamos donde estaba el tipo de movimiento
                        CL.DetaMov = "CUP"
                        vvCstock.DetaMov = "CUP"
                    End If
                    'Ponemos vacios los campos del deposito
                    'Fuera numero de lote y fuera kilos
                    
                    cDEP.QuitarAsignacionDeposito_ 1
                    Espera 0.75
                End If
            End If
        End If 'DE incializa stock
        miRsAux.MoveNext
    Wend
    miRsAux.Close
        
    Set cDEP = Nothing

    If SoloComprobar Then
        RealizarCoupageLOTES = True
        If AuxPartida <> "" Then
            AuxPartida = "-  " & Err_x_Articulo & AuxPartida & vbCrLf
            ErroresEnPartidas = ErroresEnPartidas & AuxPartida
        End If
        If ErroresEnPartidas <> "" Then
            ErroresEnPartidas = ErroresEnPartidas & "¿Continuar?"
            If MsgBox(ErroresEnPartidas, vbExclamation + vbYesNo) = vbNo Then RealizarCoupageLOTES = False
        End If
        GoTo ERealizarCUPLOTES 'para k haga los =nothing
    End If

        

    AuxPartida = ""
    
        

    'AHora comprobamos los stcosk de las entraddas , de las lineas
    cad = TransformaComasPuntos(CStr(CantidadMezcla))
    cad = "select codartic," & cad & " kilos,numlote,codalmac,deposito from olicoupage where codigo=" & RecuperaValor(Me.Intercambio, 1)
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'SOLO HAY una linea
    If Not miRsAux.EOF Then
        B = False
        If InicializarCStockCoupage(vvCstock, "E", True) Then    'Las lineas son de netrada
                
                                
                'AHora veremos los numeros de lote
                'EL nUMERO DE LOTE NO puede ser NULO
                CantidadNecesaria = vvCstock.Cantidad  'Para tenerla despues en los lotes
                
                
                                                        'Vemos si ya existe
                If cP.LeerDesdeArticulo(miRsAux!codartic, miRsAux!codAlmac, miRsAux!NUmlote) Then
                    'Ya existia(por algun motivo)
                    cP.IncrementarCantidad CantidadNecesaria
                    
                Else
                    cP.Cantidad = CantidadNecesaria
                    cP.codAlmac = miRsAux!codAlmac
                    cP.codartic = vvCstock.codartic
                    cP.codProve = 0
                    cP.Fecha = CDate(txtFecha(1).Text)
                    cP.NumAlbar = "CUP" & RecuperaValor(Me.Intercambio, 1)
                    cP.NUmlote = miRsAux!NUmlote
                    If Not cP.Insertar Then Err.Raise vbObjectError + 513, , cad
                    
                End If
                
                'Insertamos en la linea de smoval
                CL.tipoMov = 1
                CL.Cantidad = Abs(CantidadNecesaria)
                CL.codAlmac = vvCstock.codAlmac
                CL.codartic = vvCstock.codartic
                CL.NUmlote = cP.NUmlote
                CL.InsertarLote
                
                B = True
                
                Set cDEP = New cDeposito
                'Para que no de error insertando en hco
                T1 = Timer
                If Not cDEP.LeerDatos(miRsAux!Deposito, False) Then B = False
                
                AuxPartida = DevuelveDesdeBD(conAri, "factorconversion", "sartic", "codartic", miRsAux!codartic, "T")
                CantidadNecesaria = CCur(AuxPartida)
                
                
                cDEP.Kilos = CL.Cantidad
                cDEP.NUmlote = cP.NUmlote
                cDEP.IdPartida = cP.IdPartida
                Espera 0.5
                cDEP.InsertarEnDeposito 1
                
                T1 = Timer - T1
                Espera T1
        End If
    End If
        
    RealizarCoupageLOTES = B


    
ERealizarCUPLOTES:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set cP = Nothing
    Set miRsAux = Nothing
    Set vvCstock = Nothing
    Set cDEP = Nothing
End Function






Private Function ActualizarPrecio() As Boolean
Dim B As Boolean
Dim CantidadTotalAProducir As Currency 'Cuatro decimales
Dim PrecioTotal As Currency
Dim C As Currency
Dim Articulo As String
    
    
    Set miRsAux = New ADODB.Recordset
    
    'Vemos si la referencia es de esas
    cad = "select olicoupage.codartic from olicoupage where codigo=" & RecuperaValor(Me.Intercambio, 1)
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Articulo = miRsAux!codartic
    miRsAux.Close
    
    'Estos articulos me los indico ramoon en un Email
    '003500411513  003500421513 003900431513
    B = (Articulo = "003500411513") Or (Articulo = "003500421513") Or (Articulo = "003900431513")
    If Not B Then
        Set miRsAux = Nothing
        Exit Function
    End If
    
    
    'OK.Calculo el precio
    
    
    
    
    'Los mezclantes
    
    cad = "select olicoupagelin.*,preciouc, preciomp from olicoupagelin,sartic where olicoupagelin.codartic=sartic.codartic and "
    cad = cad & "  codigo = " & RecuperaValor(Me.Intercambio, 1)
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    B = False
    
    CantidadTotalAProducir = 0
    PrecioTotal = 0
    While Not miRsAux.EOF
        C = DBLet(miRsAux!precioUC, "N")
        C = miRsAux!Kilos * C
        PrecioTotal = PrecioTotal + C
        CantidadTotalAProducir = CantidadTotalAProducir + miRsAux!Kilos
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    'Si no produce nada nos piramos
    If CantidadTotalAProducir = 0 Then Exit Function
    
    PrecioTotal = Round(PrecioTotal / CantidadTotalAProducir, 4)
    
    cad = "select preciouc,ultfecco from sartic where codartic='" & Articulo & "'"
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    B = False 'Tiene que actualizar
    If IsNull(miRsAux!ultfecco) Then
        B = True
    Else
        If CDate(miRsAux!ultfecco) < CDate(txtFecha(1).Text) Then
            'OK
            'Veremos los importes
            C = DBLet(miRsAux!precioUC)
                                            'Ha cambiado
            If C <> PrecioTotal Then B = True
        End If
    End If
    miRsAux.Close
    
    
  
    If B Then
        'OK. Hay que actualizar los importes
        lbFec(1).Caption = "Act. precio"
        lbFec(1).Refresh
        Espera 0.3
        ActualizarPrecioCosteArticulo PrecioTotal, Articulo
    End If
    Set miRsAux = Nothing
End Function




Private Sub ActualizarPrecioCosteArticulo(ByRef Pre As Currency, ByRef codArt As String)


On Error GoTo EActualizarPrecioCosteArt


    cad = "UPDATE sartic set PrecioUC = " & TransformaComasPuntos(CStr(Pre))
    cad = cad & ", ultfecco = '" & Format(txtFecha(1).Text, FormatoFecha) & "'"
    cad = cad & " WHERE codartic = '" & codArt & "'"
    
    'Ejecutar
    Conn.Execute cad
    Espera 0.2
    
    
    
    
    'Para que se actualice bien
    CommitConexion
    Espera 0.1
    
    'AHora va el meollo. Si el articulo es materia prima, todos los artiuclos
    'de venta en los que el entra como materia prima deben sera actualizados
    cad = "select sartic.codartic,nomartic,codunida from sarti1,sartic where sarti1.codartic = sartic.codartic"
    cad = cad & " AND codarti1 = '" & codArt & "'"
    miRsAux.Open cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        NumRegElim = 0
        While Not miRsAux.EOF
            NumRegElim = NumRegElim + 1
            miRsAux.MoveNext
        Wend
        miRsAux.MoveFirst
        Pre = 1
        While Not miRsAux.EOF
            lbFec(1).Caption = "UPC " & CInt(Pre) & " de " & NumRegElim
            lbFec(1).Refresh
            ActualizaUPCArticuloCabecera miRsAux!codartic, CInt(miRsAux!CodUnida)
            Pre = CInt(Pre) + 1
            miRsAux.MoveNext
            If (CInt(Pre) Mod 15) = 0 Then
                Me.Refresh
                DoEvents
            End If
        Wend
     
    End If
    miRsAux.Close
    
EActualizarPrecioCosteArt:
    
    If Err.Number <> 0 Then MuestraError Err.Number, "Actualizando precio coste"
    Screen.MousePointer = vbDefault
    Set miRsAux = Nothing

End Sub




Private Sub ActualizaUPCArticuloCabecera(ByRef C As String, CodUnida As Integer)
Dim Aux As String
Dim RS As ADODB.Recordset
Dim Im0 As Currency
Dim Im1 As Currency

    On Error GoTo eActualizaUPCArticuloCabecera
    Set RS = New ADODB.Recordset
    Aux = "SELECT sarti1.codartic, numlinea, sarti1.codarti1,sartic.nomartic, sarti1.Cantidad ,"
    Aux = Aux & "sartic.preciove , sartic.precioUC, FactorConversion"
    Aux = Aux & " FROM   sarti1 INNER JOIN sartic ON sarti1.codarti1 = sartic.codArtic where sarti1.codartic='"
    Aux = Aux & C & "' ORDER BY sarti1.numlinea"
    RS.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Im1 = 0
    Aux = ""
    While Not RS.EOF
        Aux = RS!NomArtic
        Im0 = DBLet(RS!FactorConversion, "N")  'del articulo de la linea

        'COSTE
        Im0 = DBLet(RS!Cantidad, "N") * Im0
        Im0 = Im0 * DBLet(RS!precioUC, "N")
        Im1 = Im1 + Im0
        
        RS.MoveNext
    Wend

    RS.Close
    
    'El formato
    Aux = "Select sum(importe) from sunilin where codunida=" & CodUnida
    RS.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Im0 = 0
    If Not RS.EOF Then Im0 = DBLet(RS.Fields(0), "N")
    RS.Close

    'Redondeamos (al igual que en el mantenimiento de articulos) a 3 antes de sumar el formato
    Im1 = Round(Im1, 3)

    Im1 = Im1 + Im0
    Im1 = Round2(Im1, 3)
    
    'UPDATEAMOS
    Aux = "UPDATE sartic set PrecioUC = " & TransformaComasPuntos(CStr(Im1))
    Aux = Aux & ", ultfecco = '" & Format(txtFecha(1).Text, FormatoFecha) & "'"
    Aux = Aux & " WHERE codartic = '" & C & "'"
    Conn.Execute Aux
    
eActualizaUPCArticuloCabecera:
    If Err.Number <> 0 Then MuestraError Err.Number, Aux
    Set RS = Nothing
End Sub


Private Sub txtMeses_GotFocus()
    ConseguirFoco txtMeses, 3
End Sub

Private Sub txtMeses_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub


Private Sub txtMeses_LostFocus()
    txtMeses.Text = Trim(txtMeses.Text)
    If txtMeses.Text = "" Then Exit Sub
    
    If Not IsNumeric(txtMeses.Text) Then
        MsgBox "Campo numerico", vbExclamation
        txtMeses.Text = "18"
        PonerFoco txtMeses
    End If
    
    txtMeses.Text = Abs(Val(txtMeses.Text))
    
        
        
End Sub




Private Sub CargaComobosTrasiegos(Inicio As Byte, Fin As Byte)

    Set miRsAux = New ADODB.Recordset
    For i = Inicio To Fin
        cboDeposito(i).Clear
        
        If i = 0 Or i = 2 Or i = 4 Then
            cad = "SELECT proddepositos.numdeposito, spartidas.codartic, sartic.nomartic, spartidas.numlote, kilos vlitros"
            '(kilos * factorconversion) vlitros"
            cad = cad & " FROM  proddepositos left join spartidas on spartidas.numlote=proddepositos.numlote"
            cad = cad & " inner join sartic on spartidas.codartic=sartic.codartic AND sartic.factorconversion<1"
            cad = cad & " Where Not spartidas.numLote Is Null"
            cad = cad & " ORDER BY numdeposito"
    
        Else

            cad = "select * from proddepositos where numlote is null"
        
        End If
        
        miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            If i = 0 Or i = 2 Or i = 4 Then
                cad = Format(miRsAux!NumDeposito, "00") & "  L" & Mid(miRsAux!NUmlote & "       ", 1, 9) & " " & miRsAux!NomArtic & " (" & Format(miRsAux!vlitros, FormatoCantidad) & ")"
            Else
                cad = "Deposito " & miRsAux!NumDeposito
            End If
            cboDeposito(i).AddItem cad
            cboDeposito(i).ItemData(cboDeposito(i).NewIndex) = miRsAux!NumDeposito
            
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    Next i
    Set miRsAux = Nothing
End Sub





'Este trozo esta copiado de proceso produccion
'De momento solo entra aqui para materia prima
Private Sub RegularizarFinLote_Partida(ByRef cDEP As cDeposito)
Dim cPar As cPartidas

Dim cLot As cLotaje
Dim vvCstock As CStock
Dim Aux As String
Dim Donde As String
Dim Cantidad As Currency

    On Error GoTo eRegularizarFinLote_Partida

    
    
    Set cPar = New cPartidas
    Set cLot = New cLotaje
    Set vvCstock = New CStock
    
    Donde = "Leyendo clases"
    
    'select * from spartidas,sartic where spartidas.codartic=sartic.codartic and sartic.factorconversion<1 and numlote in (select numlote from proddepositos)
    Aux = "spartidas.codartic=sartic.codartic and sartic.factorconversion<1 and numlote"
    Aux = DevuelveDesdeBD(conAri, "id", "spartidas,sartic", Aux, cDEP.NUmlote, "T")
    If Aux = "" Then Err.Raise 513, , "No se encuentra la partida"
    cPar.Leer CLng(Aux)
    
    
    
        
    Set cLot = New cLotaje
    Set vvCstock = New CStock
        
   
    
    
    'Un linea mas en smoval
    vvCstock.DetaMov = "DEP"
    '0=Salida, 1=Entrada
    If cPar.Cantidad >= 0 Then
        vvCstock.tipoMov = "S"
        cLot.tipoMov = 0
    Else
        vvCstock.tipoMov = "E"
        cLot.tipoMov = 1
    End If
    vvCstock.Cantidad = Abs(cPar.Cantidad)
    vvCstock.Trabajador = TrabajadorConectado_
    'vCStock.Documento = RecuperaValor(Intercambio, 1)
    vvCstock.Fechamov = Format(Now, "dd/mm/yyyy")
    vvCstock.HoraMov = Now
    vvCstock.codAlmac = cPar.codAlmac
    vvCstock.codartic = cPar.codartic
    vvCstock.Importe = 0
    vvCstock.Documento = "FIN" & Format(cPar.IdPartida, "0000000")
    
    cLot.codAlmac = vvCstock.codAlmac
    cLot.codartic = vvCstock.codartic
    cLot.DetaMov = vvCstock.DetaMov
    cLot.Fechamov = vvCstock.Fechamov
    cLot.HoraMov = vvCstock.HoraMov
    cLot.NUmlote = cPar.NUmlote
    
    cLot.Cantidad = vvCstock.Cantidad
    cLot.LineaDocu = cDEP.NumDeposito
    cLot.Documento = vvCstock.Documento
    
    cLot.InsertarLote

    vvCstock.ActualizarStock False
    cPar.AjustarFinPartida
    
    
                        
    
eRegularizarFinLote_Partida:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set cPar = Nothing
    Set cLot = Nothing
    Set vvCstock = Nothing
    
End Sub


