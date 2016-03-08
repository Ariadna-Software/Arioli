VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "Mscomm32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTerminal 
   Caption         =   "Poste paletizado Aceites Morales"
   ClientHeight    =   4935
   ClientLeft      =   2940
   ClientTop       =   1755
   ClientWidth     =   7155
   ForeColor       =   &H00000000&
   Icon            =   "vbterm.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   7155
   Begin VB.Timer TimerServicio 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4800
      Top             =   1200
   End
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   6615
      Begin VB.Label lblDupl 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "D U P L I C A D O"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   325
         Left            =   3240
         TabIndex        =   4
         Top             =   120
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Conectar / desconectar"
         Height          =   255
         Left            =   600
         TabIndex        =   3
         Top             =   150
         Width           =   1935
      End
      Begin VB.Image imgNotConnected2 
         Height          =   240
         Left            =   120
         Picture         =   "vbterm.frx":030A
         Stretch         =   -1  'True
         ToolTipText     =   "Alternar puerto"
         Top             =   120
         Width           =   240
      End
      Begin VB.Image imgConnected 
         Height          =   240
         Left            =   120
         Picture         =   "vbterm.frx":0454
         Stretch         =   -1  'True
         ToolTipText     =   "Alternar puerto"
         Top             =   120
         Width           =   240
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   210
      Top             =   3645
   End
   Begin VB.TextBox txtTerm 
      Height          =   2370
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   1920
      Width           =   5790
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   165
      Top             =   1815
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   240
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      NullDiscard     =   -1  'True
      RThreshold      =   1
      RTSEnable       =   -1  'True
      SThreshold      =   1
      InputMode       =   1
   End
   Begin MSComDlg.CommonDialog OpenLog 
      Left            =   105
      Top             =   1170
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "LOG"
      FileName        =   "Abrir archivo de registro de comunicaciones"
      Filter          =   "Archivo de registro (*.log)|*.log;"
      FilterIndex     =   501
      FontSize        =   1,17491e-38
   End
   Begin MSComctlLib.StatusBar sbrStatus 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   4620
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "Estado:"
            TextSave        =   "Estado:"
            Key             =   "Status"
            Object.ToolTipText     =   "Estado del puerto de comunicaciones"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8283
            MinWidth        =   2
            Text            =   "Valores:"
            TextSave        =   "Valores:"
            Key             =   "Settings"
            Object.ToolTipText     =   "Valores del puerto de comunicaciones"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1244
            MinWidth        =   1244
            Key             =   "ConnectTime"
            Object.ToolTipText     =   "Tiempo de conexión"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   165
      Top             =   2445
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbterm.frx":059E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbterm.frx":08B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbterm.frx":0BD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbterm.frx":0EEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbterm.frx":1206
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbterm.frx":1520
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTerminal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------
' VBTerm - Este es un programa de demostración del
' control ActiveX de comunicaciones MSComm.
'
' Copyright (c) 1994, Crescent Software, Inc.
' por Don Malin y Carl Franklin.
'
' Actualizado por Mike Maddox
'--------------------------------------------------
Option Explicit
                        
Dim Ret As Integer      ' Entero auxiliar.
'Dim Temp As String      ' Cadena auxiliar.
Dim StartTime As Date   ' Almacena la hora de inicio del cronómetro del puerto


Dim PrimeraVez As Boolean


Private SegundosServicio As Byte
Private Reintentos As Long


Private NumAriges As Byte
Private UltVezAbierto As Date

Private UltimaEntradaInformacion As Date

Private UltimaVezConBD As Date  'Si hace mas de una hora que no hace conn.algo cerraremos y volveremos a abir
Private ConnAbierta As Boolean

Dim FiLog As Integer
Dim NomLog As String
Dim CreacionLog As Date
 
Dim ABiertoLog As Boolean

Dim SePuedeCerrar As Boolean
Dim HayCajasDuplicadas As Boolean
Dim Segundos_Duplicada As Single

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Conectar
    
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub FijarParametrosCom()
Dim CommPort As String, Handshaking As String, Settings As String



    'Settings = GetSetting(App.Title, "Properties", "Settings", "") ' frmTerminal.MSComm1.Settings]\
    Settings = "19200,n,8,1"
    If Settings <> "" Then
        MSComm1.Settings = Settings
        If Err Then
            MensajeError Error$, 2
            End
        End If
    End If



    'Puerto parametrizable.  Arichov llamado Nºpuerto.port  ej: 3.port --> com3
    CommPort = "5"
    Handshaking = Dir(App.Path & "\*.port", vbArchive)
    If Handshaking <> "" Then
        'Tiene puesto un puerto a mano
        Ret = InStr(1, Handshaking, ".")
        Handshaking = Mid(Handshaking, 1, Ret - 1)
        
        If Handshaking = "" Then
            MensajeError "Error cadena puerto(a)", 1
        Else
            If Not IsNumeric(Handshaking) Then
                MensajeError "Error cadena puerto(b)", 1
            Else
                CommPort = Handshaking
            End If
        End If
    End If
    If CommPort <> "" Then MSComm1.CommPort = CommPort

    
    Handshaking = ""
    If Handshaking <> "" Then
        MSComm1.Handshaking = Handshaking
        If Err Then
            MensajeError Error$, 2
            End
        End If
    End If
    
    'FIAJAMOS LA BD
    Settings = Dir(App.Path & "\*.ari", vbArchive)
    If Settings = "" Then Settings = "1.ari"
    Ret = InStr(1, Settings, ".")
    Settings = Mid(Settings, 1, Ret - 1)
    If Settings = "" Then Settings = "1"
    NumAriges = CByte(Settings)
    


    
    


End Sub


Private Sub Form_Load()
    
        
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    PrimeraVez = True
    App.Title = "Terminal de Poste paletizado"
    
    EsServicio = False
    If Dir(App.Path & "\EsServicio.dat", vbArchive) <> "" Then EsServicio = True
    
    'EsServicio = True
    Reintentos = 0
    FijarParametrosCom
    
    ' Establece el color predeterminado del terminal
    If Not EsServicio Then
        txtTerm.SelLength = Len(txtTerm)
        txtTerm.SelText = ""
        txtTerm.ForeColor = vbBlue
       
        ' Establece la luz indicadora de estado
        imgNotConnected2.ZOrder
       
        ' Centra el formulario
        frmTerminal.Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    
    End If
    

    
    'Handshaking = GetSetting(App.Title, "Properties", "Handshaking", "") 'frmTerminal.MSComm1.Handshaking
    
    'Echo = GetSetting(App.Title, "Properties", "Echo", "") ' Echo
    Echo = True
    HayCajasDuplicadas = False
    lblDupl.Visible = False
    
    
    
    SePuedeCerrar = True
    
    'Si es servicio abre un LOG
    If EsServicio Then
        ProcesarFicheroLog True
        Print #FiLog, "Inicio aplicacion: " & Now
    End If
        
    'ABrimos BD
    AbrirCon
    If ConnAbierta Then UltimaVezConBD = Now
    
        
        
    If EsServicio Then
        
        'Abrimos ya el puerto
        Conectar
        
        TimerServicio.Enabled = True
        If Not PuertoAbierto Then
            If Not EsServicio Then MensajeError "Imposible abrir puerto", 2
            
        Else
            'Como tenemos un conversor rs232-wifi
            'parece ser que hay que enviarle una trama para que comience la transmision
            
        End If
        
    End If
    
    
    
    On Error GoTo 0

End Sub


Private Sub EnvioInicial()
    On Error Resume Next
    MSComm1.Output = "David"
    Err.Clear
End Sub



Private Sub ProcesarFicheroLog(EsPrimeraVez As Boolean)
        
        On Error Resume Next
        If Not ABiertoLog Then FiLog = FreeFile
        
        If Not EsPrimeraVez Then
            If Not ABiertoLog Then
                If CrearNuevoLog Then EsPrimeraVez = True
            End If
        End If
        
        If EsPrimeraVez Then
        
            If Dir(App.Path & "\LOG", vbDirectory) = "" Then MkDir App.Path & "\LOG"
            NomLog = App.Path & "\LOG\" & Format(Now, "yymmdd_hhnnss") & ".log"
            Open NomLog For Output As #FiLog
            CreacionLog = Now
        Else
            If Not ABiertoLog Then Open NomLog For Append As #FiLog
        End If
        UltVezAbierto = Now
        If Err.Number <> 0 Then
            MensajeError Err.Description, 2
            'End  NO LO PARAMOS
            FuerzaCierre
        Else
             ABiertoLog = True
        End If
       
End Sub

Private Sub FuerzaCierre()
Dim i As Integer
    On Error Resume Next
    i = FiLog
    For i = FiLog To 1 Step -1
        Close #i
        Err.Clear
    Next
    ABiertoLog = False
    
End Sub
Private Sub Form_Resize()
Dim Fr As Integer
    If Me.WindowState = vbMinimized Then Exit Sub
   Fr = 60
   Frame1.Width = frmTerminal.ScaleWidth
   ' Cambia el tamaño del control Term (ventana)
   txtTerm.Move 0, Fr + Frame1.Height, frmTerminal.ScaleWidth, frmTerminal.ScaleHeight - Frame1.Height - sbrStatus.Height - Fr
   

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim Counter As Long

    If Not EsServicio Then
        If PuertoAbierto Then
            Cancel = 1
            Me.WindowState = 1
            Exit Sub
        End If
    End If
        

    If PuertoAbierto Then
       ' Espera 10 segundos para transmitir los datos.
       MSComm1.PortOpen = 0
    End If

    On Error Resume Next
    Close #FiLog
    Conn.Close
    Set Conn = Nothing
    End
    
End Sub

Private Sub Frame1_DblClick()
   ' Dim SQL
   ' SQL = ProcesarCodigoCaja("0000429000151")
End Sub

Private Sub imgConnected_Click()
    ' Llama a la rutina mnuOpen_Click para alternar entre conectar y desconectar.
    Call Abre2
End Sub

Private Sub Conectar()


        Call Abre2
        If EsServicio Then
            If PuertoAbierto Then
                'App.LogEvent "Puerto abierto", vbLogEventTypeInformation
            Else
              '  App.LogEvent "Abrir puerto", vbLogEventTypeError
            End If
        End If
       
    
    
    
End Sub



'Private Sub mnuCloseLog_Click()
'    ' Cierra el archivo de registro.
'    Close hLogFile
'    hLogFile = 0
'    mnuOpenLog.Enabled = True


Private Sub mnuProperties_Click()
  ' Muestra el formulario de propiedades de CommPort
  frmProps.Show vbModal
  
End Sub

' Alterna el estado del puerto (abierto o cerrado).
Private Sub Abre2()
    On Error Resume Next
    Dim OpenFlag

    MSComm1.PortOpen = Not MSComm1.PortOpen
    If Err Then
       If Not EsServicio Then MensajeError Error$, 2
       Err.Clear
    End If
    
    If EsServicio Then
        If Not PuertoAbierto Then
            If ABiertoLog Then Print #FiLog, "Intento " & Reintentos & "   " & Now
            'If (Reintentos Mod 60) = 0 Then MensajeError Error$, 2
            SegundosServicio = 0
            If ABiertoLog Then Close #FiLog
            ABiertoLog = False
            
            
        Else
            EnvioInicial
            If ABiertoLog Then Print #FiLog, "Puerto abierto:  " & Now
            Reintentos = 0
            UltimaEntradaInformacion = Now
            
        End If
        Err.Clear
        Exit Sub 'salimos y no hacemos nada mas
    
    End If
    
    OpenFlag = PuertoAbierto
    
        If PuertoAbierto Then
        ' Habilita el botón de marcar y el elemento de menú asociado
        
        Me.Caption = "Poste paletizado Aceites Morales"
        SePuedeCerrar = False
        imgConnected.ZOrder
        sbrStatus.Panels("Settings").Text = "Valores:" & MSComm1.Settings
        UltimaEntradaInformacion = Now
        StartTiming
    Else
        SePuedeCerrar = True
        Me.Caption = "Poste CERRADO"
      
        
        imgNotConnected2.ZOrder
        sbrStatus.Panels("Settings").Text = "Valores:"
        StopTiming
    End If


    
    
    If Err.Number <> 0 Then Err.Clear
    
End Sub




Private Sub imgNotConnected2_Click()
  Call Abre2
End Sub

' El evento OnComm se usa para interceptar eventos y errores de comunicaciones.
Private Static Sub MSComm1_OnComm()
    Dim EVMsg$
    Dim ERMsg$
    Dim Aux As String
    
    
    On Error GoTo eMscomm
    
    
    ' Bifurca según la propiedad CommEvent.
    Select Case MSComm1.CommEvent
        ' Mensajes de evento.
        Case comEvReceive
            Dim Buffer As Variant
            Buffer = MSComm1.Input
            
            Aux = StrConv(Buffer, vbUnicode)
            Debug.Print "Recibir - " & Aux
            ShowData txtTerm, Aux
           
           
           
           
        Case comEvSend
        Case comEvCTS
            EVMsg$ = "Detectado cambio en CTS"
        Case comEvDSR
            EVMsg$ = "Detectado cambio en DSR"
        Case comEvCD
            EVMsg$ = "Detectado cambio en CD"
        Case comEvRing
            EVMsg$ = "El teléfono está sonando"
        Case comEvEOF
            EVMsg$ = "Detectado el final del archivo"

        ' Mensajes de error.
        Case comBreak
            ERMsg$ = "Parada recibida"
        Case comCDTO
            ERMsg$ = "Sobrepasado el tiempo de espera de detección de portadora"
        Case comCTSTO
            ERMsg$ = "Soprepasado el tiempo de espera de CTS"
        Case comDCB
            ERMsg$ = "Error recibiendo DCB"
        Case comDSRTO
            ERMsg$ = "Sobrepasado el tiempo de espera de DSR"
        Case comFrame
            ERMsg$ = "Error de marco"
        Case comOverrun
            ERMsg$ = "Error de sobrecarga"
        Case comRxOver
            ERMsg$ = "Desbordamiento en el búfer de recepción"
        Case comRxParity
            ERMsg$ = "Error de paridad"
        Case comTxFull
            ERMsg$ = "Búfer de transmisión lleno"
        Case Else
            ERMsg$ = "Error o evento desconocido"
    End Select
    
    If Len(EVMsg$) Then
    
        If EsServicio Then
            'En el LOG meto
            Print #FiLog, "Ev.--->" & EVMsg$
        Else
            ' Muestra los mensajes de evento en la barra de estado.
            sbrStatus.Panels("Status").Text = "Estado:" & EVMsg$
        
            ' Activa el cronómetro para que el mensaje de la barra
            ' de estado se borre después de dos segundos.
            Timer2.Enabled = True
        End If
    ElseIf Len(ERMsg$) Then
        ' Muestra los mensajes de evento en la barra de estado.
        If EsServicio Then
            'En el LOG meto
            Print #FiLog, "ERR--->" & EVMsg$
        Else
            sbrStatus.Panels("Status").Text = "Estado:" & ERMsg$
        
        
            ' Muestra los mensajes de error en un cuadro de alerta.
            Beep
            'Ret = MsgBox(ERMsg$, 1, "Haga clic en Cancelar para salir, clic en Aceptar para ignorar.")
        
            ' Si el usuario hace clic en Cancelar (2)...
            If Ret = 2 Then
                MSComm1.PortOpen = False    ' Cierra el puerto y sale.
            End If
        
            ' Activa el cronómetro para que el mensaje de la barra
            ' de estado se borre después de dos segundos.
            Timer2.Enabled = True
        End If
    End If
    
    UltVezAbierto = Now 'guardo el dato
    Exit Sub
eMscomm:
    Aux = Err.Description
    Err.Clear
    If ABiertoLog Then Print #FiLog, Aux
        
        
End Sub






' Este procedimiento agrega datos a la propiedad Text del
' control Term. También filtra los caracteres de control,
' como RETROCESO, retorno de carro y avances de línea, y
' escribe datos en un archivo de registro.
' Los caracteres RETROCESO eliminan el carácter situado a
' su izquierda, ya sea en la propiedad Text o en la cadena
' pasada. Se agregan caracteres de avance de línea a todos
' los retornos de carro. El tamaño de la propiedad Text del
' control Term también se controla para que nunca exceda de
' MAXTERMSIZE caracteres.
Private Static Sub ShowData(Term As Control, Data As String)
    On Error GoTo Handler
    Const MAXTERMSIZE = 16000
    Dim TermSize As Long, i
    
    
   
    'CONNEXION BD
    InsertaEnBD Data
    
    If EsServicio Then
        If MeterDatoLeidoLog Then Print #FiLog, Data
        UltimaEntradaInformacion = Now
        Exit Sub
    Else
        If MeterDatoLeidoLog Then Term.Text = Term.Text & vbCrLf & Now & vbCrLf
        UltimaEntradaInformacion = Now
        If InStr(1, Data, vbCrLf) = 0 Then Term.Text = Term.Text & vbCrLf
    End If
    
    
    ' Se asegura que el texto existente no se haga demasiado largo.
    TermSize = Len(Term.Text)
    If TermSize > MAXTERMSIZE Then
       Term.Text = Mid$(Term.Text, 4097)
       TermSize = Len(Term.Text)
    End If

    ' Apunta al final de los datos de Term.
    Term.SelStart = TermSize

    ' Filtra y procesa los caracteres RETROCESO.
    Do
       i = InStr(Data, Chr(8))
       If i Then
          If i = 1 Then
             Term.SelStart = TermSize - 1
             Term.SelLength = 1
             Data = Mid$(Data, i + 1)
          Else
             Data = Left$(Data, i - 2) & Mid$(Data, i + 1)
          End If
       End If
    Loop While i


    Term.SelText = Data
  
   
    Term.SelStart = Len(Term.Text)
Exit Sub

Handler:
    MensajeError Error$, 1
    Resume Next
    If ABiertoLog Then Print #FiLog, "Err show data"
End Sub


Private Function MeterDatoLeidoLog() As Boolean
Dim N As Long
    
    On Error Resume Next
    MeterDatoLeidoLog = True
    If Not EsServicio Then MeterDatoLeidoLog = False
    N = DateDiff("n", UltimaEntradaInformacion, Now)
    If N > 2 Then MeterDatoLeidoLog = True
    Err.Clear
End Function

Private Sub Timer2_Timer()
    sbrStatus.Panels("Status").Text = "Estado:"
    Timer2.Enabled = False

End Sub

Private Sub ComprobarLaConexion2()
Dim N As Long

                
        N = DateDiff("n", UltimaVezConBD, Now)
        If N > 60 Then
            'Si lleva una hora sin hacer nada, cierro conn y vuelvo a abrir
            AbrirCon
            UltimaVezConBD = Now
            
            'Mando un SQL para que lo inserte
            InsertaEnBD "Test BD"
            
        End If

End Sub


Private Sub TimerServicio_Timer()
Dim B As Boolean
Dim N As Long
    DoEvents
    B = False
    SegundosServicio = SegundosServicio + 1
    
    
    MatarAplicacion
    
    If PuertoAbierto Then
        'Comprobar inactividad
        If SegundosServicio > 100 Then
            TimerServicio.Enabled = False
                SegundosServicio = 0
            
           
            
                ComprobarLaConexion2

                TiempoFicheroAbierto
            
            
                ReaabriPuerto
                
                DoEvents
            TimerServicio.Enabled = True
        End If
    
    Else
        If Reintentos < 10 Then
            'Cada 10 segundos
            
            If SegundosServicio > 30 Then
                SegundosServicio = 0
                B = True
            End If
        Else
            If Reintentos < 20 Then
                'Cada minuto
                If SegundosServicio > 90 Then
                    SegundosServicio = 0
                    B = True
                End If
            Else
                
                If SegundosServicio > 200 Then
                    SegundosServicio = 0
                    B = True
                End If
            End If
        End If
        If B Then
            TimerServicio.Enabled = False
            HacerReintento
            If PuertoAbierto Then
                ComprobarLaConexion2
                SegundosServicio = 0
            End If
            TimerServicio.Enabled = True
        End If
         DoEvents
    End If
End Sub

Private Sub HacerReintento()
    
    ProcesarFicheroLog False
    Reintentos = Reintentos + 1
    Conectar
End Sub


' Las pulsaciones interceptadas aquí se envían
' al control MSComm, donde se devuelven a través
' del evento OnComm (comEvReceive), y se muestran
' con el procedimiento ShowData.
Private Sub txtTerm_KeyPress(KeyAscii As Integer)
    ' Si el puerto está abierto...
    If MSComm1.PortOpen Then
        ' Envía la pulsación al puerto.
        MSComm1.Output = Chr(KeyAscii)
        
        ' Si el eco no está activado, no hay
        ' necesidad de que el control de texto
        ' muestre la tecla. Normalmente, el módem
        ' devolverá el carácter.
        If Not Echo Then
            ' Sitúa la posición al final del terminal
            txtTerm.SelStart = Len(txtTerm)
            KeyAscii = 0
        End If
    End If
     
End Sub






Private Sub Timer1_Timer()
Dim HabianCajas As Boolean
Dim Aux As String
    ' Muestra la hora de conexión
    sbrStatus.Panels("ConnectTime").Text = Format(Now - StartTime, "hh:nn:ss") & " "
    If HayCajasDuplicadas Then
        
        If Timer - Segundos_Duplicada > 2 Then
            'Comprobamos si siguen duplicados
            Aux = DevuelveDesdeBD("count(*)", "prodcajasduplicadas", "1", "1")
            If Aux = "" Then Aux = "0"
            HayCajasDuplicadas = Val(Aux) > 0
            'Ya han visto el error
            If Not HayCajasDuplicadas Then HabianCajas = True
               
        End If
    End If
    
    If HabianCajas Then
        lblDupl.Visible = False
        If Not HayCajasDuplicadas Then Me.Caption = "Poste paletizado Aceites Morales"
    End If
End Sub
' Llama a esta función para iniciar el cronómetro ConnectTime
Private Sub StartTiming()
    StartTime = Now
    Timer1.Enabled = True
End Sub
' Llama a esta función para detener el cronometraje
Private Sub StopTiming()
    Timer1.Enabled = False
    sbrStatus.Panels("ConnectTime").Text = ""
End Sub


Private Sub InsertaEnBD(ByVal Valor As String)
Dim C As String
    
    
    
     'El ID es autnomerico y NO lo pongo
    'insert into `prodlecturaposte` (fechahora,`lectura`) values ( '2000-01-01 12:39:49','sadsd')
    C = Replace(Valor, vbCrLf, "")
    C = Replace(C, vbLf, "")
    C = Replace(C, vbCr, "")
    Valor = DevNombreSQL(C)
    C = "insert into prodlecturaposte (fechahora,lectura) VALUES (CURRENT_TIMESTAMP (),'" & Valor & "')"
    On Error Resume Next
    
    
    Conn.Execute C
    If Err.Number <> 0 Then
        MensajeError Err.Description, 2
        Err.Clear
    Else
        UltimaVezConBD = Now   'realmente ha hecho algo
    End If
    
    
    C = ProcesarCodigoCaja(Valor)
    If C = "" Then C = "OK"
    If ABiertoLog Then Print #FiLog, C
        
    
End Sub




Private Function ProcesarCodigoCaja(CodigoCaja As String) As String
Dim C As String
Dim Aux As String
Dim OK As Boolean
Dim LineaPal As String
Dim idTraza As Long

    
    ProcesarCodigoCaja = ""
    If CodigoCaja = "" Then Exit Function
    
    
    If Len(CodigoCaja) <> 13 Then
        'MsgBox "Longitudad etiqueta incorrecta", vbExclamation
        ProcesarCodigoCaja = "Longitudad etiqueta incorrecta " & Len(CodigoCaja)
    Else
        If Not IsNumeric(CodigoCaja) Then ProcesarCodigoCaja = "Campo no numérico"
    End If
    
    If ProcesarCodigoCaja <> "" Then Exit Function
    
    'Dividimos la etiqueta leida en 2
    'los cinco ultimos son el IDCAJa
    'el resto ID trza

    C = Mid(CodigoCaja, 1, Len(CodigoCaja) - 5)
    
        
        
    idTraza = Val(C)
    'Pongo ffin de la propalettrza no la del palet. Puede que vayamos a paletizar otr
    C = DevuelveDesdeBD("prodpalets.idpalet", "prodpalets,prodpaletstraza", "prodpalets.idpalet= prodpaletstraza.idpalet and fhfin is null and lotetraza ", C)
    
    If C = "" Then
        ProcesarCodigoCaja = "No existe palet asignado para la idtraza"
        C = "NULL" 'para que meta el IDCAJA
    End If
    
    'Si es correcta metemos en la linea de produccion
    '
    C = "insert into prodcajas(lotetraza,idcaja,idpalet,fcreacion) VALUES (" & idTraza & "," & Right(CodigoCaja, 5) & "," & C & ",CURRENT_TIMESTAMP ())"
    
    
    
    'If Not Ejecuta(C) Then ProcesarCodigoCaja = ProcesarCodigoCaja & "    " & "YA existe la caja"
    Aux = InsertaEnProdcajas(C)
    If Aux <> "" Then ProcesarCodigoCaja = ProcesarCodigoCaja & "    " & Aux
    
End Function



Private Function InsertaEnProdcajas(Insert As String) As String
Dim Duplicada As Boolean
Dim Aux As String
    On Error Resume Next
    
    InsertaEnProdcajas = ""
    Duplicada = False
    Conn.Execute Insert
    If Err.Number <> 0 Then
        'Error. Veremos si es error de duplicado
        Aux = Err.Description
        Aux = UCase(Aux)
        If InStr(1, Aux, "DUPLICATE ENTRY") > 0 Then
            'Entrada duplicada
            InsertaEnProdcajas = "YA existe la caja"
            Duplicada = True
            
        Else
            InsertaEnProdcajas = Aux
        End If
        Err.Clear
    End If
    
    'prodcajasduplicadas
    If Duplicada Then
        'Insertamos en duplicada
        Aux = Insert
        Aux = Replace(Insert, "prodcajas", "prodcajasduplicadas")
        Ejecuta Aux
        
        lblDupl.Visible = True
        Caption = "ERROR ERROR Cajas duplicadas"
        HayCajasDuplicadas = True
    End If
End Function



Public Function DevuelveDesdeBD(kCampo As String, Ktabla As String, Kcodigo As String, ValorCodigo As String, Optional Tipo As String, Optional ByRef otroCampo As String) As String
    Dim Rs As Recordset
    Dim Cad As String
    Dim Aux As String
    
    On Error GoTo EDevuelveDesdeBD
    DevuelveDesdeBD = ""
    Cad = "Select " & kCampo
    If otroCampo <> "" Then Cad = Cad & ", " & otroCampo
    Cad = Cad & " FROM " & Ktabla
    Cad = Cad & " WHERE " & Kcodigo & " = "
    If Tipo = "" Then Tipo = "N"
    Select Case Tipo
    Case "N"
        'No hacemos nada
        Cad = Cad & ValorCodigo
    Case "T", "F"
        Cad = Cad & "'" & ValorCodigo & "'"
    Case Else
        MsgBox "Tipo : " & Tipo & " no definido", vbExclamation
        Exit Function
    End Select
    
'    Debug.Print cad
    
    'Creamos el sql
    Set Rs = New ADODB.Recordset
    
    
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        DevuelveDesdeBD = DBLet(Rs.Fields(0))
        If otroCampo <> "" Then otroCampo = DBLet(Rs.Fields(1))
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
EDevuelveDesdeBD:
        MensajeError "Devuelve DesdeBD." & vbCrLf & Err.Description, 1
End Function





'Ponemos todas las funciones del mscomm dentro de otra funcion

Private Function PuertoAbierto() As Boolean

    On Error Resume Next
    PuertoAbierto = MSComm1.PortOpen
    If Err.Number <> 0 Then
        Err.Clear
        PuertoAbierto = False
    End If
    
   
    
End Function



Private Sub AbrirCon()
Dim Aux As String
Dim C As Boolean

    On Error GoTo eAbrirCon

    'reutilizacion de variables
    If ConnAbierta Then
        Conn.Close
        ConnAbierta = False
    End If
    
    
    C = ABiertoLog 'si estaba cerrado
    If Not ABiertoLog Then ProcesarFicheroLog False
    
    
    
    
    
    Aux = AbrirConexion2(NumAriges)
    If Aux <> "" Then
        If EsServicio Then
            If ABiertoLog Then Print #FiLog, "Error abriendo BD: " & Err.Description
        Else
            MensajeError Aux, 2 'critical
        End If
    Else
        If EsServicio Then
            If ABiertoLog Then Print #FiLog, "Conexion BD OK "
        End If
        ConnAbierta = True
    End If
    
    If Not C Then
        'estaba cerrado, lo cierro
        Close #FiLog
        ABiertoLog = False
    End If
    
    Exit Sub
eAbrirCon:
    MensajeError "Abrir BD", 2  'critical
    Set Conn = Nothing
    ConnAbierta = False
    If ABiertoLog Then Print #FiLog, "Error abrir BD "
End Sub

Private Sub TiempoFicheroAbierto()

Dim N As Long
    
        If Not ABiertoLog Then Exit Sub
                
        N = DateDiff("n", UltVezAbierto, Now)
        If N > 15 Then
            Close #FiLog
            ABiertoLog = False
            ProcesarFicheroLog False
        End If
End Sub



Private Sub ReaabriPuerto()
Dim N As Long
    
        On Error GoTo eReaabriPuerto
                
        If Not PuertoAbierto Then Exit Sub
        
        N = DateDiff("h", UltimaEntradaInformacion, Now)
        If N > 2 Then
            Me.MSComm1.PortOpen = False
      
        End If
        Exit Sub
eReaabriPuerto:
    MensajeError "Reabrir puerto", 2  'critical
    
End Sub








Private Sub MatarAplicacion()
Dim Matar As Boolean
    On Error Resume Next
    
    
    Matar = False
    If Dir(App.Path & "\Matar.dat", vbArchive) <> "" Then Matar = True
    If Not Matar Then Exit Sub
    
    
    Name App.Path & "\Matar.dat" As App.Path & "\Matar2.dat"
    If Err.Number <> 0 Then Err.Clear
    
    
    'Si NO significa que quiero forzar a que se cierre la aplicacion
    Set Conn = Nothing
    End
    

End Sub


Private Function CrearNuevoLog() As Boolean
Dim N As Long
    
    On Error GoTo eCrearNuevoLog
    CrearNuevoLog = False
    
    N = DateDiff("h", CreacionLog, Now)
    If N > 24 Then CrearNuevoLog = True
        
eCrearNuevoLog:
    Err.Clear
End Function
