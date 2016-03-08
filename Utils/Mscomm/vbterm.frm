VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "Mscomm32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTerminal 
   Caption         =   "Terminal de Visual Basic"
   ClientHeight    =   4935
   ClientLeft      =   2940
   ClientTop       =   2055
   ClientWidth     =   7155
   ForeColor       =   &H00000000&
   Icon            =   "vbterm.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   7155
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   480
      Width           =   7215
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2640
         TabIndex        =   8
         Text            =   "1"
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   255
         Left            =   6600
         Picture         =   "vbterm.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3840
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "Texto"
         Height          =   195
         Index           =   1
         Left            =   3240
         TabIndex        =   10
         Top             =   240
         Width           =   405
      End
      Begin VB.Label Label2 
         Caption         =   "Columna"
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Envio HITACHI"
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
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   1455
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
      TabIndex        =   3
      Top             =   1920
      Width           =   5790
   End
   Begin MSComctlLib.Toolbar tbrToolBar 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "OpenLogFile"
            Description     =   "Open Log File..."
            Object.ToolTipText     =   "Abrir archivo de registro..."
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "CloseLogFile"
            Description     =   "Close Log File"
            Object.ToolTipText     =   "Cerrar archivo de registro"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DialPhoneNumber"
            Description     =   "Dial Phone Number..."
            Object.ToolTipText     =   "Marcar número de teléfono"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "HangUpPhone"
            Description     =   "Hang Up Phone"
            Object.ToolTipText     =   "Colgar teléfono"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Propiedades"
            Description     =   "Properties..."
            Object.ToolTipText     =   "Propiedades..."
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "TransmitTextFile"
            Description     =   "Transmit Text File..."
            Object.ToolTipText     =   "Transmitir archivo de texto"
            ImageIndex      =   6
         EndProperty
      EndProperty
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   240
         Left            =   4000
         TabIndex        =   2
         Top             =   75
         Width           =   240
         Begin VB.Image imgConnected 
            Height          =   240
            Left            =   0
            Picture         =   "vbterm.frx":6B5C
            Stretch         =   -1  'True
            ToolTipText     =   "Alternar puerto"
            Top             =   0
            Width           =   240
         End
         Begin VB.Image imgNotConnected 
            Height          =   240
            Left            =   0
            Picture         =   "vbterm.frx":6CA6
            Stretch         =   -1  'True
            ToolTipText     =   "Alternar puerto"
            Top             =   0
            Width           =   240
         End
      End
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
            Picture         =   "vbterm.frx":6DF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbterm.frx":710A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbterm.frx":7424
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbterm.frx":773E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbterm.frx":7A58
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbterm.frx":7D72
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuOpenLog 
         Caption         =   "&Abrir archivo de registro..."
      End
      Begin VB.Menu mnuCloseLog 
         Caption         =   "&Cerrar archivo de registro"
         Enabled         =   0   'False
      End
      Begin VB.Menu M3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSendText 
         Caption         =   "&Transmitir archivo de texto"
         Enabled         =   0   'False
      End
      Begin VB.Menu Bar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuPort 
      Caption         =   "&Puerto"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Puerto abierto"
      End
      Begin VB.Menu MBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "Pr&opiedades..."
      End
   End
   Begin VB.Menu mnuMSComm 
      Caption         =   "&MSComm"
      Begin VB.Menu mnuInputLen 
         Caption         =   "&InputLen..."
      End
      Begin VB.Menu mnuRThreshold 
         Caption         =   "&RThreshold..."
      End
      Begin VB.Menu mnuSThreshold 
         Caption         =   "&SThreshold..."
      End
      Begin VB.Menu mnuParRep 
         Caption         =   "P&arityReplace..."
      End
      Begin VB.Menu mnuDTREnable 
         Caption         =   "&DTREnable"
      End
      Begin VB.Menu Bar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHCD 
         Caption         =   "&CDHolding..."
      End
      Begin VB.Menu mnuHCTS 
         Caption         =   "CTSH&olding..."
      End
      Begin VB.Menu mnuHDSR 
         Caption         =   "DSRHo&lding..."
      End
   End
   Begin VB.Menu mnuCall 
      Caption         =   "&Llamar"
      Begin VB.Menu mnuDial 
         Caption         =   "&Marcar número de teléfono"
      End
      Begin VB.Menu mnuHangUp 
         Caption         =   "&Colgar teléfono"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnDavid 
      Caption         =   "David"
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
Dim Temp As String      ' Cadena auxiliar.
Dim hLogFile As Integer ' Controlador de archivo de registro abierto.
Dim StartTime As Date   ' Almacena la hora de inicio del cronómetro del puerto

Private Sub Command1_Click()
Dim Cad As String
    Cad = ""
    If Not MSComm1.PortOpen Then Cad = "Puerto cerrado"
    If Text2.Text = "" Then
        Cad = Cad & vbCrLf & "Columna vacia"
    Else
        If Not IsNumeric(Text2.Text) Then Cad = Cad & vbCrLf & "Columna debe ser numérica"
    End If
    If Cad <> "" Then
        MsgBox Cad, vbExclamation
        Exit Sub
    End If
    EnviaHitachi
End Sub

Private Sub EnviaHitachi()
Const STX = 2
Const ETX = 3
Const EOT = 4
Const SOH = 1
Const ENQ = 5
Const ACK = 6
Const NAK = 21
Const ETB = 23
Const ESC = 27

Dim datos As String
Dim dato As String

'Si tengo que meter el STX..ETX
'Envia la columna y el text
MSComm1.Output = Chr(STX) + Text2.Text + Text1.Text + Chr(ETX)
    
    
End Sub

Private Sub Form_Load()
    Dim CommPort As String, Handshaking As String, Settings As String
        
    On Error Resume Next
    
    ' Establece el color predeterminado del terminal
    txtTerm.SelLength = Len(txtTerm)
    txtTerm.SelText = ""
    txtTerm.ForeColor = vbBlue
       
    ' Establece el título
    App.Title = "Terminal de Visual Basic"
    
    ' Establece la luz indicadora de estado
    imgNotConnected.ZOrder
       
    ' Centra el formulario
    frmTerminal.Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    
    ' Carga la configuración del registro
    
    Settings = GetSetting(App.Title, "Properties", "Settings", "") ' frmTerminal.MSComm1.Settings]\
    If Settings <> "" Then
        MSComm1.Settings = Settings
        If Err Then
            MsgBox Error$, 48
            Exit Sub
        End If
    End If
    
    CommPort = GetSetting(App.Title, "Properties", "CommPort", "") ' frmTerminal.MSComm1.CommPort
    If CommPort <> "" Then MSComm1.CommPort = CommPort
    
    Handshaking = GetSetting(App.Title, "Properties", "Handshaking", "") 'frmTerminal.MSComm1.Handshaking
    If Handshaking <> "" Then
        MSComm1.Handshaking = Handshaking
        If Err Then
            MsgBox Error$, 48
            Exit Sub
        End If
    End If
    
    Echo = GetSetting(App.Title, "Properties", "Echo", "") ' Echo
    On Error GoTo 0

End Sub

Private Sub Form_Resize()
Dim Fr As Integer
    Frame2.Top = tbrToolBar.Height + 6
    Frame2.Width = frmTerminal.ScaleWidth - 120
    Fr = Frame2.Height + 60
   ' Cambia el tamaño del control Term (ventana)
   txtTerm.Move 0, tbrToolBar.Height + Fr, frmTerminal.ScaleWidth, frmTerminal.ScaleHeight - sbrStatus.Height - tbrToolBar.Height - Fr
   
   ' Sitúa la luz indicadora de estado
   Frame1.Left = ScaleWidth - Frame1.Width * 1.5
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim Counter As Long

    If MSComm1.PortOpen Then
       ' Espera 10 segundos para transmitir los datos.
       Counter = Timer + 10
       Do While MSComm1.OutBufferCount
          Ret = DoEvents()
          If Timer > Counter Then
             Select Case MsgBox("Imposible enviar los datos", 34)
                ' Cancelar.
                Case 3
                   Cancel = True
                   Exit Sub
                ' Reintentar.
                Case 4
                   Counter = Timer + 10
                ' Ignorar.
                Case 5
                   Exit Do
             End Select
          End If
       Loop

       MSComm1.PortOpen = 0
    End If

    ' Si el archivo de registro está abierto, vuelca y lo cierra.
    If hLogFile Then mnuCloseLog_Click
    End
End Sub

Private Sub imgConnected_Click()
    ' Llama a la rutina mnuOpen_Click para alternar entre conectar y desconectar.
    Call mnuOpen_Click
End Sub

Private Sub imgNotConnected_Click()
    ' Llama a la rutina mnuOpen_Click para alternar entre conectar y desconectar.
    Call mnuOpen_Click
End Sub

Private Sub mnDavid_Click()
Dim i As Integer
Dim Cad As String
  'Estos dos FUNCIONA
   ' MSComm1.Output = Chr(27) & "[8q" & Chr(27) & "[3q" & Chr(27) & "[7q" & Chr(13)
  ' MSComm1.Output = Chr(27) & "[7m" & Chr(13) 'reverse mode
  ' MSComm1.Output = Chr(27) & "[0m" & Chr(13) 'normal mode
  
  ' MSComm1.Output = Chr(27) & "[2J" & Chr(13)
  '  For I = 0 To 2
             
   
           
            
            'MSComm1.Output = Chr(27) & "E"
            'MSComm1.Output = "Datos: " & I & Chr(13)
           'mSComm1.Output = Chr(27) & "E"
           Cad = Chr(27) & "E" & "Datos:1 " & Chr(27) & "E" & "Dat:2 " & Chr(27) & "E" & "Datos:3 " & Chr(27) & "E" & "Dat:4 " & Chr(13)
            MSComm1.Output = Cad
  ' Next I
    ShowData txtTerm, "Lanzada secuencia" & vbCrLf
End Sub

Private Sub mnuCloseLog_Click()
    ' Cierra el archivo de registro.
    Close hLogFile
    hLogFile = 0
    mnuOpenLog.Enabled = True
    tbrToolBar.Buttons("OpenLogFile").Enabled = True
    mnuCloseLog.Enabled = False
    tbrToolBar.Buttons("CloseLogFile").Enabled = False
    frmTerminal.Caption = "Terminal de Visual Basic"
End Sub

Private Sub mnuDial_Click()
    On Local Error Resume Next
    Static Num As String
    
    Num = "1-206-936-6735" ' Este es el número de MSDN
    
    ' Obtiene un número del usuario.
    Num = InputBox$("Escriba el número de teléfono:", "Marcar número", Num)
    If Num = "" Then Exit Sub
    
    ' Abre el puerto si no está abierto ya.
    If Not MSComm1.PortOpen Then
       mnuOpen_Click
       If Err Then Exit Sub
    End If
      
    ' Habilita el botón de colgar y el elemento de menú correspondiente.
    mnuHangUp.Enabled = True
    tbrToolBar.Buttons("HangUpPhone").Enabled = True
              
    ' Marca el número.
    MSComm1.Output = "ATDT" & Num & vbCrLf
    
    ' Inicia el cronómetro del puerto.
    StartTiming
End Sub

' Alterna la propiedad DTREnabled.
Private Sub mnuDTREnable_Click()
    ' Alterna la propiedad DTREnable
    MSComm1.DTREnable = Not MSComm1.DTREnable
    mnuDTREnable.Checked = MSComm1.DTREnable
End Sub


Private Sub mnuFileExit_Click()
    ' Utiliza Form_Unload, ya que tiene código para comprobar si hay datos
    ' no enviados y si hay un archivo de registro abierto.
    Form_Unload Ret
End Sub



' Alterna la propiedad DTREnable para colgar la línea.
Private Sub mnuHangup_Click()
    On Error Resume Next
    
    MSComm1.Output = "ATH"      ' Envía la cadena de colgar
    Ret = MSComm1.DTREnable     ' Guarda el valor actual.
    MSComm1.DTREnable = True    ' Activa DTR.
    MSComm1.DTREnable = False   ' Desactiva DTR.
    MSComm1.DTREnable = Ret     ' Restablece el valor anterior.
    mnuHangUp.Enabled = False
    tbrToolBar.Buttons("HangUpPhone").Enabled = False
    
    ' Si el puerto continúa abierto, lo cierra
    If MSComm1.PortOpen Then MSComm1.PortOpen = False
    
    ' Notifica el error al usuario
    If Err Then MsgBox Error$, 48
    
    mnuSendText.Enabled = False
    tbrToolBar.Buttons("TransmitTextFile").Enabled = False
    mnuHangUp.Enabled = False
    tbrToolBar.Buttons("HangUpPhone").Enabled = False
    mnuDial.Enabled = True
    tbrToolBar.Buttons("DialPhoneNumber").Enabled = True
    sbrStatus.Panels("Settings").Text = "Valores:"
    
    ' Apaga la luz indicadora y quita la marca al menú
    mnuOpen.Checked = False
    imgNotConnected.ZOrder
            
    ' Detiene el cronómetro del puerto
    StopTiming
    sbrStatus.Panels("Status").Text = "Estado:"
    On Error GoTo 0
End Sub

' Muestra el valor de la propiedad CDHolding.
Private Sub mnuHCD_Click()
    If MSComm1.CDHolding Then
        Temp = "True"
    Else
        Temp = "False"
    End If
    MsgBox "CDHolding = " + Temp
End Sub

' Muestra el valor de la propiedad CTSHolding.
Private Sub mnuHCTS_Click()
    If MSComm1.CTSHolding Then
        Temp = "True"
    Else
        Temp = "False"
    End If
    MsgBox "CTSHolding = " + Temp
End Sub

' Muestra el valor de la propiedad DSRHolding.
Private Sub mnuHDSR_Click()
    If MSComm1.DSRHolding Then
        Temp = "True"
    Else
        Temp = "False"
    End If
    MsgBox "DSRHolding = " + Temp
End Sub

' Este procedimiento establece la propiedad InputLen, que determina
' el número de bytes de datos leídos cada vez que se usa Input para
' recuperar datos del búfer de entrada.
' Al establecer 0 en InputLen se especifica que debe leerse todo el
' contenido del búfer.
Private Sub mnuInputLen_Click()
    On Error Resume Next

    Temp = InputBox$("Escriba nuevo InputLen:", "InputLen", Str$(MSComm1.InputLen))
    If Len(Temp) Then
        MSComm1.InputLen = Val(Temp)
        If Err Then MsgBox Error$, 48
    End If
End Sub

Private Sub mnuProperties_Click()
  ' Muestra el formulario de propiedades de CommPort
  frmProps.Show vbModal
  
End Sub

' Alterna el estado del puerto (abierto o cerrado).
Private Sub mnuOpen_Click()
    On Error Resume Next
    Dim OpenFlag

    MSComm1.PortOpen = Not MSComm1.PortOpen
    If Err Then MsgBox Error$, 48
    
    OpenFlag = MSComm1.PortOpen
    
    mnuOpen.Checked = OpenFlag
    mnuSendText.Enabled = OpenFlag
    tbrToolBar.Buttons("TransmitTextFile").Enabled = OpenFlag
        
    If MSComm1.PortOpen Then
        ' Habilita el botón de marcar y el elemento de menú asociado
        mnuDial.Enabled = True
        tbrToolBar.Buttons("DialPhoneNumber").Enabled = True
        
        ' Habilita el botón de colgar y el elemento de menú asociado
        mnuHangUp.Enabled = True
        tbrToolBar.Buttons("HangUpPhone").Enabled = True
        
        imgConnected.ZOrder
        sbrStatus.Panels("Settings").Text = "Valores:" & MSComm1.Settings
        StartTiming
    Else
        ' Habilita el botón de marcar y el elemento de menú asociado
        mnuDial.Enabled = True
        tbrToolBar.Buttons("DialPhoneNumber").Enabled = True
        
        ' deshabilita el botón de colgar y el elemento de menú asociado
        mnuHangUp.Enabled = False
        tbrToolBar.Buttons("HangUpPhone").Enabled = False
        
        imgNotConnected.ZOrder
        sbrStatus.Panels("Settings").Text = "Valores:"
        StopTiming
    End If
    
End Sub

Private Sub mnuOpenLog_Click()
   Dim replace
   On Error Resume Next
   OpenLog.Flags = cdlOFNHideReadOnly Or cdlOFNExplorer
   OpenLog.CancelError = True
      
   ' Obtiene del usuario el nombre de archivo largo.
   OpenLog.DialogTitle = "Abrir archivo de registro de comunicaciones"
   OpenLog.Filter = "Archivos de registro (*.LOG)|*.log|Todos los archivos (*.*)|*.*"
   
   Do
      OpenLog.FileName = ""
      OpenLog.ShowOpen
      If Err = cdlCancel Then Exit Sub
      Temp = OpenLog.FileName

      ' Si el archivo ya existe, pregunta al usuario si desea
      ' sobrescribirlo o agregarle datos.
      Ret = Len(Dir$(Temp))
      If Err Then
         MsgBox Error$, 48
         Exit Sub
      End If
      If Ret Then
         replace = MsgBox("Reemplazar el archivo - " + Temp + "?", 35)
      Else
         replace = 0
      End If
   Loop While replace = 2

   ' El usuario ha hecho clic en el botón Sí, así que elimina el archivo.
   If replace = 6 Then
      Kill Temp
      If Err Then
         MsgBox Error$, 48
         Exit Sub
      End If
   End If

   ' Abre el archivo de registro.
   hLogFile = FreeFile
   Open Temp For Binary Access Write As hLogFile
   If Err Then
      MsgBox Error$, 48
      Close hLogFile
      hLogFile = 0
      Exit Sub
   Else
      ' Va al final del archivo para poder agregarle nuevos datos.
      Seek hLogFile, LOF(hLogFile) + 1
   End If

   frmTerminal.Caption = "Terminal de Visual Basic - " + OpenLog.FileTitle
   mnuOpenLog.Enabled = False
   tbrToolBar.Buttons("OpenLogFile").Enabled = False
   mnuCloseLog.Enabled = True
   tbrToolBar.Buttons("CloseLogFile").Enabled = True
End Sub

' Este procedimiento establece la propiedad ParityReplace, que
' contiene el carácter que reemplazará a los caracteres
' incorrectos recibidos a causa de un error de paridad.
Private Sub mnuParRep_Click()
    On Error Resume Next

    Temp = InputBox$("Escriba el carácter que desea reemplazar", "ParityReplace", frmTerminal.MSComm1.ParityReplace)
    frmTerminal.MSComm1.ParityReplace = Left$(Temp, 1)
    If Err Then MsgBox Error$, 48
End Sub

' Este procedimiento establece la propiedad RThreshlold, que
' determina el número de bytes que pueden llegar al búfer de
' recepción antes de disparar el evento OnComm y de que se
' establezca comEvReceive en la propiedad CommEvent.
Private Sub mnuRThreshold_Click()
    On Error Resume Next
    
    Temp = InputBox$("Introduzca el nuevo RThreshold:", "RThreshold", Str$(MSComm1.RThreshold))
    If Len(Temp) Then
        MSComm1.RThreshold = Val(Temp)
        If Err Then MsgBox Error$, 48
    End If

End Sub

' El evento OnComm se usa para interceptar eventos y errores de comunicaciones.
Private Static Sub MSComm1_OnComm()
    Dim EVMsg$
    Dim ERMsg$
    Dim aux As String
    ' Bifurca según la propiedad CommEvent.
    Select Case MSComm1.CommEvent
        ' Mensajes de evento.
        Case comEvReceive
            Dim Buffer As Variant
            Buffer = MSComm1.Input
            
            aux = StrConv(Buffer, vbUnicode)
            Debug.Print "Recibir - " & aux
            ShowData txtTerm, aux
            If InStr(1, aux, ">") > 0 Then mnDavid_Click
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
        ' Muestra los mensajes de evento en la barra de estado.
        sbrStatus.Panels("Status").Text = "Estado:" & EVMsg$
                
        ' Activa el cronómetro para que el mensaje de la barra
        ' de estado se borre después de dos segundos.
        Timer2.Enabled = True
        
    ElseIf Len(ERMsg$) Then
        ' Muestra los mensajes de evento en la barra de estado.
        sbrStatus.Panels("Status").Text = "Estado:" & ERMsg$
        
        ' Muestra los mensajes de error en un cuadro de alerta.
        Beep
        Ret = MsgBox(ERMsg$, 1, "Haga clic en Cancelar para salir, clic en Aceptar para ignorar.")
        
        ' Si el usuario hace clic en Cancelar (2)...
        If Ret = 2 Then
            MSComm1.PortOpen = False    ' Cierra el puerto y sale.
        End If
        
        ' Activa el cronómetro para que el mensaje de la barra
        ' de estado se borre después de dos segundos.
        Timer2.Enabled = True
    End If
End Sub

Private Sub mnuSendText_Click()
   Dim hSend, BSize, LF&
   
   On Error Resume Next
   
   mnuSendText.Enabled = False
   tbrToolBar.Buttons("TransmitTextFile").Enabled = False
   
   ' Obtiene del usuario el nombre del archivo de texto.
   OpenLog.DialogTitle = "Enviar archivo de texto"
   OpenLog.Filter = "Archivos de texto (*.TXT)|*.txt|Todos los archivos (*.*)|*.*"
   Do
      OpenLog.CancelError = True
      OpenLog.FileName = ""
      OpenLog.ShowOpen
      If Err = cdlCancel Then
        mnuSendText.Enabled = True
        tbrToolBar.Buttons("TransmitTextFile").Enabled = True
        Exit Sub
      End If
      Temp = OpenLog.FileName

      ' Si el archivo no existe, vuelve.
      Ret = Len(Dir$(Temp))
      If Err Then
         MsgBox Error$, 48
         mnuSendText.Enabled = True
         tbrToolBar.Buttons("TransmitTextFile").Enabled = True
         Exit Sub
      End If
      If Ret Then
         Exit Do
      Else
         MsgBox Temp + " no encontrado.", 48
      End If
   Loop

   ' Abre el archivo de registro.
   hSend = FreeFile
   Open Temp For Binary Access Read As hSend
   If Err Then
      MsgBox Error$, 48
   Else
      ' Muestra el cuadro de diálogo Cancelar.
      CancelSend = False
      frmCancelSend.Label1.Caption = "Transmitiendo archivo de texto - " + Temp
      frmCancelSend.Show
      
      ' Lee el archivo en bloques del tamaño del búfer de transmisión.
      BSize = MSComm1.OutBufferSize
      LF& = LOF(hSend)
      Do Until EOF(hSend) Or CancelSend
         ' No lee demasiado al final.
         If LF& - Loc(hSend) <= BSize Then
            BSize = LF& - Loc(hSend) + 1
         End If
      
         ' Lee un bloque de datos.
         Temp = Space$(BSize)
         Get hSend, , Temp
      
         ' Transmite el bloque.
         MSComm1.Output = Temp
         If Err Then
            MsgBox Error$, 48
            Exit Do
         End If
      
         ' Espera a que se envíen todos los datos.
         Do
            Ret = DoEvents()
         Loop Until MSComm1.OutBufferCount = 0 Or CancelSend
      Loop
   End If
   
   Close hSend
   mnuSendText.Enabled = True
   tbrToolBar.Buttons("TransmitTextFile").Enabled = True
   CancelSend = True
   frmCancelSend.Hide
End Sub


' Este procedimiento establece la propiedad SThreshold, que
' determina el número máximo de caracteres que deben estar
' esperando en el búfer de salida para que se establezca
' comEvSend en la propiedad CommEvent y se dispare el evento
' OnComm.
Private Sub mnuSThreshold_Click()
    On Error Resume Next
    
    Temp = InputBox$("Escriba el nuevo valor de SThreshold", "SThreshold", Str$(MSComm1.SThreshold))
    If Len(Temp) Then
        MSComm1.SThreshold = Val(Temp)
        If Err Then MsgBox Error$, 48
    End If
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

    ' Elimina los avances de línea.
    Do
       i = InStr(Data, Chr(10))
       If i Then
          Data = Left$(Data, i - 1) & Mid$(Data, i + 1)
       End If
    Loop While i

    ' Se asegura de que todos los retornos de carro tengan un
    ' avance de línea.
    i = 1
    Do
       i = InStr(i, Data, Chr(13))
       If i Then
          Data = Left$(Data, i) & Chr(10) & Mid$(Data, i + 1)
          i = i + 1
       End If
    Loop While i

    ' Agrega los datos filtrados a la propiedad SelText.
    Term.SelText = Data
  
    ' Registra los datos en un archivo si así se solicita.
    If hLogFile Then
       i = 2
       Do
          Err = 0
          Put hLogFile, , Data
          If Err Then
             i = MsgBox(Error$, 21)
             If i = 2 Then
                mnuCloseLog_Click
             End If
          End If
       Loop While i <> 2
    End If
    Term.SelStart = Len(Term.Text)
Exit Sub

Handler:
    MsgBox Error$
    Resume Next
End Sub

Private Sub Timer2_Timer()
sbrStatus.Panels("Status").Text = "Estado:"
Timer2.Enabled = False

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




Private Sub tbrToolBar_ButtonClick(ByVal Button As MSComCtlLib.Button)
Select Case Button.Key
Case "OpenLogFile"
    Call mnuOpenLog_Click
Case "CloseLogFile"
    Call mnuCloseLog_Click
Case "DialPhoneNumber"
    Call mnuDial_Click
Case "HangUpPhone"
    Call mnuHangup_Click
Case "Propiedades"
    Call mnuProperties_Click
Case "TransmitTextFile"
    Call mnuSendText_Click
End Select
End Sub

Private Sub Timer1_Timer()
    ' Muestra la hora de conexión
    sbrStatus.Panels("ConnectTime").Text = Format(Now - StartTime, "hh:nn:ss") & " "
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
