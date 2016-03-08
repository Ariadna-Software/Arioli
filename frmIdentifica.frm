VERSION 5.00
Begin VB.Form frmIdentifica 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   4920
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   4920
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ariadna Software"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   4
      Left            =   3240
      TabIndex        =   6
      Top             =   5160
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ariadna Software"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   3240
      TabIndex        =   5
      Top             =   5400
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cargando ....."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   4920
      TabIndex        =   4
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   4920
      TabIndex        =   3
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   4920
      TabIndex        =   2
      Top             =   360
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   5655
      Left            =   0
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "frmIdentifica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PrimeraVez As Boolean
Dim T1 As Single

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        
        Espera 0.5
        Me.Refresh
        
        'Vemos datos de configAriges.ini
        Set vConfig = New Configuracion
        If vConfig.Leer = 1 Then
             vConfig.SERVER = InputBox("Servidor: ")
             vConfig.User = InputBox("Usuario: ")
             vConfig.password = InputBox("Password: ")
'             vConfig.Integraciones = InputBox("Path integraciones: ")
             vConfig.Grabar
             MsgBox "Reinicie AriGes", vbCritical
             End
             Exit Sub
        End If
        
         'Abrimos conexion para comprobar el usuario
         'Luego, en funcion del nivel de usuario que tenga cerraremos la conexion
         'y la abriremos con usuario-codigo ajustado a su nivel
         If AbrirConexionUsuarios() = False Then
             MsgBox "La apliacación no puede continuar sin acceso a los datos. ", vbCritical
             End
        End If
         
         'La llave
'         '### Descomentar
'         Load frmLLave
'         If Not frmLLave.ActiveLock1.RegisteredUser Then
'             'No ESTA REGISTRADO
'             frmLLave.Show vbModal
'         Else
'             Unload frmLLave
'         End If
         '###
         
         
        'Gestionar el nombre del PC para la asignacion de PC en el entorno de red
        GestionaPC
        
         'Leemos el ultimo usuario conectado
         NumeroEmpresaMemorizar True
         
         T1 = T1 + 2.5 - Timer
         If T1 > 0 Then Espera T1

         
         PonerVisible True
         If Text1(0).Text <> "" Then
            PonerFoco Text1(1)
         Else
            PonerFoco Text1(0)
         End If
        
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    PonerVisible False
    T1 = Timer
    Text1(0).Text = ""
    Text1(1).Text = ""
    Label1(4).Caption = "Ver. " & App.Major & "." & App.Minor & "." & App.Revision & ""
    PrimeraVez = True
    CargaImagen
End Sub


Private Sub CargaImagen()
    On Error Resume Next
    Me.Image1 = LoadPicture(App.Path & "\fondo4.dat")
    If Err.Number <> 0 Then
        MsgBox Err.Description & vbCrLf & vbCrLf & "Error cargando", vbCritical
        Set Conn = Nothing
        End
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    NumeroEmpresaMemorizar False
End Sub













Private Sub Text1_GotFocus(Index As Integer)
    With Text1(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    Else
        If KeyAscii = 27 Then
            CadenaDesdeOtroForm = ""
            Unload Me
        End If
    End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).Text = Trim(Text1(Index).Text)
    
    'Comprobamos si los dos estan con datos
    If Text1(0).Text <> "" And Text1(1).Text <> "" Then
        'Probar conexion usuario
        Validar
    End If
        
End Sub



Private Sub Validar()
Dim NuevoUsu As Usuario
Dim OK As Byte

    'Validaremos el usuario y despues el password
    Set vUsu = New Usuario
    
    If vUsu.Leer(Text1(0).Text) = 0 Then
        'Con exito
        If vUsu.PasswdPROPIO = Text1(1).Text Then
            OK = 0
        Else
            OK = 1
        End If

    Else
        OK = 2
    End If
    
    If OK <> 0 Then
        MsgBox "Usuario-Clave Incorrecto", vbExclamation

            Text1(1).Text = ""
            PonerFoco Text1(0)
    Else
        'OK
        CadenaDesdeOtroForm = "OK"
        Unload Me
    End If

End Sub


Private Sub PonerVisible(visible As Boolean)
    Label1(2).visible = Not visible  'Cargando
    Text1(0).visible = visible
    Text1(1).visible = visible
    Label1(0).visible = visible
    Label1(1).visible = visible
End Sub




'Lo que haremos aqui es ver, o guardar, el ultimo numero de empresa
'a la que ha entrado, y el usuario
Private Sub NumeroEmpresaMemorizar(Leer As Boolean)
Dim NF As Integer
Dim cad As String
On Error GoTo ENumeroEmpresaMemorizar


        
    cad = App.Path & "\ultusu.dat"
    If Leer Then
        If Dir(cad) <> "" Then
            NF = FreeFile
            Open cad For Input As #NF
            Line Input #NF, cad
            Close #NF
            cad = Trim(cad)
            
                'El primer pipe es el usuario
                Text1(0).Text = cad
    
        End If
    Else 'Escribir
        NF = FreeFile
        Open cad For Output As #NF
        cad = Text1(0).Text
        Print #NF, cad
        Close #NF
    End If
ENumeroEmpresaMemorizar:
    Err.Clear
End Sub
