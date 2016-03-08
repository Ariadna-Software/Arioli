VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Identificacion"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3510
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      Picture         =   "frmLogin.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Entrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   2880
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1800
      Width           =   3015
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   240
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "(c) Ariadna Software"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3840
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cad As String
Dim Pvez As Boolean

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub Command1_Click()
Dim VolverA_Guardar As Boolean
    If Combo1.Text = "" Or Text1.Text = "" Then Exit Sub
    If vUsu.Leer(Combo1.Text) = 1 Then
        MsgBox "Usuario incorrecto", vbExclamation
        Exit Sub
    End If
    If vUsu.PasswdPROPIO <> Text1.Text Then
        MsgBox "Error password", vbExclamation
        PonerFoco Text1
        Exit Sub
    End If
    
    cad = "|" & vUsu.Login & "|"
    VolverA_Guardar = False
    If InStr(1, Combo1.Tag, cad) = 0 Then
        VolverA_Guardar = True
        Combo1.Tag = Combo1.Tag & Mid(cad, 2)
    Else
        If Text1.Tag <> Combo1.Text Then VolverA_Guardar = True
        
        
    End If
    If VolverA_Guardar Then FicheroLogin False
    
    
    
    'AHora veremos si hay varias empresas disponibles
    Dim OK As Boolean
    OK = ComprobarEmpresasAsignadas
    Unload Me
    If OK Then frmPist1.Show vbModal
    
    
End Sub



Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If Pvez Then
        If Me.Combo1.ListCount > 0 Then PonerFoco Text1
    End If
End Sub

Private Sub Form_Load()
    Pvez = True
    ProcesarUsuariosLogados
    Text1.Text = ""
End Sub



Private Sub ProcesarUsuariosLogados()
    On Error GoTo eProcesarUsuariosLogados
    Combo1.Clear
    Combo1.Tag = "|"
    FicheroLogin True

eProcesarUsuariosLogados:
    If Err.Number <> 0 Then Err.Clear
    
End Sub
Private Sub FicheroLogin(Leer As Boolean)
Dim NF As Integer

    NF = FreeFile
    cad = App.Path & "\ultusu2.dat"
    If Leer Then
        If Dir(cad, vbArchive) = "" Then Exit Sub
        Open cad For Input As #NF
        'Primera linea el ultimo usuario
        Line Input #NF, cad
        Combo1.Text = cad
        Text1.Tag = cad
        
        'Segunda linea todos los usuarios
        Line Input #NF, cad
        Close #NF
        Combo1.Tag = cad
        If cad <> "" Then cad = Mid(cad, 2)
        While cad <> ""
            NF = InStr(1, cad, "|")
            If NF = 0 Then
                cad = ""
            Else
                Combo1.AddItem Mid(cad, 1, NF - 1)
                cad = Mid(cad, NF + 1)
            End If
        Wend
        
    Else
        Open cad For Output As #NF
        Print #NF, Combo1.Text
        
        Print #NF, Combo1.Tag
      
    End If
End Sub

Private Sub Text1_GotFocus()
    ObtenerFoco Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub


Private Function ComprobarEmpresasAsignadas() As Boolean
Dim Hay As Integer
    
    CadenaDesdeOtroForm = "select * from usuarios.empresasarioli where not codempre in (select codempre from usuarios.usuarioempresasarioli where usuarioempresasarioli.codusu = " & vUsu.Codigo Mod 1000 & ")"
    Set MiRsAux = New ADODB.Recordset
    MiRsAux.Open CadenaDesdeOtroForm, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CadenaDesdeOtroForm = ""
    cad = ""
    While Not MiRsAux.EOF
        cad = cad & "X"
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & MiRsAux!codempre & " " & MiRsAux!nomempre & "|"
        MiRsAux.MoveNext
    Wend
    MiRsAux.Close
    
    If cad = "" Then
        MsgBox "Ninguna empresa seleccionable"
        ComprobarEmpresasAsignadas = False
        Exit Function
    End If
    
    
    Hay = Len(cad)
    If Hay = 1 Then
        'SOLO HAY UNA EMPRESA
        CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, 1, Len(CadenaDesdeOtroForm) - 1) 'quito el PIPE
    End If

    'Pantallita para seleccionar empresas
    If Hay > 1 Then frmSelempresa.Show vbModal
    
    If CadenaDesdeOtroForm = "" Then Exit Function 'LE HA DADO A CANCELAR
    
    Hay = InStr(1, CadenaDesdeOtroForm, " ")
    cad = Mid(CadenaDesdeOtroForm, 1, Hay - 1)
    CadenaDesdeOtroForm = ""
    Conn.Close
    If Not AbrirConexion(cad) Then Exit Function
    
    'Daremos por supuesto que tienen la conta en el mismo server que la gestion
    Set vParamAplic = New CParamAplic
    If vParamAplic.Leer = 1 Then End 'no puede dar error
    
    
    BDConta = vParamAplic.NumeroConta
    If BDConta <> "" Then
        'Tiene la conta. Voy a comprobar que funciona"
        BDConta = "conta" & BDConta
        cad = DevuelveDesdeBD(1, "fechaini", BDConta & ".parametros", "1", "1")
        If cad = "" Then
            cad = "Imposible accedera a contabilidad: " & BDConta
            MsgBox cad, vbExclamation
            BDConta = ""
        End If
    End If
    ComprobarEmpresasAsignadas = True
        
    
End Function
