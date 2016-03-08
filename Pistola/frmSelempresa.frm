VERSION 5.00
Begin VB.Form frmSelempresa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccionar "
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3570
   Icon            =   "frmSelempresa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   3570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Seleccionar"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3660
      Width           =   1935
   End
   Begin VB.CommandButton cmdCAncel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   3660
      Width           =   975
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3360
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "(c) Ariadna Software"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   4080
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "EMPRESAS"
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
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "frmSelempresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private EmpresaMemorizada As Integer

Private Sub cmdAceptar_Click()
Dim i As Integer
Dim B As Integer

    'cmdAceptar_Click
    B = -1
    For i = 0 To List1.ListCount - 1
       If List1.Selected(i) Then
            CadenaDesdeOtroForm = List1.List(i)
            If B >= 0 Then B = 3001
            B = i
        End If
    Next
    
    If B < 0 Then
        MsgBox "Selecione alguna empresa", vbExclamation
    Else
        If B > 3000 Then
            MsgBox "Raro, dos empresas seleccionadas", vbCritical
        Else
            
            MemorizarEmpresa False, B
            Unload Me
        End If
    End If
            
End Sub

Private Sub cmdCAncel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim i As Integer

     Me.Height = 320 * Screen.TwipsPerPixelX
     Me.Width = 240 * Screen.TwipsPerPixelX

    Do
        i = InStr(1, CadenaDesdeOtroForm, "|")
        If i = 0 Then
            CadenaDesdeOtroForm = ""
        Else
            List1.AddItem Mid(CadenaDesdeOtroForm, 1, i - 1)
            CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, i + 1)
            'If List1.ListCount = 1 Then List1.Selected(0) = True
        End If
    Loop Until i = 0
    EmpresaMemorizada = 0
    MemorizarEmpresa True, 0
    If List1.ListCount >= 0 Then
        On Error Resume Next
        List1.Selected(EmpresaMemorizada) = True
        If Err.Number <> 0 Then Err.Clear
    End If
End Sub

Private Sub MemorizarEmpresa(Leer As Boolean, Empresa As Integer)

Dim NF As Integer
Dim Cad As String

    On Error GoTo EM

    Cad = App.Path & "\ultemp.dat"
    NF = FreeFile
    If Leer Then
       If Dir(Cad, vbArchive) <> "" Then
            
            Open Cad For Input As #NF
            Line Input #NF, Cad
            Close #NF
            Cad = Trim(Cad)
            If Cad = "" Then Cad = "0"
            If Not IsNumeric(Cad) Then Cad = "0"
            
            EmpresaMemorizada = CInt(Cad)
            
        End If
    Else
        If EmpresaMemorizada <> Empresa Then
            Open Cad For Output As #NF
            Print #NF, Empresa
            Close #NF
        End If
    End If
        
    Exit Sub
EM:
    Err.Clear
    
End Sub


Private Sub List1_DblClick()
    cmdAceptar_Click
End Sub
