VERSION 5.00
Begin VB.Form frmProtector 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   6810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15855
   LinkTopic       =   "Form2"
   ScaleHeight     =   6810
   ScaleWidth      =   15855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   480
      Top             =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "23:59:59"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   3255
      Left            =   360
      TabIndex        =   0
      Top             =   2040
      Width           =   13215
   End
End
Attribute VB_Name = "frmProtector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim NumColores  As Integer
Dim colores As String
Dim CambiamosColor As Integer
Dim Movemos As Integer
Dim maxLeft As Integer
Dim maxTop As Integer


Private Sub Form_KeyPress(KeyAscii As Integer)
    Cerrar
End Sub

Private Sub Form_Load()
    NumColores = 7
    colores = "&HFFFFFF|&HC0C0FF|&HC0E0FF|&HC0FFC0|&HFFFFC0|&HFFC0C0|&HFFC0FF|"
    colores = Replace(colores, ",", "|")
    Label1.Caption = Format(Now, "hh:mm:ss")
    Label1.FontBold = True
    Label1.FontSize = 118
    CambiamosColor = 0
    Movemos = 0
  
    maxLeft = Screen.Width - Label1.Width - 240
    maxTop = Screen.Height - Label1.Height - 240
    
    Label1.Top = maxTop \ 2
    Label1.Left = maxLeft \ 2
    
    Movemos = 0
    Timer1.Enabled = True
    
End Sub

Private Sub Cerrar()
    'Exit Sub
    Timer1.Enabled = False
    Unload Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Cerrar
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ' Cerrar
End Sub

Private Sub Label1_Click()
    Cerrar
End Sub

Private Sub Aleatorio()
Dim J As Integer
Dim Aux As Long
        
        
    CambiamosColor = CambiamosColor + 1
    If CambiamosColor > 45 Then
        CambiamosColor = 0
        J = Int((NumColores * Rnd) + 1)
        If J > NumColores Then J = 1
    
    
        Me.Label1.ForeColor = Val(RecuperaValor(colores, J))
        Label1.Refresh
    End If
    
    
    Movemos = Movemos + 1
    If Movemos > 10 Then
        Movemos = 0
        
        J = Int((3 * Rnd) + 1)
        If J > 2 Then
            J = -1
        Else
            J = 1
        End If
        Label1.Left = Label1.Left + J
        If Label1.Left < 0 Then Label1.Left = 0
        If Label1.Left > maxLeft Then Label1.Left = Label1.Left - 3000
        
        
        J = Int((3 * Rnd) + 1)
        If J > 2 Then
            J = -1
        Else
            J = 1
        End If
         
        Label1.Top = Label1.Top + J
        If Label1.Top < 0 Then Label1.Top = 0
        If Label1.Top > maxTop Then Label1.Top = Label1.Top - 1500
        
        Label1.Refresh
        
       
    
    End If
End Sub

Private Sub Timer1_Timer()
    Label1.Caption = Format(Now, "hh:mm:ss")
    Aleatorio
End Sub
