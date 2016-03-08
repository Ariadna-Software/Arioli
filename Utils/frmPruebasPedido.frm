VERSION 5.00
Begin VB.Form frmPruebasPedido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4320
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   3510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   288
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   234
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Fondo del form"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3960
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
End
Attribute VB_Name = "frmPruebasPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SQL As String

Private Sub Form_Load()
 Me.Height = 320 * Screen.TwipsPerPixelX
 Me.Width = 240 * Screen.TwipsPerPixelX
 Text1.Text = ""
 Text2.Text = ""
 Label1.Caption = ""
End Sub

Private Sub Text1_GotFocus()
    ObtenerFoco Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub Text1_LostFocus()
Dim C As String

    Text1.Text = Trim(Text1.Text)
    C = ""
    SQL = ""
    If Text1.Text <> "" Then
        C = "nomartic"
        SQL = DevuelveDesdeBD("codartic", "sartic", "codigoea", Text1.Text, "T", C)
        If SQL = "" Then
            C = "*** NO EXISTE: " & Text1.Text
            Text1.Text = ""
            PonerFoco Text1
        End If
    End If
    Label1.Caption = C
    Label1.Tag = SQL
End Sub
