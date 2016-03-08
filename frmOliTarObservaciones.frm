VERSION 5.00
Begin VB.Form frmOliTarObservaciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Observaciones  TO"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7380
   Icon            =   "frmOliTarObservaciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Index           =   1
      Left            =   5160
      TabIndex        =   1
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   375
      Index           =   0
      Left            =   6240
      TabIndex        =   2
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   4335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmOliTarObservaciones.frx":000C
      Top             =   240
      Width           =   7095
   End
End
Attribute VB_Name = "frmOliTarObservaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event Observaciones(Valor As String)



Private Sub Command1_Click(Index As Integer)
    If Index = 1 Then RaiseEvent Observaciones(Text1.Text)
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = frmppal.Icon
    
End Sub
