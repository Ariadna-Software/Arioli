VERSION 5.00
Begin VB.Form frmPpUtuil 
   Caption         =   "Utilidades auxiliares"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5850
   Icon            =   "frmPpUtuil.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   5850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Ayuda demo"
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Pruebas pistola"
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Recalculo pesos Ficha tecnica"
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   3480
      Width           =   4695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Lotaje produccion"
      Enabled         =   0   'False
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   2760
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "S A L I R"
      Default         =   -1  'True
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   4200
      Width           =   4695
   End
   Begin VB.CommandButton cmdMovi 
      Caption         =   "Movimientos"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   2040
      Width           =   4695
   End
   Begin VB.CommandButton cmdMoval 
      Caption         =   "Stock"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1440
      Width           =   4695
   End
   Begin VB.CommandButton cmdRevInvne 
      Caption         =   "Revision inventario a fecha"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   4695
   End
End
Attribute VB_Name = "frmPpUtuil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdMoval_Click()
    frmStock.Show vbModal
End Sub

Private Sub cmdMovi_Click()
    frmComprobarMovClie.Show vbModal
End Sub

Private Sub cmdRevInvne_Click()
    frmRevisonInven.Show vbModal
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    frmLotaje.Show vbModal
End Sub

Private Sub Command3_Click()
    frmFichaTecnica.Show vbModal
End Sub

Private Sub Command4_Click()
    Me.Hide
    frmPruebasPedido.Show vbModal
    Me.Show
End Sub

Private Sub Command5_Click()

    'Me.Hide
    frmAyudaDemo.Show vbModal
    'Me.Show
End Sub
