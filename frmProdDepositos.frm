VERSION 5.00
Begin VB.Form frmProdDepositos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Depósitos"
   ClientHeight    =   9585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9585
   ScaleWidth      =   13695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Virtuales"
      Height          =   615
      Left            =   480
      TabIndex        =   55
      Top             =   8880
      Width           =   6015
      Begin VB.Image imgLinea 
         Height          =   360
         Index           =   6
         Left            =   5160
         Top             =   210
         Width           =   360
      End
      Begin VB.Image imgLinea 
         Height          =   360
         Index           =   5
         Left            =   4380
         Top             =   210
         Width           =   360
      End
      Begin VB.Image imgLinea 
         Height          =   360
         Index           =   4
         Left            =   3600
         Top             =   210
         Width           =   360
      End
      Begin VB.Image imgLinea 
         Height          =   360
         Index           =   3
         Left            =   2820
         Top             =   210
         Width           =   360
      End
      Begin VB.Image imgLinea 
         Height          =   360
         Index           =   2
         Left            =   2040
         Top             =   210
         Width           =   360
      End
      Begin VB.Image imgLinea 
         Height          =   360
         Index           =   1
         Left            =   1260
         Top             =   210
         Width           =   360
      End
      Begin VB.Image imgLinea 
         Height          =   360
         Index           =   0
         Left            =   480
         Top             =   210
         Width           =   360
      End
      Begin VB.Label lblVirtu 
         Caption         =   "3"
         Height          =   255
         Index           =   6
         Left            =   5040
         TabIndex        =   62
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblVirtu 
         Caption         =   "3"
         Height          =   255
         Index           =   5
         Left            =   4260
         TabIndex        =   61
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblVirtu 
         Caption         =   "3"
         Height          =   255
         Index           =   4
         Left            =   3480
         TabIndex        =   60
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblVirtu 
         Caption         =   "3"
         Height          =   255
         Index           =   3
         Left            =   2700
         TabIndex        =   59
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblVirtu 
         Caption         =   "3"
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   58
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblVirtu 
         Caption         =   "3"
         Height          =   255
         Index           =   1
         Left            =   1140
         TabIndex        =   57
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblVirtu 
         Caption         =   "3"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   56
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   26
      Left            =   12600
      TabIndex        =   54
      Text            =   "Text1"
      Top             =   8280
      Width           =   735
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   25
      Left            =   12600
      TabIndex        =   53
      Text            =   "Text1"
      Top             =   5400
      Width           =   735
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   24
      Left            =   12600
      TabIndex        =   52
      Text            =   "Text1"
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   23
      Left            =   11280
      TabIndex        =   51
      Text            =   "Text1"
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   22
      Left            =   11280
      TabIndex        =   50
      Text            =   "Text1"
      Top             =   5400
      Width           =   735
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   21
      Left            =   11280
      TabIndex        =   49
      Text            =   "Text1"
      Top             =   8280
      Width           =   735
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   20
      Left            =   9480
      TabIndex        =   48
      Text            =   "Text1"
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   19
      Left            =   8040
      TabIndex        =   47
      Text            =   "Text1"
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   18
      Left            =   6600
      TabIndex        =   46
      Text            =   "Text1"
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   17
      Left            =   5160
      TabIndex        =   45
      Text            =   "Text1"
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   16
      Left            =   3720
      TabIndex        =   44
      Text            =   "Text1"
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   15
      Left            =   2280
      TabIndex        =   43
      Text            =   "Text1"
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   14
      Left            =   840
      TabIndex        =   42
      Text            =   "Text1"
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   13
      Left            =   9480
      TabIndex        =   41
      Text            =   "Text1"
      Top             =   5400
      Width           =   735
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   12
      Left            =   8040
      TabIndex        =   40
      Text            =   "Text1"
      Top             =   5400
      Width           =   735
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   11
      Left            =   6600
      TabIndex        =   39
      Text            =   "Text1"
      Top             =   5400
      Width           =   735
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   10
      Left            =   5160
      TabIndex        =   38
      Text            =   "Text1"
      Top             =   5400
      Width           =   735
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   9
      Left            =   3720
      TabIndex        =   37
      Text            =   "Text1"
      Top             =   5400
      Width           =   735
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   8
      Left            =   2280
      TabIndex        =   36
      Text            =   "Text1"
      Top             =   5400
      Width           =   735
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   7
      Left            =   840
      TabIndex        =   35
      Text            =   "Text1"
      Top             =   5400
      Width           =   855
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   6
      Left            =   9480
      TabIndex        =   34
      Text            =   "Text1"
      Top             =   8280
      Width           =   735
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   5
      Left            =   8040
      TabIndex        =   33
      Text            =   "Text1"
      Top             =   8280
      Width           =   735
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   4
      Left            =   6600
      TabIndex        =   32
      Text            =   "Text1"
      Top             =   8280
      Width           =   735
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   3
      Left            =   5160
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   8280
      Width           =   735
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   2
      Left            =   3720
      TabIndex        =   30
      Text            =   "Text1"
      Top             =   8280
      Width           =   735
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   1
      Left            =   2280
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   8280
      Width           =   735
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   0
      Left            =   840
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   8280
      Width           =   855
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   12600
      TabIndex        =   27
      Top             =   9000
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   2
      Left            =   360
      Picture         =   "frmProdDepositos.frx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   1
      Left            =   840
      Picture         =   "frmProdDepositos.frx":1A72
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   0
      Left            =   0
      Picture         =   "frmProdDepositos.frx":34E4
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   2055
      Index           =   24
      Left            =   12360
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Index           =   26
      Left            =   12360
      TabIndex        =   26
      Top             =   8280
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "2"
      Height          =   195
      Index           =   25
      Left            =   12360
      TabIndex        =   25
      Top             =   5400
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Index           =   24
      Left            =   12360
      TabIndex        =   24
      Top             =   2520
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "2"
      Height          =   195
      Index           =   23
      Left            =   11040
      TabIndex        =   23
      Top             =   2520
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Index           =   22
      Left            =   11040
      TabIndex        =   22
      Top             =   5400
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "2"
      Height          =   195
      Index           =   21
      Left            =   11040
      TabIndex        =   21
      Top             =   8280
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Index           =   20
      Left            =   9240
      TabIndex        =   20
      Top             =   2520
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "2"
      Height          =   195
      Index           =   19
      Left            =   7800
      TabIndex        =   19
      Top             =   2520
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Index           =   18
      Left            =   6360
      TabIndex        =   18
      Top             =   2520
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "2"
      Height          =   195
      Index           =   17
      Left            =   4920
      TabIndex        =   17
      Top             =   2520
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Index           =   16
      Left            =   3480
      TabIndex        =   16
      Top             =   2520
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "2"
      Height          =   195
      Index           =   15
      Left            =   2040
      TabIndex        =   15
      Top             =   2520
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Index           =   14
      Left            =   600
      TabIndex        =   14
      Top             =   2520
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "2"
      Height          =   195
      Index           =   13
      Left            =   9240
      TabIndex        =   13
      Top             =   5400
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Index           =   12
      Left            =   7800
      TabIndex        =   12
      Top             =   5400
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "2"
      Height          =   195
      Index           =   11
      Left            =   6360
      TabIndex        =   11
      Top             =   5400
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Index           =   10
      Left            =   4920
      TabIndex        =   10
      Top             =   5400
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "2"
      Height          =   195
      Index           =   9
      Left            =   3480
      TabIndex        =   9
      Top             =   5400
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Index           =   8
      Left            =   2040
      TabIndex        =   8
      Top             =   5400
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "2"
      Height          =   195
      Index           =   7
      Left            =   600
      TabIndex        =   7
      Top             =   5400
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Index           =   6
      Left            =   9240
      TabIndex        =   6
      Top             =   8280
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "2"
      Height          =   195
      Index           =   5
      Left            =   7800
      TabIndex        =   5
      Top             =   8280
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Index           =   4
      Left            =   6360
      TabIndex        =   4
      Top             =   8280
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "2"
      Height          =   195
      Index           =   3
      Left            =   4920
      TabIndex        =   3
      Top             =   8280
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Index           =   2
      Left            =   3480
      TabIndex        =   2
      Top             =   8280
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "2"
      Height          =   195
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      Top             =   8280
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   8280
      Width           =   90
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   2055
      Index           =   26
      Left            =   12360
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   2055
      Index           =   25
      Left            =   12360
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   2055
      Index           =   23
      Left            =   11040
      Top             =   360
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   2055
      Index           =   22
      Left            =   11040
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   2055
      Index           =   21
      Left            =   11040
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   2055
      Index           =   20
      Left            =   9240
      Top             =   360
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   2055
      Index           =   19
      Left            =   7800
      Top             =   360
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   2055
      Index           =   18
      Left            =   6360
      Top             =   360
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   2055
      Index           =   17
      Left            =   4920
      Top             =   360
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   2055
      Index           =   16
      Left            =   3480
      Top             =   360
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   2055
      Index           =   15
      Left            =   2040
      Top             =   360
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   2055
      Index           =   14
      Left            =   600
      Top             =   360
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   2055
      Index           =   13
      Left            =   9240
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   2055
      Index           =   12
      Left            =   7800
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   2055
      Index           =   11
      Left            =   6360
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   2055
      Index           =   10
      Left            =   4920
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   2055
      Index           =   9
      Left            =   3480
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   2055
      Index           =   8
      Left            =   2040
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   2055
      Index           =   7
      Left            =   600
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   2055
      Index           =   6
      Left            =   9240
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   2055
      Index           =   5
      Left            =   7800
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   2055
      Index           =   4
      Left            =   6360
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   2055
      Index           =   3
      Left            =   4920
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   2055
      Index           =   2
      Left            =   3480
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   2055
      Index           =   1
      Left            =   2040
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   2055
      Index           =   0
      Left            =   600
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   6
      Left            =   9240
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   5
      Left            =   7800
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   4
      Left            =   6360
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   3
      Left            =   4920
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   2
      Left            =   3480
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   1
      Left            =   2040
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   0
      Left            =   600
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   6
      Left            =   9240
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   5
      Left            =   7800
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   4
      Left            =   6360
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   3
      Left            =   4920
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   2
      Left            =   3480
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   1
      Left            =   2040
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   0
      Left            =   600
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   26
      Left            =   12360
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   25
      Left            =   12360
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   22
      Left            =   11040
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   21
      Left            =   11040
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   13
      Left            =   9240
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   12
      Left            =   7800
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   11
      Left            =   6360
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   10
      Left            =   4920
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   9
      Left            =   3480
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   8
      Left            =   2040
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   7
      Left            =   600
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   23
      Left            =   11040
      Top             =   360
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   20
      Left            =   9240
      Top             =   360
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   19
      Left            =   7800
      Top             =   360
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   18
      Left            =   6360
      Top             =   360
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   17
      Left            =   4920
      Top             =   360
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   16
      Left            =   3480
      Top             =   360
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   15
      Left            =   2040
      Top             =   360
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   14
      Left            =   600
      Top             =   360
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   23
      Left            =   11040
      Top             =   360
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   20
      Left            =   9240
      Top             =   360
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   19
      Left            =   7800
      Top             =   360
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   18
      Left            =   6360
      Top             =   360
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   17
      Left            =   4920
      Top             =   360
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   16
      Left            =   3480
      Top             =   360
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   15
      Left            =   2040
      Top             =   360
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   14
      Left            =   600
      Top             =   360
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   25
      Left            =   12360
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   22
      Left            =   11040
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   13
      Left            =   9240
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   12
      Left            =   7800
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   11
      Left            =   6360
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   10
      Left            =   4920
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   9
      Left            =   3480
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   8
      Left            =   2040
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   7
      Left            =   600
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   26
      Left            =   12360
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   21
      Left            =   11040
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   24
      Left            =   12360
      Top             =   360
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   24
      Left            =   12360
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "frmProdDepositos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cad As String
Dim RS As ADODB.Recordset

Dim DepositoDblClik As Integer




Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_DblClick()
    
    If DepositoDblClik < 0 Then Exit Sub
    'If txtLote(DepositoDblClik).Text = "" Then Exit Sub
    
    
    frmProdVerUnDepo.NumDepo = DepositoDblClik + 1
    frmProdVerUnDepo.idProd = Me.txtLote(DepositoDblClik).Tag
    frmProdVerUnDepo.Show vbModal
    
    If CadenaDesdeOtroForm <> "" Then
        Screen.MousePointer = vbHourglass
        CargarDepositos
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = frmppal.Icon
    CargarDepositos
    CargarDepositosVirtuales
    'Command1.Top = Me.Height + 400
    DepositoDblClik = -1
End Sub



Private Sub CargarDepositos()
Dim KMostrar As Byte
Dim Deposito As Integer
Dim Lotes As String

    cad = "select NumDeposito,capacidad,kilos,spartidas.codartic,nomartic,factorconversion,spartidas.numlote"
    cad = cad & " from proddepositos left join spartidas on proddepositos.numlote=spartidas.numlote"
    cad = cad & " left join sartic on spartidas.codartic=sartic.codartic"
    cad = cad & " WHERE DepositoVtaDirecta=0"
    
    Set RS = New ADODB.Recordset
    RS.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        'Sera ERROR SI,
        ' o sartic.codartic=null
        ' o factorconverrsion=1
        KMostrar = 2 'todo
        If DBLet(RS!NUmlote, "T") <> "" Then
            If DBLet(RS!codArtic, "T") = "" Then
                KMostrar = 0 'ERROR
            Else
                If RS!FactorConversion = 1 Then KMostrar = 0
            End If
        Else
            KMostrar = 1 'No tiene lote. VAcio
        End If
        
        'Comun
        'If RS!NumDeposito = 24 Then Stop
        Deposito = RS!NumDeposito - 1
        Me.txtLote(Deposito).Text = ""
        txtLote(Deposito).Tag = ""
        Me.txtLote(Deposito).Locked = True
        Me.ShFondo(Deposito).BackColor = &H808080
        lblDep(Deposito) = RS!NumDeposito
        
        
        If KMostrar < 2 Then
            'LINEA CON ERROR o VACIO
            
            LimpiarDatosDeposito KMostrar = 0, RS!NumDeposito - 1
            
            
        Else
            CargarUnDeposito Deposito
            Lotes = Lotes & ", " & DBSet(RS!NUmlote, "T")
            
        End If
    
    
        RS.MoveNext
        
    
    Wend
    RS.Close
    
    
    'Vamos a ver cual esta envasando en linea de produccion
    If Lotes <> "" Then
        
        cad = "select numlote,prodlin.codigo,prodlin.idlin from prodlin,prodtrazcompo"
        cad = cad & " where prodlin.codigo= prodtrazcompo.codigo AND prodlin.idlin = "
        cad = cad & " prodtrazcompo.idlin and prodtrazcompo.cantutili is null and estado >0"
        cad = cad & " and estado<10 and numlote in"
        cad = cad & " (" & Mid(Lotes, 2) & ")"
        RS.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
            For KMostrar = 0 To MaxNumDepositos_ - 1
                If Me.txtLote(KMostrar).Text = RS!NUmlote Then
                    Me.Shape1(KMostrar).BorderWidth = 3
                    Me.Shape1(KMostrar).BorderColor = vbGreen
                    txtLote(KMostrar).Tag = RS!Codigo & "|" & RS!idlin & "|" 'Para cuando pong ver deposito
                    Exit For
                End If
            Next
            RS.MoveNext
        Wend
        RS.Close
    End If
        
    Set RS = Nothing
End Sub


Private Sub LimpiarDatosDeposito(HayError As Boolean, kDeposito As Integer)
    'Si hay textos
    Me.txtLote(kDeposito).visible = False
    
    'Los graficos
    ShDeposito(kDeposito).Height = 0
    Shape1(kDeposito).BorderColor = vbBlack
    If HayError Then
        'DATOS rs
        'NumDeposito,capacidad,kilos,spartidas.codartic,nomartic,factorconversion,spartidas.numlote
        If RS.EOF Then
            cad = "Consulta vacia (EOF)"
        Else
            cad = "Deposito:      " & DBLet(RS!NumDeposito, "T") & vbCrLf
            cad = cad & "Capacidad:     " & DBLet(RS!Capacidad, "T") & vbCrLf
            cad = cad & "Codigo:     " & DBLet(RS!codArtic, "T") & vbCrLf
            cad = cad & "Referencia:   " & DBLet(RS!NomArtic, "T") & vbCrLf
            cad = cad & "Factor conversion:      " & DBLet(RS!FactorConversion, "T") & vbCrLf
            cad = cad & "LOTE:      " & DBLet(RS!NUmlote, "T") & vbCrLf
            
            cad = "Error datos deposito: " & vbCrLf & vbCrLf & cad
    
        End If
        MsgBox cad, vbExclamation
            
    End If
End Sub


Private Sub CargarUnDeposito(kDeposito As Integer)
Dim PorcentajeLleno As Integer
Dim Cantidad As Currency

    
    Me.txtLote(kDeposito).visible = True
    Me.txtLote(kDeposito).Text = RS!NUmlote
    Me.txtLote(kDeposito).ToolTipText = RS!codArtic & "  " & RS!NomArtic
    Me.txtLote(kDeposito).Alignment = 2
    
    If DBLet(RS!Kilos, "N") = 0 Then
        Cantidad = 0
        
    Else
        Cantidad = RS!Kilos / RS!FactorConversion
    End If
    PorcentajeLleno = Round((Cantidad * 100) / RS!Capacidad, 2)
    
    If PorcentajeLleno > 100 Then
        PorcentajeLleno = 100
    ElseIf PorcentajeLleno < 0 Then PorcentajeLleno = 0
    End If
    
    
    
    
    PorcentajeLleno = CInt((Me.ShFondo(kDeposito).Height * PorcentajeLleno / 100))
    ShDeposito(kDeposito).Height = PorcentajeLleno
    'Ya tengo lo que debe medir el shape de deposito. Luego hay que moverlo un poquito hasta ajustar
    
    'If kDeposito = 25 Then Stop
        
    PorcentajeLleno = ShFondo(kDeposito).Height - PorcentajeLleno
    ShDeposito(kDeposito).Top = ShFondo(kDeposito).Top + PorcentajeLleno
    
    If PorcentajeLleno = 100 Then
        Me.Shape1(kDeposito).BorderWidth = 1
    Else
        Me.Shape1(kDeposito).BorderWidth = 3
    End If
    Me.Shape1(kDeposito).BorderColor = vbBlue
End Sub






Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim J As Integer
    
    DepositoDblClik = -1
    If Button = vbLeftButton Then
    
        For J = 0 To MaxNumDepositos_ - 1
            If Me.ShFondo(J).Left <= X Then
                If Me.ShFondo(J).Left + Me.ShFondo(J).Width >= X Then
                    'OK la x correcta. Vamos a verel eje Y
                    If Me.ShFondo(J).Top <= Y Then
                        If Me.ShFondo(J).Top + Me.ShFondo(J).Height >= Y Then
                            DepositoDblClik = J
                            Exit For
                        End If
                    End If
                End If
            End If
                            
        Next
        
    End If
    
End Sub



Private Sub CargarDepositosVirtuales()
Dim KMostrar As Byte
Dim Lotes As String
Dim QueImage As Integer

    cad = "select NumDeposito,capacidad,kilos,spartidas.codartic,nomartic,factorconversion,spartidas.numlote"
    cad = cad & " from proddepositos left join spartidas on proddepositos.numlote=spartidas.numlote"
    cad = cad & " left join sartic on spartidas.codartic=sartic.codartic"
    cad = cad & " WHERE DepositoVtaDirecta=1 ORDER BY numdeposito"
    
    QueImage = 0
    Set RS = New ADODB.Recordset
    RS.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        'Sera ERROR SI,
        ' o sartic.codartic=null
        ' o factorconverrsion=1
        KMostrar = 2 'todo
        
        
        imgLinea(QueImage).Tag = RS!NumDeposito & "|"
        If DBLet(RS!NUmlote, "T") <> "" Then
            
            If DBLet(RS!codArtic, "T") = "" Then
                KMostrar = 0 'ERROR
            Else
                If RS!FactorConversion < 1 Then
                    KMostrar = 2
                Else
                    KMostrar = 0 'ERROR
                    imgLinea(QueImage).Tag = 0
                End If
            End If
        Else
            KMostrar = 1 'No tiene lote. VAcio
        End If
        
        'Me.txtLote(kDeposito).Text = Rs!NUmlote
        
        ' Rs!NumDeposito & "||"
        
        Me.imgLinea(QueImage).visible = True
        Me.lblVirtu(QueImage).visible = True
        Me.lblVirtu(QueImage).Caption = QueImage
        
        If KMostrar < 2 Then
            'LINEA CON ERROR o VACIO
            
            Me.imgLinea(QueImage).Picture = Me.Image1(KMostrar).Picture
            imgLinea(QueImage).ToolTipText = ""
            If KMostrar = 1 Then imgLinea(QueImage).ToolTipText = "Virtual  vacio"
        Else
            Me.imgLinea(QueImage).Picture = Me.Image1(2).Picture
            Lotes = Lotes & ", " & DBSet(RS!NUmlote, "T")
            imgLinea(QueImage).ToolTipText = "Lote: " & DBSet(RS!NUmlote, "T")
        End If
    
    
        RS.MoveNext
        QueImage = QueImage + 1
    
    Wend
    RS.Close
    
    
    'Vamos a ver cual esta envasando en linea de produccion
    If Lotes <> "" Then
        
'        Cad = "select numlote,prodlin.codigo,prodlin.idlin from prodlin,prodtrazcompo"
'        Cad = Cad & " where prodlin.codigo= prodtrazcompo.codigo AND prodlin.idlin = "
'        Cad = Cad & " prodtrazcompo.idlin and prodtrazcompo.cantutili is null and estado >0"
'        Cad = Cad & " and estado<10 and numlote in"
'        Cad = Cad & " (" & Mid(Lotes, 2) & ")"
'        Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        While Not Rs.EOF
'            For KMostrar = 0 To QueImage - 1
''                If Me.txtLote(KMostrar).Text = Rs!NUmlote Then
'''                    Me.Shape1(KMostrar).BorderWidth = 3
'''                    Me.Shape1(KMostrar).BorderColor = vbGreen
'''                    txtLote(KMostrar).Tag = Rs!codigo & "|" & Rs!idlin & "|" 'Para cuando pong ver deposito
'''                    Exit For
''                End If
'            Next
'            Rs.MoveNext
'        Wend
'        Rs.Close
    End If
        
    Set RS = Nothing

End Sub

Private Sub imgLinea_DblClick(Index As Integer)
    If imgLinea(Index).visible Then
        If Val(imgLinea(Index).Tag) > 0 Then
            frmProdVerUnDepo.NumDepo = RecuperaValor(imgLinea(Index).Tag, 1)
            'frmProdVerUnDepo.idProd = Me.txtLote(DepositoDblClik).Tag
            frmProdVerUnDepo.Show vbModal
        End If
    End If
End Sub

