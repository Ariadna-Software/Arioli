VERSION 5.00
Begin VB.Form frmProdDepositosVall 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Depósitos"
   ClientHeight    =   9915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9915
   ScaleWidth      =   14025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Virtuales"
      Height          =   2055
      Left            =   10080
      TabIndex        =   37
      Top             =   120
      Width           =   1335
      Begin VB.Image imgLinea 
         Height          =   360
         Index           =   6
         Left            =   480
         Top             =   2610
         Width           =   360
      End
      Begin VB.Image imgLinea 
         Height          =   360
         Index           =   5
         Left            =   2100
         Top             =   2130
         Width           =   360
      End
      Begin VB.Image imgLinea 
         Height          =   360
         Index           =   4
         Left            =   1320
         Top             =   2130
         Width           =   360
      End
      Begin VB.Image imgLinea 
         Height          =   360
         Index           =   3
         Left            =   540
         Top             =   2130
         Width           =   360
      End
      Begin VB.Image imgLinea 
         Height          =   360
         Index           =   2
         Left            =   600
         Top             =   1320
         Width           =   360
      End
      Begin VB.Image imgLinea 
         Height          =   360
         Index           =   1
         Left            =   600
         Top             =   720
         Width           =   360
      End
      Begin VB.Image imgLinea 
         Height          =   360
         Index           =   0
         Left            =   600
         Top             =   240
         Width           =   360
      End
      Begin VB.Label lblVirtu 
         Caption         =   "3"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   44
         Top             =   2640
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblVirtu 
         Caption         =   "3"
         Height          =   255
         Index           =   5
         Left            =   1980
         TabIndex        =   43
         Top             =   2160
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblVirtu 
         Caption         =   "3"
         Height          =   255
         Index           =   4
         Left            =   1200
         TabIndex        =   42
         Top             =   2160
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblVirtu 
         Caption         =   "3"
         Height          =   255
         Index           =   3
         Left            =   420
         TabIndex        =   41
         Top             =   2160
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblVirtu 
         Caption         =   "3"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   40
         Top             =   1320
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblVirtu 
         Caption         =   "3"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   39
         Top             =   840
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblVirtu 
         Caption         =   "3"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   38
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   17
      Left            =   10200
      TabIndex        =   36
      Text            =   "Text1"
      Top             =   6360
      Width           =   1335
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   16
      Left            =   7920
      TabIndex        =   35
      Text            =   "Text1"
      Top             =   9480
      Width           =   1095
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   15
      Left            =   6120
      TabIndex        =   34
      Text            =   "Text1"
      Top             =   9480
      Width           =   1095
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   14
      Left            =   4320
      TabIndex        =   33
      Text            =   "Text1"
      Top             =   9480
      Width           =   1095
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   13
      Left            =   2640
      TabIndex        =   32
      Text            =   "Text1"
      Top             =   9480
      Width           =   1095
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   12
      Left            =   960
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   9480
      Width           =   1095
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   11
      Left            =   720
      TabIndex        =   30
      Text            =   "Text1"
      Top             =   6480
      Width           =   1095
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   10
      Left            =   2280
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   6480
      Width           =   1095
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   9
      Left            =   3840
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   6480
      Width           =   1095
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   8
      Left            =   5400
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   6480
      Width           =   1095
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   7
      Left            =   6960
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   6480
      Width           =   1095
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   6
      Left            =   8520
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   6480
      Width           =   1095
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   5
      Left            =   8520
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   4
      Left            =   6960
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   3
      Left            =   5400
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   2
      Left            =   3840
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   1
      Left            =   2280
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Index           =   0
      Left            =   720
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   10200
      TabIndex        =   18
      Top             =   9120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   " Embasado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   9960
      TabIndex        =   45
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   2
      Left            =   360
      Picture         =   "frmProdDepositosVall.frx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   1
      Left            =   840
      Picture         =   "frmProdDepositosVall.frx":1A72
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   0
      Left            =   0
      Picture         =   "frmProdDepositosVall.frx":34E4
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "2"
      Height          =   195
      Index           =   17
      Left            =   9960
      TabIndex        =   17
      Top             =   6360
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Index           =   16
      Left            =   7680
      TabIndex        =   16
      Top             =   9480
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "2"
      Height          =   195
      Index           =   15
      Left            =   5880
      TabIndex        =   15
      Top             =   9480
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Index           =   14
      Left            =   4080
      TabIndex        =   14
      Top             =   9480
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "2"
      Height          =   195
      Index           =   13
      Left            =   2400
      TabIndex        =   13
      Top             =   9480
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Index           =   12
      Left            =   720
      TabIndex        =   12
      Top             =   9480
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "2"
      Height          =   195
      Index           =   11
      Left            =   480
      TabIndex        =   11
      Top             =   6480
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Index           =   10
      Left            =   2040
      TabIndex        =   10
      Top             =   6480
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "2"
      Height          =   195
      Index           =   9
      Left            =   3480
      TabIndex        =   9
      Top             =   6480
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Index           =   8
      Left            =   5040
      TabIndex        =   8
      Top             =   6480
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "2"
      Height          =   195
      Index           =   7
      Left            =   6600
      TabIndex        =   7
      Top             =   6480
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Index           =   6
      Left            =   8160
      TabIndex        =   6
      Top             =   6480
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "2"
      Height          =   195
      Index           =   5
      Left            =   8280
      TabIndex        =   5
      Top             =   2880
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Index           =   4
      Left            =   6720
      TabIndex        =   4
      Top             =   2880
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "2"
      Height          =   195
      Index           =   3
      Left            =   5160
      TabIndex        =   3
      Top             =   2880
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Index           =   2
      Left            =   3600
      TabIndex        =   2
      Top             =   2880
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "2"
      Height          =   195
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      Top             =   2880
      Width           =   90
   End
   Begin VB.Label lblDep 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   2880
      Width           =   90
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   1815
      Index           =   17
      Left            =   10200
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   1920
      Index           =   16
      Left            =   7680
      Top             =   7420
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   1920
      Index           =   15
      Left            =   5880
      Top             =   7420
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   1920
      Index           =   14
      Left            =   4080
      Top             =   7420
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   1920
      Index           =   13
      Left            =   2400
      Top             =   7420
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   1920
      Index           =   12
      Left            =   720
      Top             =   7420
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   2775
      Index           =   11
      Left            =   480
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   2055
      Index           =   10
      Left            =   2040
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   2055
      Index           =   9
      Left            =   3600
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   2055
      Index           =   8
      Left            =   5160
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   2055
      Index           =   7
      Left            =   6720
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   2055
      Index           =   6
      Left            =   8280
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   2655
      Index           =   5
      Left            =   8280
      Top             =   120
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   2655
      Index           =   4
      Left            =   6720
      Top             =   120
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   2055
      Index           =   3
      Left            =   5160
      Top             =   720
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   2055
      Index           =   2
      Left            =   3600
      Top             =   720
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   2055
      Index           =   1
      Left            =   2040
      Top             =   720
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   2055
      Index           =   0
      Left            =   480
      Top             =   720
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   6
      Left            =   8280
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2655
      Index           =   5
      Left            =   8280
      Top             =   120
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2655
      Index           =   4
      Left            =   6720
      Top             =   120
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   3
      Left            =   5160
      Top             =   720
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   2
      Left            =   3600
      Top             =   720
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   1
      Left            =   2040
      Top             =   720
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   0
      Left            =   480
      Top             =   720
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   6
      Left            =   8280
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   2655
      Index           =   5
      Left            =   8280
      Top             =   120
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   2655
      Index           =   4
      Left            =   6720
      Top             =   120
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   3
      Left            =   5160
      Top             =   720
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   2
      Left            =   3600
      Top             =   720
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   1
      Left            =   2040
      Top             =   720
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   0
      Left            =   480
      Top             =   720
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   1920
      Index           =   13
      Left            =   2400
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   1920
      Index           =   12
      Left            =   720
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2775
      Index           =   11
      Left            =   480
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   10
      Left            =   2040
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   9
      Left            =   3600
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   8
      Left            =   5160
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   7
      Left            =   6720
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   1815
      Index           =   17
      Left            =   10200
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   1920
      Index           =   16
      Left            =   7680
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   1920
      Index           =   15
      Left            =   5880
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Shape ShDeposito 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   1920
      Index           =   14
      Left            =   4080
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   1815
      Index           =   17
      Left            =   10200
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   1920
      Index           =   16
      Left            =   7680
      Top             =   7420
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   1920
      Index           =   15
      Left            =   5880
      Top             =   7420
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   1920
      Index           =   14
      Left            =   4080
      Top             =   7420
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   1920
      Index           =   13
      Left            =   2400
      Top             =   7420
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   1920
      Index           =   12
      Left            =   720
      Top             =   7420
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   2775
      Index           =   11
      Left            =   480
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   10
      Left            =   2040
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   9
      Left            =   3600
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   8
      Left            =   5160
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Shape ShFondo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   7
      Left            =   6720
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000080&
      Height          =   3015
      Left            =   9840
      Top             =   3840
      Width           =   1815
   End
End
Attribute VB_Name = "frmProdDepositosVall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Cad As String
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

    
    '8 depositos de 45000 litros, Ajusto de tamaño los de la fila intermedia
    For Deposito = 6 To 11
        Me.Shape1(Deposito).Top = 3680
        Me.ShFondo(Deposito).Top = 3680
        Me.ShDeposito(Deposito).Top = 3680
        
        Me.Shape1(Deposito).Height = 2655
        Me.ShFondo(Deposito).Height = 2655
        Me.ShDeposito(Deposito).Height = 2655
    Next




    
    
    
    
    


    Cad = "select NumDeposito,capacidad,kilos,spartidas.codartic,nomartic,factorconversion,spartidas.numlote"
    Cad = Cad & " from proddepositos left join spartidas on proddepositos.numlote=spartidas.numlote"
    Cad = Cad & " left join sartic on spartidas.codartic=sartic.codartic"
    Cad = Cad & " WHERE DepositoVtaDirecta=0"
    
    Set RS = New ADODB.Recordset
    RS.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        'Sera ERROR SI,
        ' o sartic.codartic=null
        ' o factorconverrsion=1
        KMostrar = 2 'todo
        If DBLet(RS!NUmlote, "T") <> "" Then
            If DBLet(RS!codartic, "T") = "" Then
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
    If Lotes <> "" And vParamAplic.QUE_EMPRESA = 0 Then
        
        Cad = "select numlote,prodlin.codigo,prodlin.idlin from prodlin,prodtrazcompo"
        Cad = Cad & " where prodlin.codigo= prodtrazcompo.codigo AND prodlin.idlin = "
        Cad = Cad & " prodtrazcompo.idlin and prodtrazcompo.cantutili is null and estado >0"
        Cad = Cad & " and estado<10 and numlote in"
        Cad = Cad & " (" & Mid(Lotes, 2) & ")"
        RS.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
            Cad = "Consulta vacia (EOF)"
        Else
            Cad = "Deposito:      " & DBLet(RS!NumDeposito, "T") & vbCrLf
            Cad = Cad & "Capacidad:     " & DBLet(RS!Capacidad, "T") & vbCrLf
            Cad = Cad & "Codigo:     " & DBLet(RS!codartic, "T") & vbCrLf
            Cad = Cad & "Referencia:   " & DBLet(RS!NomArtic, "T") & vbCrLf
            Cad = Cad & "Factor conversion:      " & DBLet(RS!FactorConversion, "T") & vbCrLf
            Cad = Cad & "LOTE:      " & DBLet(RS!NUmlote, "T") & vbCrLf
            
            Cad = "Error datos deposito: " & vbCrLf & vbCrLf & Cad
    
        End If
        MsgBox Cad, vbExclamation
            
    End If
End Sub


Private Sub CargarUnDeposito(kDeposito As Integer)
Dim PorcentajeLleno As Integer
Dim Cantidad As Currency

    
    Me.txtLote(kDeposito).visible = True
    Me.txtLote(kDeposito).Text = RS!NUmlote
    If UCase(Mid(RS!NUmlote, 1, 6)) = "MOSTRA" Then Me.txtLote(kDeposito).Text = Mid(RS!NUmlote, 7)
    
    Me.txtLote(kDeposito).ToolTipText = RS!codartic & "  " & RS!NomArtic
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

    Cad = "select NumDeposito,capacidad,kilos,spartidas.codartic,nomartic,factorconversion,spartidas.numlote"
    Cad = Cad & " from proddepositos left join spartidas on proddepositos.numlote=spartidas.numlote"
    Cad = Cad & " left join sartic on spartidas.codartic=sartic.codartic"
    Cad = Cad & " WHERE DepositoVtaDirecta=1 ORDER BY numdeposito"
    
    QueImage = 0
    Set RS = New ADODB.Recordset
    RS.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        'Sera ERROR SI,
        ' o sartic.codartic=null
        ' o factorconverrsion=1
        KMostrar = 2 'todo
        
        
        imgLinea(QueImage).Tag = RS!NumDeposito & "|"
        If DBLet(RS!NUmlote, "T") <> "" Then
            
            If DBLet(RS!codartic, "T") = "" Then
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

