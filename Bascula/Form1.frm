VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   ClientHeight    =   11430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18960
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11430
   ScaleWidth      =   18960
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   3120
   End
   Begin VB.Timer Timer11 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   0
      Top             =   3600
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   120
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Frame FramelecturaBascula 
      Height          =   10935
      Left            =   2160
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   14895
      Begin VB.Frame FrameResultados 
         Height          =   4575
         Left            =   720
         TabIndex        =   47
         Top             =   5400
         Visible         =   0   'False
         Width           =   13815
         Begin MSComctlLib.ListView lw1 
            Height          =   4095
            Left            =   360
            TabIndex        =   66
            Top             =   360
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   7223
            View            =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Peso"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.CommandButton cmdValirdarLecturas 
            Caption         =   "G U A R D A R"
            Height          =   495
            Index           =   1
            Left            =   11280
            TabIndex        =   48
            Top             =   3840
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.Frame FrameResultProd 
            Height          =   3495
            Left            =   4800
            TabIndex        =   49
            Top             =   240
            Width           =   8895
            Begin VB.TextBox txtPesada 
               BackColor       =   &H80000018&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   585
               Index           =   12
               Left            =   2640
               TabIndex        =   64
               Text            =   "Text1"
               Top             =   2640
               Width           =   1935
            End
            Begin VB.CommandButton cmdValirdarLecturas 
               Caption         =   "Ampliación"
               Height          =   495
               Index           =   2
               Left            =   6360
               TabIndex        =   63
               Top             =   2760
               Visible         =   0   'False
               Width           =   2415
            End
            Begin VB.TextBox txtPesada 
               BackColor       =   &H80000018&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   585
               Index           =   11
               Left            =   6600
               TabIndex        =   61
               Text            =   "Text1"
               Top             =   1320
               Width           =   2055
            End
            Begin VB.TextBox txtPesada 
               BackColor       =   &H80000018&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   585
               Index           =   10
               Left            =   2280
               TabIndex        =   59
               Text            =   "Text1"
               Top             =   1320
               Width           =   2655
            End
            Begin VB.TextBox txtPesada 
               BackColor       =   &H80000018&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   585
               Index           =   9
               Left            =   6120
               TabIndex        =   57
               Text            =   "Text1"
               Top             =   360
               Width           =   1095
            End
            Begin VB.TextBox txtPesada 
               BackColor       =   &H80000018&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   585
               Index           =   8
               Left            =   1800
               TabIndex        =   55
               Text            =   "Text1"
               Top             =   360
               Width           =   1215
            End
            Begin VB.Line Line2 
               X1              =   120
               X2              =   8640
               Y1              =   2280
               Y2              =   2280
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Med Peso"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   27.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   675
               Index           =   16
               Left            =   120
               TabIndex        =   65
               Top             =   2520
               Width           =   2340
            End
            Begin VB.Line Line1 
               X1              =   120
               X2              =   8760
               Y1              =   1080
               Y2              =   1080
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Desv.:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   27.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   675
               Index           =   15
               Left            =   5040
               TabIndex        =   62
               Top             =   1320
               Width           =   1545
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Med Vol"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   27.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   675
               Index           =   14
               Left            =   120
               TabIndex        =   60
               Top             =   1320
               Width           =   1950
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "EMPx2"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   27.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   675
               Index           =   13
               Left            =   4320
               TabIndex        =   58
               Top             =   240
               Width           =   1620
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "EMP"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   27.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   675
               Index           =   12
               Left            =   120
               TabIndex        =   56
               Top             =   240
               Width           =   1050
            End
         End
         Begin VB.Frame FrameResultMaAux 
            Height          =   3495
            Left            =   4800
            TabIndex        =   50
            Top             =   240
            Width           =   8295
            Begin VB.TextBox txtPesada 
               BackColor       =   &H00C0C0FF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   24
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Index           =   7
               Left            =   4200
               TabIndex        =   53
               Text            =   "Text1"
               Top             =   1200
               Width           =   3615
            End
            Begin VB.TextBox txtPesada 
               BackColor       =   &H80000018&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   24
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Index           =   6
               Left            =   120
               TabIndex        =   51
               Text            =   "Text1"
               Top             =   1200
               Width           =   3615
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Desviacion"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   36
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   870
               Index           =   11
               Left            =   4200
               TabIndex        =   54
               Top             =   240
               Width           =   3375
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Peso medio"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   36
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   870
               Index           =   10
               Left            =   120
               TabIndex        =   52
               Top             =   240
               Width           =   3630
            End
         End
      End
      Begin VB.CommandButton cmdValirdarLecturas 
         Caption         =   "Cancelar"
         Height          =   495
         Index           =   0
         Left            =   12000
         TabIndex        =   46
         Top             =   10080
         Width           =   2535
      End
      Begin VB.TextBox txtPesada 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   5
         Left            =   720
         TabIndex        =   45
         Text            =   "Text1"
         Top             =   4320
         Width           =   13935
      End
      Begin VB.TextBox txtPesada 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   4
         Left            =   13560
         TabIndex        =   43
         Text            =   "Text1"
         Top             =   3480
         Width           =   1095
      End
      Begin VB.TextBox txtPesada 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   3
         Left            =   7920
         TabIndex        =   41
         Text            =   "Text1"
         Top             =   3480
         Width           =   975
      End
      Begin VB.TextBox txtPesada 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   2
         Left            =   2760
         TabIndex        =   39
         Text            =   "Text1"
         Top             =   3480
         Width           =   615
      End
      Begin VB.TextBox txtPesada 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   2760
         TabIndex        =   37
         Text            =   "Text1"
         Top             =   2160
         Width           =   4335
      End
      Begin VB.TextBox txtPesada 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   2640
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   960
         Width           =   11775
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         BorderWidth     =   5
         Height          =   10575
         Left            =   120
         Top             =   240
         Width           =   14775
      End
      Begin VB.Label lblPeso 
         Alignment       =   2  'Center
         BackColor       =   &H00FDE6F0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1695
         Left            =   720
         TabIndex        =   44
         Top             =   5520
         Width           =   12495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Contador:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Index           =   9
         Left            =   9960
         TabIndex        =   42
         Top             =   3360
         Width           =   3150
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "NºPesadas:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   8
         Left            =   4080
         TabIndex        =   40
         Top             =   3360
         Width           =   3615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Serie:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   7
         Left            =   720
         TabIndex        =   38
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lote:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   6
         Left            =   720
         TabIndex        =   35
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Refer:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   5
         Left            =   720
         TabIndex        =   34
         Top             =   840
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4695
      Left            =   120
      TabIndex        =   19
      Top             =   7680
      Width           =   11415
      Begin VB.Image Image2 
         Height          =   4620
         Left            =   4200
         Picture         =   "Form1.frx":6852
         Stretch         =   -1  'True
         Top             =   240
         Width           =   4020
      End
      Begin VB.Label Label1 
         Caption         =   "Producción"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6975
      Left            =   6240
      TabIndex        =   21
      Top             =   0
      Width           =   5295
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tapón"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   2775
      End
      Begin VB.Image Image3 
         Height          =   2160
         Left            =   1200
         Picture         =   "Form1.frx":7BAA
         Top             =   1560
         Width           =   3240
      End
   End
   Begin VB.Frame FrameBotella 
      Height          =   6975
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   6015
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Envases"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   5775
      End
      Begin VB.Image Image1 
         Height          =   4800
         Left            =   960
         Picture         =   "Form1.frx":108BF
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   3945
      End
   End
   Begin VB.Frame FrameMateriaAuxiliar 
      Height          =   11655
      Left            =   11640
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   8895
      Begin VB.CommandButton cmdMoverVector 
         Height          =   855
         Index           =   1
         Left            =   6960
         Picture         =   "Form1.frx":1237C
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   9720
         Width           =   1095
      End
      Begin VB.CommandButton cmdMoverVector 
         Height          =   855
         Index           =   0
         Left            =   600
         Picture         =   "Form1.frx":12C46
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   9720
         Width           =   1095
      End
      Begin VB.CommandButton cmdMatAux 
         Caption         =   "Command1"
         Height          =   855
         Index           =   4
         Left            =   480
         TabIndex        =   31
         Top             =   7680
         Width           =   7935
      End
      Begin VB.CommandButton cmdMatAux 
         Caption         =   "Command1"
         Height          =   855
         Index           =   3
         Left            =   480
         TabIndex        =   30
         Top             =   6240
         Width           =   7935
      End
      Begin VB.CommandButton cmdMatAux 
         Caption         =   "Command1"
         Height          =   855
         Index           =   2
         Left            =   480
         TabIndex        =   29
         Top             =   5040
         Width           =   7935
      End
      Begin VB.CommandButton cmdMatAux 
         Caption         =   "Command1"
         Height          =   855
         Index           =   1
         Left            =   480
         TabIndex        =   28
         Top             =   3720
         Width           =   7935
      End
      Begin VB.CommandButton cmdMatAux 
         Caption         =   "Command1"
         Height          =   855
         Index           =   0
         Left            =   480
         TabIndex        =   27
         Top             =   2280
         Width           =   7935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Líneas producción"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   870
         Index           =   4
         Left            =   1200
         TabIndex        =   26
         Top             =   600
         Width           =   5700
      End
   End
   Begin VB.Frame FrameLineasProd 
      Height          =   12975
      Left            =   11640
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   8895
      Begin VB.CommandButton cmdLinea 
         Caption         =   "Command1"
         Height          =   735
         Index           =   9
         Left            =   960
         TabIndex        =   68
         Top             =   11880
         Width           =   7215
      End
      Begin VB.CommandButton cmdLinea 
         Caption         =   "Command1"
         Height          =   735
         Index           =   8
         Left            =   960
         TabIndex        =   67
         Top             =   10800
         Width           =   7215
      End
      Begin VB.CommandButton cmdLinea 
         Caption         =   "Command1"
         Height          =   735
         Index           =   0
         Left            =   960
         TabIndex        =   8
         Top             =   2160
         Width           =   7215
      End
      Begin VB.CommandButton cmdLinea 
         Caption         =   "Command1"
         Height          =   735
         Index           =   1
         Left            =   960
         TabIndex        =   7
         Top             =   3240
         Width           =   7215
      End
      Begin VB.CommandButton cmdLinea 
         Caption         =   "Command1"
         Height          =   735
         Index           =   2
         Left            =   960
         TabIndex        =   6
         Top             =   4320
         Width           =   7215
      End
      Begin VB.CommandButton cmdLinea 
         Caption         =   "Command1"
         Height          =   735
         Index           =   3
         Left            =   960
         TabIndex        =   5
         Top             =   5400
         Width           =   7215
      End
      Begin VB.CommandButton cmdLinea 
         Caption         =   "Command1"
         Height          =   735
         Index           =   4
         Left            =   960
         TabIndex        =   4
         Top             =   6480
         Width           =   7215
      End
      Begin VB.CommandButton cmdLinea 
         Caption         =   "Command1"
         Height          =   735
         Index           =   5
         Left            =   960
         TabIndex        =   3
         Top             =   7560
         Width           =   7215
      End
      Begin VB.CommandButton cmdLinea 
         Caption         =   "Command1"
         Height          =   735
         Index           =   6
         Left            =   960
         TabIndex        =   2
         Top             =   8640
         Width           =   7215
      End
      Begin VB.CommandButton cmdLinea 
         Caption         =   "Command1"
         Height          =   735
         Index           =   7
         Left            =   960
         TabIndex        =   1
         Top             =   9720
         Width           =   7215
      End
      Begin VB.Label lblLinea 
         Caption         =   "M2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   9
         Left            =   120
         TabIndex        =   70
         Top             =   11880
         Width           =   735
      End
      Begin VB.Label lblLinea 
         Caption         =   "M1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   8
         Left            =   120
         TabIndex        =   69
         Top             =   10800
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Líneas producción"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   870
         Index           =   3
         Left            =   1080
         TabIndex        =   23
         Top             =   720
         Width           =   5700
      End
      Begin VB.Label lblLinea 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label lblLinea 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label lblLinea 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   2
         Left            =   240
         TabIndex        =   14
         Top             =   4200
         Width           =   615
      End
      Begin VB.Label lblLinea 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   3
         Left            =   240
         TabIndex        =   13
         Top             =   5280
         Width           =   615
      End
      Begin VB.Label lblLinea 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   4
         Left            =   240
         TabIndex        =   12
         Top             =   6360
         Width           =   615
      End
      Begin VB.Label lblLinea 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   5
         Left            =   240
         TabIndex        =   11
         Top             =   7440
         Width           =   615
      End
      Begin VB.Label lblLinea 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   6
         Left            =   240
         TabIndex        =   10
         Top             =   8520
         Width           =   615
      End
      Begin VB.Label lblLinea 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   7
         Left            =   240
         TabIndex        =   9
         Top             =   9600
         Width           =   615
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const Pruebas = False



Dim miRsAux As ADODB.Recordset
Dim SQL As String

Dim colMA As Collection
Dim Puntero As Integer

Dim ParaElInsertSQL As String
Dim CadenaINsertFinal As String
Dim Buffer As String

'Es de parametros y para todos los formatos
Dim EtiquetaEnBD As Currency
Dim Etiqueta2 As Currency
Dim VolumenProd As Currency
Dim PesoBotella As Currency
Dim PesoTapon As Currency
Dim Retractil As Currency
Dim EMP_ As Currency


Dim ClaveReferenciaPesada As String

Dim LanzarProtector As Integer

Dim PesoDePrueba As Currency

Dim ResultadoPesadasIncorrecto As Boolean

Private Sub cmdLinea_Click(Index As Integer)
Dim TextoAux As String
Dim EsperoBotella As Boolean
Dim Ok As Boolean
Dim N As Integer
Dim VaOK As Boolean
Dim RT As ADODB.Recordset
Dim EsUnaLata2 As Boolean
Dim SinPesarBotellas As Boolean
        'Ponemos a pesar
    Set miRsAux = New ADODB.Recordset
    Set RT = New ADODB.Recordset
    VaOK = False
    
    
    If EtiquetaEnBD < 0 Then
        EtiquetaEnBD = 0
        SQL = "Select PesoEtiqueta from spara1"
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then EtiquetaEnBD = DBLet(miRsAux!pesoetiqueta, "N")
        miRsAux.Close
    End If
            
    
    
    SQL = "Select litrosunidad,nomartic,codtipar from sartic where codartic=" & DBSet(RecuperaValor(cmdLinea(Index).Tag, 1), "T")
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then MsgBox "EOF Reg": Stop
    
    
    Me.txtPesada(0).Text = miRsAux!nomartic
    txtPesada(0).Tag = ""
    TextoAux = "Volumen " & miRsAux!litrosunidad & " Litros  "
    VolumenProd = miRsAux!litrosunidad * 1000  'Ml, y gramos
    
    EsUnaLata2 = DBLet(miRsAux!codtipar, "T") = "09"
    
    miRsAux.Close
    
    txtPesada(1).Text = RecuperaValor(cmdLinea(Index).Tag, 2)
    txtPesada(2).Text = ""
    txtPesada(3).Text = 50 'Unidaes a pesar
    txtPesada(4).Text = 0 'Unidaes a pesar
    lw1.ListItems.Clear
    
    
    'Vamos a leer pesos
    SQL = "select codigo,idlin,numlote,tipartic,nomartic,prodtrazcompo.codartic from prodtrazcompo,sartic  where prodtrazcompo.codartic=sartic.codartic AND"
    SQL = SQL & " codigo= " & RecuperaValor(cmdLinea(Index).Tag, 3) & "  and idlin=" & RecuperaValor(cmdLinea(Index).Tag, 4)
    SQL = SQL & " AND lotetraza =" & DBSet(RecuperaValor(cmdLinea(Index).Tag, 2), "T")
    SQL = SQL & " and tipartic in (2,3) ORDER BY tipartic"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    EsperoBotella = True
    N = 0
    PesoBotella = -1
    PesoTapon = -1
    Retractil = 0
    Etiqueta2 = EtiquetaEnBD  'Valor enla BD
    'Si el formato es LATA, entonces ponemos el peso a 0
    If EsUnaLata2 Then Etiqueta2 = 0
    
    EMP_ = DiferenciaPermitidaFormato2(VolumenProd)
    
    
    
    If Not miRsAux.EOF Then
        '  prodlinpesos(codigo,idlin,serie,pesoBotella,pesoTapon,pesoEtiqueta,pesoOtro)
        ParaElInsertSQL = miRsAux!codigo & "," & miRsAux!idlin & ","
        While Not miRsAux.EOF
            If N > 1 Then
                MsgBox "Mas de dos componentes tapon-botella", vbExclamation
            Else
                If miRsAux!tipartic = 2 Then
                    SinPesarBotellas = True
                    EsperoBotella = False 'ESTA es la botella
                    If DBLet(miRsAux!numlote, "T") = "" Then
                        TextoAux = TextoAux & "Bot. sin lote"
                    Else
                        'Veremos PESO
                        SQL = "select  avg(peso),count(*) from spartidaspesos ,spartidas where id=idpartida and "
                        SQL = SQL & " codartic=" & DBSet(miRsAux!codartic, "T") & " and numlote=" & DBSet(miRsAux!numlote, "T")
                        RT.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                        If RT.EOF Then
                            TextoAux = TextoAux & "Botella sin peso(2)"
                            SinPesarBotellas = False
                        Else
                            If DBLet(RT.Fields(1), "N") = 0 Then
                                TextoAux = TextoAux & "  Bot.(" & miRsAux!numlote & ") Botella sin peso(3)"
                                SinPesarBotellas = False
                            Else
                                PesoBotella = RT.Fields(0)
                                TextoAux = TextoAux & "  Bot. " & PesoBotella
                            End If
                        End If
                        RT.Close
                    End If
                    
                    If Not SinPesarBotellas Then
                        SQL = "insert into `slog` (`fecha`,`accion`,`usuario`,`pc`,`descripcion`) values ( "
                        SQL = SQL & " now(),14,'Operario','BASCULA',"
                        SQL = SQL & DBSet("Sin peso botella: " & RecuperaValor(cmdLinea(Index).Tag, 2), "T") & ")"
                        EjecutaSQL SQL, False
                        
                        SQL = String(30, "*") & vbCrLf & vbCrLf
                        
                        SQL = SQL & "Botella sin pesar" & vbCrLf & vbCrLf & SQL
                        MsgBox SQL, vbExclamation
                    End If
                End If
                If miRsAux!tipartic = 3 Then
                    'TAPON
                    If EsperoBotella Then TextoAux = TextoAux & " ERROR botella "
                    
                    If DBLet(miRsAux!numlote, "T") = "" Then
                        TextoAux = TextoAux & "Tapon sin lote"
                    Else
                        'Veremos PESO
                        'Veremos PESO
                        SQL = "select  avg(peso),count(*) from spartidaspesos ,spartidas where id=idpartida and "
                        SQL = SQL & " codartic=" & DBSet(miRsAux!codartic, "T") & " and numlote=" & DBSet(miRsAux!numlote, "T")
                        RT.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                        If RT.EOF Then
                            TextoAux = TextoAux & "  -   Tap sin peso(2)"
                        Else
                            If DBLet(RT.Fields(1), "N") = 0 Then
                                TextoAux = TextoAux & "    -   Tap.(" & miRsAux!numlote & ")"
                            Else
                                PesoTapon = RT.Fields(0)
                                TextoAux = TextoAux & "    - Tap.: " & PesoTapon
                            End If
                        End If
                        RT.Close
                        
                    End If
                    
                    
                End If
            End If
            miRsAux.MoveNext
            N = N + 1
        Wend
    End If
    miRsAux.Close
    txtPesada(5).Text = TextoAux
    txtPesada(5).Visible = True
    lblPeso.Caption = "0,0000"
    SQL = "Select max(serie) from prodlinpesos where "
    SQL = SQL & " codigo= " & RecuperaValor(cmdLinea(Index).Tag, 3) & "  and idlin=" & RecuperaValor(cmdLinea(Index).Tag, 4)
    SQL = SQL & " AND lotetraza= " & RecuperaValor(cmdLinea(Index).Tag, 2)

    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = "1"
    If Not miRsAux.EOF Then SQL = DBLet(miRsAux.Fields(0), "N") + 1
    miRsAux.Close
    txtPesada(2).Text = SQL
 
        
    
    
    'Retractil
    SQL = "select pesoneto,tipartic,nomartic,sartic.codartic from sarti4,sarti1,sartic where sarti4.codartic = sarti1.codarti1 And sartic.codartic = sarti4.codartic"
    SQL = SQL & " AND tipartic=8 and sarti1.codartic=" & DBSet(RecuperaValor(cmdLinea(Index).Tag, 1), "T")
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then Retractil = DBLet(miRsAux!pesoneto, "N")
    miRsAux.Close
    
    
    
    
    
    'Veamos si todos los datos son correctos para poder empezar a pesar
    If txtPesada(2).Text <> "" Then
        'Ok la serie de pesaje
        If N > 0 And N < 3 Then
            '1 tapon y una botella
            If PesoBotella > 0 And PesoTapon > 0 Then
                VaOK = True
                cmdValirdarLecturas(1).Tag = cmdLinea(Index).Tag
            End If
        End If
    End If
    
    
    
    
    
'
'    Stop   'QUITAR TODOS ESTOS
'    VolumenProd = 500
'    PesoTapon = 0.025
'    PesoBotella = 400
'    EMP = Val(DiferenciaPermitidaFormato(VolumenProd))

    
    'Objetivo final:
    ' Insertar aqui
    '  prodlinpesos(codigo,idlin,serie,lotetraza,pesoBotella,pesoTapon,pesoEtiqueta,pesoEtiqueta,secuencial,fechahora,pesoLleno,CumpleEMP,Cumple2EMP)
    ParaElInsertSQL = " (" & RecuperaValor(cmdLinea(Index).Tag, 3) & "," & RecuperaValor(cmdLinea(Index).Tag, 4) & "," & Me.txtPesada(2).Text & ","
    ParaElInsertSQL = ParaElInsertSQL & RecuperaValor(cmdLinea(Index).Tag, 2) & "," & DBSet(PesoBotella, "N") & "," & DBSet(PesoTapon, "N") & ","
    ParaElInsertSQL = ParaElInsertSQL & DBSet(Etiqueta2, "N") & "," & DBSet(Retractil, "N") & ","
    
    
    
    
    'prodlinpesos
    
    
    
    HabilitarPesada True
    Set miRsAux = Nothing
    Set RT = Nothing
    cmdValirdarLecturas(1).Visible = VaOK
    If VaOK Then
        Timer11.Enabled = True
        Timer2.Enabled = False
    End If
End Sub

Private Sub cmdMatAux_Click(Index As Integer)
Dim Formato As String  'para las pruebas

    'Ponemos a pesar
    Set miRsAux = New ADODB.Recordset
    
    SQL = "Select id,nomartic,numlote,spartidas.codartic,codunida,codfamia from spartidas,sartic where spartidas.codartic=sartic.codartic and id=" & Me.cmdMatAux(Index).Tag
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    ClaveReferenciaPesada = Me.cmdMatAux(Index).Tag
    
    If miRsAux.EOF Then MsgBox "EOF Reg": Stop
    Me.txtPesada(0).Text = miRsAux!nomartic
    txtPesada(0).Tag = miRsAux!id
    txtPesada(1).Text = miRsAux!numlote
    txtPesada(2).Text = 1 'serie
    If UCase(Label1(4).Caption) = "TAPONES" Then
        txtPesada(3).Text = 10 'Unidaes a pesar
    Else
    
        'Botellas de plastico, pesan 10 unidades tambien
        If miRsAux!codfamia = 18 Then
            'Plastico
            txtPesada(3).Text = 10
        Else
            'Resto de envases
            txtPesada(3).Text = 25 'Unidaes a pesar
        End If
    End If
    txtPesada(4).Text = 0 'Unidaes a pesar
    txtPesada(5).Visible = False
    
    lw1.ListItems.Clear
    lblPeso.Caption = "0,0000"
    
    
    'En pruebas el peso viene de un randmo, leere de la ficha tecnica
    SQL = miRsAux!codartic
    Formato = miRsAux!codunida
    
    'Para el INSERT
    'spartidaspesos(idpartida,secuencial,fechahora,peso)
    'tmppesadasBascula(codigo,secuencial,fechahora,pesoLleno)
    ParaElInsertSQL = " (" & miRsAux!id & ","
    CadenaINsertFinal = ""
    HabilitarPesada True
    cmdValirdarLecturas(1).Visible = True
    cmdValirdarLecturas(1).Tag = cmdMatAux(Index).Tag  'Me guardo para luego saber dentro de COLMA cual era
    miRsAux.Close
    
    If Pruebas Then
        'En sql he guardado el codartic
        PesoDePrueba = 0
        SQL = "select * from sarti4 where codartic = " & DBSet(SQL, "T")
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        If Not miRsAux.EOF Then
            If Not IsNull(miRsAux!pesoneto) Then PesoDePrueba = miRsAux!pesoneto
        End If
        miRsAux.Close
        
        
        'Si es CER=, para las pruebas, saco la media del formato
        If PesoDePrueba = 0 Then
            If UCase(Label1(4).Caption) = "TAPONES" Then
                SQL = 32
            Else
                SQL = 24
            End If
            SQL = "select avg(pesoneto) from sarti4,sartic where sarti4.codartic=sartic.codartic and codunida=" & Formato & " and codmarca=" & SQL
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            If Not miRsAux.EOF Then PesoDePrueba = DBLet(miRsAux.Fields(0), "N")
            miRsAux.Close
        End If
        
        PesoDePrueba = PesoDePrueba * 1000
    End If
    
    Timer11.Enabled = True
    Timer2.Enabled = False
    Set miRsAux = Nothing
    
End Sub

Private Sub cmdMoverVector_Click(Index As Integer)
    If Index = 0 Then
        If Puntero < 5 Then
            Puntero = 0
        Else
            Puntero = Puntero - 5
        End If
    Else
        If Puntero >= colMA.Count - 1 Then
            Puntero = colMA.Count - 1
        Else
            Puntero = Puntero + 5
            If Puntero >= colMA.Count - 1 Then Puntero = colMA.Count - 1
        End If
    End If
    PonerDatosEnBotonesMaux
    LanzarProtector = 0
End Sub

Private Sub cmdValirdarLecturas_Click(Index As Integer)
Dim J As Integer


    '50 pesadas mas
    If Index = 2 Then
        FrameResultados.Visible = False
        Me.txtPesada(3).Text = 100
        Timer11.Enabled = True
        Exit Sub
    End If


    If Index = 1 Then
        
        
        If ResultadoPesadasIncorrecto Then MsgBox "Los datos serán guardados. Avise al responsable que hay errores", vbExclamation
        
        
        
        
        'INSERTAMOS EN peso y mostraremos el resumen
        If Me.txtPesada(5).Visible Then
            'Produccion
            
            SQL = "codigo,idlin,serie,pesoBotella,pesoTapon,pesoEtiqueta,pesoOtro,secuencial,fechahora,pesoLleno,volumenLlenado,EMP,CumpleEMP,Cumple2EMP,lotetraza"
            SQL = "INSERT INTO prodlinpesos(" & SQL & ") SELECT " & SQL & " FROM tmppesadasBascula"

            
        Else
        
            SQL = "INSERT INTO spartidaspesos(idpartida,secuencial,fechahora,peso)  "
            SQL = SQL & "Select codigo,secuencial,fechahora,pesoLleno from  tmppesadasBascula order by secuencial"
        End If
        Conn.Execute SQL
        
        If txtPesada(5).Visible = True Then
            'Hemos pesado produccion
            
        Else
            For J = colMA.Count To 1 Step -1
                If cmdValirdarLecturas(1).Tag = RecuperaValor(colMA(J), 1) Then
                    colMA.Remove J
                    Exit For
                End If
            Next J
            
            
            PonerDatosEnBotonesMaux
            
            
        End If
    End If
    ParaElInsertSQL = ""
    CadenaINsertFinal = ""
    Timer11.Enabled = False  'Por si acaso
    If Not Pruebas Then
        If MSComm1.PortOpen Then MSComm1.PortOpen = False
    End If
    HabilitarPesada False
End Sub



Private Sub Form_Load()
Dim H As Integer
     HabilitarPesada False
     
     
     Me.Caption = "Aceites Morales. Bascula producción   Ver: " & App.Major & "." & App.Minor & "." & App.Revision
     
     H = Me.Height - Me.FramelecturaBascula.Height
     H = H / 2
     Me.FramelecturaBascula.Top = H
     H = Me.Width - Me.FramelecturaBascula.Width
     H = H / 2
     Me.FramelecturaBascula.Left = H
     
     Me.FrameResultMaAux.BorderStyle = 0
     Me.FrameResultProd.BorderStyle = 0
     FrameResultados.Visible = False
     
     MSComm1.CommPort = Config.kCOMM
     MSComm1.Settings = Config.Velocidad & ",n,8,1"
     EtiquetaEnBD = -1 'para forzar lectura peso la primera vez
     Timer2.Enabled = True
End Sub

Private Sub LineasProduccionVisible(Linea As Integer, Visible As Boolean)
   ' Me.lblLinea(Linea).ForeColor = IIf(Visible, vbBlack, vbgray)
    Me.cmdLinea(Linea).Visible = Visible
     Me.lblLinea(Linea).Visible = True
End Sub





Private Sub Image1_Click()
    CargaMateriaAuxiliar False
End Sub

Private Sub Image3_Click()
    CargaMateriaAuxiliar True
    
End Sub


Private Sub Image2_Click()
    'Leemos lo que se esta producciendo AHORA
    LineasProduccion
    'Ponemos visible el frame
    FrameMateriaAuxiliar.Visible = False
    FrameLineasProd.Visible = True
    Set colMA = Nothing
    LanzarProtector = 0
End Sub


Private Sub LineasProduccion()
Dim J As Integer
    Set miRsAux = New ADODB.Recordset
    
    For J = 0 To 9
        cmdLinea(J).Tag = ""
        Me.cmdLinea(J).Visible = False
        Me.lblLinea(J).Visible = False
        LeerLinea J
    Next J
    
    Set miRsAux = Nothing
End Sub

Private Sub LeerLinea(KLinea As Integer)
    On Error GoTo eLee
    
    
    
    SQL = "select codartic,lotetraza,prodlin.codigo,prodlin.idlin  from prodlin,prodtrazlin  where prodlin.codigo= prodtrazlin.codigo AND prodlin.idlin = prodtrazlin.idlin"
    SQL = SQL & " and lineaprod = " & KLinea & " and estado >0 and estado<10 ORDER BY lotetraza DESC"  'Pq puede que haya varios cambios de trazabilidad para la misma linea
    miRsAux.Open SQL, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        'ya tenemos que es lo que esta produciendose. Consulto nombre ...
        SQL = miRsAux.Fields(0) & "|" & miRsAux.Fields(1) & "|" & miRsAux.Fields(2) & "|" & miRsAux.Fields(3) & "|"
        cmdLinea(KLinea).Tag = SQL
        SQL = "Select nomartic from sartic where codartic=" & DBSet(miRsAux!codartic, "T")
        miRsAux.Close
        
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        cmdLinea(KLinea).Caption = miRsAux.Fields(0)  'no puede ser EOF
    
    
        Me.cmdLinea(KLinea).Visible = True
        Me.lblLinea(KLinea).Visible = True
    End If
    miRsAux.Close
    
    Exit Sub
eLee:
    MuestraError Err.Number
    Set miRsAux = Nothing
    Set miRsAux = New ADODB.Recordset
End Sub




Private Sub CargaMateriaAuxiliar(Tapones As Boolean)
    
    If Tapones Then
        SQL = "3"
        Label1(4).Caption = "Tapones"
    Else
        SQL = "2"
        Label1(4).Caption = "Envases"
    End If
    SQL = "sartic.codartic=spartidas.codartic and tipartic=" & SQL & " and cantotal<>0"
    SQL = "select nomartic,id,numlote  from spartidas,sartic where " & SQL
    SQL = SQL & " AND not id IN (select distinct idpartida from spartidaspesos)"
    SQL = SQL & " order by fecha desc"
    Set miRsAux = New ADODB.Recordset
    Set colMA = New Collection
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        
        colMA.Add miRsAux!id & "|" & miRsAux!nomartic & "  (" & miRsAux!numlote & ")|"
        
        'pruebas
        'colMA.Add colMA.Count + 1 & "|" & colMA.Count + 1 & "|"
        'If colMA.Count > 52 Then
        '    While Not miRsAux.EOF
        '        miRsAux.MoveNext
        '    Wend
        'Else
        
            miRsAux.MoveNext
       ' End If
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    Me.FrameMateriaAuxiliar.Visible = True
    Puntero = 0
    PonerDatosEnBotonesMaux
    LanzarProtector = 0
    
    
    

End Sub

Private Sub PonerDatosEnBotonesMaux()
Dim N As Integer

    For N = 0 To 4
        If Puntero + N >= colMA.Count Then
            Me.cmdMatAux(N).Visible = False
            Me.cmdMatAux(N).Tag = ""
        Else
            Me.cmdMatAux(N).Tag = RecuperaValor(colMA(Puntero + 1 + N), 1)
            Me.cmdMatAux(N).Caption = RecuperaValor(colMA(Puntero + 1 + N), 2)
            Me.cmdMatAux(N).Visible = True
        End If
    Next N
        
    
End Sub



Private Sub HabilitarPesada(Si As Boolean)
    Me.FrameBotella.Enabled = Not Si
    Me.FrameLineasProd.Enabled = Not Si
    Me.FrameLineasProd.Enabled = Not Si
    Me.FrameMateriaAuxiliar.Enabled = Not Si
    Me.FramelecturaBascula.Visible = Si
    If Not Si Then
        Timer11.Enabled = False
        If Not Pruebas Then
            If MSComm1.PortOpen Then MSComm1.PortOpen = False
        End If
        FrameLineasProd.Visible = False
        LanzarProtector = 0
        Timer2.Enabled = True
    Else
        'Borro datos anteriores
        Timer2.Enabled = False
        FrameResultados.Visible = False
        Conn.Execute "DELETE from tmppesadasBascula"
        Buffer = ""
        If Not Pruebas Then
            If Not MSComm1.PortOpen Then PuertoComm True
        End If
    End If
    
End Sub









Private Sub lw1_KeyPress(KeyAscii As Integer)
Dim cad As String
Dim Peso As Currency
Dim Ok1 As Boolean
Dim Ok2 As Boolean


    If lw1.SelectedItem Is Nothing Then Exit Sub
    
    'If lw1.SelectedItem.ForeColor = vbBlack Then Exit Sub
    
    
    
    'La tecla tiene que ser la P
    If Chr(KeyAscii) = "P" Or Chr(KeyAscii) = "p" Then
    
    
        cad = InputBox("Pes.", "Valor", lw1.SelectedItem.Text)
        If cad = "" Then Exit Sub
        
        Peso = CCur(cad)
        lw1.SelectedItem.Text = Peso
        
        If Me.txtPesada(5).Visible Then
            cad = "UPDATE tmppesadasBascula set pesolleno=" & DBSet(Peso, "N")
            Peso = (Peso - (PesoBotella + PesoTapon + Etiqueta2 + Retractil))
            Peso = Peso / 0.916
            cad = cad & ", volumenllenado=" & DBSet(Peso, "N")
            Ok1 = True
            Ok2 = True
            If VolumenProd - Peso > EMP_ Then
                lw1.SelectedItem.Bold = True
                If VolumenProd - Peso > 2 * EMP_ Then
                    lw1.SelectedItem.ForeColor = vbRed
                    Ok2 = False
                Else
                    lw1.SelectedItem.ForeColor = vbGreen
                    Ok1 = False
                End If
    
            Else
                lw1.SelectedItem.ForeColor = vbBlack
                lw1.SelectedItem.Bold = False
            End If
             
        Else
        
            lw1.SelectedItem.ForeColor = vbBlack
            lw1.SelectedItem.Bold = False
            Ok1 = True
            Ok2 = True
        End If
        
        cad = cad & ",CumpleEMP=" & Abs(Ok1)
        cad = cad & ",Cumple2EMP=" & Abs(Ok2)
        cad = cad & " WHERE secuencial=" & lw1.SelectedItem.Index
        Conn.Execute cad
        espera 0.5
        
        HacerTotales
    
    End If
    
End Sub

'---------------------------------------------------------------------
' Sobre la bascula. COMM

Private Static Sub MSComm1_OnComm_NO_LO_UTILIZO()
    Dim EVMsg$
    Dim ERMsg$
    Dim Aux As String
    ' Bifurca según la propiedad CommEvent.
    Select Case MSComm1.CommEvent
        ' Mensajes de evento.
        Case comEvReceive
            Aux = MSComm1.Input
            Aux = StrConv(Buffer, vbUnicode)
            Buffer = Buffer & Aux
            
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
'''        ' Muestra los mensajes de evento en la barra de estado.
'''        sbrStatus.Panels("Status").Text = "Estado:" & EVMsg$
'''
'''        ' Activa el cronómetro para que el mensaje de la barra
'''        ' de estado se borre después de dos segundos.
'''        Timer2.Enabled = True
        
    ElseIf Len(ERMsg$) Then
        ' Muestra los mensajes de evento en la barra de estado.
'        sbrStatus.Panels("Status").Text = "Estado:" & ERMsg$
'
'        ' Muestra los mensajes de error en un cuadro de alerta.
'        Beep
'        Ret = MsgBox(ERMsg$, 1, "Haga clic en Cancelar para salir, clic en Aceptar para ignorar.")
'
'        ' Si el usuario hace clic en Cancelar (2)...
'        If Ret = 2 Then
'            MSComm1.PortOpen = False    ' Cierra el puerto y sale.
'        End If
'
'        ' Activa el cronómetro para que el mensaje de la barra
'        ' de estado se borre después de dos segundos.
'        Timer2.Enabled = True
    End If
End Sub



Private Sub PuertoComm(Abrir As Boolean)
        
    
    If Pruebas Then Exit Sub
    MSComm1.PortOpen = True
    MSComm1.InputLen = 0

        
End Sub


Private Sub Timer11_Timer()
Dim PesoInterMedio As Currency
Dim TeoricoPruebas As Currency
Dim OtroAux As Currency
Dim K As Integer
Dim Ok As Boolean
Dim J As Integer
Dim cad As String

    Timer11.Enabled = False
    If Pruebas Then
    
        If Me.txtPesada(5).Visible Then
            'Produccion
            'Ejemplo ramon
            If False Then
                    Select Case Val(txtPesada(4).Text)
                    Case 0
                        Buffer = 838
                    Case 1
                        Buffer = 836
                    Case 2
                        Buffer = 902
                    Case 3
                        Buffer = 837  'Si queremos poner un tercer EMP
                    Case Else
                        Buffer = 856
                    End Select
            Else
                'Para cualquier ejemplo
               
                PesoInterMedio = Retractil + Etiqueta2 + PesoBotella + PesoTapon
                
                
                'Uno de cada mil casos, 2*E;P
                If Int((1000 * Rnd) + 1) >= 1000 Then
                    TeoricoPruebas = VolumenProd - (2 * EMP_) - Rnd - 0.45

                Else
                    'Uno de cada 50 datos es EMP
                    If Int((55 * Rnd) + 1) >= 54 Then
                        TeoricoPruebas = VolumenProd - EMP_ - Rnd - 0.1
                        
                    Else
                        'Peso OK
                        'Cada 5 botellas, 1 peasra un poco menos
                        If Int((5 * Rnd) + 1) > 1 Then
                            TeoricoPruebas = VolumenProd + Rnd
                        Else
                            OtroAux = Rnd
                            If OtroAux * 100 > 50 Then OtroAux = OtroAux - 0.45
                            TeoricoPruebas = VolumenProd - Rnd
                        End If
                    End If
                End If
                TeoricoPruebas = (TeoricoPruebas * 0.916)
                PesoInterMedio = TeoricoPruebas + PesoInterMedio
                Buffer = PesoInterMedio
            End If
        Else
            'Materia auxiliar
            If PesoDePrueba < 10 Then
                PesoInterMedio = Rnd / 100
            Else
                PesoInterMedio = Rnd / 3
            End If
            Buffer = PesoDePrueba + PesoInterMedio
        End If
    Else
    
        
    
        Buffer = ""
        
        If MSComm1.PortOpen Then
        
            MSComm1.Output = "P" + Chr(13) + Chr(10)
            espera 0.3
            K = 0
            Ok = False
            Do
                cad = MSComm1.Input
                Buffer = Buffer & cad
                K = K + 1
                espera 0.2
                DoEvents
            
                If Buffer <> "" Then
                    K = 0
                    Buffer = Trim(Buffer)
                    J = InStr(1, Buffer, "g")
                    If J > 0 Then
                        
                        If InStr(1, Buffer, "?") > 0 Then
                            'Ok Ha llegado el peso
                            'Pero NO es estable
                            Buffer = ""
                            Ok = True
                        Else
                            Buffer = Mid(Buffer, 1, J - 1)
                            If Not IsNumeric(Buffer) Then
                                MsgBox "Campo no numerico: " & Buffer, vbExclamation
                            Else
                                Buffer = TransformaPuntosComas(Buffer)
                                If CCur(Buffer) = 0 Then Buffer = ""
                                Ok = True
                            End If
                        End If
                    End If
                Else
                    'If K > 5 Then
                    '    MsgBox "No llega dato", vbExclamation
                    '    espera 0.5
                    '    Buffer = ""
                    '    Ok = True
                    'End If
                End If
            
            
            Loop Until Ok
        
        
        End If
    
    End If
    
    
    If Buffer <> "" Then
        'ha llegado un peso
        ProcesarBuffer
        
        espera 0.4
        lblPeso.Caption = ""
        If Pruebas Then
            espera 0.5
        Else
            For K = 1 To 3
                espera 0.8
                lblPeso.Caption = lblPeso.Caption & "."
                lblPeso.Refresh
            Next
        End If
        Me.lblPeso.BackColor = 16639728
        lblPeso.ForeColor = vbBlack
        
        
        
        
        lblPeso.Caption = ""
        If Val(txtPesada(4).Text) >= Val(txtPesada(3).Text) Then
            'Ya estan todas las pesadas
             
            'Pondremos frame visible
            HacerTotales
            Exit Sub
        End If
    End If
    Timer11.Enabled = True
    Timer2.Enabled = False
End Sub


Private Sub HacerTotales()
Dim Valor As Currency
Dim EmpOk As Boolean
Dim EmpOk2 As Boolean
Dim MediaOK As Boolean
Dim AmpliarNumeroPesadas As Boolean








    


    FrameResultProd.Visible = False
    Me.FrameResultMaAux.Visible = False
    Set miRsAux = New ADODB.Recordset
    cmdValirdarLecturas(1).Visible = False
    cmdValirdarLecturas(2).Visible = False
    AmpliarNumeroPesadas = False
    If txtPesada(5).Visible Then
        'Produccion
        
        'Volumen llenado
        'Maximo peso permitido por formato
        
        
        
        SQL = CStr(DiferenciaPermitidaFormato2(VolumenProd))
        Valor = CCur(SQL)
        
        SQL = "UPDATE tmppesadasBascula SET EMP =" & DBSet(Valor, "N")
        SQL = SQL & ", volumenllenado = (pesolleno- (pesobotella+pesotapon +pesoetiqueta+pesootro)) / 0.916 "
        Conn.Execute SQL
        espera 0.5
        
      
 
        SQL = "select avg(volumenLlenado),std(volumenLlenado),sum(if(CumpleEMP=0,1,0)),sum(if(Cumple2EMP=0,1,0)),avg(pesoLleno) from tmppesadasbascula"
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        'No puede ser eof
        txtPesada(8).Text = miRsAux.Fields(2)  'no emp
        txtPesada(9).Text = miRsAux.Fields(3) 'no dobleemp
        txtPesada(10).Text = Round(miRsAux.Fields(0), 2)
        txtPesada(11).Text = Round(miRsAux.Fields(1), 4) 'desviacion
        txtPesada(12).Text = miRsAux.Fields(4) 'no dobleemp
        
        
      
        
        EmpOk2 = True
        EmpOk = True
        MediaOK = True
        
        
        If Val(txtPesada(9).Text) > 0 Then
            'ERROR GRAVE. Lote mal pesado
            'NO PUEDE guardarse, hay un valor que excede de EMp*2
            EmpOk2 = False
           ' MsgBox "Error grave 2*EMP", vbCritical
        End If
        
        'Emp
        If Val(txtPesada(10).Text) > 0 Then
            If Val(txtPesada(3).Text) = 50 Then
                'Primerea serie. Haremos una segunda prueba de muestreo
                If Val(txtPesada(8).Text) > 2 Then
                    If Val(txtPesada(8).Text) >= 5 Then
                        'LOTE RECHAZAdo
                        EmpOk = False
                    Else
                        'Ampliar muestra
                        'Ampliaremos muestra
                        EmpOk = False
                        
                        'haremos 50 mediciones mas
                        cmdValirdarLecturas(2).Visible = True
                        AmpliarNumeroPesadas = True
                    End If
                    
                Else
                    'empok=true
                End If
            Else
                'Segundo muestreo. Total 100 uds
                If Val(txtPesada(8).Text) >= 7 Then
                    'lote rechazado
                    EmpOk = False
                End If
    
            End If
        End If
    
        
        Me.txtPesada(8).BackColor = IIf(EmpOk, &HC0FFC0, &HC0C0FF)
        Me.txtPesada(9).BackColor = IIf(EmpOk2, &HC0FFC0, &HC0C0FF)
        
        'Falataria ver por desviacion
        'SI(media>=formato-(0,379*desviacion);"CONFORME";"NO CONFORME")
        Valor = 0.379 * miRsAux.Fields(1)
        Valor = VolumenProd - Valor
        
        If miRsAux.Fields(0) > Valor Then
            MediaOK = True
        Else
            MediaOK = False
        End If
        Me.txtPesada(11).BackColor = IIf(MediaOK, &HC0FFC0, &HC0C0FF)
        miRsAux.Close
        
        
        If EmpOk And EmpOk2 And MediaOK Then
            cmdValirdarLecturas(1).Visible = True
            ResultadoPesadasIncorrecto = False
        Else
            ResultadoPesadasIncorrecto = True
            If Not AmpliarNumeroPesadas Then cmdValirdarLecturas(1).Visible = True
        End If
        FrameResultProd.Visible = True
        FrameResultados.Visible = True
        
    Else
        ResultadoPesadasIncorrecto = False
    
        'Materia auxiliar
        SQL = "select avg(pesolleno),std(pesolleno) from tmppesadasBascula"
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        'No puede ser eof
        txtPesada(6).Text = miRsAux.Fields(0)
        txtPesada(7).Text = miRsAux.Fields(1) 'desviacion
    
    
        'Si es BOTELLA veo la desveiacion frente al formato
        'Vere el volumen
        If UCase(Label1(4).Caption) <> "TAPONES" Then
            miRsAux.Close
            SQL = "Select codartic from spartidas where id=" & ClaveReferenciaPesada
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            SQL = miRsAux!codartic
            miRsAux.Close
            SQL = "select litrosunidad from sartic,sarti1 where sartic.codartic=sarti1.codartic and codarti1 =" & DBSet(SQL, "T") & " ORDER BY 1"
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Valor = -1
            CadenaINsertFinal = ""
            SQL = ""
            While Not miRsAux.EOF
                'DEBERIAN SER TODOS DEL MISMO FORMATO (50,250...)
                If Valor <> miRsAux.Fields(0) Then
                    Valor = miRsAux.Fields(0)
                    SQL = SQL & "X"
                    CadenaINsertFinal = CadenaINsertFinal & "  --  " & miRsAux.Fields(0)
                End If
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            
            If Len(SQL) <> 1 Then
                CadenaINsertFinal = Mid(CadenaINsertFinal, 5)
                Valor = InStr(1, CadenaINsertFinal, "--")
                SQL = Trim(Mid(CadenaINsertFinal, 1, CInt(Valor) - 1))
                CadenaINsertFinal = "Error en formatos. Introduzca los LITROS del envase : " & vbCrLf & CadenaINsertFinal
           
                Valor = 0
                Do
                    SQL = InputBox(CadenaINsertFinal, , SQL)
                    If SQL = "" Then
                        Valor = -1
                    Else
                        Valor = CCur(SQL)
                        
                        If MsgBox("Formato envase: " & Valor & " LITROS ?", vbQuestion + vbYesNo) = vbYes Then SQL = ""
                    End If
                    
                Loop Until SQL = ""
                If Valor = -1 Then Exit Sub
            End If
                
            'La desviacion tuene que ser inferior a 0.25 el maximo por formato
            Valor = Valor * 1000
            SQL = CStr(DiferenciaPermitidaFormato2(Valor))
            Valor = CCur(SQL) * Valor
                
            If CCur(txtPesada(7).Text) > Valor Then
                'ERRROR. Se va. NO CONFORME
                EmpOk = False
                ResultadoPesadasIncorrecto = True
            Else
                'Lote OK
               ' MsgBox "OK"
               EmpOk = True
            End If
        Else
            EmpOk = True
            
        End If
        
        Me.txtPesada(7).BackColor = IIf(EmpOk, &HC0FFC0, &HC0C0FF)
        Me.FrameResultMaAux.Visible = True
         Me.FrameResultados.Visible = True
         
        cmdValirdarLecturas(1).Visible = True
    End If
    Set miRsAux = Nothing
   
    
End Sub

Private Function DiferenciaPermitidaFormato2(Formato As Currency) As Currency
Dim Aux As Currency
    'Entre 0.. 499 = 9
    '      500.1000=15
    '      1001 en adelante= 15% formato  -->
    '           2000=30
    '           2500=37.5
    '           4000=60
    '           5000=75
               
    
    If Formato < 500 Then
        DiferenciaPermitidaFormato2 = 9
    Else
        If Formato <= 1000 Then
            DiferenciaPermitidaFormato2 = 15
        Else
            Aux = Formato / 1000
            Aux = Aux * 15  'el 15 %
            DiferenciaPermitidaFormato2 = Round(Aux, 1)
        End If
    End If
End Function


Private Sub ProcesarBuffer()
Dim Peso As Currency
Dim Ok1 As Boolean
Dim Ok2 As Boolean
Dim K As Integer

    If IsNumeric(Buffer) Then
        'Es un numero
                
        lblPeso.Caption = Format(Buffer, "0.00")
        lblPeso.Refresh
        
        
        If Me.txtPesada(5).Visible Then
            'Produccion
            CadenaINsertFinal = "INSERT INTO tmppesadasBascula(codigo,idlin,serie,lotetraza,pesoBotella,pesoTapon,pesoEtiqueta,pesoOtro,secuencial,fechahora,pesoLleno,CumpleEMP,Cumple2EMP) VALUES "
            
        Else
            'ParaElInsertSQL
            'tmppesadasBascula(codigo,secuencial,fechahora,pesoLleno)
            CadenaINsertFinal = "INSERT INTO tmppesadasBascula(codigo,secuencial,fechahora,pesoLleno) VALUES "
        End If
        CadenaINsertFinal = CadenaINsertFinal & ParaElInsertSQL & lw1.ListItems.Count + 1 & "," & DBSet(Now, "FH") & "," & DBSet(CCur(lblPeso.Caption), "N")
        
        
         
         
         
        lw1.ListItems.Add , , lblPeso.Caption
        K = lw1.ListItems.Count
        If Me.txtPesada(5).Visible Then
            'Para cada BOTELLA comprobamos, YA MISMO, el EMP  y DOBLE EMP
            'Es de parametros y para todos los formatos
            'Maximo peso permitido por formato
            Ok1 = True
            Ok2 = True
            
            Peso = CCur(Buffer)
            Peso = (Peso - (PesoBotella + PesoTapon + Etiqueta2 + Retractil))
            Peso = Peso / 0.916
            
            
            If VolumenProd - Peso > EMP_ Then
                Beep
                If VolumenProd - Peso > 2 * EMP_ Then
                    Me.lblPeso.BackColor = vbRed
                    lblPeso.ForeColor = vbWhite
                    lblPeso.Refresh
                    espera 1
                    Ok2 = False
                    Beep
                Else
                    Me.lblPeso.BackColor = vbYellow
                    lblPeso.ForeColor = vbBlack
                    Ok1 = False
                End If
                lblPeso.Refresh

                espera 1
                
                lw1.ListItems(K).ForeColor = IIf(Not Ok2, vbRed, vbGreen)
                lw1.ListItems(K).Bold = True
            Else
                Me.lblPeso.BackColor = 16639728
                lblPeso.ForeColor = vbBlack
            End If
            
            'CumpleEMP,Cumple2EMP
            CadenaINsertFinal = CadenaINsertFinal & "," & Abs(Ok1) & "," & Abs(Ok2)
            
         Else
            'Nada
        End If
        
        Conn.Execute CadenaINsertFinal & ")"
        
        txtPesada(4).Text = lw1.ListItems.Count
        txtPesada(4).Refresh
         
    End If
End Sub

Private Sub Timer2_Timer()
    
    LanzarProtector = LanzarProtector + 1
    If LanzarProtector > Config.Segundos Then
        Timer2.Enabled = False
        LanzarProtector = 0
        frmProtector.Show vbModal
        Me.FrameLineasProd.Visible = False
        Timer2.Enabled = True
    End If
End Sub
