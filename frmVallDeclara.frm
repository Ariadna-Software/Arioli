VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmVallDeclara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Declaración mensual"
   ClientHeight    =   10980
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   18240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10980
   ScaleWidth      =   18240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   1680
      TabIndex        =   52
      Top             =   10320
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   495
      Left            =   240
      TabIndex        =   51
      Top             =   10320
      Width           =   1335
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar periodo"
      Height          =   495
      Left            =   6240
      TabIndex        =   50
      Top             =   10320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   48
      Text            =   "Text2"
      Top             =   9120
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   27
      Left            =   14640
      Locked          =   -1  'True
      TabIndex        =   42
      Text            =   "Text1"
      Top             =   10560
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   26
      Left            =   14400
      Locked          =   -1  'True
      TabIndex        =   41
      Text            =   "Text1"
      Top             =   10560
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   25
      Left            =   13800
      Locked          =   -1  'True
      TabIndex        =   40
      Text            =   "Text1"
      Top             =   10560
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   24
      Left            =   12480
      Locked          =   -1  'True
      TabIndex        =   39
      Text            =   "Text1"
      Top             =   10560
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   23
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   38
      Text            =   "Text1"
      Top             =   10560
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   22
      Left            =   9000
      Locked          =   -1  'True
      TabIndex        =   37
      Text            =   "Text1"
      Top             =   10560
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   21
      Left            =   13320
      Locked          =   -1  'True
      TabIndex        =   36
      Text            =   "Text1"
      Top             =   10080
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   20
      Left            =   16200
      Locked          =   -1  'True
      TabIndex        =   35
      Text            =   "Text1"
      Top             =   10080
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   19
      Left            =   14640
      Locked          =   -1  'True
      TabIndex        =   34
      Text            =   "Text1"
      Top             =   10080
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   12240
      Locked          =   -1  'True
      TabIndex        =   33
      Text            =   "Text1"
      Top             =   10080
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   32
      Text            =   "Text1"
      Top             =   10080
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   9000
      Locked          =   -1  'True
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   10080
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4800
      Top             =   3120
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   9615
      Left            =   7920
      TabIndex        =   26
      Top             =   360
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   16960
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   5160
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   5160
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   5160
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   5160
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00404000&
      BorderWidth     =   5
      X1              =   2520
      X2              =   7560
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label Label1 
      Caption         =   "Observaciones"
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
      Index           =   18
      Left            =   240
      TabIndex        =   49
      Top             =   8760
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Mes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   17
      Left            =   7920
      TabIndex        =   47
      Top             =   10200
      Width           =   450
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   7440
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Label Label1 
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   120
      TabIndex        =   46
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Mes actual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   120
      TabIndex        =   45
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Hasta mes  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   14
      Left            =   120
      TabIndex        =   44
      Top             =   4080
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   13
      Left            =   7920
      TabIndex        =   43
      Top             =   10560
      Width           =   555
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404000&
      BorderWidth     =   5
      X1              =   2520
      X2              =   7560
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "Salidas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Index           =   12
      Left            =   15720
      TabIndex        =   30
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Productos obtenidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Index           =   11
      Left            =   11880
      TabIndex        =   29
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Aceitunas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   375
      Index           =   10
      Left            =   9120
      TabIndex        =   28
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label LabelFechaMes 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Movimientos aceite"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   27
      Top             =   240
      Width           =   6735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   120
      Top             =   120
      Width           =   7575
   End
   Begin VB.Label Label1 
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   6480
      TabIndex        =   13
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Alm . y patrimo."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   4440
      TabIndex        =   12
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Otras entidades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   2880
      TabIndex        =   11
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Envasa. propia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   1320
      TabIndex        =   10
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Detalles de aceite"
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
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Stock final"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   5760
      TabIndex        =   8
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Salidas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3960
      TabIndex        =   7
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Producido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2280
      TabIndex        =   6
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Existencias ant"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   480
      TabIndex        =   5
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Movimientos aceite"
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
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   2535
   End
End
Attribute VB_Name = "frmVallDeclara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Cad As String



Private Sub Command2_Click()

End Sub

Private Sub cmdCerrar_Click()
    
    
    'Algunas cosillas
    '1.- Todas las entradas de oliva estan "albaranadas"
    
    Cad = "Select count(*) from vallentradacamion where year(fechaentrada)=" & Year(LabelFechaMes.Tag)
    Cad = Cad & " AND entradafinalizada=0 "
    Cad = Cad & " AND month(fechaentrada)=" & Month(LabelFechaMes.Tag)
    miRsAux.Open Cad, conn, adOpenKeyset, adCmdText
    Cad = ""
    If Not miRsAux.EOF Then
        Cad = DBLet(miRsAux.Fields(0), "N")
        If Val(Cad) > 0 Then
            Cad = "Existen entradas de oliva pendiente de generar el albaran"
        Else
            Cad = ""
        End If
    End If
    miRsAux.Close
    If Cad <> "" Then
        MsgBox Cad, vbExclamation
        Exit Sub
    End If
    
    'Todos los procesos de molturacin INICIO en el mes estan cerrados
    
    Cad = "Select count(*) from vallalmazaraproceso where year(fecha)=" & Year(LabelFechaMes.Tag)
    Cad = Cad & " AND month(fecha)=" & Month(LabelFechaMes.Tag) & " and fechafin<>null"
    miRsAux.Open Cad, conn, adOpenKeyset, adCmdText
    Cad = ""
    If Not miRsAux.EOF Then
        Cad = DBLet(miRsAux.Fields(0), "N")
        If Val(Cad) > 0 Then
            Cad = "Existen procesos de almazara pendiente de molturar"
        Else
            Cad = ""
        End If
    End If
    miRsAux.Close
    If Cad <> "" Then
        MsgBox Cad, vbExclamation
        Exit Sub
    End If
    If MsgBox("Desea cerrar el periodo?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    conn.BeginTrans
    If CerrarProcesoMensual Then
        conn.CommitTrans
        Unload Me
    Else
        conn.RollbackTrans
    End If
End Sub

Private Sub cmdImprimir_Click()

    Cad = "{tmpnlotes.codusu}=" & vUsu.Codigo
     
     LlamaImprimirGral Cad, "", 0, "valldeclaramenActu.rpt", "Declaracion mensual(Previa)"
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Load()
Dim Importe As Currency
Dim Importe2 As Currency
Dim UltFecPresentada As Date

Dim I As Integer

    Me.Icon = frmppal.Icon
    
    limpiar Me
    'Cargamos todos los campos
    Set miRsAux = New ADODB.Recordset
    UltFecPresentada = vParamAplic.FechaActiva
    Cad = "Select max(anyo * 100 + mes) from valldeclara  order by 1 desc "
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        Cad = miRsAux.Fields(0)
        I = DiasMes(CByte(Mid(Cad, 5, 2)), CInt(Mid(Cad, 1, 4)))
        Cad = Format(I, "00") & "/" & Mid(Cad, 5, 2) & "/" & Mid(Cad, 1, 4)
        UltFecPresentada = CDate(Cad)
        UltFecPresentada = DateAdd("d", 1, UltFecPresentada)
        If UltFecPresentada < vParamAplic.FechaActiva Then UltFecPresentada = vParamAplic.FechaActiva
    End If
    miRsAux.Close
    
    Cad = "Select * from tmpnlotes where codusu =" & vUsu.Codigo
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Si el periodo que esta cerrando es un mes mas que el que esta activo
    If UltFecPresentada = miRsAux!FechaAlb Then
        Label1(18).visible = True
        Text2.visible = True
        Me.cmdCerrar.visible = True
    End If
    'no puede ser eof
    LabelFechaMes.Caption = MonthName(Month(miRsAux!FechaAlb)) & "  " & Year(miRsAux!FechaAlb)
    LabelFechaMes.Tag = miRsAux!FechaAlb
    
    Text1(1).Text = Format(miRsAux!Cantidad, "##,##0")
    If IsNull(miRsAux!numlotes) Then
        Importe = 0
    Else
        Importe = CCur(TransformaPuntosComas(miRsAux!numlotes))
    End If
    Text1(2).Text = Format(Importe, "##,##0")
    If Not IsNull(miRsAux!Cantidad) Then Importe = miRsAux!Cantidad - Importe
    miRsAux.Close
    
    'Dato anterior
    Cad = DateAdd("d", -1, CDate(LabelFechaMes.Tag))
    Cad = "Select aceite_stfinal from valldeclara where mes=" & Month(CDate(Cad)) & " AND anyo=" & Year(CDate(Cad))
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Text1(0).Text = "0"
    If Not miRsAux.EOF Then
        Text1(0).Text = Format(miRsAux!aceite_stfinal, "##,##0")
        Importe = Importe + miRsAux!aceite_stfinal
    End If
    miRsAux.Close
    'Total la suma de todos
    Text1(3).Text = Format(Importe, "##,##0")
    
    
    
    'Totales anteriores periodo declarado
    If Month(CDate(LabelFechaMes.Tag)) < 10 Then
        Cad = Year(CDate(LabelFechaMes.Tag)) - 1
    Else
        Cad = Year(CDate(LabelFechaMes.Tag))
    End If
    
    'Desde inicio de "ejercicio" , fecha activa, hasta fecha
    ' y fecha menor que mesaño actual
    
    Cad = "( anyo =" & Year(vParamAplic.FechaActiva) & " AND mes>=" & Month(vParamAplic.FechaActiva)
    Cad = Cad & ") AND (anyo =" & Year(CDate(LabelFechaMes.Tag)) & " AND mes<" & Month(CDate(LabelFechaMes.Tag)) & ")"
    Cad = "  from valldeclara WHERE " & Cad
    Cad = " Sum(salida_aceite), Sum(salida_orujo) " & Cad
    Cad = " Sum(aceituna_entrada), Sum(aceituna_molturada), Sum(aceite_obtenido), Sum(orujo_ontenido)," & Cad
    Cad = "select sum(detasalida_envprop),sum(detasalida_otrasent),sum(detasalida_otraptri), " & Cad
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'NO debe ser eof
    
    Importe2 = 0
    For I = 0 To 2
        Importe = DBLet(miRsAux.Fields(I), "N")
        Importe2 = Importe2 + Importe
        Text1(I + 4).Text = Format(Importe, "##,##0")
    Next
    Text1(7).Text = Format(Importe2, "##,##0")
        
        
    
    
    'sumatorios del grid
    Importe = 0
    For I = 22 To 27
        Text1(I).Tag = 0
    Next
    
    If Not miRsAux.EOF Then
        For I = 4 To 6
            If Not IsNull(miRsAux.Fields(I - 4)) Then
               Importe = Importe + miRsAux.Fields(I - 4)
               Text1(I).Text = Format(miRsAux.Fields(I - 4), "##,##0")
            End If
        Next
        Text1(7).Text = Format(Importe, "##,##0")
        'Sumatorios anteriores del grid
        For I = 22 To 27
            Text1(I).Tag = DBLet(miRsAux.Fields(I - 19), "N")
        Next
    
    End If
    miRsAux.Close

    '1 .- Coupage Entrada
    '   3 .- Trasiego entrada
    '   4 .-    "     salida
    '   7 .- Forzar vaciado
   '   9 .-   "    salida
    'select * from proddepositoshco where horamovi between '2017-01-01' and '2017-01-30'  and numdeposito=18
    'and tipoaccion in (1,3,7,9)
    
    
    'Los sums de los productos del grid
    Cad = "select sum(importe1),sum(importe2),sum(importe3),sum(importe4),sum(importe5),sum(importeb1)"
    Cad = Cad & " from tmpinformes where codusu=" & vUsu.Codigo
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'NO puede ser eof
    For I = 16 To 21
        Text1(I).Text = Format(miRsAux.Fields(I - 16), "##,##0")
        Text1(I + 6).Text = Format(miRsAux.Fields(I - 16) + Text1(I + 6).Tag, "##,##0")
    Next I
    miRsAux.Close
    
    
        
    
    
    
    

    'por utlimo cargamos el grid
    CargaGrid
    
End Sub



Private Sub CargaGrid()
Dim I As Integer
   DataGrid1.Enabled = False

    Cad = " SELECT codigo1,importe1,importe2,importe3,importe4,importe5,importeb1  FROM tmpinformes where codusu = " & vUsu.Codigo
    Cad = Cad & " ORDER BY codigo1"
    
    CargaGridGnral DataGrid1, Me.Adodc1, Cad, False
    'Numero Linea
    DataGrid1.Columns(0).Caption = "Dia"
    DataGrid1.Columns(0).Width = 600
    
    
    For I = 1 To 6
        DataGrid1.Columns(I).Caption = RecuperaValor("Entrada|Molturada|Aceite|Orujo|Aceite|Orujo|", I)
        DataGrid1.Columns(I).Alignment = dbgRight
        DataGrid1.Columns(I).Width = 1450
        DataGrid1.Columns(I).NumberFormat = "##,##0"
    Next
    
    For I = 1 To 6
        Me.Text1(I + 15).Left = DataGrid1.Columns(I).Left + DataGrid1.Left
        Me.Text1(I + 15).Width = DataGrid1.Columns(I).Width
           
        Me.Text1(I + 21).Left = DataGrid1.Columns(I).Left + DataGrid1.Left
        Me.Text1(I + 21).Width = DataGrid1.Columns(I).Width
           
           
    Next I
   
End Sub





Private Function CerrarProcesoMensual() As Boolean
    On Error GoTo eCerrarProcesoMensual
    CerrarProcesoMensual = False


    'insertamos en la cabecera
    Cad = "INSERT INTO valldeclara(mes,anyo,fechahora,aceite_existencia,aceite_producido,aceite_salidas,aceite_stfinal,"
    Cad = Cad & "detasalida_envprop,detasalida_otrasent,detasalida_otraptri,aceituna_entrada,aceituna_molturada,"
    Cad = Cad & "aceite_obtenido,orujo_ontenido,salida_aceite,salida_orujo,observa) VALUES ("
    Cad = Cad & Month(LabelFechaMes.Tag) & "," & Year(LabelFechaMes.Tag) & "," & DBSet(Now, "FH") & ","
    'aceite_existencia,aceite_producido,aceite_salidas,aceite_stfinal
    Cad = Cad & DBSet(Text1(0).Text, "N", "N") & "," & DBSet(Text1(1).Text, "N", "N") & ","
    Cad = Cad & DBSet(Text1(2).Text, "N", "N") & "," & DBSet(Text1(3).Text, "N", "N") & ","
    'detasalida_envprop,detasalida_otrasent,detasalida_otraptri
    Cad = Cad & DBSet(Text1(8).Text, "N", "N") & "," & DBSet(Text1(9).Text, "N", "N") & "," & DBSet(Text1(10).Text, "N", "N") & ","
    'aceituna_entrada , aceituna_molturada, aceite_obtenido,orujo_ontenido,
    Cad = Cad & DBSet(Text1(16).Text, "N", "N") & "," & DBSet(Text1(17).Text, "N", "N") & ","
    Cad = Cad & DBSet(Text1(18).Text, "N", "N") & "," & DBSet(Text1(19).Text, "N", "N") & ","
    'salida_aceite,salida_orujo,observa
    Cad = Cad & DBSet(Text1(20).Text, "N", "N") & "," & DBSet(Text1(21).Text, "N", "N") & "," & DBSet(Text2.Text, "T") & ")"
    conn.Execute Cad
    
    'Insertamos lineas
    Cad = "INSERT INTO valldeclaralin(anyo,mes,dia,aceitu_ent,aceitu_molt,obtenido_aceite,obtenido_orujo,salida_aceite,salida_orujo) "
    Cad = Cad & " select year(fecha1) anyo,campo1 mes, codigo1 dia,importe1 aceitu_ent"
    Cad = Cad & " ,importe2 aceitu_molt,importe3 obtenido_aceite ,importe4 obtenido_orujo"
    Cad = Cad & " ,importe5 salida_aceite,importeb1 salida_orujo"
    Cad = Cad & " from tmpinformes where codusu =" & vUsu.Codigo & " order by codigo1"
    conn.Execute Cad
    
    CerrarProcesoMensual = True
    
eCerrarProcesoMensual:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description

End Function
