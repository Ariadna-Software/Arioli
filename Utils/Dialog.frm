VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Batchelor"
   ClientHeight    =   3870
   ClientLeft      =   3975
   ClientTop       =   4380
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   450
      Left            =   2610
      TabIndex        =   0
      Top             =   3210
      Width           =   1470
   End
   Begin VB.CheckBox chkIgnoreAllErrors 
      Caption         =   "Ignore all Automation Errors.  (If ignored all labels may not print.)"
      Height          =   405
      Left            =   780
      TabIndex        =   1
      Top             =   2670
      Width           =   4860
   End
   Begin VB.Label lblErrorDescription 
      Height          =   2280
      Left            =   840
      TabIndex        =   2
      Top             =   255
      Width           =   5535
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private Sub Form_Load()
'Centers the utility on the screen
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
 
End Sub
 
Private Sub OKButton_Click()
m_iIgnoreAllErrorsFlag = chkIgnoreAllErrors.Value
Unload Me
End Sub
