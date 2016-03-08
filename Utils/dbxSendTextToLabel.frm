VERSION 5.00
Begin VB.Form dbxSendTextToLabel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Send Text To Label"
   ClientHeight    =   2610
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Enter Info for Sub-Strings"
      Height          =   1815
      Left            =   135
      TabIndex        =   2
      Top             =   210
      Width           =   4860
      Begin VB.TextBox txtDataToBeSent 
         Height          =   795
         Left            =   1800
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   810
         Width           =   2805
      End
      Begin VB.ComboBox cboSelectSubString 
         Height          =   315
         Left            =   1815
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   330
         Width           =   2805
      End
      Begin VB.Label lblSelectSubString 
         Caption         =   "Select Sub-String"
         Height          =   210
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1515
      End
      Begin VB.Label lblTypeDataToBeSent 
         Caption         =   "Data To Be Sent"
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   810
         Width           =   1470
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3780
      TabIndex        =   1
      Top             =   2145
      Width           =   1215
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   375
      Left            =   2445
      TabIndex        =   0
      Top             =   2145
      Width           =   1215
   End
End
Attribute VB_Name = "dbxSendTextToLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private Sub CancelButton_Click()
'Unload the form without transfering any variables.
Unload Me
 
End Sub
 
Private Sub cboSelectSubString_Click()
'Clears the "Data To Be Sent" dialog box.
txtDataToBeSent.Text = ""
'Gets the Data from the sub-string that's listed in the Sub-string list box.
txtDataToBeSent.Text = BtFormat.GetNamedSubStringValue(cboSelectSubString.Text)
 
End Sub
 
 
Private Sub cmdSend_Click()
 Dim Cad
'Send what the user has typed in the box over to the label.
BtFormat.SetNamedSubStringValue cboSelectSubString.Text, txtDataToBeSent.Text

 BtFormat.SetNamedSubStringValue "IdCaja", """"
 BtFormat.SetNamedSubStringValue "Producto", ""
 BtFormat.SetNamedSubStringValue "CodigoDUN", ""
 
  BtFormat.SetNamedSubStringValue "IdCaja", "la caja"
 BtFormat.SetNamedSubStringValue "Producto", "er producto"
 BtFormat.SetNamedSubStringValue "CodigoDUN", "dun"
 
 
 txtDataToBeSent.Text = BtFormat.GetNamedSubStringValue("IdCaja")
 txtDataToBeSent.Text = BtFormat.GetNamedSubStringValue("Producto")
 
 
 Exit Sub
'export image to clipboard
BtFormat.ExportToClipboard btColors16, btResolutionScreen
 
'Get the data from the clip board to display this.
Batchelor.imgBTFormat.Picture = Clipboard.GetData(vbCFDIB)
 
Unload Me
 
End Sub
 
Private Sub Form_Load()
 
'Centers the Form on the screen
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
 
'This Variable is be used as a format name in the below procedure.
Dim X As SubString
 
'Here We're making sure that a format exists in the list.
If BtFormat.NamedSubStrings.Count > 0 Then
 
'Here we're getting all the sub-strings from the
'selected format and putting them into a drop down list.
    For Each X In BtFormat.NamedSubStrings
        cboSelectSubString.AddItem X.Name
        cboSelectSubString.Text = cboSelectSubString.List(0)
    Next
     
    'gets the data from the sub-string that's
    'first listed in the sub-string drop down list.
    txtDataToBeSent.Text = BtFormat.GetNamedSubStringValue(cboSelectSubString.Text)
     
End If
 
 
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
 
'Clears out the "Select Sub-String" Drop down list.
cboSelectSubString.Clear
 
End Sub
