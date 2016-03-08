VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Batchelor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pruebas Impresion"
   ClientHeight    =   8310
   ClientLeft      =   2685
   ClientTop       =   3255
   ClientWidth     =   12585
   Icon            =   "frmBarTender.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   12585
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDavid 
      Caption         =   "Impr. David"
      Height          =   375
      Left            =   4680
      TabIndex        =   29
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox txtCantidad 
      Height          =   285
      Left            =   1200
      TabIndex        =   28
      Text            =   "1"
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Frame Frame5 
      Caption         =   "Etiquetas aceites morales"
      Height          =   3495
      Left            =   6360
      TabIndex        =   15
      Top             =   240
      Width           =   6135
      Begin VB.Frame FrameCaja 
         Caption         =   "CAJAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2295
         Left            =   240
         TabIndex        =   19
         Top             =   840
         Width           =   5775
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   360
            TabIndex        =   24
            Top             =   1800
            Width           =   5055
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   360
            TabIndex        =   22
            Top             =   1200
            Width           =   5055
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   360
            TabIndex        =   20
            Top             =   600
            Width           =   5055
         End
         Begin VB.Label Label5 
            Caption         =   "DUN"
            Height          =   255
            Left            =   360
            TabIndex        =   25
            Top             =   1560
            Width           =   2175
         End
         Begin VB.Label Label4 
            Caption         =   "IDCaja"
            Height          =   255
            Left            =   360
            TabIndex        =   23
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Pruducto"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   21
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Materia auxiliar"
         Height          =   255
         Index           =   2
         Left            =   4080
         TabIndex        =   18
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Palets"
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   17
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Cajas"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.Frame FramePalet 
         Caption         =   "PALETS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   2295
         Left            =   240
         TabIndex        =   26
         Top             =   840
         Visible         =   0   'False
         Width           =   5775
      End
      Begin VB.Frame FrameMatAux 
         Caption         =   "Materia auxiliar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   2295
         Left            =   240
         TabIndex        =   27
         Top             =   840
         Visible         =   0   'False
         Width           =   5775
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   8835
      Top             =   -30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBarTender.frx":0A02
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBarTender.frx":0A60
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBarTender.frx":0ABE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBarTender.frx":0B1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBarTender.frx":0B7A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog comDlgFormats 
      Left            =   8160
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "Printer"
      Height          =   795
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   6195
      Begin VB.CheckBox chkUsePrinterSpecifiedinFormat 
         Caption         =   "Usar la impr. del archivo"
         Height          =   225
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Value           =   1  'Checked
         Width           =   2250
      End
      Begin VB.ComboBox cboWhatPrinter 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2400
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   360
         Width           =   3675
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "BarTender Formats"
      Height          =   1875
      Left            =   75
      TabIndex        =   0
      Top             =   240
      Width           =   6240
      Begin VB.ListBox list_Formats 
         Height          =   645
         Left            =   195
         TabIndex        =   12
         Top             =   285
         Width           =   5865
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   390
         Left            =   4920
         TabIndex        =   9
         Top             =   1200
         Width           =   1125
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   405
         Left            =   3600
         TabIndex        =   1
         Top             =   1200
         Width           =   1125
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Copies"
      Height          =   1575
      Left            =   960
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   5970
      Begin VB.CheckBox chkUseSpecifiedAmountinFormat 
         Caption         =   "Use Specified Amount in Format"
         Height          =   435
         Left            =   225
         TabIndex        =   14
         Top             =   330
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.TextBox txtNumberOfBatchCopies 
         Height          =   315
         Left            =   4440
         MaxLength       =   6
         TabIndex        =   5
         Text            =   "1"
         Top             =   1110
         Width           =   810
      End
      Begin VB.TextBox txtNumberOfIdenticalCopies 
         Height          =   315
         Left            =   4440
         MaxLength       =   6
         TabIndex        =   4
         Text            =   "1"
         Top             =   270
         Width           =   810
      End
      Begin VB.TextBox txtNumberOfSerializedCopies 
         Height          =   315
         Left            =   4440
         MaxLength       =   6
         TabIndex        =   3
         Text            =   "1"
         Top             =   720
         Width           =   810
      End
      Begin VB.Label Label1 
         Caption         =   "Number of Batch Copies"
         Height          =   255
         Left            =   2385
         TabIndex        =   8
         Top             =   1155
         Width           =   1740
      End
      Begin VB.Label lblnumberofidentical 
         Caption         =   "Number of Identical Copies"
         Height          =   240
         Left            =   2385
         TabIndex        =   7
         Top             =   330
         Width           =   1965
      End
      Begin VB.Label Label2 
         Caption         =   "Number of Serialized Copies"
         Height          =   210
         Left            =   2385
         TabIndex        =   6
         Top             =   765
         Width           =   1995
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Cantidad"
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   30
      Top             =   3360
      Width           =   645
   End
   Begin VB.Image imgBTFormat 
      Height          =   4260
      Left            =   240
      Top             =   3840
      Width           =   12015
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000003&
      BackStyle       =   1  'Opaque
      Height          =   4575
      Left            =   120
      Top             =   3720
      Width           =   12315
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      WindowList      =   -1  'True
      Begin VB.Menu mnuVisible 
         Caption         =   "&Visible"
      End
      Begin VB.Menu mnuExportOpen 
         Caption         =   "&Export / Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPrintBatch 
         Caption         =   "&Print"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuSendText 
         Caption         =   "&Send Text"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Batchelor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkUsePrinterSpecifiedinFormat_Click()
'If the "Use Printer Specified in Format" check box is checked then uncheck it
'and let the user pick which printer they want to print to.
If chkUsePrinterSpecifiedinFormat.Value = 1 Then
    cboWhatPrinter.Enabled = False
Else
    cboWhatPrinter.Enabled = True
End If
 
End Sub
 
Private Sub chkUseSpecifiedAmountinFormat_Click()
 
If chkUseSpecifiedAmountinFormat.Value = 1 Then 'if this check box is unchecked then check it and...
    txtNumberOfIdenticalCopies.Enabled = False 'disable the Number of Identical copies text box
    txtNumberOfSerializedCopies.Enabled = False 'disable the Number of serialized copies text box
    txtNumberOfSerializedCopies.BackColor = -2147483644 'Change the background of the text boxes
    txtNumberOfIdenticalCopies.BackColor = -2147483644 'to a grey color
Else
    txtNumberOfIdenticalCopies.Enabled = True 'Enable the Number of identical copies for manipulation.
    txtNumberOfSerializedCopies.Enabled = True 'Enable the Number of Serialized Copies for manipulation.
    txtNumberOfSerializedCopies.BackColor = -2147483643 'set the back ground of this text boxes to white.
    txtNumberOfIdenticalCopies.BackColor = -2147483643 'set the background of this text boxes to white.
    txtNumberOfIdenticalCopies.SetFocus 'Set the cursor focus to be in the identical copies of label text box.
End If
 
End Sub
 
Private Sub cmdAdd_Click()
 
On Error GoTo ErrorHandler
    comDlgFormats.FileName = ""
    'Set filters
    comDlgFormats.Filter = "BarTender Label Formats (*.btw)|*.btw|"
    'Specify default filter
    comDlgFormats.FilterIndex = 2
    'Display the Open dialog box
    comDlgFormats.ShowOpen
     
    'sends the format file name and path to a public Variable called m_sNewFormat
    m_sNewFormat = comDlgFormats.FileName
     
    'Add the file selected as long as the name is blank...
    If m_sNewFormat <> "" Then list_Formats.AddItem m_sNewFormat
     
    ' if there's just 1 entry in the formats then high-light that format
    If list_Formats.ListCount = 1 Then
        list_Formats.ListIndex = 0
        mnuExportOpen_Click
    End If
     
    ' if there's 1 entry or more in the formats list
    ' Enable the print options other wise leave them disabled.
    If list_Formats.ListCount > 0 Then
        mnuPrintBatch.Enabled = True
        'Toolbar1.Buttons.Item(5).Enabled = True '5 = Print Batch
    Else
        mnuPrintBatch.Enabled = False
        'Toolbar1.Buttons.Item(5).Enabled = False '5 = Print Batch
    End If
     
     
         
ErrorHandler:
    Exit Sub
End Sub
 
 
Public Sub p_OpenFormat()
 
On Error GoTo ErrorHandler
 
'if the "Use Printer Specified in Format" check box is checked then'
'open the format with no specified printer if or else open the label with
'the printer that's specified from the printer drop down list.
 
If chkUsePrinterSpecifiedinFormat.Value = 1 Then
    'open the format with the default printer while closing any labels which may be up
    Set BtFormat = BtApp.Formats.Open(m_sSelectedFormat, True)
Else
    'Open the format with the selected printer in the printer drop down list.
    Set BtFormat = BtApp.Formats.Open(m_sSelectedFormat, True, cboWhatPrinter.Text)
End If
 
' if the code gets to here then the format has opened successfully
m_bOpenFormatFailed = False
 
 
Exit Sub
 
ErrorHandler:
 
m_bOpenFormatFailed = True 'If an Error has occured then the format opened has failed setting the flag...
p_ErrorHandler 'Go to the Error Handler procedure which is in the Flags Module
 
     
End Sub
 
 
Private Sub cmdDavid_Click()
 Dim I As Integer
Dim J As Integer

    If BtFormat Is Nothing Then Exit Sub
    
    If Val(Me.txtCantidad.Text) = 0 Then Exit Sub
    
    If Me.Option1(0).Value Then
        'Etiquetas de cajas
    
        BtFormat.SetNamedSubStringValue "Producto", "er producto esta muuu jodido"
        BtFormat.SetNamedSubStringValue "CodigoDUN", "8100002512"
 
    ElseIf Me.Option1(1).Value Then
        'Etiquetas de palet
    
        BtFormat.SetNamedSubStringValue "Producto", "er producto esta muuu jodido"
        BtFormat.SetNamedSubStringValue "CodigoDUN", "810000251"
 
    Else
        'etiqueta de materia auxiliar
    
    
    
    End If
    BtFormat.NumberSerializedLabels = 50
    
    
    
    
    
    
    
    
    
    For I = 0 To Val(Me.txtCantidad.Text) - 1
        J = 50 * I
        J = J + 1
        If Me.Option1(0).Value Then
            'ID CAJA
            BtFormat.SetNamedSubStringValue "IdCaja", "0000600" & Format(J, "00000")
            
            
            
            
            
        ElseIf Me.Option1(1).Value Then
            'PALET
            
        Else
            'Mat aux
        End If
        BtFormat.PrintOut False, False
    Next I

        
End Sub

Private Sub mnuAbout_Click()
    MsgBox "Programa pruebas impresion etiquetas" & vbCrLf & vbCrLf & "Ariadna Software", vbInformation
 
End Sub
 
Private Sub mnuExportOpen_Click()
 
On Error GoTo ErrorHandler
 
'As long as there's a format name in the list then go on to Open and then Export this label format.
If list_Formats.ListIndex <> -1 Then
     
    m_sSelectedFormat = list_Formats.List(list_Formats.ListIndex)
 
 
    'Go to the p_OpenFormat procedure to Open the Selected label formats
    p_OpenFormat
 
    'Set imgBTFormat to become visible on the form.
    imgBTFormat.Visible = True
     
    'If the label format opens then Export the image to the form
    If m_bOpenFormatFailed = False Then
        BtFormat.ExportToClipboard btColors16, btResolutionScreen 'export image to clipboard
        imgBTFormat.Picture = Clipboard.GetData(vbCFDIB) 'Get the data from the clip board to display this.
    End If
     

     
    'Enable the Send Text, Print and Save menu options.
    mnuSendText.Enabled = True
    mnuSave.Enabled = True
    mnuPrintBatch.Enabled = True
     
Else
    imgBTFormat.Visible = False 'if there are no formats selected the throw up a message box
    MsgBox "You must first high-light a Format to be previewed from the BarTender Formats list.", vbInformation, "Batchelor"
     
End If
 
Exit Sub
ErrorHandler:
    p_ErrorHandler
     
 
End Sub
 
 
 
Private Sub cmdPrint_Click()
 
On Error GoTo ErrorHandler
 
Dim GetListCount As Integer ' define an integer to count the formats in BT.
Dim X As Long 'Define an Integer to use to count batch copies
Dim Y As Long 'Define an Integer to use to count how many labels it's printed
 
GetListCount = Batchelor.list_Formats.ListCount
 
m_iIgnoreAllErrorsFlag = 0 'Resets the ignore all flags Variable to off.
m_bOpenFormatFailed = False 'If it gets to this point then it succeeded in opening the format.
 
If GetListCount > 0 Then    'if there's anything in the Formats List then...
    For Y = 1 To txtNumberOfBatchCopies.Text 'keeps track of how many batches you've printed.
        For X = 0 To (GetListCount - 1) 'For Each format in the list print them out.
            m_sSelectedFormat = list_Formats.List(X)   'get the name of the format from the list
            p_OpenFormat    'Go to the OpenFormat procedure and open the labels.
             
            If m_bOpenFormatFailed = False Then 'If the format opens then procede.
                If chkUseSpecifiedAmountinFormat.Value = 0 Then
                    BtFormat.IdenticalCopiesOfLabel = txtNumberOfIdenticalCopies.Text
                    BtFormat.NumberSerializedLabels = txtNumberOfSerializedCopies.Text
                End If
                BtFormat.PrintOut 'Print the Label Format out
            End If
             
            m_bOpenFormatFailed = False
        Next
    Next
End If
 
Exit Sub
 
ErrorHandler:
    'Goto the Error
    p_ErrorHandler
    Resume Next
     
         
End Sub
 
Private Sub cmdRemove_Click()
 
'Remove the highlighted item in the "BarTender Formats" list
'as long as there's something in the list.
If list_Formats.ListIndex <> -1 Then
    list_Formats.RemoveItem (list_Formats.ListIndex)
End If
 
If list_Formats.ListCount = 0 Then
         
 '   Toolbar1.Buttons.Item(3).Enabled = False '3 = Send Text
 '   Toolbar1.Buttons.Item(4).Enabled = False '4 = Save Format
 '   Toolbar1.Buttons.Item(5).Enabled = False '5 = Print Batch
    imgBTFormat.Visible = False
     
    'Disable the Send Text, Print and Save menu options.
    mnuSendText.Enabled = False
    mnuSave.Enabled = False
    mnuPrintBatch.Enabled = False
End If
     
 
End Sub
 
Private Sub cmdSubStringsProcedure_Click()
 
On Error GoTo ErrorHandler
 
'Checks to make sure that there's a format path and name in the Format's list
If list_Formats.ListIndex <> -1 Then
        If BtApp.Formats.Count > 0 Then 'check to make sure that a label format is acutally open in BarTender.
            If BtFormat.NamedSubStrings.Count > 0 Then 'Checks to make sure that the format has Sub-strings
                dbxSendTextToLabel.Show vbModal, Me ' if all conditions pass then show the dialog box.
            Else
                'if the selected format doesn't have any sub strings then show this message...
                MsgBox "The format selected contains no Named Sub-Strings.", vbExclamation, "Batchelor"
            End If
             
        Else
            'If there are no label formats open within Bartender then display this message...
            MsgBox "Please select a valid label format from the Label Formats List", vbExclamation, "Batchelor"
        End If
Else
    imgBTFormat.Visible = False 'if there are no formats selected the throw up a message box
    MsgBox "You must first high-light a Format to be previewed from the BarTender Formats list.", vbInformation, "Batchelor"
     
End If
 
Exit Sub
 
ErrorHandler:
 
    p_ErrorHandler
    Exit Sub
     
 
 
End Sub
 
 
Private Sub cmdVisible_Click()
'Make the Bartender visible if it is Not Visible OR Else
'Make the BarTender Invisible if it is Visible
 
If m_bBtVisibility = False Then
    BtApp.Visible = True
    m_bBtVisibility = True
Else
    BtApp.Visible = False
    m_bBtVisibility = False
End If
 
 
End Sub
 
Private Sub Form_Load()
'On loading the form if we get an error
'Goto the Error Handler at the bottom of this procedure.
On Error GoTo ErrorHandler
 
'Here is where we're Starting the Bartender Application
'At the time of loading this program. We've delclared
'the variable in the "PublicVariables Module".
'****************************************************
Set BtApp = CreateObject("Bartender.Application")  '*
'****************************************************
 
'Make BarTender not seen
BtApp.Visible = False
 
'Flag that Bartender is invisible
m_bBtVisibility = False
 
'Declaring a variable for the printers
Dim X As Printer
 
'Get the names of printers installed on the system
For Each X In Printers
    cboWhatPrinter.AddItem X.DeviceName
Next
 
'sets the combo box to the current printer.
cboWhatPrinter.Text = Printer.DeviceName
 
'Centers the Form on the screen
'Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
 
'Sets the "Suppress Page Setup Error" to not skip the error message box.
m_iIgnoreAllErrorsFlag = 0
 
'Disable and Change the color of the Background of the identical and Serialize copies field.
txtNumberOfIdenticalCopies.Enabled = False
txtNumberOfSerializedCopies.Enabled = False
txtNumberOfSerializedCopies.BackColor = -2147483644
txtNumberOfIdenticalCopies.BackColor = -2147483644
 
 
Exit Sub
 
ErrorHandler:
 
'If the Bartender Pro/integrator or Enterprise software is not on
'the system then give this error message.
    If Err.Number = 429 Then
        MsgBox "You have not installed the BarTender Professional Integrator Package " & _
        vbCrLf & "or the BarTender Enterprise Edition software.  Please install either one" & vbCrLf & _
        "of these packages before starting the Batchelor.", vbExclamation, "Batchelor"
        Unload Batchelor
    End If
 
'if no printer drivers are installed on this machine then come up with the
'error message
    If Err.Number = 484 Then
        MsgBox "You have not installed a Printer on this system, therefore the Batchelor" & vbCrLf & _
        "cannot print.  Please install a printer driver and restart the Batchelor.", vbExclamation, "Batchelor"
        Unload Batchelor
    Else
        MsgBox "Error Number:    " & Err.Number & vbCrLf & vbCrLf & vbCrLf & Err.Description, vbExclamation, "Batchelor"
    End If
     
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
 
'Closing the BarTender Application and the format if there is one.

BtFormat.Close btDoNotSaveChanges
BtApp.Quit
 
End Sub
 
Private Sub imgBTFormat_DblClick()
'Item(3) is = SendSub-string button.
'If Toolbar1.Buttons.Item(3).Enabled = True Then
    mnuSendText_Click
'End If
End Sub
 
Private Sub list_Formats_DblClick()
'Runs the ExportOpen menu option which Exports then opens the format.
mnuExportOpen_Click
End Sub
 
Private Sub mnuExit_Click()
'Unloads the form and the label formats.
Unload Me
End Sub
 
 
Private Sub mnuPrintBatch_Click()
 
On Error GoTo ErrorHandler
 
'
Dim GetListCount As Integer
Dim X As Long
Dim Y As Long
 
GetListCount = Batchelor.list_Formats.ListCount
 
m_iIgnoreAllErrorsFlag = 0
m_bOpenFormatFailed = False
 
If GetListCount > 0 Then    'if there's anything in the Formats List then...
    For Y = 1 To txtNumberOfBatchCopies.Text
        For X = 0 To (GetListCount - 1)
            m_sSelectedFormat = list_Formats.List(X)   'get the name of the format from the list
            p_OpenFormat    'Go to the OpenFormat procedure and open the labels
             
            If m_bOpenFormatFailed = False Then
                If chkUseSpecifiedAmountinFormat.Value = 0 Then
                    BtFormat.IdenticalCopiesOfLabel = txtNumberOfIdenticalCopies.Text
                    BtFormat.NumberSerializedLabels = txtNumberOfSerializedCopies.Text
                End If
                BtFormat.PrintOut 'Print the Label Format out
            End If
             
            m_bOpenFormatFailed = False
        Next
    Next
End If
 
Exit Sub
 
ErrorHandler:
         
    p_ErrorHandler
    Resume Next
 
End Sub
 
Private Sub mnuSave_Click()
    BtFormat.Save
End Sub
 
Private Sub mnuSendText_Click()
 
On Error GoTo ErrorHandler
 
'Checks to make sure that there's a format path and name in the Format's list
If list_Formats.ListIndex <> -1 Then
        If BtApp.Formats.Count > 0 Then 'check to make sure that a label format is acutally open in BarTender.
            If BtFormat.NamedSubStrings.Count > 0 Then 'Checks to make sure that the format has Sub-strings
                dbxSendTextToLabel.Show vbModal, Me ' if all conditions pass then show the dialog box.
            Else
                'if the selected format doesn't have any sub strings then show this message...
                MsgBox "The format selected contains no Named Sub-Strings.", vbExclamation, "Batchelor"
            End If
             
        Else
            'If there are no label formats open within Bartender then display this message...
            MsgBox "Please select a valid label format from the Label Formats List", vbExclamation, "Batchelor"
        End If
Else
    imgBTFormat.Visible = False 'if there are no formats selected the throw up a message box
    MsgBox "You must first high-light a Format to be previewed from the BarTender Formats list.", vbInformation, "Batchelor"
     
End If
 
Exit Sub
 
ErrorHandler:
 
    p_ErrorHandler
    Exit Sub
     
 
 
End Sub
 
Private Sub mnuVisible_Click()
 
'Make the Bartender visible if it is Not Visible OR Else
'Make the BarTender Invisible if it is Visible
If m_bBtVisibility = False Then
    BtApp.Visible = True
    m_bBtVisibility = True
Else
    BtApp.Visible = False
    m_bBtVisibility = False
End If
 
 
End Sub
 
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
 
Select Case Button.Key
Case "Visibility": mnuVisible_Click ' use the function stored in the mnuVisible Procedure
Case "Export": mnuExportOpen_Click ' use the function stored in the mnuExport Procedure
Case "SendText": mnuSendText_Click ' use the function stored in the mnuSendText Procedure
Case "Print": mnuPrintBatch_Click ' use the function stored in the mnuPrintBatch Procedure
Case "Save": mnuSave_Click ' use the function stored in the mnuSave Procedure
 
End Select
 
End Sub
 
Private Sub Option1_Click(Index As Integer)
    Me.FrameCaja.Visible = Index = 0
    Me.FramePalet.Visible = Index = 1
    Me.FrameMatAux.Visible = Index = 2
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    TeclaNUmerica KeyAscii
End Sub

Private Sub txtNumberOfBatchCopies_KeyPress(KeyAscii As Integer)
    TeclaNUmerica KeyAscii
End Sub

Private Sub TeclaNUmerica(ByRef KeyAscii As Integer)
'Allow only numerics to be typed in the Number of Batch Copies text box
If (KeyAscii <> 48) And _
(KeyAscii <> 49) And _
(KeyAscii <> 50) And _
(KeyAscii <> 51) And _
(KeyAscii <> 52) And _
(KeyAscii <> 53) And _
(KeyAscii <> 54) And _
(KeyAscii <> 55) And _
(KeyAscii <> 56) And _
(KeyAscii <> 57) And _
(KeyAscii <> 8) Then
    KeyAscii = 0
End If
 
End Sub
 
Private Sub txtNumberOfIdenticalCopies_KeyPress(KeyAscii As Integer)
 
'Allow only numerics to be typed in the Number of Identical Copies text box.
If (KeyAscii <> 48) And _
(KeyAscii <> 49) And _
(KeyAscii <> 50) And _
(KeyAscii <> 51) And _
(KeyAscii <> 52) And _
(KeyAscii <> 53) And _
(KeyAscii <> 54) And _
(KeyAscii <> 55) And _
(KeyAscii <> 56) And _
(KeyAscii <> 57) And _
(KeyAscii <> 8) Then
    KeyAscii = 0
End If
 
End Sub
 
Private Sub txtNumberOfSerializedCopies_KeyPress(KeyAscii As Integer)
 
'Allow only numerics to be typed in the Number of Serialized Copies text box.
If (KeyAscii <> 48) And _
(KeyAscii <> 49) And _
(KeyAscii <> 50) And _
(KeyAscii <> 51) And _
(KeyAscii <> 52) And _
(KeyAscii <> 53) And _
(KeyAscii <> 54) And _
(KeyAscii <> 55) And _
(KeyAscii <> 56) And _
(KeyAscii <> 57) And _
(KeyAscii <> 8) Then
    KeyAscii = 0
End If
 
End Sub

