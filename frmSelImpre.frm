VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelImpre 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Index           =   1
      Left            =   4440
      TabIndex        =   2
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Index           =   0
      Left            =   3360
      TabIndex        =   1
      Top             =   3840
      Width           =   975
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   5741
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Impresora"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmSelImpre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim PrimVez As Boolean
Dim ImpOriginal As Integer
Dim P As Printer

Private Sub Command1_Click(Index As Integer)


    If Index = 0 Then
        If Me.ListView1.ListItems.Count = 0 Then Exit Sub
        If Me.ListView1.SelectedItem Is Nothing Then Exit Sub
        
        If ImpOriginal <> Me.ListView1.SelectedItem.Index Then
        
             For Each P In Printers
               If P.DeviceName = ListView1.SelectedItem.Text Then
                ' La define como predeterminada del sistema.
                   Set Printer = P
                   'Espera 0.5
                   CambiarImprDefecto
                   ' Sale del bucle.
                   Exit For
                End If
            Next
        End If
        
        
    End If
    
    Unload Me
    Set P = Nothing
End Sub

Private Sub Form_Activate()
    If PrimVez Then
        PrimVez = False
        
        CargaImpresoras
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Me.Icon = frmppal.Icon
    PrimVez = True
    Me.ListView1.SmallIcons = frmppal.imgListComun
    
End Sub



Private Sub CargaImpresoras()

Dim IT As ListItem

    On Error GoTo eEstablecerImpresoraAnterior
        Me.ListView1.ColumnHeaders(1).Width = Me.ListView1.Width - 400
        
        For Each P In Printers
            Set IT = Me.ListView1.ListItems.Add()
            IT.Text = P.DeviceName
            IT.SmallIcon = 16
            If P.DeviceName = Printer.DeviceName Then
                IT.Bold = True
                Me.ListView1.SelectedItem = IT
                ImpOriginal = IT.Index
            End If
            
        Next


    
Exit Sub
eEstablecerImpresoraAnterior:
    Err.Clear
    

End Sub


Private Sub CambiarImprDefecto()
Const HWND_BROADCAST = &HFFFF&
Const WM_WININICHANGE = &H1A
Dim di As Variant
Dim L As Variant

di = WriteProfileString("WINDOWS", "DEVICE", Me.ListView1.SelectedItem.Text & " ,winspool, Ne05:")
L = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, ByVal "windows")

End Sub
