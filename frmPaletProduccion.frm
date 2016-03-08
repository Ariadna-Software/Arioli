VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPaletProduccion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Palet"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   6405
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameNuevo 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin VB.Frame FrameTapaLineas 
         Caption         =   "Lineas paletizado"
         Height          =   975
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Visible         =   0   'False
         Width           =   5895
         Begin VB.OptionButton optPalet 
            Caption         =   "Pal. 6"
            Height          =   255
            Index           =   5
            Left            =   4920
            TabIndex        =   26
            Top             =   480
            Width           =   855
         End
         Begin VB.OptionButton optPalet 
            Caption         =   "Pal. 5"
            Height          =   255
            Index           =   4
            Left            =   3960
            TabIndex        =   25
            Top             =   480
            Width           =   975
         End
         Begin VB.OptionButton optPalet 
            Caption         =   "Pal. 4"
            Height          =   255
            Index           =   3
            Left            =   3000
            TabIndex        =   24
            Top             =   480
            Width           =   975
         End
         Begin VB.OptionButton optPalet 
            Caption         =   "Pal. 3"
            Height          =   255
            Index           =   2
            Left            =   2040
            TabIndex        =   23
            Top             =   480
            Width           =   975
         End
         Begin VB.OptionButton optPalet 
            Caption         =   "Pal. 2"
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   22
            Top             =   480
            Width           =   975
         End
         Begin VB.OptionButton optPalet 
            Caption         =   "Pal. 1"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   21
            Top             =   480
            Width           =   975
         End
      End
      Begin VB.TextBox txtPalet 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1800
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.CheckBox chkReabrir 
         Caption         =   "Reabrir palet"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton cmdcancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   0
         Left            =   4920
         TabIndex        =   11
         Top             =   2280
         Width           =   975
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   10
         Top             =   2280
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "L7"
         Height          =   255
         Index           =   7
         Left            =   5520
         TabIndex        =   9
         Top             =   960
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "L6"
         Height          =   255
         Index           =   6
         Left            =   4764
         TabIndex        =   8
         Top             =   960
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   "L5"
         Height          =   255
         Index           =   5
         Left            =   4010
         TabIndex        =   7
         Top             =   960
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   "L4"
         Height          =   255
         Index           =   4
         Left            =   3256
         TabIndex        =   6
         Top             =   960
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   "L3"
         Height          =   255
         Index           =   3
         Left            =   2502
         TabIndex        =   5
         Top             =   960
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   "L2"
         Height          =   255
         Index           =   2
         Left            =   1748
         TabIndex        =   4
         Top             =   960
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   "L1"
         Height          =   255
         Index           =   1
         Left            =   994
         TabIndex        =   3
         Top             =   960
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   "L0"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   615
      End
      Begin VB.Image imgPalet 
         Height          =   240
         Left            =   1800
         Picture         =   "frmPaletProduccion.frx":0000
         Top             =   1800
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Nuevo palet en producción"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Frame FrameCierrePalet 
      Height          =   2895
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   6135
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   1575
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   2778
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
      End
      Begin VB.CheckBox ChkCierreProd 
         Caption         =   "Seguir con otro palet"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2400
         Width           =   2655
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3840
         TabIndex        =   15
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton cmdcancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   5040
         TabIndex        =   14
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Cierre palet en producción"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   345
         Left            =   960
         TabIndex        =   13
         Top             =   120
         Width           =   4095
      End
   End
End
Attribute VB_Name = "frmPaletProduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Varios As String
    'Si es nuevo pondra:
    '  LIENAPALETIZACION|lineas produccion disponibles|
Public TrazaEnLineas As String
Public cPal As CPalet


'Desde NUEVALINEA, nada mas acaba de poner la linea en produccion llamamos a este
'llevara la disponibilidad de lineas de paletizacion
' DesdeNuevaProd:  0011  las lineas 3 y 4 NO pueden paletizarse
Public DesdeNuevaProd As String

Private WithEvents frmPa As frmProdPalets
Attribute frmPa.VB_VarHelpID = -1


Dim C As Collection
Dim i As Integer
Dim Total As Integer

Private Sub chkReabrir_Click()
    Me.imgPalet.visible = chkReabrir.Value = 1
    txtPalet.visible = imgPalet.visible
    If Me.chkReabrir.Value = 0 Then Me.txtPalet.Text = ""
        
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    If DesdeNuevaProd <> "" Then
        If MsgBox("DEBERIA INICIAR PALETIZACION. Salir igualmente?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    CadenaDesdeOtroForm = ""
    Unload Me
End Sub

Private Sub cmdCerrar_Click()


    If TreeView1.Nodes.Count = 0 Then
        
        TrazaEnLineas = DevuelveDesdeBD(conAri, "count(*)", "prodcajas", "idpalet", CStr(cPal.ID))
        If TrazaEnLineas = "" Then TrazaEnLineas = "0"
        If Val(TrazaEnLineas) > 0 Then
            MsgBox "Hay cajas asignadas", vbExclamation
            Exit Sub
        End If
        If MsgBox("No hay cajas.  Eliminar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        Conn.Execute "DELETE FROM prodpalets WHERE idpalet = " & cPal.ID
        
        
        'prodpaletstraza
        EjecutaSQL conAri, "DELETE FROM prodpaletstraza WHERE idpalet = " & cPal.ID
        
        Set cPal = Nothing
        CadenaDesdeOtroForm = "OK"
        Unload Me
        Exit Sub
    End If
    
    
    'Abril 2012
    'NO DEJO CERRAR PALET DESDE AQUI, para eso esta la pistola
    MsgBox "No se puede cerrar el palet cuando tiene carga. Debe hacerse desde la pistola", vbExclamation
    Exit Sub
    
    'TrazaEnLineas: Esta variable esta "" y en el cierre NO la utilizo
    TrazaEnLineas = "¿Seguro que desea cerrar el palet "
    If Me.ChkCierreProd.Value = 1 Then TrazaEnLineas = TrazaEnLineas & " y continuar con otro"
    TrazaEnLineas = TrazaEnLineas & "?"
    If MsgBox(TrazaEnLineas, vbQuestion + vbYesNo) = vbYes Then
    
            
            
            
            If cPal.CerrarPalet(Total) Then
            
                If Me.ChkCierreProd.Value = 1 Then CopiarPalet
                
                'OK YA ESTA CERRADO. Ahora a imprimir la sabana santa
                Varios = "{tmppartidas.codusu}=" & vUsu.Codigo
                LlamaImprimirGral Varios, "", 0, "EtiqPalet.rpt", "Etiqueta palet: " & cPal.ID
                Varios = ""
            
                Set cPal = Nothing
                CadenaDesdeOtroForm = "OK"
                Unload Me
            End If
    End If
    TrazaEnLineas = ""
    
    
End Sub

Private Sub CopiarPalet()
Dim CPN As CPalet
    
        
    Set CPN = New CPalet
    For i = 0 To 7
        CPN.LineasProd(i) = cPal.LineasProd(i)
    Next i
    
    CPN.FechaInicio = Now
    CPN.LineaPeletizacion = cPal.LineaPeletizacion
    'Mal. Hay que ver QUE esta paletizanod AHORA, sin nada mas
    TrazaEnLineas = " Not FFin Is Null AND idPalet "
    TrazaEnLineas = DevuelveDesdeBD(conAri, "lotetraza", "prodpaletstraza", TrazaEnLineas, cPal.ID)
    'Si no encotramos NINGUNA ponermos la que habia
    If TrazaEnLineas = "" Then
        TrazaEnLineas = cPal.TrazabilidadPaletizando
    Else
        TrazaEnLineas = TrazaEnLineas & "|"
    End If
    While TrazaEnLineas <> ""
        i = InStr(1, TrazaEnLineas, "|")
        If i = 0 Then
            TrazaEnLineas = ""
        Else
            CPN.AñadirIdTraza CLng(Mid(TrazaEnLineas, 1, i - 1))
            TrazaEnLineas = Mid(TrazaEnLineas, i + 1)
        End If
    Wend
    
    CPN.CrearPalet
    Set CPN = Nothing
    
End Sub

Private Sub cmdNuevo_Click()
Dim b As Boolean
Dim TrazaAntes As Long
    For i = 0 To 7
        If Me.Check1(i).Value = 1 Then Exit For
    Next i

    If i > 7 Then
        MsgBox "Seleccione alguna linea de producción", vbExclamation
        Exit Sub
    End If

    
    'Si es del ALTA
    If Me.DesdeNuevaProd <> "" Then
        'La linea de produccion la asigno ahora
        Label3.Tag = -1
        For i = 0 To optPalet.Count - 1
            If Me.optPalet(i).Value Then
                Label3.Tag = i + 1
                Exit For
            End If
        Next i
    
        If Label3.Tag = -1 Then
            MsgBox "Error asignado linea paletizacion", vbExclamation
            Exit Sub
        End If
    End If
        


    If Me.chkReabrir.Value = 1 Then
        If Me.txtPalet.Text = "" Then
            MsgBox "Seleccione el palet que desea reabrir", vbExclamation
            Exit Sub
        End If
        If MsgBox("Volver a abrir palet " & txtPalet.Text & " ?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Else
        If MsgBox("Nuevo palet?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If








    Set cPal = New CPalet
    
    If Me.chkReabrir.Value = 1 Then
        
        'Id palet
        CadenaDesdeOtroForm = Trim(Mid(txtPalet.Text, 4)) 'quito ID:
        i = InStr(1, CadenaDesdeOtroForm, " ")
        If i = 0 Then
            CadenaDesdeOtroForm = -1 'error
        Else
            CadenaDesdeOtroForm = Trim(Mid(CadenaDesdeOtroForm, 1, i))
        End If
        
        b = cPal.Leer(CLng(CadenaDesdeOtroForm))
        
        cPal.LineaPeletizacion = Label3.Tag 'Nueva linea de paletizacion
        'Veremos lo que habia antes
        TrazaAntes = Len(cPal.TrazabilidadPaletizando)
            'lo que estaba paletizando lo tendre que quitar ya que tendre que saber cuales son los nuevos ID de palet
            
        
        AsignarTrazabilidad
        
        'VEo cuales son los nuevos
        CadenaDesdeOtroForm = Mid(cPal.TrazabilidadPaletizando, TrazaAntes + 1)
        
        If b Then b = cPal.ReAbrirPalet(CadenaDesdeOtroForm)
        
    Else
        'NUEVO
        cPal.FechaInicio = Now
        cPal.LineaPeletizacion = Label3.Tag
        
        
        AsignarTrazabilidad
        
    
        b = cPal.CrearPalet
    
    End If
    CadenaDesdeOtroForm = ""
    If b Then
        CadenaDesdeOtroForm = "OK"
        Unload Me
    Else
        Set cPal = Nothing
    End If
End Sub


Private Sub AsignarTrazabilidad()
Dim L As Long
    For i = 0 To 7
            cPal.LineasProd(i) = Me.Check1(i).Value = 1
            If Me.Check1(i).Value = 1 Then
                CadenaDesdeOtroForm = RecuperaValor(TrazaEnLineas, i + 1) 'el primero es la linea de paleticacion
                If CadenaDesdeOtroForm = "" Then
                    CadenaDesdeOtroForm = "0"
                Else
                    If Mid(CadenaDesdeOtroForm, 1, 1) = "#" Then CadenaDesdeOtroForm = "0"
                End If
                If Val(CadenaDesdeOtroForm) > 0 Then
                    L = CLng(CadenaDesdeOtroForm)
                    cPal.AñadirIdTraza CLng(CadenaDesdeOtroForm)
                    'Ahora, esperamos 1 momento
                    Espera 0.5
                    '
                    CadenaDesdeOtroForm = "prodlin.codigo = prodtrazlin.codigo  and prodlin.idlin=prodtrazlin.idlin AND lotetraza"
                    CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "tipoimpresionpalet", _
                        "prodlin , prodtrazlin", CadenaDesdeOtroForm, CStr(L))

                    
                    If CadenaDesdeOtroForm <> "" Then cPal.TipoImpresion = Val(CadenaDesdeOtroForm)
                    CadenaDesdeOtroForm = L
                    
                End If
                    
                    
            End If
    Next i
End Sub

Private Sub Form_Load()
Dim cad As String

    Me.Icon = frmppal.Icon
    
    If cPal Is Nothing Then
        Me.FrameNuevo.visible = True
        FrameCierrePalet.visible = False
        Me.cmdcancelar(0).Cancel = True
        
        
        Label3.Tag = RecuperaValor(Varios, 1)
        Varios = RecuperaValor(Varios, 2)
        
        For i = 0 To 7
            Me.Check1(i).Enabled = False
            
            If Mid(Varios, i + 1, 1) = "0" Then
                Me.Check1(i).Value = 0
            Else
                cad = RecuperaValor(Me.TrazaEnLineas, i + 1)  'Veo si Esta ya produciendose
                If Mid(cad, 1, 1) = "#" Then
                    'YA esta asignada a una linea
                    Me.Check1(i).Value = 2
                Else
                    Me.Check1(i).Enabled = True
                    Me.Check1(i).Value = 0
                End If
                
                
                'Acaba de dar de alta la produccion y va a  paletizarse YA
                If DesdeNuevaProd <> "" Then Me.Check1(i).Value = 1
                
                
            End If
        Next
        
        
        'Ahora si viene de Nueva produiccion
        Me.FrameTapaLineas.visible = DesdeNuevaProd <> ""
        If DesdeNuevaProd <> "" Then
            Label3.Tag = -1
            cad = ""
            For i = 0 To 3
                If Mid(DesdeNuevaProd, i + 1, 1) = "0" Then
                    
                    optPalet(i).Enabled = True
                    If cad = "" Then
                        optPalet(i).Value = True
                        cad = "YA"
                    End If
                Else
                    optPalet(i).Enabled = False
                End If
            Next i
            
           
            
        End If
    Else
        Me.FrameNuevo.visible = False
        FrameCierrePalet.visible = True
        Me.cmdcancelar(1).Cancel = True
        
        CargaDatosPalet2
        
    End If
    
    
    
    
    
    
    
End Sub


Private Sub CargaDatosPalet2()
Dim SQL As String
Dim cad As String
Dim N As Node
Dim i As Integer

    On Error GoTo ECargaDatosPalet

    Set miRsAux = New ADODB.Recordset
    Me.TreeView1.Nodes.Clear
    Total = 0
    cPal.CargaDatosPalet C, True, Total, False
    
    cad = ""
    For i = 1 To C.Count
        SQL = RecuperaValor(C(i), 1)
        
        If SQL <> cad Then
            cad = SQL
            Set N = TreeView1.Nodes.Add(, , "A" & cad, RecuperaValor(C(i), 2))
            
        End If
        SQL = "Lote: " & Format(RecuperaValor(C(i), 3), "000000") & "    Cajas: " & RecuperaValor(C(i), 4)
        Set N = TreeView1.Nodes.Add("A" & cad, tvwChild, , SQL)
        N.EnsureVisible
    Next i
ECargaDatosPalet:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set C = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DesdeNuevaProd = "" 'por si acaso
End Sub

Private Sub frmPa_DatoSeleccionado(CadenaSeleccion As String)
    Me.txtPalet.Text = CadenaSeleccion
End Sub

Private Sub imgPalet_Click()
    Set frmPa = New frmProdPalets
    frmPa.DatosADevolverBusqueda2 = "0"
    frmPa.Show vbModal
    Set frmPa = Nothing
End Sub
