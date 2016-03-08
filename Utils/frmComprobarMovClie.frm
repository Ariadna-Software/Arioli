VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComprobarMovClie 
   Caption         =   "Form1"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   7905
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   6480
      TabIndex        =   4
      Top             =   3720
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2775
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "codartic"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "nomartic"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Incid"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Descrip"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Desde smoval"
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Desde albaranes"
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Text            =   "30/10/2009"
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Art."
      Height          =   255
      Left            =   5160
      TabIndex        =   7
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   3720
      Width           =   5895
   End
End
Attribute VB_Name = "frmComprobarMovClie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DesdeDocumento()
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim R2 As ADODB.Recordset
Dim vAr As String
Dim Fin As Boolean
Dim Situacion As String
Dim aUX As String


    If Text1.Text = "" Then Exit Sub
    
    If Not IsDate(Text1.Text) Then Exit Sub
    
    Label1.Caption = ""
    ListView1.ListItems.Clear
    'Albaranes
    SQL = "Select slialb.*,fechaalb from slialb,scaalb,sartic where "
    SQL = SQL & " scaalb.codtipom=slialb.codtipom and scaalb.numalbar=slialb.numalbar and"
    SQL = SQL & " sartic.codartic = slialb.codartic And sartic.ctrstock = 1 and codalmac=1"
    If Text2.Text <> "" Then SQL = SQL & " AND sartic.codartic = '" & Text2.Text & "'"
    Set Rs = New ADODB.Recordset
    Set R2 = New ADODB.Recordset
    
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    vAr = ""
    While Not Rs.EOF
        'Caption
        Label1.Caption = Rs!nomartic
        Label1.Refresh
        
        If vAr <> Rs!codartic Then
            If vAr <> "" Then R2.Close
            SQL = "Select * from smoval where codalmac=1 and codartic='" & Rs!codartic & "' order by document,numlinea,fechamov"
            R2.Open SQL, Conn, adOpenKeyset, adLockPessimistic, adCmdText
            
            If Not R2.EOF Then
                R2.MoveFirst
                vAr = Rs!codartic
            End If
        End If
        
        'Codartic
        If R2.EOF Then
            Fin = True
        Else
            Fin = False
        End If
        Situacion = "NO encontrado" 'NO ESTA
        aUX = Format(Rs!numalbar, "0000000")
        While Not Fin
        
            If R2!document = aUX Then
                If Rs!numlinea = R2!numlinea Then
                        'OK. Encontrada
                        'Veamos diferencias
                  
                        
                        If Rs!fechaalb <> R2!fechamov Then
                            Situacion = "Fecha"
                        Else
                            If Rs!cantidad <> R2!cantidad Then
                                Situacion = "Cantidad"
                            Else
                                Situacion = "" 'OK es esta
                            End If
                        End If
                        
                         If Situacion <> "" Then
                            
                            InsertaItem Rs, Situacion
                                
                    
                        End If
                         
                        
                        Fin = True
                End If
            End If
            If Not Fin Then
                R2.MoveNext
                Fin = R2.EOF
                If Fin Then InsertaItem Rs, "NO encontrado"
            Else
                R2.MoveFirst
            End If
                
    
        Wend
       
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    Set R2 = Nothing
    
    
    
    
    
    '**************************************************************
    'Facturas
    SQL = "select  sartic.codartic,sartic.nomartic,codalmac,c2.fechaalb,c2.codtipoa codtipom,l.cantidad,"
    SQL = SQL & " c2.numalbar,numlinea  from "
    SQL = SQL & " scafac c,scafac1 c2,slifac l,sartic where "
    SQL = SQL & " sartic.codartic=l.codartic and codalmac=1 AND"
    SQL = SQL & " c.codtipom=c2.codtipom and c.numfactu=c2.numfactu and  "
    SQL = SQL & " c.fecfactu=c2.fecfactu and l.codtipom=c2.codtipom and  l.numfactu=c2.numfactu and l.fecfactu=c2.fecfactu and l.codtipoa=c2.codtipoa and l.numalbar=c2.numalbar  "
    SQL = SQL & " and sartic.ctrstock=1 and c2.fechaalb>='" & Format(Text1.Text, FormatoFecha) & "'"

    If Text2.Text <> "" Then SQL = SQL & " AND sartic.codartic = '" & Text2.Text & "'"

    Set Rs = New ADODB.Recordset
    Set R2 = New ADODB.Recordset
    
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    vAr = ""
    While Not Rs.EOF
        'Caption
        Label1.Caption = Rs!nomartic
        Label1.Refresh
        
        If vAr <> Rs!codartic Then
            If vAr <> "" Then R2.Close
            SQL = "Select * from smoval where  codartic='" & Rs!codartic & "' order by document,numlinea,fechamov"
            R2.Open SQL, Conn, adOpenKeyset, adLockPessimistic, adCmdText
            
            If Not R2.EOF Then
                R2.MoveFirst
                vAr = Rs!codartic
            End If
        End If
        
        'Codartic
        If R2.EOF Then
            Fin = True
        Else
            Fin = False
        End If
        Situacion = "NO ESTA" 'NO ESTA
        aUX = Format(Rs!numalbar, "0000000")
        While Not Fin
        
            If R2!document = aUX Then
                If Rs!numlinea = R2!numlinea Then
                        'OK. Encontrada
                        'Veamos diferencias
                  
                        
                        If Rs!fechaalb <> R2!fechamov Then
                            Situacion = "Fecha"
                        Else
                            If Rs!cantidad <> R2!cantidad Then
                                Situacion = "Cantidad"
                            Else
                                Situacion = "" 'OK es esta
                            End If
                        End If
                        
                         If Situacion <> "" Then
                            
                            InsertaItem Rs, Situacion
                                
                        Else
    
                        End If
                         
                        
                        Fin = True
                End If
            End If
            If Not Fin Then
                R2.MoveNext
                Fin = R2.EOF
                If Fin Then InsertaItem Rs, "NO encontrado"
            Else
                R2.MoveFirst
            End If
                
    
        Wend
       
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    
    
    
    
    
End Sub


Private Sub InsertaItem(ByRef R As ADODB.Recordset, Incidencia As String)
    
Dim It As ListItem
          Set It = ListView1.ListItems.Add()
          It.Text = R!codartic
          It.SubItems(1) = R!nomartic
          It.SubItems(2) = Incidencia
          
          If Option1(0).Value Then
            It.SubItems(3) = R!numalbar & "  -  " & R!numlinea
          Else
            It.SubItems(3) = R!document & " " & R!numlinea
          End If
          
End Sub

Private Sub cmdAceptar_Click()
    If Me.Option1(0).Value Then
        DesdeDocumento
    Else
        DesdeMovimientos
    End If
    Label1.Caption = ""
End Sub




Private Sub DesdeMovimientos()


Dim Rs As ADODB.Recordset
Dim SQL As String
Dim RAlb As ADODB.Recordset
Dim RFra As ADODB.Recordset
Dim vAr As String
Dim Fin As Boolean
Dim Situacion As String
Dim aUX As String


    If Text1.Text = "" Then Exit Sub
    
    If Not IsDate(Text1.Text) Then Exit Sub
    
    Label1.Caption = ""
    ListView1.ListItems.Clear
    'Albaranes
    SQL = "Select smoval.*,nomartic from smoval,sartic where "
    SQL = SQL & " sartic.codartic = smoval.codartic And sartic.ctrstock = 1"
    SQL = SQL & " AND smoval.codalmac=1"
    SQL = SQL & " AND smoval.fechamov>='" & Format(Text1.Text, FormatoFecha) & "'"
    SQL = SQL & " AND detamovi IN ('ALZ','ALV') ORDER BY codartic "
    If Text2.Text <> "" Then SQL = SQL & " AND sartic.codartic = '" & Text2.Text & "'"
    Set Rs = New ADODB.Recordset
    Set RAlb = New ADODB.Recordset
    Set RFra = New ADODB.Recordset
    
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    vAr = ""
    While Not Rs.EOF
        'Caption
        Label1.Caption = Rs!nomartic
        Label1.Refresh
        
        If vAr <> Rs!codartic Then
            If vAr <> "" Then RAlb.Close
            
            'Albaranes
            SQL = "Select slialb.*,fechaalb from slialb,scaalb,sartic where "
            SQL = SQL & " scaalb.codtipom=slialb.codtipom and scaalb.numalbar=slialb.numalbar and"
            SQL = SQL & " sartic.codartic = slialb.codartic And sartic.codartic='" & Rs!codartic & "' and codalmac=1"
            SQL = SQL & " ORDER BY codtipom,numalbar,numlinea"
            RAlb.Open SQL, Conn, adOpenKeyset, adLockPessimistic, adCmdText
            
            
            'Fras
            SQL = "select  c2.fechaalb,c2.codtipoa codtipom,l.cantidad,"
            SQL = SQL & " c2.numalbar,numlinea  from "
            SQL = SQL & " scafac c,scafac1 c2,slifac l,sartic where "
            SQL = SQL & " sartic.codartic=l.codartic and codalmac=1 AND"
            SQL = SQL & " c.codtipom=c2.codtipom and c.numfactu=c2.numfactu and  "
            SQL = SQL & " c.fecfactu=c2.fecfactu and l.codtipom=c2.codtipom and  l.numfactu=c2.numfactu and l.fecfactu=c2.fecfactu and l.codtipoa=c2.codtipoa and l.numalbar=c2.numalbar  "
            SQL = SQL & " ORDER BY codtipom,numalbar,numlinea"
            If vAr <> "" Then RFra.Close
            RFra.Open SQL, Conn, adOpenKeyset, adLockPessimistic, adCmdText
            
            
            
            
            
            vAr = Rs!codartic

        End If
        
        'En el albaran
        Fin = RAlb.EOF
        Situacion = "NO encontrado" 'NO ESTA
        While Not Fin
        
        
                'SI
                    
                    aUX = Format(RAlb!numalbar, "0000000")
                    If Rs!detamovi = RAlb!codtipom And Rs!document = aUX Then
                        If Rs!numlinea = RAlb!numlinea Then
                                'OK. Encontrada
                                'Veamos diferencias
                          
                                
                                If Rs!fechamov <> RAlb!fechaalb Then
                                    Situacion = "Fecha"
                                Else
                                    If Rs!cantidad <> RAlb!cantidad Then
                                        Situacion = "Cantidad"
                                    Else
                                        Situacion = "" 'OK es esta
                                    End If
                                End If
                                
                                 If Situacion <> "" Then
                                    
                                    InsertaItem Rs, Situacion
                                        
                            
                                End If
                                 
                                
                                Fin = True
                        End If
                    End If
                    If Not Fin Then
                        RAlb.MoveNext
                        Fin = RAlb.EOF
                        
                        If Fin Then
                            'InsertaItem RS, "NO encontrado"
                            RAlb.MoveFirst
                        End If
                    Else
                        RAlb.MoveFirst
                    End If
                        
        
        Wend
        
        
        'Ahora las facturas
        If Situacion = "NO encontrado" Then
            'miro en las facturas
            Fin = RFra.EOF
            While Not Fin
        
        
                    'SI
                        aUX = Format(RFra!numalbar, "0000000")
                        If Rs!detamovi = RFra!codtipom And aUX = Rs!document Then
                            If Rs!numlinea = RFra!numlinea Then
                                    'OK. Encontrada
                                    'Veamos diferencias
                              
                                    
                                    If RFra!fechaalb <> Rs!fechamov Then
                                        Situacion = "Fecha"
                                    Else
                                        If Rs!cantidad <> RFra!cantidad Then
                                            Situacion = "Cantidad"
                                        Else
                                            Situacion = "" 'OK es esta
                                        End If
                                    End If
                                    
                                     If Situacion <> "" Then
                                        
                                        InsertaItem Rs, Situacion
                                            
                                
                                    End If
                                     
                                    
                                    Fin = True
                            End If
                        End If
                        If Not Fin Then
                            RFra.MoveNext
                            Fin = RFra.EOF
                            'Ahora no, miraremos en fras
                            If Fin Then
                                InsertaItem Rs, "NO encontrado"
                                RFra.MoveFirst
                            End If
                        Else
                            RFra.MoveFirst
                        End If
        
        
            Wend
           
           End If
       
            Rs.MoveNext
    Wend 'de moval
    Rs.Close
    Set Rs = Nothing
    Set RFra = Nothing
    Set RAlb = Nothing



End Sub


Private Sub Form_Load()
    Me.Icon = frmPpUtuil.Icon
End Sub

Private Sub ListView1_DblClick()
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    InputBox ListView1.SelectedItem.SubItems(1), , ListView1.SelectedItem.Text
End Sub


