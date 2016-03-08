VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmUtDeclara 
   Caption         =   "Declarar LOM"
   ClientHeight    =   2160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4050
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Index           =   1
      Left            =   2400
      TabIndex        =   6
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   65863681
      CurrentDate     =   39259
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   65863681
      CurrentDate     =   39259
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdComenzar 
      Caption         =   "Declaración"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Hasta Fecha"
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Desde Fecha:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblInf 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   3375
   End
End
Attribute VB_Name = "frmUtDeclara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As BaseDatos
Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim sql As String
Dim i As Integer
Dim cantidad As Double
Dim resto As Double

Private Sub cmdComenzar_Click()
Dim nomDocum As String

    '-- comprobamos que las fechas de paso son as correctas
    If DTPicker1(0).Value > DTPicker1(1).Value Then
        MsgBox "La fecha desde ha de ser menor o igual que la fecha hasta", vbInformation
        Exit Sub
    End If
    '-- Eliminamos posibles declaraciones anteriores
    sql = "delete from declaralom"
    db.ejecutar sql
    '-- Antes de empezar y como vamos a hacer uso de canasign, lo limpiamos
    sql = "update slotes set canasign = 0"
    db.ejecutar sql
    '-- Ahora vamos a por el gran mogollón
    sql = "select a.codtipom, a.numfactu, a.fecfactu, a.codartic, a.nomartic, a.cantidad " & _
            ",b.nomclien, b.nifclien" & _
            ",d.descateg" & _
            " from slifac as a, scafac as b, sartic as c, scateg as d" & _
            " where a.codartic in" & _
            " (select codartic from sartic" & _
            " where codcateg in (select codcateg from scateg where ctrlotes = 1))" & _
            " and b.codtipom = a.codtipom" & _
            " and b.numfactu = a.numfactu" & _
            " and b.fecfactu = a.fecfactu" & _
            " and c.codartic = a.codartic" & _
            " and d.codcateg = c.codcateg" & _
            " order by a.fecfactu"
    Set rs = db.cursor(sql)
    If Not rs.EOF Then
        rs.MoveFirst
        While Not rs.EOF
            i = i + 1
            lblInf.Caption = "Procesando.. " & CStr(i)
            lblInf.Refresh
            DoEvents
            cantidad = rs!cantidad
            resto = cantidad
            '-- Ahora que conocemos la cantidad, vamos a ver a que lotes la aplicamos
'            sql = "select distinct a.codartic, a.numlotes, a.fecentra, a.canentra, a.canasign, b.document, c.nomprove" & _
'                    " from slotes as a, smoval as b, sprove as c" & _
'                    " where a.codartic = " & db.Texto(rs!codartic) & _
'                    " and (a.canentra - a.canasign > 0)" & _
'                    " and a.fecentra <= " & db.fecha(rs!fecfactu) & _
'                    " and a.codartic = b.codartic" & _
'                    " and a.fecentra = b.fechamov" & _
'                    " and c.codprove = b.codigope" & _
'                    " and b.tipomovi = 1" & _
'                    " and b.detamovi = 'ALC'" & _
'                    " order by fecentra"

            sql = "select a.codartic, a.numlotes, a.fecentra, a.canentra, a.canasign, b.numserie from slotes as a, sartic as b" & _
                    " where a.codartic = " & db.Texto(rs!codArtic) & _
                    " and (a.canentra - a.canasign > 0)" & _
                    " and a.fecentra <= " & db.Fecha(rs!FecFactu) & _
                    " and b.codartic = a.codartic" & _
                    " order by a.fecentra"
            
            Set rs2 = db.cursor(sql)
            If Not rs2.EOF Then
                rs2.MoveFirst
                While Not rs2.EOF
                    If resto Then
                        If rs2!canentra - rs2!canasign >= resto Then
                            sql = "insert into declaralom(FechaVenta, NombreComercial, Registro, Categoria, Lote, Cantidad, NombreSocio, NIF, NumFactura)"
                            sql = sql & " values("
                            sql = sql & db.Fecha(rs!FecFactu) & "," ' FechaVenta
                            sql = sql & db.Texto(rs!NomArtic) & "," ' NombreComercial
                            'sql = sql & db.Texto(" ") & "," ' Registro
                            sql = sql & db.Texto(rs2!numSerie) & "," ' Registro
                            sql = sql & db.Texto(rs!descateg) & "," ' Categoria
                            sql = sql & db.Texto(rs2!numlotes) & "," ' Lote
                            sql = sql & TransformaComasPuntos(db.numero(resto)) & "," ' Cantidad
                            sql = sql & db.Texto(rs!nomclien) & "," ' NombreSocio
                            sql = sql & db.Texto(rs!nifClien) & "," ' NIF
                            sql = sql & db.Texto("V-" & Format(rs!NumFactu, "0000000")) & ")" ' NumFactura
                            db.ejecutar sql
                            '-- actualizamos las cantidades
                            sql = "update slotes set canasign = " & TransformaComasPuntos(db.numero(rs2!canasign + resto))
                            sql = sql & " where codartic = " & db.Texto(rs2!codArtic)
                            sql = sql & " and numlotes = " & db.Texto(rs2!numlotes)
                            sql = sql & " and fecentra = " & db.Fecha(rs2!fecentra)
                            db.ejecutar sql
                            resto = 0
                        Else
                            sql = "insert into declaralom(FechaVenta, NombreComercial, Registro, Categoria, Lote, Cantidad, NombreSocio, NIF, NumFactura)"
                            sql = sql & " values("
                            sql = sql & db.Fecha(rs!FecFactu) & "," ' FechaVenta
                            sql = sql & db.Texto(rs!NomArtic) & "," ' NombreComercial
                            sql = sql & db.Texto(" ") & "," ' Registro
                            sql = sql & db.Texto(rs!descateg) & "," ' Categoria
                            sql = sql & db.Texto(rs2!numlotes) & "," ' Lote
                            sql = sql & TransformaComasPuntos(db.numero(rs2!canentra - rs2!canasign)) & "," ' Cantidad
                            sql = sql & db.Texto(rs!nomclien) & "," ' NombreSocio
                            sql = sql & db.Texto(rs!nifClien) & "," ' NIF
                            sql = sql & db.Texto("V-" & Format(rs!NumFactu, "0000000")) & ")" ' NumFactura
                            db.ejecutar sql
                            '-- actualizamos las cantidades
                            sql = "update slotes set canasign = " & TransformaComasPuntos(db.numero(rs2!canentra))
                            sql = sql & " where codartic = " & db.Texto(rs2!codArtic)
                            sql = sql & " and numlotes = " & db.Texto(rs2!numlotes)
                            sql = sql & " and fecentra = " & db.Fecha(rs2!fecentra)
                            db.ejecutar sql
                            resto = resto - (rs2!canentra - rs2!canasign)
                        End If
                    End If
                    rs2.MoveNext
                Wend
            End If
            If resto > 0 Then
                '-- Quiere decir que tras haber consumido todos los lotes todavía nos queda resto
                '   lo grabamos pero sin ningún lote asociado.
                sql = "insert into declaralom(FechaVenta, NombreComercial, Registro, Categoria, Lote, Cantidad, NombreSocio, NIF, NumFactura)"
                sql = sql & " values("
                sql = sql & db.Fecha(rs!FecFactu) & "," ' FechaVenta
                sql = sql & db.Texto(rs!NomArtic) & "," ' NombreComercial
                sql = sql & db.Texto(" ") & "," ' Registro
                sql = sql & db.Texto(rs!descateg) & "," ' Categoria
                sql = sql & db.Texto(" ") & "," ' Lote
                sql = sql & db.numero(resto) & "," ' Cantidad
                sql = sql & db.Texto(rs!nomclien) & "," ' NombreSocio
                sql = sql & db.Texto(rs!nifClien) & "," ' NIF
                sql = sql & db.Texto("V-" & Format(rs!NumFactu, "0000000")) & ")" ' NumFactura
                db.ejecutar sql
            End If
            rs.MoveNext
        Wend
        '-- Por último actualizamos las compras
'        sql = "insert into declaralom (FechaVenta,NombreComercial,Registro,Categoria,Lote,Cantidad,NombreSocio,NIF,NumFactura,CanCompra)"
'        sql = sql & " select a.fecentra, b.nomartic, b.numserie, c.descateg, a.numlotes, 0, '---', '---','----', a.canentra"
'        sql = sql & " from slotes as a, sartic as b, scateg as c"
'        sql = sql & " where b.codartic = a.codartic"
'        sql = sql & " and c.codcateg = b.codcateg"
        
        
        sql = "insert into declaralom (FechaVenta,NombreComercial,Registro,Categoria,Lote,Cantidad,NombreSocio,NIF,NumFactura,CanCompra)"
        sql = sql & "select distinct a.fecentra, b.nomartic, b.numserie, c.descateg, a.numlotes, 0, e.nomprove, e.nifprove, d.document, a.canentra" & _
                " from slotes as a, sartic as b, scateg as c, smoval as d, sprove as e" & _
                " where b.codartic = a.codartic" & _
                " and c.codcateg = b.codcateg" & _
                " and d.codartic = a.codartic" & _
                " and d.fechamov = a.fecentra" & _
                " and d.tipomovi = 1" & _
                " and d.detamovi = 'ALC'" & _
                " and e.codprove = d.codigope"
        
        
        db.ejecutar sql
        lblInf.Caption = "Fin de generación. Lanzando informe..."
        lblInf.Refresh
        DoEvents
        '-- Llamar al informe
        Dim Desde As Date
        Dim Hasta As Date
        Desde = DTPicker1(0).Value
        Hasta = DTPicker1(1).Value
        frmVisReport.OtrosParametros = "|FecDesde=Date(" & Format(Desde, "yyyy") & _
                                            "," & Format(Desde, "mm") & _
                                            "," & Format(Desde, "dd") & ")|" & _
                                       "FecHasta=Date(" & Format(Hasta, "yyyy") & _
                                            "," & Format(Hasta, "mm") & _
                                            "," & Format(Hasta, "dd") & ")|"
        frmVisReport.NumeroParametros = 2
'        frmVisReport.Informe = App.Path & "\Informes\" & "declaracion_lom.rpt"
        
        'Añade los parametros de la tabla scrystal para el informe
        If Not PonerParamRPT(31, "", 0, nomDocum) Then
            Exit Sub
        End If
        frmVisReport.Informe = App.Path & "\Informes\" & nomDocum
        
        frmVisReport.FormulaSeleccion = "{declaralom.FechaVenta} in " & _
                                            "Date(" & Format(Desde, "yyyy") & _
                                            "," & Format(Desde, "mm") & _
                                            "," & Format(Desde, "dd") & ")" & _
                                            " to" & _
                                            " Date(" & Format(Hasta, "yyyy") & _
                                            "," & Format(Hasta, "mm") & _
                                            "," & Format(Hasta, "dd") & ")"
        frmVisReport.Show vbModal
        '--
        lblInf.Caption = "Proceso terminado."
        lblInf.Refresh
        DoEvents
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    '-- Abrimos la base de datos para trabajar con ella
    Set db = New BaseDatos
'    db.abrir "vAriges", "root", "aritel"
    db.asignar conn
    
    db.Tipo = "MYSQL"
    '-- Por defecto desde y hasta fecha de hoy
    DTPicker1(0).Value = Format(Date, "dd/mm/yyyy")
    DTPicker1(1).Value = Format(Date, "dd/mm/yyyy")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set db = Nothing
End Sub
