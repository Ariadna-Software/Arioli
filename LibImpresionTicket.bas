Attribute VB_Name = "LibImpresionTicket"
Option Explicit


'David
'Llamara a esta funcion. Si el tipo de documento 32 (tickets) pone impresion directa, lo dejamos como esta, si no...
' hay que hacer a traves del rpt
Public Sub ImprimirTicketDirecto(NumTicket As String, FechaTicket As Date, Optional Entregado As Currency, Optional Cambio As Currency)   ' (RAFA/ALZIRA 05092006)
Dim Directo As Boolean
Dim Cadparam As String
Dim NumParam As Byte
Dim cadNomRPT As String
    
    Directo = True
    If Not PonerParamRPT(32, Cadparam, NumParam, cadNomRPT, Directo) Then Directo = True
        
    

    If Directo Then
        'Impresion directa
        ImprimirElTicketDirecto NumTicket, FechaTicket, Entregado, Cambio
        
    Else
        'Con cristal
        With frmImprimir
            .FormulaSeleccion = " {scafac.codtipom} = 'FTI'" & _
                " and {scafac.numfactu} = " & CStr(NumTicket) & _
                " and {scafac.fecfactu} = Date(" & Year(FechaTicket) & "," & Month(FechaTicket) & "," & Day(FechaTicket) & ")"
                
            .OtrosParametros = ""
            .NumeroParametros = 0
            .SoloImprimir = True
            .EnvioEMail = False
            .Opcion = 93
            .Titulo = "Ticket"
            .NombreRPT = cadNomRPT
            .ConSubInforme = True
            .Show vbModal
         End With
        
        
        
        'sI ABRE EL CAJON
        If vParamTPV.AbreCajon Then ImprimePorLaCom ""
               
    End If
End Sub




'Obligo la fecha. Antes NO y la cogia de rsventa
'Public Sub ImprimirTicketDirecto(NumTicket As String, NumAlbTicket1 As String, FechaTicket As Date)  ' (RAFA/ALZIRA 05092006)
Public Sub ImprimirElTicketDirecto(NumTicket As String, FechaTicket As Date, Optional Entregado As Currency, Optional Cambio As Currency)   ' (RAFA/ALZIRA 05092006)
    Dim NomImpre As String
  '  Dim FechaT As Date
    Dim rs1 As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim rs3 As ADODB.Recordset
    Dim rs4 As ADODB.Recordset
    Dim SQL As String
    Dim Lin As String ' línea de impresión
    Dim I As Integer
    Dim N As Integer
    Dim ImporteIva As Currency
    Dim EnEfectivo As Boolean
    
On Error GoTo EImpTickD
    'Antes DAVID
'    If FechaTicket = "" Then
'        FechaT = RSVenta!fecventa
'    Else
'        FechaT = CDate(FechaTicket)
'    End If
'       --> Como desde aqui no se ve el rsventa entonces OBLIGAMOS a que se traiga la fecha

'
'   Printer.Font = "Courier New"
    
    'Establecemos la impresora de ticket
    If vParamTPV.NomImpresora <> "" Then
        If Printer.DeviceName <> vParamTPV.NomImpresora Then
            'guardamos la impresora que habia
            NomImpre = Printer.DeviceName
            'establecemos la de ticket
            EstablecerImpresora vParamTPV.NomImpresora
        End If
    End If
    
    '-- Obtenemos cabeceras y pies en un recordset (rs1)
    SQL = "select * from spatpvg"
    Set rs1 = New ADODB.Recordset
    rs1.Open SQL, conn, adOpenForwardOnly
    If Not rs1.EOF Then
        ' Ahora buscamos las cabecera de ticket
        SQL = "select * from scafac where codtipom = 'FTI'" & _
                " and numfactu = " & CStr(NumTicket) & _
                " and fecfactu = '" & Format(FechaTicket, "yyyy-mm-dd") & "'"
        Set rs2 = New ADODB.Recordset
        rs2.Open SQL, conn, adOpenForwardOnly
        If Not rs2.EOF Then
            '-- Consultamos la forma de pago pa 2 cosas
            '   Para imprimirla en el pie y para en el caso de contado mostrar entregado
            '   y cambio.
            SQL = "select * from sforpa where codforpa = " & CStr(rs2!codforpa)
            Set rs4 = New ADODB.Recordset
            rs4.Open SQL, conn, adOpenForwardOnly
            If Not rs4.EOF Then
                If rs4!tipforpa = 0 Then EnEfectivo = True
            End If
            '-- Montar las líneas
            SQL = "select * from slifac where codtipom = 'FTI'" & _
                    " and numfactu = " & CStr(NumTicket) & _
                    " and fecfactu = '" & Format(FechaTicket, "yyyy-mm-dd") & "'"
            Set rs3 = New ADODB.Recordset
            rs3.Open SQL, conn, adOpenForwardOnly
            If Not rs3.EOF Then
                '-- Impresión de la cabecera
'                Lin = "         1         2         3         4"
'                Printer.Print Lin
'                Lin = "1234567890123456789012345678901234567890"
'                Printer.Print Lin
                For I = 1 To 5
                    If Not IsNull(rs1.Fields("cabtick" & CStr(I))) Then
                        Lin = LineaCentrada(rs1.Fields("cabtick" & CStr(I)))
                        If Lin <> "" Then Printer.Print Lin
                    End If
                Next I
                Lin = CuadraParteI(20, "Tiquet:" & Format(NumTicket, "0000000")) & _
                      CuadraParteD(20, "Fecha: " & Format(FechaTicket, "dd/mm/yyyy"))
                Printer.Print Lin
                rs3.MoveFirst
                Printer.Print ""
                Lin = CuadraParteI(40, "CLIENTE: " & Format(rs2!CodClien, "000000") & "  " & rs2!nomclien)
                Printer.Print Lin
                Printer.Print ""
                Lin = LineaCentrada("IVA INCLUIDO")
                Printer.Print Lin
                Lin = String(40, "-")
                Printer.Print Lin
                Lin = CuadraParteI(16, "DESCRIPCION") & _
                        CuadraParteD(6, " CANT") & _
                        CuadraParteD(8, "PVP") & _
                        CuadraParteD(10, "IMPORTE")
                Printer.Print Lin
                Lin = String(40, "-")
                Printer.Print Lin
                While Not rs3.EOF
                    '-- Una línea de impresión
                    
                    Lin = CuadraParteI(16, Mid(rs3!NomArtic, 1, 16)) & _
                            CuadraParteD(6, Format(rs3!Cantidad, "##0.00")) & _
                            CuadraParteD(8, Format(rs3!precioiv, "#,##0.00")) & _
                            CuadraParteD(10, Format(Round(rs3!Cantidad * rs3!precioiv, 2), "###,##0.00"))
                    Printer.Print Lin
                    rs3.MoveNext
                Wend
                '-- Impresion del total
                Printer.Print String(40, " ")
                Lin = CuadraParteI(20, "Total ticket: ") & CuadraParteD(20, Format(rs2!TotalFac, "###,###,#0.00"))
                Printer.Print Lin
                
                'Si deglosa IVAS
                If vParamTPV.DesglosaIVATicket Then
                    'Linea en blanco
                    'Lin = String(40, " ")
                    Printer.Print ""
                    
                    'Los tpios de IVA
                    Printer.Print "Detalle desglose IVAs"
                    
                    For I = 1 To 3
                        If Not IsNull(rs2.Fields("porciva" & CStr(I))) Then
                            'Lleva TIPO IVA
                            SQL = Format(DBLet(rs2.Fields("porciva" & CStr(I)), "N"), "0.00") & "%"
                            Lin = CuadraParteD(6, SQL)
                            'base
                            ImporteIva = DBLet(rs2.Fields("baseimp" & CStr(I)), "N")
                            SQL = Format(ImporteIva, "0.00")
                            Lin = Lin & CuadraParteD(10, SQL)
                            
                            'iva
                            SQL = Format(DBLet(rs2.Fields("imporiv" & CStr(I)), "N"), "0.00")
                            ImporteIva = ImporteIva + DBLet(rs2.Fields("imporiv" & CStr(I)), "N")
                            Lin = Lin & CuadraParteD(10, SQL)
                            'total
                            
                            SQL = Format(ImporteIva, FormatoImporte)
                            Lin = Lin & CuadraParteD(14, SQL)
                            Printer.Print Lin
                        End If
                    Next I
                End If
                '-- (RAFA 15/05/2008) -- Para Quatretonda
                '-- Imprimir la forma de pago
                Lin = CuadraParteI(40, "Forma de pago: " & DBLet(rs4!nomforpa, "T"))
                Printer.Print Lin
                If EnEfectivo Then
                    '-- Si han pagado en efectivo mostramos entregado y cambio.
                    Printer.Print String(40, " ")
                    SQL = Format(Entregado, "0.00")
                    Lin = CuadraParteI(20, "Entregado: " & SQL)
                    SQL = Format(Cambio, "0.00")
                    Lin = Lin & CuadraParteD(20, "Cambio: " & SQL)
                    Printer.Print Lin
                End If
                '-- Impresion del pie
                Printer.Print String(40, " ")
                For I = 1 To 3
                    If Not IsNull(rs1.Fields("pietick" & CStr(I))) Then
                        Lin = LineaCentrada(rs1.Fields("pietick" & CStr(I)))
                        If Lin <> "" Then Printer.Print Lin
                    End If
                Next I
                For I = 1 To 8
                    Printer.Print String(40, " ")
                Next I
                
                '-- Fin de impresión
                Printer.NewPage
                Printer.EndDoc

                
                
                'Abrir cajon
                            'El primer numero es el numero de caracteres de secuencia.
                            'Ej: 5|27|p|0|25|250|
                            '   Son 5:  27;p;0;25:250
                If vParamTPV.AbreCajon Then
                    
    
                        'De momento lo pongo a piñon
                        'N = RecuperaValor(vParamTPV.SecuenciaCajon, 1)
                        'Lin = ""
                        'For i = 1 To N
                        '    SQL = RecuperaValor(vParamTPV.SecuenciaCajon, i + 1)
                        '    If IsNumeric(SQL) Then
                        '        Lin = Lin & Chr(SQL)
                        '    Else
                        '        Lin = Lin & SQL
                        '    End If
                        'Next i
                        'Printer.Print Lin
                        ImprimePorLaCom ""
                End If
                
                
                
            Else
                MsgBox "No se han encontrado lineas del ticket " & CStr(NumTicket) & " de " & Format(FechaTicket, "dd/mm/yyyy"), vbCritical
            End If
            rs3.Close
        Else
            MsgBox "No se ha encontrado el ticket " & CStr(NumTicket) & " de " & Format(FechaTicket, "dd/mm/yyyy"), vbCritical
        End If
        rs2.Close
    Else
        MsgBox "Faltan los parámetros para la impresión del ticket", vbCritical
    End If
    rs1.Close
    
    
    'Volver la impresora a la predeterminada
    EstablecerImpresora NomImpre
    Exit Sub
EImpTickD:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Imprimir ticket."
End Sub


Private Sub ImprimePorLaCom(Cadena As String)
    On Error GoTo EI
    
    
    
    Dim nFicSalCajon As Integer
    Dim Puerto As String
    Puerto = "COM1"
    nFicSalCajon = FreeFile
    
    Open Puerto For Output As #nFicSalCajon
    'If Check1.Value = 1 Then
        Print #nFicSalCajon, Chr$(27); "p"; Chr$(0); Chr$(25); Chr$(250)
    'Else
    '    Print #nFicSalCajon, Cadena
    'End If
    Close nFicSalCajon
    
    Exit Sub
EI:
    Cadena = "Error en COM: " & vbCrLf & vbCrLf & Err.Description
    MsgBox Cadena, vbCritical
End Sub



Private Function LineaCentrada(Lin As String) As String
    Dim queda As Integer
    Dim parte As Integer
    queda = 40 - Len(Lin)
    parte = queda / 2
    If parte Then
        LineaCentrada = String(parte, " ") & Lin & String(queda - parte, " ")
    Else
        LineaCentrada = Lin
    End If
End Function

Private Function CuadraParteD(longitud As Integer, Cadena As String) As String
    CuadraParteD = Right(String(longitud, " ") & Cadena, longitud)
End Function

Private Function CuadraParteI(longitud As Integer, Cadena As String) As String
    CuadraParteI = Left(Cadena & String(longitud, " "), longitud)
End Function

