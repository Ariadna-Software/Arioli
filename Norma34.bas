Attribute VB_Name = "Norma34"
Option Explicit

Dim AuxD As String
Private NumeroTransferencia As Integer
'----------------------------------------------------------------------
'  Copia fichero generado bajo
Public Sub CopiarFicheroNorma43(Destino As String)

    
    'If Not CopiarEnDisquette(True, 3) Then
        AuxD = Destino
        CopiarEnDisquette False, 0  'A disco
    
        
End Sub

Private Function CopiarEnDisquette(A_disquetera As Boolean, Intentos As Byte) As Boolean
Dim I As Integer
Dim Cad As String

On Error Resume Next

    CopiarEnDisquette = False
    
    If A_disquetera Then
        For I = 1 To Intentos
            Cad = "Introduzca un disco vacio. (" & I & ")"
            MsgBox Cad, vbInformation
            FileCopy App.Path & "\norma34.txt", "a:\norma34.txt"
            If Err.Number <> 0 Then
                MuestraError Err.Number, "Copiar En Disquette"
            Else
                CopiarEnDisquette = True
                Exit For
            End If
        Next I
    Else
        If AuxD = "" Then
            Cad = Format(Now, "ddmmyyhhnn")
            Cad = App.Path & "\" & Cad & ".txt"
        Else
            Cad = AuxD
        End If
        FileCopy App.Path & "\norma34.txt", Cad
        If Err.Number <> 0 Then
            MsgBox "Error creando copia fichero. Consulte soporte técnico." & vbCrLf & Err.Description, vbCritical
        Else
            MsgBox "El fichero esta guardado como: " & Cad, vbInformation
        End If
            
    End If
End Function



'----------------------------------------------------------------------
'----------------------------------------------------------------------
'----------------------------------------------------------------------
'Cuenta propia tendra empipados entidad|sucursal|cc|cuenta|
Public Function GeneraFicheroNorma34(CIF As String, Fecha As Date, CuentaPropia As String, ConceptoTransferencia As String, vNumeroTransferencia As Integer, ByRef ConceptoTr As String, Pagos As Boolean) As Boolean
Dim NFich As Integer
Dim Regs As Integer
Dim CodigoOrdenante As String
Dim Importe As Currency
Dim Im As Currency
Dim RS As ADODB.Recordset
Dim Aux As String
Dim Cad As String


    On Error GoTo EGen
    GeneraFicheroNorma34 = False
    
    NumeroTransferencia = vNumeroTransferencia
    
    
    'Cargamos la cuenta
    Cad = "Select * from ctabancaria where codmacta='" & CuentaPropia & "'"
    Set RS = New ADODB.Recordset
    RS.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Aux = Right("    " & CIF, 10)
    If RS.EOF Then
        Cad = ""
    Else
        If IsNull(RS!entidad) Then
            Cad = ""
        Else
            Cad = Format(RS!entidad, "0000") & "|" & Format(DBLet(RS!oficina, "T"), "0000") & "|" & DBLet(RS!Control, "T") & "|" & Format(DBLet(RS!CtaBanco, "T"), "0000000000") & "|"
            CuentaPropia = Cad
        End If
        
        
        'Identificador norma bancaria
        If Not IsNull(RS!idnorma34) Then Aux = RS!idnorma34
    End If
    RS.Close
    Set RS = Nothing
    If Cad = "" Then
        MsgBox "Error leyendo datos para: " & CuentaPropia, vbExclamation
        Exit Function
    End If
    
    NFich = FreeFile
    Open App.Path & "\norma34.txt" For Output As #NFich
    
    
    
    
    
    'Codigo ordenante
    '---------------------------------------------------
    'Si el banco tiene puesto si ID de norma34 entonces
    'la pongo aquin. Lo he cargado antes sobre la variable AUX
    CodigoOrdenante = Left(Aux & "          ", 10)  'CIF EMPRESA
    
    
    'CABECERA
    Cabecera1 NFich, CodigoOrdenante, Fecha, CuentaPropia, Cad
    Cabecera2 NFich, CodigoOrdenante, Cad
    Cabecera3 NFich, CodigoOrdenante, Cad
    Cabecera4 NFich, CodigoOrdenante, Cad
    
    
    
    'Imprimimos las lineas
    'Para ello abrimos la tabla tmpNorma34
    Set RS = New ADODB.Recordset
    If Pagos Then
        Aux = "Select spagop.*,nommacta,dirdatos,codposta,dirdatos,despobla from spagop,cuentas"
        Aux = Aux & " where codmacta=ctaprove and transfer =" & NumeroTransferencia
    Else
        'ABONOS
         '
        Aux = "Select scobro.codbanco as entidad,scobro.codsucur as oficina,scobro.cuentaba,scobro.digcontr as CC"
        Aux = Aux & ",nommacta,dirdatos,codposta,dirdatos,despobla,impvenci,scobro.codmacta from scobro,cuentas"
        Aux = Aux & " where cuentas.codmacta=scobro.codmacta and transfer =" & NumeroTransferencia
    End If
    RS.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Importe = 0
    If RS.EOF Then
        'No hayningun registro
        
    Else
        Regs = 0
        While Not RS.EOF
            If Pagos Then
                Im = DBLet(RS!imppagad, "N")
                Im = RS!impefect - Im
                Aux = RellenaAceros(RS!CtaProve, False, 12)
            
            Else
                Im = Abs(RS!ImpVenci)
                Aux = RellenaAceros(RS!Codmacta, False, 12)
            End If
            
            'Cad = "06"
            'Cad = Cad & "56"
            'Cad = Cad & " "
            'Aux = "06" & "56" & " " & CodigoOrdenante & Aux  'Ordenante y socio juntos
        
            Aux = "06" & "56" & CodigoOrdenante & Aux   'Ordenante y socio juntos
        
            Linea1 NFich, Aux, RS, Im, Cad, ConceptoTransferencia
            Linea2 NFich, Aux, RS, Cad
            Linea3 NFich, Aux, RS, Cad
            Linea4 NFich, Aux, RS, Cad
            Linea5 NFich, Aux, RS, Cad
            Linea6 NFich, Aux, RS, Cad, ConceptoTr, Pagos
            If Pagos Then Linea7 NFich, Aux, RS, Cad
        
        
        
        
            Importe = Importe + Im
            Regs = Regs + 1
            RS.MoveNext
        Wend
        'Imprimimos totales
        Totales NFich, CodigoOrdenante, Importe, Regs, Cad, Pagos
    End If
    RS.Close
    Set RS = Nothing
    Close (NFich)
    If Regs > 0 Then GeneraFicheroNorma34 = True
    Exit Function
EGen:
    MuestraError Err.Number, Err.Description

End Function


'----------------------------------------------------------------------
'----------------------------------------------------------------------
'----------------------------------------------------------------------
'Cuenta propia tendra empipados entidad|sucursal|cc|cuenta|
Public Function GeneraFicheroNorma34New(CIF As String, Fecha As Date, CuentaPropia As String, ConceptoTransferencia As String, vNumeroTransferencia As Integer, ByRef ConceptoTr As String, Pagos As Boolean) As Boolean
Dim NFich As Integer
Dim Regs As Integer
Dim CodigoOrdenante As String
Dim Importe As Currency
Dim Im As Currency
Dim RS As ADODB.Recordset
Dim Aux As String
Dim Cad As String


    On Error GoTo EGen
    GeneraFicheroNorma34New = False
    
    NumeroTransferencia = vNumeroTransferencia
    
    
'    'Cargamos la cuenta
'    cad = "Select * from ctabancaria where codmacta='" & CuentaPropia & "'"
'    Set Rs = New ADODB.Recordset
'    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Aux = Right("    " & CIF, 10)
'    If Rs.EOF Then
'        cad = ""
'    Else
'        If IsNull(Rs!entidad) Then
'            cad = ""
'        Else
'            cad = Format(Rs!entidad, "0000") & "|" & Format(DBLet(Rs!oficina, "T"), "0000") & "|" & DBLet(Rs!Control, "T") & "|" & Format(DBLet(Rs!CtaBanco, "T"), "0000000000") & "|"
'            CuentaPropia = cad
'        End If
'
'
'        'Identificador norma bancaria
'        If Not IsNull(Rs!idnorma34) Then Aux = Rs!idnorma34
'    End If
'    Rs.Close
'    Set Rs = Nothing
'    If cad = "" Then
'        MsgBox "Error leyendo datos para: " & CuentaPropia, vbExclamation
'        Exit Function
'    End If

    NFich = FreeFile
    Open App.Path & "\norma34.txt" For Output As #NFich
    
    
    'Codigo ordenante
    '---------------------------------------------------
    'Si el banco tiene puesto si ID de norma34 entonces
    'la pongo aquin. Lo he cargado antes sobre la variable AUX
    CodigoOrdenante = Left(Aux & "          ", 10)  'CIF EMPRESA
    
    
    'CABECERA
    Cabecera1 NFich, CodigoOrdenante, Fecha, CuentaPropia, Cad
    Cabecera2 NFich, CodigoOrdenante, Cad
    Cabecera3 NFich, CodigoOrdenante, Cad
    Cabecera4 NFich, CodigoOrdenante, Cad
    
    
    
    'Imprimimos las lineas
    'Para ello abrimos la tabla tmpNorma34
    Set RS = New ADODB.Recordset
    Aux = "Select codtraba,sum(impnomi), sum(impgasto)"
    
'    Aux = "select tmpimpor.*, straba.codbanco as entidad, straba.codsucur as oficina, straba.digcontr as CC, straba.cuentaba as cuentaba, "
'    Aux = Aux & " straba.nomtraba as nommacta, straba.domtraba as dirdatos, straba.codpobla as codposta, straba.pobtraba as despobla "
'    Aux = Aux & " from tmpimpor, straba where tmpimpor.codtraba = straba.codtraba "
    
''    If Pagos Then
''        Aux = "Select spagop.*,nommacta,dirdatos,codposta,dirdatos,despobla from spagop,cuentas"
''        Aux = Aux & " where codmacta=ctaprove and transfer =" & NumeroTransferencia
''    Else
''        'ABONOS
''         '
''        Aux = "Select scobro.codbanco as entidad,scobro.codsucur as oficina,scobro.cuentaba,scobro.digcontr as CC"
''        Aux = Aux & ",nommacta,dirdatos,codposta,dirdatos,despobla,impvenci,scobro.codmacta from scobro,cuentas"
''        Aux = Aux & " where cuentas.codmacta=scobro.codmacta and transfer =" & NumeroTransferencia
''    End If



    RS.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Importe = 0
    If RS.EOF Then
        'No hayningun registro
        
    Else
        Regs = 0
        While Not RS.EOF
'            If Pagos Then
'                Im = DBLet(Rs!imppagad, "N")
'                Im = Rs!impefect - Im
'                Aux = RellenaAceros(Rs!CtaProve, False, 12)
'
'            Else
'                Im = Abs(Rs!ImpVenci)
'                Aux = RellenaAceros(Rs!Codmacta, False, 12)
'            End If

            Im = DBLet(RS!Importe, "N")
            Aux = RellenaAceros("0", False, 12) 'Rs!Codmacta, False, 12)
            
            'Cad = "06"
            'Cad = Cad & "56"
            'Cad = Cad & " "
            'Aux = "06" & "56" & " " & CodigoOrdenante & Aux  'Ordenante y socio juntos
        
            Aux = "06" & "56" & CodigoOrdenante & Aux   'Ordenante y socio juntos
        
            Linea1 NFich, Aux, RS, Im, Cad, ConceptoTransferencia
            Linea2 NFich, Aux, RS, Cad
            Linea3 NFich, Aux, RS, Cad
            Linea4 NFich, Aux, RS, Cad
            Linea5 NFich, Aux, RS, Cad
            Linea6 NFich, Aux, RS, Cad, ConceptoTr, Pagos
            If Pagos Then Linea7 NFich, Aux, RS, Cad
        
            Importe = Importe + Im
            Regs = Regs + 1
            RS.MoveNext
        Wend
        'Imprimimos totales
        Totales NFich, CodigoOrdenante, Importe, Regs, Cad, Pagos
    End If
    RS.Close
    Set RS = Nothing
    Close (NFich)
    If Regs > 0 Then GeneraFicheroNorma34New = True
    Exit Function
    
EGen:
    MuestraError Err.Number, Err.Description
End Function





'----------------------------------------------------------------------
'----------------------------------------------------------------------
'----------------------------------------------------------------------
'Cuenta propia tendra empipados entidad|sucursal|cc|cuenta|
Public Function GeneraFicheroNorma34_ARIGES(CIF As String, Fecha As Date, CuentaPropia As String, ConceptoTrans As String, vNumeroTransferencia As Integer, ByRef DescripcionTrans As String, Pagos As Boolean, cadSQL As String) As Boolean
Dim NFich As Integer
Dim nomFich As String
Dim Regs As Integer
Dim CodigoOrdenante As String
Dim Importe As Currency
Dim Im As Currency
Dim RS As ADODB.Recordset
Dim Aux As String
Dim Cad As String


    On Error GoTo EGen
    GeneraFicheroNorma34_ARIGES = False
    
    NumeroTransferencia = vNumeroTransferencia
    
    

    NFich = FreeFile
    Open App.Path & "\norma34.txt" For Output As #NFich
    
    
    'Codigo ordenante
    '---------------------------------------------------
    'Si el banco tiene puesto si ID de norma34 entonces
    'la pongo aquin. Lo he cargado antes sobre la variable AUX

    CodigoOrdenante = Left(CIF & "          ", 10)  'CIF EMPRESA
    
    
    'CABECERA
    Cabecera1 NFich, CodigoOrdenante, Fecha, CuentaPropia, Cad
    Cabecera2 NFich, CodigoOrdenante, Cad
    Cabecera3 NFich, CodigoOrdenante, Cad
    Cabecera4 NFich, CodigoOrdenante, Cad
    
    
    
    'Imprimimos las lineas
    'Para ello abrimos la tabla tmpNorma34
    Set RS = New ADODB.Recordset
    Aux = cadSQL
    

    RS.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Importe = 0
    If RS.EOF Then
        'No hayningun registro
        
    Else
        Regs = 0
        While Not RS.EOF
            frmListadoNomi.lblProgreso.Caption = "Trabajador: " & DBLet(RS!nommacta, "T") & "     " & Regs + 1 & " de " & frmListadoNomi.ProgressBar1.Max
'            If Pagos Then
'                Im = DBLet(Rs!imppagad, "N")
'                Im = Rs!impefect - Im
'                Aux = RellenaAceros(Rs!CtaProve, False, 12)
'
'            Else
'                Im = Abs(Rs!ImpVenci)
'                Aux = RellenaAceros(Rs!Codmacta, False, 12)
'            End If

            Im = DBLet(RS!Importe, "N")
            
            'Codigo beneficiario
            '---------------------------------------------------
            'Aux = Left(DBLet(RS!refbenef, "T") & "          ", 12) 'NIF beneficiari (trabajador)
            Aux = RellenaABlancos(DBLet(RS!refbenef, "T"), True, 12) 'NIF beneficiari (trabajador)
            
'            Aux = RellenaAceros("0", False, 12) 'Rs!Codmacta, False, 12)
            
            'Cad = "06"
            'Cad = Cad & "56"
            'Cad = Cad & " "
            'Aux = "06" & "56" & " " & CodigoOrdenante & Aux  'Ordenante y socio juntos
        
            Aux = "06" & "56" & CodigoOrdenante & Aux   'Ordenante y socio juntos
        
            Linea1 NFich, Aux, RS, Im, Cad, ConceptoTrans
            Linea2 NFich, Aux, RS, Cad
            Linea3 NFich, Aux, RS, Cad
            Linea4 NFich, Aux, RS, Cad
            Linea5 NFich, Aux, RS, Cad
            Linea6 NFich, Aux, RS, Cad, DescripcionTrans, Pagos
            If Pagos Then Linea7 NFich, Aux, RS, Cad
        
            Importe = Importe + Im
            Regs = Regs + 1
            IncrementarProgresNew frmListadoNomi.ProgressBar1, 1
            
            RS.MoveNext
        Wend
        'Imprimimos totales
        frmListadoNomi.lblProgreso.Caption = "Total..."
        Totales NFich, CodigoOrdenante, Importe, Regs, Cad, Pagos
    End If
    
    RS.Close
    Set RS = Nothing
    Close (NFich)
    frmListadoNomi.lblProgreso.Caption = "Proceso finalizado."
    
    If Regs > 0 Then GeneraFicheroNorma34_ARIGES = True
    Exit Function
    
EGen:
    MuestraError Err.Number, Err.Description
End Function





Private Function RellenaABlancos(Cadena As String, PorLaDerecha As Boolean, longitud As Integer) As String
Dim Cad As String
    
    Cad = Space(longitud)
    If PorLaDerecha Then
        Cad = Cadena & Cad
        RellenaABlancos = Left(Cad, longitud)
    Else
        Cad = Cad & Cadena
        RellenaABlancos = Right(Cad, longitud)
    End If
    
End Function



Private Function RellenaAceros(Cadena As String, PorLaDerecha As Boolean, longitud As Integer) As String
Dim Cad As String
    
    Cad = Mid("00000000000000000000", 1, longitud)
    If PorLaDerecha Then
        Cad = Cadena & Cad
        RellenaAceros = Left(Cad, longitud)
    Else
        Cad = Cad & Cadena
        RellenaAceros = Right(Cad, longitud)
    End If
    
End Function




Private Sub Cabecera1(NF As Integer, ByRef CodOrde As String, Fecha As Date, cta As String, ByRef Cad As String)

    Cad = "03"
    Cad = Cad & "56"
    'cad = cad & " "
    Cad = Cad & CodOrde
    Cad = Cad & Space(12) & "001"
    Cad = Cad & Format(Now, "ddmmyy")
    Cad = Cad & Format(Fecha, "ddmmyy")
    'Cuenta bancaria
    Cad = Cad & RecuperaValor(cta, 1)
    Cad = Cad & RecuperaValor(cta, 2)
    Cad = Cad & RecuperaValor(cta, 4)
    Cad = Cad & "0"  'Sin relacion
    Cad = Cad & "   " & RecuperaValor(cta, 3)  'Digito de control bancario
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub



Private Sub Cabecera2(NF As Integer, ByRef CodOrde As String, ByRef Cad As String)
    Cad = "03"
    Cad = Cad & "56"
    'cad = cad & " "
    Cad = Cad & CodOrde
    Cad = Cad & Space(12) & "002"
    
    Cad = Cad & RellenaABlancos(vParam.NombreEmpresa, True, 30)   'Nombre empresa
  
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub


Private Sub Cabecera3(NF As Integer, ByRef CodOrde As String, ByRef Cad As String)
    Cad = "03"
    Cad = Cad & "56"
    'cad = cad & " "
    Cad = Cad & CodOrde
    Cad = Cad & Space(12) & "003"
    
    
'    AuxD = DevuelveDesdeBD("direccion", "empresa2", "codigo", 1, "N")
    Cad = Cad & RellenaABlancos(vParam.DomicilioEmpresa, True, 30) 'AuxD, True, 30)   'Nombre empresa
    Cad = Cad & RellenaABlancos("", True, 30)   'Nombre empresa
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub



Private Sub Cabecera4(NF As Integer, ByRef CodOrde As String, ByRef Cad As String)

    Cad = "03"
    Cad = Cad & "56"
    'cad = cad & " "
    Cad = Cad & CodOrde
    Cad = Cad & Space(12) & "004"
    
'    AuxD = DevuelveDesdeBD("codpos", "empresa2", "codigo", 1, "N")
    Cad = Cad & RellenaABlancos(vParam.CPostal, False, 5) '   AuxD, False, 5)
    Cad = Cad & " "
'    AuxD = DevuelveDesdeBD("provincia", "empresa2", "codigo", 1, "N")
    Cad = Cad & RellenaABlancos(vParam.Provincia, True, 30) 'AuxD, True, 30)
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub



'ConceptoTransferencia
'1.- Abono nomina
'9.- Transferencia ordinaria
Private Sub Linea1(NF As Integer, ByRef CodOrde As String, ByRef rs1 As ADODB.Recordset, ByRef Importe1 As Currency, ByRef Cad As String, vConceptoTransferencia As String)


   
    '
    Cad = CodOrde   'llevara tb la ID del socio
    Cad = Cad & "010"
    Cad = Cad & RellenaAceros(CStr(Round(Importe1, 2) * 100), False, 12)
    
    Cad = Cad & RellenaAceros(CStr(rs1!entidad), False, 4)     'Entidad
    Cad = Cad & RellenaAceros(CStr(rs1!oficina), False, 4)   'Sucur
    Cad = Cad & RellenaAceros(CStr(rs1!cuentaba), False, 10)  'Cta
    Cad = Cad & "1" & vConceptoTransferencia
    Cad = Cad & "  "
    Cad = Cad & RellenaAceros(CStr(rs1!CC), False, 2)  'CC
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub


Private Sub Linea2(NF As Integer, ByRef CodOrde As String, ByRef rs1 As ADODB.Recordset, ByRef Cad As String)
    Cad = CodOrde    'llevara tb la ID del socio
    Cad = Cad & "011"
    Cad = Cad & RellenaABlancos(rs1!nommacta, False, 36)
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub


Private Sub Linea3(NF As Integer, ByRef CodOrde As String, ByRef rs1 As ADODB.Recordset, ByRef Cad As String)
    Cad = CodOrde    'llevara tb la ID del socio
    Cad = Cad & "012"
    Cad = Cad & RellenaABlancos(DBLet(rs1!dirdatos, "T"), False, 36)
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub


Private Sub Linea4(NF As Integer, ByRef CodOrde As String, ByRef rs1 As ADODB.Recordset, ByRef Cad As String)
    Cad = CodOrde    'llevara tb la ID del socio
    Cad = Cad & "013"
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub


Private Sub Linea5(NF As Integer, ByRef CodOrde As String, ByRef rs1 As ADODB.Recordset, ByRef Cad As String)
    Cad = CodOrde    'llevara tb la ID del socio
    Cad = Cad & "014"
    Cad = Cad & RellenaABlancos(DBLet(rs1!codposta, "T"), False, 5) & " "
    Cad = Cad & RellenaABlancos(DBLet(rs1!despobla, "T"), False, 30)
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub


Private Sub Linea6(NF As Integer, ByRef CodOrde As String, ByRef rs1 As ADODB.Recordset, ByRef Cad As String, ByRef ConceptoT As String, Pagos As Boolean)
Dim Aux As String

    Aux = ConceptoT
'    If Pagos Then
'        'Tiene dos campos para las descripcion. Si no tiene nada pondre la descripcion de la transferencia
'        Aux = Trim(DBLet(RS1!text1csb, "T"))
'        If Aux = "" Then Aux = ConceptoT
'    End If

    Cad = CodOrde    'llevara tb la ID del socio
    Cad = Cad & "016"
    Cad = Cad & RellenaABlancos(Aux, False, 35)
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub


Private Sub Linea7(NF As Integer, ByRef CodOrde As String, ByRef rs1 As ADODB.Recordset, ByRef Cad As String)


    Cad = CodOrde    'llevara tb la ID del socio
    Cad = Cad & "017"
    Cad = Cad & RellenaABlancos(DBLet(rs1!text2csb, "T"), False, 35)
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub




Private Sub Totales(NF As Integer, ByRef CodOrde As String, Total As Currency, Registros As Integer, ByRef Cad As String, Pagos As Boolean)
    Cad = "08" & "56"
    Cad = Cad & CodOrde    'llevara tb la ID del socio
    Cad = Cad & Space(15)
    Cad = Cad & RellenaAceros(CStr(Int(Round(Total * 100, 2))), False, 12)
    Cad = Cad & RellenaAceros(CStr(Registros), False, 8)
    If Pagos Then
        Cad = Cad & RellenaAceros(CStr((Registros * 7) + 4 + 1), False, 10)
    Else
        Cad = Cad & RellenaAceros(CStr((Registros * 6) + 4 + 1), False, 10)
    End If
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub






'**************************************************************************************
'**************************************************************************************
'**************************************************************************************


'   XML n34

'**************************************************************************************
'**************************************************************************************
'**************************************************************************************


Public Function GeneraFicheroNorma34SEPA_XML(CIF As String, Fecha As Date, CuentaPropia2 As String, cadSQL As String, DescripcionTrans As String) As Boolean
Dim Regs As Integer
Dim Importe As Currency
Dim Im As Currency
Dim Cad As String
Dim Aux As String
Dim SufijoOEM As String
Dim NFic As Integer
Dim EsPersonaJuridica2 As Boolean

Dim ContabilidadNueva As Boolean  'Llegado el momento habra que poner   vParamAplic.ContabilidadNueva


    On Error GoTo EGen3
    GeneraFicheroNorma34SEPA_XML = False
    ContabilidadNueva = False
    NFic = -1
    
    
    'Cargamos la cuenta
    
     
    'Cargamos la cuenta
    Cad = "cuentaba='" & RecuperaValor(CuentaPropia2, 4) & "' and codbanco=" & RecuperaValor(CuentaPropia2, 1) & " and iban='" & RecuperaValor(CuentaPropia2, 5) & "' AND codsucur"
    Cad = DevuelveDesdeBD(conAri, "idnorma34", "sbanpr", Cad, RecuperaValor(CuentaPropia2, 2))
    SufijoOEM = Right("000" & Cad, 3)
 
    Cad = RecuperaValor(CuentaPropia2, 5) & RecuperaValor(CuentaPropia2, 1) & RecuperaValor(CuentaPropia2, 2) & RecuperaValor(CuentaPropia2, 3) & RecuperaValor(CuentaPropia2, 4)
        
    If Len(Cad) <> 24 Then
        MsgBox "Error leyendo datos para: " & CuentaPropia2, vbExclamation
        Exit Function
    End If
    CuentaPropia2 = Cad
    
    
   
    
    NFic = FreeFile
    Open App.Path & "\norma34.txt" For Output As NFic
    
    
    Print #NFic, "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
    Print #NFic, "<Document xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""urn:iso:std:iso:20022:tech:xsd:pain.001.001.03"">"
    Print #NFic, "<CstmrCdtTrfInitn>"
    Print #NFic, "   <GrpHdr>"
    Cad = "TRANPAG000000" & Format(Now, "yymmdd") & "F" & Format(Now, "yyyymmddThhnnss")
    Print #NFic, "      <MsgId>" & Cad & "</MsgId>"
    Print #NFic, "      <CreDtTm>" & Format(Now, "yyyy-mm-ddThh:nn:ss") & "</CreDtTm>"
    
    Cad = cadSQL
    Set miRsAux = New ADODB.Recordset
    
   
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Im = 0
    Regs = 0
    While Not miRsAux.EOF
        Regs = Regs + 1
        Im = Im + miRsAux!Importe
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    Aux = Regs & "|" & Format(Im, "#.00") & "|"
    Print #NFic, "      <NbOfTxs>" & RecuperaValor(Aux, 1) & "</NbOfTxs>"
    Print #NFic, "      <CtrlSum>" & TransformaComasPuntos(RecuperaValor(Aux, 2)) & "</CtrlSum>"
    Print #NFic, "      <InitgPty>"
    Print #NFic, "         <Nm>" & XML(vEmpresa.nomempre) & "</Nm>"
    Print #NFic, "         <Id>"
    Cad = Mid(CIF, 1, 1)
    
    EsPersonaJuridica2 = Not IsNumeric(Cad)

    
    
    
    Cad = "PrvtId"
    If EsPersonaJuridica2 Then Cad = "OrgId"
    
    Print #NFic, "           <" & Cad & ">"
    Print #NFic, "               <Othr>"
    Print #NFic, "                  <Id>" & CIF & SufijoOEM & "</Id>"
    Print #NFic, "               </Othr>"
    Print #NFic, "           </" & Cad & ">"
    
    Print #NFic, "         </Id>"
    Print #NFic, "      </InitgPty>"
    Print #NFic, "   </GrpHdr>"

    Print #NFic, "   <PmtInf>"
    
    Print #NFic, "      <PmtInfId>" & Format(Now, "yyyymmddhhnnss") & CIF & "</PmtInfId>"
    Print #NFic, "      <PmtMtd>TRF</PmtMtd>"
    Print #NFic, "      <ReqdExctnDt>" & Format(Fecha, "yyyy-mm-dd") & "</ReqdExctnDt>"
    Print #NFic, "      <Dbtr>"
    
     'Nombre


    miRsAux.Open "Select domempre ,codpobla,pobempre,proempre from sparam"
    
    If miRsAux.EOF Then Err.Raise 513, , "Error obteniendo datos empresa(empresa2)"
    
    Print #NFic, "         <Nm>" & XML(vEmpresa.nomempre) & "</Nm>"
    Print #NFic, "         <PstlAdr>"
    Print #NFic, "            <Ctry>ES</Ctry>"

    Cad = miRsAux!domempre & " "
    Cad = Cad & Trim(DBLet(miRsAux!codpobla, "T") & " " & miRsAux!pobempre) & " "
    Cad = Cad & DBLet(miRsAux!proempre, "T")
    miRsAux.Close
    Print #NFic, "            <AdrLine>" & XML(Trim(Cad)) & "</AdrLine>"
    
    Print #NFic, "         </PstlAdr>"
    Print #NFic, "         <Id>"
    
    Aux = "PrvtId"
    If EsPersonaJuridica2 Then Aux = "OrgId"
   
    
    Print #NFic, "            <" & Aux & ">"
    
    Print #NFic, "               <Othr>"
    Print #NFic, "                  <Id>" & CIF & SufijoOEM & "</Id>"
    Print #NFic, "               </Othr>"
    Print #NFic, "            </" & Aux & ">"
    Print #NFic, "         </Id>"
    Print #NFic, "    </Dbtr>"
    
    
    Print #NFic, "    <DbtrAcct>"
    Print #NFic, "       <Id>"
    Print #NFic, "          <IBAN>" & Trim(CuentaPropia2) & "</IBAN>"
    Print #NFic, "       </Id>"
    Print #NFic, "       <Ccy>EUR</Ccy>"
    Print #NFic, "    </DbtrAcct>"
    Print #NFic, "    <DbtrAgt>"
    Print #NFic, "       <FinInstnId>"
    
    Cad = Mid(CuentaPropia2, 5, 4)
    'Contabilidad nueva
    If ContabilidadNueva Then
        Cad = DevuelveDesdeBD(conConta, "bic", "bics", "entidad", Cad)
    Else
        Cad = DevuelveDesdeBD(conConta, "bic", "sbic", "entidad", Cad)
    End If
    Print #NFic, "          <BIC>" & Trim(Cad) & "</BIC>"
    Print #NFic, "       </FinInstnId>"
    Print #NFic, "    </DbtrAgt>"
    
    
    
    'sql
    Cad = cadSQL
    miRsAux.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
    Regs = 0
    While Not miRsAux.EOF
        Print #NFic, "   <CdtTrfTxInf>"
        Print #NFic, "      <PmtId>"
        
        Aux = miRsAux!refbenef
         
        
        Print #NFic, "         <EndToEndId>" & Aux & "</EndToEndId>"
        Print #NFic, "      </PmtId>"
        Print #NFic, "      <PmtTpInf>"
        
        Im = Abs(miRsAux!Importe)
    

        
        'Persona fisica o juridica
        Cad = Mid(miRsAux!refbenef, 1, 1)
        EsPersonaJuridica2 = Not IsNumeric(Cad)

        
        
        Importe = Importe + Im
        Regs = Regs + 1
        
        Print #NFic, "          <SvcLvl><Cd>SEPA</Cd></SvcLvl>"

        Aux = "SALA"
'        ElseIf ConceptoTr = "0" Then
'            AUX = "PENS"
'        Else
'            AUX = "TRAD"
'        End If
        Print #NFic, "          <CtgyPurp><Cd>" & Aux & "</Cd></CtgyPurp>"
        Print #NFic, "       </PmtTpInf>"
        Print #NFic, "       <Amt>"
        Print #NFic, "          <InstdAmt Ccy=""EUR"">" & TransformaComasPuntos(CStr(Im)) & "</InstdAmt>"
        Print #NFic, "       </Amt>"
        Print #NFic, "       <CdtrAgt>"
        Print #NFic, "          <FinInstnId>"
        Cad = DBLet(miRsAux!entidad, "T")
        If Cad <> "" Then
            If ContabilidadNueva Then
                Cad = DevuelveDesdeBD(conConta, "bic", "bics", "entidad", Cad)
            Else
                Cad = DevuelveDesdeBD(conConta, "bic", "sbic", "entidad", Cad)
            End If
        End If
        If Cad = "" Then Err.Raise 513, , "No existe BIC: " & miRsAux!nommacta & vbCrLf & "Entidad: " & Cad
        
        Print #NFic, "             <BIC>" & Cad & "</BIC>"
        Print #NFic, "          </FinInstnId>"
        Print #NFic, "       </CdtrAgt>"
        Print #NFic, "       <Cdtr>"
        Print #NFic, "          <Nm>" & XML(miRsAux!nommacta) & "</Nm>"
        
        

        
        
        Print #NFic, "           <Id>"
        Aux = "PrvtId"
        If EsPersonaJuridica2 Then Aux = "OrgId"
      
        Print #NFic, "               <" & Aux & ">"
        Print #NFic, "                  <Othr>"
        Print #NFic, "                     <Id>" & miRsAux!refbenef & "</Id>"
        'Da problemas.... con Cajamar
        'Print #NFic, "                     <Issr>NIF</Issr>"
        Print #NFic, "                  </Othr>"
        Print #NFic, "               </" & Aux & ">"
        Print #NFic, "           </Id>"
        Print #NFic, "        </Cdtr>"
        Print #NFic, "        <CdtrAcct>"
        Print #NFic, "           <Id>"
        Print #NFic, "              <IBAN>" & IBAN_Destino() & "</IBAN>"
        Print #NFic, "           </Id>"
        Print #NFic, "        </CdtrAcct>"
        Print #NFic, "      <Purp>"
        
       ' If ConceptoTr = "1" Then
            Aux = "SALA"
       ' ElseIf ConceptoTr = "0" Then
       '     AUX = "PENS"
       ' Else
       '     AUX = "TRAD"
       ' End If
        
        Print #NFic, "         <Cd>" & Aux & "</Cd>"
        Print #NFic, "      </Purp>"
        Print #NFic, "      <RmtInf>"
        'Print #NFic, "      <Ustrd>ESTE ES EL CONCEPTO, POR TANTO NO SE SI SERA EL TEXTO QUE IRA DONDE TIENE QUE IR, O EN OTRO LADAO... A SABER. LO QUE ESTA CLARO ES QUE VA.</Ustrd>

        
        Aux = DescripcionTrans
        If Trim(Aux) = "" Then Aux = miRsAux!nommacta
        Print #NFic, "         <Ustrd>" & XML(Trim(Aux)) & "</Ustrd>"
        Print #NFic, "      </RmtInf>"
        Print #NFic, "   </CdtTrfTxInf>"
 
       
    
            
        miRsAux.MoveNext
    Wend
    Print #NFic, "   </PmtInf>"
    Print #NFic, "</CstmrCdtTrfInitn></Document>"
    
    
    miRsAux.Close
    Set miRsAux = Nothing
    Close (NFic)
    NFic = -1
    If Regs > 0 Then GeneraFicheroNorma34SEPA_XML = True
    Exit Function
EGen3:
    MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
    If NFic > 0 Then Close (NFic)
End Function





Private Function XML(Cadena As String) As String
Dim I As Integer
Dim Aux As String
Dim Le As String
Dim C As Integer
    'Carácter no permitido en XML  Representación ASCII
    '& (ampersand)          &amp;
    '< (menor que)          &lt;
    ' > (mayor que)         &gt;
    '“ (dobles comillas)    &quot;
    '' (apóstrofe)          &apos;
    
    'La ISO recomienda trabajar con los carcateres:
    'a b c d e f g h i j k l m n o p q r s t u v w x y z
    'A B C D E F G H I J K L M N O P Q R S T U V W X Y Z
    '0 1 2 3 4 5 6 7 8 9
    '/ - ? : ( ) . , ' +
    'Espacio
    Aux = ""
    For I = 1 To Len(Cadena)
        Le = Mid(Cadena, I, 1)
        C = Asc(Le)
        
        
        Select Case C
        Case 40 To 57
            'Caracteres permitidos y numeros
            
        Case 65 To 90
            'Letras mayusculas
            
        Case 97 To 122
            'Letras minusculas
            
        Case 32
            'espacio en balanco
            
        Case Else
            Le = " "
        End Select
        Aux = Aux & Le
    Next
    XML = Aux
End Function


Private Function IBAN_Destino() As String
Dim Aux As String

        IBAN_Destino = FrmtStr(DBLet(miRsAux!IBAN, "T"), 4) ' ES00
        IBAN_Destino = IBAN_Destino & Format(miRsAux!entidad, "0000") ' Código de entidad receptora
        IBAN_Destino = IBAN_Destino & Format(miRsAux!oficina, "0000") ' Código de oficina receptora
        IBAN_Destino = IBAN_Destino & Format(miRsAux!CC, "00") ' Dígitos de control
        IBAN_Destino = IBAN_Destino & Format(miRsAux!cuentaba, "0000000000") ' Código de cuenta
        
        
        
        
        
End Function


Public Function FrmtStr(campo As String, longitud As Integer) As String
    FrmtStr = Mid(Trim(campo) & Space(longitud), 1, longitud)
End Function

