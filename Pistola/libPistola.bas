Attribute VB_Name = "libPistola"
Option Explicit

Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

Public Const NumeroDeDecimales = 2
Public vUsu As Usuario
Public conAri As Byte
Public Log As cLOG
Public ConnConta As Connection  'De momento siempre NOTHING
Public vParamAplic As CParamAplic
Public Const ValorNulo = "Null"
Public BDConta As String  'Para el iva...
Public ContadorActualizaciones As Integer   'Para lleva un conteo de cuatas articulos actualizado hoy

'Formato de fecha
Public FormatoFecha As String
Public FormatoFechaHora As String

    '#,##0.00
Public FormatoImporte As String 'Decimal(12,2)
Public FormatoPrecio As String 'Decimal(10,4)txtAnterior
Public FormatoCantidad As String 'Decimal(10,2)
Public FormatoCantidad2 As String 'Decimal(8,2)
Public FormatoDescuento As String 'Decimal(4,2)
Public FormatoKms As String 'Decimal(8,4)
Public FormatoPorcen As String 'Decimal(5,2)

Public CadenaDesdeOtroForm As String

Public UltimaEntradaEnHcoDepositos As Date


Public NombrePistola As String   'Llevara una especie de numeracion las pistolas. De momento no


'Conexión a la BD Ariges de la empresa
Public Conn As ADODB.Connection
Public MiRsAux As ADODB.Recordset




Public Sub Main()

'    If Dir(App.Path & "\aqui.txt", vbArchive) = "" Then
'        cad = InputBox("Base datos nº: ", "Conexion", "1")
'    Else
'        cad = InputBox("Base datos nº: ", "Conexion", "9")
'    End If

    If Not AbrirConexion("") Then End
    
    OtrasAcciones
    
    NombrePistola = "P1"
    
    Set vUsu = New Usuario
    frmLogin.Show
End Sub

Private Sub OtrasAcciones()
On Error Resume Next

    FormatoFecha = "yyyy-mm-dd"
    FormatoFechaHora = "yyyy-mm-dd hh:mm:ss"
    FormatoImporte = "#,###,###,##0.00"  'Decimal(12,2)
    FormatoPrecio = "###,##0.0000"  'Decimal(10,4)
    
    'Por si acasomcambaimos la aplicacion los numeros de decimales
    'ANTES
    'FormatoCantidad = "##,###,##0.00"   'Decimal(10,2)
    'FormatoCantidad2 = "###,##0.00"   'Decimal(8,2)
    'Ahora
    FormatoCantidad = "##,###,##0." & String(NumeroDeDecimales, "0")
    FormatoCantidad2 = "###,##0." & String(NumeroDeDecimales, "0")
    
    FormatoDescuento = "#0.00" 'Decima(4,2)
    FormatoKms = "#,##0.00##" 'Decimal(8,4)
    FormatoPorcen = "##0.00" 'Decima(5,2)
    
    UltimaEntradaEnHcoDepositos = Now  'Lo necesitamos porque es una variable de ARIOLI
    
End Sub




Public Function AbrirConexion(Kariges As String) As Boolean
Dim cad As String
On Error GoTo EAbrirConexion

    
    AbrirConexion = False
    Set Conn = Nothing
    Set Conn = New Connection

    Conn.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente

    
    If Dir(App.Path & "\ODBC2.dat") <> "" Then
        'Aqui en nuestro servidor. El ODBC del arioli es distinto para no colisionar con el de ariges
        cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATA SOURCE= vAceite;"
    Else
        cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATA SOURCE= vAriges;"
    End If
    If Kariges <> "" Then cad = cad & "DATABASE=Ariges" & Kariges & ";"
   ' cad = cad & ";UID=root"
   ' cad = cad & ";PWD=aritel"
    cad = cad & ";Persist Security Info=true"
    
    Conn.ConnectionString = cad
    Conn.Open
    Conn.Execute "Set AUTOCOMMIT = 1"
    AbrirConexion = True
    Exit Function
    
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexión BD:Ariges.", Err.Description
End Function





Public Sub MuestraError(numero As Long, Optional Cadena As String, Optional Desc As String)
    Dim cad As String
    Dim Aux As String
    'Con este sub pretendemos unificar el msgbox para todos los errores
    'que se produzcan
    On Error Resume Next
    cad = "Se ha producido un error: " & vbCrLf
    If Cadena <> "" Then
        cad = cad & vbCrLf & Cadena & vbCrLf & vbCrLf
    End If
    Aux = ""

    If Aux <> "" Then Desc = Aux
    If Desc <> "" Then cad = cad & vbCrLf & Desc & vbCrLf & vbCrLf
    If Aux = "" Then cad = cad & "Número: " & numero & vbCrLf & "Descripción: " & Error(numero)
    MsgBox cad, vbExclamation
End Sub


Public Function Espera(Segundos As Single)
Dim T1
    T1 = Timer
    Do
    Loop Until Timer - T1 > Segundos
End Function





'recupera valor desde una cadena con pipes(acabada en pipes)
'Para ello le decimos el orden  y ya ta
Public Function RecuperaValor(ByRef Cadena As String, Orden As Integer) As String
Dim i As Integer
Dim J As Integer
Dim cont As Integer
Dim cad As String

    i = 0
    cont = 1
    cad = ""
    Do
        J = i + 1
        i = InStr(J, Cadena, "|")
        If i > 0 Then
            If cont = Orden Then
                cad = Mid(Cadena, J, i - J)
                i = Len(Cadena) 'Para salir del bucle
                Else
                    cont = cont + 1
            End If
        End If
    Loop Until i = 0
    RecuperaValor = cad
End Function


'Cambia los puntos de los numeros decimales
'por comas
Public Function TransformaPuntosComas(Cadena As String) As String
    Dim i As Integer
    Do
        i = InStr(1, Cadena, ".")
        If i > 0 Then
            Cadena = Mid(Cadena, 1, i - 1) & "," & Mid(Cadena, i + 1)
        End If
        Loop Until i = 0
    TransformaPuntosComas = Cadena
End Function


'Cambia los puntos de los numeros decimales
'por comas
Public Function TransformaComasPuntos(Cadena As String) As String
Dim i As Integer
    Do
        i = InStr(1, Cadena, ",")
        If i > 0 Then
            Cadena = Mid(Cadena, 1, i - 1) & "." & Mid(Cadena, i + 1)
        End If
    Loop Until i = 0
    TransformaComasPuntos = Cadena
End Function






Public Function ImporteFormateado(Importe As String) As Currency
Dim i As Integer

    If Importe = "" Then
        ImporteFormateado = 0
    Else
        'Primero quitamos los puntos
        Do
            i = InStr(1, Importe, ".")
            If i > 0 Then Importe = Mid(Importe, 1, i - 1) & Mid(Importe, i + 1)
        Loop Until i = 0
        ImporteFormateado = Importe
    End If
End Function

Public Function EsFechaOKTex(ByRef T As TextBox) As Boolean
Dim Te As String
    
    
    EsFechaOKTex = False
    Te = T.Text
    If EsFechaOK(Te) Then
        T.Text = Te
        EsFechaOKTex = True
    Else
       T.Text = ""
    End If
End Function


Public Function EsFechaOK(T As String) As Boolean
Dim cad As String
Dim mes As String, dia As String
    
    cad = T
    If InStr(1, cad, "/") = 0 Then
       'debe ser una cadena tipo:020105 y la convertimos a 02/01/05
       If Not IsNumeric(cad) Then
            EsFechaOK = False
            Exit Function
       End If
        
      '==== Anade: Laura 04/02/2005 =============
        If Len(cad) < 6 Then
            EsFechaOK = False
            Exit Function
        End If
        
        'Comprobar que el dia es correcto, valores entre 1-31
        dia = Mid(cad, 1, 2)
        If dia < 1 Or dia > 31 Then
            EsFechaOK = False
            Exit Function
        End If
        
        'Comprobar que el mes es correcto, valores entre 1-12
        mes = Mid(cad, 3, 2)
        If mes < 1 Or mes > 12 Then
            EsFechaOK = False
            Exit Function
        End If
      '============================================
        
        If Len(T) = 8 Then
            cad = Mid(cad, 1, 2) & "/" & Mid(cad, 3, 2) & "/" & Mid(cad, 5)
        Else
            If Len(T) = 6 Then cad = Mid(cad, 1, 2) & "/" & Mid(cad, 3, 2) & "/" & Mid(cad, 5)
        End If
    Else
        dia = Mid(cad, 1, 2)
        mes = Mid(cad, 4, 2)
    End If
    
    If IsDate(cad) Then
        EsFechaOK = True
        T = Format(cad, "dd/mm/yyyy")
      '==== Añade: Laura 08/02/2005
        If Month(T) <> Val(mes) Then EsFechaOK = False
        If Day(T) <> Val(dia) Then EsFechaOK = False
      '====
    Else
        EsFechaOK = False
    End If
End Function


Public Sub PonerFoco(ByRef T As TextBox)
    On Error Resume Next
    T.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub ObtenerFoco(ByRef T As TextBox)
    On Error Resume Next
    T.SelStart = 0
    T.SelLength = Len(T.Text)
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Function DBLet(vData As Variant, Optional Tipo As String) As Variant
    If IsNull(vData) Then
        DBLet = ""
        If Tipo <> "" Then
            Select Case Tipo
                Case "T"    'Texto
                    DBLet = ""
                Case "N"    'Numero
                    DBLet = "0"
                Case "F"    'Fecha
                     '==David
'                    DBLet = "0:00:00"
                     '==Laura
'                     DBLet = "0000-00-00"
                      DBLet = ""
                Case "D"
                    DBLet = 0
                Case "B"  'Boolean
                    DBLet = False
                Case Else
                    DBLet = ""
            End Select
        End If
    Else
        DBLet = vData
    End If
End Function





Public Function DevuelveDesdeBD(NoUtilizado As Byte, kCampo As String, Ktabla As String, Kcodigo As String, ValorCodigo As String, Optional Tipo As String, Optional ByRef otroCampo As String) As String
    Dim RS As Recordset
    Dim cad As String
    Dim Aux As String
    
    On Error GoTo EDevuelveDesdeBD
    DevuelveDesdeBD = ""
    cad = "Select " & kCampo
    If otroCampo <> "" Then cad = cad & ", " & otroCampo
    cad = cad & " FROM " & Ktabla
    cad = cad & " WHERE " & Kcodigo & " = "
    If Tipo = "" Then Tipo = "N"
    Select Case Tipo
    Case "N"
        'No hacemos nada
        cad = cad & ValorCodigo
    Case "T", "F"
        cad = cad & "'" & ValorCodigo & "'"
    Case Else
        MsgBox "Tipo : " & Tipo & " no definido", vbExclamation
        Exit Function
    End Select
    
'    Debug.Print cad
    
    'Creamos el sql
    Set RS = New ADODB.Recordset
    

    RS.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    
    If Not RS.EOF Then
        
        DevuelveDesdeBD = DBLet(RS.Fields(0))
        If otroCampo <> "" Then otroCampo = DBLet(RS.Fields(1))
    End If
    RS.Close
    Set RS = Nothing
    Exit Function
EDevuelveDesdeBD:
        MuestraError Err.Number, "Devuelve DesdeBD.", Err.Description
End Function


Public Sub KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

Public Function QuitarCaracterNULL(vcad As String) As String
Dim i As Integer

    Do
        i = InStr(1, vcad, vbNullChar)
        If i > 0 Then 'Hay null
            vcad = Mid(vcad, 1, i - 1) & Mid(vcad, i + 2)
        End If
    Loop Until i = 0
    QuitarCaracterNULL = vcad
End Function



Public Function PonerFormatoDecimal(ByRef T As TextBox, tipoF As Single) As Boolean
'tipoF: tipo de Formato a aplicar
'  1 -> Decimal(12,2)
'  2 -> Decimal(10,4)
'  3 -> Decimal(10,2)    '¡FORMATO CANTIDAD
'  4 -> Decimal(4,2)
'  5 -> Decimal(8,4)
'  6 -> Decimal(8,2)
'  7 -> Decimal(5,2)
'
'
'  8 -> Lo que ponga en su TAG   NO ESTA
Dim Valor As Currency
Dim PEntera As Currency
Dim NoOK As Boolean
'Dim Tg As CTag
Dim FormatoTag As String

    PonerFormatoDecimal = False
    If T.Text = "" Then Exit Function
    NoOK = False
    With T
        'If Not EsNumerico(.Text) Then
        If Not IsNumeric(.Text) Then
'            .Text = ""
            PonerFoco T
        Else
            If InStr(1, .Text, ",") > 0 Then
                Valor = ImporteFormateado(.Text)
            Else
                Valor = CCur(TransformaPuntosComas(.Text))
            End If

            'Comprobar la longitud de la Parte Entera
            PEntera = Int(Valor)
            Select Case tipoF 'Comprobar longitud
                Case 1 'Decimal(12,2)
                    If Len(PEntera) > 10 Then
                        MsgBox "El valor no puede ser mayor de 9999999999,99", vbExclamation
                        NoOK = True
                    End If
                Case 2 'Decimal(10,4)
                    If Len(CStr(PEntera)) > 6 Then
                        MsgBox "El valor no puede ser mayor de 999999,9999", vbExclamation
                        NoOK = True
                    End If
                Case 3 'Decimal(10,2)
                    If Len(CStr(PEntera)) > 6 Then
                        MsgBox "El valor no puede ser mayor de 999999,99", vbExclamation
                        NoOK = True
                    End If
                Case 4 'Decimal(4,2)
                    If Len(CStr(PEntera)) > 2 Then
                        MsgBox "El valor no puede ser mayor de 99,99", vbExclamation
                        NoOK = True
                    End If
                Case 5 'Decimal(8,4)
                    If Len(CStr(PEntera)) > 4 Then
                        MsgBox "El valor no puede ser mayor de 9999,9999", vbExclamation
                        NoOK = True
                    End If
                Case 6 'Decimal(8,2)
                    If Len(CStr(PEntera)) > 6 Then
                        MsgBox "El valor no puede ser mayor de 999999,99", vbExclamation
                        NoOK = True
                    End If
                Case 7 'Decimal(5,2)
                    '---- Laura: 05/10/2006
                    '# ANTES:   If Len(CStr(PEntera)) > 3 Then
                    If Len(CStr(Abs(PEntera))) > 3 Then
                    '----
                        MsgBox "El valor no puede ser mayor de 100,00", vbExclamation
                        NoOK = True
                    End If
                    
                Case 8
                    'David 12 Feb 07
                    'Lo que ponga en su tag
'                    Set Tg = New CTag
'                    If Not Tg.Cargar(T) Then NoOK = True
'                    FormatoTag = Tg.Formato
'                    Set Tg = Nothing
            End Select
            
            If NoOK Then
                .Text = ""
                T.SetFocus
                PonerFormatoDecimal = False
                Exit Function
            Else
                PonerFormatoDecimal = True
            End If

            'Poner el Formato
            Select Case tipoF
                Case 1 'Formato Decimal(12,2)
                    .Text = Format(Valor, FormatoImporte)
                Case 2 'Formato Decimal(10,4)
                    .Text = Format(Valor, FormatoPrecio)
                Case 3 'Formato Decimal(10,2)
                    .Text = Format(Valor, FormatoCantidad)
                Case 4 'Formato Decimal(4,2)
                    .Text = Format(Valor, FormatoDescuento)
                Case 5 'Formato Decimal(8,4)
                    .Text = Format(Valor, FormatoKms)
                Case 6 'Formato Decimal(8,2)
                    .Text = Format(Valor, FormatoCantidad2)
                Case 7 'Formato Decimal(5,2)
                    .Text = Format(Valor, "##0.00")
                Case 8
                    .Text = Format(Valor, FormatoTag)
            End Select
        End If
    End With
End Function

Public Function BloqueoManual(cadTabla As String, cadWhere As String, Optional OcultarMSG As Boolean) As Boolean
Dim Aux As String

On Error GoTo EBLOQ
    BloqueoManual = False
    If cadWhere = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
    Else
        Aux = "INSERT INTO zbloqueos(codusu,tabla,clave) VALUES(" & vUsu.Codigo & ",'" & cadTabla
        Aux = Aux & "',""" & cadWhere & """)"
        Conn.Execute Aux
        BloqueoManual = True
    End If
EBLOQ:
    If Err.Number <> 0 Then
        Aux = ""
        If Conn.Errors.Count > 0 Then
            If Conn.Errors(0).NativeError = 1062 Then
                '¡Ya existe el registro, luego esta bloqueada
                Aux = "BLOQUEO"
            End If
        End If
        
        If Aux = "" Then
            MuestraError Err.Number, "Bloqueo tabla"
        Else
            If Not OcultarMSG Then MsgBox "Registro bloqueado por otro usuario", vbExclamation
        End If
    End If
'    Screen.MousePointer = AntiguoCursor
End Function


Public Function DesBloqueoManual(cadTabla As String) As Boolean
Dim SQL As String

'Solo me interesa la tabla
On Error Resume Next

        SQL = "DELETE FROM zbloqueos WHERE codusu=" & vUsu.Codigo & " and tabla='" & cadTabla & "'"
        Conn.Execute SQL
        If Err.Number <> 0 Then
            Err.Clear
        End If
End Function



Public Function DBSet(vData As Variant, Tipo As String, Optional esNULO As String) As Variant
'Establece el valor del dato correcto antes de Insertar en la BD
'Tipos
'       T
'       N
'       F
'       H
'       FH
'       B
'       S   single O DOUBLE. sINGLE DE MOMENTO.    MAYO 2009
Dim cad As String
Dim ValorNumericoCero As Boolean

    On Error GoTo Error1

        If IsNull(vData) Then
            DBSet = "NULL"
            Exit Function
        End If
    
        If Tipo <> "" Then
            Select Case Tipo
                Case "T"    'Texto
                    If vData = "" Then
                        If esNULO = "N" Then
                            DBSet = "''"
                        Else
                            DBSet = "NULL"
                        End If
                    Else
                        cad = (CStr(vData))
                        NombreSQL cad
                        DBSet = "'" & cad & "'"
                    End If
                    
                Case "N", "S"   'Numero  y  SINGLE
                    
                    If CStr(vData) = "" Then
                        ValorNumericoCero = True
                    
                    Else
                        If Tipo = "S" Then
                            ValorNumericoCero = CSng(vData) = 0
                        Else
                            ValorNumericoCero = CCur(vData) = 0
                        End If
                    End If
                    
                    If ValorNumericoCero Then
                        If esNULO <> "" Then
                            If esNULO = "S" Then
                                DBSet = "NULL"
                            Else
                                DBSet = 0
                            End If
                        Else
                            DBSet = 0
                        End If
                    Else
                        If Tipo = "N" Then
                            cad = CStr(ImporteFormateado(CStr(vData)))
                        Else
                            'Sngle
'                            Cad = CStr(ImporteFormateadoSingle(CStr(vData)))
                           cad = CStr(ImporteFormateado(CStr(vData)))
                        End If
                        DBSet = TransformaComasPuntos(cad)
                    End If
                    
                Case "F"    'Fecha
'                     '==David
''                    DBLet = "0:00:00"
'                     '==Laura
                    If vData = "" Then
                        If esNULO = "S" Then
                            DBSet = "NULL"
                        Else
                            DBSet = "'1900-01-01'"
                        End If
                    Else
                        DBSet = "'" & Format(vData, FormatoFecha) & "'"
                    End If

                Case "FH" 'Fecha/Hora
                    If vData = "" Then
                        If esNULO = "S" Then DBSet = "NULL"
                    Else
                        DBSet = "'" & Format(vData, "yyyy-mm-dd hh:mm:ss") & "'"
                    End If

                Case "H" 'Hora
                    If vData = "" Then
                    Else
                        DBSet = "'" & Format(vData, "hh:mm:ss") & "'"
                    End If
                
                Case "B"  'Boolean
                    If vData Then
                        DBSet = 1
                    Else
                        DBSet = 0
                    End If
            End Select
        End If
Error1:
    If Err.Number <> 0 Then MuestraError Err.Number, "Formato para la BD.", Err.Description
End Function



'Para los nombre que pueden tener ' . Para las comillas habra que hacer dentro otro INSTR
Public Sub NombreSQL(ByRef Cadena As String)
Dim J As Integer
Dim i As Integer
Dim Aux As String

    J = 1
    '-- (RAFA/ALZIRA) 07052006
    Do
        i = InStr(J, Cadena, "\")
        If i > 0 Then
            Aux = Mid(Cadena, 1, i - 1) & "\"
            Cadena = Aux & Mid(Cadena, i)
            J = i + 2
        End If
    Loop Until i = 0
    

    J = 1
    Do
        i = InStr(J, Cadena, "'")
        If i > 0 Then
            Aux = Mid(Cadena, 1, i - 1) & "\"
            Cadena = Aux & Mid(Cadena, i)
            J = i + 2
        End If
    Loop Until i = 0
    
End Sub



Public Function EjecutaSQL(vBD As Byte, ByRef vSQL As String, Optional VerError As Boolean) As Boolean
    On Error Resume Next
    
    
        Conn.Execute vSQL
 
    If Err.Number <> 0 Then
        If VerError Then MuestraError Err.Number, Err.Description
        Err.Clear
        EjecutaSQL = False
    Else
        EjecutaSQL = True
    End If
End Function


Public Sub limpiar(ByRef formulario As Form)
    Dim Control As Object

    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Control.Text = ""
        End If
    Next Control
End Sub



Public Sub ConseguirFoco(ByRef Text As TextBox, Modo As Byte, Optional cadkey As Integer)
'Acciones que se realizan en el evento:GotFocus de los TextBox:Text1
'en los formularios de Mantenimiento
On Error Resume Next

    If Modo = 5 Then Exit Sub
    
    If (Modo <> 0 And Modo <> 2) Then
        If Modo = 1 Then
            Text.BackColor = vbYellow  'Modo 1: Busqueda
        Else
            If Text.Locked Then 'si el control esta bloqueado pasamos el foco al sig. campo
                Text.BackColor = &H80000018 'amarillo claro
                 If cadkey = 0 Then cadkey = 40
                 KEYdown cadkey
                 Exit Sub
            Else
                Text.BackColor = vbWhite
            End If
        End If
        Text.SelStart = 0
        Text.SelLength = Len(Text.Text)
    End If
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub KEYdown(KeyCode As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
On Error Resume Next
    Select Case KeyCode
        Case 38 'Desplazamieto Fecha Hacia Arriba
            SendKeys "+{tab}"
        Case 40 'Desplazamiento Flecha Hacia Abajo
            SendKeys "{tab}"
    End Select
    If Err.Number <> 0 Then Err.Clear
End Sub




Public Sub KEYpressGnral(KeyAscii As Integer, Modo As Byte, cerrar As Boolean)
'IN: codigo keyascii tecleado, y modo en que esta el formulario
'OUT: si se tiene que cerrar el formulario o no
    cerrar = False
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then 'ESC
        If (Modo = 0 Or Modo = 2) Then cerrar = True
    End If
End Sub



Public Function PonerFormatoEntero(ByRef T As TextBox) As Boolean
'Comprueba que el valor del textbox es un entero y le pone el formato
Dim mTag As CTag
Dim cad As String
Dim Formato As String
On Error GoTo EPonerFormato

    If T.Text = "" Then Exit Function
    PonerFormatoEntero = True
    
    Set mTag = New CTag
    mTag.Cargar T
    If mTag.Cargado Then
       cad = mTag.Nombre 'descripcion del campo
       Formato = mTag.Formato
    End If
    Set mTag = Nothing

    If Not EsEnteroNew(T.Text) Then
        PonerFormatoEntero = False
        MsgBox "El campo " & cad & " tiene que ser un número entero.", vbExclamation
        PonerFoco T
    Else
         T.Text = Format(T.Text, Formato)
    End If
    
EPonerFormato:
    If Err.Number <> 0 Then Err.Clear
End Function



Public Function EsEnteroNew(Texto As String) As Boolean
Dim i As Integer
Dim C As Integer
Dim L As Integer
Dim res As Boolean

    res = True
    EsEnteroNew = False

    If Not IsNumeric(Texto) Then
        res = False
    Else
        'Vemos si ha puesto mas de un punto
        C = 0
        L = 1
        Do
            i = InStr(L, Texto, ".")
            If i > 0 Then
                L = i + 1
                C = C + 1
            End If
        Loop Until i = 0
        If C > 0 Then res = False
        
        'Si ha puesto mas de una coma y no tiene puntos
        If C = 0 Then
            L = 1
            Do
                i = InStr(L, Texto, ",")
                If i > 0 Then
                    L = i + 1
                    C = C + 1
                End If
            Loop Until i = 0
            If C > 0 Then res = False
        End If
    End If
    EsEnteroNew = res
End Function



Public Sub PonerFocoBtn(ByRef btn As CommandButton)
On Error Resume Next
    If btn.Visible Then btn.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Function DevNombreSQL(Cadena As String) As String
Dim J As Integer
Dim i As Integer
Dim Aux As String
    J = 1
    Do
        i = InStr(J, Cadena, "'")
        If i > 0 Then
            Aux = Mid(Cadena, 1, i - 1) & "\"
            Cadena = Aux & Mid(Cadena, i)
            J = i + 2
        End If
    Loop Until i = 0
    DevNombreSQL = Cadena
End Function



Public Function CalcularImporte(Cantidad As String, Precio As String, Dto1 As String, Dto2 As String, TipoDto As Byte) As String
'Calcula el Importe de una linea de Oferta, Pedido, Albaran, ...
'Importe=cantidad * precio - (descuentos)
'Si DtoProv=sprove.tipodtos, calcular Importe para Proveedores y obtener el tipo de descuento
'del campo sprove.tipodtos, si es para Clientes obtener el tipo de descuento del
'parametro spara1.tipodtos
'Tipo Descuento: 0=aditivo, 1=sobre resto
Dim vImp As Currency
Dim vDto1 As Currency, vDto2 As Currency
Dim vPre As Currency
On Error Resume Next

    'Como son de tipo string comprobar que si vale "" lo ponemos a 0
    Cantidad = ComprobarCero(Cantidad)
    vPre = ComprobarCero(Precio)
    Dto1 = ComprobarCero(Dto1)
    Dto2 = ComprobarCero(Dto2)
    
    vImp = CCur(Cantidad) * CCur(vPre)
    If TipoDto = 0 Then 'Dto Aditivo
        vDto1 = (CCur(Dto1) * vImp) / 100
        vDto2 = (CCur(Dto2) * vImp) / 100
        vImp = vImp - vDto1 - vDto2
    ElseIf TipoDto = 1 Then 'Sobre Resto
        vDto1 = (CCur(Dto1) * vImp) / 100
        vImp = vImp - vDto1
        vDto2 = (CCur(Dto2) * vImp) / 100
        vImp = vImp - vDto2
    End If
    '// Enero 2009.  Hacia mal el redondeo pq ahora cantidad lleva decimales
    '   Ponemos Round2
    vImp = Round2(vImp, 2)
    CalcularImporte = CStr(vImp)
End Function


Public Function ComprobarCero(Valor As String) As String
    If Valor = "" Then
        ComprobarCero = "0"
    Else
        ComprobarCero = Valor
    End If
End Function


Public Function Round2(Number As Variant, Optional NumDigitsAfterDecimals As Long) As Variant
Dim ent As Integer
Dim cad As String
  
  ' Comprobaciones
  If Not IsNumeric(Number) Then
    Err.Raise 13, "Round2", "Error de tipo. Ha de ser un número."
    Exit Function
  End If
  If NumDigitsAfterDecimals < 0 Then
    Err.Raise 0, "Round2", "NumDigitsAfterDecimals no puede ser negativo."
    Exit Function
  End If
  
  ' Redondeo.
  cad = "0"
  If NumDigitsAfterDecimals <> 0 Then cad = cad & "." & String(NumDigitsAfterDecimals, "0")
  Round2 = Format(Number, cad)
  
End Function



Public Function DevuelveDesdeBDNew(vBD As Byte, Ktabla As String, kCampo As String, Kcodigo1 As String, valorCodigo1 As String, Optional tipo1 As String, Optional ByRef otroCampo As String, Optional KCodigo2 As String, Optional ValorCodigo2 As String, Optional tipo2 As String, Optional KCodigo3 As String, Optional ValorCodigo3 As String, Optional tipo3 As String) As String
'IN: vBD --> Base de Datos a la que se accede
Dim RS As Recordset
Dim cad As String
Dim Aux As String
    
On Error GoTo EDevuelveDesdeBDnew
    DevuelveDesdeBDNew = ""
'    If valorCodigo1 = "" And ValorCodigo2 = "" Then Exit Function
    cad = "Select " & kCampo
    If otroCampo <> "" Then cad = cad & ", " & otroCampo
    cad = cad & " FROM " & Ktabla
    If Kcodigo1 <> "" Then
        cad = cad & " WHERE " & Kcodigo1 & " = "
        If tipo1 = "" Then tipo1 = "N"
    Select Case tipo1
        Case "N"
            'No hacemos nada
            cad = cad & Val(valorCodigo1)
        Case "T"
            cad = cad & DBSet(valorCodigo1, "T")
        Case "F"
            cad = cad & "'" & valorCodigo1 & "'"
        Case Else
            MsgBox "Tipo : " & tipo1 & " no definido", vbExclamation
            Exit Function
    End Select
    End If
    
    If KCodigo2 <> "" Then
        cad = cad & " AND " & KCodigo2 & " = "
        If tipo2 = "" Then tipo2 = "N"
        Select Case tipo2
        Case "N"
            'No hacemos nada
            If ValorCodigo2 = "" Then
                cad = cad & "-1"
            Else
                cad = cad & Val(ValorCodigo2)
            End If
        Case "T"
'            cad = cad & "'" & ValorCodigo2 & "'"
            cad = cad & DBSet(ValorCodigo2, "T")
        Case "F"
            cad = cad & "'" & Format(ValorCodigo2, FormatoFecha) & "'"
        Case Else
            MsgBox "Tipo : " & tipo2 & " no definido", vbExclamation
            Exit Function
        End Select
    End If
    
    If KCodigo3 <> "" Then
        cad = cad & " AND " & KCodigo3 & " = "
        If tipo3 = "" Then tipo3 = "N"
        Select Case tipo3
        Case "N"
            'No hacemos nada
            If ValorCodigo3 = "" Then
                cad = cad & "-1"
            Else
                cad = cad & Val(ValorCodigo3)
            End If
        Case "T"
            cad = cad & "'" & ValorCodigo3 & "'"
        Case "F"
            cad = cad & "'" & Format(ValorCodigo3, FormatoFecha) & "'"
        Case Else
            MsgBox "Tipo : " & tipo3 & " no definido", vbExclamation
            Exit Function
        End Select
    End If
    
    
    'Creamos el sql
    Set RS = New ADODB.Recordset
    
    If vBD = conAri Then 'BD 1: Ariges
        RS.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Else    'BD 2: Conta
        RS.Open cad, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
    End If
    
    If Not RS.EOF Then
        DevuelveDesdeBDNew = DBLet(RS.Fields(0))
        If otroCampo <> "" Then otroCampo = DBLet(RS.Fields(1))
    End If
    RS.Close
    Set RS = Nothing
    Exit Function
    
EDevuelveDesdeBDnew:
        MuestraError Err.Number, "Devuelve DesdeBD.", Err.Description
End Function



Public Function EntreFechas(FIni As String, FechaComp As String, FFin As String) As Boolean
Dim B As Boolean
    B = False
    If FIni <> "" And FFin <> "" Then
        If EsFechaIgualPosterior(FIni, FechaComp, False) And EsFechaIgualPosterior(FechaComp, FFin, False) Then
            B = True
        End If
    ElseIf FIni = "" And FFin <> "" Then
        If EsFechaIgualPosterior(FechaComp, FFin, False) Then
            B = True
        End If
    ElseIf FIni <> "" And FFin = "" Then
        If EsFechaIgualPosterior(FIni, FechaComp, False) Then
            B = True
        End If
    End If
    EntreFechas = B
End Function


Public Function EsFechaIgualPosterior(FIni As String, FFin As String, MError As Boolean, Optional Men As String) As Boolean
'Comprueba que la Fecha Fin es igual o posterior a la Fecha de Inicio
'Si se pasa un cadena Men, se muestra esta como Mensaje de Error
'(IN) -> FIni: fecha inicio
'(IN) -> FFin: fecha fin
'(IN) -> MError: mostrar mensaje de error si/no
'(IN) -> Men: cadena mensaje de error
'(OUT) -> true: FFin >= Fini

    On Error GoTo ErrFec

'    EsFechaIgualPosterior = True
    
    If Trim(FIni) <> "" And Trim(FFin) <> "" Then
        If CDate(FIni) > CDate(FFin) Then
            EsFechaIgualPosterior = False
            
            If MError Then 'mostrar error
                If Men <> "" Then
                    'mostrar mensaje especifico q pasamos como parametro
                    MsgBox Men, vbInformation
                Else
                    'mostrar mensaje general
                    MsgBox "La Fecha Fin debe ser igual o posterior a la Fecha Inicio", vbInformation
                End If
            End If
        Else
            EsFechaIgualPosterior = True
        End If
    Else
        EsFechaIgualPosterior = True
    End If
    
    Exit Function
    
ErrFec:
    MuestraError Err.Number, "", Err.Description
End Function




Public Sub BloquearTxt(ByRef Text As TextBox, B As Boolean, Optional EsContador As Boolean)
'Bloquea un control de tipo TextBox
'Si lo bloquea lo pone de color amarillo claro sino lo pone en color blanco (sino es contador)
'pero si es contador lo pone color azul claro
On Error Resume Next

    Text.Locked = B
    If Not B And Text.Enabled = False Then Text.Enabled = True
    If B Then
        If EsContador Then
            'Si Es un campo que se obtiene de un contador poner color azul
'            Text.BackColor = &H80000013 'Azul Claro
            Text.BackColor = &HFFFFC0   'Azul claro con vista
        Else
            Text.BackColor = &H80000018 'Amarillo Claro
        End If
    Else
        Text.BackColor = vbWhite
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub

