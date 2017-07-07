Attribute VB_Name = "bus"
Option Explicit

Public vUsu As Usuario  'Datos usuario
Public vEmpresa As Cempresa 'Los datos de la empresa
Public vParam As Cparametros  'Parametros Generales de la Empresa (nombre, direc.,...
Public vParamAplic As CParamAplic 'Parametros Aplicación
Public vConfig As Configuracion 'Parametros Configuracion

Public vParamTPV As CParamTPV 'Parametros para el TPV

'LOG de acciones relevantes
Public LOG As cLOG   'Se instancia , se ejecuta LOG.insertar y se elimina :LOG=nothing   Ver ejemplo borre facturas


Public Const NumeroDeDecimales = 2
Public Const SerieFraPro = "1"

'Formato de fecha
Public FormatoFecha As String
Public FormatoFechaHora As String

    '#,##0.00
Public FormatoImporte As String 'Decimal(12,2)
Public FormatoPrecio As String 'Decimal(10,4)
Public FormatoCantidad As String 'Decimal(10,2)
Public FormatoCantidad2 As String 'Decimal(8,2)
Public FormatoDescuento As String 'Decimal(4,2)
Public FormatoKms As String 'Decimal(8,4)
Public FormatoPorcen As String 'Decimal(5,2)

Public CadenaDesdeOtroForm As String


'Conexión a la BD Ariges de la empresa
Public conn As ADODB.Connection

'Conexión a la BD de Usuarios
Public ConnUsuarios As ADODB.Connection

'Conexión a la BD de Contabilidad
Public ConnConta As ADODB.Connection

'Que conexion a base de datos se va a utilizar
Public Const conAri As Byte = 1 'Si conAri entonces trabajaremos con conexion conn a la BD ARIGES
Public Const conConta As Byte = 2 'Si conConta entonces trabajaremos con conexion connConta a la BD CONTA


'Para las formas de pago.  David
Public Const vbFPTransferencia = 1
Public Const vbCrearNuevaCta = "### CREAR CTA CONTAB. ###"


'Global para nº de registro eliminado
Public NumRegElim  As Long

'Para algunos campos de texto suletos controlarlos
'Public miTag As CTag

'Variable para saber si se ha actualizado algun asiento
'Public AlgunAsientoActualizado As Boolean
'Public TieneIntegracionesPendientes As Boolean

'Public miRsAux As ADODB.Recordset

Public AnchoLogin As String  'Para fijar los anchos de columna



'Empresas AVAB    --> 30 Mayo 2011. Parametro que indica si es "exportadora" (tb llamada  avab)
Public EmprAVAB As Integer   'Para no tener que calcularlo cada vez
Public EmprMorales As Integer 'Solo se utilizara para cuando la empresa sea AVAB


Public MaxNumDepositos_ As Integer



'Inicio Aplicación
Public Sub Main()
Dim T1 As Single
        

        'Aqui dentro fijamos la variable EmpresaAVAB
        
       Load frmIdentifica
       CadenaDesdeOtroForm = ""
       
       'Necesitaremos el archivo arifon.dat
       frmIdentifica.Show vbModal
               
       If CadenaDesdeOtroForm = "" Then
            'NO se ha identificado
            Set conn = Nothing
            End
       End If
       
       
       EmprAVAB = -1 'Luego fiajara cual es la empresa
       CadenaDesdeOtroForm = ""
       frmLogin.Show vbModal
       
        If CadenaDesdeOtroForm = "" Then
            'No ha seleccionado ninguna empresa
            Set conn = Nothing
            End
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass

'        LeerEmpresa 'Carga los Datos de la empresa
        'Carga los Datos Básicos de la empresa
        LeerDatosEmpresa
        
        'Cerramos la conexion con BD: Usuarios
        conn.Close

        'Abre la conexión a BDatos:Ariges
        If AbrirConexion() = False Then
            MsgBox "La apliacación no puede continuar sin acceso a los datos. ", vbCritical
            End
        Else
            'Carga Parametros Generales y Contables de la empresa
            LeerParametros
        End If
                
        'Abrir conexión a la BDatos de Contabilidad para acceder a
        'Tablas: Cuentas, Tipos IVA
        If AbrirConexionConta(False) = False Then
            MsgBox "La apliacación no puede continuar sin acceso a los datos. ", vbCritical
            End
        End If
        
        vParamAplic.SII_FijarValores
        'Carga los Niveles de cuentas de Contabilidad de la empresa y las fechasINICIO y FIN
        LeerNivelesEmpresa
        
'        'Gestionar el nombre del PC para la asignacion de PC en el entorno de red
'        GestionaPC
        
        'Otras acciones
        OtrasAcciones
         
        frmppal.Show

'ANTES
'Exit Sub

'       Load frmInicio
'       frmInicio.Show
'       frmInicio.Refresh
'       T1 = Timer
'       Set vConfig = New Configuracion
'       If vConfig.Leer = 1 Then
'            vConfig.SERVER = InputBox("Servidor: ")
'            vConfig.User = InputBox("Usuario: ")
'            vConfig.password = InputBox("Password: ")
''            vConfig.Integraciones = InputBox("Path integraciones: ")
'            vConfig.Grabar
'            MsgBox "Reinicie la contabilidad", vbCritical
'            End
'            Exit Sub
'       End If
 
'
'        If AbrirConexion() = False Then
'            MsgBox "La apliacación no puede continuar sin acceso a los datos. ", vbCritical
'            End
'        End If
        
'        'La llave
'        Load frmLLave
'        If Not frmLLave.ActiveLock1.RegisteredUser Then
'            'No ESTA REGISTRADO
'            frmLLave.Show vbModal
'        Else
'            Unload frmLLave
'        End If
        
        
'
'        'Que se vea un momentito
'        T1 = Timer - T1
'        If T1 < 0.5 Then
'            T1 = 0.5 - T1
'            espera T1
'        End If
        
'        'Descargamos inicio
'        Unload frmInicio
'
'
'        CadenaDesdeOtroForm = ""
'        frmLogin.Show vbModal
'        If vUsu Is Nothing Then
'            'Esto significa que no se ha identifcado como usuario
'            'luego finaliza la aplicacion
'            End
'        End If

'        Screen.MousePointer = vbHourglass

        'Cerramos la conexion
'        Conn.Close

'
'        If AbrirConexion() = False Then
'            MsgBox "La apliacación no puede continuar sin acceso a los datos. ", vbCritical
'            End
'        End If
        
'        LeerEmpresaParametros
        
        'Gestionar el nombre del PC para la asignacion de PC en el entorno de red
'        GestionaPC
        
        'Otras acciones
'        OtrasAcciones
         
'        frmPpal.Show
End Sub


Public Function LeerDatosEmpresa()
 'Crea instancia de la clase Cempresa con los valores en
 'Tabla: ArigesEmpresa
 'BDatos: Usuarios
 
        Set vEmpresa = New Cempresa
        If vEmpresa.LeerDatos = 1 Then
            MsgBox "No se han podido cargar datos empresa (BD:usuarios). Debe configurar la aplicación.", vbExclamation
            Set vEmpresa = Nothing
        End If
        
End Function


Public Function LeerNivelesEmpresa()
 'Crea instancia de la clase Cempresa con los valores en
 'Tabla: Empresa
 'BDatos: Conta
 
        If vEmpresa.LeerNiveles = 1 Then
            MsgBox "No se han podido cargar los niveles de la contabilidad de la empresa. Debe configurar la aplicación.", vbExclamation
'            Set vEmpresa = Nothing
        End If
            
End Function


Public Function LeerParametros()
'Crea instancia de la clase CParametros con los valores en
'Tabla: sparam
'BDatos: Ariges
 Dim Devuelve As String
 
    'Parametros Generales
    Set vParam = New Cparametros
    If vParam.Leer() = 1 Then
        Devuelve = "No se han podido cargar los Parámetros Generales.(sparam)" & vbCrLf
        MsgBox Devuelve & " Debe configurar la aplicación.", vbExclamation
        Set vParam = Nothing
    End If
        
    'Parametros Aplicacion
    Set vParamAplic = New CParamAplic
    If vParamAplic.Leer() = 1 Then
        Devuelve = "No se han podido cargar los Parámetros de la Aplicación.(spara1)" & vbCrLf
        MsgBox Devuelve & "Debe configurar la aplicación.", vbExclamation
        Set vParamAplic = Nothing
    End If
                
    'Si
    If Not vParam Is Nothing Then
        If Not vParamAplic Is Nothing Then
            If EmprAVAB = -1 Then EmprAVAB = FijaEmpresaAvab   'Si es -1 es la primera vez
            
            
        End If
    End If
    
    MaxNumDepositos_ = 27
    If vParamAplic.QUE_EMPRESA = 4 Then MaxNumDepositos_ = 18
       
End Function


'/////////////////////////////////////////////////////////////////
'// Se trata de identificar el PC en la BD. Asi conseguiremos tener
'// los nombres de los PC para poder asignarles un codigo
'// UNa vez asignado el codigo  se lo sumaremos (x 1000) al codusu
'// con lo cual el usuario sera distinto( aunque sea con el mismo codigo de entrada)
'// dependiendo desde k PC trabaje

Public Sub GestionaPC()
Dim miRsAux As ADODB.Recordset




CadenaDesdeOtroForm = ComputerName
If CadenaDesdeOtroForm <> "" Then
    'conAri=1: conexion a BD Ariges
    FormatoFecha = DevuelveDesdeBD(conAri, "codpc", "usuarios.pcs", "nompc", CadenaDesdeOtroForm, "T")
    If FormatoFecha = "" Then
        NumRegElim = 0
        FormatoFecha = "Select max(codpc) from usuarios.pcs"
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open FormatoFecha, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then
            NumRegElim = DBLet(miRsAux.Fields(0), "N")
        End If
        miRsAux.Close
        Set miRsAux = Nothing
        NumRegElim = NumRegElim + 1
        If NumRegElim > 9999 Then
            MsgBox "Error en numero de PC's activos. Demasiados PC en BD. Llame a soporte técnico.", vbCritical
            End
        End If
        FormatoFecha = "INSERT INTO usuarios.pcs (codpc, nompc) VALUES (" & NumRegElim & ", '" & CadenaDesdeOtroForm & "')"
        conn.Execute FormatoFecha
    End If
    
    
    
    'If CadenaDesdeOtroForm = "PCDAVID" Then EmpresaAVAB = 10
    
    
    
    
End If
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
    
    'Borramos uno de los archivos temporales
    If Dir(App.Path & "\ErrActua.txt") <> "" Then Kill App.Path & "\ErrActua.txt"
    
    
    'Borramos tmp bloqueos
    'Borramos temporal
    CadenaDesdeOtroForm = OtrosPCsContraContabiliad
    NumRegElim = Len(CadenaDesdeOtroForm)
    If NumRegElim = 0 Then
        CadenaDesdeOtroForm = ""
    Else
        CadenaDesdeOtroForm = " WHERE codusu = " & vUsu.Codigo
    End If
    conn.Execute "Delete from zbloqueos " & CadenaDesdeOtroForm
    CadenaDesdeOtroForm = ""
    NumRegElim = 0
    
    'Trabajador que mete en ALMACEN B
    CadenaDesdeOtroForm = DevuelveDesdeBDNew(conAri, "straba", "presupuesto", "login", vUsu.Login, "T")
    vUsu.TrabajadorB = CadenaDesdeOtroForm = "1"  'Trabajador de almacen en B
    
    
    vUsu.FijarCodigoTrabajador
        
        
    
        
    
    CadenaDesdeOtroForm = ""
    NumRegElim = 0
    
End Sub





'Usuario As String, Pass As String --> Directamente el usuario
Public Function AbrirConexion() As Boolean
Dim Cad As String
On Error GoTo EAbrirConexion

    
    AbrirConexion = False
    Set conn = Nothing
    Set conn = New Connection
'    Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    conn.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente

'        cad = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=accUPVMED"
'        cad = cad & ";UID=" & Usuario
'        cad = cad & ";PWD=" & Pass
'        Conn.ConnectionString = cad
    
    'cad = "DSN=plannertours;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=plannertours;UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
    
    '---- Laura: 17/10/2006
    Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATA SOURCE= vAriges;DATABASE=" & vUsu.CadenaConexion
    Cad = Cad & ";UID=" & vConfig.User
    Cad = Cad & ";PWD=" & vConfig.password
    Cad = Cad & ";Persist Security Info=true"
    
    conn.ConnectionString = Cad
    conn.Open
    conn.Execute "Set AUTOCOMMIT = 1"
    AbrirConexion = True
    Exit Function
    
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexión BD:Ariges.", Err.Description
End Function





Public Function AbrirConexionUsuarios() As Boolean
Dim Cad As String
On Error GoTo EAbrirConexion


    AbrirConexionUsuarios = False
    Set conn = Nothing
    Set conn = New Connection
    'Conn.CursorLocation = adUseClient
    conn.CursorLocation = adUseServer
    'Cad = "DSN=vUsuarios;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=usuarios;"
    'Cad = Cad & "SERVER=" & vConfig.SERVER & ";UID=" & vConfig.User & ";PASSWORD=" & vConfig.password & ";PORT=3306;OPTION=3;STMT=;"

    Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=usuarios;SERVER=" & vConfig.SERVER

    Cad = Cad & ";UID=" & vConfig.User
    Cad = Cad & ";PWD=" & vConfig.password
    Cad = Cad & ";OPTION=3;STMT=;Persist Security Info=true"

    conn.ConnectionString = Cad
    conn.Open
    AbrirConexionUsuarios = True
    Exit Function
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexión usuarios.", Err.Description
End Function



Public Function AbrirConexionConta(ContabilidadEnB As Boolean) As Boolean
'Abre

Dim Cad As String
On Error GoTo EAbrirConexion

    
    AbrirConexionConta = False
    Set ConnConta = Nothing
    Set ConnConta = New Connection
'    Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    ConnConta.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente
                        
    If vParamAplic.ContabilidadNueva Then
        Cad = "ariconta"
    Else
        Cad = "conta"
    End If
    Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=" & Cad
    If ContabilidadEnB Then
        Cad = Cad & vParamAplic.ContabilidadB
    Else
        Cad = Cad & vParamAplic.NumeroConta
    End If
    If vParamAplic.ServidorConta = "" Then
        Cad = Cad & ";SERVER=" & vConfig.SERVER & ";"
    Else
        Cad = Cad & ";SERVER=" & vParamAplic.ServidorConta & ";"
    End If
    Cad = Cad & ";UID=" & vParamAplic.UsuarioConta
    
    Cad = Cad & ";PWD=" & vParamAplic.PasswordConta
    '---- Laura: 29/09/2006
    Cad = Cad & ";PORT=3306;OPTION=3;STMT="
    '----
    Cad = Cad & ";Persist Security Info=true"
    ConnConta.ConnectionString = Cad
    ConnConta.Open
    ConnConta.Execute "Set AUTOCOMMIT = 1"
    AbrirConexionConta = True
    Exit Function
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexión.", Err.Description
End Function



Public Function CerrarConexionConta()
  'Cerramos la conexion con BD: Contabilidad
  On Error Resume Next
   ConnConta.Close
   If Err.Number <> 0 Then Err.Clear
End Function




'Para las cosas que tengan que ver con aridoc
'Utilizaremos la conexion de conta
Public Function AbrirConexionAridoc() As Boolean
Dim Cad As String
On Error GoTo EAbrirConexion

    
    AbrirConexionAridoc = False
    Set ConnConta = Nothing
    Set ConnConta = New Connection
    ConnConta.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente

    Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATA SOURCE= Aridoc;DATABASE=Aridoc"
    'Cad = Cad & ";UID=" & vConfig.User
    'Cad = Cad & ";PWD=" & vConfig.password
    Cad = Cad & ";Persist Security Info=true"
    
    ConnConta.ConnectionString = Cad
    ConnConta.Open
    ConnConta.Execute "Set AUTOCOMMIT = 1"
    AbrirConexionAridoc = True
    Exit Function
    
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexión BD:Aridoc.", Err.Description
End Function



Public Function Conexion_Aridoc_(abrir As Boolean) As Boolean
Dim Bien As Boolean
    Conexion_Aridoc_ = False
    CerrarConexionConta
    If abrir Then
        Bien = AbrirConexionAridoc()
    Else
        'Reabrimos la conexion conta
        Bien = AbrirConexionConta(False)
    End If
    If Not Bien Then
        If Not abrir Then
            MsgBox "EL PRORGRAMA FINALIZARA", vbExclamation
            End
        End If
    Else
        Conexion_Aridoc_ = True
    End If
    
End Function














'/////////////////////////////////////////////////
'   Esto lo ejecutaremos justo antes de bloquear
'   Prepara la conexion para bloquear
Public Sub PreparaBloquear()
    conn.Execute "commit"
    conn.Execute "set autocommit=0"
End Sub

'/////////////////////////////////////////////////
'   Esto lo ejecutaremos justo despues de un bloque
'   Prepara la conexion para bloquear
Public Sub TerminaBloquear()
    conn.Execute "commit"
    conn.Execute "set autocommit=1"
End Sub


'Cambia los puntos de los numeros decimales
'por comas
Public Function TransformaPuntosComas(Cadena As String) As String
    Dim I As Integer
    Do
        I = InStr(1, Cadena, ".")
        If I > 0 Then
            Cadena = Mid(Cadena, 1, I - 1) & "," & Mid(Cadena, I + 1)
        End If
        Loop Until I = 0
    TransformaPuntosComas = Cadena
End Function


'Cambia los puntos de los numeros decimales
'por comas
Public Function TransformaComasPuntos(Cadena As String) As String
Dim I As Integer
    Do
        I = InStr(1, Cadena, ",")
        If I > 0 Then
            Cadena = Mid(Cadena, 1, I - 1) & "." & Mid(Cadena, I + 1)
        End If
    Loop Until I = 0
    TransformaComasPuntos = Cadena
End Function



'Cambia los puntos de los numeros decimales
'por comas
Public Function TransformaPuntosHoras(Cadena As String) As String
    Dim I As Integer
    Do
        I = InStr(1, Cadena, ".")
        If I > 0 Then
            Cadena = Mid(Cadena, 1, I - 1) & ":" & Mid(Cadena, I + 1)
        End If
    Loop Until I = 0
    TransformaPuntosHoras = Cadena
End Function


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



Public Function DBLetMemo(vData As Variant) As String
    On Error Resume Next
    
    DBLetMemo = vData
    
'    If IsNull(DBLetMemo) Then DBLetMemo = ""
    
    If Err.Number <> 0 Then
        Err.Clear
        DBLetMemo = ""
    End If
End Function

'11 Marzo 2009 (Cumple de Llum, por cierto. Dos años)
'
'Añadimos el tipo "S" para que el casting(moldeo) del tipo de datos numerico lo haga sobre un single
' no sobre un ccur
Public Function DBSet(vData As Variant, Tipo As String, Optional EsNulo As String) As Variant
'Establece el valor del dato correcto antes de Insertar en la BD
Dim Cad As String
Dim EsCero As Boolean
    On Error GoTo Error1

        If IsNull(vData) Then
            DBSet = ValorNulo
            Exit Function
        End If
    
        If Tipo <> "" Then
            Select Case Tipo
                Case "T"    'Texto
                    If vData = "" Then
                        If EsNulo = "N" Then
                            DBSet = "''"
                        Else
                            DBSet = ValorNulo
                        End If
                    Else
                        Cad = (CStr(vData))
                        NombreSQL Cad
                        DBSet = "'" & Cad & "'"
                    End If
                    
                Case "N", "S"   'Numero  "S"ingle
                    If CStr(vData) = "" Then
                        If EsNulo <> "" Then
                            If EsNulo = "S" Then
                                DBSet = ValorNulo
                            Else
                                DBSet = 0
                            End If
                        Else
                            DBSet = 0
                        End If
                        
                    Else
                        If Tipo = "S" Then
                            EsCero = CSng(vData) = 0
                        Else
                            EsCero = CCur(vData) = 0
                        End If
                        If EsCero Then
                            If EsNulo <> "" Then
                                If EsNulo = "S" Then
                                    DBSet = ValorNulo
                                Else
                                    DBSet = 0
                                End If
                            Else
                                DBSet = 0
                            End If
                        Else
                            If Tipo = "S" Then
                                Cad = CStr(ImporteFormateadoSingle(CStr(vData)))
                            Else
                                Cad = CStr(ImporteFormateado(CStr(vData)))
                            End If
                            DBSet = TransformaComasPuntos(Cad)
                        End If
                    End If
                    
                Case "F"    'Fecha
'                     '==David
''                    DBLet = "0:00:00"
'                     '==Laura
                    If vData = "" Then
                        If EsNulo = "S" Then
                            DBSet = ValorNulo
                        Else
                            DBSet = "'1900-01-01'"
                        End If
                    Else
                        DBSet = "'" & Format(vData, FormatoFecha) & "'"
                    End If

                Case "FH" 'Fecha/Hora
                    If vData = "" Then
                        If EsNulo = "S" Then DBSet = ValorNulo
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




'Public Function FechaCorrecta(vFecha As Date) As Byte
''--------------------------------------------------------
''   Dada una fecha dira si pertenece o no
''   al intervalo de fechas que maneja la apliacion
''   Resultados:
''       0 .- Año actual
''       1 .- Siguiente
''       2 .- Anterior al inicio
''       3 .- Posterior al fin
''--------------------------------------------------------
'    FechaCorrecta = 2
'    If vFecha >= vParam.fechaini Then
'        If vFecha <= vParam.fechafin Then
'            FechaCorrecta = 0
'        Else
'            'Compruebo si el año siguiente
'            If vFecha <= DateAdd("yyyy", 1, vParam.fechafin) Then
'                FechaCorrecta = 1
'            Else
'                FechaCorrecta = 3
'            End If
'        End If
'    End If
'End Function


Public Sub MuestraError(numero As Long, Optional Cadena As String, Optional Desc As String)
    Dim Cad As String
    Dim Aux As String
    'Con este sub pretendemos unificar el msgbox para todos los errores
    'que se produzcan
    On Error Resume Next
    Cad = "Se ha producido un error: " & vbCrLf
    If Cadena <> "" Then
        Cad = Cad & vbCrLf & Cadena & vbCrLf & vbCrLf
    End If
    'Numeros de errores que contolamos
    If conn.Errors.Count > 0 Then
        ControlamosError Aux
        conn.Errors.Clear
    Else
        Aux = ""
    End If
    If Aux <> "" Then Desc = Aux
    If Desc <> "" Then Cad = Cad & vbCrLf & Desc & vbCrLf & vbCrLf
    If Aux = "" Then Cad = Cad & "Número: " & numero & vbCrLf & "Descripción: " & Error(numero)
    MsgBox Cad, vbExclamation
End Sub


Public Function Espera(Segundos As Single)
Dim T1
    T1 = Timer
    Do
    Loop Until Timer - T1 > Segundos
End Function


Public Function RellenaCodigoCuenta(vCodigo As String) As String
'Rellena con ceros hasta poner una cuenta.
'Ejemplo: 43.1 --> 430000001
Dim I As Integer
Dim J As Integer
Dim cont As Integer
Dim Cad As String

    RellenaCodigoCuenta = vCodigo
    If Len(vCodigo) > vEmpresa.DigitosUltimoNivel Then Exit Function
    
    I = 0: cont = 0
    Do
        I = I + 1
        I = InStr(I, vCodigo, ".")
        If I > 0 Then
            If cont > 0 Then cont = 1000
            cont = cont + I
        End If
    Loop Until I = 0

    'Habia mas de un punto
    If cont > 1000 Or cont = 0 Then Exit Function

    'Cambiamos el punto por 0's  .-Utilizo la variable maximocaracteres, para no tener k definir mas
    I = Len(vCodigo) - 1 'el punto lo quito
    J = vEmpresa.DigitosUltimoNivel - I
    Cad = ""
    For I = 1 To J
        Cad = Cad & "0"
    Next I

    Cad = Mid(vCodigo, 1, cont - 1) & Cad
    Cad = Cad & Mid(vCodigo, cont + 1)
    RellenaCodigoCuenta = Cad
End Function



Public Function DevuelveDesdeBD(vBD As Byte, Kcampo As String, KTabla As String, Kcodigo As String, ValorCodigo As String, Optional Tipo As String, Optional ByRef otroCampo As String) As String
    Dim RS As Recordset
    Dim Cad As String
    Dim Aux As String
    
    On Error GoTo EDevuelveDesdeBD
    DevuelveDesdeBD = ""
    Cad = "Select " & Kcampo
    If otroCampo <> "" Then Cad = Cad & ", " & otroCampo
    Cad = Cad & " FROM " & KTabla
    Cad = Cad & " WHERE " & Kcodigo & " = "
    If Tipo = "" Then Tipo = "N"
    Select Case Tipo
    Case "N"
        'No hacemos nada
        Cad = Cad & ValorCodigo
    Case "T", "F"
        Cad = Cad & "'" & ValorCodigo & "'"
    Case Else
        MsgBox "Tipo : " & Tipo & " no definido", vbExclamation
        Exit Function
    End Select
    
'    Debug.Print cad
    
    'Creamos el sql
    Set RS = New ADODB.Recordset
    
    If vBD = 1 Then 'BD 1: Ariges
        RS.Open Cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Else    'BD 2: Conta
        RS.Open Cad, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
    End If
    
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


'Este metodo sustituye a DevuelveDesdeBD
'Funciona para claves primarias formadas por 2 campos
Public Function DevuelveDesdeBDNew(vBD As Byte, KTabla As String, Kcampo As String, Kcodigo1 As String, valorCodigo1 As String, Optional tipo1 As String, Optional ByRef otroCampo As String, Optional KCodigo2 As String, Optional ValorCodigo2 As String, Optional tipo2 As String, Optional KCodigo3 As String, Optional ValorCodigo3 As String, Optional tipo3 As String) As String
'IN: vBD --> Base de Datos a la que se accede
Dim RS As Recordset
Dim Cad As String
Dim Aux As String
    
On Error GoTo EDevuelveDesdeBDnew
    DevuelveDesdeBDNew = ""
'    If valorCodigo1 = "" And ValorCodigo2 = "" Then Exit Function
    Cad = "Select " & Kcampo
    If otroCampo <> "" Then Cad = Cad & ", " & otroCampo
    Cad = Cad & " FROM " & KTabla
    If Kcodigo1 <> "" Then
        Cad = Cad & " WHERE " & Kcodigo1 & " = "
        If tipo1 = "" Then tipo1 = "N"
    Select Case tipo1
        Case "N"
            'No hacemos nada
            Cad = Cad & Val(valorCodigo1)
        Case "T"
            Cad = Cad & DBSet(valorCodigo1, "T")
        Case "F"
            Cad = Cad & "'" & valorCodigo1 & "'"
        Case Else
            MsgBox "Tipo : " & tipo1 & " no definido", vbExclamation
            Exit Function
    End Select
    End If
    
    If KCodigo2 <> "" Then
        Cad = Cad & " AND " & KCodigo2 & " = "
        If tipo2 = "" Then tipo2 = "N"
        Select Case tipo2
        Case "N"
            'No hacemos nada
            If ValorCodigo2 = "" Then
                Cad = Cad & "-1"
            Else
                Cad = Cad & Val(ValorCodigo2)
            End If
        Case "T"
'            cad = cad & "'" & ValorCodigo2 & "'"
            Cad = Cad & DBSet(ValorCodigo2, "T")
        Case "F"
            Cad = Cad & "'" & Format(ValorCodigo2, FormatoFecha) & "'"
        Case Else
            MsgBox "Tipo : " & tipo2 & " no definido", vbExclamation
            Exit Function
        End Select
    End If
    
    If KCodigo3 <> "" Then
        Cad = Cad & " AND " & KCodigo3 & " = "
        If tipo3 = "" Then tipo3 = "N"
        Select Case tipo3
        Case "N"
            'No hacemos nada
            If ValorCodigo3 = "" Then
                Cad = Cad & "-1"
            Else
                Cad = Cad & Val(ValorCodigo3)
            End If
        Case "T"
            Cad = Cad & "'" & ValorCodigo3 & "'"
        Case "F"
            Cad = Cad & "'" & Format(ValorCodigo3, FormatoFecha) & "'"
        Case Else
            MsgBox "Tipo : " & tipo3 & " no definido", vbExclamation
            Exit Function
        End Select
    End If
    
    
    'Creamos el sql
    Set RS = New ADODB.Recordset
    
    If vBD = conAri Then 'BD 1: Ariges
        RS.Open Cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Else    'BD 2: Conta
        RS.Open Cad, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
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


Public Function EjecutaSQL(vBD As Byte, ByRef vSQL As String, Optional VerError As Boolean) As Boolean
    On Error Resume Next
    
    If vBD = conAri Then
        conn.Execute vSQL
    Else
        ConnConta.Execute vSQL
    End If
    If Err.Number <> 0 Then
        If VerError Then MuestraError Err.Number, Err.Description
        Err.Clear
        EjecutaSQL = False
    Else
        EjecutaSQL = True
    End If
End Function



'Obvio
Public Function EsCuentaUltimoNivel(Cuenta As String) As Boolean
    EsCuentaUltimoNivel = (Len(Cuenta) = vEmpresa.DigitosUltimoNivel)
End Function


Public Function CuentaCorrectaUltimoNivel(ByRef Cuenta As String, ByRef Devuelve As String) As Boolean
'Comprueba si es numerica
Dim SQL As String
Dim otroCampo As String

CuentaCorrectaUltimoNivel = False
If Cuenta = "" Then
    Devuelve = "Cuenta vacia"
    Exit Function
End If

If Not IsNumeric(Cuenta) Then
    Devuelve = "La cuenta debe de ser numérica: " & Cuenta
    Exit Function
End If

'Rellenamos si procede
Cuenta = RellenaCodigoCuenta(Cuenta)

'==========
If Not EsCuentaUltimoNivel(Cuenta) Then
    Devuelve = "No es cuenta de último nivel: " & Cuenta
    Exit Function
End If
'==================

otroCampo = "apudirec"
'BD 2: conexion a BD Conta
SQL = DevuelveDesdeBD(conConta, "nommacta", "cuentas", "codmacta", Cuenta, "T", otroCampo)
If SQL = "" Then
    Devuelve = "No existe la cuenta : " & Cuenta
    CuentaCorrectaUltimoNivel = True
    Exit Function
End If

'Llegados aqui, si que existe la cuenta
If otroCampo = "S" Then 'Si es apunte directo
    CuentaCorrectaUltimoNivel = True
    Devuelve = SQL
Else
    Devuelve = "No es apunte directo: " & Cuenta
End If

End Function

'-------------------------------------------------------------------------
'
'   Es la misma solo k no si no existe cuenta no da error
'Public Function CuentaCorrectaUltimoNivelSIN(ByRef Cuenta As String, ByRef devuelve As String) As Byte
''Comprueba si es numerica
'Dim SQL As String
'
'CuentaCorrectaUltimoNivelSIN = 0
'If Cuenta = "" Then
'    devuelve = "Cuenta vacia"
'    Exit Function
'End If
'If Not IsNumeric(Cuenta) Then
'    devuelve = "La cuenta debe de ser numérica: " & Cuenta
'    Exit Function
'End If
'
''Rellenamos si procede
'Cuenta = RellenaCodigoCuenta(Cuenta)
'
'CuentaCorrectaUltimoNivelSIN = 1
'If Not EsCuentaUltimoNivel(Cuenta) Then
'    SQL = "No es cuenta de último nivel"
'Else
'    'BD 2: conexion a BD Conta
'    SQL = DevuelveDesdeBD(2, "nommacta", "cuentas", "codmacta", Cuenta, "T")
'    If SQL = "" Then
'        SQL = "No existe la cuenta  "
'    Else
'        CuentaCorrectaUltimoNivelSIN = 2
'    End If
'End If
'
''Llegados aqui, si que existe la cuenta
'devuelve = SQL
'End Function


'Devuelve, para un nivel determinado, cuantos digitos tienen las cuentas
' a ese nivel
'Public Function DigitosNivel(numnivel As Integer) As Integer
'    Select Case numnivel
'    Case 1
'        DigitosNivel = vEmpresa.numdigi1
'
'    Case 2
'        DigitosNivel = vEmpresa.numdigi2
'
'    Case 3
'        DigitosNivel = vEmpresa.numdigi3
'
'    Case 4
'        DigitosNivel = vEmpresa.numdigi4
'
'    Case 5
'        DigitosNivel = vEmpresa.numdigi5
'
'    Case 6
'        DigitosNivel = vEmpresa.numdigi6
'
'    Case 7
'        DigitosNivel = vEmpresa.numdigi7
'
'    Case 8
'        DigitosNivel = vEmpresa.numdigi8
'
'    Case 9
'        DigitosNivel = vEmpresa.numdigi9
'
'    Case 10
'        DigitosNivel = vEmpresa.numdigi10
'
'    Case Else
'        DigitosNivel = -1
'    End Select
'End Function


'Public Function NivelCuenta(CodigoCuenta As String) As Integer
'Dim lon As Integer
'Dim niv As Integer
'Dim I As Integer
'    NivelCuenta = -1
'    lon = Len(CodigoCuenta)
'    I = 0
'    Do
'       I = I + 1
'       niv = DigitosNivel(I)
'       If niv > 0 Then
'            If niv = lon Then
'                NivelCuenta = I
'                I = 11 'para salir del bucle
'            End If
'        Else
'            I = 11 'salimos pq ya no hay nveles para las cuentas de longitud lon
'        End If
'    Loop Until I > 10
'End Function


'Public Function ExistenSubcuentas(ByRef Cuenta As String, Nivel As Integer) As Boolean
'Dim I As Integer
'Dim b As Boolean
'Dim Cad As String
'
'    I = DigitosNivel(Nivel)
'    Cad = Mid(Cuenta, 1, I)
'    Cad = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Cad, "T")
'    If Cad = "" Then
'        'NO existe la subcuenta de nivel N
'        'salimos
'        ExistenSubcuentas = False
'        Exit Function
'    End If
'    If Nivel > 1 Then
'        ExistenSubcuentas = ExistenSubcuentas(Cuenta, Nivel - 1)
'    Else
'        ExistenSubcuentas = True
'    End If
'End Function


'Public Function CreaSubcuentas(ByRef Cuenta, HastaNivel As Integer, TEXTO As String) As Boolean
'Dim I As Integer
'Dim J As Integer
'Dim Cad As String
'Dim Cta As String
'
'On Error GoTo ECreaSubcuentas
'CreaSubcuentas = False
'For I = 1 To HastaNivel
'    J = DigitosNivel(I)
'    Cta = Mid(Cuenta, 1, J)
'    Cad = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Cta, "T")
'    If Cad = "" Then
'        'CreaCuenta
'        Cad = "INSERT INTO cuentas (codmacta, nommacta, apudirec, model347, razosoci, "
'        Cad = Cad & " dirdatos, codposta, despobla, desprovi, nifdatos, maidatos, webdatos,"
'        Cad = Cad & " obsdatos) VALUES ("
'        Cad = Cad & " '" & Cta
'        Cad = Cad & " ', '" & TEXTO
'        Cad = Cad & " ', "
'        Cad = Cad & " 'N', 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL)"
'        Conn.Execute Cad
'    End If
'Next I
'CreaSubcuentas = True
'Exit Function
'ECreaSubcuentas:
'    MuestraError Err.Number, "Creando subcuentas", Err.Description
'End Function




Public Function CambiarBarrasPATH2(ParaGuardarBD As Boolean, Cadena) As String
Dim I As Integer
Dim Ch As String
Dim Ch2 As String

If ParaGuardarBD Then
    Ch = "\"
    Ch2 = "/"
Else
    Ch = "/"
    Ch2 = "\"
End If
I = 0
Do
    I = I + 1
    I = InStr(1, Cadena, Ch)
    If I > 0 Then Cadena = Mid(Cadena, 1, I - 1) & Ch2 & Mid(Cadena, I + 1)
Loop Until I = 0
CambiarBarrasPATH2 = Cadena
End Function


Public Function ImporteSinFormato(Cadena As String) As String
Dim I As Integer
    'Quitamos puntos
    Do
        I = InStr(1, Cadena, ".")
        If I > 0 Then Cadena = Mid(Cadena, 1, I - 1) & Mid(Cadena, I + 1)
    Loop Until I = 0
    ImporteSinFormato = TransformaPuntosComas(Cadena)
End Function




'Public Sub SaldoHistorico(Cuenta As String)
'Dim RS As Recordset
'Dim SQL As String
'Dim RC2 As String
'    Screen.MousePointer = vbHourglass
'    SQL = "Select Sum(timporteD),sum(timporteH) from hlinapu"
'    SQL = SQL & " WHERE codmacta='" & Cuenta & "'"
'    SQL = SQL & " AND fechaent>='" & Format(vParam.fechaini, FormatoFecha) & "' AND punteada "
'    Set RS = New ADODB.Recordset
'    RC2 = Cuenta & "|"
'    'PUNTEADO
'    RS.Open SQL & "='S';", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'    If Not RS.EOF Then
'       RC2 = RC2 & Format(RS.Fields(0), FormatoImporte) & "|"
'       RC2 = RC2 & Format(RS.Fields(1), FormatoImporte) & "|"
'    Else
'        RC2 = RC2 & "||"
'    End If
'    RS.Close
'    'SIN puntear
'    RS.Open SQL & "<>'S';", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'    If Not RS.EOF Then
'       RC2 = RC2 & Format(RS.Fields(0), FormatoImporte) & "|"
'       RC2 = RC2 & Format(RS.Fields(1), FormatoImporte) & "|"
'    Else
'        RC2 = RC2 & "||"
'    End If
'    RS.Close
'    Set RS = Nothing
'    'Mostramos la ventanita de mesaje
'    frmMensajes.Opcion = 1
'    frmMensajes.Parametros = RC2
'    frmMensajes.Show vbModal
'
'End Sub

'Lo que hace es comprobar que si la resolucion es mayor
'que 800x600 lo pone en el 400
Public Sub AjustarPantalla(ByRef formulario As Form)
'    If Screen.Width > 13000 Then
'        formulario.Top = 400
'        formulario.Left = 400
'    Else
'        formulario.Top = 0
'        formulario.Left = 0
'    End If
'    formulario.Width = 12000
'    formulario.Height = 9000
End Sub


'///////////////////////////////////////////////////////////////
'
'   Cogemos un numero formateado: 1.256.256,98  y deevolvemos 1256256,98
'   Tiene que venir numérico
Public Function ImporteFormateado(Importe As String) As Currency
Dim I As Integer

    If Importe = "" Then
        ImporteFormateado = 0
    Else
        'Primero quitamos los puntos
        Do
            I = InStr(1, Importe, ".")
            If I > 0 Then Importe = Mid(Importe, 1, I - 1) & Mid(Importe, I + 1)
        Loop Until I = 0
        ImporteFormateado = Importe
    End If
End Function
Public Function ImporteFormateadoSingle(Importe As String) As Single
Dim I As Integer

    If Importe = "" Then
        ImporteFormateadoSingle = 0
    Else
        'Primero quitamos los puntos
        Do
            I = InStr(1, Importe, ".")
            If I > 0 Then Importe = Mid(Importe, 1, I - 1) & Mid(Importe, I + 1)
        Loop Until I = 0
        ImporteFormateadoSingle = Importe
    End If
End Function






Public Function ComprobarEmpresaBloqueada(CodUsu As Long, ByRef Empresa As String) As Boolean
'Dim cad As String
'Dim miRsAux As ADODB.Recordset
'
'ComprobarEmpresaBloqueada = False
'
''Antes de nada, borramos las entradas de usuario, por si hubiera kedado algo
'Conn.Execute "Delete from Usuarios.vBloqBD where codusu=" & CodUsu
'
''Ahora comprobamos k nadie bloquea la BD
''BD 1: conexion a BD Ariges
'cad = DevuelveDesdeBD(conAri, "codusu", "Usuarios.vBloqBD", "conta", Empresa, "T")
'If cad <> "" Then
'    'En teoria esta bloqueada. Puedo comprobar k no se haya kedado el bloqueo a medias
'
'    Set miRsAux = New ADODB.Recordset
'    cad = "show processlist"
'    miRsAux.Open cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
'    cad = ""
'    While Not miRsAux.EOF
'        If miRsAux.Fields(3) = Empresa Then
'            cad = miRsAux.Fields(2)
'            miRsAux.MoveLast
'        End If
'
'        'Siguiente
'        miRsAux.MoveNext
'    Wend
'    miRsAux.Close
'    Set miRsAux = Nothing
'
'    If cad = "" Then
'        'Nadie esta utilizando la aplicacion, luego se puede borrar la tabla
'        Conn.Execute "Delete from Usuarios.vBloqBD where conta ='" & Empresa & "'"
'
'    Else
'        MsgBox "BD bloqueada.", vbCritical
'        ComprobarEmpresaBloqueada = True
'    End If
'End If
'
'Conn.Execute "commit"
End Function


'Public Function Bloquear_DesbloquearBD(Bloquear As Boolean) As Boolean
'
'On Error GoTo EBLo
'    Bloquear_DesbloquearBD = False
'    If Bloquear Then
'        CadenaDesdeOtroForm = "INSERT INTO usuarios.vBloqBD (codusu, conta) VALUES (" & vUsu.Codigo & ",'" & vUsu.CadenaConexion & "')"
'    Else
'        CadenaDesdeOtroForm = "DELETE FROM  usuarios.vBloqBD WHERE codusu =" & vUsu.Codigo & " AND conta = '" & vUsu.CadenaConexion & "'"
'    End If
'    Conn.Execute CadenaDesdeOtroForm
'    Bloquear_DesbloquearBD = True
'    Exit Function
'EBLo:
'    'MuestraError Err.Number, "Bloq. BD"
'    Err.Clear
'End Function


Public Function OtrosPCsContraContabiliad() As String
Dim MiRS As Recordset
Dim Cad As String
Dim Equipo As String
Dim EquipoConBD As Boolean


    Set MiRS = New ADODB.Recordset
    EquipoConBD = (vUsu.PC = vConfig.SERVER Or LCase(vConfig.SERVER = "localhost"))
    Cad = "show processlist"
    MiRS.Open Cad, conn, adOpenKeyset, adLockOptimistic, adCmdText
    Cad = ""
    While Not MiRS.EOF
        If UCase(MiRS.Fields(3)) = UCase(vUsu.CadenaConexion) Then
            Equipo = MiRS.Fields(2)
            'Primero quitamos los dos puntos del puerot
            NumRegElim = InStr(1, Equipo, ":")
            If NumRegElim > 0 Then Equipo = Mid(Equipo, 1, NumRegElim - 1)
            
            'El punto del dominio
            NumRegElim = InStr(1, Equipo, ".")
            If NumRegElim > 0 Then Equipo = Mid(Equipo, 1, NumRegElim - 1)
            
            Equipo = UCase(Equipo)
            
            If Equipo <> vUsu.PC Then
                    
                    NumRegElim = 0
                    If Equipo <> "LOCALHOST" Then
                        'Si no es localhost
                        NumRegElim = 1
                    Else
                        'HAy un proceso de loclahost. Luego, si el equipo no tiene la BD
                        If Not EquipoConBD Then NumRegElim = 1
                    End If

                    
                    'Si hay que insertar
                    If NumRegElim = 1 Then
                        If InStr(1, Cad, Equipo & "|") = 0 Then Cad = Cad & Equipo & "|"
                    End If
            End If
        End If
        'Siguiente
        MiRS.MoveNext
    Wend
    NumRegElim = 0
    MiRS.Close
    Set MiRS = Nothing
    OtrosPCsContraContabiliad = Cad

End Function


Public Function EsNumerico(Texto As String) As Boolean
Dim I As Integer
Dim C As Integer
Dim L As Integer
Dim Cad As String
Dim b As Boolean
    
    EsNumerico = False
    b = True
    Cad = ""
    If Not IsNumeric(Texto) Then
        Cad = "El campo debe ser numérico"
        b = False
        '======= Añade Laura
        'formato: (.25)
        I = InStr(1, Texto, ".")
        If I = 1 Then
            If IsNumeric(Mid(Texto, 2, Len(Texto))) Then b = True
        End If
        '======================
    Else
        'Vemos si ha puesto mas de un punto
        C = 0
        L = 1
        Do
            I = InStr(L, Texto, ".")
            If I > 0 Then
                L = I + 1
                C = C + 1
            End If
        Loop Until I = 0
        If C > 1 Then
            'JUNIO 2011
            'Si ha puesto mas de un punto, pero HAY una coma por lo menos puede que este bien
            If InStr(1, Texto, ",") = 0 Then
                Cad = "Numero de puntos incorrecto"
                b = False
            End If
        End If
        'Si ha puesto mas de una coma y no tiene puntos
        If C = 0 Then
            L = 1
            Do
                I = InStr(L, Texto, ",")
                If I > 0 Then
                    L = I + 1
                    C = C + 1
                End If
            Loop Until I = 0
            If C > 1 Then
                Cad = "Numero incorrecto"
                b = False
            End If
        End If
    End If
    If Not b Then
        MsgBox Cad, vbExclamation
    Else
        EsNumerico = b
    End If
End Function







'==== Laura==
'Public Function EsPorcentajeOK(ByRef T As TextBox) As Boolean
'Dim cad As String
'Dim OK As Boolean
'
'    cad = TransformaPuntosComas(T.Text)
'
'    OK = False
'    If InStr(1, cad, ",") = 0 Then 'No hay decimales
'        If Len(T.Text) = 5 Then
'            cad = Mid(cad, 1, 2) & "," & Mid(cad, 3, 2)
'            OK = True
'        Else
'            If Len(T.Text) = 4 Then cad = Mid(cad, 1, 2) & "," & Mid(cad, 3, 2)
'            OK = True
'        End If
'    ElseIf InStr(1, cad, ",") = 1 Or InStr(1, cad, ",") = 2 Or InStr(1, cad, ",") = 3 Then 'Hay punto
'        OK = True
'    End If
'    If OK Then T.Text = cad
'    EsPorcentajeOK = OK
''    If IsDate(Cad) Then
''        EsFechaOK = True
''        T.Text = Format(Cad, "dd/mm/yyyy")
''    Else
''        EsFechaOK = False
''    End If
'
'End Function
'============




'Devuelve si hay archivos
'                                                        Llevara la forma: 01, 02  para la empresa 1 o 2..
'Public Function BuscarIntegraciones(Errores As Boolean, Empresa As String) As Boolean
'Dim cad As String
'On Error GoTo Ebuscarintegraciones
'
'    BuscarIntegraciones = False
'    If vConfig.Integraciones = "" Then Exit Function
'
'    cad = vConfig.Integraciones
'    If Right(cad, 1) <> "\" Then cad = cad & "\"
'    If Dir(cad, vbDirectory) = "" Then
'        MsgBox "Carpeta de errores no encontrada: " & vConfig.Integraciones, vbExclamation
'        Exit Function
'    End If
'
'    If Errores Then
'        cad = vConfig.Integraciones & "\ERRORES"
'    Else
'        cad = vConfig.Integraciones & "\INTEGRA"
'    End If
'
'    'Facturas clientes
'    If Dir(cad & "\FRACLI\*.?" & Empresa) <> "" Then
'        BuscarIntegraciones = True
'        Exit Function
'    End If
'
'    'Facturas Proveedores
'    If Dir(cad & "\FRAPRO\*.?" & Empresa) <> "" Then
'        BuscarIntegraciones = True
'        Exit Function
'    End If
'
'    'Asientos al diario
'    If Dir(cad & "\ASIDIA\*.?" & Empresa) <> "" Then
'        BuscarIntegraciones = True
'        Exit Function
'    End If
'
'    'Asientos al historico
'    If Dir(cad & "\ASIHCO\*.?" & Empresa) <> "" Then
'        BuscarIntegraciones = True
'        Exit Function
'    End If
'
'    Exit Function
'Ebuscarintegraciones:
'    MuestraError Err.Number, Err.Description, "Buscar archivos integraciones" & vbCrLf
'End Function


'Para los nombre que pueden tener ' . Para las comillas habra que hacer dentro otro INSTR
Public Sub NombreSQL(ByRef Cadena As String)
Dim J As Integer
Dim I As Integer
Dim Aux As String

    J = 1
    '-- (RAFA/ALZIRA) 07052006
    Do
        I = InStr(J, Cadena, "\")
        If I > 0 Then
            Aux = Mid(Cadena, 1, I - 1) & "\"
            Cadena = Aux & Mid(Cadena, I)
            J = I + 2
        End If
    Loop Until I = 0
    

    J = 1
    Do
        I = InStr(J, Cadena, "'")
        If I > 0 Then
            Aux = Mid(Cadena, 1, I - 1) & "\"
            Cadena = Aux & Mid(Cadena, I)
            J = I + 2
        End If
    Loop Until I = 0
    
End Sub

Public Function DevNombreSQL(Cadena As String) As String
Dim J As Integer
Dim I As Integer
Dim Aux As String
    J = 1
    Do
        I = InStr(J, Cadena, "'")
        If I > 0 Then
            Aux = Mid(Cadena, 1, I - 1) & "\"
            Cadena = Aux & Mid(Cadena, I)
            J = I + 2
        End If
    Loop Until I = 0
    DevNombreSQL = Cadena
End Function



'Para los balnces
'Public Function FechaInicioIGUALinicioEjerecicio(FecIni As Date, EjerciciosCerrados1 As Boolean) As Byte
'Dim Fecha As Date
'Dim Salir As Boolean
'Dim I As Integer
'On Error GoTo EfechaInicioIGUALinicioEjerecicio
'
'    FechaInicioIGUALinicioEjerecicio = 1
'    If EjerciciosCerrados1 Then
'        I = -1 'En ejercicios cerrados empèzamos mirando un año por debajo fecini
'    Else
'        I = 1
'    End If
'    Fecha = DateAdd("yyyy", I, vParam.fechaini)
'    Salir = False
'    While Not Salir
'        If FecIni = Fecha Then
'            'Fecha inicio del listado contiene es fecha incio ejercicio
'            FechaInicioIGUALinicioEjerecicio = 0
'            Salir = True
'        Else
'            If FecIni < Fecha Then
'                Fecha = DateAdd("yyyy", -1, Fecha)
'            Else
'                Salir = True
'            End If
'        End If
'    Wend
'
'    Exit Function
'EfechaInicioIGUALinicioEjerecicio:
'    Err.Clear  'No tiene importancia
'End Function





'Public Function DevuelveDigitosNivelAnterior() As Integer
'Dim J As Integer
'    DevuelveDigitosNivelAnterior = 3
'    If vEmpresa Is Nothing Then Exit Function
'    If vEmpresa.numnivel < 2 Then Exit Function
'    J = vEmpresa.numnivel - 1
'    J = DigitosNivel(J)
'    If J < 3 Then J = 3
'    DevuelveDigitosNivelAnterior = J
'End Function



'--------------------------------------------------------------------------
' Los numeros vendran formateados o sin formatear, pero siempre viene texto
'
Public Function CadenaCurrency(Texto As String, ByRef Importe As Currency) As Boolean
Dim I As Integer
On Error GoTo ECadenaCurrency
    
    Importe = 0
    CadenaCurrency = False
    If Not IsNumeric(Texto) Then Exit Function
    I = InStr(1, Texto, ",")
    If I = 0 Then
        'Significa k el numero no esta  formateado y como mucho tiene punto
        Importe = CCur(TransformaPuntosComas(Texto))
    Else
        Importe = ImporteFormateado(Texto)
    End If
    
    CadenaCurrency = True
    
    Exit Function
ECadenaCurrency:
    Err.Clear
End Function


Public Sub CommitConexion()
On Error Resume Next
    conn.Execute "Commit"
    If Err.Number <> 0 Then Err.Clear
End Sub






'------------------------------------------------------------------
'   Comprobara si una daterminada fecha esta o no en los ejercicios
'   contables (actual y siguiente)
'   Dando un O: SI. Correcto. Ok
'            1: Inferior
'            2: Superior

Public Function EsFechaOKConta(Fecha As Date) As Boolean

    EsFechaOKConta = False
    If Fecha < vEmpresa.FechaIni Then
       MsgBox "Fecha anterior ejercicios", vbExclamation
       Exit Function
    End If
    
    If Fecha > DateAdd("yyyy", 1, vEmpresa.FechaFin) Then
        MsgBox "Fecha posterior ejercicios abiertos", vbExclamation
        Exit Function
    End If

    If vEmpresa.SiiTiene Then
        If Fecha >= vEmpresa.SiiFechaInicio Then
            If Fecha < DateAdd("d", -vEmpresa.SiiDiasPlazo, Now) Then
                If vUsu.Nivel = 0 Then
                    If MsgBox("Fecha fuera de plazo SII. ¿Continuar?", vbQuestion + vbYesNo) = vbYes Then EsFechaOKConta = True
                Else
                    MsgBox "Fecha fuera de plazo SII.", vbCritical
                End If
            Else
                EsFechaOKConta = True
            End If
        Else
            EsFechaOKConta = True
        End If
    Else
        EsFechaOKConta = True
    End If
        
    
    
    

End Function



'--------------------------------------------------------------------
'-------------------------------------------------------------------
'Para el envio de los mails
Public Function PrepararCarpetasEnvioMail() As Boolean
    On Error GoTo EPrepararCarpetasEnvioMail
    PrepararCarpetasEnvioMail = False

    If Dir(App.Path & "\temp", vbDirectory) = "" Then
        MkDir App.Path & "\temp"
    Else
        If Dir(App.Path & "\temp\*.*", vbArchive) <> "" Then Kill App.Path & "\temp\*.*"
    End If


    PrepararCarpetasEnvioMail = True
    Exit Function
EPrepararCarpetasEnvioMail:
    MuestraError Err.Number, "", "Preparar Carpetas Envio Mail "
End Function



Public Function TieneAvisosPendientes() As Boolean
Dim CW As String
Dim F As Date
    On Error GoTo ETieneAvisosPendientes
    TieneAvisosPendientes = False
    
    
    'Alabaranes clientes
    If vParamAplic.avialbcli > 0 Then
        DoEvents
        F = DateAdd("d", -vParamAplic.avialbcli, Now)
        CW = " fechaalb <= '" & Format(F, FormatoFecha) & "'"
        If HayRegParaInforme("scaalb", CW, True) Then
            'No hace falta que siga puesto que si que hay alertar
            TieneAvisosPendientes = True
            Exit Function
        End If
    End If
    
    
    'Albaranes proveedores
    If vParamAplic.avialbpro > 0 Then
        DoEvents
        F = DateAdd("d", -vParamAplic.avialbpro, Now)
        CW = " fechaalb <= '" & Format(F, FormatoFecha) & "'"
        If HayRegParaInforme("scaalp", CW, True) Then
            'No hace falta que siga puesto que si que hay alertar
            TieneAvisosPendientes = True
            Exit Function
        End If
    End If
    
    'Pedidos  cli
    '
    If vParamAplic.avipedcli > 0 Then
        DoEvents
        F = DateAdd("d", -vParamAplic.avipedcli, Now)
        CW = " fecpedcl <= '" & Format(F, FormatoFecha) & "'"
        If HayRegParaInforme("scaped", CW, True) Then
            'No hace falta que siga puesto que si que hay alertar
            TieneAvisosPendientes = True
            Exit Function
        End If
    End If
    
    
    
    'Pedidos  cli
    '
    If vParamAplic.avipedpro > 0 Then
        DoEvents
        F = DateAdd("d", -vParamAplic.avipedpro, Now)
        CW = " fecpedpr <= '" & Format(F, FormatoFecha) & "'"
        If HayRegParaInforme("scappr", CW, True) Then
            'No hace falta que siga puesto que si que hay alertar
            TieneAvisosPendientes = True
            Exit Function
        End If
    End If
    
    'Avisos clientes
        
    'Pedidos  cli
    '
    If vParamAplic.aviavisos > 0 Then
        DoEvents
        F = DateAdd("d", -vParamAplic.aviavisos, Now)
        CW = " fechaavi <= '" & Format(F, FormatoFecha) & "' and situacio =0"
        If HayRegParaInforme("scaavi", CW, True) Then
            'No hace falta que siga puesto que si que hay alertar
            TieneAvisosPendientes = True
            Exit Function
        End If
    End If
    
    
    
    
    'Para los mantenimientos esta masss jodido, la verdad
    If vParamAplic.avimanteni > 0 Then
        
        'Fecha a partir de la cual reclamar
        F = DateAdd("d", -vParamAplic.avimanteni, Now)
        
        'Mensual
        CW = " WHERE (tipopago = 0 And ulmesfac < " & Month(F) & ")"
        'Trimestral
        CW = CW & " OR (tipopago = 1 And ulmesfac < " & Month(F) - 3 & ")"
        'Semestral
        CW = CW & " OR (tipopago = 2 And ulmesfac < " & Month(F) - 6 & ")"
        'Anual
        CW = CW & " OR (tipopago = 3 And ulmesfac =0)"
        If HayRegParaInforme("scaman", CW, True) Then
            'No hace falta que siga puesto que si que hay alertar
            TieneAvisosPendientes = True
            Exit Function
        End If
    
    End If
    Exit Function
ETieneAvisosPendientes:
    MuestraError Err.Number, Err.Description
End Function





'--------------------  ELIMINAR ARTICULO
Public Function SePuedeEliminarArticulo(ByVal Articulo As String, ByRef L1 As Label) As String
On Error GoTo Salida
Dim SQL As String
Dim RS As ADODB.Recordset
Dim I As Integer
Dim C As String
Dim nt As Integer

    SePuedeEliminarArticulo = ""
    Set RS = New ADODB.Recordset
    Articulo = "'" & DevNombreSQL(Articulo) & "'"
    
    
    'Clientes
    DevuelveTablasBorre 0, C, SQL, nt
    For I = 1 To nt
        L1.Caption = RecuperaValor(SQL, I) & " (Clientes)"
        L1.Refresh
        If TieneDatosSQLCount(RS, "SELECT count(*) from " & RecuperaValor(C, I) & " where codartic = " & Articulo, 0) Then
            SePuedeEliminarArticulo = SePuedeEliminarArticulo & "    -" & L1.Caption & vbCrLf
            
        End If
    Next I
    If SePuedeEliminarArticulo <> "" Then SePuedeEliminarArticulo = SePuedeEliminarArticulo & vbCrLf & vbCrLf
    
    'Si llega aqui comprobamos en  proveedores
    'PROVEEDORES
    DevuelveTablasBorre 1, C, SQL, nt
    For I = 1 To nt
        L1.Caption = RecuperaValor(SQL, I) & " (Proveedores)"
        L1.Refresh
        If TieneDatosSQLCount(RS, "SELECT count(*) from " & RecuperaValor(C, I) & " where codartic = " & Articulo, 0) Then
            SePuedeEliminarArticulo = SePuedeEliminarArticulo & "    -" & L1.Caption & vbCrLf
        
        End If
    Next I
    If SePuedeEliminarArticulo <> "" Then SePuedeEliminarArticulo = SePuedeEliminarArticulo & vbCrLf & vbCrLf
    
    'Varios
    DevuelveTablasBorre 2, C, SQL, nt
    For I = 1 To nt
        L1.Caption = RecuperaValor(SQL, I) & " (Varios)"
        L1.Refresh
        If TieneDatosSQLCount(RS, "SELECT count(*) from " & RecuperaValor(C, I) & " where codartic = " & Articulo, 0) Then
            SePuedeEliminarArticulo = SePuedeEliminarArticulo & "    -" & L1.Caption & vbCrLf
            
        End If
    Next I
    
        
        
    
    
    
    
    
    
    
    
    
    
    
Salida:
    If Err.Number <> 0 Then
        SePuedeEliminarArticulo = "Error: " & Err.Description
        Err.Clear
    End If
End Function



Private Function TieneDatosSQLCount(ByRef RS As ADODB.Recordset, vSQL As String, IndexdelCount As Integer) As Boolean
    TieneDatosSQLCount = False
    RS.Open vSQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(IndexdelCount)) Then If RS.Fields(IndexdelCount) > 0 Then TieneDatosSQLCount = True
    End If
        
    RS.Close

End Function



Public Function EliminarArticulo(ByVal codartic As String, L1 As Label) As Boolean
Dim nt As Integer
Dim Tablas As String
Dim Dsc As String

    On Error GoTo EEliminarArticulo
    
    EliminarArticulo = False
    
    codartic = " WHERE codartic = '" & DevNombreSQL(codartic) & "'"
    
    'Borraremos de tablas que se inserta autmaticamente
    'Ejm: slistas, precios especiales......
    DevuelveTablasBorre 3, Tablas, Dsc, nt
    Do
        Debug.Print RecuperaValor(Tablas, nt)
        L1.Caption = RecuperaValor(Dsc, nt)
        L1.Refresh
        conn.Execute "DELETE FROM " & RecuperaValor(Tablas, nt) & codartic
        nt = nt - 1
    Loop Until nt = 0
    
    L1.Caption = "Ficha tecnica"
    L1.Refresh
    
    'BORRAMOS EN FICH TECNICA
    conn.Execute "DELETE FROM sarti4" & codartic
    
    'BORRAMOS EN img fichtec
    conn.Execute "DELETE FROM sfichtecdocs" & codartic
    
    
    'BORRAMOS EL ARTICULO
    L1.Caption = Mid(codartic, 19)
    L1.Refresh
    conn.Execute "DELETE FROM sartic " & codartic
    
    EliminarArticulo = True
    
    Exit Function
EEliminarArticulo:
    MuestraError Err.Number, Err.Description
End Function


'Opcion
'   0- Clientes
'   1- Proveedores
'   2- Varios
'   ---------
'   3.- Tabas que cuando eliminen el articulo tendre que borrar yo
Public Sub DevuelveTablasBorre(Opcion As Byte, ByRef Tablas As String, ByRef Descripcion As String, ByRef NumeroTablas As Integer)

    If Opcion = 0 Then
        'CLIENTES
        Tablas = "slhalb|slhped|slhpre|slialb|slifac|sliordpr|sliped|slipre|sliven|slirep|"
        Descripcion = "Hco albaranes|Hco pedidos|Hco ofertas|Albaranes|Facturas|produccion|"
        Descripcion = Descripcion & "Pedidos|Ofertas|TPV|Reparaciones|"
        NumeroTablas = 10
    ElseIf Opcion = 1 Then
        'PROVEEDRORES
        Tablas = "slhalp|slhppr|slialp|slifpc|slippr|"
        Descripcion = "Hco albaranes|Hco pedidos|Albaranes|Facturas|Pedidos|"
        NumeroTablas = 5
        
        
    ElseIf Opcion = 2 Then
        'VARIOS
        Tablas = "slhmov|sarti2|slhtra|slimov|slitra|slotes|smoval|sserie|stipco|shinve|"
        Descripcion = "Hco Lineas Movimientos Almacen|Instalaciones|hco traspaso almacen|"
        Descripcion = Descripcion & "Lin mov almacen|Traspaso almacen|Nº lotes|Mov almacen|Nº serie|Tipos contrato|Hco inventario|"
        NumeroTablas = 10
        If vParamAplic.Produccion Then
            Tablas = Tablas & "sarti1|"
            Descripcion = Descripcion & "Artic. produccion|"
            NumeroTablas = NumeroTablas + 1
        End If
        
    Else
        'Tablas que al eliminar el articulo voy a tener que borrar
        'Esta salmac. Antes de lanzar el proceso hay que comprobar que la suma de stock es CERO
        Tablas = "slipla|slisp1|slispr|sbonif|slist1|slista|spree1|sprees|spromo|salmac|sarti5|"
        Descripcion = "Plantillas|Precios proveedor|cab. precios provee.|bonificacion facturas|"
        Descripcion = Descripcion & "Hco tarifas|Tarifas|Hco precios especiales|Precios especiales|Promociones|Articulos x Almacen|precios proveedor|"
        NumeroTablas = 11
        
    End If
    
End Sub





'------------------------------------------------------------------------------------------------
'
'       UpdateaPesoNeto:  Por si el UPDATE que hace la final tb tiene que updatear el pesonetoaceite. Por si acaso no lo teien en la ficha
Public Function RecalcularPesoArticulo(Articulo As String, Unicajas As Integer, CajasPalet As Integer, PesoNetoAceite As Currency, UpdateaPesoNeto As Boolean) As Boolean
Dim R As ADODB.Recordset
Dim SQL As String
Dim PesoTapon2 As Currency
Dim PesoBotella As Currency
Dim OtrosPesos As Currency
Dim PesoBrutoBotella As Currency
Dim Aux As Currency
Dim CajaVacia As Currency
Dim PesoBrutoCaja As Currency
Dim PesoNetoCaja As Currency
Dim PesoRetractil As Currency
Dim A As String

    On Error GoTo ERecalcularPesoArticulo

    PesoTapon2 = 0
    PesoBotella = 0
    CajaVacia = 0
    OtrosPesos = 0
    PesoRetractil = 0
    Set R = New ADODB.Recordset
    SQL = "select sarti4.*,tipartic,cantidad,nomartic from sarti4,sarti1,sartic where sarti4.codartic=sarti1.codarti1 and sartic.codartic = sarti4.codartic and sarti1.codartic='" & Articulo & "'"
    R.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not R.EOF
        Select Case Val(R!tipartic)
        Case 3
            'Modificacion Marzo 2012
            'Puede haber mas de un tapon
            PesoTapon2 = PesoTapon2 + DBLet(R!pesoneto, "N")
        Case 2
            PesoBotella = DBLet(R!pesoneto, "N")
           
        Case 8
            'retractil
            PesoRetractil = DBLet(R!pesoneto, "N")    'va por cada botella
        Case Else
            If Val(R!tipartic) = 6 Then
                'Caja
               CajaVacia = DBLet(R!caj_vacia, "N")
            Else
                If Val(R!tipartic) = 9 Then
                    'ESTUCHE
                    Aux = Unicajas
                    Aux = Aux * DBLet(R!pesoneto, "N")
                    OtrosPesos = OtrosPesos + Aux
                Else
                
                    If Val(R!tipartic) > 1 Then
                        Aux = 1  'R!cantidad
                        Aux = Aux * DBLet(R!pesoneto, "N")
                        OtrosPesos = OtrosPesos + Aux
                    End If
                End If
            End If
        End Select
        R.MoveNext
    Wend
    R.Close
    
    'If PesoTapon = 0 Or PesoBotella = 0 Then It.ForeColor = vbRed
    PesoBrutoBotella = PesoTapon2 + PesoBotella + PesoNetoAceite + PesoRetractil
    'It.SubItems(2) = Format(PesoNetoAceite, FormatoPrecio)
    'It.SubItems(3) = Format(PesoBrutoBotella, FormatoPrecio)
    
    
    'CAJA
    'Vamos a calcular la caja
    'neto caja
    PesoNetoCaja = PesoNetoAceite * Unicajas
    
    'bruto caja
    PesoBrutoCaja = (PesoBrutoBotella * Unicajas) + CajaVacia + OtrosPesos
    'It.SubItems(4) = Format(PesoNetoCaja, FormatoPrecio)
    'It.SubItems(5) = Format(PesoBrutoCaja, FormatoPrecio)
    
    
    
    
    '----------------------------------------------------------------------------
    'UPDATEAMOS en sarti4
    'Updaeatemos del codaric de venta peso bruto, peso bruto palet y pesoneto palet
    A = TransformaComasPuntos(CStr(PesoBrutoBotella))
    SQL = "UPDATE sarti4 SET pesobruto=" & A
    
    If UpdateaPesoNeto Then SQL = SQL & ", pesoneto =" & TransformaComasPuntos(CStr(PesoNetoAceite))
    
    'pal_pneto pal_pbruto
    A = TransformaComasPuntos(CStr(PesoNetoCaja * CajasPalet))
    SQL = SQL & ", pal_pneto =" & A
    A = TransformaComasPuntos(CStr(PesoBrutoCaja * CajasPalet))
    SQL = SQL & ", pal_pbruto =" & A
    
    'IMPORTANTE.Copiar esto en Arioli
    'Para el producto venta , los campos de sarti4  ret_medid ret_seriT
    'seran el peso brut y el peso neto de la caja
    A = Format(PesoNetoCaja, FormatoPrecio)
    A = "'" & A & " Kg'"
    SQL = SQL & ", ret_medid = " & A
    A = Format(PesoBrutoCaja, FormatoPrecio)
    A = "'" & A & " Kg'"
    SQL = SQL & ", ret_seriT =" & A
    
    SQL = SQL & " WHERE codartic = '" & Articulo & "'"
    
    conn.Execute SQL
    
    'Junio 2011
    If Not vParamAplic.EsAVAB Then
        If EmprAVAB > 0 Then
              A = Replace(SQL, "UPDATE sarti4 SET", "UPDATE ariges" & EmprAVAB & ".sarti4 SET")
              EjecutaSQL conAri, A, True
              
        End If
    End If
    
    
    
    RecalcularPesoArticulo = True
    
ERecalcularPesoArticulo:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set R = Nothing
End Function





Private Function FijaEmpresaAvab() As Integer
Dim RT As ADODB.Recordset
Dim Cad As String
Dim Aux2 As String

    'Aqui fijare tb la empresa MORALES.
    'Hay un campo en spara1 que nos dira el codmpresea morales
    
    EmprMorales = -1

    FijaEmpresaAvab = 0
    Set RT = New ADODB.Recordset
    'Cad = "Select * from usuarios.empresasarioli where codempre <> " & vEmpresa.codempre
    Cad = "Select * from usuarios.empresasarioli ORDER BY codempre"
    RT.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RT.EOF
        Aux2 = "CodEmpresaMorales"
        Cad = DevuelveDesdeBD(conAri, "EsEmpresaExportadora", RT!AriGes & ".spara1", "1", "1", "N", Aux2)
        'Si es empresa produccion
        If EmprMorales < 1 Then
            If Aux2 = "" Then Aux2 = "0"
            If Val(Aux2) > 0 Then EmprMorales = Val(Aux2)
        End If
        
        
        If Cad = "1" Then
            'OK esta es la empresa exportadora(o lo que es lo mismo, el AVAB
            FijaEmpresaAvab = RT!codempre
            While Not RT.EOF
                RT.MoveNext
            Wend
        Else
            'If EmprMorales < 0 Then EmprMorales = RT!codempre   FALTA### Morales, de momento siempre 1
            RT.MoveNext
        End If
        
    Wend
    RT.Close
    Set RT = Nothing

End Function




'Trozo copiado de pistola
    'If Len(CodigoCaja) <> 13 Then
        'MsgBox "Longitudad etiqueta incorrecta", vbExclamation
        
     
        'Dividimos la etiqueta leida en 2
        'los cinco ultimos son el IDCAJa
        'el resto ID trza
        'select * from prodcajas where lotetraza=27 and idcaja=5
        
Public Function LeerCaja(CodigoCaja As String) As String
Dim SQL As String
Dim RS As ADODB.Recordset

        LeerCaja = ""

        If Len(CodigoCaja) <> 13 Then
            LeerCaja = "Longitudad etiqueta incorrecta"
            Exit Function
        End If
                    
        If Not IsNumeric(CodigoCaja) Then
            LeerCaja = "Codigo caja NO numerico"
            Exit Function
        End If

        'Ahora veremos si pertence a una orden de carga
        Set RS = New ADODB.Recordset
        SQL = Mid(CodigoCaja, 1, 8) & " AND idcaja = " & Val(Mid(CodigoCaja, 9))
        SQL = "Select * from prodcajas WHERE lotetraza = " & SQL
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        If RS.EOF Then
            LeerCaja = "NO existe la caja en el sistema"
        Else
            'Vemos si lo que lee es lo que escribe ;
            'Si realmente la caja es de lo que me ha dicho en el albarna
            SQL = "select * from prodlin,prodtrazlin  where prodlin.codigo= prodtrazlin.codigo AND prodlin.idlin = prodtrazlin.idlin and lotetraza=" & Mid(CodigoCaja, 1, 8)
        End If
        RS.Close
        
        If SQL = "" Then Exit Function
        
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        If RS.EOF Then
            LeerCaja = "Error en lote de trazabilidad"
        Else
            'El codartic esta en nla primera linea del label, desde la posicion 4 hasta el ·
            'SQL = Mid(Label4(1).Caption, 1, InStr(Label4(1).Caption, "·") - 1)
            CodigoCaja = RS!Codigo & "|" & RS!idlin & "|"
            'SQL = "select * from srepartolot where idreparto=" & Label5(0).Tag & " and numalbar=" & Label5(1).Tag & " and codartic=" & SQL
        End If
        RS.Close
        
        
        
    





End Function


'Hay una tabla que se llama spermisos
'Para cada usuario y "proceso-formulario"
' que definamos tendremos si tiene permisos
'   Ejemplo.  Para el formulario frmproduccion vamos a
'   dar permisos a unos usuarios. Para ello meteremos
'   en la tabla los valores
'   usuario    accion     v1   v2  v3  v4  v5
'   ---------------------------------------------
'   1        frmprodou..   1    1   0   0   0    'Usuario1. Tiene permisos en v1 y v2 que dentro del form se define como lineas y planning
'   2           "          0    1   0   0   0    '       2   "              en v2 solo
'               el resto, al no estar definido NO tiene permisos
'
'Sere el vodigotrabajador
'
'       devolvera un char 010001 con los permisos para v1v2..5
Public Function TienePermiso(Accion As String, ByRef CadenaPermisos As String) As Boolean
Dim R As ADODB.Recordset
Dim I As Byte

    Set R = New ADODB.Recordset
    
    CadenaPermisos = ""
    If (vUsu.Codigo Mod 1000) = 0 Then
        TienePermiso = True
        Exit Function
    End If
    TienePermiso = False
    R.Open "Select valor1,valor2,valor3,valor4,valor5 from spermisos WHERE usuario = " & vUsu.CodigoTrabajador, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not R.EOF Then
        For I = 0 To 4
            CadenaPermisos = CadenaPermisos & DBLet(R.Fields(CInt(I)), "N")
        Next I
        If Val(CadenaPermisos) > 0 Then TienePermiso = True
    End If
    R.Close
    Set R = Nothing
    
End Function



Public Sub CargaComboTipoImpresionPalet(ByRef CBO As ComboBox)
    CBO.Clear
    CBO.AddItem "Normal"
    CBO.ItemData(CBO.NewIndex) = 0
    
    CBO.AddItem "Olive line"
    CBO.ItemData(CBO.NewIndex) = 1
End Sub



