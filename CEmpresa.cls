VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cempresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'EMPRESA
'TABLA:


'ATRIBUTOS
'Variables locales que contienen valores de propiedad
Private mvarCodempre As String 'copia local
Private mvarNomempre As String 'copia local
Private mvarNomresum As String 'copia local
Private mvarAriges As String

Private mvarnumnivel As Integer 'copia local
Private mvarnumdigi1 As Integer 'copia local
Private mvarnumdigi2 As Integer 'copia local
Private mvarnumdigi3 As Integer 'copia local
Private mvarnumdigi4 As Integer 'copia local
Private mvarnumdigi5 As Integer 'copia local
Private mvarnumdigi6 As Integer 'copia local
Private mvarnumdigi7 As Integer 'copia local
Private mvarnumdigi8 As Integer 'copia local
Private mvarnumdigi9 As Integer 'copia local
Private mvarnumdigi10 As Integer 'copia local



'VARIABLES
'Campo calculado de numero de digitos en la ULTIMA NIVEL CONTABLE
Private mvarDigitosUltimoNivel As Byte

'David
'-----------------------------------------------
'Para  no tener que leer "cienes y cienes" de veces
'las fechas de inicio y fin en la contabilidad
'las leo en esta primera lectura de datos
'Ya que si tiene esta tabla(empresa), tiene la de parametros
Private mvarFechaIni As Date
Private mvarFechaFin As Date
'Tambien a�adire si lleva analitica
Private mvarTieneAnalitica As Boolean

Private mvarSiiTiene As Boolean
Private mvarSiiFechaI As Date
Private mvarDiasSii As Integer




Private SQL As String  'Para la sentencia
Private Mens As String    'Para los mensajes
Private Rs As ADODB.Recordset
Private Rc As Byte




'------------------------------------------------
'Propiedades del modulo Cempresa
'------------------------------------------------

'*** codigo de la empresa

Public Property Let codempre(ByVal vData As String)
     mvarCodempre = vData
End Property

Public Property Get codempre() As String
     codempre = mvarCodempre
End Property


'*** nombre de la empresa

Public Property Let nomempre(ByVal vData As String)
     mvarNomempre = vData
End Property

Public Property Get nomempre() As String
     nomempre = mvarNomempre
End Property


'*** nombre resumen de la empresa
Public Property Let nomresum(ByVal vData As String)
     mvarNomresum = vData
End Property

Public Property Get nomresum() As String
     nomresum = mvarNomresum
End Property


'*** Base de Datos de la empresa

Public Property Let BDAriges(ByVal vData As String)
     mvarAriges = vData
End Property

Public Property Get BDAriges() As String
     BDAriges = mvarAriges
End Property


'*** nivel de las cuentas de contabilidad

Public Property Let numnivel(ByVal vData As Integer)
     mvarnumnivel = vData
End Property

Public Property Get numnivel() As Integer
     numnivel = mvarnumnivel
End Property


'*** digito de 1er nivel de las cuentas de contabilidad
Public Property Let numdigi1(ByVal vData As Integer)
     mvarnumdigi1 = vData
End Property

Public Property Get numdigi1() As Integer
     numdigi1 = mvarnumdigi1
End Property

'*** digito de nivel 2

Public Property Let numdigi2(ByVal vData As Integer)
     mvarnumdigi2 = vData
End Property

Public Property Get numdigi2() As Integer
     numdigi2 = mvarnumdigi2
End Property

'*** digito de nivel 3

Public Property Let numdigi3(ByVal vData As Integer)
     mvarnumdigi3 = vData
End Property

Public Property Get numdigi3() As Integer
     numdigi3 = mvarnumdigi3
End Property

'*** digito de nivel 4
Public Property Let numdigi4(ByVal vData As Integer)
     mvarnumdigi4 = vData
End Property

Public Property Get numdigi4() As Integer
     numdigi4 = mvarnumdigi4
End Property


'*** digito de nivel 8
Public Property Let numdigi8(ByVal vData As Integer)
     mvarnumdigi8 = vData
End Property

Public Property Get numdigi8() As Integer
     numdigi8 = mvarnumdigi8
End Property


'*** digito de nivel *

Public Property Let numdigi9(ByVal vData As Integer)
     mvarnumdigi9 = vData
End Property

Public Property Get numdigi9() As Integer
     numdigi9 = mvarnumdigi9
End Property

'*** digito de nivel 10

Public Property Let numdigi10(ByVal vData As Integer)
     mvarnumdigi10 = vData
End Property

Public Property Get numdigi10() As Integer
     numdigi10 = mvarnumdigi10
End Property


'*** digitos de la cuenta de ultimo nivel

Public Property Get DigitosUltimoNivel() As Byte
     DigitosUltimoNivel = mvarDigitosUltimoNivel
End Property


Public Property Let FechaIni(ByVal vData As Date)
     mvarFechaIni = vData
End Property

Public Property Get FechaIni() As Date
     FechaIni = mvarFechaIni
End Property

Public Property Let FechaFin(ByVal vData As Date)
     mvarFechaFin = vData
End Property

Public Property Get FechaFin() As Date
     FechaFin = mvarFechaFin
End Property




Public Property Let TieneAnalitica(ByVal vData As Boolean)
     mvarTieneAnalitica = vData
End Property

Public Property Get TieneAnalitica() As Boolean
     TieneAnalitica = mvarTieneAnalitica
End Property


Public Property Get SiiTiene() As Boolean
     SiiTiene = mvarSiiTiene
End Property

Public Property Get SiiFechaInicio() As Date
     SiiFechaInicio = mvarSiiFechaI
End Property

Public Property Get SiiDiasPlazo() As Integer
     SiiDiasPlazo = mvarDiasSii
End Property


        
'------------------------------------------------
'FUNCIONES del modulo Cempresa
'------------------------------------------------

''****   AGREGAR  ***
'Public Function Agregar() As Byte
''A�adir una nueva empresa en la BD de usuarios tabla:empresasariges
'
'    On Error GoTo Err1
'    'La Base de datos "Usuarios" esta en el mismo servidor que BD "Ariges"
'    SQL = "INSERT INTO usuarios.empresasariges (codempre, nomempre, nomresum)"
'    SQL = SQL & " VALUES ("
'    'Datos
'    SQL = SQL & mvarCodempre
'    SQL = SQL & "," & DBSet(mvarNomempre, "T")
'    SQL = SQL & "," & DBSet(mvarNomresum, "T")
'    SQL = SQL & ")"
'
'    Conn.Execute SQL
'    Agregar = 0
'    Exit Function
'
'Err1:
'    Mens = "Se ha producido un error." & vbCrLf
'    Mens = Mens & "N�mero: " & Err.Number & vbCrLf
'    Mens = Mens & "Descripci�n: " & Err.Description
'    MsgBox Mens, vbExclamation
'    Agregar = 1
'End Function



'****   LEER  ***
Public Function LeerDatos() As Byte
'Leer los parametros de una empresa
'Lee de la BD: Usuarios, Tabla: EmpresasAriges

    On Error GoTo Err3
    SQL = "SELECT * FROM usuarios.empresasarioli WHERE ariges='" & vUsu.CadenaConexion & "'"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Mens = ""
    
    If Rs.EOF Then
        LeerDatos = 1
    Else
        Mens = "Datos basicos"
        mvarCodempre = Rs!codempre
        mvarNomempre = Rs!nomempre
        mvarNomresum = DBLet(Rs!nomresum)
        mvarAriges = DBLet(Rs!AriGes)
        Mens = "Datos de Gesti�n"
        LeerDatos = 0
    End If

    Rs.Close
    Set Rs = Nothing
    Exit Function

Err3:
    Mens = "Se ha producido un error. " & Mens & vbCrLf
    Mens = Mens & "N�mero: " & Err.Number & vbCrLf
    Mens = Mens & "Descripci�n: " & Err.Description
    MsgBox Mens, vbExclamation
    Set Rs = Nothing
    LeerDatos = 1
End Function


'****   LEER  ***
Public Function LeerNiveles() As Byte
'Leer los parametros de la empresa en BDatos "Conta" correspondiente
'a la empresa de "Ariges" que estamos conectados

    On Error GoTo Err3
    
    SQL = "SELECT * FROM empresa "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
    Mens = ""
    
    If Rs.EOF Then
        LeerNiveles = 1
    Else
        Mens = "Datos basicos empresa en contabilidad."
        mvarnumnivel = Rs!numnivel
        mvarnumdigi1 = Rs!numdigi1
        mvarnumdigi2 = DBLet(Rs!numdigi2, "N")
        mvarnumdigi3 = DBLet(Rs!numdigi3, "N")
        mvarnumdigi4 = DBLet(Rs!numdigi4, "N")
        mvarnumdigi5 = DBLet(Rs!numdigi5, "N")
        mvarnumdigi6 = DBLet(Rs!numdigi6, "N")
        mvarnumdigi7 = DBLet(Rs!numdigi7, "N")
        mvarnumdigi8 = DBLet(Rs!numdigi8, "N")
        mvarnumdigi9 = DBLet(Rs!numdigi9, "N")
        mvarnumdigi10 = DBLet(Rs!numdigi10, "N")
        'Fijamos la longitud maxima de digitos en ctas de ultimo nivel
        FijaLongitudMaxima
        LeerNiveles = 0
    End If
    Rs.Close
    
    
    'Asigno los valores por defecto a las variables de fechaini y fechafin
    If LeerNiveles = 0 Then
        LeerNiveles = 2
        SQL = "Select * from parametros"
        Rs.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Rs.EOF Then
            MsgBox "Fecha inicio y fin de ejercicio", vbExclamation
        Else
            mvarFechaIni = Rs!FechaIni
            mvarFechaFin = Rs!FechaFin
            mvarTieneAnalitica = Rs!Autocoste = 1
            
            FijarValoresSii
            
            
            
            
        End If
        Rs.Close
    End If
    
    
    Set Rs = Nothing
    Exit Function

Err3:
    Mens = "Se ha producido un error. " & Mens & vbCrLf
    Mens = Mens & "N�mero: " & Err.Number & vbCrLf
    Mens = Mens & "Descripci�n: " & Err.Description
    MsgBox Mens, vbExclamation
    Set Rs = Nothing
    LeerNiveles = 1
End Function


Private Sub FijarValoresSii()
     On Error GoTo eSIIFechaInicio
     mvarSiiTiene = False
     
     If Not IsNull(Rs!SiiFechaInicio) Then
        mvarSiiTiene = True
        mvarSiiFechaI = Rs!SiiFechaInicio
        mvarDiasSii = Rs!SIIDiasAviso
     End If

eSIIFechaInicio:
    Err.Clear
End Sub

'****   MODIFICAR  ***
Public Function ModificarDatos() As Byte
'Modifica los parametros de una empresa
'Modifica la Tabla:empresasariges en la BD: Usuarios
    
    On Error GoTo Err4

    'Modificamos en la tabla Usuarios.empresasariges
    SQL = "UPDATE usuarios.empresasarioli SET nomempre= '" & mvarNomempre
    SQL = SQL & "', nomresum= '" & mvarNomresum
    SQL = SQL & "' WHERE codempre=" & mvarCodempre
    conn.Execute SQL
    
    ModificarDatos = 0
    Exit Function

Err4:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "N�mero: " & Err.Number & vbCrLf
    Mens = Mens & "Descripci�n: " & Err.Description
    MsgBox Mens, vbExclamation
    Rs.Close
    ModificarDatos = 1
End Function



'---------------------------------------------------------
'FUNCIONES del modulo Cempresa referentes a Contabilidad
'---------------------------------------------------------

Private Sub FijaLongitudMaxima()
Dim i As Byte

    Select Case mvarnumnivel
        Case 1
            i = mvarnumdigi1
        Case 2
            i = mvarnumdigi2
        Case 3
            i = mvarnumdigi3
        Case 4
            i = mvarnumdigi4
        Case 5
            i = mvarnumdigi5
        Case 6
            i = mvarnumdigi6
        Case 7
            i = mvarnumdigi7
        Case 8
            i = mvarnumdigi8
        Case 9
            i = mvarnumdigi9
        Case 10
            i = mvarnumdigi10
    End Select
    mvarDigitosUltimoNivel = i
End Sub

