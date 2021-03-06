VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'CLIENTES
'TABLA: sclien


'ATRIBUTOS
'Variables locales que contienen valores de propiedad
Private mCodClien As String 'Cod. cliente
Private mNomClien As String 'Nombre cliente
Private mNomComer As String 'Nombre comercial cliente
Private mDomClien As String 'Domicilio cliente
Private mCodpobla As String 'cod. postal cliente
Private mPobClien As String 'poblacion cliente
Private mProClien As String 'Provincia cliente
Private mNIFClien As String 'NIF cliente
Private mTfnClien As String 'Telefono administracion
Private mMaiClien As String 'Mail administracion

Private mCodEnvio As String 'cod. envio
Private mCodAgent As String 'cod. agente
Private mCodForpa As String 'cod. forma de pago
Private mDtoPpago As String 'descuento pronto pago
Private mDtoGnral As String 'descuento general
Private mTipoFact As Byte   'tipo facturacion (0:factura colectiva, 1:facturaxAlbaran

Private mCodbanco As String 'cod. banco
Private mCodsucur As String 'cod. sucur
Private mDigContr As String 'digito control
Private mCuentaba As String 'cuenta bancaria

Private mObservac As String 'Observaciones cliente
Private mCliVario As Boolean 'si es cliente de varios

Private mCodTarif As Integer 'cod. tarifa de precios
Private mCodSitua As Byte 'codigo de situacion
Private mKilometr As Integer 'kimlometros

Private mFecUltMovim As String 'Fecha del ult. movimiento del cliente

Private mWwwClien As String 'web del cliente


'Dic2013
Private mIban As String

'Jun17
Private mPais As String

'------------------------------------------------
'Propiedades del modulo CCliente
'------------------------------------------------

'**** Codigo del cliente

Public Property Let Codigo(ByVal vData As String)
     mCodClien = vData
End Property

Public Property Get Codigo() As String
     Codigo = mCodClien
End Property


'**** Nombre del cliente

Public Property Let Nombre(ByVal vData As String)
     mNomClien = vData
End Property

Public Property Get Nombre() As String
     Nombre = mNomClien
End Property


'**** Domicilio del cliente

Public Property Let Domicilio(ByVal vData As String)
     mDomClien = vData
End Property

Public Property Get Domicilio() As String
     Domicilio = mDomClien
End Property

'**** CPostal del cliente

Public Property Let CPostal(ByVal vData As String)
     mCodpobla = vData
End Property

Public Property Get CPostal() As String
     CPostal = mCodpobla
End Property

'**** Poblacion del cliente

Public Property Let Poblacion(ByVal vData As String)
     mPobClien = vData
End Property

Public Property Get Poblacion() As String
     Poblacion = mPobClien
End Property

'**** Provincia del cliente

Public Property Let Provincia(ByVal vData As String)
     mProClien = vData
End Property

Public Property Get Provincia() As String
     Provincia = mProClien
End Property

'**** NIF del cliente

Public Property Let NIF(ByVal vData As String)
     mNIFClien = vData
End Property

Public Property Get NIF() As String
     NIF = mNIFClien
End Property

'**** Telefono administracion del cliente

Public Property Let TfnoClien(ByVal vData As String)
     mTfnClien = vData
End Property

Public Property Get TfnoClien() As String
     TfnoClien = mTfnClien
End Property

'**** e-MAil administracion del cliente

Public Property Let EMailAdm(ByVal vData As String)
     mMaiClien = vData
End Property

Public Property Get EMailAdm() As String
     EMailAdm = mMaiClien
End Property


'**** Web del cliente

Public Property Let WebClien(ByVal vData As String)
     mWwwClien = vData
End Property

Public Property Get WebClien() As String
     WebClien = mWwwClien
End Property



'**** Tipo Facturacion del cliente

Public Property Let TipoFactu(ByVal vData As Byte)
     mTipoFact = vData
End Property

Public Property Get TipoFactu() As Byte
     TipoFactu = mTipoFact
End Property

'**** Cod. Forma Envio del Cliente

Public Property Let FEnvio(ByVal vData As String)
     mCodEnvio = vData
End Property

Public Property Get FEnvio() As String
     FEnvio = mCodEnvio
End Property


'**** Agente del Cliente

Public Property Let Agente(ByVal vData As String)
     mCodAgent = vData
End Property

Public Property Get Agente() As String
     Agente = mCodAgent
End Property


'**** Forma de pago del Cliente

Public Property Let ForPago(ByVal vData As String)
     mCodForpa = vData
End Property

Public Property Get ForPago() As String
     ForPago = mCodForpa
End Property


'**** Descuento pronto pago del Cliente

Public Property Let DtoPPago(ByVal vData As String)
     mDtoPpago = vData
End Property

Public Property Get DtoPPago() As String
     DtoPPago = mDtoPpago
End Property

'**** Descuento general del Cliente

Public Property Let DtoGnral(ByVal vData As String)
     mDtoGnral = vData
End Property

Public Property Get DtoGnral() As String
     DtoGnral = mDtoGnral
End Property


'**** Banco del Cliente

Public Property Let Banco(ByVal vData As String)
     mCodbanco = vData
End Property

Public Property Get Banco() As String
     Banco = mCodbanco
End Property


'**** Sucursal del Cliente

Public Property Let Sucursal(ByVal vData As String)
     mCodsucur = vData
End Property

Public Property Get Sucursal() As String
     Sucursal = mCodsucur
End Property


'**** Digito Control de la cuenta del cliente

Public Property Let DigControl(ByVal vData As String)
     mDigContr = vData
End Property

Public Property Get DigControl() As String
     DigControl = mDigContr
End Property


'**** Cuenta Bancaria de la Factura

Public Property Let CuentaBan(ByVal vData As String)
     mCuentaba = vData
End Property

Public Property Get CuentaBan() As String
     CuentaBan = mCuentaba
End Property


'**** Observaciones del cliente

Public Property Let Observaciones(ByVal vData As String)
     mObservac = vData
End Property

Public Property Get Observaciones() As String
     Observaciones = mObservac
End Property


'**** cliente de varios (SI/NO)

Public Property Let DeVarios(ByVal vData As Boolean)
     mCliVario = vData
End Property

Public Property Get DeVarios() As Boolean
     DeVarios = mCliVario
End Property



'**** Tarifa precios del cliente

Public Property Let Tarifa(ByVal vData As Integer)
     mCodTarif = vData
End Property

Public Property Get Tarifa() As Integer
     Tarifa = mCodTarif
End Property



'**** situacion cliente

Public Property Let Situacion(ByVal vData As Byte)
     mCodSitua = vData
End Property

Public Property Get Situacion() As Byte
     Situacion = mCodSitua
End Property


'**** Kilometros cliente

Public Property Let Kilometros(ByVal vData As Integer)
     mKilometr = vData
End Property

Public Property Get Kilometros() As Integer
     Kilometros = mKilometr
End Property



'**** Fecha ult. movimiento del cliente

Public Property Let FechaUltMov(ByVal vData As String)
     mFecUltMovim = vData
End Property

Public Property Get FechaUltMov() As String
     FechaUltMov = mFecUltMovim
End Property




Public Property Let IBAN(ByVal vData As String)
     mIban = vData
End Property

Public Property Get IBAN() As String
     IBAN = mIban
End Property


Public Property Get Pais() As String
     Pais = mPais
End Property


'------------------------------------------------
'Procedimientos del modulo CCliente
'------------------------------------------------

Public Function Existe(vCodClien As String) As Boolean
'Comprueba si existe el cliente en la BD
Dim Devuelve As String

    On Error GoTo EExiste
    
    Devuelve = DevuelveDesdeBD(conAri, "codclien", "sclien", "codclien", vCodClien, "N")
    If Devuelve = "" Then
        Existe = False
        Devuelve = "No existe el Cliente " & vCodClien & ". "
        MsgBox Devuelve, vbExclamation
    Else
        Existe = True
    End If
    
EExiste:
    If Err.Number <> 0 Then Existe = False
End Function



Public Function LeerDatos(vCodClien As String) As Boolean
'Leer los datos de un cliente dado
'Lee de la BD: Ariges, Tabla: sclien
'OUT: True si lee los datos correctamente
Dim RS As ADODB.Recordset
Dim SQL As String

    On Error GoTo ELeer
    
    LeerDatos = False
    
    SQL = "SELECT codclien,nomclien,nomcomer,codpobla,domclien,pobclien,proclien,nifclien,wwwclien,codenvio,codagent,codforpa,codbanco,codsucur,digcontr,cuentaba,telclie1,maiclie1,clivario,dtoppago,dtognral,codtarif,codsitua,tipofact,kilometr,fechamov,observac,iban "
    If vParamAplic.ContabilidadNueva Then SQL = SQL & ", codpais"
    SQL = SQL & " FROM sclien "
    SQL = SQL & " WHERE codclien=" & vCodClien
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    If RS.EOF Then
        LeerDatos = False
    Else
        mCodClien = CStr(RS!CodClien)
        
        mNomClien = RS!nomclien
        mNomComer = RS!nomcomer
        mDomClien = DBLet(RS!domclien, "T")
        mCodpobla = DBLet(RS!codpobla, "T")
        mPobClien = DBLet(RS!pobclien, "T")
        mProClien = DBLet(RS!proclien, "T")
        mNIFClien = DBLet(RS!nifClien, "T")
        mWwwClien = DBLet(RS!wwwclien, "T")
        mTfnClien = DBLet(RS!telclie1, "T")
        mMaiClien = DBLet(RS!maiclie1, "T")
       
        mCodEnvio = CStr(DBLet(RS!codenvio, "N"))
        mCodAgent = CStr(RS!codagent)
        mCodForpa = CStr(RS!codforpa)
        mDtoPpago = CStr(RS!DtoPPago)
        mDtoGnral = CStr(RS!DtoGnral)
        mTipoFact = RS!TipoFact
        
        mCodbanco = CStr(DBLet(RS!codbanco, "N"))
        mCodsucur = CStr(DBLet(RS!codsucur, "N"))
        mDigContr = DBLet(RS!digcontr, "T")
        mCuentaba = DBLet(RS!cuentaba, "T")
        
        mIban = DBLet(RS!IBAN, "T")
        
        mObservac = DBLet(RS!observac, "T")
        mCliVario = RS!CliVario
        
        Tarifa = RS!codTarif
        Situacion = RS!codsitua
        mKilometr = DBLet(RS!kilometr, "N")

        mFecUltMovim = DBLet(RS!Fechamov, "T")
        
        If vParamAplic.ContabilidadNueva Then mPais = DBLet(RS!codpais, "T")
        
        
        LeerDatos = True
    End If

    RS.Close
    Set RS = Nothing
    Exit Function

ELeer:
    SQL = "Se ha producido un error. " & "Datos Cliente" & vbCrLf
    SQL = SQL & "N�mero: " & Err.Number & vbCrLf
    SQL = SQL & "Descripci�n: " & Err.Description
    MsgBox SQL, vbExclamation
    Set RS = Nothing
    LeerDatos = False
End Function



Public Function LeerDatosCliVario(vNIF As String) As Boolean
'si es cliente de varios recuperamos los datos nombre,direc,etc.
'de la tabla de varios: sclvar
Dim RS As ADODB.Recordset
Dim SQL As String
    
    On Error GoTo ELeer
    LeerDatosCliVario = False
    
    SQL = "SELECT nifclien,nomclien,domclien,codpobla,pobclien,proclien,telclien "
    SQL = SQL & " FROM sclvar "
    SQL = SQL & " WHERE nifclien=" & DBSet(vNIF, "T")
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    If RS.EOF Then
        LeerDatosCliVario = False
        
        mNomClien = ""
        mNomComer = ""
        mDomClien = ""
        mCodpobla = ""
        mPobClien = ""
        mProClien = ""
        mTfnClien = ""
        MsgBox "No existe el cliente vario: " & vNIF, vbInformation
    Else
        mNIFClien = RS!nifClien
        mNomClien = RS!nomclien
        mNomComer = ""
        mDomClien = RS!domclien
        mCodpobla = DBLet(RS!codpobla, "T")
        mPobClien = DBLet(RS!pobclien, "T")
        mProClien = DBLet(RS!proclien, "T")
        mTfnClien = DBLet(RS!telclien, "T")
        
        LeerDatosCliVario = True
    End If

    RS.Close
    Set RS = Nothing
    Exit Function

ELeer:
    SQL = "Se ha producido un error. " & "Datos Cliente de Varios." & vbCrLf
    SQL = SQL & "N�mero: " & Err.Number & vbCrLf
    SQL = SQL & "Descripci�n: " & Err.Description
    MsgBox SQL, vbExclamation
    Set RS = Nothing
    LeerDatosCliVario = False
End Function



Public Function ActualizarClienteV(vNIF As String) As Boolean
'Actualiza los datos del cliente de varios en la tabla: sclvar
Dim SQL As String
Dim Devuelve As String

    On Error GoTo EActualizarCliente
    ActualizarClienteV = False
    
    Devuelve = DevuelveDesdeBDNew(conAri, "sclvar", "nifclien", "nifclien", vNIF, "T")
    
    If Devuelve <> "" Then 'Existe el cliente de varios y lo actualizamos
        SQL = "UPDATE sclvar SET nomclien=" & DBSet(Nombre, "T") & ", domclien=" & DBSet(Domicilio, "T") & ", "
        SQL = SQL & " codpobla=" & DBSet(CPostal, "T") & ", pobclien=" & DBSet(Poblacion, "T") & ", "
        SQL = SQL & " proclien=" & DBSet(Provincia, "T") & ", telclien=" & DBSet(TfnoClien, "T")
        SQL = SQL & " WHERE nifclien=" & DBSet(vNIF, "T")
    Else 'No existe el cliente y lo a�adimos
        SQL = "INSERT INTO sclvar (nifclien, nomclien, domclien, codpobla, pobclien, proclien, telclien) "
        SQL = SQL & " VALUES (" & DBSet(vNIF, "T") & ", " & DBSet(Nombre, "T") & ", " & DBSet(Domicilio, "T") & ", "
        SQL = SQL & DBSet(CPostal, "T") & ", " & DBSet(Poblacion, "T") & ", " & DBSet(Provincia, "T") & ", " & DBSet(TfnoClien, "T") & ")"
    End If
    conn.Execute SQL, , adCmdText
    
    ActualizarClienteV = True
    
EActualizarCliente:
    If Err.Number <> 0 Then
        ActualizarClienteV = False
        MuestraError Err.Number, "Actualizando Cliente de Varios.", Err.Description
    End If
End Function


Public Function ClienteBloqueado() As Boolean
'devuelve true si el cliente esta bloqueado
'si la situaci�n del cliente es distinta de NORMAL(codsitua=0) entonces
'mostrar un mensaje con la situaci�n especial del cliente
Dim Tipo As String
Dim Devuelve As String

    On Error GoTo EBloqueado
    ClienteBloqueado = False
    
    If Situacion <> 0 Then '0: situacion normal
        Tipo = "tipositu"
        Devuelve = DevuelveDesdeBDNew(conAri, "ssitua", "nomsitua", "codsitua", Situacion, "N", Tipo)
        
        If Tipo = "1" Then 'Cliente Bloqueado por Situaci�n Especial.
            MsgBox UCase("Cliente Bloqueado por: ") & vbCrLf & Devuelve, vbInformation, "Situaci�n Especial del Cliente."
            ClienteBloqueado = True
        Else
            MsgBox Devuelve, vbInformation, "Situaci�n Especial del Cliente."
        End If
    End If
    
EBloqueado:
    If Err.Number <> 0 Then Err.Clear
End Function



Public Function DptoCliente(vCodDpto As String, NomDpto As String) As Boolean
'comprueba si existe el departamento para el cliente y devuelve su nombre
'IN: vCodDpto= codigo del departamento o direccion del cliente
'OUT: NomDppto= devuelve el nombre del dpto o direccion si lo encuentra
Dim Devuelve As String
Dim cadDpto As String
    
    If Codigo = "" Then
        Devuelve = "Debe introducir valor al campo Cliente, para cargar sus "
        If vParamAplic.Departamento Then 'Se trabaja con dpto
            Devuelve = Devuelve & "Departamentos. "
        Else 'la aplicacion trabaja con direcciones
            Devuelve = Devuelve & "Direcciones. "
        End If
        MsgBox Devuelve, vbInformation
        Exit Function
    End If

    'Comprobar que el cliente seleccionado tiene esa direccion
    Devuelve = DevuelveDesdeBDNew(conAri, "sdirec", "nomdirec", "codclien", Codigo, "N", , "coddirec", vCodDpto, "N")
    NomDpto = Devuelve
    If Devuelve = "" Then
        DptoCliente = False
        
        If vParamAplic.Departamento Then 'Se trabaja con dpto
            cadDpto = " el Departamento "
        Else 'la aplicacion trabaja con direcciones
            cadDpto = " la Direcci�n "
        End If
        
        Devuelve = "No existe" & cadDpto & vCodDpto & " para el cliente: "
        Devuelve = Devuelve & Codigo
        MsgBox Devuelve, vbInformation
    Else
        DptoCliente = True
    End If
End Function


Public Function TieneRefObligatoria(vReferen As String) As Boolean
'Comprueba si el campo de referencia obligatoria es necesario que se rellene
'al introducir datos de ese cliente
Dim Devuelve As String

    Devuelve = DevuelveDesdeBDNew(conAri, "sclien", "referobl", "codclien", Codigo, "N")
    If Devuelve = "1" And vReferen = "" Then 'Referencia Obligatoria
        TieneRefObligatoria = True
        Devuelve = "La Referencia del Cliente es Obligatoria."
        MsgBox Devuelve, vbInformation
    Else
        TieneRefObligatoria = False
    End If
End Function


Public Function exentoIVA2() As Byte
'Comprueba si el cliente esta exento de IVA para calcular el %IVA de las ofertas,pedidos
'albaranes, facturas,...
'Dim devuelve As String

    exentoIVA2 = DevuelveDesdeBDNew(conAri, "sclien", "tipoiva", "codclien", Codigo, "N")
    'exentoIVA = devuelve  '2 exento   0 normal  1 RE
    'If devuelve = "2" Then 'exento de iva
    '    exentoIVA = True
    'Else
    '    exentoIVA = False
    'End If
End Function


Public Function Bonifica() As Boolean
'Indica si al cliente se le aplican bonificaciones o no
Dim Cad As String

    On Error GoTo EBonifica
    Bonifica = False
    
    Cad = DevuelveDesdeBDNew(conAri, "sclien", "codtarif", "codclien", Codigo, "N")
    If Cad <> "" Then
        Cad = DevuelveDesdeBDNew(conAri, "starif", "bonifica", "codlista", Cad, "N")
        If Cad <> "" Then Bonifica = CBool(Cad)
    End If
    
EBonifica:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar bonificaci�n del cliente.", Err.Description
    End If
End Function


Public Function ActualizaUltFecMovim(vFecha As String) As Boolean
'Actualiza la fecha del ultimo movimiento del cliente si la fecha del movimiento
'que se pasa por parametro es posterior a la que tiene el cliente
Dim SQL As String
Dim b As Boolean
Dim actualiza As Boolean

    On Error GoTo EActFecha
    
    If FechaUltMov = "" Then _
        FechaUltMov = DevuelveDesdeBDNew(conAri, "sclien", "fechamov", "codclien", Codigo, "N")
    
    
    If FechaUltMov = "" Then
        actualiza = True
    ElseIf CDate(vFecha) > CDate(FechaUltMov) Then
        actualiza = True
    Else
        actualiza = False
    End If
    
    If actualiza Then
        SQL = "UPDATE sclien SET fechamov=" & DBSet(vFecha, "F")
        SQL = SQL & " WHERE codclien=" & Codigo
        conn.Execute SQL, , adCmdText
    End If
    b = True
        
EActFecha:
    If Err.Number <> 0 Then b = False
    ActualizaUltFecMovim = b
End Function




Public Sub MostrarObservaciones()
'Recuperamos el texto de las observaciones del cliente, y si tiene
'valor lo mostramos en un mensaje
    If Me.Observaciones <> "" Then _
         MsgBox Me.Observaciones, vbInformation, "Observaciones del cliente"
End Sub




Public Function ComprobarCtaBancaria(cadMen As String) As Boolean
'comprueba q el cliente tenga cuenta bancaria y que sea correcta
Dim cta As String

    If Me.Banco = 0 And Me.Sucursal = 0 And Me.DigControl = "" And Me.CuentaBan = "" Then
        ComprobarCtaBancaria = False
        cadMen = "El cliente no tiene asignada cuenta bancaria."
        Exit Function
    End If
    
    
    
    cta = Format(Me.Banco, "0000") & Format(Me.Sucursal, "0000") & Format(Me.DigControl, "00") & Format(Me.CuentaBan, "0000000000")
    
    If Val(ComprobarCero(cta)) = 0 Then
        ComprobarCtaBancaria = False
        cadMen = "El cliente no tiene asignada cuenta bancaria."
        Exit Function
    End If
    
    
    If Not Comprueba_CC(cta) Then
        cadMen = "La cuenta bancaria del cliente no es correcta."
        ComprobarCtaBancaria = False
    Else
        ComprobarCtaBancaria = True
    End If
End Function
