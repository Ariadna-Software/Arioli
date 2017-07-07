VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'VENTAS
'FACTURA: se corresponde con la tabla ariges.scafac (cabecera de facturas)
'-------------------------------------------------------------------------

'ATRIBUTOS
'Variables locales que contienen valores de propiedad
Private mCodTipoM As String 'Tipo de Movimiento="FAV", "FAR" (Factura venta o reparacion)
Private mNumFactu As Long 'Nº Factura
Private mFecFactu As String 'FEcha Factura
Private mCodClien As String 'Cod. cliente

Private mNomClien As String 'nombre del cliente
Private mDomClien As String 'domicilio
Private mCodpobla As String 'cpostal
Private mPobClien As String 'poblacion
Private mProClien As String 'provincia
Private mNIFClien As String 'NIF proveedor
Private mTelclien As String 'telefono

Private mCoddirec As String 'codigo direc./dpto
Private mNomdirec As String  'nombre direc./dpto
Private mCodAgent As Integer 'codigo de agente


Private mCodbanco As String 'cod. banco
Private mCodsucur As String 'cod. sucur
Private mDigContr As String 'digito control
Private mCuentaba As String 'cuenta bancaria

Private mCodForpa As String 'Cod. Forma de pago
Private mTipForpa As Byte 'tipo de forma de pago: efectivo,...

Private mDtoPpago As Currency 'descuento pronto pago
Private mDtoGnral As Currency 'descuento general
Private mBrutofac As Currency 'importe bruto de la factura
Private mImpPPago As Currency 'Importe al aplicar al bruto el dto pronto pago
Private mImpGnral As Currency  'Importe al aplicar al bruto el dto general
Private mBaseImp As Currency  'Base imponible de la factura (bruto - dtopp - dtogn)

Private mBaseiva1 As Currency  'base imponible tipo IVA 1
Private mBaseiva2 As Currency  'base imponible tipo IVA 2
Private mBaseiva3 As Currency  'base imponible tipo IVA 3
Private mTipoiva1 As Byte   'Cod. tipo de iva 1
Private mTipoiva2 As Byte   'Cod. tipo de iva 2
Private mTipoiva3 As Byte   'Cod. tipo de iva 3
Private mPorciva1 As Currency    '% de iva 1
Private mPorciva2 As Currency    '% de iva 2
Private mPorciva3 As Currency    '% de iva 3

Private mImpoiva1 As Currency   'Importe de iva 1
Private mImpoiva2 As Currency   'Importe de iva 2
Private mImpoiva3 As Currency   'Importe de iva 3


'DAVID.   ENERO 2008
Private mPorciva1RE As Currency    '% de iva 1
Private mPorciva2RE As Currency    '% de iva 2
Private mPorciva3RE As Currency    '% de iva 3

Private mImpoiva1RE As Currency   'Importe de iva 1
Private mImpoiva2RE As Currency   'Importe de iva 2
Private mImpoiva3RE As Currency   'Importe de iva 3





Private mTotalfac As String 'total factura

Private mLetraSer As String 'Letra de serie para los cobros
Private mBanPr As String 'Banco propio que ponemos como cuenta prevista de cobro
Private mCtaPrev As String 'Cuenta prevista de cobro

Private mNumtermi As Integer 'terminal de venta del TPV
Private mNumventa As Long 'Nº de venta del TPV


'David.
'Para las facturas de mantenimeinto, hay un campo que es "concepto factura" varchar(60)
'Si tiene valor entonces ese valor es el que grabaremos en la linea y como obervacion de la factura
'meteremos lo de antes ("texto" de: mes de año.  Ejemplo: manteimiento de Junio de 2007
'Si no tiene valor seguira como estaba
Private mObservacion As String






'Dic. 2013
'IBAN
Private mIban As String

Private mPais As String


' VBLES para los Mantenimientos -------------------------------------------
Dim TipCoMan As String 'tipo de contrato para Facturas de Manteniminetos
Dim OpeFactu As String 'Operador conectado o el que pasamos como parametro
Dim AlmFactu As String 'Cod Almacen del trabajador conectado, para insertar lines mantenim, desplaz, bonificaciones,....
Dim MesFactu As String
Dim numMante As String
Dim ArticFac As String 'cod articulo que se va a utilizar para facturar mantenimientos
'--------------------------------------------------------------------------

' VBLES para los Desplazamientos ----------------------------------
Dim TotalKm As Integer
Dim PrecioKm As Currency
Dim ImporteL As Currency
Dim Ampliacion As String  'ampliacion de la linea: "Dia,Dia,Dia..." de los albaranes con desplazamiento
Dim codartic As String
Dim despIVA As Integer 'tipo de IVA del articulo para desplazamientos
'------------------------------------------------------------------

' VBLES para las Bonificaciones -----------------------------------
Dim TotBonif1 As Currency 'Total de bonificacion a factura con articulo de tipo de IVA1
Dim IVABonif1 As Byte 'tipo de IVA
Dim TotBonif2 As Currency
Dim IVABonif2 As Byte
Dim TotBonif3 As Currency
Dim IVABonif3 As Byte
Dim SQLBonif As String
'------------------------------------------------------------------

' VBLES para los Cheques Regalo -----------------------------------
Dim ImpCheque As Currency
'------------------------------------------------------------------




'vbles para las facturas que vienen de una venta de ticket ---------
Private NumTicket As String
Private NumAlbTicket As String '01/09/06 laura
'-------------------------------------------------------------------

Private mAportacion As Currency   'Viene de Radival, pero se podria utilizar en mas
                                 ' Es la aportacion que hace Telefonica, en los moviles
                                 

'# Laura 14/11/2006 recuperar facturas Alzira
Private EstaRecuFac As Boolean


'------------------------------------------------
'Propiedades del modulo CFactura
'------------------------------------------------

'**** Tipo de Movimiento de la Factura

Public Property Let Codtipom(ByVal vData As String)
     mCodTipoM = vData
End Property

Public Property Get Codtipom() As String
     Codtipom = mCodTipoM
End Property


'**** Nº de la Factura

Public Property Let NumFactu(ByVal vData As Long)
     mNumFactu = vData
End Property

Public Property Get NumFactu() As Long
     NumFactu = mNumFactu
End Property


'**** Fecha de la Factura

Public Property Let Fecfactu(ByVal vData As String)
     mFecFactu = vData
End Property

Public Property Get Fecfactu() As String
     Fecfactu = mFecFactu
End Property


'**** Cliente de la Factura

Public Property Let Cliente(ByVal vData As String)
     mCodClien = vData
End Property

Public Property Get Cliente() As String
     Cliente = mCodClien
End Property


'**** Nombre del Cliente de la Factura

Public Property Let NombreClien(ByVal vData As String)
     mNomClien = vData
End Property

Public Property Get NombreClien() As String
     NombreClien = mNomClien
End Property


'**** Domicilio del Cliente de la Factura

Public Property Let DomicilioClien(ByVal vData As String)
     mDomClien = vData
End Property

Public Property Get DomicilioClien() As String
     DomicilioClien = mDomClien
End Property


'**** CPostal del Cliente de la Factura

Public Property Let CPostal(ByVal vData As String)
     mCodpobla = vData
End Property

Public Property Get CPostal() As String
     CPostal = mCodpobla
End Property


'**** Poblacion del Cliente de la Factura

Public Property Let Poblacion(ByVal vData As String)
     mPobClien = vData
End Property

Public Property Get Poblacion() As String
     Poblacion = mPobClien
End Property


'**** Provincia del Cliente de la Factura

Public Property Let Provincia(ByVal vData As String)
     mProClien = vData
End Property

Public Property Get Provincia() As String
     Provincia = mProClien
End Property


'**** NIF del Cliente de la Factura

Public Property Let NIF(ByVal vData As String)
     mNIFClien = vData
End Property

Public Property Get NIF() As String
     NIF = mNIFClien
End Property


'**** Telefono del Cliente de la Factura

Public Property Let Telefono(ByVal vData As String)
     mTelclien = vData
End Property

Public Property Get Telefono() As String
     Telefono = mTelclien
End Property


'**** Direc./Dpto de Cliente de la Factura

Public Property Let DirDpto(ByVal vData As String)
     mCoddirec = vData
End Property

Public Property Get DirDpto() As String
     DirDpto = mCoddirec
End Property


'**** Nombre Direc./Dpto de Cliente de la Factura

Public Property Let NombreDirDpto(ByVal vData As String)
     mNomdirec = vData
End Property

Public Property Get NombreDirDpto() As String
     NombreDirDpto = mNomdirec
End Property


'**** Agente de Cliente de la Factura

Public Property Let Agente(ByVal vData As Integer)
     mCodAgent = vData
End Property

Public Property Get Agente() As Integer
     Agente = mCodAgent
End Property


'**** Forma Pago de la Factura

Public Property Let ForPago(ByVal vData As String)
     mCodForpa = vData
End Property

Public Property Get ForPago() As String
     ForPago = mCodForpa
End Property


'**** Tipo de Forma Pago de la Factura
'(para tickets saber si insertar en tesoreria o no: si es efectivo no se inserta en scobro)

Public Property Let TipForPago(ByVal vData As Byte)
     mTipForpa = vData
End Property

Public Property Get TipForPago() As Byte
     TipForPago = mTipForpa
End Property


'**** Cuenta Prevista de cobro de la Factura

Public Property Let LetraSerie(ByVal vData As String)
     mLetraSer = vData
End Property

Public Property Get LetraSerie() As String
     LetraSerie = mLetraSer
End Property


'**** Banco propio de pago de la Factura

Public Property Let BancoPr(ByVal vData As String)
     mBanPr = vData
End Property

Public Property Get BancoPr() As String
     BancoPr = mBanPr
End Property


'**** Cuenta Prevista de cobro de la Factura

Public Property Let CuentaPrev(ByVal vData As String)
     mCtaPrev = vData
End Property

Public Property Get CuentaPrev() As String
     CuentaPrev = mCtaPrev
End Property


'**** Banco de la Factura

Public Property Let Banco(ByVal vData As String)
     mCodbanco = vData
End Property

Public Property Get Banco() As String
     Banco = mCodbanco
End Property


'**** Sucursal de la Factura

Public Property Let Sucursal(ByVal vData As String)
     mCodsucur = vData
End Property

Public Property Get Sucursal() As String
     Sucursal = mCodsucur
End Property


'**** Digito Control de la Factura

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


'**** Descuento pronto pago de la Factura

Public Property Let DtoPPago(ByVal vData As Currency)
     mDtoPpago = vData
End Property

Public Property Get DtoPPago() As Currency
     DtoPPago = mDtoPpago
End Property


'**** Descuento general de la Factura

Public Property Let DtoGnral(ByVal vData As Currency)
     mDtoGnral = vData
End Property

Public Property Get DtoGnral() As Currency
     DtoGnral = mDtoGnral
End Property


'**** Importe Bruto Factura

Public Property Let BrutoFac(ByVal vData As Currency)
     mBrutofac = vData
End Property

Public Property Get BrutoFac() As Currency
     BrutoFac = mBrutofac
End Property


'**** Importe pronto pago de la Factura

Public Property Let ImpPPago(ByVal vData As Currency)
     vData = vData
     mImpPPago = vData
End Property

Public Property Get ImpPPago() As Currency
     ImpPPago = mImpPPago
End Property


'**** Importe general de la Factura

Public Property Let ImpGnral(ByVal vData As Currency)
     mImpGnral = vData
End Property

Public Property Get ImpGnral() As Currency
     ImpGnral = mImpGnral
End Property


'**** Importe Base Imponible Factura

Public Property Let BaseImp(ByVal vData As Currency)
     mBaseImp = vData
End Property

Public Property Get BaseImp() As Currency
     BaseImp = mBaseImp
End Property


'**** Base IVA1 de la Factura

Public Property Let BaseIVA1(ByVal vData As Currency)
     mBaseiva1 = vData
End Property

Public Property Get BaseIVA1() As Currency
     BaseIVA1 = mBaseiva1
End Property


'**** Base IVA2 de la Factura

Public Property Let BaseIVA2(ByVal vData As Currency)
     mBaseiva2 = vData
End Property

Public Property Get BaseIVA2() As Currency
     BaseIVA2 = mBaseiva2
End Property


'**** Base IVA3 de la Factura

Public Property Let BaseIVA3(ByVal vData As Currency)
     mBaseiva3 = vData
End Property

Public Property Get BaseIVA3() As Currency
     BaseIVA3 = mBaseiva3
End Property


'**** Tipo iva 1 de la Factura

Public Property Let TipoIVA1(ByVal vData As Byte)
     mTipoiva1 = vData
End Property

Public Property Get TipoIVA1() As Byte
     TipoIVA1 = mTipoiva1
End Property


'**** Tipo iva 2 de la Factura

Public Property Let TipoIVA2(ByVal vData As Byte)
     mTipoiva2 = vData
End Property

Public Property Get TipoIVA2() As Byte
     TipoIVA2 = mTipoiva2
End Property


'**** Tipo iva 3 de la Factura

Public Property Let TipoIVA3(ByVal vData As Byte)
     mTipoiva3 = vData
End Property

Public Property Get TipoIVA3() As Byte
     TipoIVA3 = mTipoiva3
End Property


'**** % de iva 1 de la Factura

Public Property Let PorceIVA1(ByVal vData As Currency)
     mPorciva1 = vData
End Property

Public Property Get PorceIVA1() As Currency
     PorceIVA1 = mPorciva1
End Property


'**** % de iva 2 de la Factura

Public Property Let PorceIVA2(ByVal vData As Currency)
     mPorciva2 = vData
End Property

Public Property Get PorceIVA2() As Currency
     PorceIVA2 = mPorciva2
End Property


'**** % de iva 3 de la Factura

Public Property Let PorceIVA3(ByVal vData As Currency)
     mPorciva3 = vData
End Property

Public Property Get PorceIVA3() As Currency
     PorceIVA3 = mPorciva3
End Property


'**** Importe de IVA 1 de la Factura

Public Property Let ImpIVA1(ByVal vData As Currency)
     mImpoiva1 = vData
End Property

Public Property Get ImpIVA1() As Currency
     ImpIVA1 = mImpoiva1
End Property


'**** Importe de IVA 2 de la Factura

Public Property Let ImpIVA2(ByVal vData As Currency)
     mImpoiva2 = vData
End Property

Public Property Get ImpIVA2() As Currency
     ImpIVA2 = mImpoiva2
End Property


'**** Importe de IVA 3 de la Factura

Public Property Let ImpIVA3(ByVal vData As Currency)
     mImpoiva3 = vData
End Property

Public Property Get ImpIVA3() As Currency
     ImpIVA3 = mImpoiva3
End Property




'*****************************************************++
'   Recargo de equivalencia

'**** % de iva 1 de la Factura

Public Property Let PorceIVA1RE(ByVal vData As Currency)
     mPorciva1RE = vData
End Property

Public Property Get PorceIVA1RE() As Currency
     PorceIVA1RE = mPorciva1RE
End Property


'**** % de iva 2 de la Factura

Public Property Let PorceIVA2RE(ByVal vData As Currency)
     mPorciva2RE = vData
End Property

Public Property Get PorceIVA2RE() As Currency
     PorceIVA2RE = mPorciva2RE
End Property


'**** % de iva 3 de la Factura

Public Property Let PorceIVA3RE(ByVal vData As Currency)
     mPorciva3RE = vData
End Property

Public Property Get PorceIVA3RE() As Currency
     PorceIVA3RE = mPorciva3RE
End Property


'**** Importe de IVA 1 de la Factura

Public Property Let ImpIVA1RE(ByVal vData As Currency)
     mImpoiva1RE = vData
End Property

Public Property Get ImpIVA1RE() As Currency
     ImpIVA1RE = mImpoiva1RE
End Property


'**** Importe de IVA 2 de la Factura

Public Property Let ImpIVA2RE(ByVal vData As Currency)
     mImpoiva2RE = vData
End Property

Public Property Get ImpIVA2RE() As Currency
     ImpIVA2RE = mImpoiva2RE
End Property


'**** Importe de IVA 3 de la Factura

Public Property Let ImpIVA3RE(ByVal vData As Currency)
     mImpoiva3RE = vData
End Property

Public Property Get ImpIVA3RE() As Currency
     ImpIVA3RE = mImpoiva3RE
End Property












'**** Total Factura

Public Property Let TotalFac(ByVal vData As String)
     mTotalfac = vData
End Property

Public Property Get TotalFac() As String
     TotalFac = mTotalfac
End Property


'**** Nº terminal de la venta del TPV

Public Property Let NumTerminal(ByVal vData As Integer)
     mNumtermi = vData
End Property

Public Property Get NumTerminal() As Integer
     NumTerminal = mNumtermi
End Property



'**** Nº venta del TPV

Public Property Let NumVenta(ByVal vData As Long)
     mNumventa = vData
End Property

Public Property Get NumVenta() As Long
     NumVenta = mNumventa
End Property


'**** Para los MANTENIMIENTOS.  sera 1 la linea. Si esta vacio lo cargaremos con la amplicaicon normal

Public Property Let Observacion(ByVal vData As String)
     mObservacion = vData
End Property

Public Property Get Observacion() As String
     Observacion = mObservacion
End Property


'-----   Aportacion
Public Property Let Aportacion(ByVal vData As Currency)
     mAportacion = vData
End Property

Public Property Get Aportacion() As Currency
     Aportacion = mAportacion
End Property



Public Property Let iBan(ByVal vData As String)
     mIban = vData
End Property

Public Property Get iBan() As String
     iBan = mIban
End Property


Public Property Get Pais() As String
    Pais = mPais
End Property

'------------------------------------------------
'Procedimientos del modulo CFactura
'------------------------------------------------

Public Function LeerDatosN(vCodtipom As String, vNumfactu As String, vFecfactu As String) As Boolean
'Leer los datos de una factura almacenada en la tabla scafac
'Lee de la BD: Ariges, Tabla: scafac
'OUT: True si lee los datos correctamente
Dim RS As ADODB.Recordset
Dim Mens As String
Dim SQL As String

    On Error GoTo ELeer
    
    
    LeerDatosN = False
    SQL = "SELECT codtipom,numfactu,fecfactu,codclien,codforpa,codbanco,codsucur,digcontr,cuentaba,brutofac, totalfac,iban "
    If vParamAplic.ContabilidadNueva Then SQL = SQL & ",nomclien,domclien,pobclien,proclien,nifclien,codpobla"
    
    
    SQL = SQL & "  FROM scafac WHERE codtipom='" & vCodtipom & "'"
    SQL = SQL & " AND numfactu=" & vNumfactu
    SQL = SQL & " AND fecfactu='" & Format(vFecfactu, FormatoFecha) & "'"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Mens = ""
    
    If RS.EOF Then
        LeerDatosN = False
    Else
        Mens = "Datos Factura"
        mCodTipoM = RS!Codtipom
        mNumFactu = CStr(RS!NumFactu)
        mFecFactu = CStr(RS!Fecfactu)
        mCodClien = CStr(RS!CodClien)
        ForPago = CStr(RS!codforpa)
        TipForPago = DevuelveDesdeBDNew(conAri, "sforpa", "tipforpa", "codforpa", ForPago, "N")
        mCodbanco = CStr(DBLet(RS!codbanco, "N"))
        mCodsucur = CStr(DBLet(RS!codsucur, "N"))
        mDigContr = DBLet(RS!digcontr, "T")
        mCuentaba = DBLet(RS!cuentaba, "T")
        mIban = DBLet(RS!iBan, "T")
        mBrutofac = CStr(RS!BrutoFac)
        mTotalfac = CStr(RS!TotalFac)

        If vParamAplic.ContabilidadB Then
            '    nomclien , domclien, pobclien, proclien, nifClien, codpobla
            mNomClien = DBLet(RS!nomclien, "T")
            mDomClien = DBLet(RS!domclien, "T")
            mPobClien = DBLet(RS!pobclien, "T")
            mProClien = DBLet(RS!proclien, "T")
            mCodpobla = DBLet(RS!codpobla, "T")
            mNIFClien = DBLet(RS!nifClien, "T")
            mPais = DevuelveDesdeBD(conAri, "codpais", "sclien", "codclien", RS!CodClien)
            If mPais = "" Then mPais = "ES"
        End If
        mLetraSer = ObtenerLetraSerie(mCodTipoM)
        'falta### mPais = RS!S
        
        'si venimos de proceso de insertar tenemos la cuenta prevista de cobro que
        'se introdujo en el formulario de pasar albaranes a factura
        If CuentaPrev = "" Then
          
          
          If vParamAplic.ContabilidadNueva Then
                SQL = "SELECT COUNT(*) FROM cobros WHERE numserie='" & LetraSerie & "' and numfactu=" & NumFactu
                SQL = SQL & " AND fecfactu=" & DBSet(Fecfactu, "F")
            Else
                SQL = "SELECT COUNT(*) FROM scobro WHERE numserie='" & LetraSerie & "' and codfaccl=" & NumFactu
                SQL = SQL & " AND fecfaccl=" & DBSet(Fecfactu, "F")
            End If
            If RegistrosAListar(SQL, conConta) > 0 Then
                If vParamAplic.ContabilidadNueva Then
                    SQL = DevuelveDesdeBDNew(conConta, "cobros", "ctabanc1", "numserie", mLetraSer, "T", , "numfactu", vNumfactu, "N", "fecfactu", vFecfactu, "F")
                Else
                    SQL = DevuelveDesdeBDNew(conConta, "scobro", "ctabanc1", "numserie", mLetraSer, "T", , "codfaccl", vNumfactu, "N", "fecfaccl", vFecfactu, "F")
                End If
                If SQL <> "" Then
                    mCtaPrev = SQL
                    LeerDatosN = True
                Else
                    LeerDatosN = False
                    Mens = "La cuenta prevista de cobro para la factura no puede ser nula."
                End If
            Else
                LeerDatosN = True
            End If
        
          
        Else
            LeerDatosN = True
        End If
    End If

    RS.Close
    Set RS = Nothing
    Exit Function

ELeer:
    Mens = "Se ha producido un error. " & Mens & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Set RS = Nothing
    LeerDatosN = False
End Function


'======================================================================
'GRABAR EN TESORERIA
'======================================================================

Public Function InsertarEnTesoreria(vTextosCSB As String, MenError As String) As Boolean
'Guarda datos de Tesoreria en tablas: ariges.svenci y en conta.scobros
Dim b As Boolean
Dim RS As ADODB.Recordset
Dim rsVenci As ADODB.Recordset
Dim SQL As String, Codmacta As String, textcsb33 As String
Dim CadValues As String, cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAuxConta As String 'para insertar en conta.scobro
Dim CadValues3 As String
Dim FecVenci As Date, FecVenci1 As Date
Dim ImpVenci As Single 'importe para insertar en la svenci
Dim ImpVenci2 As Single 'importe para insertar en conta.scobro
Dim I As Byte
Dim TotalFactura2 As Currency   'Por si acaso lleva aportacion al terminal

Dim IbanAux As String
Dim InsertFinalContaNueva As String

    On Error GoTo EInsertarTesoreria



    If vTextosCSB = "" Then
        If vParamAplic.ContabilidadNueva Then
            vTextosCSB = "NULL"
        Else
            vTextosCSB = "NULL,NULL,NULL"
        End If
    Else
        
        If vParamAplic.ContabilidadNueva Then
            CadValues = Trim(RecuperaValor(vTextosCSB, 1))
            CadValues2 = Trim(RecuperaValor(vTextosCSB, 2))
            
                
            vTextosCSB = DBSet(Trim(Mid(CadValues & " " & CadValues2, 1, 60)), "T", "S")
        
        Else
            CadValues = DBSet(RecuperaValor(vTextosCSB, 1), "T", "S")
            CadValues2 = DBSet(RecuperaValor(vTextosCSB, 2), "T", "S")
            CadValues3 = DBSet(RecuperaValor(vTextosCSB, 3), "T", "S")
                     
            vTextosCSB = CadValues & "," & CadValues2 & "," & CadValues3
        End If
    End If
    b = False
    InsertarEnTesoreria = False
    CadValues3 = ""
    CadValues = ""
    CadValues2 = ""
    
    IbanAux = mIban
    If vParamAplic.ContabilidadNueva Then
        If mCodbanco <> "" And mCodsucur <> "" Then IbanAux = IbanAux & Right("0000" & mCodbanco, 4) & Right("0000" & mCodsucur, 4) & Right("00" & mDigContr, 2) & Right("0000000000" & mCuentaba, 10)            '
    
    
        'Trozo fianl insert cobros nueva conta
        ', nifclien,codusu,codpais,nomclien,domclien,pobclien, cpclien,proclien)"
        InsertFinalContaNueva = "," & DBSet(mNIFClien, "T", "S") & "," & vUsu.Id_Usu & "," & DBSet(mPais, "T", "S") & ","
        InsertFinalContaNueva = InsertFinalContaNueva & DBSet(mNomClien, "T", "S") & "," & DBSet(mDomClien, "T", "S") & ","
        InsertFinalContaNueva = InsertFinalContaNueva & DBSet(mPobClien, "T", "S") & "," & DBSet(mCodpobla, "T", "S")
        InsertFinalContaNueva = InsertFinalContaNueva & "," & DBSet(mProClien, "T", "S")

    End If

    
    
    
    'campo para insertar en conta.scobro de Tesoreria
    textcsb33 = "'FACTURA: " & LetraSerie & "-" & Format(NumFactu, "0000000") & " de Fecha " & Format(Fecfactu, "dd mmm yyyy") & "'"

    'Obtener el Nº de Vencimientos de la forma de pago
    SQL = "SELECT numerove, primerve, restoven FROM sforpa WHERE codforpa=" & ForPago
    Set rsVenci = New ADODB.Recordset
    rsVenci.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not rsVenci.EOF Then
        If rsVenci!numerove > 0 Then
            'Obtener los dias de pago del cliente
            SQL = " SELECT  diapago1, diapago2, diapago3, mesnogir, diavtoat, codmacta "
            If vParamAplic.ContabilidadNueva Then SQL = SQL & ",codpais"
            SQL = SQL & " FROM sclien "
            SQL = SQL & " WHERE codclien=" & Cliente
            Set RS = New ADODB.Recordset
            RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
            Codmacta = DBSet(RS!Codmacta, "T")
'            textcsb33 = "'FACTURA: " & LetraSerie & "-" & Format(NumFactu, "0000000") & " de Fecha " & Format(FecFactu, "dd,mm,yyyy") & "'"
            
            If Not RS.EOF Then
                cadValuesAux = "('" & Codtipom & "', " & NumFactu & ", '" & Format(Fecfactu, FormatoFecha) & "', "
                CadValuesAuxConta = "('" & LetraSerie & "', " & NumFactu & ", '" & Format(Fecfactu, FormatoFecha) & "', "
                '                    Añadire a la cadena fija esta los valores de textcsb41,txcs
                CadValuesAuxConta = CadValuesAuxConta & vTextosCSB & ","
                '-------- Primer Vencimiento
                I = 1
                'FECHA VTO
                FecVenci = CDate(Fecfactu)
                FecVenci = DateAdd("d", DBLet(rsVenci!primerve, "N"), FecVenci)
                '===
                'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                If TipForPago <> 0 Then
                    FecVenci = ComprobarFechaVenci(FecVenci, DBLet(RS!DiaPago1, "N"), DBLet(RS!DiaPago2, "N"), DBLet(RS!DiaPago3, "N"))
                Else
                    FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
                End If
                'Comprobar si cliente tiene mes a no girar
                FecVenci1 = FecVenci
                If CInt(DBLet(RS!mesnogir, "N")) <> 0 Then
                    FecVenci1 = ComprobarMesNoGira(FecVenci1, DBLet(RS!mesnogir, "N"), DBLet(RS!DiaVtoAt, "N"), DBLet(RS!DiaPago1, "N"), DBLet(RS!DiaPago2, "N"), DBLet(RS!DiaPago3, "N"))
                End If
                
                'Comprobar si cliente tiene dia de vencimiento atrasado
                CadValues = cadValuesAux & I & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                CadValues2 = CadValuesAuxConta & I & ", "
                CadValues2 = CadValues2 & Codmacta & ", " & ForPago & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                
                'IMPORTE del Vencimiento
                TotalFactura2 = TotalFac - Aportacion
                If rsVenci!numerove = 1 Then
                    ImpVenci = TotalFactura2
                    ImpVenci2 = TotalFactura2 - ImpCheque
                Else
                    
                    ImpVenci = Round2(TotalFactura2 / rsVenci!numerove, 2)
                    ImpVenci2 = Round2((TotalFactura2 - ImpCheque) / rsVenci!numerove, 2)
                    'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                    If ImpVenci * rsVenci!numerove <> TotalFactura2 Then
                        ImpVenci = Round(ImpVenci + (TotalFactura2 - ImpVenci * rsVenci!numerove), 2)
                    End If
                    'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                    If (ImpVenci2 * rsVenci!numerove) + ImpCheque <> TotalFactura2 Then
                        ImpVenci2 = Round(ImpVenci2 + (TotalFactura2 - ImpCheque - (ImpVenci2 * rsVenci!numerove)), 2)
                    End If
                End If
                CadValues = CadValues & DBSet(ImpVenci, "N") & ")"
                If vParamAplic.ContabilidadNueva Then
                    CadValues2 = CadValues2 & DBSet(ImpVenci2, "N") & ", '" & CuentaPrev & "'," & DBSet(IbanAux, "T", "S") & ", " & textcsb33 & ", " & DBSet(Agente, "N")
                        
                    'departamento y transfer
                    CadValues2 = CadValues2 & "," & DBSet(Me.DirDpto, "N", "S") & ",NULL" & InsertFinalContaNueva & ")"
                Else
                    CadValues2 = CadValues2 & DBSet(ImpVenci2, "N") & ", '" & CuentaPrev & "', " & DBSet(Banco, "N") & ", " & DBSet(Sucursal, "N") & ", " & DBSet(DigControl, "T") & ", " & DBSet(CuentaBan, "T") & "," & DBSet(iBan, "T", "S") & ", " & textcsb33 & ", " & DBSet(Agente, "N")
                        
                    'departamento y transfer
                    CadValues2 = CadValues2 & "," & DBSet(Me.DirDpto, "N", "S") & ",NULL)"
                End If
                
                
                'Resto Vencimientos
                '--------------------------------------------------------------------
                For I = 2 To rsVenci!numerove
                   'FECHA Resto Vencimientos
                    '=== Laura 23/01/2007
                    'FecVenci = FecVenci + DBSet(rsVenci!restoven, "N")
                    FecVenci = DateAdd("d", DBLet(rsVenci!restoven, "N"), FecVenci)
                    '===
                    'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                    If TipForPago <> 0 Then
                        FecVenci = ComprobarFechaVenci(FecVenci, DBLet(RS!DiaPago1, "N"), DBLet(RS!DiaPago2, "N"), DBLet(RS!DiaPago3, "N"))
                    Else
                        FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
                    End If
                    'Comprobar si cliente tiene mes a no girar
                    FecVenci1 = FecVenci
                    If DBLet(RS!mesnogir, "N") <> "0" Then
                        FecVenci1 = ComprobarMesNoGira(FecVenci1, DBLet(RS!mesnogir, "N"), DBLet(RS!DiaVtoAt, "N"), DBLet(RS!DiaPago1, "N"), DBLet(RS!DiaPago2, "N"), DBLet(RS!DiaPago3, "N"))
                    End If
                        
                    CadValues = CadValues & ", " & cadValuesAux & I & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                    CadValues2 = CadValues2 & ", " & CadValuesAuxConta & I & ", " & Codmacta & ", " & ForPago & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                    
                    'IMPORTE Resto de Vendimientos
                    ImpVenci = Round2(TotalFactura2 / rsVenci!numerove, 2)
                    ImpVenci2 = Round2((TotalFactura2 - ImpCheque) / rsVenci!numerove, 2)
                    CadValues = CadValues & DBSet(ImpVenci, "N") & ")"
                    
                    If vParamAplic.ContabilidadNueva Then
                        CadValues2 = CadValues2 & DBSet(ImpVenci2, "N") & ", '" & CuentaPrev & "'," & DBSet(IbanAux, "T", "S") & ", " & textcsb33 & ", " & DBSet(Agente, "N")
                            
                        'departamento y transfer
                        CadValues2 = CadValues2 & "," & DBSet(Me.DirDpto, "N", "S") & ",NULL" & InsertFinalContaNueva & ")"
                              
                    Else
                
                        CadValues2 = CadValues2 & DBSet(ImpVenci2, "N") & ", " & DBSet(CuentaPrev, "T") & ", " & DBSet(Banco, "N") & ", " & DBSet(Sucursal, "N") & ", " & DBSet(DigControl, "T") & ", " & DBSet(CuentaBan, "T") & "," & DBSet(iBan, "T", "S") & ", " & textcsb33 & ", " & DBSet(Agente, "N") & ", "
                        CadValues2 = CadValues2 & DBSet(Me.DirDpto, "N", "S") & ",NULL)"
                    End If
                Next I
                
                '--- Cheque regalo: laura 1/12/2006   y/o    Aportacion terminal
                'si hay cheque regalo insertar una linea más para la forma de pago correspondiente y el importe del cheque
                If ImpCheque > 0 Then
                
                    
                    CadValues2 = CadValues2 & ", " & CadValuesAuxConta & I & "," & Codmacta & ", " & vParamAplic.ForPagoChequeRegalo & ", "
                    
                    'FECHA VTO
                    CadValuesAuxConta = "primerve"
                    TipForPago = DevuelveDesdeBDNew(conAri, "sforpa", "tipforpa", "codforpa", vParamAplic.ForPagoChequeRegalo, "N", CadValuesAuxConta)
                    FecVenci = CDate(Fecfactu)
                    FecVenci = FecVenci + CByte(CadValuesAuxConta)
                    'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                    If TipForPago <> 0 Then
            '            FecVenci = ComprobarFechaVenci(FecVenci, DBLet(RS!DiaPago1, "N"), DBLet(RS!DiaPago2, "N"), DBLet(RS!DiaPago3, "N"))
                        MsgBox "FALTA cheque regalo con forma de pago no en EFECTIVO", vbInformation
                    Else
                        FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
                    End If
                    
                    CadValues2 = CadValues2 & DBSet(FecVenci, "F") & ", "
                    CadValues2 = CadValues2 & DBSet(ImpCheque, "N") & ", '" & CuentaPrev & "', " & DBSet(Banco, "N") & ", " & DBSet(Sucursal, "N") & ", " & DBSet(DigControl, "T") & ", " & DBSet(CuentaBan, "T") & ", " & textcsb33 & ", " & DBSet(Agente, "N")
                    'departamento
                    CadValues2 = CadValues2 & "," & DBSet(Me.DirDpto, "N", "S") & ",NULL)"
                End If
                
                
                'Aportacion al terminal
                'Si tiene cuenta aportacion entonces añadiremos un cobro
                If Me.Aportacion > 0 Then
                    If vParamAplic.ctaAportacion = "" Then
                        MsgBox "Error cta aportacion NULL", vbExclamation
                        Exit Function
                    End If
                        
                   'Montamos EL SQL para el cobro de la aportacion al termina
                    'ForPago
                    
                    CadValues2 = CadValues2 & ", " & CadValuesAuxConta & I & ",'" & vParamAplic.ctaAportacion & "', " & ForPago & ", "
                    
                    
                    CadValuesAuxConta = "primerve"
                    TipForPago = DevuelveDesdeBDNew(conAri, "sforpa", "tipforpa", "codforpa", ForPago, "N", CadValuesAuxConta)
                    FecVenci = CDate(Fecfactu)
                    FecVenci = FecVenci + CByte(CadValuesAuxConta)
                    'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                    If TipForPago <> 0 Then
            '            FecVenci = ComprobarFechaVenci(FecVenci, DBLet(RS!DiaPago1, "N"), DBLet(RS!DiaPago2, "N"), DBLet(RS!DiaPago3, "N"))
                        MsgBox "FALTA config con forma de pago no en EFECTIVO", vbInformation
                    Else
                        FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
                    End If
                    
                    CadValues2 = CadValues2 & DBSet(FecVenci, "F") & ", "
                    CadValues2 = CadValues2 & DBSet(Aportacion, "N") & ", '" & CuentaPrev & "', " & DBSet(Banco, "N") & ", " & DBSet(Sucursal, "N") & ", " & DBSet(DigControl, "T") & ", " & DBSet(CuentaBan, "T") & ", " & textcsb33 & ", " & DBSet(Agente, "N")
                    'departamento
                    CadValues2 = CadValues2 & "," & DBSet(Me.DirDpto, "N", "S") & ",NULL)"
                
                End If
                
            End If
        End If
        RS.Close
        Set RS = Nothing
    End If
    rsVenci.Close
    Set rsVenci = Nothing
    
    If CadValues <> "" Then
        SQL = "INSERT INTO svenci (codtipom, numfactu, fecfactu, ordefect, fecefect, impefect)"
        SQL = SQL & " VALUES " & CadValues
        conn.Execute SQL
    End If
    
    
    'Grabar tabla scobro de la CONTABILIDAD
    '-------------------------------------------------
    If CadValues2 <> "" Then
        '01/09/06
'        If (NumTicket = "") Or (NumTicket <> "" And TipForPago <> 0) Then
            If CuentaPrev <> "" Then
                'antes de grabar en la scobro comprobar que existe en conta.sforpa la
                'forma de pago de la factura. Sino existe insertarla
                'vemos si existe en la conta
                b = InsertarFormaPagoEnConta(ForPago, MenError)
                
                'si hay cheque regalo comprobar q existe su forma de pago en conta
                'sino insertarla
                If b And ImpCheque > 0 Then b = InsertarFormaPagoEnConta(vParamAplic.ForPagoChequeRegalo, MenError)
                
                
                
                
                
                If b Then
                    'Insertamos en la tabla scobro de la CONTA
                    If vParamAplic.ContabilidadNueva Then
                        SQL = "INSERT INTO cobros (numserie, numfactu, fecfactu,text41csb, "
                        SQL = SQL & "numorden , Codmacta, codforpa, FecVenci, ImpVenci, ctabanc1, "
                        SQL = SQL & "iban,text33csb,agente,departamento,transfer "
                        SQL = SQL & ", nifclien,codusu,codpais"
                        SQL = SQL & ",nomclien,domclien,pobclien, cpclien,proclien)"
                    Else
                        'Insertamos en la tabla scobro de la CONTA
                        SQL = " INSERT INTO scobro (numserie, codfaccl, fecfaccl,text41csb ,text42csb, text43csb, "
                        SQL = SQL & " numorden, codmacta, codforpa, fecvenci, impvenci,ctabanc1, codbanco, codsucur, digcontr, cuentaba,"
                        SQL = SQL & " iban,text33csb,agente,departamento,transfer)"
                    End If
                    SQL = SQL & " VALUES " & CadValues2
                    ConnConta.Execute SQL
                End If
                

                
                
                
                
                
            Else
                'DAVID ####
                'ENERO 2008
                'Si no inserto en tesoreria por que ctaprvsta="" entonces dejo continuar
                b = True
        
        End If





    End If
    
'    If b Then b = True
    
EInsertarTesoreria:
    If Err.Number <> 0 Then
        b = False
        MenError = "Insertar en Tesoreria: " & vbCrLf & Err.Description
    End If
    InsertarEnTesoreria = b
End Function





'###### Laura 15/11/2006 Recuperar facturas ALZIRA
Public Function InsertarEnTesoreria_RecupFac(MenError As String) As Boolean
'Guarda datos de Tesoreria en tablas: ariges.svenci y en conta.scobros
Dim b As Boolean
Dim RS As ADODB.Recordset
Dim rsVenci As ADODB.Recordset
Dim SQL As String, Codmacta As String, textcsb33 As String
Dim CadValues As String, cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim FecVenci As Date, FecVenci1 As Date
Dim ImpVenci As Single
Dim I As Byte


    On Error GoTo EInsertarTesoreria

    b = False
    InsertarEnTesoreria_RecupFac = False
    CadValues = ""
    CadValues2 = ""

    'Obtener el Nº de Vencimientos de la forma de pago
    SQL = "SELECT numerove, primerve, restoven FROM sforpa WHERE codforpa=" & ForPago
    Set rsVenci = New ADODB.Recordset
    rsVenci.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not rsVenci.EOF Then
        If rsVenci!numerove > 0 Then
            'Obtener los dias de pago del cliente
            SQL = " SELECT  diapago1, diapago2, diapago3, mesnogir, diavtoat, codmacta "
            SQL = SQL & " FROM sclien "
            SQL = SQL & " WHERE codclien=" & Cliente
            Set RS = New ADODB.Recordset
            RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
            Codmacta = DBSet(RS!Codmacta, "T")
         
            textcsb33 = "'FACTURA: " & LetraSerie & "-" & Format(NumFactu, "0000000") & " de Fecha " & Format(Fecfactu, "dd,mm,yyyy") & "'"
            
            If Not RS.EOF Then
               cadValuesAux = "('" & Codtipom & "', " & NumFactu & ", '" & Format(Fecfactu, FormatoFecha) & "', "
               CadValuesAux2 = "('" & LetraSerie & "', " & NumFactu & ", '" & Format(Fecfactu, FormatoFecha) & "', "

              'Primer Vencimiento
              '------------------------------------------------------------
                I = 1
                'FECHA VTO
                FecVenci = CDate(Fecfactu)
                FecVenci = FecVenci + CByte(DBLet(rsVenci!primerve, "N"))
                'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                If TipForPago <> 0 Then
                    FecVenci = ComprobarFechaVenci(FecVenci, DBLet(RS!DiaPago1, "N"), DBLet(RS!DiaPago2, "N"), DBLet(RS!DiaPago3, "N"))
                Else
                    FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
                End If
                'Comprobar si cliente tiene mes a no girar
                FecVenci1 = FecVenci
                If CInt(DBLet(RS!mesnogir, "N")) <> 0 Then
                    FecVenci1 = ComprobarMesNoGira(FecVenci1, DBLet(RS!mesnogir, "N"), DBLet(RS!DiaVtoAt, "N"), DBLet(RS!DiaPago1, "N"), DBLet(RS!DiaPago2, "N"), DBLet(RS!DiaPago3, "N"))
                End If
                
                'Comprobar si cliente tiene dia de vencimiento atrasado
                CadValues = cadValuesAux & I & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                CadValues2 = CadValuesAux2 & I & ", "
                CadValues2 = CadValues2 & Codmacta & ", " & ForPago & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                
                'IMPORTE del Vencimiento
                If rsVenci!numerove = 1 Then
                    ImpVenci = TotalFac
                Else
                    ImpVenci = Round(TotalFac / rsVenci!numerove, 2)
                    'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                    If ImpVenci * rsVenci!numerove <> TotalFac Then
'                        If ImpVenci * rsVenci!numerove <> TotalFac2 Then
                            ImpVenci = Round(ImpVenci + (TotalFac - ImpVenci * rsVenci!numerove), 2)
'                        End If
                    End If
                End If
                CadValues = CadValues & DBSet(ImpVenci, "N") & ")"
                CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", '" & CuentaPrev & "', " & DBSet(Banco, "N") & ", " & DBSet(Sucursal, "N") & ", " & DBSet(DigControl, "T") & ", " & DBSet(CuentaBan, "T") & ", " & textcsb33 & ", " & DBSet(Agente, "N")
                'departamento
                CadValues2 = CadValues2 & "," & DBSet(Me.DirDpto, "N", "S") & ")"
                
                
              'Resto Vencimientos
              '--------------------------------------------------------------------
                For I = 2 To rsVenci!numerove
                  'FECHA Resto Vencimientos
                    FecVenci = FecVenci + DBSet(rsVenci!restoven, "N")
                    'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                    If TipForPago <> 0 Then
                        FecVenci = ComprobarFechaVenci(FecVenci, DBLet(RS!DiaPago1, "N"), DBLet(RS!DiaPago2, "N"), DBLet(RS!DiaPago3, "N"))
                    Else
                        FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
                    End If
                    'Comprobar si cliente tiene mes a no girar
                    FecVenci1 = FecVenci
                    If DBLet(RS!mesnogir, "N") <> "0" Then
                        FecVenci1 = ComprobarMesNoGira(FecVenci1, DBLet(RS!mesnogir, "N"), DBLet(RS!DiaVtoAt, "N"), DBLet(RS!DiaPago1, "N"), DBLet(RS!DiaPago2, "N"), DBLet(RS!DiaPago3, "N"))
                    End If
                        
                    CadValues = CadValues & ", " & cadValuesAux & I & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                    CadValues2 = CadValues2 & ", " & CadValuesAux2 & I & ", " & Codmacta & ", " & ForPago & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                    
                  'IMPORTE Resto de Vendimientos
                    ImpVenci = Round(TotalFac / rsVenci!numerove, 2)
                    CadValues = CadValues & DBSet(ImpVenci, "N") & ")"
                    CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", '" & CuentaPrev & "', " & DBSet(Banco, "N") & ", " & DBSet(Sucursal, "N") & ", " & DBSet(DigControl, "T") & ", " & DBSet(CuentaBan, "T") & ", " & textcsb33 & ", " & DBSet(Agente, "N") & ", "
                    CadValues2 = CadValues2 & DBSet(Me.DirDpto, "N", "S") & ")"
                Next I
            End If
        End If
        RS.Close
        Set RS = Nothing
    End If
    rsVenci.Close
    Set rsVenci = Nothing
    
    If CadValues <> "" Then
        SQL = "INSERT INTO svenci (codtipom, numfactu, fecfactu, ordefect, fecefect, impefect)"
        SQL = SQL & " VALUES " & CadValues
        conn.Execute SQL
    End If
    
    
    'Grabar tabla scobro de la CONTABILIDAD
    '-------------------------------------------------
'    If CadValues2 <> "" Then
'        If CuentaPrev <> "" Then
'            'antes de grabar en la scobro comprobar que existe en conta.sforpa la
'            'forma de pago de la factura. Sino existe insertarla
'            'vemos si existe en la conta
'            CadValuesAux2 = DevuelveDesdeBDNew(conConta, "sforpa", "codforpa", "codforpa", ForPago, "N")
'            'si no existe la forma de pago en conta, insertamos la de ariges
'            If CadValuesAux2 = "" Then
'                cadValuesAux = "tipforpa"
'                CadValuesAux2 = DevuelveDesdeBDNew(conAri, "sforpa", "nomforpa", "codforpa", ForPago, "N", cadValuesAux)
'                'insertamos e sforpa de la CONTA
'                SQL = "INSERT INTO sforpa(codforpa,nomforpa,tipforpa)"
'                SQL = SQL & " VALUES(" & ForPago & ", " & DBSet(CadValuesAux2, "T") & ", " & cadValuesAux & ")"
'                ConnConta.Execute SQL
'            End If
'
'            'Insertamos en la tabla scobro de la CONTA
'            SQL = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci,ctabanc1, codbanco, codsucur, digcontr, cuentaba,text33csb,agente,departamento) "
'            SQL = SQL & " VALUES " & CadValues2
'            ConnConta.Execute SQL
'        End If
'    End If

    b = True
    
EInsertarTesoreria:
    If Err.Number <> 0 Then b = False
    InsertarEnTesoreria_RecupFac = b
End Function
















Private Sub ReiniciarImportesFactura()
    Me.BaseImp = 0
    Me.BaseIVA1 = 0
    Me.BaseIVA2 = 0
    Me.BaseIVA3 = 0
    
    Me.BrutoFac = 0
    
    Me.ImpIVA1 = 0
    Me.ImpIVA2 = 0
    Me.ImpIVA3 = 0
    
    Me.PorceIVA1 = 0
    Me.PorceIVA2 = 0
    Me.PorceIVA3 = 0
    
    
    Me.ImpIVA1RE = 0
    Me.ImpIVA2RE = 0
    Me.ImpIVA3RE = 0
    
    Me.PorceIVA1RE = 0
    Me.PorceIVA2RE = 0
    Me.PorceIVA3RE = 0
    
    
    Me.TipoIVA1 = 0
    Me.TipoIVA2 = 0
    Me.TipoIVA3 = 0
    
    Me.ImpGnral = 0
    Me.ImpPPago = 0
    Me.TotalFac = 0
    

End Sub


Public Function CalcularDatosFactura(cadWhere As String, NomTabla As String, NomTablaLin As String, CambiarIVA As Boolean) As Boolean
'cadWhere: cad para la where de la SQL que selecciona las lineas del albaran o la factura
'nomTabla: nombre de la tabla de albaranes(scaalp) o de AlbaranesXFactura(scafpa)
'           segun llamemos desde recepcion de facturas o desde Hco de Facturas
Dim RS As ADODB.Recordset
Dim I As Integer

Dim SQL As String
Dim cadAux As String

'Aqui vamos acumulando los totales
Dim TotBruto As Currency
Dim TotImpIVA As Currency

Dim ImpAux As Currency
Dim ImpIVA As Currency
Dim ImpBImIVA As Currency 'Importe Base imponible a la que hay q aplicar el IVA

Dim vBruto As Currency

Dim exento_IVA As Byte   ' 0.-IVA normal      1.- Recargo equivalencia    2.- Exento  3.-Exento INTRACOM
Dim conDesplaz As Boolean
    
    
Dim TipoIVAAplicable As Integer
Dim RE1 As String
Dim impRE As Currency
Dim TotRe As Currency


Dim ImporteParaComparar As Currency
Dim NueroIvas As Integer


    CalcularDatosFactura = False
    On Error GoTo ECalcular

    If Codtipom <> "FAM" Then ReiniciarImportesFactura

   '------------------------------------------------------------------------
   'En estos dos procedimientos solo entra cuando va a generar la factura y
   'codTipoMov="FAV", cuando llamamos desde pedidos, albaranes,etc, para mostrar
   'el resumen de la factura codTipoMov="ALV","PEV",... y no entra
   '-------------------------------------------------------------------------
   
   'comprobar si hay que facturar DESPLAZAMIENTOS
   '2012 Abril 18
   ' NO entramos aqui... NUNCA
   
        'CalcularDatosFactura = CalcularDesplazamiento(cadWhere, NomTabla)
        'If Not CalcularDatosFactura Then Exit Function
   
   
   
   'comprobar si hay que facturar BONIFICACIONES
    CalcularDatosFactura = CalcularBonificacion(cadWhere, NomTablaLin)
   If Not CalcularDatosFactura Then Exit Function
   '-------------------------------------------------------------------------
   
   'comprobar si el cliente de la factura esta EXENTO de IVA
   If Codtipom = "ALZ" Or Codtipom = "FAZ" Or Codtipom = "FAI" Or Codtipom = "ALI" Then
        'Albaranes de B. No llevan IVA
        exento_IVA = 2   'Exento IVA
   Else
        exento_IVA = ClienteExentoIVA2
   End If
   If exento_IVA = 2 Or exento_IVA = 3 Then
        cadWhere = Replace(cadWhere, NomTabla, NomTablaLin)
        SQL = "SELECT sum(importel) as bruto"
        SQL = SQL & " FROM " & NomTablaLin
        SQL = SQL & " WHERE " & cadWhere
   Else
    'Agrupar el importe bruto por tipos de iva
       If Codtipom = "FAM" Then
            'para los mantenimientos obtenemos de la tabla de tipos de contrato(stipco.codartic)
            'el articulo que se va a utilizar para facturar los mantenimientos
    '        CadAux = DevuelveDesdeBDNew(conAri, "stipco", "codartic", "codtipco", TipCoMan, "T")
            If ArticFac <> "" Then 'tenemos el articulo
                SQL = "SELECT codartic,nomartic,sartic.codigiva, tiposiva.porceiva "
                SQL = SQL & " FROM sartic INNER JOIN conta" & vParamAplic.NumeroConta & ".tiposiva ON sartic.codigiva=conta" & vParamAplic.NumeroConta & ".tiposiva.codigiva"
                SQL = SQL & " WHERE codartic=" & DBSet(ArticFac, "T")
            End If
       Else
            cadWhere = Replace(cadWhere, NomTabla, NomTablaLin)
            SQL = "SELECT sartic.codigiva,sum(importel) as bruto"
            SQL = SQL & " FROM " & NomTablaLin & " LEFT OUTER JOIN sartic ON " & NomTablaLin & ".codartic=sartic.codartic"
            SQL = SQL & " WHERE " & cadWhere
            SQL = SQL & " GROUP BY codigiva "
            SQL = SQL & " ORDER BY codigiva "
        End If
    End If
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenKeyset, adLockOptimistic, adCmdText
    NueroIvas = 0 'Numero IVAS
    ImporteParaComparar = 0
    
    'VEre cuantos IVAS tiene y el bruto. Para cuadrar los decimales
    If Codtipom = "FAM" Then
        'No hago la comrobacion para cuadre decimales
        NueroIvas = 1
    Else
            
        While Not RS.EOF
            TotBruto = DBLet(RS!bruto, "N")
            ImporteParaComparar = ImporteParaComparar + TotBruto
            NueroIvas = NueroIvas + 1
            RS.MoveNext
        Wend
    
    End If
    If NueroIvas > 0 Then
        TotImpIVA = CalcularPorcentaje(ImporteParaComparar, DtoPPago, 2)
        TotBruto = CalcularPorcentaje(ImporteParaComparar, DtoGnral, 2)
        ImporteParaComparar = ImporteParaComparar - TotImpIVA - TotBruto '(bruto - DtoPP -dtogral

        RS.MoveFirst
    End If
    TotBruto = 0
    TotImpIVA = 0
    TotRe = 0
    I = 1
    conDesplaz = False

    While Not RS.EOF
        'Aqui vamos acumulando la suma del importe bruto de las lineas
        If Codtipom = "FAM" Then
            vBruto = BrutoFac
        Else
            vBruto = RS!bruto
        End If
        
        TotBruto = TotBruto + vBruto
        ImpBImIVA = vBruto
        
        If exento_IVA < 2 Then     'Con IVA normal o RE
            'comprobar que si hay km de DESPLAZAMIENTO se añadan al tipo de IVA correspondiente
            If TotalKm > 0 Then
                If despIVA = RS!codigiva Then
                    vBruto = vBruto + ImporteL
                    TotBruto = TotBruto + ImporteL
                    conDesplaz = True
                End If
            End If
        
            'comprobar si hay BONIFICACIONES que aplicar a la Factura
            If IVABonif1 = RS!codigiva Then
                vBruto = vBruto - TotBonif1
                TotBruto = TotBruto - TotBonif1
            ElseIf IVABonif2 = RS!codigiva Then
                vBruto = vBruto - TotBonif2
                TotBruto = TotBruto - TotBonif2
            ElseIf IVABonif3 = RS!codigiva Then
                vBruto = vBruto - TotBonif3
                TotBruto = TotBruto - TotBonif3
            End If
            
        
            'Aplicarle el dto ppago
            '---- Laura: 05/10/2006
            ImpAux = CalcularPorcentaje(vBruto, DtoPPago, 2)
'            ImpAux = CCur(CalcularDto(CStr(vBruto), CStr(DtoPPago)))
            '---- Laura: 27/09/2006
'            ImpAux = Round(ImpAux, 2)
            '----
            ImpBImIVA = vBruto - ImpAux '(bruto - DtoPP)
            
            'Aplicarle el dto grnal
            '---- Laura: 05/10/2006
            ImpAux = CalcularPorcentaje(vBruto, DtoGnral, 2)
            ImpBImIVA = ImpBImIVA - ImpAux '(bruto - Dtogn)
            
            
            'NUEVO. 22 Mayo 2008
            'Si tiene mas de un tipo de iva, compruebo en la ultima de las bases
            'Si el total sera igual al total a comparar
            If I > 1 And I = NueroIvas Then
                'ImporteParaComparar
                If ImporteParaComparar - ImpBImIVA <> 0 Then
                    ImpBImIVA = ImporteParaComparar
                End If
            Else
                ImporteParaComparar = ImporteParaComparar - ImpBImIVA  'Para comparar en la ultima (si ha lugar)
            End If
            
            'Obtener el % de IVA
            cadAux = RS!codigiva    'Por si no tiene  RE o si teniendo RE no tiene el correspondiente
            If exento_IVA = 1 Then cadAux = CStr(vParamAplic.DevuleveTipoIVA_RE(RS!codigiva))
            TipoIVAAplicable = CInt(cadAux)
            
            
            'Sera viejo o nuevo
            
            
            
            
            ' If IVA_Antiguo And exento_IVA < 2 Then
            If CambiarIVA And exento_IVA < 2 Then
               
                    'JUNIO 2010
                    'Debido al cambio de IVA, los albaranes rectificativos podran ir en el IVA que llevaran en su momento
                    'Es decir, que aunque ahora sea al 18(a partir de Julio), para un albr rectifi podria ser
                    'con iva 16
                    cadAux = CStr(vParamAplic.DevuelveTipoIVA_OLD(exento_IVA = 1, RS!codigiva))
                    TipoIVAAplicable = CInt(cadAux)
            End If
            
            
            
            
            
            
            
            'ANTES
            'cadAux = DevuelveDesdeBDNew(conConta, "tiposiva", "porceiva", "codigiva", CStr(RS!Codigiva), "N")
            RE1 = "porcerec"
            cadAux = DevuelveDesdeBDNew(conConta, "tiposiva", "porceiva", "codigiva", cadAux, "N", RE1)
            If exento_IVA <> 1 Then
                RE1 = "0"
                impRE = 0
            Else
                impRE = CalcularPorcentaje(ImpBImIVA, CCur(RE1), 2)
                TotRe = TotRe + impRE
            End If
            

            
            
            'aplicar el IVA a la base imponible de ese tipo
            '---- Laura: 05/10/2006
            ImpIVA = CalcularPorcentaje(ImpBImIVA, CCur(cadAux), 2)
            
            'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
            'los vamos acumulando
            TotImpIVA = TotImpIVA + ImpIVA
            
            
            Select Case I
                Case 1  'IVA 1
                    TipoIVA1 = TipoIVAAplicable    'Antes Rs!codigiva
                    
                    BaseIVA1 = ImpBImIVA 'BASE IMPONIBLE
                    
                    PorceIVA1 = cadAux '% de IVA
                    
                    'Importe total con IVA
                    ImpIVA1 = ImpIVA
                    
                    'Recargo equivalencia
                    PorceIVA1RE = RE1
                    ImpIVA1RE = impRE
                     
                Case 2  'IVA 2
                    TipoIVA2 = TipoIVAAplicable     'Antes Rs!codigiva
                    
                    BaseIVA2 = ImpBImIVA 'BASE IMPONIBLE
                    
                    PorceIVA2 = cadAux '% de IVA
                    
                    'Importe total con IVA
                    ImpIVA2 = ImpIVA
    
                    'Recargo equivalencia
                    PorceIVA2RE = RE1
                    ImpIVA2RE = impRE
    
    
                Case 3  'IVA 3
                    TipoIVA3 = TipoIVAAplicable     'Antes Rs!codigiva
                    
                    BaseIVA3 = ImpBImIVA 'BASE IMPONIBLE
                    
                    PorceIVA3 = cadAux '% de IVA
                    
                    'Importe total con IVA
                    ImpIVA3 = ImpIVA
                    
                    'Recargo equivalencia
                    PorceIVA3RE = RE1
                    ImpIVA3RE = impRE
                    
            End Select
            I = I + 1
        Else 'EXENTO DE IVA
            'si hay que factura KM añadirlos al bruto
            If TotalKm > 0 Then vBruto = vBruto + ImporteL
            
            If exento_IVA = 2 Then
                'Exento normal
                If vParamAplic.IVA_Exento2 = 0 Then MsgBox "Falta configurar el IVA exento", vbExclamation
                TipoIVA1 = vParamAplic.IVA_Exento2
            Else
                'exento intracom
                If vParamAplic.IVA_Intracomunitario = 0 Then MsgBox "Falta configurar el IVA intracomunitario", vbExclamation
                TipoIVA1 = vParamAplic.IVA_Intracomunitario
                
            End If
        
        
        
        
                ImpAux = CalcularPorcentaje(vBruto, DtoPPago, 2)
                ImpBImIVA = vBruto - ImpAux '(bruto - DtoPP)

                ImpAux = CalcularPorcentaje(vBruto, DtoGnral, 2)
                ImpBImIVA = ImpBImIVA - ImpAux '(bruto - Dtogn)
                
        
        
        
                
                BaseIVA1 = ImpBImIVA 'BASE IMPONIBLE
                
                PorceIVA1 = 0 '% de IVA
                
                'Importe total con IVA
                ImpIVA1 = 0
                
                'Recargo equivalencia
                PorceIVA1RE = 0
                ImpIVA1RE = 0

        End If
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing

    'si tiene desplazamiento comprobar que se ha añadido en el IVA
    'si no calcularlo
    If TotalKm > 0 And conDesplaz = False Then
    
        'FALTA##### con la conversion del Tipo de iva a RE
    
        TotBruto = TotBruto + ImporteL
        vBruto = ImporteL
        ImpBImIVA = ImporteL
        
        'Aplicarle el dto ppago
        '---- Laura: 05/10/2006
        ImpAux = CalcularPorcentaje(vBruto, DtoPPago, 2)
'        ImpAux = CCur(CalcularDto(CStr(vBruto), CStr(DtoPPago)))
        ImpBImIVA = ImpBImIVA - ImpAux '(bruto - DtoPP)
            
        'Aplicarle el dto grnal
        '---- Laura: 05/10/2006
        ImpAux = CalcularPorcentaje(vBruto, DtoGnral, 2)
'        ImpAux = CCur(CalcularDto(CStr(vBruto), CStr(DtoGnral)))
        ImpBImIVA = ImpBImIVA - ImpAux '(bruto - Dtogn)
            
        'Obtener el % de IVA
        cadAux = DevuelveDesdeBDNew(conConta, "tiposiva", "porceiva", "codigiva", CStr(despIVA), "N")
            
        'cuota IVA: aplicar el IVA a la base imponible de ese tipo
        'Laura: 05/10/2006
        ImpAux = CalcularPorcentaje(ImpBImIVA, CCur(cadAux), 2)
'        ImpAux = CalcularDto(CStr(ImpBImIVA), cadAux)
        ImpIVA = ImpAux
            
        'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
        'los vamos acumulando
        TotImpIVA = TotImpIVA + ImpIVA
        
        Select Case I
            Case 2 'IVA 2
                TipoIVA2 = despIVA
                
                BaseIVA2 = ImpBImIVA 'BASE IMPONIBLE
                
                PorceIVA2 = cadAux '% de IVA
                
                'Importe total con IVA
                ImpIVA2 = ImpIVA
                
            Case 3 'IVA 3
                TipoIVA3 = RS!codigiva
                
                BaseIVA3 = ImpBImIVA 'BASE IMPONIBLE
                
                PorceIVA3 = cadAux '% de IVA
                
                'Importe total con IVA
                ImpIVA3 = ImpIVA
        End Select
    End If
    
    BrutoFac = TotBruto
    
    'Aplicarle el dto ppago
    '---- Laura: 05/10/2006
    ImpPPago = CalcularPorcentaje(TotBruto, DtoPPago, 2)
'    ImpPPago = CCur(CalcularDto(CStr(TotBruto), CStr(DtoPPago)))
    '---- Laura: 27/09/2006
'    ImpPPago = Round(ImpPPago, 2)
    '----
    
    'Aplicarle el dto general
    '---- Laura: 05/10/2006
    ImpGnral = CalcularPorcentaje(TotBruto, DtoGnral, 2)
'    ImpGnral = CCur(CalcularDto(CStr(TotBruto), CStr(DtoGnral)))
    '---- Laura: 27/09/2006
'    ImpGnral = Round(ImpGnral, 2)
    '----
    
    
    'Base Imponible
    BaseImp = TotBruto - ImpPPago - ImpGnral
    
    'TOTAL de la factura
    TotalFac = BaseImp + TotImpIVA + TotRe
    
    CalcularDatosFactura = True
    
ECalcular:
    If Err.Number <> 0 Then
        CalcularDatosFactura = False
    Else
        CalcularDatosFactura = True
    End If
End Function



Public Function PasarAlbaranesAFactura2(TipoAlb As String, cadSQL As String, textoCSB As String, Optional ErroresAux As String, Optional EstaRecu As Boolean) As Boolean
'IN ->  cadSQL: cadena para seleccion de los Albaranes que vamos a Facturar
'       TextosCSB:  Para la tesoreria
'Desde Albaranes Genera las Facturas correspondientes
Dim b As Boolean
Dim MenError As String

    On Error GoTo ETraspasoAlbFac

    PasarAlbaranesAFactura2 = False
    b = False
    ErroresAux = ""
    '#Laura 14/11/2006 recuperar facturas de Alzira
    'si esta recuperando facturas
    EstaRecuFac = EstaRecu
    
    Select Case TipoAlb
        Case "ALV", "ALM": 'ALV: Albaranes venta a clientes
                    'ALM: Albaran mostrador
            Codtipom = "FAV"
        Case "ALR": 'Albaranes de reparacion en clientes
            Codtipom = "FAR"
        Case "ART": 'Albaranes de factura rectificativa
            Codtipom = "FRT"
        Case "ALS": 'Albaranes de servicios [SERVICIOS]
            Codtipom = "FAS"
        Case "ALZ"  'Albaranes "presupuestos". Es decir, el "B"
            Codtipom = "FAZ"
        Case "ALI"
            Codtipom = "FAI"
    End Select
      
     
    OpeFactu = PonerTrabajadorConectado("")
    If OpeFactu <> "" Then
        AlmFactu = DevuelveDesdeBDNew(conAri, "straba", "codalmac", "codtraba", OpeFactu, "N")
    Else
        AlmFactu = "1"
    End If
      
    'Aqui empieza transaccion
    conn.BeginTrans
    ConnConta.BeginTrans
   
    
    MenError = "Error al Insertar en tablas de factura."
    b = InsertarFactura(cadSQL)
    
    
    'Insertar los Vencimientos de la Factura. Tabla: svenci
    'Grabar en TESORERIA. Tabla de Contabilidad: sconta.scobros
    '#Laura: 14/11/2006 Recuperar facturas Alzira
    If b Then
        If EstaRecuFac = False Then
            MenError = "Error al actualizar en Tesoreria."
            b = InsertarEnTesoreria(textoCSB, MenError)
        Else
            MenError = "Error al actualizar en Tesoreria."
            b = InsertarEnTesoreria_RecupFac(MenError)
        End If
    End If
    
    'Eliminar los Albaranes incluidos en la Factura
    If b Then
        MenError = "Error al eliminar los albaranes."
        b = EliminarAlbaranes(cadSQL)
    End If
    
    'Si alguna linea de albaran tiene asociados num. serie actualizar
    'el campo numfactu de dichos num serie
    If b Then
        MenError = "Error al actualizar los Nº de serie."
        b = ActualizarNumSeries(cadSQL)
    End If
    
    
ETraspasoAlbFac:
    If Err.Number <> 0 Then
        ErroresAux = Err.Description
        Err.Clear
        b = False
    End If
  
    If b Then
        conn.CommitTrans
        ConnConta.CommitTrans
        PasarAlbaranesAFactura2 = True
    Else
        conn.RollbackTrans
        ConnConta.RollbackTrans
        PasarAlbaranesAFactura2 = False
        ErroresAux = MenError & vbCrLf & ErroresAux & vbCrLf
    End If
    Espera 0.4
End Function



Public Function PasarMtosAFactura(TipCo As String, Operador As String, mes As String, nMante As String) As Boolean
'IN -> cadSQL: cadena para seleccion de los Mantenimientos que vamos a Facturar
'Desde Mantenimientos Genera las Facturas correspondientes
Dim b As Boolean
Dim MenError As String

    On Error GoTo ETraspasoFac

    PasarMtosAFactura = False
    b = False
      
    Codtipom = "FAM"
    
    TipCoMan = TipCo
    OpeFactu = Operador
    MesFactu = mes
    numMante = nMante
    
    ArticFac = DevuelveDesdeBDNew(conAri, "stipco", "codartic", "codtipco", TipCoMan, "T")
    If ArticFac = "" Then
        MenError = "El tipo de contrato no tiene articulo para facturar."
        Exit Function
    End If
      
    
    If OpeFactu <> "" Then
        AlmFactu = DevuelveDesdeBDNew(conAri, "straba", "codalmac", "codtraba", OpeFactu, "N")
    Else
        AlmFactu = "1"
    End If
    
      
    'Aqui empieza transaccion
    conn.BeginTrans
    ConnConta.BeginTrans
    
       
    'Insertar la Factura (scafac,scafac1,slifac)
    MenError = "Error al Insertar en tablas de factura"
    b = InsertarFactura("")
    
    
    If b Then
        'Insertar en la Tabla para detalle equipos, para reimprimir
       MenError = "Error al insertar detalle de equipos."
       InsertarDetEquip
    End If
    
    
    'Insertar los Vencimientos de la Factura. Tabla: svenci
    'Grabar en TESORERIA. Tabla de Contabilidad: sconta.scobros
    If b Then
        MenError = "Error al actualizar en Tesoreria"
        b = InsertarEnTesoreria("", MenError)
    End If
    
   
   
    'Actualizo en la tabla mantenimientos la fecha ultimio mantenimiento
    If b Then
        MenError = "Actualizar fecha ultimo mantenimiento"
        ArticFac = "UPDATE scaman SET ulmesfac = " & mes & " WHERE codclien = " & Cliente & " AND nummante= '" & numMante & "'"
        conn.Execute ArticFac
    End If
   
ETraspasoFac:
    If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        MuestraError Err.Number, "Pasando Mantenimientos a Factura.", Err.Description
        b = False
    End If
    If b Then
        conn.CommitTrans
        ConnConta.CommitTrans
        PasarMtosAFactura = True
    Else
        conn.RollbackTrans
        ConnConta.RollbackTrans
        PasarMtosAFactura = False
    End If
    Espera 0.2
End Function



Public Function PasarTicketAFactura(cadSQL As String, Optional ErroresAux As String, Optional NTicket As String, Optional NAlbTicket As String, Optional impRegalo As String) As Boolean
'IN -> cadSQL: cadena para seleccion de las ventas del ticket que vamos a Facturar
'Desde ventas Genera las Facturas de ticket correspondiente
Dim b As Boolean
Dim MenError As String

    On Error GoTo ETrasTicFac

    OpeFactu = ErroresAux 'Aqui hemos pasado el trabajador
    NumTicket = NTicket
    NumAlbTicket = NAlbTicket '01/09/06 laura
    ImpCheque = CCur(ComprobarCero(impRegalo)) 'importe del cheque regalo (si hay)
    
    PasarTicketAFactura = False
    b = False
    ErroresAux = ""
      
      
    'por defecto se coge el almacen con el q trabaja el empleado
    If OpeFactu <> "" Then
        AlmFactu = DevuelveDesdeBDNew(conAri, "straba", "codalmac", "codtraba", OpeFactu, "N")
    Else
        AlmFactu = "1"
    End If
    
    
    'Insertar la Factura (scafac,scafac1,slifac)
    MenError = "Error al Insertar en tablas de factura."
    b = InsertarFacturaTicket(cadSQL)
    
    
    'Insertar los Vencimientos de la Factura. Tabla: svenci
    'Grabar en TESORERIA. Tabla de Contabilidad: sconta.scobros
    If b Then
        MenError = "Error al actualizar en Tesoreria."
        b = InsertarEnTesoreria("", MenError)
    End If
    
    
    'Eliminar los Albaranes incluidos en la Factura
    If b Then
        MenError = "Error al eliminar la venta."
        b = EliminarVenta(cadSQL)
    End If
    
    
    
ETrasTicFac:
    If Err.Number <> 0 Then
        ErroresAux = Err.Description
        Err.Clear
        b = False
    End If
  
    If Not b Then ErroresAux = MenError & vbCrLf & ErroresAux & vbCrLf
    PasarTicketAFactura = b
End Function



Private Function InsertarFactura(cadSQL As String, Optional cadNumFac As String, Optional HacerCambiosIVA As Boolean) As Boolean
'Insertamos en las tablas de factura de cliente: scafac, scafac1, slifac
Dim b As Boolean
Dim vTipoMov As CTiposMov
Dim Existe As Boolean
Dim Devuelve As String

    InsertarFactura = False
    
    
    'Obtener el Contador de FACTURA
    '--------------------------------------------------------------
    Set vTipoMov = New CTiposMov
        
    If vTipoMov.Leer(Codtipom) Then
        '#Laura: 14/11/2006 Recuperar facturas ALZIRA
        If EstaRecuFac = False Then
            'Comprobar si mientras tanto se incremento el contador de Facturas
            'para ello vemos si existe una Factura con ese contador y si existe la incrementamos
            Do
                NumFactu = vTipoMov.ConseguirContador(Codtipom)
                Devuelve = DevuelveDesdeBDNew(conAri, "scafac", "numfactu", "codtipom", Codtipom, "T", , "numfactu", CStr(NumFactu), "N", "fecfactu", Fecfactu, "F")
                If Devuelve <> "" Then
                    'Ya existe el contador incrementarlo
                    Existe = True
                    vTipoMov.IncrementarContador (Codtipom)
                    NumFactu = vTipoMov.ConseguirContador(Codtipom)
                Else
                    Existe = False
                End If
            Loop Until Not Existe
        End If
    Else
        Exit Function
    End If
        
    LetraSerie = vTipoMov.LetraSerie 'la guardamos para INSERT en scobro de la Conta
    
    'calcular el total de la factura antes de insertar
    b = CalcularDatosFactura(cadSQL, "scaalb", "slialb", HacerCambiosIVA)
    If Not b Then Exit Function
    
    '### Laura: 14/11/2006 Recuperar facturas Alzira
    Devuelve = ""
    If EstaRecuFac = True Then
        'comprobar q el total de la factura a insertar coincide con el total de la
        'factura en conta.cabfact si esta existe en conta
        Devuelve = "SELECT COUNT(*) FROM cabfact WHERE numserie=" & DBSet(Me.LetraSerie, "T")
        Devuelve = Devuelve & " AND codfaccl=" & Me.NumFactu & " AND anofaccl=" & Year(Fecfactu)
        If (RegistrosAListar(Devuelve, conConta) > 0) Then
            Devuelve = DevuelveDesdeBDNew(conConta, "cabfact", "totfaccl", "numserie", Me.LetraSerie, "T", , "codfaccl", Me.NumFactu, "N", "anofaccl", Year(Me.Fecfactu), "N")
            If Devuelve <> "" Then
                If CCur(Devuelve) <> Me.TotalFac Then
                    If MsgBox("El total de la factura no coincide con el total en la Contabilidad" & vbCrLf & "¿Desea continuar?", vbQuestion + vbYesNo) = vbNo Then
                        b = False
                    End If
                End If
            End If
            Devuelve = "EXISTE"
        End If
        If Not b Then Exit Function
    End If
    '#####
    
    'Cabecera Factura
    b = InsertarCabeceraFactu
    If Not b Then Exit Function
    
    'Cabeceras Albaranes de la Factura
    b = InsertarCabAlbaranesFactu(cadSQL)
    If Not b Then Exit Function
    
    'Insertar lineas de la factura
    b = InsertarLineasFactu(cadSQL)
    If Not b Then Exit Function
    
    'Comprobar si los albaranes tienen km de desplazamiento e insertar
    'lineas de factura del desplazamiento e incrementar el bruto de la factura
    'Insertar lineas de desplazamiento
    b = InsertarDesplaz
    If Not b Then Exit Function
    
    'si hay bonificaciones en la factura
    If SQLBonif <> "" Then
        b = InsertarBonificaciones
        If Not b Then Exit Function
    End If
    
   
    '#Laura: 14/11/2006 Recuperar facturas Alzira
    If EstaRecuFac = True And Devuelve = "EXISTE" Then
        b = MarcarContabilizada
        If Not b Then Exit Function
    End If
    
    
    'Incrementar contador del tipo de movimiento
    '-------------------------------------------------------------
    '#Laura: 14/11/2006 Recuperar facturas Alzira
    If EstaRecuFac = False Then vTipoMov.IncrementarContador (Codtipom)
    Set vTipoMov = Nothing
      
    InsertarFactura = True
End Function




Private Function InsertarFacturaTicket(cadSQL As String) As Boolean
'Insertamos en las tablas de factura de cliente: scafac, scafac1, slifac
'Es un ticket de venta
Dim b As Boolean
Dim Devuelve As String

    InsertarFacturaTicket = False
    
    
    'Calcular los campos de la factura y cuadrarla al
    'pasar las lineas sin IVA
    b = CalcularDatosFacturaTicket(cadSQL)
    If Not b Then Exit Function

    
    'Cabecera Factura
    b = InsertarCabeceraFactu
    If Not b Then Exit Function
    
    'Cabeceras Albaranes de la Factura
    b = InsertarCabAlbaranesFactu(cadSQL)
    If Not b Then Exit Function

    'Insertar lineas de la factura
    b = InsertarLineasFactu(cadSQL)
    If Not b Then Exit Function
    
       
'    'Incrementar contador del tipo de movimiento
'    '-------------------------------------------------------------
'    vTipoMov.IncrementarContador (codTipoM)
'    Set vTipoMov = Nothing
      
    InsertarFacturaTicket = True
End Function




Private Function InsertarCabeceraFactu() As Boolean
'Inserta la cabecera de la factura en la tabla: scafac
Dim SQL As String
Dim vClien As CCliente
Dim Nulo2 As String
Dim Nulo3 As String
Dim ctaDpto As Boolean
Dim Cad As String, cad2 As String

    On Error GoTo EInsertar

    
    'Insertar en la tabla cabecera de la factura de compras
    SQL = "INSERT INTO scafac (codtipom,numfactu,fecfactu,codclien,nomclien,domclien,codpobla,pobclien,proclien,nifclien,telclien,coddirec,nomdirec,"
    SQL = SQL & "codagent,codforpa,dtoppago,dtognral,codbanco,codsucur,digcontr,cuentaba,brutofac,impdtopp,impdtogr,baseimp1,baseimp2,baseimp3, "
    SQL = SQL & "codigiv1,codigiv2,codigiv3,porciva1,porciva2,porciva3,imporiv1,imporiv2,imporiv3,totalfac,intconta "
            ' Añadimos IVA recargo de equivalencia                               02 Junio 08
    SQL = SQL & ",porciva1re,imporiv1re, porciva2re,imporiv2re, porciva3re,imporiv3re,aportacion,aridoc,iban)"
            
    
    
    SQL = SQL & " VALUES (" & DBSet(Codtipom, "T") & "," & DBSet(NumFactu, "N") & "," & DBSet(Fecfactu, "F") & "," & DBSet(Cliente, "N") & ","
    SQL = SQL & DBSet(NombreClien, "T") & "," & DBSet(DomicilioClien, "T", "N") & "," & DBSet(CPostal, "T", "N") & "," & DBSet(Poblacion, "T", "N") & "," & DBSet(Provincia, "T", "N") & ","
    SQL = SQL & DBSet(NIF, "T", "N") & "," & DBSet(Telefono, "T") & "," & DBSet(DirDpto, "N", "S") & "," & DBSet(NombreDirDpto, "T") & "," & DBSet(Agente, "N") & ","
    SQL = SQL & ForPago & "," & DBSet(DtoPPago, "N") & "," & DBSet(DtoGnral, "N") & ","
    
    If Codtipom = "FAV" Or Codtipom = "FRT" Or Codtipom = "FAR" Or Codtipom = "FAS" Or Codtipom = "FAI" Then
        'si son facturas de VENTA o Factura Rectificativa
        'Cargar los datos del banco del cliente
        'si se trabaja con departamentos y la factura va a un departamento
        'comprobar si tiene banco y asignar el del departamento
        ctaDpto = False
        If vParamAplic.Departamento Then
            If Me.DirDpto <> "" Then
                'comprobar que el dpto del cliente tiene cta bancaria
                cad2 = "digcontr"
                Cad = DevuelveDesdeBDNew(conAri, "sdirec", "cuentaba", "codclien", Me.Cliente, "N", cad2, "coddirec", Me.DirDpto, "N")
                If Cad <> "" And cad2 <> "" Then ctaDpto = True
            End If
        End If
        
        If ctaDpto Then 'cogemos la cta bancaria del departamento
            Me.CuentaBan = Cad
'            cad = DevuelveDesdeBDNew(conAri, "sdirec", "digcontr", "codclien", Me.Cliente, "N", , "coddirec", Me.DirDpto, "N")
            Me.DigControl = cad2
            cad2 = "codsucur"
            Cad = DevuelveDesdeBDNew(conAri, "sdirec", "codbanco", "codclien", Me.Cliente, "N", cad2, "coddirec", Me.DirDpto, "N")
            Me.Banco = Cad
            Me.Sucursal = cad2
            Me.iBan = ""
        Else 'cuenta bancaria del cliente
            Set vClien = New CCliente
            If vClien.LeerDatos(Cliente) Then
                Banco = vClien.Banco
                Sucursal = vClien.Sucursal
                DigControl = vClien.DigControl
                CuentaBan = vClien.CuentaBan
                iBan = vClien.iBan
                mPais = vClien.Pais
                
            Else
                Banco = ValorNulo
                Sucursal = ValorNulo
                DigControl = ValorNulo
                CuentaBan = ValorNulo
                iBan = ValorNulo
                mPais = "ES"
            End If
            Set vClien = Nothing
        End If
    End If
    
    SQL = SQL & DBSet(Banco, "N") & "," & DBSet(Sucursal, "N") & "," & DBSet(DigControl, "T") & "," & DBSet(CuentaBan, "T") & ","
    
    SQL = SQL & DBSet(BrutoFac, "N") & ","
    SQL = SQL & DBSet(ImpPPago, "N") & "," & DBSet(ImpGnral, "N") & ","
    SQL = SQL & DBSet(BaseIVA1, "N") & "," & DBSet(BaseIVA2, "N", "S") & "," & DBSet(BaseIVA3, "N", "S") & ","
    Nulo2 = "N"
    Nulo3 = "N"
    If BaseIVA2 = 0 Then Nulo2 = "S"
    If BaseIVA3 = 0 Then Nulo3 = "S"
    SQL = SQL & DBSet(TipoIVA1, "N") & "," & DBSet(TipoIVA2, "N", Nulo2) & "," & DBSet(TipoIVA3, "N", Nulo3) & ","
    SQL = SQL & DBSet(PorceIVA1, "N") & "," & DBSet(PorceIVA2, "N", Nulo2) & "," & DBSet(PorceIVA3, "N", Nulo3) & ","
     
    SQL = SQL & DBSet(ImpIVA1, "N") & "," & DBSet(ImpIVA2, "N", Nulo2) & "," & DBSet(ImpIVA3, "N", Nulo3) & ","
     
    SQL = SQL & DBSet(TotalFac, "N") & ",0"
    
    'AÑADIMOS EL recargo de equivalencia en el INSERT  porciva1re,imporiv1re .......
    SQL = SQL & "," & DBSet(PorceIVA1RE, "N", "S") & "," & DBSet(ImpIVA1RE, "N", "S")
    SQL = SQL & "," & DBSet(PorceIVA2RE, "N", "S") & "," & DBSet(ImpIVA2RE, "N", "S")
    SQL = SQL & "," & DBSet(PorceIVA3RE, "N", "S") & "," & DBSet(ImpIVA3RE, "N", "S") & ","
    
    'aportacion,aridoc
    SQL = SQL & DBSet(mAportacion, "N", "N") & ",NULL,"
    'Dic2013
    SQL = SQL & DBSet(mIban, "T", "S") & ")"
    
    conn.Execute SQL
    InsertarCabeceraFactu = True
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabeceraFactu = False
        MuestraError Err.Number, "Insertar cabecera factura.", Err.Description
    Else
        InsertarCabeceraFactu = True
    End If
End Function


Private Function InsertarCabAlbaranesFactu(cadSQL As String) As Boolean
'Insertar las cabeceras de los Albaranes de la factura en la tabla: scafac1
'NumTicket: para cuando de venta generamos Factura en numalbar guardamos el ticket que la genero
Dim SQL As String
Dim Devuelve As String
Dim MantenimientoObservacion As String

    On Error GoTo EInsertarAlb

    Devuelve = ""
    If Codtipom = "FTI" Then
'        cadSQL = Replace(cadSQL, "sliven", "scaven")
    Else
        If InStr(cadSQL, "scaven") > 0 Then 'viene de generar factura en ticket
'            codTipoM = "FTI"
        Else
            cadSQL = Replace(cadSQL, "slialb", "scaalb")
        End If
    End If
    
    
    SQL = "INSERT INTO scafac1 (codtipom,numfactu,fecfactu,codtipoa,numalbar,fechaalb,numpedcl,fecpedcl,sementre,numofert,fecofert,referenc,codenvio,codtraba,codtrab1,codtrab2,observa1,observa2,observa3,observa4,observa5,numtermi,numventa,observa6"
    'Junio16
    SQL = SQL & ",FechaCarga , Muestra, Deposito, TransEmpresa, TransMatricula, TransConductor, TransCondDNI, TransNumBocas, TransBruto,"
    SQL = SQL & "TransTara, TransObsPrecintos, TransMatRemolque, TransMercancia, TransAcidez, TransDestino, TransLacradasCoop,"
    SQL = SQL & "TransLacradasCompr, TransTicketBas, TransCMR, TransCertLim, TransOtros,hora)"
    
    If (Codtipom = "FAV" And NumTicket <> "") Then
        'Facturas generadas desde una venta de tickets: en numalbar guardamos numticket
'        SQL = SQL & " VALUES('" & codTipoM & "', " & NumFactu & ", '" & Format(FecFactu, FormatoFecha) & "','FTI'," & NumTicket & ","
        '01/09/06 laura
        SQL = SQL & " VALUES('" & Codtipom & "', " & NumFactu & ", '" & Format(Fecfactu, FormatoFecha) & "','ATI'," & NumTicket & ","
        
        SQL = SQL & "'" & Format(Fecfactu, FormatoFecha) & "'," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," 'datos de pedido a nulo
        SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," 'datos de oferta a nulo
        'obtenemos el codigo de envio de cliente
        Devuelve = DevuelveDesdeBDNew(conAri, "sclien", "codenvio", "codclien", Cliente, "N")
        SQL = SQL & DBSet(Devuelve, "N") & ","
        'trabajadores
        SQL = SQL & DBSet(OpeFactu, "N") & "," & ValorNulo & "," & DBSet(OpeFactu, "N") & ","
        'observaciones
        '---- Laura 10/10/06: poner en la observacion 1 el valor de scaven.observa1
        Devuelve = DevuelveDesdeBDNew(conAri, "scaven", "observa1", "numtermi", Me.NumTerminal, "N", , "numventa", Me.NumVenta, "N", "fecventa", Me.Fecfactu, "F")
        SQL = SQL & DBSet(Devuelve, "T") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
        
        'Dic 2009.  observa6
        'SQL = SQL & NumTerminal & "," & NumVenta & ")"
        SQL = SQL & NumTerminal & "," & NumVenta & ",NULL"
        
        SQL = SQL & ",FechaCarga , Muestra, Deposito, TransEmpresa, TransMatricula, TransConductor, TransCondDNI, TransNumBocas, TransBruto,"
        SQL = SQL & "TransTara, TransObsPrecintos, TransMatRemolque, TransMercancia, TransAcidez, TransDestino, TransLacradasCoop,"
        SQL = SQL & "TransLacradasCompr, TransTicketBas, TransCMR, TransCertLim, TransOtros,hora)"
        
    ElseIf (Codtipom <> "FAM" And Codtipom <> "FTI") And cadSQL <> "" Then 'Facturas de VEnta
        'Dic 2009
        'SQL = SQL & " SELECT '" & codTipoM & "' as codtipom," & DBSet(NumFactu, "N") & " as numfactu," & DBSet(FecFactu, "F") & " as fecfactu," & "codtipom as codtipoa, numalbar,fechaalb,numpedcl,fecpedcl,sementre,numofert,fecofert,referenc,codenvio,codtraba,codtrab1,codtrab2,observa01,observa02,observa03,observa04,observa05,numtermi,numventa "
        SQL = SQL & " SELECT '" & Codtipom & "' as codtipom," & DBSet(NumFactu, "N") & " as numfactu," & DBSet(Fecfactu, "F") & " as fecfactu," & "codtipom as codtipoa, numalbar,fechaalb,numpedcl,fecpedcl,sementre,numofert,fecofert,referenc,codenvio,codtraba,codtrab1,codtrab2,observa01,observa02,observa03,observa04,observa05,numtermi,numventa,observa6 "
        
        
        SQL = SQL & ",FechaCarga , Muestra, Deposito, TransEmpresa, TransMatricula, TransConductor, TransCondDNI, TransNumBocas, TransBruto,"
        SQL = SQL & "TransTara, TransObsPrecintos, TransMatRemolque, TransMercancia, TransAcidez, TransDestino, TransLacradasCoop,"
        SQL = SQL & "TransLacradasCompr, TransTicketBas, TransCMR, TransCertLim, TransOtros,hora"
        
        
        SQL = SQL & " FROM scaalb WHERE " & cadSQL
        
        
    Else 'Bonificaciones/Desplazamientos
        'Facturas de Mantenimiento
        'Facturas de Ticket del TPV
        SQL = SQL & " VALUES('" & Codtipom & "', " & NumFactu & ", '" & Format(Fecfactu, FormatoFecha) & "',"
'        If (codTipoM = "FAM" Or codTipoM = "FTI") Then
        '01/09/06 laura
        If (Codtipom = "FAM") Then
            SQL = SQL & "'',0,"
        ElseIf Codtipom = "FTI" Then
            SQL = SQL & "'ATI'," & NumAlbTicket & ","
        Else
            SQL = SQL & "'',9999999,"
        End If
        SQL = SQL & "'" & Format(Fecfactu, FormatoFecha) & "'," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," 'datos de pedido a nulo
        SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," 'datos de oferta a nulo
        'obtenemos el codigo de envio de cliente
        Devuelve = DevuelveDesdeBDNew(conAri, "sclien", "codenvio", "codclien", Cliente, "N")
        SQL = SQL & DBSet(Devuelve, "N") & ","
        'trabajadores
        SQL = SQL & DBSet(OpeFactu, "N") & "," & ValorNulo & "," & DBSet(OpeFactu, "N") & ","
        'observaciones 1
        If Codtipom = "FTI" Then
            '---- Laura 19/12/06: poner en la observacion 1 el valor de scaven.observa1
            Devuelve = DevuelveDesdeBDNew(conAri, "scaven", "observa1", "numtermi", Me.NumTerminal, "N", , "numventa", Me.NumVenta, "N", "fecventa", Me.Fecfactu, "F")
        Else
            'La linea viene en el campo concefac que hemos guardado en observaciones
            If Codtipom = "FAM" Then
                
                MantenimientoObservacion = DevuelveDesdeBDNew(conAri, "scaman", "tipopago", "codclien", Me.Cliente, "N", , "nummante", numMante, "T")
                If MantenimientoObservacion = "1" Then 'Trimestral
                    MantenimientoObservacion = "TRIMESTRAL DE " & Year(Fecfactu)
                ElseIf MantenimientoObservacion = "2" Then 'Semestral
                    MantenimientoObservacion = "SEMESTRAL DE " & Year(Fecfactu)
                ElseIf MantenimientoObservacion = "3" Then 'Anual
                    MantenimientoObservacion = "ANUAL " & Year(Fecfactu)
                Else
                    'si es mensual: ej: "JUNIO de 2005" (ponemos el mes seleccionado para facturar)
                    MantenimientoObservacion = MonthName(CLng(MesFactu)) & " de " & Year(Fecfactu)
                    MantenimientoObservacion = UCase(MantenimientoObservacion)
                End If
                                                            
                If mObservacion <> "" Then
                    Devuelve = MantenimientoObservacion
                    
                Else
                    Devuelve = ""
                    mObservacion = MantenimientoObservacion
                End If
                
            Else
                Devuelve = ""
            End If
        End If
        SQL = SQL & DBSet(Devuelve, "T") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo
        'DIC 2009
        'SQL = SQL & "," & DBSet(NumTerminal, "N", "S") & "," & DBSet(NumVenta, "N", "S") & ")"
        SQL = SQL & "," & DBSet(NumTerminal, "N", "S") & "," & DBSet(NumVenta, "N", "S") & ",NULL"
        
        SQL = SQL & ",NULL , NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL,"
        SQL = SQL & "NULL, NULL, NULL, NULL, NULL, NULL, NULL,"
        SQL = SQL & "NULL, NULL, NULL, NULL, NULL,NULL)"
        
        
        
    End If
    
    conn.Execute SQL

    InsertarCabAlbaranesFactu = True
    
EInsertarAlb:
    If Err.Number <> 0 Then
        InsertarCabAlbaranesFactu = False
        MuestraError Err.Number, "Insertar cabecera albaranes factura.", Err.Description
    Else
        InsertarCabAlbaranesFactu = True
    End If
End Function


Private Function InsertarLineasFactu(cadSQL As String) As Boolean
'Insertar las lineas de los Albaranes de la factura en la tabla: slifac
Dim SQL As String
Dim Cad As String
Dim cad2 As String
Dim cArt As CArticulo
Dim InsertaTbLotes As Boolean

    On Error GoTo EInsertarLin

    '##########################################
    'Cambios para llevar la TRAZA con codprove
    '   si cojo slialb:  tengo la columna
    '   RESTO DE CASOS!!!!! FALTA###
    

    'Otra modificacion:   16 Mayo 08
    ' Se permiten dtos, con lo cual la donde ponia 0 as dto ya no vale
    
    'Mod  11 Junio 2009.  Tendremos que insertar tb en slifaclotes
    
    SQL = "INSERT INTO slifac (codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel,origpre"
    If NumTicket <> "" Then SQL = SQL & ",precioiv"
    SQL = SQL & ",preciomp,preciost,preciouc,codprovex,palets,hectogrado)"     'ENERO 2008.  Traza sobre codprove. palets y Nov2011 hectogrado
    
    InsertaTbLotes = False
    
    If Codtipom = "FAM" Then 'facturas de mantenimientos
        SQL = SQL & " VALUES('" & Codtipom & "'," & NumFactu & ",'" & Format(Fecfactu, FormatoFecha) & "','',0,1,"
            
        SQL = SQL & DBSet(AlmFactu, "N") & ","
            
        'Articulo para facturar mantenimientos de la tabla de tipos contrato: stipco.codartic
        SQL = SQL & DBSet(ArticFac, "T") & ", "
            
        'nombre del articulo para facturar,preciomp,preciost, preciouc
        If ArticFac <> "" Then
            Set cArt = New CArticulo
            If cArt.LeerDatos(ArticFac) Then
                Cad = cArt.Nombre   'Aunque luego machaco cad con la nueva amplicacion: contrato + numerocontrato
            End If
'            cad2 = "preciomp"
'            cad = DevuelveDesdeBDNew(conAri, "sartic", "nomartic", "codartic", ArticFac, "T", cad2)
        End If
        
        Cad = DevuelveDesdeBDNew(conAri, "stipco", "nomtipco", "codtipco", TipCoMan, "T")
        Cad = Cad & "  Nº: " & numMante
        
        SQL = SQL & DBSet(Cad, "T") & ","
            
        'campo Ampliacion linea segun el tipopago del mantenimientos
        '-----------------------
'        'ESTO YA LO HACE EN LA insercion de scafac1.  Con lo cual aqui pondremos lo que haya en obsraciones
'       ------------------------
'        cad = DevuelveDesdeBDNew(conAri, "scaman", "tipopago", "codclien", Me.Cliente, "N", , "nummante", numMante, "T")
'        If cad = "1" Then 'Trimestral
'            cad = "TRIMESTRAL DE " & Year(FecFactu)
'        ElseIf cad = "2" Then 'Semestral
'            cad = "SEMESTRAL DE " & Year(FecFactu)
'        ElseIf cad = "3" Then 'Anual
'            cad = "ANUAL " & Year(FecFactu)
'        Else
'            'si es mensual: ej: "JUNIO de 2005" (ponemos el mes seleccionado para facturar)
'            cad = MonthName(CLng(MesFactu)) & " de " & Year(FecFactu)
'            cad = UCase(cad)
'        End If
        Cad = mObservacion
        
        SQL = SQL & DBSet(Cad, "T") & ","
            
        'cantidad, precio, dtos
        SQL = SQL & "1," & DBSet(BrutoFac, "N") & ", 0,0, "
        'importe, origen precio
        SQL = SQL & DBSet(BrutoFac, "N") & ", 'M',"
        
        'guardar los precios del articulo
        SQL = SQL & DBSet(cArt.PrecioMedPon, "N") & "," & DBSet(cArt.PrecioStan, "N") & "," & DBSet(cArt.PrecioUltCom, "N") & ","
        'traza ENERO 2008     Palets MAYO 2011 y hecto 2011
        SQL = SQL & DBSet(cArt.codProve, "N") & ",0,1)"
            
        
        Set cArt = Nothing
        
    ElseIf (Codtipom = "FTI") Then 'factura de ticket
        'se entra en esta opcion cuando desde el TPV generamos un ticket
        SQL = SQL & " SELECT '" & Codtipom & "' as codtipom," & DBSet(NumFactu, "N") & " as numfactu," & DBSet(Fecfactu, "F") & " as fecfactu, 'ATI' as codtipoa," & NumAlbTicket & " as numalbar," & "numlinea,"
        SQL = SQL & AlmFactu & " as codalmac," & "sliven.codartic,sliven.nomartic," & ValorNulo & " as ampliaci,cantidad,precioar,"
        'David 16 Mayo 08
        '0 as dtoline1,0 as dtoline2, cantidad*precioar as importel,'' as origpre,precioiv "
        SQL = SQL & " dto1,dto2, implineareal,'' as origpre,precioiv "
        
         '-- Laura 15/09/2006: añadir precios del articulo a las lineas de factura
        SQL = SQL & ", sartic.preciomp,sartic.preciost,sartic.preciouc,sartic.codprove,0 palets ,1 hectog"
        '--
        SQL = SQL & " FROM sliven "
        '-- Laura 15/09/2006: añadir precios del articulo a las lineas de factura
        SQL = SQL & " INNER JOIN sartic ON sliven.codartic=sartic.codartic "
        '--
        SQL = SQL & " WHERE " & Replace(cadSQL, "scaven", "sliven")
        
    ElseIf NumTicket <> "" Then 'Ticket de venta genera factura
        SQL = SQL & " SELECT '" & Codtipom & "' as codtipom," & DBSet(NumFactu, "N") & " as numfactu," & DBSet(Fecfactu, "F") & " as fecfactu, 'ATI' as codtipoa," & NumTicket & " as numalbar," & "numlinea,"
        SQL = SQL & AlmFactu & " as codalmac," & "sliven.codartic,sliven.nomartic," & ValorNulo & " as ampliaci,cantidad,precioar,"
        
        'David 16 Mayo 08
        '0 as dtoline1,0 as dtoline2, cantidad*precioar as importel,'' as origpre,precioiv "
        SQL = SQL & " dto1,dto2, implineareal,'' as origpre,precioiv "
        
        
        
        '-- Laura 15/09/2006: añadir precios del articulo a las lineas de factura
        SQL = SQL & ", sartic.preciomp,sartic.preciost,sartic.preciouc,sartic.codprove,0 palets ,1 hectog"
        '--
        SQL = SQL & " FROM sliven "
        '-- Laura 15/09/2006: añadir precios del articulo a las lineas de factura
        SQL = SQL & " INNER JOIN sartic ON sliven.codartic=sartic.codartic "
        '--
        SQL = SQL & " WHERE " & Replace(cadSQL, "scaven", "sliven")
    
    Else 'pasamos lineas de albaranes a lineas factura
        SQL = SQL & " SELECT '" & Codtipom & "' as codtipom," & DBSet(NumFactu, "N") & " as numfactu," & DBSet(Fecfactu, "F") & " as fecfactu," & "codtipom as codtipoa, numalbar,numlinea,codalmac,slialb.codartic,slialb.nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel,origpre "
        '-- Laura 15/09/2006: añadir precios del articulo a las lineas de factura 'ENERO 2008. Traza cond codprove
        SQL = SQL & ", sartic.preciomp,sartic.preciost,sartic.preciouc,codprovex,palets,hectogrado" 'NAYO 2011  palets  Nov hectogrado
        '--
        SQL = SQL & " FROM slialb "
        '-- Laura 15/09/2006: añadir precios del articulo a las lineas de factura
        SQL = SQL & " INNER JOIN sartic ON slialb.codartic=sartic.codartic "
        '--
        SQL = SQL & " WHERE " & Replace(cadSQL, "scaalb", "slialb")
        InsertaTbLotes = True 'Para que luego haga el insert de lotes
        
    End If
    
    conn.Execute SQL

    If InsertaTbLotes Then
        'Insertara tm en lotes InsertaTbLotes desde slialblotes
        SQL = "INSERT INTO `slifaclotes` (`codtipom`,`numfactu`,`fecfactu`,`codtipoa`,`numalbar`,`numlinea`,`linea`,`numlote`,cantidad)"
        SQL = SQL & " SELECT '" & Codtipom & "' as codtipom," & DBSet(NumFactu, "N") & " as numfactu,"
        SQL = SQL & DBSet(Fecfactu, "F") & " as fecfactu," & "codtipom as codtipoa, numalbar,numlinea,linea,numlote,cantidad"
        SQL = SQL & " FROM slialblotes "
        SQL = SQL & " WHERE " & Replace(cadSQL, "scaalb", "slialblotes")
        EjecutaSQL conAri, SQL, True
    End If
    InsertarLineasFactu = True
    
EInsertarLin:
    If Err.Number <> 0 Then
        InsertarLineasFactu = False
        MuestraError Err.Number, "Insertar lineas factura.", Err.Description
    Else
        InsertarLineasFactu = True
    End If
End Function


Private Function EliminarAlbaranes(cadSQL As String) As Boolean
'Eliminamos de las tablas de Albaranes: scaalb, slialb
Dim SQL As String

    On Error GoTo EEliminar

    EliminarAlbaranes = False
    
    'Eliminamos los lotes
    SQL = "DELETE FROM slialblotes "
    SQL = SQL & " WHERE " & Replace(cadSQL, "scaalb", "slialblotes")
    conn.Execute SQL
    
    'ELiminar lineas albaranes
    SQL = "DELETE FROM slialb "
    SQL = SQL & " WHERE " & Replace(cadSQL, "scaalb", "slialb")
    conn.Execute SQL
    
    Espera 0.1
    
    'Eliminar Cabeceras Albaranes
    SQL = "DELETE FROM scaalb "
    SQL = SQL & " WHERE " & Replace(cadSQL, "slialb", "scaalb")
    conn.Execute SQL
        
    EliminarAlbaranes = True

EEliminar:
    If Err.Number <> 0 Then
        EliminarAlbaranes = False
    Else
        EliminarAlbaranes = True
    End If
End Function



Private Function InsertarDetEquip() As Boolean
Dim SQL As String
Dim Cad As String

    'si son facturas de MAntenimiento y parametro de stipco: detallar equipos en la factura=true
    'guardar en una tabla el resumen por tipo de articulo de los Nº Series de equipos
    'asociados al MAntenimiento
    InsertarDetEquip = False
    
    On Error GoTo EDetalle
       
    SQL = DevuelveDesdeBDNew(conAri, "stipco", "detequip", "codtipco", TipCoMan, "T")
    
    If SQL = "1" Then 'Detalla Equipo
        Cad = ObtenerSQLcomponentes(" WHERE codclien=" & Cliente & " AND nummante=" & DBSet(numMante, "T"))
        SQL = "Select distinct '" & Codtipom & "' as codtipom," & NumFactu & " as numfactu,'" & Format(Fecfactu, FormatoFecha) & "' as fecfactu,"
        Cad = Replace(Cad, "Select distinct", SQL)
        SQL = "INSERT INTO srecom (codtipom,numfactu,fecfactu,codtipar,nomtipar,cantidad) "
        SQL = SQL & Cad

        conn.Execute SQL
    End If
    
EDetalle:
    If Err.Number <> 0 Then
        InsertarDetEquip = False
    Else
        InsertarDetEquip = True
    End If
End Function


'ANTES:  ClienteExentoIVA  as boolean
'Ahora:  ClienteExentoIVA2 as byte
'           0.- Exento
'           1.- RE
'           2.- EXNTO
Private Function ClienteExentoIVA2() As Byte
Dim vClien As CCliente

    Set vClien = New CCliente
    vClien.Codigo = Cliente
    ClienteExentoIVA2 = vClien.exentoIVA2
    Set vClien = Nothing
    
    'ClienteExentoIVA = b
End Function



Private Function InsertarDesplaz() As Boolean
'Si los albaranes a facturar tienen KM de desplazamiento añadirlos a la factura
'Añadir cantkm * preciokm al bruto de la factura
'Insertar linea en tabla slifac con el importe del desplazamiento
Dim SQL As String
Dim Cad As String
Dim b As Boolean

    On Error GoTo EDesplaz
    InsertarDesplaz = False
    
    If TotalKm > 0 Then
        'Insertar en scafac1 un albaran a 9999999, para el desplazamiento (y servira tb para bonificaciones)
        '-----------------------------------------------------------
        b = InsertarCabAlbaranesFactu("")
        If Not b Then Exit Function
            
               
        'Insertar linea de Desplazamiento en tabla: slifac (lineas factura)
        '------------------------------------------------------------------
        SQL = "INSERT INTO slifac (codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,"
        SQL = SQL & "cantidad,precioar,dtoline1,dtoline2,importel,origpre) "
        'codtipom,numfactu,fecfactu,codtipoa,numalbar
        SQL = SQL & " VALUES (" & DBSet(Codtipom, "T") & "," & NumFactu & "," & DBSet(Fecfactu, "F") & ",'',9999999,"
        'numlinea
        SQL = SQL & "1,"
        SQL = SQL & DBSet(AlmFactu, "N") & ","
            
        'codartic,nomartic,ampliaci
        Cad = DevuelveDesdeBDNew(conAri, "sartic", "nomartic", "codartic", codartic, "T")
        SQL = SQL & DBSet(codartic, "T") & "," & DBSet(Cad, "T") & "," & DBSet(Ampliacion, "T") & ","
        'cantidad,precioar,dtoline1,dtoline2,importel,origpre
        SQL = SQL & DBSet(TotalKm, "N") & "," & DBSet(PrecioKm, "N") & ",0,0," & DBSet(ImporteL, "N") & ",'M')"
        
        conn.Execute SQL
        
    End If
    InsertarDesplaz = True
     
EDesplaz:
    If Err.Number <> 0 Then
        InsertarDesplaz = False
    Else
        InsertarDesplaz = True
    End If
End Function


Private Function CalcularDesplazamiento(cadWhere As String, NomTabla As String) As Boolean
'comprobar si los albaranes a facturar tienen DESPLAZAMIENTOS
'calcula en vbles: Totalkm, Ampliacion los valores para insertar en linea factura
Dim RS As ADODB.Recordset
Dim SQL As String
    
    CalcularDesplazamiento = False
    On Error GoTo EDesp
    
    If Codtipom = "FAV" Or Codtipom = "FAR" Then
        'Obtenemos el total de km a facturar en los albaranes seleccionados
        SQL = "SELECT fechaalb,sum(cantidkm)" ',sum(sclien.kilometr) "
        SQL = SQL & " FROM " & NomTabla '& " INNER JOIN sclien ON " & NomTabla & ".codclien=sclien.codclien "
        SQL = SQL & " WHERE " & cadWhere '& " AND facturkm=1 "
        SQL = SQL & " GROUP BY fechaalb"
        SQL = SQL & " ORDER BY fechaalb"
        
        Set RS = New ADODB.Recordset
        RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        TotalKm = 0
        Ampliacion = ""
        While Not RS.EOF
            If RS.Fields(1).Value <> 0 Then
                TotalKm = TotalKm + RS.Fields(1).Value
                Ampliacion = Ampliacion & Day(RS.Fields(0).Value) & ", "
            End If
            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
        If Ampliacion <> "" Then 'quitar la ultima coma
            Ampliacion = Mid(Ampliacion, 1, Len(Ampliacion) - 2)
            Ampliacion = "DIAS: " & Ampliacion
        End If
        
        If TotalKm > 0 Then
            'obtenemos el precio del km de los parametros, y el cod Articulo
            codartic = "codartid"
            SQL = DevuelveDesdeBDNew(conAri, "spara1", "preukmcl", "codigo", "1", "N", codartic)
            PrecioKm = CCur(ComprobarCero(SQL))
            ImporteL = Round(TotalKm * PrecioKm, 2)
            
            'obtenemos el tipo IVA del articulo de desplazamiento
            SQL = DevuelveDesdeBDNew(conAri, "sartic", "codigiva", "codartic", codartic, "T")
            despIVA = CInt(ComprobarCero(SQL))
        End If
    End If
    CalcularDesplazamiento = True
    
EDesp:
    If Err.Number <> 0 Then
        CalcularDesplazamiento = False
    Else
        CalcularDesplazamiento = True
    End If
End Function



Private Function CalcularBonificacion(cadWhere As String, NomTablaLin As String) As Boolean
'comprobar los articulos de los albaranes a facturar tienen BONIFICACION
Dim RS As ADODB.Recordset
Dim RSaux As ADODB.Recordset
Dim SQL As String, cadIVA As String
Dim vClien As CCliente
Dim EsExcluyente As Boolean
Dim Bonif1 As Currency, Bonif2 As Currency
Dim Canti1 As Single, Canti2 As Single
Dim numlinea As Integer

    CalcularBonificacion = False
    On Error GoTo EBonif
    
    If Codtipom = "FAV" Then 'Factura de Venta
        'Obtener el cod. tarifa del cliente al que se va a facturar y
        'comprobar si para este tipo de tarifa se aplica Bonificacion
        Set vClien = New CCliente
        vClien.Codigo = Cliente
        SQLBonif = ""
        TotBonif1 = 0
        TotBonif2 = 0
        TotBonif3 = 0
        IVABonif1 = 0
        IVABonif2 = 0
        IVABonif3 = 0
            
        If vClien.Bonifica Then
            'Comprobar para los codartic de las lineas de albaran que vamos
            'a factura, para cuales hay bonificacion y la cantidad de la linea
            'cumple la cantidad para que se le aplica la bonificacion
            SQL = "SELECT slialb.codartic, sum(slialb.cantidad) as cantidad "
            SQL = SQL & "FROM " & NomTablaLin
            SQL = SQL & " INNER JOIN sbonif ON " & NomTablaLin & ".codartic=sbonif.codartic "
            SQL = SQL & " WHERE " & Replace(cadWhere, "scaalb", NomTablaLin)
            SQL = SQL & " GROUP BY " & NomTablaLin & ".codartic "
            SQL = SQL & " order by " & NomTablaLin & ".codartic "
            
           
            Set RS = New ADODB.Recordset
            RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            numlinea = 1
            
            While Not RS.EOF 'para cada codartic
                'Para cada linea de slialb a la que se le aplica bonificacion
                'obtener la bonif. que se le va a aplicar

                'mirar primero si hay algun hastaca2<=cantidad
                '--------------------------------------------------
                SQL = "SELECT * from sbonif "
                SQL = SQL & " WHERE codartic=" & DBSet(RS!codartic, "T") '& " AND (" & RS!Cantidad & "<=sbonif.hastaca1)"
                SQL = SQL & " order by hastaca1 desc "
                Set RSaux = New ADODB.Recordset
                RSaux.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                Bonif1 = 0
                Bonif2 = 0
                If Not RSaux.EOF Then
'                    If (1 <= Rs!Cantidad) And (Rs!Cantidad <= RSaux!hastaca1) Then
                    If (0 < RS!Cantidad) And (RS!Cantidad <= RSaux!hastaca1) Then
                        Bonif1 = RSaux!impboni1
                        Canti1 = RS!Cantidad
                    End If
                    If (RSaux!hastaca1 < RS!Cantidad) And (RS!Cantidad <= RSaux!hastaca2) Then
                        Bonif1 = RSaux!impboni1
                        Canti1 = RS!Cantidad
                        Bonif2 = RSaux!impboni2
                        Canti2 = RS!Cantidad
                    End If
                    
                    
                    'comprobar si la bonificacion es excluyente
                    EsExcluyente = CBool(RSaux!excluyen)
                    
                    If EsExcluyente Then
                    'se aplicara una bonificacion u otra
                       If Bonif2 > 0 Then Bonif1 = 0
                    End If
                    
                    If Bonif1 > 0 Then
                        numlinea = numlinea + 1
                        SQLBonif = SQLBonif & "('" & Codtipom & "'," & NumFactu & "," & DBSet(Fecfactu, "F") & ",'',9999999," & numlinea & "," & AlmFactu & "," & DBSet(RSaux!codarti1, "T") & ","
                        cadIVA = "codigiva"
                        SQL = DevuelveDesdeBDNew(conAri, "sartic", "nomartic", "codartic", RSaux!codarti1, "T", cadIVA)
                        SQLBonif = SQLBonif & DBSet(SQL, "T") & ",''," & DBSet(Canti1, "N") & "," & DBSet(-Bonif1, "N") & ",0,0," & DBSet(0 - Round(Canti1 * Bonif1, 2), "N") & ",'M'),"
                        If IVABonif1 = 0 Or IVABonif1 = CByte(cadIVA) Then
                            IVABonif1 = CByte(cadIVA)
                            TotBonif1 = TotBonif1 + Round(Canti1 * Bonif1, 2)
                        ElseIf IVABonif2 = 0 Or IVABonif2 = CByte(cadIVA) Then
                            IVABonif2 = CByte(cadIVA)
                            TotBonif2 = TotBonif2 + Round((Canti1 * Bonif1), 2)
                        ElseIf IVABonif3 = 0 Or IVABonif3 = CByte(cadIVA) Then
                            IVABonif3 = CByte(cadIVA)
                            TotBonif3 = TotBonif3 + Round(Canti1 * Bonif1, 2)
                        End If
                    End If
                    
                    If Bonif2 > 0 Then
'                        SQL = "(codtipom='" & codTipoM & "' and numfactu=" & NumFactu & " and fecfactu=" & DBSet(FecFactu, "F") & " and codtipoa='' and numalbar=0)"
                        numlinea = numlinea + 1
                        SQLBonif = SQLBonif & "('" & Codtipom & "'," & NumFactu & "," & DBSet(Fecfactu, "F") & ",'',9999999," & numlinea & "," & AlmFactu & "," & DBSet(RSaux!codarti2, "T") & ","
                        SQL = DevuelveDesdeBDNew(conAri, "sartic", "nomartic", "codartic", RSaux!codarti2, "T")
                        SQLBonif = SQLBonif & DBSet(SQL, "T") & ",''," & DBSet(Canti2, "N") & "," & DBSet(-Bonif2, "N") & ",0,0," & DBSet(0 - Round(Canti2 * Bonif2, 2), "N") & ",'M'),"
                        If IVABonif1 = 0 Or IVABonif1 = CByte(cadIVA) Then
                            IVABonif1 = CByte(cadIVA)
                            TotBonif1 = TotBonif1 + Round((Canti2 * Bonif2), 2)
                        ElseIf IVABonif2 = 0 Or IVABonif2 = CByte(cadIVA) Then
                            IVABonif2 = CByte(cadIVA)
                            TotBonif2 = TotBonif2 + Round((Canti2 * Bonif2), 2)
                        ElseIf IVABonif3 = 0 Or IVABonif3 = CByte(cadIVA) Then
                            IVABonif3 = CByte(cadIVA)
                            TotBonif3 = TotBonif3 + Round((Canti2 * Bonif2), 2)
                        End If
                    End If
'                    RSaux.MoveNext
                End If
                RSaux.Close
                Set RSaux = Nothing
                RS.MoveNext
            Wend
            RS.Close
            Set RS = Nothing
            If SQLBonif <> "" Then 'quitamos ultima coma
                SQLBonif = Mid(SQLBonif, 1, Len(SQLBonif) - 1)
            End If
        End If
    End If
    Set vClien = Nothing
    CalcularBonificacion = True
    
    
EBonif:
    If Err.Number <> 0 Then
        CalcularBonificacion = False
    Else
        CalcularBonificacion = True
    End If
End Function



Private Function InsertarBonificaciones() As Boolean
'Si los albaranes a facturar tienen lineas de articulos con bonificacion
'añadirlos a la factura
'Insertar lineas en tabla slifac con el importe de las bonificaciones
Dim SQL As String
Dim b As Boolean

    On Error GoTo EBonifica

    InsertarBonificaciones = False
    
    If SQLBonif > "" Then
        'Insertar en scafac1 un albaran a 0, para la bonificacion
        'si no se inserto antes para los desplazamientos
        '-----------------------------------------------------------
        SQL = "(codtipom='" & Codtipom & "' and numfactu=" & NumFactu & " and fecfactu=" & DBSet(Fecfactu, "F") & " and codtipoa='' and numalbar=9999999)"
        SQL = "Select count(*) from scafac1 where " & SQL
        If RegistrosAListar(SQL) = 0 Then
            b = InsertarCabAlbaranesFactu("")
            If Not b Then Exit Function
        End If
               
        'Insertar lineas de Desplazamiento en tabla: slifac (lineas factura)
        '------------------------------------------------------------------
        SQL = "INSERT INTO slifac (codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,"
        SQL = SQL & "cantidad,precioar,dtoline1,dtoline2,importel,origpre) "
        SQL = SQL & " VALUES " & SQLBonif
        conn.Execute SQL
        
    End If
    InsertarBonificaciones = True
     
EBonifica:
    If Err.Number <> 0 Then
        InsertarBonificaciones = False
    Else
        InsertarBonificaciones = True
    End If
End Function




Private Function ActualizarNumSeries(cadSQL As String) As Boolean
'Actualiza el campo "numfactu" de la tabla de num. serie: sserie
'Para ponerle la factura asociada
Dim SQL As String

    On Error GoTo EActualizar

    SQL = "UPDATE sserie SET numfactu= " & NumFactu
    SQL = SQL & " WHERE " & Replace(cadSQL, "scaalb", "sserie")
    conn.Execute SQL

    ActualizarNumSeries = True
    
EActualizar:
    If Err.Number <> 0 Then
        ActualizarNumSeries = False
    Else
        ActualizarNumSeries = True
    End If
End Function



Private Function CalcularDatosFacturaTicket(cadSQL As String) As Boolean
Dim b As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
Dim I As Byte
Dim ImpTotal As Currency
Dim ImpDif As Currency

    On Error GoTo ETicket
    
'    'Bruto de la factura sera la suma de cantidad*precioar de las linea
'    'importe sin iva
'    SQL = "SELECT sum(cantidad*precioar) FROM sliven WHERE "
'    SQL = SQL & Replace(cadSQL, "scaven", "sliven")
'    Set RS = New ADODB.Recordset
'    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    If Not RS.EOF Then
'        BrutoFac = Round(RS.Fields(0).Value, 2)
'        BaseImp = BrutoFac
'    End If
'    RS.Close
'    Set RS = Nothing
    
    
    
    'desglosar las bases imponibles por tipo de IVA
'    SQL = "SELECT codigiva,sum(cantidad*precioar) as baseimp from sliven "
    '01/09/06
    SQL = "SELECT codigiva,sum(importel) as totalimp from sliven "
    
    SQL = SQL & " WHERE " & Replace(cadSQL, "scaven", "sliven")
    SQL = SQL & " group by codigiva"
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 1
    While Not RS.EOF And I <= 3
        SQL = DevuelveDesdeBDNew(conConta, "tiposiva", "porceiva", "codigiva", CStr(RS!codigiva), "N")
        Select Case I
            Case 1
                TipoIVA1 = RS.Fields(0).Value
                PorceIVA1 = CCur(SQL)
                BaseIVA1 = Round2(RS.Fields(1).Value / (1 + (PorceIVA1 / 100)), 2)
'                SQL = CalcularDto(CStr(BaseIVA1), CStr(PorceIVA1))
'                ImpIVA1 = Round(CCur(SQL), 2)
                ImpIVA1 = RS.Fields(1).Value - BaseIVA1
            Case 2
                TipoIVA2 = RS.Fields(0).Value
                PorceIVA2 = CCur(SQL)
'                BaseIVA2 = RS.Fields(1).Value
'                SQL = CalcularDto(CStr(BaseIVA2), CStr(PorceIVA2))
'                ImpIVA2 = CCur(SQL)
                BaseIVA2 = Round2(RS.Fields(1).Value / (1 + (PorceIVA2 / 100)), 2)
                ImpIVA2 = RS.Fields(1).Value - BaseIVA2
            Case 3
                TipoIVA3 = RS.Fields(0).Value
                PorceIVA3 = CCur(SQL)
'                BaseIVA3 = RS.Fields(1).Value
'                SQL = CalcularDto(CStr(BaseIVA3), CStr(PorceIVA3))
'                ImpIVA3 = CCur(SQL)
                BaseIVA3 = Round2(RS.Fields(1).Value / (1 + (PorceIVA3 / 100)), 2)
                ImpIVA3 = RS.Fields(1).Value - BaseIVA3
        End Select
        RS.MoveNext
        I = I + 1
    Wend
    RS.Close
    Set RS = Nothing
    
    
    
    
    'comprobar que suma bases imp + imporiva = total factura
    ImpTotal = (BaseIVA1 + ImpIVA1) + (BaseIVA2 + ImpIVA2) + (BaseIVA3 + ImpIVA3)
    BrutoFac = BaseIVA1 + BaseIVA2 + BaseIVA3
    BaseImp = BrutoFac
    
    If Not (ImpTotal = TotalFac) Then
        ImpDif = TotalFac - ImpTotal
        ImpIVA1 = ImpIVA1 + ImpDif
'        If BaseIVA2 = 0 And BaseIVA3 = 0 Then ImpIVA1 = ImpIVA1 + ImpDif
'        ImpTotal = (BaseIVA1 + ImpIVA1) + (BaseIVA2 + ImpIVA2) + (BaseIVA3 + ImpIVA3)
    End If
    
    b = True
    
    
ETicket:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Calculando datos de ticket a factura.", Err.Description
        b = False
    End If
    CalcularDatosFacturaTicket = b
End Function


Private Function InsertarFormaPagoEnConta(nForPa As String, cadErr As String) As Boolean
Dim cadAux As String
Dim cadAux2 As String
Dim SQL As String

    On Error GoTo ErrInsForpa
    InsertarFormaPagoEnConta = False
    
    'antes de grabar en la scobro comprobar que existe en conta.sforpa la
    'forma de pago de la factura. Sino existe insertarla
    
    'vemos si existe en la conta
    'vemos si existe en la conta
    If vParamAplic.ContabilidadNueva Then
        cadAux = DevuelveDesdeBDNew(conConta, "formapago", "codforpa", "codforpa", nForPa, "N")
        If cadAux = "" Then Err.Raise 513, , "No existe forma de pago: " & nForPa   'para que no

    Else
        cadAux = DevuelveDesdeBDNew(conConta, "sforpa", "codforpa", "codforpa", nForPa, "N")
    End If
    'si no existe la forma de pago en conta, insertamos la de ariges
    If cadAux = "" Then
        cadAux2 = "tipforpa"
        cadAux = DevuelveDesdeBDNew(conAri, "sforpa", "nomforpa", "codforpa", nForPa, "N", cadAux2)
        If cadAux <> "" Then
            'insertamos e sforpa de la CONTA
            SQL = "INSERT INTO sforpa(codforpa,nomforpa,tipforpa)"
            SQL = SQL & " VALUES(" & nForPa & ", " & DBSet(cadAux, "T") & ", " & cadAux2 & ")"
            ConnConta.Execute SQL
            InsertarFormaPagoEnConta = True
        Else
            InsertarFormaPagoEnConta = False
        End If
    Else
        InsertarFormaPagoEnConta = True
    End If
    
    
    Exit Function
    
ErrInsForpa:
    InsertarFormaPagoEnConta = False
    cadErr = "Insertar forma de pago en Contablilidad: " & vbCrLf & Err.Description
End Function



'#### Laura 15/11/2006 Recuperar facturas ALZIRA
Public Function MarcarContabilizada(Optional cadErr As String) As Boolean
'Poner la factura como contabilizada
Dim SQL As String

    On Error GoTo EActualizar
    MarcarContabilizada = False
    
    SQL = "UPDATE scafac SET intconta=1 "
    SQL = SQL & " WHERE codtipom=" & DBSet(Me.Codtipom, "T") & " AND numfactu=" & Me.NumFactu & " AND fecfactu=" & DBSet(Me.Fecfactu, "F")
    conn.Execute SQL
    
    MarcarContabilizada = True
    Exit Function
    
EActualizar:
    MarcarContabilizada = False
    cadErr = Err.Description
End Function




