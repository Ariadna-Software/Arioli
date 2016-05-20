VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CParamAplic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Variables locales que contienen valores de propiedad
'Variables asociadas a cada campo de la tabla de parametros SPARA1
Private mvarTipoDtos As Byte      'Tipo de Descuentos: 0= Aditivo 1=Sobre Resto
Private mvarOrdenDtos As Byte   'Orden de Dtos para obtener los dtos a aplicar al precio segun la tabla: sdtofm
                                'Si F=0004 y M=0004 y existe dtos para (F,M)=(0000,0004) y para (F,M)=(0004,0000) siendo 0000=todas las marcas/todas las familias
                                'si ordendto=0(Fam/Marca) aplicaremos el dto (F,M)=(0004,0000)
                                'si ordendto=1 (Marca/Fam) aplicaremos el dto (F,M)=(0000,0004)
                                
Private mvarCodTarif As Integer   'Cod. Tarifa de PVP
Private mvarCajaComp As Boolean      'Cajas completas precios (S/N)
Private mvarPreuKmCli As Single 'Precio Km Desplazamiento Clientes
Private mvarPreuKmTec As Single 'Precio Km Desplazamiento Técnicos
Private mvarnomGerente As String  'Nombre del Director Gerente
Private mvarnomAdmin As String  'Nombre del Responsable Administración
Private mvarHayMante As Boolean   'Si hay Mantenimiento (S/N)
Private mvarHayRepar As Boolean   'Si hay Reparaciones (S/N)
Private mvarHayFrecu As Boolean   'Si hay Frecuencias (S/N)
Private mvarHayServicio As Boolean   'Si hay Servicios (S/N)
Private mvarHayDepar As Boolean   'Si hay Departamento o Direccion
Private mvarCtrStock As Boolean   'Si hay Control de Stock (S/N)
Private mvarInventar As Boolean  'Si se realiza Inventario ordenando listados por Proveedor o no
Private mvarHaynserie As Boolean 'Si hay Nº de Serie en Compras (S/N)
Private mvarDiasSiMan As Integer 'Nº de dias si hay mantenimiento
Private mvarDiasNoMan As Integer 'Nº de dias si NO hay mantenimiento

'Datos Facturacion
Private mvarCodArticDes As String 'Articulo para facturar desplazamientos
Private mvarDiaPago1 As Byte 'Dia de pago 1 para compras
Private mvarDiaPago2 As Byte 'Dia de pago 2 para compras
Private mvarDiaPago3 As Byte 'Dia de pago 3 para compras
Private mvarMesnogir As Byte 'Mes a no girar para compras
Private mvarCodForPa As String 'forma de pago para cheque regalo

'Datos de Internet
Private mvarDireMail As String 'Direccion email
Private mvarSMTPhost As String 'Host SMTP para envio de email
Private mvarSMTPuser As String 'Usuario SMTP para envio de email
Private mvarSMTPpass As String 'Password SMTP para envio de email
Private mvarWebSoporte As String 'Web de soporte
Private mvarMailSoporte As String 'Mail de soporte
Private mvarWebVersion As String 'Version de la web

'Datos de la Contabilidad
Private mvarSerConta As String 'Servidor BD de Contabilidad
Private mvarUsuConta As String 'Usuario BD de Contabilidad
Private mvarPasConta As String 'Password BD de Contabilidad
Private mvarNumConta As Byte 'Numero BD de Contabilidad


'Valores por defecto.   Toca meter un variable pq en el matenimiento (en el form) no deja modificar si no es con el objeto
'-----------------
'
'      Clientes
Private mvarDefActiv As Integer
Private mvarDefEnvio As Integer
Private mvarDefZona As Integer
Private mvarDefRuta As Integer
Private mvarDefSitu As Integer
Private mvarDefTarifa As Integer
Private mvarDefAgente As Integer

'Facturacion recargas telefonia
Private mvarCodarticTfnia As String



    'aVISOS  al entrar
    Private mvaravipedcli  As Integer
    Private mvaravipedpro As Integer
    Private mvaravialbcli As Integer
    Private mvaravialbpro As Integer
    Private mvaravimanteni As Integer
    Private mvaravirepara As Integer
    Private mvaraviavisos As Integer
    Private mvarTieneAvisos As Boolean 'De este no hay valor en BD. Simplememnte si alguno de ellos tiene valor, el tendra valor


'Si LLEVA REcargo de equivalencia
Private mvarRecargoEquivalencia As Boolean
Private mvarTipoIVAre1 As Integer
Private mvarTipoIVAre2 As Integer
Private mvarTipoIVAre3 As Integer
Private mvarIVA1 As Integer
Private mvarIVA2 As Integer
Private mvarIVA3 As Integer


'Retencion y REA
Private mvarIVAREA As Integer
Private mvarCtaReten As String
Private mvarPorReten As Currency

Private mvarIVA_Exento As Integer
Private mvarIVA_Intracomunitario As Integer

'Tickets TPV agrupados a contabilidad
Private mvarContabilizarTicketAgrupados As Boolean

'Albaranes y facturas en B
Private mvarContabilidadB As Byte

'Cueenta para la aportacion al terminal en RADIVAL
Private mvarctaaportacion As String

'Produccion
Private mvarProduccion As Boolean

'Descriptores
Private mvarDescriptores As Boolean

'Que llevara al campo observaciones de la factura cli/pro
Private mvarObsFactura As Byte

'Reciclado / PUNTO VERDE
Private mvarArtReciclado As String

'Almacen del B      FALTA###   De momento esta a piñon el 2
Private mvarAlmacenB As Integer


Private mvarOLDIVA1 As Integer
Private mvarOLDIVA2 As Integer
Private mvarOLDIVAre1 As Integer
Private mvarOLDIVAre2 As Integer

'Abril 2011
Private mvarProduccionNueva As Boolean
'Mayo 2011
Private mvarEsAVAB As Boolean 'empresa exportadora

'Junio 2011
Private mvarCodEmpresaMorales As Byte 'NO lleva NI GET ni LET. No se utilza de momento.Se cambia por el YOG.  En todas llevara el codempresa que es MOrales (el UNO vamos)
Private mvarNoUtilizaTOs As Boolean   'Solo lleva GET
Private mvarContabilizacionMoixent As Boolean
        'Si contabiliza como MOrales(ariges) o contabiliza como moixent, es decir
        'para cada fam,marc,unid,tipar unas cuentas

'Noviembre 2011
' Identificador de empresa.  Podriamos ver de quitar es empresa AVAB, morales... Con este lo tendriamos todo
Private mvarqueempresaes As Byte   '0-1 Morales AVAB   2 Bodega Moixent  3 Almazara Moixent



Private SQL As String  'Para la sentencia
Private Mens As String    'Para los mensajes






'FALTA###
Public Property Get AlmacenB() As Integer
     'AlmacenB = mvarAlmacenB
     AlmacenB = 2
End Property





'------------------------------------------------
'Propiedades del modulo CParamAplic
'------------------------------------------------

Public Property Let TipoDtos(ByVal vData As Byte)
     mvarTipoDtos = vData
End Property

Public Property Get TipoDtos() As Byte
     TipoDtos = mvarTipoDtos
End Property


Public Property Let OrdenDtos(ByVal vData As Byte)
     mvarOrdenDtos = vData
End Property

Public Property Get OrdenDtos() As Byte
     OrdenDtos = mvarOrdenDtos
End Property

Public Property Let CodTarifa(ByVal vData As Integer)
     mvarCodTarif = vData
End Property

Public Property Get CodTarifa() As Integer
     CodTarifa = mvarCodTarif
End Property


Public Property Let CajasCompletas(ByVal vData As Boolean)
     mvarCajaComp = vData
End Property

Public Property Get CajasCompletas() As Boolean
     CajasCompletas = mvarCajaComp
End Property


Public Property Let PrecioKmClientes(ByVal vData As Single)
     mvarPreuKmCli = vData
End Property

Public Property Get PrecioKmClientes() As Single
     PrecioKmClientes = mvarPreuKmCli
End Property


Public Property Let PrecioKmTecnicos(ByVal vData As Single)
     mvarPreuKmTec = vData
End Property

Public Property Get PrecioKmTecnicos() As Single
     PrecioKmClientes = mvarPreuKmTec
End Property


Public Property Let NomGerente(ByVal vData As String)
     mvarnomGerente = vData
End Property


Public Property Get NomGerente() As String
     NomGerente = mvarnomGerente
End Property


Public Property Let NomRespAdmin(ByVal vData As String)
     mvarnomAdmin = vData
End Property

Public Property Get NomRespAdmin() As String
     NomRespAdmin = mvarnomAdmin
End Property


Public Property Let Mantenimientos(ByVal vData As Boolean)
     mvarHayMante = vData
End Property

Public Property Get Mantenimientos() As Boolean
     Mantenimientos = mvarHayMante
End Property


Public Property Let Reparaciones(ByVal vData As Boolean)
     mvarHayRepar = vData
End Property

Public Property Get Reparaciones() As Boolean
     Reparaciones = mvarHayRepar
End Property

Public Property Let Frecuencias(ByVal vData As Boolean)
'    If vData Then
'        MsgBox "Lo sentimos todavía no puede activar las frecuencias. En próximas versiones de la aplicación será posible", vbExclamation
'        vData = False
'    End If
     mvarHayFrecu = vData
End Property

Public Property Get Servicios() As Boolean
     Servicios = mvarHayServicio
End Property

Public Property Let Servicios(ByVal vData As Boolean)
    Dim Mens As String
    If vData Then
        If MovimientosParaServicios() Then
            vData = True
        Else
            Mens = "Para poder activar los servicios es necesario que los movimientos ALS y FAS estén dados de alta." & vbCrLf & _
                    "Délos de alta ahora y luego vuelva a intentar activar los servicios. Gracias."
            MsgBox Mens, vbExclamation
            vData = False
        End If
    End If
     mvarHayServicio = vData
End Property

Public Property Get Frecuencias() As Boolean
     Frecuencias = mvarHayFrecu
End Property

Public Property Let Departamento(ByVal vData As Boolean)
     mvarHayDepar = vData
End Property

Public Property Get Departamento() As Boolean
     Departamento = mvarHayDepar
End Property


Public Property Let ControlStock(ByVal vData As Boolean)
     mvarCtrStock = vData
End Property

Public Property Get ControlStock() As Boolean
     ControlStock = mvarCtrStock
End Property

Public Property Let InventarioxProv(ByVal vData As Boolean)
     mvarInventar = vData
End Property

Public Property Get InventarioxProv() As Boolean
     InventarioxProv = mvarInventar
End Property

Public Property Let NumSeries(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.contabil = 5
     mvarHaynserie = vData
End Property

Public Property Get NumSeries() As Boolean
     NumSeries = mvarHaynserie
End Property


Public Property Let DiasSiMante(ByVal vData As Integer)
     mvarDiasSiMan = vData
End Property

Public Property Get DiasSiMante() As Integer
     DiasSiMante = mvarDiasSiMan
End Property


Public Property Let DiasNoMante(ByVal vData As Integer)
     mvarDiasNoMan = vData
End Property

Public Property Get DiasNoMante() As Integer
     DiasNoMante = mvarDiasNoMan
End Property


Public Property Let ArticDesplaz(ByVal vData As String)
     mvarCodArticDes = vData
End Property

Public Property Get ArticDesplaz() As String
     ArticDesplaz = mvarCodArticDes
End Property


Public Property Let DiaPago1(ByVal vData As Byte)
     mvarDiaPago1 = vData
End Property

Public Property Get DiaPago1() As Byte
     DiaPago1 = mvarDiaPago1
End Property

Public Property Let DiaPago2(ByVal vData As Byte)
     mvarDiaPago2 = vData
End Property

Public Property Get DiaPago2() As Byte
     DiaPago2 = mvarDiaPago2
End Property

Public Property Let DiaPago3(ByVal vData As Byte)
     mvarDiaPago3 = vData
End Property

Public Property Get DiaPago3() As Byte
     DiaPago3 = mvarDiaPago3
End Property


Public Property Let MesNoGirar(ByVal vData As Byte)
     mvarMesnogir = vData
End Property

Public Property Get MesNoGirar() As Byte
     MesNoGirar = mvarMesnogir
End Property

'forma de pago de los cheques regalo
Public Property Let ForPagoChequeRegalo(ByVal vData As String)
     mvarCodForPa = vData
End Property

Public Property Get ForPagoChequeRegalo() As String
     ForPagoChequeRegalo = mvarCodForPa
End Property




Public Property Let DireMail(ByVal vData As String)
     mvarDireMail = vData
End Property

Public Property Get DireMail() As String
     DireMail = mvarDireMail
End Property

Public Property Let SMTPhost(ByVal vData As String)
     mvarSMTPhost = vData
End Property

Public Property Get SMTPhost() As String
     SMTPhost = mvarSMTPhost
End Property

Public Property Let SMTPuser(ByVal vData As String)
     mvarSMTPuser = vData
End Property

Public Property Get SMTPuser() As String
     SMTPuser = mvarSMTPuser
End Property

Public Property Let SMTPpass(ByVal vData As String)
     mvarSMTPpass = vData
End Property

Public Property Get SMTPpass() As String
     SMTPpass = mvarSMTPpass
End Property

Public Property Let WebSoporte(ByVal vData As String)
     mvarWebSoporte = vData
End Property

Public Property Get WebSoporte() As String
     WebSoporte = mvarWebSoporte
End Property

Public Property Let MailSoporte(ByVal vData As String)
     mvarMailSoporte = vData
End Property

Public Property Get MailSoporte() As String
     MailSoporte = mvarMailSoporte
End Property

Public Property Let WebVersion(ByVal vData As String)
     mvarWebVersion = vData
End Property

Public Property Get WebVersion() As String
     WebVersion = mvarWebVersion
End Property



'servidor conta
Public Property Let ServidorConta(ByVal vData As String)
     mvarSerConta = vData
End Property

Public Property Get ServidorConta() As String
     ServidorConta = mvarSerConta
End Property



Public Property Let UsuarioConta(ByVal vData As String)
     mvarUsuConta = vData
End Property

Public Property Get UsuarioConta() As String
     UsuarioConta = mvarUsuConta
End Property


Public Property Let PasswordConta(ByVal vData As String)
     mvarPasConta = vData
End Property

Public Property Get PasswordConta() As String
     PasswordConta = mvarPasConta
End Property


Public Property Let NumeroConta(ByVal vData As Byte)
     mvarNumConta = vData
End Property

Public Property Get NumeroConta() As Byte
     NumeroConta = mvarNumConta
End Property




Public Property Let PorDefecto_Activ(ByVal vData As Integer)
     mvarDefActiv = vData
End Property

Public Property Get PorDefecto_Activ() As Integer
     PorDefecto_Activ = mvarDefActiv
End Property
Public Property Let PorDefecto_Envio(ByVal vData As Integer)
     mvarDefEnvio = vData
End Property

Public Property Get PorDefecto_Envio() As Integer
     PorDefecto_Envio = mvarDefEnvio
End Property
Public Property Let PorDefecto_Zona(ByVal vData As Integer)
     mvarDefZona = vData
End Property

Public Property Get PorDefecto_Zona() As Integer
     PorDefecto_Zona = mvarDefZona
End Property
Public Property Let PorDefecto_Ruta(ByVal vData As Integer)
     mvarDefRuta = vData
End Property

Public Property Get PorDefecto_Ruta() As Integer
     PorDefecto_Ruta = mvarDefRuta
End Property
Public Property Let PorDefecto_Situ(ByVal vData As Integer)
     mvarDefSitu = vData
End Property

Public Property Get PorDefecto_Situ() As Integer
     PorDefecto_Situ = mvarDefSitu
End Property

Public Property Let PorDefecto_Tarifa(ByVal vData As Integer)
     mvarDefTarifa = vData
End Property

Public Property Get PorDefecto_Tarifa() As Integer
     PorDefecto_Tarifa = mvarDefTarifa
End Property
Public Property Let PorDefecto_Agente(ByVal vData As Integer)
     mvarDefAgente = vData
End Property

Public Property Get PorDefecto_Agente() As Integer
     PorDefecto_Agente = mvarDefAgente
End Property



Public Property Let CodarticTfnia(ByVal vData As String)
     mvarCodarticTfnia = vData
End Property

Public Property Get CodarticTfnia() As String
     CodarticTfnia = mvarCodarticTfnia
End Property




Public Property Let avipedcli(ByVal vData As Integer)
     mvaravipedcli = vData
End Property

Public Property Get avipedcli() As Integer
     avipedcli = mvaravipedcli
End Property
    
Public Property Let avipedpro(ByVal vData As Integer)
     mvaravipedpro = vData
End Property

Public Property Get avipedpro() As Integer
     avipedpro = mvaravipedpro
End Property

Public Property Let avialbcli(ByVal vData As Integer)
     mvaravialbcli = vData
End Property

Public Property Get avialbcli() As Integer
     avialbcli = mvaravialbcli
End Property
Public Property Let avialbpro(ByVal vData As Integer)
     mvaravialbpro = vData
End Property

Public Property Get avialbpro() As Integer
     avialbpro = mvaravialbpro
End Property
    

Public Property Let avimanteni(ByVal vData As Integer)
     mvaravimanteni = vData
End Property

Public Property Get avimanteni() As Integer
     avimanteni = mvaravimanteni
End Property
    
Public Property Let avirepara(ByVal vData As Integer)
     mvaravirepara = vData
End Property

Public Property Get avirepara() As Integer
     avirepara = mvaravirepara
End Property

Public Property Let aviavisos(ByVal vData As Integer)
     mvaraviavisos = vData
End Property

Public Property Get aviavisos() As Integer
     aviavisos = mvaraviavisos
End Property

'Solo tiene GET
Public Property Get TieneAvisos() As Integer
     TieneAvisos = mvarTieneAvisos
End Property



'---------------------------------------
'RECARGO EQUIVALENCIA
Public Property Get TipoIVAre1() As Integer
     TipoIVAre1 = mvarTipoIVAre1
End Property

Public Property Let TipoIVAre1(ByVal vData As Integer)
     mvarTipoIVAre1 = vData
End Property
Public Property Get TipoIVAre2() As Integer
     TipoIVAre2 = mvarTipoIVAre2
End Property

Public Property Let TipoIVAre2(ByVal vData As Integer)
     mvarTipoIVAre2 = vData
End Property

Public Property Get TipoIVAre3() As Integer
     TipoIVAre3 = mvarTipoIVAre3
End Property

Public Property Let TipoIVAre3(ByVal vData As Integer)
     mvarTipoIVAre3 = vData
End Property

Public Property Get TipoIVA1() As Integer
     TipoIVA1 = mvarIVA1
End Property

Public Property Let TipoIVA1(ByVal vData As Integer)
     mvarIVA1 = vData
End Property
Public Property Get TipoIVA2() As Integer
     TipoIVA2 = mvarIVA2
End Property

Public Property Let TipoIVA2(ByVal vData As Integer)
     mvarIVA2 = vData
End Property

Public Property Get TipoIVA3() As Integer
     TipoIVA3 = mvarIVA3
End Property

Public Property Let TipoIVA3(ByVal vData As Integer)
     mvarIVA3 = vData
End Property

Public Property Get RecargoEquivalencia() As Integer
     RecargoEquivalencia = mvarRecargoEquivalencia
End Property



'RETENCION Y REGIMEN ESPECIAL AGRARIO
'---------------------------------------
Public Property Get IVA_REA() As Integer
     IVA_REA = mvarIVAREA
End Property

Public Property Let IVA_REA(ByVal vData As Integer)
     mvarIVAREA = vData
End Property


Public Property Get CtaReten() As String
     CtaReten = mvarCtaReten
End Property

Public Property Let CtaReten(ByVal vData As String)
     mvarCtaReten = vData
End Property

Public Property Get PorReten() As Currency
     PorReten = mvarPorReten
End Property

Public Property Let PorReten(ByVal vData As Currency)
     mvarPorReten = vData
End Property

'IVA exento
Public Property Get IVA_Exento2() As Integer
     IVA_Exento2 = mvarIVA_Exento
End Property

Public Property Let IVA_Exento2(ByVal vData As Integer)
     mvarIVA_Exento = vData
End Property



'IVA intracomunitario
Public Property Get IVA_Intracomunitario() As Integer
     IVA_Intracomunitario = mvarIVA_Intracomunitario
End Property

Public Property Let IVA_Intracomunitario(ByVal vData As Integer)
     mvarIVA_Intracomunitario = vData
End Property



'Contabilizacioon del Ticket del TPV a grupado
Public Property Get ContabilizarTicketAgrupados() As Boolean
     ContabilizarTicketAgrupados = mvarContabilizarTicketAgrupados
End Property

Public Property Let ContabilizarTicketAgrupados(ByVal vData As Boolean)
     mvarContabilizarTicketAgrupados = vData
End Property


'Numero contabilidad para las albaranes /facturas en B
Public Property Get ContabilidadB() As Byte
     ContabilidadB = mvarContabilidadB
End Property

Public Property Let ContabilidadB(ByVal vData As Byte)
     mvarContabilidadB = vData
End Property


'APortacion al terminal. Ej. Facturas de mostrador en RADIVAL
Public Property Get ctaAportacion() As String
     ctaAportacion = mvarctaaportacion
End Property

Public Property Let ctaAportacion(ByVal vData As String)
     mvarctaaportacion = vData
End Property


'Tiene produccion
Public Property Get Produccion() As Boolean
     Produccion = mvarProduccion
End Property

Public Property Let Produccion(ByVal vData As Boolean)
     mvarProduccion = vData
End Property

'Tiene Descriptores
Public Property Get Descriptores() As Boolean
     Descriptores = mvarDescriptores
End Property

Public Property Let Descriptores(ByVal vData As Boolean)
     mvarDescriptores = vData
End Property


'Campo observaciones en factura
Public Property Get ObsFactura() As Byte
     ObsFactura = mvarObsFactura
End Property

Public Property Let ObsFactura(ByVal vData As Byte)
     mvarObsFactura = vData
End Property


'Articulo de Recilado o punto Verde
Public Property Get ArtReciclado() As String
     ArtReciclado = mvarArtReciclado
End Property

Public Property Let ArtReciclado(ByVal vData As String)
     mvarArtReciclado = vData
End Property

'PARA el cambio de IVA
Public Property Get OLDIVA1() As Integer
     OLDIVA1 = mvarOLDIVA1
End Property

Public Property Let OLDIVA1(ByVal vData As Integer)
     mvarOLDIVA1 = vData
End Property

Public Property Get OLDIVA2() As Integer
     OLDIVA2 = mvarOLDIVA2
End Property

Public Property Let OLDIVA2(ByVal vData As Integer)
     mvarOLDIVA2 = vData
End Property
Public Property Get OLDIVAre1() As Integer
     OLDIVAre1 = mvarOLDIVAre1
End Property

Public Property Let OLDIVAre1(ByVal vData As Integer)
     mvarOLDIVAre1 = vData
End Property

Public Property Get OLDIVAre2() As Integer
     OLDIVAre2 = mvarOLDIVAre2
End Property

Public Property Let OLDIVAre2(ByVal vData As Integer)
     mvarOLDIVAre2 = vData
End Property


'Para caudno empecemos con las pistolas
'pondremos esta marca a true
Public Property Get ProduccionNueva() As Boolean
     ProduccionNueva = mvarProduccionNueva
End Property

Public Property Let ProduccionNueva(ByVal vData As Boolean)
     mvarProduccionNueva = vData
End Property

'Esta sera la empresa AVAB
Public Property Get EsAVAB() As Boolean
     EsAVAB = mvarEsAVAB
End Property

Public Property Let EsAVAB(ByVal vData As Boolean)
     mvarEsAVAB = vData
End Property


' mvarCodEmpresaMorales (Byte) 'NO lleva NI GET ni LET. No se utilza de momento.Se cambia por el YOG.  En todas llevara el codempresa que es MOrales (el UNO vamos)
' mvarNoUtilizaTOs As Boolean   'Solo lleva GET  --> Moixent NO tuliza TOS
Public Property Get NoUtilizaTOs() As Boolean
     NoUtilizaTOs = mvarNoUtilizaTOs
End Property




'Si contabiliza como MOrales(ariges) o contabiliza como moixent, es decir
'para cada fam,marc,unid,tipar unas cuentas
Public Property Get ContabilizacionMoixent() As Boolean
     ContabilizacionMoixent = mvarContabilizacionMoixent
End Property


'Con este podriamos quitar alguno del los de arriba (empresaAvab, emprMorales..
Public Property Get QUE_EMPRESA() As Byte
     QUE_EMPRESA = mvarqueempresaes   '0-1 Morales AVAB   2 Bodega Moxient  3 Almazara Moixent     4.- QuatreTonda
End Property







'------------------------------------------------
'FUNCIONES del modulo Cparametros
'------------------------------------------------

'****   LEER  ***
Public Function Leer() As Byte
Dim RS As ADODB.Recordset
On Error GoTo Err3
    
    SQL = "SELECT * FROM spara1 "
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic

    If RS.EOF Then
        Leer = 1
    Else
        mvarTipoDtos = RS!TipoDtos
        mvarOrdenDtos = RS!ordendto
        mvarCodTarif = RS!codTarif
        mvarCajaComp = RS!cajacomp  'Cajas completas precios (S/N)
        mvarPreuKmCli = DBLet(RS!preukmcl, "N") 'Precio Km Desplazamiento Clientes
        mvarPreuKmTec = DBLet(RS!preukmtc, "N") 'Precio Km Desplazamiento Técnicos
        mvarnomGerente = DBLet(RS!nomgeren, "T") 'Nombre del Director Gerente
        mvarnomAdmin = DBLet(RS!nomadmin, "T") 'Nombre del Responsable Administración
        mvarHayMante = RS!haymante  'Si hay Mantenimiento (S/N)
        mvarHayRepar = RS!hayrepar  'Si hay Reparaciones (S/N)
        mvarHayFrecu = RS!hayfrecu  'Si hay Frecuencias (S/N)
        mvarHayServicio = RS!hayservicio  'Si hay Servicios (S/N)
        mvarHayDepar = RS!haydepar  'Si hay Departamento o Direccion
        mvarCtrStock = RS!CtrStock  'Si hay Control de Stock (S/N)
        mvarInventar = RS!inventar
        mvarHaynserie = RS!haynserie 'Si hay Nº Serie en Compras (S/N)
        mvarDiasSiMan = RS!diassiman 'Dias reparacion si hay Mantenimiento
        mvarDiasNoMan = RS!diasnoman 'Dias reparacion si NO hay Mantenimiento
        Me.ForPagoChequeRegalo = DBLet(RS!codforpa, "T")
        
        'Datos Facturacion
        mvarCodArticDes = DBLet(RS!codartid, "T") 'articulo para facturar mantenimientos
        mvarDiaPago1 = DBLet(RS!DiaPago1, "N") 'dia pago 1 para compras
        mvarDiaPago2 = DBLet(RS!DiaPago2, "N") 'dia pago 2 para compras
        mvarDiaPago3 = DBLet(RS!DiaPago3, "N") 'dia pago 3 para compras
        mvarMesnogir = DBLet(RS!mesnogir, "N") 'mes a no girar para compras
        
        'Datos de Internet
        mvarDireMail = DBLet(RS!DireMail, "T") 'Direccion de email
        mvarSMTPhost = DBLet(RS!SMTPhost, "T") 'SMTP host para el envio de email
        mvarSMTPuser = DBLet(RS!SMTPuser, "T") 'SMTP usuario para el envio de email
        mvarSMTPpass = DBLet(RS!SMTPpass, "T") 'SMTP password para el envio de email
        mvarWebSoporte = DBLet(RS!WebSoporte, "T") 'Web de soporte
        mvarMailSoporte = DBLet(RS!MailSoporte, "T") 'email de soporte
        mvarWebVersion = DBLet(RS!WebVersion, "T") 'Version de la Web
        
        
        'Datos Contabilidad
        mvarSerConta = DBLet(RS!serconta, "T") 'servidor conta
        mvarUsuConta = DBLet(RS!UsuConta, "T") 'usuario conta
        mvarPasConta = DBLet(RS!pasconta, "T") 'password
        mvarNumConta = DBLet(RS!NumConta, "N") 'numero conta
        
        
        'Valores por defecto
        mvarDefActiv = DBLet(RS!defactividad, "N") 'numero conta
        mvarDefEnvio = DBLet(RS!defenvio, "N") 'numero conta
        mvarDefZona = DBLet(RS!defzona, "N") 'numero conta
        mvarDefRuta = DBLet(RS!defruta, "N") 'numero conta
        mvarDefSitu = DBLet(RS!defstituacion, "N") 'numero conta
        mvarDefTarifa = DBLet(RS!deftarifa, "N") 'numero conta
        mvarDefAgente = DBLet(RS!defagente, "N") 'numero conta


        mvaravipedcli = DBLet(RS!avipedcli, "N")
        mvaravipedpro = DBLet(RS!avipedpro, "N")
        mvaravialbcli = DBLet(RS!avialbcli, "N")
        mvaravialbpro = DBLet(RS!avialbpro, "N")
        mvaravimanteni = DBLet(RS!avimanteni, "N")
        mvaravirepara = DBLet(RS!avirepara, "N")
        mvaraviavisos = DBLet(RS!aviavios, "N")
        PonVariableTienAvisos

        
        

        mvarTipoIVAre1 = DBLet(RS!IVAre1eq, "N")
        mvarTipoIVAre2 = DBLet(RS!IVAre2eq, "N")
        mvarTipoIVAre3 = DBLet(RS!IVAre3eq, "N")
        mvarRecargoEquivalencia = (mvarTipoIVAre1 > 0 Or mvarTipoIVAre2 > 0 Or mvarTipoIVAre3 > 0)
        
        mvarIVA1 = DBLet(RS!IVAre1, "N")
        mvarIVA2 = DBLet(RS!IVAre2, "N")
        mvarIVA3 = DBLet(RS!IVAre3, "N")

        'El articulo de telefonia no lo pongo como parametros
        'codartictel
        mvarCodarticTfnia = DBLet(RS!codartictel, "T") 'numero conta
                
                
        'Retencion
        mvarIVAREA = DBLet(RS!IVA_REA, "N")
        mvarCtaReten = DBLet(RS!CtaReten, "T")
        mvarPorReten = DBLet(RS!PorReten, "N")
                
                
        'IVA exento
        mvarIVA_Exento = DBLet(RS!ivaexento, "N")
        
        
        
        'Contabilidad
        mvarContabilidadB = DBLet(RS!conta_B, "N")
        
        'Contabiliza los tickets resumidos
        mvarContabilizarTicketAgrupados = (DBLet(RS!conttickagrupado, "N") = 1)
        
        'Cta aportacion
        mvarctaaportacion = DBLet(RS!ctaAportacion, "T")
        
        'produccion
        mvarProduccion = DBLet(RS!Produccion, "N") = 1
        'descriptores
        mvarDescriptores = DBLet(RS!Descriptores, "N") = 1
        'obs factura
        mvarObsFactura = DBLet(RS!ObsFactura, "N")
        '
        mvarIVA_Intracomunitario = DBLet(RS!IvaIntracom, "N")
        'Recliclado Punto verde
        mvarArtReciclado = DBLet(RS!ArtReciclado, "T")


        mvarOLDIVA1 = DBLet(RS!iva_old1, "N")
        mvarOLDIVA2 = DBLet(RS!iva_old2, "N")
        mvarOLDIVAre1 = DBLet(RS!iva_oldre1, "N")
        mvarOLDIVAre2 = DBLet(RS!iva_oldre2, "N")
        
        mvarProduccionNueva = DBLet(RS!ProduccionNueva, "N") = 1
        mvarEsAVAB = DBLet(RS!EsEmpresaExportadora, "N") = 1
        
        'CodEmpresaMorales. Este campo en la BD no tiene referencia en los parametros.
        '                   Solo sirve para marcar cual es la empresa de MORALES. Por YOG
        
        mvarNoUtilizaTOs = DBLet(RS!NoUtilizaTOs, "N") = 1
        mvarContabilizacionMoixent = DBLet(RS!ContabilizacionMoixe, "N") = 1
        
        '0-1 Morales AVAB   2 Bodega Moxient  3 Almazara Moixent
        mvarqueempresaes = DBLet(RS!QueEmpresaEs, "N")

        Leer = 0
    End If

    RS.Close
    Set RS = Nothing
    Exit Function

Err3:
    Mens = "Se ha producido un error(Leer Parametros)" & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    RS.Close
    Set RS = Nothing
    Leer = 1
End Function


Public Function Modificar(Codigo As Byte) As Byte
'Modifica los parametros de la Aplicación
'Modifica la Tabla: spara1, BD: Ariges
On Error GoTo EModificar

    SQL = "UPDATE spara1 SET "
    SQL = SQL & " tipodtos= " & mvarTipoDtos
    SQL = SQL & ", ordendto= " & mvarOrdenDtos
    SQL = SQL & ", codtarif= " & mvarCodTarif
    SQL = SQL & ", cajacomp= " & DBSet(mvarCajaComp, "B")
    SQL = SQL & ", preukmcl= " & DBSet(mvarPreuKmCli, "N")
    SQL = SQL & ", preukmtc= " & DBSet(mvarPreuKmTec, "N")
    SQL = SQL & ", nomgeren= " & DBSet(mvarnomGerente, "T")
    SQL = SQL & ", nomadmin= " & DBSet(mvarnomAdmin, "T")
    SQL = SQL & ", haymante = " & DBSet(mvarHayMante, "B")
    SQL = SQL & ", hayrepar = " & DBSet(mvarHayRepar, "B")
    SQL = SQL & ", hayfrecu = " & DBSet(mvarHayFrecu, "B")
    SQL = SQL & ", hayservicio = " & DBSet(mvarHayServicio, "B")
    SQL = SQL & ", haydepar = " & DBSet(mvarHayDepar, "B")
    SQL = SQL & ", ctrstock = " & DBSet(mvarCtrStock, "B")
    SQL = SQL & ", inventar= " & DBSet(mvarInventar, "B")
    SQL = SQL & ", haynserie= " & DBSet(mvarHaynserie, "B")
    SQL = SQL & ", diassiman= " & DBSet(mvarDiasSiMan, "N")
    SQL = SQL & ", diasnoman= " & DBSet(mvarDiasNoMan, "N")
    SQL = SQL & ", codartid= " & DBSet(mvarCodArticDes, "T")
    SQL = SQL & ", diapago1= " & DBSet(mvarDiaPago1, "N")
    SQL = SQL & ", diapago2= " & DBSet(mvarDiaPago2, "N")
    SQL = SQL & ", diapago3= " & DBSet(mvarDiaPago3, "N")
    SQL = SQL & ", mesnogir= " & DBSet(mvarMesnogir, "N")
    SQL = SQL & ", codforpa= " & DBSet(mvarCodForPa, "N", "S")
    SQL = SQL & ", diremail= " & DBSet(mvarDireMail, "T")
    SQL = SQL & ", smtphost= " & DBSet(mvarSMTPhost, "T")
    SQL = SQL & ", smtpuser= " & DBSet(mvarSMTPuser, "T")
    SQL = SQL & ", smtppass= " & DBSet(mvarSMTPpass, "T")
    SQL = SQL & ", websoporte= " & DBSet(mvarWebSoporte, "T")
    SQL = SQL & ", mailsoporte= " & DBSet(mvarMailSoporte, "T")
    SQL = SQL & ", webversion= " & DBSet(mvarWebVersion, "T")
    SQL = SQL & ", serconta= " & DBSet(mvarSerConta, "T")
    SQL = SQL & ", usuconta= " & DBSet(mvarUsuConta, "T")
    SQL = SQL & ", pasconta= " & DBSet(mvarPasConta, "T")
    SQL = SQL & ", numconta= " & DBSet(mvarNumConta, "N", "S")
    
    'Valore por defecto
    SQL = SQL & ", defactividad= " & DBSet(mvarDefActiv, "N", "S")
    SQL = SQL & ", defenvio= " & DBSet(mvarDefEnvio, "N", "S")
    SQL = SQL & ", defzona= " & DBSet(mvarDefZona, "N", "S")
    SQL = SQL & ", defruta= " & DBSet(mvarDefRuta, "N", "S")
    'SI K PUEDE SER 0
    SQL = SQL & ", defstituacion= " & DBSet(mvarDefSitu, "N", "N")
    SQL = SQL & ", deftarifa= " & DBSet(mvarDefTarifa, "N", "S")
    SQL = SQL & ", defagente= " & DBSet(mvarDefAgente, "N", "N")
      
    'Telefonia
    SQL = SQL & ", codartictel= " & DBSet(mvarCodarticTfnia, "T", "S")
      
      
      
      
    'Avisos
    SQL = SQL & ", avipedcli= " & DBSet(mvaravipedcli, "N", "S")
    SQL = SQL & ", avipedpro= " & DBSet(mvaravipedpro, "N", "S")
    SQL = SQL & ", avialbcli= " & DBSet(mvaravialbcli, "N", "S")
    SQL = SQL & ", avialbpro= " & DBSet(mvaravialbpro, "N", "S")
    SQL = SQL & ", avimanteni= " & DBSet(mvaravimanteni, "N", "S")
    SQL = SQL & ", avirepara= " & DBSet(mvaravirepara, "N", "S")
    SQL = SQL & ", aviavios= " & DBSet(mvaraviavisos, "N", "S")
      
    PonVariableTienAvisos

    SQL = SQL & ", IVAre1eq= " & DBSet(mvarTipoIVAre1, "N", "S")
    SQL = SQL & ", IVAre2eq= " & DBSet(mvarTipoIVAre2, "N", "S")
    SQL = SQL & ", IVAre3eq= " & DBSet(mvarTipoIVAre3, "N", "S")
    mvarRecargoEquivalencia = (mvarTipoIVAre1 > 0 Or mvarTipoIVAre2 > 0 Or mvarTipoIVAre3 > 0)
    
    SQL = SQL & ", IVAre1= " & DBSet(mvarIVA1, "N", "S")
    SQL = SQL & ", IVAre2= " & DBSet(mvarIVA2, "N", "S")
    SQL = SQL & ", IVAre3= " & DBSet(mvarIVA3, "N", "S")

    'Retencion y REA
    SQL = SQL & ", IVA_REA = " & DBSet(mvarIVAREA, "N", "S")
    SQL = SQL & ", CtaReten = " & DBSet(mvarCtaReten, "T", "S")
    SQL = SQL & ", PorReten = " & DBSet(mvarPorReten, "N", "S")

    'IVA exento
    SQL = SQL & ", ivaexento = " & DBSet(mvarIVA_Exento, "N", "S")


    'Tickets agrupados
    SQL = SQL & ", conttickagrupado = " & DBSet(mvarContabilizarTicketAgrupados, "B")

    'Contabilidad "B". Radival
    SQL = SQL & ", conta_B  = " & DBSet(mvarContabilidadB, "N")

    'Ctaa aportacion termial
    SQL = SQL & ", ctaaportacion  = " & DBSet(mvarctaaportacion, "T")
    'Produccion
    SQL = SQL & ", Produccion  = " & DBSet(mvarProduccion, "B")
    'Descriptores
    SQL = SQL & ", descriptores = " & DBSet(mvarDescriptores, "B")
    'Observaciones factura proveedor/cliente
    SQL = SQL & ", ObsFactura  = " & DBSet(mvarObsFactura, "N")
    SQL = SQL & ", IvaIntracom  = " & DBSet(mvarIVA_Intracomunitario, "N")

    'Tasa recliclado PUNTO VERDE
    SQL = SQL & ", ArtReciclado  = " & DBSet(mvarArtReciclado, "T", "S")
    

    SQL = SQL & ", iva_old1  = " & DBSet(mvarOLDIVA1, "T", "S")
    SQL = SQL & ", iva_old2  = " & DBSet(mvarOLDIVA2, "T", "S")
    SQL = SQL & ", iva_oldre1  = " & DBSet(mvarOLDIVAre1, "T", "S")
    SQL = SQL & ", iva_oldre2  = " & DBSet(mvarOLDIVAre2, "T", "S")

    SQL = SQL & ", ProduccionNueva = " & DBSet(mvarProduccionNueva, "B")
    SQL = SQL & ", EsEmpresaExportadora = " & DBSet(mvarEsAVAB, "B")
    
    'Las TOS y ContabilizacionMoixent no van en el update
    
    
    '-----------------
    'where
    SQL = SQL & " WHERE codigo =" & Codigo & ";"

    conn.Execute SQL



    'Desde los ivas de retencion, hay una tabla(sivatra) que tiene la relacion entre
    ' iva sin retencion, iva con retencion
    'Entocnes cuando modifico, borro las entradas y las vuevlo a generer
    CargarTablaTraduccionIVAS
    
    
    'Todo bien
    Modificar = 1

    Exit Function

EModificar:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Modificar = 0
End Function

Private Sub PonVariableTienAvisos()
    mvarTieneAvisos = (mvaravipedcli + mvaravipedpro + mvaravialbcli + mvaravialbpro + mvaravimanteni _
        + mvaravirepara + mvaraviavisos > 0)
End Sub

Private Function MovimientosParaServicios() As Boolean
    '-- Esta función comprueba si los movimientos ALS y FAS están creados
    '   este es un requisito previo si se desa activar la opción de servicios
    Dim SQL As String
    Dim RS As ADODB.Recordset
    SQL = "select count(*) from stipom where codtipom = 'ALS' or codtipom = 'FAS'"
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly
    If RS.Fields(0) = 2 Then
        MovimientosParaServicios = True
    Else
        MovimientosParaServicios = False
    End If
End Function



Public Function DevuleveTipoIVA_RE(TipoIVAOrigen As Integer) As Integer
Dim devuelve As Integer
    
    'IVA con recargo de equivalencia
    devuelve = TipoIVAOrigen
    If TipoIVAOrigen = TipoIVA1 Then
        devuelve = TipoIVAre1
    ElseIf TipoIVAOrigen = vParamAplic.TipoIVA2 Then
        devuelve = TipoIVAre2
    ElseIf TipoIVAOrigen = TipoIVA3 Then
        devuelve = TipoIVAre3
    Else
        MsgBox "NO se ha encontrado la correspondencia para el IVA RE: " & TipoIVAOrigen, vbExclamation
    End If
    DevuleveTipoIVA_RE = devuelve
End Function





Private Sub CargarTablaTraduccionIVAS()
Dim RS As ADODB.Recordset

    On Error GoTo ECargarTablaTraduccionIVAS
    If ConnConta Is Nothing Then
        MsgBox "Reinicie ARIGES y vuelva a modificar los parametros para cargar los tipos de IVA", vbExclamation
        Exit Sub
    End If
    
    SQL = "DELETE FROM sivatra"
    conn.Execute SQL
    
    
    Set RS = New ADODB.Recordset
    SQL = "select codigiva from tiposiva"
    RS.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        SQL = "INSERT INTO sivatra VALUES (" & RS!codigiva & "," & RS!codigiva & ")"
        
        conn.Execute SQL
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    'Ahora acutalizo la relacion
    If mvarTipoIVAre1 > 0 Then
        SQL = "UPDATE sivatra set codigiva2=" & mvarTipoIVAre1 & " WHERE codigiva = " & mvarIVA1
        conn.Execute SQL
    End If
    
    If mvarTipoIVAre2 > 0 Then
        SQL = "UPDATE sivatra set codigiva2=" & mvarTipoIVAre2 & " WHERE codigiva = " & mvarIVA2
        conn.Execute SQL
    End If
        
    If mvarTipoIVAre3 > 0 Then
        SQL = "UPDATE sivatra set codigiva2=" & mvarTipoIVAre3 & " WHERE codigiva = " & mvarIVA3
        conn.Execute SQL
    End If
        
    Exit Sub
ECargarTablaTraduccionIVAS:
    MuestraError Err.Number, Err.Description, "Cargar Tabla Traduccion IVAS"
End Sub





Public Function DevuelveTipoIVA_OLD(ConRecargo As Boolean, TipoIVAOrigen As Integer) As Integer
Dim devuelve As Integer
    
    devuelve = TipoIVAOrigen
    
    If ConRecargo Then
        'IVA con recargo de equivalencia
        If TipoIVAOrigen = TipoIVA1 Then
            devuelve = Me.OLDIVAre1
        ElseIf TipoIVAOrigen = vParamAplic.TipoIVA2 Then
            devuelve = Me.OLDIVAre2
        Else
            MsgBox "NO se ha encontrado la correspondencia para el ANTIGUO IVA RE: " & TipoIVAOrigen, vbExclamation
        End If
    
    
    Else
        If TipoIVAOrigen = TipoIVA1 Then
            devuelve = OLDIVA1
        ElseIf TipoIVAOrigen = vParamAplic.TipoIVA2 Then
            devuelve = OLDIVA2
        Else
            MsgBox "NO se ha encontrado la correspondencia para el ANTIGUO IVA: " & TipoIVAOrigen, vbExclamation
        End If
    
    
    
    End If
    DevuelveTipoIVA_OLD = devuelve
End Function
