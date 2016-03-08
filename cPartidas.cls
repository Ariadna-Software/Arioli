VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPartidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit




'Variables para las propiedades
Private mIdPartida As Long
Private mCodartic As String
Private mCodalmac As Integer
Private mNumlote As String
Private mCodProve As Long
Private mNumalbar As String
Private mFecha As Date
Private mCantidad As Currency

Private mEtiq As Integer

Dim Rs As ADODB.Recordset
Dim SQL As String




'**** Identificador de la partida
Public Property Let IdPartida(ByVal vData As Long)
     mIdPartida = vData
End Property

Public Property Get IdPartida() As Long
     IdPartida = mIdPartida
End Property



'**** Articulo de la partida
Public Property Let codArtic(ByVal vData As String)
     mCodartic = vData
End Property

Public Property Get codArtic() As String
     codArtic = mCodartic
End Property



'**** almacen de la partida
Public Property Let codAlmac(ByVal vData As Integer)
     mCodalmac = vData
End Property

Public Property Get codAlmac() As Integer
     codAlmac = mCodalmac
End Property



'**** almacen de la partida
Public Property Let NUmlote(ByVal vData As String)
     mNumlote = vData
End Property

Public Property Get NUmlote() As String
     NUmlote = mNumlote
End Property


'**** almacen de la partida
Public Property Let codProve(ByVal vData As Long)
     mCodProve = vData
End Property

Public Property Get codProve() As Long
     codProve = mCodProve
End Property



'**** almacen de la partida
Public Property Let NumAlbar(ByVal vData As String)
     mNumalbar = vData
End Property

Public Property Get NumAlbar() As String
     NumAlbar = mNumalbar
End Property


'**** almacen de la partida
Public Property Let Fecha(ByVal vData As Date)
     mFecha = vData
End Property

Public Property Get Fecha() As Date
     Fecha = mFecha
End Property


'**** cantidad
Public Property Let Cantidad(ByVal vData As Currency)
     mCantidad = vData
End Property

Public Property Get Cantidad() As Currency
     Cantidad = mCantidad
End Property



'***************************************************************************
'       METODOS
'***************************************************************************


'***************************************************************************
'LEER
Public Function Leer(ID As Long) As Boolean
    On Error GoTo ELee
    Leer = False
    Set Rs = New ADODB.Recordset
    SQL = "Select * from spartidas where id =" & ID
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Rs.EOF Then
        'MAAAAl. No existe
        
    Else
        AsignarValoresClase   'para poder llamar las sentecias en mas sitios
        Leer = True
    End If
    Rs.Close
        
    
ELee:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbExclamation
        Err.Clear
    End If
    Set Rs = Nothing
End Function

'Leer la clase con los valores del articulo
Public Function LeerDesdeArticulo(vCodArtic As String, vCodAlmac As Integer, cLote As String) As Boolean
    On Error GoTo ELeerDesdeArticulo
    LeerDesdeArticulo = False
    Set Rs = New ADODB.Recordset
    SQL = "Select * from spartidas where codartic = " & DBSet(vCodArtic, "T")
    SQL = SQL & " AND codalmac = " & CStr(vCodAlmac) & " AND numlote = " & DBSet(cLote, "T")
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Rs.EOF Then
        'Con los datos NO encuetra nada.
        
    Else
        AsignarValoresClase   'para poder llamar las sentecias en mas sitios
        LeerDesdeArticulo = True
    End If
    Rs.Close
    
ELeerDesdeArticulo:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set Rs = Nothing
End Function


'Leer la clase con los valores del articulo
'  Hasta el 4 de Julio, se fueron leyendo cajas palets pero no se daba
' de alta en lotes. Se hacia una produccion antigua.
' Cunado estamos en expediccion puede ser que el lote no lo tengamos con la codificacion
' de la TRZ bilidad nueva. Habria que ver si  lo tenemos en modo antiguo (4001 2012/05/24)
' Es decir
Public Function LeerDesdeExpedicion(vCodArtic As String, vCodAlmac As Integer, LoteExp As String) As Boolean
Dim OK As Boolean

    On Error GoTo ELeerDesdeArticulo
    LeerDesdeExpedicion = False
    Set Rs = New ADODB.Recordset
    SQL = "Select * from spartidas where codartic = " & DBSet(vCodArtic, "T")
    SQL = SQL & " AND codalmac = " & CStr(vCodAlmac) & " AND numlote = " & DBSet(LoteExp, "T")
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    OK = False
    If Rs.EOF Then
        'Con los datos NO encuetra nada.
        
    Else
        AsignarValoresClase   'para poder llamar las sentecias en mas sitios
        LeerDesdeExpedicion = True
        OK = True
    End If
    Rs.Close
    
    
    
    'Si no se ha encontrado NADA, miramos a ver si el parte de produccion se puso de
    'manera manual. Al estilo 4001 2012/05/24
    If Not OK Then
            SQL = "'" & Val(LoteExp) & " 2012/%'"
            SQL = " AND codalmac = " & CStr(vCodAlmac) & " AND numlote like " & SQL
            SQL = "Select * from spartidas where codartic = " & DBSet(vCodArtic, "T") & SQL
            
            Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            OK = False
            If Rs.EOF Then
                'Con los datos NO encuetra nada.
                
            Else
                AsignarValoresClase   'para poder llamar las sentecias en mas sitios
                LeerDesdeExpedicion = True
                OK = True
            End If
            Rs.Close
    End If
    
ELeerDesdeArticulo:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set Rs = Nothing
End Function




Public Function LeerDesdeAlbaran(vCodArtic As String, vCodAlmac As Integer, cLote As String, NumAlbar As String, codProve As Long) As Boolean
    On Error GoTo ELeerDesdeArticulo
    LeerDesdeAlbaran = False
    Set Rs = New ADODB.Recordset
    SQL = "Select * from spartidas where codartic = " & DBSet(vCodArtic, "T")
    SQL = SQL & " AND codalmac = " & CStr(vCodAlmac) & " AND numlote = " & DBSet(cLote, "T")
    SQL = SQL & " AND numalbar = " & DBSet(NumAlbar, "T") & " AND codprove = " & CStr(codProve)
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Rs.EOF Then
        'Con los datos NO encuetra nada.
        
    Else
        AsignarValoresClase   'para poder llamar las sentecias en mas sitios
        LeerDesdeAlbaran = True
    End If
    Rs.Close
    
ELeerDesdeArticulo:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set Rs = Nothing
End Function





Public Sub IncrementarCantidad(Incremento As Currency)
    On Error GoTo EIns
    
    SQL = "UPDATE spartidas set cantotal=cantotal "
    If Incremento >= 0 Then SQL = SQL & "+ "
    SQL = SQL & TransformaComasPuntos(CStr(Incremento))
    SQL = SQL & " WHERE id = " & mIdPartida
    Conn.Execute SQL
    Exit Sub
EIns:
    MsgBox Err.Description, vbExclamation
    
End Sub



'CUBAS-DEPOSITOS
'Cuando es fin deposito, se regulzarizara, y la cantidad hay que dejarla a cero
Public Sub FinPartida()
    On Error GoTo EIns
    
    SQL = "UPDATE spartidas set cantotal=0 "
    SQL = SQL & " WHERE id = " & mIdPartida
    Conn.Execute SQL
    Exit Sub
EIns:
    MsgBox Err.Description, vbExclamation
    
End Sub


Public Function Insertar() As Boolean
    On Error GoTo EIns
    Insertar = False
    
    'Siguiente
    mIdPartida = Siguiente
    'Insetamos
    SQL = "insert into spartidas (id,codartic,codalmac,numlote,codprove,numalbar,fecha,cantotal) values ("
    SQL = SQL & mIdPartida & "," & DBSet(mCodartic, "T") & "," & mCodalmac & ","
    SQL = SQL & DBSet(mNumlote, "T") & "," & DBSet(mCodProve, "N")
    SQL = SQL & "," & DBSet(mNumalbar, "T") & "," & DBSet(mFecha, "F") & "," & DBSet(mCantidad, "N") & ")"
    Conn.Execute SQL
    Insertar = True 'todo ok
    Exit Function
EIns:
    MsgBox Err.Description, vbExclamation
    
End Function



Public Function Eliminar() As Boolean
    On Error GoTo EIns
    Eliminar = False
    
    SQL = "DELETE FROM spartidaslin WHERE id = " & mIdPartida
    Conn.Execute SQL
    
    SQL = "DELETE FROM spartidas WHERE id = " & mIdPartida
    Conn.Execute SQL
    Eliminar = True 'todo ok
    Exit Function
EIns:
    MsgBox Err.Description, vbExclamation
    
End Function



'Devuelve.
'           0. Ok
'           1. No tenemos para todos
'           2. Error
Public Function RecuperarLotes(vCodArtic As String, vCodAlmac As Integer, ByVal CantidadNecesaria As Currency, ByRef kLotes As Collection) As Byte
'   CantidadNecesaria. La pasamos byVal para poder modificarla aqui dentro
Dim Ca As Currency
Dim Fin As Boolean
Dim AlgunoAsignado As Boolean
    

    On Error GoTo ERecuperarLotes
    RecuperarLotes = 2
    Set Rs = New ADODB.Recordset
    SQL = "Select * from spartidas where codartic = " & DBSet(vCodArtic, "T")
    SQL = SQL & " AND codalmac = " & CStr(vCodAlmac)
    SQL = SQL & " ORDER BY fecha asc"
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        'Con los datos NO encuetra nada.
        AsignarValoresClase   'para poder llamar las sentecias en mas sitios
        AlgunoAsignado = False
        'Ahora veremos, a partir de la cantidad QUE necesitamos, si podemos darle los numeros de serie
        Set kLotes = New Collection
        Fin = False
        While Not Fin
            Ca = Rs!cantotal
            If Ca > 0 Then
                If Ca < CantidadNecesaria Then
                    'En este lote articulo no tengo bastante
                    CantidadNecesaria = CantidadNecesaria - Ca
                    
                    
                Else
                    'Tengo de sobra
                    Ca = CantidadNecesaria
                    CantidadNecesaria = 0 'Para saber que hemos tenido bien
                    Fin = True
                End If
                AlgunoAsignado = True
                SQL = Rs!ID & "|" & Ca & "|" & Rs!NUmlote & "|"
                kLotes.Add SQL
            End If
            
            Rs.MoveNext
            If Rs.EOF Then Fin = True
            
        Wend
        
        
        
        If CantidadNecesaria > 0 Then
            'No tenemos bastante del articulo para sacar
         '   MsgBox "No hay suficiente cantidad de lotes para el articulo. ", vbExclamation
            If AlgunoAsignado Then RecuperarLotes = 1
        Else
            'Todo ok
            RecuperarLotes = 0
        End If
        
    End If
    Rs.Close
    
ERecuperarLotes:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set Rs = Nothing
End Function



'Cargamos la tabla tmporal para la muestra de datos relacionados con la partida o  lote/articulo
Public Function CargaDatosParaListar() As Boolean
Dim J As Integer

    On Error GoTo ECargaDatosParaListar
    CargaDatosParaListar = False
    'Se da por ententido que la ya se ha hecho el metodo. leer
    
    'Ira comprobando en las tablas:
    ' olicoupagelinlotes, slialblotes,slifaclotes,sliordpr2lotes,sliordprlotes,slialp,slifpc
    'JUNIO 14. Ira a smovalotes para DEP, que son los finales de deposito
    
    'insert into `tmppartidas` (`codusu`,`idpartida`,`codartic`,`numlote`,fecha,`idOperacion`,`idNumOperacion`,`cantidad`) values
    
    Set Rs = New ADODB.Recordset
    
    
    For J = 1 To 12
        'olicoupage,olicoupagelinlotes , slialblotes, slifaclotes, sliordpr2lotes, sliordprlotes, slialp,
        'slifpc,slhmov,prodtrazalin,prodtrazcompo,smovalotes
        LoteEnTabla2 J
    Next
    Set Rs = Nothing
    CargaDatosParaListar = True
    Exit Function
ECargaDatosParaListar:
    Set Rs = Nothing
    MuestraError Err.Number, Err.Description
End Function




'Comprobara la suma de cantidades con la cantidad almacenada
Public Sub AjustarCantidad(Preguntar As Boolean)
Dim Real As Currency

    'Borro tmp
    SQL = "DELETE from tmppartidas where codusu = " & vUsu.codigo
    Conn.Execute SQL
    If CargaDatosParaListar Then
        'Ya tengo cargado el
        
        
        Set Rs = New ADODB.Recordset
        SQL = "Select sum(cantidad) from tmppartidas where codusu = " & vUsu.codigo & " AND idpartida=" & Me.IdPartida
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Real = 0
        If Not Rs.EOF Then Real = DBLet(Rs.Fields(0), "N")
        Rs.Close
        Set Rs = Nothing
        
        If Real <> Me.Cantidad Then
            If Preguntar Then
                SQL = "Cantidades distintas." & vbCrLf
                SQL = SQL & "Partida: " & Format(Me.Cantidad, FormatoCantidad) & vbCrLf
                SQL = SQL & "Mov.: " & Format(Real, FormatoCantidad) & vbCrLf
                SQL = SQL & "¿Continuar?"
                If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then SQL = ""
            Else
                SQL = ""
            End If
            If SQL = "" Then
                    SQL = "UPDATE spartidas set cantotal=" & TransformaComasPuntos(CStr(Real))
                    SQL = SQL & " WHERE id = " & mIdPartida
                    EjecutaSQL conAri, SQL, True
            End If
        End If
    End If
    
End Sub








'#################################################################################################
'Privados

Private Sub AsignarValoresClase()
    'No lleva error. Que se dispare el erro en el sub que lo llama
    mIdPartida = Rs!ID
    mCodartic = Rs!codArtic
    mCodalmac = Rs!codAlmac
    mNumlote = Rs!NUmlote
    mCodProve = Rs!codProve
    mNumalbar = Rs!NumAlbar
    mFecha = Rs!Fecha
    mCantidad = Rs!cantotal
End Sub


Private Function Siguiente() As Long
    SQL = "select max(id) from spartidas"
    Set Rs = New ADODB.Recordset
    Siguiente = 1
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then Siguiente = Rs.Fields(0) + 1
    End If
    Rs.Close
    Set Rs = Nothing
End Function


'EN SQL tengo el trozo de insert into tabla/v1,v2) vALUES
'Tablas:
'   olicoupage,olicoupagelinlotes , slialblotes, slifaclotes, sliordpr2lotes, sliordprlotes, slialp, slifpc
'
'
' El insert
'   `codusu`,`idpartida`,`codartic`,`numlote`,fecha,
'   `idOperacion`,`idNumOperacion`,  idOperacion:  que es             idNumOperacion clave
' `idReferencia`,`Referencia `,`cantidad`     Ref: cliente/proveedor/trabajador
Private Sub LoteEnTabla2(Num As Integer)
Dim Cadena1 As String
Dim Desc As String
Dim miSql As String
Dim Ca As Currency
Dim SumaLotes As Boolean
    
    Select Case Num
    Case 1
        'Lineas de coupage
        Desc = "'COUPAGE'"
        miSql = "select olicoupage.codigo clave, sum(kilos) cantmostrar ,olicoupage.fecha mifecha, 0 idRefd,' ' vRefd from"
        miSql = miSql & " olicoupage,olicoupagelin where olicoupage.codigo=olicoupagelin.codigo"
        'El arti/lote/almacen
        miSql = miSql & " AND olicoupage.codartic = " & DBSet(mCodartic, "T") & " AND olicoupage.numlote = " & DBSet(mNumlote, "T")
        miSql = miSql & " and codalmac = " & mCodalmac
        miSql = miSql & " group by 1,3"
        SumaLotes = True
    Case 2
    'Lineas de coupage
        Desc = "'COUPAGE Lin'"
        miSql = "select olicoupagelinlotes.codigo clave,cantlote cantmostrar,fecha mifecha ,0 idRefd,' ' vRefd"
        miSql = miSql & " FROM olicoupagelinlotes,olicoupage WHERE  olicoupagelinlotes.codigo=olicoupage.codigo"
        'El arti/lote/almacen
        miSql = miSql & " AND olicoupagelinlotes.codartic = " & DBSet(mCodartic, "T") & " AND olicoupagelinlotes.numlote = " & DBSet(mNumlote, "T")
        miSql = miSql & " and codalmac = " & mCodalmac
        SumaLotes = False
    
    Case 3
        'Lineas albaran
        Desc = "'Albaran cliente'"
        miSql = "select concat(slialblotes.codtipom,slialblotes.numalbar) clave,slialblotes.cantidad cantmostrar,"
        miSql = miSql & " fechaalb mifecha, codclien idrefd,nomclien vrefd "
        miSql = miSql & " from scaalb,slialblotes,slialb where"
        miSql = miSql & " scaalb.codtipom=slialb.codtipom and scaalb.numalbar=slialblotes.numalbar and"
        miSql = miSql & " slialblotes.codTipoM = slialb.codTipoM And slialb.NumAlbar = slialblotes.NumAlbar"
        miSql = miSql & " and slialb.numlinea=slialblotes.numlinea"
        'El arti/lote/almacen
        miSql = miSql & " AND slialb.codartic = " & DBSet(mCodartic, "T") & " AND numlote = " & DBSet(mNumlote, "T")
        miSql = miSql & " and codalmac = " & mCodalmac
        SumaLotes = False
        
    Case 4
        'lineas facturas
        
        
         
        
        
        Desc = "'FACTURAS CLIENTE'"
        miSql = "select concat(slifaclotes.codtipom,slifaclotes.numfactu) clave,slifaclotes.cantidad cantmostrar,"
        miSql = miSql & " scafac.fecfactu mifecha, codclien idrefd,nomclien vrefd  from"
       ' miSQL = miSQL & " scafac,slifac,slifaclotes where scafac.codtipom=slifac.codtipom"
       ' miSQL = miSQL & " and scafac.numfactu=slifac.numalbar and scafac.fecfactu=slifac.numalbar"
       ' miSQL = miSQL & " and scafac.numfactu=slifaclotes.numfactu and scafac.fecfactu=slifaclotes.fecfactu and"
       ' miSQL = miSQL & " scafac.codTipoM = slifaclotes.codTipoM and slifac.numlinea=slifaclotes.numlinea "
        'AHORA MAYO 2010
        miSql = miSql & " (`scafac` `scafac` INNER JOIN `slifac` `slifac` ON ((`scafac`.`codtipom`=`slifac`.`codtipom`) "
        miSql = miSql & " AND (`scafac`.`numfactu`=`slifac`.`numfactu`)) AND (`scafac`.`fecfactu`=`slifac`.`fecfactu`))"
        miSql = miSql & " INNER JOIN `slifaclotes` `slifaclotes` ON (((((`slifac`.`codtipom`=`slifaclotes`.`codtipom`)"
        miSql = miSql & " AND (`slifac`.`numfactu`=`slifaclotes`.`numfactu`)) AND (`slifac`.`fecfactu`=`slifaclotes`.`fecfactu`))"
        miSql = miSql & " AND (`slifac`.`codtipoa`=`slifaclotes`.`codtipoa`)) AND (`slifac`.`numalbar`=`slifaclotes`.`numalbar`))"
        miSql = miSql & " AND (`slifac`.`numlinea`=`slifaclotes`.`numlinea`)"

       
        'El arti/lote/almacen
        miSql = miSql & " AND slifac.codartic = " & DBSet(mCodartic, "T") & " AND numlote = " & DBSet(mNumlote, "T")
        miSql = miSql & " and codalmac = " & mCodalmac
        SumaLotes = False
        
    Case 5
        'PRoduccion
        'lineas SUMA
        Desc = "'PRODUCCION'"
        miSql = "select sordprod.codigo clave,cantlote cantmostrar,fecproduccion mifecha ,0 idRefd,' ' vRefd"
        miSql = miSql & " FROM sordprod,sliordprlotes WHERE sordprod.codigo = sliordprlotes.codigo"
        'Para que saque solo los que estan producidos
        miSql = miSql & " AND fecproduccion>'1900-01-01'"
        
        'El arti/lote/almacen
        miSql = miSql & " AND sliordprlotes.codartic = " & DBSet(mCodartic, "T") & " AND numlote = " & DBSet(mNumlote, "T")
        miSql = miSql & " and codalmac = " & mCodalmac
        SumaLotes = True
    Case 6
        'Sub lineas restan
        Desc = "'PRODUCCION compo.'"
        miSql = "select sordprod.codigo clave,cantlote cantmostrar"
        miSql = miSql & ",fecproduccion mifecha ,0 idRefd,' ' vRefd"
        miSql = miSql & " FROM sordprod,sliordpr2lotes,sartic WHERE sordprod.codigo = sliordpr2lotes.codigo"
        miSql = miSql & " AND sliordpr2lotes.codarti2=sartic.codartic"
        'Para que saque solo los que estan producidos
        miSql = miSql & " AND fecproduccion>'1900-01-01'"
        
        'El arti/lote/almacen
        miSql = miSql & " AND sliordpr2lotes.codarti2 = " & DBSet(mCodartic, "T") & " AND numlote = " & DBSet(mNumlote, "T")
        miSql = miSql & " and codalmac = " & mCodalmac
        SumaLotes = False
        
        
    Case 7
        'ALbaranes proveedor
        Desc = "'Albaranes proveedor'"
        miSql = "select slialp.numalbar clave,cantidad cantmostrar,scaalp.fechaalb mifecha ,scaalp.codprove idRefD,scaalp.nomprove vrefd"
        miSql = miSql & " from slialp,scaalp where scaalp.codprove=slialp.codprove and scaalp.fechaalb=slialp.fechaalb and scaalp.numalbar=slialp.numalbar"
        'El arti/lote/almacen                                               'ojo a la S
        miSql = miSql & " AND slialp.codartic = " & DBSet(mCodartic, "T") & " AND numlotes = " & DBSet(mNumlote, "T")
        miSql = miSql & " and codalmac = " & mCodalmac
        SumaLotes = True
        
    Case 8
        'Facturas proveedor
         Desc = "'FACTURAS PROVEEDOR'"
        miSql = "select scafpc.numfactu clave,cantidad cantmostrar,scafpc.fecfactu mifecha,scafpc.codprove idrefd,scafpc.nomprove vrefd"
        miSql = miSql & " FROM slifpc,scafpc where scafpc.codprove=slifpc.codprove and scafpc.fecfactu=slifpc.fecfactu"
        miSql = miSql & " and scafpc.numfactu=slifpc.numfactu "
        'El arti/lote/almacen                                                     'ojo a la S de numloteSSSS
        miSql = miSql & " AND slifpc.codartic = " & DBSet(mCodartic, "T") & " AND numlotes = " & DBSet(mNumlote, "T")
        miSql = miSql & " and codalmac = " & mCodalmac
        SumaLotes = True
        
        
        
    Case 9
        'REGULARIZACION
         Desc = "'REGULARIZACION'"
        miSql = "select schmov.codmovim clave,if(tipomovi=0,-cantidad,cantidad) cantmostrar,schmov.fecmovim mifecha,schmov.codtraba idrefd,nomtraba vrefd"
        miSql = miSql & " from slhmov,schmov left join straba on schmov.codtraba=straba.codtraba WHERE slhmov.codmovim=schmov.codmovim"
 
        'El arti/lote/almacen
        miSql = miSql & " AND slhmov.codartic = " & DBSet(mCodartic, "T") & " AND numlote = " & DBSet(mNumlote, "T")
        miSql = miSql & " and schmov.codalmac = " & mCodalmac
        SumaLotes = True
        
        
    Case 10
        'NUEA PRODUCCION
        Desc = "'PRO.TRAZABILIDAD'"
        
        'FALTA###
        'comprobar que es numerico
        If Not IsNumeric(mNumlote) Then Exit Sub
        'prodtrazlin
        miSql = "select concat(prodcab.codigo,""/"",prodlin.idlin) clave,prodtrazlin.cantprodu cantmostrar,fhinicio mifecha,0 idRefd,' ' vRefd"
        miSql = miSql & " from prodtrazlin,prodlin,prodcab where prodcab.codigo=prodlin.codigo  and prodtrazlin.codigo=prodlin.codigo and prodtrazlin.idlin=prodlin.idlin"
        'El arti/lote/almacen
        miSql = miSql & " AND prodlin.codartic = " & DBSet(mCodartic, "T") & " AND lotetraza = " & Val(mNumlote)
        miSql = miSql & " and almacen2 = " & mCodalmac
        SumaLotes = True
    Case 11
        Desc = "'PRO.TRAZA.COMPO.'"
        
        'FALTA###
        'comprobar que es numerico
        
        'prodtrazlin
        miSql = "select concat(prodcab.codigo,""/"",prodtrazcompo.idlin,"" "") clave,prodtrazcompo.cantutili cantmostrar,"
        miSql = miSql & "fhinicio mifecha,prodtrazcompo.lineaprod idRefd,concat('Idtraza: ',prodtrazcompo.lotetraza) vRefd"
        miSql = miSql & " from prodtrazcompo,prodlin,prodcab where prodcab.codigo=prodlin.codigo  and prodtrazcompo.codigo=prodlin.codigo and prodtrazcompo.idlin=prodlin.idlin"
        'El arti/lote/almacen
        miSql = miSql & " AND prodtrazcompo.codartic = " & DBSet(mCodartic, "T") & " AND numlote = " & DBSet(mNumlote, "T")
        miSql = miSql & " and almacen2 = " & mCodalmac
        miSql = miSql & " and cantutili>0"
        SumaLotes = False
        
        
    Case 12
        'FIN DEPOSITO
        Desc = "'FIN DEPOSITO CUP'"
        miSql = "select document clave,if(tipomovi=1,-cantidad,cantidad) cantmostrar,fechamov mifecha ,0 idRefd,' ' vRefd"
        miSql = miSql & " FROM smovalotes WHERE  "
        miSql = miSql & " detamovi='DEP' AND codartic = " & DBSet(mCodartic, "T") & " AND numlote = " & DBSet(mNumlote, "T")
        miSql = miSql & " and codalmac = " & mCodalmac
        SumaLotes = False
        
        
    End Select
    Rs.Open miSql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    ' El insert
'   `codusu`,`idpartida`,`codartic`,`numlote`,fecha,`idOperacion`,`idNumOperacion`,
' `idReferencia`,`Referencia `,`cantidad` ,abs_cantidad    Ref: cliente/proveedor/trabajador
    SQL = ", (" & vUsu.codigo & "," & mIdPartida & "," & DBSet(mCodartic, "T") & "," & DBSet(mNumlote, "T") & ","
    miSql = ""
    While Not Rs.EOF
        miSql = miSql & SQL 'codusus,partida,numlote
    
        'fecha, idoperacion,idnumoperacion
        miSql = miSql & DBSet(Rs!mifecha, "F") & "," & Desc & "," & DBSet(Rs!Clave, "T")
        'idreferencia, referencia
        miSql = miSql & "," & Rs!idrefd & "," & DBSet(Rs!vrefd, "T") & ","
        'cantidad, y valor absoluto de la cantidad
        Ca = Rs!cantmostrar
        If Not SumaLotes Then Ca = -1 * Ca
        miSql = miSql & DBSet(Ca, "N") & "," & DBSet(Rs!cantmostrar, "N") & ")"
        Rs.MoveNext
    Wend
    Rs.Close
    
    If miSql <> "" Then
        SQL = "insert into `tmppartidas` (`codusu`,`idpartida`,`codartic`,`numlote`,`fecha`,`idOperacion`,`idNumOperacion`,`idReferencia`,`Referencia`,`cantidad`,abs_cantidad) values"
        SQL = SQL & Mid(miSql, 2)
        Conn.Execute SQL
    End If
    
End Sub




'Ya debe estar hecho el .Leer
Public Function TrazbilidadDesdeCompra(ByRef L As Label, DetallarCoupages As Boolean) As Boolean
Dim vContador As Long
Dim MiNivel As Integer
        
    MiNivel = 0
    vContador = 0
    
    GeneracionProcesosDesdeCompra MiNivel, vContador, Me.codArtic, Me.NUmlote, L, DetallarCoupages
    
End Function



Public Sub GeneracionProcesosDesdeCompra(Nivel As Integer, ByRef Contador As Long, ByVal vArticulo As String, ByVal vLote As String, ByRef LB As Label, DetallaLosCoup As Boolean)
Dim Col As Collection
Dim RN As ADODB.Recordset
Dim Aux As String
Dim J As Integer
Dim H As Integer
Dim Col2 As Collection
Dim Cantidad As Currency

Dim Cant_Conver As Currency
Dim Aux2 As String
Dim FactorMultiplicador1 As Currency

    LB.Caption = Nivel & "-> " & vArticulo
    LB.Refresh
    'Si esta en compras, solo el primer LOTE
    'Debug.Print vArticulo & " " & vLote
    If Nivel = 0 Then
        'Las compras del articulo ORIGEN solo
        EnAlbaranCompras Nivel, vArticulo, vLote, Contador
    Else
      '  Stop
    End If
         
    'En COUPAGES
    Set Col = New Collection
    Set RN = New ADODB.Recordset
    
    Aux = "select olicoupage.codartic,olicoupage.numlote, olicoupagelinlotes.cantlote,olicoupage.codigo from"
    Aux = Aux & " olicoupage , olicoupagelinlotes where olicoupage.codigo = olicoupagelinlotes.codigo"
    Aux = Aux & " AND olicoupagelinlotes.codartic = " & DBSet(vArticulo, "T")
    Aux = Aux & " AND olicoupagelinlotes.numlote = " & DBSet(vLote, "T")
    RN.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RN.EOF
        Aux = RN!codArtic & "|" & RN!NUmlote & "|" & RN!cantlote & "|" & RN!codigo & "|"
        Col.Add Aux
        RN.MoveNext
    Wend
    RN.Close
    If Col.Count > 0 Then
        

        For J = 1 To Col.Count
            Aux = "Coupage (" & RecuperaValor(Col.Item(J), 4) & ")"
            InsertarLinea Aux, Nivel, CStr(RecuperaValor(Col.Item(J), 1)), CStr(RecuperaValor(Col.Item(J), 2)), Contador, CCur(CStr(RecuperaValor(Col.Item(J), 3))), ""
            If DetallaLosCoup Then GeneracionProcesosDesdeCompra Nivel + 1, Contador, CStr(RecuperaValor(Col.Item(J), 1)), CStr(RecuperaValor(Col.Item(J), 2)), LB, True
        Next J
    End If
    
    'AHORA produccion
    Aux = "select sliordprlotes.codartic,sliordprlotes.numlote,sliordprlotes.cantlote,sliordprlotes.codigo from sliordprlotes ,sliordpr2lotes"
    Aux = Aux & " Where sliordprlotes.codigo = sliordpr2lotes.codigo"
    Aux = Aux & " and sliordprlotes.codalmac =sliordpr2lotes.codalmac"
    Aux = Aux & " and sliordprlotes.codartic =sliordpr2lotes.codartic"
    Aux = Aux & " AND sliordpr2lotes.codarti2 = " & DBSet(vArticulo, "T")
    Aux = Aux & " AND sliordpr2lotes.numlote = " & DBSet(vLote, "T")
    RN.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set Col = New Collection
    While Not RN.EOF
        Aux = RN!codArtic & "|" & RN!NUmlote & "|" & RN!cantlote & "|" & RN!codigo & "|"
        Col.Add Aux
        RN.MoveNext
    Wend
    RN.Close
    
    If Col.Count > 0 Then
        
        For J = 1 To Col.Count
            Aux = "Produc. (" & RecuperaValor(Col.Item(J), 4) & ")"
            InsertarLinea Aux, Nivel, CStr(RecuperaValor(Col.Item(J), 1)), CStr(RecuperaValor(Col.Item(J), 2)), Contador, CCur(CStr(RecuperaValor(Col.Item(J), 3))), ""
            GeneracionProcesosDesdeCompra Nivel + 1, Contador, CStr(RecuperaValor(Col.Item(J), 1)), CStr(RecuperaValor(Col.Item(J), 2)), LB, True
        Next J
    End If
    Set Col = Nothing
    
    ' If Nivel = 0 Then Stop
    Aux = "select * from prodtrazcompo where"
    Aux = Aux & " prodtrazcompo.codartic = " & DBSet(vArticulo, "T")
    Aux = Aux & " AND prodtrazcompo.numlote = " & DBSet(vLote, "T")
    'Veremos todos los lotes nuevos
    
    Set Col2 = New Collection
    Set Col = New Collection
    RN.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RN.EOF
        Col2.Add CStr(RN!lotetraza)
        RN.MoveNext
    Wend
    RN.Close
    
    For H = 1 To Col2.Count
        Aux = "select prodlin.codartic,right(concat(""0000000000"",prodtrazlin.lotetraza),10) el_lote,"
        Aux = Aux & " prodtrazlin.cantprodu,concat(prodtrazlin.idlin,'/',prodtrazlin.codigo) elcodigo"
        Aux = Aux & " from prodlin,prodtrazlin where prodtrazlin.codigo=prodlin.codigo and"
        Aux = Aux & " prodtrazlin.idlin = prodlin.idlin And lotetraza = " & Col2.Item(H)
        Aux = Aux & " order by prodtrazlin.lotetraza"
        RN.Open Aux, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not RN.EOF
            Aux = RN!codArtic & "|" & RN!el_lote & "|" & RN!cantprodu & "|" & RN!elcodigo & "|"
            Col.Add Aux
            RN.MoveNext
        Wend
        RN.Close
    Next H
    
    If Col.Count > 0 Then
        
        For J = 1 To Col.Count
            Aux = RecuperaValor(Col.Item(J), 3)
            If Aux = "" Then Aux = "0"
            Cantidad = CCur(Aux)
            
            
            'select cantidad*factorconversion from sarti1,sartic where sarti1.codarti1=sartic.codartic
            'and sarti1.codartic='002000013502' and codarti1='003700441513'
            Aux = "factorconversion"
            Aux2 = RecuperaValor(Col.Item(J), 1)
            Aux2 = "sarti1.codarti1=sartic.codartic AND sarti1.codartic= '" & Aux2 & "' and codarti1"
            Aux2 = DevuelveDesdeBD(conAri, "cantidad*factorconversion", "sarti1,sartic", Aux2, vArticulo, "T", Aux)
            If Aux2 = "" Then Aux2 = "1"
            Cant_Conver = CCur(Aux2)
            FactorMultiplicador1 = Cant_Conver

            Cantidad = Round(Cantidad * Cant_Conver, 2)
            
            Aux = "Prod TRZ. (" & RecuperaValor(Col.Item(J), 4) & ")"
            
            InsertarLinea Aux, Nivel, CStr(RecuperaValor(Col.Item(J), 1)), CStr(RecuperaValor(Col.Item(J), 2)), Contador, Cantidad, ""
            GeneracionProcesosDesdeCompra Nivel + 1, Contador, CStr(RecuperaValor(Col.Item(J), 1)), CStr(RecuperaValor(Col.Item(J), 2)), LB, True
        Next J
    End If
 
 
 
 
    Set Col = Nothing
    
    
    'Veremos para este articulo LOTE las ventas
    If Nivel = 0 Then
        Aux = "Select cantidad ,tipomovi,numlinea from smovalotes where codartic=" & DBSet(vArticulo, "T") & " AND numlote = " & DBSet(vLote, "T")
        Aux = Aux & " AND detamovi in ('DEP')"
        RN.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RN.EOF Then
            Aux = "Regularizacion deposito " & RN!numlinea
            Cantidad = RN!Cantidad
            InsertarLinea Aux, Nivel, vArticulo, "", Contador, Cantidad, ""
                
        End If
    End If
    
    
    
    'Si tiene alguna regularizacion
    Set RN = Nothing
    
    
    
    
    
End Sub

Private Function EnAlbaranVentas_(vNivel As Integer, vArt As String, NLote As String, ByRef Contador As Long, FactorMultiplicador As Currency) As Boolean
Dim RN As ADODB.Recordset
Dim S As String
Dim ProveClie As String

    On Error GoTo EEnAlbaranVentas
    EnAlbaranVentas_ = False
    Set RN = New ADODB.Recordset
    S = "Select * from slialb,slialblotes    Where slialb.codTipoM = slialblotes.codTipoM"
    S = S & " AND codartic = '" & vArt & "' and codalmac = " & Me.codAlmac
    S = S & " AND numlote = " & DBSet(NLote, "T")


    RN.Open S, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RN.EOF
        ProveClie = "numalbar = " & RN!NumAlbar & " AND codtipom"
        ProveClie = DevuelveDesdeBD(conAri, "nomclien", "scaalb", ProveClie, RN!codTipoM, "T")
        InsertarLinea RN!codTipoM & RN!NumAlbar, vNivel, vArt, "", Contador, RN!Cantidad, ProveClie

        RN.MoveNext
    Wend
    RN.Close
    
    'En factras
    S = "Select * from slifac,slifaclotes    Where   slifac.codTipoM = slifaclotes.codTipoM"
    S = S & " and slifac.numfactu  = slifaclotes.numfactu and slifac.fecfactu  = slifaclotes.fecfactu"
    S = S & " and slifac.codtipoa = slifaclotes.codtipoa and slifac.numalbar = slifaclotes.numalbar"
    S = S & " and slifac.numlinea = slifaclotes.numlinea  "
    S = S & " AND codartic = '" & vArt & "' and codalmac = " & Me.codAlmac
    S = S & " AND numlote = " & DBSet(NLote, "T")
    

    RN.Open S, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RN.EOF
        
        ProveClie = "codtipom = '" & RN!codTipoM & "' AND fecfactu =" & DBSet(RN!FecFactu, "F") & " AND numfactu"
        ProveClie = DevuelveDesdeBD(conAri, "nomclien", "scafac", ProveClie, RN!NumFactu, "T")
        
        
        InsertarLinea RN!codTipoM & RN!NumFactu, vNivel, vArt, "", Contador, RN!Cantidad, ProveClie

        RN.MoveNext
    Wend
    RN.Close
    
    
    Exit Function
EEnAlbaranVentas:
    MuestraError Err.Number
End Function



Private Function EnAlbaranCompras(vNivel As Integer, vArt As String, NLote As String, ByRef Contador As Long) As Boolean
Dim RN As ADODB.Recordset
Dim S As String
Dim DesdeUnAlbaran As Boolean

    On Error GoTo EEnAlbaranCompras
    EnAlbaranCompras = False
    DesdeUnAlbaran = False
    Set RN = New ADODB.Recordset
    S = "select slialp.numalbar clave,cantidad cantmostrar,scaalp.fechaalb mifecha ,scaalp.codprove idRefD,scaalp.nomprove vrefd"
    S = S & " from slialp,scaalp where scaalp.codprove=slialp.codprove and scaalp.fechaalb=slialp.fechaalb and scaalp.numalbar=slialp.numalbar"
    'El arti/lote/almacen                                               'ojo a la S
    S = S & " AND slialp.codartic = " & DBSet(vArt, "T") & "AND numlotes = " & DBSet(NLote, "T")
    S = S & " and codalmac = " & mCodalmac



    RN.Open S, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RN.EOF
        DesdeUnAlbaran = True
        InsertarLinea "ALP " & RN!Clave, vNivel, vArt, "", Contador, RN!cantmostrar, RN!vrefd

        RN.MoveNext
    Wend
    RN.Close
    
    'En factras
    S = "select scafpc.numfactu clave,cantidad cantmostrar,scafpc.fecfactu mifecha,scafpc.codprove idrefd,scafpc.nomprove vrefd"
    S = S & " FROM slifpc,scafpc where scafpc.codprove=slifpc.codprove and scafpc.fecfactu=slifpc.fecfactu"
    S = S & " and scafpc.numfactu=slifpc.numfactu "
    'El arti/lote/almacen                                                     'ojo a la S
    S = S & " AND slifpc.codartic = " & DBSet(vArt, "T") & "AND numlotes = " & DBSet(NLote, "T")
    S = S & " and codalmac = " & mCodalmac
    RN.Open S, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RN.EOF
        InsertarLinea "F.PRO " & RN!Clave, vNivel, vArt, "", Contador, RN!cantmostrar, RN!vrefd
        DesdeUnAlbaran = True
        RN.MoveNext
    Wend
    RN.Close
    
    
    If Not DesdeUnAlbaran Then
        'NO es de un albaran. Podriamos ver, si el lote acaba en C
        'que viene de un copupage
        If Right(NLote, 1) = "C" Then
            S = "select olicoupage.codigo,sum(kilos) cantmostrar,deposito from olicoupage,olicoupagelin WHERE "
            S = S & " olicoupage.codigo=olicoupagelin.codigo AND "
            S = S & " numlote =" & DBSet(NUmlote, "T") & " And codAlmac = " & mCodalmac & "  group by 1"
            RN.Open S, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RN.EOF Then
                S = "CUP" & Format(RN!codigo, "0000")
                If RN!Deposito > 0 Then S = S & "[" & RN!Deposito & "]"
                InsertarLinea S, vNivel, vArt, "", Contador, RN!cantmostrar, S
                
                RN.MoveNext
            End If
            RN.Close
        End If
    End If
    
    
    
    
    EnAlbaranCompras = True
    Exit Function
EEnAlbaranCompras:
    MuestraError Err.Number
End Function







Private Sub InsertarLinea(Clave As String, vNivel As Integer, vArt As String, vL As String, ByRef Contador As Long, Cantidad As Currency, NomClienProv As String)
Dim SQL As String

    SQL = "insert into `tmptraza` (`codusu`,`contador`,`artppal`,`numloteppal`,"
    SQL = SQL & "`nivle`,`idOperacion`,`artic2`,`numlote2`,`cantidad`,`abs_cantidad`,nomclien) values ("
    'Valroes
    SQL = SQL & vUsu.codigo & "," & Contador & "," & DBSet(Me.codArtic, "T") & "," & DBSet(Me.NUmlote, "T")
    SQL = SQL & "," & vNivel & "," & DBSet(Clave, "T") & "," & DBSet(vArt, "T") & "," & DBSet(vL, "T", "N")
    SQL = SQL & "," & TransformaComasPuntos(CStr(Cantidad)) & ",0," & DBSet(NomClienProv, "T", "S") & ")"
    Conn.Execute SQL
    Contador = Contador + 1
End Sub




'-----------------------------------
'Trazbilidad 2
'
'
' Veremos para un lote comprado
'       -Ventas directas del lote
'       -producciones realizadas del lote
'           -Ventas de esa produccion
'           -mas producciones realizadas
'       -cupajes realizados
'           -ventas del cupaje
'           -produccion del cupaje
Public Function TrazabilidadDesdeCompra() As Boolean
Dim ColComp As Collection
Dim ColVentas As Collection
Dim TrozoLineaIntermedia
Dim J As Integer

    TrazabilidadDesdeCompra = False

    'Albaran compra y ventas directas del lote
    Set ColComp = New Collection
    
    If Not EnAlbaranCompras2(ColComp) Then Exit Function
    
    
    'INSERT
    'insert into `tmptraza2` (`codusu`,`contador`,`artppal`,`numloteppal`,`cantidad1`,
    '`albprov`,`proveedor`,`idCup`,`fechacup`,`idProd`,`fechaprod`,
    '`albventa`,fechaventa,`nomclien`,`cantidad2`)
    
    'OK. Tengo el /los albaranes de entrada
    'VENTAS DIRECTAS
    
    For J = 1 To ColComp.Count
        Set ColVentas = New Collection
        EnAlbaranVentas2 ColVentas, Me.codArtic, NUmlote
        
    Next J
    
    
    
    'AHora veremos si esta cupado, y/o en produccion
        


End Function



'Le pasamos una coleccion para que la rellene
Private Function EnAlbaranCompras2(ByRef vLista As Collection) As Boolean
Dim RN As ADODB.Recordset
Dim S As String


    
    On Error GoTo EEnAlbaranCompras
    EnAlbaranCompras2 = False
    Set RN = New ADODB.Recordset
    S = "select slialp.numalbar clave,cantidad cantmostrar,scaalp.fechaalb mifecha ,scaalp.codprove idRefD,scaalp.nomprove vrefd"
    S = S & " from slialp,scaalp where scaalp.codprove=slialp.codprove and scaalp.fechaalb=slialp.fechaalb and scaalp.numalbar=slialp.numalbar"
    'El arti/lote/almacen                                               'ojo a la S
    S = S & " AND slialp.codartic = " & DBSet(Me.codArtic, "T") & "AND numlotes = " & DBSet(Me.NUmlote, "T")
    S = S & " and codalmac = " & mCodalmac



    RN.Open S, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RN.EOF
        S = RN!NumAlbar
        vLista.Add S

        RN.MoveNext
    Wend
    RN.Close
    
    'En factras
    S = "select scafpc.numfactu clave,cantidad cantmostrar,scafpc.fecfactu mifecha,scafpc.codprove idrefd,scafpc.nomprove vrefd"
    S = S & " FROM slifpc,scafpc where scafpc.codprove=slifpc.codprove and scafpc.fecfactu=slifpc.fecfactu"
    S = S & " and scafpc.numfactu=slifpc.numfactu "
    'El arti/lote/almacen                                                     'ojo a la S
    S = S & " AND slifpc.codartic = " & DBSet(codArtic, "T") & "AND numlotes = " & DBSet(NUmlote, "T")
    S = S & " and codalmac = " & mCodalmac
    RN.Open S, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RN.EOF
        S = RN!Clave & "|" & RN!cantmostrar & "|"
        vLista.Add S
        RN.MoveNext
    Wend
    RN.Close
    
    EnAlbaranCompras2 = True
    Exit Function
EEnAlbaranCompras:
    MuestraError Err.Number
End Function


'En collection metere lineas con el trzo de sql
' '`albventa`,fechaventa,`nomclien`,`cantidad2`
Private Function EnAlbaranVentas2(cVentas As Collection, vArt As String, NLote As String) As Boolean
Dim RN As ADODB.Recordset
Dim S As String

    '`albventa`,fechaventa,`nomclien`,`cantidad2`

    On Error GoTo EEnAlbaranVentas
    EnAlbaranVentas2 = False
    Set RN = New ADODB.Recordset
    S = "Select * from slialb,slialblotes,scaalb    Where  slialb.CodTipom = scaalb.CodTipom"
    S = S & " AND slialb.numalbar = scaalb.numalbar AND slialb.codTipoM = slialblotes.codTipoM"
    S = S & " AND codartic = '" & vArt & "' and codalmac = " & Me.codAlmac
    S = S & " AND numlote = " & DBSet(NLote, "T") & " ORDER BY fechaalb"


    RN.Open S, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RN.EOF
        'De momento NOMBRE NO LO TRATP
        S = "'" & RN!codTipoM & Format(RN!NumAlbar, "00000") & "'," & DBSet(RN!FechaAlb, "F")
        S = S & ",NULL," & TransformaComasPuntos(CStr(RN!Cantidad))
        cVentas.Add S

        RN.MoveNext
    Wend
    RN.Close
    
    'En factras
    S = "Select * from slifac,slifaclotes    Where   slifac.codTipoM = slifaclotes.codTipoM"
    S = S & " and slifac.numfactu  = slifaclotes.numfactu and slifac.fecfactu  = slifaclotes.fecfactu"
    S = S & " and slifac.codtipoa = slifaclotes.codtipoa and slifac.numalbar = slifaclotes.numalbar"
    S = S & " and slifac.numlinea = slifaclotes.numlinea  "
    S = S & " AND codartic = '" & vArt & "' and codalmac = " & Me.codAlmac
    S = S & " AND numlote = " & DBSet(NLote, "T") & " ORDER BY slifac.fecfactu,slifac.numfactu"
    

    RN.Open S, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RN.EOF
        S = "'" & RN!codTipoM & Format(RN!NumFactu, "00000") & "'," & DBSet(RN!FecFactu, "F")
        S = S & ",NULL," & TransformaComasPuntos(CStr(RN!Cantidad))
        cVentas.Add S

        RN.MoveNext
    Wend
    RN.Close
    
    
    Exit Function
EEnAlbaranVentas:
    MuestraError Err.Number
End Function








Public Sub GeneracionProcesosDesdeCompra3(ByRef Contador As Long, ByVal vArticulo As String, ByVal vLote As String, ByRef LB As Label)
Dim Col As Collection
Dim RN As ADODB.Recordset
Dim Aux As String
Dim J As Integer
Dim FactorMultiplicador1 As Currency
Dim Nivel As Integer

         
    'En COUPAGES
    Set Col = New Collection
    Set RN = New ADODB.Recordset
    
    
    
    Aux = "select olicoupage.codartic,olicoupage.numlote, olicoupagelinlotes.cantlote,olicoupage.codigo from"
    Aux = Aux & " olicoupage , olicoupagelinlotes where olicoupage.codigo = olicoupagelinlotes.codigo"
    Aux = Aux & " AND olicoupagelinlotes.codartic = " & DBSet(vArticulo, "T")
    Aux = Aux & " AND olicoupagelinlotes.numlote = " & DBSet(vLote, "T")
    RN.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RN.EOF
        Aux = RN!codArtic & "|" & RN!NUmlote & "|" & RN!cantlote & "|" & RN!codigo & "|"
        Col.Add Aux
        RN.MoveNext
    Wend
    RN.Close
    If Col.Count > 0 Then
        

        For J = 1 To Col.Count
            Aux = "Coupage (" & RecuperaValor(Col.Item(J), 4) & ")"
            InsertarLinea Aux, Nivel, CStr(RecuperaValor(Col.Item(J), 1)), CStr(RecuperaValor(Col.Item(J), 2)), Contador, CCur(CStr(RecuperaValor(Col.Item(J), 3))), ""
            GeneracionProcesosDesdeCompra Nivel + 1, Contador, CStr(RecuperaValor(Col.Item(J), 1)), CStr(RecuperaValor(Col.Item(J), 2)), LB, True
        Next J
    End If
    
    'AHORA produccion
    Aux = "select sliordprlotes.codartic,sliordprlotes.numlote,sliordprlotes.cantlote,sliordprlotes.codigo from sliordprlotes ,sliordpr2lotes"
    Aux = Aux & " Where sliordprlotes.codigo = sliordpr2lotes.codigo"
    Aux = Aux & " and sliordprlotes.codalmac =sliordpr2lotes.codalmac"
    Aux = Aux & " and sliordprlotes.codartic =sliordpr2lotes.codartic"
    Aux = Aux & " AND sliordpr2lotes.codarti2 = " & DBSet(vArticulo, "T")
    Aux = Aux & " AND sliordpr2lotes.numlote = " & DBSet(vLote, "T")
    RN.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set Col = New Collection
    While Not RN.EOF
        Aux = RN!codArtic & "|" & RN!NUmlote & "|" & RN!cantlote & "|" & RN!codigo & "|"
        Col.Add Aux
        RN.MoveNext
    Wend
    RN.Close
    
    If Col.Count > 0 Then
        
        For J = 1 To Col.Count
            Aux = "Produc. (" & RecuperaValor(Col.Item(J), 4) & ")"
            InsertarLinea Aux, Nivel, CStr(RecuperaValor(Col.Item(J), 1)), CStr(RecuperaValor(Col.Item(J), 2)), Contador, CCur(CStr(RecuperaValor(Col.Item(J), 3))), ""
            GeneracionProcesosDesdeCompra Nivel + 1, Contador, CStr(RecuperaValor(Col.Item(J), 1)), CStr(RecuperaValor(Col.Item(J), 2)), LB, True
        Next J
    End If
    
    
        
    
    
    Set Col = Nothing
    Set RN = Nothing
    
    'Veremos para este articulo LOTE las ventas
    Stop
    EnAlbaranVentas_ Nivel, vArticulo, vLote, Contador, FactorMultiplicador1
    
    
End Sub














'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'
' Desde venta
'
'
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------

Public Function TrazbilidadDesdeVenta(SoloAceite As Boolean, MostrarMensaje As Boolean) As Boolean
Dim vContador As Long
Dim MiNivel As Integer
Dim Embasado As Boolean
Dim Produccion As Long
Dim RCompo As ADODB.Recordset
Dim Cad As String
Dim Art1 As String


Dim Encontrado As Boolean
Dim BUscarEnNuevaProduccion As Boolean
Dim LinProd As Integer



    'Insertamo el lote y el articulo
    MiNivel = 0
    vContador = 0
    'Vamos por caso
    'Si es materia embasada:
    Cad = DevuelveDesdeBD(conAri, "conjunto", "sartic", "codartic", Me.codArtic, "T")
    If Cad = "" Then Cad = "0"
    Embasado = Cad = "1"
    
    If Embasado Then
    
        'Proceso de envasado
        ' vemos la produccion donde esta
        SQL = "codalmac"
        Cad = "codartic ='" & codArtic & "' and numlote"
        Cad = DevuelveDesdeBD(conAri, "codigo", "sliordprlotes", Cad, Me.NUmlote, "T", SQL)
        Encontrado = False
        BUscarEnNuevaProduccion = False
        If Cad = "" Then
            'NO encontramos la produccion
            
            BUscarEnNuevaProduccion = True
        Else
            Encontrado = True
            Produccion = Val(Cad)
            codAlmac = Val(SQL)
        End If
        
        Set RCompo = New ADODB.Recordset
        If BUscarEnNuevaProduccion Then
            'Vemos si hay que buscar en nueva produccion
            Cad = "select * from prodlin,prodtrazlin,prodtrazcompo where prodtrazlin.codigo=prodtrazcompo.codigo and"
            Cad = Cad & " prodtrazlin.idlin = prodtrazcompo.idlin And prodtrazlin.Codigo = prodlin.Codigo And prodtrazlin.idlin = prodlin.idlin"
            Cad = Cad & " AND prodlin.codartic = " & DBSet(codArtic, "T") & "  AND prodtrazcompo.lotetraza = " & DBSet(NUmlote, "T")
            RCompo.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RCompo.EOF Then
                Encontrado = True
                Produccion = RCompo!codigo
                LinProd = RCompo!idlin
            Else
                MsgBox "No se ha encontrado en lineas de produccion"
            
            End If
            RCompo.Close
            
        End If
        
        
        
        
        If Encontrado Then
            If Not BUscarEnNuevaProduccion Then
                '**********************************
                '
                '   Esta en produccion
            
            
                Cad = "select * from sliordpr2lotes,sartic  where sliordpr2lotes.codarti2=sartic.codartic"
                Cad = Cad & " AND codigo=" & Produccion & " and sliordpr2lotes.codartic=" & DBSet(codArtic, "T") & " AND codalmac= " & codAlmac
                Cad = Cad & " ORDER BY FactorConversion desc"
                RCompo.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not RCompo.EOF
                    Cad = ""
                    
                    If RCompo!FactorConversion <> 1 Then
        
                        'Vemos si viene de un cupaje, o viene de compra directa
                        'Veremos si para un aceite y lote este, ha sido desde coupaje
                        ComprobarAceite RCompo!codarti2, DBLet(RCompo!NUmlote), True, MiNivel, vContador
                      
                        
                    Else
                        If Not SoloAceite Then
                            'Intentamos ver en que ALBARAN esta comprado
                            materiaPrimaProduccion RCompo!codarti2, DBLet(RCompo!NUmlote), MiNivel, vContador
                        
                        End If
                    End If
                    RCompo.MoveNext
                
                Wend
                RCompo.Close
                
        
            Else
                Cad = "select prodtrazcompo.*,FactorConversion from prodtrazcompo,sartic where "
                Cad = Cad & " sartic.codartic=prodtrazcompo.codartic and codigo=" & Produccion
                Cad = Cad & " and idlin=" & LinProd & " and lotetraza=" & DBSet(NUmlote, "T")
                Cad = Cad & " ORDER BY FactorConversion desc"
                RCompo.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not RCompo.EOF
                    Cad = ""
                    
                    If RCompo!FactorConversion <> 1 Then
        
                        'Vemos si viene de un cupaje, o viene de compra directa
                        'Veremos si para un aceite y lote este, ha sido desde coupaje
                        ComprobarAceite RCompo!codArtic, DBLet(RCompo!NUmlote), True, MiNivel, vContador
                      
                        
                    Else
                        If Not SoloAceite Then
                            'Intentamos ver en que ALBARAN esta comprado
                            materiaPrimaProduccion RCompo!codArtic, DBLet(RCompo!NUmlote), MiNivel, vContador
                        
                        End If
                    End If
                    RCompo.MoveNext
                
                Wend
                RCompo.Close
            
            
            End If
            Set RCompo = Nothing
        Else
            'NO HA SIDO ENCONTRADP
        End If
    
    Else
        'No es embasado
        If MostrarMensaje Then MsgBox "No embasado", vbExclamation
    End If
    Set RCompo = Nothing
End Function

Private Sub materiaPrimaProduccion(ByRef Art As String, ByRef NLote As String, Nivel As Integer, ByRef Contador As Long)
Dim T As ADODB.Recordset



    'En facturas. Si esta no busco mas
    Set T = New ADODB.Recordset
    SQL = "Select *  FROM slifpc,scafpc where scafpc.codprove=slifpc.codprove and "
    SQL = SQL & " scafpc.fecfactu=slifpc.fecfactu and scafpc.numfactu=slifpc.numfactu "
    SQL = SQL & " AND slifpc.codartic = " & DBSet(Art, "T") & " AND numlotes = " & DBSet(NLote, "T")
    SQL = SQL & " and codalmac = " & Me.codAlmac
    
    T.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    If Not T.EOF Then
        SQL = T!NumFactu & " " & T!FecFactu & " (" & T!codProve & ")"
        InsertarLinea SQL, Nivel, Art, NLote, Contador, 0, T!nomprove
        
    End If
    T.Close
    
    
    'Si no encotramos probamos en albaranes
    If SQL = "" Then
            'NO ha encontrado nada
            
            SQL = "select scaalp.*"
            SQL = SQL & " from slialp,scaalp where scaalp.codprove=slialp.codprove and scaalp.fechaalb=slialp.fechaalb and scaalp.numalbar=slialp.numalbar"
            'El arti/lote/almacen                                               'ojo a la S
            SQL = SQL & " AND slialp.codartic = " & DBSet(Art, "T") & "AND numlotes = " & DBSet(NLote, "T")
            SQL = SQL & " and codalmac = " & codAlmac
            T.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            SQL = ""
            If Not T.EOF Then
                SQL = T!NumAlbar & " " & T!FechaAlb & " (" & T!codProve & ")*"
                InsertarLinea SQL, Nivel, Art, NLote, Contador, 0, T!nomprove
                T.MoveNext
            End If
            T.Close
            
            
    End If
    Set T = Nothing
    
    
    
    
    
    
End Sub














Private Sub ComprobarAceite(ByRef Art As String, ByRef NLote As String, ComprobarCoupaje As Boolean, Nivel As Integer, ByRef Contador As Long)
Dim T As ADODB.Recordset
Dim Col As Collection
Dim i As Integer
Dim Seguir As Boolean
Dim C As String
    'Veo si esta en coupaje
    Seguir = True
    If ComprobarCoupaje Then
            SQL = "select olicoupage.codigo from"
            SQL = SQL & " olicoupage,olicoupagelin where olicoupage.codigo=olicoupagelin.codigo"
            'El arti/lote/almacen
            SQL = SQL & " AND olicoupage.codartic = " & DBSet(Art, "T") & " AND olicoupage.numlote = " & DBSet(NLote, "T")
            SQL = SQL & " and codalmac = " & mCodalmac
            SQL = SQL & " group by 1"
            Set T = New ADODB.Recordset
            T.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            SQL = ""
            If Not T.EOF Then
                SQL = T!codigo
            End If
            T.Close
          
            If SQL <> "" Then

                SQL = "select * from olicoupagelinlotes where codigo=" & SQL
                T.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                Set Col = New Collection
                While Not T.EOF
                    C = T!codArtic & "|" & T!NUmlote & "|"
                    Col.Add C
                    SQL = T!codigo
                    T.MoveNext
                Wend
                
                'Comprobamos si viene de coupar o no
                If Col.Count = 0 Then
                    MsgBox "Error leyendo coupage", vbExclamation
                Else
                    SQL = "CUP " & SQL
                    InsertarLinea SQL, Nivel, Art, NLote, Contador, 0, ""
                    Nivel = Nivel + 1
                    For i = 1 To Col.Count
                        C = Col.Item(i)
                        ComprobarAceite RecuperaValor(C, 1), RecuperaValor(C, 2), True, Nivel, Contador
                    Next i
                End If
                Set Col = Nothing
                Seguir = False
            Else
                
                'Continuamos abajo
            End If
            Set T = Nothing
    End If
    
    'Si no viene de coupar
    If Seguir Then

            'El arti/lote/almacen
            materiaPrimaProduccion Art, NLote, Nivel, Contador

            
            
    End If
    
    
    
    
    
    
    
    
End Sub









'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'
' Impresion etiquetas
'
'
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
Public Sub EstablecerEtiquetas(CuantasEtiquetas As Integer)
Dim i As Integer

    SQL = "DELETE FROM spartidaslin WHERE id = " & mIdPartida
    Conn.Execute SQL
    If CuantasEtiquetas = 0 Then Exit Sub  'Si no hay salgo
    
    SQL = ""
    For i = 1 To CuantasEtiquetas
        SQL = SQL & ", (" & mIdPartida & "," & i & ",0)"
    Next
    SQL = Mid(SQL, 2) 'QUITO la primera coma
    SQL = "INSERT INTO spartidaslin(id,bulto,cantidad) VALUES " & SQL
    EjecutaSQL conAri, SQL, True
End Sub







Public Function CuantasEtiquetas() As Integer
Dim i As Integer

    Set Rs = New ADODB.Recordset
    SQL = "select count(*) from spartidaslin WHERE id = " & mIdPartida
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CuantasEtiquetas = 0
    If Not Rs.EOF Then CuantasEtiquetas = DBLet(Rs.Fields(0), "N")
    Rs.Close
    Set Rs = Nothing
'    EjecutaSQL conAri, SQL, True
End Function


Public Function TieneEtiquetasYAenProduccion() As Boolean
Dim i As Integer

    Set Rs = New ADODB.Recordset
    SQL = "select count(*) from spartidaslin WHERE id = " & mIdPartida & " AND fechaulizada > '1900-01-01'"
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    TieneEtiquetasYAenProduccion = False
    If Not Rs.EOF Then TieneEtiquetasYAenProduccion = DBLet(Rs.Fields(0), "N") > 0
    Rs.Close
    Set Rs = Nothing
 '   EjecutaSQL conAri, SQL, True
End Function





Public Function AjustarFinPartida()
    SQL = "UPDATE spartidas set cantotal=0 "
    SQL = SQL & " WHERE id = " & mIdPartida
    Conn.Execute SQL
End Function
