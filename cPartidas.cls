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
Private midPartida As Long
Private mCodartic As String
Private mCodalmac As Integer
Private mNumlote As String
Private mCodProve As Long
Private mNumalbar As String
Private mFecha As Date
Private mCantidad As Currency

Private mEtiq As Integer


Private mPesoMedio As Currency
Private mDesviacion As Currency


Dim RS As ADODB.Recordset
Dim SQL As String




'**** Identificador de la partida
Public Property Let idPartida(ByVal vData As Long)
     midPartida = vData
End Property

Public Property Get idPartida() As Long
     idPartida = midPartida
End Property



'**** Articulo de la partida
Public Property Let codartic(ByVal vData As String)
     mCodartic = vData
End Property

Public Property Get codartic() As String
     codartic = mCodartic
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



Public Property Get PesoMedio() As Currency
     PesoMedio = mPesoMedio
End Property

Public Property Get Desviacion() As Currency
     Desviacion = mDesviacion
End Property



'***************************************************************************
'       METODOS
'***************************************************************************


'***************************************************************************
'LEER
Public Function Leer(ID As Long) As Boolean
    On Error GoTo ELee
    Leer = False
    Set RS = New ADODB.Recordset
    SQL = "Select * from spartidas where id =" & ID
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        'MAAAAl. No existe
        
    Else
        AsignarValoresClase   'para poder llamar las sentecias en mas sitios
        Leer = True
    End If
    RS.Close
        
    
ELee:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbExclamation
        Err.Clear
    End If
    Set RS = Nothing
End Function

'Leer la clase con los valores del articulo
Public Function LeerDesdeArticulo(vCodArtic As String, vCodAlmac As Integer, cLote As String) As Boolean
    On Error GoTo ELeerDesdeArticulo
    LeerDesdeArticulo = False
    Set RS = New ADODB.Recordset
    SQL = "Select * from spartidas where codartic = " & DBSet(vCodArtic, "T")
    SQL = SQL & " AND codalmac = " & CStr(vCodAlmac) & " AND numlote = " & DBSet(cLote, "T")
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        'Con los datos NO encuetra nada.
        
    Else
        AsignarValoresClase   'para poder llamar las sentecias en mas sitios
        LeerDesdeArticulo = True
    End If
    RS.Close
    
ELeerDesdeArticulo:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set RS = Nothing
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
    Set RS = New ADODB.Recordset
    SQL = "Select * from spartidas where codartic = " & DBSet(vCodArtic, "T")
    SQL = SQL & " AND codalmac = " & CStr(vCodAlmac) & " AND numlote = " & DBSet(LoteExp, "T")
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    OK = False
    If RS.EOF Then
        'Con los datos NO encuetra nada.
        
    Else
        AsignarValoresClase   'para poder llamar las sentecias en mas sitios
        LeerDesdeExpedicion = True
        OK = True
    End If
    RS.Close
    
    
    
    'Si no se ha encontrado NADA, miramos a ver si el parte de produccion se puso de
    'manera manual. Al estilo 4001 2012/05/24
    If Not OK Then
            SQL = "'" & Val(LoteExp) & " 2012/%'"
            SQL = " AND codalmac = " & CStr(vCodAlmac) & " AND numlote like " & SQL
            SQL = "Select * from spartidas where codartic = " & DBSet(vCodArtic, "T") & SQL
            
            RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            OK = False
            If RS.EOF Then
                'Con los datos NO encuetra nada.
                
            Else
                AsignarValoresClase   'para poder llamar las sentecias en mas sitios
                LeerDesdeExpedicion = True
                OK = True
            End If
            RS.Close
    End If
    
ELeerDesdeArticulo:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set RS = Nothing
End Function




Public Function LeerDesdeAlbaran(vCodArtic As String, vCodAlmac As Integer, cLote As String, NumAlbar As String, codProve As Long) As Boolean
    On Error GoTo ELeerDesdeArticulo
    LeerDesdeAlbaran = False
    Set RS = New ADODB.Recordset
    SQL = "Select * from spartidas where codartic = " & DBSet(vCodArtic, "T")
    SQL = SQL & " AND codalmac = " & CStr(vCodAlmac) & " AND numlote = " & DBSet(cLote, "T")
    SQL = SQL & " AND numalbar = " & DBSet(NumAlbar, "T") & " AND codprove = " & CStr(codProve)
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        'Con los datos NO encuetra nada.
        
    Else
        AsignarValoresClase   'para poder llamar las sentecias en mas sitios
        LeerDesdeAlbaran = True
    End If
    RS.Close
    
ELeerDesdeArticulo:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set RS = Nothing
End Function





Public Sub IncrementarCantidad(Incremento As Currency)
    On Error GoTo EIns
    
    SQL = "UPDATE spartidas set cantotal=cantotal "
    If Incremento >= 0 Then SQL = SQL & "+ "
    SQL = SQL & TransformaComasPuntos(CStr(Incremento))
    SQL = SQL & " WHERE id = " & midPartida
    conn.Execute SQL
    Exit Sub
EIns:
    MsgBox Err.Description, vbExclamation
    
End Sub



'CUBAS-DEPOSITOS
'Cuando es fin deposito, se regulzarizara, y la cantidad hay que dejarla a cero
Public Sub FinPartida()
    On Error GoTo EIns
    
    SQL = "UPDATE spartidas set cantotal=0 "
    SQL = SQL & " WHERE id = " & midPartida
    conn.Execute SQL
    Exit Sub
EIns:
    MsgBox Err.Description, vbExclamation
    
End Sub


Public Function Insertar() As Boolean
    On Error GoTo EIns
    Insertar = False
    
    'Siguiente
    midPartida = Siguiente
    'Insetamos
    SQL = "insert into spartidas (id,codartic,codalmac,numlote,codprove,numalbar,fecha,cantotal) values ("
    SQL = SQL & midPartida & "," & DBSet(mCodartic, "T") & "," & mCodalmac & ","
    SQL = SQL & DBSet(mNumlote, "T") & "," & DBSet(mCodProve, "N")
    SQL = SQL & "," & DBSet(mNumalbar, "T") & "," & DBSet(mFecha, "F") & "," & DBSet(mCantidad, "N") & ")"
    conn.Execute SQL
    Insertar = True 'todo ok
    Exit Function
EIns:
    MsgBox Err.Description, vbExclamation
    
End Function



Public Function Eliminar() As Boolean
    On Error GoTo EIns
    Eliminar = False
    
    SQL = "DELETE FROM spartidaslin WHERE id = " & midPartida
    conn.Execute SQL
    
    SQL = "DELETE FROM spartidaspesos WHERE idpartida = " & midPartida
    conn.Execute SQL
    
    SQL = "DELETE FROM spartidas WHERE id = " & midPartida
    conn.Execute SQL
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
    Set RS = New ADODB.Recordset
    SQL = "Select * from spartidas where codartic = " & DBSet(vCodArtic, "T")
    SQL = SQL & " AND codalmac = " & CStr(vCodAlmac)
    SQL = SQL & " ORDER BY fecha asc"
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        'Con los datos NO encuetra nada.
        AsignarValoresClase   'para poder llamar las sentecias en mas sitios
        AlgunoAsignado = False
        'Ahora veremos, a partir de la cantidad QUE necesitamos, si podemos darle los numeros de serie
        Set kLotes = New Collection
        Fin = False
        While Not Fin
            Ca = RS!cantotal
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
                SQL = RS!ID & "|" & Ca & "|" & RS!NUmlote & "|"
                kLotes.Add SQL
            End If
            
            RS.MoveNext
            If RS.EOF Then Fin = True
            
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
    RS.Close
    
ERecuperarLotes:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set RS = Nothing
End Function



'Para la VALL.
'El aceite , por defecto, viene del deposito 18
'Con lo cual el lote debe traer ese
Public Function RecuperarLotesprodVall(vCodArtic As String, vCodAlmac As Integer, ByVal CantidadNecesaria As Currency, ByRef kLotes As Collection) As Byte
'   CantidadNecesaria. La pasamos byVal para poder modificarla aqui dentro
Dim Ca As Currency
Dim Fin As Boolean
Dim AlgunoAsignado As Boolean
Dim LoteDeposito As String

    On Error GoTo ERecuperarLotes
    RecuperarLotesprodVall = 2
    Set RS = New ADODB.Recordset
    
    SQL = "Select numlote from proddepositos where numdeposito=18"
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    LoteDeposito = ""
    If Not RS.EOF Then LoteDeposito = DBLet(RS!NUmlote, "T")
    RS.Close
    
    
    If LoteDeposito = "" Then
        SQL = "1 "
    Else
        SQL = "if(numlote=" & DBSet(LoteDeposito, "T") & ",0,1)"
    End If
    SQL = "Select spartidas.*," & SQL & " ElLote"
    SQL = SQL & " from spartidas where codartic = " & DBSet(vCodArtic, "T")
    SQL = SQL & " AND codalmac = " & CStr(vCodAlmac)
    SQL = SQL & " ORDER BY ElLote, fecha asc"
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        'Con los datos NO encuetra nada.
        AsignarValoresClase   'para poder llamar las sentecias en mas sitios
        AlgunoAsignado = False
        'Ahora veremos, a partir de la cantidad QUE necesitamos, si podemos darle los numeros de serie
        Set kLotes = New Collection
        Fin = False
        While Not Fin
            Ca = RS!cantotal
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
                SQL = RS!ID & "|" & Ca & "|" & RS!NUmlote & "|"
                kLotes.Add SQL
            End If
            
            RS.MoveNext
            If RS.EOF Then Fin = True
            
        Wend
        
        
        
        If CantidadNecesaria > 0 Then
            'No tenemos bastante del articulo para sacar
         '   MsgBox "No hay suficiente cantidad de lotes para el articulo. ", vbExclamation
            If AlgunoAsignado Then RecuperarLotesprodVall = 1
        Else
            'Todo ok
            RecuperarLotesprodVall = 0
        End If
        
    End If
    RS.Close
    
ERecuperarLotes:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set RS = Nothing
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
    
    Set RS = New ADODB.Recordset

    
    For J = 1 To 14
        'olicoupage,olicoupagelinlotes , slialblotes, slifaclotes, sliordpr2lotes, sliordprlotes, slialp,
        'slifpc,slhmov,prodtrazalin,prodtrazcompo,smovalotes, smovallotes para MLTs
        '
        LoteEnTabla2 J
    Next
    Set RS = Nothing
    CargaDatosParaListar = True
    Exit Function
ECargaDatosParaListar:
    Set RS = Nothing
    MuestraError Err.Number, Err.Description
End Function




'Comprobara la suma de cantidades con la cantidad almacenada
Public Sub AjustarCantidad(Preguntar As Boolean)
Dim Real As Currency

    'Borro tmp
    SQL = "DELETE from tmppartidas where codusu = " & vUsu.Codigo
    conn.Execute SQL
    If CargaDatosParaListar Then
        'Ya tengo cargado el
        
        
        Set RS = New ADODB.Recordset
        SQL = "Select sum(cantidad) from tmppartidas where codusu = " & vUsu.Codigo & " AND idpartida=" & Me.idPartida
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Real = 0
        If Not RS.EOF Then Real = DBLet(RS.Fields(0), "N")
        RS.Close
        Set RS = Nothing
        
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
                    SQL = SQL & " WHERE id = " & midPartida
                    EjecutaSQL conAri, SQL, True
            End If
        End If
    End If
    
End Sub








'#################################################################################################
'Privados

Private Sub AsignarValoresClase()
    'No lleva error. Que se dispare el erro en el sub que lo llama
    midPartida = RS!ID
    mCodartic = RS!codartic
    mCodalmac = RS!codAlmac
    mNumlote = RS!NUmlote
    mCodProve = RS!codProve
    mNumalbar = RS!NumAlbar
    mFecha = RS!Fecha
    mCantidad = RS!cantotal
    mPesoMedio = 0   'Estos se pondran cuando se solicite el evento leerpesos
    mDesviacion = 0
End Sub


'Devolvera el que le tiocaria
Public Function Siguiente() As Long
    SQL = "select max(id) from spartidas"
    Set RS = New ADODB.Recordset
    Siguiente = 1
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then Siguiente = RS.Fields(0) + 1
    End If
    RS.Close
    Set RS = Nothing
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
Dim miSQL As String
Dim Ca As Currency
Dim SumaLotes As Boolean
    
    Select Case Num
    Case 1
        'Lineas de coupage
        Desc = "'COUPAGE'"
        miSQL = "select olicoupage.codigo clave, sum(kilos) cantmostrar ,olicoupage.fecha mifecha, 0 idRefd,' ' vRefd from"
        miSQL = miSQL & " olicoupage,olicoupagelin where olicoupage.codigo=olicoupagelin.codigo"
        'El arti/lote/almacen
        miSQL = miSQL & " AND olicoupage.codartic = " & DBSet(mCodartic, "T") & " AND olicoupage.numlote = " & DBSet(mNumlote, "T")
        miSQL = miSQL & " and codalmac = " & mCodalmac
        miSQL = miSQL & " group by 1,3"
        SumaLotes = True
    Case 2
    'Lineas de coupage
        Desc = "'COUPAGE Lin'"
        miSQL = "select olicoupagelinlotes.codigo clave,cantlote cantmostrar,fecha mifecha ,0 idRefd,' ' vRefd"
        miSQL = miSQL & " FROM olicoupagelinlotes,olicoupage WHERE  olicoupagelinlotes.codigo=olicoupage.codigo"
        'El arti/lote/almacen
        miSQL = miSQL & " AND olicoupagelinlotes.codartic = " & DBSet(mCodartic, "T") & " AND olicoupagelinlotes.numlote = " & DBSet(mNumlote, "T")
        miSQL = miSQL & " and codalmac = " & mCodalmac
        SumaLotes = False
    
    Case 3
        'Lineas albaran
        Desc = "'Albaran cliente'"
        miSQL = "select concat(slialblotes.codtipom,slialblotes.numalbar) clave,slialblotes.cantidad cantmostrar,"
        miSQL = miSQL & " fechaalb mifecha, codclien idrefd,nomclien vrefd "
        miSQL = miSQL & " from scaalb,slialblotes,slialb where"
        miSQL = miSQL & " scaalb.codtipom=slialb.codtipom and scaalb.numalbar=slialblotes.numalbar and"
        miSQL = miSQL & " slialblotes.codTipoM = slialb.codTipoM And slialb.NumAlbar = slialblotes.NumAlbar"
        miSQL = miSQL & " and slialb.numlinea=slialblotes.numlinea"
        'El arti/lote/almacen
        miSQL = miSQL & " AND slialb.codartic = " & DBSet(mCodartic, "T") & " AND numlote = " & DBSet(mNumlote, "T")
        miSQL = miSQL & " and codalmac = " & mCodalmac
        SumaLotes = False
        
    Case 4
        'lineas facturas
        
        
         
        
        
        Desc = "'FACTURAS CLIENTE'"
        miSQL = "select concat(slifaclotes.codtipom,slifaclotes.numfactu) clave,slifaclotes.cantidad cantmostrar,"
        miSQL = miSQL & " scafac.fecfactu mifecha, codclien idrefd,nomclien vrefd  from"
       ' miSQL = miSQL & " scafac,slifac,slifaclotes where scafac.codtipom=slifac.codtipom"
       ' miSQL = miSQL & " and scafac.numfactu=slifac.numalbar and scafac.fecfactu=slifac.numalbar"
       ' miSQL = miSQL & " and scafac.numfactu=slifaclotes.numfactu and scafac.fecfactu=slifaclotes.fecfactu and"
       ' miSQL = miSQL & " scafac.codTipoM = slifaclotes.codTipoM and slifac.numlinea=slifaclotes.numlinea "
        'AHORA MAYO 2010
        miSQL = miSQL & " (`scafac` `scafac` INNER JOIN `slifac` `slifac` ON ((`scafac`.`codtipom`=`slifac`.`codtipom`) "
        miSQL = miSQL & " AND (`scafac`.`numfactu`=`slifac`.`numfactu`)) AND (`scafac`.`fecfactu`=`slifac`.`fecfactu`))"
        miSQL = miSQL & " INNER JOIN `slifaclotes` `slifaclotes` ON (((((`slifac`.`codtipom`=`slifaclotes`.`codtipom`)"
        miSQL = miSQL & " AND (`slifac`.`numfactu`=`slifaclotes`.`numfactu`)) AND (`slifac`.`fecfactu`=`slifaclotes`.`fecfactu`))"
        miSQL = miSQL & " AND (`slifac`.`codtipoa`=`slifaclotes`.`codtipoa`)) AND (`slifac`.`numalbar`=`slifaclotes`.`numalbar`))"
        miSQL = miSQL & " AND (`slifac`.`numlinea`=`slifaclotes`.`numlinea`)"

       
        'El arti/lote/almacen
        miSQL = miSQL & " AND slifac.codartic = " & DBSet(mCodartic, "T") & " AND numlote = " & DBSet(mNumlote, "T")
        miSQL = miSQL & " and codalmac = " & mCodalmac
        SumaLotes = False
        
    Case 5
        'PRoduccion
        'lineas SUMA
        Desc = "'PRODUCCION'"
        miSQL = "select sordprod.codigo clave,cantlote cantmostrar,fecproduccion mifecha ,0 idRefd,' ' vRefd"
        miSQL = miSQL & " FROM sordprod,sliordprlotes WHERE sordprod.codigo = sliordprlotes.codigo"
        'Para que saque solo los que estan producidos
        miSQL = miSQL & " AND fecproduccion>'1900-01-01'"
        
        'El arti/lote/almacen
        miSQL = miSQL & " AND sliordprlotes.codartic = " & DBSet(mCodartic, "T") & " AND numlote = " & DBSet(mNumlote, "T")
        miSQL = miSQL & " and codalmac = " & mCodalmac
        SumaLotes = True
    Case 6
        'Sub lineas restan
        Desc = "'PRODUCCION compo.'"
        miSQL = "select sordprod.codigo clave,cantlote cantmostrar"
        miSQL = miSQL & ",fecproduccion mifecha ,0 idRefd,' ' vRefd"
        miSQL = miSQL & " FROM sordprod,sliordpr2lotes,sartic WHERE sordprod.codigo = sliordpr2lotes.codigo"
        miSQL = miSQL & " AND sliordpr2lotes.codarti2=sartic.codartic"
        'Para que saque solo los que estan producidos
        miSQL = miSQL & " AND fecproduccion>'1900-01-01'"
        
        'El arti/lote/almacen
        miSQL = miSQL & " AND sliordpr2lotes.codarti2 = " & DBSet(mCodartic, "T") & " AND numlote = " & DBSet(mNumlote, "T")
        miSQL = miSQL & " and codalmac = " & mCodalmac
        SumaLotes = False
        
        
    Case 7
        'ALbaranes proveedor
        Desc = "'Albaranes proveedor'"
        miSQL = "select slialp.numalbar clave,cantidad cantmostrar,scaalp.fechaalb mifecha ,scaalp.codprove idRefD,scaalp.nomprove vrefd"
        miSQL = miSQL & " from slialp,scaalp where scaalp.codprove=slialp.codprove and scaalp.fechaalb=slialp.fechaalb and scaalp.numalbar=slialp.numalbar"
        'El arti/lote/almacen                                               'ojo a la S
        miSQL = miSQL & " AND slialp.codartic = " & DBSet(mCodartic, "T") & " AND numlotes = " & DBSet(mNumlote, "T")
        miSQL = miSQL & " and codalmac = " & mCodalmac
        SumaLotes = True
        
    Case 8
        'Facturas proveedor
         Desc = "'FACTURAS PROVEEDOR'"
        miSQL = "select scafpc.numfactu clave,cantidad cantmostrar,scafpc.fecfactu mifecha,scafpc.codprove idrefd,scafpc.nomprove vrefd"
        miSQL = miSQL & " FROM slifpc,scafpc where scafpc.codprove=slifpc.codprove and scafpc.fecfactu=slifpc.fecfactu"
        miSQL = miSQL & " and scafpc.numfactu=slifpc.numfactu "
        'El arti/lote/almacen                                                     'ojo a la S de numloteSSSS
        miSQL = miSQL & " AND slifpc.codartic = " & DBSet(mCodartic, "T") & " AND numlotes = " & DBSet(mNumlote, "T")
        miSQL = miSQL & " and codalmac = " & mCodalmac
        SumaLotes = True
        
        
        
    Case 9
        'REGULARIZACION
         Desc = "'REGULARIZACION'"
        miSQL = "select schmov.codmovim clave,if(tipomovi=0,-cantidad,cantidad) cantmostrar,schmov.fecmovim mifecha,schmov.codtraba idrefd,nomtraba vrefd"
        miSQL = miSQL & " from slhmov,schmov left join straba on schmov.codtraba=straba.codtraba WHERE slhmov.codmovim=schmov.codmovim"
 
        'El arti/lote/almacen
        miSQL = miSQL & " AND slhmov.codartic = " & DBSet(mCodartic, "T") & " AND numlote = " & DBSet(mNumlote, "T")
        miSQL = miSQL & " and schmov.codalmac = " & mCodalmac
        SumaLotes = True
        
        
    Case 10
        'NUEA PRODUCCION
        Desc = "'PRO.TRAZABILIDAD'"
        
        'FALTA###
        'comprobar que es numerico
        If Not IsNumeric(mNumlote) Then Exit Sub
        'prodtrazlin
        miSQL = "select concat(prodcab.codigo,""/"",prodlin.idlin) clave,prodtrazlin.cantprodu cantmostrar,fhinicio mifecha,0 idRefd,' ' vRefd"
        miSQL = miSQL & " from prodtrazlin,prodlin,prodcab where prodcab.codigo=prodlin.codigo  and prodtrazlin.codigo=prodlin.codigo and prodtrazlin.idlin=prodlin.idlin"
        'El arti/lote/almacen
        miSQL = miSQL & " AND prodlin.codartic = " & DBSet(mCodartic, "T") & " AND lotetraza = " & Val(mNumlote)
        miSQL = miSQL & " and almacen2 = " & mCodalmac
        SumaLotes = True
    Case 11
        Desc = "'PRO.TRAZA.COMPO.'"
        
        'FALTA###
        'comprobar que es numerico
        
        'prodtrazlin
        miSQL = "select concat(prodcab.codigo,""/"",prodtrazcompo.idlin,"" "") clave,prodtrazcompo.cantutili cantmostrar,"
        miSQL = miSQL & "fhinicio mifecha,prodtrazcompo.lineaprod idRefd,concat('Idtraza: ',prodtrazcompo.lotetraza) vRefd"
        miSQL = miSQL & " from prodtrazcompo,prodlin,prodcab where prodcab.codigo=prodlin.codigo  and prodtrazcompo.codigo=prodlin.codigo and prodtrazcompo.idlin=prodlin.idlin"
        'El arti/lote/almacen
        miSQL = miSQL & " AND prodtrazcompo.codartic = " & DBSet(mCodartic, "T") & " AND numlote = " & DBSet(mNumlote, "T")
        miSQL = miSQL & " and almacen2 = " & mCodalmac
        miSQL = miSQL & " and cantutili>0"
        SumaLotes = False
        
        
    Case 12
        'FIN DEPOSITO
        Desc = "'FIN DEPOSITO CUP'"
        miSQL = "select document clave,if(tipomovi=1,-cantidad,cantidad) cantmostrar,fechamov mifecha ,0 idRefd,' ' vRefd"
        miSQL = miSQL & " FROM smovalotes WHERE  "
        miSQL = miSQL & " detamovi='DEP' AND codartic = " & DBSet(mCodartic, "T") & " AND numlote = " & DBSet(mNumlote, "T")
        miSQL = miSQL & " and codalmac = " & mCodalmac
        SumaLotes = False
        
    Case 13
        Desc = "'MOLTURACION'"
      
        miSQL = "select document clave,if(tipomovi=1,cantidad,-cantidad) cantmostrar,fechamov mifecha ,codprocliope idRefd,nomtraba vRefd"
        miSQL = miSQL & " FROM smovalotes left join straba on codtraba=codprocliope WHERE  "
        miSQL = miSQL & " detamovi='MLT' AND codartic = " & DBSet(mCodartic, "T") & " AND numlote = " & DBSet(mNumlote, "T")
        miSQL = miSQL & " and smovalotes.codalmac = " & mCodalmac
        SumaLotes = True
        
    Case 14
        Desc = "'TRASIEGO'"
      
        miSQL = "select document clave,if(tipomovi=1,cantidad,-cantidad) cantmostrar,fechamov mifecha ,codprocliope idRefd,nomtraba vRefd"
        miSQL = miSQL & " FROM smovalotes left join straba on codtraba=codprocliope WHERE  "
        miSQL = miSQL & " detamovi='TRO' AND codartic = " & DBSet(mCodartic, "T") & " AND numlote = " & DBSet(mNumlote, "T")
        miSQL = miSQL & " and smovalotes.codalmac = " & mCodalmac
        SumaLotes = True
        
        
    End Select
    RS.Open miSQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    ' El insert
'   `codusu`,`idpartida`,`codartic`,`numlote`,fecha,`idOperacion`,`idNumOperacion`,
' `idReferencia`,`Referencia `,`cantidad` ,abs_cantidad    Ref: cliente/proveedor/trabajador
    SQL = ", (" & vUsu.Codigo & "," & midPartida & "," & DBSet(mCodartic, "T") & "," & DBSet(mNumlote, "T") & ","
    miSQL = ""
    While Not RS.EOF
        miSQL = miSQL & SQL 'codusus,partida,numlote
    
        'fecha, idoperacion,idnumoperacion
        miSQL = miSQL & DBSet(RS!mifecha, "F") & "," & Desc & "," & DBSet(RS!Clave, "T")
        'idreferencia, referencia
        miSQL = miSQL & "," & RS!idrefd & "," & DBSet(RS!vrefd, "T") & ","
        'cantidad, y valor absoluto de la cantidad
        Ca = RS!cantmostrar
        If Not SumaLotes Then Ca = -1 * Ca
        miSQL = miSQL & DBSet(Ca, "N") & "," & DBSet(RS!cantmostrar, "N") & ")"
        RS.MoveNext
    Wend
    RS.Close
    
    If miSQL <> "" Then
        SQL = "insert into `tmppartidas` (`codusu`,`idpartida`,`codartic`,`numlote`,`fecha`,`idOperacion`,`idNumOperacion`,`idReferencia`,`Referencia`,`cantidad`,abs_cantidad) values"
        SQL = SQL & Mid(miSQL, 2)
        conn.Execute SQL
    End If
    
End Sub




'Ya debe estar hecho el .Leer
Public Function TrazbilidadDesdeCompra(ByRef L As Label, DetallarCoupages As Boolean) As Boolean
Dim vContador As Long
Dim MiNivel As Integer
        
    MiNivel = 0
    vContador = 0
    
    GeneracionProcesosDesdeCompra MiNivel, vContador, Me.codartic, Me.NUmlote, L, DetallarCoupages
    
End Function



Public Sub GeneracionProcesosDesdeCompra(Nivel As Integer, ByRef contador As Long, ByVal vArticulo As String, ByVal vLote As String, ByRef LB2 As Label, DetallaLosCoup As Boolean)
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

    If Not LB2 Is Nothing Then
        LB2.Caption = Nivel & "-> " & vArticulo
        LB2.Refresh
    End If
    Set RN = New ADODB.Recordset
    
    'Si esta en compras, solo el primer LOTE
    'Debug.Print vArticulo & " " & vLote
    If Nivel = 0 Then
        'Las compras del articulo ORIGEN solo
        
        EnAlbaranCompras Nivel, vArticulo, vLote, contador
   
        If vParamAplic.QUE_EMPRESA = 4 Then
            'En LA VALL, hay que ver si viene de una molturacion
            Aux = "Select id,HoraFin,deposito,litros,kilos,codtrabaRec,nomtraba from vallalmazaraproceso left join straba on codtrabaRec=codtraba "
            Aux = Aux & " where loteproducido=" & DBSet(vLote, "T") & " AND artproducido =" & DBSet(vArticulo, "T")
            RN.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RN.EOF Then
                'Kilos a litros
                Aux = DevuelveDesdeBD(conAri, "factorconversion", "sartic", "codartic", vArticulo, "T")
                If Aux = "" Then Aux = "1"
                Cantidad = Round(CCur(Aux) * RN!Litros, 2)
            
                'PErfecto. Vienen de una molturacion
                Aux = "MOLT: " & RN!ID & " - " & Format(RN!HoraFin, "dd/mm/yy") & " Dep:" & RN!Deposito & "   "
                Aux = Aux & "Kilos oliva " & RN!Kilos
                
                InsertarLinea Aux, Nivel, vArticulo, vLote, contador, Cantidad, RN!codtrabaRec & " " & DBLet(RN!NomTraba, "T")
            End If
            RN.Close
            
            
            
            
            If contador = 0 Then
                'Viene de un coupage inicial
                Aux = "select olicoupage.codigo,olicoupage.fecha,olicoupage.Deposito, sum(olicoupagelinlotes.cantlote) cantidad from"
                Aux = Aux & " olicoupage , olicoupagelinlotes where olicoupage.codigo = olicoupagelinlotes.codigo"
                Aux = Aux & " AND olicoupage.codartic = " & DBSet(vArticulo, "T")
                Aux = Aux & " AND olicoupage.numlote = " & DBSet(vLote, "T")
                RN.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not RN.EOF Then
                
                    'PErfecto. Vienen de una molturacion
                    Aux = "CUP: " & RN!Codigo & " - " & Format(RN!Fecha, "dd/mm/yy") & " Dep:" & RN!Deposito & "   "
      
                    InsertarLinea Aux, Nivel, vArticulo, vLote, contador, DBLet(RN!Cantidad, "N"), ""
                End If
                RN.Close
            End If
        End If
    End If
         
    'En COUPAGES
    Set Col = New Collection
    
    
    Aux = "select olicoupage.codartic,olicoupage.numlote, olicoupagelinlotes.cantlote,olicoupage.codigo from"
    Aux = Aux & " olicoupage , olicoupagelinlotes where olicoupage.codigo = olicoupagelinlotes.codigo"
    Aux = Aux & " AND olicoupagelinlotes.codartic = " & DBSet(vArticulo, "T")
    Aux = Aux & " AND olicoupagelinlotes.numlote = " & DBSet(vLote, "T")
    RN.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RN.EOF
        
    
        Aux = RN!codartic & "|" & RN!NUmlote & "|" & RN!cantlote & "|" & RN!Codigo & "|"
        Col.Add Aux
        RN.MoveNext
    Wend
    RN.Close
    If Col.Count > 0 Then
        

        For J = 1 To Col.Count
            Aux = "Coupage (" & RecuperaValor(Col.Item(J), 4) & ")"
            InsertarLinea Aux, Nivel, CStr(RecuperaValor(Col.Item(J), 1)), CStr(RecuperaValor(Col.Item(J), 2)), contador, CCur(CStr(RecuperaValor(Col.Item(J), 3))), ""
            If DetallaLosCoup Then GeneracionProcesosDesdeCompra Nivel + 1, contador, CStr(RecuperaValor(Col.Item(J), 1)), CStr(RecuperaValor(Col.Item(J), 2)), LB2, True
        Next J
    Else
        
    End If
    
    'AHORA produccion
    Aux = "select sliordprlotes.codartic,sliordprlotes.numlote,sliordprlotes.cantlote,sliordprlotes.codigo from sliordprlotes ,sliordpr2lotes"
    Aux = Aux & " Where sliordprlotes.codigo = sliordpr2lotes.codigo"
    Aux = Aux & " and sliordprlotes.codalmac =sliordpr2lotes.codalmac"
    Aux = Aux & " and sliordprlotes.codartic =sliordpr2lotes.codartic"
    Aux = Aux & " AND sliordpr2lotes.codarti2 = " & DBSet(vArticulo, "T")
    Aux = Aux & " AND sliordpr2lotes.numlote = " & DBSet(vLote, "T")
    RN.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set Col = New Collection
    While Not RN.EOF
        Aux = RN!codartic & "|" & RN!NUmlote & "|" & RN!cantlote & "|" & RN!Codigo & "|"
        Col.Add Aux
        RN.MoveNext
    Wend
    RN.Close
    
    If Col.Count > 0 Then
        
        For J = 1 To Col.Count
            Aux = "Produc. (" & RecuperaValor(Col.Item(J), 4) & ")"
            InsertarLinea Aux, Nivel, CStr(RecuperaValor(Col.Item(J), 1)), CStr(RecuperaValor(Col.Item(J), 2)), contador, CCur(CStr(RecuperaValor(Col.Item(J), 3))), ""
            GeneracionProcesosDesdeCompra Nivel + 1, contador, CStr(RecuperaValor(Col.Item(J), 1)), CStr(RecuperaValor(Col.Item(J), 2)), LB2, True
        Next J
    End If
    Set Col = Nothing
    
    ' If Nivel = 0 Then
    Aux = "select * from prodtrazcompo where"
    Aux = Aux & " prodtrazcompo.codartic = " & DBSet(vArticulo, "T")
    Aux = Aux & " AND prodtrazcompo.numlote = " & DBSet(vLote, "T")
    'Veremos todos los lotes nuevos
    
    Set Col2 = New Collection
    Set Col = New Collection
    RN.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
        RN.Open Aux, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not RN.EOF
            Aux = RN!codartic & "|" & RN!el_lote & "|" & RN!cantprodu & "|" & RN!elcodigo & "|"
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
            
            InsertarLinea Aux, Nivel, CStr(RecuperaValor(Col.Item(J), 1)), CStr(RecuperaValor(Col.Item(J), 2)), contador, Cantidad, ""
            GeneracionProcesosDesdeCompra Nivel + 1, contador, CStr(RecuperaValor(Col.Item(J), 1)), CStr(RecuperaValor(Col.Item(J), 2)), LB2, True
        Next J
    End If
 
 

 
 
    Set Col = Nothing
    
    
    
    
    
    
    'Veremos para este articulo LOTE las ventas
    If Nivel = 0 Then
        Aux = "Select cantidad ,tipomovi,numlinea from smovalotes where codartic=" & DBSet(vArticulo, "T") & " AND numlote = " & DBSet(vLote, "T")
        Aux = Aux & " AND detamovi in ('DEP')"
        RN.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RN.EOF Then
            Aux = "Regularizacion deposito " & RN!numlinea
            Cantidad = RN!Cantidad
            InsertarLinea Aux, Nivel, vArticulo, "", contador, Cantidad, ""
                
        End If
        RN.Close
        
            
       If vParamAplic.QUE_EMPRESA > 2 Then
            Aux = "Select * from slifac,slifaclotes    Where   slifac.codTipoM = slifaclotes.codTipoM"
            Aux = Aux & " and slifac.numfactu  = slifaclotes.numfactu and slifac.fecfactu  = slifaclotes.fecfactu"
            Aux = Aux & " and slifac.codtipoa = slifaclotes.codtipoa and slifac.numalbar = slifaclotes.numalbar"
            Aux = Aux & " and slifac.numlinea = slifaclotes.numlinea  "
            Aux = Aux & " AND codartic = '" & vArticulo & "' and codalmac =1"
            Aux = Aux & " AND numlote = " & DBSet(vLote, "T")
            
            RN.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RN.EOF
                Aux = RN!Codtipom & Format(RN!NumFactu, "000000")
                Cantidad = RN!Cantidad
                InsertarLinea Aux, Nivel, vArticulo, "", contador, Cantidad, ""
                RN.MoveNext
            Wend
            RN.Close
            
            
            Aux = "Select * from slialb,slialblotes    Where   slialb.codTipoM = slialblotes.codTipoM"
            Aux = Aux & " and slialb.codtipom = slialblotes.codtipom and slialb.numalbar = slialblotes.numalbar"
            Aux = Aux & " and slialb.numlinea = slialblotes.numlinea  "
            Aux = Aux & " AND codartic = '" & vArticulo & "' and codalmac =1"
            Aux = Aux & " AND numlote = " & DBSet(vLote, "T")
            
            RN.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RN.EOF
                Aux = RN!Codtipom & Format(RN!NumAlbar, "000000")
                Cantidad = RN!Cantidad
                InsertarLinea Aux, Nivel, vArticulo, "", contador, Cantidad, ""
                RN.MoveNext
            Wend
            RN.Close
            
            
            
        End If
    


        
        
        
    End If
    
    
    'Si tiene alguna regularizacion
    Set RN = Nothing
    
    
    
    
    
End Sub

Private Function EnAlbaranVentas_(vNivel As Integer, vArt As String, NLote As String, ByRef contador As Long, FactorMultiplicador As Currency) As Boolean
Dim RN As ADODB.Recordset
Dim s As String
Dim ProveClie As String

    On Error GoTo EEnAlbaranVentas
    EnAlbaranVentas_ = False
    Set RN = New ADODB.Recordset
    s = "Select * from slialb,slialblotes    Where slialb.codTipoM = slialblotes.codTipoM"
    s = s & " AND codartic = '" & vArt & "' and codalmac = " & Me.codAlmac
    s = s & " AND numlote = " & DBSet(NLote, "T")


    RN.Open s, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RN.EOF
        ProveClie = "numalbar = " & RN!NumAlbar & " AND codtipom"
        ProveClie = DevuelveDesdeBD(conAri, "nomclien", "scaalb", ProveClie, RN!Codtipom, "T")
        InsertarLinea RN!Codtipom & RN!NumAlbar, vNivel, vArt, "", contador, RN!Cantidad, ProveClie

        RN.MoveNext
    Wend
    RN.Close
    
    'En factras
    s = "Select * from slifac,slifaclotes    Where   slifac.codTipoM = slifaclotes.codTipoM"
    s = s & " and slifac.numfactu  = slifaclotes.numfactu and slifac.fecfactu  = slifaclotes.fecfactu"
    s = s & " and slifac.codtipoa = slifaclotes.codtipoa and slifac.numalbar = slifaclotes.numalbar"
    s = s & " and slifac.numlinea = slifaclotes.numlinea  "
    s = s & " AND codartic = '" & vArt & "' and codalmac = " & Me.codAlmac
    s = s & " AND numlote = " & DBSet(NLote, "T")
    

    RN.Open s, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RN.EOF
        
        ProveClie = "codtipom = '" & RN!Codtipom & "' AND fecfactu =" & DBSet(RN!Fecfactu, "F") & " AND numfactu"
        ProveClie = DevuelveDesdeBD(conAri, "nomclien", "scafac", ProveClie, RN!NumFactu, "T")
        
        
        InsertarLinea RN!Codtipom & RN!NumFactu, vNivel, vArt, "", contador, RN!Cantidad, ProveClie

        RN.MoveNext
    Wend
    RN.Close
    
    
    Exit Function
EEnAlbaranVentas:
    MuestraError Err.Number
End Function



Private Function EnAlbaranCompras(vNivel As Integer, vArt As String, NLote As String, ByRef contador As Long) As Boolean
Dim RN As ADODB.Recordset
Dim s As String
Dim DesdeUnAlbaran As Boolean

    On Error GoTo EEnAlbaranCompras
    EnAlbaranCompras = False
    DesdeUnAlbaran = False
    Set RN = New ADODB.Recordset
    s = "select slialp.numalbar clave,cantidad cantmostrar,scaalp.fechaalb mifecha ,scaalp.codprove idRefD,scaalp.nomprove vrefd"
    s = s & " from slialp,scaalp where scaalp.codprove=slialp.codprove and scaalp.fechaalb=slialp.fechaalb and scaalp.numalbar=slialp.numalbar"
    'El arti/lote/almacen                                               'ojo a la S
    s = s & " AND slialp.codartic = " & DBSet(vArt, "T") & "AND numlotes = " & DBSet(NLote, "T")
    s = s & " and codalmac = " & mCodalmac



    RN.Open s, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RN.EOF
        DesdeUnAlbaran = True
        InsertarLinea "ALP " & RN!Clave, vNivel, vArt, "", contador, RN!cantmostrar, RN!vrefd

        RN.MoveNext
    Wend
    RN.Close
    
    'En factras
    s = "select scafpc.numfactu clave,cantidad cantmostrar,scafpc.fecfactu mifecha,scafpc.codprove idrefd,scafpc.nomprove vrefd"
    s = s & " FROM slifpc,scafpc where scafpc.codprove=slifpc.codprove and scafpc.fecfactu=slifpc.fecfactu"
    s = s & " and scafpc.numfactu=slifpc.numfactu "
    'El arti/lote/almacen                                                     'ojo a la S
    s = s & " AND slifpc.codartic = " & DBSet(vArt, "T") & "AND numlotes = " & DBSet(NLote, "T")
    s = s & " and codalmac = " & mCodalmac
    RN.Open s, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RN.EOF
        InsertarLinea "F.PRO " & RN!Clave, vNivel, vArt, "", contador, RN!cantmostrar, RN!vrefd
        DesdeUnAlbaran = True
        RN.MoveNext
    Wend
    RN.Close
    
    
    If Not DesdeUnAlbaran Then
        'NO es de un albaran. Podriamos ver, si el lote acaba en C
        'que viene de un copupage
        If Right(NLote, 1) = "C" Then
            s = "select olicoupage.codigo,sum(kilos) cantmostrar,deposito from olicoupage,olicoupagelin WHERE "
            s = s & " olicoupage.codigo=olicoupagelin.codigo AND "
            s = s & " numlote =" & DBSet(NUmlote, "T") & " And codAlmac = " & mCodalmac & "  group by 1"
            RN.Open s, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RN.EOF Then
                s = "CUP" & Format(RN!Codigo, "0000")
                If RN!Deposito > 0 Then s = s & "[" & RN!Deposito & "]"
                InsertarLinea s, vNivel, vArt, "", contador, RN!cantmostrar, s
                
                RN.MoveNext
            End If
            RN.Close
            
        Else
                'M0IXENT    . No compra. Es una almazara. Mete las existencias por regularizacion
                If vParamAplic.QUE_EMPRESA >= 2 Then
                    s = "select * from slhmov where "
                    'El arti/lote/almacen                                                     'ojo a la S
                    s = s & "  slhmov.codartic = " & DBSet(codartic, "T") & "AND slhmov.numlote = " & DBSet(NUmlote, "T")
                    'S = S & " and codalmac = " & mCodalmac
                    RN.Open s, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    While Not RN.EOF
                        s = "REG" & Format(RN!codMovim, "00000")
                        InsertarLinea s, vNivel, vArt, "", contador, RN!Cantidad, "ALMAZARA"
                        RN.MoveNext
                    Wend
                    RN.Close
                        
                    
                End If

        
        End If
    End If
    
    
    
    
    EnAlbaranCompras = True
    Exit Function
EEnAlbaranCompras:
    MuestraError Err.Number
End Function







Private Sub InsertarLinea(Clave As String, vNivel As Integer, vArt As String, vL As String, ByRef contador As Long, Cantidad As Currency, NomClienProv As String)
Dim SQL As String

    SQL = "insert into `tmptraza` (`codusu`,`contador`,`artppal`,`numloteppal`,"
    SQL = SQL & "`nivle`,`idOperacion`,`artic2`,`numlote2`,`cantidad`,`abs_cantidad`,nomclien) values ("
    'Valroes
    SQL = SQL & vUsu.Codigo & "," & contador & "," & DBSet(Me.codartic, "T") & "," & DBSet(Me.NUmlote, "T")
    SQL = SQL & "," & vNivel & "," & DBSet(Clave, "T") & "," & DBSet(vArt, "T") & "," & DBSet(vL, "T", "N")
    SQL = SQL & "," & TransformaComasPuntos(CStr(Cantidad)) & ",0," & DBSet(NomClienProv, "T", "S") & ")"
    conn.Execute SQL
    contador = contador + 1
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
        EnAlbaranVentas2 ColVentas, Me.codartic, NUmlote
        
    Next J
    
    
    
    'AHora veremos si esta cupado, y/o en produccion
        


End Function



'Le pasamos una coleccion para que la rellene
Private Function EnAlbaranCompras2(ByRef vLista As Collection) As Boolean
Dim RN As ADODB.Recordset
Dim s As String


    
    On Error GoTo EEnAlbaranCompras
    EnAlbaranCompras2 = False
    Set RN = New ADODB.Recordset
    s = "select slialp.numalbar clave,cantidad cantmostrar,scaalp.fechaalb mifecha ,scaalp.codprove idRefD,scaalp.nomprove vrefd"
    s = s & " from slialp,scaalp where scaalp.codprove=slialp.codprove and scaalp.fechaalb=slialp.fechaalb and scaalp.numalbar=slialp.numalbar"
    'El arti/lote/almacen                                               'ojo a la S
    s = s & " AND slialp.codartic = " & DBSet(Me.codartic, "T") & "AND numlotes = " & DBSet(Me.NUmlote, "T")
    s = s & " and codalmac = " & mCodalmac



    RN.Open s, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RN.EOF
        s = RN!NumAlbar
        vLista.Add s

        RN.MoveNext
    Wend
    RN.Close
    
    'En factras
    s = "select scafpc.numfactu clave,cantidad cantmostrar,scafpc.fecfactu mifecha,scafpc.codprove idrefd,scafpc.nomprove vrefd"
    s = s & " FROM slifpc,scafpc where scafpc.codprove=slifpc.codprove and scafpc.fecfactu=slifpc.fecfactu"
    s = s & " and scafpc.numfactu=slifpc.numfactu "
    'El arti/lote/almacen                                                     'ojo a la S
    s = s & " AND slifpc.codartic = " & DBSet(codartic, "T") & "AND numlotes = " & DBSet(NUmlote, "T")
    s = s & " and codalmac = " & mCodalmac
    RN.Open s, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RN.EOF
        s = RN!Clave & "|" & RN!cantmostrar & "|"
        vLista.Add s
        RN.MoveNext
    Wend
    RN.Close
    
    
    
    'M0IXENT    . No compra. Es una almazara. Mete las existencias por regularizacion
    If vParamAplic.QUE_EMPRESA >= 2 Then
        s = "select * from slhmov where "
        'El arti/lote/almacen                                                     'ojo a la S
        s = s & " AND slhmov.codartic = " & DBSet(codartic, "T") & "AND slhmov = " & DBSet(NUmlote, "T")
        'S = S & " and codalmac = " & mCodalmac
        RN.Open s, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RN.EOF
            s = RN!Clave & "|" & RN!cantmostrar & "|"
            vLista.Add s
            RN.MoveNext
        Wend
        RN.Close
            
        
    End If
    
    EnAlbaranCompras2 = True
    Exit Function
EEnAlbaranCompras:
    MuestraError Err.Number
End Function


'En collection metere lineas con el trzo de sql
' '`albventa`,fechaventa,`nomclien`,`cantidad2`
Private Function EnAlbaranVentas2(cVentas As Collection, vArt As String, NLote As String) As Boolean
Dim RN As ADODB.Recordset
Dim s As String

    '`albventa`,fechaventa,`nomclien`,`cantidad2`

    On Error GoTo EEnAlbaranVentas
    EnAlbaranVentas2 = False
    Set RN = New ADODB.Recordset
    s = "Select * from slialb,slialblotes,scaalb    Where  slialb.CodTipom = scaalb.CodTipom"
    s = s & " AND slialb.numalbar = scaalb.numalbar AND slialb.codTipoM = slialblotes.codTipoM"
    s = s & " AND codartic = '" & vArt & "' and codalmac = " & Me.codAlmac
    s = s & " AND numlote = " & DBSet(NLote, "T") & " ORDER BY fechaalb"


    RN.Open s, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RN.EOF
        'De momento NOMBRE NO LO TRATP
        s = "'" & RN!Codtipom & Format(RN!NumAlbar, "00000") & "'," & DBSet(RN!FechaAlb, "F")
        s = s & ",NULL," & TransformaComasPuntos(CStr(RN!Cantidad))
        cVentas.Add s

        RN.MoveNext
    Wend
    RN.Close
    
    'En factras
    s = "Select * from slifac,slifaclotes    Where   slifac.codTipoM = slifaclotes.codTipoM"
    s = s & " and slifac.numfactu  = slifaclotes.numfactu and slifac.fecfactu  = slifaclotes.fecfactu"
    s = s & " and slifac.codtipoa = slifaclotes.codtipoa and slifac.numalbar = slifaclotes.numalbar"
    s = s & " and slifac.numlinea = slifaclotes.numlinea  "
    s = s & " AND codartic = '" & vArt & "' and codalmac = " & Me.codAlmac
    s = s & " AND numlote = " & DBSet(NLote, "T") & " ORDER BY slifac.fecfactu,slifac.numfactu"
    

    RN.Open s, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RN.EOF
        s = "'" & RN!Codtipom & Format(RN!NumFactu, "00000") & "'," & DBSet(RN!Fecfactu, "F")
        s = s & ",NULL," & TransformaComasPuntos(CStr(RN!Cantidad))
        cVentas.Add s

        RN.MoveNext
    Wend
    RN.Close
    
    
    Exit Function
EEnAlbaranVentas:
    MuestraError Err.Number
End Function








Public Sub GeneracionProcesosDesdeCompra3(ByRef contador As Long, ByVal vArticulo As String, ByVal vLote As String, ByRef LB2 As Label)
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
    RN.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RN.EOF
        Aux = RN!codartic & "|" & RN!NUmlote & "|" & RN!cantlote & "|" & RN!Codigo & "|"
        Col.Add Aux
        RN.MoveNext
    Wend
    RN.Close
    If Col.Count > 0 Then
        

        For J = 1 To Col.Count
            Aux = "Coupage (" & RecuperaValor(Col.Item(J), 4) & ")"
            InsertarLinea Aux, Nivel, CStr(RecuperaValor(Col.Item(J), 1)), CStr(RecuperaValor(Col.Item(J), 2)), contador, CCur(CStr(RecuperaValor(Col.Item(J), 3))), ""
            GeneracionProcesosDesdeCompra Nivel + 1, contador, CStr(RecuperaValor(Col.Item(J), 1)), CStr(RecuperaValor(Col.Item(J), 2)), LB2, True
        Next J
    End If
    
    'AHORA produccion
    Aux = "select sliordprlotes.codartic,sliordprlotes.numlote,sliordprlotes.cantlote,sliordprlotes.codigo from sliordprlotes ,sliordpr2lotes"
    Aux = Aux & " Where sliordprlotes.codigo = sliordpr2lotes.codigo"
    Aux = Aux & " and sliordprlotes.codalmac =sliordpr2lotes.codalmac"
    Aux = Aux & " and sliordprlotes.codartic =sliordpr2lotes.codartic"
    Aux = Aux & " AND sliordpr2lotes.codarti2 = " & DBSet(vArticulo, "T")
    Aux = Aux & " AND sliordpr2lotes.numlote = " & DBSet(vLote, "T")
    RN.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set Col = New Collection
    While Not RN.EOF
        Aux = RN!codartic & "|" & RN!NUmlote & "|" & RN!cantlote & "|" & RN!Codigo & "|"
        Col.Add Aux
        RN.MoveNext
    Wend
    RN.Close
    
    If Col.Count > 0 Then
        
        For J = 1 To Col.Count
            Aux = "Produc. (" & RecuperaValor(Col.Item(J), 4) & ")"
            InsertarLinea Aux, Nivel, CStr(RecuperaValor(Col.Item(J), 1)), CStr(RecuperaValor(Col.Item(J), 2)), contador, CCur(CStr(RecuperaValor(Col.Item(J), 3))), ""
            GeneracionProcesosDesdeCompra Nivel + 1, contador, CStr(RecuperaValor(Col.Item(J), 1)), CStr(RecuperaValor(Col.Item(J), 2)), LB2, True
        Next J
    End If
    
    
        
    
    
    Set Col = Nothing
    Set RN = Nothing
    
    'Veremos para este articulo LOTE las ventas
    
    EnAlbaranVentas_ Nivel, vArticulo, vLote, contador, FactorMultiplicador1
    
    
End Sub


'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'
' Obtener albaranes molturacion
'
'
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------

Public Sub GeneracionHastaMolturacion_()
Dim contador As Long


         
    'obtener desde albaranes
    contador = 0
    ComprobarAceite mCodartic, mNumlote, True, 0, contador, -1
    
    
    



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
    Cad = DevuelveDesdeBD(conAri, "conjunto", "sartic", "codartic", Me.codartic, "T")
    If Cad = "" Then Cad = "0"
    Embasado = Cad = "1"
    
    If Embasado Then
    
        'Proceso de envasado
        ' vemos la produccion donde esta
        SQL = "codalmac"
        Cad = "codartic ='" & codartic & "' and numlote"
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
            Cad = Cad & " AND prodlin.codartic = " & DBSet(codartic, "T") & "  AND prodtrazcompo.lotetraza = " & DBSet(NUmlote, "T")
            RCompo.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RCompo.EOF Then
                Encontrado = True
                Produccion = RCompo!Codigo
                LinProd = RCompo!idlin
            Else
                If vParamAplic.QUE_EMPRESA < 2 Then MsgBox "No se ha encontrado en lineas de produccion"
            
            End If
            RCompo.Close
            
        End If
        
        
        
        
        If Encontrado Then
            If Not BUscarEnNuevaProduccion Then
                '**********************************
                '
                '   Esta en produccion
            
            
                Cad = "select * from sliordpr2lotes,sartic  where sliordpr2lotes.codarti2=sartic.codartic"
                Cad = Cad & " AND codigo=" & Produccion & " and sliordpr2lotes.codartic=" & DBSet(codartic, "T") & " AND codalmac= " & codAlmac
                Cad = Cad & " ORDER BY FactorConversion desc"
                RCompo.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not RCompo.EOF
                    Cad = ""
                    
                    If RCompo!FactorConversion <> 1 Then
        
                        'Vemos si viene de un cupaje, o viene de compra directa
                        'Veremos si para un aceite y lote este, ha sido desde coupaje
                        MiNivel = MiNivel - 1
                        ComprobarAceite RCompo!codarti2, DBLet(RCompo!NUmlote), True, MiNivel, vContador, -1
                      
                        
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
                RCompo.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not RCompo.EOF
                    Cad = ""
                    
                    If RCompo!FactorConversion <> 1 Then
        
                        'Vemos si viene de un cupaje, o viene de compra directa
                        'Veremos si para un aceite y lote este, ha sido desde coupaje
                        ComprobarAceite RCompo!codartic, DBLet(RCompo!NUmlote), True, MiNivel, vContador, -1
                      
                        
                    Else
                        If Not SoloAceite Then
                            'Intentamos ver en que ALBARAN esta comprado
                            materiaPrimaProduccion RCompo!codartic, DBLet(RCompo!NUmlote), MiNivel, vContador
                        
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
        If vParamAplic.QUE_EMPRESA = 4 Then
            'COOP LA VALL
            'Esta vendiendo directamente "granel"
            
            
            GeneracionHastaMolturacion_
            
        Else
            If MostrarMensaje Then MsgBox "No embasado", vbExclamation
        End If
    End If
    Set RCompo = Nothing
End Function

Private Sub materiaPrimaProduccion(ByRef Art As String, ByRef NLote As String, Nivel As Integer, ByRef contador As Long)
Dim T As ADODB.Recordset



    'En facturas. Si esta no busco mas
    Set T = New ADODB.Recordset
    SQL = "Select *  FROM slifpc,scafpc where scafpc.codprove=slifpc.codprove and "
    SQL = SQL & " scafpc.fecfactu=slifpc.fecfactu and scafpc.numfactu=slifpc.numfactu "
    SQL = SQL & " AND slifpc.codartic = " & DBSet(Art, "T") & " AND numlotes = " & DBSet(NLote, "T")
    SQL = SQL & " and codalmac = " & Me.codAlmac
    
    T.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    If Not T.EOF Then
        SQL = T!NumFactu & " " & T!Fecfactu & " (" & T!codProve & ")"
        InsertarLinea SQL, Nivel, Art, NLote, contador, 0, T!nomprove
        
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
            T.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            SQL = ""
            If Not T.EOF Then
                SQL = T!NumAlbar & " " & T!FechaAlb & " (" & T!codProve & ")*"
                InsertarLinea SQL, Nivel, Art, NLote, contador, 0, T!nomprove
                T.MoveNext
            End If
            T.Close
            
            
    End If
    Set T = Nothing
    
    
    
    
    
    
End Sub







Private Sub materiaPrimaMolturacion(ByRef Art As String, ByRef NLote As String, Nivel As Integer, ByRef contador As Long)
Dim T As ADODB.Recordset
Dim idMolturacion As Long
Dim C As String
Dim OK As Boolean
Dim MiNivel As Integer
Dim T2 As ADODB.Recordset
Dim K As Integer

    'En facturas. Si esta no busco mas
    Set T = New ADODB.Recordset
    Set T2 = New ADODB.Recordset
    'Obtenemos la molturacion
    C = "codartic=" & DBSet(Art, "T") & " AND numlote"
    C = DevuelveDesdeBD(conAri, "numalbar", "spartidas", C, NLote, "T")
    
    If Mid(C, 1, 3) = "REG" Then
        InsertarLinea "Regularizacion", Nivel, Art, NLote, contador, 0, C
        Exit Sub
    End If
    
    
    
    
    OK = False
    idMolturacion = Val(C)
    If idMolturacion = 0 Then
         InsertarLinea "Err", Nivel, Art, NLote, contador, 0, "Error obt. molturacion"
    
    Else
    
    
        C = "Select * from vallalmazaraproceso where loteproducido= " & DBSet(NLote, "T")
        T2.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        
        While Not T2.EOF
            
            idMolturacion = T2!ID
            C = "select fecha,deposito,litros,FechaFin,HoraFin,horainicio from vallalmazaraproceso where id=" & idMolturacion
            T.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If T.EOF Then
                 InsertarLinea "err", Nivel, Art, NLote, contador, 0, "Error obt. molturacion II"
            Else
            
                C = "Id:" & idMolturacion & "  Dep:" & T!Deposito & "  F:" & Format(T!Fecha, "dd/mm") & Format(T!horainicio, " hh:nn")
                InsertarLinea C, Nivel, Art, NLote, contador, 0, "Fin: " & Format(T!FechaFin, "dd/mm") & " " & Format(T!HoraFin, " hh:nn")
                OK = True
             
             End If
             T.Close
        
    
            SQL = "SELECT scaalp.numalbar, scaalp.fechaalb, slialp.nomartic,scaalp.codprove, scaalp.nomprove ,slialp.codartic,cantidad"
            SQL = SQL & "  FROM scaalp,slialp,sartic,sfamia WHERE"
            SQL = SQL & " scaalp.numalbar= slialp.numalbar and scaalp.fechaalb  = slialp.fechaalb and scaalp.codprove = slialp.codprove and slialp.codartic=sartic.codartic and"
            SQL = SQL & " sfamia.codfamia=sartic.codfamia and tipfamia=30 and"
            SQL = SQL & " (scaalp.numalbar,scaalp.fechaalb,scaalp.codprove) in"
            SQL = SQL & " (select numalbar,fechaalb,codprove from vallalmazaraprocesoalb where id=" & idMolturacion & " )"
            
            
            T.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            SQL = ""
            MiNivel = Nivel + 1
            If T.EOF Then
                T.Close
                'Puede que haya facturado el albaran, con lo cual esta en scafapc
                SQL = "SELECT scafpa.numalbar, scafpa.fechaalb, slifpc.nomartic,scafpa.codprove, nomprove ,slifpc.codartic,cantidad"
                SQL = SQL & " fROM scafpa,slifpc,sartic,sfamia,sprove WHERE"
                SQL = SQL & " scafpa.numfactu= slifpc.numfactu and scafpa.numalbar= slifpc.numalbar and"
                SQL = SQL & " scafpa.FecFactu = slifpc.FecFactu And scafpa.codProve = slifpc.codProve"
                SQL = SQL & " and slifpc.codartic=sartic.codartic and sfamia.codfamia=sartic.codfamia"
                SQL = SQL & " and sfamia.codfamia=sartic.codfamia and tipfamia=30"
                SQL = SQL & " and scafpa.codprove=sprove.codprove AND "
                SQL = SQL & " (scafpa.numalbar,scafpa.fechaalb,scafpa.codprove) in"
                SQL = SQL & " (select numalbar,fechaalb,codprove from vallalmazaraprocesoalb where id=" & idMolturacion & " )"
                    
                T.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                SQL = ""
                
                
            End If
            While Not T.EOF
                SQL = T!FechaAlb & " "
                InsertarLinea SQL, MiNivel, T!codartic, T!NumAlbar, contador, T!Cantidad, T!nomprove
                
                T.MoveNext
            Wend
            T.Close
        
    
    
            T2.MoveNext
        Wend
        T2.Close
        Set T2 = Nothing
    End If
              
    
    
    
        
   
    Set T = Nothing
    Set T2 = Nothing
    
    
    
End Sub

















Private Sub ComprobarAceite(ByRef Art As String, ByRef NLote As String, ComprobarCoupaje As Boolean, Nivel As Integer, ByRef contador As Long, ByRef DepositoDestino As Integer)
Dim T As ADODB.Recordset
Dim Col As Collection
Dim I As Integer
Dim Seguir As Boolean
Dim C As String
Dim MiNivel As Integer
Dim Deposito As Integer
Dim ParaElLote As String
Dim Fecha As Date
Dim SubLotes As String

    If vParamAplic.QUE_EMPRESA = 4 Then
    
        'puede ser que sea un trasiego
        I = InStr(6, NLote, "T")   'ejemplo lote: MOSTRA45-2016T1
        If I > 0 Then
            'Verifiquemos que es un trasiego
            SQL = "codartic=" & DBSet(Art, "T") & " AND numlote"
            SQL = DevuelveDesdeBD(conAri, "numalbar", "spartidas", SQL, NLote, "T")
            If Mid(SQL, 1, 3) = "TRS" Then
                'Es un trasiego
                'El lote original
                I = InStr(6, NLote, "T")   'ejemplo lote: MOSTRA45-2016T1
                If I = 0 Then
                    'MAL
                     
                Else
                    C = Mid(NLote, 1, I - 1)
                    NLote = C
                End If
            End If
        End If
    End If


    'Veo si esta en coupaje
    Seguir = True
    If ComprobarCoupaje Then
            SQL = "select olicoupage.codigo,deposito,fecha from"
            SQL = SQL & " olicoupage,olicoupagelin where olicoupage.codigo=olicoupagelin.codigo"
            'El arti/lote/almacen
            SQL = SQL & " AND olicoupage.codartic = " & DBSet(Art, "T") & " AND olicoupage.numlote = " & DBSet(NLote, "T")
            SQL = SQL & " and codalmac = " & mCodalmac
            SQL = SQL & " group by 1"
            Set T = New ADODB.Recordset
            T.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            SQL = ""
            If Not T.EOF Then
                SQL = T!Codigo
                Deposito = T!Deposito
                Fecha = T!Fecha
            End If
            T.Close
            'If SQL = "248" Then
            If SQL <> "" Then

                SQL = "select * from olicoupagelinlotes where codigo=" & SQL
                T.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                Set Col = New Collection
                SubLotes = ""
                While Not T.EOF
                
                           
                    C = T!NUmlote
                    I = InStr(6, C, "T")   'ejemplo lote: MOSTRA45-2016T1
                    If I > 0 Then
                        C = Mid(C, 1, I - 1)
                    End If
                    
                    If InStr(1, SubLotes, C) = 0 Then
                        SubLotes = SubLotes & C
                        C = T!codartic & "|" & C & "|"
                        Col.Add C
                    Else
                        '
                    End If
                    SQL = T!Codigo
                    T.MoveNext
                Wend
                
                'Comprobamos si viene de coupar o no
                If Col.Count = 0 Then
                    MsgBox "Error leyendo coupage", vbExclamation
                Else
                    ParaElLote = NLote
                    If vParamAplic.QUE_EMPRESA = 4 Then
                        If Nivel = 0 Then
                            ParaElLote = NLote
                        Else
                            
                        End If
                        If Format(Fecha, "hh:nn") = "00:00" Then
                        
                            ParaElLote = ParaElLote & " " & Format(Fecha, "dd/mm/yyyy")
                        Else
                            ParaElLote = ParaElLote & " " & Format(Fecha, "dd/mm/yy hh:nn")
                        End If
                    End If
                
                    
                    SQL = "CUP " & SQL & " (Dep:" & Deposito & " )"
                    MiNivel = Nivel + 1
                    InsertarLinea SQL, MiNivel, Art, ParaElLote, contador, 0, ""
                    For I = 1 To Col.Count
                        C = Col.Item(I)
                        ComprobarAceite RecuperaValor(C, 1), RecuperaValor(C, 2), True, MiNivel, contador, Deposito
                    Next I
                    
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
            
            
            
            'LA VALL.
            ' No compran el aceit, lo producen
            'por lo tanto, para la vall, el proceso este NO es de ir a comprar, es de ir a molturar
            If vParamAplic.QUE_EMPRESA = 4 Then
                MiNivel = Nivel + 1
                materiaPrimaMolturacion Art, NLote, MiNivel, contador
              
            Else
            'El arti/lote/almacen
                materiaPrimaProduccion Art, NLote, Nivel, contador
            
            End If
            
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
Dim I As Integer

    SQL = "DELETE FROM spartidaslin WHERE id = " & midPartida
    conn.Execute SQL
    If CuantasEtiquetas = 0 Then Exit Sub  'Si no hay salgo
    
    SQL = ""
    For I = 1 To CuantasEtiquetas
        SQL = SQL & ", (" & midPartida & "," & I & ",0)"
    Next
    SQL = Mid(SQL, 2) 'QUITO la primera coma
    SQL = "INSERT INTO spartidaslin(id,bulto,cantidad) VALUES " & SQL
    EjecutaSQL conAri, SQL, True
End Sub







Public Function CuantasEtiquetas() As Integer
Dim I As Integer

    Set RS = New ADODB.Recordset
    SQL = "select count(*) from spartidaslin WHERE id = " & midPartida
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CuantasEtiquetas = 0
    If Not RS.EOF Then CuantasEtiquetas = DBLet(RS.Fields(0), "N")
    RS.Close
    Set RS = Nothing
'    EjecutaSQL conAri, SQL, True
End Function


Public Function TieneEtiquetasYAenProduccion() As Boolean
Dim I As Integer

    Set RS = New ADODB.Recordset
    SQL = "select count(*) from spartidaslin WHERE id = " & midPartida & " AND fechaulizada > '1900-01-01'"
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    TieneEtiquetasYAenProduccion = False
    If Not RS.EOF Then TieneEtiquetasYAenProduccion = DBLet(RS.Fields(0), "N") > 0
    RS.Close
    Set RS = Nothing
 '   EjecutaSQL conAri, SQL, True
End Function





Public Function AjustarFinPartida()
    SQL = "UPDATE spartidas set cantotal=0 "
    SQL = SQL & " WHERE id = " & midPartida
    conn.Execute SQL
End Function





Public Sub LeerDatosPesajes()

    mPesoMedio = 0   'Estos se pondran cuando se solicite el evento leerpesos
    mDesviacion = 0


    On Error GoTo ePesadas
    
    Set RS = New ADODB.Recordset
    'S e pueden hacer sum, reondando....
    ',std(round(peso,1)),sum(peso),count(*)
    SQL = "select avg(peso),std(peso) from spartidaspesos  where idpartida=" & midPartida
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then mPesoMedio = RS.Fields(0)
        If Not IsNull(RS.Fields(1)) Then mDesviacion = Round2(RS.Fields(1), 4)
    End If

    RS.Close
ePesadas:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set RS = Nothing


End Sub


