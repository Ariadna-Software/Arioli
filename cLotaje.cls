VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLotaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'=======================================================================
'MODULO para el control de lotes
'Cada vez que se asigne un lote se insertara una linea que recogera el movimiento
'de ese lote

'=======================================================================


'------------------------------------------------
'Variables del modulo CStock
'------------------------------------------------
Private mvarcodartic As String
Private mvarCodAlmac As Integer
Private mvarNumLote As String
Private mvarFechaMov As String
Private mvarHoraMov As String
Private mvarTipoMov As Byte '0=Salida, 1=Entrada
Private mvarDetaMov As String 'referencia a codtipom de la tabla stipom   LOT: Insercion Manual
Private mvarCantidad As Currency 'Cantidad de Stock a mover

Private mvarDocum As String
    'Codigo de la tabla de Movimientos, Traspaso, Albaran,... al que va asociado el movimiento
    'Sera para
    '   PRODUCCION.  codprod
    '   COUPA        codigo
    '   CLIENTES     Numero albaran o fatura
    '   PROVEEDOR      "          "
    '   REG          NUmero regularicacion mov almacen


Private mvarLinea As Integer 'Numero de linea del documento linfac...


'Variables NUEVAS

Private mvarSubLinea As Integer   'Para una misma linea de albaran pueden ir dos lotes diferentes. Para saber cual de los dos es
Private mvarcodArti2 As String    'para las sublineas de produccion
Private mvarProvCliTra As Long
    'Codprove en albra/fra proveedor
    'Clie        en cliente
    'Oper        en PRoD y coupa  y REG



'------------------------------------------------
'Propiedades del modulo CStock
'------------------------------------------------

Public Property Let codartic(ByVal vData As String)
     mvarcodartic = vData
End Property

Public Property Get codartic() As String
     codartic = mvarcodartic
End Property


Public Property Let codAlmac(ByVal vData As Integer)
     mvarCodAlmac = vData
End Property

Public Property Get codAlmac() As Integer
     codAlmac = mvarCodAlmac
End Property


Public Property Let Fechamov(ByVal vData As String)
     mvarFechaMov = vData
End Property

Public Property Get Fechamov() As String
     Fechamov = mvarFechaMov
End Property



Public Property Let HoraMov(ByVal vData As String)
     mvarHoraMov = vData
End Property

Public Property Get HoraMov() As String
     HoraMov = mvarHoraMov
End Property

'0=Salida, 1=Entrada
Public Property Let tipoMov(ByVal vData As String)
     mvarTipoMov = vData
End Property

Public Property Get tipoMov() As String
     tipoMov = mvarTipoMov
End Property

Public Property Let DetaMov(ByVal vData As String)
     mvarDetaMov = vData
End Property
Public Property Get DetaMov() As String
     DetaMov = mvarDetaMov
End Property


Public Property Let Cantidad(ByVal vData As Single)
     mvarCantidad = vData
End Property

Public Property Get Cantidad() As Single
     Cantidad = mvarCantidad
End Property



Public Property Let Documento(ByVal vData As String)
     mvarDocum = vData
End Property

Public Property Get Documento() As String
     Documento = mvarDocum
End Property


Public Property Let LineaDocu(ByVal vData As Integer)
     mvarLinea = vData
End Property

Public Property Get LineaDocu() As Integer
     LineaDocu = mvarLinea
End Property



'Para una misma linea de albaran pueden ir dos lotes diferentes. Para saber cual de los dos es
Public Property Let SubLinea(ByVal vData As Integer)
     mvarSubLinea = vData
End Property

Public Property Get SubLinea() As Integer
     SubLinea = mvarSubLinea
End Property


'para las sublineas de produccion
Public Property Let codarti2(ByVal vData As String)
     mvarcodArti2 = vData
End Property

Public Property Get codarti2() As String
     codarti2 = mvarcodArti2
End Property



Public Property Let ProvCliTra(ByVal vData As Long)
     mvarProvCliTra = vData
End Property

Public Property Get ProvCliTra() As Long
     ProvCliTra = mvarProvCliTra
End Property



Public Property Let NUmlote(ByVal vData As String)
     mvarNumLote = vData
End Property

Public Property Get NUmlote() As String
     NUmlote = mvarNumLote
End Property




Private Function MontaWHERE(DesdeElDocumento As Boolean) As String
Dim SQL As String
    SQL = " codartic = " & DBSet(Me.codartic, "T")
    If Not DesdeElDocumento Then SQL = SQL & " AND numlote = " & DBSet(Me.NUmlote, "T")
    SQL = SQL & " AND codalmac = " & Me.codAlmac
    SQL = SQL & " AND detamovi = " & DBSet(Me.DetaMov, "T")
    'En funcion del tipo
    Select Case Me.DetaMov
    Case "PRO"
        SQL = SQL & " AND document = '" & Val(Me.Documento) & "'"
        
        'Es de subproductos
        If Me.codarti2 <> "" Then
            SQL = SQL & " AND codarti2 = " & DBSet(Me.codarti2, "T")
            SQL = SQL & " AND numlinea = " & Me.LineaDocu
        End If
        
    Case "CUP"
        SQL = SQL & " AND codtipom = " & Val(Me.Documento)
        If Me.LineaDocu > 0 Then SQL = SQL & " AND numlinea = " & Me.LineaDocu
    
    Case "ALC"
        'Compras
        SQL = SQL & " AND document = " & DBSet(Me.Documento, "T")
        SQL = SQL & " AND numlinea = " & Me.LineaDocu
        SQL = SQL & " AND codprocliope = " & Me.ProvCliTra
        SQL = SQL & " AND fechamov = " & DBSet(Me.Fechamov, "F")
    Case Else
        If Mid(Me.DetaMov, 1, 2) = "AL" Then
            'Albarabes y/o facturas
            SQL = SQL & " AND document = '" & Val(Me.Documento) & "'"
            SQL = SQL & " AND numlinea = " & Me.LineaDocu
            If Not DesdeElDocumento Then SQL = SQL & " AND linea2 = " & Me.SubLinea
        Else
            'NO se puede leer el movimiento
            MsgBox "Tipo no tratado: " & Me.DetaMov, vbExclamation
            SQL = SQL & " AND codartic ='ee@@'" 'para que no devuelve nada
        End If
        
    End Select
    MontaWHERE = SQL
    
End Function



Public Function ModificarMovimArticulosLotaje(ModificaHoramov As Boolean) As Boolean
'Modifica la cantidad  en la tabla "smovalotes"
Dim SQL As String
On Error Resume Next
    
    
    SQL = "UPDATE smovalotes SET cantidad=" & DBSet(Cantidad, "N")
    If ModificaHoramov Then SQL = SQL & ", Horamovi = " & DBSet(mvarHoraMov, "FH")
    SQL = SQL & " WHERE " & MontaWHERE(False)  'Lee lote y lineas
    conn.Execute SQL
    
    If Err.Number <> 0 Then
        'Hay error ,
        ModificarMovimArticulosLotaje = False
    Else
        ModificarMovimArticulosLotaje = True
    End If
End Function





'Para leer un registro de los movimientos.
' Para leer es importante que tenga puesto el detamovi
'el articulo, lote
'Ya que en funcion del mismo monta un select u otro
Public Function Leer() As Boolean
Dim RT As ADODB.Recordset
Dim SQL As String
Dim Tipo
    
    On Error GoTo ELeer
    Leer = False
    SQL = MontaWHERE(False)
    SQL = "select * from smovalotes WHERE " & SQL
    
    Set RT = New ADODB.Recordset
    RT.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RT.EOF Then
        'NO EXISTE
    
    Else
        'SI EXISTE
        Leer = True
        
    End If
    RT.Close
    Exit Function
ELeer:
    MuestraError Err.Number, Err.Description
    Set RT = Nothing
End Function




Public Function InsertarLote() As Boolean
Dim SQL As String
    On Error Resume Next
    InsertarLote = False
    SQL = "insert into `smovalotes` (`codartic`,`numlote`,`codalmac`,`fechamov`,`horamovi`,"
    SQL = SQL & "`tipomovi`,`detamovi`,`cantidad`,`document`,`numlinea`,`linea2`,`codarti2`,`codprocliope`) "
    SQL = SQL & " VALUES ("
    'Comun
    SQL = SQL & DBSet(mvarcodartic, "T")
    SQL = SQL & ", " & DBSet(mvarNumLote, "T")
    SQL = SQL & ", " & mvarCodAlmac
    SQL = SQL & ", " & DBSet(mvarFechaMov, "F")
    SQL = SQL & ", " & DBSet(mvarHoraMov, "FH")
    SQL = SQL & ", " & mvarTipoMov
    SQL = SQL & ", " & DBSet(mvarDetaMov, "T")
    SQL = SQL & ", " & DBSet(mvarCantidad, "N")
    SQL = SQL & ", " & DBSet(mvarDocum, "T")
    SQL = SQL & ", " & mvarLinea
    SQL = SQL & ", " & mvarSubLinea
    SQL = SQL & ", " & DBSet(mvarcodArti2, "T", "S")
    SQL = SQL & ", " & ProvCliTra & ")"
    conn.Execute SQL
    If Err.Number <> 0 Then
        MuestraError Err.Number
    Else
        InsertarLote = True
    End If
End Function


'DesdeElDocumento
'Indica que quiere borrar todos los movimientos asociados a un documento en particular, con lo cual NO hyabque indicar LOTE
Public Function EliminarMovimArticulosLotaje(DesdeElDocumento As Boolean) As Boolean
'Modifica la cantidad  en la tabla "smovalotes"
Dim SQL As String
On Error Resume Next
    
    SQL = MontaWHERE(DesdeElDocumento)
    SQL = "DELETE FROM smovalotes  WHERE " & SQL
    conn.Execute SQL
    
    If Err.Number <> 0 Then
        'Hay error ,
        EliminarMovimArticulosLotaje = False
    Else
        EliminarMovimArticulosLotaje = True
    End If
End Function
