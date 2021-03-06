VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CNumSerie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'TABLA: sserie
'N� SERIE DE EQUIPOS

'Variables locales que contienen valores de propiedad
Private mNumserie As String 'num. serie
Private mCodartic As String 'cod. articulo
Private mCodtipar As String 'cod. tipo articulo

'Private mCodClien As Integer 'cod. cliente
Private mCodClien As Long 'cod. cliente  FAVID. En la bD es un INT

Private mCoddirec As String 'cod. direc./dpto
Private mTieneman As Boolean 'tiene manteniento S/N
Private mNummante As String 'num. mantenimiento
Private mUltrepar As String 'ult. repar
Private mFingaran As String 'fin garantia

Private mCodTipoM As String 'tipo movim.
Private mFechavta As String 'Fecha de venta
Private mNumalbar As String 'Numero de albaran venta
Private mNumFactu As String 'Numero de factura venta
Private mNumLine1 As String 'Num. linea del albaran

'Private mCodProve As Integer 'cod. proveedor
Private mCodProve As Long 'cod. proveedor DAVID

Private mNumalbpr As String 'mum. Albaran prove
Private mFechaCom As String 'Fecha de compra
Private mNumline2 As String 'Num. linea Albaran prove



'------------------------------------------------
'Propiedades del modulo CSerie
'------------------------------------------------

'**** num. Serie

Public Property Let numSerie(ByVal vData As String)
     mNumserie = vData
End Property

Public Property Get numSerie() As String
     numSerie = mNumserie
End Property


'**** cod. Artic

Public Property Let Articulo(ByVal vData As String)
     mCodartic = vData
End Property

Public Property Get Articulo() As String
     Articulo = mCodartic
End Property


'**** Tipo Articulo

Public Property Let TipoArt(ByVal vData As String)
     mCodtipar = vData
End Property

Public Property Get TipoArt() As String
     TipoArt = mCodtipar
End Property


'**** Fecha fin garantia
Public Property Let FinGarantia(ByVal vData As String)
     mFingaran = vData
End Property

Public Property Get FinGarantia() As String
     FinGarantia = mFingaran
End Property



'**** Cod. Cliente

Public Property Let Cliente(ByVal vData As Long)
     mCodClien = vData
End Property

Public Property Get Cliente() As Long
     Cliente = mCodClien
End Property


'**** Cod. Direc/dpto

Public Property Let DirDpto(ByVal vData As String)
     mCoddirec = vData
End Property

Public Property Get DirDpto() As String
     DirDpto = mCoddirec
End Property


'**** Tiene mantenimiento

Public Property Let conMante(ByVal vData As Boolean)
     mTieneman = vData
End Property

Public Property Get conMante() As Boolean
     conMante = mTieneman
End Property


'**** Num. MAntenimiento

Public Property Let numMante(ByVal vData As String)
     mNummante = vData
End Property

Public Property Get numMante() As String
     numMante = mNummante
End Property


'**** Cod. tipo movim.

Public Property Let tipoMov(ByVal vData As String)
     mCodTipoM = vData
End Property

Public Property Get tipoMov() As String
     tipoMov = mCodTipoM
End Property


'**** Fecha Venta

Public Property Let FechaVta(ByVal vData As String)
     mFechavta = vData
End Property

Public Property Get FechaVta() As String
     FechaVta = mFechavta
End Property


'**** Num. Albaran

Public Property Let NumAlbaran(ByVal vData As String)
     mNumalbar = vData
End Property

Public Property Get NumAlbaran() As String
     NumAlbaran = mNumalbar
End Property


'**** Num. Linea Albaran

Public Property Let NumLinAlb(ByVal vData As String)
     mNumLine1 = vData
End Property

Public Property Get NumLinAlb() As String
     NumLinAlb = mNumLine1
End Property



'**** Num. Factura

Public Property Let NumFactura(ByVal vData As String)
     mNumFactu = vData
End Property

Public Property Get NumFactura() As String
     NumFactura = mNumFactu
End Property



'**** Cod. Proveedor

Public Property Let Proveedor(ByVal vData As Long)
     mCodProve = vData
End Property

Public Property Get Proveedor() As Long
     Proveedor = mCodProve
End Property

'**** Num. Albaran Proveedor (compras)

Public Property Let NumAlbProve(ByVal vData As String)
     mNumalbpr = vData
End Property

Public Property Get NumAlbProve() As String
     NumAlbProve = mNumalbpr
End Property


'**** Fecha Compra

Public Property Let fechaCom(ByVal vData As String)
     mFechaCom = vData
End Property

Public Property Get fechaCom() As String
     fechaCom = mFechaCom
End Property



'**** Num. Linea Albaran Proveedor

Public Property Let NumLinAlbPr(ByVal vData As String)
     mNumline2 = vData
End Property

Public Property Get NumLinAlbPr() As String
     NumLinAlbPr = mNumline2
End Property


'------------------------------------------------
'Procedimientos del modulo CNumSerie
'------------------------------------------------

Public Function InsertarNumSerie() As Boolean
Dim SQL As String

    On Error GoTo EInsertar

    If TipoArt = "" Then
        TipoArt = DevuelveDesdeBDNew(conAri, "sartic", "codtipar", "codartic", Articulo, "T")
    End If

    'Insertar en la tabla sserie
    SQL = "INSERT INTO sserie (numserie, codartic, codtipar, codclien, coddirec,tieneman, nummante, ultrepar, fingaran, "
    SQL = SQL & " codtipom, numfactu, fechavta, numalbar, numline1, codprove, numalbpr, fechacom, numline2) "
    
    SQL = SQL & " VALUES (" & DBSet(numSerie, "T") & ", " & DBSet(Articulo, "T") & ", " & DBSet(TipoArt, "T") & ","
    SQL = SQL & DBSet(Cliente, "N", "S") & "," & DBSet(DirDpto, "N", "S") & ","
    SQL = SQL & DBSet(conMante, "B") & "," & DBSet(numMante, "T") & ","
    SQL = SQL & ValorNulo & "," & DBSet(FinGarantia, "F") & "," 'ult. repar, fin garantia
    SQL = SQL & DBSet(tipoMov, "T") & "," & ValorNulo & "," & DBSet(FechaVta, "F", "S") & "," 'codtipom,numfactu, fechavta
    SQL = SQL & DBSet(NumAlbaran, "N", "S") & "," & DBSet(NumLinAlb, "N", "S") & "," 'numalbar, numline1
    SQL = SQL & DBSet(Proveedor, "N", "S") & ","
    SQL = SQL & DBSet(NumAlbProve, "T") & "," & DBSet(fechaCom, "F") & "," & DBSet(NumLinAlbPr, "N", "S") & ")"
    
    Conn.Execute SQL
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarNumSerie = False
        MsgBox "Error al insertar el N� de Serie: " & numSerie & " del art�culo: " & Articulo, vbExclamation
    Else
        InsertarNumSerie = True
    End If
End Function



Public Function ActualizarNumSerie(DeVentas As Boolean) As Boolean
'UPDATE en la tabla sserie de los campos de venta o compra
Dim SQL As String

    On Error GoTo EActualizar

    'Insertar en la tabla sserie
    If DeVentas Then 'Se llama desde Ventas a Clientes
        SQL = "UPDATE sserie SET codclien=" & DBSet(Cliente, "N", "S") & ","
        SQL = SQL & " coddirec=" & DBSet(DirDpto, "N", "S") & ","
        SQL = SQL & " tieneman=" & DBSet(conMante, "B") & ","
        SQL = SQL & " fingaran=" & DBSet(FinGarantia, "F", "S") & ","
        SQL = SQL & " codtipom=" & DBSet(tipoMov, "T", "S") & ","
        SQL = SQL & " fechavta=" & DBSet(FechaVta, "F", "S") & ","
        SQL = SQL & " numalbar=" & DBSet(NumAlbaran, "N", "S") & ","
        SQL = SQL & " numfactu=" & DBSet(NumFactura, "N", "S") & ","
        SQL = SQL & " numline1=" & DBSet(NumLinAlb, "N", "S")
        
    Else 'se llama desde Compras a Proveedores
        SQL = "UPDATE sserie SET codprove=" & DBSet(Proveedor, "N", "S") & ","
        SQL = SQL & "numalbpr=" & DBSet(NumAlbProve, "T") & ","
        SQL = SQL & "fechacom=" & DBSet(fechaCom, "F", "S") & ","
        SQL = SQL & "numline2=" & DBSet(NumLinAlbPr, "N", "S")
    End If
    
    SQL = SQL & " WHERE numserie=" & DBSet(numSerie, "T") & " AND codartic=" & DBSet(Articulo, "T")
      
    Conn.Execute SQL
    
EActualizar:
    If Err.Number <> 0 Then
        ActualizarNumSerie = False
    Else
        ActualizarNumSerie = True
    End If
End Function




Public Function EliminarNumSerie() As Boolean
'Elimina un n� serie de un articulo
Dim SQL As String
    
    On Error GoTo EEliminar

    SQL = "DELETE FROM sserie "
    SQL = SQL & " WHERE numserie=" & DBSet(numSerie, "T") & " AND codartic=" & DBSet(Articulo, "T")

    Conn.Execute SQL

EEliminar:
    If Err.Number <> 0 Then
        EliminarNumSerie = False
    Else
        EliminarNumSerie = True
    End If
End Function


Public Function BorrarNumSeriesAlbVta() As Boolean
'Actualiza los campos cliente,codtipom,numfactu,numalbar,fehcavta,numline1
'de todos los n� serie de un articulo asignados a una determinada linea de un albaran de venta
Dim SQL As String
    
    On Error GoTo EBorrar

    SQL = "UPDATE sserie SET  codclien=" & ValorNulo & ", codtipom=" & ValorNulo & ","
    SQL = SQL & "numfactu=" & ValorNulo & ", fechavta=" & ValorNulo & ", numalbar= " & ValorNulo & ",numline1=" & ValorNulo
    SQL = SQL & " WHERE codtipom=" & DBSet(tipoMov, "T") & " AND numalbar=" & NumAlbaran & " AND numline1=" & NumLinAlb & " and year(fechavta)=" & Year(Me.FechaVta)

    Conn.Execute SQL

EBorrar:
    If Err.Number <> 0 Then
        BorrarNumSeriesAlbVta = False
    Else
        BorrarNumSeriesAlbVta = True
    End If
End Function




Public Function TieneAsignadaVenta() As Boolean
'Comprueba si el n� de serie de un articulo ha sido asignado a algun albaran de venta
'o factura de venta en cuyo caso no se podr� eliminar el n� de serie
Dim SQL As String
    
    On Error GoTo EAsignada
    
    If Me.NumAlbaran <> "" Or Me.NumFactura <> "" Then
        TieneAsignadaVenta = True
    Else
        TieneAsignadaVenta = False
    End If
    
    Exit Function
    
EAsignada:
    TieneAsignadaVenta = False
    MuestraError Err.Number, "Comprovando si n� de serie asignado a una venta", Err.Description
End Function






Public Sub ObtenFechaFinGarantia(codArt As String, Fecha As String)
'calcula la fecha fin de garantia a partir de los dias de garantia del articulo
Dim cad As String
    On Error GoTo EFechaFin

    'obtenemos los dias de garantia del articulo
    cad = DevuelveDesdeBDNew(conAri, "sartic", "garantia", "codartic", codArt, "T")
    'fin garantia= fecha albaran + dias de garantia
    If Fecha <> "" Then FinGarantia = CStr(CDate(Fecha) + CInt(ComprobarCero(cad)))
EFechaFin:
    If Err.Number <> 0 Then Err.Clear
End Sub
