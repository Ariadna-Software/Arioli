VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLOG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'ACCIONES DISPONIBLES
'
'       1.- Eliminar facturas
'       2.- Facturacion
'       3.- Contabilizar facturas
'       4.- Correccion precios
'       5.- Facturacion desde hasta COMPRAS
'       6.- Facturacion TICKETS agrupados
'
'       7.- Eliminar articulo
'       8.- Updatear en referencias
'       9.- Partidas(LOTES) Solo podran usuarios nivel=0

'       10.- Ajuste cantidad albaran desde pistola de expedicion
'       11.- Cambio proveedor en albaranes prov
'       12.- Cambiar fecha factura
'       13.- Reestablecer albaran
'       14.- Pesaje. Pesando produccion sin haber pesado botellas
'       15.- Liquidacion VALL
Public Function Insertar(Accion As Byte, ByRef ElUsuario As Usuario, Descripcion As String) As Boolean
Dim SQL As String

    On Error GoTo EI
    Insertar = False
    
    SQL = "insert into `slog` (`fecha`,`accion`,`usuario`,`pc`,`descripcion`) values ( "
    SQL = SQL & " now()," & Accion & " ,'" & DevNombreSQL(ElUsuario.Login) & "','"
    SQL = SQL & DevNombreSQL(ElUsuario.PC) & "',"
    If Descripcion = "" Then
        SQL = SQL & "NULL)"
    Else
        SQL = SQL & "'" & DevNombreSQL(Descripcion) & "')"
    End If
    conn.Execute SQL
    
    
    Insertar = True
    Exit Function
EI:
    MuestraError Err.Number, Err.Description

End Function




Public Function DevuelveAcciones(ByRef ListaAcciones As Collection) As Boolean
 On Error GoTo EDevuelveAcciones
    ListaAcciones.Add "1|Eliminar facturas|"
    ListaAcciones.Add "2|Facturación|"
    ListaAcciones.Add "3|Contabilizar facturas|"
    ListaAcciones.Add "4|Corrección precios|"
    ListaAcciones.Add "5|Facturacion COMPRAS|"
    ListaAcciones.Add "6|Facturacion TICKETS agrupados|"
    ListaAcciones.Add "7|Eliminar articulo|"
    ListaAcciones.Add "8|Updatear referencias|"
    ListaAcciones.Add "9|Mto partidas-lotes|"
    ListaAcciones.Add "10|Ajuste albaran desde expedicion|"
    ListaAcciones.Add "11|Cambio proveedor en albaran|"
    ListaAcciones.Add "12|Cambiar fecha factura|"
    ListaAcciones.Add "13|Reestablecer albaran|"
    ListaAcciones.Add "14|Botellas sin pesar|"
    ListaAcciones.Add "15|Liquidacion|"
    DevuelveAcciones = True
    Exit Function
EDevuelveAcciones:
    MuestraError Err.Number
End Function

