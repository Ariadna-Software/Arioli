VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLineaProCompo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'variables locales para almacenar los valores de las propiedades
'Private mvarCodProduccion As Long 'copia local   INTERNA
'Private mvaridLiProd As Long 'copia local        INTERNA
'Private mvarLINEA As Byte 'copia local           INTERNA
'Private mvarLoteTrazabilidad As Long 'c          INTERNA
Private mvarcodarticCompo As String 'copia local
Private mvarNomArticCompo As String 'copia local
Private mvarEsMateriaPrima As Boolean
Private mvarLoteMateria As String 'copia local
Private mvarCantidadEstimada As Currency

Private mvarFactorMultiplicador As Currency   'Sera el factor de conversion *  la cantidad que tenga el articulo en sarti1
                                              'Es decir.  Si es un aceite en botella de 250ml seran
                                              ' 0.956 * 0.250
                                              'Si es una botella sera 1*1
Private mvarFactorConversion As Currency
Private mvarCantidadSarti1 As Currency



Private mvarCantidadUtilizada As Currency


Public Property Get NomArticCompo() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.NomArtic
    NomArticCompo = mvarNomArticCompo
End Property



Public Property Get codarticCompo() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.codartic
    codarticCompo = mvarcodarticCompo
End Property


Public Property Get EsMateriaPrima() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.codartic
    EsMateriaPrima = mvarEsMateriaPrima
End Property

Public Property Get LoteMateria() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.codartic
    LoteMateria = mvarLoteMateria
End Property

Public Property Get CantidadEstimada() As Currency
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.codartic
    CantidadEstimada = mvarCantidadEstimada
End Property


Public Property Get CantidadUtilizada() As Currency
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.codartic
    CantidadUtilizada = mvarCantidadUtilizada
End Property

Public Property Let CantidadUtilizada(ByVal vData As Currency)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.NomArtic = 5
    mvarCantidadUtilizada = vData
End Property


''''Public Property Get FactorConversion() As Currency
'''''se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'''''Syntax: Debug.Print X.codartic
''''    FactorConversion = mvarFactorConversion
''''End Property


'Sera el factor de conversion *  la cantidad que tenga el articulo en sarti1
'Es decir.  Si es un aceite en botella de 250ml seran
' 0.956 * 0.250
'Si es una botella sera 1*1
Public Property Get FactorMultiplicador() As Currency
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.codartic
    FactorMultiplicador = mvarFactorMultiplicador
End Property



Public Function AsignarValoresComponentes(codart As String, NomArt As String, CantNecesaria As Currency, vFactorConversion As Currency, vCantidadEnSarti1 As Currency) As Boolean
Dim C As Currency

    mvarcodarticCompo = codart
    mvarNomArticCompo = NomArt
    mvarEsMateriaPrima = vFactorConversion <> 1
    mvarCantidadEstimada = CantNecesaria
    ''''mvarFactorConversion = vFactorConversion
    C = 1
    If vFactorConversion <> 1 Then
        C = vFactorConversion
    End If
    If vCantidadEnSarti1 <> 1 Then C = C * vCantidadEnSarti1
    mvarFactorMultiplicador = C
    
    'AQUI NO ASIGNO EL NUMERO DE LOTE
    
End Function


Public Function AsignarLote(Lot As String)
    mvarLoteMateria = Lot
End Function
