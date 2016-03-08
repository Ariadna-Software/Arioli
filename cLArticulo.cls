VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLArticulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvarcodartic As String 'nombre del cliente
Private mvarcantidadnecesaria As Single
Private mvarcantidadStock As Single



Public Property Let codartic(ByVal vData As String)
     mvarcodartic = vData
End Property

Public Property Get codartic() As String
     codartic = mvarcodartic
End Property


Public Property Let CantidadNecesaria(ByVal vData As Single)
     mvarcantidadnecesaria = vData
End Property

Public Property Get CantidadNecesaria() As Single
     CantidadNecesaria = mvarcantidadnecesaria
End Property


Public Property Let cantidadStock(ByVal vData As Single)
     mvarcantidadStock = vData
End Property

Public Property Get cantidadStock() As Single
     cantidadStock = mvarcantidadStock
End Property


Public Sub Reestablecer()
    mvarcantidadStock = 0
    mvarcantidadnecesaria = 0
    mvarcodartic = ""
End Sub
