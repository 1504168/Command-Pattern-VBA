VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("CommandPattern.Request")
'@PredeclaredId
'@Description : Parameter for the concrete implementation of the Command class(BuStock and SellStock).
'                       Plain Object( Model)

Option Explicit

Private Type TStock
    StockName As String
    Quantity As Long
End Type

Private This As TStock

Public Property Get StockName() As String
    StockName = This.StockName
End Property

Public Property Let StockName(ByVal RHS As String)
    This.StockName = RHS
End Property

Public Property Get Quantity() As Long
    Quantity = This.Quantity
End Property

Public Property Let Quantity(ByVal RHS As Long)
    This.Quantity = RHS
End Property

Public Function Create(GivenStockName As String, GivenQuantity As Long) As Stock
    Dim CurrentStock As Stock
    Set CurrentStock = New Stock
    With CurrentStock
        .StockName = GivenStockName
        .Quantity = GivenQuantity
    End With
    Set Create = CurrentStock
End Function


